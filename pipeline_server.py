from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import threading
import time
import uuid
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any

from flask import jsonify, request

import app as core

APP_VERSION = "1.5"
core.APP_VERSION = APP_VERSION
flask_app = core.app
# Large journal issues can easily exceed 1 GB once hundreds of PDFs are selected.
flask_app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024 * 1024

_PIPELINES: dict[str, dict[str, Any]] = {}
_PIPE_LOCK = threading.RLock()
_JOB_FILE_LOCK = threading.RLock()
_WORD_LOCAL = threading.local()
# Newer app.py exposes ROOT; fall back to this file's directory for older cores (e.g. v1.2).
_CORE_ROOT = getattr(core, "ROOT", None) or Path(__file__).resolve().parent
_CACHE_ROOT = Path(_CORE_ROOT) / "runtime" / "pipeline_cache_v15"
_CACHE_WORD = _CACHE_ROOT / "word"
_CACHE_REVISED = _CACHE_ROOT / "revised"
for _d in (_CACHE_WORD, _CACHE_REVISED):
    _d.mkdir(parents=True, exist_ok=True)

_ORIG_PREPARE_WORD = core.prepare_word_document
_ORIG_SAVE_JOB = getattr(core, "_save_job", None)


def _atomic_write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def _safe_asdict(obj: Any) -> dict:
    if is_dataclass(obj):
        return asdict(obj)
    if isinstance(obj, dict):
        return obj
    out = {}
    for name in ("first_author", "title", "article_folder", "figures", "warnings"):
        if hasattr(obj, name):
            value = getattr(obj, name)
            if name == "figures":
                value = [_safe_asdict(x) for x in value]
            out[name] = value
    return out


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _slug_name(name: str, limit: int = 120) -> str:
    try:
        return core.sanitize_filename(name, limit)
    except Exception:
        name = re.sub(r"[\\/:*?\"<>|]+", "_", name).strip()
        return (name[:limit] or "file")


def _json_clone(value: Any) -> Any:
    return json.loads(json.dumps(value, ensure_ascii=False, default=str))


def _get_pipeline(pid: str) -> dict | None:
    with _PIPE_LOCK:
        state = _PIPELINES.get(pid)
        return _json_clone(state) if state else None


def _mutate_pipeline(pid: str, fn) -> None:
    with _PIPE_LOCK:
        state = _PIPELINES.get(pid)
        if state is None:
            return
        fn(state)
        state["updated_ts"] = time.time()


def _set_pipeline(pid: str, **values) -> None:
    def mut(state):
        state.update(values)
    _mutate_pipeline(pid, mut)


def _article_update(pid: str, article_id: str, **values) -> None:
    def mut(state):
        for article in state.get("articles", []):
            if article.get("id") == article_id:
                article.update(values)
                break
    _mutate_pipeline(pid, mut)


def _recalc_pipeline(pid: str) -> None:
    def mut(state):
        arts = state.get("articles", [])
        total = max(1, len(arts))
        done = sum(1 for a in arts if a.get("status") in {"done", "warning", "error"})
        ready = sum(int(a.get("ready_pairs", 0) or 0) for a in arts)
        figures = sum(int(a.get("figure_total", 0) or 0) for a in arts)
        state["completed_articles"] = done
        state["ready_pairs"] = ready
        state["discovered_figures"] = figures
        state["percent"] = min(99, int(done * 100 / total)) if state.get("status") == "running" else 100
        started = float(state.get("started_ts", time.time()))
        state["elapsed_seconds"] = max(0, int(time.time() - started))
        if done and done < total:
            per = state["elapsed_seconds"] / done
            state["eta_seconds"] = max(0, int(per * (total - done)))
        else:
            state["eta_seconds"] = 0
    _mutate_pipeline(pid, mut)


def _public_pair(pair: dict) -> dict:
    keys = [
        "id", "author", "title", "figure_no", "caption", "original_method",
        "original_confidence", "revised_source", "revised_warnings", "status", "comparison",
    ]
    return {k: pair.get(k) for k in keys}


def _merge_live_comparisons(job_id: str, job_data: dict) -> None:
    """Preserve AI results if the editor starts comparing while later articles are still processing."""
    p = Path(core.JOBS_DIR) / job_id / "job.json"
    if not p.exists():
        return
    try:
        live = json.loads(p.read_text(encoding="utf-8"))
        by_id = {x.get("id"): x for x in live.get("pairs", []) if x.get("id")}
        for pair in job_data.get("pairs", []):
            old = by_id.get(pair.get("id"))
            if old and old.get("comparison") is not None:
                pair["comparison"] = old.get("comparison")
                pair["status"] = old.get("status", pair.get("status"))
    except Exception:
        pass


def _save_incremental_job(job_id: str, job_data: dict) -> None:
    with _JOB_FILE_LOCK:
        _merge_live_comparisons(job_id, job_data)
        _atomic_write_json(Path(core.JOBS_DIR) / job_id / "job.json", job_data)


if _ORIG_SAVE_JOB is not None:
    def _locked_core_save_job(job_id: str, data: dict):
        with _JOB_FILE_LOCK:
            return _ORIG_SAVE_JOB(job_id, data)
    core._save_job = _locked_core_save_job


class _ReusableWord:
    def __init__(self):
        self.word = None
        self.pythoncom = None

    def start(self):
        if self.word is not None:
            return
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        self.pythoncom = pythoncom
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.ScreenUpdating = False
        except Exception:
            pass
        self.word = word

    def restart(self):
        self.close()
        self.start()

    def convert(self, src: Path, out_pdf: Path, normalized_docx: Path) -> list[str]:
        self.start()
        warnings: list[str] = []
        last_exc = None
        for attempt in range(2):
            doc = None
            try:
                out_pdf.parent.mkdir(parents=True, exist_ok=True)
                normalized_docx.parent.mkdir(parents=True, exist_ok=True)
                doc = self.word.Documents.Open(
                    str(src.resolve()), ReadOnly=True, AddToRecentFiles=False,
                    ConfirmConversions=False, Visible=False,
                )
                # Always create a disposable DOCX copy. It lets the proven extraction code
                # inspect the actual embedded Word/OLE/Visio object before each caption.
                doc.SaveAs2(str(normalized_docx.resolve()), FileFormat=16, AddToRecentFiles=False)
                doc.ExportAsFixedFormat(
                    OutputFileName=str(out_pdf.resolve()),
                    ExportFormat=17,  # wdExportFormatPDF
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    Range=0,
                    Item=0,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=1,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False,
                )
                return warnings
            except Exception as exc:
                last_exc = exc
                warnings.append(f"Word 批处理第 {attempt + 1} 次失败：{exc}")
                try:
                    if doc is not None:
                        doc.Close(False)
                except Exception:
                    pass
                doc = None
                if attempt == 0:
                    self.restart()
                    continue
                raise
            finally:
                try:
                    if doc is not None:
                        doc.Close(False)
                except Exception:
                    pass
        if last_exc:
            raise last_exc
        return warnings

    def close(self):
        if self.word is not None:
            try:
                self.word.Quit()
            except Exception:
                pass
            self.word = None
        if self.pythoncom is not None:
            try:
                self.pythoncom.CoUninitialize()
            except Exception:
                pass
            self.pythoncom = None


def _pipeline_prepare_word(src: Path, out_pdf: Path, normalized_docx: Path):
    session = getattr(_WORD_LOCAL, "session", None)
    if session is None:
        return _ORIG_PREPARE_WORD(src, out_pdf, normalized_docx)
    return session.convert(Path(src), Path(out_pdf), Path(normalized_docx))


# The proven process_one_word still owns figure detection/cropping. Only the Word lifecycle is reused.
core.prepare_word_document = _pipeline_prepare_word


def _word_cache_key(src: Path) -> str:
    return "wordobj-v15-" + _sha256(src)


def _revised_cache_key(src: Path) -> str:
    return "revised-v15-" + _sha256(src)


def _process_word_cached(src: Path, original_root: Path, work: Path) -> tuple[dict, bool]:
    key = _word_cache_key(src)
    cache = _CACHE_WORD / key
    meta = cache / "article.json"
    cached_originals = cache / "originals"
    if meta.exists() and cached_originals.exists():
        data = json.loads(meta.read_text(encoding="utf-8"))
        folder = str(data.get("article_folder", ""))
        if folder:
            dest = original_root / folder
            if dest.exists():
                shutil.rmtree(dest, ignore_errors=True)
            shutil.copytree(cached_originals, dest)
            return data, True
    article_obj = core.process_one_word(src, original_root, work)
    data = _safe_asdict(article_obj)
    folder = str(data.get("article_folder", ""))
    try:
        if cache.exists():
            shutil.rmtree(cache, ignore_errors=True)
        cache.mkdir(parents=True, exist_ok=True)
        meta.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        source_folder = original_root / folder
        if folder and source_folder.exists():
            shutil.copytree(source_folder, cached_originals)
    except Exception:
        core.logging.exception("写入 Word 提取缓存失败：%s", src)
    return data, False


def _render_revised_cached(src: Path, out_png: Path) -> tuple[list[str], bool]:
    key = _revised_cache_key(src)
    cache_png = _CACHE_REVISED / f"{key}.png"
    cache_meta = _CACHE_REVISED / f"{key}.json"
    if cache_png.exists():
        out_png.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(cache_png, out_png)
        warnings = []
        if cache_meta.exists():
            try:
                warnings = json.loads(cache_meta.read_text(encoding="utf-8")).get("warnings", [])
            except Exception:
                pass
        return warnings, True
    warnings = core.render_revised_file(src, out_png)
    try:
        shutil.copy2(out_png, cache_png)
        cache_meta.write_text(json.dumps({"warnings": warnings}, ensure_ascii=False), encoding="utf-8")
    except Exception:
        core.logging.exception("写入重制图缓存失败：%s", src)
    return warnings, False


def _figure_no(text: str) -> str | None:
    try:
        return core._figure_no_from_name(text)
    except Exception:
        m = re.search(r"图\s*([0-9０-９]+)", text or "", re.I)
        return core.normalize_digits(m.group(1)) if m else None


def _match_revised(article: dict, fig: dict, revised_items: list[dict], used: set[int], word_total: int) -> tuple[dict | None, str]:
    author = str(article.get("first_author") or "").strip()
    no = str(fig.get("figure_no") or "").strip()
    candidates = []
    for item in revised_items:
        if item["idx"] in used or item.get("figure_no") != no:
            continue
        label = item.get("relative_path") or item.get("source_name") or ""
        if author and author != "未知作者" and author in label:
            candidates.append(item)
    if len(candidates) == 1:
        return candidates[0], "author+figure"
    if len(candidates) > 1:
        return None, "同一作者+图号命中多个重制文件，无法安全自动配对"
    if word_total == 1:
        candidates = [x for x in revised_items if x["idx"] not in used and x.get("figure_no") == no]
        if len(candidates) == 1:
            return candidates[0], "single-article-figure"
    return None, "未找到唯一的“第一作者+图号”重制图"


def _article_status_payload(article: dict) -> dict:
    return {
        "id": article.get("id"),
        "source_name": article.get("source_name"),
        "status": article.get("status"),
        "author": article.get("author", ""),
        "title": article.get("title", ""),
        "figure_total": article.get("figure_total", 0),
        "ready_pairs": article.get("ready_pairs", 0),
        "unmatched_count": article.get("unmatched_count", 0),
        "cache_hit": article.get("cache_hit", False),
        "error": article.get("error", ""),
        "pairs": article.get("pairs", []),
    }


def _run_pipeline(pid: str, word_entries: list[dict], revised_items: list[dict]) -> None:
    job = Path(core.JOBS_DIR) / pid
    original_root = job / "originals"
    revised_root = job / "revised_rendered"
    work = job / "work"
    for d in (original_root, revised_root, work):
        d.mkdir(parents=True, exist_ok=True)

    job_data = {
        "job_id": pid,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "articles": [],
        "word_failures": [],
        "pairs": [],
        "unmatched_originals": [],
        "unmatched_revised": [],
        "pipeline_version": APP_VERSION,
    }
    _save_incremental_job(pid, job_data)
    used_revised: set[int] = set()
    session = _ReusableWord()
    _WORD_LOCAL.session = session
    try:
        for ai, entry in enumerate(word_entries):
            article_id = entry["article_id"]
            src = Path(entry["path"])
            _set_pipeline(pid, current_article_id=article_id, phase="word", detail=f"正在提取第 {ai + 1}/{len(word_entries)} 篇：{entry['source_name']}")
            _article_update(pid, article_id, status="extracting", error="")
            try:
                article, cache_hit = _process_word_cached(src, original_root, work)
            except Exception as exc:
                core.logging.exception("流水线 Word 处理失败：%s", src)
                job_data["word_failures"].append({"file": entry["source_name"], "error": str(exc)})
                _article_update(pid, article_id, status="error", error=str(exc)[:500])
                _recalc_pipeline(pid)
                _save_incremental_job(pid, job_data)
                continue

            author = str(article.get("first_author") or "未知作者")
            title = str(article.get("title") or "")
            figures = list(article.get("figures") or [])
            article["pipeline_article_id"] = article_id
            job_data["articles"].append(article)
            _article_update(
                pid, article_id,
                status="matching", author=author, title=title,
                figure_total=len(figures), ready_pairs=0, unmatched_count=0,
                cache_hit=cache_hit, pairs=[],
            )
            _set_pipeline(pid, phase="revised", detail=f"{author}：原稿提取完成，开始逐图匹配与渲染")
            _save_incremental_job(pid, job_data)

            unmatched_count = 0
            for fi, fig in enumerate(figures):
                no = str(fig.get("figure_no") or "")
                caption = str(fig.get("caption") or "")
                candidate, reason = _match_revised(article, fig, revised_items, used_revised, len(word_entries))
                if candidate is None:
                    unmatched_count += 1
                    un = {
                        "article_index": ai, "figure_index": fi, "author": author, "title": title,
                        "article_folder": article.get("article_folder", ""), "figure_no": no,
                        "caption": caption,
                    }
                    job_data["unmatched_originals"].append(un)
                    _article_update(pid, article_id, unmatched_count=unmatched_count)
                    continue

                _set_pipeline(pid, detail=f"{author}：正在准备图{no}（{fi + 1}/{len(figures)}）")
                out = revised_root / f"{article_id}_fig_{_slug_name(no, 24)}_{candidate['idx']:04d}.png"
                try:
                    warnings, revised_cache_hit = _render_revised_cached(Path(candidate["path"]), out)
                except Exception as exc:
                    core.logging.exception("流水线重制图渲染失败：%s", candidate["path"])
                    unmatched_count += 1
                    candidate["render_error"] = str(exc)
                    job_data["unmatched_originals"].append({
                        "article_index": ai, "figure_index": fi, "author": author, "title": title,
                        "article_folder": article.get("article_folder", ""), "figure_no": no,
                        "caption": caption,
                    })
                    _article_update(pid, article_id, unmatched_count=unmatched_count)
                    continue

                used_revised.add(candidate["idx"])
                original_path = original_root / str(article.get("article_folder", "")) / str(fig.get("filename", ""))
                pair_id = f"p{len(job_data['pairs']) + 1:05d}"
                pair = {
                    "article_index": ai,
                    "figure_index": fi,
                    "author": author,
                    "title": title,
                    "article_folder": article.get("article_folder", ""),
                    "figure_no": no,
                    "caption": caption,
                    "original_path": str(original_path.resolve()),
                    "original_method": fig.get("extraction_method"),
                    "original_confidence": fig.get("confidence"),
                    "id": pair_id,
                    "revised_path": str(out.resolve()),
                    "revised_source": candidate.get("relative_path") or candidate.get("source_name"),
                    "revised_warnings": warnings,
                    "status": "PENDING",
                    "comparison": None,
                    "pipeline_article_id": article_id,
                    "revised_cache_hit": revised_cache_hit,
                }
                job_data["pairs"].append(pair)
                _save_incremental_job(pid, job_data)

                def add_pair(state):
                    for a in state.get("articles", []):
                        if a.get("id") == article_id:
                            a.setdefault("pairs", []).append(_public_pair(pair))
                            a["ready_pairs"] = len(a["pairs"])
                            a["status"] = "ready_partial" if len(a["pairs"]) < len(figures) else "ready"
                            a["unmatched_count"] = unmatched_count
                            break
                _mutate_pipeline(pid, add_pair)
                _recalc_pipeline(pid)

            final_status = "warning" if unmatched_count else "done"
            _article_update(pid, article_id, status=final_status, unmatched_count=unmatched_count)
            _recalc_pipeline(pid)
            _save_incremental_job(pid, job_data)

        # Anything never consumed is a revised image with no safe original match.
        for item in revised_items:
            if item["idx"] not in used_revised:
                reason = item.get("render_error") or "未被任何已识别文章的“第一作者+图号”唯一匹配"
                job_data["unmatched_revised"].append({
                    "source_name": item.get("source_name"),
                    "relative_path": item.get("relative_path"),
                    "reason": reason,
                    "warnings": [],
                })
        _save_incremental_job(pid, job_data)
        _set_pipeline(pid, status="done", phase="done", detail="全部文章处理完成。已完成的文章在处理过程中即可提前查看。", current_article_id="")
        _recalc_pipeline(pid)
    except Exception as exc:
        core.logging.exception("流水线任务异常")
        _set_pipeline(pid, status="error", phase="error", detail=str(exc)[:600], error=str(exc)[:1000])
        _recalc_pipeline(pid)
    finally:
        try:
            session.close()
        finally:
            _WORD_LOCAL.session = None


def _relative_paths(field: str, count: int) -> list[str]:
    try:
        vals = json.loads(request.form.get(field, "[]"))
        if not isinstance(vals, list):
            vals = []
        vals = [str(x) for x in vals]
    except Exception:
        vals = []
    vals.extend([""] * max(0, count - len(vals)))
    return vals[:count]


def pipeline_prepare_compare():
    try:
        core._cleanup_old_jobs()
    except Exception:
        pass
    word_files = [f for f in request.files.getlist("word_files") if Path(f.filename).suffix.lower() in core.WORD_EXTS]
    revised_files = [f for f in request.files.getlist("revised_files") if Path(f.filename).suffix.lower() in core.REVISED_EXTS]
    if not word_files:
        return jsonify({"error": "没有选择 .doc / .docx 原稿。"}), 400
    if not revised_files:
        return jsonify({"error": "没有选择 PDF / JPG / PNG 重制图。"}), 400

    requested = str(request.form.get("progress_id", "")).strip().lower()
    pid = requested if re.fullmatch(r"[0-9a-f]{12,40}", requested) else uuid.uuid4().hex[:20]
    job = Path(core.JOBS_DIR) / pid
    incoming_words = job / "incoming_words"
    incoming_revised = job / "incoming_revised"
    for d in (incoming_words, incoming_revised):
        d.mkdir(parents=True, exist_ok=True)

    revised_relpaths = _relative_paths("revised_relpaths", len(revised_files))
    word_entries: list[dict] = []
    articles_state: list[dict] = []
    for idx, f in enumerate(word_files):
        name = Path(f.filename.replace("\\", "/")).name
        src = incoming_words / f"{idx:03d}_{_slug_name(name, 140)}"
        f.save(src)
        aid = f"a{idx + 1:04d}"
        word_entries.append({"article_id": aid, "source_name": name, "path": str(src.resolve())})
        articles_state.append({
            "id": aid, "source_name": name, "status": "waiting", "author": "", "title": "",
            "figure_total": 0, "ready_pairs": 0, "unmatched_count": 0, "cache_hit": False,
            "pairs": [], "error": "",
        })

    revised_items: list[dict] = []
    for idx, f in enumerate(revised_files):
        orig_name = Path(f.filename.replace("\\", "/")).name
        ext = Path(orig_name).suffix.lower()
        src = incoming_revised / f"{idx:04d}_{_slug_name(Path(orig_name).stem, 100)}{ext}"
        f.save(src)
        rel = revised_relpaths[idx] or orig_name
        revised_items.append({
            "idx": idx,
            "source_name": orig_name,
            "relative_path": rel,
            "path": str(src.resolve()),
            "figure_no": _figure_no(rel),
        })

    with _PIPE_LOCK:
        _PIPELINES[pid] = {
            "job_id": pid,
            "status": "running",
            "phase": "starting",
            "detail": "文件已接收，正在启动按文章流水线。",
            "started_ts": time.time(),
            "updated_ts": time.time(),
            "elapsed_seconds": 0,
            "eta_seconds": 0,
            "percent": 0,
            "total_articles": len(word_entries),
            "completed_articles": 0,
            "total_revised": len(revised_items),
            "discovered_figures": 0,
            "ready_pairs": 0,
            "current_article_id": "",
            "articles": articles_state,
            "error": "",
        }

    thread = threading.Thread(target=_run_pipeline, args=(pid, word_entries, revised_items), daemon=True, name=f"jfe-pipeline-{pid[:6]}")
    thread.start()
    return jsonify({
        "job_id": pid,
        "pipeline": True,
        "status": "running",
        "article_count": len(word_entries),
        "revised_count": len(revised_items),
    })


# Replace only the long-running prepare endpoint. All proven image/AI/report endpoints remain from app.py.
flask_app.view_functions["prepare_compare"] = pipeline_prepare_compare


@flask_app.get("/api/pipeline/<pid>/status")
def pipeline_status(pid: str):
    if not re.fullmatch(r"[0-9a-f]{12,40}", pid or ""):
        return jsonify({"error": "无效任务 ID"}), 400
    state = _get_pipeline(pid)
    if not state:
        return jsonify({"error": "任务不存在或服务已重启"}), 404
    if state.get("started_ts"):
        state["elapsed_seconds"] = max(0, int(time.time() - float(state["started_ts"])))
    state.pop("started_ts", None)
    state.pop("updated_ts", None)
    state["articles"] = [_article_status_payload(x) for x in state.get("articles", [])]
    return jsonify(state)


@flask_app.after_request
def inject_pipeline_ui(response):
    if request.path == "/" and response.mimetype == "text/html":
        try:
            text = response.get_data(as_text=True)
            # Whatever older label the base template contains, expose the actual runtime version.
            text = re.sub(r"v1\.[0-9]+", f"v{APP_VERSION}", text)
            script = f'<script src="/static/pipeline.js?v={APP_VERSION}"></script>'
            if script not in text:
                text = text.replace("</body>", script + "\n</body>")
            response.set_data(text)
            response.headers["Content-Length"] = str(len(response.get_data()))
            response.headers["Cache-Control"] = "no-store"
        except Exception:
            core.logging.exception("注入 v1.5 流水线 UI 失败")
    return response


if __name__ == "__main__":
    port = int(os.environ.get("JFE_PORT", "8765"))
    flask_app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
