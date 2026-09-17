from __future__ import annotations

import hashlib
import json
import multiprocessing
import os
import re
import shutil
import sys
import threading
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import asdict
from pathlib import Path

from flask import jsonify, request

import app as core
from batch_worker import render_revised_worker

APP_VERSION = "1.4"
core.APP_VERSION = APP_VERSION
flask_app = core.app
# Local-only tool: allow genuinely large editorial batches.
flask_app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024 * 1024

_PROGRESS: dict[str, dict] = {}
_PROGRESS_LOCK = threading.Lock()
_PROGRESS_MAX_AGE = 6 * 3600
_BATCH_LOCK = threading.Lock()

ROOT = Path(__file__).resolve().parent
CACHE_ROOT = ROOT / "runtime" / "cache_v14"
WORD_CACHE = CACHE_ROOT / "word-object-first-v08"
REVISED_CACHE = CACHE_ROOT / "revised-render-v1"
WORD_CACHE.mkdir(parents=True, exist_ok=True)
REVISED_CACHE.mkdir(parents=True, exist_ok=True)

_ORIGINAL_PREPARE_WORD = core.prepare_word_document


def _clean_progress() -> None:
    cutoff = time.time() - _PROGRESS_MAX_AGE
    with _PROGRESS_LOCK:
        stale = [k for k, v in _PROGRESS.items() if float(v.get("updated_ts", 0)) < cutoff]
        for k in stale:
            _PROGRESS.pop(k, None)


def _get_state(pid: str) -> dict | None:
    with _PROGRESS_LOCK:
        state = _PROGRESS.get(pid)
        return dict(state) if state else None


def _set_state(pid: str, **values) -> None:
    if not pid:
        return
    with _PROGRESS_LOCK:
        state = _PROGRESS.setdefault(pid, {})
        state.update(values)
        state["updated_ts"] = time.time()


def _display_name(path) -> str:
    name = Path(str(path)).name
    return re.sub(r"^\d{3,4}_", "", name)


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _fmt_eta(seconds: float | int | None) -> str:
    if seconds is None:
        return ""
    seconds = max(0, int(seconds))
    if seconds < 60:
        return f"预计剩余约 {seconds} 秒"
    return f"预计剩余约 {seconds // 60} 分 {seconds % 60:02d} 秒"


def _phase_eta(started: float, done: int, total: int) -> int | None:
    if done <= 0 or total <= done:
        return 0 if total and done >= total else None
    elapsed = max(0.01, time.time() - started)
    return int((elapsed / done) * (total - done))


class _WordBatchSession:
    """Keep one hidden WINWORD.EXE alive for the entire batch.

    v1.2 launched and quit Word once per document. With 67 manuscripts that startup
    overhead dominates. This class preserves the exact ExportAsFixedFormat / SaveAs2
    behavior but reuses the same Word.Application instance, with one automatic restart
    if Office becomes unhealthy.
    """

    def __init__(self):
        self.pythoncom = None
        self.win32 = None
        self.word = None
        self.com_initialized = False

    def start(self) -> None:
        if not sys.platform.startswith("win"):
            return
        if not self.com_initialized:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            self.pythoncom = pythoncom
            self.win32 = win32com.client
            self.com_initialized = True
        if self.word is None:
            self.word = self.win32.DispatchEx("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = 0
            try:
                self.word.AutomationSecurity = 3
            except Exception:
                pass
            for attr in ("CheckSpellingAsYouType", "CheckGrammarAsYouType"):
                try:
                    setattr(self.word.Options, attr, False)
                except Exception:
                    pass

    def _drop_word(self) -> None:
        if self.word is not None:
            try:
                self.word.Quit()
            except Exception:
                pass
            self.word = None

    def close(self) -> None:
        self._drop_word()
        if self.com_initialized and self.pythoncom is not None:
            try:
                self.pythoncom.CoUninitialize()
            except Exception:
                pass
        self.com_initialized = False

    def render(self, src: Path, out_pdf: Path, normalized_docx: Path):
        if not sys.platform.startswith("win"):
            return _ORIGINAL_PREPARE_WORD(src, out_pdf, normalized_docx)

        last_exc = None
        for attempt in (1, 2):
            doc = None
            try:
                self.start()
                doc = self.word.Documents.Open(
                    str(src.resolve()),
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False,
                    ConfirmConversions=False,
                )
                doc.ExportAsFixedFormat(str(out_pdf.resolve()), 17)
                warnings = []
                if src.suffix.lower() == ".docx":
                    shutil.copy2(src, normalized_docx)
                else:
                    try:
                        doc.SaveAs2(str(normalized_docx.resolve()), FileFormat=16, AddToRecentFiles=False)
                    except Exception as exc:
                        warnings.append(f"临时 DOC→DOCX 归一化失败，将仅使用 PDF 回退：{exc}")
                try:
                    doc.Close(False)
                finally:
                    doc = None
                if out_pdf.exists() and out_pdf.stat().st_size > 0:
                    return warnings
                raise RuntimeError("Word 未生成有效 PDF")
            except Exception as exc:
                last_exc = exc
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                self._drop_word()
                if attempt == 1:
                    continue
                raise RuntimeError(f"Microsoft Word 转 PDF 失败：{exc}") from exc
        raise RuntimeError(f"Microsoft Word 转 PDF 失败：{last_exc}")


_WORD_CTX = threading.local()


def _batch_prepare_word_document(src: Path, out_pdf: Path, normalized_docx: Path):
    session = getattr(_WORD_CTX, "session", None)
    if session is not None:
        return session.render(src, out_pdf, normalized_docx)
    return _ORIGINAL_PREPARE_WORD(src, out_pdf, normalized_docx)


core.prepare_word_document = _batch_prepare_word_document


def _article_from_cache(cache_dir: Path, dest_root: Path, source_name: str):
    meta_path = cache_dir / "article.json"
    figures_dir = cache_dir / "figures"
    if not meta_path.exists() or not figures_dir.exists():
        return None
    try:
        data = json.loads(meta_path.read_text(encoding="utf-8"))
        figures = [core.FigureInfo(**x) for x in data.get("figures", [])]
        article = core.ArticleInfo(
            source_file=source_name,
            first_author=data["first_author"],
            title=data["title"],
            article_folder=data["article_folder"],
            figures=figures,
            warnings=list(data.get("warnings", [])),
        )
        dest = dest_root / article.article_folder
        if dest.exists():
            shutil.rmtree(dest, ignore_errors=True)
        shutil.copytree(figures_dir, dest)
        try:
            os.utime(cache_dir, None)
        except Exception:
            pass
        return article
    except Exception:
        core.logging.exception("读取 Word 缓存失败: %s", cache_dir)
        return None


def _save_article_cache(cache_dir: Path, article, source_root: Path) -> None:
    try:
        tmp = cache_dir.with_name(cache_dir.name + ".tmp-" + os.urandom(4).hex())
        shutil.rmtree(tmp, ignore_errors=True)
        tmp.mkdir(parents=True, exist_ok=True)
        src_figures = source_root / article.article_folder
        shutil.copytree(src_figures, tmp / "figures")
        (tmp / "article.json").write_text(
            json.dumps(asdict(article), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        if cache_dir.exists():
            shutil.rmtree(tmp, ignore_errors=True)
        else:
            tmp.replace(cache_dir)
    except Exception:
        core.logging.exception("写入 Word 缓存失败: %s", cache_dir)


def _revised_cache_paths(file_hash: str) -> tuple[Path, Path]:
    return REVISED_CACHE / f"{file_hash}.png", REVISED_CACHE / f"{file_hash}.json"


def _restore_revised_cache(file_hash: str, out_png: Path):
    png, meta = _revised_cache_paths(file_hash)
    if not png.exists():
        return None
    try:
        shutil.copy2(png, out_png)
        warnings = []
        if meta.exists():
            obj = json.loads(meta.read_text(encoding="utf-8"))
            warnings = list(obj.get("warnings", []))
        try:
            os.utime(png, None)
        except Exception:
            pass
        return warnings
    except Exception:
        core.logging.exception("读取重制图缓存失败: %s", png)
        return None


def _save_revised_cache(file_hash: str, out_png: Path, warnings) -> None:
    png, meta = _revised_cache_paths(file_hash)
    try:
        if not png.exists():
            tmp = png.with_suffix(".tmp.png")
            shutil.copy2(out_png, tmp)
            try:
                tmp.replace(png)
            except FileExistsError:
                tmp.unlink(missing_ok=True)
        meta.write_text(json.dumps({"warnings": warnings or []}, ensure_ascii=False), encoding="utf-8")
    except Exception:
        core.logging.exception("写入重制图缓存失败: %s", png)


def _cleanup_cache(max_age_days: int = 45) -> None:
    cutoff = time.time() - max_age_days * 86400
    try:
        for p in WORD_CACHE.iterdir():
            if p.is_dir() and p.stat().st_mtime < cutoff:
                shutil.rmtree(p, ignore_errors=True)
        for p in REVISED_CACHE.iterdir():
            if p.is_file() and p.stat().st_mtime < cutoff:
                try:
                    p.unlink()
                except Exception:
                    pass
    except Exception:
        core.logging.exception("清理批处理缓存失败")


def _public_pairs(pairs):
    return [
        {k: p0.get(k) for k in [
            "id", "author", "title", "figure_no", "caption", "original_method",
            "original_confidence", "revised_source", "revised_warnings", "status", "comparison"
        ]}
        for p0 in pairs
    ]


def _fast_prepare_compare(pid: str):
    core._cleanup_old_jobs()
    _cleanup_cache()

    word_files = [f for f in request.files.getlist("word_files") if Path(f.filename).suffix.lower() in core.WORD_EXTS]
    revised_files = [f for f in request.files.getlist("revised_files") if Path(f.filename).suffix.lower() in core.REVISED_EXTS]
    if not word_files:
        return jsonify({"error": "没有选择 .doc / .docx 原稿。"}), 400
    if not revised_files:
        return jsonify({"error": "没有选择 PDF / JPG 重制图。"}), 400

    revised_relpaths = core._relative_upload_paths("revised_relpaths", len(revised_files))
    job_id = os.urandom(10).hex()
    job = core.JOBS_DIR / job_id
    incoming_words = job / "incoming_words"
    incoming_revised = job / "incoming_revised"
    original_root = job / "originals"
    revised_root = job / "revised_rendered"
    work = job / "work"
    for d in (incoming_words, incoming_revised, original_root, revised_root, work):
        d.mkdir(parents=True, exist_ok=True)

    word_total = len(word_files)
    revised_total = len(revised_files)
    started = time.time()
    _set_state(
        pid, status="running", phase="word", phase_label="提取 Word 原稿",
        detail=f"批量模式启动：{word_total} 篇 Word + {revised_total} 张重制图",
        percent=4, word_total=word_total, word_done=0, word_failed=0, word_cached=0,
        revised_total=revised_total, revised_done=0, revised_failed=0, revised_cached=0,
        pair_state="waiting", current_file="", started_ts=started,
    )

    articles = []
    failures = []
    word_cached = 0
    word_phase_started = time.time()
    session = _WordBatchSession()
    _WORD_CTX.session = session
    try:
        for idx, f in enumerate(word_files):
            name = Path(f.filename.replace("\\", "/")).name
            src = incoming_words / f"{idx:03d}_{core.sanitize_filename(name, 140)}"
            f.save(src)
            file_hash = _sha256_file(src)
            cache_dir = WORD_CACHE / file_hash

            _set_state(
                pid, phase="word", phase_label="提取 Word 原稿",
                detail=f"正在处理第 {idx + 1}/{word_total} 篇：{name}",
                current_file=name,
                percent=5 + int(52 * idx / max(word_total, 1)),
            )
            try:
                article = _article_from_cache(cache_dir, original_root, name)
                if article is not None:
                    word_cached += 1
                else:
                    article = core.process_one_word(src, original_root, work)
                    article.source_file = name
                    _save_article_cache(cache_dir, article, original_root)
                articles.append(article)
            except Exception as exc:
                core.logging.exception("高吞吐准备：处理 Word 失败 %s", name)
                failures.append({"file": name, "error": str(exc)})

            done = idx + 1
            eta = _phase_eta(word_phase_started, done, word_total)
            _set_state(
                pid, word_done=done, word_failed=len(failures), word_cached=word_cached,
                percent=5 + int(52 * done / max(word_total, 1)),
                detail=(
                    f"Word 已处理 {done}/{word_total}，缓存命中 {word_cached}，失败 {len(failures)}"
                    + (f" · {_fmt_eta(eta)}" if eta else "")
                ),
            )
    finally:
        try:
            session.close()
        finally:
            _WORD_CTX.session = None

    if not articles:
        shutil.rmtree(job, ignore_errors=True)
        return jsonify({"error": "所有 Word 原稿处理失败。", "failures": failures}), 500

    tasks = []
    revised_by_idx: dict[int, dict] = {}
    revised_cached = 0
    _set_state(
        pid, phase="revised", phase_label="准备重制图",
        detail=f"正在检查 {revised_total} 个重制文件的缓存……",
        percent=58, current_file="",
    )
    for idx, f in enumerate(revised_files):
        orig_name = Path(f.filename.replace("\\", "/")).name
        ext = Path(orig_name).suffix.lower()
        src = incoming_revised / f"{idx:04d}_{core.sanitize_filename(Path(orig_name).stem, 100)}{ext}"
        f.save(src)
        out = revised_root / f"r{idx:04d}.png"
        rel = revised_relpaths[idx] or orig_name
        file_hash = _sha256_file(src)
        cached_warnings = _restore_revised_cache(file_hash, out)
        if cached_warnings is not None:
            revised_cached += 1
            revised_by_idx[idx] = {
                "source_name": orig_name,
                "relative_path": rel,
                "rendered_path": str(out.resolve()),
                "warnings": cached_warnings,
            }
        else:
            tasks.append({
                "idx": idx, "src": src, "out": out, "source_name": orig_name,
                "relative_path": rel, "hash": file_hash,
            })

    done_count = revised_cached
    _set_state(
        pid, revised_done=done_count, revised_cached=revised_cached,
        detail=f"重制图缓存命中 {revised_cached}/{revised_total}；待并行渲染 {len(tasks)} 张",
        percent=58 + int(5 * done_count / max(revised_total, 1)),
    )

    render_workers = max(1, int(os.environ.get(
        "JFE_RENDER_WORKERS",
        str(min(6, max(2, (os.cpu_count() or 4) // 2)))
    )))
    if tasks:
        render_phase_started = time.time()
        ctx = multiprocessing.get_context("spawn")
        with ProcessPoolExecutor(max_workers=render_workers, mp_context=ctx) as pool:
            future_map = {
                pool.submit(render_revised_worker, str(t["src"]), str(t["out"])): t
                for t in tasks
            }
            for future in as_completed(future_map):
                t = future_map[future]
                idx = t["idx"]
                try:
                    warns = future.result()
                    revised_by_idx[idx] = {
                        "source_name": t["source_name"],
                        "relative_path": t["relative_path"],
                        "rendered_path": str(t["out"].resolve()),
                        "warnings": warns,
                    }
                    _save_revised_cache(t["hash"], t["out"], warns)
                except Exception as exc:
                    core.logging.exception("并行渲染重制图失败 %s", t["source_name"])
                    revised_by_idx[idx] = {
                        "source_name": t["source_name"],
                        "relative_path": t["relative_path"],
                        "rendered_path": "",
                        "warnings": [str(exc)],
                        "render_failed": True,
                        "reason": "重制图无法渲染：" + str(exc),
                    }

                done_count += 1
                failed_count = sum(1 for x in revised_by_idx.values() if not x.get("rendered_path"))
                rendered_misses_done = max(0, done_count - revised_cached)
                eta = _phase_eta(render_phase_started, rendered_misses_done, len(tasks))
                _set_state(
                    pid, phase="revised", phase_label=f"并行渲染重制图（{render_workers} 进程）",
                    revised_done=done_count, revised_failed=failed_count, revised_cached=revised_cached,
                    current_file=t["source_name"],
                    percent=63 + int(29 * done_count / max(revised_total, 1)),
                    detail=(
                        f"重制图已完成 {done_count}/{revised_total}，缓存 {revised_cached}，失败 {failed_count}"
                        + (f" · {_fmt_eta(eta)}" if eta else "")
                    ),
                )

    revised_items = [revised_by_idx[i] for i in range(revised_total)]
    render_ok = [x for x in revised_items if x.get("rendered_path")]
    render_bad = [x for x in revised_items if not x.get("rendered_path")]

    _set_state(
        pid, phase="pair", phase_label="自动配对", pair_state="running",
        percent=96, detail="正在建立作者 + 图号索引并自动配对……", current_file="",
    )
    pairs, unmatched_originals, unmatched_revised = core._pair_revised(articles, original_root, render_ok)
    unmatched_revised.extend(render_bad)

    job_data = {
        "job_id": job_id,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "articles": [asdict(a) for a in articles],
        "word_failures": failures,
        "pairs": pairs,
        "unmatched_originals": unmatched_originals,
        "unmatched_revised": unmatched_revised,
        "performance": {
            "word_cache_hits": word_cached,
            "revised_cache_hits": revised_cached,
            "render_workers": render_workers,
            "elapsed_seconds": int(time.time() - started),
        },
    }
    (job / "job.json").write_text(json.dumps(job_data, ensure_ascii=False, indent=2), encoding="utf-8")

    _set_state(
        pid, status="done", phase="done", phase_label="完成", pair_state="done",
        percent=100, current_file="",
        detail=(
            f"完成：{len(articles)} 篇 / {sum(len(a.figures) for a in articles)} 张原稿图 / "
            f"{revised_total} 张重制图；缓存命中 Word {word_cached}、重制图 {revised_cached}"
        ),
    )

    return jsonify({
        "job_id": job_id,
        "article_count": len(articles),
        "original_figure_count": sum(len(a.figures) for a in articles),
        "revised_count": len(revised_files),
        "matched_count": len(pairs),
        "unmatched_original_count": len(unmatched_originals),
        "unmatched_revised_count": len(unmatched_revised),
        "pairs": _public_pairs(pairs),
        "unmatched_originals": [
            {k: x.get(k) for k in ["author", "title", "figure_no", "caption"]}
            for x in unmatched_originals
        ],
        "unmatched_revised": [
            {k: x.get(k) for k in ["relative_path", "source_name", "reason", "warnings"]}
            for x in unmatched_revised
        ],
        "word_failures": failures,
        "performance": job_data["performance"],
    })


def tracked_prepare_compare():
    _clean_progress()
    pid = str(request.form.get("progress_id", "")).strip().lower()
    if not re.fullmatch(r"[0-9a-f]{12,40}", pid):
        pid = os.urandom(10).hex()

    word_total = len([f for f in request.files.getlist("word_files") if Path(f.filename).suffix.lower() in core.WORD_EXTS])
    revised_total = len([f for f in request.files.getlist("revised_files") if Path(f.filename).suffix.lower() in core.REVISED_EXTS])
    started = time.time()
    _set_state(
        pid, status="running", phase="queue", phase_label="批量任务已接收",
        detail=f"准备处理 {word_total} 篇 Word + {revised_total} 张重制图……",
        percent=3, word_total=word_total, word_done=0, word_failed=0, word_cached=0,
        revised_total=revised_total, revised_done=0, revised_failed=0, revised_cached=0,
        pair_state="waiting", current_file="", started_ts=started,
    )

    with _BATCH_LOCK:
        try:
            rv = _fast_prepare_compare(pid)
            response = flask_app.make_response(rv)
            if response.status_code >= 400:
                _set_state(
                    pid, status="error", phase="error", phase_label="处理失败",
                    percent=100, detail=f"处理失败（HTTP {response.status_code}），请查看页面错误信息。",
                )
            response.headers["X-JFE-Progress-ID"] = pid
            return response
        except Exception as exc:
            _set_state(
                pid, status="error", phase="error", phase_label="处理失败",
                percent=100, detail=str(exc)[:500],
            )
            raise


flask_app.view_functions["prepare_compare"] = tracked_prepare_compare


@flask_app.get("/api/prepare-progress/<pid>")
def prepare_progress(pid: str):
    if not re.fullmatch(r"[0-9a-f]{12,40}", pid or ""):
        return jsonify({"error": "无效进度 ID"}), 400
    state = _get_state(pid)
    if not state:
        return jsonify({"status": "waiting", "percent": 0, "detail": "等待服务器接收文件……"}), 200
    if state.get("started_ts"):
        state["elapsed_seconds"] = max(0, int(time.time() - float(state["started_ts"])))
    state.pop("updated_ts", None)
    state.pop("started_ts", None)
    return jsonify(state)


@flask_app.after_request
def inject_progress_ui(response):
    if request.path == "/" and response.mimetype == "text/html":
        try:
            text = response.get_data(as_text=True)
            text = text.replace("v1.2", f"v{APP_VERSION}").replace("v1.3", f"v{APP_VERSION}")
            script = f'<script src="/static/progress.js?v={APP_VERSION}"></script>'
            if script not in text:
                text = text.replace("</body>", script + "\n</body>")
            old = "这一步不调用 AI。继续使用 v0.8 已验证的“Word 对象优先”提取逻辑，再将重制 PDF/JPG 渲染成图并按作者+图号配对。"
            new = (
                "这一步不调用 AI。高吞吐模式会复用同一个 Word 进程、并行渲染重制图，并缓存已处理文件；"
                "面向几十篇文章、数百张重制图的整期批量处理。"
            )
            text = text.replace(old, new)
            response.set_data(text)
            response.headers["Content-Length"] = str(len(response.get_data()))
            response.headers["Cache-Control"] = "no-store"
        except Exception:
            core.logging.exception("注入进度 UI 失败")
    return response


if __name__ == "__main__":
    port = int(os.environ.get("JFE_PORT", "8765"))
    flask_app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
