from __future__ import annotations

import os
import re
import sys
import time
import threading
from pathlib import Path

from flask import jsonify, request

import app as core

APP_VERSION = "1.3"
core.APP_VERSION = APP_VERSION
flask_app = core.app

_PROGRESS: dict[str, dict] = {}
_PROGRESS_LOCK = threading.Lock()
_PROGRESS_CTX = threading.local()
_PROGRESS_MAX_AGE = 6 * 3600


def _clean_progress() -> None:
    cutoff = time.time() - _PROGRESS_MAX_AGE
    with _PROGRESS_LOCK:
        stale = [k for k, v in _PROGRESS.items() if float(v.get("updated_ts", 0)) < cutoff]
        for k in stale:
            _PROGRESS.pop(k, None)


def _progress_id() -> str | None:
    return getattr(_PROGRESS_CTX, "progress_id", None)


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


def _update_current(**values) -> None:
    pid = _progress_id()
    if pid:
        _set_state(pid, **values)


def _percent_for_word(done: int, total: int) -> int:
    return 5 + int(55 * min(done, max(total, 1)) / max(total, 1))


def _percent_for_revised(done: int, total: int) -> int:
    return 60 + int(31 * min(done, max(total, 1)) / max(total, 1))


def _display_name(path) -> str:
    name = Path(str(path)).name
    return re.sub(r"^\d{3,4}_", "", name)


# Wrap the proven v0.8/v1.2 processing functions instead of changing extraction logic.
_original_process_one_word = core.process_one_word
_original_render_revised_file = core.render_revised_file
_original_pair_revised = core._pair_revised
_original_prepare_compare = flask_app.view_functions["prepare_compare"]


def _tracked_process_one_word(src, *args, **kwargs):
    pid = _progress_id()
    if not pid:
        return _original_process_one_word(src, *args, **kwargs)
    state = _get_state(pid) or {}
    total = int(state.get("word_total", 0) or 0)
    done = int(state.get("word_done", 0) or 0)
    _update_current(
        phase="word",
        phase_label="提取 Word 原稿",
        detail=f"正在处理第 {done + 1}/{total} 篇：{_display_name(src)}",
        current_file=_display_name(src),
        percent=max(5, _percent_for_word(done, total)),
    )
    failed = False
    try:
        return _original_process_one_word(src, *args, **kwargs)
    except Exception:
        failed = True
        raise
    finally:
        state = _get_state(pid) or {}
        new_done = int(state.get("word_done", 0) or 0) + 1
        new_failed = int(state.get("word_failed", 0) or 0) + (1 if failed else 0)
        _update_current(
            word_done=new_done,
            word_failed=new_failed,
            percent=_percent_for_word(new_done, total),
            detail=f"Word 原稿已处理 {new_done}/{total}" + (f"（失败 {new_failed}）" if new_failed else ""),
        )


def _tracked_render_revised_file(src, *args, **kwargs):
    pid = _progress_id()
    if not pid:
        return _original_render_revised_file(src, *args, **kwargs)
    state = _get_state(pid) or {}
    total = int(state.get("revised_total", 0) or 0)
    done = int(state.get("revised_done", 0) or 0)
    _update_current(
        phase="revised",
        phase_label="渲染重制图",
        detail=f"正在渲染第 {done + 1}/{total} 张：{_display_name(src)}",
        current_file=_display_name(src),
        percent=max(60, _percent_for_revised(done, total)),
    )
    failed = False
    try:
        return _original_render_revised_file(src, *args, **kwargs)
    except Exception:
        failed = True
        raise
    finally:
        state = _get_state(pid) or {}
        new_done = int(state.get("revised_done", 0) or 0) + 1
        new_failed = int(state.get("revised_failed", 0) or 0) + (1 if failed else 0)
        _update_current(
            revised_done=new_done,
            revised_failed=new_failed,
            percent=_percent_for_revised(new_done, total),
            detail=f"重制图已渲染 {new_done}/{total}" + (f"（失败 {new_failed}）" if new_failed else ""),
        )


def _tracked_pair_revised(*args, **kwargs):
    pid = _progress_id()
    if pid:
        _update_current(
            phase="pair",
            phase_label="自动配对",
            pair_state="running",
            percent=94,
            detail="正在按正文第一作者 + 图号自动配对……",
            current_file="",
        )
    result = _original_pair_revised(*args, **kwargs)
    if pid:
        pairs, unmatched_originals, unmatched_revised = result
        _update_current(
            pair_state="done",
            percent=98,
            detail=f"配对完成：{len(pairs)} 对成功，{len(unmatched_originals) + len(unmatched_revised)} 项未匹配",
        )
    return result


core.process_one_word = _tracked_process_one_word
core.render_revised_file = _tracked_render_revised_file
core._pair_revised = _tracked_pair_revised


def tracked_prepare_compare():
    _clean_progress()
    pid = str(request.form.get("progress_id", "")).strip().lower()
    if not re.fullmatch(r"[0-9a-f]{12,40}", pid):
        pid = os.urandom(10).hex()

    word_total = len([f for f in request.files.getlist("word_files") if Path(f.filename).suffix.lower() in core.WORD_EXTS])
    revised_total = len([f for f in request.files.getlist("revised_files") if Path(f.filename).suffix.lower() in core.REVISED_EXTS])
    started = time.time()
    _set_state(
        pid,
        status="running",
        phase="word",
        phase_label="准备处理",
        detail="文件已上传，正在初始化 Word 提取……",
        percent=3,
        word_total=word_total,
        word_done=0,
        word_failed=0,
        revised_total=revised_total,
        revised_done=0,
        revised_failed=0,
        pair_state="waiting",
        current_file="",
        started_ts=started,
        elapsed_seconds=0,
    )

    _PROGRESS_CTX.progress_id = pid
    try:
        rv = _original_prepare_compare()
        response = flask_app.make_response(rv)
        elapsed = max(0, int(time.time() - started))
        if response.status_code >= 400:
            _set_state(
                pid,
                status="error",
                phase="error",
                phase_label="处理失败",
                percent=100,
                elapsed_seconds=elapsed,
                detail=f"处理失败（HTTP {response.status_code}），请查看页面错误信息。",
            )
        else:
            _set_state(
                pid,
                status="done",
                phase="done",
                phase_label="完成",
                percent=100,
                elapsed_seconds=elapsed,
                detail="全部提取、渲染与配对已完成。",
                current_file="",
            )
        response.headers["X-JFE-Progress-ID"] = pid
        return response
    except Exception as exc:
        _set_state(
            pid,
            status="error",
            phase="error",
            phase_label="处理失败",
            percent=100,
            elapsed_seconds=max(0, int(time.time() - started)),
            detail=str(exc)[:500],
        )
        raise
    finally:
        _PROGRESS_CTX.progress_id = None


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
    # Keep the existing proven front end untouched; inject the progress enhancement at runtime.
    if request.path == "/" and response.mimetype == "text/html":
        try:
            text = response.get_data(as_text=True)
            text = text.replace("v1.2", f"v{APP_VERSION}")
            script = f'<script src="/static/progress.js?v={APP_VERSION}"></script>'
            if script not in text:
                text = text.replace("</body>", script + "\n</body>")
            response.set_data(text)
            response.headers["Content-Length"] = str(len(response.get_data()))
            response.headers["Cache-Control"] = "no-store"
        except Exception:
            core.logging.exception("注入进度 UI 失败")
    return response


if __name__ == "__main__":
    port = int(os.environ.get("JFE_PORT", "8765"))
    flask_app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
