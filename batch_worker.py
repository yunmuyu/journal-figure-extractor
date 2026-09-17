from __future__ import annotations

from pathlib import Path


def render_revised_worker(src: str, out_png: str):
    """Process-pool worker for revised PDF/JPG/PNG rendering.

    Import app inside the child process so PyMuPDF/Pillow state is isolated per process.
    This preserves the proven render_revised_file behavior while allowing safe parallelism
    across independent files on Windows.
    """
    import app as core

    warnings = core.render_revised_file(Path(src), Path(out_png))
    return warnings
