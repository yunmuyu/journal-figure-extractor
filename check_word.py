from __future__ import annotations

import sys


def main() -> int:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        print(f"WORD_COM_OK version={word.Version}")
        return 0
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"WORD_COM_FAILED: {type(exc).__name__}: {exc}", file=sys.stderr)
        raise
