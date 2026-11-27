# -*- coding: utf-8 -*-
"""Check Windows, Python, Excel COM, pywin32, and openpyxl."""
import argparse
import platform
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import say, write_text, wrote
from excel_retry import BUSY_MESSAGE, ExcelShutdownBusy, busy, cleanup_excel, quit_excel, retry_excel


def check_excel():
    try:
        from win32com.client import DispatchEx
    except Exception as exc:
        return False, "pywin32 import failed: " + type(exc).__name__
    try:
        excel = retry_excel(lambda: DispatchEx("Excel.Application"))
        with cleanup_excel(lambda: quit_excel(lambda: excel.Quit())):
            retry_excel(lambda: setattr(excel, "Visible", False))
            retry_excel(lambda: setattr(excel, "DisplayAlerts", False))
            version = str(retry_excel(lambda: excel.Version))
        return True, version
    except Exception as exc:
        return False, BUSY_MESSAGE if busy(exc) or isinstance(exc, ExcelShutdownBusy) else type(exc).__name__


def main():
    parser = argparse.ArgumentParser(description="Check the model build environment.")
    parser.add_argument("--out", type=Path, help="UTF-8 markdown report")
    args = parser.parse_args()
    rows = []
    ok = True

    windows = platform.system() == "Windows"
    rows.append(("Windows", windows, platform.platform()))
    ok = ok and windows

    py_ok = sys.version_info >= (3, 10)
    rows.append(("Python>=3.10", py_ok, sys.version.split()[0]))
    ok = ok and py_ok

    try:
        import openpyxl
        rows.append(("openpyxl", True, getattr(openpyxl, "__version__", "installed")))
    except Exception:
        rows.append(("openpyxl", False, "python -m pip install openpyxl"))
        ok = False

    try:
        import win32com.client  # noqa: F401
        rows.append(("pywin32", True, "installed"))
        pywin = True
    except Exception:
        rows.append(("pywin32", False, "python -m pip install pywin32"))
        ok = False
        pywin = False

    if windows and pywin:
        excel_ok, detail = check_excel()
        rows.append(("Excel COM", excel_ok, detail if excel_ok or detail == BUSY_MESSAGE else "Install desktop Excel. " + detail))
        ok = ok and excel_ok
    else:
        rows.append(("Excel COM", False, "Skipped until Windows and pywin32 are available."))
        ok = False

    lines = ["# 환경 확인", "", "결과: " + ("통과" if ok else "부족"), ""]
    lines.append("| 항목 | 상태 | 내용 |")
    lines.append("|---|---|---|")
    for name, passed, detail in rows:
        status = "확인 못 함" if detail == BUSY_MESSAGE else ("있음" if passed else "없음")
        lines.append(f"| {name} | {status} | {detail} |")
    lines.append("")
    if any(not passed and detail != BUSY_MESSAGE for _, passed, detail in rows):
        lines.append("## 설치")
        lines.append("")
        if not py_ok:
            lines.append("- Python 3.10 이상을 설치한다.")
        if not any(name == "openpyxl" and passed for name, passed, _detail in rows):
            lines.append("- `python -m pip install openpyxl`")
        if not any(name == "pywin32" and passed for name, passed, _detail in rows):
            lines.append("- `python -m pip install pywin32`")
        if not any(name == "Excel COM" and (passed or detail == BUSY_MESSAGE) for name, passed, detail in rows):
            lines.append("- Microsoft Excel 데스크톱을 설치한다. COM 자동화에는 Excel이 필요하다.")
    if args.out:
        write_text(args.out, "\n".join(lines) + "\n")
        wrote(args.out, "env")
    for name, passed, detail in rows:
        text = str(detail).encode("ascii", "ignore").decode("ascii").strip()
        if detail == BUSY_MESSAGE:
            say(f"{name}: BUSY. Retry when Excel is ready.")
        else:
            say(f"{name}: {'OK' if passed else 'MISSING'} {text}")
    say("RESULT " + ("PASS" if ok else "FAIL"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
