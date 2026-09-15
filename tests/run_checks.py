"""Прогон проверки/исправления/примечаний по файлу: python tests/run_checks.py path.docx"""

from __future__ import annotations

import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.stdout.reconfigure(encoding="utf-8")

from doxibot import report, service  # noqa: E402


def main(path: str) -> None:
    data = Path(path).read_bytes()
    t = time.perf_counter()
    analysis = service.check(data)
    print(f"check: {time.perf_counter() - t:.2f}s")
    print(re.sub(r"<[^>]+>", "", report.full_text_report(analysis, Path(path).name)))

    t = time.perf_counter()
    result = service.fix(data)
    print(f"fix: {time.perf_counter() - t:.2f}s applied={result.applied} failed={result.failed}")
    out = Path(path).with_name(Path(path).stem + "_fixed.docx")
    out.write_bytes(result.data)
    print(re.sub(r"<[^>]+>", "", report.fix_summary(result.applied, result.failed, result.after)))
    print(report.full_text_report(result.after, out.name))

    annotated, _ = service.annotate(data)
    Path(path).with_name(Path(path).stem + "_notes.docx").write_bytes(annotated)
    print("annotated ok")


if __name__ == "__main__":
    main(sys.argv[1])
