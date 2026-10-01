"""월초 작업 모음 - workbench "월초 작업" 버튼이 호출한다.

첫 번째 작업: 퀵 금액 정리 (별도 저장소 quick_settlement, 코드만 git 관리).
이 모듈은 그 스크립트를 숨김 창으로 실행하고 마지막 줄 RESULT_JSON 을 파싱할 뿐이다.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
QUICK_SETTLEMENT_DIR = BASE_DIR.parent / "quick_settlement"


def run_quick_settlement(month: str, mode: str):
    """month='YYYY-MM', mode='draft'|'apply'. 항상 dict 반환(예외 던지지 않음)."""
    if not re.fullmatch(r"\d{4}-\d{2}", month or ""):
        return {"ok": False, "error": f"월 형식이 올바르지 않습니다: {month!r} (YYYY-MM)"}
    if mode not in ("draft", "apply"):
        return {"ok": False, "error": f"mode 는 draft/apply 만 가능: {mode!r}"}
    script = QUICK_SETTLEMENT_DIR / "quick_settlement.py"
    if not script.exists():
        return {"ok": False, "error": f"스크립트가 없습니다: {script}"}
    try:
        proc = subprocess.run(
            [sys.executable, str(script), "--month", month, "--mode", mode],
            cwd=str(QUICK_SETTLEMENT_DIR), capture_output=True, timeout=600,
            creationflags=subprocess.CREATE_NO_WINDOW, env={**__import__("os").environ, "PYTHONIOENCODING": "utf-8"},
        )
    except subprocess.TimeoutExpired:
        return {"ok": False, "error": "퀵 금액 정리가 10분 안에 끝나지 않았습니다 (Excel 창/프로세스 확인 필요)"}
    out = proc.stdout.decode("utf-8", "replace")
    for line in reversed(out.splitlines()):
        if line.startswith("RESULT_JSON:"):
            return json.loads(line[len("RESULT_JSON:"):])
    return {"ok": False, "error": "결과를 읽지 못했습니다", "stderr": proc.stderr.decode("utf-8", "replace")[-800:]}
