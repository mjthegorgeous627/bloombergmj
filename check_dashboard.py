"""workbench 화면 문법 검사 - dashboard.html 수정 후, 서버 재시작 전에 실행.
각 JS 조각과 합친 전체 <script>를 node --check 로 검사한다. 실패하면 종료코드 1."""

import re
import subprocess
import sys
import tempfile
from pathlib import Path

import dashboard_assets

sys.stdout.reconfigure(encoding="utf-8")


def _check_js(label, code):
    with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as f:
        f.write(code)
        path = f.name
    r = subprocess.run(["node", "--check", path], capture_output=True, text=True, encoding="utf-8")
    Path(path).unlink(missing_ok=True)
    if r.returncode != 0:
        print(f"[실패] {label}\n{r.stderr.strip()}")
        return False
    print(f"[통과] {label}")
    return True


def main():
    ok = True
    for p in sorted((dashboard_assets.TEMPLATES_DIR / "dashboard").glob("*.js")):
        ok &= _check_js(p.name, p.read_text(encoding="utf-8"))
    html = dashboard_assets.assemble()
    if dashboard_assets._INCLUDE_RE.search(html):
        print("[실패] 끼워 넣지 못한 #include 줄이 남아 있음")
        ok = False
    scripts = re.findall(r"<script>(.*?)</script>", html, re.S)
    ok &= _check_js("합친 전체 script", "\n".join(scripts).replace("__BOARD_DATA__", "[]")
                    .replace("__CUSTOM_TYPES__", "[]").replace("__NONWORKING_DAYS__", "{}")
                    .replace("__DELETED_ITEM_RETENTION_DAYS__", "7"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
