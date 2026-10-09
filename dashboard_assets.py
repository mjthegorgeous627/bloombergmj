"""workbench 화면(dashboard.html)은 한 파일이 3천 줄이라 작은 문자열 실수 하나로 화면 전체가
백지가 되는 사고가 두 번 있었다. 그래서 CSS/JS를 templates/dashboard/ 아래 조각 파일로 나눴고,
이 모듈이 서버 시작 때 조각을 원래 자리에 끼워 넣어 예전과 똑같은 한 덩어리 HTML을 만든다.
(브라우저가 받는 내용은 분리 전과 완전히 같다.)

dashboard.html 안의 `//#include dashboard/xx.js` 또는 `/*#include dashboard/xx.css*/` 한 줄이
해당 파일 내용으로 바뀐다. JS 조각은 모두 같은 <script> 안에 순서대로 들어가므로 변수/함수는
예전처럼 서로 공유된다 - 조각 순서(번호)를 바꾸면 안 된다.

검사: `python check_dashboard.py` (각 조각과 합친 결과를 node --check 로 문법 검사)."""

import re
from pathlib import Path

TEMPLATES_DIR = Path(__file__).resolve().parent / "templates"
_INCLUDE_RE = re.compile(r"^(?://|/\*)#include (\S+?)(?:\*/)?(\r?)$", re.M)


def assemble(name="dashboard.html", base=TEMPLATES_DIR):
    text = (base / name).read_text(encoding="utf-8")

    def repl(m):
        return (base / m.group(1)).read_text(encoding="utf-8") + m.group(2)

    return _INCLUDE_RE.sub(repl, text)
