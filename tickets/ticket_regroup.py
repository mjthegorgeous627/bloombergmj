"""카드의 '다시 정리': 이미 저장된 티켓 내용 + 사용자의 지시 한 줄 -> Claude가 TO DO/배송장 제안만 다시 만든다.
터미널은 쓰지 않는다(저장된 글만 사용, 배송장 검색 `board`만 허용). 보드/DB는 수정하지 않는다."""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
sys.path.insert(0, BASE)
import ticket_store  # noqa: E402

CLAUDE_EXE = os.path.join(os.environ.get("APPDATA", ""), "npm", "node_modules", "@anthropic-ai", "claude-code", "bin", "claude.exe")
ALLOWED = ["Bash(python -X utf8 tickets/ticket_tools.py board:*)"]
LOG = os.path.join(HERE, "runs", "corrections.jsonl")
NL = chr(10)

RULES = """너는 SAP/물류 담당자(MINJUNG SUH, 서울 창고팀)의 비서다. 아래 티켓 내용과 사용자의 지시를 보고 **배송장 반영 제안만** 다시 만든다. 터미널은 쓰지 않는다.
- 배송장(workbench)은 내가 해야 할 일의 날짜별 개인 목록이다. 나에게 요구사항이 있으면 반영한다.
- board.action: `append_memo`(row_id + memo_append) / `set_date`(row_id + date YYYY-MM-DD) / `new_row` / `none`.
- new_row이면 board는 {"action":"new_row","new_row":{...필드...},"row_label":"","question":""} 형태로 필드를 반드시 `new_row` 객체 안에 넣는다.
- `new_row` 필드: type(배송/회수/업무), order_no("SDSK 숫자"), item, mn, sn, customer, phone, address, memo, date(YYYY-MM-DD, 없으면 빈칸).
- **S/N이 있는 요청(배송·회수·픽업·수출·수입 등 이동/처리)은 사용자가 말하지 않아도 스스로** 정리한다: `sn`은 S/N 한 줄에 하나씩(줄바꿈), `mn`은 각 S/N과 **같은 줄 순서**로 매칭한다(티켓에 MN/모델이 있으면 그 줄에 맞춰 넣고, 모르면 빈 줄로 둬 줄 수를 맞춘다. 같은 MN이면 줄마다 반복). `type`은 요청 동작에 맞춰 배송/회수/수출/수입 등으로(픽업·반납함 수거=회수), `item`은 품목·수량(예: HDD 8개), 받는 곳·주소는 티켓에서. 날짜가 없으면 빈칸.
- 사용자 지시가 비어 있어도 위 규칙은 적용한다.
- 기존 행에 붙일지 확인하려면 `python -X utf8 tickets/ticket_tools.py board <검색어>`를 쓴다.
- **사용자의 지시가 기본 규칙과 다르면 지시를 따른다.**
- 출력은 JSON 한 개만(설명 없이): {"todo":"내가 할 일 한 줄 이상(줄바꿈 가능)","due":"YYYY-MM-DD 또는 빈칸","board":{"action":..., ...,"row_label":"","question":""}}
"""


def regroup(key, instruction):
    t = next((x for x in ticket_store.pending() + ticket_store.watching() if x["key"] == key), None)
    if not t:
        raise ValueError("저장소에서 티켓을 찾지 못했습니다.")
    ins = str(instruction or "").strip()[:1000]
    if not ins:
        ins = "(지시 없음 - 기본 규칙대로 S/N·MN을 줄 단위로 다시 정리)"
    try:
        with open(os.path.join(HERE, "PLAYBOOK.md"), "r", encoding="utf-8") as fp_:
            rules_extra = NL + fp_.read() + NL
    except Exception:
        rules_extra = ""
    posts = [{k: p.get(k) for k in ("n", "who", "when", "text")} for p in (t.get("posts") or [])]
    prompt = (RULES + rules_extra + NL + "## 티켓" + NL + json.dumps({"key": key, "title": t.get("title"), "srep": t.get("srep"), "posts": posts,
              "현재 todo": t.get("todo"), "현재 제안": t.get("board")}, ensure_ascii=False, indent=1) + NL + NL
              + "## 사용자의 지시" + NL + ins + NL)
    cmd = [CLAUDE_EXE, "-p", "--permission-mode", "dontAsk", "--output-format", "text", "--allowedTools"] + ALLOWED
    proc = subprocess.run(cmd, input=prompt, capture_output=True, text=True, encoding="utf-8", cwd=BASE, timeout=300,
                          creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    out = (proc.stdout or "").strip()
    if proc.returncode != 0 or not out:
        raise RuntimeError("Claude 실행 실패: " + (proc.stderr or out)[-300:])
    data, _ = json.JSONDecoder().raw_decode(out[out.find("{"):])
    bd = data.get("board")
    if isinstance(bd, dict) and bd.get("action") == "new_row" and not isinstance(bd.get("new_row"), dict):
        fields = ("type", "order_no", "item", "mn", "sn", "customer", "phone", "address", "memo", "date")
        bd["new_row"] = {k: bd.pop(k) for k in fields if k in bd}
    try:   # 사용자가 고친 내역을 쌓아 둔다(반복되면 PLAYBOOK 규칙 후보로 제안하기 위해)
        os.makedirs(os.path.dirname(LOG), exist_ok=True)
        with open(LOG, "a", encoding="utf-8") as lf:
            lf.write(json.dumps({"at": __import__("datetime").datetime.now().isoformat(timespec="seconds"), "key": key, "title": t.get("title"),
                                 "instruction": ins, "before": t.get("board"), "after": data.get("board")}, ensure_ascii=False) + NL)
    except Exception:
        pass
    return ticket_store.set_proposal(key, data.get("todo"), data.get("board"), data.get("due"))
