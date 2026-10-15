"""백그라운드 Claude(ticket_check_run.py가 띄움)가 쓰는 티켓 확인 도구 모음 (2026-10-15).

백그라운드 Claude는 이 스크립트 외의 명령은 실행할 수 없게 제한된다(--allowedTools). 보드(workbench.db)는 읽기만 한다.
사용:
  python -X utf8 tickets/ticket_tools.py inbox                 수신함을 열고 화면 캡처 -> 경로 출력
  python -X utf8 tickets/ticket_tools.py open <행y> <이름>      수신함 줄(y좌표) 클릭 -> 메일 -> SDSK 링크 -> 티켓 (읽음 처리됨)
  python -X utf8 tickets/ticket_tools.py mail <행y> <이름>      일반 메일 줄 클릭 -> 메일 내용만 OCR (링크 안 누름)
  python -X utf8 tickets/ticket_tools.py read <SDSK번호>        티켓 본문 전체 캡처(스크롤, 수 분 걸림)
  python -X utf8 tickets/ticket_tools.py board <검색어>         보드에서 SDSK/오더번호/이름/S/N 검색(읽기 전용)
"""
import os
import re
import sqlite3
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
NL = chr(10)
sys.path.insert(0, HERE)
OUT = r"C:\1\티켓관리\inbox\run"
DB = os.path.join(os.path.dirname(HERE), "workbench.db")
ROW0_Y, ROW_PITCH = 222, 26.65     # 수신함 1번째 줄 y, 줄 간격 (2026-10-15 실측)


def cmd_inbox():
    import term_ops as t
    os.makedirs(OUT, exist_ok=True)
    t.goto("MSG", wait=6)
    p = os.path.join(OUT, "inbox.png")
    t.shot(p)
    print(p)
    print(f"행 n의 y좌표 = {ROW0_Y} + round({ROW_PITCH} * (n-1))")


MAIL_REGIONS = [("메일 머리(보낸사람/제목/날짜)", (0, 150, 1920, 350)), ("메일 본문", (0, 350, 1920, 1030))]
TICKET_REGIONS = [("티켓 머리(Summary/Status/Rep)", (0, 175, 1000, 250)),
                  ("티켓 속성(Due Date/Country/Office)", (1400, 290, 1910, 500)),
                  ("티켓 글 목록(Key Details > Details, 화면에 보이는 최신 글들)", (180, 360, 1390, 980))]


def _ocr_regions(img, regions):
    import ocr
    crops = [img.crop(box) for _, box in regions]
    res = ocr.ocr_images(crops)
    return [ocr.block(title, {lg: res[lg][i] for lg in res}) for i, (title, _) in enumerate(regions)]


def cmd_open(row_y, name):
    import term_ops as t
    os.makedirs(OUT, exist_ok=True)
    t.goto("MSG", wait=5)
    mail, ticket = t.open_ticket_from_inbox(int(row_y))
    name = re.sub(r"[^0-9A-Za-z_-]", "", name) or "x"
    mp, tp = os.path.join(OUT, f"mail_{name}.png"), os.path.join(OUT, f"ticket_{name}.png")
    mail.save(mp)
    ticket.save(tp)
    print(f"(이미지 원본: {mp} , {tp} - OCR이 이상할 때만 Read로 확인)")
    print((NL * 2).join(_ocr_regions(mail, MAIL_REGIONS) + _ocr_regions(ticket, TICKET_REGIONS)))


def cmd_mail(row_y, name):
    import term_ops as t
    os.makedirs(OUT, exist_ok=True)
    t.goto("MSG", wait=5)
    mail = t.open_mail_from_inbox(int(row_y))
    name = re.sub(r"[^0-9A-Za-z_-]", "", name) or "x"
    mp = os.path.join(OUT, f"mail_{name}.png")
    mail.save(mp)
    print(f"(이미지 원본: {mp} - OCR이 이상할 때만 Read로 확인)")
    print((NL * 2).join(_ocr_regions(mail, MAIL_REGIONS)))


def cmd_read(num):
    import ocr
    import read_ticket
    from PIL import Image
    folder, frames = read_ticket.read_ticket(f"SDSK {num}")
    print(f"(이미지 원본 폴더: {folder} , {len(frames)} frames)")
    chrono = list(reversed(frames))                      # frame_NN(큰 번호)=오래된 글(위쪽) -> frame_01=최신(아래쪽)
    imgs = [Image.open(f).convert("RGB") for f in chrono]
    head = Image.open(os.path.join(folder, "header.png")).convert("RGB")
    hres = ocr.ocr_images([head.crop((0, 175, 1000, 250)), head.crop((1400, 290, 1910, 500))])
    print(ocr.block("티켓 머리(Summary/Status/Rep)", {lg: hres[lg][0] for lg in hres}))
    print(ocr.block("티켓 속성(Due Date/Country/Office)", {lg: hres[lg][1] for lg in hres}))
    res = ocr.ocr_images(imgs)
    print("### 티켓 글 전체(위=오래된 글 → 아래=최신 글, 겹친 줄 제거됨)")
    for lg, lists in res.items():
        print(f"[{'English OCR' if lg.startswith('en') else '한국어 OCR'}]")
        print(NL.join(ocr.merge_frames(lists)))


def cmd_board(q):
    c = sqlite3.connect(f"file:{DB}?mode=ro", uri=True)
    like = f"%{q}%"
    rows = c.execute(
        """select o.id, o.order_no, o.order_type, o.customer, o.phone, o.memo, o.extra_codes, o.delivery_no,
                  i.item_date, i.board_status, i.description, i.material, i.serial
           from orders o left join order_items i on i.order_id = o.id and i.deleted_at is null
           where o.order_no like ? or o.order_type like ? or o.customer like ? or o.memo like ?
                 or i.serial like ? or i.description like ? or o.extra_codes like ? or o.delivery_no like ?
           order by o.id desc limit 30""", (like,) * 8).fetchall()
    if not rows:
        print("(보드에 일치하는 행 없음)")
    for r in rows:
        print(f"id={r[0]} order_no={r[1]!r} type={r[2]!r} customer={r[3]!r} phone={r[4]!r} date={r[8]} status={r[9]!r}"
              f" item={r[10]!r} M/N={r[11]!r} S/N={r[12]!r}" + NL + f"    Order# 아래 코드: OBD={r[7]!r} / {r[6]!r}" + NL + f"    memo={r[5]!r}")


if __name__ == "__main__":
    a = sys.argv[1:]
    if not a:
        raise SystemExit(__doc__)
    if a[0] == "inbox":
        cmd_inbox()
    elif a[0] == "open" and len(a) >= 3:
        cmd_open(a[1], a[2])
    elif a[0] == "mail" and len(a) >= 3:
        cmd_mail(a[1], a[2])
    elif a[0] == "read" and len(a) >= 2:
        cmd_read(re.sub(r"\D", "", a[1]))
    elif a[0] == "board" and len(a) >= 2:
        cmd_board(" ".join(a[1:]))
    else:
        raise SystemExit(__doc__)
