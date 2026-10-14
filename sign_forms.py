"""배송/회수 서명 용지(POD/POC) 자동 작성+인쇄 (workbench ver.2, 2026-10-14).

서식 파일: C:\\1\\ZOR 배송 확인 싸인 용지 DN&PN-SEOUL .xlsx  (시트 'POD' = Proof of Delivery, 'POC' = Proof of Collection)
- 원본 파일은 건드리지 않는다: 임시 폴더로 복사한 뒤 숨은 Excel(별도 인스턴스)로 열어 값을 채우고 인쇄한다.
- 기본 프린터(A4)로 인쇄. mode='pdf' 이면 인쇄 대신 PDF로 저장(시안/검증용).
- 한 장에 다 안 들어가면 "(1/2)" "(2/2)" 표시를 제목에 붙여 여러 장을 인쇄한다.
  POD 표는 5줄, POC 표는 7줄(왼쪽 고객용/오른쪽 우리용 같은 내용).

CLI(서버가 호출): python sign_forms.py <job.json>
job = {"mode": "print"|"pdf", "kind": "pod"|"poc", "order": {...}, "pdfPath": "..."}
order = {"prefix": "ZOR", "number": "7866546", "date": "2026-10-14", "company": "...", "address": "...", "attn": "...",
         "items": [{"material": "10045196", "description": "KEYBOARD ...", "qty": 1}]}
"""
from __future__ import annotations

import json
import logging
import os
import shutil
import sys
import tempfile

TEMPLATE_PATH = r"C:\1\ZOR 배송 확인 싸인 용지 DN&PN-SEOUL .xlsx"
logger = logging.getLogger("sign_forms")

# 시트별 셀 위치(2026-10-14 서식 기준). 서식 칸 위치가 바뀌면 여기만 고친다.
LAYOUT = {
    "pod": {
        "sheet": "POD", "rows": 5, "first_row": 22, "title": "D6", "title_text": "Proof of Delivery",
        "date": ["D7"], "company": ["D10"], "address": ["D11"], "attn": ["D16"], "order": ["D19"],
        "cols": [("D", "material"), ("E", "description"), ("H", "qty"), ("I", "unit")],
    },
    "poc": {
        "sheet": "POC", "rows": 7, "first_row": 21, "title": "D6", "title_text": "Proof of Collection",
        "date": ["D7"], "company": ["D10", "K10"], "address": ["D11", "K11"], "attn": ["D15"], "order": ["D18"],
        "cols": [("D", "material"), ("E", "description"), ("H", "qty"), ("I", "unit"),
                 ("K", "material"), ("L", "description"), ("O", "qty"), ("P", "unit")],
    },
}


def _clean_address(address):
    lines = [ln.strip() for ln in str(address or "").splitlines() if ln.strip()]
    out = []
    for ln in lines:
        low = ln.lower()
        if "cust#" in low:
            ln = ln[: low.index("cust#")].rstrip(" (").strip()
        if ln:
            out.append(ln)
    return out


def _chunks(items, n):
    return [items[i:i + n] for i in range(0, len(items), n)] or [[]]


def _cell_value(col_key, item):
    if col_key == "material":
        m = str(item.get("material") or "").strip()
        return int(m) if m.isdigit() else m
    if col_key == "description":
        return str(item.get("description") or "").strip()
    if col_key == "qty":
        return item.get("qty") or 1
    return "EA"


def fill_sheet(ws, kind, order, page_items, page_no, pages):
    lay = LAYOUT[kind]
    lines = _clean_address(order.get("address"))
    given = str(order.get("company") or "").strip()
    company = given or (lines[0] if lines else "")
    addr_lines = lines if given else lines[1:]
    pre = str(order.get("prefix") or "").strip()
    num = str(order.get("number") or "").strip()
    for c in lay["date"]:
        ws.Range(c).Value = "Date :  " + str(order.get("date") or "")
    for c in lay["company"]:
        ws.Range(c).Value = company
    for c in lay["address"]:
        ws.Range(c).Value = "\n".join(addr_lines)
    for c in lay["attn"]:
        ws.Range(c).Value = "Attn:  " + str(order.get("attn") or "").strip()
    for c in lay["order"]:
        ws.Range(c).Value = f"{pre}# {num}".strip()
    ws.Range(lay["title"]).Value = lay["title_text"] + (f"  ({page_no}/{pages})" if pages > 1 else "")
    for i in range(lay["rows"]):
        r = lay["first_row"] + i
        item = page_items[i] if i < len(page_items) else None
        for col, key in lay["cols"]:
            ws.Range(f"{col}{r}").Value = _cell_value(key, item) if item else ""


def run_job(job):
    import win32com.client

    kind = job["kind"]
    lay = LAYOUT[kind]
    order = job["order"]
    # 같은 자재+설명은 수량 합산
    merged = {}
    for it in order.get("items") or []:
        k = (str(it.get("material") or "").strip(), str(it.get("description") or "").strip())
        merged[k] = merged.get(k, 0) + int(it.get("qty") or 1)
    items = [{"material": k[0], "description": k[1], "qty": q} for k, q in merged.items()]
    pages = _chunks(items, lay["rows"])

    tmp_dir = tempfile.mkdtemp(prefix="signform_")
    tmp_xlsx = os.path.join(tmp_dir, "form.xlsx")
    shutil.copyfile(TEMPLATE_PATH, tmp_xlsx)
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    wb = None
    try:
        wb = xl.Workbooks.Open(tmp_xlsx, 0, False)
        ws = wb.Worksheets(lay["sheet"])
        pdf_paths = []
        for idx, page_items in enumerate(pages, start=1):
            fill_sheet(ws, kind, order, page_items, idx, len(pages))
            if job.get("mode") == "pdf":
                path = job["pdfPath"]
                if len(pages) > 1:
                    base, ext = os.path.splitext(path)
                    path = f"{base}_{idx}{ext}"
                ws.ExportAsFixedFormat(0, path)
                pdf_paths.append(path)
            else:
                ws.PrintOut(Copies=1)
        return {"pages": len(pages), "pdf": pdf_paths}
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        finally:
            xl.Quit()
            shutil.rmtree(tmp_dir, ignore_errors=True)


def main():
    with open(sys.argv[1], encoding="utf-8") as fh:
        job = json.load(fh)
    try:
        res = run_job(job)
        logger.info("sign form done: %s", res)
        return 0
    except Exception:
        logger.exception("sign form failed")
        try:
            from win_notify import send_windows_notification
            send_windows_notification("서명 용지 인쇄 실패", "sign_forms.log 를 확인하세요", duration="short")
        except Exception:
            pass
        return 1


if __name__ == "__main__":
    logging.basicConfig(
        filename=os.path.join(os.path.dirname(os.path.abspath(__file__)), "sign_forms.log"),
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8",
    )
    sys.exit(main())
