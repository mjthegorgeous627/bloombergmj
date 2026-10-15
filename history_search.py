"""Past-delivery history search for the workbench (2026-10-01).

마감하면 workbench.db에서 행이 사라지고 엑셀에만 남아서, 예전 시리얼/오더번호를
찾으려면 연도별 배송장 엑셀을 열어 Ctrl+F 해야 했다. 이 모듈은 C:\\1\\배송장의
연도별 엑셀(2017~)과 2026 마감/공유파일의 "날짜별 시트"를 읽어 별도의
history.db(SQLite)에 색인해 두고, workbench에서 바로 검색/상세 확인하게 한다.

- workbench.db는 건드리지 않는다(읽기 전용 색인, 엑셀이 원본).
- 컬럼 위치는 연도별로 제각각(S/N 칸이 밀린 행 등)이라 필드 파싱에 의존하지 않고
  행의 모든 셀 텍스트를 그대로 검색 대상으로 삼는다. 표시용 필드는 best-effort.
- 시트 판별: 상단 4행 안에 'M/N'과 'S/N' 헤더가 있는 시트만 색인한다. (시트가
  아닌 'Backup'처럼 계정정보가 적힌 시트가 섞여 있어 헤더 없는 시트는 무조건 제외)
- 같은 날짜의 내용이 동일한 시트(예: '2019 배송장_2' 복사본)는 해시로 중복 제거.
"""

from __future__ import annotations

import hashlib
import logging
import re
import shutil
import sqlite3
import tempfile
import threading
from pathlib import Path

import openpyxl

logger = logging.getLogger("workbench")

BASE_DIR = Path(__file__).resolve().parent
HISTORY_DB_PATH = BASE_DIR / "history.db"
SHIPPING_DIR = Path(r"C:\1\배송장")
_MAX_COLS = 14
_SKIP_NAME_PARTS = ("before_last", "실시간", "DN&PN", "_마감")  # _마감: 2026-10-08 폐지(배송장 파일 하나로 통합)
_INDEX_LOCK = threading.Lock()

_ORDER_RE = re.compile(r"(?<!\d)(\d{7,9})(?!\d)")
_MD_RE = re.compile(r"(\d{1,2})\s*월\s*(\d{1,2})\s*일")
_DASH_RE = re.compile(r"(?:(\d{4})[-./])?(\d{1,2})[-./](\d{1,2})")
_DASH_LOOSE_RE = re.compile(r"(?<!\d)(\d{1,2})\s*[-./]\s*(\d{1,2})(?!\d)")

SCHEMA = """
CREATE TABLE IF NOT EXISTS sheets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_file TEXT NOT NULL,
    sheet_name TEXT NOT NULL,
    sheet_date TEXT,
    content_hash TEXT NOT NULL,
    row_count INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_sheets_file ON sheets(source_file);
CREATE INDEX IF NOT EXISTS idx_sheets_hash ON sheets(sheet_date, content_hash);
CREATE TABLE IF NOT EXISTS rows (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet_id INTEGER NOT NULL,
    group_no INTEGER NOT NULL,
    excel_row INTEGER NOT NULL,
    kind TEXT,
    order_no TEXT,
    cells_json TEXT NOT NULL,
    search_text TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_rows_sheet ON rows(sheet_id, group_no);
CREATE TABLE IF NOT EXISTS index_state (
    source_file TEXT PRIMARY KEY,
    mtime REAL NOT NULL,
    size INTEGER NOT NULL,
    indexed_at TEXT
);
"""


def _connect():
    conn = sqlite3.connect(HISTORY_DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    # 2026-10-15: 색/병합/헤더없는 시트 표시용 컬럼 (기존 DB에는 ALTER로 추가)
    for table, col, ddl in (("rows", "fills_json", "TEXT"), ("sheets", "merges_json", "TEXT"), ("sheets", "no_header", "INTEGER DEFAULT 0")):
        have = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
        if col not in have:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ddl}")
    conn.commit()
    return conn


def _json_dumps(o):
    import json
    return json.dumps(o)


def _fill_hex(cell):
    """칸 배경색(단색 RGB)만 'RRGGBB'로. 흰색/없음/테마색은 None."""
    f = getattr(cell, "fill", None)
    if not f or f.fill_type != "solid":
        return None
    col = f.fgColor
    if col is None or col.type != "rgb" or not isinstance(col.rgb, str):
        return None
    rgb = col.rgb[-6:].upper()
    return None if rgb in ("FFFFFF", "000000") else rgb


def _sheet_merges(xlsx_path):
    """시트 XML의 <mergeCells>를 직접 읽는다(openpyxl read_only는 병합 정보를 안 줌).
    반환: {시트명: [[r1,c1,r2,c2], ...]} (열 _MAX_COLS 이내)."""
    import zipfile
    import xml.etree.ElementTree as ET
    from openpyxl.utils.cell import range_boundaries

    out = {}
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    with zipfile.ZipFile(xlsx_path) as z:
        wbx = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = {r.get("Id"): r.get("Target") for r in rels}
        for s in wbx.find("m:sheets", ns):
            rid = s.get("{%s}id" % ns["r"])
            t = (target.get(rid) or "").lstrip("/")
            path = t if t.startswith("xl/") else "xl/" + t
            try:
                with z.open(path) as fh:
                    merges = []
                    for _, el in ET.iterparse(fh, events=("end",)):
                        tag = el.tag.rsplit("}", 1)[-1]
                        if tag == "mergeCell":
                            c1, r1, c2, r2 = range_boundaries(el.get("ref"))
                            if c1 <= _MAX_COLS and (r2 > r1 or c2 > c1):
                                merges.append([r1, c1, r2, min(c2, _MAX_COLS)])
                        elif tag == "row":
                            el.clear()
                    out[s.get("name")] = merges
            except KeyError:
                out[s.get("name")] = []
    return out


def _cell_text(v):
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _file_year(path: Path):
    m = re.search(r"(20\d\d)", path.name)
    return int(m.group(1)) if m else None


def _parse_sheet_date(title, sheet_name, file_year, sheet_idx):
    """시트 제목(A1)이나 시트명에서 날짜를 뽑는다. 못 뽑으면 None."""
    month = day = year = None
    for text in (title or "", sheet_name or ""):
        m = _MD_RE.search(text)
        if m:
            month, day = int(m.group(1)), int(m.group(2))
            ym = re.search(r"(20\d\d)", text)
            year = int(ym.group(1)) if ym else None
            break
        m = _DASH_RE.search(text)
        if m:
            year = int(m.group(1)) if m.group(1) else None
            month, day = int(m.group(2)), int(m.group(3))
            break
    if not month and sheet_name:
        # 2026-10-15: "10 -12", "6- 13"처럼 숫자 사이에 공백이 낀 시트명(2017~2018 파일에 많음).
        # 제목(A1)에서는 오탐 위험이 있어 시트명에만 느슨한 규칙을 쓴다.
        m = _DASH_LOOSE_RE.match(sheet_name.strip()) or _DASH_LOOSE_RE.search(sheet_name)
        if m:
            month, day = int(m.group(1)), int(m.group(2))
    if not month or not day or not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    if year is None:
        year = file_year
        # 연초 파일 맨 앞에 전년도 12월 시트가 끼어 있는 경우(예: 2021 파일의 12-31).
        if year and month == 12 and sheet_idx < 3:
            year -= 1
    if not year:
        return None
    return f"{year:04d}-{month:02d}-{day:02d}"


def _header_row_index(first_rows):
    for i, row in enumerate(first_rows):
        up = [str(c).strip().upper() for c in row if c is not None]
        if "M/N" in up and "S/N" in up:
            return i
    return None


def _kind_of(first_cell):
    t = first_cell or ""
    if "회수" in t:
        return "회수"
    if "배송" in t:
        return "배송"
    return ""


def _index_workbook(conn, path: Path):
    """파일 하나의 모든 날짜 시트를 읽어 색인(해당 파일의 기존 색인은 교체)."""
    tmp = Path(tempfile.gettempdir()) / f"history_idx_{abs(hash(str(path)))}.xlsx"
    shutil.copy2(path, tmp)  # 엑셀이 열어둔 파일도 읽을 수 있게 복사본으로
    file_year = _file_year(path)
    key = str(path)
    conn.execute("DELETE FROM rows WHERE sheet_id IN (SELECT id FROM sheets WHERE source_file=?)", (key,))
    conn.execute("DELETE FROM sheets WHERE source_file=?", (key,))
    try:
        merges_by_sheet = _sheet_merges(tmp)
    except Exception:
        logger.warning("merge scan failed: %s", path, exc_info=True)
        merges_by_sheet = {}
    wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
    stats = {"sheets": 0, "rows": 0, "skipped_no_header": 0, "skipped_dup": 0}
    try:
        for idx, name in enumerate(wb.sheetnames):
            ws = wb[name]
            first = [list(r[:_MAX_COLS]) for r in ws.iter_rows(min_row=1, max_row=4, values_only=True)]
            hi = _header_row_index(first)
            no_header = 0
            if hi is None:
                # 2026-10-15: 헤더(M/N,S/N) 없는 시트도 메모/이벤트/주소록 등으로 쓰였으므로 전부 색인한다.
                no_header = 1
                hi = -1
                stats["skipped_no_header"] += 1   # (이름은 그대로 두되 이제는 '헤더 없는 시트 수'로 집계)
            title = _cell_text(first[0][0]) if first and first[0] else ""
            sheet_date = _parse_sheet_date(title, name, file_year, idx)
            data = []
            fills = {}
            for excel_row, r in enumerate(
                ws.iter_rows(min_row=hi + 2, max_row=hi + 2 + 600, max_col=_MAX_COLS), start=hi + 2
            ):
                cells = [_cell_text(c.value) for c in r]
                while cells and not cells[-1]:
                    cells.pop()
                if any(cells):
                    data.append((excel_row, cells))
                    fm = {i: _fill_hex(c) for i, c in enumerate(r)}
                    fm = {i: v for i, v in fm.items() if v}
                    if fm:
                        fills[excel_row] = fm
            if not data:
                continue
            digest = hashlib.md5(repr([c for _, c in data]).encode("utf-8")).hexdigest()
            dup = conn.execute(
                "SELECT 1 FROM sheets WHERE sheet_date IS ? AND content_hash=? AND source_file<>?",
                (sheet_date, digest, key),
            ).fetchone()
            if dup:
                stats["skipped_dup"] += 1
                continue
            cur = conn.execute(
                "INSERT INTO sheets(source_file, sheet_name, sheet_date, content_hash, row_count, merges_json, no_header) VALUES (?,?,?,?,?,?,?)",
                (key, name, sheet_date, digest, len(data), _json_dumps(merges_by_sheet.get(name) or []), no_header),
            )
            sheet_id = cur.lastrowid
            group_no, cur_order = 0, None
            import json as _json

            for excel_row, cells in data:
                first_cell = cells[0] if cells else ""
                m = _ORDER_RE.search(first_cell)
                order_no = m.group(1) if m else ""
                if first_cell and (order_no != cur_order or not order_no):
                    group_no += 1
                    cur_order = order_no
                elif group_no == 0:
                    group_no = 1
                conn.execute(
                    "INSERT INTO rows(sheet_id, group_no, excel_row, kind, order_no, cells_json, search_text, fills_json) "
                    "VALUES (?,?,?,?,?,?,?,?)",
                    (
                        sheet_id,
                        group_no,
                        excel_row,
                        _kind_of(first_cell),
                        order_no,
                        _json.dumps(cells, ensure_ascii=False),
                        " | ".join(c for c in cells if c).lower(),
                        _json.dumps(fills[excel_row]) if excel_row in fills else None,
                    ),
                )
            stats["sheets"] += 1
            stats["rows"] += len(data)
    finally:
        wb.close()
        try:
            tmp.unlink()
        except OSError:
            pass
    st = path.stat()
    conn.execute(
        "INSERT INTO index_state(source_file, mtime, size, indexed_at) VALUES (?,?,?,datetime('now','localtime')) "
        "ON CONFLICT(source_file) DO UPDATE SET mtime=excluded.mtime, size=excluded.size, indexed_at=excluded.indexed_at",
        (key, st.st_mtime, st.st_size),
    )
    conn.commit()
    return stats


def source_files():
    files = []
    for pattern_dir in (SHIPPING_DIR, SHIPPING_DIR / "backup"):
        if not pattern_dir.exists():
            continue
        for p in sorted(pattern_dir.glob("*.xlsx")):
            if p.name.startswith("~$") or any(s in p.name for s in _SKIP_NAME_PARTS):
                continue
            files.append(p)
    return files


def reindex(only_changed=True, files=None):
    """전체(또는 지정한 파일들) 색인. 이미 같은 mtime/size로 색인된 파일은 건너뜀."""
    if not _INDEX_LOCK.acquire(blocking=False):
        return {"status": "busy"}
    try:
        conn = _connect()
        report = {}
        try:
            for p in files or source_files():
                st = p.stat()
                prev = conn.execute(
                    "SELECT mtime, size FROM index_state WHERE source_file=?", (str(p),)
                ).fetchone()
                if only_changed and prev and prev["mtime"] == st.st_mtime and prev["size"] == st.st_size:
                    continue
                try:
                    report[p.name] = _index_workbook(conn, p)
                    logger.info("history index %s: %s", p.name, report[p.name])
                except Exception as exc:
                    conn.rollback()
                    report[p.name] = {"error": str(exc)}
                    logger.error("history index failed: %s", p, exc_info=True)
        finally:
            conn.close()
        return {"status": "ok", "files": report}
    finally:
        _INDEX_LOCK.release()


def reindex_async(files=None):
    """마감 직후처럼 호출자를 막으면 안 되는 곳에서 쓰는 백그라운드 색인."""
    t = threading.Thread(target=lambda: reindex(only_changed=False, files=files), daemon=True)
    t.start()


def _row_dict(r):
    import json as _json

    cells = _json.loads(r["cells_json"])
    return {
        "rowId": r["id"],
        "groupNo": r["group_no"],
        "excelRow": r["excel_row"],
        "kind": r["kind"] or "",
        "orderNo": r["order_no"] or "",
        "cells": cells,
        "fills": (_json.loads(r["fills_json"]) if r["fills_json"] else {}),
    }


def search(query, limit=200):
    terms = [t for t in re.split(r"\s+", (query or "").strip().lower()) if t]
    if not terms:
        return {"results": [], "total": 0}
    conn = _connect()
    try:
        where = " AND ".join("r.search_text LIKE ? ESCAPE '\\'" for _ in terms)
        params = ["%" + t.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%" for t in terms]
        total = conn.execute(
            f"SELECT COUNT(*) FROM rows r WHERE {where}", params
        ).fetchone()[0]
        rows = conn.execute(
            f"""SELECT r.*, s.sheet_date, s.sheet_name, s.source_file
                FROM rows r JOIN sheets s ON s.id = r.sheet_id
                WHERE {where}
                ORDER BY s.sheet_date DESC, r.sheet_id DESC, r.excel_row
                LIMIT ?""",
            params + [int(limit)],
        ).fetchall()
        out = []
        for r in rows:
            d = _row_dict(r)
            d.update(
                sheetDate=r["sheet_date"] or "",
                sheetName=r["sheet_name"],
                sourceFile=Path(r["source_file"]).name,
                sheetId=r["sheet_id"],
            )
            out.append(d)
        return {"results": out, "total": total}
    finally:
        conn.close()


def group_detail(sheet_id, group_no):
    """검색 결과 한 줄이 속한 오더 묶음(배송/회수 행 전체)을 돌려준다."""
    conn = _connect()
    try:
        sh = conn.execute("SELECT * FROM sheets WHERE id=?", (sheet_id,)).fetchone()
        if not sh:
            return None
        rows = conn.execute(
            "SELECT * FROM rows WHERE sheet_id=? AND group_no=? ORDER BY excel_row",
            (sheet_id, group_no),
        ).fetchall()
        return {
            "sheetDate": sh["sheet_date"] or "",
            "sheetName": sh["sheet_name"],
            "sourceFile": Path(sh["source_file"]).name,
            "rows": [_row_dict(r) for r in rows],
        }
    finally:
        conn.close()


def restore_to_workbench(row_ids, target_date=None):
    """선택한 이력 행을 workbench 보드에 메모 행으로 되돌린다(2026-10-06 사용자 요청).

    items.add_manual_note()/copy_items()와 같은 안전 규칙: order_no는 항상 빈값이라
    과거 실제 오더를 되돌려도 자동화(Portal/ZREC/POD)가 살아있는 오더로 오인하지
    않는다. 원래 오더번호는 Order# 라벨(order_type)에 글자로만 남긴다. 컬럼 위치가
    연도별로 밀려 있으므로 best-effort 매핑: 0=오더/구분(접두 배송|회수 제거),
    1=제목(자재 설명), 2=M/N, 3=S/N, 4=고객, 5=연락처, 6=주소, 7=메모.
    """
    import json as _json
    from items import add_manual_note

    ids = [int(i) for i in (row_ids or []) if str(i).strip()]
    if not ids:
        return {"restored": 0}
    conn = _connect()
    try:
        ph = ",".join("?" for _ in ids)
        rows = conn.execute(f"SELECT * FROM rows WHERE id IN ({ph}) ORDER BY sheet_id, excel_row", ids).fetchall()
    finally:
        conn.close()
    created = []
    for r in rows:
        c = [str(x or "").strip() for x in _json.loads(r["cells_json"])] + [""] * 8
        label = re.sub(r"^(배송|회수)\s*", "", c[0]).strip().replace("\n", " ")
        title = c[1] or c[2] or label
        if not title:
            continue
        res = add_manual_note(
            [title, c[2], c[3], c[4], c[5], c[6], c[7]], target_date, r["kind"] or "배송", label or "memo", 7,
        )
        created.append(res["orderId"])
    return {"restored": len(created), "orderIds": created}


def day_detail(date_iso=None, sheet_id=None):
    """달력에서 마감된 과거 날짜(또는 날짜 없는 시트 하나)의 시트와 행 전체(+색/병합 정보)."""
    import json as _json

    conn = _connect()
    try:
        if sheet_id:
            shs = conn.execute("SELECT * FROM sheets WHERE id=?", (int(sheet_id),)).fetchall()
        else:
            shs = conn.execute("SELECT * FROM sheets WHERE sheet_date=? ORDER BY no_header, id", (date_iso,)).fetchall()
        out = []
        for sh in shs:
            rows = conn.execute("SELECT * FROM rows WHERE sheet_id=? ORDER BY excel_row", (sh["id"],)).fetchall()
            out.append({
                "sheetId": sh["id"],
                "sheetName": sh["sheet_name"],
                "sourceFile": Path(sh["source_file"]).name,
                "noHeader": bool(sh["no_header"]),
                "merges": _json.loads(sh["merges_json"]) if sh["merges_json"] else [],
                "rows": [_row_dict(r) for r in rows],
            })
        return {"date": date_iso or "", "sheets": out}
    finally:
        conn.close()


def misc_sheets():
    """날짜를 못 읽은 시트(메모/이벤트/주소록 등) 목록."""
    conn = _connect()
    try:
        return {"sheets": [
            {"sheetId": r["id"], "sheetName": r["sheet_name"], "sourceFile": Path(r["source_file"]).name, "rows": r["row_count"]}
            for r in conn.execute("SELECT * FROM sheets WHERE sheet_date IS NULL ORDER BY source_file DESC, id").fetchall()
        ]}
    finally:
        conn.close()


_EDIT_LOCK = threading.Lock()


def _open_workbook_com(path: str):
    """이미 열려 있으면 그 통합문서에 붙고, 아니면 숨김 Excel로 연다. (wb, 정리함수) 반환."""
    import os
    import win32com.client

    try:
        wb = win32com.client.GetObject(os.path.abspath(path))   # 사용자가 열어둔 파일(ROT)
        if getattr(wb, "Name", None):
            return wb, (lambda: None)
    except Exception:
        pass
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=0)
    except Exception:
        xl.Quit()
        raise

    def cleanup():
        try:
            wb.Close(SaveChanges=False)
        finally:
            xl.Quit()

    return wb, cleanup


def update_cells(edits):
    """이력 화면에서 고친 셀만 원본 엑셀에 반영한다(시트 전체를 다시 쓰지 않음).

    edits: [{rowId, col(0부터), value}]. 엑셀의 현재 값이 색인해 둔 값과 다르면
    (그 사이 누가 고쳤거나 행이 밀린 경우) 덮어쓰지 않고 conflict로 돌려준다.
    성공하면 history.db도 같은 값으로 갱신하고 index_state를 새 파일 상태로 맞춰
    불필요한 전체 재색인을 피한다.
    """
    import json as _json
    from order_lock import ExcelWriteLock

    results = []
    todo = {}   # source_file -> sheet_name -> [(rowId, excel_row, col, old, new)]
    conn = _connect()
    try:
        for e in edits or []:
            rid, col, new = int(e.get("rowId")), int(e.get("col")), str(e.get("value") if e.get("value") is not None else "")
            r = conn.execute(
                "SELECT r.*, s.sheet_name, s.source_file FROM rows r JOIN sheets s ON s.id=r.sheet_id WHERE r.id=?", (rid,)
            ).fetchone()
            if not r or not (0 <= col < _MAX_COLS):
                results.append({"rowId": rid, "col": col, "ok": False, "error": "대상 행을 찾을 수 없음"})
                continue
            cells = _json.loads(r["cells_json"])
            old = cells[col] if col < len(cells) else ""
            todo.setdefault(r["source_file"], {}).setdefault(r["sheet_name"], []).append((rid, r["excel_row"], col, old, new))
    finally:
        conn.close()

    with _EDIT_LOCK:
        for path, sheets in todo.items():
            done = []   # 엑셀에 실제로 쓴 것
            wb = cleanup = None
            try:
                if path not in {str(p) for p in source_files()}:
                    raise RuntimeError("원본 파일을 찾을 수 없음")
                with ExcelWriteLock():
                    wb, cleanup = _open_workbook_com(path)
                    if wb.ReadOnly:
                        raise RuntimeError("엑셀 파일이 읽기 전용으로 열려 있어 저장할 수 없음 (다른 사람이 쓰는 중일 수 있음)")
                    for sheet_name, items in sheets.items():
                        ws = wb.Worksheets(sheet_name)
                        for rid, xr, col, old, new in items:
                            cell = ws.Cells(xr, col + 1)
                            cur = _cell_text(cell.Value)
                            if cur != old:
                                results.append({"rowId": rid, "col": col, "ok": False, "conflict": True,
                                                "error": "엑셀의 현재 값이 달라서 덮어쓰지 않았습니다 (현재: %s)" % cur[:40]})
                                continue
                            if new == old:
                                results.append({"rowId": rid, "col": col, "ok": True})
                                continue
                            cell.Value = new if new != "" else None
                            done.append((rid, col, new))
                            results.append({"rowId": rid, "col": col, "ok": True})
                    if done:
                        wb.Save()
            except Exception as exc:
                logger.error("history edit failed: %s", path, exc_info=True)
                results = [x for x in results if not any(x["rowId"] == d[0] and x["col"] == d[1] for d in done)]
                done = []
                for items in sheets.values():
                    for rid, xr, col, old, new in items:
                        if not any(x["rowId"] == rid and x["col"] == col for x in results):
                            results.append({"rowId": rid, "col": col, "ok": False, "error": str(exc)[:120]})
            finally:
                if cleanup:
                    try:
                        cleanup()
                    except Exception:
                        pass
            if done:
                conn = _connect()
                try:
                    for rid, col, new in done:
                        r = conn.execute("SELECT * FROM rows WHERE id=?", (rid,)).fetchone()
                        cells = _json.loads(r["cells_json"])
                        while len(cells) <= col:
                            cells.append("")
                        cells[col] = new
                        while cells and not cells[-1]:
                            cells.pop()
                        first = cells[0] if cells else ""
                        m = _ORDER_RE.search(first)
                        conn.execute(
                            "UPDATE rows SET cells_json=?, search_text=?, kind=?, order_no=? WHERE id=?",
                            (_json.dumps(cells, ensure_ascii=False), " | ".join(c for c in cells if c).lower(),
                             _kind_of(first) if col == 0 else r["kind"], (m.group(1) if m else "") if col == 0 else r["order_no"], rid),
                        )
                    st = Path(path).stat()
                    conn.execute("UPDATE index_state SET mtime=?, size=?, indexed_at=datetime('now','localtime') WHERE source_file=?",
                                 (st.st_mtime, st.st_size, path))
                    conn.commit()
                finally:
                    conn.close()
    return {"results": results}


def stats():
    conn = _connect()
    try:
        n_rows = conn.execute("SELECT COUNT(*) FROM rows").fetchone()[0]
        n_sheets = conn.execute("SELECT COUNT(*) FROM sheets").fetchone()[0]
        rng = conn.execute("SELECT MIN(sheet_date), MAX(sheet_date) FROM sheets").fetchone()
        return {"rows": n_rows, "sheets": n_sheets, "from": rng[0], "to": rng[1], "indexing": _INDEX_LOCK.locked()}
    finally:
        conn.close()


if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO)
    print(reindex(only_changed="--force" not in sys.argv))
    print(stats())
