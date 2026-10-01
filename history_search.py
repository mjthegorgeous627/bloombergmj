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
_SKIP_NAME_PARTS = ("before_last", "실시간", "DN&PN")
_INDEX_LOCK = threading.Lock()

_ORDER_RE = re.compile(r"(?<!\d)(\d{7,9})(?!\d)")
_MD_RE = re.compile(r"(\d{1,2})\s*월\s*(\d{1,2})\s*일")
_DASH_RE = re.compile(r"(?:(\d{4})[-./])?(\d{1,2})[-./](\d{1,2})")

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
    return conn


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
    wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
    stats = {"sheets": 0, "rows": 0, "skipped_no_header": 0, "skipped_dup": 0}
    try:
        for idx, name in enumerate(wb.sheetnames):
            ws = wb[name]
            first = [list(r[:_MAX_COLS]) for r in ws.iter_rows(min_row=1, max_row=4, values_only=True)]
            hi = _header_row_index(first)
            if hi is None:
                stats["skipped_no_header"] += 1
                continue
            title = _cell_text(first[0][0]) if first and first[0] else ""
            sheet_date = _parse_sheet_date(title, name, file_year, idx)
            data = []
            for excel_row, r in enumerate(
                ws.iter_rows(min_row=hi + 2, max_row=hi + 2 + 600, values_only=True), start=hi + 2
            ):
                cells = [_cell_text(c) for c in list(r[:_MAX_COLS])]
                while cells and not cells[-1]:
                    cells.pop()
                if any(cells):
                    data.append((excel_row, cells))
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
                "INSERT INTO sheets(source_file, sheet_name, sheet_date, content_hash, row_count) VALUES (?,?,?,?,?)",
                (key, name, sheet_date, digest, len(data)),
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
                    "INSERT INTO rows(sheet_id, group_no, excel_row, kind, order_no, cells_json, search_text) "
                    "VALUES (?,?,?,?,?,?,?)",
                    (
                        sheet_id,
                        group_no,
                        excel_row,
                        _kind_of(first_cell),
                        order_no,
                        _json.dumps(cells, ensure_ascii=False),
                        " | ".join(c for c in cells if c).lower(),
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
