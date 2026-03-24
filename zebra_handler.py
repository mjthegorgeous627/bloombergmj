"""
Zebra 라벨 프린터 핸들러.
ZDesigner GK420t, USB 연결, 8cm × 5cm 라벨.

레이아웃: 좌측 QR코드 / 우측 이름 + 회사명
"""

import re
import win32print

PRINTER_NAME = "ZDesigner GK420t"

# 203 DPI: 8cm(가로)=640dots, 5cm(세로)=400dots  — 가로 라벨
LABEL_WIDTH  = 640
LABEL_HEIGHT = 400


# ── ZPL 파싱 ────────────────────────────────────────────────────────────────

def parse_qr_data_from_zpl(zpl_content):
    """
    다운로드된 .zpl 파일에서 QR코드 데이터 문자열 추출.
    ZPL 패턴: ^BQ...^FDQA,{data}^FS
    """
    # 패턴 1: ^FDQA,{data}^FS
    m = re.search(r'\^FD(?:QA,|MM,A,)(.+?)\^FS', zpl_content, re.DOTALL)
    if m:
        return m.group(1).strip()
    # 패턴 2: ^BQ 바로 다음 ^FD{data}^FS
    m = re.search(r'\^BQ[^\^]*\^FD(.+?)\^FS', zpl_content, re.DOTALL)
    if m:
        return m.group(1).strip()
    return ''


# ── 텍스트 줄바꿈 ─────────────────────────────────────────────────────────────

def _wrap_text(text, max_chars):
    """긴 텍스트를 max_chars 기준으로 단어 단위 줄 분리."""
    if not text:
        return []
    words = text.split()
    lines, current = [], ''
    for word in words:
        candidate = f"{current} {word}".strip() if current else word
        if len(candidate) <= max_chars:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


# ── ZPL 생성 ────────────────────────────────────────────────────────────────

def build_zpl_label(qr_data, name, company):
    """
    QR데이터 + 이름 + 회사명 → 8cm×5cm 가로 라벨 ZPL 생성.

    레이아웃 (겹침 없음):
      ┌──────────┬──────────────────┐
      │          │                  │
      │ QR CODE  │  NAME            │
      │  (좌측)  │  COMPANY         │
      │          │                  │
      └──────────┴──────────────────┘
    203 DPI: 8cm=640dots(가로), 5cm=400dots(세로)
    QR mag 5 → 최대 245dots (세로 400 안에 안전)
    텍스트 영역: x=270 고정 → QR과 겹침 불가
    """
    # QR: 좌측 고정
    # mag 7: version6(41*7=287), version8(49*7=343) → 400 세로 안전
    QR_MAG = 7
    QR_X   = 10
    QR_Y   = 10

    # 텍스트: QR 추정 끝(~300) ~ 라벨 끝(630) 중앙 → x≈465
    # x=400 시작, 사용 가능 폭: 640-400-10 = 230 dots
    TEXT_X    = 400

    # 이름: 한 줄 유지
    # 230dots / 20per char = 11자
    NAME_MAX_CHARS = 25
    FONT_H_NAME    = 52
    FONT_W_NAME    = 20

    # 회사: 최대 2줄
    CO_MAX_CHARS = 13
    FONT_H_CO    = 44
    FONT_W_CO    = 17

    LINE_GAP = 12

    name_lines = _wrap_text(name, NAME_MAX_CHARS)[:1]   # 강제 1줄
    co_lines   = _wrap_text(company, CO_MAX_CHARS)[:2]  # 최대 2줄

    # 텍스트 블록 수직 중앙 정렬
    total_h = (
        len(name_lines) * (FONT_H_NAME + LINE_GAP) +
        (14 if name_lines and co_lines else 0) +
        len(co_lines) * (FONT_H_CO + LINE_GAP)
    )
    y = max(20, (LABEL_HEIGHT - total_h) // 2)

    lines = [
        "^XA",
        f"^PW{LABEL_WIDTH}",
        f"^LL{LABEL_HEIGHT}",
        "^LH0,0",
        "^CI28",

        # QR 코드 (좌측)
        f"^FO{QR_X},{QR_Y}",
        f"^BQN,2,{QR_MAG}",
        f"^FDQA,{qr_data}^FS",
    ]

    # 이름 (우측)
    for line in name_lines:
        lines += [
            f"^FO{TEXT_X},{y}",
            f"^A0N,{FONT_H_NAME},{FONT_W_NAME}",
            f"^FD{line}^FS",
        ]
        y += FONT_H_NAME + LINE_GAP

    if name_lines and co_lines:
        y += 14

    # 회사명 (우측)
    for line in co_lines:
        lines += [
            f"^FO{TEXT_X},{y}",
            f"^A0N,{FONT_H_CO},{FONT_W_CO}",
            f"^FD{line}^FS",
        ]
        y += FONT_H_CO + LINE_GAP

    lines.append("^XZ")
    return "\n".join(lines)


# ── 출력 ────────────────────────────────────────────────────────────────────

def print_zpl(zpl_string, printer_name=PRINTER_NAME):
    """ZPL 문자열을 USB Zebra 프린터로 직접 전송 (드라이버 우회, RAW 모드)."""
    try:
        h = win32print.OpenPrinter(printer_name)
    except Exception as e:
        print(f"[Zebra] 프린터 '{printer_name}' 연결 실패: {e}")
        print("[Zebra] 연결된 프린터 목록:")
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1):
            print(f"  - {p[2]}")
        return False

    try:
        win32print.StartDocPrinter(h, 1, ("Zebra Label", None, "RAW"))
        try:
            win32print.StartPagePrinter(h)
            win32print.WritePrinter(h, zpl_string.encode('utf-8'))
            win32print.EndPagePrinter(h)
        finally:
            win32print.EndDocPrinter(h)
        print("[Zebra] 출력 완료")
        return True
    except Exception as e:
        print(f"[Zebra] 출력 실패: {e}")
        return False
    finally:
        win32print.ClosePrinter(h)


# ── PDF 라벨 파싱 (Bloomberg .ZPL 파일은 실제 PDF) ───────────────────────────

def parse_pdf_label(pdf_path):
    """
    Bloomberg 포털에서 다운로드된 .ZPL 파일(실제 PDF) 파싱.
    반환: {'qr_data': '...', 'name': '...', 'company': '...', 'phone': '...'}

    Bloomberg 라벨 구조 (PDF 좌표 기준):
      - QR 코드: 상단 (x≈231, y≈705 in PDF coords)
      - Ship To 데이터: 하단 좌측 (x≈68, y≈360~432 in PDF coords)
        → PyMuPDF 좌표 (y=0 상단): y≈359~432
    """
    import fitz
    from pyzbar.pyzbar import decode as pyzbar_decode
    from PIL import Image

    result = {'qr_data': '', 'name': '', 'company': '', 'phone': ''}
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]

        # 1. QR 코드 디코딩: 3x 확대 렌더링 → pyzbar
        mat = fitz.Matrix(3, 3)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        for d in pyzbar_decode(img):
            if d.type == 'QRCODE':
                result['qr_data'] = d.data.decode('utf-8', errors='replace')
                break

        # 2. Ship To 텍스트 파싱
        # Bloomberg 라벨: Ship To 데이터는 x≈68~310, y≈345~450 (PyMuPDF 좌표)
        # "Ship To:" 레이블(x≈17)은 clip으로 제외
        clip = fitz.Rect(60, 345, 310, 450)
        section = page.get_text("text", clip=clip, sort=True)
        lines = [l.strip() for l in section.splitlines() if l.strip()]

        if lines:
            result['name'] = lines[0]
        if len(lines) > 1:
            result['phone'] = lines[1]
        if len(lines) > 2:
            result['company'] = lines[2]

        # 폴백: 좌표로 못 찾으면 "Ship To:" 마커 기반 파싱
        if not result['name']:
            _parse_pdf_ship_to_fallback(page, result)

        doc.close()
    except Exception as e:
        logger.error(f"PDF 파싱 실패 ({pdf_path}): {e}")

    return result


def _parse_pdf_ship_to_fallback(page, result):
    """전체 텍스트에서 'Ship To:' 마커 이후 섹션 파싱."""
    full = page.get_text("text", sort=True)
    lines = [l.strip() for l in full.splitlines() if l.strip()]
    for i, line in enumerate(lines):
        if line == "Ship To:":
            after = lines[i + 1:]
            if after:
                result['name'] = after[0]
            if len(after) > 1:
                result['phone'] = after[1]
            if len(after) > 2:
                result['company'] = after[2]
            return
    # 마지막 폴백: 전화번호 패턴으로 위치 추정
    import re
    phone_re = re.compile(r'^\+\d[\d\-]{8,}$')
    for i, line in enumerate(lines):
        if phone_re.match(line):
            if i > 0:
                result['name'] = lines[i - 1]
            result['phone'] = line
            if i + 1 < len(lines):
                result['company'] = lines[i + 1]
            break


# ── 공개 함수 ────────────────────────────────────────────────────────────────

def print_label(qr_data, name, company):
    """QR 데이터 + 이름 + 회사명 → 라벨 생성 → Zebra 출력."""
    zpl = build_zpl_label(qr_data, name, company)
    print(f"[Zebra] 라벨: {name} / {company}")
    return print_zpl(zpl)
