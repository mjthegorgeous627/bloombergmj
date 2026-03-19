"""
Zebra 라벨 프린터 핸들러.
ZDesigner GK420t, USB 연결, 8cm × 5cm 라벨.

레이아웃: 좌측 QR코드 / 우측 이름 + 회사명
"""

import re
import win32print

PRINTER_NAME = "ZDesigner GK420t"

# 203 DPI: 8cm=640dots, 5cm=400dots
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
    QR데이터 + 이름 + 회사명 → 8cm×5cm 라벨 ZPL 문자열 생성.

    레이아웃 (이미지 기준):
      ┌─────────────┬──────────────┐
      │             │  NAME        │
      │   QR CODE   │              │
      │             │  COMPANY     │
      └─────────────┴──────────────┘
    """
    # 텍스트 영역: x=310~630 (약 320dots 폭)
    # 글자 너비 26 기준 → 최대 12자 → 단어 줄바꿈
    MAX_CHARS   = 9
    TEXT_X      = 315

    FONT_H_NAME = 68   # 이름 폰트 높이 (dots)
    FONT_W_NAME = 34   # 이름 폰트 너비 (dots)
    FONT_H_CO   = 56   # 회사명 폰트 높이 (dots)
    FONT_W_CO   = 28   # 회사명 폰트 너비 (dots)
    LINE_GAP    = 10   # 줄 간격

    name_lines = _wrap_text(name, MAX_CHARS)
    co_lines   = _wrap_text(company, MAX_CHARS)

    # 텍스트 블록 총 높이 → 수직 중앙 정렬
    total_h = (
        len(name_lines) * (FONT_H_NAME + LINE_GAP) +
        (12 if name_lines and co_lines else 0) +  # 이름-회사 간격
        len(co_lines) * (FONT_H_CO + LINE_GAP)
    )
    y = max(15, (LABEL_HEIGHT - total_h) // 2)

    lines = [
        "^XA",
        f"^PW{LABEL_WIDTH}",
        f"^LL{LABEL_HEIGHT}",
        "^LH0,0",
        "^CI28",            # UTF-8

        # QR 코드 (좌측)
        "^FO15,15",
        "^BQN,2,10",        # QR code, normal, magnification 10
        f"^FDQA,{qr_data}^FS",
    ]

    # 이름
    for line in name_lines:
        lines += [
            f"^FO{TEXT_X},{y}",
            f"^A0N,{FONT_H_NAME},{FONT_W_NAME}",
            f"^FD{line}^FS",
        ]
        y += FONT_H_NAME + LINE_GAP

    # 이름-회사 간격
    if name_lines and co_lines:
        y += 12

    # 회사명
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


# ── 공개 함수 ────────────────────────────────────────────────────────────────

def print_label(qr_data, name, company):
    """QR 데이터 + 이름 + 회사명 → 라벨 생성 → Zebra 출력."""
    zpl = build_zpl_label(qr_data, name, company)
    print(f"[Zebra] 라벨: {name} / {company}")
    return print_zpl(zpl)
