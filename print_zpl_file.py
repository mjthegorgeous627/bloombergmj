"""
ZPL 파일 직접 출력 스크립트.
Bloomberg 포털에서 다운로드한 .ZPL 파일을 Zebra 프린터로 바로 전송.

사용법:
  python print_zpl_file.py "C:\...\SHIPLABEL_1 - 2026-03-23T140743.617.ZPL"
  python print_zpl_file.py   (경로 미입력 시 Downloads 폴더 최신 .ZPL 자동 선택)
"""

import sys
import os
import glob

import win32print

PRINTER_NAME = "ZDesigner GK420t"


def find_latest_zpl():
    """Downloads 폴더에서 가장 최신 .ZPL 파일 반환."""
    downloads = os.path.expanduser("~/Downloads")
    files = glob.glob(os.path.join(downloads, "*.ZPL")) + glob.glob(os.path.join(downloads, "*.zpl"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def is_pdf(path):
    with open(path, "rb") as f:
        return f.read(4) == b"%PDF"


def send_raw(data: bytes, printer_name: str) -> bool:
    """bytes 데이터를 RAW 모드로 프린터에 직접 전송."""
    try:
        h = win32print.OpenPrinter(printer_name)
    except Exception as e:
        print(f"[오류] 프린터 '{printer_name}' 연결 실패: {e}")
        print("[참고] 연결된 프린터 목록:")
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1):
            print(f"  - {p[2]}")
        return False

    try:
        win32print.StartDocPrinter(h, 1, ("ZPL Label", None, "RAW"))
        try:
            win32print.StartPagePrinter(h)
            win32print.WritePrinter(h, data)
            win32print.EndPagePrinter(h)
        finally:
            win32print.EndDocPrinter(h)
        print("[완료] 출력 전송 성공")
        return True
    except Exception as e:
        print(f"[오류] 출력 실패: {e}")
        return False
    finally:
        win32print.ClosePrinter(h)


def print_as_zpl(path: str) -> bool:
    """파일 내용을 그대로 ZPL로 전송."""
    with open(path, "rb") as f:
        data = f.read()
    print(f"[ZPL] 파일 크기: {len(data)} bytes")
    return send_raw(data, PRINTER_NAME)


def print_as_pdf(path: str) -> bool:
    """PDF를 파싱 → QR+이름+회사 추출 → 커스텀 ZPL 생성 후 출력."""
    from zebra_handler import parse_pdf_label, build_zpl_label

    info = parse_pdf_label(path)
    print(f"[PDF 파싱] name={info['name']!r}  company={info['company']!r}  qr={info['qr_data'][:40]!r}...")

    if not info['qr_data']:
        print("[경고] QR 데이터 파싱 실패 — PDF 원본을 ZPL로 직접 전송 시도")
        return print_as_zpl(path)

    zpl = build_zpl_label(info['qr_data'], info['name'], info['company'])
    return send_raw(zpl.encode('utf-8'), PRINTER_NAME)


def main():
    if len(sys.argv) >= 2:
        path = " ".join(sys.argv[1:]).strip('"').strip("'")
    else:
        path = find_latest_zpl()
        if not path:
            print("[오류] Downloads 폴더에 .ZPL 파일이 없습니다. 경로를 직접 지정하세요.")
            sys.exit(1)
        print(f"[자동 선택] {path}")

    if not os.path.exists(path):
        print(f"[오류] 파일 없음: {path}")
        sys.exit(1)

    print(f"[파일] {os.path.basename(path)}")

    if is_pdf(path):
        print("[타입] PDF (Bloomberg 포털 라벨) → 파싱 후 출력")
        success = print_as_pdf(path)
    else:
        print("[타입] ZPL 텍스트 → 직접 전송")
        success = print_as_zpl(path)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
