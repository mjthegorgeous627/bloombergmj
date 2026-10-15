"""Windows 내장 OCR로 터미널 캡처를 글자로 바꾸는 도구 (2026-10-15).

터미널은 복사(Ctrl+C)와 창 글자 읽기(UI Automation)가 모두 막혀 있어(Chromium 렌더링) 글자를 직접 못 가져온다.
그래서 캡처 영역만 잘라 OCR한다. 영어 엔진(en-US)은 영어가 정확하고 한글을 못 읽고, 한국어 엔진(ko)은 한글은 읽지만
영어에 오타가 섞인다(예: A닙B, tO) -> 두 결과를 모두 돌려주고 읽는 쪽(Claude)이 알맞은 쪽을 쓴다.
"""
import difflib
import os
import re
import subprocess
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PS1 = os.path.join(HERE, "ocr.ps1")


def ocr_files(paths, lang):
    """이미지 파일 경로 목록 -> 파일별 줄 목록(같은 순서)."""
    if not paths:
        return []
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", PS1, "-Lang", lang, "-Files", ";".join(paths)]
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=120,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if p.returncode != 0:
        raise RuntimeError(f"OCR 실패({lang}): {(p.stderr or p.stdout)[-300:]}")
    out, cur = [], None
    for ln in p.stdout.splitlines():
        if ln.startswith("====="):
            cur = []
            out.append(cur)
        elif cur is not None:
            cur.append(ln)
    return out


def ocr_images(images, langs=("en-US", "ko")):
    """PIL 이미지 목록 -> {lang: [줄목록, ...]} (이미지 순서 유지)."""
    tmp = []
    try:
        for im in images:
            f = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
            f.close()
            im.save(f.name)
            tmp.append(f.name)
        return {lg: ocr_files(tmp, lg) for lg in langs}
    finally:
        for f in tmp:
            try:
                os.remove(f)
            except OSError:
                pass


def _norm(s):
    return re.sub(r"[^0-9a-z가-힣]", "", s.lower())


def merge_frames(frames_lines, min_overlap=2):
    """스크롤하며 겹쳐 찍힌 프레임들(시간 순서: 위쪽 -> 아래쪽)의 OCR 줄을 이어 붙이며 겹친 줄을 제거한다."""
    merged = []
    for lines in frames_lines:
        lines = [x for x in lines if x.strip() and x.strip() != "Return to Top"]
        if not merged:
            merged = list(lines)
            continue
        best = 0
        for k in range(min(len(lines), len(merged), 60), min_overlap - 1, -1):
            a = [_norm(x) for x in merged[-k:]]
            b = [_norm(x) for x in lines[:k]]
            ok = sum(1 for x, y in zip(a, b) if x == y or difflib.SequenceMatcher(None, x, y).ratio() >= 0.88)
            if ok >= max(min_overlap, int(k * 0.8)):
                best = k
                break
        merged.extend(lines[best:])
    return merged


def block(title, per_lang):
    """{lang: 줄목록} -> 읽기 좋은 문자열 블록."""
    parts = [f"### {title}"]
    for lg, lines in per_lang.items():
        parts.append(f"[{ 'English OCR' if lg.startswith('en') else '한국어 OCR' }]")
        parts.extend(lines)
    return "\n".join(parts)
