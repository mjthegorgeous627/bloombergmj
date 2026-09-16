@echo off
REM 2026-08-28: 작업표시줄 트레이 아이콘 실행. 콘솔 창 없이(pythonw) 백그라운드로
REM 뜨고, 이미 떠있으면 tray_icon.py 자신의 단일 인스턴스 잠금이 조용히 종료한다
REM (두 번 눌러도 안전).
cd /d "%~dp0"
start "" pythonw.exe tray_icon.py
