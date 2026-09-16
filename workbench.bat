@echo off
REM 2026-08-25: pythonw.exe로 콘솔 창 없이 백그라운드 실행 - "콘솔 창을 실수로
REM 클릭/드래그하면 Quick Edit Mode 때문에 프로세스가 그대로 멈춰버리는" 위험을
REM 없애기 위함. 크래시/에러는 이제 콘솔 창이 아니라 workbench_error.log에
REM 기록되고(2026-08-24 도입), 시작될 때마다 윈도우 토스트 알림이 뜨므로
REM (workbench_app.py의 serve() 참고) 창이 없어도 "떠 있나"는 확인 가능하다.
REM 이 cmd 창 자체는 pythonw를 띄우자마자 바로 닫힌다(대기하지 않음).
cd /d "%~dp0"
start "" pythonw.exe workbench_app.py
