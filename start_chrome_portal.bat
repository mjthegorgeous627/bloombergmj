@echo off
echo Bloomberg Portal용 Chrome 디버그 모드 실행 중...
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%TEMP%\chrome_portal_debug" "https://bsp.btogo.com/supplier/login"
echo Chrome이 열리면 로그인하고, 이후 test_portal_print.py를 실행하세요.
pause
