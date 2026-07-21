@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo    미국장 단타 추천 - EXE 만들기
echo ============================================
echo.
echo [1/2] 필요한 파일 설치 중... (최초 1회, 몇 분 걸릴 수 있어요)
call npm install
if errorlevel 1 goto err
echo.
echo [2/2] EXE 빌드 중... (Node 바이너리를 처음 받을 때 시간이 걸립니다)
call npm run build:exe
if errorlevel 1 goto err
echo.
echo ============================================
echo    완성!  dist\미국장단타추천.exe
echo    이 파일을 더블클릭하면 오늘의 추천이 나옵니다.
echo ============================================
echo.
pause
exit /b 0

:err
echo.
echo [오류] 문제가 발생했습니다. 위 메시지를 확인하세요.
echo   - Node.js 가 설치되어 있는지 (명령창에서 node -v)
echo   - 인터넷 연결이 되는지 확인하세요.
pause
exit /b 1
