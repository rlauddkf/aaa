@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "EXE=%~dp0dist\stock-recommend.exe"

if not exist "%EXE%" (
  echo [INFO] Build the EXE first by running "1_EXE만들기.bat".
  echo.
  pause
  exit /b 1
)

echo Registering a daily task at 17:30 (5:30 PM)...
schtasks /create /tn "StockRecommend" /tr "\"%EXE%\" once-quiet" /sc daily /st 17:30 /f
if errorlevel 1 (
  echo.
  echo [ERROR] Failed to register the task.
  pause
  exit /b 1
)
echo.
echo ============================================
echo    DONE! It will run every day at 17:30.
echo    Results are saved as text files in dist\logs\.
echo    To cancel, run "자동실행_해제.bat".
echo ============================================
echo.
pause
exit /b 0
