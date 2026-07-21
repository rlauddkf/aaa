@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "EXE=%~dp0dist\미국장단타추천.exe"

if not exist "%EXE%" (
  echo [안내] 먼저 "1_EXE만들기.bat" 을 실행해서 exe를 만들어 주세요.
  echo.
  pause
  exit /b 1
)

echo 매일 저녁 5시 30분에 자동 실행되도록 등록합니다...
schtasks /create /tn "미국장단타추천" /tr "\"%EXE%\" once-quiet" /sc daily /st 17:30 /f
if errorlevel 1 (
  echo.
  echo [오류] 등록에 실패했습니다.
  pause
  exit /b 1
)
echo.
echo ============================================
echo    등록 완료! 매일 17:30 자동 분석됩니다.
echo    결과는 dist\logs\ 폴더의 txt 파일로 저장돼요.
echo    (해제하려면 "자동실행_해제.bat" 실행)
echo ============================================
echo.
pause
exit /b 0
