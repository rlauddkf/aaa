@echo off
chcp 65001 >nul
schtasks /delete /tn "미국장단타추천" /f
if errorlevel 1 (
  echo [안내] 등록된 자동 실행이 없거나 해제에 실패했습니다.
) else (
  echo 자동 실행이 해제되었습니다.
)
echo.
pause
