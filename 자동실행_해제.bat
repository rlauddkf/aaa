@echo off
chcp 65001 >nul
schtasks /delete /tn "StockRecommend" /f
if errorlevel 1 (
  echo [INFO] No scheduled task found, or it could not be removed.
) else (
  echo Daily auto-run has been canceled.
)
echo.
pause
