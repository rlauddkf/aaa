@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo    Stock Recommender - Build EXE
echo ============================================
echo.
echo [1/2] Installing packages (first time only, a few minutes)...
call npm install
if errorlevel 1 goto err
echo.
echo [2/2] Building EXE (downloading Node binary on first run)...
call npm run build:exe
if errorlevel 1 goto err
echo.
echo ============================================
echo    DONE!  ^> dist\stock-recommend.exe
echo    Double-click that file to see today's pick.
echo ============================================
echo.
pause
exit /b 0

:err
echo.
echo [ERROR] Something went wrong. See the messages above.
echo   - Is Node.js installed?   (open cmd and run:  node -v )
echo   - Are you connected to the internet?
pause
exit /b 1
