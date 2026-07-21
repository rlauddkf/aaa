@echo off
REM ===================================================================
REM  무한 피하기 게임 -> Windows 실행파일(.exe) 빌드 스크립트
REM  Windows에서 이 파일을 더블클릭하거나 명령창에서 실행하세요.
REM ===================================================================

echo [1/3] 필요한 패키지 설치 중...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto error

echo [2/3] exe 빌드 중...
pyinstaller --onefile --windowed --name Dodge game.py
if errorlevel 1 goto error

echo [3/3] 완료!
echo 실행파일 위치: dist\Dodge.exe
pause
exit /b 0

:error
echo.
echo 빌드 중 오류가 발생했습니다. Python이 설치되어 있는지 확인하세요.
pause
exit /b 1
