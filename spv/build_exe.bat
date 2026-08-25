@echo off
REM ===================================================================
REM  SPV 분석기 -> Windows EXE 빌드 스크립트
REM  이 파일을 Windows 에서 더블클릭하면 dist\ 폴더에 EXE 가 생깁니다.
REM  (Python 3.8+ 이 설치되어 있어야 합니다: https://www.python.org)
REM ===================================================================
setlocal
cd /d "%~dp0"
chcp 65001 >nul

echo.
echo [1/3] Python 확인...
where python >nul 2>&1
if errorlevel 1 (
  echo   [오류] Python 을 찾을 수 없습니다. https://www.python.org 에서 설치하세요.
  echo         설치 시 "Add Python to PATH" 를 체크하세요.
  pause
  exit /b 1
)
python --version

echo.
echo [2/3] PyInstaller 준비...
python -m pip install --upgrade pip >nul
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
  echo   PyInstaller 설치 중...
  python -m pip install pyinstaller
)

echo.
echo [3/3] EXE 빌드 중... (몇 분 걸릴 수 있습니다)
REM  GUI 버전: 창만 뜨는 단일 실행파일. spv_analyzer.py 를 함께 포함.
python -m PyInstaller --onefile --windowed --clean --name "SPV분석기" ^
  --add-data "spv_analyzer.py;." ^
  spv_analyzer_gui.py

REM  명령줄(CLI) 버전도 함께 빌드 (선택)
python -m PyInstaller --onefile --console --clean --name "SPV분석기-CLI" ^
  spv_analyzer.py

echo.
if exist "dist\SPV분석기.exe" (
  echo  [완료] dist\SPV분석기.exe        (더블클릭해서 실행)
  echo  [완료] dist\SPV분석기-CLI.exe    (명령줄용)
) else (
  echo  [실패] 빌드 로그를 확인하세요.
)
echo.
pause
