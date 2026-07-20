@echo off
REM Windows 에서 직접 EXE 를 빌드하려면 이 파일을 더블클릭하거나 실행하세요.
REM (Python 3.8+ 필요)

python -m pip install --upgrade pip
pip install -r requirements-pptx.txt
pip install pyinstaller

pyinstaller --onefile --windowed ^
  --name PPTImageReplacer ^
  --hidden-import win32timezone ^
  --hidden-import win32com.client ^
  pptx_image_replacer.py

echo.
echo 빌드 완료: dist\PPTImageReplacer.exe
pause
