# 무한 피하기 게임 (Dodge)

Python + pygame 으로 만든 무한 피하기 게임입니다.
위에서 떨어지는 빨간 상자를 피하세요! **ESC 키를 누르면 프로그램이 즉시 종료됩니다.**

## 조작법

| 키 | 동작 |
|----|------|
| ← → (또는 A / D) | 좌우 이동 |
| Space / Enter | 시작 · 다시 시작 |
| **ESC** | **프로그램 종료** |

## 바로 실행해보기 (소스 실행)

```bash
pip install -r requirements.txt
python game.py
```

## exe(실행파일) 만들기

### 방법 1) 내 Windows PC에서 직접 빌드
1. [Python](https://www.python.org/downloads/) 설치 (설치 시 "Add Python to PATH" 체크)
2. 이 `dodge` 폴더에서 `build.bat` 더블클릭
3. 완료되면 `dist\Dodge.exe` 가 생성됩니다.

수동으로 하려면:
```bat
pip install -r requirements.txt
pyinstaller --onefile --windowed --name Dodge game.py
```

### 방법 2) GitHub Actions로 자동 빌드 (PC에 설치 없이)
이 저장소에 push 하면 `.github/workflows/build-dodge.yml` 이 자동으로 빌드합니다.

1. GitHub 저장소 → **Actions** 탭 → **Build Dodge Game** 실행 결과
2. 하단 **Artifacts** 에서 다운로드:
   - **Dodge-Windows-exe** → 압축 풀면 `Dodge.exe`
   - **Dodge-macOS-AppleSilicon** → Apple Silicon 맥용 `Dodge.app`

> 맥 앱은 서명이 없어 처음엔 **우클릭 → 열기**로 실행해야 합니다.
