# 테트리스 (Tetris)

Python + pygame 으로 만든 테트리스 게임입니다.
**ESC 키를 누르면 프로그램이 즉시 종료됩니다.**

## 조작법

| 키 | 동작 |
|----|------|
| ← → | 좌우 이동 (누르고 있으면 연속 이동) |
| ↑ | 회전 |
| ↓ | 소프트 드롭 (누르고 있으면 계속 내려감) |
| Space | 하드 드롭 |
| R | 다시 시작 (게임 오버 시) |
| **ESC** | **프로그램 종료** |

## 바로 실행해보기 (소스 실행)

```bash
pip install -r requirements.txt
python tetris.py
```

## exe(실행파일) 만들기

`.exe`는 **Windows 실행파일**이라 반드시 **Windows 환경**에서 빌드해야 합니다.
방법은 두 가지입니다.

### 방법 1) 내 Windows PC에서 직접 빌드 (가장 간단)

1. [Python](https://www.python.org/downloads/) 설치 (설치 시 "Add Python to PATH" 체크)
2. 이 `tetris` 폴더에서 `build.bat` 더블클릭
3. 완료되면 `dist\Tetris.exe` 가 생성됩니다. 더블클릭하면 바로 실행돼요.

수동으로 하려면:

```bat
pip install -r requirements.txt
pyinstaller --onefile --windowed --name Tetris tetris.py
```

### 방법 2) GitHub Actions로 자동 빌드 (PC에 설치 없이)

이 저장소에 push 하면 `.github/workflows/build-exe.yml` 이 **Windows 러너에서
자동으로 exe를 빌드**합니다.

1. GitHub 저장소 → **Actions** 탭
2. **Build Tetris EXE** 워크플로 실행 결과 클릭
3. 하단 **Artifacts → Tetris-Windows-exe** 다운로드 → 압축 풀면 `Tetris.exe`

> `--windowed` 옵션 때문에 exe 실행 시 검은 콘솔창 없이 게임 창만 뜹니다.
