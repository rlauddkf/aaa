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

이 저장소에 push 하면 `.github/workflows/build-exe.yml` 이 **Windows·macOS 러너에서
자동으로 실행파일을 빌드**합니다.

1. GitHub 저장소 → **Actions** 탭
2. **Build Tetris EXE** 워크플로 실행 결과 클릭
3. 하단 **Artifacts** 에서 원하는 OS용 파일 다운로드:
   - **Tetris-Windows-exe** → 압축 풀면 `Tetris.exe`
   - **Tetris-macOS-AppleSilicon** → Apple Silicon(M1/M2/M3…) 맥용
   - **Tetris-macOS-Intel** → Intel 맥용

> `--windowed` 옵션 때문에 실행 시 콘솔창 없이 게임 창만 뜹니다.

## 맥(macOS)에서 실행하기

### 가장 간단 — 소스로 바로 실행

```bash
pip3 install pygame
python3 tetris.py
```

### 맥용 앱(.app)으로 실행

위 **방법 2**의 Actions Artifacts에서 본인 맥에 맞는 파일을 받습니다
(Apple Silicon / Intel 구분). 압축을 풀면 `Tetris.app` 이 나옵니다.

> ⚠️ 서명되지 않은 앱이라 처음 실행 시 "확인되지 않은 개발자" 경고가 뜹니다.
> **`Tetris.app` 우클릭 → 열기 → (다시) 열기** 를 누르면 실행됩니다.
> 또는 *시스템 설정 → 개인정보 보호 및 보안* 에서 "확인 없이 열기"를 허용하세요.

### 맥에서 직접 .app 빌드하기

```bash
pip3 install pygame pyinstaller
pyinstaller --onefile --windowed --name Tetris tetris.py
# dist/Tetris.app 생성
```

> 맥용 앱은 반드시 **맥에서** 빌드해야 합니다 (Windows/Linux에서는 못 만듭니다).
