# Google Sheets Editor

Google Sheets 링크를 URL에 붙여넣으면 누구나 수정할 수 있는 웹 편집기입니다.  
서비스 계정(Service Account) 방식으로 개인 로그인 없이 동작합니다.

## 설정 (최초 1회)

### 1. Google Cloud 설정
1. [Google Cloud Console](https://console.cloud.google.com) 접속
2. **APIs & Services → 라이브러리** → `Google Sheets API` 활성화
3. **APIs & Services → 사용자 인증 정보 → 서비스 계정 만들기**
4. 서비스 계정 생성 후 → **키 추가 → JSON** 다운로드
5. 다운로드한 파일 이름을 `credentials.json`으로 바꿔서 이 폴더에 복사

### 2. 스프레드시트 공유
- Google Sheets 열기 → **공유** 클릭
- `credentials.json` 안의 `client_email` 값을 복사해서 **편집자** 권한으로 공유

### 3. 서버 실행
```bash
npm install
npm start
```

브라우저에서 `http://localhost:3000` 접속 후  
Google Sheets URL을 붙여넣으면 바로 편집 가능합니다.

## 다른 사람과 공유

서버를 실행한 뒤 IP/도메인 주소를 공유하면  
**링크를 아는 누구나 로그인 없이 편집** 할 수 있습니다.

```
http://서버IP:3000
```

## 사용법

| 동작 | 방법 |
|------|------|
| 시트 열기 | URL 붙여넣고 **열기** 클릭 |
| 셀 편집 | 셀 클릭 후 바로 입력 (노란색 = 미저장) |
| 저장 | **저장** 버튼 또는 `Ctrl+S` |
| 시트 전환 | 상단 드롭다운에서 탭 선택 |
| 행 추가 | 테이블 하단 **+ 행 추가** 클릭 |
