# RND Dashboard

`projects` 폴더에 있는 과제별 성과지표 엑셀 파일을 읽어 시각화하는 Streamlit 앱입니다.

## 기능

- 프로젝트별 총 사업비/정량지표 달성 현황 요약
- 파일/시트 선택 기반 데이터 탐색
- 헤더 행 지정(병합 셀 형태 엑셀 대응)
- 라인/막대/산점도 차트 생성

## 실행 방법

1. 의존성 설치

```bash
py -m pip install -r requirements.txt
```

2. 앱 실행

```bash
py -m streamlit run app.py
```

3. 브라우저에서 표시되는 주소로 접속

## HTML로 빠르게 확인하기

단일 파일 대시보드인 `dashboard.html`도 함께 제공합니다.

### 방법 1) 서버 없이 바로 확인

1. `dashboard.html`을 브라우저로 엽니다.
2. `파일 읽기` 전에 `projects` 폴더의 엑셀 파일을 선택해서 업로드합니다.

### 방법 2) 샘플 자동 불러오기까지 사용

```bash
py -m http.server 8080
```

브라우저에서 `http://localhost:8080/dashboard.html` 접속

## 파일 변경 시 자동 새로고침 + Git 자동 커밋/푸시

### 1) 선행 조건 (최초 1회)

`auto_sync.py`는 Git 저장소/원격이 준비되어 있어야 push까지 수행합니다.

```bash
git init
git remote add origin <원격저장소URL>
git branch -M main
```

### 2) 대시보드 서버 실행

```bash
py -m http.server 8080
```

### 3) 감시 스크립트 실행 (별도 터미널)

```bash
py auto_sync.py
```

동작 방식:

- `C:/Copilot/RND-Dashboard` 하위 파일 추가/수정/삭제를 2초 간격으로 감지
- 감지 시 `.watch/reload-token.json` 갱신 → 열려 있는 `dashboard.html` 자동 새로고침
- 같은 시점에 `git add -A` → `git commit` → `git push origin <현재브랜치>` 자동 실행

## 데이터 위치

- 엑셀 파일 폴더: `projects/`
- 지원 확장자: `.xlsx`, `.xls`, `.xlsm`

## 참고

- 일부 시트는 머리글이 2행 이상이거나 병합되어 있어 자동 인식이 완벽하지 않을 수 있습니다.
- 이 경우 앱의 `헤더 행 번호`를 조정해서 원하는 컬럼 구조로 맞출 수 있습니다.
