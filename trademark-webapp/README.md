# 상표 OA 분석기 (웹앱)

특허청 상표 의견제출통지서(OA) ZIP 파일을 브라우저에 업로드하면
출원상표/선행상표 이미지 + 호칭(AI 분석) + OA 요약이 정리된 엑셀을 생성해줍니다.

내 컴퓨터에서만 돌아가는 로컬 서버입니다. 파일과 API 키는 Anthropic API 호출을
제외하고는 외부로 전송되지 않습니다.

---

## 1. 준비물

### Python 3.10 이상
- Mac: 터미널에서 `python3 --version` 으로 확인. 없으면 [python.org](https://www.python.org/downloads/) 에서 설치.
- Windows: [python.org](https://www.python.org/downloads/) 에서 설치 시 **"Add python.exe to PATH"** 체크 필수.

### poppler (PDF 텍스트 추출용, `pdftotext` 명령이 필요합니다)
- Mac (Homebrew 필요): `brew install poppler`
- Windows: https://github.com/oschwartz10612/poppler-windows/releases 에서 최신 zip
  다운로드 → 압축 해제 → `bin` 폴더 경로를 환경변수 PATH에 추가
- Ubuntu/Debian: `sudo apt install poppler-utils`

## 2. Anthropic API 키 발급 방법

상표 이미지의 한국어/영어 호칭을 AI로 분석하는 데 사용됩니다. 지금은 API 키가
없으시다고 하셨으니 아래 순서대로 발급받으시면 됩니다.

1. https://console.anthropic.com 접속 후 회원가입/로그인
2. 좌측 메뉴에서 **"Settings" → "API Keys"** 이동 (또는 바로 https://console.anthropic.com/settings/keys)
3. **"Create Key"** 클릭 → 이름 지정 (예: `trademark-oa-tool`) → 생성
4. 생성된 키(`sk-ant-...`로 시작)를 복사해둡니다. **다시 볼 수 없으니 안전한 곳에 저장하세요.**
5. 결제 수단이 등록되어 있지 않다면 **"Billing"** 메뉴에서 카드 등록 및 크레딧 충전이 필요합니다.
   (사용한 만큼만 과금되는 종량제이며, 상표 이미지 1건 분석에 드는 비용은 매우 적습니다.)

이 키는 웹앱 화면에서 직접 입력하며, 서버 파일에 저장되지 않습니다
(입력할 때마다 새로 넣어야 합니다. 매번 입력이 번거로우시면 아래 "선택: API 키 자동 입력"
항목을 참고하세요).

## 3. 설치

터미널(Mac) 또는 명령 프롬프트(Windows)에서 이 폴더로 이동한 뒤:

```bash
pip install -r requirements.txt
```

## 4. 실행

```bash
python app.py
```

터미널에 `Running on http://127.0.0.1:5000` 이 뜨면, 브라우저에서 아래 주소로 접속하세요.

```
http://localhost:5000
```

종료하려면 터미널 창에서 `Ctrl + C`.

## 5. 사용법

1. 화면에 Anthropic API 키 입력
2. 특허청에서 다운로드한 OA ZIP 파일을 업로드 (여러 건의 OA PDF가 들어있는 zip)
3. 처리 건수는 기본 `0`(전체) — 테스트해보고 싶으면 `3` 정도로 줄여서 먼저 확인해보세요
4. "엑셀 생성 시작" 클릭 → 진행률 표시 → 완료되면 다운로드 버튼 클릭

## 6. 참고 사항

- 처리 속도는 OA 건수와 이미지 분석(AI 호출) 수에 비례합니다. 건수가 많으면
  시간이 걸릴 수 있으니, 먼저 소량(`count=3~5`)으로 테스트해보시길 권합니다.
- 신형(2023년~)/구형(~2022년) PDF 레이아웃을 모두 처리하도록 되어 있지만,
  특허청 서식이 바뀌면 이미지/텍스트 추출 정규식을 다시 조정해야 할 수 있습니다.
- 이 도구는 로컬 1인 사용을 기준으로 만들어졌습니다. 사무실 내 여러 명이
  동시에 쓰시려면 서버 배포(별도 안내 가능) 및 작업 큐 개선이 필요합니다.

## 7. 선택: API 키 자동 입력 (매번 입력이 번거로울 경우)

프로젝트 폴더에 `.env` 파일을 만들고 아래처럼 적어두면, 웹 화면의 API 키
입력란에 자동으로 채워지도록 코드를 조금만 손보면 됩니다. 필요하시면 말씀해주세요.

```
ANTHROPIC_API_KEY=sk-ant-여기에-키-입력
```

## 폴더 구조

```
trademark-webapp/
├── app.py              # Flask 서버 (업로드/진행률/다운로드 라우트)
├── core.py             # PDF 파싱, 이미지 추출, AI 호칭 분석, 엑셀 생성 로직
├── requirements.txt
├── templates/
│   └── index.html      # 업로드 화면
└── work/                # 처리 중 임시 파일 (자동 생성/정리)
```
