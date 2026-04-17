# Certificate Generator (상장/수료증 자동 제작)

템플릿 이미지(상장 배경) 위에 **엑셀/CSV 명단의 텍스트를 지정 좌표에 찍어서** 참가자별 PNG 또는 합본 PDF를 생성합니다.

## (추천) 참가 증명서 웹앱

`app.py`는 **템플릿 이미지 + 이름 직접 입력 또는 엑셀 불러오기** 방식으로 참가 증명서를 만드는 웹앱입니다.

- 학교 입력 없이 이름만 한글로 입력
- 여러 명이면 한 줄에 한 명씩 직접 입력하거나 엑셀에서 이름 컬럼 선택
- 이름 글자의 위치, 글꼴, 글자 크기, 색상 조정
- 실시간 미리보기
- `PNG ZIP`, `합본 PDF`, `둘 다` 저장

### 실행

```bash
cd certificate-generator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## 문서 제작용 웹앱 (기준 DOCX 분석 + 요약서 기반 생성)

`document_studio.py`는 기준이 되는 `DOCX` 문서를 먼저 분석한 뒤, 요약서를 업로드하면 같은 흐름의 새 문서를 만들어 주는 웹앱입니다.

- 기준 문서의 섹션 순서, 표 수, 사진 위치 분석
- 요약서(`txt`, `md`, `docx`) 업로드
- 사진이 들어갈 자리는 `[사진 추가 예정]` 공란으로 생성
- `DOCX`, `HWPX` 다운로드 지원

### 실행

```bash
cd certificate-generator
source .venv/bin/activate
pip install -r requirements.txt
streamlit run document_studio.py
```

### 요약서 예시

```text
작품명: 스마트 냉각 쉼터
출품학생: 홍길동
지도교원: 김선생
출품분야: 생활과학
날짜: 2026. 3. 23.

발명 동기
무더운 날 대기 공간에서 느낀 불편함이 시작점이었다.

문제 해결
센서와 팬 제어를 결합해 필요할 때만 작동하도록 설계했다.
```

> 참고: 이 앱은 한글에서 바로 열 수 있는 `HWPX`를 생성합니다. 전통적인 `.hwp` 5.x 바이너리 저장은 별도 한글 프로그램 연동이 필요합니다.

## 외부 접속용 배포

이 프로젝트는 이제 **Docker/Render 배포 가능 상태**입니다. 가장 쉬운 방법은 GitHub에 올린 뒤 Render로 배포하는 방식입니다.

### 방법 1: Render로 배포

프로젝트에 아래 파일이 포함되어 있습니다.

- `Dockerfile`
- `render.yaml`
- `.streamlit/config.toml`

배포 순서:

1. 이 폴더를 GitHub 저장소로 올립니다.
2. [Render](https://render.com)에서 `New Web Service`를 선택합니다.
3. GitHub 저장소를 연결합니다.
4. Render가 `render.yaml`과 `Dockerfile`을 읽어서 자동 배포합니다.
5. 배포가 끝나면 `https://...onrender.com` 주소로 외부 접속할 수 있습니다.

### 방법 2: Docker 서버에 직접 배포

```bash
docker build -t certificate-generator .
docker run -p 8501:8501 certificate-generator
```

서버에 올린 뒤 `http://서버IP:8501` 또는 연결한 도메인으로 접속하면 됩니다.

### 입력 방식

- 이름만 직접 입력하거나 엑셀에서 불러옵니다.
- 직접 입력은 여러 명을 줄바꿈으로 구분합니다.
- 엑셀 업로드 후 이름 컬럼을 선택합니다.
- 한글 폰트를 업로드하면 미리보기와 결과물을 더 정확하게 맞출 수 있습니다.

> 템플릿은 **PNG/JPG**를 권장합니다. PDF 템플릿은 먼저 이미지로 변환해 주세요.

## 준비물

- Python 3.9+ 권장
- 상장 템플릿 이미지: `template.png` (또는 jpg)
- 명단 파일: 웹앱은 이름 직접 입력 또는 Excel(`.xlsx`, `.xlsm`), CLI는 CSV(UTF-8 권장)
- (권장) 한글 폰트 파일: `.ttf` 또는 `.otf`

## 설치

```bash
cd certificate-generator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 빠른 시작

1) 샘플 CSV 확인: `recipients.sample.csv`

2) 실행(기본 좌표/기본 필드 사용):

```bash
python certificate_generator.py \
  --template template.png \
  --csv recipients.sample.csv \
  --out out \
  --pdf
```

출력:

- `out/홍길동.png` 같은 개별 PNG
- `out/certificates.pdf` 합본 PDF

## 좌표/폰트/필드 커스터마이즈

텍스트 한 줄 배치 규칙은 `--text FIELD X Y SIZE COLOR ANCHOR` 입니다. 여러 번 반복할 수 있어요.

예시:

```bash
python certificate_generator.py \
  --template template.png \
  --csv recipients.sample.csv \
  --out out \
  --font "/System/Library/Fonts/Supplemental/AppleGothic.ttf" \
  --text award_title 800 430 80 "#111111" mm \
  --text name       800 600 110 "#111111" mm \
  --text org        800 780 44 "#333333" mm \
  --text date       800 850 36 "#333333" mm \
  --pdf
```

- `X, Y`: 템플릿 이미지 **왼쪽 위 기준 픽셀 좌표**
- `ANCHOR`: `mm`(가운데), `lm`(왼쪽-가운데), `rm`(오른쪽-가운데) 등 Pillow 앵커

## CSV 컬럼

기본 예시는 아래 컬럼을 기대합니다.

- `name`
- `award_title`
- `date`
- `org`

다른 컬럼명을 쓰면 `--text`의 `FIELD`를 그 컬럼명으로 맞추면 됩니다.

## 팁(좌표 찾는 방법)

- 템플릿 이미지를 미리보기(Preview)로 열고, 대략 위치를 잡은 뒤 `--text`로 조금씩 조정하는 게 가장 빠릅니다.
- 한글이 네모(□)로 나오면 `--font`로 한글 지원 폰트를 지정하세요.
