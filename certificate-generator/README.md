# Certificate Generator (상장/수료증 자동 제작)

템플릿 이미지(상장 배경) 위에 **엑셀/CSV 명단의 텍스트를 지정 좌표에 찍어서** 참가자별 PNG 또는 합본 PDF를 생성합니다.

## (추천) 업로드 방식 웹앱 (엑셀 업로드)

`app.py`는 **템플릿 이미지 + 엑셀(xlsx/xlsm)** 을 업로드해서 아래 작업을 할 수 있는 웹앱입니다.

- 상장 제목 직접 입력
- 이름 / 학교 / 지도교사 컬럼 매핑
- 제목 / 이름 / 학교 / 지도교사별 위치, 글꼴, 글자 크기, 색상 설정
- `PNG ZIP`, `합본 PDF`, `둘 다` 저장

### 실행

```bash
cd certificate-generator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

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

### 엑셀 컬럼

- 필수: `이름` (컬럼명은 앱에서 선택 가능)
- 선택: `학교`, `지도교사`
- 권장: 파일명에 쓸 컬럼을 `이름`으로 지정

> 템플릿은 **PNG/JPG**를 권장합니다. PDF 템플릿은 먼저 이미지로 변환해 주세요.

## 준비물

- Python 3.9+ 권장
- 상장 템플릿 이미지: `template.png` (또는 jpg)
- 명단 파일: 웹앱은 Excel(`.xlsx`, `.xlsm`), CLI는 CSV(UTF-8 권장)
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
