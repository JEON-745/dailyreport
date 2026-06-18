# 광고 리포트 자동화 웹앱

브랜드를 선택하고 파일을 올리면, 리포트가 자동으로 채워지는 내부용 웹앱입니다.
차트·서식·썸네일·메모를 보존한 채 값만 채웁니다.

## 폴더 구성
- `app.py` — 웹앱 화면
- `build_report.py` — 매체(통합보고서 → 리포트) 자동입력 엔진
- `sns_auto.py` — SNS 소재별 효율 + 메모 엔진
- `requirements.txt` — 필요한 파이썬 패키지
- `.streamlit/secrets.toml` — 접속 비밀번호 (배포 시 별도 등록)

## 로컬에서 먼저 실행해보기
1. 파이썬 설치 (3.10+)
2. 터미널에서 이 폴더로 이동 후:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. 브라우저가 열리면 `secrets.toml`의 비밀번호로 접속

## 외부에서 접속 가능하게 배포 (Streamlit Community Cloud — 무료)
1. 이 폴더를 GitHub 저장소(private 권장)에 올립니다.
   - ⚠️ `.streamlit/secrets.toml`은 올리지 마세요(비밀번호 노출). `.gitignore`에 추가.
2. https://share.streamlit.io 접속 → GitHub 연결 → 이 저장소의 `app.py` 선택.
3. 앱 설정의 **Secrets**에 아래를 등록:
   ```
   APP_PASSWORD = "원하는_비밀번호"
   ```
4. 배포되면 `https://<이름>.streamlit.app` 주소가 생기고, 외부 누구나 이 URL + 비밀번호로 접속 가능합니다.

## 다른 브랜드 추가
`app.py`의 `PROFILES`에 새 항목을 추가하고, 그 브랜드용 처리 함수를 만들면 됩니다.
방식은 같으니, 해당 브랜드의 리포트 양식 + 소스 파일만 있으면 프로필을 추가할 수 있습니다.

## 보안 메모
광고주 데이터를 다루므로, 민감하면 Streamlit Cloud 대신 **사내 서버**나 접근이 제한된 환경에 배포하는 것을 권장합니다.
