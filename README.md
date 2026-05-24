# Material ERP Dashboard

구매/자재 Excel 데이터를 업로드해 자재 부족, PO별 필요 자재, 날짜별 입출고 이벤트, 공급업체별 액션 리스트, 추가 물량 리스크를 확인하는 Streamlit 대시보드입니다.

## GitHub 업로드 파일 구성

```text
material-erp-dashboard-streamlit/
├─ app.py
├─ requirements.txt
├─ README.md
├─ .gitignore
└─ .streamlit/
   └─ config.toml
```

## 로컬 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud 연결 방법

1. GitHub에서 새 Repository를 생성합니다.
2. 이 폴더 안의 파일들을 그대로 업로드합니다.
3. Streamlit Cloud에서 `New app`을 누릅니다.
4. Repository를 선택합니다.
5. Main file path에 아래처럼 입력합니다.

```text
app.py
```

6. Deploy를 누르면 실행됩니다.

## 필요한 Excel 시트

앱은 업로드된 Excel 파일에서 다음 성격의 시트를 자동 탐색해 사용합니다.

- 구매데이터 또는 자재품목리스트
- 생산계획
- 입고 내역
- 출고 내역
- 운송요율: 선택 사항

운송요율 시트가 없으면 운송 리스크 계산은 비활성화되고, 자재/생산 리스크 중심으로 작동합니다.

## 배포 시 주의사항

- 파일명은 반드시 `app.py`로 유지하세요.
- `requirements.txt`를 함께 올려야 Streamlit Cloud에서 필요한 라이브러리가 설치됩니다.
- Excel 파일은 앱 화면의 사이드바에서 업로드하는 방식입니다. GitHub에 민감한 실제 데이터 파일을 올리지 않는 것을 권장합니다.
