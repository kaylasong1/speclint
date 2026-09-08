# speclint

화면정의서(.pptx)를 업로드하면 사람이 눈으로 찾기 힘든 불일치를 자동으로 찾아주는 검사 도구입니다.

## 검사 항목

| issue_type | 설명 |
| --- | --- |
| `broken_reference` | 다른 화면을 참조하는 코드가 있는데, 그 코드가 어느 화면의 정식 코드로도 정의되어 있지 않은 경우 |
| `duplicate_screen` | 같은 화면 코드와 화면명을 가진 장표가 2개 이상 있는 경우 (같은 화면을 실수로 두 번 정의했을 가능성) |
| `missing_screen_code` | 기능 정의로 보이는 페이지(설명 문구가 있고 텍스트가 충분히 긴 페이지)인데 화면 코드가 비어 있는 경우 |

## 로컬 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```

브라우저가 자동으로 열리지 않으면 터미널에 출력된 URL(기본값: http://localhost:8501)로 접속합니다.

화면에서 검사할 `.pptx` 파일을 업로드하고 [검사 실행] 버튼을 누르면 규칙 검사 결과가 표로 표시되고, CSV로 내려받을 수 있습니다.

### 커맨드라인으로 실행하기

웹 UI 없이 터미널에서만 검사하고 싶다면 아래 두 단계로 실행할 수 있습니다.

```bash
python extract_pptx.py <화면정의서.pptx> -o out/screens.json
python check_spec.py out/screens.json -o out/issues.csv
```
