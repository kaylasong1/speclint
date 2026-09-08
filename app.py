"""
speclint 웹 화면 (Streamlit)

화면정의서(.pptx)를 업로드하면 기존 추출 로직(extract_pptx.py)과 규칙 검사
로직(check_spec.py)을 그대로 재사용해서 이슈를 찾아 화면에 보여준다.

사용법:
    streamlit run app.py
"""

import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from pptx import Presentation

from check_spec import FIELDNAMES, run_all_checks
from extract_pptx import extract_slide

st.set_page_config(page_title="speclint", layout="wide")

st.title("speclint - 화면정의서 검사")


def run_extraction(pptx_path):
    """extract_pptx.py의 extract_slide를 재사용해 screens 리스트를 만든다."""
    prs = Presentation(pptx_path)
    src = Path(pptx_path).name
    slide_height = prs.slide_height

    screens = []
    for i, slide in enumerate(prs.slides, start=1):
        screen = extract_slide(slide, i, slide_height)
        screen["source"]["file"] = src
        screens.append(screen)
    return screens


uploaded_file = st.file_uploader("화면정의서 .pptx 파일을 업로드하세요", type=["pptx"])

# 업로드된 파일이 바뀌면(새 파일 첨부, 파일 제거 등) 이전 검사 결과와 필터
# 선택을 초기화한다. file_id는 같은 파일을 다시 올려도 매번 새로 발급되므로
# "재첨부"도 새 파일로 취급된다.
current_file_id = uploaded_file.file_id if uploaded_file is not None else None
if current_file_id != st.session_state.get("last_file_id"):
    st.session_state.pop("screens", None)
    st.session_state.pop("issues", None)
    st.session_state.pop("error", None)
    st.session_state["last_file_id"] = current_file_id

if uploaded_file is None:
    st.info("검사할 .pptx 파일을 업로드해주세요.")

run_clicked = st.button("검사 실행", disabled=uploaded_file is None)

if run_clicked and uploaded_file is not None:
    st.session_state.pop("error", None)
    try:
        with st.spinner("검사 중입니다..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_path = Path(tmpdir) / uploaded_file.name
                tmp_path.write_bytes(uploaded_file.getvalue())

                screens = run_extraction(str(tmp_path))

                screens_json_path = Path(tmpdir) / "screens.json"
                screens_json_path.write_text(
                    json.dumps(screens, ensure_ascii=False, indent=2), encoding="utf-8"
                )

                issues = run_all_checks(screens)

        st.session_state["screens"] = screens
        st.session_state["issues"] = issues
    except Exception as e:
        st.session_state.pop("screens", None)
        st.session_state.pop("issues", None)
        st.session_state["error"] = str(e)

if st.session_state.get("error"):
    st.error(f"파일 처리 중 오류가 발생했습니다: {st.session_state['error']}")

tab_rule, tab_llm = st.tabs(["규칙 검사", "LLM 검토"])

with tab_rule:
    if "screens" not in st.session_state:
        st.write("파일을 업로드하고 [검사 실행] 버튼을 눌러주세요.")
    else:
        screens = st.session_state["screens"]
        issues = st.session_state["issues"]

        total_slides = len(screens)
        slides_with_code = sum(1 for s in screens if s.get("screen_code"))
        code_ratio = slides_with_code / total_slides if total_slides else 0

        no_code_warning = (
            "이 문서에서 화면코드를 찾지 못했습니다.\n"
            "화면코드 형식이 지원 패턴과 다를 수 있습니다."
        )

        if slides_with_code == 0:
            # 화면코드가 하나도 없으면 검사 결과 자체가 의미 없으므로 경고만 보여준다.
            st.warning(no_code_warning)
        else:
            # 검출은 됐지만 비율이 낮으면(10% 미만) 놓친 화면코드가 있을 수 있다는
            # 경고를 결과와 함께 보여준다(결과 자체는 숨기지 않는다).
            if code_ratio < 0.1:
                st.warning(no_code_warning)

            issue_count = len(issues)

            col1, col2, col3 = st.columns(3)
            col1.metric("전체 슬라이드 수", total_slides)
            col2.metric("화면코드가 있는 슬라이드 수", slides_with_code)
            col3.metric("검출 이슈 건수", issue_count)

            if issue_count == 0:
                st.success("검출된 이슈가 없습니다")
            else:
                df = pd.DataFrame(issues, columns=FIELDNAMES)

                issue_types = sorted(df["issue_type"].unique().tolist())
                # key를 업로드된 파일에 묶어서, 새 파일이 오면 필터 위젯도 새로
                # 시작하도록 한다(이전 파일에서 고른 선택이 남아있지 않게).
                selected_types = st.multiselect(
                    "검사 항목 필터",
                    options=issue_types,
                    key=f"issue_type_filter_{current_file_id}",
                    placeholder="전체 표시 중 (항목을 선택하면 필터링됩니다)",
                )
                # 선택한 항목이 없으면 "필터 없음 = 전체 보기"로 취급한다.
                filtered_df = df if not selected_types else df[df["issue_type"].isin(selected_types)]

                st.dataframe(filtered_df, use_container_width=True)

                csv_bytes = filtered_df.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "CSV 다운로드",
                    data=csv_bytes,
                    file_name="issues.csv",
                    mime="text/csv",
                )

with tab_llm:
    st.info("준비 중입니다")
