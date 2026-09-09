"""
speclint 웹 화면 (Streamlit)

화면정의서(.pptx)를 업로드하면 기존 추출 로직(extract_pptx.py)과 규칙 검사
로직(check_spec.py)을 그대로 재사용해서 이슈를 찾아 화면에 보여준다.

화면코드 형식이 문서마다 다르므로(F.mMY.PP_MOBILE.1.0 / BO_M_0607 ...),
업로드 직후 detect_code_pattern.py로 후보를 뽑아 사용자가 고르게 한다.
기획자가 정규식을 직접 입력할 일은 없다.

사용법:
    streamlit run app.py
"""

import json
import re
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from pptx import Presentation

from check_spec import FIELDNAMES, run_all_checks
from detect_code_pattern import detect_from_screens
from extract_pptx import DEFAULT_SCREEN_CODE_PATTERN, extract_slide

st.set_page_config(page_title="speclint", layout="wide")

st.title("speclint - 화면정의서 검사")

DEFAULT_LABEL = "기본 패턴 (F.로 시작하는 코드)"


def run_extraction(pptx_path, code_pattern):
    """extract_pptx.py의 extract_slide를 재사용해 screens 리스트를 만든다."""
    code_re = re.compile(code_pattern)
    prs = Presentation(pptx_path)
    src = Path(pptx_path).name
    slide_height = prs.slide_height

    screens = []
    for i, slide in enumerate(prs.slides, start=1):
        screen = extract_slide(slide, i, slide_height, code_re)
        screen["source"]["file"] = src
        screens.append(screen)
    return screens


def analyze(pptx_bytes, filename, code_pattern):
    """추출 -> 후보 패턴 감지 -> 규칙 검사. (screens, issues, candidates) 반환."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir) / filename
        tmp_path.write_bytes(pptx_bytes)

        screens = run_extraction(str(tmp_path), code_pattern)

        screens_json_path = Path(tmpdir) / "screens.json"
        screens_json_path.write_text(
            json.dumps(screens, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    # 후보 감지는 blocks 원문만 보므로 어떤 코드 패턴으로 뽑았든 결과가 같다.
    candidates = detect_from_screens(screens, top=5)
    issues = run_all_checks(screens)
    return screens, issues, candidates


uploaded_file = st.file_uploader("화면정의서 .pptx 파일을 업로드하세요", type=["pptx"])

# 업로드된 파일이 바뀌면(새 파일 첨부, 파일 제거 등) 이전 검사 결과와 필터
# 선택을 초기화한다. file_id는 같은 파일을 다시 올려도 매번 새로 발급되므로
# "재첨부"도 새 파일로 취급된다.
current_file_id = uploaded_file.file_id if uploaded_file is not None else None
if current_file_id != st.session_state.get("last_file_id"):
    for key in ("screens", "issues", "candidates", "error", "pptx_bytes",
                "filename", "code_pattern", "computed_key"):
        st.session_state.pop(key, None)
    st.session_state["last_file_id"] = current_file_id

if uploaded_file is None:
    st.info("검사할 .pptx 파일을 업로드해주세요.")

run_clicked = st.button("검사 실행", disabled=uploaded_file is None)

if run_clicked and uploaded_file is not None:
    st.session_state["pptx_bytes"] = uploaded_file.getvalue()
    st.session_state["filename"] = uploaded_file.name
    st.session_state["code_pattern"] = DEFAULT_SCREEN_CODE_PATTERN
    st.session_state.pop("computed_key", None)

# 파일 또는 선택한 코드 패턴이 바뀌었으면 다시 계산한다.
if "pptx_bytes" in st.session_state:
    key = (current_file_id, st.session_state["code_pattern"])
    if st.session_state.get("computed_key") != key:
        st.session_state.pop("error", None)
        try:
            with st.spinner("검사 중입니다..."):
                screens, issues, candidates = analyze(
                    st.session_state["pptx_bytes"],
                    st.session_state["filename"],
                    st.session_state["code_pattern"],
                )
            st.session_state["screens"] = screens
            st.session_state["issues"] = issues
            st.session_state["candidates"] = candidates
            st.session_state["computed_key"] = key
        except Exception as e:
            for k in ("screens", "issues", "candidates"):
                st.session_state.pop(k, None)
            st.session_state["error"] = str(e)
            st.session_state["computed_key"] = key

if st.session_state.get("error"):
    st.error(f"파일 처리 중 오류가 발생했습니다: {st.session_state['error']}")


def render_pattern_picker(candidates, expanded):
    """감지된 후보 중에서 화면코드 패턴을 고르게 한다."""
    options = [DEFAULT_LABEL]
    pattern_by_label = {DEFAULT_LABEL: DEFAULT_SCREEN_CODE_PATTERN}
    for c in candidates:
        label = "%s… — %d개 슬라이드에서 %d종 (예: %s)" % (
            c["pattern"], c["slides"], c["values"], c["examples"][0]
        )
        options.append(label)
        pattern_by_label[label] = c["regex"]

    current = st.session_state["code_pattern"]
    current_index = 0
    for i, label in enumerate(options):
        if pattern_by_label[label] == current:
            current_index = i
            break

    with st.expander("화면코드 패턴", expanded=expanded):
        if not candidates:
            st.write("이 문서에서 코드처럼 보이는 패턴을 찾지 못했습니다.")
        choice = st.radio(
            "이 문서의 화면코드로 쓸 패턴을 골라주세요.",
            options=options,
            index=current_index,
            key=f"pattern_choice_{current_file_id}",
        )
        chosen = pattern_by_label[choice]
        if chosen != current:
            st.session_state["code_pattern"] = chosen
            st.rerun()
        st.caption(f"적용 중인 정규식: `{chosen}`")


tab_rule, tab_llm = st.tabs(["규칙 검사", "LLM 검토"])

with tab_rule:
    if "screens" not in st.session_state:
        st.write("파일을 업로드하고 [검사 실행] 버튼을 눌러주세요.")
    else:
        screens = st.session_state["screens"]
        issues = st.session_state["issues"]
        candidates = st.session_state.get("candidates", [])

        total_slides = len(screens)
        slides_with_code = sum(1 for s in screens if s.get("screen_code"))
        code_ratio = slides_with_code / total_slides if total_slides else 0

        no_code_warning = (
            "이 문서에서 화면코드를 찾지 못했습니다.\n"
            "아래 [화면코드 패턴]에서 이 문서에 맞는 형식을 골라주세요."
        )

        # 코드를 못 찾았거나 비율이 낮으면 패턴 선택을 펼친 상태로 먼저 보여준다.
        needs_attention = slides_with_code == 0 or code_ratio < 0.1
        if needs_attention:
            st.warning(no_code_warning)
        render_pattern_picker(candidates, expanded=needs_attention)

        if slides_with_code == 0:
            # 화면코드가 하나도 없으면 검사 결과 자체가 의미 없으므로 여기서 멈춘다.
            st.stop()

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
