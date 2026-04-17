from __future__ import annotations

from pathlib import Path

import streamlit as st

from document_tools import (
    analyze_reference_docx,
    build_docx_document,
    build_hwpx_document,
    parse_summary_document,
)


BASE_DIR = Path(__file__).resolve().parent
HWPX_TEMPLATE_DIR = BASE_DIR / "hwpx_template"


def _show_reference_analysis(analysis) -> None:
    st.subheader("2) 기준 문서 분석 결과")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("섹션 수", len(analysis.sections))
    col2.metric("표 개수", analysis.total_tables)
    col3.metric("사진 자리", analysis.total_images)
    col4.metric("기준 문서", analysis.source_name)

    st.write(f"대회명 추정: `{analysis.competition_title}`")
    st.write(
        "문서 크기: "
        f"{analysis.page_width_cm}cm x {analysis.page_height_cm}cm, "
        f"여백(상/하/좌/우): {analysis.margin_top_cm}/{analysis.margin_bottom_cm}/"
        f"{analysis.margin_left_cm}/{analysis.margin_right_cm}cm"
    )

    if analysis.cover_fields:
        st.write("표지 기본 정보 항목: " + ", ".join(analysis.cover_fields))

    rows = []
    for section in analysis.sections:
        rows.append(
            {
                "섹션": section.full_title,
                "문단 수": section.paragraph_count,
                "표 수": section.table_count,
                "사진 공란 수": section.image_slots,
            }
        )

    st.dataframe(rows, width="stretch", hide_index=True)


def _show_summary_guidance() -> None:
    with st.expander("요약서 작성 팁", expanded=False):
        st.markdown(
            """
요약서는 `TXT`, `MD`, `DOCX`로 올릴 수 있습니다. 아래처럼 주면 가장 잘 맞습니다.

```text
작품명: 스마트 냉각 쉼터
출품학생: 홍길동
지도교원: 김선생
출품분야: 생활과학
날짜: 2026. 3. 23.

발명 동기
무더운 날 대기 공간에서 열기가 심해...

문제 해결
초음파 센서와 팬 제어를 결합해...

발명품 제작
1차 모형은 ...
```

구조가 단순한 요약서여도 생성은 가능하지만, 섹션 제목을 같이 적어주면 기준 문서와 더 비슷하게 맞춰집니다.
"""
        )


def main() -> None:
    st.set_page_config(page_title="문서 제작기", layout="wide")
    st.title("문서 제작기")
    st.markdown(
        """
기준이 되는 `DOCX` 문서를 먼저 분석한 뒤, 요약서를 업로드하면 같은 흐름의 새 문서를 만듭니다.

- 기준 문서의 섹션 순서, 표 수, 사진 위치를 분석합니다.
- 요약서를 바탕으로 새 `DOCX` 문서를 생성합니다.
- 사진이 들어갈 자리는 `[사진 추가 예정]` 공란으로 남깁니다.
- 다운로드는 `DOCX`와 `HWPX`를 함께 제공합니다.
"""
    )

    _show_summary_guidance()

    col1, col2 = st.columns(2, gap="large")
    with col1:
        st.subheader("1) 파일 업로드")
        reference_file = st.file_uploader(
            "기준 문서 업로드 (DOCX)",
            type=["docx"],
            accept_multiple_files=False,
            key="reference_docx",
        )
        summary_file = st.file_uploader(
            "요약서 업로드 (TXT / MD / DOCX)",
            type=["txt", "md", "docx"],
            accept_multiple_files=False,
            key="summary_doc",
        )

    with col2:
        st.subheader("작업 메모")
        st.info(
            "현재 HWP 저장은 한글에서 바로 열 수 있는 `HWPX`로 제공합니다. "
            "`.hwp` 5.x 바이너리 저장은 별도 한글 프로그램 연동이 필요합니다."
        )

    if reference_file is None:
        st.stop()

    analysis = analyze_reference_docx(reference_file.getvalue(), reference_file.name)
    _show_reference_analysis(analysis)

    if summary_file is None:
        st.stop()

    summary = parse_summary_document(summary_file.name, summary_file.getvalue(), analysis)

    st.subheader("3) 요약서 해석 결과")
    preview_cols = st.columns(5)
    preview_cols[0].text_input("작품명", value=summary.project_title, disabled=True)
    preview_cols[1].text_input("대회명", value=summary.competition_title, disabled=True)
    preview_cols[2].text_input("출품학생", value=summary.student_name, disabled=True)
    preview_cols[3].text_input("지도교원", value=summary.advisor_name, disabled=True)
    preview_cols[4].text_input("출품분야", value=summary.category, disabled=True)

    section_preview_rows = []
    for section in analysis.sections:
        source_lines = summary.section_text.get(section.title, [])
        section_preview_rows.append(
            {
                "섹션": section.full_title,
                "직접 매핑된 줄 수": len(source_lines),
                "미리보기": " ".join(source_lines[:2])[:120],
            }
        )
    st.dataframe(section_preview_rows, width="stretch", hide_index=True)

    st.subheader("4) 문서 생성")
    generate = st.button("문서 만들기", type="primary", width="stretch")
    if not generate:
        st.stop()

    if not HWPX_TEMPLATE_DIR.exists():
        st.error("HWPX 템플릿 파일이 없습니다. `hwpx_template` 폴더를 확인해 주세요.")
        st.stop()

    with st.spinner("기준 문서 결을 반영해 새 문서를 만드는 중입니다..."):
        docx_bytes = build_docx_document(analysis, summary)
        hwpx_bytes = build_hwpx_document(analysis, summary, HWPX_TEMPLATE_DIR)

    st.success("문서 생성이 완료되었습니다.")

    output_name = summary.project_title.strip() or "generated_document"
    safe_name = output_name.replace("/", "_").replace("\\", "_")

    download_col1, download_col2 = st.columns(2)
    with download_col1:
        st.download_button(
            "DOCX 다운로드",
            data=docx_bytes,
            file_name=f"{safe_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            width="stretch",
        )
    with download_col2:
        st.download_button(
            "HWPX 다운로드",
            data=hwpx_bytes,
            file_name=f"{safe_name}.hwpx",
            mime="application/zip",
            width="stretch",
        )


if __name__ == "__main__":
    main()
