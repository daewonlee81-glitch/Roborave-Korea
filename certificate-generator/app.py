from __future__ import annotations

import io
import os
import tempfile
import zipfile
from dataclasses import asdict
from typing import Optional

import pandas as pd
import streamlit as st
from PIL import Image

from certificate_generator import TextSpec, _render_one, _safe_filename


def _get_uploaded_font_path(key: str, uploaded_file) -> Optional[str]:
    path_key = f"{key}_uploaded_font_path"
    name_key = f"{key}_uploaded_font_name"

    if path_key not in st.session_state:
        st.session_state[path_key] = None
    if name_key not in st.session_state:
        st.session_state[name_key] = None

    if uploaded_file is None:
        old = st.session_state.get(path_key)
        if old and os.path.exists(old):
            try:
                os.remove(old)
            except OSError:
                pass
        st.session_state[path_key] = None
        st.session_state[name_key] = None
        return None

    if st.session_state.get(name_key) != uploaded_file.name or not st.session_state.get(path_key):
        old = st.session_state.get(path_key)
        if old and os.path.exists(old):
            try:
                os.remove(old)
            except OSError:
                pass

        suffix = "." + uploaded_file.name.split(".")[-1].lower()
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_file.getvalue())
            st.session_state[path_key] = tmp.name
            st.session_state[name_key] = uploaded_file.name

    return st.session_state.get(path_key)


def _read_excel(data: bytes, sheet: str | int | None = None) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(data), sheet_name=sheet, dtype=str)
    if isinstance(df, dict):
        first_sheet = next(iter(df.keys()))
        return df[first_sheet]
    return df


def _coerce_rows(df: pd.DataFrame) -> list[dict[str, str]]:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.fillna("")
    return [{k: str(v).strip() for k, v in row.items()} for row in df.to_dict(orient="records")]


def _build_outputs(
    template_img: Image.Image,
    rows: list[dict[str, str]],
    specs: list[TextSpec],
    filename_field: str,
    export_png: bool,
    export_pdf: bool,
) -> tuple[Optional[bytes], Optional[bytes]]:
    png_buf = io.BytesIO() if export_png else None
    pdf_buf = io.BytesIO() if export_pdf else None
    rendered_rgb: list[Image.Image] = []
    zip_file = (
        zipfile.ZipFile(png_buf, mode="w", compression=zipfile.ZIP_DEFLATED)
        if png_buf is not None
        else None
    )

    try:
        for i, row in enumerate(rows, start=1):
            img = _render_one(template_img, row, specs)

            if export_pdf:
                rendered_rgb.append(img.convert("RGB"))

            if export_png and zip_file is not None:
                base = _safe_filename(row.get(filename_field, "")) or f"row_{i}"
                one = io.BytesIO()
                img.save(one, format="PNG")
                zip_file.writestr(f"{base}.png", one.getvalue())

        if export_pdf and pdf_buf is not None and rendered_rgb:
            first, rest = rendered_rgb[0], rendered_rgb[1:]
            first.save(pdf_buf, format="PDF", save_all=True, append_images=rest)
    finally:
        if zip_file is not None:
            zip_file.close()

    return (
        png_buf.getvalue() if png_buf is not None else None,
        pdf_buf.getvalue() if pdf_buf is not None else None,
    )


def _default_anchor_options() -> list[str]:
    return ["mm", "lm", "rm", "mt", "mb", "lt", "lb", "rt", "rb"]


def _style_editor(
    label: str,
    key_prefix: str,
    template_img: Image.Image,
    default_x: int,
    default_y: int,
    default_size: int,
    default_color: str,
    enabled: bool = True,
) -> TextSpec:
    st.markdown(f"**{label}**")
    x = st.number_input(
        f"{label} X",
        min_value=0,
        max_value=template_img.width,
        value=default_x,
        step=1,
        disabled=not enabled,
        key=f"{key_prefix}_x",
    )
    y = st.number_input(
        f"{label} Y",
        min_value=0,
        max_value=template_img.height,
        value=default_y,
        step=1,
        disabled=not enabled,
        key=f"{key_prefix}_y",
    )
    size = st.slider(
        f"{label} 글자 크기",
        min_value=10,
        max_value=300,
        value=default_size,
        step=1,
        disabled=not enabled,
        key=f"{key_prefix}_size",
    )
    color = st.color_picker(
        f"{label} 색상",
        value=default_color,
        disabled=not enabled,
        key=f"{key_prefix}_color",
    )
    anchor = st.selectbox(
        f"{label} 정렬(anchor)",
        options=_default_anchor_options(),
        index=0,
        disabled=not enabled,
        key=f"{key_prefix}_anchor",
    )
    font_file = st.file_uploader(
        f"{label} 폰트 업로드 (선택)",
        type=["ttf", "otf", "ttc"],
        accept_multiple_files=False,
        disabled=not enabled,
        help="비워두면 시스템 기본 폰트를 사용합니다.",
        key=f"{key_prefix}_font_file",
    )
    font_path = _get_uploaded_font_path(key_prefix, font_file)

    return TextSpec(
        field=key_prefix,
        x=int(x),
        y=int(y),
        size=int(size),
        color=str(color),
        anchor=str(anchor),
        font_path=font_path,
    )


st.set_page_config(page_title="상장/인증서 제작기", layout="wide")
st.title("상장/인증서 제작기")

st.markdown(
    """
상장 배경 이미지와 엑셀 명단을 업로드하면, **상장 제목 / 이름 / 학교 / 지도교사**를 한 번에 배치해서 저장할 수 있습니다.

- 상장 제목은 화면에서 직접 입력합니다.
- 이름, 학교, 지도교사는 엑셀 컬럼에서 불러옵니다.
- PNG ZIP, 합본 PDF, 또는 둘 다 저장할 수 있습니다.
"""
)

col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    st.subheader("1) 파일 업로드")
    template_file = st.file_uploader(
        "상장 배경 이미지 업로드 (PNG/JPG)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=False,
    )
    excel_file = st.file_uploader(
        "명단 엑셀 업로드 (XLSX/XLSM)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
    )

with col_right:
    st.subheader("2) 템플릿 미리보기")
    template_img: Optional[Image.Image] = None
    if template_file is not None:
        template_img = Image.open(template_file).convert("RGBA")
        st.image(
            template_img,
            caption=f"{template_file.name} ({template_img.width}x{template_img.height})",
            width="stretch",
        )
    else:
        st.info("왼쪽에서 상장 배경 이미지를 업로드하세요.")

if excel_file is None or template_img is None:
    st.stop()

df = _read_excel(excel_file.getvalue())
df.columns = [str(c).strip() for c in df.columns]
cols = list(df.columns)

if not cols:
    st.error("엑셀에서 컬럼을 찾지 못했습니다. 첫 줄(헤더)을 확인해 주세요.")
    st.stop()

st.subheader("3) 상장 제목과 컬럼 매핑")
title_text = st.text_input("상장 제목", value="상장")
title_value = title_text.strip()

map_col1, map_col2, map_col3, map_col4 = st.columns([1, 1, 1, 1], gap="large")
with map_col1:
    name_col = st.selectbox(
        "이름 컬럼",
        options=cols,
        index=cols.index("이름") if "이름" in cols else 0,
    )
with map_col2:
    include_school = st.checkbox("학교 출력", value=True)
    school_col = st.selectbox(
        "학교 컬럼",
        options=cols,
        index=cols.index("학교") if "학교" in cols else min(1, len(cols) - 1),
        disabled=not include_school,
    )
with map_col3:
    default_advisor_index = cols.index("지도교사") if "지도교사" in cols else min(2, len(cols) - 1)
    include_advisor = st.checkbox("지도교사 출력", value="지도교사" in cols)
    advisor_col = st.selectbox(
        "지도교사 컬럼",
        options=cols,
        index=default_advisor_index,
        disabled=not include_advisor,
    )
with map_col4:
    filename_col = st.selectbox(
        "파일명에 쓸 컬럼",
        options=cols,
        index=cols.index(name_col),
    )

rows_raw = _coerce_rows(df)
rows: list[dict[str, str]] = []
for row in rows_raw:
    name = row.get(name_col, "")
    school = row.get(school_col, "") if include_school else ""
    advisor = row.get(advisor_col, "") if include_advisor else ""
    filename = row.get(filename_col, "") or name

    if not any([title_value, name, school, advisor]):
        continue

    rows.append(
        {
            "title": title_value,
            "name": name,
            "school": school,
            "advisor": advisor,
            "filename": filename,
        }
    )

if not rows:
    st.warning("생성할 데이터가 없습니다. 엑셀 컬럼과 값이 비어 있지 않은지 확인해 주세요.")
    st.stop()

st.subheader("4) 텍스트 위치/스타일 설정")
st.caption("좌표는 템플릿 이미지의 왼쪽 위를 (0, 0)으로 하는 픽셀 기준입니다.")

style_col1, style_col2 = st.columns([1, 1], gap="large")
with style_col1:
    title_spec = _style_editor(
        label="상장 제목",
        key_prefix="title",
        template_img=template_img,
        default_x=int(template_img.width * 0.5),
        default_y=int(template_img.height * 0.24),
        default_size=84,
        default_color="#111111",
    )
    name_spec = _style_editor(
        label="이름",
        key_prefix="name",
        template_img=template_img,
        default_x=int(template_img.width * 0.5),
        default_y=int(template_img.height * 0.46),
        default_size=110,
        default_color="#111111",
    )

with style_col2:
    school_spec = _style_editor(
        label="학교",
        key_prefix="school",
        template_img=template_img,
        default_x=int(template_img.width * 0.5),
        default_y=int(template_img.height * 0.58),
        default_size=54,
        default_color="#333333",
        enabled=include_school,
    )
    advisor_spec = _style_editor(
        label="지도교사",
        key_prefix="advisor",
        template_img=template_img,
        default_x=int(template_img.width * 0.5),
        default_y=int(template_img.height * 0.68),
        default_size=46,
        default_color="#333333",
        enabled=include_advisor,
    )

specs = [title_spec, name_spec]
if include_school:
    specs.append(school_spec)
if include_advisor:
    specs.append(advisor_spec)

st.subheader("5) 미리보기")
preview_row = rows[0]
preview_img = _render_one(template_img, preview_row, specs)
st.image(
    preview_img,
    caption=(
        f"미리보기: 제목={preview_row.get('title', '')}, "
        f"이름={preview_row.get('name', '')}, "
        f"학교={preview_row.get('school', '')}, "
        f"지도교사={preview_row.get('advisor', '')}"
    ),
    width="stretch",
)

missing_fonts = [spec.field for spec in specs if spec.font_path is None]
if missing_fonts:
    st.info(
        "일부 항목은 업로드한 폰트 없이 기본 폰트로 미리보기/생성됩니다. "
        "한글 모양이나 크기 표현이 중요하면 해당 항목 폰트를 올려 주세요."
    )

st.subheader("6) 생성/다운로드")
st.write(f"총 {len(rows)}건 생성 예정")

export_option = st.radio(
    "저장 형식",
    options=["PNG ZIP", "PDF", "PDF + PNG ZIP"],
    horizontal=True,
)
export_png = export_option in {"PNG ZIP", "PDF + PNG ZIP"}
export_pdf = export_option in {"PDF", "PDF + PNG ZIP"}

generate = st.button("상장 생성하기", type="primary", width="stretch")
if generate:
    with st.spinner("상장을 생성하는 중입니다..."):
        png_bytes, pdf_bytes = _build_outputs(
            template_img=template_img,
            rows=rows,
            specs=specs,
            filename_field="filename",
            export_png=export_png,
            export_pdf=export_pdf,
        )

    st.success("생성이 완료되었습니다.")

    if export_png and png_bytes is not None:
        st.download_button(
            "PNG ZIP 다운로드",
            data=png_bytes,
            file_name="certificates_png.zip",
            mime="application/zip",
            width="stretch",
        )

    if export_pdf and pdf_bytes is not None:
        st.download_button(
            "합본 PDF 다운로드",
            data=pdf_bytes,
            file_name="certificates.pdf",
            mime="application/pdf",
            width="stretch",
        )

    st.subheader("현재 설정(JSON)")
    st.json(
        {
            "title_text": title_text,
            "title": asdict(title_spec),
            "name": asdict(name_spec),
            "school": asdict(school_spec) if include_school else None,
            "advisor": asdict(advisor_spec) if include_advisor else None,
            "excel_columns": {
                "name_col": name_col,
                "school_col": school_col if include_school else None,
                "advisor_col": advisor_col if include_advisor else None,
                "filename_col": filename_col,
            },
            "export_option": export_option,
        }
    )
