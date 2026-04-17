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

from certificate_generator import TextSpec, _render_one, _safe_filename, _unique_filename_base


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


def _default_anchor_options() -> list[str]:
    return ["mm", "lm", "rm", "mt", "mb", "lt", "lb", "rt", "rb"]


def _read_excel(data: bytes, sheet: str | int | None = None) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(data), sheet_name=sheet, dtype=str)
    if isinstance(df, dict):
        first_sheet = next(iter(df.keys()))
        return df[first_sheet]
    return df


def _rows_from_dataframe(df: pd.DataFrame, name_col: str) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    values = df.fillna("")[name_col].tolist()
    for i, value in enumerate(values, start=1):
        name = str(value).strip()
        if not name:
            continue
        rows.append(
            {
                "name": name,
                "filename": _safe_filename(name) or f"name_{i}",
            }
        )
    return rows


def _parse_names(raw_text: str) -> list[dict[str, str]]:
    names = [line.strip() for line in raw_text.splitlines() if line.strip()]
    rows: list[dict[str, str]] = []
    for i, name in enumerate(names, start=1):
        rows.append(
            {
                "name": name,
                "filename": _safe_filename(name) or f"name_{i}",
            }
        )
    return rows


def _duplicate_name_count(rows: list[dict[str, str]]) -> int:
    counts: dict[str, int] = {}
    duplicates = 0
    for row in rows:
        name = row.get("name", "")
        counts[name] = counts.get(name, 0) + 1
        if counts[name] > 1:
            duplicates += 1
    return duplicates


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
    used_filenames: dict[str, int] = {}
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
                base = _unique_filename_base(row.get(filename_field, ""), used_filenames, f"row_{i}")
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


def _style_editor(
    label: str,
    key_prefix: str,
    template_img: Image.Image,
    default_x: int,
    default_y: int,
    default_size: int,
    default_color: str,
) -> TextSpec:
    x = st.number_input(
        f"{label} X",
        min_value=0,
        max_value=template_img.width,
        value=default_x,
        step=1,
        key=f"{key_prefix}_x",
    )
    y = st.number_input(
        f"{label} Y",
        min_value=0,
        max_value=template_img.height,
        value=default_y,
        step=1,
        key=f"{key_prefix}_y",
    )
    size = st.slider(
        f"{label} 글자 크기",
        min_value=10,
        max_value=300,
        value=default_size,
        step=1,
        key=f"{key_prefix}_size",
    )
    color = st.color_picker(
        f"{label} 색상",
        value=default_color,
        key=f"{key_prefix}_color",
    )
    anchor = st.selectbox(
        f"{label} 정렬(anchor)",
        options=_default_anchor_options(),
        index=0,
        key=f"{key_prefix}_anchor",
    )
    font_file = st.file_uploader(
        f"{label} 폰트 업로드 (선택)",
        type=["ttf", "otf", "ttc"],
        accept_multiple_files=False,
        help="한글 미리보기가 중요하면 사용하려는 폰트를 업로드하세요.",
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


st.set_page_config(page_title="참가 증명서 제작기", layout="wide")
st.title("참가 증명서 제작기")

st.markdown(
    """
참가 증명서 배경 이미지 위에 **이름만 한글로 직접 입력하거나 엑셀에서 불러와서** 바로 미리보고 저장할 수 있습니다.

- 학교 입력은 제외했습니다.
- 여러 명이면 이름을 한 줄에 한 명씩 입력하거나, 엑셀에서 이름 컬럼만 불러오면 됩니다.
- 이름 글자의 폰트, 크기, 색상, 위치를 조정할 수 있습니다.
- PNG ZIP, 합본 PDF, 또는 둘 다 저장할 수 있습니다.
"""
)

col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    st.subheader("1) 배경 파일 업로드")
    template_file = st.file_uploader(
        "참가 증명서 배경 이미지 업로드 (PNG/JPG)",
        type=["png", "jpg", "jpeg"],
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
        st.info("왼쪽에서 참가 증명서 배경 이미지를 업로드하세요.")

if template_img is None:
    st.stop()

st.subheader("3) 이름 입력 방식")
input_method = st.radio(
    "이름 불러오기 방식",
    options=["직접 입력", "엑셀 불러오기"],
    horizontal=True,
)

if input_method == "직접 입력":
    name_text = st.text_area(
        "이름 입력 (한 줄에 한 명씩, 한글 권장)",
        value="홍길동",
        height=180,
        help="예시:\n홍길동\n김민수\n박서연",
    )
    rows = _parse_names(name_text)
else:
    excel_file = st.file_uploader(
        "이름 명단 엑셀 업로드 (XLSX/XLSM)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        help="이름이 들어 있는 열 하나만 있어도 됩니다.",
    )
    if excel_file is None:
        st.info("이름 목록이 들어 있는 엑셀 파일을 업로드해 주세요.")
        st.stop()

    df = _read_excel(excel_file.getvalue())
    df.columns = [str(col).strip() for col in df.columns]
    cols = list(df.columns)

    if not cols:
        st.error("엑셀에서 컬럼을 찾지 못했습니다. 첫 줄(헤더)을 확인해 주세요.")
        st.stop()

    default_index = cols.index("이름") if "이름" in cols else 0
    name_col = st.selectbox(
        "이름 컬럼 선택",
        options=cols,
        index=default_index,
    )
    st.caption(f"불러온 컬럼: {name_col} / 행 수: {len(df)}")
    rows = _rows_from_dataframe(df, name_col)

if not rows:
    st.warning("최소 한 명 이상의 이름을 입력하거나 엑셀에서 불러와 주세요.")
    st.stop()

duplicate_count = _duplicate_name_count(rows)
if duplicate_count:
    st.info(
        f"동명이인 또는 중복 이름 {duplicate_count}건이 있습니다. "
        "파일 저장 시 자동으로 `_2`, `_3` 같은 번호를 붙여 모두 생성합니다."
    )

st.subheader("4) 이름 글자 위치/스타일 설정")
st.caption("좌표는 배경 이미지의 왼쪽 위를 (0, 0)으로 하는 픽셀 기준입니다.")

name_spec = _style_editor(
    label="이름",
    key_prefix="name",
    template_img=template_img,
    default_x=int(template_img.width * 0.5),
    default_y=int(template_img.height * 0.55),
    default_size=96,
    default_color="#111111",
)
specs = [name_spec]

st.subheader("5) 미리보기")
preview_index = st.selectbox(
    "미리볼 이름",
    options=list(range(len(rows))),
    index=0,
    format_func=lambda idx: rows[idx]["name"],
)
preview_row = rows[preview_index]
preview_img = _render_one(template_img, preview_row, specs)
st.image(
    preview_img,
    caption=f"미리보기: 이름={preview_row['name']}",
    width="stretch",
)

if name_spec.font_path is None:
    st.info("한글 폰트를 맞춰 보고 싶다면 이름 폰트를 업로드해 주세요.")

st.subheader("6) 생성/다운로드")
st.write(f"총 {len(rows)}명 생성 예정")

export_option = st.radio(
    "저장 형식",
    options=["PNG ZIP", "PDF", "PDF + PNG ZIP"],
    horizontal=True,
)
export_png = export_option in {"PNG ZIP", "PDF + PNG ZIP"}
export_pdf = export_option in {"PDF", "PDF + PNG ZIP"}

generate = st.button("참가 증명서 생성하기", type="primary", width="stretch")
if generate:
    with st.spinner("참가 증명서를 생성하는 중입니다..."):
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
            file_name="participation_certificates_png.zip",
            mime="application/zip",
            width="stretch",
        )

    if export_pdf and pdf_bytes is not None:
        st.download_button(
            "합본 PDF 다운로드",
            data=pdf_bytes,
            file_name="participation_certificates.pdf",
            mime="application/pdf",
            width="stretch",
        )

    st.subheader("현재 설정(JSON)")
    st.json(
        {
            "names": [row["name"] for row in rows],
            "name": asdict(name_spec),
            "export_option": export_option,
        }
    )
