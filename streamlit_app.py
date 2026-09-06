"""정비통제업무일지 초안 생성기 (Streamlit).

Daily Work Order 엑셀을 올리면 '항공기 주요 작업 사항 (정비 1/정비 2)' 칸에
작업 제목이 채워진 업무일지 docx 초안을 만들어 준다.
작업 결과(=> 뒤)는 근무 종료 후 직접 입력한다.
"""

from __future__ import annotations

import io
import os
import tempfile

import streamlit as st

from worksheet_auto import (
    build_worksheet,
    collect_warnings,
    merge_blocks,
    output_filename,
    parse_workorder,
    render_all,
)

BASE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(BASE, "template", "정비통제업무일지_TEMPLATE.docx")

st.set_page_config(page_title="정비통제업무일지 초안 생성기", page_icon="✈️", layout="wide")
st.title("✈️ 정비통제업무일지 초안 생성기")
st.caption(
    "Daily Work Order 엑셀 → 업무일지 docx 초안. "
    "정비 1(GMP) / 정비 2(ICN) 작업 제목만 채워지고, 결과는 근무 후 직접 입력합니다."
)


def _save_upload(upload, suffix: str) -> str:
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(upload.getbuffer())
    return path


with st.sidebar:
    st.header("설정")
    arrow = st.checkbox("작업 제목 뒤에 '=>' 붙이기", value=True,
                        help="결과를 바로 이어 쓸 수 있게 화살표를 미리 넣어 둡니다.")
    clear_previous = st.checkbox("템플릿의 지난 내용 비우기", value=True,
                                 help="인수인계자·기상·FLT PLAN·근무인원·IRR·SEAT BLOCK·국내/해외 칸을 비웁니다.")
    st.divider()
    st.subheader("업무일지 템플릿")
    template_upload = st.file_uploader(
        "양식이 바뀌면 새 업무일지 docx를 올리세요", type=["docx"], key="tpl"
    )
    if template_upload is None:
        st.caption(f"기본 템플릿 사용: `{os.path.basename(DEFAULT_TEMPLATE)}`")

xlsx_upload = st.file_uploader("Daily Work Order 엑셀 업로드", type=["xlsx", "xlsm"])

if xlsx_upload is None:
    st.info("Daily Work Order 엑셀 파일을 올려 주세요.")
    st.stop()

xlsx_path = _save_upload(xlsx_upload, ".xlsx")
template_path = _save_upload(template_upload, ".docx") if template_upload else DEFAULT_TEMPLATE

try:
    blocks = parse_workorder(xlsx_path)
except Exception as exc:  # noqa: BLE001
    st.error(f"엑셀을 읽지 못했습니다: {exc}")
    st.stop()

merged = merge_blocks(blocks)
if not merged:
    st.error("Daily Work Order 블록(‘Date … Shift’ 행)을 찾지 못했습니다. 양식을 확인해 주세요.")
    st.stop()

options = {
    f"{b.key}   (작업 {len(b.entries)}건 · STA {', '.join(b.stations()) or '-'})": key
    for key, b in merged.items()
}
choice = st.radio("생성할 근무 선택", list(options.keys()), horizontal=False)
block = merged[options[choice]]

rendered = render_all(block.entries, arrow=arrow)
warnings = collect_warnings(block.entries)

col1, col2 = st.columns(2)
for col, name in ((col1, "정비 1"), (col2, "정비 2")):
    with col:
        sta = "GMP" if name == "정비 1" else "ICN"
        text = rendered.get(name, "")
        count = len([l for l in text.splitlines() if not l.startswith("[")])
        st.subheader(f"항공기 주요 작업 사항 ({name})")
        st.caption(f"{sta} · {count}건")
        st.text_area(name, value=text or "(해당 작업 없음)", height=460,
                     label_visibility="collapsed")

if warnings:
    with st.expander(f"확인 필요 {len(warnings)}건", expanded=True):
        for note in warnings:
            st.write("• " + note)

st.caption("‘국내. 해외’ 칸은 IRR 사항이라 초안을 만들지 않습니다. IRR·기상·근무인원도 직접 입력해 주세요.")

out_name = output_filename(block.date, block.shift)
out_path = os.path.join(tempfile.mkdtemp(), out_name)
build_worksheet(
    template_path,
    out_path,
    block.entries,
    block.date,
    block.shift,
    arrow=arrow,
    clear_previous=clear_previous,
)
with open(out_path, "rb") as f:
    st.download_button(
        f"📄 {out_name} 내려받기",
        data=f.read(),
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )
