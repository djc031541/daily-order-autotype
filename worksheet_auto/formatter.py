"""파싱된 작업 목록을 정비통제업무일지 본문 형식으로 렌더링.

출력 예)
    [B738]
    HL8547 1) CL-BI-WEEKLY CHECK=>
             2) AD-ULTRASONIC AND EXTERNAL DETAILED INSPECTION - ...=>
    [B38M]
    HL8751 1) TRP-REPLACEMENT - AIRCRAFT BATTERY=>
"""

from __future__ import annotations

import re
from collections import OrderedDict
from typing import Iterable

from .parser import Entry

# 업무일지 섹션 <- Daily Work Order 의 STA
SECTION_BY_STATION = {
    "GMP": "정비 1",
    "ICN": "정비 2",
}
# IRR 성격이라 초안을 만들지 않는 섹션 (지방/해외 스테이션)
OUTSTATION_SECTION = "국내. 해외"

SECTION_ORDER = ["정비 1", "정비 2"]

# 샘플 업무일지의 이어지는 항목 들여쓰기 (공백 9칸)
CONT_INDENT = " " * 9

# 엑셀에 기번이 비어 있는 행(기종 공통 작업 등)에 쓰는 표시
NO_REG = "N/A"

_DEF_CLR_RE = re.compile(r"^\[?\s*DEF\s*(?:CLR|CLEAR)\s*\]?\s*[:\-]?\s*(.*)$", re.I)


def section_of(entry: Entry) -> str:
    return SECTION_BY_STATION.get(entry.station.upper(), OUTSTATION_SECTION)


def entry_body(entry: Entry, ) -> str:
    """'NRC-TITLE' / 'DEF CLR: TITLE' 형태의 본문 한 줄."""
    wtype = (entry.type or "").strip().upper()
    desc = " ".join((entry.desc or "").split()).strip()
    if not desc:
        return ""

    matched = _DEF_CLR_RE.match(desc)
    if wtype == "DEF" or matched:
        if matched:
            return "DEF CLR: " + matched.group(1).strip()
        return "DEF-" + desc
    if not wtype:
        return desc
    return f"{wtype}-{desc}"


def build_sections(
    entries: Iterable[Entry],
    include_outstation: bool = False,
) -> "OrderedDict[str, OrderedDict[str, OrderedDict[str, list[Entry]]]]":
    """{섹션: {기종: {기번: [Entry, ...]}}} — 모두 엑셀 등장 순서 유지."""
    sections: OrderedDict[str, OrderedDict[str, OrderedDict[str, list[Entry]]]] = OrderedDict()
    for name in SECTION_ORDER:
        sections[name] = OrderedDict()
    if include_outstation:
        sections[OUTSTATION_SECTION] = OrderedDict()

    for entry in entries:
        name = section_of(entry)
        if name not in sections:
            continue
        if not entry_body(entry):
            continue
        model = (entry.model or NO_REG).upper()
        reg = entry.reg or NO_REG
        sections[name].setdefault(model, OrderedDict()).setdefault(reg, []).append(entry)

    return sections


def render_section(models: "OrderedDict[str, OrderedDict[str, list[Entry]]]",
                   arrow: bool = True) -> str:
    """한 섹션(정비 1 / 정비 2)의 본문 텍스트."""
    lines: list[str] = []
    tail = "=>" if arrow else ""
    for model, regs in models.items():
        lines.append(f"[{model}]")
        for reg, items in regs.items():
            for i, entry in enumerate(items, start=1):
                body = entry_body(entry)
                if i == 1:
                    lines.append(f"{reg} {i}) {body}{tail}")
                else:
                    lines.append(f"{CONT_INDENT}{i}) {body}{tail}")
    return "\n".join(lines)


def render_all(entries: Iterable[Entry], arrow: bool = True) -> "OrderedDict[str, str]":
    sections = build_sections(entries)
    return OrderedDict((name, render_section(models, arrow)) for name, models in sections.items())


def collect_warnings(entries: Iterable[Entry]) -> list[str]:
    """초안에 넣긴 했지만 사람이 한 번 봐야 하는 행들."""
    notes: list[str] = []
    for entry in entries:
        if section_of(entry) == OUTSTATION_SECTION:
            continue
        where = f"엑셀 {entry.row}행"
        if not entry.reg:
            notes.append(f"{where}: 기번이 비어 있어 '{NO_REG}'로 표기 - {entry.desc[:60]}")
        if not entry.type:
            notes.append(f"{where}: Type이 비어 있음 - {entry.desc[:60]}")
        if not entry.model:
            notes.append(f"{where}: 기종이 비어 있음 - {entry.desc[:60]}")
    return notes
