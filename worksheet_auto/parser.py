"""Daily Work Order 엑셀 파서.

한 파일 안에 (날짜 + Shift) 단위 블록이 여러 개 세로로 쌓여 있다.
각 블록은 A열 'Date' 행으로 시작하고, 그 아래 "REG' NO." 헤더 행이 나온 뒤
실제 작업 행들이 이어진다. 컬럼 위치는 헤더 텍스트로 매핑한다
(엑셀 양식이 바뀌어도 헤더만 같으면 동작하도록).
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import Any, Iterable

import openpyxl

# 헤더를 못 찾았을 때 쓰는 기본 컬럼 (1-based)
DEFAULT_COLS = {
    "sta": 1,
    "model": 2,
    "reg": 3,
    "arr": 4,
    "spot": 5,
    "dep": 6,
    "wo": 7,
    "type": 8,
    "desc": 9,
    "mh": 14,
    "insp": 15,
    "remark": 16,
}

_EXACT_HEADERS = {
    "sta": "sta",
    "sta'": "sta",
    "model": "model",
    "reg' no.": "reg",
    "reg no.": "reg",
    "reg no": "reg",
    "도착시각": "arr",
    "spot no.": "spot",
    "출발시각": "dep",
    "type": "type",
    "std m/h": "mh",
    "insp'": "insp",
    "insp": "insp",
    "자재 / 비고": "remark",
    "자재/비고": "remark",
}


def norm(value: Any) -> str:
    """셀 값을 정규화한 문자열로. 전각공백/NBSP 제거, 연속 공백 축약."""
    if value is None:
        return ""
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).replace("　", " ").replace("\xa0", " ")
    return " ".join(text.split()).strip()


def _flatten(value: Any) -> str:
    """줄바꿈만 공백으로 바꾸고 나머지는 유지."""
    if value is None:
        return ""
    text = str(value).replace("　", " ").replace("\xa0", " ")
    return " ".join(text.split()).strip()


@dataclass
class Entry:
    """작업 한 건 (엑셀 한 행)."""

    row: int
    station: str
    model: str
    reg: str
    wo: str
    type: str
    desc: str
    mh: str = ""
    insp: str = ""
    remark: str = ""


@dataclass
class Block:
    """(날짜, Shift) 단위 Daily Work Order 블록."""

    date: _dt.date | None
    shift: str
    start_row: int
    end_row: int
    entries: list[Entry] = field(default_factory=list)

    @property
    def key(self) -> str:
        d = self.date.strftime("%Y-%m-%d") if self.date else "?"
        return f"{d} {self.shift or '?'}"

    @property
    def label(self) -> str:
        d = self.date.strftime("%d-%b-%y").upper() if self.date else "?"
        return f"({d}){self.shift}"

    def stations(self) -> list[str]:
        seen: list[str] = []
        for e in self.entries:
            if e.station and e.station not in seen:
                seen.append(e.station)
        return seen


def _row_values(ws, row: int, max_col: int) -> list[Any]:
    return [ws.cell(row, c).value for c in range(1, max_col + 1)]


def _is_date_row(values: list[Any]) -> bool:
    return norm(values[0]).lower() == "date" if values else False


def _find_labeled(values: list[Any], label: str) -> Any:
    """label 셀을 찾아 그 오른쪽 첫 비어있지 않은 값을 반환."""
    target = label.lower()
    for i, v in enumerate(values):
        if norm(v).lower() == target:
            for j in range(i + 1, len(values)):
                if norm(values[j]):
                    return values[j]
            return None
    return None


def _to_date(value: Any) -> _dt.date | None:
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    text = norm(value)
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return _dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _map_columns(values: list[Any]) -> dict[str, int] | None:
    """헤더 행이면 {필드: 1-based 컬럼} 반환, 아니면 None."""
    cols: dict[str, int] = {}
    for i, v in enumerate(values):
        text = norm(v).lower()
        if not text:
            continue
        col = i + 1
        if text in _EXACT_HEADERS:
            cols.setdefault(_EXACT_HEADERS[text], col)
        elif "w/o no" in text or "erp no" in text:
            cols.setdefault("wo", col)
        elif "order description" in text or "title / action" in text or "title/action" in text:
            cols.setdefault("desc", col)
    if "reg" not in cols or "model" not in cols:
        return None
    # Description 헤더가 비어 있는 양식이 있어 Type 바로 옆 컬럼으로 보정
    if "desc" not in cols and "type" in cols:
        cols["desc"] = cols["type"] + 1
    for key, col in DEFAULT_COLS.items():
        cols.setdefault(key, col)
    return cols


def _normalize_reg(reg: str) -> str:
    text = norm(reg).upper().replace(" ", "")
    if not text:
        return ""
    if text.startswith("HL"):
        return text
    if re.fullmatch(r"\d{3,5}", text):
        return "HL" + text
    return text


def parse_workorder(path: str) -> list[Block]:
    """Daily Work Order 엑셀을 (날짜, Shift) 블록 목록으로 파싱."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    max_col = max(ws.max_column, 19)
    max_row = ws.max_row

    rows = {r: _row_values(ws, r, max_col) for r in range(1, max_row + 1)}

    starts = [r for r in range(1, max_row + 1) if _is_date_row(rows[r])]
    blocks: list[Block] = []

    for idx, start in enumerate(starts):
        end = starts[idx + 1] - 1 if idx + 1 < len(starts) else max_row
        header = rows[start]
        block = Block(
            date=_to_date(_find_labeled(header, "Date")),
            shift=norm(_find_labeled(header, "Shift")).upper(),
            start_row=start,
            end_row=end,
        )

        cols: dict[str, int] | None = None
        station = model = reg = last_type = ""

        for r in range(start, end + 1):
            values = rows[r]
            if cols is None:
                cols = _map_columns(values)
                if cols is not None:
                    station = model = reg = last_type = ""
                continue

            def cell(key: str) -> str:
                col = cols[key]
                return _flatten(values[col - 1]) if col <= len(values) else ""

            sta = norm(cell("sta")).upper()
            if sta:
                station = sta
            row_model = norm(cell("model")).upper()
            row_reg = cell("reg")
            if row_model or row_reg:
                model = row_model or model
                reg = _normalize_reg(row_reg) or reg
                last_type = ""

            wo = cell("wo")
            wtype = norm(cell("type")).upper()
            desc = cell("desc")

            if not desc:
                continue  # 작업 내용이 없으면 일지에 올릴 것이 없다
            if wtype:
                last_type = wtype
            else:
                wtype = last_type  # 같은 항공기에서 Type을 생략한 연속 행

            block.entries.append(
                Entry(
                    row=r,
                    station=station,
                    model=model,
                    reg=reg,
                    wo=wo,
                    type=wtype,
                    desc=desc,
                    mh=cell("mh"),
                    insp=cell("insp"),
                    remark=cell("remark"),
                )
            )

        blocks.append(block)

    return blocks


def merge_blocks(blocks: Iterable[Block]) -> dict[str, Block]:
    """같은 (날짜, Shift)를 가진 블록들을 하나로 합친다 (GMP/ICN/지방 블록이 분리돼 있음)."""
    merged: dict[str, Block] = {}
    for b in blocks:
        if b.key in merged:
            target = merged[b.key]
            target.entries.extend(b.entries)
            target.end_row = max(target.end_row, b.end_row)
        else:
            merged[b.key] = Block(
                date=b.date,
                shift=b.shift,
                start_row=b.start_row,
                end_row=b.end_row,
                entries=list(b.entries),
            )
    return merged
