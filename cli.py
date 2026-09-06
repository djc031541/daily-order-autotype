#!/usr/bin/env python3
"""Daily Work Order 엑셀 -> 정비통제업무일지 초안(docx) 생성 CLI.

사용 예)
    python cli.py "Daily Work Order (2026-09-02).xlsx"            # 블록 목록 보기
    python cli.py "..." --shift D                                  # D 근무 초안 생성
    python cli.py "..." --date 2026-09-02 --shift N --print        # 본문만 출력
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys

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
DEFAULT_OUTDIR = os.path.join(BASE, "output")


def pick_block(blocks, date_str: str | None, shift: str | None):
    merged = merge_blocks(blocks)
    candidates = list(merged.values())
    if date_str:
        want = _dt.datetime.strptime(date_str, "%Y-%m-%d").date()
        candidates = [b for b in candidates if b.date == want]
    if shift:
        candidates = [b for b in candidates if b.shift == shift.upper()]
    return candidates


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Daily Work Order -> 정비통제업무일지 초안")
    ap.add_argument("xlsx", help="Daily Work Order 엑셀 경로")
    ap.add_argument("--date", help="대상 날짜 (YYYY-MM-DD). 생략 시 전체")
    ap.add_argument("--shift", choices=["D", "N", "d", "n"], help="대상 근무 (D/N)")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE, help="업무일지 템플릿 docx")
    ap.add_argument("--outdir", default=DEFAULT_OUTDIR, help="저장 폴더")
    ap.add_argument("--print", dest="print_only", action="store_true",
                    help="docx를 만들지 않고 본문만 출력")
    ap.add_argument("--no-arrow", action="store_true", help="제목 뒤 '=>' 를 붙이지 않음")
    ap.add_argument("--keep-previous", action="store_true",
                    help="템플릿의 기존 IRR/기상/인원 값을 지우지 않음")
    args = ap.parse_args(argv)

    blocks = parse_workorder(args.xlsx)
    if not blocks:
        print("Daily Work Order 블록을 찾지 못했습니다.", file=sys.stderr)
        return 1

    targets = pick_block(blocks, args.date, args.shift)
    if not targets:
        print("조건에 맞는 블록이 없습니다. 이 파일에 들어있는 블록:", file=sys.stderr)
        for b in merge_blocks(blocks).values():
            print(f"  - {b.key}  (작업 {len(b.entries)}건, STA {', '.join(b.stations())})",
                  file=sys.stderr)
        return 1

    if not args.date and not args.shift and len(targets) > 1:
        print("이 파일에 들어있는 블록:")
        for b in targets:
            print(f"  - {b.key}  (작업 {len(b.entries)}건, STA {', '.join(b.stations())})")
        print("\n--date / --shift 로 대상을 지정하세요.")
        return 0

    os.makedirs(args.outdir, exist_ok=True)
    for block in targets:
        rendered = render_all(block.entries, arrow=not args.no_arrow)
        if args.print_only:
            print(f"===== {block.key} =====")
            for name, text in rendered.items():
                print(f"--- 항공기 주요 작업 사항 ({name}) ---")
                print(text or "(해당 없음)")
                print()
            continue

        out_path = os.path.join(args.outdir, output_filename(block.date, block.shift))
        build_worksheet(
            args.template,
            out_path,
            block.entries,
            block.date,
            block.shift,
            arrow=not args.no_arrow,
            clear_previous=not args.keep_previous,
        )
        counts = {name: len([l for l in t.splitlines() if not l.startswith("[")])
                  for name, t in rendered.items()}
        print(f"생성: {out_path}")
        print(f"  정비 1 {counts.get('정비 1', 0)}건 / 정비 2 {counts.get('정비 2', 0)}건")
        for note in collect_warnings(block.entries):
            print(f"  [확인] {note}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
