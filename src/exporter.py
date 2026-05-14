from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .parser import (
    Item,
    KeyValueItem,
    ListItem,
    Section,
    SubSection,
    TaskItem,
    classify_tag,
    extract_assignees,
    extract_date_range,
    extract_thursday_date,
    extract_vacations,
)

# ── AO운영 고정 텍스트 ────────────────────────────────────────────────────
AO_REGULAR_TASKS = (
    "정기 업무(일/주/월)\n"
    "- 상용 모니터링 및 대응 [일작업]\n"
    "- 사업팀 문의 응대 [일작업]\n"
    "- 일일점검 및 일대사(원천사, PG사) 확인 [일작업]\n"
    "- VOC 조치 및 대응 [일작업]\n"
    "- 클립포인트/PG 일대사 불일치 내역 분석 대응 [일작업]\n"
    "- 카드사 시스템점검 일정 확인[일작업]\n"
    "- 주간 가맹점 현황 보고 [주별]\n"
    "- ktds 담당/본부 기준 조근점검 진행\n"
    "- 휴면고객 (5년경과 데이터삭제 내역 확인) [일작업]\n"
    "\n"
    "월 정기업무\n"
    "- 서버백신\n"
    "- 카드사 월정사\n"
    "- pg사 월정사"
)

# 소분류 행 정의: (소분류명, 중분류명, 셀 내 하위항목 목록)
_DEV_ROWS = [
    ("포인트허브 개발",   "개발",   ["일반 개발 항목", "보안 및 인프라"]),
    ("운영 및 장애 대응", "개발",   ["배포 및 모니터링", "버그 수정", "DB 및 서버 관리"]),
    ("기타",              "개발",   []),
    ("정기 업무",         "AO운영", []),
    ("비정기 업무",       "AO운영", []),
]

# ── 색상 ──────────────────────────────────────────────────────────────────
C_HEADER = "D9D9D9"   # 개발파트 헤더 회색
C_TITLE = "DEEBF7"    # 이슈사항 제목 하늘
C_GREEN = "92D050"    # 개발부문/휴가 연두
C_RED = "F66674"      # 운영부문 분홍
C_YELLOW = "FFFF00"   # 이슈사항 항목 헤더 노랑
C_WHITE = "FFFFFF"

# ── 스타일 헬퍼 ───────────────────────────────────────────────────────────

def _fill(color: str) -> PatternFill:
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def _font(size: float = 11, bold: bool = False, color: str = "000000") -> Font:
    return Font(name="맑은 고딕", size=size, bold=bold, color=color)


def _border() -> Border:
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def _align(h: str = "center", v: str = "center", wrap: bool = True) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _set(ws, row: int, col: int, value, fill=None, font=None, align=None, border=True):
    cell = ws.cell(row=row, column=col, value=value)
    if fill:
        cell.fill = fill
    if font:
        cell.font = font
    if align:
        cell.alignment = align
    if border:
        cell.border = _border()
    cell.number_format = "@"
    return cell


def _apply_text_format(ws, max_row: int, max_col: int) -> None:
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            if cell.__class__.__name__ != "MergedCell":
                cell.number_format = "@"


# ── 텍스트 포맷 헬퍼 ─────────────────────────────────────────────────────

_SUBTITLE_MARKER = "▶"
_TIME_MARK_RE = re.compile(
    r"\s*⏱\s*[-\d.]+\s*h?\s*/\s*[-\d.]+\s*h?(?:\s*\([^)]*\))?"
)


def _clean_export_text(text: str) -> str:
    return _TIME_MARK_RE.sub("", text).strip()


def _estimate_text_lines(text: str, col_width: float) -> int:
    if not text:
        return 1
    width = max(int(col_width), 1)
    total = 0
    for ln in text.splitlines():
        total += max(1, -(-len(ln) // width))
    return max(total, 1)


def _comfortable_height(lines: int, minimum: float = 50.0) -> float:
    return max(minimum, lines * 17.0 + 18.0)


def _is_vacation(item: Item) -> bool:
    return isinstance(item, TaskItem) and item.status == "취소" and "휴가" in item.text


def _fmt_task(item: TaskItem) -> str:
    parts = [_clean_export_text(item.text)]
    if item.assignees:
        parts.append(f"({', '.join(item.assignees)})")
    return " ".join(parts)


def _fmt_list(item: ListItem) -> str:
    text = _clean_export_text(item.text)
    if item.note:
        text += f" — {_clean_export_text(item.note)}"
    if item.assignees:
        text += f" ({', '.join(item.assignees)})"
    return text


def _items_to_lines(items: list[Item], prefix: str = "o ") -> list[str]:
    lines = []
    for item in items:
        if isinstance(item, TaskItem):
            lines.append(f"{prefix}{_fmt_task(item)}")
        elif isinstance(item, ListItem):
            lines.append(f"{prefix}{_fmt_list(item)}")
        elif isinstance(item, KeyValueItem):
            lines.append(f"• {item.key}: {item.value or '-'}")
    return lines


def _collect_all_items(sections: list[Section], keywords: list[str]) -> list[Item]:
    """섹션 키워드에 해당하는 모든 항목 수집 (서브섹션 포함)"""
    result: list[Item] = []
    for kw in keywords:
        for sec in sections:
            if kw in sec.title:
                result.extend(sec.items)
                for sub in sec.subsections:
                    result.extend(sub.items)
                    for sub4 in sub.subsections:
                        result.extend(sub4.items)
                break
    return result


def _build_tagged_cell(items: list[Item], sub_items: list[str]) -> str:
    """태그 기반으로 sub_items 블록 구성. 상세 항목 2칸 들여쓰기, 휴가/중복 항목 제외."""
    from collections import defaultdict
    groups: dict[str, list[str]] = defaultdict(list)
    seen: set[str] = set()

    for item in items:
        if _is_vacation(item):
            continue
        if isinstance(item, TaskItem):
            key = item.text
            if key in seen:
                continue
            seen.add(key)
            cat = classify_tag(item.tags)
            groups[cat].append(f"  o {_fmt_task(item)}")
        elif isinstance(item, ListItem):
            key = item.text
            if key in seen:
                continue
            seen.add(key)
            cat = classify_tag(item.tags)
            groups[cat].append(f"  o {_fmt_list(item)}")

    if not sub_items:
        lines = groups.get("기타", [])
        return "\n".join(lines)

    parts: list[str] = []
    for sub in sub_items:
        parts.append(f"{_SUBTITLE_MARKER} {sub}")
        parts.extend(groups.get(sub, []))
        parts.append("")
    return "\n".join(parts).strip()


def _count_tagged_items(items: list[Item], sub_items: list[str]) -> int:
    """해당 소분류 행에 속하는 항목 수 (휴가/중복 제외)"""
    count = 0
    seen: set[str] = set()
    for item in items:
        if _is_vacation(item):
            continue
        if isinstance(item, (TaskItem, ListItem)):
            if item.text in seen:
                continue
            seen.add(item.text)
            cat = classify_tag(item.tags)
        else:
            continue
        if not sub_items or cat in sub_items or (not sub_items and cat == "기타"):
            count += 1
    return count


def _build_업무리스트(sections: list[Section]) -> str:
    """G열: 이번 주 목표 / 완료 업무 / 진행 중 업무 / 다음 주 예정"""
    keywords = ["이번 주 목표", "완료 업무", "진행 중 업무", "다음 주 예정"]
    parts: list[str] = []
    for kw in keywords:
        for sec in sections:
            if kw in sec.title:
                lines = _items_to_lines(sec.items)
                if lines:
                    parts.append(f"[{kw}]")
                    parts.extend(lines)
                    parts.append("")
                break
    return "\n".join(parts).strip()


def _build_금주수행업무(sections: list[Section]) -> str:
    """I열: 업무 수행 내역 (요일별)"""
    for sec in sections:
        if "업무 수행 내역" in sec.title:
            parts: list[str] = []
            for sub in sec.subsections:
                lines = _items_to_lines(sub.items)
                if lines:
                    parts.append(f"[{sub.title}]")
                    parts.extend(lines)
                    parts.append("")
            return "\n".join(parts).strip()
    return ""


def _build_이슈사항_금주(sections: list[Section]) -> str:
    """D열 금주 이슈: 특이사항/이슈 섹션"""
    for sec in sections:
        if "특이사항" in sec.title or "이슈" in sec.title:
            parts: list[str] = []
            for sub3 in sec.subsections:
                parts.append(f"[{sub3.title}]")
                for sub4 in sub3.subsections:
                    parts.append(f"  <{sub4.title}>")
                    parts.extend(f"  {l}" for l in _items_to_lines(sub4.items))
                lines3 = _items_to_lines(sub3.items)
                if lines3:
                    parts.extend(lines3)
                parts.append("")
            return "\n".join(parts).strip()
    return "o 특이사항 없음"


def _build_차주이슈(sections: list[Section]) -> str:
    """E열 차주 이슈: 다음 주 예정"""
    for sec in sections:
        if "다음 주 예정" in sec.title:
            lines = _items_to_lines(sec.items)
            return "\n".join(lines) if lines else "o 특이사항 없음"
    return "o 특이사항 없음"


def _issue_division(title: str) -> str | None:
    if "운영" in title or "AO" in title.upper():
        return "op"
    if "개발" in title:
        return "dev"
    return None


def _issue_timing(title: str) -> str | None:
    normalized = title.replace(" ", "")
    if "차주" in normalized or "다음주" in normalized:
        return "next"
    if "금주" in normalized or "이번주" in normalized:
        return "current"
    return None


def _is_issue_context_title(title: str) -> bool:
    return _issue_division(title) is not None or _issue_timing(title) is not None


def _format_issue_heading(title: str, level: int) -> str:
    return f"[{title}]" if level <= 3 else f"  <{title}>"


def _issue_lines_from_subsection(sub: SubSection) -> list[str]:
    lines: list[str] = []
    if not _is_issue_context_title(sub.title):
        lines.append(_format_issue_heading(sub.title, sub.level))

    item_prefix = "  " if lines else ""
    lines.extend(f"{item_prefix}{line}" for line in _items_to_lines(sub.items))

    for child in sub.subsections:
        lines.extend(_issue_lines_from_subsection(child))
    return lines


def _issue_context_item_text(item: Item) -> str:
    if isinstance(item, (TaskItem, ListItem)):
        return item.text.strip()
    if isinstance(item, KeyValueItem):
        return item.key.strip()
    return ""


def _append_issue_items_by_context(
    groups: dict[tuple[str, str], list[str]],
    items: list[Item],
    base_div: str,
    base_time: str,
) -> None:
    cur_div = base_div
    cur_time = base_time
    for item in items:
        text = _issue_context_item_text(item)
        next_div = _issue_division(text)
        next_time = _issue_timing(text)
        if next_div is not None or next_time is not None:
            cur_div = next_div or cur_div
            cur_time = next_time or cur_time
            continue
        groups[(cur_div, cur_time)].extend(_items_to_lines([item]))


def _split_issue_cells(sections: list[Section]) -> dict[tuple[str, str], str]:
    groups: dict[tuple[str, str], list[str]] = {
        ("dev", "current"): [],
        ("dev", "next"): [],
        ("op", "current"): [],
        ("op", "next"): [],
    }

    for sec in sections:
        if "특이사항" not in sec.title and "이슈" not in sec.title:
            continue

        if sec.items:
            _append_issue_items_by_context(groups, sec.items, "dev", "current")

        for sub3 in sec.subsections:
            div3 = _issue_division(sub3.title)
            time3 = _issue_timing(sub3.title)

            if div3 is None and time3 is None:
                groups[("dev", "current")].extend(_issue_lines_from_subsection(sub3))
                continue

            base_div = div3 or "dev"
            base_time = time3 or "current"
            if sub3.items:
                _append_issue_items_by_context(groups, sub3.items, base_div, base_time)

            for sub4 in sub3.subsections:
                div4 = _issue_division(sub4.title) or base_div
                time4 = _issue_timing(sub4.title) or base_time
                groups[(div4, time4)].extend(_issue_lines_from_subsection(sub4))
        break

    result = {key: "\n".join(lines).strip() for key, lines in groups.items()}
    if not result[("dev", "current")]:
        result[("dev", "current")] = "o 특이사항 없음"
    if not result[("dev", "next")]:
        result[("dev", "next")] = _build_차주이슈(sections)
    if not result[("op", "current")]:
        result[("op", "current")] = "o 특이사항 없음"
    if not result[("op", "next")]:
        result[("op", "next")] = "o 특이사항 없음"
    return result


def _parse_title_info(title: str, d_start: str, d_end: str) -> dict:
    m = re.match(r"(\d+)년\s+(\d+)월\s+(\d+)주", title)
    if m:
        month, week = m.group(2), m.group(3)
        year = m.group(1)
        range_str = f"({d_start} ~ {d_end})" if d_start else ""
        week_display = f"{month}월 {week}주차\n주간보고\n{range_str}".strip()
        period = f"{year}/{month.zfill(2)} ~"
        return {"week_display": week_display, "period": period}
    return {"week_display": title, "period": ""}


def _item_text(item: Item) -> str:
    if isinstance(item, TaskItem):
        return _fmt_task(item)
    if isinstance(item, ListItem):
        return _fmt_list(item)
    if isinstance(item, KeyValueItem):
        return f"{item.key}: {item.value or '-'}"
    return ""


def _collect_vacation_lines(sections: list[Section], year: str) -> list[str]:
    lines: list[str] = []

    for sec in sections:
        if "휴가" not in sec.title:
            continue
        lines.extend(_item_text(item) for item in sec.items if _item_text(item))
        for sub in sec.subsections:
            sub_lines = [_item_text(item) for item in sub.items if _item_text(item)]
            if sub_lines:
                lines.append(f"{sub.title}: {', '.join(sub_lines)}")
            for sub4 in sub.subsections:
                sub4_lines = [_item_text(item) for item in sub4.items if _item_text(item)]
                if sub4_lines:
                    lines.append(f"{sub.title} / {sub4.title}: {', '.join(sub4_lines)}")

    if lines:
        return lines

    for name, date_info in extract_vacations(sections):
        if date_info:
            lines.append(f"{year}.{date_info.replace('/', '.')} {name} 휴가")
        else:
            lines.append(f"{name} 휴가")
    return lines


# ── 개발파트 시트 ─────────────────────────────────────────────────────────

_DEV_COL_WIDTHS = {
    1: 2.5,    # A spacer
    2: 21.5,   # B 대분류
    3: 15.2,   # C 중분류
    4: 32.3,   # D 소분류
    5: 21.2,   # E 업무 수행 주기
    6: 12.7,   # F 작업자
    7: 71.7,   # G 업무리스트
    8: 11.3,   # H 진행률 업무리스트
    9: 71.7,   # I 금주 수행 업무
    10: 15.2,  # J 진행률 금주
    11: 60.3,  # K 비고
    12: 70.3,  # L 업무지시
}

_DEV_HEADERS = {
    2: "대분류",
    3: "중분류",
    4: "소분류",
    5: "업무 수행 주기 \n및 기간",
    6: "작업자",
    7: "업무리스트",
    8: "진행률(%)\n(업무리스트)",
    9: "금주 수행 업무",
    10: "진행률(%)\n(금주수행업무)",
    11: "비고",
    12: "업무지시 및 체크포인트",
}


def _write_개발파트(ws, doc_title: str, sections: list[Section]) -> None:
    d_start, d_end = extract_date_range(sections)
    info = _parse_title_info(doc_title, d_start, d_end)
    assignees = extract_assignees(sections)

    # 컬럼 너비
    for col, w in _DEV_COL_WIDTHS.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    # Row 1: 빈 행
    ws.row_dimensions[1].height = 7.5

    # Row 2: 헤더
    ws.row_dimensions[2].height = 49.25
    hdr_fill = _fill(C_HEADER)
    hdr_font = _font(10, bold=True)
    for col, text in _DEV_HEADERS.items():
        _set(ws, 2, col, text,
             fill=hdr_fill, font=hdr_font,
             align=_align("center", "center"))

    # G열(업무수행내역), I열(완료업무) 원본 수집
    g_items = _collect_all_items(sections, ["업무 수행 내역"])
    i_items = _collect_all_items(sections, ["완료 업무"])

    작업자_str = "\n".join(assignees) if assignees else ""
    data_font = _font(11)
    gray_font = _font(11, color="808080")

    DATA_START = 3
    DATA_END = DATA_START + len(_DEV_ROWS) - 1  # row 7

    # 대분류 병합
    _set(ws, DATA_START, 2, info["week_display"],
         font=data_font, align=_align("center", "center"))
    ws.merge_cells(start_row=DATA_START, start_column=2,
                   end_row=DATA_END, end_column=2)

    # 중분류 병합: 개발(rows 3-5), AO운영(rows 6-7)
    개발_end = DATA_START + 2
    ao_start = DATA_START + 3
    ao_end = DATA_END
    _set(ws, DATA_START, 3, "개발",
         font=data_font, align=_align("center", "center"))
    ws.merge_cells(start_row=DATA_START, start_column=3,
                   end_row=개발_end, end_column=3)
    _set(ws, ao_start, 3, "AO운영",
         font=data_font, align=_align("center", "center"))
    ws.merge_cells(start_row=ao_start, start_column=3,
                   end_row=ao_end, end_column=3)

    for idx, (소분류, 중분류, sub_items) in enumerate(_DEV_ROWS):
        row = DATA_START + idx
        is_first = idx == 0
        is_ao_regular = 소분류 == "정기 업무"
        content_font = gray_font if is_ao_regular else data_font

        # G열(업무리스트): 업무수행내역 태그 기반
        if is_ao_regular:
            g_content = AO_REGULAR_TASKS
        elif 소분류 == "비정기 업무":
            g_content = ""
        else:
            g_content = _build_tagged_cell(g_items, sub_items)

        # I열(금주수행업무): 완료업무 태그 기반
        if is_ao_regular:
            i_content = AO_REGULAR_TASKS
        elif 소분류 == "비정기 업무":
            i_content = ""
        else:
            i_content = _build_tagged_cell(i_items, sub_items)

        # H열: 소분류 행별 완료수/전체수, J열: 100%
        if is_first:
            g_cnt = _count_tagged_items(g_items, sub_items)
            i_cnt = _count_tagged_items(i_items, sub_items)
            h_pct = round(i_cnt / g_cnt * 100) if g_cnt > 0 else 0
            h_val = f"{h_pct}%"
        else:
            h_val = "100%" if (g_content or i_content) else "0%"
        j_val = "100%" if i_content else "0%"

        _set(ws, row, 4, 소분류,
             font=data_font, align=_align("center", "center"))
        _set(ws, row, 5, "", font=data_font, align=_align("center", "center"))
        _set(ws, row, 6, "", font=data_font, align=_align("center", "center"))
        _set(ws, row, 7, g_content,
             font=content_font, align=_align("left", "top"))
        _set(ws, row, 8, h_val,
             font=data_font, align=_align("center", "center"))
        _set(ws, row, 9, i_content,
             font=content_font, align=_align("left", "top"))
        _set(ws, row, 10, j_val,
             font=data_font, align=_align("center", "center"))
        _set(ws, row, 11, "", font=data_font, align=_align("left", "top"))
        _set(ws, row, 12, "", font=data_font, align=_align("left", "top"))

        lines = max(
            _estimate_text_lines(g_content, _DEV_COL_WIDTHS[7]),
            _estimate_text_lines(i_content, _DEV_COL_WIDTHS[9]),
            2,
        )
        ws.row_dimensions[row].height = _comfortable_height(lines)

    _apply_text_format(ws, DATA_END, max(_DEV_COL_WIDTHS))


# ── 이슈사항 시트 ─────────────────────────────────────────────────────────

_ISS_COL_WIDTHS = {
    1: 3.2,   # A 부문
    2: 26.2,  # B 항목
    3: 28.5,  # C 담당자
    4: 89.7,  # D 금주 이슈
    5: 67.0,  # E 차주 이슈
    6: 46.7,  # F 비고
}

_ISS_ITEM_HEADERS = ["항목", "담당자", "금주 이슈", "차주 이슈", "비고"]


def _write_이슈사항(ws, doc_title: str, sections: list[Section]) -> None:
    assignees = extract_assignees(sections)
    thursday = extract_thursday_date(sections)

    # 컬럼 너비
    for col, w in _ISS_COL_WIDTHS.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    # ── Row 1: 제목 (목요일 날짜 사용) ───────────────────────────────────
    year_m = re.match(r"(\d{4})년", doc_title)
    year = year_m.group(1) if year_m else "2026"
    date_str = f"{year}-{thursday.replace('/', '-')}" if thursday else ""
    title_text = f"주간업무 이슈사항 보고  ({date_str})" if date_str else "주간업무 이슈사항 보고"
    _set(ws, 1, 1, title_text,
         fill=_fill(C_TITLE),
         font=_font(16, bold=True),
         align=_align("center", "center"))
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 30

    def _write_section_block(
        section_row: int,
        section_label: str,
        section_fill_color: str,
        data_row: int,
        항목: str,
        담당자: str,
        금주이슈: str,
        차주이슈: str,
        비고: str = "",
    ) -> None:
        _set(ws, section_row, 1, section_label,
             fill=_fill(section_fill_color),
             font=_font(11, bold=True),
             align=_align("center", "center"))
        ws.merge_cells(
            start_row=section_row, start_column=1,
            end_row=data_row, end_column=1
        )

        hdr_fill = _fill(C_YELLOW)
        hdr_font = _font(11, bold=True)
        for i, h in enumerate(_ISS_ITEM_HEADERS, start=2):
            _set(ws, section_row, i, h,
                 fill=hdr_fill, font=hdr_font,
                 align=_align("center", "center"))

        d_font = _font(11)
        _set(ws, data_row, 2, 항목,
             font=d_font, align=_align("center", "center"))
        _set(ws, data_row, 3, 담당자,
             font=d_font, align=_align("center", "center"))
        _set(ws, data_row, 4, 금주이슈,
             font=d_font, align=_align("left", "top"))
        _set(ws, data_row, 5, 차주이슈,
             font=d_font, align=_align("left", "top"))
        _set(ws, data_row, 6, 비고,
             font=d_font, align=_align("left", "top"))

        ws.row_dimensions[section_row].height = 24
        max_lines = max(
            _estimate_text_lines(항목, _ISS_COL_WIDTHS[2]),
            _estimate_text_lines(담당자, _ISS_COL_WIDTHS[3]),
            _estimate_text_lines(금주이슈, _ISS_COL_WIDTHS[4]),
            _estimate_text_lines(차주이슈, _ISS_COL_WIDTHS[5]),
            _estimate_text_lines(비고, _ISS_COL_WIDTHS[6]),
            2,
        )
        ws.row_dimensions[data_row].height = _comfortable_height(max_lines)

    issue_cells = _split_issue_cells(sections)
    담당자_str = "\n".join(assignees) if assignees else ""

    # ── 개발부문: D3/E3 고정 ─────────────────────────────────────────────
    _write_section_block(
        section_row=2,
        section_label="개발부문",
        section_fill_color=C_GREEN,
        data_row=3,
        항목="포인트허브",
        담당자=담당자_str,
        금주이슈=issue_cells[("dev", "current")],
        차주이슈=issue_cells[("dev", "next")],
    )

    # ── 운영부문: D5/E5 고정 ─────────────────────────────────────────────
    _write_section_block(
        section_row=4,
        section_label="운영부문",
        section_fill_color=C_RED,
        data_row=5,
        항목="포인트허브",
        담당자=담당자_str,
        금주이슈=issue_cells[("op", "current")],
        차주이슈=issue_cells[("op", "next")],
    )

    # ── 휴가: row 6 한 줄 고정 ───────────────────────────────────────────
    vacation_row = 6
    vacation_text = " / ".join(_collect_vacation_lines(sections, year))
    _set(ws, vacation_row, 1, "휴가",
         fill=_fill(C_GREEN),
         font=_font(11, bold=True),
         align=_align("center", "center"))
    _set(ws, vacation_row, 2, vacation_text,
         font=_font(11),
         align=_align("left", "center"))
    ws.merge_cells(
        start_row=vacation_row, start_column=2,
        end_row=vacation_row, end_column=6
    )
    vacation_width = sum(_ISS_COL_WIDTHS[c] for c in range(2, 7))
    ws.row_dimensions[vacation_row].height = _comfortable_height(
        _estimate_text_lines(vacation_text, vacation_width),
        minimum=28.0,
    )
    _apply_text_format(ws, vacation_row, max(_ISS_COL_WIDTHS))


# ── 진입점 ────────────────────────────────────────────────────────────────

def export_to_xlsx(
    doc_title: str,
    sections: list[Section],
    output_path: Path,
) -> list[str]:
    wb = Workbook()

    ws_dev = wb.active
    ws_dev.title = "개발파트"
    _write_개발파트(ws_dev, doc_title, sections)

    ws_iss = wb.create_sheet("포인트허브_이슈사항")
    _write_이슈사항(ws_iss, doc_title, sections)

    wb.save(output_path)
    return [ws.title for ws in wb.worksheets]
