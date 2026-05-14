from __future__ import annotations

from datetime import date, datetime, timedelta
import re
from dataclasses import dataclass, field
from typing import Union

CHECKBOX_STATUS: dict[str, str] = {
    "x": "완료",
    " ": "미완",
    "-": "취소",
    "/": "처리",
}

TABLE_STATUS: dict[str, str] = {
    "[x]": "완료",
    "[ ]": "미완",
    "[/]": "처리",
    "[-]": "취소",
    "[>]": "미완",
    "[<]": "미완",
    "[!]": "미완",
    "[?]": "미완",
}

CHECKBOX_ICON: dict[str, str] = {
    "완료": "✅",
    "미완": "⬜",
    "취소": "➖",
    "처리": "✔",
}


@dataclass
class TaskItem:
    status: str
    text: str
    hours_est: float | None = None
    hours_act: float | None = None
    assignees: list[str] = field(default_factory=list)
    tags: list[str] = field(default_factory=list)


@dataclass
class ListItem:
    text: str
    note: str = ""
    assignees: list[str] = field(default_factory=list)
    tags: list[str] = field(default_factory=list)


@dataclass
class KeyValueItem:
    key: str
    value: str


Item = Union[TaskItem, ListItem, KeyValueItem]


@dataclass
class SubSection:
    title: str
    level: int
    items: list[Item] = field(default_factory=list)
    subsections: list["SubSection"] = field(default_factory=list)


@dataclass
class Section:
    title: str
    items: list[Item] = field(default_factory=list)
    subsections: list[SubSection] = field(default_factory=list)


def _parse_task_item(line: str) -> TaskItem | None:
    m = re.match(r"^-\s+\[([x \-/])\]\s+(.*)", line)
    if not m:
        return None
    status = CHECKBOX_STATUS.get(m.group(1), "미완")
    rest = m.group(2).strip()

    tags = re.findall(r"#(\w+)", rest)
    rest = re.sub(r"\s*#\w+", "", rest).strip()

    hours_est = hours_act = None
    assignees: list[str] = []

    timer = re.search(
        r"⏱\s*([-\d.]+)\s*h?\s*/\s*([-\d.]+)\s*h?(?:\s*\(([^)]*)\))?",
        rest,
    )
    if timer:
        if timer.group(1) != "-":
            hours_est = float(timer.group(1))
        if timer.group(2) != "-":
            hours_act = float(timer.group(2))
        if timer.group(3):
            assignees = [a.strip() for a in timer.group(3).split(",")]
        rest = rest[: timer.start()].strip()

    return TaskItem(
        status=status,
        text=rest,
        hours_est=hours_est,
        hours_act=hours_act,
        assignees=assignees,
        tags=tags,
    )


def _parse_list_item(line: str) -> ListItem | None:
    m = re.match(r"^-\s+(.*)", line)
    if not m:
        return None
    text = m.group(1).strip()
    if not text:
        return None

    tags = re.findall(r"#(\w+)", text)
    text = re.sub(r"\s*#\w+", "", text).strip()

    note = ""
    if " — " in text:
        text, note = text.split(" — ", 1)
        text = text.strip()
        note = note.strip()

    # 담당자: 줄 끝 짧은 한국어 이름만 e.g. "(김학진)" or "(김학진, 고영중)"
    # 공백 포함 문장(설명)은 제외
    assignees: list[str] = []
    end_paren = re.search(r"\(([가-힣]{2,4}(?:,\s*[가-힣]{2,4})*)\)$", text)
    if end_paren:
        assignees = [a.strip() for a in end_paren.group(1).split(",")]
        text = text[: end_paren.start()].strip()

    return ListItem(text=text, note=note, assignees=assignees, tags=tags)


def _parse_kv(line: str) -> KeyValueItem | None:
    m = re.match(r"^\*\*([^*]+)\*\*:\s*(.*)", line)
    if m:
        return KeyValueItem(key=m.group(1), value=m.group(2).strip())
    return None


def _parse_markdown_sections(text: str) -> tuple[str, list[Section]]:
    lines = text.splitlines()
    doc_title = ""
    sections: list[Section] = []
    cur_sec: Section | None = None
    cur_sub3: SubSection | None = None
    cur_sub4: SubSection | None = None

    for raw in lines:
        line = raw.strip()

        if line.startswith("# ") and not line.startswith("## "):
            doc_title = line[2:].strip()
            continue

        if line.startswith("## "):
            cur_sub3 = cur_sub4 = None
            cur_sec = Section(title=line[3:].strip())
            sections.append(cur_sec)
            continue

        if line.startswith("### "):
            cur_sub4 = None
            cur_sub3 = SubSection(title=line[4:].strip(), level=3)
            if cur_sec is not None:
                cur_sec.subsections.append(cur_sub3)
            continue

        if line.startswith("#### "):
            cur_sub4 = SubSection(title=line[5:].strip(), level=4)
            if cur_sub3 is not None:
                cur_sub3.subsections.append(cur_sub4)
            continue

        if line in ("---", ""):
            continue

        target: Section | SubSection | None = cur_sub4 or cur_sub3 or cur_sec
        if target is None:
            continue

        task = _parse_task_item(line)
        if task:
            target.items.append(task)
            continue

        kv = _parse_kv(line)
        if kv:
            target.items.append(kv)
            continue

        item = _parse_list_item(line)
        if item:
            target.items.append(item)

    return doc_title, sections


def _parse_frontmatter(text: str) -> tuple[dict[str, str], list[str]]:
    lines = text.splitlines()
    meta: dict[str, str] = {}
    if not lines or lines[0].strip() != "---":
        return meta, lines

    for idx, line in enumerate(lines[1:], start=1):
        if line.strip() == "---":
            return meta, lines[idx + 1:]
        if ":" in line:
            key, value = line.split(":", 1)
            meta[key.strip()] = value.strip()

    return meta, lines


def _is_work_management_document(text: str) -> bool:
    return (
        "# 업무 관리" in text
        and "| 순번" in text
        and "| 시작일" in text
        and "| 업무내용" in text
    )


def _split_table_row(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def _is_table_separator(cells: list[str]) -> bool:
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell.strip()) for cell in cells)


def _parse_task_table(lines: list[str]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    headers: list[str] = []
    in_task_table = False

    for line in lines:
        if not line.lstrip().startswith("|"):
            if in_task_table:
                break
            continue

        cells = _split_table_row(line)
        if not in_task_table:
            if {"시작일", "종료일", "상태", "업무내용"}.issubset(set(cells)):
                headers = cells
                in_task_table = True
            continue

        if _is_table_separator(cells):
            continue

        row = {header: cells[idx] if idx < len(cells) else "" for idx, header in enumerate(headers)}
        if row.get("업무내용", "").strip():
            rows.append(row)

    return rows


def _parse_period(period: str) -> tuple[date | None, date | None]:
    m = re.search(
        r"(\d{4})-(\d{1,2})-(\d{1,2})\s*~\s*(\d{4})-(\d{1,2})-(\d{1,2})",
        period,
    )
    if not m:
        return None, None
    start = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    end = date(int(m.group(4)), int(m.group(5)), int(m.group(6)))
    return start, end


def _parse_table_date(value: str, year: int) -> date | None:
    m = re.search(r"(\d{1,2})[-/](\d{1,2})", value)
    if not m:
        return None
    return date(year, int(m.group(1)), int(m.group(2)))


def _parse_hours(value: str) -> tuple[float | None, float | None]:
    m = re.search(r"([-\d.]+)\s*/\s*([-\d.]+)", value)
    if not m:
        return None, None

    def _to_float(raw: str) -> float | None:
        return None if raw == "-" else float(raw)

    return _to_float(m.group(1)), _to_float(m.group(2))


def _table_row_to_task(row: dict[str, str]) -> TaskItem:
    tags = re.findall(r"#(\w+)", row.get("태그", ""))
    hours_est, hours_act = _parse_hours(row.get("예상/실제", ""))
    assignees = [
        assignee.strip()
        for assignee in re.split(r"[,/]", row.get("작업자", ""))
        if assignee.strip()
    ]
    return TaskItem(
        status=TABLE_STATUS.get(row.get("상태", "").strip(), "미완"),
        text=row.get("업무내용", "").strip(),
        hours_est=hours_est,
        hours_act=hours_act,
        assignees=assignees,
        tags=tags,
    )


def _task_to_list_item(task: TaskItem) -> ListItem:
    return ListItem(
        text=task.text,
        assignees=list(task.assignees),
        tags=list(task.tags),
    )


def _copy_task(task: TaskItem) -> TaskItem:
    return TaskItem(
        status=task.status,
        text=task.text,
        hours_est=task.hours_est,
        hours_act=task.hours_act,
        assignees=list(task.assignees),
        tags=list(task.tags),
    )


def _date_range(start: date, end: date) -> list[date]:
    days: list[date] = []
    cur = start
    while cur <= end:
        days.append(cur)
        cur += timedelta(days=1)
    return days


def _build_work_management_sections(
    rows: list[dict[str, str]],
    period_start: date | None,
    period_end: date | None,
) -> list[Section]:
    year = period_start.year if period_start else datetime.today().year
    tasks = [(_table_row_to_task(row), row) for row in rows]

    goals = Section(title="📌 이번 주 목표")
    goals.items = [_copy_task(task) for task, _row in tasks]

    daily = Section(title="📋 업무 수행 내역")
    if period_start and period_end:
        weekdays = ["월", "화", "수", "목", "금", "토", "일"]
        for day in _date_range(period_start, period_end):
            sub = SubSection(
                title=f"{weekdays[day.weekday()]} ({day.strftime('%m/%d')})",
                level=3,
            )
            for task, row in tasks:
                task_start = _parse_table_date(row.get("시작일", ""), year)
                task_end = _parse_table_date(row.get("종료일", ""), year) or task_start
                if task_start and task_end and task_start <= day <= task_end:
                    sub.items.append(_copy_task(task))
            daily.subsections.append(sub)

    completed = Section(title="✅ 완료 업무")
    in_progress = Section(title="🔄 진행 중 업무")
    for task, _row in tasks:
        if task.status == "완료":
            completed.items.append(_task_to_list_item(task))
        else:
            in_progress.items.append(_task_to_list_item(task))

    return [goals, daily, completed, in_progress]


def _parse_work_management_document(text: str) -> tuple[str, list[Section]]:
    meta, lines = _parse_frontmatter(text)
    week = meta.get("주차", "").strip()
    if week:
        doc_title = week if "주간 보고" in week else f"{week} 주간 보고"
    else:
        doc_title = "주간 보고"
    period_start, period_end = _parse_period(meta.get("기간", ""))

    generated_sections = _build_work_management_sections(
        _parse_task_table(lines),
        period_start,
        period_end,
    )
    _ignored_title, existing_sections = _parse_markdown_sections(text)
    passthrough_sections = [
        section
        for section in existing_sections
        if "대기/조건부 작업" not in section.title
    ]
    return doc_title, [*generated_sections, *passthrough_sections]


def parse_document(text: str) -> tuple[str, list[Section]]:
    if _is_work_management_document(text):
        return _parse_work_management_document(text)
    return _parse_markdown_sections(text)


def extract_date_range(sections: list[Section]) -> tuple[str, str]:
    for sec in sections:
        if "업무 수행 내역" in sec.title:
            dates = [
                m.group(1)
                for sub in sec.subsections
                if (m := re.search(r"\((\d{2}/\d{2})\)", sub.title))
            ]
            if dates:
                return dates[0], dates[-1]
    return "", ""


def extract_assignees(sections: list[Section]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []

    def _collect(items: list[Item]) -> None:
        for item in items:
            names: list[str] = []
            if isinstance(item, TaskItem):
                names = item.assignees
            elif isinstance(item, ListItem):
                names = item.assignees
            for name in names:
                if name and name not in seen and name.lower() != "all":
                    seen.add(name)
                    result.append(name)

    def _walk(sec: Section | SubSection) -> None:
        _collect(sec.items)
        for sub in sec.subsections:
            _collect(sub.items)
            for sub4 in sub.subsections:
                _collect(sub4.items)

    for sec in sections:
        _walk(sec)

    return result


_CATEGORY_PRIORITY = ["보안", "배포", "버그", "db", "기타"]
_CATEGORY_TAG_MAP: dict[str, str] = {
    "보안": "보안 및 인프라",
    "배포": "배포 및 모니터링",
    "버그": "버그 수정",
    "db": "DB 및 서버 관리",
    "기타": "기타",
}


def classify_tag(tags: list[str]) -> str:
    """태그 목록에서 분류 카테고리 결정. 우선순위: 보안>배포>버그>db>기타>기본"""
    for p in _CATEGORY_PRIORITY:
        if p in tags:
            return _CATEGORY_TAG_MAP[p]
    return "일반 개발 항목"


def extract_thursday_date(sections: list[Section]) -> str:
    """업무 수행 내역에서 목요일 날짜(MM/DD) 추출"""
    for sec in sections:
        if "업무 수행 내역" in sec.title:
            for sub in sec.subsections:
                if sub.title.startswith("목"):
                    m = re.search(r"\((\d{2}/\d{2})\)", sub.title)
                    if m:
                        return m.group(1)
    return ""


def extract_vacations(sections: list[Section]) -> list[tuple[str, str]]:
    result: list[tuple[str, str]] = []
    for sec in sections:
        if "업무 수행 내역" not in sec.title:
            continue
        for sub in sec.subsections:
            dm = re.search(r"\((\d{2}/\d{2})\)", sub.title)
            date = dm.group(1) if dm else ""
            for item in sub.items:
                if isinstance(item, TaskItem) and item.status == "취소" and "휴가" in item.text:
                    name = item.text.replace("휴가", "").strip()
                    if name:
                        result.append((name, date))
    return result
