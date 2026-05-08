from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Union

CHECKBOX_STATUS: dict[str, str] = {
    "x": "완료",
    " ": "미완",
    "-": "취소",
    "/": "처리",
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


def parse_document(text: str) -> tuple[str, list[Section]]:
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
