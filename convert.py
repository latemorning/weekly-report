#!/usr/bin/env python3
from __future__ import annotations

import re
from pathlib import Path

import click

from src.parser import parse_document, extract_thursday_date
from src.exporter import export_to_xlsx

_BASE_DIR = Path.home() / "Documents" / "크레디프" / "주간보고"


def _resolve_output(doc_title: str, sections) -> Path:
    """목요일 날짜 기준으로 출력 경로 결정."""
    year_m = re.search(r"(\d{4})년", doc_title)
    year = year_m.group(1) if year_m else "2026"

    thursday = extract_thursday_date(sections)  # "MM/DD"
    if thursday:
        month, day = thursday.split("/")
        filename = f"포인트허브_주간보고_{year}{month}{day}.xlsx"
    else:
        filename = "포인트허브_주간보고.xlsx"
        month = "01"

    out_dir = _BASE_DIR / year / month
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / filename


@click.command()
@click.argument("input_md", type=click.Path(exists=True, path_type=Path))
@click.argument("output_xlsx", type=click.Path(path_type=Path), required=False)
def main(input_md: Path, output_xlsx: Path | None) -> None:
    """주간보고 MD 파일을 XLSX로 변환합니다."""
    text = input_md.read_text(encoding="utf-8")
    doc_title, sections = parse_document(text)

    if output_xlsx is None:
        output_xlsx = _resolve_output(doc_title, sections)

    sheets = export_to_xlsx(doc_title, sections, output_xlsx)

    click.echo(f"✓ {output_xlsx}")
    for name in sheets:
        click.echo(f"  · {name}")


if __name__ == "__main__":
    main()
