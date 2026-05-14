#!/usr/bin/env python3
from __future__ import annotations

import re
from pathlib import Path

import click

from src.parser import parse_document, extract_thursday_date
from src.exporter import export_to_xlsx

_BASE_DIR = Path.home() / "Documents" / "크레디프" / "주간보고"
_SOURCE_MD = "업무관리.md"
_LEGACY_SOURCE_MD = "00.md"


def _resolve_input(input_path: Path) -> Path:
    """주간 폴더 또는 기존 00.md 입력에서 실제 변환할 MD 파일 결정."""
    if input_path.is_dir():
        source = input_path / _SOURCE_MD
        if source.exists():
            return source
        legacy_source = input_path / _LEGACY_SOURCE_MD
        if legacy_source.exists():
            return legacy_source
        raise click.ClickException(
            f"{input_path} 안에서 {_SOURCE_MD} 또는 {_LEGACY_SOURCE_MD} 파일을 찾을 수 없습니다."
        )

    if input_path.name == _LEGACY_SOURCE_MD:
        source = input_path.with_name(_SOURCE_MD)
        if source.exists():
            return source

    return input_path


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
    """주간보고 MD 파일 또는 주간 폴더를 XLSX로 변환합니다."""
    source_md = _resolve_input(input_md)
    text = source_md.read_text(encoding="utf-8")
    doc_title, sections = parse_document(text)

    if output_xlsx is None:
        output_xlsx = _resolve_output(doc_title, sections)

    sheets = export_to_xlsx(doc_title, sections, output_xlsx)

    click.echo(f"✓ {output_xlsx}")
    if source_md != input_md:
        click.echo(f"  · 입력: {source_md}")
    for name in sheets:
        click.echo(f"  · {name}")


if __name__ == "__main__":
    main()
