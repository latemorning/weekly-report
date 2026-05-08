# Repository Guidelines

## Project Structure & Module Organization

This repository is a Python CLI tool that converts Obsidian-style weekly report Markdown into Excel workbooks.

- `convert.py`: main Click CLI entry point.
- `src/parser.py`: parses Markdown into dataclass-based report models.
- `src/exporter.py`: builds and styles XLSX output with `openpyxl`.
- `samples/`: sample Markdown and workbook inputs for manual validation.
- `output/`: generated workbook output; avoid committing ad hoc generated files unless they are intentional fixtures.
- `pointhub_weekly_report_sample_1.xlsx`: reference workbook template.

## Build, Test, and Development Commands

- `uv sync`: install dependencies from `pyproject.toml` and `uv.lock`.
- `uv run python convert.py samples/sample.md`: convert a sample report and write to the default dated output path.
- `uv run python convert.py samples/sample.md output/report.xlsx`: convert a sample report to an explicit output file.
- `uv run python main.py`: run the placeholder entry point; currently prints a greeting only.

There is no packaging or build command configured yet.

## Coding Style & Naming Conventions

Use Python 3.9+ syntax and four-space indentation. Keep type annotations on public helpers and dataclass fields, following the existing style in `src/parser.py`. Use `snake_case` for functions and variables, `PascalCase` for dataclasses, `UPPER_CASE` for constants, and a leading underscore for private helpers. Korean sheet names, labels, and domain comments are expected where they match the report format.

## Testing Guidelines

No automated test suite is currently present. For changes to parsing or Excel rendering, manually validate with:

```bash
uv run python convert.py samples/sample.md output/report.xlsx
```

When adding tests, place them under `tests/`, name files `test_*.py`, and prefer focused parser tests plus workbook assertions for exporter behavior. Add `pytest` to the project before relying on `uv run pytest`.

## Commit & Pull Request Guidelines

The current `develop` branch has no commits, so no project-specific commit convention is established. Use concise imperative messages such as `Add parser tests` or `Fix issue sheet assignee export`.

Pull requests should include a short summary, manual validation command and result, linked issue if applicable, and screenshots or attached XLSX samples when output layout changes.

## Agent-Specific Instructions

Keep generated workbook changes out of commits unless requested. Preserve user edits in the worktree.

Issue sheet layout is fixed: development issues only use `D3`/`E3`; operations issues only use `D5`/`E5`. Never mix development and operations content or repeat `개발부문`/`운영부문` inside those cells. Keep issue subsection titles in cell text, and write vacation output as one row from weekly report vacation content.
