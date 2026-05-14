Obsidian에 등록한 주간 업무 내용을 보고서용 Excel 파일로 변환하는 프로젝트입니다.

기본 입력은 각 주간 폴더의 `업무관리.md`입니다. 폴더 경로를 넘기면 내부의 `업무관리.md`를 찾아 변환하고, 기존처럼 `00.md`를 넘긴 경우에도 같은 폴더에 `업무관리.md`가 있으면 그 파일을 우선 사용합니다.

```bash
uv run python convert.py /path/to/week-folder
uv run python convert.py /path/to/week-folder/업무관리.md output/report.xlsx
```
