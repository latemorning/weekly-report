# 저장소 지침

## 프로젝트 구조 및 모듈 구성

이 저장소는 Obsidian 스타일의 주간 보고서 Markdown을 Excel 통합 문서로 변환하는 Python CLI 도구입니다.

- `convert.py`: 기본 Click CLI 진입점입니다.
- `src/parser.py`: Markdown을 dataclass 기반 보고서 모델로 파싱합니다.
- `src/exporter.py`: `openpyxl`로 XLSX 출력을 생성하고 스타일을 적용합니다.
- `samples/`: 수동 검증에 사용하는 샘플 Markdown 및 통합 문서 입력 파일입니다.
- `output/`: 생성된 통합 문서 출력 위치입니다. 의도한 fixture가 아니라면 임의 생성 파일은 커밋하지 마세요.
- `pointhub_weekly_report_sample_1.xlsx`: 참조용 통합 문서 템플릿입니다.

## 빌드, 테스트, 개발 명령

- `uv sync`: `pyproject.toml` 및 `uv.lock`의 의존성을 설치합니다.
- `uv run python convert.py samples/sample.md`: 샘플 보고서를 변환하고 기본 날짜 기반 출력 경로에 저장합니다.
- `uv run python convert.py /path/to/week-folder`: 주간 폴더 안의 `업무관리.md`를 변환하고 기본 날짜 기반 출력 경로에 저장합니다.
- `uv run python convert.py samples/sample.md output/report.xlsx`: 샘플 보고서를 명시한 출력 파일로 변환합니다.
- `uv run python main.py`: placeholder 진입점을 실행합니다. 현재는 인사말만 출력합니다.

아직 패키징 또는 빌드 명령은 구성되어 있지 않습니다.

## 코딩 스타일 및 명명 규칙

Python 3.9+ 문법과 네 칸 들여쓰기를 사용하세요. `src/parser.py`의 기존 스타일을 따라 공개 helper와 dataclass 필드에는 타입 주석을 유지하세요. 함수와 변수는 `snake_case`, dataclass는 `PascalCase`, 상수는 `UPPER_CASE`, private helper는 앞에 밑줄을 붙이는 방식을 사용합니다. 보고서 형식과 맞는 경우 한국어 시트명, 레이블, 도메인 주석을 사용해도 됩니다.

## 테스트 지침

현재 자동화된 테스트 스위트는 없습니다. 파싱 또는 Excel 렌더링을 변경한 경우 다음 명령으로 수동 검증하세요.

```bash
uv run python convert.py samples/sample.md output/report.xlsx
uv run python convert.py samples/업무관리.md output/report.xlsx
```

테스트를 추가할 때는 `tests/` 아래에 두고 파일명은 `test_*.py`로 지정하세요. exporter 동작은 focused parser test와 workbook assertion을 우선 사용하세요. `uv run pytest`에 의존하기 전에 프로젝트에 `pytest`를 추가하세요.

## 커밋 및 Pull Request 지침

현재 `develop` 브랜치에는 커밋이 없으므로 프로젝트 고유의 커밋 규칙은 아직 정해져 있지 않습니다. `Add parser tests` 또는 `Fix issue sheet assignee export`처럼 간결한 명령형 메시지를 사용하세요.

Pull Request에는 짧은 요약, 수동 검증 명령과 결과, 관련 issue가 있다면 링크, 출력 레이아웃이 바뀐 경우 스크린샷 또는 첨부 XLSX 샘플을 포함하세요.

## Agent 전용 지침

요청이 없는 한 생성된 통합 문서 변경 사항은 커밋에 포함하지 마세요. 작업 트리의 사용자 수정 사항을 보존하세요.

이슈 시트 레이아웃은 고정되어 있습니다. 개발 이슈는 `D3`/`E3`만 사용하고, 운영 이슈는 `D5`/`E5`만 사용합니다. 해당 셀 안에서 개발과 운영 내용을 섞거나 `개발부문`/`운영부문`을 반복하지 마세요. 이슈 하위 섹션 제목은 셀 텍스트 안에 유지하고, 휴가 출력은 주간 보고서의 휴가 내용을 한 행으로 작성하세요.
