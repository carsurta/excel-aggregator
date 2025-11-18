from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from textwrap import indent


ROOT = Path(__file__).resolve().parent
PLAN_FILE = ROOT / "requirements_test_cases.md"
SECTION_RE = re.compile(r"^##\s+([A-Z0-9\-]+)\s+-\s+(.+)$", re.MULTILINE)


def _load_sections() -> list[dict]:
    if not PLAN_FILE.exists():
        raise FileNotFoundError(f"테스트 케이스 문서를 찾을 수 없습니다: {PLAN_FILE}")
    text = PLAN_FILE.read_text(encoding="utf-8")
    matches = list(SECTION_RE.finditer(text))
    sections: list[dict] = []
    for idx, match in enumerate(matches):
        start = match.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        body = text[start:end].strip()
        sections.append(
            {
                "id": match.group(1),
                "title": match.group(2).strip(),
                "body": body,
            }
        )
    return sections


def list_sections(sections: list[dict]) -> None:
    print("등록된 요구사항 ID:")
    for sec in sections:
        print(f" - {sec['id']}: {sec['title']}")


def show_section(sections: list[dict], requirement_id: str) -> None:
    for sec in sections:
        if sec["id"].lower() == requirement_id.lower():
            print(f"## {sec['id']} - {sec['title']}")
            print(sec["body"])
            return
    print(f"[WARN] ID '{requirement_id}' 를 찾을 수 없습니다.", file=sys.stderr)
    list_sections(sections)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="requirements.md 기반 테스트 케이스 출력 도구"
    )
    parser.add_argument(
        "--id",
        help="특정 요구사항 ID만 출력 (예: FR-FILE-01)",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="요구사항 ID 목록만 표시",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv or sys.argv[1:])
    sections = _load_sections()
    if args.list:
        list_sections(sections)
        return
    if args.id:
        show_section(sections, args.id)
        return

    print(
        "모든 요구사항과 테스트 케이스는 tests/requirements_test_cases.md 파일에 정리되어 있습니다.\n"
        "자주 사용하는 명령 예시:\n"
        "  python tests/run_requirements_suite.py --list\n"
        "  python tests/run_requirements_suite.py --id FR-FILE-01\n"
        "위 명령으로 필요한 케이스를 바로 확인한 뒤 수동 테스트를 진행하세요.\n"
    )
    list_sections(sections)


if __name__ == "__main__":
    main()
