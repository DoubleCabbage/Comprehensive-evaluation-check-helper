"""Minimal test runner used when pytest is unavailable offline."""

import ast
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from activity_matcher.catalog import build_activity_tree  # noqa: E402
from activity_matcher.normalize import (  # noqa: E402
    normalize_activity_name,
    normalize_reference_score,
    normalize_text,
)
from activity_matcher.retrieval import MatchQuery, retrieve  # noqa: E402


def main() -> int:
    python_files = (
        list((PROJECT_ROOT / "src").glob("**/*.py"))
        + list((PROJECT_ROOT / "tests").glob("*.py"))
        + list((PROJECT_ROOT / "experiments").glob("*.py"))
        + [PROJECT_ROOT / "run.py"]
    )
    for path in python_files:
        ast.parse(path.read_text(encoding="utf-8"), filename=str(path))

    assert normalize_text("  Ａ　 B  ") == "A B"
    assert normalize_activity_name("1-测试活动") == "测试活动"
    assert normalize_reference_score("10") == 10.0

    records = [
        _record("a", "新年合唱", "演员", 2.0, 2),
        _record("b", "新年合唱", "观众", 0.5, 3),
        _record("c", "物理竞赛", "一等奖", 14.0, 4),
    ]
    tree = build_activity_tree(records)
    choir = next(node for node in tree if node.activity_name == "新年合唱")
    assert {role.label for role in choir.roles} == {"演员", "观众"}
    result = retrieve(MatchQuery(activity_name="新年合唱晚会"), tree, limit=1)
    assert result[0].activity.activity_name == "新年合唱"

    conflicting = build_activity_tree(
        [
            _record("x", "创新训练", "获评项目", 10.0, 5),
            _record("y", "创新训练", "获评项目", 8.0, 6),
        ]
    )
    assert len(conflicting) == 2
    assert all(node.needs_review for node in conflicting)

    print(f"offline_smoke_checks_ok files={len(python_files)}")
    return 0


def _record(
    record_id: str,
    activity_name: str,
    role_level: str,
    score: float,
    source_row: int,
) -> dict[str, object]:
    return {
        "record_id": record_id,
        "major_category": "测试大类",
        "sub_category": "测试小类",
        "activity_name": activity_name,
        "role_level": role_level,
        "reference_score": score,
        "source_row": source_row,
    }


if __name__ == "__main__":
    raise SystemExit(main())
