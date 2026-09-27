"""Run reproducible retrieval checks against a normalized activity catalog."""

import argparse
import csv
import json
import re
import sys
from dataclasses import asdict
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from activity_matcher.catalog import build_activity_tree, load_normalized_records  # noqa: E402
from activity_matcher.retrieval import MatchQuery, retrieve  # noqa: E402


_YEAR_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})(年?)")
_EDITION_RE = re.compile(r"第\s*(\d+)\s*届")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run retrieval baseline checks")
    parser.add_argument(
        "--catalog",
        type=Path,
        default=PROJECT_ROOT / "data" / "processed" / "historical_activities.csv",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / "data" / "interim" / "baseline_retrieval_report.json",
    )
    parser.add_argument(
        "--cases",
        type=Path,
        help="Optional CSV containing query and expected_activity columns",
    )
    parser.add_argument("--summary-only", action="store_true")
    parser.add_argument(
        "--fail-on-miss",
        action="store_true",
        help="Return a non-zero exit code when any Top-1 case misses",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    activities = build_activity_tree(load_normalized_records(args.catalog))
    test_cases = _load_cases(args.cases) if args.cases else _build_automatic_cases(activities)

    cases: list[dict[str, object]] = []
    for test_case in test_cases:
        query = str(test_case["query"])
        expected_name = str(test_case["expected_activity"])
        results = retrieve(
            MatchQuery(
                activity_name=query,
                major_category=str(test_case.get("major_category", "")),
                sub_category=str(test_case.get("sub_category", "")),
            ),
            activities,
            limit=3,
        )
        result_names = [result.activity.activity_name for result in results]
        cases.append(
            {
                **test_case,
                "hit_at_1": bool(result_names and result_names[0] == expected_name),
                "hit_at_3": expected_name in result_names,
                "results": [asdict(result) for result in results],
            }
        )

    report = {
        "mode": "labeled_cases" if args.cases else "automatic_smoke_cases",
        "catalog": str(args.catalog),
        "activity_node_count": len(activities),
        "test_case_count": len(cases),
        "top_1_hits": sum(bool(case["hit_at_1"]) for case in cases),
        "top_3_hits": sum(bool(case["hit_at_3"]) for case in cases),
        "cases": cases,
    }
    report["top_1_misses"] = report["test_case_count"] - report["top_1_hits"]
    report["top_3_misses"] = report["test_case_count"] - report["top_3_hits"]
    report["status"] = "pass" if report["top_1_misses"] == 0 else "needs_review"
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    summary = {
        key: report[key]
        for key in (
            "mode",
            "activity_node_count",
            "test_case_count",
            "top_1_hits",
            "top_3_hits",
            "top_1_misses",
            "top_3_misses",
            "status",
        )
    }
    print(json.dumps(summary if args.summary_only else report, ensure_ascii=False, indent=2))
    if args.fail_on_miss and report["top_1_misses"]:
        return 1
    return 0


def _load_cases(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as source:
        rows = list(csv.DictReader(source))
    required = {"query", "expected_activity"}
    if not rows:
        raise ValueError(f"测试用例文件为空: {path}")
    if not required.issubset(rows[0]):
        missing = sorted(required - set(rows[0]))
        raise ValueError(f"测试用例文件缺少列: {', '.join(missing)}")
    return rows


def _build_automatic_cases(activities: list[object]) -> list[dict[str, str]]:
    unique: dict[tuple[str, str, str], object] = {}
    for activity in activities:
        key = (
            activity.major_category,
            activity.sub_category,
            activity.activity_name,
        )
        unique.setdefault(key, activity)

    return [
        {
            "query": _make_smoke_query(activity.activity_name),
            "expected_activity": activity.activity_name,
            "major_category": activity.major_category,
            "sub_category": activity.sub_category,
        }
        for activity in unique.values()
    ]


def _make_smoke_query(activity_name: str) -> str:
    year_match = _YEAR_RE.search(activity_name)
    if year_match:
        replacement = f"{int(year_match.group(1)) + 1}{year_match.group(2)}"
        return _YEAR_RE.sub(replacement, activity_name, count=1)
    edition_match = _EDITION_RE.search(activity_name)
    if edition_match:
        replacement = f"第{int(edition_match.group(1)) + 1}届"
        return _EDITION_RE.sub(replacement, activity_name, count=1)
    return activity_name


if __name__ == "__main__":
    raise SystemExit(main())
