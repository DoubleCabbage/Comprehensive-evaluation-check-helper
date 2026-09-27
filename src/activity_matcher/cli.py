"""Command-line entry point reserved for the PC prototype."""

import argparse
from dataclasses import asdict
import json
import sys
from pathlib import Path

from .catalog import build_activity_tree, load_normalized_records
from .import_excel import inspect_workbook
from .pipeline import clean_workbook
from .retrieval import MatchQuery, retrieve


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="activity-matcher",
        description="Offline historical activity-score matcher prototype",
    )
    subparsers = parser.add_subparsers(dest="command")

    inspect_parser = subparsers.add_parser("inspect", help="检查 Excel 结构")
    inspect_parser.add_argument("source", type=Path)

    clean_parser = subparsers.add_parser("clean", help="清洗 Excel 并输出 CSV 和报告")
    clean_parser.add_argument("source", type=Path)
    clean_parser.add_argument(
        "--output",
        type=Path,
        default=Path("data/processed/historical_activities.csv"),
    )
    clean_parser.add_argument(
        "--report",
        type=Path,
        default=Path("data/interim/import_report.json"),
    )
    clean_parser.add_argument(
        "--tree",
        type=Path,
        default=Path("data/processed/activity_tree.json"),
    )

    match_parser = subparsers.add_parser("match", help="运行文字匹配基线")
    match_parser.add_argument("query")
    match_parser.add_argument(
        "--catalog",
        type=Path,
        default=Path("data/processed/historical_activities.csv"),
    )
    match_parser.add_argument("--major-category", default="")
    match_parser.add_argument("--sub-category", default="")
    match_parser.add_argument("--top", type=int, default=3)
    return parser


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = build_parser()
    args = parser.parse_args()
    if args.command is None:
        parser.print_help()
        return 0
    if args.command == "inspect":
        print(json.dumps(inspect_workbook(args.source), ensure_ascii=False, indent=2))
        return 0
    if args.command == "clean":
        report = clean_workbook(
            args.source,
            processed_path=args.output,
            report_path=args.report,
            tree_path=args.tree,
        )
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0
    if args.command == "match":
        activities = build_activity_tree(load_normalized_records(args.catalog))
        results = retrieve(
            MatchQuery(
                activity_name=args.query,
                major_category=args.major_category,
                sub_category=args.sub_category,
            ),
            activities,
            limit=args.top,
        )
        print(json.dumps([asdict(result) for result in results], ensure_ascii=False, indent=2))
        return 0
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
