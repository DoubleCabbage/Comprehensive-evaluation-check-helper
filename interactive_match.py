"""Interactive local tester that mimics the app's search and dropdown flow."""

import argparse
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from activity_matcher.catalog import (  # noqa: E402
    ActivityNode,
    build_activity_tree,
    load_normalized_records,
)
from activity_matcher.retrieval import MatchQuery, MatchResult, retrieve  # noqa: E402


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Interactive activity matcher")
    parser.add_argument("--catalog", type=Path, help="Normalized historical activity CSV")
    parser.add_argument("--top", type=int, default=5, help="Number of matches to show")
    return parser


def main() -> int:
    _configure_console()
    args = build_parser().parse_args()
    if args.top < 1:
        raise SystemExit("--top 必须至少为 1")

    catalog_path = args.catalog or _choose_catalog()
    if catalog_path is None:
        print("没有找到已清洗的数据。请先运行 run_all_tests.cmd。")
        return 1

    activities = build_activity_tree(load_normalized_records(catalog_path))
    print("\n学生素养分活动匹配：交互测试")
    print(f"已加载 {len(activities)} 个活动节点。")
    print("输入 q 可随时退出。\n")

    while True:
        query = _read("请输入活动名称：").strip()
        if query.lower() in {"q", "quit", "exit"}:
            print("测试结束。")
            return 0
        if not query:
            continue

        results = retrieve(MatchQuery(activity_name=query), activities, limit=args.top)
        _show_results(results)
        selection = _read("选择活动序号展开职务；直接回车输入下一条：").strip()
        if selection.lower() in {"q", "quit", "exit"}:
            print("测试结束。")
            return 0
        if not selection:
            print()
            continue

        result_index = _parse_choice(selection, len(results))
        if result_index is None:
            print("序号无效，继续输入下一条。\n")
            continue

        chosen_activity = results[result_index].activity
        _show_role_options(chosen_activity)
        role_selection = _read("选择职务/等级序号；直接回车返回：").strip()
        if role_selection.lower() in {"q", "quit", "exit"}:
            print("测试结束。")
            return 0
        if role_selection:
            role_index = _parse_choice(role_selection, len(chosen_activity.roles))
            if role_index is None:
                print("序号无效。")
            else:
                role = chosen_activity.roles[role_index]
                print("\n本次选择")
                print(f"活动：{chosen_activity.activity_name}")
                if chosen_activity.variant_label:
                    print(f"变体：{chosen_activity.variant_label}")
                print(f"职务/等级：{role.label or '未填写'}")
                print(f"历史参考分数：{_format_score(role.reference_score)}")
        _read("按回车输入下一条……")
        print()


def _choose_catalog() -> Path | None:
    candidates = _discover_catalogs()
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]

    print("检测到多份已清洗数据：")
    for index, path in enumerate(candidates, start=1):
        run_name = path.parent.name if path.parent.name != "processed" else "默认数据"
        print(f"  [{index}] {run_name}")

    while True:
        choice = _read("请选择数据序号；输入 q 退出：").strip()
        if choice.lower() in {"q", "quit", "exit"}:
            return None
        selected = _parse_choice(choice, len(candidates))
        if selected is not None:
            return candidates[selected]
        print("请输入列表中的有效序号。")


def _discover_catalogs() -> list[Path]:
    candidates = list(
        (PROJECT_ROOT / "data" / "processed" / "runs").glob(
            "*/historical_activities.csv"
        )
    )
    default_catalog = PROJECT_ROOT / "data" / "processed" / "historical_activities.csv"
    if default_catalog.is_file():
        candidates.append(default_catalog)
    return sorted(
        {path.resolve() for path in candidates if path.is_file()},
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def _show_results(results: list[MatchResult]) -> None:
    print("\n匹配结果：")
    if not results:
        print("  没有找到候选活动。\n")
        return

    for index, result in enumerate(results, start=1):
        activity = result.activity
        review_mark = " [待确认变体]" if activity.needs_review else ""
        print(
            f"  [{index}] {activity.activity_name}{review_mark}  "
            f"匹配度 {result.score * 100:.1f}%"
        )
        if activity.variant_label:
            print(f"      {activity.variant_label}")
        category = " / ".join(
            value
            for value in (activity.major_category, activity.sub_category)
            if value
        )
        if category:
            print(f"      分类：{category}")
        role_summary = "；".join(
            f"{role.label or '未填写'} {_format_score(role.reference_score)}"
            for role in activity.roles
        )
        print(f"      可选项：{role_summary or '无'}")
    print()


def _show_role_options(activity: ActivityNode) -> None:
    print(f"\n活动：{activity.activity_name}")
    if activity.variant_label:
        print(f"变体：{activity.variant_label}")
    print("职务/获奖等级下拉选项：")
    for index, role in enumerate(activity.roles, start=1):
        print(
            f"  [{index}] {role.label or '未填写'} → "
            f"{_format_score(role.reference_score)}"
        )


def _parse_choice(value: str, item_count: int) -> int | None:
    try:
        index = int(value) - 1
    except ValueError:
        return None
    return index if 0 <= index < item_count else None


def _format_score(score: float | None) -> str:
    return "未填写" if score is None else f"{score:g} 分"


def _read(prompt: str) -> str:
    try:
        return input(prompt)
    except (EOFError, KeyboardInterrupt):
        print("\n测试结束。")
        raise SystemExit(0) from None


def _configure_console() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
