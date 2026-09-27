"""Small, reproducible Excel-to-normalized-data pipeline."""

import csv
import json
from collections import defaultdict
from pathlib import Path

from .catalog import build_activity_tree, save_activity_tree
from .import_excel import inspect_workbook, iter_source_rows
from .normalize import normalize_activity_name, normalize_reference_score, normalize_text
from .schema import NORMALIZED_COLUMNS


def clean_workbook(
    source_path: Path,
    *,
    processed_path: Path,
    report_path: Path,
    tree_path: Path | None = None,
) -> dict[str, object]:
    """Normalize one workbook, preserving rows and reporting ambiguous duplicates."""
    records: list[dict[str, object]] = []
    issues: list[dict[str, object]] = []

    for sheet_name, row_number, source in iter_source_rows(source_path):
        try:
            reference_score = normalize_reference_score(source.get("参考分数"))
        except ValueError as error:
            reference_score = None
            issues.append(
                {
                    "type": "invalid_reference_score",
                    "sheet": sheet_name,
                    "row": row_number,
                    "message": str(error),
                }
            )

        record = {
            "record_id": f"{source_path.stem}:{sheet_name}:{row_number}",
            "major_category": normalize_text(source.get("指标名称")),
            "sub_category": normalize_text(source.get("活动分类")),
            "activity_name": normalize_activity_name(source.get("附加分名称")),
            "role_level": normalize_text(source.get("获奖等级")),
            "reference_score": reference_score,
            "source_file": source_path.name,
            "source_sheet": sheet_name,
            "source_row": row_number,
        }
        if not record["activity_name"]:
            issues.append(
                {
                    "type": "missing_activity_name",
                    "sheet": sheet_name,
                    "row": row_number,
                }
            )
        records.append(record)

    conflict_groups = _find_score_conflicts(records)
    issues.extend(conflict_groups)

    processed_path.parent.mkdir(parents=True, exist_ok=True)
    with processed_path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=NORMALIZED_COLUMNS)
        writer.writeheader()
        writer.writerows(records)

    resolved_tree_path = tree_path or processed_path.with_name("activity_tree.json")
    activity_tree = build_activity_tree(records)
    save_activity_tree(activity_tree, resolved_tree_path)

    report = {
        "source": str(source_path),
        "workbook": inspect_workbook(source_path),
        "records_written": len(records),
        "issue_count": len(issues),
        "issues": issues,
        "output": str(processed_path),
        "activity_nodes_written": len(activity_tree),
        "tree_output": str(resolved_tree_path),
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return report


def _find_score_conflicts(records: list[dict[str, object]]) -> list[dict[str, object]]:
    groups: dict[tuple[str, str, str, str], list[dict[str, object]]] = defaultdict(list)
    for record in records:
        key = (
            str(record["major_category"]),
            str(record["sub_category"]),
            str(record["activity_name"]),
            str(record["role_level"]),
        )
        groups[key].append(record)

    conflicts: list[dict[str, object]] = []
    for key, group in groups.items():
        scores = sorted(
            {record["reference_score"] for record in group if record["reference_score"] is not None}
        )
        if len(scores) > 1:
            conflicts.append(
                {
                    "type": "conflicting_duplicate_scores",
                    "key": {
                        "major_category": key[0],
                        "sub_category": key[1],
                        "activity_name": key[2],
                        "role_level": key[3],
                    },
                    "scores": scores,
                    "source_rows": [record["source_row"] for record in group],
                }
            )
    return conflicts
