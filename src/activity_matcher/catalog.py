"""Build the activity -> role/level -> reference-score tree used by the UI."""

import csv
import hashlib
import json
from collections import defaultdict
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable


@dataclass(frozen=True, slots=True)
class RoleOption:
    label: str
    reference_score: float | None
    record_ids: tuple[str, ...]
    source_rows: tuple[int, ...]


@dataclass(frozen=True, slots=True)
class ActivityNode:
    activity_id: str
    activity_name: str
    major_category: str
    sub_category: str
    variant_label: str
    needs_review: bool
    roles: tuple[RoleOption, ...]


def load_normalized_records(path: Path) -> list[dict[str, object]]:
    """Load the normalized CSV while restoring numeric fields."""
    with path.open(encoding="utf-8-sig", newline="") as source:
        rows = list(csv.DictReader(source))

    records: list[dict[str, object]] = []
    for row in rows:
        records.append(
            {
                **row,
                "reference_score": (
                    float(row["reference_score"]) if row["reference_score"] else None
                ),
                "source_row": int(row["source_row"]),
            }
        )
    return records


def build_activity_tree(records: Iterable[dict[str, object]]) -> list[ActivityNode]:
    """Group ordinary role options and split conflicting same-role scores into variants."""
    activity_groups: dict[tuple[str, str, str], list[dict[str, object]]] = defaultdict(list)
    for record in records:
        key = (
            str(record["major_category"]),
            str(record["sub_category"]),
            str(record["activity_name"]),
        )
        activity_groups[key].append(record)

    nodes: list[ActivityNode] = []
    for key, group in activity_groups.items():
        scores_by_role: dict[str, set[float | None]] = defaultdict(set)
        for record in group:
            scores_by_role[str(record["role_level"])].add(record["reference_score"])
        has_conflict = any(len(scores) > 1 for scores in scores_by_role.values())

        if has_conflict:
            sorted_group = sorted(
                group,
                key=lambda row: int(row["source_row"]),
            )
            for variant_number, record in enumerate(sorted_group, start=1):
                score = record["reference_score"]
                nodes.append(
                    ActivityNode(
                        activity_id=f"variant:{record['record_id']}",
                        activity_name=key[2],
                        major_category=key[0],
                        sub_category=key[1],
                        variant_label=(
                            f"待确认变体 {variant_number}（{record['role_level']}，"
                            f"{_format_score(score)}分）"
                        ),
                        needs_review=True,
                        roles=(
                            RoleOption(
                                label=str(record["role_level"]),
                                reference_score=score,
                                record_ids=(str(record["record_id"]),),
                                source_rows=(int(record["source_row"]),),
                            ),
                        ),
                    )
                )
            continue

        role_groups: dict[tuple[str, float | None], list[dict[str, object]]] = defaultdict(list)
        for record in group:
            role_groups[(str(record["role_level"]), record["reference_score"])].append(record)
        roles = tuple(
            RoleOption(
                label=role,
                reference_score=score,
                record_ids=tuple(str(row["record_id"]) for row in rows),
                source_rows=tuple(int(row["source_row"]) for row in rows),
            )
            for (role, score), rows in sorted(
                role_groups.items(),
                key=lambda item: (
                    item[0][1] is None,
                    -float(item[0][1]) if item[0][1] is not None else 0.0,
                    min(int(row["source_row"]) for row in item[1]),
                    item[0][0],
                ),
            )
        )
        digest = hashlib.sha1("\x1f".join(key).encode("utf-8")).hexdigest()[:12]
        nodes.append(
            ActivityNode(
                activity_id=f"activity:{digest}",
                activity_name=key[2],
                major_category=key[0],
                sub_category=key[1],
                variant_label="",
                needs_review=False,
                roles=roles,
            )
        )

    return sorted(
        nodes,
        key=lambda node: (
            node.major_category,
            node.sub_category,
            node.activity_name,
            node.variant_label,
        ),
    )


def save_activity_tree(nodes: Iterable[ActivityNode], path: Path) -> None:
    """Save a UI-ready tree as UTF-8 JSON."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps([asdict(node) for node in nodes], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _format_score(score: object) -> str:
    if score is None:
        return "未填写"
    number = float(score)
    return f"{number:g}"
