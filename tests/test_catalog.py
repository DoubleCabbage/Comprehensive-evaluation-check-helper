from activity_matcher.catalog import build_activity_tree


def _record(record_id: str, role: str, score: float, row: int) -> dict[str, object]:
    return {
        "record_id": record_id,
        "major_category": "文体素养",
        "sub_category": "校园文化",
        "activity_name": "新年合唱",
        "role_level": role,
        "reference_score": score,
        "source_row": row,
    }


def test_normal_roles_become_dropdown_options() -> None:
    tree = build_activity_tree(
        [_record("a", "演员", 2.0, 2), _record("b", "观众", 0.5, 3)]
    )
    assert len(tree) == 1
    assert {role.label for role in tree[0].roles} == {"演员", "观众"}
    assert tree[0].needs_review is False


def test_conflicting_same_role_scores_split_into_activity_variants() -> None:
    tree = build_activity_tree(
        [_record("a", "演员", 2.0, 2), _record("b", "演员", 3.0, 3)]
    )
    assert len(tree) == 2
    assert all(node.needs_review for node in tree)
    assert {node.roles[0].reference_score for node in tree} == {2.0, 3.0}
