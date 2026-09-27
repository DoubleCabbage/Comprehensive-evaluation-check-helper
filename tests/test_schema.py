from activity_matcher.schema import NORMALIZED_COLUMNS, SOURCE_COLUMN_MAP


def test_required_source_columns_are_mapped() -> None:
    assert SOURCE_COLUMN_MAP == {
        "指标名称": "major_category",
        "活动分类": "sub_category",
        "附加分名称": "activity_name",
        "获奖等级": "role_level",
        "参考分数": "reference_score",
    }


def test_normalized_schema_keeps_source_provenance() -> None:
    assert {"source_file", "source_sheet", "source_row"}.issubset(NORMALIZED_COLUMNS)

