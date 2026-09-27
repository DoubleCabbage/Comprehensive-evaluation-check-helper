"""Shared source-to-normalized data contract."""

SOURCE_COLUMN_MAP: dict[str, str] = {
    "指标名称": "major_category",
    "活动分类": "sub_category",
    "附加分名称": "activity_name",
    "获奖等级": "role_level",
    "参考分数": "reference_score",
}

IGNORED_SOURCE_COLUMNS: tuple[str, ...] = (
    "指标库",
    "所属院系",
    "所属年级",
)

NORMALIZED_COLUMNS: tuple[str, ...] = (
    "record_id",
    "major_category",
    "sub_category",
    "activity_name",
    "role_level",
    "reference_score",
    "source_file",
    "source_sheet",
    "source_row",
)

