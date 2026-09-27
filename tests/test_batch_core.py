from pathlib import Path
from zipfile import ZipFile
import json

from openpyxl import Workbook, load_workbook

from activity_matcher.batch_core import (
    FormatConfig, Reference, Submission, canonical_field, category_key,
    clean_category_label, field_key,
    export_review_copies, load_alias_settings,
    load_references, load_submissions, rank, restore_progress, save_alias_settings,
    save_progress, status,
)
import pytest


def make_files(tmp_path: Path):
    reference = Workbook()
    for sheet_name, name, score in [("学术创新", "大学生创新创业训练计划结题", 10),
                                    ("身心健康", "校园体育活动", 2)]:
        sheet = reference.active if sheet_name == "学术创新" else reference.create_sheet()
        sheet.title = sheet_name
        sheet.append(["指标库", "所属院系", "所属年级", "指标名称", "活动分类", "附加分名称", "获奖等级", "参考分数"])
        sheet.append(["模拟", "电子工程学院", 2025, sheet_name, "小类", name, "获评项目", score])
    reference_path = tmp_path / "reference.xlsx"
    reference.save(reference_path)

    source = Workbook()
    summary = source.active
    summary.title = "总表"
    summary.append(["示例"])
    summary.append(["指标", "序号", "活动名称", "加分等级", "加分分值", "备注"])
    summary.append(["学术创新（14分）", 1, "1-大学生创新创业训练计划结题", "获评项目", 10, "原备注"])
    summary.append([None, 2, "创业训练", "获评项目", 3, None])
    summary.append(["身心健康（10分）", 1, "校园体育活动", "获评项目", 3, None])
    for name in ["学术创新活动证明", "身心健康活动证明"]:
        proof = source.create_sheet(name)
        proof.append(["序号", "证明材料"])
        proof.append([1, None])
        proof.append([2, None])
    source_path = tmp_path / "student.xlsx"
    source.save(source_path)
    return reference_path, source_path


def test_batch_match_progress_and_marked_copy(tmp_path: Path):
    reference_path, source_path = make_files(tmp_path)
    config = FormatConfig()
    references = load_references(reference_path, config)
    items = load_submissions([source_path], config)
    assert len(references) == 2
    assert len(items) == 3
    assert [status(item, references)[0] for item in items] == ["matched", "pending", "wrong"]
    assert items[0].image_warning == "对应序号没有可读取的单元格图片"
    assert (items[0].evidence_sheet, items[0].evidence_row) == ("学术创新活动证明", 2)
    assert (items[1].evidence_sheet, items[1].evidence_row) == ("学术创新活动证明", 3)
    assert rank(items[1], references)

    items[1].decision = ""
    items[0].proof = "insufficient"
    progress = tmp_path / "progress.json"
    save_progress(progress, reference_path, config, items)
    restored_ref, restored_config, restored = restore_progress(progress)
    assert restored_ref == reference_path
    assert restored_config == config
    assert restored[1].decision == ""
    assert restored[0].proof == "insufficient"

    original = source_path.read_bytes()
    result = export_review_copies(tmp_path / "copies", config, items, references)[0]
    second = export_review_copies(tmp_path / "copies", config, items, references)[0]
    assert second != result and result.exists()
    assert source_path.read_bytes() == original
    with ZipFile(result) as archive:
        assert archive.testzip() is None
    reviewed = load_workbook(result, read_only=True)
    try:
        assert reviewed["总表"]["F3"].value == "原备注；证明不足"
        assert reviewed["总表"]["C3"].value == "1-大学生创新创业训练计划结题"
        assert reviewed["总表"]["C3"].fill.fgColor.rgb == "FFE2F0D9"
        assert reviewed["总表"]["C4"].fill.fgColor.rgb == "FFF4CCCC"
        assert reviewed["总表"]["F3"].fill.fgColor.rgb == "FFFFF2CC"
    finally:
        reviewed.close()


def test_text_reference_score_is_a_separate_review_case(tmp_path: Path):
    reference_path, source_path = make_files(tmp_path)
    reference = load_workbook(reference_path)
    reference["身心健康"].append(["模拟", "电子工程学院", 2025, "身心健康",
        "体育锻炼", "校园跑步活动", "参与", "评审认定"])
    reference.save(reference_path)
    source = load_workbook(source_path)
    source["总表"].append([None, 2, "校园跑步活动", "参与", 2, None])
    source.save(source_path)

    config = FormatConfig()
    references = load_references(reference_path, config)
    items = load_submissions([source_path], config)
    item = items[-1]
    assert references[-1].score == "评审认定"
    assert status(item, references) == ("review_score", references[-1])
    assert "评审认定" in rank(item, references)[0].differences
    item.decision = references[-1].id
    assert status(item, references)[0] == "review_score"

    copy = export_review_copies(tmp_path / "copies", config, items, references)[0]
    reviewed = load_workbook(copy, read_only=True)
    try:
        assert reviewed["总表"]["F6"].value == "参考分数需人工认定：评审认定"
        assert reviewed["总表"]["C6"].fill.fgColor.rgb == "FFFCE4D6"
    finally:
        reviewed.close()


def test_persistent_category_alias_ignores_configured_difference(tmp_path: Path):
    reference_path, source_path = make_files(tmp_path)
    source = load_workbook(source_path)
    source["总表"]["A5"] = "身心健康活动成绩（10分）"
    source["总表"]["E5"] = 2
    source.save(source_path)
    groups = {"major": [["身心健康", "身心健康活动成绩", "身心健康成绩"]]}
    settings = tmp_path / "batch_review_settings.json"
    save_alias_settings(settings, groups)
    assert load_alias_settings(settings) == groups
    assert canonical_field("身心健康活动成绩（10分）", "major", groups) == canonical_field("身心健康", "major", groups)
    assert clean_category_label("身心健康（满分 10 分）") == "身心健康"
    assert category_key("身心健康(12.5 分)") == "身心健康"

    references = load_references(reference_path, FormatConfig())
    without_rule = load_submissions([source_path], FormatConfig())[-1]
    assert status(without_rule, references)[0] == "pending"
    config = FormatConfig(alias_groups=load_alias_settings(settings))
    item = load_submissions([source_path], config)[-1]
    assert item.major == "身心健康活动成绩"
    assert item.image_warning == "对应序号没有可读取的单元格图片"
    assert status(item, references, alias_groups=config.alias_groups)[0] == "matched"
    assert "大类" not in rank(item, references, alias_groups=config.alias_groups)[0].differences

    progress = tmp_path / "progress.json"
    save_progress(progress, reference_path, config, [item])
    _, loaded_config, loaded_items = restore_progress(progress, alias_groups=groups)
    assert loaded_config.alias_groups == groups
    assert status(loaded_items[0], references, alias_groups=loaded_config.alias_groups)[0] == "matched"

    copy = export_review_copies(tmp_path / "copies", config, [item], references)[0]
    reviewed = load_workbook(copy, read_only=True)
    try:
        assert reviewed["总表"]["C5"].fill.fgColor.rgb == "FFE2F0D9"
    finally:
        reviewed.close()


def test_cross_field_groups_with_three_names_each(tmp_path: Path):
    reference_path, source_path = make_files(tmp_path)
    reference = load_workbook(reference_path)
    reference["学术创新"]["E2"] = "创新创业"
    reference.save(reference_path)
    source = load_workbook(source_path)
    summary = source["总表"]
    summary["A3"] = "学术创造（14分）"
    summary["C3"] = "1-双创训练结项"
    summary["D3"] = "项目通过"
    summary["G3"] = "创业创新（4分）"
    source.save(source_path)
    groups = {
        "major": [["学术创新", "学术创造", "科创"]],
        "minor": [["创新创业", "创业创新", "双创"]],
        "name": [["大学生创新创业训练计划结题", "双创训练结项", "大创项目结题"]],
        "role": [["获评项目", "项目通过", "结题项目"]],
    }
    config = FormatConfig(summary_minor=6, alias_groups=groups)
    references = load_references(reference_path, config)
    item = load_submissions([source_path], config)[0]
    assert item.major == "学术创造"
    assert item.minor == "创业创新"
    assert item.image_warning == "对应序号没有可读取的单元格图片"
    assert status(item, references)[0] == "pending"
    assert status(item, references, alias_groups=groups)[0] == "matched"
    assert rank(item, references, alias_groups=groups)[0].differences == "一致"


def test_invalid_or_overlapping_alias_groups_rejected(tmp_path: Path):
    with pytest.raises(ValueError, match="多个别名组"):
        save_alias_settings(tmp_path / "settings.json",
            {"major": [["身心健康", "体育"], ["体育", "身体健康"]]})


def test_version_one_alias_config_migrates_without_rewriting(tmp_path: Path):
    path = tmp_path / "settings.json"
    payload = {"version": 1, "category_aliases": {"身心健康活动成绩": "身心健康"}}
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    assert load_alias_settings(path) == {"major": [["身心健康", "身心健康活动成绩"]]}
    assert json.loads(path.read_text(encoding="utf-8")) == payload


def test_grade_plus_is_not_the_same_as_plain_grade(tmp_path: Path):
    item = Submission(tmp_path / "student.xlsx", 59, "组织协调能力", "", 1,
        "院级学生组织", "第三档A", 3.4)
    references = [
        Reference("组织协调能力", 82, "组织协调能力", "", "院级学生组织", "第三档A+", 4.0),
        Reference("组织协调能力", 83, "组织协调能力", "", "院级学生组织", "第三档A", 3.4),
    ]
    assert field_key("第三档A+", "role") != field_key("第三档A", "role")
    assert field_key("第三档Ａ＋", "role") == field_key("第三档A+", "role")
    assert status(item, references) == ("matched", references[1])
    matches = rank(item, references)
    assert matches[0].target == references[1]
    assert matches[0].similarity == 1.0
    assert matches[0].differences == "一致"
    assert matches[1].similarity < 1.0
    assert "等级" in matches[1].differences
    assert "分数" in matches[1].differences
