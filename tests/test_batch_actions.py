from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook

from activity_matcher.batch_core import (
    FormatConfig, Reference, add_equivalence_group, apply_edits_to_original,
    backup_workbooks, file_fingerprint, load_references, load_submissions,
    propose_original_edits, restore_progress, save_progress, text_differences,
)
from test_batch_core import make_files


def test_backup_every_workbook_without_overwriting(tmp_path: Path):
    reference, source = make_files(tmp_path)
    destination = tmp_path / "backups"
    first = backup_workbooks([reference, source], destination)
    second = backup_workbooks([reference, source], destination)
    assert len(first) == len(second) == 2
    assert set(first).isdisjoint(second)
    assert first[0].read_bytes() == second[0].read_bytes() == reference.read_bytes()
    assert first[1].read_bytes() == second[1].read_bytes() == source.read_bytes()


def test_quick_ignore_builds_a_group_and_keeps_score_difference(tmp_path: Path):
    reference, source = make_files(tmp_path)
    config = FormatConfig()
    item = load_submissions([source], config)[0]
    target = Reference("学术创新", 9, "学术创新", "小类", "双创训练结项", "获评项目", 8)
    assert text_differences(item, target, config) == [
        ("name", "大学生创新创业训练计划结题", "双创训练结项")]
    groups = add_equivalence_group({}, "name", item.name, target.name)
    groups = add_equivalence_group(groups, "name", "大创结题", target.name)
    assert groups == {"name": [["双创训练结项", "大学生创新创业训练计划结题", "大创结题"]]}
    assert config.alias_groups == {}


def test_apply_to_original_preserves_other_zip_members_and_updates_progress(tmp_path: Path):
    reference, source = make_files(tmp_path)
    with ZipFile(source, "a") as archive:
        archive.writestr("xl/media/proof.png", b"image bytes stay untouched")
    config = FormatConfig()
    item = load_submissions([source], config)[0]
    target = Reference("学术创新", 9, "学术创新", "小类",
        "大学生创新创业训练计划结项", "获评项目", 8)
    edits = propose_original_edits(item, target, config)
    assert [(edit.field_name, edit.address, edit.new) for edit in edits] == [
        ("name", "C3", "1-大学生创新创业训练计划结项"),
        ("score", "E3", 8),
    ]
    progress = tmp_path / "progress.json"
    save_progress(progress, reference, config, [item])
    original_hash = file_fingerprint(source)
    apply_edits_to_original(source, "总表", edits)
    assert file_fingerprint(source) != original_hash
    with ZipFile(source) as archive:
        assert archive.testzip() is None
        assert archive.read("xl/media/proof.png") == b"image bytes stay untouched"
    workbook = load_workbook(source, read_only=True)
    try:
        assert workbook["总表"]["C3"].value == "1-大学生创新创业训练计划结项"
        assert workbook["总表"]["E3"].value == 8
    finally:
        workbook.close()
    save_progress(progress, reference, config, load_submissions([source], config),
        refresh_fingerprints=True)
    assert len(restore_progress(progress)[2]) == 3


def test_merged_category_edit_exposes_affected_range(tmp_path: Path):
    reference, source = make_files(tmp_path)
    workbook = load_workbook(source)
    workbook["总表"].merge_cells("A3:A4")
    workbook.save(source)
    item = load_submissions([source], FormatConfig())[0]
    target = Reference("学术创造", 2, "学术创造", "小类", item.name, item.role, item.score)
    edits = propose_original_edits(item, target, FormatConfig())
    major = next(edit for edit in edits if edit.field_name == "major")
    assert major.address == "A3"
    assert major.affected_range == "A3:A4"
    assert major.new == "学术创造（14分）"
