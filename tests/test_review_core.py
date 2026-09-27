from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
import pytest

from activity_matcher.review_core import (
    ScoreRecord, _read_photo_map, automatic_status, load_phone_export, rank_candidates,
    load_score_table, sheet_headers, suggest_columns, workbook_sheets,
)
from activity_matcher.review_session import ReviewSession


def make_workbooks(tmp_path: Path) -> tuple[Path, Path]:
    source = Workbook()
    summary = source.active
    summary.title = "总表"
    summary.append(["领域", "序号", "活动名", "获奖等级", "分数", "条目ID"])
    summary.append(["学术创新", 1, "物理竞赛", "一等奖", 3, "41"])
    summary.append([None, 2, "数学比赛", "参与者", 2, "42"])
    summary.append([None, 3, "校级演讲活动", "参与者", 1, "43"])
    source.create_sheet("学术创新").append(["序号", "凭证1"])
    source_path = tmp_path / "手机导出.xlsx"
    source.save(source_path)

    target = Workbook()
    sheet = target.active
    sheet.title = "新标准"
    sheet.append(["活动名称", "所属维度", "获奖等级", "认定分数"])
    sheet.append(["物理竞赛", "学术创新", "一等奖", 3])
    sheet.append(["数学比赛", "学术创新", "参与者", 4])
    sheet.append(["校级演讲比赛", "学术创新", "参与者", 1])
    sheet.merge_cells("B2:B4")
    target_path = tmp_path / "新评分表.xlsx"
    target.save(target_path)
    return source_path, target_path


def test_import_match_and_export(tmp_path: Path) -> None:
    source_path, target_path = make_workbooks(tmp_path)
    headers = sheet_headers(target_path, "新标准", 1)
    mapping = suggest_columns(headers)
    assert mapping == {"name": 0, "dimension": 1, "role": 2, "score": 3}
    session = ReviewSession.create(source_path, target_path, "新标准", 1, mapping)
    assert [item.record_id for item in session.sources] == ["41", "42", "43"]
    assert [session.status(item) for item in session.sources] == ["相同", "分数不同", "近似待确认"]
    assert session.sources[0].dimension == session.sources[1].dimension == "学术创新"

    candidates = rank_candidates(session.sources[2], session.targets)
    assert candidates[0].target.name == "校级演讲比赛"
    assert automatic_status(candidates, 1)[0] == "近似待确认"
    session.decisions["43"] = candidates[0].target.record_id
    assert session.status(session.sources[2]) == "已确认"
    session.decisions["42"] = None
    assert session.status(session.sources[1]) == "确认无对应"

    progress = tmp_path / "progress.json"
    session.save(progress)
    restored = ReviewSession.open(progress)
    assert restored.decisions == session.decisions
    result_path = tmp_path / "结果.xlsx"
    restored.export_result(result_path)
    result = load_workbook(result_path, read_only=True, data_only=True)
    assert result.sheetnames == ["核对结果", "新表未对应"]
    assert result["核对结果"]["A4"].value == "已确认"
    assert result["核对结果"]["L4"].value == 0
    assert result["新表未对应"].max_row == 2
    result.close()


def test_changed_workbook_invalidates_progress(tmp_path: Path) -> None:
    source_path, target_path = make_workbooks(tmp_path)
    session = ReviewSession.create(source_path, target_path, "新标准", 1, suggest_columns(sheet_headers(target_path, "新标准", 1)))
    progress = tmp_path / "progress.json"
    session.save(progress)
    with target_path.open("ab") as output:
        output.write(b"changed")
    with pytest.raises(ValueError, match="内容已变化"):
        ReviewSession.open(progress)


def test_duplicate_exact_matches_require_review(tmp_path: Path) -> None:
    source_path, target_path = make_workbooks(tmp_path)
    source = load_phone_export(source_path)[0]
    targets = [
        ScoreRecord(f"S:{row}", "学术创新", "物理竞赛", "一等奖", score, "S", row)
        for row, score in ((2, 3.0), (3, 4.0))
    ]
    assert automatic_status(rank_candidates(source, targets), source.score) == ("重复候选", None)


def test_new_scoring_table_with_invalid_fill_still_imports(tmp_path: Path) -> None:
    _, target_path = make_workbooks(tmp_path)
    broken_path = tmp_path / "样式异常的新评分表.xlsx"
    with ZipFile(target_path) as source, ZipFile(broken_path, "w") as destination:
        for member in source.infolist():
            data = source.read(member.filename)
            if member.filename == "xl/styles.xml":
                root = ET.fromstring(data)
                fills = root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}fills")
                assert fills is not None
                fills[0].clear()
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            destination.writestr(member, data)
    original_bytes = broken_path.read_bytes()
    with pytest.raises(TypeError, match="expected.*Fill"):
        load_workbook(broken_path, read_only=True, data_only=True)

    assert workbook_sheets(broken_path) == ["新标准"]
    mapping = suggest_columns(sheet_headers(broken_path, "新标准", 1))
    rows = load_score_table(broken_path, "新标准", 1, mapping)
    assert len(rows) == 3
    assert rows[1].name == "数学比赛"
    assert broken_path.read_bytes() == original_bytes


def test_existing_cell_images_are_resolved() -> None:
    example = Path(__file__).resolve().parents[3] / "outputs" / "activity-score-android" / "build" / "test-export" / "sample.xlsx"
    if not example.is_file():
        pytest.skip("Android export smoke sample unavailable")
    workbook = load_workbook(example, read_only=True)
    photo_map = _read_photo_map(example, workbook.sheetnames)
    workbook.close()
    members = {photo.archive_member for photos in photo_map.values() for photo in photos}
    assert members == {"xl/media/image1.jpg", "xl/media/image2.png", "xl/media/image3.jpg"}
    records = load_phone_export(example)
    assert len(records) == 4
    assert records[0].record_id == "41"
    assert len(records[0].photos) == 2
    assert len(records[-1].photos) == 1
