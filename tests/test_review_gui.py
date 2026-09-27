import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from activity_matcher.review_core import sheet_headers, suggest_columns
from activity_matcher.review_session import ReviewSession
from review_desktop import ReviewWindow
from test_review_core import make_workbooks


def test_review_window_confirms_a_candidate(tmp_path: Path) -> None:
    app = QApplication.instance() or QApplication([])
    source_path, target_path = make_workbooks(tmp_path)
    mapping = suggest_columns(sheet_headers(target_path, "新标准", 1))
    window = ReviewWindow()
    window.session = ReviewSession.create(source_path, target_path, "新标准", 1, mapping)
    window.source_path = source_path
    window.project_path = tmp_path / "progress.json"
    window.update_enabled()
    window.refresh_sources()
    assert window.source_table.rowCount() == 3
    window.source_table.selectRow(2)
    assert window.candidate_table.rowCount() > 0
    assert "校级演讲活动" in window.photo_owner.text()
    window.target_search.setText("物理竞赛")
    assert window.candidate_table.rowCount() == 1
    assert window.displayed_candidates[0].target.name == "物理竞赛"
    window.target_search.clear()
    window.candidate_table.selectRow(0)
    window.confirm_selected()
    assert window.session.decisions["43"] == "新标准:4"
    assert window.project_path.is_file()
    assert window.source_table.item(2, 0).text() == "已确认"
    window.close()
    app.processEvents()
