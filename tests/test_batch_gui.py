import os
from dataclasses import replace
import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QProcess, Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QApplication, QDialog, QGraphicsView, QMessageBox, QTableWidgetItem
from openpyxl import load_workbook

from activity_matcher.batch_core import (
    FormatConfig, ImageRef, Match, ProposedEdit, Reference, load_alias_settings,
    load_submissions, save_alias_settings,
)
from batch_review_desktop import (
    AliasDialog, ApplyOriginalDialog, BatchReviewWindow, IgnoreDifferenceDialog, PhotoView,
    migrate_alias_settings, show_loaded_alias_warning,
)
from test_batch_core import make_files


def test_startup_warns_only_when_alias_rules_loaded(monkeypatch):
    app = QApplication.instance() or QApplication([])
    window = BatchReviewWindow()
    shown = []

    def capture(message):
        shown.append((message.windowTitle(), message.text(), message.informativeText()))
        return 0

    monkeypatch.setattr(QMessageBox, "exec", capture)
    window.config = replace(window.config, alias_groups={}, category_aliases={})
    show_loaded_alias_warning(window)
    assert shown == []

    window.config = replace(window.config,
        alias_groups={"major": [["身心健康", "身心健康活动成绩", "身心健康成绩"]]})
    show_loaded_alias_warning(window)
    assert len(shown) == 1
    assert "1 组" in shown[0][1]
    assert "身心健康 = 身心健康活动成绩 = 身心健康成绩" in shown[0][2]
    window.close()
    app.processEvents()


def test_alias_settings_migration_copies_once_without_overwriting(tmp_path):
    old = tmp_path / "old" / "batch_review_settings.json"
    new = tmp_path / "user" / "batch_review_settings.json"
    original = {"major": [["身心健康", "身心健康活动成绩"]]}
    save_alias_settings(old, original)
    assert migrate_alias_settings(new, [old]) == old
    assert new.read_bytes() == old.read_bytes()

    save_alias_settings(old, {"name": [["旧名称", "新名称"]]})
    assert migrate_alias_settings(new, [old]) is None
    assert load_alias_settings(new) == original

    invalid = tmp_path / "invalid.json"
    invalid.write_text("not json", encoding="utf-8")
    target = tmp_path / "another" / "batch_review_settings.json"
    with pytest.raises(ValueError):
        migrate_alias_settings(target, [invalid])
    assert not target.exists()


def test_alias_editor_accepts_multiple_names_and_fields():
    app = QApplication.instance() or QApplication([])
    dialog = AliasDialog({"major": [["身心健康", "身心健康活动成绩"]]})
    dialog.table.item(0, 2).setText("身心健康活动成绩；身心健康成绩；体育健康")
    dialog.add_row("name", "大学生创新创业训练计划结题", "双创结题；大创结题")
    dialog.validate_and_accept()
    assert dialog.result_alias_groups == {
        "major": [["身心健康", "身心健康活动成绩", "身心健康成绩", "体育健康"]],
        "name": [["大学生创新创业训练计划结题", "双创结题", "大创结题"]],
    }
    dialog.close()
    app.processEvents()


def test_quick_action_dialogs_preview_rules_and_original_cells(tmp_path):
    app = QApplication.instance() or QApplication([])
    _, source = make_files(tmp_path)
    item = load_submissions([source], FormatConfig())[0]
    target = Reference("学术创新", 9, "学术创新", "小类", "双创训练结项", "获评项目", 8)
    ignore = IgnoreDifferenceDialog(item, target, FormatConfig())
    assert len(ignore.checks) == 1
    assert "双创训练结项" in ignore.preview.toPlainText()
    ignore.validate_and_accept()
    assert ignore.result_groups == {"name": [["双创训练结项", item.name]]}
    ignore.close()

    edits = [
        ProposedEdit("major", "A3", "学术创新（14分）", "学术创造（14分）", "A3:A12"),
        ProposedEdit("name", "C3", "旧活动", "新活动"),
    ]
    apply = ApplyOriginalDialog(item, target, edits)
    assert apply.table.item(0, 0).checkState() == Qt.Unchecked
    assert apply.table.item(1, 0).checkState() == Qt.Checked
    apply.validate_and_accept()
    assert apply.selected_edits == edits[1:]
    apply.close()
    app.processEvents()


def test_import_warning_offers_backup_or_continue(monkeypatch, tmp_path):
    app = QApplication.instance() or QApplication([])
    window = BatchReviewWindow()
    _, source = make_files(tmp_path)
    shown = []

    def continue_import(message):
        shown.append(message.text())
        next(button for button in message.buttons() if button.text() == "暂不备份，继续导入").click()
        return 0

    monkeypatch.setattr(QMessageBox, "exec", continue_import)
    assert window._prompt_import_backup([source])
    assert "直接修改" in shown[0]
    window.close()
    app.processEvents()


def test_quick_actions_update_rules_and_original_file(monkeypatch, tmp_path):
    app = QApplication.instance() or QApplication([])
    reference, source = make_files(tmp_path)
    window = BatchReviewWindow()
    window.config = FormatConfig()
    window.reference_path = reference
    window.source_paths = [source]
    window.scan()
    window.source_table.selectRow(0)
    item = window.current
    target = Reference("学术创新", 9, "学术创新", "小类", "双创训练结项", "获评项目", 10)
    settings = tmp_path / "rules.json"
    monkeypatch.setattr("batch_review_desktop.SETTINGS_PATH", settings)

    def select_candidate():
        window.candidates = [Match(target, .8, "活动名", False)]
        window.candidate_table.setRowCount(1)
        window.candidate_table.setItem(0, 0, QTableWidgetItem("80%"))
        window.candidate_table.selectRow(0)

    def confirm_ignore(dialog):
        dialog.validate_and_accept()
        return QDialog.Accepted

    select_candidate()
    monkeypatch.setattr(IgnoreDifferenceDialog, "exec", confirm_ignore)
    window.ignore_difference()
    assert settings.is_file()
    assert window.config.alias_groups["name"][0] == ["双创训练结项", item.name]

    select_candidate()
    monkeypatch.setattr(ApplyOriginalDialog, "exec", lambda dialog: setattr(dialog, "selected_edits", dialog.edits) or QDialog.Accepted)
    window.apply_to_original()
    workbook = load_workbook(source, read_only=True)
    try:
        assert workbook["总表"]["C3"].value == "1-双创训练结项"
    finally:
        workbook.close()
    window.dirty = False
    window.close()
    app.processEvents()


def test_proof_decision_advances_to_next_pending_with_filter(tmp_path):
    app = QApplication.instance() or QApplication([])
    reference, source = make_files(tmp_path)
    window = BatchReviewWindow()
    window.config = FormatConfig()
    window.reference_path = reference
    window.source_paths = [source]
    window.scan()
    window.filter.setCurrentText("证明待审")
    window.source_table.selectRow(0)
    assert "font-size: 18pt" in window.proof_title.styleSheet()
    assert window.items[0].name in window.proof_title.text()

    window.set_proof("insufficient")
    assert window.items[0].proof == "insufficient"
    assert window.current.id == window.items[1].id
    window.set_proof("insufficient")
    assert window.current.id == window.items[2].id
    window.set_proof("insufficient")
    assert not window.visible
    assert "证明已全部审核" in window.statusBar().currentMessage()
    window.dirty = False
    window.close()
    app.processEvents()


def test_photo_view_preserves_full_image_and_can_zoom_and_pan():
    app = QApplication.instance() or QApplication([])
    view = PhotoView()
    view.resize(600, 500)
    view.show()
    app.processEvents()
    view.set_photo(QPixmap(2400, 1600))
    assert view.sceneRect().width() == 2400
    assert view.sceneRect().height() == 1600
    assert view.dragMode() == QGraphicsView.ScrollHandDrag
    fit_scale = view.transform().m11()
    assert 0 < fit_scale < 1
    view.zoom(1.2)
    assert view.transform().m11() > fit_scale
    view.actual_size()
    assert view.transform().m11() == 1
    view.fit_photo()
    assert view.transform().m11() < 1
    view.close()
    app.processEvents()


def test_excel_navigation_targets_summary_and_proof_even_without_image(monkeypatch, tmp_path):
    app = QApplication.instance() or QApplication([])
    reference, source = make_files(tmp_path)
    window = BatchReviewWindow()
    window.config = FormatConfig()
    window.reference_path = reference
    window.source_paths = [source]
    window.scan()
    targets = []
    monkeypatch.setattr(window, "_navigate_excel", lambda *args: targets.append(args))
    item = window.items[0]

    window.open_summary_excel(item)
    window.open_evidence_excel(item)
    assert targets == [(source, "总表", "C3"), (source, "学术创新活动证明", "B2")]

    item.photos = [ImageRef("学术创新活动证明", "D2", "xl/media/image1.png")]
    window.open_evidence_excel(item)
    assert targets[-1] == (source, "学术创新活动证明", "D2")

    assert window.source_table.contextMenuPolicy() == Qt.CustomContextMenu
    menu = window._source_menu(item)
    assert [action.text() for action in menu.actions()] == ["在 Excel 中定位总表条目", "在 Excel 中定位证明"]
    for action in menu.actions():
        action.trigger()
    assert targets[-2:] == [(source, "总表", "C3"), (source, "学术创新活动证明", "D2")]
    window.source_table.itemDoubleClicked.emit(window.source_table.item(0, 0))
    assert targets[-1] == (source, "总表", "C3")
    window.dirty = False
    window.close()
    app.processEvents()


def test_excel_navigation_process_receives_exact_location(monkeypatch, tmp_path):
    app = QApplication.instance() or QApplication([])
    _, source = make_files(tmp_path)
    window = BatchReviewWindow()
    monkeypatch.setattr(QProcess, "start", lambda _process: None)
    window._navigate_excel(source, "学术创新活动证明", "B2")
    assert len(window.excel_processes) == 1
    process = next(iter(window.excel_processes))
    assert process.program() == "powershell.exe"
    assert process.arguments()[-6:] == ["-WorkbookPath", str(source.resolve()),
        "-SheetName", "学术创新活动证明", "-CellAddress", "B2"]
    window._excel_navigation_finished(process, 0, source, "学术创新活动证明", "B2")
    assert not window.excel_processes
    assert "学术创新活动证明!B2" in window.statusBar().currentMessage()
    window.close()
    app.processEvents()
