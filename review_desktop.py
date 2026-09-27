"""Offline desktop review for phone exports and current scoring sheets."""

from __future__ import annotations

from pathlib import Path
import sys
from zipfile import ZipFile

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QKeySequence, QPixmap, QShortcut
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
    QHBoxLayout, QLabel, QLineEdit, QListWidget, QMainWindow, QMessageBox, QPushButton,
    QScrollArea, QSpinBox, QSplitter, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget, QHeaderView,
)

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
from activity_matcher.review_core import (  # noqa: E402
    Candidate, SourceRecord, rank_candidates, sheet_headers, suggest_columns, workbook_sheets,
)
from activity_matcher.review_session import ReviewSession  # noqa: E402


EXCEL_FILTER = "Excel 工作簿 (*.xlsx *.xlsm)"
FIELD_NAMES = {"name": "活动名称 *", "dimension": "所属领域", "role": "职务/获奖等级", "score": "分数"}


class MappingDialog(QDialog):
    def __init__(self, path: Path, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.path = path
        self.setWindowTitle("指定新评分表的列")
        self.resize(650, 410)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("选择工作表和表头行，再确认字段对应关系。领域、职务和分数可以留空。"))
        form = QFormLayout()
        self.sheet = QComboBox()
        self.sheet.addItems(workbook_sheets(path))
        self.header_row = QSpinBox()
        self.header_row.setRange(1, 30)
        self.fields: dict[str, QComboBox] = {}
        form.addRow("工作表", self.sheet)
        form.addRow("表头所在行", self.header_row)
        for field, label in FIELD_NAMES.items():
            combo = QComboBox()
            self.fields[field] = combo
            form.addRow(label, combo)
        layout.addLayout(form)
        layout.addWidget(QLabel("表头预览"))
        self.preview = QLabel()
        self.preview.setWordWrap(True)
        self.preview.setMinimumHeight(55)
        layout.addWidget(self.preview)
        layout.addStretch()
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.sheet.currentIndexChanged.connect(self.update_headers)
        self.header_row.valueChanged.connect(self.update_headers)
        self.update_headers()

    def update_headers(self) -> None:
        try:
            headers = sheet_headers(self.path, self.sheet.currentText(), self.header_row.value())
        except (KeyError, StopIteration, ValueError) as error:
            self.preview.setText(str(error))
            return
        suggested = suggest_columns(headers)
        self.preview.setText("  |  ".join(f"{i + 1}. {title or '空列'}" for i, title in enumerate(headers[:20])))
        for field, combo in self.fields.items():
            combo.clear()
            combo.addItem("不使用", None)
            for index, title in enumerate(headers):
                combo.addItem(f"{index + 1}. {title or '空列'}", index)
            index = suggested[field]
            combo.setCurrentIndex(index + 1 if index is not None else 0)

    def mapping(self) -> dict[str, int | None]:
        return {field: combo.currentData() for field, combo in self.fields.items()}

    def validate(self) -> None:
        if self.mapping()["name"] is None:
            QMessageBox.warning(self, "缺少活动名称", "请选择新评分表的活动名称列。")
            return
        self.accept()


class ReviewWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("活动评分核对")
        self.resize(1450, 850)
        self.source_path: Path | None = None
        self.session: ReviewSession | None = None
        self.project_path: Path | None = None
        self.visible_sources: list[SourceRecord] = []
        self.displayed_candidates: list[Candidate] = []
        self.active_source: SourceRecord | None = None
        self._pixmap: QPixmap | None = None

        base = QWidget()
        self.setCentralWidget(base)
        layout = QVBoxLayout(base)
        bar = QHBoxLayout()
        self.source_button = QPushButton("打开手机导出 Excel")
        self.target_button = QPushButton("打开新评分表")
        self.open_button = QPushButton("打开核对进度")
        self.save_button = QPushButton("保存核对进度")
        self.export_button = QPushButton("导出核对结果")
        for button in (
            self.source_button, self.target_button, self.open_button,
            self.save_button, self.export_button,
        ):
            bar.addWidget(button)
        layout.addLayout(bar)
        self.summary = QLabel("先打开手机导出的 Excel，再打开新评分表。原文件不会被修改。")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)

        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter, 1)
        left = QWidget()
        left_layout = QVBoxLayout(left)
        filter_bar = QHBoxLayout()
        filter_bar.addWidget(QLabel("筛选"))
        self.filter = QComboBox()
        self.filter.addItems(["全部", "待确认", "相同", "分数不同", "已确认", "未找到/无对应"])
        filter_bar.addWidget(self.filter, 1)
        left_layout.addLayout(filter_bar)
        self.source_table = QTableWidget(0, 6)
        self.source_table.setHorizontalHeaderLabels([
            "状态", "领域/序号", "导出活动名", "职务/等级", "原分", "对应新表",
        ])
        self.source_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.source_table.setSelectionMode(QTableWidget.SingleSelection)
        self.source_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.source_table.setColumnWidth(0, 105)
        self.source_table.setColumnWidth(1, 105)
        self.source_table.setColumnWidth(3, 100)
        self.source_table.setColumnWidth(4, 65)
        self.source_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.source_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        left_layout.addWidget(self.source_table)
        splitter.addWidget(left)

        middle = QWidget()
        middle_layout = QVBoxLayout(middle)
        self.current_title = QLabel("选择左侧条目")
        self.current_title.setWordWrap(True)
        middle_layout.addWidget(self.current_title)
        middle_layout.addWidget(QLabel("新评分表候选条目（按相似程度排列）"))
        self.target_search = QLineEdit()
        self.target_search.setPlaceholderText("没有正确候选？搜索新表活动名、领域或等级")
        middle_layout.addWidget(self.target_search)
        self.candidate_table = QTableWidget(0, 5)
        self.candidate_table.setHorizontalHeaderLabels([
            "相似度", "新表活动名", "领域/等级", "分数", "差异说明",
        ])
        self.candidate_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.candidate_table.setSelectionMode(QTableWidget.SingleSelection)
        self.candidate_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.candidate_table.setColumnWidth(0, 70)
        self.candidate_table.setColumnWidth(2, 115)
        self.candidate_table.setColumnWidth(3, 65)
        self.candidate_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.candidate_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        middle_layout.addWidget(self.candidate_table, 1)
        action_bar = QHBoxLayout()
        self.confirm_button = QPushButton("确认选中对应项")
        self.no_match_button = QPushButton("确认无对应")
        self.clear_button = QPushButton("清除人工确认")
        for button in (self.confirm_button, self.no_match_button, self.clear_button):
            action_bar.addWidget(button)
        middle_layout.addLayout(action_bar)
        self.target_note = QLabel("选中候选后可确认；近似项不会自动确认。")
        self.target_note.setWordWrap(True)
        middle_layout.addWidget(self.target_note)
        splitter.addWidget(middle)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        self.photo_owner = QLabel("选择条目后显示对应的照片凭证。")
        self.photo_owner.setWordWrap(True)
        right_layout.addWidget(self.photo_owner)
        self.photo_list = QListWidget()
        self.photo_list.setMaximumHeight(145)
        right_layout.addWidget(self.photo_list)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.photo_preview = QLabel("暂无照片")
        self.photo_preview.setAlignment(Qt.AlignCenter)
        self.photo_preview.setMinimumSize(300, 320)
        scroll.setWidget(self.photo_preview)
        right_layout.addWidget(scroll, 1)
        splitter.addWidget(right)
        splitter.setSizes([620, 520, 310])

        self.source_button.clicked.connect(self.choose_source)
        self.target_button.clicked.connect(self.choose_target)
        self.open_button.clicked.connect(self.open_project)
        self.save_button.clicked.connect(self.save_project)
        self.export_button.clicked.connect(self.export_result)
        self.filter.currentIndexChanged.connect(self.refresh_sources)
        self.source_table.itemSelectionChanged.connect(self.source_selected)
        self.candidate_table.itemSelectionChanged.connect(self.candidate_selected)
        self.target_search.textChanged.connect(self.update_candidates)
        self.photo_list.currentRowChanged.connect(self.show_photo)
        self.confirm_button.clicked.connect(self.confirm_selected)
        self.no_match_button.clicked.connect(self.mark_no_match)
        self.clear_button.clicked.connect(self.clear_decision)
        QShortcut(QKeySequence("Alt+C"), self, activated=self.confirm_selected)
        QShortcut(QKeySequence("Alt+N"), self, activated=self.mark_no_match)
        self.statusBar().showMessage("离线运行；原始 Excel 保持不变")
        self.update_enabled()

    def error(self, title: str, error: Exception) -> None:
        QMessageBox.critical(self, title, str(error))

    def update_enabled(self) -> None:
        has_session = self.session is not None
        self.target_button.setEnabled(self.source_path is not None)
        for button in (self.save_button, self.export_button, self.confirm_button, self.no_match_button, self.clear_button):
            button.setEnabled(has_session)

    def choose_source(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择手机导出的 Excel", "", EXCEL_FILTER)
        if selected:
            self.source_path = Path(selected)
            self.session = None
            self.project_path = None
            self.source_table.setRowCount(0)
            self.candidate_table.setRowCount(0)
            self.summary.setText(f"已选择手机导出文件：{self.source_path.name}。现在打开新评分表。")
            self.update_enabled()

    def choose_target(self) -> None:
        if self.source_path is None:
            return
        selected, _ = QFileDialog.getOpenFileName(self, "选择新的评分表", "", EXCEL_FILTER)
        if not selected:
            return
        try:
            dialog = MappingDialog(Path(selected), self)
            if dialog.exec() != QDialog.Accepted:
                return
            self.session = ReviewSession.create(
                self.source_path, Path(selected), dialog.sheet.currentText(),
                dialog.header_row.value(), dialog.mapping(),
            )
            self.project_path = None
            self.update_enabled()
            self.refresh_sources()
        except Exception as error:
            self.error("导入失败", error)

    def open_project(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "打开核对进度", "", "核对进度 (*.json)")
        if not selected:
            return
        try:
            self.session = ReviewSession.open(Path(selected))
            self.source_path = self.session.source_path
            self.project_path = Path(selected)
            self.update_enabled()
            self.refresh_sources()
        except Exception as error:
            self.error("无法打开核对进度", error)

    def save_project(self) -> None:
        if self.session is None:
            return
        path = self.project_path
        if path is None:
            selected, _ = QFileDialog.getSaveFileName(self, "保存核对进度", "核对进度.json", "核对进度 (*.json)")
            if not selected:
                return
            path = Path(selected)
        try:
            self.session.save(path)
            self.project_path = path
            self.statusBar().showMessage(f"核对进度已保存：{path}", 5000)
        except Exception as error:
            self.error("保存失败", error)

    def export_result(self) -> None:
        if self.session is None:
            return
        selected, _ = QFileDialog.getSaveFileName(self, "导出核对结果", "核对结果.xlsx", EXCEL_FILTER)
        if not selected:
            return
        try:
            self.session.export_result(Path(selected))
            self.statusBar().showMessage(f"核对结果已导出：{selected}", 6000)
        except Exception as error:
            self.error("导出失败", error)

    def _passes_filter(self, status: str) -> bool:
        selected = self.filter.currentText()
        if selected == "全部":
            return True
        if selected == "待确认":
            return status in {"近似待确认", "重复候选", "未找到"}
        if selected == "相同":
            return status == "相同"
        if selected == "分数不同":
            return "分数不同" in status
        if selected == "已确认":
            return status.startswith("已确认")
        return status in {"未找到", "确认无对应"}

    def refresh_sources(self) -> None:
        if self.session is None:
            return
        previous_id = self.active_source.record_id if self.active_source else None
        self.source_table.blockSignals(True)
        self.source_table.setRowCount(0)
        self.visible_sources = []
        counts: dict[str, int] = {}
        for source in self.session.sources:
            status = self.session.status(source)
            counts[status] = counts.get(status, 0) + 1
            if not self._passes_filter(status):
                continue
            target = self.session.linked_target(source)
            top_candidates = self.session.candidates(source, 1)
            suggested = target.name if target else (top_candidates[0].target.name if top_candidates else "")
            values = [status, f"{source.dimension} {source.number}", source.name,
                      source.role, "" if source.score is None else str(source.score), suggested]
            row = self.source_table.rowCount()
            self.source_table.insertRow(row)
            self.visible_sources.append(source)
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if status in {"近似待确认", "重复候选", "未找到"}:
                    item.setBackground(QColor("#FFF2CC"))
                elif "分数不同" in status:
                    item.setBackground(QColor("#FCE4D6"))
                elif status in {"相同", "已确认"}:
                    item.setBackground(QColor("#E2F0D9"))
                self.source_table.setItem(row, column, item)
        self.source_table.blockSignals(False)
        self.summary.setText(
            f"导出 {len(self.session.sources)} 条；新表 {len(self.session.targets)} 条；"
            + "；".join(f"{status} {count}" for status, count in counts.items())
        )
        if self.visible_sources:
            row = next((i for i, item in enumerate(self.visible_sources) if item.record_id == previous_id), 0)
            self.source_table.selectRow(row)
        else:
            self.active_source = None
            self.current_title.setText("当前筛选没有条目")
            self.candidate_table.setRowCount(0)
            self.photo_list.clear()
            self.photo_preview.setText("暂无照片")

    def source_selected(self) -> None:
        row = self.source_table.currentRow()
        if self.session is None or not 0 <= row < len(self.visible_sources):
            return
        source = self.visible_sources[row]
        self.active_source = source
        if self.target_search.text():
            self.target_search.blockSignals(True)
            self.target_search.clear()
            self.target_search.blockSignals(False)
        self.current_title.setText(
            f"{source.dimension} · {source.number} · {source.name}\n"
            f"职务/等级：{source.role or '未填写'}    原分数：{source.score if source.score is not None else '未填写'}"
        )
        self.update_candidates()
        linked = self.session.linked_target(source)
        self.photo_owner.setText(
            f"凭证属于：{source.dimension} · {source.number} · {source.name}\n"
            f"对应新表：{linked.name if linked else '尚未确认'}"
        )
        self.photo_list.clear()
        for index, photo in enumerate(source.photos, start=1):
            self.photo_list.addItem(f"凭证 {index} · {photo.sheet}!{photo.cell}")
        if source.photos:
            self.photo_list.setCurrentRow(0)
        else:
            self._pixmap = None
            self.photo_preview.setPixmap(QPixmap())
            self.photo_preview.setText("这条活动没有照片凭证")

    def update_candidates(self) -> None:
        if self.session is None or self.active_source is None:
            return
        source = self.active_source
        query = self.target_search.text().strip().lower()
        if query:
            matches = [
                target for target in self.session.targets
                if query in " ".join((target.name, target.dimension, target.role)).lower()
            ]
            ranked = rank_candidates(source, matches, limit=50)
            included = {candidate.target.record_id for candidate in ranked}
            ranked.extend(Candidate(target, 0.0, False, "手动搜索") for target in matches if target.record_id not in included)
            candidates = ranked[:50]
        else:
            candidates = self.session.candidates(source)
        self.displayed_candidates = candidates
        self.candidate_table.blockSignals(True)
        self.candidate_table.setRowCount(0)
        linked = self.session.linked_target(source)
        for candidate in candidates:
            target = candidate.target
            row_index = self.candidate_table.rowCount()
            self.candidate_table.insertRow(row_index)
            values = [f"{candidate.similarity:.0%}" if candidate.similarity else "搜索", target.name,
                      " / ".join(part for part in (target.dimension, target.role) if part),
                      "" if target.score is None else str(target.score),
                      f"{candidate.reason} · {target.sheet}!{target.row}"]
            for column, value in enumerate(values):
                self.candidate_table.setItem(row_index, column, QTableWidgetItem(value))
            if linked and linked.record_id == target.record_id:
                self.candidate_table.selectRow(row_index)
        self.candidate_table.blockSignals(False)
        if linked and not candidates:
            self.target_note.setText(f"已确认：{linked.name}（{linked.sheet}!{linked.row}）")
        elif source.record_id in self.session.decisions and linked is None:
            self.target_note.setText("人工确认：新评分表没有对应条目。")
        else:
            self.candidate_selected()

    def candidate_selected(self) -> None:
        if self.session is None or self.active_source is None:
            return
        row = self.candidate_table.currentRow()
        candidates = self.displayed_candidates
        if 0 <= row < len(candidates):
            target = candidates[row].target
            usage = self.session.target_usage()[target.record_id]
            warning = f"；已有 {usage} 条导出记录指向它" if usage > 1 else ""
            self.target_note.setText(f"候选：{target.name} · {target.sheet}!{target.row}{warning}")
        elif self.active_source.record_id not in self.session.decisions:
            self.target_note.setText("选中候选后可确认；近似项不会自动确认。")

    def show_photo(self, index: int) -> None:
        if self.session is None or self.active_source is None or index < 0 or index >= len(self.active_source.photos):
            return
        try:
            photo = self.active_source.photos[index]
            with ZipFile(self.session.source_path) as archive:
                image_data = archive.read(photo.archive_member)
            pixmap = QPixmap()
            if not pixmap.loadFromData(image_data):
                raise ValueError(f"无法预览图片：{photo.archive_member}")
            self._pixmap = pixmap
            self.photo_preview.setText("")
            self.photo_preview.setPixmap(pixmap.scaled(650, 650, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        except Exception as error:
            self.error("凭证预览失败", error)

    def _commit(self, target_id: str | None) -> None:
        if self.session is None or self.active_source is None:
            return
        self.session.decisions[self.active_source.record_id] = target_id
        self._save_if_possible()
        self.refresh_sources()

    def _save_if_possible(self) -> None:
        if self.session is not None and self.project_path is not None:
            try:
                self.session.save(self.project_path)
            except Exception as error:
                self.error("自动保存失败", error)
        else:
            self.statusBar().showMessage("确认结果尚未保存，请点“保存核对进度”", 6000)

    def confirm_selected(self) -> None:
        if self.session is None or self.active_source is None:
            return
        row = self.candidate_table.currentRow()
        candidates = self.displayed_candidates
        if not 0 <= row < len(candidates):
            QMessageBox.information(self, "选择候选", "先在中间选择一个新评分表条目。")
            return
        self._commit(candidates[row].target.record_id)

    def mark_no_match(self) -> None:
        self._commit(None)

    def clear_decision(self) -> None:
        if self.session is None or self.active_source is None:
            return
        self.session.decisions.pop(self.active_source.record_id, None)
        self._save_if_possible()
        self.refresh_sources()

    def closeEvent(self, event) -> None:  # type: ignore[override]
        if self.session and self.session.decisions and self.project_path is None:
            choice = QMessageBox.question(
                self, "核对进度尚未保存", "人工确认结果尚未保存。现在关闭会丢失这些结果，确定关闭吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if choice != QMessageBox.Yes:
                event.ignore()
                return
        event.accept()


def main() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft YaHei UI", 10))
    window = ReviewWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
