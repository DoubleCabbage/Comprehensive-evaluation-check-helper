"""Offline batch review of student workbooks, independent of the phone app."""

from __future__ import annotations

from dataclasses import replace
import os
from pathlib import Path
import sys
from zipfile import ZipFile

from PySide6.QtCore import QProcess, Qt
from PySide6.QtGui import QColor, QFont, QKeySequence, QPainter, QPixmap, QShortcut
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout, QLineEdit,
    QGraphicsScene, QGraphicsView, QGroupBox, QHBoxLayout, QHeaderView, QLabel, QMainWindow, QMenu, QMessageBox,
    QPlainTextEdit, QPushButton, QSpinBox, QSplitter, QTableWidget, QTableWidgetItem,
    QTabWidget, QVBoxLayout, QWidget,
)
from openpyxl.utils.cell import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
from activity_matcher.batch_core import (  # noqa: E402
    ALIAS_FIELDS, FormatConfig, ProposedEdit, Reference, Submission, add_equivalence_group,
    apply_edits_to_original, backup_workbooks, export_review_copies, file_fingerprint,
    load_alias_settings, load_references, load_submissions, propose_original_edits,
    rank, restore_progress, save_alias_settings, save_progress, status,
    text_differences, validate_alias_groups,
)
from activity_matcher.review_core import sheet_headers, workbook_sheets  # noqa: E402


EXCEL_FILTER = "Excel 工作簿 (*.xlsx *.xlsm)"
SETTINGS_PATH = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming") / \
    "CollegeReview" / "batch_review_settings.json"
EXCEL_NAVIGATION_SCRIPT = Path(__file__).resolve().parent / "excel_navigation.ps1"
REF_FIELDS = [("reference_major", "大类"), ("reference_minor", "小类"),
    ("reference_name", "活动名称"), ("reference_role", "等级"), ("reference_score", "参考分数")]
SOURCE_FIELDS = [("summary_major", "大类"), ("summary_number", "序号"),
    ("summary_name", "活动名称"), ("summary_role", "等级"),
    ("summary_score", "申报分数"), ("summary_remark", "备注")]


def legacy_settings_paths() -> list[Path]:
    project_config = Path(__file__).resolve().parent / "configs" / "batch_review_settings.json"
    if getattr(sys, "frozen", False):
        beside_exe = Path(sys.executable).resolve().parent / "configs" / "batch_review_settings.json"
        return list(dict.fromkeys((beside_exe, project_config)))
    return [project_config]


def migrate_alias_settings(target: Path, legacy_paths: list[Path]) -> Path | None:
    """Copy valid old rules once; never replace the user's AppData rules."""
    if target.exists():
        return None
    for legacy in legacy_paths:
        if not legacy.is_file():
            continue
        load_alias_settings(legacy)
        contents = legacy.read_bytes()
        target.parent.mkdir(parents=True, exist_ok=True)
        try:
            with target.open("xb") as saved:
                saved.write(contents)
        except FileExistsError:
            return None
        except OSError:
            target.unlink(missing_ok=True)
            raise
        return legacy
    return None


class FormatDialog(QDialog):
    def __init__(self, config: FormatConfig, reference: Path | None, sample: Path | None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置本批表格格式")
        self.resize(730, 700)
        self.config = config
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("列号从 1 开始。参考大表的所有 Sheet 使用同一套列；个人表格本批使用同一套格式。"))
        self.controls = {}
        ref_box = QGroupBox("参考大表")
        ref_form = QFormLayout(ref_box)
        self._spin(ref_form, "reference_header", "表头行", config.reference_header, True)
        for field, label in REF_FIELDS:
            self._spin(ref_form, field, label, getattr(config, field) + 1)
        layout.addWidget(ref_box)
        source_box = QGroupBox("个人提交总表")
        source_form = QFormLayout(source_box)
        self.sheet = QComboBox()
        if sample:
            self.sheet.addItems(workbook_sheets(sample))
        else:
            self.sheet.addItem(config.summary_sheet)
        self.sheet.setCurrentText(config.summary_sheet)
        source_form.addRow("总表 Sheet", self.sheet)
        self._spin(source_form, "summary_header", "表头行", config.summary_header, True)
        for field, label in SOURCE_FIELDS:
            self._spin(source_form, field, label, getattr(config, field) + 1)
        self._spin(source_form, "summary_minor", "小类（可选）",
            0 if config.summary_minor is None else config.summary_minor + 1, "optional")
        layout.addWidget(source_box)
        self.mapping_table = QTableWidget(0, 2)
        self.mapping_table.setHorizontalHeaderLabels(["总表大类", "分类证明 Sheet"])
        self.mapping_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.mapping_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        layout.addWidget(QLabel("分类证明对应关系（自动识别失败时可在这里指定）"))
        reload_mapping = QPushButton("按当前列设置刷新分类")
        reload_mapping.clicked.connect(lambda: self._refresh_evidence(sample))
        layout.addWidget(reload_mapping)
        layout.addWidget(self.mapping_table, 1)
        self.mapping_status = QLabel()
        layout.addWidget(self.mapping_status)
        self._refresh_evidence(sample)
        self.preview = QLabel(f"参考文件：{reference.name if reference else '未选'}；个人样表：{sample.name if sample else '未选'}")
        layout.addWidget(self.preview)
        self.header_preview = QLabel()
        self.header_preview.setWordWrap(True)
        layout.addWidget(self.header_preview)
        self.sheet.currentIndexChanged.connect(lambda _: self._update_preview(reference, sample))
        self.controls["reference_header"][0].valueChanged.connect(lambda _: self._update_preview(reference, sample))
        self.controls["summary_header"][0].valueChanged.connect(lambda _: self._update_preview(reference, sample))
        self._update_preview(reference, sample)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _update_preview(self, reference: Path | None, sample: Path | None):
        lines = []
        try:
            if reference:
                first = workbook_sheets(reference)[0]
                headers = sheet_headers(reference, first, self.controls["reference_header"][0].value())
                lines.append("参考表：" + "  ".join(f"{i + 1}.{value}" for i, value in enumerate(headers[:12])))
            if sample:
                headers = sheet_headers(sample, self.sheet.currentText(), self.controls["summary_header"][0].value())
                lines.append("个人总表：" + "  ".join(f"{i + 1}.{value}" for i, value in enumerate(headers[:12])))
        except Exception as error:
            lines.append(f"表头预览失败：{error}")
        self.header_preview.setText("\n".join(lines))

    def _refresh_evidence(self, sample: Path | None):
        if not sample:
            self.mapping_status.setText("加入个人样表后可设置分类证明 Sheet。")
            return
        try:
            configured = self.result_config()
            categories = list(dict.fromkeys(item.major for item in load_submissions([sample], configured)))
            sheets = workbook_sheets(sample)
            self.mapping_table.setRowCount(len(categories))
            for row, major in enumerate(categories):
                self.mapping_table.setItem(row, 0, QTableWidgetItem(major))
                combo = QComboBox()
                combo.addItem("自动识别", "")
                for name in sheets:
                    if name != configured.summary_sheet:
                        combo.addItem(name, name)
                index = combo.findData(configured.evidence_sheets.get(major, ""))
                combo.setCurrentIndex(max(0, index))
                self.mapping_table.setCellWidget(row, 1, combo)
            self.mapping_status.setText(f"识别出 {len(categories)} 个大类。")
        except Exception as error:
            self.mapping_status.setText(f"分类预览暂不可用：{error}")

    def _spin(self, form, field: str, label: str, value: int, direct: bool = False):
        control = QSpinBox()
        control.setRange(0 if direct == "optional" else 1, 200)
        if direct == "optional":
            control.setSpecialValueText("不使用")
        control.setValue(value)
        self.controls[field] = (control, direct)
        form.addRow(label, control)

    def result_config(self) -> FormatConfig:
        values = {field: (None if widget.value() == 0 else widget.value() - 1)
            if direct == "optional" else (widget.value() if direct else widget.value() - 1)
            for field, (widget, direct) in self.controls.items()}
        mapping = {}
        for row in range(self.mapping_table.rowCount()):
            major = self.mapping_table.item(row, 0).text()
            sheet = self.mapping_table.cellWidget(row, 1).currentData()
            if sheet:
                mapping[major] = sheet
        return replace(self.config, summary_sheet=self.sheet.currentText(),
            evidence_sheets=mapping, **values)


class AliasDialog(QDialog):
    def __init__(self, groups: dict[str, list[list[str]]], parent=None):
        super().__init__(parent)
        self.setWindowTitle("长期设置 · 等价写法")
        self.resize(850, 460)
        self.result_alias_groups = groups.copy()
        layout = QVBoxLayout(self)
        tip = QLabel("每行是一组等价写法：选择字段，填写统一名称，再用中文分号“；”分隔其他写法。一组可包含多个别名。")
        tip.setWordWrap(True)
        layout.addWidget(tip)
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["适用字段", "统一名称", "其他等价写法（用；分隔）"])
        self.table.setColumnWidth(0, 110)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        layout.addWidget(self.table, 1)
        for field_name, field_groups in groups.items():
            for group in field_groups:
                self.add_row(field_name, group[0], "；".join(group[1:]))
        row_buttons = QHBoxLayout()
        add = QPushButton("添加一组")
        add.clicked.connect(lambda: self.add_row())
        remove = QPushButton("删除选中行")
        remove.clicked.connect(self.remove_row)
        row_buttons.addWidget(add)
        row_buttons.addWidget(remove)
        layout.addLayout(row_buttons)
        layout.addWidget(QLabel(f"配置文件：{SETTINGS_PATH}"))
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def add_row(self, field_name: str = "major", canonical: str = "", aliases: str = ""):
        row = self.table.rowCount()
        self.table.insertRow(row)
        field = QComboBox()
        for key, label in ALIAS_FIELDS.items():
            field.addItem(label, key)
        field.setCurrentIndex(max(0, field.findData(field_name)))
        self.table.setCellWidget(row, 0, field)
        self.table.setItem(row, 1, QTableWidgetItem(canonical))
        self.table.setItem(row, 2, QTableWidgetItem(aliases))

    def remove_row(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def validate_and_accept(self):
        groups: dict[str, list[list[str]]] = {}
        try:
            for row in range(self.table.rowCount()):
                canonical = self.table.item(row, 1).text().strip() if self.table.item(row, 1) else ""
                others = self.table.item(row, 2).text().strip() if self.table.item(row, 2) else ""
                if not canonical and not others:
                    continue
                if not canonical or not others:
                    raise ValueError(f"第 {row + 1} 组需要填写统一名称和至少一个其他写法")
                aliases = [part.strip() for part in others.split("；")]
                if not all(aliases):
                    raise ValueError(f"第 {row + 1} 组有空别名")
                field_name = self.table.cellWidget(row, 0).currentData()
                groups.setdefault(field_name, []).append([canonical, *aliases])
            validate_alias_groups(groups)
            self.result_alias_groups = groups
            self.accept()
        except ValueError as error:
            QMessageBox.warning(self, "别名设置有误", str(error))


class IgnoreDifferenceDialog(QDialog):
    def __init__(self, item: Submission, target: Reference, config: FormatConfig, parent=None):
        super().__init__(parent)
        self.setWindowTitle("忽略本差异 · 预览长期规则")
        self.resize(760, 500)
        self.config = config
        self.result_groups = config.alias_groups
        self.differences = text_differences(item, target, config)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"申报：{item.name}   →   参考：{target.name}（{target.id}）"))
        layout.addWidget(QLabel("勾选要视为等价的文字差异。确认后会写入长期规则，并影响后续所有表格的匹配。"))
        self.checks = []
        for field_name, submitted, official in self.differences:
            check = QCheckBox(f"{ALIAS_FIELDS[field_name]}：{submitted}  ⇄  {official}")
            check.setChecked(True)
            check.toggled.connect(self.update_preview)
            layout.addWidget(check)
            self.checks.append(check)
        if item.score != target.score:
            layout.addWidget(QLabel("分数差异不会加入文字等价规则，仍需单独核实。"))
        layout.addWidget(QLabel("将保存的规则预览"))
        self.preview = QPlainTextEdit()
        self.preview.setReadOnly(True)
        layout.addWidget(self.preview, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("确认加入长期规则")
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.update_preview()

    def selected_groups(self) -> dict[str, list[list[str]]]:
        groups = self.config.alias_groups
        for check, (field_name, submitted, official) in zip(self.checks, self.differences):
            if check.isChecked():
                groups = add_equivalence_group(groups, field_name, submitted, official)
        return groups

    def update_preview(self):
        try:
            groups = self.selected_groups()
            lines = [f"{ALIAS_FIELDS[field]}：{' = '.join(group)}"
                for field, field_groups in groups.items() for group in field_groups]
            self.preview.setPlainText("\n".join(lines) or "未选择任何可加入的文字差异")
        except ValueError as error:
            self.preview.setPlainText(str(error))

    def validate_and_accept(self):
        try:
            groups = self.selected_groups()
            if groups == self.config.alias_groups:
                raise ValueError("没有选中新的文字差异")
            self.result_groups = groups
            self.accept()
        except ValueError as error:
            QMessageBox.warning(self, "无法加入规则", str(error))


class ApplyOriginalDialog(QDialog):
    def __init__(self, item: Submission, target: Reference, edits: list[ProposedEdit], parent=None):
        super().__init__(parent)
        self.setWindowTitle("应用到原表 · 核对写入内容")
        self.resize(900, 500)
        self.edits = edits
        self.selected_edits: list[ProposedEdit] = []
        layout = QVBoxLayout(self)
        warning = QLabel(f"将直接修改：{item.path}\n请确认已备份，并逐项核对新旧值。图片等其他工作表会保留。")
        warning.setWordWrap(True)
        layout.addWidget(warning)
        self.table = QTableWidget(len(edits), 5)
        self.table.setHorizontalHeaderLabels(["应用", "字段", "原表位置", "原值", "参考表值"])
        self.table.setColumnWidth(0, 55)
        self.table.setColumnWidth(1, 90)
        self.table.setColumnWidth(2, 145)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        layout.addWidget(self.table, 1)
        for row, edit in enumerate(edits):
            check = QTableWidgetItem()
            check.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            check.setCheckState(Qt.Unchecked if edit.affected_range else Qt.Checked)
            self.table.setItem(row, 0, check)
            values = [ALIAS_FIELDS.get(edit.field_name, "分数"),
                edit.address + (f"（合并 {edit.affected_range}）" if edit.affected_range else ""),
                "" if edit.old is None else str(edit.old), str(edit.new)]
            for col, value in enumerate(values, 1):
                self.table.setItem(row, col, QTableWidgetItem(value))
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("确认写入原始文件")
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def validate_and_accept(self):
        self.selected_edits = [edit for row, edit in enumerate(self.edits)
            if self.table.item(row, 0).checkState() == Qt.Checked]
        if not self.selected_edits:
            QMessageBox.information(self, "没有选中修改", "请勾选至少一项要写入原表的内容。")
            return
        self.accept()


class PhotoView(QGraphicsView):
    """Show evidence at its original resolution, with fit, zoom and drag controls."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setScene(QGraphicsScene(self))
        self.setRenderHint(QPainter.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorViewCenter)
        self.setMinimumSize(580, 480)
        self._has_photo = False
        self._fit_mode = True
        self.show_message("暂无图片")

    def show_message(self, message: str):
        self.scene().clear()
        self._has_photo = False
        self._fit_mode = True
        self.resetTransform()
        label = self.scene().addText(message or "暂无图片")
        label.setDefaultTextColor(QColor("#526477"))
        self.scene().setSceneRect(label.boundingRect().adjusted(-20, -20, 20, 20))

    def set_photo(self, pixmap: QPixmap):
        self.scene().clear()
        self.scene().addPixmap(pixmap)
        self.scene().setSceneRect(0, 0, pixmap.width(), pixmap.height())
        self._has_photo = True
        self.fit_photo()

    def fit_photo(self):
        if not self._has_photo:
            return
        self._fit_mode = True
        self.resetTransform()
        self.fitInView(self.sceneRect(), Qt.KeepAspectRatio)

    def actual_size(self):
        if not self._has_photo:
            return
        self._fit_mode = False
        self.resetTransform()
        self.centerOn(self.sceneRect().center())

    def zoom(self, factor: float):
        if not self._has_photo:
            return
        scale = self.transform().m11() * factor
        if 0.02 <= scale <= 20:
            self._fit_mode = False
            self.scale(factor, factor)

    def wheelEvent(self, event):
        if self._has_photo and event.angleDelta().y():
            self.zoom(1.2 if event.angleDelta().y() > 0 else 1 / 1.2)
            event.accept()
        else:
            super().wheelEvent(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._fit_mode and self._has_photo:
            self.fitInView(self.sceneRect(), Qt.KeepAspectRatio)


class BatchReviewWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("学院批量核对 · 独立电脑端")
        self.resize(1580, 900)
        self.reference_path: Path | None = None
        self.source_paths: list[Path] = []
        try:
            groups = load_alias_settings(SETTINGS_PATH)
        except (OSError, ValueError) as error:
            groups = {}
            QMessageBox.warning(self, "长期设置读取失败", str(error))
        self.config = replace(FormatConfig(), alias_groups=groups)
        self.references: list[Reference] = []
        self.items: list[Submission] = []
        self.visible: list[Submission] = []
        self.candidates = []
        self.current: Submission | None = None
        self.progress_path: Path | None = None
        self.dirty = False
        self.source_hashes: dict[Path, str] = {}
        self.force_fingerprints = False
        self.excel_processes: set[QProcess] = set()
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        bar = QHBoxLayout()
        for label, action in [
            ("选择参考大表", self.choose_reference), ("加入个人表格", self.add_sources),
            ("备份全部表格", self.backup_selected_workbooks),
            ("设置格式", self.configure), ("等价写法（长期）", self.configure_aliases),
            ("开始扫描", self.scan),
            ("打开进度", self.open_progress), ("保存进度", self.save_progress),
            ("导出审核副本", self.export_copies),
        ]:
            button = QPushButton(label)
            button.clicked.connect(action)
            bar.addWidget(button)
        layout.addLayout(bar)
        self.info = QLabel("先选择参考大表和全部个人表格，再设置格式、开始扫描。")
        self.info.setWordWrap(True)
        layout.addWidget(self.info)
        split = QSplitter(Qt.Horizontal)
        layout.addWidget(split, 1)
        left = QWidget()
        left_layout = QVBoxLayout(left)
        self.filter = QComboBox()
        self.filter.addItems(["全部", "待确认", "参考分数待评审", "匹配", "有问题", "证明待审", "证明不足"])
        self.filter.currentIndexChanged.connect(self.refresh)
        left_layout.addWidget(self.filter)
        left_layout.addWidget(QLabel("双击条目定位原表总表；右键可定位总表或证明。"))
        self.source_table = QTableWidget(0, 6)
        self.source_table.setHorizontalHeaderLabels(["状态", "文件", "大类/序号", "活动名称", "等级/分数", "证明"])
        self._table(self.source_table)
        self.source_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.source_table.itemSelectionChanged.connect(self.select_source)
        self.source_table.itemDoubleClicked.connect(lambda cell: self.open_summary_excel(self.visible[cell.row()]))
        self.source_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.source_table.customContextMenuRequested.connect(self.source_context_menu)
        left_layout.addWidget(self.source_table)
        split.addWidget(left)
        tabs = QTabWidget()
        match_tab = QWidget()
        match_layout = QVBoxLayout(match_tab)
        self.source_title = QLabel("选择左侧条目")
        self.source_title.setWordWrap(True)
        match_layout.addWidget(self.source_title)
        match_layout.addWidget(QLabel("条目匹配度用于定位候选；分数是否正确请看右侧“差异”列。"))
        self.candidate_table = QTableWidget(0, 6)
        self.candidate_table.setHorizontalHeaderLabels(["条目匹配度", "Sheet/行", "大类/小类", "活动", "等级/分数", "差异"])
        self._table(self.candidate_table)
        self.candidate_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.candidate_table.itemDoubleClicked.connect(lambda _: self.jump_reference())
        match_layout.addWidget(self.candidate_table, 1)
        actions = QHBoxLayout()
        for label, action in [("确认对应", self.confirm_match), ("标记不匹配", self.mark_wrong),
                              ("暂时跳过", self.skip_match), ("定位参考项", self.jump_reference),
                              ("忽略本差异", self.ignore_difference),
                              ("应用到原表", self.apply_to_original)]:
            button = QPushButton(label)
            button.clicked.connect(action)
            actions.addWidget(button)
        match_layout.addLayout(actions)
        self.reference_search = QLineEdit()
        self.reference_search.setPlaceholderText("搜索全部参考条目：活动、大类、小类或等级")
        self.reference_search.textChanged.connect(self._fill_reference_table)
        match_layout.addWidget(self.reference_search)
        self.reference_table = QTableWidget(0, 5)
        self.reference_table.setHorizontalHeaderLabels(["位置", "大类/小类", "活动", "等级", "参考分数"])
        self._table(self.reference_table)
        self.reference_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        match_layout.addWidget(QLabel("参考大表定位视图"))
        match_layout.addWidget(self.reference_table, 1)
        full_match = QPushButton("确认选中的参考条目")
        full_match.clicked.connect(self.confirm_reference)
        match_layout.addWidget(full_match)
        tabs.addTab(match_tab, "活动与分数")
        proof_tab = QWidget()
        proof_layout = QVBoxLayout(proof_tab)
        self.proof_title = QLabel("选择左侧条目查看序号对应的证明")
        self.proof_title.setWordWrap(True)
        self.proof_title.setStyleSheet(
            "font-size: 18pt; font-weight: 700; color: #164C70; "
            "background: #EAF4FA; padding: 10px; border-radius: 6px;"
        )
        proof_layout.addWidget(self.proof_title)
        self.proof_meta = QLabel("图片可用鼠标滚轮缩放、按住拖动查看。")
        self.proof_meta.setWordWrap(True)
        self.proof_meta.setStyleSheet("font-size: 11pt; color: #43596B;")
        proof_layout.addWidget(self.proof_meta)
        self.photo_list = QComboBox()
        self.photo_list.currentIndexChanged.connect(self.show_photo)
        proof_layout.addWidget(self.photo_list)
        photo_controls = QHBoxLayout()
        for label, action in [("适应窗口", self.photo_fit), ("100%", self.photo_actual_size),
                              ("放大 +", self.photo_zoom_in), ("缩小 −", self.photo_zoom_out)]:
            button = QPushButton(label)
            button.clicked.connect(action)
            photo_controls.addWidget(button)
        photo_controls.addStretch(1)
        proof_layout.addLayout(photo_controls)
        self.photo = PhotoView()
        proof_layout.addWidget(self.photo, 1)
        proof_actions = QHBoxLayout()
        for label, action in [("证明足够", lambda: self.set_proof("enough")),
                              ("证明不足", lambda: self.set_proof("insufficient")),
                              ("下一条待审", self.next_proof)]:
            button = QPushButton(label)
            button.clicked.connect(action)
            proof_actions.addWidget(button)
        proof_layout.addLayout(proof_actions)
        tabs.addTab(proof_tab, "证明审核")
        split.addWidget(tabs)
        split.setSizes([680, 900])
        QShortcut(QKeySequence("Alt+Y"), self, activated=lambda: self.set_proof("enough"))
        QShortcut(QKeySequence("Alt+N"), self, activated=lambda: self.set_proof("insufficient"))

    @staticmethod
    def _table(table):
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.setEditTriggers(QTableWidget.NoEditTriggers)

    def _error(self, title: str, error: Exception):
        QMessageBox.critical(self, title, str(error))

    def source_context_menu(self, position):
        cell = self.source_table.itemAt(position)
        if cell is None or not 0 <= cell.row() < len(self.visible):
            return
        item = self.visible[cell.row()]
        self.source_table.selectRow(cell.row())
        menu = self._source_menu(item)
        menu.exec(self.source_table.viewport().mapToGlobal(position))
        menu.deleteLater()

    def _source_menu(self, item: Submission) -> QMenu:
        menu = QMenu(self)
        menu.addAction("在 Excel 中定位总表条目", lambda: self.open_summary_excel(item))
        menu.addAction("在 Excel 中定位证明", lambda: self.open_evidence_excel(item))
        return menu

    def open_summary_excel(self, item: Submission):
        address = f"{get_column_letter(self.config.summary_name + 1)}{item.row}"
        self._navigate_excel(item.path, self.config.summary_sheet, address)

    def open_evidence_excel(self, item: Submission):
        if item.evidence_sheet is None or item.evidence_row is None:
            QMessageBox.information(self, "无法定位证明",
                f"{item.path.name} · {item.major} 序号 {item.number}：{item.image_warning or '未找到唯一的证明行'}。"
                "请检查分类证明 Sheet 映射及序号。")
            return
        image_index = self.photo_list.currentIndex() if self.current is item else 0
        if item.photos:
            address = item.photos[max(0, min(image_index, len(item.photos) - 1))].cell
        else:
            address = f"{get_column_letter(self.config.evidence_image + 1)}{item.evidence_row}"
        self._navigate_excel(item.path, item.evidence_sheet, address)

    def _navigate_excel(self, path: Path, sheet: str, address: str):
        if not path.is_file():
            QMessageBox.warning(self, "无法打开原表", f"找不到文件：{path}")
            return
        process = QProcess(self)
        process.setProgram("powershell.exe")
        process.setArguments(["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
            "-File", str(EXCEL_NAVIGATION_SCRIPT), "-WorkbookPath", str(path.resolve()),
            "-SheetName", sheet, "-CellAddress", address])
        process.finished.connect(lambda code, _status, p=process:
            self._excel_navigation_finished(p, code, path, sheet, address))
        process.errorOccurred.connect(lambda _error, p=process:
            self._excel_navigation_error(p))
        self.excel_processes.add(process)
        self.statusBar().showMessage(f"正在 Excel 中定位：{path.name} · {sheet}!{address}")
        process.start()

    def _excel_navigation_error(self, process: QProcess):
        if process.property("excelReported"):
            return
        process.setProperty("excelReported", True)
        self.excel_processes.discard(process)
        QMessageBox.warning(self, "无法启动 Excel 定位", process.errorString())
        process.deleteLater()

    def _excel_navigation_finished(self, process: QProcess, code: int,
                                   path: Path, sheet: str, address: str):
        if process.property("excelReported"):
            return
        self.excel_processes.discard(process)
        if code == 0:
            self.statusBar().showMessage(f"已在 Excel 中定位：{path.name} · {sheet}!{address}", 5000)
        else:
            detail = bytes(process.readAllStandardError()).decode("utf-8", errors="replace").strip()
            QMessageBox.warning(self, "Excel 定位失败", detail[-1000:] or
                "请确认电脑已安装 Microsoft Excel，且目标文件可以正常打开。")
        process.deleteLater()

    def _discard_review(self) -> bool:
        if self.dirty:
            answer = QMessageBox.question(self, "未保存的核对进度",
                "当前人工判定尚未保存。继续会丢失这些判定，确定继续吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if answer != QMessageBox.Yes:
                return False
        self.items = []
        self.references = []
        self.current = None
        self.progress_path = None
        self.dirty = False
        self.source_hashes = {}
        self.force_fingerprints = False
        self.source_table.setRowCount(0)
        self.candidate_table.setRowCount(0)
        self.reference_table.setRowCount(0)
        return True

    def choose_reference(self):
        selected, _ = QFileDialog.getOpenFileName(self, "选择参考大表", "", EXCEL_FILTER)
        if selected:
            if not self._discard_review(): return
            self.reference_path = Path(selected)
            self.info.setText(f"参考大表：{self.reference_path.name}；已选个人表格 {len(self.source_paths)} 份")

    def add_sources(self):
        selected, _ = QFileDialog.getOpenFileNames(self, "选择全部待确认表格", "", EXCEL_FILTER)
        if selected:
            candidates = list(dict.fromkeys(self.source_paths + [Path(p) for p in selected]))
            if not self._prompt_import_backup(candidates): return
            if not self._discard_review(): return
            self.source_paths = candidates
            self.info.setText(f"参考大表：{self.reference_path.name if self.reference_path else '未选'}；已选个人表格 {len(self.source_paths)} 份")

    def _all_selected_workbooks(self, sources: list[Path] | None = None,
                                reference: Path | None = None) -> list[Path]:
        selected_reference = reference if reference is not None else self.reference_path
        files = [selected_reference] if selected_reference else []
        files.extend(sources if sources is not None else self.source_paths)
        return list(dict.fromkeys(files))

    def _choose_backup_folder(self, sources: list[Path] | None = None,
                              reference: Path | None = None) -> bool:
        folder = QFileDialog.getExistingDirectory(self, "选择全部表格的备份文件夹")
        if not folder:
            return False
        try:
            copied = backup_workbooks(self._all_selected_workbooks(sources, reference), Path(folder))
            QMessageBox.information(self, "备份完成", f"已备份 {len(copied)} 份 Excel 到：\n{folder}")
            return True
        except Exception as error:
            self._error("备份失败", error)
            return False

    def backup_selected_workbooks(self):
        if not self._all_selected_workbooks():
            QMessageBox.information(self, "没有文件", "请先选择参考大表和个人表格。")
            return
        self._choose_backup_folder()

    def _prompt_import_backup(self, sources: list[Path], reference: Path | None = None) -> bool:
        prompt = QMessageBox(self)
        prompt.setIcon(QMessageBox.Warning)
        prompt.setWindowTitle("导入前建议备份")
        prompt.setText("审核时使用“应用到原表”会直接修改个人提交的 Excel。建议现在备份。")
        prompt.setInformativeText(
            f"将导入 {len(sources)} 份个人表格；备份操作也会包含已选的参考大表。导入和扫描本身不会修改文件。"
        )
        backup_button = prompt.addButton("选择文件夹并备份", QMessageBox.AcceptRole)
        continue_button = prompt.addButton("暂不备份，继续导入", QMessageBox.DestructiveRole)
        prompt.addButton("取消导入", QMessageBox.RejectRole)
        prompt.exec()
        if prompt.clickedButton() == backup_button:
            return self._choose_backup_folder(sources, reference)
        return prompt.clickedButton() == continue_button

    def configure(self):
        dialog = FormatDialog(self.config, self.reference_path,
            self.source_paths[0] if self.source_paths else None, self)
        if dialog.exec() == QDialog.Accepted:
            new_config = dialog.result_config()
            if new_config != self.config and not self._discard_review(): return
            self.config = new_config
            self.statusBar().showMessage("本批格式设置已更新", 4000)

    def configure_aliases(self):
        dialog = AliasDialog(self.config.alias_groups, self)
        if dialog.exec() != QDialog.Accepted:
            return
        groups = dialog.result_alias_groups
        if groups == self.config.alias_groups:
            return
        if not self._discard_review():
            return
        try:
            save_alias_settings(SETTINGS_PATH, groups)
            self.config = replace(self.config, alias_groups=groups, category_aliases={})
            self.info.setText("等价写法已长期保存。请重新扫描，使新规则应用于本批表格。")
        except Exception as error:
            self._error("保存等价写法失败", error)

    def scan(self):
        if not self.reference_path or not self.source_paths:
            QMessageBox.information(self, "先选择文件", "请选择参考大表和至少一份个人表格。")
            return
        try:
            references = load_references(self.reference_path, self.config)
            items = load_submissions(self.source_paths, self.config)
            if self.items and not self._discard_review(): return
            self.references = references
            self.items = items
            self.source_hashes = {path: file_fingerprint(path) for path in self.source_paths}
            self.progress_path = None
            self.dirty = False
            self._fill_reference_table()
            self.refresh()
        except Exception as error:
            self._error("扫描失败", error)

    def _fill_reference_table(self):
        self.reference_table.setRowCount(0)
        self.visible_references = []
        query = self.reference_search.text().strip().lower()
        for ref in self.references:
            if query and query not in " ".join((ref.sheet, ref.major, ref.minor, ref.name, ref.role)).lower():
                continue
            row = self.reference_table.rowCount()
            self.reference_table.insertRow(row)
            self.visible_references.append(ref)
            for col, text in enumerate([ref.id, f"{ref.major}/{ref.minor}", ref.name, ref.role,
                                        "" if ref.score is None else str(ref.score)]):
                self.reference_table.setItem(row, col, QTableWidgetItem(text))

    def refresh(self):
        selected_id = self.current.id if self.current else ""
        self.source_table.blockSignals(True)
        self.source_table.setRowCount(0)
        self.visible = []
        counts = {"matched": 0, "pending": 0, "review_score": 0, "wrong": 0}
        for item in self.items:
            state, _ = status(item, self.references, self.config.category_aliases, self.config.alias_groups)
            counts[state] += 1
            selected = self.filter.currentText()
            if selected == "待确认" and state not in {"pending", "review_score"}: continue
            if selected == "参考分数待评审" and state != "review_score": continue
            if selected == "匹配" and state != "matched": continue
            if selected == "有问题" and state != "wrong": continue
            if selected == "证明待审" and item.proof != "pending": continue
            if selected == "证明不足" and item.proof != "insufficient": continue
            row = self.source_table.rowCount()
            self.source_table.insertRow(row)
            self.visible.append(item)
            values = [{"matched": "匹配", "pending": "待确认", "review_score": "分数待评审", "wrong": "有问题"}[state],
                item.path.name, f"{item.major} / {item.number}", item.name,
                f"{item.role} / {item.score if item.score is not None else '空'}",
                {"pending": "待审", "enough": "足够", "insufficient": "不足"}.get(item.proof, item.proof)]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                cell.setBackground(QColor({"matched": "#E2F0D9", "pending": "#FCE4D6",
                    "review_score": "#FCE4D6", "wrong": "#F4CCCC"}[state]))
                if col == 5 and item.proof == "insufficient":
                    cell.setBackground(QColor("#FFF2CC"))
                self.source_table.setItem(row, col, cell)
        self.source_table.blockSignals(False)
        self.info.setText(f"参考条目 {len(self.references)}；个人表格 {len(self.source_paths)} 份、活动 {len(self.items)} 条；"
            f"匹配 {counts['matched']}，待确认 {counts['pending']}，参考分数待评审 {counts['review_score']}，"
            f"有问题 {counts['wrong']}。")
        if self.visible:
            row = next((i for i, item in enumerate(self.visible) if item.id == selected_id), 0)
            self.source_table.selectRow(row)
        else:
            self.current = None
            self.candidate_table.setRowCount(0)
            self.proof_title.setText("当前筛选无条目")
            self.proof_meta.clear()
            self.photo_list.blockSignals(True)
            self.photo_list.clear()
            self.photo_list.blockSignals(False)
            self.photo.show_message("当前筛选无条目")

    def select_source(self):
        row = self.source_table.currentRow()
        if not 0 <= row < len(self.visible): return
        self.current = self.visible[row]
        item = self.current
        self.source_title.setText(f"{item.path.name} · 总表第 {item.row} 行 · {item.major} 序号 {item.number}\n"
            f"{item.name} / {item.role} / 申报 {item.score}")
        self.candidates = rank(item, self.references, aliases=self.config.category_aliases,
            alias_groups=self.config.alias_groups)
        self.candidate_table.setRowCount(0)
        for match in self.candidates:
            ref = match.target
            row = self.candidate_table.rowCount()
            self.candidate_table.insertRow(row)
            values = [f"{match.similarity:.0%}", ref.id, f"{ref.major}/{ref.minor}", ref.name,
                f"{ref.role} / {ref.score}", match.differences]
            for col, value in enumerate(values):
                self.candidate_table.setItem(row, col, QTableWidgetItem(value))
            if item.decision == ref.id:
                self.candidate_table.selectRow(row)
        self.proof_title.setText(f"序号 {item.number} · {item.name}")
        self.proof_meta.setText(f"{item.path.name} · {item.major} · "
            f"证明页：{item.photos[0].sheet if item.photos else '未识别'}；"
            f"{item.image_warning or f'找到 {len(item.photos)} 张单元格图片'}"
            "。滚轮缩放，按住图片拖动查看。")
        self.photo_list.blockSignals(True)
        self.photo_list.clear()
        for image in item.photos:
            self.photo_list.addItem(f"{image.sheet}!{image.cell}")
        self.photo_list.blockSignals(False)
        self.show_photo(0)

    def show_photo(self, index: int):
        if not self.current or not 0 <= index < len(self.current.photos):
            self.photo.show_message(self.current.image_warning if self.current else "暂无图片")
            return
        try:
            image = self.current.photos[index]
            with ZipFile(self.current.path) as archive:
                data = archive.read(image.member)
            pixmap = QPixmap()
            if not pixmap.loadFromData(data):
                raise ValueError("图片编码不受 Qt 支持")
            self.photo.set_photo(pixmap)
        except Exception as error:
            self.photo.show_message("图片预览失败，请查看原始 Excel")
            self._error("预览证明失败", error)

    def photo_fit(self):
        self.photo.fit_photo()

    def photo_actual_size(self):
        self.photo.actual_size()

    def photo_zoom_in(self):
        self.photo.zoom(1.2)

    def photo_zoom_out(self):
        self.photo.zoom(1 / 1.2)

    def confirm_match(self):
        if not self.current: return
        row = self.candidate_table.currentRow()
        if not 0 <= row < len(self.candidates):
            QMessageBox.information(self, "选择候选", "请先选中一个参考大表候选条目。")
            return
        self.current.decision = self.candidates[row].target.id
        self._changed()

    def _selected_candidate_target(self) -> Reference | None:
        row = self.candidate_table.currentRow()
        if self.current is None or not 0 <= row < len(self.candidates):
            QMessageBox.information(self, "选择候选", "请先选择左侧活动，再选中一条参考大表候选。")
            return None
        return self.candidates[row].target

    def _reload_with_config(self, config: FormatConfig) -> list[Submission]:
        states = {item.id: (item.decision, item.proof) for item in self.items}
        refreshed = load_submissions(self.source_paths, config)
        for item in refreshed:
            if item.id in states:
                item.decision, item.proof = states[item.id]
        return refreshed

    def ignore_difference(self):
        target = self._selected_candidate_target()
        if target is None:
            return
        dialog = IgnoreDifferenceDialog(self.current, target, self.config, self)
        if not dialog.differences:
            QMessageBox.information(self, "没有可加入的文字差异",
                "该候选没有尚未忽略的文字差异。分数差异不会加入等价规则。")
            return
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            config = replace(self.config, alias_groups=dialog.result_groups, category_aliases={})
            refreshed = self._reload_with_config(config)
            save_alias_settings(SETTINGS_PATH, dialog.result_groups)
            self.config = config
            self.items = refreshed
            prior_dirty = self.dirty
            self._changed()
            if self.progress_path is None:
                self.dirty = prior_dirty
            self.statusBar().showMessage("已加入长期规则，并重新匹配当前批次", 6000)
        except Exception as error:
            self._error("加入规则失败", error)

    def apply_to_original(self):
        target = self._selected_candidate_target()
        if target is None:
            return
        item = self.current
        try:
            expected = self.source_hashes.get(item.path)
            if expected and file_fingerprint(item.path) != expected:
                raise ValueError(f"{item.path.name} 在扫描后被其他程序修改。请重新扫描后再写入。")
            edits = propose_original_edits(item, target, self.config)
            if not edits:
                QMessageBox.information(self, "没有可写入的差异",
                    "该候选的原表文字和数值已与参考表相同，或参考分数需要人工认定。")
                return
            dialog = ApplyOriginalDialog(item, target, edits, self)
            if dialog.exec() != QDialog.Accepted:
                return
            apply_edits_to_original(item.path, self.config.summary_sheet, dialog.selected_edits)
            self.source_hashes[item.path] = file_fingerprint(item.path)
            states = {existing.id: (existing.decision, existing.proof) for existing in self.items}
            reloaded = load_submissions([item.path], self.config)
            for refreshed in reloaded:
                if refreshed.id in states:
                    refreshed.decision, refreshed.proof = states[refreshed.id]
            self.items = [existing for existing in self.items if existing.path != item.path] + reloaded
            self.items.sort(key=lambda row: (self.source_paths.index(row.path), row.row))
            self.force_fingerprints = True
            self._changed()
            self.statusBar().showMessage(f"已写入原表：{item.path.name}", 6000)
        except Exception as error:
            self._error("应用到原表失败", error)

    def mark_wrong(self):
        if self.current:
            self.current.decision = ""
            self._changed()

    def skip_match(self):
        if self.current:
            self.current.decision = None
            self._changed()

    def jump_reference(self):
        row = self.candidate_table.currentRow()
        if not 0 <= row < len(self.candidates): return
        target = self.candidates[row].target
        if self.reference_search.text():
            self.reference_search.clear()
        index = next((i for i, ref in enumerate(self.visible_references) if ref.id == target.id), -1)
        if index >= 0:
            self.reference_table.selectRow(index)
            self.reference_table.scrollToItem(self.reference_table.item(index, 0))
            self.statusBar().showMessage(f"参考大表位置：{target.id}", 5000)

    def confirm_reference(self):
        if not self.current: return
        row = self.reference_table.currentRow()
        if not 0 <= row < len(self.visible_references):
            QMessageBox.information(self, "选择参考项", "请先在参考大表定位视图中选择一行。")
            return
        self.current.decision = self.visible_references[row].id
        self._changed()

    def set_proof(self, value: str):
        if not self.current: return
        if value == "enough" and not self.current.photos:
            answer = QMessageBox.question(self, "请先核实原表",
                "没有可预览的单元格图片。若已在原始 Excel 中人工核实证明足够，点击“是”继续。",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if answer != QMessageBox.Yes: return
        reviewed_index = self.items.index(self.current)
        self.current.proof = value
        self._changed()
        self._select_next_pending(reviewed_index)

    def next_proof(self):
        if not self.items: return
        self._select_next_pending(self.items.index(self.current) if self.current in self.items else -1)

    def _select_next_pending(self, after_index: int):
        next_item = next((self.items[i % len(self.items)]
            for i in range(after_index + 1, after_index + 1 + len(self.items))
            if self.items[i % len(self.items)].proof == "pending"), None)
        if next_item:
            if next_item not in self.visible:
                self.filter.setCurrentText("证明待审")
            row = next((i for i, item in enumerate(self.visible) if item.id == next_item.id), -1)
            if row >= 0:
                self.source_table.selectRow(row)
                self.source_table.scrollToItem(self.source_table.item(row, 0))
        else:
            self.statusBar().showMessage("证明已全部审核", 5000)

    def _changed(self):
        self.dirty = True
        if self.progress_path:
            try:
                save_progress(self.progress_path, self.reference_path, self.config, self.items,
                    refresh_fingerprints=self.force_fingerprints)
                self.dirty = False
                self.force_fingerprints = False
            except Exception as error:
                self._error("自动保存失败", error)
        self.refresh()

    def save_progress(self):
        if not self.items or not self.reference_path: return
        path = self.progress_path
        if path is None:
            selected, _ = QFileDialog.getSaveFileName(self, "保存核对进度", "批量核对进度.json", "JSON (*.json)")
            if not selected: return
            path = Path(selected)
        try:
            save_progress(path, self.reference_path, self.config, self.items,
                refresh_fingerprints=self.force_fingerprints)
            self.progress_path = path
            self.dirty = False
            self.force_fingerprints = False
            self.statusBar().showMessage(f"进度已保存：{path}", 5000)
        except Exception as error:
            self._error("保存进度失败", error)

    def open_progress(self):
        selected, _ = QFileDialog.getOpenFileName(self, "打开核对进度", "", "JSON (*.json)")
        if not selected: return
        try:
            reference_path, config, items = restore_progress(Path(selected), alias_groups=self.config.alias_groups)
            references = load_references(reference_path, config)
            paths = list(dict.fromkeys(item.path for item in items))
            if not self._prompt_import_backup(paths, reference_path): return
            if not self._discard_review(): return
            self.reference_path, self.config, self.items = reference_path, config, items
            self.references = references
            self.source_paths = paths
            self.source_hashes = {path: file_fingerprint(path) for path in self.source_paths}
            self.progress_path = Path(selected)
            self.dirty = False
            self.force_fingerprints = False
            self._fill_reference_table()
            self.refresh()
        except Exception as error:
            self._error("打开进度失败", error)

    def export_copies(self):
        if not self.items: return
        selected = QFileDialog.getExistingDirectory(self, "选择审核副本输出文件夹")
        if not selected: return
        try:
            output = export_review_copies(Path(selected), self.config, self.items, self.references)
            self.statusBar().showMessage(f"已导出 {len(output)} 份审核副本到 {selected}", 8000)
            QMessageBox.information(self, "导出完成", f"已生成 {len(output)} 份审核副本。原始提交文件未修改。")
        except Exception as error:
            self._error("导出失败", error)

    def closeEvent(self, event):  # type: ignore[override]
        if self.dirty:
            answer = QMessageBox.question(self, "核对进度尚未保存",
                "当前人工判定尚未保存。现在关闭会丢失这些判定，确定关闭吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if answer != QMessageBox.Yes:
                event.ignore()
                return
        event.accept()


def show_loaded_alias_warning(window: BatchReviewWindow) -> None:
    groups = window.config.alias_groups
    rules = [f"{ALIAS_FIELDS[field]}：{' = '.join(group)}"
        for field, field_groups in groups.items() for group in field_groups]
    if not rules:
        return
    visible = rules[:8]
    if len(rules) > 8:
        visible.append(f"……另有 {len(rules) - 8} 条，点击“显示详细信息”查看全部")
    warning = QMessageBox(window)
    warning.setIcon(QMessageBox.Warning)
    warning.setWindowTitle("已加载长期等价写法规则")
    warning.setTextFormat(Qt.PlainText)
    warning.setText(
        f"本次启动加载了 {len(rules)} 组长期等价写法规则。请先确认它们适用于今年，再开始扫描。"
    )
    warning.setInformativeText("\n".join(visible) +
        "\n\n若规则已过期，请在“等价写法（长期）”中修改或删除。")
    warning.setDetailedText("配置文件：" + str(SETTINGS_PATH) + "\n\n" + "\n".join(rules))
    warning.addButton("已核对", QMessageBox.AcceptRole)
    warning.exec()


def main() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft YaHei UI", 10))
    migration_error = None
    try:
        migrated_from = migrate_alias_settings(SETTINGS_PATH, legacy_settings_paths())
    except (OSError, ValueError) as error:
        migrated_from = None
        migration_error = error
    window = BatchReviewWindow()
    window.show()
    if migrated_from:
        window.statusBar().showMessage(f"旧规则已复制到：{SETTINGS_PATH}", 10000)
    if migration_error:
        QMessageBox.warning(window, "旧规则迁移失败",
            f"未能将旧规则复制到用户配置目录：{migration_error}\n旧文件仍保留，请检查后重试。")
    show_loaded_alias_warning(window)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
