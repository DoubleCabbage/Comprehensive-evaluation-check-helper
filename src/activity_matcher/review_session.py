"""Human review state and its export. Original workbooks are never modified."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from hashlib import sha256
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .review_core import (
    Candidate,
    ScoreRecord,
    SourceRecord,
    automatic_status,
    load_phone_export,
    load_score_table,
    rank_candidates,
)


def fingerprint(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


@dataclass
class ReviewSession:
    source_path: Path
    target_path: Path
    sheet: str
    header_row: int
    mapping: dict[str, int | None]
    sources: list[SourceRecord]
    targets: list[ScoreRecord]
    decisions: dict[str, str | None] = field(default_factory=dict)
    source_hash: str = ""
    target_hash: str = ""
    _candidate_cache: dict[str, list[Candidate]] = field(default_factory=dict, repr=False)

    @classmethod
    def create(
        cls,
        source_path: Path,
        target_path: Path,
        sheet: str,
        header_row: int,
        mapping: dict[str, int | None],
    ) -> ReviewSession:
        sources = load_phone_export(source_path)
        targets = load_score_table(target_path, sheet, header_row, mapping)
        return cls(
            source_path, target_path, sheet, header_row, mapping, sources, targets,
            source_hash=fingerprint(source_path), target_hash=fingerprint(target_path),
        )

    @classmethod
    def open(cls, path: Path) -> ReviewSession:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if payload.get("version") != 1:
            raise ValueError("核对进度文件版本不受支持")
        source_path = Path(payload["source_path"])
        target_path = Path(payload["target_path"])
        if not source_path.is_file() or not target_path.is_file():
            raise ValueError("核对进度所引用的 Excel 文件已移动或删除")
        if fingerprint(source_path) != payload["source_hash"] or fingerprint(target_path) != payload["target_hash"]:
            raise ValueError("Excel 文件内容已变化，请重新导入，以免旧确认结果对应错行")
        session = cls.create(
            source_path, target_path, payload["sheet"], payload["header_row"], payload["mapping"]
        )
        valid_source_ids = {source.record_id for source in session.sources}
        valid_target_ids = {target.record_id for target in session.targets}
        for source_id, target_id in payload.get("decisions", {}).items():
            if source_id not in valid_source_ids or target_id is not None and target_id not in valid_target_ids:
                raise ValueError("核对进度包含无效的条目 ID")
        session.decisions = payload.get("decisions", {})
        return session

    def save(self, path: Path) -> None:
        payload = {
            "version": 1,
            "source_path": str(self.source_path.resolve()),
            "target_path": str(self.target_path.resolve()),
            "source_hash": self.source_hash,
            "target_hash": self.target_hash,
            "sheet": self.sheet,
            "header_row": self.header_row,
            "mapping": self.mapping,
            "decisions": self.decisions,
        }
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def candidates(self, source: SourceRecord, limit: int = 5) -> list[Candidate]:
        if limit < 1:
            raise ValueError("limit must be at least 1")
        if limit > 5:
            return rank_candidates(source, self.targets, limit)
        if source.record_id not in self._candidate_cache:
            self._candidate_cache[source.record_id] = rank_candidates(source, self.targets, 5)
        return self._candidate_cache[source.record_id][:limit]

    def linked_target(self, source: SourceRecord) -> ScoreRecord | None:
        if source.record_id in self.decisions:
            selected_id = self.decisions[source.record_id]
        else:
            _, selected_id = automatic_status(self.candidates(source), source.score)
        return next((item for item in self.targets if item.record_id == selected_id), None)

    def status(self, source: SourceRecord) -> str:
        if source.record_id in self.decisions:
            target = self.linked_target(source)
            if target is None:
                return "确认无对应"
            return "已确认·分数不同" if source.score != target.score else "已确认"
        return automatic_status(self.candidates(source), source.score)[0]

    def target_usage(self) -> Counter[str]:
        return Counter(
            target.record_id
            for source in self.sources
            if (target := self.linked_target(source)) is not None
        )

    def export_result(self, path: Path) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "核对结果"
        sheet.append([
            "状态", "领域", "序号", "导出活动名", "导出职务/等级", "导出分数",
            "新表活动名", "新表领域", "新表职务/等级", "新表分数", "新表位置",
            "照片数量", "凭证位置", "导出条目ID", "说明",
        ])
        usage = self.target_usage()
        used: set[str] = set()
        for source in self.sources:
            target = self.linked_target(source)
            if target:
                used.add(target.record_id)
            notes = []
            if target and usage[target.record_id] > 1:
                notes.append("多条导出记录对应同一新表条目，请核对")
            if source.record_id not in self.decisions and self.status(source) in {"近似待确认", "重复候选"}:
                notes.append("尚未人工确认")
            sheet.append([
                self.status(source), source.dimension, source.number, source.name, source.role,
                source.score, target.name if target else None, target.dimension if target else None,
                target.role if target else None, target.score if target else None,
                f"{target.sheet}!{target.row}" if target else None,
                len(source.photos), "、".join(f"{p.sheet}!{p.cell}" for p in source.photos),
                source.record_id, "；".join(notes),
            ])
        extra = workbook.create_sheet("新表未对应")
        extra.append(["领域", "活动名", "职务/等级", "分数", "原表位置"])
        for target in self.targets:
            if target.record_id not in used:
                extra.append([target.dimension, target.name, target.role, target.score, f"{target.sheet}!{target.row}"])
        for worksheet in workbook:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for cell in worksheet[1]:
                cell.fill = PatternFill("solid", fgColor="234B63")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(vertical="center")
            worksheet.row_dimensions[1].height = 27
            widths = [18, 16, 9, 38, 24, 13, 38, 16, 24, 13, 18, 12, 42, 20, 45]
            for index in range(1, worksheet.max_column + 1):
                worksheet.column_dimensions[worksheet.cell(1, index).column_letter].width = widths[index - 1]
        workbook.save(path)
