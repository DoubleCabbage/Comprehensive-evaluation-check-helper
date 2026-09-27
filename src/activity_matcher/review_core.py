"""Import and compare a phone export with a new scoring workbook."""

from __future__ import annotations

from dataclasses import dataclass
from difflib import SequenceMatcher
from functools import lru_cache
from io import BytesIO
from pathlib import Path
import posixpath
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.styles.fills import Fill

from .normalize import normalize_activity_name, normalize_reference_score, normalize_text


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RICH_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
NS = {"x": MAIN_NS, "p": PKG_NS, "r": REL_NS, "rv": RICH_NS}

HEADER_ALIASES = {
    "name": ("活动名", "活动名称", "附加分名称", "项目名称", "加分项目", "评分项目", "项目"),
    "dimension": ("领域", "所属维度", "维度", "指标名称", "类别", "一级分类"),
    "role": ("获奖等级", "职务/获奖等级", "职务", "等级", "身份", "获奖级别"),
    "score": ("分数", "参考分数", "认定分数", "得分", "评分", "加分"),
}


@dataclass(frozen=True)
class PhotoRef:
    sheet: str
    cell: str
    archive_member: str


@dataclass(frozen=True)
class SourceRecord:
    record_id: str
    dimension: str
    number: int
    name: str
    role: str
    score: float | None
    sheet: str
    row: int
    photos: tuple[PhotoRef, ...] = ()


@dataclass(frozen=True)
class ScoreRecord:
    record_id: str
    dimension: str
    name: str
    role: str
    score: float | None
    sheet: str
    row: int


@dataclass(frozen=True)
class Candidate:
    target: ScoreRecord
    similarity: float
    exact: bool
    reason: str


def normalize_header(value: object) -> str:
    return re.sub(r"[\s/（）()：:·_-]", "", normalize_text(value)).lower()


def suggest_columns(headers: list[str]) -> dict[str, int | None]:
    result: dict[str, int | None] = {}
    normalized = [normalize_header(header) for header in headers]
    for field, aliases in HEADER_ALIASES.items():
        choices = {normalize_header(alias) for alias in aliases}
        result[field] = next((i for i, title in enumerate(normalized) if title in choices), None)
    return result


def _data_workbook(path: Path):
    """Read cell values even when a workbook contains an invalid fill style.

    Some Excel-compatible writers emit empty/unsupported ``<fill>`` entries.
    openpyxl rejects those before it can read any cells. Repair only invalid fill
    definitions in a temporary in-memory ZIP; never rewrite the input file.
    """
    try:
        return load_workbook(path, read_only=True, data_only=True)
    except TypeError as error:
        if "expected <class 'openpyxl.styles.fills.Fill'>" not in str(error):
            raise
        repaired = BytesIO()
        with ZipFile(path) as source, ZipFile(repaired, "w") as destination:
            for member in source.infolist():
                data = source.read(member.filename)
                if member.filename == "xl/styles.xml":
                    data = _repair_invalid_fills(data)
                destination.writestr(member, data)
        repaired.seek(0)
        try:
            return load_workbook(repaired, read_only=True, data_only=True)
        except Exception as repair_error:
            raise ValueError(
                "新评分表的填充样式无法解析。请用 Excel 或 WPS 打开后另存为新的 .xlsx，再导入。"
            ) from repair_error


def _repair_invalid_fills(styles_xml: bytes) -> bytes:
    root = ET.fromstring(styles_xml)
    fills = root.find(f"{{{MAIN_NS}}}fills")
    if fills is None:
        return styles_xml
    changed = False
    for fill in fills:
        try:
            parsed = Fill.from_tree(fill)
        except (TypeError, ValueError, IndexError):
            parsed = None
        if not isinstance(parsed, Fill):
            fill.clear()
            ET.SubElement(fill, f"{{{MAIN_NS}}}patternFill", {"patternType": "none"})
            changed = True
    return ET.tostring(root, encoding="utf-8", xml_declaration=True) if changed else styles_xml


def workbook_sheets(path: Path) -> list[str]:
    workbook = _data_workbook(path)
    try:
        return workbook.sheetnames
    finally:
        workbook.close()


def sheet_headers(path: Path, sheet: str, row: int) -> list[str]:
    workbook = _data_workbook(path)
    try:
        values = next(workbook[sheet].iter_rows(min_row=row, max_row=row, values_only=True))
        return [normalize_text(value) for value in values]
    finally:
        workbook.close()


def load_score_table(
    path: Path, sheet: str, header_row: int, mapping: dict[str, int | None]
) -> list[ScoreRecord]:
    if mapping.get("name") is None:
        raise ValueError("请指定新评分表的活动名称列")
    workbook = _data_workbook(path)
    try:
        worksheet = workbook[sheet]
        records: list[ScoreRecord] = []
        current_dimension = ""
        for row_number, values in enumerate(
            worksheet.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1
        ):
            def field(key: str) -> object:
                index = mapping.get(key)
                return values[index] if index is not None and index < len(values) else None

            name = normalize_activity_name(field("name"))
            if not name:
                continue
            # Scoring tables often merge the category cell across several activity rows.
            dimension_value = normalize_text(field("dimension"))
            if dimension_value:
                current_dimension = dimension_value
            try:
                score = normalize_reference_score(field("score"))
            except ValueError as error:
                raise ValueError(f"{sheet} 第 {row_number} 行：{error}") from error
            records.append(
                ScoreRecord(
                    record_id=f"{sheet}:{row_number}",
                    dimension=current_dimension,
                    name=name,
                    role=normalize_text(field("role")),
                    score=score,
                    sheet=sheet,
                    row=row_number,
                )
            )
        if not records:
            raise ValueError("所选工作表没有可读取的活动条目")
        return records
    finally:
        workbook.close()


def load_phone_export(path: Path) -> list[SourceRecord]:
    workbook = _data_workbook(path)
    try:
        if len(workbook.worksheets) < 2:
            raise ValueError("手机导出文件应包含总表和凭证工作表")
        summary = workbook.worksheets[0]
        header = [normalize_header(value) for value in next(summary.iter_rows(max_row=1, values_only=True))]
        expected = [normalize_header(value) for value in ("领域", "序号", "活动名", "获奖等级", "分数")]
        if header[:5] != expected:
            raise ValueError("所选文件不是本项目的手机导出 Excel（总表表头不匹配）")
        photo_map = _read_photo_map(path, workbook.sheetnames)
        dimension = ""
        records: list[SourceRecord] = []
        for row_number, values in enumerate(summary.iter_rows(min_row=2, values_only=True), start=2):
            if values[0] is not None:
                dimension = normalize_text(values[0])
            if len(values) < 3 or values[1] is None or not normalize_text(values[2]):
                continue
            try:
                number = int(float(values[1]))
            except (TypeError, ValueError) as error:
                raise ValueError(f"总表第 {row_number} 行序号无效") from error
            try:
                score = normalize_reference_score(values[4] if len(values) > 4 else None)
            except ValueError as error:
                raise ValueError(f"总表第 {row_number} 行：{error}") from error
            stable_id = normalize_text(values[5]) if len(values) > 5 else ""
            record_id = stable_id or f"{dimension}:{number}"
            records.append(
                SourceRecord(
                    record_id=record_id,
                    dimension=dimension,
                    number=number,
                    name=normalize_activity_name(values[2]),
                    role=normalize_text(values[3] if len(values) > 3 else None),
                    score=score,
                    sheet=summary.title,
                    row=row_number,
                    photos=tuple(photo_map.get((dimension, number), ())),
                )
            )
        if not records:
            raise ValueError("手机导出文件总表没有活动条目")
        if len({record.record_id for record in records}) != len(records):
            raise ValueError("手机导出文件包含重复条目 ID 或重复的领域序号")
        return records
    finally:
        workbook.close()


def _read_photo_map(path: Path, sheet_names: list[str]) -> dict[tuple[str, int], list[PhotoRef]]:
    result: dict[tuple[str, int], list[PhotoRef]] = {}
    with ZipFile(path) as archive:
        member_by_vm = _rich_image_members(archive)
        names = set(archive.namelist())
        for sheet_index, dimension in enumerate(sheet_names[1:], start=2):
            sheet_path = f"xl/worksheets/sheet{sheet_index}.xml"
            if sheet_path not in names:
                continue
            root = ET.fromstring(archive.read(sheet_path))
            for row in root.findall("./x:sheetData/x:row", NS):
                row_number = int(row.get("r", "0"))
                if row_number < 2:
                    continue
                cells = row.findall("x:c", NS)
                sequence = next((cell for cell in cells if cell.get("r") == f"A{row_number}"), None)
                if sequence is None:
                    continue
                value = sequence.find("x:v", NS)
                if value is None or not value.text:
                    continue
                number = int(float(value.text))
                photos: list[PhotoRef] = []
                for cell in cells:
                    vm = cell.get("vm")
                    if not vm:
                        continue
                    member = member_by_vm.get(int(vm))
                    if member and member in names:
                        photos.append(PhotoRef(dimension, cell.get("r", ""), member))
                result[(dimension, number)] = photos
    return result


def _rich_image_members(archive: ZipFile) -> dict[int, str]:
    """Resolve Excel place-in-cell image metadata to ZIP members."""
    needed = {
        "xl/metadata.xml",
        "xl/richData/rdrichvalue.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
    }
    if not needed.issubset(archive.namelist()):
        return {}
    metadata = ET.fromstring(archive.read("xl/metadata.xml"))
    value_metadata = metadata.find("x:valueMetadata", NS)
    rich = ET.fromstring(archive.read("xl/richData/rdrichvalue.xml"))
    rich_rels = ET.fromstring(archive.read("xl/richData/richValueRel.xml"))
    relationships = ET.fromstring(archive.read("xl/richData/_rels/richValueRel.xml.rels"))
    if value_metadata is None:
        return {}
    relation_targets = {
        node.get("Id"): posixpath.normpath(posixpath.join("xl/richData", node.get("Target", "")))
        for node in relationships.findall("p:Relationship", NS)
    }
    relation_ids = [node.get(f"{{{REL_NS}}}id") for node in rich_rels]
    rich_values = rich.findall("rv:rv", NS)
    result: dict[int, str] = {}
    for vm, bucket in enumerate(value_metadata.findall("x:bk", NS), start=1):
        rc = bucket.find("x:rc", NS)
        if rc is None:
            continue
        rich_index = int(rc.get("v", "-1"))
        if rich_index < 0 or rich_index >= len(rich_values):
            continue
        image_index = rich_values[rich_index].find("rv:v", NS)
        if image_index is None or image_index.text is None:
            continue
        relation_index = int(image_index.text)
        if 0 <= relation_index < len(relation_ids):
            member = relation_targets.get(relation_ids[relation_index])
            if member:
                result[vm] = member
    return result


_YEAR = re.compile(r"(?<!\d)(?:19|20)\d{2}年?")
_EDITION = re.compile(r"第\s*\d+\s*届")
_PUNCTUATION = re.compile(r"[^0-9a-z\u4e00-\u9fff]")


@lru_cache(maxsize=100_000)
def _key(value: str, *, relaxed: bool = False) -> str:
    text = normalize_activity_name(value).lower()
    if relaxed:
        text = _YEAR.sub("", _EDITION.sub("", text))
    return _PUNCTUATION.sub("", text)


def rank_candidates(source: SourceRecord, targets: list[ScoreRecord], limit: int = 5) -> list[Candidate]:
    source_name = _key(source.name)
    source_relaxed = _key(source.name, relaxed=True)
    ranked: list[Candidate] = []
    for target in targets:
        target_name = _key(target.name)
        name_same = bool(source_name) and source_name == target_name
        dimension_same = bool(source.dimension and target.dimension) and _key(source.dimension) == _key(target.dimension)
        role_same = _key(source.role) == _key(target.role)
        # Empty fields on either side cannot establish an exact identity.
        exact = name_same and dimension_same and role_same
        relaxed_name = _key(target.name, relaxed=True)
        name_similarity = SequenceMatcher(None, source_relaxed, relaxed_name).ratio()
        if source_relaxed and relaxed_name and (source_relaxed in relaxed_name or relaxed_name in source_relaxed):
            name_similarity = max(name_similarity, 0.82)
        similarity = 0.76 * name_similarity
        similarity += 0.16 if dimension_same else (0.04 if not target.dimension else 0)
        similarity += 0.08 if role_same else (0.02 if not target.role else 0)
        if exact:
            similarity = 1.0
        if similarity < 0.34:
            continue
        reasons = []
        if not name_same:
            reasons.append("活动名称不同")
        if not dimension_same:
            reasons.append("领域不同或缺失")
        if not role_same:
            reasons.append("职务/等级不同")
        if source.score != target.score:
            reasons.append("分数不同")
        ranked.append(Candidate(target, round(similarity, 3), exact, "、".join(reasons) or "字段一致"))
    ranked.sort(key=lambda item: (-item.similarity, item.target.sheet, item.target.row))
    return ranked[:limit]


def automatic_status(candidates: list[Candidate], source_score: float | None) -> tuple[str, str | None]:
    exact = [candidate for candidate in candidates if candidate.exact]
    if len(exact) > 1:
        return "重复候选", None
    if exact:
        target = exact[0].target
        return ("相同" if source_score == target.score else "分数不同"), target.record_id
    if candidates:
        return "近似待确认", None
    return "未找到", None
