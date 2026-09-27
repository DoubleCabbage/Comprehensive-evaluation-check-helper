"""Independent, batch-oriented review of personal workbooks against a master table.

The submitted XLSX files are read only.  Review copies are edited at the OOXML
level so WPS and Excel in-cell image parts survive the round trip unchanged.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from difflib import SequenceMatcher
from hashlib import sha256
from pathlib import Path
import copy
import json
import os
import posixpath
import re
import shutil
import tempfile
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl.utils.cell import range_boundaries

from .normalize import normalize_reference_score, normalize_text
from .review_core import MAIN_NS, PKG_NS, REL_NS, _data_workbook, _rich_image_members


NS = {"x": MAIN_NS, "p": PKG_NS, "r": REL_NS}
DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
COLORS = {"matched": "E2F0D9", "pending": "FCE4D6", "wrong": "F4CCCC", "proof": "FFF2CC"}
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


def key(value: object) -> str:
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", normalize_text(value).lower())


_CATEGORY_SCORE_SUFFIX = re.compile(
    r"\s*[（(]\s*(?:总分|满分)?\s*\d+(?:\.\d+)?\s*分\s*[）)]\s*$"
)


def clean_category_label(value: object) -> str:
    return _CATEGORY_SCORE_SUFFIX.sub("", normalize_text(value)).strip()


def category_key(value: object) -> str:
    return key(clean_category_label(value)).removesuffix("活动证明").removesuffix("证明")


def clean_name(value: object) -> str:
    # Submission and reference names often include an ordinal prefix.
    return re.sub(r"^\s*\d+\s*[-－—、.]\s*", "", normalize_text(value))


@dataclass(frozen=True)
class FormatConfig:
    reference_header: int = 1
    reference_major: int = 3
    reference_minor: int = 4
    reference_name: int = 5
    reference_role: int = 6
    reference_score: int = 7
    summary_sheet: str = "总表"
    summary_header: int = 2
    summary_major: int = 0
    summary_minor: int | None = None
    summary_number: int = 1
    summary_name: int = 2
    summary_role: int = 3
    summary_score: int = 4
    summary_remark: int = 5
    evidence_number: int = 0
    evidence_image: int = 1
    evidence_sheets: dict[str, str] = field(default_factory=dict)
    alias_groups: dict[str, list[list[str]]] = field(default_factory=dict)
    # Kept for progress files created before grouped, field-specific aliases.
    category_aliases: dict[str, str] = field(default_factory=dict)


ALIAS_FIELDS = {"major": "大类", "minor": "小类", "name": "活动名称", "role": "等级"}


def field_key(value: object, field_name: str) -> str:
    if field_name == "major":
        return category_key(value)
    if field_name == "minor":
        return key(clean_category_label(value))
    if field_name == "name":
        return key(clean_name(value))
    if field_name == "role":
        # + / - distinguish grades such as A+ and A; stripping them creates
        # false duplicate exact matches across adjacent score bands.
        text = normalize_text(value).lower().translate(str.maketrans({"−": "-", "–": "-", "—": "-"}))
        return re.sub(r"[^0-9a-z\u4e00-\u9fff+\-./]", "", text)
    return key(value)


def canonical_field(value: object, field_name: str,
                    groups: dict[str, list[list[str]]] | None = None,
                    legacy_major_aliases: dict[str, str] | None = None) -> str:
    current = field_key(value, field_name)
    if field_name == "major" and legacy_major_aliases:
        current = canonical_category(value, legacy_major_aliases)
    for group in (groups or {}).get(field_name, []):
        normalized = [field_key(name, field_name) for name in group]
        if current in normalized:
            return normalized[0]
    return current


def validate_alias_groups(groups: dict[str, list[list[str]]]) -> None:
    if not isinstance(groups, dict) or any(field not in ALIAS_FIELDS for field in groups):
        raise ValueError("别名规则包含不支持的字段")
    for field_name, field_groups in groups.items():
        if not isinstance(field_groups, list):
            raise ValueError(f"{ALIAS_FIELDS[field_name]}规则必须是分组列表")
        seen = set()
        for group in field_groups:
            if not isinstance(group, list) or len(group) < 2 or not all(isinstance(name, str) for name in group):
                raise ValueError(f"{ALIAS_FIELDS[field_name]}每组至少需要两个名称")
            normalized = [field_key(name, field_name) for name in group]
            if not all(normalized) or len(set(normalized)) != len(normalized):
                raise ValueError(f"{ALIAS_FIELDS[field_name]}有空名称或同组重复名称")
            if seen.intersection(normalized):
                raise ValueError(f"{ALIAS_FIELDS[field_name]}有名称同时属于多个别名组")
            seen.update(normalized)


def _groups_from_legacy(aliases: dict[str, str]) -> dict[str, list[list[str]]]:
    _validate_aliases(aliases)
    groups: list[list[str]] = []
    for alias, canonical in aliases.items():
        matching = [group for group in groups if canonical in group or alias in group]
        if matching:
            group = matching[0]
            for name in (canonical, alias):
                if name not in group:
                    group.append(name)
        else:
            groups.append([canonical, alias])
    result = {"major": groups} if groups else {}
    validate_alias_groups(result)
    return result


def canonical_category(value: object, aliases: dict[str, str] | None = None) -> str:
    """Resolve a major category through persistent, user-defined aliases."""
    current = category_key(value)
    normalized = {category_key(alias): category_key(canonical)
        for alias, canonical in (aliases or {}).items()}
    seen = set()
    while current in normalized:
        if current in seen:
            raise ValueError("大类别名配置存在循环映射")
        seen.add(current)
        current = normalized[current]
    return current


def _validate_aliases(aliases: dict[str, str]) -> None:
    normalized = {}
    for alias, canonical in aliases.items():
        if not isinstance(alias, str) or not isinstance(canonical, str) or not category_key(alias) or not category_key(canonical):
            raise ValueError("大类别名和标准名称都必须是非空文字")
        alias_key, canonical_key = category_key(alias), category_key(canonical)
        if alias_key in normalized and normalized[alias_key] != canonical_key:
            raise ValueError(f"大类别名重复且指向不同标准名称：{alias}")
        normalized[alias_key] = canonical_key
    for alias in aliases:
        canonical_category(alias, aliases)


def load_alias_settings(path: Path) -> dict[str, list[list[str]]]:
    if not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError(f"别名配置格式无效：{path}")
    if payload.get("version") == 1 and isinstance(payload.get("category_aliases"), dict):
        return _groups_from_legacy(payload["category_aliases"])
    if payload.get("version") == 2:
        groups = payload.get("alias_groups")
        validate_alias_groups(groups)
        return groups
    raise ValueError(f"别名配置格式无效：{path}")


def save_alias_settings(path: Path, groups: dict[str, list[list[str]]]) -> None:
    # Validate before replacing the user's existing settings.
    validate_alias_groups(groups)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps({"version": 2, "alias_groups": groups},
        ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


@dataclass(frozen=True)
class Reference:
    sheet: str
    row: int
    major: str
    minor: str
    name: str
    role: str
    score: float | str | None

    @property
    def id(self) -> str:
        return f"{self.sheet}!{self.row}"


@dataclass(frozen=True)
class ImageRef:
    sheet: str
    cell: str
    member: str


@dataclass
class Submission:
    path: Path
    row: int
    major: str
    minor: str
    number: int
    name: str
    role: str
    score: float | None
    photos: list[ImageRef] = field(default_factory=list)
    evidence_sheet: str | None = None
    evidence_row: int | None = None
    image_warning: str = ""
    decision: str | None = None  # reference id, or empty string for confirmed wrong
    proof: str = "pending"  # pending / enough / insufficient

    @property
    def id(self) -> str:
        return f"{self.path.resolve()}|{self.row}"


@dataclass(frozen=True)
class Match:
    target: Reference
    similarity: float
    differences: str
    identity_exact: bool


def _value(values: tuple, index: int) -> object:
    return values[index] if 0 <= index < len(values) else None


def _score(value: object, label: str) -> float | None:
    try:
        return normalize_reference_score(value)
    except ValueError as error:
        raise ValueError(f"{label}：{error}") from error


def _reference_score(value: object) -> float | str | None:
    text = normalize_text(value)
    if not text:
        return None
    try:
        return normalize_reference_score(value)
    except ValueError:
        # The official table may prescribe a review instead of a fixed number.
        return text


def load_references(path: Path, config: FormatConfig) -> list[Reference]:
    workbook = _data_workbook(path)
    try:
        result = []
        for sheet in workbook:
            major = ""
            minor = ""
            for row, values in enumerate(sheet.iter_rows(min_row=config.reference_header + 1, values_only=True), config.reference_header + 1):
                name = clean_name(_value(values, config.reference_name))
                if not name:
                    continue
                major = normalize_text(_value(values, config.reference_major)) or major or sheet.title
                minor = normalize_text(_value(values, config.reference_minor)) or minor
                result.append(Reference(sheet.title, row, major, minor, name,
                    normalize_text(_value(values, config.reference_role)),
                    _reference_score(_value(values, config.reference_score))))
        if not result:
            raise ValueError("参考大表没有找到活动条目；请检查表头行和列设置")
        return result
    finally:
        workbook.close()


def _part_maps(archive: ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target", "") for rel in rels.findall("p:Relationship", NS)}
    result = {}
    for sheet in workbook.findall("./x:sheets/x:sheet", NS):
        target = targets.get(sheet.get(f"{{{REL_NS}}}id"))
        if target:
            result[sheet.get("name", "")] = (target.lstrip("/") if target.startswith("/") else
                posixpath.normpath(posixpath.join("xl", target)))
    return result


def _relationships(archive: ZipFile, part: str) -> dict[str, str]:
    rel_path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if rel_path not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(rel_path))
    result = {}
    for rel in root.findall("p:Relationship", NS):
        target = rel.get("Target", "")
        result[rel.get("Id", "")] = (target.lstrip("/") if target.startswith("/") else
            posixpath.normpath(posixpath.join(posixpath.dirname(part), target)))
    return result


def _wps_images(archive: ZipFile) -> dict[str, str]:
    if "xl/cellimages.xml" not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read("xl/cellimages.xml"))
    rels = _relationships(archive, "xl/cellimages.xml")
    result = {}
    for pic in root.findall(f".//{{{XDR_NS}}}pic"):
        prop = pic.find(f".//{{{XDR_NS}}}cNvPr")
        blip = pic.find(f".//{{{DRAW_NS}}}blip")
        if prop is not None and blip is not None:
            member = rels.get(blip.get(f"{{{REL_NS}}}embed"))
            if member:
                result[prop.get("name", "")] = member
    return result


def _image_cells(archive: ZipFile, part: str, sheet: str, wps: dict[str, str], rich: dict[int, str]) -> tuple[dict[int, list[ImageRef]], bool]:
    root = ET.fromstring(archive.read(part))
    names = set(archive.namelist())
    result: dict[int, list[ImageRef]] = {}
    for cell in root.findall(".//x:sheetData/x:row/x:c", NS):
        address = cell.get("r", "")
        match = re.search(r"\d+$", address)
        if not match:
            continue
        member = rich.get(int(cell.get("vm"))) if cell.get("vm") else None
        formula = cell.find("x:f", NS)
        if formula is not None and formula.text:
            image_id = re.search(r'DISPIMG\s*\(\s*"([^"]+)"', formula.text, re.I)
            if image_id:
                member = wps.get(image_id.group(1))
        if member and member in names:
            result.setdefault(int(match.group()), []).append(ImageRef(sheet, address, member))
    # Floating drawings are deliberately not assigned by proximity to a sequence.
    floating = root.find(f"{{{MAIN_NS}}}drawing") is not None
    return result, floating


def _evidence_sheet(major: str, sheets: list[str], config: FormatConfig) -> str | None:
    explicit = config.evidence_sheets.get(major)
    if explicit in sheets:
        return explicit
    major_key = canonical_field(major, "major", config.alias_groups, config.category_aliases)
    matches = [name for name in sheets if name != config.summary_sheet and
        (major_key in canonical_field(name, "major", config.alias_groups, config.category_aliases)
            or canonical_field(name, "major", config.alias_groups, config.category_aliases) in major_key)]
    return matches[0] if len(matches) == 1 else None


def load_submissions(paths: list[Path], config: FormatConfig) -> list[Submission]:
    result = []
    for path in paths:
        workbook = _data_workbook(path)
        try:
            if config.summary_sheet not in workbook.sheetnames:
                raise ValueError(f"{path.name}：没有 {config.summary_sheet} 工作表")
            summary = workbook[config.summary_sheet]
            rows = []
            major = ""
            for row, values in enumerate(summary.iter_rows(min_row=config.summary_header + 1, values_only=True), config.summary_header + 1):
                major = clean_category_label(_value(values, config.summary_major)) or major
                name = clean_name(_value(values, config.summary_name))
                if not name:
                    continue
                raw_number = _value(values, config.summary_number)
                try:
                    number = int(float(raw_number))
                except (ValueError, TypeError) as error:
                    raise ValueError(f"{path.name} 总表第 {row} 行序号无效") from error
                rows.append(Submission(path, row, major,
                    clean_category_label(_value(values, config.summary_minor)) if config.summary_minor is not None else "",
                    number, name,
                    normalize_text(_value(values, config.summary_role)),
                    _score(_value(values, config.summary_score), f"{path.name} 总表第 {row} 行")))
            if not rows:
                raise ValueError(f"{path.name}：总表没有找到活动条目；请检查表头行和列设置")
            with ZipFile(path) as archive:
                parts = _part_maps(archive)
                wps = _wps_images(archive)
                rich = _rich_image_members(archive)
                sheet_images = {}
                for name, part in parts.items():
                    if name != config.summary_sheet and part in archive.namelist():
                        sheet_images[name] = _image_cells(archive, part, name, wps, rich)
            evidence_rows = {}
            for sheet in workbook:
                if sheet.title == config.summary_sheet:
                    continue
                sequence = {}
                for row, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
                    raw = _value(values, config.evidence_number)
                    try:
                        if raw is not None:
                            sequence[row] = int(float(raw))
                    except (TypeError, ValueError):
                        continue
                evidence_rows[sheet.title] = sequence
            for item in rows:
                evidence = _evidence_sheet(item.major, workbook.sheetnames, config)
                if evidence is None:
                    item.image_warning = "未找到唯一对应的分类证明 Sheet，请设置工作表映射"
                    continue
                item.evidence_sheet = evidence
                images, floating = sheet_images.get(evidence, ({}, False))
                matching_rows = [r for r, n in evidence_rows[evidence].items() if n == item.number]
                if len(matching_rows) != 1:
                    item.image_warning = "证明序号缺失或重复，请人工查看"
                    continue
                item.evidence_row = matching_rows[0]
                item.photos = images.get(item.evidence_row, [])
                if not item.photos:
                    item.image_warning = ("含浮动图片，需打开原表确认" if floating else "对应序号没有可读取的单元格图片")
            result.extend(rows)
        finally:
            workbook.close()
    return result


def rank(item: Submission, references: list[Reference], limit: int = 8,
         aliases: dict[str, str] | None = None,
         alias_groups: dict[str, list[list[str]]] | None = None) -> list[Match]:
    result = []
    for target in references:
        name_same = canonical_field(item.name, "name", alias_groups) == canonical_field(target.name, "name", alias_groups)
        major_same = canonical_field(item.major, "major", alias_groups, aliases) in {
            canonical_field(target.major, "major", alias_groups, aliases),
            canonical_field(target.sheet, "major", alias_groups, aliases)}
        minor_same = not item.minor or canonical_field(item.minor, "minor", alias_groups) == canonical_field(target.minor, "minor", alias_groups)
        role_same = canonical_field(item.role, "role", alias_groups) == canonical_field(target.role, "role", alias_groups)
        similarity = 1.0 if name_same else SequenceMatcher(None, key(item.name), key(target.name)).ratio()
        if item.minor:
            score = .69 * similarity + .15 * major_same + .08 * minor_same + .08 * role_same
        else:
            score = .76 * similarity + .16 * major_same + .08 * role_same
        if name_same and major_same and minor_same and role_same:
            score = 1.0
        if score < .35:
            continue
        differences = []
        if not name_same: differences.append("活动名")
        if not major_same: differences.append("大类")
        if not minor_same: differences.append("小类")
        if not role_same: differences.append("等级")
        if isinstance(target.score, str):
            differences.append(f"参考分数需人工认定：{target.score}")
        elif item.score != target.score:
            differences.append("分数")
        result.append(Match(target, round(score, 3), "、".join(differences) or "一致",
            name_same and major_same and minor_same and role_same))
    result.sort(key=lambda m: (-m.similarity, m.target.sheet, m.target.row))
    return result[:limit]


def status(item: Submission, references: list[Reference],
           aliases: dict[str, str] | None = None,
           alias_groups: dict[str, list[list[str]]] | None = None) -> tuple[str, Reference | None]:
    lookup = {ref.id: ref for ref in references}
    if item.decision is not None:
        if item.decision == "":
            return "wrong", None
        target = lookup.get(item.decision)
        if target:
            return _score_status(item, target), target
    source_name = canonical_field(item.name, "name", alias_groups)
    source_major = canonical_field(item.major, "major", alias_groups, aliases)
    source_minor = canonical_field(item.minor, "minor", alias_groups) if item.minor else ""
    source_role = canonical_field(item.role, "role", alias_groups)
    exact = [ref for ref in references if canonical_field(ref.name, "name", alias_groups) == source_name
        and (canonical_field(ref.major, "major", alias_groups, aliases) == source_major
            or canonical_field(ref.sheet, "major", alias_groups, aliases) == source_major)
        and (not source_minor or canonical_field(ref.minor, "minor", alias_groups) == source_minor)
        and canonical_field(ref.role, "role", alias_groups) == source_role]
    if len(exact) == 1:
        return _score_status(item, exact[0]), exact[0]
    return "pending", None


def _score_status(item: Submission, target: Reference) -> str:
    if isinstance(target.score, str):
        return "review_score"
    if item.score is None or target.score is None:
        return "pending"
    return "matched" if item.score == target.score else "wrong"


def save_progress(path: Path, reference_path: Path, config: FormatConfig,
                  items: list[Submission], refresh_fingerprints: bool = False) -> None:
    from dataclasses import asdict
    source_paths = sorted({str(item.path) for item in items})
    previous = {}
    if path.exists():
        try:
            old = json.loads(path.read_text(encoding="utf-8"))
            if old.get("reference") == str(reference_path) and old.get("sources") == source_paths:
                previous = old.get("fingerprints", {})
        except (OSError, ValueError):
            pass
    files = [reference_path, *(Path(s) for s in source_paths)]
    payload = {"version": 1, "reference": str(reference_path), "config": asdict(config),
        "sources": source_paths,
        "fingerprints": {str(p): (_fingerprint(p) if refresh_fingerprints else
            previous.get(str(p)) or _fingerprint(p)) for p in files},
        "decisions": {item.id: {"decision": item.decision, "proof": item.proof} for item in items}}
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def restore_progress(path: Path, category_aliases: dict[str, str] | None = None,
                     alias_groups: dict[str, list[list[str]]] | None = None) -> tuple[Path, FormatConfig, list[Submission]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("version") != 1:
        raise ValueError("不支持的核对进度版本")
    for filename, expected in payload.get("fingerprints", {}).items():
        if _fingerprint(Path(filename)) != expected:
            raise ValueError(f"文件已改变，请重新扫描：{filename}")
    config = FormatConfig(**payload["config"])
    if category_aliases is not None:
        from dataclasses import replace
        config = replace(config, category_aliases=category_aliases)
    if alias_groups is not None:
        from dataclasses import replace
        config = replace(config, alias_groups=alias_groups, category_aliases={})
    items = load_submissions([Path(p) for p in payload["sources"]], config)
    for item in items:
        saved = payload.get("decisions", {}).get(item.id)
        if saved:
            item.decision, item.proof = saved.get("decision"), saved.get("proof", "pending")
    return Path(payload["reference"]), config, items


def _fingerprint(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_fingerprint(path: Path) -> str:
    return _fingerprint(path)


def backup_workbooks(paths: list[Path], folder: Path) -> list[Path]:
    """Copy every selected workbook without overwriting an existing backup."""
    folder.mkdir(parents=True, exist_ok=True)
    output = []
    for source in dict.fromkeys(path.resolve() for path in paths):
        if not source.is_file():
            raise FileNotFoundError(source)
        destination = folder / source.name
        number = 2
        while destination.exists() or destination.resolve() == source:
            destination = folder / f"{source.stem}_备份_{number}{source.suffix}"
            number += 1
        shutil.copy2(source, destination)
        output.append(destination)
    return output


def text_differences(item: Submission, target: Reference, config: FormatConfig) -> list[tuple[str, str, str]]:
    """Text differences eligible for a persistent equivalence rule."""
    pairs = [
        ("major", item.major, target.major or target.sheet),
        ("name", item.name, target.name),
        ("role", item.role, target.role),
    ]
    if item.minor:
        pairs.insert(1, ("minor", item.minor, target.minor))
    result = []
    for field_name, submitted, official in pairs:
        if not submitted or not official:
            continue
        source_key = canonical_field(submitted, field_name, config.alias_groups, config.category_aliases)
        official_key = canonical_field(official, field_name, config.alias_groups, config.category_aliases)
        if field_name == "major" and source_key == canonical_field(target.sheet, field_name,
            config.alias_groups, config.category_aliases):
            continue
        if source_key != official_key:
            result.append((field_name, submitted, official))
    return result


def add_equivalence_group(groups: dict[str, list[list[str]]], field_name: str,
                          submitted: str, official: str) -> dict[str, list[list[str]]]:
    """Return rules that make the two names equivalent, merging groups if needed."""
    if field_name not in ALIAS_FIELDS:
        raise ValueError(f"不支持的等价字段：{field_name}")
    updated = copy.deepcopy(groups)
    field_groups = updated.setdefault(field_name, [])
    submitted_key, official_key = field_key(submitted, field_name), field_key(official, field_name)
    if not submitted_key or not official_key:
        raise ValueError("不能把空文字加入等价规则")
    if submitted_key == official_key:
        return updated
    touched = [index for index, group in enumerate(field_groups)
        if submitted_key in {field_key(name, field_name) for name in group}
        or official_key in {field_key(name, field_name) for name in group}]
    if not touched:
        field_groups.append([official, submitted])
    else:
        merged = [official]
        for index in touched:
            merged.extend(field_groups[index])
        merged.append(submitted)
        unique = []
        seen = set()
        for name in merged:
            normalized = field_key(name, field_name)
            if normalized not in seen:
                unique.append(name)
                seen.add(normalized)
        for index in reversed(touched):
            del field_groups[index]
        field_groups.append(unique)
    validate_alias_groups(updated)
    return updated


@dataclass(frozen=True)
class ProposedEdit:
    field_name: str
    address: str
    old: object
    new: object
    affected_range: str = ""


def _anchor_for(root: ET.Element, column: int, row: int) -> tuple[str, str]:
    for node in root.findall("./x:mergeCells/x:mergeCell", NS):
        ref = node.get("ref", "")
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        if min_col <= column + 1 <= max_col and min_row <= row <= max_row:
            return f"{_column(min_col - 1)}{min_row}", ref
    return f"{_column(column)}{row}", ""


def propose_original_edits(item: Submission, target: Reference, config: FormatConfig) -> list[ProposedEdit]:
    with ZipFile(item.path) as archive:
        part = _part_maps(archive)[config.summary_sheet]
        root = ET.fromstring(archive.read(part))
    workbook = _data_workbook(item.path)
    try:
        sheet = workbook[config.summary_sheet]
        choices = [
            ("major", config.summary_major, target.major or target.sheet),
            ("name", config.summary_name, target.name),
            ("role", config.summary_role, target.role),
        ]
        if config.summary_minor is not None:
            choices.insert(1, ("minor", config.summary_minor, target.minor))
        if isinstance(target.score, (int, float)):
            choices.append(("score", config.summary_score, target.score))
        result = []
        for field_name, column, official in choices:
            if official is None or official == "":
                continue
            address, affected = _anchor_for(root, column, item.row)
            match = re.match(r"([A-Z]+)(\d+)", address)
            anchor_col = 0
            for char in match.group(1):
                anchor_col = anchor_col * 26 + ord(char) - 64
            old = sheet.cell(int(match.group(2)), anchor_col).value
            new = official
            if field_name in {"major", "minor"}:
                if clean_category_label(old) == clean_category_label(official):
                    continue
                suffix = _CATEGORY_SCORE_SUFFIX.search(str(old) if old is not None else "")
                new = clean_category_label(official) + (suffix.group(0) if suffix else "")
            elif field_name == "name":
                if clean_name(old) == clean_name(official):
                    continue
                prefix = re.match(r"^\s*\d+\s*[-－—、.]\s*", str(old) if old is not None else "")
                new = (prefix.group(0) if prefix else "") + official
            elif field_name == "role":
                if normalize_text(old) == normalize_text(official):
                    continue
            elif field_name == "score" and item.score == float(official):
                continue
            cell = root.find(f".//x:sheetData/x:row/x:c[@r='{address}']", NS)
            if cell is not None and cell.find("x:f", NS) is not None:
                raise ValueError(f"{item.path.name} {config.summary_sheet}!{address} 含公式，不能自动覆盖")
            result.append(ProposedEdit(field_name, address, old, new, affected))
        if len({edit.address for edit in result}) != len(result):
            raise ValueError("格式设置把多个字段指向同一单元格，不能安全应用到原表")
        return result
    finally:
        workbook.close()


def apply_edits_to_original(path: Path, summary_sheet: str, edits: list[ProposedEdit]) -> None:
    if not edits:
        raise ValueError("没有选中要写入原表的差异")
    handle, temporary_name = tempfile.mkstemp(prefix=f".{path.stem}.review-", suffix=path.suffix,
        dir=path.parent)
    os.close(handle)
    temporary = Path(temporary_name)
    try:
        with ZipFile(path) as original, ZipFile(temporary, "w") as saved:
            part = _part_maps(original)[summary_sheet]
            for info in original.infolist():
                data = original.read(info.filename)
                if info.filename == part:
                    data = _replace_cells(data, edits)
                saved.writestr(info, data)
        with ZipFile(temporary) as check:
            if check.testzip() is not None:
                raise ValueError("修改后的 Excel 压缩包校验失败，原文件未改动")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _replace_cells(xml: bytes, edits: list[ProposedEdit]) -> bytes:
    root = ET.fromstring(xml)
    rows = {int(row.get("r", "0")): row for row in root.findall("./x:sheetData/x:row", NS)}
    for edit in edits:
        match = re.fullmatch(r"([A-Z]+)(\d+)", edit.address)
        row_number = int(match.group(2))
        row = rows.get(row_number)
        if row is None:
            raise ValueError(f"总表缺少第 {row_number} 行，原文件未改动")
        cell = next((child for child in row.findall("x:c", NS) if child.get("r") == edit.address), None)
        if cell is None:
            cell = ET.SubElement(row, f"{{{MAIN_NS}}}c", {"r": edit.address})
        for child in list(cell):
            if child.tag in {f"{{{MAIN_NS}}}f", f"{{{MAIN_NS}}}v", f"{{{MAIN_NS}}}is"}:
                cell.remove(child)
        if isinstance(edit.new, (int, float)):
            cell.attrib.pop("t", None)
            ET.SubElement(cell, f"{{{MAIN_NS}}}v").text = str(edit.new)
        else:
            cell.set("t", "inlineStr")
            inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
            ET.SubElement(inline, f"{{{MAIN_NS}}}t").text = str(edit.new)
        row[:] = sorted((c for c in row if c.tag == f"{{{MAIN_NS}}}c"),
            key=lambda c: _cell_sort(c.get("r", ""))) + [
            c for c in row if c.tag != f"{{{MAIN_NS}}}c"]
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _style_ids(styles: ET.Element) -> dict[tuple[int, str], int]:
    fills = styles.find(f"{{{MAIN_NS}}}fills")
    xfs = styles.find(f"{{{MAIN_NS}}}cellXfs")
    if fills is None or xfs is None:
        raise ValueError("原文件缺少标准样式表，无法安全标色")
    new_fills = {}
    for name, rgb in COLORS.items():
        fill = ET.SubElement(fills, f"{{{MAIN_NS}}}fill")
        pattern = ET.SubElement(fill, f"{{{MAIN_NS}}}patternFill", {"patternType": "solid"})
        ET.SubElement(pattern, f"{{{MAIN_NS}}}fgColor", {"rgb": "FF" + rgb})
        ET.SubElement(pattern, f"{{{MAIN_NS}}}bgColor", {"indexed": "64"})
        new_fills[name] = len(fills) - 1
    fills.set("count", str(len(fills)))
    original = list(xfs)
    ids = {}
    for base_id, base in enumerate(original):
        for name, fill_id in new_fills.items():
            clone = copy.deepcopy(base)
            clone.set("fillId", str(fill_id))
            clone.set("applyFill", "1")
            xfs.append(clone)
            ids[(base_id, name)] = len(xfs) - 1
    xfs.set("count", str(len(xfs)))
    return ids


def _column(index: int) -> str:
    value = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        value = chr(65 + remainder) + value
    return value


def _mark_sheet(xml: bytes, items: list[Submission], config: FormatConfig, references: list[Reference], style_ids: dict[tuple[int, str], int], remarks: dict[int, str]) -> bytes:
    root = ET.fromstring(xml)
    by_row = {item.row: item for item in items}
    for row in root.findall("./x:sheetData/x:row", NS):
        item = by_row.get(int(row.get("r", "0")))
        if item is None:
            continue
        state, target = status(item, references, config.category_aliases, config.alias_groups)
        # Major category cells are merged over several rows, so color the row's
        # own sequence through remark cells without changing merged regions.
        first = min(config.summary_number, config.summary_name, config.summary_role, config.summary_score)
        last = max(config.summary_remark, config.summary_score)
        cells = {cell.get("r"): cell for cell in row.findall("x:c", NS)}
        for col in range(first, last + 1):
            address = f"{_column(col)}{item.row}"
            cell = cells.get(address)
            if cell is None:
                cell = ET.Element(f"{{{MAIN_NS}}}c", {"r": address})
                row.append(cell)
                cells[address] = cell
            fill = "proof" if col == config.summary_remark and item.proof == "insufficient" else (
                "pending" if state == "review_score" else state)
            cell.set("s", str(style_ids[(int(cell.get("s", "0")), fill)]))
        notes = []
        if state == "review_score" and target is not None:
            notes.append(f"参考分数需人工认定：{target.score}")
        if item.proof == "insufficient":
            notes.append("证明不足")
        if notes:
            address = f"{_column(config.summary_remark)}{item.row}"
            cell = cells[address]
            old = remarks.get(item.row, "")
            text = old
            for note in notes:
                if note not in text:
                    text += ("；" if text else "") + note
            for child in list(cell):
                cell.remove(child)
            cell.set("t", "inlineStr")
            inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
            ET.SubElement(inline, f"{{{MAIN_NS}}}t").text = text
        cells_sorted = sorted((c for c in row if c.tag == f"{{{MAIN_NS}}}c"),
            key=lambda c: _cell_sort(c.get("r", "")))
        row[:] = cells_sorted + [c for c in row if c.tag != f"{{{MAIN_NS}}}c"]
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _cell_sort(address: str) -> int:
    result = 0
    for char in re.match(r"[A-Z]+", address).group(0):
        result = result * 26 + ord(char) - 64
    return result


def export_review_copies(folder: Path, config: FormatConfig, items: list[Submission], references: list[Reference]) -> list[Path]:
    folder.mkdir(parents=True, exist_ok=True)
    output = []
    for source_path in sorted({item.path for item in items}):
        subset = [item for item in items if item.path == source_path]
        destination = folder / f"{source_path.stem}_审核标记{source_path.suffix}"
        if destination.resolve() == source_path.resolve():
            raise ValueError("审核副本不能覆盖原始提交文件")
        if destination.exists():
            number = 2
            while True:
                candidate = folder / f"{source_path.stem}_审核标记_{number}{source_path.suffix}"
                if not candidate.exists():
                    destination = candidate
                    break
                number += 1
        workbook = _data_workbook(source_path)
        try:
            remarks = {item.row: normalize_text(workbook[config.summary_sheet].cell(item.row, config.summary_remark + 1).value)
                for item in subset}
        finally:
            workbook.close()
        with ZipFile(source_path) as original, ZipFile(destination, "w") as saved:
            parts = _part_maps(original)
            summary_part = parts[config.summary_sheet]
            styles = ET.fromstring(original.read("xl/styles.xml"))
            ids = _style_ids(styles)
            for info in original.infolist():
                data = original.read(info.filename)
                if info.filename == "xl/styles.xml":
                    data = ET.tostring(styles, encoding="utf-8", xml_declaration=True)
                elif info.filename == summary_part:
                    data = _mark_sheet(data, subset, config, references, ids, remarks)
                saved.writestr(info, data)
        output.append(destination)
    return output
