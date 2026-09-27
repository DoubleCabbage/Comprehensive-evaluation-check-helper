"""Read-only Excel discovery and row import."""

from pathlib import Path
from typing import Iterator

from openpyxl import load_workbook

from .normalize import normalize_text
from .schema import IGNORED_SOURCE_COLUMNS, SOURCE_COLUMN_MAP


SUPPORTED_EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
REQUIRED_SOURCE_COLUMNS = tuple(SOURCE_COLUMN_MAP)


def discover_excel_files(raw_dir: Path) -> list[Path]:
    """Return candidate Excel files in stable name order."""
    return sorted(
        path
        for path in raw_dir.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXCEL_SUFFIXES
    )


def inspect_workbook(path: Path) -> dict[str, object]:
    """Describe sheets, headers and constant metadata without changing the workbook."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheets: list[dict[str, object]] = []

    for worksheet in workbook.worksheets:
        header_row, headers = _find_header(worksheet)
        header_set = set(headers)
        metadata_values = {
            column: sorted(
                {
                    normalize_text(row[headers.index(column)])
                    for row in worksheet.iter_rows(
                        min_row=header_row + 1,
                        values_only=True,
                    )
                    if headers.index(column) < len(row)
                    and normalize_text(row[headers.index(column)])
                }
            )
            for column in IGNORED_SOURCE_COLUMNS
            if column in header_set
        }
        sheets.append(
            {
                "name": worksheet.title,
                "rows": worksheet.max_row,
                "columns": worksheet.max_column,
                "header_row": header_row,
                "headers": headers,
                "missing_required_columns": [
                    column for column in REQUIRED_SOURCE_COLUMNS if column not in header_set
                ],
                "metadata_unique_values": metadata_values,
            }
        )

    return {"file": path.name, "sheets": sheets}


def iter_source_rows(path: Path) -> Iterator[tuple[str, int, dict[str, object]]]:
    """Yield source rows as dictionaries with their sheet name and Excel row number."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    for worksheet in workbook.worksheets:
        header_row, headers = _find_header(worksheet)
        missing = [column for column in REQUIRED_SOURCE_COLUMNS if column not in headers]
        if missing:
            raise ValueError(f"{worksheet.title!r} 缺少必要列: {', '.join(missing)}")

        for row_number, values in enumerate(
            worksheet.iter_rows(min_row=header_row + 1, values_only=True),
            start=header_row + 1,
        ):
            if not any(value not in (None, "") for value in values):
                continue
            yield worksheet.title, row_number, {
                header: values[index] if index < len(values) else None
                for index, header in enumerate(headers)
                if header
            }


def _find_header(worksheet: object, *, search_rows: int = 20) -> tuple[int, list[str]]:
    """Locate the row containing all required source columns."""
    best_row = 0
    best_headers: list[str] = []
    best_matches = -1

    for row_number, values in enumerate(
        worksheet.iter_rows(min_row=1, max_row=search_rows, values_only=True),
        start=1,
    ):
        headers = [normalize_text(value) for value in values]
        matches = sum(column in headers for column in REQUIRED_SOURCE_COLUMNS)
        if matches > best_matches:
            best_row, best_headers, best_matches = row_number, headers, matches
        if matches == len(REQUIRED_SOURCE_COLUMNS):
            return row_number, headers

    return best_row, best_headers
