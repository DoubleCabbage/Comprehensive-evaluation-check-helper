from pathlib import Path

from activity_matcher.import_excel import discover_excel_files


def test_discovers_supported_excel_files_only(tmp_path: Path) -> None:
    (tmp_path / "去年评分.xlsx").touch()
    (tmp_path / "启用宏.xlsm").touch()
    (tmp_path / "说明.txt").touch()

    assert [path.name for path in discover_excel_files(tmp_path)] == [
        "去年评分.xlsx",
        "启用宏.xlsm",
    ]
