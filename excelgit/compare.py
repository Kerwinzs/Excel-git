"""Compare cell contents, keeping formulas as formulas rather than cached results."""

import csv
import json
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from io import BytesIO
from pathlib import Path
from zipfile import BadZipFile
from xml.etree.ElementTree import ParseError

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.exceptions import InvalidFileException

from .common import ExcelGitError, new_output

MAX_CELLS = 1_000_000  # A deliberately small-workbook MVP; bound rectangular scans.


@dataclass(frozen=True)
class CellValue:
    kind: str
    value: object


@dataclass(frozen=True)
class Change:
    sheet: str
    cell: str
    change: str
    before: CellValue | None = None
    after: CellValue | None = None


def open_workbook(data: bytes):
    """Read a copy in memory. Never call save() on the user's workbook."""
    try:
        return load_workbook(BytesIO(data), read_only=True, data_only=False)
    except (BadZipFile, InvalidFileException, KeyError, ValueError, OSError, ParseError) as exc:
        raise ExcelGitError(f"无法读取 .xlsx 工作簿：{exc}") from exc


def read_xlsx(path: Path) -> bytes:
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise ExcelGitError("第一版只支持 .xlsx 文件。")
    return path.read_bytes()


def _contents(data: bytes) -> dict[str, dict[str, CellValue]]:
    workbook = open_workbook(data)
    result = {}
    scanned = 0
    try:
        for sheet in workbook.worksheets:
            scanned += (sheet.max_row or 1) * (sheet.max_column or 1)
            if scanned > MAX_CELLS:
                raise ExcelGitError("工作簿扫描范围超过 100 万格；请缩小有效区域后比较。")
            cells = {}
            visited = 0
            for row in sheet.iter_rows():
                visited += len(row)
                if visited > MAX_CELLS:
                    raise ExcelGitError("工作表扫描范围超过 100 万格。")
                for cell in row:
                    if cell.value is None:
                        continue  # Formatting-only cells do not count as content.
                    if cell.data_type == "f" and not isinstance(cell.value, str):
                        raise ExcelGitError("第一版暂不比较数组公式或数据表公式。")
                    cells[cell.coordinate] = CellValue(cell.data_type, cell.value)
            result[sheet.title] = cells
    finally:
        workbook.close()
    return result


def compare_workbooks(before: bytes, after: bytes) -> list[Change]:
    """Match by sheet name and A1 coordinate, not by business row identity."""
    old, new = _contents(before), _contents(after)
    changes = []
    for name in sorted(old.keys() | new.keys()):
        if name not in old:
            changes.append(Change(name, "", "sheet_added"))
        elif name not in new:
            changes.append(Change(name, "", "sheet_removed"))
        left, right = old.get(name, {}), new.get(name, {})
        # Dictionaries preserve worksheet row order; sort explicitly for stable output.
        for address in sorted(left.keys() | right.keys(), key=coordinate_to_tuple):
            a, b = left.get(address), right.get(address)
            if a == b:
                continue
            kind = "cell_added" if a is None else "cell_removed" if b is None else "cell_changed"
            changes.append(Change(name, address, kind, a, b))
    return changes


def _display(cell: CellValue | None) -> str:
    if cell is None:
        return ""
    value = cell.value
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, timedelta):
        return str(value)
    # Quoted strings distinguish text from numbers and keep formulas literal in CSV.
    return json.dumps(value, ensure_ascii=False)


def _safe_text(value: str) -> str:
    # Sheet names can begin with '='. CSV quoting alone cannot prevent evaluation.
    return "'" + value if value.lstrip().startswith(("=", "+", "-", "@")) else value


def write_report(changes: list[Change], output: Path) -> None:
    """UTF-8 BOM makes Chinese text easier to open directly in Excel."""
    with new_output(output) as stream:
        writer = csv.writer(stream)
        writer.writerow(["工作表", "单元格", "变化", "旧类型", "旧值", "新类型", "新值"])
        for item in changes:
            writer.writerow([_safe_text(item.sheet), item.cell, item.change,
                             item.before.kind if item.before else "", _safe_text(_display(item.before)),
                             item.after.kind if item.after else "", _safe_text(_display(item.after))])
