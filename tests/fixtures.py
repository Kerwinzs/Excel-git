"""Tiny in-memory workbooks: tests never need the user's files."""

from io import BytesIO

from openpyxl import Workbook


def workbook_bytes(sheets=None):
    book = Workbook()
    book.remove(book.active)
    for name, cells in (sheets or {"数据": {}}).items():
        sheet = book.create_sheet(name)
        for address, value in cells.items():
            sheet[address] = value
    stream = BytesIO()
    book.save(stream)
    book.close()
    return stream.getvalue()
