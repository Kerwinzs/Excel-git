"""Run from the project root: python examples/make_samples.py --output demo/input."""

import argparse
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


def create_sample(version: int) -> bytes:
    book = Workbook()
    sheet = book.active
    sheet.title = "采购"
    sheet.append(["品名", "数量", "单价", "金额"])
    sheet.append(["笔记本", 3 if version == 1 else 4, 12.5,
                  "=B2*C2" if version == 1 else "=ROUND(B2*C2,2)"])
    sheet.append(["签字笔", 2, 5, "=B3*C3"])
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="214E70")
    for column in ("A", "B", "C", "D"):
        sheet.column_dimensions[column].width = 20
    for address in ("C2", "D2", "C3", "D3"):
        sheet[address].number_format = "0.00"
    sheet.freeze_panes = "A2"
    notes = book.create_sheet("备注" if version == 1 else "审批")
    notes["A1"] = "首次录入" if version == 1 else "已核对数量"
    stream = BytesIO()
    book.save(stream)
    book.close()
    return stream.getvalue()


def main():
    app = argparse.ArgumentParser(description="创建两份练习用工作簿；已有文件不会被覆盖。")
    app.add_argument("--output", type=Path, default=Path("demo/input"))
    args = app.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    for version in (1, 2):
        target = args.output / f"budget_v{version}.xlsx"
        try:
            with target.open("xb") as stream:
                stream.write(create_sample(version))
        except FileExistsError:
            app.exit(1, f"目标已存在，拒绝覆盖：{target}\n")
        print(f"已创建：{target}")


if __name__ == "__main__":
    main()
