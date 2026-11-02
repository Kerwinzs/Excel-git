"""A small argparse interface. Business logic remains usable from Python."""

import argparse
import sys
from pathlib import Path

from .common import ExcelGitError
from .compare import compare_workbooks, read_xlsx, write_report
from .repository import Repository


def parser() -> argparse.ArgumentParser:
    app = argparse.ArgumentParser(description="用本地 Git 管理 Excel 快照，不改写原工作簿。")
    commands = app.add_subparsers(dest="command", required=True)
    init = commands.add_parser("init", help="在新目录或空目录中创建仓库")
    init.add_argument("folder", type=Path)
    init.add_argument("--user-name", default="ExcelGit Local")
    init.add_argument("--user-email", default="excelgit@localhost")
    snapshot = commands.add_parser("snapshot", help="复制原文件并提交命名工作簿的快照")
    snapshot.add_argument("source", type=Path)
    snapshot.add_argument("--name", required=True, help="工作簿名称，不带 .xlsx；后续版本使用同一名称")
    snapshot.add_argument("-m", "--message", required=True)
    history = commands.add_parser("history", help="列出一个工作簿的提交历史")
    history.add_argument("name")
    export = commands.add_parser("export", help="将历史版本导出到一个新文件")
    export.add_argument("name")
    export.add_argument("revision", help="提交哈希、HEAD、HEAD~1 等 Git 版本")
    export.add_argument("-o", "--output", type=Path, required=True)
    diff = commands.add_parser("diff", help="比较仓库内同一工作簿的两个版本")
    diff.add_argument("name")
    diff.add_argument("before")
    diff.add_argument("after")
    diff.add_argument("-o", "--output", type=Path, required=True)
    compare = commands.add_parser("compare", help="直接比较两个 .xlsx 文件，无需仓库")
    compare.add_argument("before", type=Path)
    compare.add_argument("after", type=Path)
    compare.add_argument("-o", "--output", type=Path, required=True)
    for command in (snapshot, history, export, diff):
        command.add_argument("--repo", type=Path, required=True, help="ExcelGit 仓库根目录")
    return app


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    try:
        if args.command == "init":
            repo = Repository.init(args.folder, args.user_name, args.user_email)
            print(f"仓库已就绪：{repo.root}")
        elif args.command == "compare":
            _report(read_xlsx(args.before), read_xlsx(args.after), args.output)
        else:
            repo = Repository(args.repo)
            if args.command == "snapshot":
                commit = repo.snapshot(args.source, args.name, args.message)
                print(f"已保存：{commit}" if commit else "文件字节没有变化，未创建重复快照。")
            elif args.command == "history":
                versions = repo.history(args.name)
                print("提交哈希\t时间\t说明")
                for version in versions:
                    print(f"{version.commit}\t{version.date}\t{version.message}")
                if not versions:
                    print("没有该工作簿的历史记录。")
            elif args.command == "export":
                repo.export(args.name, args.revision, args.output)
                print(f"已导出：{args.output}")
            elif args.command == "diff":
                _report(repo.version_bytes(args.name, args.before),
                        repo.version_bytes(args.name, args.after), args.output)
        return 0
    except (ExcelGitError, OSError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1


def _report(before: bytes, after: bytes, output: Path) -> None:
    if output.suffix.lower() != ".csv":
        raise ExcelGitError("差异报告请使用 .csv 后缀。")
    changes = compare_workbooks(before, after)
    write_report(changes, output)
    print(f"共 {len(changes)} 条差异，报告：{output}")
