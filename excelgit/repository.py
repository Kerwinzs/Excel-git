"""Git stores original bytes. Only our clean managed copy may be replaced."""

import json
import os
import re
import subprocess
import tempfile
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

from .common import ExcelGitError, new_output
from .compare import open_workbook, read_xlsx

MARKER = ".excelgit.json"
FORMAT = {"format": 1}


@dataclass(frozen=True)
class Version:
    commit: str
    date: str
    message: str


class Repository:
    def __init__(self, path: Path):
        self.root = Path(path).resolve()
        marker = self.root / MARKER
        if not (self.root / ".git").is_dir() or (self.root / ".git").is_symlink():
            raise ExcelGitError("请先用 init 创建独立仓库，并指定该仓库的根目录。")
        try:
            valid = not marker.is_symlink() and json.loads(marker.read_text()) == FORMAT
        except (OSError, ValueError):
            valid = False
        if not valid:
            raise ExcelGitError("这不是 ExcelGit 创建的仓库；不会修改普通 Git 仓库。")
        if Path(self.git("rev-parse", "--show-toplevel").decode().strip()).resolve() != self.root:
            raise ExcelGitError("必须指定独立仓库根目录。")
        if (self.root / "workbooks").is_symlink():
            raise ExcelGitError("workbooks 目录不能是符号链接。")

    def git(self, *args: str) -> bytes:
        # Argument lists avoid shell interpolation; clear inherited Git path overrides.
        env = {key: value for key, value in os.environ.items() if not key.startswith("GIT_")}
        try:
            result = subprocess.run(
                ["git", "--literal-pathspecs", "-C", str(self.root),
                 "-c", "commit.gpgsign=false", *args],
                capture_output=True, env=env, check=False,
            )
        except FileNotFoundError as exc:
            raise ExcelGitError("未找到 Git，请先安装 Git。") from exc
        if result.returncode:
            raise ExcelGitError(result.stderr.decode("utf-8", errors="replace").strip() or "Git 操作失败。")
        return result.stdout

    @classmethod
    def init(cls, path: Path, user_name: str = "ExcelGit Local",
             user_email: str = "excelgit@localhost") -> "Repository":
        root = Path(path).resolve()
        if root.exists() and any(root.iterdir()):
            if (root / MARKER).exists():
                return cls(root)  # Idempotent, without changing its identity or files.
            raise ExcelGitError("初始化需要空目录或新目录；请选择独立文件夹。")
        root.mkdir(parents=True, exist_ok=True)
        # Bootstrap only; all later calls require the marker and exact Git root.
        repo = object.__new__(cls)
        repo.root = root
        repo.git("init", "--quiet", "--template=")
        repo.git("config", "--local", "user.name", user_name)
        repo.git("config", "--local", "user.email", user_email)
        (root / MARKER).write_text(json.dumps(FORMAT) + "\n", encoding="utf-8")
        (root / ".gitattributes").write_text("workbooks/*.xlsx -text\n", encoding="utf-8")
        (root / "workbooks").mkdir()
        repo.git("add", "--", MARKER, ".gitattributes")
        repo.git("commit", "--quiet", "-m", "Initialize ExcelGit repository")
        return cls(root)

    @staticmethod
    def workbook_path(name: str) -> str:
        if (not name or name.startswith(".") or len(name) > 80
                or not re.fullmatch(r"[\w -]+", name, flags=re.UNICODE)
                or name != name.strip()):
            raise ExcelGitError("名称限 1–80 个中英文字母、数字、下划线、空格或短横线，不能以点或空格开头。")
        return f"workbooks/{name}.xlsx"

    @contextmanager
    def _lock(self):
        lock = self.root / ".git" / "excelgit.lock"
        try:
            stream = lock.open("x")
        except FileExistsError as exc:
            raise ExcelGitError("仓库正在保存快照；若上次意外退出，请确认没有进程运行后删除 .git/excelgit.lock。") from exc
        try:
            with stream:
                yield
        finally:
            lock.unlink()

    def snapshot(self, source: Path, name: str, message: str) -> str | None:
        relative = self.workbook_path(name)
        if not message.strip():
            raise ExcelGitError("请提供非空的提交说明。")
        data = read_xlsx(source)
        book = open_workbook(data)
        book.close()  # Validate, but never re-save: Git gets the exact source bytes.
        target = self.root / relative
        with self._lock():
            if self.git("status", "--porcelain", "--untracked-files=all"):
                raise ExcelGitError("仓库中有未保存的修改或未跟踪文件；请先移出/备份它们，快照不会覆盖。")
            tracked = bool(self.git("ls-files", "--", relative))
            if target.is_symlink() or (target.exists() and not tracked):
                raise ExcelGitError("目标不是安全的已管理副本，拒绝覆盖。")
            previous = target.read_bytes() if target.exists() else None
            if data == previous:
                return None
            target.parent.mkdir(exist_ok=True)
            published = False
            try:
                with tempfile.TemporaryDirectory(prefix="excelgit-", dir=self.root / ".git") as folder:
                    temporary = Path(folder) / "snapshot.xlsx"
                    temporary.write_bytes(data)
                    os.replace(temporary, target)
                    published = True
                self.git("add", "--", relative)
                self.git("commit", "--quiet", "-m", message, "--", relative)
            except BaseException as original:
                # Our index was clean on entry; restore the previous managed state.
                if published:
                    reset_error = None
                    try:
                        self.git("reset", "--quiet", "HEAD", "--", relative)
                    except ExcelGitError as exc:
                        reset_error = exc
                    finally:
                        if previous is None:
                            target.unlink(missing_ok=True)
                        else:
                            target.write_bytes(previous)
                    if reset_error:
                        raise ExcelGitError(f"已恢复托管文件，但暂存区回滚失败：{reset_error}") from original
                raise
            return self.git("rev-parse", "HEAD").decode().strip()

    def history(self, name: str) -> list[Version]:
        path = self.workbook_path(name)
        output = self.git("log", "--format=%H%x00%aI%x00%s", "--", path).decode("utf-8")
        return [Version(*line.split("\0", 2)) for line in output.splitlines()]

    def version_bytes(self, name: str, revision: str) -> bytes:
        path = self.workbook_path(name)
        # Resolve to a commit hash before composing the object:path selector.
        commit = self.git("rev-parse", "--verify", "--end-of-options",
                          f"{revision}^{{commit}}").decode().strip()
        return self.git("show", f"{commit}:{path}")

    def export(self, name: str, revision: str, output: Path) -> None:
        if Path(output).suffix.lower() != ".xlsx":
            raise ExcelGitError("历史版本请导出为 .xlsx 文件。")
        data = self.version_bytes(name, revision)
        with new_output(output, binary=True) as stream:
            stream.write(data)
