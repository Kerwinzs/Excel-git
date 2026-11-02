"""Shared errors and write-once output files."""

from contextlib import contextmanager
from pathlib import Path


class ExcelGitError(Exception):
    """An expected error that the CLI can explain without a traceback."""


@contextmanager
def new_output(path: Path, *, binary: bool = False):
    """Exclusive creation also refuses existing files and symlinks."""
    path = Path(path).absolute()
    path.parent.mkdir(parents=True, exist_ok=True)
    options = {} if binary else {"encoding": "utf-8-sig", "newline": ""}
    try:
        stream = path.open("xb" if binary else "x", **options)
    except FileExistsError as exc:
        raise ExcelGitError(f"目标已存在，拒绝覆盖：{path}") from exc
    try:
        with stream:
            yield stream
    except BaseException:
        path.unlink(missing_ok=True)  # Remove only our incomplete new output.
        raise
