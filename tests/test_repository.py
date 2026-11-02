import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from excelgit.common import ExcelGitError
from excelgit.compare import compare_workbooks
from excelgit.repository import Repository
from .fixtures import workbook_bytes


class RepositoryTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.addCleanup(self.folder.cleanup)
        self.root = Path(self.folder.name)
        self.repo = Repository.init(self.root / "版本库")
        self.source = self.root / "原始 文件.xlsx"
        self.v1 = workbook_bytes({"数据": {"A1": "初版", "B1": 10}})
        self.v2 = workbook_bytes({"数据": {"A1": "二版", "B1": 20}})
        self.source.write_bytes(self.v1)

    def save_first(self):
        return self.repo.snapshot(self.source, "月度 报表", "第一次保存")

    def test_full_history_export_and_comparison(self):
        first = self.save_first()
        self.assertEqual(self.source.read_bytes(), self.v1)
        self.source.write_bytes(self.v2)
        second = self.repo.snapshot(self.source, "月度 报表", "更新数字")
        self.assertEqual(self.source.read_bytes(), self.v2)
        history = self.repo.history("月度 报表")
        self.assertEqual([v.commit for v in history], [second, first])
        self.assertEqual([v.message for v in history], ["更新数字", "第一次保存"])
        output = self.root / "恢复副本.xlsx"
        self.repo.export("月度 报表", first[:12], output)
        self.assertEqual(output.read_bytes(), self.v1)
        self.assertEqual(self.source.read_bytes(), self.v2)
        self.assertEqual(len(compare_workbooks(self.repo.version_bytes("月度 报表", first),
                                                self.repo.version_bytes("月度 报表", "HEAD"))), 2)
        self.assertEqual(self.repo.git("status", "--porcelain"), b"")

    def test_identical_snapshot_is_a_noop(self):
        self.save_first()
        self.assertIsNone(self.repo.snapshot(self.source, "月度 报表", "重复"))
        self.assertEqual(len(self.repo.history("月度 报表")), 1)

    def test_export_never_overwrites(self):
        first = self.save_first()
        with self.assertRaises(ExcelGitError):
            self.repo.export("月度 报表", first, self.source)
        self.assertEqual(self.source.read_bytes(), self.v1)

    def test_dirty_managed_copy_is_preserved(self):
        self.save_first()
        managed = self.repo.root / "workbooks/月度 报表.xlsx"
        managed.write_bytes(self.v2)
        with self.assertRaises(ExcelGitError):
            self.repo.snapshot(self.source, "月度 报表", "不能覆盖")
        self.assertEqual(managed.read_bytes(), self.v2)

    def test_unrelated_staged_and_untracked_files_are_preserved(self):
        unrelated = self.repo.root / "my-notes.txt"
        unrelated.write_text("不要提交", encoding="utf-8")
        with self.assertRaises(ExcelGitError):
            self.save_first()
        self.repo.git("add", "--", "my-notes.txt")
        with self.assertRaises(ExcelGitError):
            self.save_first()
        self.assertIn(b"my-notes.txt", self.repo.git("diff", "--cached", "--name-only"))

    def test_refuses_nonempty_folders_and_ordinary_git_repositories(self):
        for path in (self.root, self.root / "ordinary"):
            path.mkdir(exist_ok=True)
            if path.name == "ordinary":
                subprocess.run(["git", "init", "--quiet", str(path)], check=True)
            with self.assertRaises(ExcelGitError):
                Repository.init(path)
            with self.assertRaises(ExcelGitError):
                Repository(path)
        self.assertEqual(self.source.read_bytes(), self.v1)

    def test_invalid_names_cannot_escape_the_repository(self):
        for name in ("../outside", "/tmp/x", ".git", "a/b", "a\\b", "", " x", "x "):
            with self.subTest(name=name), self.assertRaises(ExcelGitError):
                self.repo.snapshot(self.source, name, "说明")

    def test_bad_revision_and_missing_workbook_are_read_only(self):
        self.save_first()
        before = self.repo.git("rev-parse", "HEAD")
        for name, revision in [("月度 报表", "--help"), ("不存在", "HEAD")]:
            with self.subTest(name=name), self.assertRaises(ExcelGitError):
                self.repo.export(name, revision, self.root / "new.xlsx")
        self.assertFalse((self.root / "new.xlsx").exists())
        self.assertEqual(self.repo.git("rev-parse", "HEAD"), before)

    def test_commit_failure_restores_managed_bytes_and_index(self):
        first = self.save_first()
        self.source.write_bytes(self.v2)
        real_git = self.repo.git

        def fail_commit(*args):
            if args[0] == "commit":
                raise ExcelGitError("模拟提交失败")
            return real_git(*args)

        with patch.object(self.repo, "git", side_effect=fail_commit):
            with self.assertRaises(ExcelGitError):
                self.repo.snapshot(self.source, "月度 报表", "将失败")
        self.assertEqual((self.repo.root / "workbooks/月度 报表.xlsx").read_bytes(), self.v1)
        self.assertEqual(self.source.read_bytes(), self.v2)
        self.assertEqual(self.repo.git("status", "--porcelain"), b"")
        self.assertEqual(self.repo.git("rev-parse", "HEAD").decode().strip(), first)

    def test_lock_prevents_concurrent_saves(self):
        lock = self.repo.root / ".git/excelgit.lock"
        lock.touch()
        with self.assertRaises(ExcelGitError):
            self.save_first()
        self.assertFalse((self.repo.root / "workbooks/月度 报表.xlsx").exists())

    def test_git_index_lock_failure_still_preserves_managed_copy(self):
        self.save_first()
        self.source.write_bytes(self.v2)
        (self.repo.root / ".git/index.lock").touch()
        with self.assertRaises(ExcelGitError):
            self.repo.snapshot(self.source, "月度 报表", "保存受阻")
        self.assertEqual((self.repo.root / "workbooks/月度 报表.xlsx").read_bytes(), self.v1)
        self.assertEqual(self.source.read_bytes(), self.v2)

    def test_init_is_idempotent_and_identity_is_local(self):
        first = self.save_first()
        reopened = Repository.init(self.repo.root)
        self.assertEqual(reopened.version_bytes("月度 报表", first), self.v1)
        self.assertEqual(reopened.git("config", "--local", "user.email").strip(),
                         b"excelgit@localhost")

    def test_export_refuses_symlink_to_existing_source(self):
        first = self.save_first()
        alias = self.root / "alias.xlsx"
        alias.symlink_to(self.source)
        with self.assertRaises(ExcelGitError):
            self.repo.export("月度 报表", first, alias)
        self.assertEqual(self.source.read_bytes(), self.v1)

    def test_invalid_file_and_empty_message_do_not_modify_repository(self):
        head = self.repo.git("rev-parse", "HEAD")
        for content, message in [(b"invalid xlsx", "说明"), (self.v1, " ")]:
            self.source.write_bytes(content)
            with self.assertRaises(ExcelGitError):
                self.repo.snapshot(self.source, "月度 报表", message)
        self.assertEqual(self.repo.git("rev-parse", "HEAD"), head)
        self.assertEqual(self.repo.git("status", "--porcelain"), b"")
