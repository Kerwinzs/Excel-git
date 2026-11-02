import csv
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from .fixtures import workbook_bytes


class CliTests(unittest.TestCase):
    def invoke(self, *args):
        return subprocess.run([sys.executable, "-m", "excelgit", *map(str, args)],
                              capture_output=True, text=True)

    def test_cli_end_to_end_with_chinese_and_spaces(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            repo, source = root / "my repo", root / "报表.xlsx"
            v1 = workbook_bytes({"数据": {"A1": 1}})
            source.write_bytes(v1)
            commands = [
                ("init", repo),
                ("snapshot", source, "--name", "报表", "-m", "初版", "--repo", repo),
            ]
            for command in commands:
                result = self.invoke(*command)
                self.assertEqual(result.returncode, 0, result.stderr)
            source.write_bytes(workbook_bytes({"数据": {"A1": 2}}))
            result = self.invoke("snapshot", source, "--name", "报表", "-m", "二版", "--repo", repo)
            self.assertEqual(result.returncode, 0, result.stderr)
            result = self.invoke("history", "报表", "--repo", repo)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn("初版", result.stdout)
            self.assertIn("二版", result.stdout)
            result = self.invoke("export", "报表", "HEAD~1", "-o", root / "old.xlsx", "--repo", repo)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual((root / "old.xlsx").read_bytes(), v1)
            for command in [
                ("diff", "报表", "HEAD~1", "HEAD", "--repo", repo, "-o", root / "history.csv"),
                ("compare", root / "old.xlsx", source, "-o", root / "files.csv"),
            ]:
                result = self.invoke(*command)
                self.assertEqual(result.returncode, 0, result.stderr)
            with (root / "history.csv").open(encoding="utf-8-sig", newline="") as stream:
                self.assertEqual(len(list(csv.reader(stream))), 2)
            current = source.read_bytes()
            result = self.invoke("export", "报表", "HEAD~1", "-o", source, "--repo", repo)
            self.assertEqual(result.returncode, 1)
            self.assertIn("拒绝覆盖", result.stderr)
            self.assertEqual(source.read_bytes(), current)

    def test_wrong_extension_and_missing_file_are_friendly_errors(self):
        for name in ("missing.xlsx", "not-supported.xls"):
            result = self.invoke("compare", name, name, "-o", "unused.csv")
            self.assertEqual(result.returncode, 1)
            self.assertNotIn("Traceback", result.stderr)
