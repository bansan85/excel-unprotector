import unittest
import pathlib
import shutil
import subprocess
import sys
import tempfile
import zipfile


class TestPackaging(unittest.TestCase):
    def test_wheel_content(self):
        project = pathlib.Path(__file__).parent.parent.resolve()
        with tempfile.TemporaryDirectory() as tmp:
            # Build from a copy so the build artifacts stay out of the project
            source = pathlib.Path(tmp) / "src"
            source.mkdir()
            for name in ["pyproject.toml", "README.md", "LICENSE", "main.py"]:
                shutil.copy(project / name, source)
            # tests is copied to check that it does not end up in the wheel
            for folder in ["core", "tests"]:
                shutil.copytree(
                    project / folder,
                    source / folder,
                    ignore=shutil.ignore_patterns("__pycache__"),
                )

            subprocess.run(
                [sys.executable, "-m", "pip", "wheel", "--no-deps", "-q"]
                + ["-w", str(pathlib.Path(tmp) / "dist"), str(source)],
                check=True,
            )
            (wheel,) = (pathlib.Path(tmp) / "dist").glob("*.whl")
            with zipfile.ZipFile(wheel) as z:
                names = z.namelist()
                entry_points = next(
                    z.read(name).decode()
                    for name in names
                    if name.endswith("entry_points.txt")
                )

        self.assertIn("excel-unprotector = main:main", entry_points)
        self.assertIn("main.py", names)
        self.assertIn("core/excel_unprotector.py", names)
        self.assertIn("core/vba_grep.py", names)
        self.assertFalse([name for name in names if name.startswith("tests/")])


if __name__ == "__main__":
    unittest.main()
