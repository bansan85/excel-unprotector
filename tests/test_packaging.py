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
            for name in ["pyproject.toml", "README.md", "LICENSE"]:
                shutil.copy(project / name, source)
            # tests is copied to check that it does not end up in the wheel
            for folder in ["excel_unprotector", "tests"]:
                shutil.copytree(
                    project / folder,
                    source / folder,
                    ignore=shutil.ignore_patterns("__pycache__"),
                )

            result = subprocess.run(
                [sys.executable, "-m", "pip", "wheel", "--no-deps", "-v"]
                + ["-w", str(pathlib.Path(tmp) / "dist"), str(source)],
                check=True,
                capture_output=True,
                text=True,
                errors="replace",
            )
            self.assertNotIn(
                "SetuptoolsDeprecationWarning", result.stdout + result.stderr
            )
            (wheel,) = (pathlib.Path(tmp) / "dist").glob("*.whl")
            with zipfile.ZipFile(wheel) as z:
                names = z.namelist()
                entry_points = next(
                    z.read(name).decode()
                    for name in names
                    if name.endswith("entry_points.txt")
                )
                metadata = next(
                    z.read(name).decode()
                    for name in names
                    if name.endswith(".dist-info/METADATA")
                ).splitlines()

        self.assertIn("excel-unprotector = excel_unprotector.cli:main", entry_points)
        self.assertIn("excel_unprotector/__init__.py", names)
        self.assertIn("excel_unprotector/__main__.py", names)
        self.assertIn("excel_unprotector/cli.py", names)
        self.assertIn("excel_unprotector/unprotector.py", names)
        self.assertIn("excel_unprotector/vba_grep.py", names)
        self.assertEqual(
            {name.split("/")[0] for name in names if ".dist-info/" not in name},
            {"excel_unprotector"},
        )
        self.assertIn("License-Expression: MIT", metadata)
        self.assertIn("Requires-Python: >=3.9", metadata)
        self.assertFalse(
            any(line.startswith("Classifier: License ::") for line in metadata)
        )
        self.assertTrue(
            any(name.endswith(".dist-info/licenses/LICENSE") for name in names)
        )


if __name__ == "__main__":
    unittest.main()
