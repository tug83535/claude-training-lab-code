import tempfile
import unittest
from pathlib import Path

from tools.control_evidence_pack import run


class TestControlEvidencePack(unittest.TestCase):
    def test_creates_zip(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            input_dir = base / "in"
            output_dir = base / "out"
            input_dir.mkdir()
            (input_dir / "a.txt").write_text("hello", encoding="utf-8")

            zip_path = run(input_dir=input_dir, output_dir=output_dir, pack_name="test_pack")
            self.assertTrue(zip_path.exists())
            self.assertEqual(zip_path.suffix, ".zip")

    def test_ignores_existing_zip_in_output(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            io_dir = base / "io"
            io_dir.mkdir()
            (io_dir / "a.txt").write_text("hello", encoding="utf-8")
            (io_dir / "old.zip").write_bytes(b"zip-bytes")

            zip_path = run(input_dir=io_dir, output_dir=io_dir, pack_name="test_pack")
            self.assertTrue(zip_path.exists())


if __name__ == "__main__":
    unittest.main()
