import tempfile
import unittest
from pathlib import Path

from tools.telemetry_logger import log_event


class TestTelemetryLogger(unittest.TestCase):
    def test_log_event_creates_file(self):
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "telemetry" / "tool_usage.csv"
            out = log_event(
                telemetry_path=p,
                command="triage",
                status="SUCCESS",
                duration_ms=123,
                output_ref="out.csv",
            )
            self.assertTrue(out.exists())
            content = out.read_text(encoding="utf-8")
            self.assertIn("command", content)
            self.assertIn("triage", content)


if __name__ == "__main__":
    unittest.main()
