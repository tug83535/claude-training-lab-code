import unittest

import pandas as pd

from tools.cfo_pulse_report import build_pulse


class TestCFOPulseReport(unittest.TestCase):
    def test_status_assignment(self):
        df = pd.DataFrame(
            {
                "kpi": ["close_sla_pct", "unreconciled_amount"],
                "value": [0.9, 30000],
            }
        )
        thresholds = {
            "kpis": {
                "close_sla_pct": {
                    "direction": "higher_is_better",
                    "green": 0.95,
                    "yellow": 0.85,
                },
                "unreconciled_amount": {
                    "direction": "lower_is_better",
                    "green": 10000,
                    "yellow": 50000,
                },
            }
        }

        out = build_pulse(df, thresholds)
        statuses = {row["kpi"]: row["status"] for _, row in out.iterrows()}
        self.assertEqual(statuses["close_sla_pct"], "YELLOW")
        self.assertEqual(statuses["unreconciled_amount"], "YELLOW")

    def test_non_numeric_value_raises(self):
        df = pd.DataFrame({"kpi": ["close_sla_pct"], "value": ["oops"]})
        thresholds = {"kpis": {}}
        with self.assertRaises(ValueError):
            build_pulse(df, thresholds)


if __name__ == "__main__":
    unittest.main()
