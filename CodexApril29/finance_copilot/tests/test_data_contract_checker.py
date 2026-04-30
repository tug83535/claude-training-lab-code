import tempfile
import unittest
from pathlib import Path

import pandas as pd

from tools.data_contract_checker import evaluate_dataset


class TestDataContractChecker(unittest.TestCase):
    def test_pass_contract(self):
        df = pd.DataFrame(
            {
                "entity": ["A", "B"],
                "account": ["4000", "5000"],
                "period": ["2026-01", "2026-01"],
                "amount": [10.0, 20.0],
            }
        )
        contract = {
            "required_columns": ["entity", "account", "period", "amount"],
            "column_types": {"amount": "number"},
            "constraints": {"unique_keys": ["entity", "account", "period"]},
            "null_threshold_pct": 10.0,
        }
        result = evaluate_dataset(df, contract)
        self.assertEqual(result.status, "PASS")
        self.assertEqual(result.summary["issues_count"], 0)

    def test_fail_missing_columns(self):
        df = pd.DataFrame({"entity": ["A"]})
        contract = {"required_columns": ["entity", "amount"], "null_threshold_pct": 0}
        result = evaluate_dataset(df, contract)
        self.assertEqual(result.status, "FAIL")
        self.assertTrue(any("Missing required columns" in issue for issue in result.issues))


if __name__ == "__main__":
    unittest.main()
