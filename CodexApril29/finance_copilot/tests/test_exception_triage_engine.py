import unittest

import pandas as pd

from tools.exception_triage_engine import compute_scores


class TestExceptionTriageEngine(unittest.TestCase):
    def test_rank_order(self):
        df = pd.DataFrame(
            {
                "id": [1, 2, 3],
                "impact": [0.9, 0.5, 0.1],
                "confidence": [0.8, 0.5, 0.2],
                "days_open": [1, 5, 20],
            }
        )
        weights = {
            "impact_weight": 0.5,
            "confidence_weight": 0.3,
            "recency_weight": 0.2,
            "days_scale": 30,
            "required_columns": ["id", "impact", "confidence", "days_open"],
        }
        out = compute_scores(df, weights)
        self.assertEqual(int(out.iloc[0]["id"]), 1)
        self.assertEqual(int(out.iloc[0]["priority_rank"]), 1)


if __name__ == "__main__":
    unittest.main()
