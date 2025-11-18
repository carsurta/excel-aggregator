from __future__ import annotations

import unittest
from pathlib import Path
import sys

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from header_multirow import load_sheet_merge_aware, detect_header_band_and_build


class HeaderDetectionTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        base = Path(__file__).resolve().parents[1]
        cls.data_dir = base / "data"
        cls.labels = pd.read_csv(cls.data_dir / "header_labels.csv")

    def test_labeled_headers_detected_exactly(self):
        mismatches = []
        for row in self.labels.itertuples(index=False):
            file_path = self.data_dir / getattr(row, "file_name")
            sheet = getattr(row, "sheet_name")
            expected_start = int(getattr(row, "header_start_1based")) - 1
            expected_end = int(getattr(row, "header_end_1based")) - 1
            df = load_sheet_merge_aware(str(file_path), sheet)
            start, end, _ = detect_header_band_and_build(df)
            if (start, end) != (expected_start, expected_end):
                mismatches.append(
                    (
                        file_path.name,
                        sheet,
                        (expected_start + 1, expected_end + 1),
                        (start + 1, end + 1),
                    )
                )
        self.assertFalse(
            mismatches,
            f"헤더 검출 실패 케이스: {mismatches}",
        )


if __name__ == "__main__":
    unittest.main()
