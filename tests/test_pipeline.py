import io
import logging
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from amex_tool.pipeline import (
    DEFAULT_TEMPLATE_COLUMNS,
    clean_statement,
    apply_corporate_cards,
    load_corporate_mapping,
    save_claim_frames,
)

logging.getLogger("amex_tool.pipeline").setLevel(logging.ERROR)


def _frame():
    return pd.DataFrame([{column: "" for column in DEFAULT_TEMPLATE_COLUMNS}])


class CorporateCardPipelineTests(unittest.TestCase):
    def test_corporate_mapping_uses_named_columns_when_order_changes(self):
        data = (
            "Name,Active,Corporate Card ID,Card Number\n"
            "Jane Doe,Yes,CARD-001,9999\n"
        ).encode("utf-8")

        mapping = load_corporate_mapping(io.BytesIO(data), "cards.csv")
        frames = {"Doe": _frame()}

        apply_corporate_cards(frames, mapping)

        self.assertEqual(frames["Doe"]["Corporate Card"].iloc[0], "CARD-001 - Jane Doe")

    def test_unmatched_corporate_card_is_left_blank(self):
        data = "Name,Corporate Card ID\nJane Doe,CARD-001\n".encode("utf-8")

        mapping = load_corporate_mapping(io.BytesIO(data), "cards.csv")
        frames = {"Smith": _frame()}

        apply_corporate_cards(frames, mapping)

        self.assertEqual(frames["Smith"]["Corporate Card"].iloc[0], "")

    def test_missing_corporate_mapping_leaves_cards_blank(self):
        frames = {"Doe": _frame()}

        apply_corporate_cards(frames, None)

        self.assertEqual(frames["Doe"]["Corporate Card"].iloc[0], "")

    def test_corporate_csv_delimiter_is_detected(self):
        data = "Name\tCorporate Card ID\nJane Doe\tCARD-001\n".encode("utf-8")

        mapping = load_corporate_mapping(io.BytesIO(data), "cards.csv")
        frames = {"Doe": _frame()}

        apply_corporate_cards(frames, mapping)

        self.assertEqual(frames["Doe"]["Corporate Card"].iloc[0], "CARD-001 - Jane Doe")

    def test_clean_statement_preserves_reference_numbers(self):
        statement = pd.DataFrame(
            {
                "Transaction Amount USD": ["12.34"],
                "Transaction Description 4": ["INV-98765"],
            }
        )

        cleaned = clean_statement(statement)

        self.assertEqual(cleaned["Transaction Description 4"].iloc[0], "INV-98765")

    def test_save_claim_frames_sanitizes_filenames(self):
        frames = {"../outside": _frame()}

        with tempfile.TemporaryDirectory() as temp_dir:
            exported = save_claim_frames(frames, temp_dir, "csv")
            output_root = Path(temp_dir).resolve()

            self.assertEqual(len(exported), 1)
            self.assertEqual(exported[0].parent.resolve(), output_root)
            self.assertTrue(exported[0].name.endswith("_AMEX_Claim.csv"))
            self.assertNotIn("..", exported[0].name)
            self.assertNotIn("/", exported[0].name)


if __name__ == "__main__":
    unittest.main()
