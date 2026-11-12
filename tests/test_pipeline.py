import io
import logging
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

import pandas as pd

from amex_tool.pipeline import (
    DEFAULT_TEMPLATE_COLUMNS,
    clean_statement,
    apply_corporate_cards,
    build_claim_zip,
    load_corporate_mapping,
    plan_claim_file_names,
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


class ClaimExportTests(unittest.TestCase):
    def test_plan_rejects_ambiguous_names_with_a_useful_error(self):
        cases = (
            (("A B", "A_B"), "A_B_AMEX_Claim"),
            (("../Same", "folder/Same"), "Same_AMEX_Claim"),
            (("!!!", "???"), "cardholder_AMEX_Claim"),
            (("Smith", "smith"), "smith_AMEX_Claim"),
        )
        for export_format, suffix in (("csv", ".csv"), ("excel", ".xlsx")):
            for names, stem in cases:
                with self.subTest(format=export_format, names=names):
                    with self.assertRaises(ValueError) as caught:
                        plan_claim_file_names(dict.fromkeys(names), export_format)
                    message = str(caught.exception)
                    self.assertIn("output filename", message)
                    self.assertIn(stem + suffix, message)
                    for name in names:
                        self.assertIn(repr(name), message)

    def test_plan_preserves_names_order_and_csv_fallback(self):
        for export_format, suffix in (("csv", ".csv"), ("excel", ".xlsx"), ("other", ".csv")):
            with self.subTest(format=export_format):
                planned = plan_claim_file_names(
                    dict.fromkeys(("Smith", "Adams")), export_format
                )
                self.assertEqual(list(planned.items()), [
                    ("Smith", "Smith_AMEX_Claim" + suffix),
                    ("Adams", "Adams_AMEX_Claim" + suffix),
                ])

    def test_disk_rejects_late_collisions_before_serializing_or_creating_outputs(self):
        for export_format in ("csv", "excel"):
            for names in (("Normal", "A B", "A_B"), ("Normal", "Smith", "smith")):
                for existing in (False, True):
                    with self.subTest(format=export_format, names=names, existing=existing):
                        frames = {
                            name: pd.DataFrame({"Marker": [i]})
                            for i, name in enumerate(names)
                        }
                        with tempfile.TemporaryDirectory() as temp_dir:
                            output = Path(temp_dir) / "claims"
                            if existing:
                                output.mkdir()
                                (output / "sentinel").write_bytes(b"keep me")
                            with patch.object(pd.DataFrame, "to_csv") as csv, patch.object(
                                pd.DataFrame, "to_excel"
                            ) as excel:
                                with self.assertRaisesRegex(ValueError, "output filename"):
                                    save_claim_frames(frames, output, export_format)
                                csv.assert_not_called()
                                excel.assert_not_called()
                            if existing:
                                self.assertEqual(list(output.iterdir()), [output / "sentinel"])
                                self.assertEqual((output / "sentinel").read_bytes(), b"keep me")
                            else:
                                self.assertFalse(output.exists())

    def test_zip_rejects_late_collisions_before_serialization_or_member_writes(self):
        for export_format in ("csv", "excel"):
            for names in (("Normal", "A B", "A_B"), ("Normal", "Smith", "smith")):
                with self.subTest(format=export_format, names=names):
                    frames = {
                        name: pd.DataFrame({"Marker": [i]})
                        for i, name in enumerate(names)
                    }
                    with patch.object(pd.DataFrame, "to_csv") as csv, patch.object(
                        pd.DataFrame, "to_excel"
                    ) as excel, patch.object(ZipFile, "writestr") as write:
                        with self.assertRaisesRegex(ValueError, "output filename"):
                            build_claim_zip(frames, export_format)
                        csv.assert_not_called()
                        excel.assert_not_called()
                        write.assert_not_called()

    def test_valid_disk_and_zip_exports_preserve_each_groups_contents(self):
        frames = {
            "Smith": pd.DataFrame({"Marker": ["Smith café"], "Amount": [12.5]}, index=[7]),
            "Adams": pd.DataFrame({"Marker": ["Adams marker"], "Amount": [24.75]}, index=[9]),
        }
        for export_format, suffix in (("csv", ".csv"), ("excel", ".xlsx")):
            with self.subTest(format=export_format), tempfile.TemporaryDirectory() as temp_dir:
                expected_names = ["Smith_AMEX_Claim" + suffix, "Adams_AMEX_Claim" + suffix]
                (Path(temp_dir) / expected_names[0]).write_bytes(b"prior run")
                paths = save_claim_frames(frames, temp_dir, export_format)
                self.assertEqual([path.name for path in paths], expected_names)
                with ZipFile(io.BytesIO(build_claim_zip(frames, export_format))) as archive:
                    self.assertEqual(archive.namelist(), expected_names)
                    self.assertEqual(len(set(archive.namelist())), 2)
                    for path, expected in zip(paths, frames.values()):
                        read = pd.read_excel if export_format == "excel" else pd.read_csv
                        pd.testing.assert_frame_equal(read(path), expected.reset_index(drop=True))
                        pd.testing.assert_frame_equal(
                            read(io.BytesIO(archive.read(path.name))), expected.reset_index(drop=True)
                        )

    def test_empty_batches_preserve_disk_directory_creation_and_valid_empty_zip(self):
        self.assertEqual(plan_claim_file_names({}), {})
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "empty"
            self.assertEqual(save_claim_frames({}, output), [])
            self.assertTrue(output.is_dir())
            self.assertEqual(list(output.iterdir()), [])
        with ZipFile(io.BytesIO(build_claim_zip({}))) as archive:
            self.assertEqual(archive.namelist(), [])


if __name__ == "__main__":
    unittest.main()
