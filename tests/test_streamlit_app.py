import ast
import io
from pathlib import Path
import runpy
import sys
import unittest
from unittest.mock import MagicMock, patch
from zipfile import ZipFile

import pandas as pd


class StreamlitAppTests(unittest.TestCase):
    def run_submitted_app(self, names, export_format):
        st = MagicMock()
        st.file_uploader.side_effect = [MagicMock(name="synthetic.csv"), None, None]
        st.selectbox.return_value = export_format
        st.form_submit_button.return_value = True
        frames = {name: pd.DataFrame({"Marker": [name], "Amount": [1]}) for name in names}
        with patch.dict(sys.modules, {"streamlit": st}), patch(
            "amex_tool.pipeline.load_statement", return_value=pd.DataFrame()
        ), patch("amex_tool.pipeline.clean_statement", return_value=pd.DataFrame()), patch(
            "amex_tool.pipeline.generate_claim_frames", return_value=frames
        ):
            runpy.run_path("streamlit_app.py")
        return st

    def test_collision_shows_error_without_announcing_success_or_download(self):
        for export_format in ("csv", "excel"):
            with self.subTest(format=export_format):
                st = self.run_submitted_app(("Normal", "Smith", "smith"), export_format)
                st.error.assert_called_once()
                self.assertIn("Processing failed:", st.error.call_args.args[0])
                self.assertIn("output filename", st.error.call_args.args[0])
                st.success.assert_not_called()
                st.download_button.assert_not_called()

    def test_valid_batch_downloads_unique_zip_with_group_markers(self):
        for export_format, suffix in (("csv", ".csv"), ("excel", ".xlsx")):
            with self.subTest(format=export_format):
                st = self.run_submitted_app(("Smith", "Adams"), export_format)
                st.error.assert_not_called()
                st.success.assert_called_once()
                st.download_button.assert_called_once()
                kwargs = st.download_button.call_args.kwargs
                self.assertEqual(kwargs["file_name"], "amex_claims.zip")
                self.assertEqual(kwargs["mime"], "application/zip")
                with ZipFile(io.BytesIO(kwargs["data"])) as archive:
                    self.assertEqual(archive.namelist(), [
                        "Smith_AMEX_Claim" + suffix, "Adams_AMEX_Claim" + suffix
                    ])
                    for name in ("Smith", "Adams"):
                        read = pd.read_excel if export_format == "excel" else pd.read_csv
                        actual = read(io.BytesIO(archive.read(name + "_AMEX_Claim" + suffix)))
                        self.assertEqual(actual.to_dict("list"), {"Marker": [name], "Amount": [1]})

    def test_deployment_version_marker_is_declared(self):
        source = Path("streamlit_app.py").read_text()
        tree = ast.parse(source)
        assignments = {
            target.id: node.value.value
            for node in tree.body
            if isinstance(node, ast.Assign)
            for target in node.targets
            if isinstance(target, ast.Name)
            and isinstance(node.value, ast.Constant)
            and isinstance(node.value.value, str)
        }

        self.assertEqual(
            assignments.get("APP_VERSION"),
            "cff90c0 - corporate card import fix",
        )
        self.assertIn("APP_VERSION", source)


if __name__ == "__main__":
    unittest.main()
