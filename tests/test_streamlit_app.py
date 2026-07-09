import ast
from pathlib import Path
import unittest


class StreamlitAppTests(unittest.TestCase):
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
