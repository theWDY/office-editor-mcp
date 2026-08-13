import os
import unittest
from pathlib import Path
from unittest.mock import patch

import general_server


class WorkspacePathTests(unittest.TestCase):
    def setUp(self):
        self.original_output_dir = general_server.OUTPUT_DIR
        general_server.OUTPUT_DIR = str(Path.cwd() / "test-output")

    def tearDown(self):
        general_server.OUTPUT_DIR = self.original_output_dir

    def test_relative_path_is_resolved_inside_workspace(self):
        resolved = Path(general_server._resolve_workspace_path("reports/test.docx"))
        self.assertEqual(
            resolved,
            (Path.cwd() / "test-output" / "reports" / "test.docx").resolve(),
        )

    def test_parent_traversal_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "OFFICE_EDIT_PATH"):
            general_server._resolve_workspace_path("../outside.txt")

    def test_absolute_path_outside_workspace_is_rejected(self):
        outside = Path.cwd().parent / "outside.txt"
        with self.assertRaisesRegex(ValueError, "OFFICE_EDIT_PATH"):
            general_server._resolve_workspace_path(str(outside))


class DestructiveOperationTests(unittest.TestCase):
    def setUp(self):
        self.original_output_dir = general_server.OUTPUT_DIR
        general_server.OUTPUT_DIR = str(Path.cwd())

    def tearDown(self):
        general_server.OUTPUT_DIR = self.original_output_dir

    @patch.dict(os.environ, {}, clear=False)
    @patch("general_server.os.path.exists", return_value=True)
    @patch("general_server.os.remove")
    def test_delete_is_disabled_by_default(self, remove_mock, _exists_mock):
        os.environ.pop("OFFICE_ALLOW_DESTRUCTIVE", None)
        result = general_server.general_file_operations("delete", "protected.txt")
        self.assertFalse(result["success"])
        self.assertIn("OFFICE_ALLOW_DESTRUCTIVE", result["message"])
        remove_mock.assert_not_called()


if __name__ == "__main__":
    unittest.main()
