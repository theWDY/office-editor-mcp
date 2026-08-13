import unittest

import general_server


class BatchLimitTests(unittest.TestCase):
    def test_document_count_is_bounded_before_file_access(self):
        result = general_server.batch_create_documents("missing.docx", "output", 101)
        self.assertFalse(result["success"])
        self.assertIn("1到100", result["message"])

    def test_batch_file_count_is_bounded(self):
        result = general_server.batch_process_documents(
            ["file.txt"] * 101, "encrypt_document"
        )
        self.assertFalse(result["success"])
        self.assertIn("1到100", result["message"])

    def test_worker_count_is_bounded(self):
        result = general_server.batch_process_documents(
            ["file.txt"], "encrypt_document", max_workers=17
        )
        self.assertFalse(result["success"])
        self.assertIn("1到16", result["message"])


if __name__ == "__main__":
    unittest.main()
