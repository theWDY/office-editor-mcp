import subprocess
import sys
import unittest


class StdioCleanlinessTests(unittest.TestCase):
    def test_server_imports_do_not_write_to_stdout(self):
        for module in (
            "word_server",
            "excel_server",
            "powerpoint_server",
            "general_server",
        ):
            with self.subTest(module=module):
                process = subprocess.run(
                    [sys.executable, "-c", f"import {module}"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertEqual(process.returncode, 0, process.stderr)
                self.assertEqual(process.stdout, "")


if __name__ == "__main__":
    unittest.main()
