import json
import unittest
from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]


class ConfigurationTests(unittest.TestCase):
    def test_cursor_configuration_references_existing_servers(self):
        config = json.loads((REPOSITORY_ROOT / ".cursor" / "mcp.json").read_text())
        servers = config["mcpServers"]
        self.assertEqual(
            set(servers),
            {"office-word", "office-excel", "office-powerpoint", "office-general"},
        )

        for server in servers.values():
            script_name = Path(server["args"][0]).name
            self.assertTrue((REPOSITORY_ROOT / script_name).is_file(), script_name)

    def test_documentation_does_not_reference_missing_entrypoint(self):
        documentation = "\n".join(
            (REPOSITORY_ROOT / name).read_text(encoding="utf-8")
            for name in ("README.md", "README_CN.md", "docs/INSTALL.md")
        )
        self.assertNotIn("/office_server.py", documentation)


if __name__ == "__main__":
    unittest.main()
