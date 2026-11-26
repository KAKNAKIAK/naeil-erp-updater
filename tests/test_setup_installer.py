import importlib.util
import json
from pathlib import Path
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("setup_installer", ROOT / "packaging" / "setup_installer.py")
setup_installer = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(setup_installer)


class SetupInstallerTest(unittest.TestCase):
    def _migrate(self, config):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "config.json"
            path.write_text(json.dumps(config), encoding="utf-8")
            setup_installer.merge_v2_selectors(path)
            return json.loads(path.read_text(encoding="utf-8"))

    def test_install_migrates_previous_default_port_and_preserves_other_config(self):
        config = self._migrate({"erp_debugger_address": "127.0.0.1:9223", "erp_url": "custom"})
        self.assertEqual(config["erp_debugger_address"], "127.0.0.1:9224")
        self.assertEqual(config["erp_url"], "custom")

    def test_install_preserves_custom_erp_port(self):
        config = self._migrate({"erp_debugger_address": "127.0.0.1:9333"})
        self.assertEqual(config["erp_debugger_address"], "127.0.0.1:9333")
