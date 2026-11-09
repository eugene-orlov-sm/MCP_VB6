"""Startup regression: argparse must accept literal percent signs on Python 3.14."""
import pathlib
import subprocess
import sys
import unittest


class BridgeStartupTests(unittest.TestCase):
    def test_help_and_config_do_not_require_an_ide(self):
        bridge = pathlib.Path(__file__).resolve().parents[1] / "bridge" / "vb6mcp_bridge.py"
        for argument in ("--help", "--print-config"):
            with self.subTest(argument=argument):
                result = subprocess.run(
                    [sys.executable, str(bridge), argument],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    universal_newlines=True, timeout=15,
                )
                self.assertEqual(result.returncode, 0, result.stderr)
                if argument == "--help":
                    self.assertIn("%TEMP%", result.stdout)


if __name__ == "__main__":
    unittest.main()
