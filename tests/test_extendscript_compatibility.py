"""Keep Photoshop ExtendScript files compatible with its ES3-era parser."""
from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = (
    ROOT / "L0_Windows命令行版" / "_internal" / "batch_template.jsx",
    ROOT / "L0_Windows命令行版" / "_internal" / "template_prepare.jsx",
)
UNSUPPORTED_TOKENS = ("(?:", "(?<", "=>", "const ", "let ")


class ExtendScriptCompatibilityTest(unittest.TestCase):
    def test_jsx_files_do_not_use_modern_parser_features(self) -> None:
        for script in SCRIPTS:
            source = script.read_text(encoding="utf-8")
            for token in UNSUPPORTED_TOKENS:
                self.assertNotIn(token, source, f"{script.name} must remain ExtendScript-compatible: {token}")


if __name__ == "__main__":
    unittest.main()
