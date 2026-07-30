from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
INTERNAL = ROOT / "L0_Windows命令行版" / "_internal"
sys.path.insert(0, str(INTERNAL))

from template_identity import TemplateIdentityError, validate_template_identity


class TemplateIdentityTest(unittest.TestCase):
    def _approved_template(self, directory: Path) -> Path:
        psd = directory / "approved.psd"
        psd.write_bytes(b"approved psd bytes")
        manifest = {
            "template_id": "paper-jd-self-main-800-v1",
            "profile_id": "legacy-v1",
            "profile_version": "1.0.0",
            "variant": "main-800",
            "psd_sha256": hashlib.sha256(psd.read_bytes()).hexdigest(),
        }
        (directory / "approved.psd.template.json").write_text(json.dumps(manifest), encoding="utf-8")
        return psd

    def test_matching_sidecar_allows_template_before_photoshop(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            result = validate_template_identity(self._approved_template(Path(name)), "legacy-v1", "main-800")
        self.assertEqual(result["template_id"], "paper-jd-self-main-800-v1")

    def test_missing_or_wrong_sidecar_is_blocked(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            directory = Path(name)
            psd = self._approved_template(directory)
            (directory / "approved.psd.template.json").unlink()
            with self.assertRaisesRegex(TemplateIdentityError, "E_TEMPLATE_IDENTITY_MISSING"):
                validate_template_identity(psd, "legacy-v1", "main-800")

    def test_same_size_or_layer_superset_cannot_bypass_hash_identity(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            psd = self._approved_template(Path(name))
            psd.write_bytes(b"wrong psd with an extra layer")
            with self.assertRaisesRegex(TemplateIdentityError, "E_TEMPLATE_IDENTITY_MISMATCH"):
                validate_template_identity(psd, "legacy-v1", "main-800")


if __name__ == "__main__":
    unittest.main()
