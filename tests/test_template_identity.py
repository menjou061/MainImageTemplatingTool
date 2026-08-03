from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
INTERNAL = ROOT / "L0_Windows命令行版" / "_internal"
sys.path.insert(0, str(INTERNAL))

from template_identity import (
    TemplateIdentityError,
    sha256_file,
    validate_template_identity,
    write_template_identity,
)


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

    def test_absolute_script_invocation_can_import_sibling_profile_module(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            psd = self._approved_template(Path(name))
            environment = os.environ.copy()
            environment.pop("PYTHONPATH", None)
            result = subprocess.run(
                [
                    sys.executable,
                    str(INTERNAL / "template_identity.py"),
                    "--psd",
                    str(psd),
                    "--profile",
                    "legacy-v1",
                    "--variant",
                    "main-800",
                ],
                cwd=name,
                env=environment,
                capture_output=True,
                text=True,
                check=False,
            )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_powershell_utf8_bom_sidecar_is_accepted(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            directory = Path(name)
            psd = self._approved_template(directory)
            sidecar = psd.with_name(psd.name + ".template.json")
            sidecar.write_text(sidecar.read_text(encoding="utf-8"), encoding="utf-8-sig")
            result = validate_template_identity(psd, "legacy-v1", "main-800")
        self.assertEqual(result["template_id"], "paper-jd-self-main-800-v1")

    def test_prepared_copy_can_receive_a_fresh_hash_sidecar(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            directory = Path(name)
            source = self._approved_template(directory)
            prepared = directory / "approved_套版模板.psd"
            prepared.write_bytes(source.read_bytes() + b" prepared copy")

            written = write_template_identity(
                prepared,
                "legacy-v1",
                "main-800",
                source_psd_path=source,
            )
            verified = validate_template_identity(prepared, "legacy-v1", "main-800")
            manifest = json.loads((directory / "approved_套版模板.psd.template.json").read_text(encoding="utf-8"))
            prepared_hash = sha256_file(prepared)

        self.assertTrue(written["manifest"].endswith("approved_套版模板.psd.template.json"))
        self.assertEqual(verified["template_id"], "paper-jd-self-main-800-v1")
        self.assertEqual(manifest["psd_sha256"], prepared_hash)

    def test_generated_identity_cannot_be_written_to_source_psd(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            directory = Path(name)
            source = directory / "source.psd"
            source.write_bytes(b"source bytes")

            with self.assertRaisesRegex(TemplateIdentityError, "E_TEMPLATE_IDENTITY_INVALID"):
                write_template_identity(
                    source,
                    "legacy-v1",
                    "main-800",
                    source_psd_path=source,
                )

            self.assertFalse((directory / "source.psd.template.json").exists())

    def test_generated_identity_keeps_an_unsigned_source_unchanged(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            directory = Path(name)
            source = directory / "source.psd"
            source.write_bytes(b"source bytes")
            prepared = directory / "source_套版模板.psd"
            prepared.write_bytes(b"prepared bytes")
            source_hash = sha256_file(source)

            write_template_identity(
                prepared,
                "legacy-v1",
                "main-800",
                source_psd_path=source,
            )

            self.assertEqual(sha256_file(source), source_hash)
            self.assertFalse((directory / "source.psd.template.json").exists())
            self.assertEqual(
                validate_template_identity(prepared, "legacy-v1", "main-800")["template_id"],
                "paper-jd-self-main-800-v1",
            )

    def test_generated_identity_never_overwrites_an_existing_sidecar(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            directory = Path(name)
            source = directory / "source.psd"
            prepared = directory / "source_套版模板.psd"
            source.write_bytes(b"source bytes")
            prepared.write_bytes(b"prepared bytes")
            existing = directory / "source_套版模板.psd.template.json"
            existing.write_text('{"preserve": true}\n', encoding="utf-8")

            with self.assertRaisesRegex(TemplateIdentityError, "E_TEMPLATE_IDENTITY_INVALID"):
                write_template_identity(
                    prepared,
                    "legacy-v1",
                    "main-800",
                    source_psd_path=source,
                )

            self.assertEqual(existing.read_text(encoding="utf-8"), '{"preserve": true}\n')


if __name__ == "__main__":
    unittest.main()
