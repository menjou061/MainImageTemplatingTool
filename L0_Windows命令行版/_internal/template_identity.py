#!/usr/bin/env python3
"""Validate the approved sidecar identity for a Photoshop template.

An approved PSD is accompanied by ``<template>.template.json``.  The sidecar
is deliberately checked before Photoshop starts so a same-size or
layer-superset PSD cannot be mistaken for an approved channel template.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

# The embedded Windows runtime uses ``python311._pth`` and does not add the
# script directory to ``sys.path`` when this file is invoked by absolute path.
SCRIPT_DIR = str(Path(__file__).resolve().parent)
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from channel_profile import ProfileError, get_profile


class TemplateIdentityError(ValueError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(f"{code}: {message}")
        self.code = code


def manifest_path(psd_path: Path) -> Path:
    return psd_path.with_name(psd_path.name + ".template.json")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_template_identity(psd_path: Path, profile_id: str, variant: str) -> dict:
    if not psd_path.is_file():
        raise TemplateIdentityError("E_TEMPLATE_IDENTITY_MISSING", f"PSD 不存在：{psd_path}")
    profile = get_profile(profile_id, variant)
    expected_template_id = str(profile.get("template_id") or "")
    if not expected_template_id:
        raise TemplateIdentityError("E_CONFIG_MISMATCH", f"{profile_id}/{variant} 未配置 template_id")
    sidecar = manifest_path(psd_path)
    if not sidecar.is_file():
        raise TemplateIdentityError("E_TEMPLATE_IDENTITY_MISSING", f"缺少批准模板身份文件：{sidecar.name}")
    try:
        manifest = json.loads(sidecar.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise TemplateIdentityError("E_TEMPLATE_IDENTITY_INVALID", f"无法读取模板身份文件：{error}") from error
    expected = {
        "template_id": expected_template_id,
        "profile_id": profile_id,
        "profile_version": str(profile["profile_version"]),
        "variant": variant,
    }
    mismatches = [key for key, value in expected.items() if str(manifest.get(key, "")) != value]
    if mismatches:
        raise TemplateIdentityError("E_TEMPLATE_IDENTITY_MISMATCH", "模板身份不匹配：" + "、".join(mismatches))
    supplied_hash = str(manifest.get("psd_sha256", "")).lower()
    if len(supplied_hash) != 64 or any(char not in "0123456789abcdef" for char in supplied_hash):
        raise TemplateIdentityError("E_TEMPLATE_IDENTITY_INVALID", "模板身份文件缺少有效 psd_sha256")
    if sha256_file(psd_path) != supplied_hash:
        raise TemplateIdentityError("E_TEMPLATE_IDENTITY_MISMATCH", "PSD 内容与批准模板身份文件不一致")
    return {"template_id": expected_template_id, "manifest": str(sidecar)}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--psd", required=True)
    parser.add_argument("--profile", required=True)
    parser.add_argument("--variant", required=True)
    args = parser.parse_args()
    try:
        result = validate_template_identity(Path(args.psd), args.profile, args.variant)
    except (TemplateIdentityError, ProfileError) as error:
        print(str(error), file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
