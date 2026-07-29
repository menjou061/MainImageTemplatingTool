"""Versioned channel profile loading and schema validation for L0."""
from __future__ import annotations
import json
from pathlib import Path
from typing import Any

PROFILE_PATH = Path(__file__).with_name("channel_profiles.json")

class ProfileError(ValueError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(f"{code}: {message}")
        self.code = code


def get_variant_for_sheet(profile: dict[str, Any], sheet_name: str) -> str:
    """Resolve the single output variant selected by a visible Excel sheet."""
    variants = profile.get("variants", {})
    named = {
        variant_id: config.get("sheet_name")
        for variant_id, config in variants.items()
        if config.get("sheet_name")
    }
    if not named:
        return str(profile.get("default_variant") or "")
    matches = [variant_id for variant_id, configured_sheet in named.items() if configured_sheet == sheet_name]
    if len(matches) != 1:
        raise ProfileError(
            "E_PROFILE_SHEET_MISMATCH",
            f"Sheet {sheet_name!r} 未匹配到唯一的模板规格。请使用渠道配置中的运营 Sheet。",
        )
    return matches[0]

def load_profiles() -> dict[str, dict[str, Any]]:
    with PROFILE_PATH.open(encoding="utf-8") as handle:
        document = json.load(handle)
    return {profile["profile_id"]: profile for profile in document["profiles"]}

def get_profile(profile_id: str | None, variant: str | None = None, *, require_enabled: bool = True) -> dict[str, Any]:
    selected_id = profile_id or "legacy-v1"
    profile = load_profiles().get(selected_id)
    if not profile:
        raise ProfileError("E_PROFILE_UNSUPPORTED", f"不支持的 profile：{selected_id}")
    if require_enabled and profile.get("status") != "enabled":
        raise ProfileError("E_PROFILE_UNSUPPORTED", profile.get("approval_note", "该 profile 尚未批准使用。"))
    selected_variant = variant or profile.get("default_variant")
    if selected_variant not in profile.get("variants", {}):
        raise ProfileError("E_PROFILE_UNSUPPORTED", f"profile {selected_id} 不支持 variant：{selected_variant}")
    result = dict(profile)
    result["variant"] = selected_variant
    selected_variant_config = profile["variants"][selected_variant]
    result["target_size"] = {
        "width": selected_variant_config["width"],
        "height": selected_variant_config["height"],
    }
    for key in ("export_size", "sheet_name", "template_bindings", "output_label"):
        if key in selected_variant_config:
            result[key] = selected_variant_config[key]
    return result

def validate_vertical_schema(headers: list[str], profile: dict[str, Any]) -> None:
    missing = sorted(set(profile.get("sheet", {}).get("required_headers", [])) - set(headers))
    if missing:
        raise ProfileError("E_PROFILE_SCHEMA_MISMATCH", "缺少列：" + "、".join(missing))

def map_vertical_headers(headers: list[str], profile: dict[str, Any]) -> dict[str, str]:
    validate_vertical_schema(headers, profile)
    mapping = profile.get("mapping", {})
    ignored = set(profile.get("ignored_headers", []))
    unknown = [name for name in headers if name and name not in mapping and name not in ignored and name != "变量名称"]
    if unknown:
        raise ProfileError("E_PROFILE_SCHEMA_MISMATCH", "存在未声明列：" + "、".join(unknown))
    return {name: target for name, target in mapping.items() if name in headers}
