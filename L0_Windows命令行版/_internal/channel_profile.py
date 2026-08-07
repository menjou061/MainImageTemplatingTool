"""Versioned channel profile loading and schema validation for L0."""
from __future__ import annotations
import json
from pathlib import Path
from typing import Any
import re

PROFILE_PATH = Path(__file__).with_name("channel_profiles.json")

class ProfileError(ValueError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(f"{code}: {message}")
        self.code = code


def get_variant_for_sheet(profile: dict[str, Any], sheet_name: str) -> str:
    """Resolve a variant while allowing business-defined record-row Sheet names.

    Profiles with explicitly named variants keep strict Sheet-to-variant
    matching. Profiles without named variants use their default variant; the
    selected Sheet is then validated against the profile's header contract.
    """
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
    for key in ("export_size", "sheet_name", "template_bindings", "output_label", "template_id"):
        if key in selected_variant_config:
            result[key] = selected_variant_config[key]
    # Some channels keep one menu entry while their 750/800 PSDs expose
    # different fields. Apply those variant-specific contracts before
    # preflight and template preparation run.
    for key in (
        "required_psd_variables", "mapping", "sheet", "text_fit", "ignored_headers",
        "record_layout", "fields", "matching", "static_support_art", "static_product_art",
    ):
        if key in selected_variant_config:
            result[key] = selected_variant_config[key]
    return result


def normalize_field_alias(value: Any) -> str:
    """Normalize a human-facing table header or PSD token for matching.

    The visible labels remain Chinese and readable for operators/designers;
    matching only removes binding punctuation and spacing so aliases such as
    ``价格1``, ``@价格1`` and ``价格 1`` resolve to the same field.
    """
    text = str(value or "").strip().replace("\u3000", "")
    while text[:1] in {"@", "!", "#"}:
        text = text[1:]
    text = re.sub(r"[\s_\-:/\\（）()【】\[\]{}<>]", "", text)
    return text.casefold()


def dynamic_field_key(value: Any) -> str:
    """Return a stable CSV key for an unregistered field.

    v1.3 deliberately accepts new labels without a script change. The
    original readable label is retained in data.csv so JSX can bind a layer
    with the same name immediately.
    """
    text = str(value or "").strip().replace("\u3000", " ")
    while text[:1] in {"@", "!", "#"}:
        text = text[1:]
    return re.sub(r"\s+", " ", text).strip()


def field_definitions(profile: dict[str, Any]) -> list[dict[str, Any]]:
    """Return top-level plus variant-specific field definitions."""
    definitions: list[dict[str, Any]] = []
    for item in profile.get("fields", []) or []:
        if isinstance(item, dict):
            definitions.append(item)
    variant = profile.get("variants", {}).get(profile.get("variant"), {})
    for item in variant.get("fields", []) or []:
        if not isinstance(item, dict):
            continue
        field_id = item.get("field_id")
        definitions = [existing for existing in definitions if existing.get("field_id") != field_id]
        definitions.append(item)
    return definitions


def field_alias_map(profile: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Build normalized alias -> field definition, including legacy mapping."""
    result: dict[str, dict[str, Any]] = {}
    for definition in field_definitions(profile):
        output_key = str(definition.get("output_key") or definition.get("label") or definition.get("field_id") or "").strip()
        if not output_key:
            continue
        aliases = list(definition.get("aliases", []) or [])
        aliases.extend([definition.get("field_id"), definition.get("label"), output_key])
        for alias in aliases:
            key = normalize_field_alias(alias)
            if key:
                result.setdefault(key, {**definition, "output_key": output_key})
    # Profiles before v1.3 only carry mapping. Keep those aliases working.
    for source, target in (profile.get("mapping", {}) or {}).items():
        key = normalize_field_alias(source)
        if key and key not in result:
            result[key] = {"field_id": target, "label": target, "output_key": target, "aliases": [source]}
        elif key:
            # Preserve an explicitly declared legacy mapping for old tables;
            # canonical v1.3 labels still resolve through ``fields``.
            result[key] = {"field_id": target, "label": target, "output_key": target, "aliases": [source]}
    return result


def resolve_field_name(value: Any, profile: dict[str, Any]) -> tuple[str, dict[str, Any] | None]:
    """Resolve a table/layer name to its canonical CSV key and definition."""
    raw = str(value or "").strip()
    aliases = field_alias_map(profile)
    definition = aliases.get(normalize_field_alias(raw))
    if definition:
        return str(definition["output_key"]), definition
    if profile.get("matching", {}).get("allow_unknown_fields"):
        return dynamic_field_key(raw), None
    return raw, None

def validate_vertical_schema(headers: list[str], profile: dict[str, Any]) -> None:
    required_headers = list(profile.get("sheet", {}).get("required_headers", []))
    if profile.get("matching", {}).get("mode") == "field_id_alias":
        aliases = field_alias_map(profile)
        normalized_headers = {normalize_field_alias(header) for header in headers if header}
        missing = sorted(
            header
            for header in required_headers
            if normalize_field_alias(header) not in normalized_headers
            and not any(
                normalize_field_alias(alias) in normalized_headers
                and aliases.get(normalize_field_alias(alias), {}).get("output_key")
                == aliases.get(normalize_field_alias(header), {}).get("output_key")
                for alias in aliases
            )
        )
    else:
        missing = sorted(set(required_headers) - set(headers))
    if missing:
        raise ProfileError("E_PROFILE_SCHEMA_MISMATCH", "缺少列：" + "、".join(missing))

def map_vertical_headers(headers: list[str], profile: dict[str, Any]) -> dict[str, str]:
    validate_vertical_schema(headers, profile)
    mapping = profile.get("mapping", {})
    ignored = set(profile.get("ignored_headers", []))
    resolved: dict[str, str] = {}
    unknown = []
    for name in headers:
        if not name or name in ignored:
            continue
        target, definition = resolve_field_name(name, profile)
        if definition or name in mapping or profile.get("matching", {}).get("allow_unknown_fields"):
            resolved[name] = target
        elif name != "变量名称":
            unknown.append(name)
    if unknown:
        raise ProfileError("E_PROFILE_SCHEMA_MISMATCH", "存在未声明列：" + "、".join(unknown))
    return resolved
