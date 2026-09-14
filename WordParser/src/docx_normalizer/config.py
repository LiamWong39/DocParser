from __future__ import annotations

import tomllib
from importlib import resources
from pathlib import Path
from typing import Any


def default_config_text() -> str:
    return resources.files("docx_normalizer").joinpath("default_config.toml").read_text("utf-8")


def _merge(base: dict, override: dict, path: str = "") -> dict:
    out = dict(base)
    for key, value in override.items():
        if key not in base:
            raise ValueError(f"Unknown config key: {path}{key}")
        if isinstance(base[key], dict):
            if not isinstance(value, dict):
                raise ValueError(f"Config key {path}{key} must be a table")
            out[key] = _merge(base[key], value, f"{path}{key}.")
        else:
            out[key] = value
    return out


def load_config(path: str | Path | None = None, overrides: dict[str, Any] | None = None) -> dict:
    cfg = tomllib.loads(default_config_text())
    if path:
        with open(path, "rb") as f:
            cfg = _merge(cfg, tomllib.load(f))
    if overrides:
        cfg = _merge(cfg, overrides)
    return cfg
