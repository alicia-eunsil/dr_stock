from __future__ import annotations

import json
import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[2]
DEFAULT_CONFIG_PATH = BASE_DIR / "config.json"
DEFAULT_SECRETS_PATH = BASE_DIR / "secrets.json"


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def load_config(config_path: str | None = None) -> dict:
    path = Path(config_path) if config_path else DEFAULT_CONFIG_PATH
    config = load_json(path)
    env_base_url = os.getenv("KIS_BASE_URL")
    if env_base_url:
        config.setdefault("kis", {})
        config["kis"]["base_url"] = env_base_url
    return config


def load_secrets(secrets_path: str | None = None) -> dict:
    env_app_key = os.getenv("KIS_APP_KEY")
    env_app_secret = os.getenv("KIS_APP_SECRET")
    if env_app_key and env_app_secret:
        return {
            "app_key": env_app_key,
            "app_secret": env_app_secret,
        }

    path = Path(secrets_path) if secrets_path else DEFAULT_SECRETS_PATH
    return load_json(path)
