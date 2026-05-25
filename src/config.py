import json
from pathlib import Path

_ARQUIVO = ".config"


def carregar_api_key(root: Path) -> str:
    try:
        with open(root / _ARQUIVO, encoding="utf-8") as f:
            return json.load(f).get("api_key", "")
    except (FileNotFoundError, json.JSONDecodeError):
        return ""


def salvar_api_key(root: Path, key: str) -> None:
    with open(root / _ARQUIVO, "w", encoding="utf-8") as f:
        json.dump({"api_key": key}, f)
