"""Preferencias del usuario y registro (log) en %APPDATA% / %LOCALAPPDATA%."""
from __future__ import annotations

import json
import logging
import os
import sys
from dataclasses import asdict, dataclass, fields
from logging.handlers import RotatingFileHandler
from pathlib import Path

from . import APP_NAME


def _base_dir(env: str) -> Path:
    root = os.environ.get(env) or str(Path.home())
    path = Path(root) / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    return path


def config_dir() -> Path:
    return _base_dir("APPDATA")


def log_dir() -> Path:
    path = _base_dir("LOCALAPPDATA") / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def log_file() -> Path:
    return log_dir() / "kardexpress.log"


def resource_path(*parts: str) -> Path:
    """Ruta a recursos empaquetados (funciona en desarrollo y dentro del .exe)."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base.joinpath(*parts)


@dataclass
class Settings:
    last_dir: str = ""
    last_requisitions: str = ""
    last_kardex: str = ""
    output_mode: str = "report"         # "report" | "copy" | "overwrite"
    group_scoped_filldown: bool = True
    open_when_done: bool = False
    appearance: str = "system"          # "light" | "dark" | "system"

    @classmethod
    def load(cls) -> "Settings":
        try:
            data = json.loads((config_dir() / "settings.json").read_text(encoding="utf-8"))
            known = {f.name for f in fields(cls)}
            return cls(**{k: v for k, v in data.items() if k in known})
        except Exception:  # noqa: BLE001 - archivo inexistente o corrupto: valores por defecto
            return cls()

    def save(self) -> None:
        try:
            (config_dir() / "settings.json").write_text(
                json.dumps(asdict(self), indent=2, ensure_ascii=False), encoding="utf-8")
        except OSError:
            logging.getLogger(__name__).warning("No se pudieron guardar las preferencias", exc_info=True)


def setup_logging() -> None:
    handler = RotatingFileHandler(log_file(), maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(handler)
