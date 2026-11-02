"""Genera dist/config-hosted.json para pegar en MyGeotab (Administración → Add-Ins).

Uso: python build.py [URL_BASE_HOSTING]
"""
import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).parent
DIST = ROOT / "dist"
DIST.mkdir(exist_ok=True)

BASE_URL = (sys.argv[1] if len(sys.argv) > 1
            else "https://apicesetti.github.io/GeotabIDMR").rstrip("/")

config = {
    "name": "Movimientos de equipos",
    "supportEmail": "soporte@cesetti.com.ar",
    "version": "2.1.0",
    "items": [
        {
            "url": BASE_URL + "/movimientos.html",
            "path": "ActivityLink/",
            "menuName": {"es": "Movimientos de equipos", "en": "Device movements"},
            "svgIcon": BASE_URL + "/images/icon.svg",
        }
    ],
}

out = DIST / "config-hosted.json"
out.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print("OK ->", out)
