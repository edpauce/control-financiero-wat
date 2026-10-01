"""
tools/lib/env.py — Lector del `.env` del proyecto (formato dotenv), sin dependencias.

Reemplaza los parsers duplicados que tenían tools/scrape/_browser.py y sunat_api.py, que
truncaban el valor en el primer `#`: una contraseña con `#` (ej. `abc#123`) quedaba cortada
y el login fallaba sin decir por qué. Regla dotenv: `#` abre comentario solo al inicio del valor
o si va precedido de un espacio/tab; pegado al valor (`pa#ss`) es parte del valor. Si el valor
completo va entre comillas simples o dobles, se quitan y dentro nunca se corta.

La IA NUNCA ve los valores: este módulo se usa en runtime por los scripts; `.env` está gitignored.

Uso:
    from lib import env
    vals = env.cargar_env()                    # <ROOT>/.env → {"EFACT_ALT_USER": ..., ...}
    vals = env.cargar_env("/ruta/otro.env")
"""
from __future__ import annotations

import os
import re

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
ENV = os.path.join(ROOT, ".env")

# `export CLAVE=valor` también es dotenv válido
_LINEA = re.compile(r"^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$")


def _valor(crudo: str) -> str:
    """Valor de una línea ya partida en `=`: quita comillas envolventes o el comentario en línea."""
    v = crudo.strip()
    if not v or v.startswith("#"):
        return ""
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "'\"":
        return v[1:-1]
    for i in range(1, len(v)):
        if v[i] == "#" and v[i - 1] in " \t":
            return v[:i].rstrip()
    return v


def cargar_env(ruta: str | None = None) -> dict:
    """{CLAVE: valor} del archivo dotenv `ruta` (por defecto <ROOT>/.env). Si no existe → {}."""
    ruta = ruta or ENV
    vals: dict[str, str] = {}
    if not os.path.exists(ruta):
        return vals
    with open(ruta, encoding="utf-8") as fh:
        for lin in fh:
            lin = lin.strip()
            if not lin or lin.startswith("#"):
                continue
            m = _LINEA.match(lin)
            if not m:
                continue
            vals[m.group(1)] = _valor(m.group(2))
    return vals
