#!/usr/bin/env python3
"""
tools/auditar_fuentes.py — Auditoría de conectores: ¿qué fuentes del dashboard/tools están al día?

Revisa cada fuente de tools/lib/sources.py: resuelve el archivo más reciente, detecta el periodo que
cubre y avisa si está DESACTUALIZADO respecto al mes objetivo (por defecto el mes en curso del
proyecto). También indica si el dashboard la usa. Pensado para correr al inicio de cada cierre y
no dejar ningún conector rezagado.

Uso:  python3 tools/auditar_fuentes.py [--mes 2026-06]
"""
from __future__ import annotations
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import sources as S   # noqa: E402

MES_ABR = {"ene": "01", "feb": "02", "mar": "03", "abr": "04", "may": "05", "mayo": "05",
           "jun": "06", "jul": "07", "ago": "08", "set": "09", "sep": "09",
           "oct": "10", "nov": "11", "dic": "12"}

# Mes por nombre (abreviado o completo) SOLO como palabra aparte: sin los anclajes, "PAGOS" daba
# "ago" y cualquier nombre con esas tres letras dentro se leía como un mes.
_MES_RE = re.compile(
    r'(?<![a-z])(ene(?:ro)?|feb(?:rero)?|mar(?:zo)?|abr(?:il)?|mayo?|jun(?:io)?|jul(?:io)?|'
    r'ago(?:sto)?|sep(?:tiembre)?|set(?:iembre)?|oct(?:ubre)?|nov(?:iembre)?|dic(?:iembre)?)(?![a-z])',
    re.I)


def periodo_de(nombre: str, anio: str) -> str:
    m = re.search(r'(20\d{2}-\d{2})', nombre)
    if m:
        return m.group(1)
    m = _MES_RE.search(nombre)
    if m:
        return f"{anio}-{MES_ABR[m.group(1)[:3].lower()]}"
    return "(sin fecha)"


def main():
    from datetime import date
    mes = date.today().strftime("%Y-%m")        # por defecto, el mes en curso
    if "--mes" in sys.argv:
        mes = sys.argv[sys.argv.index("--mes") + 1]
    ruta_dash = os.path.join(os.path.dirname(__file__), "generar_dashboard.py")
    dash = open(ruta_dash).read() if os.path.exists(ruta_dash) else ""

    print(f"AUDITORÍA DE FUENTES · objetivo {mes}\n")
    print(f"{'CLAVE':14s} {'ARCHIVO':44s} {'PERIODO':11s} {'DASH':5s} ESTADO")
    rezagadas = []
    for k in S.PATHS:
        p = S.latest(k)
        if not p:
            print(f"{k:14s} {'❌ NO ENCONTRADO':44s}")
            rezagadas.append((k, "falta"))
            continue
        per = periodo_de(os.path.basename(p), mes[:4])
        usa = "sí" if (f'"{k}"' in dash or f"'{k}'" in dash) else "—"
        if per in ("(sin fecha)",) or per >= mes:
            estado = "✅ al día"
        else:
            estado = f"⚠️ {per} DESACTUALIZADO"
            rezagadas.append((k, per))
        print(f"{k:14s} {os.path.basename(p)[:44]:44s} {per:11s} {usa:5s} {estado}")

    print()
    if rezagadas:
        print(f"⚠️ {len(rezagadas)} conector(es) rezagado(s): " +
              ", ".join(f"{k} ({v})" for k, v in rezagadas))
    else:
        print("✅ Todos los conectores al día.")


if __name__ == "__main__":
    main()
