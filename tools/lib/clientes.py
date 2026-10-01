#!/usr/bin/env python3
"""
tools/lib/clientes.py — Qué clientes son del Estado y qué plazo de pago les corresponde.

Por qué existe: la **Ley de Contrataciones del Estado** obliga a la entidad a pagar dentro de los
**10 días**. Si el aging usa el vencimiento comercial anotado en la factura —30 días o fin de
mes—, facturas al Estado ya exigibles figuran como "por vencer": una factura emitida el 30/07
aparecería venciendo el 31/08 cuando su plazo legal se cumplió el 09/08.

Ojo con el matiz legal: el plazo corre desde la **conformidad** del servicio, no desde la emisión
de la factura. Aquí se calcula desde la emisión porque es el dato que el maestro tiene; para una
factura recién emitida puede adelantar unos días el vencimiento real. Para reclamar formalmente,
la fecha que vale es la de la conformidad.

Los patrones son PREFIJOS cortos a propósito: el maestro guarda la razón social truncada
(«…MUNICIPAL DE AGUA VALLEGRA»), así que un patrón con el nombre completo no encuentra nada.

Identificación por NOMBRE, no por RUC: el maestro de cobros guarda la razón social, no el RUC.

Las entidades de este archivo son FICTICIAS: ilustran la forma del padrón.
"""
from __future__ import annotations

import re

PLAZO_ESTADO_DIAS = 10

# Clientes del Estado en la cartera. El patrón es lo que se busca en la razón social del
# maestro (mayúsculas, sin tildes); el valor es (RUC si se conoce, descripción).
ESTATALES = {
    "MUNICIPALIDAD DISTRITAL DE":            ("", "gobierno local"),
    "GOBIERNO REGIONAL DE":                  ("", "gobierno regional"),
    "UNIVERSIDAD NACIONAL":                  ("", "universidad pública"),
    "MUNICIPAL DE AGUA VALLEGRA":            ("", "empresa municipal (nombre truncado en el maestro)"),
}


def _norm(s: str) -> str:
    s = (s or "").upper()
    for a, b in (("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U"), ("Ñ", "N")):
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s)


def es_estatal(cliente: str) -> bool:
    """¿La factura va a una entidad del Estado? (por razón social)"""
    n = _norm(cliente)
    return any(pat in n for pat in ESTATALES)


def detalle_estatal(cliente: str) -> tuple[str, str] | None:
    n = _norm(cliente)
    for pat, info in ESTATALES.items():
        if pat in n:
            return info
    return None


def plazo_dias(cliente: str) -> int | None:
    """Días de plazo legal si es del Estado; None si se rige por el vencimiento comercial."""
    return PLAZO_ESTADO_DIAS if es_estatal(cliente) else None
