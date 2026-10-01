"""
tools/lib/operaciones.py — Unir las dos patas de un traspaso por el NÚMERO DE OPERACIÓN.

EL HALLAZGO (16/09/2026). Los dos primeros dígitos del «Operación - Número» del BCP son el
CANAL, no el número: el mismo traspaso sale como `01482615` en la cuenta que carga y entra como
`02482615` en la que abona. Los últimos seis dígitos son el identificador, y son iguales en
ambas patas. La constancia de Telecrédito lo escribe con otro prefijo todavía (`00482615`).

POR QUÉ IMPORTA MÁS DE LO QUE PARECE. Un abono inter-compañía llegó a MERAK con la
glosa «DE OTRA CUENTA» —anónima— y el número de operación lo resolvió
sin correo, sin deducción y sin preguntarle a nadie. Los traspasos internos de una quincena completa se unen así. Es evidencia del banco, no una conclusión del maestro.

QUÉ NO HACE, A PROPÓSITO:
1. **Nunca suma monedas distintas.** Un par que saca S/55.00 de la cuenta en soles y mete US$15.97 en
   la de dólares es un CAMBIO DE MONEDA a 3.4440, no un descuadre. Sumar las dos patas reportaría
   una brecha falsa de 39.03 cada vez que haya un traspaso a dólares. Un par de monedas mixtas
   devuelve `neto=None` y se informa con su tasa implícita.
2. **No cierra por sufijo a secas.** Seis dígitos colisionan: dos operaciones sin relación pueden
   compartirlos. Un grupo solo se acepta como par si las patas están en cuentas DISTINTAS y
   dentro de `DIAS_MAXIMOS`. Lo que no pasa el filtro se DEVUELVE como descartado — no se
   tira en silencio, porque un descarte invisible es un hueco que nadie audita.
3. **No clasifica ni escribe.** Devuelve pares; quién los anota y con qué categoría es del
   conciliador, que además pasa por el gate.

Es pura: entra una lista de movimientos, sale la unión. No lee archivos ni conoce rutas.
"""
from __future__ import annotations

import datetime as dt
import re
from collections import defaultdict

# Los batch del banco (ITF, portes, mantenimiento) llevan el número en cero: no son operaciones
# y agruparlos por «000000» juntaría todo el mes en un solo montón.
BATCH = "000000"
DIAS_MAXIMOS = 3          # un traspaso entre cuentas propias liquida el mismo día o al siguiente
TOL = 0.005

INTERCIA, PROPIO = "INTERCIA", "PROPIO"
NEUTRO, DESCUADRE, CAMBIO = "neutro", "descuadre", "cambio de moneda"


def sufijo(op) -> str | None:
    """Los últimos 6 dígitos del número de operación, o None si no identifica nada.

    `01482615` → `482615` · `02482615` → `482615` · `00482615` (constancia) → `482615`.
    """
    digitos = re.sub(r"\D", "", str(op or ""))
    if len(digitos) < 6:
        return None
    ultimos = digitos[-6:]
    return None if ultimos == BATCH else ultimos


def _dias(patas) -> int:
    fechas = [p["fecha"] for p in patas if isinstance(p.get("fecha"), dt.date)]
    return (max(fechas) - min(fechas)).days if len(fechas) > 1 else 0


def _tasa(patas):
    """Tasa implícita de un par de monedas distintas, o None si no es un 1:1 limpio."""
    if len(patas) != 2:
        return None
    pen = next((p for p in patas if p.get("moneda") == "PEN"), None)
    usd = next((p for p in patas if p.get("moneda") == "USD"), None)
    if not pen or not usd or abs(usd["monto"]) < TOL:
        return None
    return round(abs(pen["monto"]) / abs(usd["monto"]), 4)


def emparejar(movimientos) -> tuple[list[dict], list[dict]]:
    """`(pares, descartados)` uniendo por sufijo de operación.

    Cada movimiento: `{cuenta, empresa, moneda, fecha, monto, glosa, op}` — `monto` con signo
    (+ entra, − sale). Cada par: `{suf, tipo, patas, monedas, neto, estado, tasa}`.
    """
    por_suf = defaultdict(list)
    for m in movimientos:
        s = sufijo(m.get("op"))
        if s:
            por_suf[s].append(m)

    pares, descartados = [], []
    for suf, patas in por_suf.items():
        if len(patas) < 2:
            continue
        cuentas = {p.get("cuenta") for p in patas}
        if len(cuentas) < 2:
            # el mismo número en la misma cuenta no es una pata que viaja: es una coincidencia
            descartados.append({"suf": suf, "patas": patas, "motivo": "una sola cuenta"})
            continue
        if _dias(patas) > DIAS_MAXIMOS:
            descartados.append({"suf": suf, "patas": patas,
                                "motivo": f"patas a {_dias(patas)} días — sufijo colisionado"})
            continue

        monedas = {p.get("moneda") for p in patas}
        empresas = {p.get("empresa") for p in patas}
        if len(monedas) > 1:
            neto, estado = None, CAMBIO          # regla 1: jamás sumar monedas distintas
        else:
            neto = round(sum(p["monto"] for p in patas), 2)
            estado = NEUTRO if abs(neto) < TOL else DESCUADRE
        pares.append({"suf": suf, "tipo": INTERCIA if len(empresas) > 1 else PROPIO,
                      "patas": sorted(patas, key=lambda p: -p["monto"]),
                      "monedas": sorted(monedas), "neto": neto, "estado": estado,
                      "tasa": _tasa(patas) if len(monedas) > 1 else None})

    orden = {DESCUADRE: 0, CAMBIO: 1, NEUTRO: 2}
    pares.sort(key=lambda p: (orden[p["estado"]], p["patas"][0]["fecha"]))
    return pares, descartados


def internos(movimientos) -> set:
    """Las llaves `(cuenta, fecha, monto, op)` que un par ya explicó.

    Sirve para lo que motivó todo esto: un abono con glosa anónima que resultó ser una cuenta
    nuestra deja de ser «por identificar» y no vuelve a la pila del cobrador.
    """
    llaves = set()
    for par in emparejar(movimientos)[0]:
        for p in par["patas"]:
            llaves.add((p.get("cuenta"), p.get("fecha"), round(p["monto"], 2), p.get("op")))
    return llaves
