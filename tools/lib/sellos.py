"""
tools/lib/sellos.py — Mejora 4 (10/08): SELLO DE INTEGRIDAD del mes cerrado.

Al cerrar un mes se guarda una huella (hash sha256 + conteo + sumas) de los NÚMEROS DUROS
de cada mes en las hojas de movimientos: fecha, descripción, abono, cargo y saldo — lo que
dice el banco y no debe cambiar jamás después del cierre. La capa de CLASIFICACIÓN
(categoría, referencia, observaciones, colores) queda deliberadamente FUERA del sello:
es la que se enriquece en sitio (cola de identificación, cruce factura↔movimiento,
formato manual del usuario) sin romper integridad.

La huella vive en una hoja oculta '🔒 Sellos' DEL PROPIO maestro: viaja con el archivo,
sobrevive copias y regeneraciones (base_master arrastra todas las hojas). Nace del
incidente del 10/08: una regeneración pisó el formato manual de julio — con sellos, una
regeneración o edición que altere NÚMEROS de un mes sellado se delata en la siguiente
auditoría. Re-sellar (tools/sellar_mes.py) es el acto consciente de aceptar un cambio.
"""
from __future__ import annotations

import hashlib

HOJA_SELLOS = "🔒 Sellos"
HOJAS_MOV = {"ALTAMIRA": "📊 ALTAMIRA - Movimientos", "MERAK": "📊 MERAK - Movimientos"}
MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def _num(v) -> float:
    return round(float(v), 2) if isinstance(v, (int, float)) else 0.0


def huella(wb, empresa: str, mes: str) -> dict | None:
    """Huella de los números duros del mes: {n, abonos, cargos, sha}. None si no hay filas."""
    hoja = HOJAS_MOV[empresa]
    if hoja not in wb.sheetnames:
        return None
    filas = []
    for r in wb[hoja].iter_rows(min_row=5, values_only=True):
        if not r or len(r) < 7 or str(r[1] or "") != mes:
            continue
        if str(r[0] or "").startswith("TOTALES"):
            continue
        # r: 0 fecha · 2 descripción · 4 abono · 5 cargo. La clasificación NO entra; el
        # SALDO tampoco: es derivado y en meses viejos vive como FÓRMULA (cambia de cara
        # según el modo de lectura del workbook y daría falsos positivos).
        filas.append(f"{r[0]}|{r[2]}|{_num(r[4]):.2f}|{_num(r[5]):.2f}")
    if not filas:
        return None
    abonos = sum(float(f.split("|")[2]) for f in filas)
    cargos = sum(float(f.split("|")[3]) for f in filas)
    sha = hashlib.sha256("\n".join(filas).encode("utf-8")).hexdigest()[:16]
    return {"n": len(filas), "abonos": round(abonos, 2), "cargos": round(cargos, 2),
            "sha": sha}


def _hoja_sellos(wb):
    if HOJA_SELLOS not in wb.sheetnames:
        ws = wb.create_sheet(HOJA_SELLOS)
        ws.append(["Empresa", "Mes", "Movs", "Abonos S/", "Cargos S/", "Huella sha256",
                   "Sellado"])
        ws.sheet_state = "hidden"
    return wb[HOJA_SELLOS]


def leer_sellos(wb) -> dict:
    """{(empresa, mes): {n, abonos, cargos, sha, cuando}} desde la hoja oculta."""
    if HOJA_SELLOS not in wb.sheetnames:
        return {}
    out = {}
    for r in wb[HOJA_SELLOS].iter_rows(min_row=2, values_only=True):
        if not r or not r[0]:
            continue
        out[(str(r[0]), str(r[1]))] = {"n": int(r[2] or 0), "abonos": _num(r[3]),
                                       "cargos": _num(r[4]), "sha": str(r[5] or ""),
                                       "cuando": str(r[6] or "")}
    return out


def sellar(wb, meses: list | None = None, cuando: str = "") -> list:
    """Escribe/actualiza el sello de cada (empresa, mes) con filas. Devuelve lo sellado."""
    ws = _hoja_sellos(wb)
    existentes = {(str(ws.cell(r, 1).value), str(ws.cell(r, 2).value)): r
                  for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value}
    sellados = []
    for emp in HOJAS_MOV:
        for mes in (meses or MESES_ES):
            h = huella(wb, emp, mes)
            if h is None:
                continue
            fila = existentes.get((emp, mes)) or ws.max_row + 1
            for c, v in enumerate((emp, mes, h["n"], h["abonos"], h["cargos"], h["sha"],
                                   cuando), 1):
                ws.cell(fila, c).value = v
            existentes[(emp, mes)] = fila
            sellados.append((emp, mes, h["sha"]))
    if meses is None:
        # re-sello TOTAL = aceptación consciente: los sellos de meses que ya no tienen
        # filas se retiran (si fue un accidente, la violación ya gritó antes de llegar aquí)
        for (emp, mes), fila in existentes.items():
            if huella(wb, emp, mes) is None:
                for c in range(1, 8):
                    ws.cell(fila, c).value = None
    return sellados


def verificar(wb) -> list[str]:
    """
    Compara la huella ACTUAL de cada mes sellado contra su sello. Devuelve violaciones
    legibles (vacío = integridad intacta). La clasificación puede haber cambiado — aquí
    solo hablan los números del banco.
    """
    violaciones = []
    for (emp, mes), sello in leer_sellos(wb).items():
        h = huella(wb, emp, mes)
        if h is None:
            violaciones.append(f"🔴 {emp} {mes}: el mes sellado YA NO TIENE FILAS "
                               f"(tenía {sello['n']}) — ¿se borró o regeneró mal?")
            continue
        if h["sha"] == sello["sha"]:
            continue
        detalle = []
        if h["n"] != sello["n"]:
            detalle.append(f"movs {sello['n']}→{h['n']}")
        if h["abonos"] != sello["abonos"]:
            detalle.append(f"abonos {sello['abonos']:,.2f}→{h['abonos']:,.2f}")
        if h["cargos"] != sello["cargos"]:
            detalle.append(f"cargos {sello['cargos']:,.2f}→{h['cargos']:,.2f}")
        violaciones.append(
            f"🔴 {emp} {mes}: NÚMEROS ALTERADOS tras el sello ({', '.join(detalle) or 'mismos totales, filas distintas'}) "
            f"— si el cambio es legítimo, re-sella con: python3 tools/sellar_mes.py")
    return violaciones
