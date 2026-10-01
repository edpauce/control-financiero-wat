#!/usr/bin/env python3
"""
tools/integrar_detracciones.py — Integra el MES al SEGUIMIENTO_DETRACCIONES (revisión nueva).

- Hojas 04 MERAK BN / 05 ALTAMIRA BN: anexa los movimientos BN reales del mes con la convención de
  la hoja (título en banda, datos en blanco, ROSADO para auto-depósitos, fila Subtotal).
- Hoja 09 Pendientes: actualiza saldos BN al corte y (si hay) la nota del snapshot.
Parte SIEMPRE del master base SIN el mes (no duplica). No deshace merges. Guarda
SEGUIMIENTO_DETRACCIONES_{REV_FECHA}.xlsx (original intacto).

Uso:  python3 tools/integrar_detracciones.py --mes 2026-07
"""
from __future__ import annotations
import glob
import os
import re
import sys
from datetime import date

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import cierre as CI                 # noqa: E402
from scrape import banco_nacion as BN        # noqa: E402

ROOT = CI.ROOT
HOJA_BN = {"MERAK": "04 MERAK BN", "ALTAMIRA": "05 ALTAMIRA BN"}
ROSADO = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")


def _fmt_fecha(s):
    s = str(s).strip()
    m = re.match(r'(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})', s)         # YYYY.MM.DD
    if m:
        return f"{int(m.group(3)):02d}/{int(m.group(2)):02d}/{m.group(1)}"
    m = re.match(r'(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})', s)         # DD/MM/YYYY
    if m:
        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}"
    return s


def base_seguimiento(cfg):
    """SEGUIMIENTO más reciente SIN la sección del mes (para no duplicar).

    Prefiere una REVISIÓN fechada sobre el master base `_{ANIO}.xlsx`: el paso previo
    (detalle_detracciones) escribe in-place sobre el base y le refresca el mtime, así que por
    mtime el base siempre ganaría — y como nunca lleva bandas BN, cada revisión mensual
    perdería las bandas de los meses anteriores. Ordenar por (es_revisión, mtime) lo evita.
    """
    marca = f"{cfg.MES_NOMBRE} {cfg.ANIO}"
    base_master_name = f"SEGUIMIENTO_DETRACCIONES_{cfg.ANIO}.xlsx"
    revs = sorted([f for f in glob.glob(os.path.join(ROOT, "04. FINANZAS/10. DETRACCIONES/SEGUIMIENTO_DETRACCIONES_*.xlsx"))
                   if not os.path.basename(f).startswith(("~$", ".~lock"))],
                  key=lambda f: (os.path.basename(f) != base_master_name, os.path.getmtime(f)))
    for f in reversed(revs):
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        ws = wb["04 MERAK BN"]
        tiene = any(r and r[0] and str(r[0]).startswith(marca)
                    and any("Movimientos BN" in str(c) for c in r if c)
                    for r in ws.iter_rows(values_only=True))
        wb.close()
        if not tiene:
            return f
    sys.exit(f"❌ No hay SEGUIMIENTO previo SIN la sección {marca} para usar de base. "
             f"Aborto para NO duplicar.")


def sincronizar_hoja09(wb, cfg):
    """
    La hoja '09 Pendientes Auto-Det' se mantiene A MANO sobre el master base
    SEGUIMIENTO_DETRACCIONES_{ANIO}.xlsx — esa es su única fuente de verdad. Pero las
    revisiones fechadas se regeneran desde la REVISIÓN del mes anterior (base_seguimiento
    prefiere revisiones para no perder las bandas BN), así que sin esta sincronización cada
    regeneración resucita la hoja 09 vieja: así se perdió una auto-detracción
    pendiente. Se copia entera (valores, estilos, merges, anchos) del master.
    """
    from copy import copy as _c
    ruta = os.path.join(ROOT, "04. FINANZAS", "10. DETRACCIONES", f"SEGUIMIENTO_DETRACCIONES_{cfg.ANIO}.xlsx")
    if not os.path.exists(ruta) or "09 Pendientes Auto-Det" not in wb.sheetnames:
        return
    src = openpyxl.load_workbook(ruta)["09 Pendientes Auto-Det"]
    dst = wb["09 Pendientes Auto-Det"]
    for m in list(dst.merged_cells.ranges):          # unmerge ANTES de borrar filas
        dst.unmerge_cells(str(m))
    dst.delete_rows(1, dst.max_row)
    for row in src.iter_rows():
        for cell in row:
            d = dst.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                d.font, d.border, d.fill = _c(cell.font), _c(cell.border), _c(cell.fill)
                d.number_format, d.alignment = cell.number_format, _c(cell.alignment)
    for m in src.merged_cells.ranges:
        dst.merge_cells(str(m))
    for letra, dim in src.column_dimensions.items():
        if dim.width:
            dst.column_dimensions[letra].width = dim.width


def main():
    cfg = CI.cargar_config(CI.arg_mes())
    if not cfg.SALDO_BN:
        sys.exit(f"❌ Falta SALDO_BN en tools/config/{cfg.PERIODO}.py (saldos de las constancias BN "
                 f"al corte). Complétalo antes de integrar detracciones.")
    R = CI.rutas(cfg)
    src = base_seguimiento(cfg)
    print("Base:", os.path.basename(src))
    wb = openpyxl.load_workbook(src)

    for emp in ("MERAK", "ALTAMIRA"):
        bn_file = CI.first(R["bn"][emp])
        if not bn_file:
            sys.exit(f"❌ Falta el EECC BN de {emp} para {cfg.PERIODO} ({R['bn'][emp]}).")
        ws = wb[HOJA_BN[emp]]

        def _fill_hex(r):
            f = ws.cell(r, 1).fill
            return str(f.start_color.rgb)[-6:] if f and f.patternType else ""
        fdato = next((i for i in range(5, ws.max_row + 1)
                      if "/" in str(ws.cell(i, 2).value or "") and "20" in str(ws.cell(i, 1).value or "")
                      and _fill_hex(i) != "F4CCCC"), 5)
        fsub = next((i for i in range(1, ws.max_row + 1)
                     if str(ws.cell(i, 1).value or "").startswith("Subtotal")), fdato)
        fhdr = next((i for i in range(1, ws.max_row + 1)
                     if str(ws.cell(i, 1).value or "").startswith("TOTAL Ene")), fsub)
        rep = BN.parse_movimientos_bn(bn_file)
        movs = list(reversed(rep["movimientos"]))   # cronológico ascendente
        ins = ws.max_row + 2                        # APPEND al final (sin insert_rows)
        CI.copiar_estilo(ws, fhdr, ins, ncols=7)
        ws.cell(ins, 1).value = f"{cfg.MES_NOMBRE} {cfg.ANIO} (al {cfg.CORTE})"
        ws.cell(ins, 4).value = f"— Movimientos BN {cfg.MES_NOMBRE.lower()}"
        for k, m in enumerate(movs, 1):
            r = ins + k
            CI.copiar_estilo(ws, fdato, r, ncols=7)
            ws.cell(r, 1).value = f"{cfg.MES_ABREV} {cfg.ANIO}"
            ws.cell(r, 2).value = _fmt_fecha(m["fecha"])
            ws.cell(r, 3).value = (m["documento"] or "") + ("/" + m["oficina"] if m.get("oficina") else "")
            ws.cell(r, 4).value = m["trans"]
            ws.cell(r, 5).value = round(m["cargo"], 2) if m["cargo"] else None
            ws.cell(r, 6).value = round(m["abono"], 2) if m["abono"] else None
            if str(m["documento"]) in cfg.AUTODEP_DOCS:   # rosado = auto-depósito de la empresa
                for c in range(1, 8):
                    ws.cell(r, c).fill = ROSADO
        r = ins + len(movs) + 1
        CI.copiar_estilo(ws, fsub, r, ncols=7)
        ws.cell(r, 1).value = f"Subtotal {cfg.MES_ABREV}{cfg.ANIO}"
        ws.cell(r, 2).value = f"{len(movs)} movimientos"
        ws.cell(r, 5).value = abs(rep["cargos"])
        ws.cell(r, 6).value = rep["abonos"]
        ws.cell(r, 7).value = cfg.SALDO_BN[emp]

    # 09 Pendientes: primero sincronizar desde el master base (fuente de verdad de la hoja),
    # luego saldos al corte + nota (celdas superior-izquierda; NO deshacer merges)
    sincronizar_hoja09(wb, cfg)
    ws = wb["09 Pendientes Auto-Det"]
    if cfg.NOTA_09:
        ws.cell(2, 1).value = cfg.NOTA_09
    for i in range(1, ws.max_row + 1):
        t = str(ws.cell(i, 1).value or "")
        if "Saldo BN MERAK" in t:
            ws.cell(i, 1).value = f"Saldo BN MERAK al {cfg.CORTE}/{cfg.ANIO} (S/)"
            ws.cell(i, 5).value = cfg.SALDO_BN["MERAK"]
        elif "Saldo BN ALTAMIRA" in t:
            ws.cell(i, 1).value = f"Saldo BN ALTAMIRA al {cfg.CORTE}/{cfg.ANIO} (S/)"
            ws.cell(i, 5).value = cfg.SALDO_BN["ALTAMIRA"]

    rev = cfg.REV_FECHA or date.today().isoformat()
    out = os.path.join(ROOT, "04. FINANZAS", "10. DETRACCIONES", f"SEGUIMIENTO_DETRACCIONES_{rev}.xlsx")
    wb.save(out)
    print("✅ Generado:", os.path.basename(out))
    for emp in ("MERAK", "ALTAMIRA"):
        rep = BN.parse_movimientos_bn(CI.first(R["bn"][emp]))
        print(f"  {emp} BN {cfg.MES_NOMBRE.lower()}: {rep['n']} movs · abonos +{rep['abonos']:,.2f} · "
              f"cargos {rep['cargos']:,.2f}")
    return out


if __name__ == "__main__":
    main()
