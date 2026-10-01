#!/usr/bin/env python3
"""
tools/cruzar_tributos.py — Cruza el REPORTE DE PAGOS de SUNAT ("Relación de constancias de
pago", consulta SOL) contra el banco. Es el enlace tributario↔financiero que faltaba:

  SUNAT dice QUÉ se declaró y pagó por período (0601 PLAME · 0621 IGV-Renta · 1662 boletas ·
  depósitos de detracción) → el banco debe mostrar la SALIDA de ese dinero:
    · formularios pagados en "Banco de la Nación" → cargo 78787878 en la cuenta de detracciones
      (los tributos se pagan desde el BN)
    · detracciones pagadas en "Banco de Crédito" → salida `PAGO DETRAC` en la conciliación BCP
      (son auto-detracciones nuestras o detracción como comprador)

Entrada: el xlsx exportado de SUNAT (hoja con 'Período | Formulario | … | Monto pago').
Se detecta la empresa por el RUC impreso en la cabecera. Salida: matriz por período en
04. FINANZAS/05. INFORMES/CRUCE_TRIBUTOS_{EMP}.md con ✓ (evidencia en banco) / ⚠️ (sin evidencia).

Uso:  python3 tools/cruzar_tributos.py "05. TRIBUTARIO/MERAK/2026/06.JUNIO/REPORTE-TRIBUTOS_*.xlsx"
      python3 tools/cruzar_tributos.py            # busca todos los REPORTE-TRIBUTOS_* del árbol
"""
from __future__ import annotations

import csv
import glob
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import sources as S            # noqa: E402
from lib import empresas as E           # noqa: E402

ROOT = S.ROOT
BN_DIR = {"ALTAMIRA": "01. ALTAMIRA/03. DETRACCIONES/2026/01. EECC_BN",
          "MERAK": "02. MERAK/03. DETRACCIONES/2026/01. EECC_BN"}
HOJA_CONC = {"ALTAMIRA": "📊 ALTAMIRA - Movimientos", "MERAK": "📊 MERAK - Movimientos"}
OUT_DIR = os.path.join(ROOT, "04. FINANZAS", "05. INFORMES")


def _monto(s) -> float | None:
    m = re.search(r"S/\s*([\d,]+(?:\.\d+)?)", str(s or ""))
    return float(m.group(1).replace(",", "")) if m else None


def _fecha(s):
    for fmt in ("%d/%m/%Y", "%Y.%m.%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(s).strip()[:10], fmt).date()
        except ValueError:
            continue
    return None


# ── 1) Leer el reporte SUNAT ────────────────────────────────────────────────
def leer_reporte(path: str):
    wb = openpyxl.load_workbook(path, data_only=True)
    ruc, pagos = None, []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            vals = [v for v in row if v not in (None, "")]
            if not vals:
                continue
            txt = " ".join(str(v) for v in vals)
            if ruc is None:
                m = re.search(r"\b(20\d{9})\b", txt)
                if m:
                    ruc = m.group(1)
            # fila de pago: ID · Período mm/yyyy · Formulario (4 dígitos o 'detr') · …
            if len(vals) >= 6 and re.fullmatch(r"\d{2}/\d{4}", str(vals[1]).strip()) \
                    and re.fullmatch(r"\d{4}|detr", str(vals[2]).strip().lower()):
                per = str(vals[1]).strip()
                form = str(vals[2]).strip().lower()
                fecha = next((_fecha(v) for v in vals[4:] if _fecha(v)), None)
                banco = "BN" if any("naci" in str(v).lower() for v in vals) else \
                        ("BCP" if any("cr" in str(v).lower() and "dito" in str(v).lower() for v in vals) else "")
                monto = next((_monto(v) for v in vals[::-1] if _monto(v) is not None), 0.0)
                pagos.append({"per": f"{per[3:]}-{per[:2]}", "form": form, "desc": str(vals[3])[:28],
                              "fecha": fecha, "banco": banco, "monto": monto or 0.0})
    emp = next((e for e, r in E.RUC.items() if r == ruc), None)
    return emp, ruc, pagos


# ── 2) Evidencia bancaria ───────────────────────────────────────────────────
def cargos_bn(emp: str) -> list:
    out = []
    for f in glob.glob(os.path.join(ROOT, BN_DIR[emp], "BN_*_2026-*_movimientos.csv")):
        for row in csv.reader(open(f, encoding="utf-8", errors="ignore")):
            if len(row) >= 7 and "78787878" in row[3]:
                try:
                    out.append({"f": _fecha(row[1]), "m": abs(float(row[6] or 0)), "usado": False})
                except ValueError:
                    pass
    return [c for c in out if c["f"] and c["m"]]


def pago_detrac_bcp(emp: str) -> list:
    out = []
    wb = S.load("conciliacion")
    if HOJA_CONC[emp] not in wb.sheetnames:
        return out
    for row in wb[HOJA_CONC[emp]].iter_rows(min_row=5, values_only=True):
        if not row or len(row) < 6:
            continue
        if "PAGO DETRAC" in str(row[2] or "").upper() and isinstance(row[5], (int, float)) and row[5]:
            out.append({"f": _fecha(row[0]) if not hasattr(row[0], "year") else row[0].date()
                        if hasattr(row[0], "date") else row[0], "m": float(row[5]), "usado": False})
    return [c for c in out if c["f"]]


def _calza(monto, fecha, candidatos, tol=1.0, dias=4):
    for c in candidatos:
        if c["usado"] or abs(c["m"] - monto) > tol:
            continue
        if fecha and c["f"] and abs((c["f"] - fecha).days) > dias:
            continue
        c["usado"] = True
        return True
    return False


def casar(pagos, bn, bcp):
    """Marca p['ok'] = True/False/None(sin cobertura). Primero cada pago solo; lo que queda,
    agrupado por (fecha, banco): el BN suele cobrar varias boletas del mismo día en UN cargo."""
    desde = {"BN": min((c["f"] for c in bn), default=None), "BCP": min((c["f"] for c in bcp), default=None)}
    for p in pagos:
        p["ok"] = None
        if p["monto"] <= 0 or not p["banco"]:
            continue
        if desde[p["banco"]] and p["fecha"] and p["fecha"] < desde[p["banco"]]:
            p["ok"] = "sin_eecc"                       # antes de la cobertura del extracto
            continue
        p["ok"] = _calza(p["monto"], p["fecha"], bn if p["banco"] == "BN" else bcp)
    grupos = defaultdict(list)
    for p in pagos:
        if p["ok"] is False:
            grupos[(p["fecha"], p["banco"])].append(p)
    for (f, b), ps in grupos.items():
        if len(ps) > 1 and _calza(sum(p["monto"] for p in ps), f, bn if b == "BN" else bcp):
            for p in ps:
                p["ok"] = True
    return desde


# ── 3) Cruce y reporte ──────────────────────────────────────────────────────
def cruzar(path: str):
    emp, ruc, pagos = leer_reporte(path)
    if not emp or not pagos:
        print(f"⚠️ {os.path.basename(path)}: no pude leer empresa/pagos (RUC {ruc}, {len(pagos)} filas)")
        return
    bn, bcp = cargos_bn(emp), pago_detrac_bcp(emp)
    desde = casar(pagos, bn, bcp)
    por_per = defaultdict(list)
    for p in pagos:
        por_per[p["per"]].append(p)

    L = [f"# Cruce tributario↔banco — {emp} (RUC {ruc})\n",
         f"Fuente SUNAT: `{os.path.basename(path)}` · evidencia: cargos 78787878 del BN y "
         f"`PAGO DETRAC` del BCP · ✓ = salida encontrada en banco (±1 sol, ±4 días; boletas del mismo "
         f"día se aceptan sumadas en un cargo) · ⚠️ = sin evidencia · ◌ = anterior al extracto disponible "
         f"(BN desde {desde['BN'] or '—'}, BCP desde {desde['BCP'] or '—'})\n",
         "| Período | PLAME 0601 | IGV-Renta 0621 | Boletas 1662 (pagos) | Detracciones depositadas | Total pagado |",
         "|---|---|---|---|---|--:|"]
    tot_ok = tot_no = 0
    for per in sorted(por_per):
        ps = por_per[per]
        def _col(filtro):
            xs = [p for p in ps if filtro(p)]
            if not xs:
                return "—"
            partes = []
            for p in xs:
                if p["monto"] <= 0:
                    partes.append("presentado (S/0)")
                else:
                    partes.append(f"{p['monto']:,.0f} " + ("✓" if p["ok"] is True else "◌" if p["ok"] == "sin_eecc" else "⚠️"))
            return " · ".join(partes)
        total = sum(p["monto"] for p in ps)
        L.append(f"| {per} | {_col(lambda p: p['form']=='0601')} | {_col(lambda p: p['form']=='0621')} "
                 f"| {_col(lambda p: p['form']=='1662')} | {_col(lambda p: p['form']=='detr')} | {total:,.0f} |")
        tot_ok += sum(1 for p in ps if p["ok"] is True); tot_no += sum(1 for p in ps if p["ok"] is False)
    tot_sin = sum(1 for p in pagos if p["ok"] == "sin_eecc")
    L.append(f"\n**{tot_ok} pagos con evidencia bancaria · {tot_no} sin evidencia · {tot_sin} fuera de la "
             f"cobertura del extracto** (⚠️ = revisar fecha fuera de ±4 días, pago desde otra cuenta o "
             f"extracto incompleto; ◌ = pedir el extracto BN de esos meses para cerrar el cruce).")
    # sentido inverso: salidas a SUNAT en el banco que el reporte NO explica
    huerfanos = [c for c in bn if not c["usado"]]
    if huerfanos:
        L.append("\n**Cargos BN 78787878 sin constancia en el reporte SUNAT** (períodos posteriores al "
                 "reporte o pagos que el contador debe explicar):")
        for c in sorted(huerfanos, key=lambda c: c["f"]):
            L.append(f"- {c['f']:%d/%m/%Y} · S/ {c['m']:,.2f}")
    L.append("\n_Lectura: la fila muestra lo que SUNAT registró; el ✓ prueba que el dinero salió de "
             "nuestras cuentas. Formularios con S/0 = declaración presentada sin pago (normal en 0621 "
             "cuando el IGV va por boleta 1662 aparte)._")
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, f"CRUCE_TRIBUTOS_{emp[:3]}.md")
    open(out, "w", encoding="utf-8").write("\n".join(L) + "\n")
    print("\n".join(L))
    print(f"\n💾 {out}")


def main():
    rutas = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not rutas:
        rutas = glob.glob(os.path.join(ROOT, "05. TRIBUTARIO", "*", "2026", "*", "REPORTE-TRIBUTOS_*.xlsx"))
    if not rutas:
        sys.exit("❌ No hay REPORTE-TRIBUTOS_*.xlsx (consulta SOL 'Relación de constancias de pago').")
    for r in rutas:
        for f in glob.glob(r):
            cruzar(f)


if __name__ == "__main__":
    main()
