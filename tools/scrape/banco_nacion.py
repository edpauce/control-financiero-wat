#!/usr/bin/env python3
"""
tools/scrape/banco_nacion.py — Movimientos de las cuentas de detracciones (Banco de la Nación).

FLUJO VERIFICADO (2026-06-22, vía Claude in Chrome — login asistido por el usuario; el resto se
automatizó sin problemas). El portal MultiRed Virtual del BN es un formulario CLÁSICO (no Angular):
acepta fechas escritas y se deja operar bien. Es el más automatizable de los portales hasta ahora.

── Flujo (Chrome MCP) ───────────────────────────────────────────────────────
1. El USUARIO abre el portal del BN (https://bancaporinternet.bn.com.pe) e inicia sesión con
   clave + token. (BN es por RUC/usuario → una sesión por empresa.)
2. "Consulta de Cuentas" → en la sección "CUENTAS CORRIENTES DETRACCIONES", abrir el selector
   "TIPO DE CONSULTA" de la cuenta y elegir **Movimientos**.
3. Fecha de Inicio (escribir DD/MM/YYYY) y Fecha de Término (calendario emergente, clic en el día).
   Transacción = "Todas" (incluye ingresos y egresos). Máx. 30 días atrás. Clic **CONSULTAR**.
4. En los resultados, clic **"Exportar a: Excel"** → descarga `Movimientos.xls` en ~/Downloads.
5. Archivar/parsear con las funciones de abajo.

⚠️ GOTCHAS:
- El formulario está dentro de un **iframe** → read_page/find NO lo ven; operar por COORDENADAS
  (screenshot + click). Los screenshots de este portal SÍ funcionan.
- La conexión de la extensión cae a veces a mitad de batch → reintentar; el formulario conserva estado.
- `Movimientos.xls` es **BIFF real** (firma OLE2 \\xd0\\xcf\\x11\\xe0) → leer con **xlrd** (no openpyxl).
- Los montos vienen como **texto con coma de miles** ("1,260.00") en columnas separadas **Cargo** y
  **Abono** → quitar la coma antes de float (ver _num).
- Cada empresa tiene su propia sesión → repetir el flujo por empresa.

Códigos del EECC BN (ver lib/empresas.py): la columna Oficina (`1234`, `78788888`, `EE 0456`)
describe el CANAL por el que entró el movimiento y NO dice quién depositó — la heurística
"1234 = cliente / 78788888 = auto-detracción" quedó desmentida por la evidencia. La
autoría se decide con la salida `PAGO DETRAC` del BCP o el rol en la constancia SUNAT
(`lib.empresas.quien_deposito`). Cargos `78787878` = pagos virtuales de tributos. La columna RUC es
la del adquiriente (cliente), no la del depositante.
Columnas del export: Nro, Fecha, Trans (NOT.ABON/NOT.CARG), Documento, RUC, Oficina, Cargo, Abono.
"""
from __future__ import annotations
import csv
import glob
import os
import shutil
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from lib import empresas as E       # noqa: E402

PORTAL = "banco_nacion"
ENV_PREFIX = "BN_{ALT,MER}"
PORTAL_URL = "https://bancaporinternet.bn.com.pe"

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

# Mapa empresa → (cuenta detracciones, sufijo de nombre, carpeta EECC_BN relativa a la raíz).
# La cuenta sale de lib/empresas.py (fuente única; antes estaba duplicada aquí).
EMPRESAS = {
    "ALTAMIRA": {"cuenta": E.CUENTA_BN["ALTAMIRA"], "suf": "ALT",
                 "eecc": "01. ALTAMIRA/03. DETRACCIONES/2026/01. EECC_BN"},
    "MERAK":    {"cuenta": E.CUENTA_BN["MERAK"], "suf": "MER",
                 "eecc": "02. MERAK/03. DETRACCIONES/2026/01. EECC_BN"},
}


def _num(x) -> float:
    try:
        return float(str(x).replace(",", "").strip()) if str(x).strip() else 0.0
    except (ValueError, AttributeError):
        return 0.0


def parse_movimientos_bn(path: str) -> dict:
    """
    Lee el export del BN (Movimientos.xls BIFF, un .xlsx, o un .csv ya convertido) y devuelve
    {movimientos:[{nro,fecha,trans,documento,ruc,oficina,cargo,abono}], abonos, cargos, neto, n}.
    """
    rows = []
    low = path.lower()
    if low.endswith(".xlsx"):
        # xlrd 2.x solo abre BIFF (.xls): un .xlsx aquí reventaba con XLRDError → openpyxl
        import openpyxl
        ws = openpyxl.load_workbook(path, data_only=True, read_only=True).active
        rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
    elif low.endswith(".xls"):
        import xlrd                        # Movimientos.xls es BIFF real (firma OLE2)
        sh = xlrd.open_workbook(path).sheet_by_index(0)
        rows = [[str(sh.cell_value(i, c)).strip() for c in range(sh.ncols)] for i in range(sh.nrows)]
    else:
        text = open(path, "rb").read().decode("latin-1", errors="replace")
        rows = list(csv.reader(text.splitlines()))
    # Localizar la fila de cabecera y quedarnos con las filas de datos (Nro numérico)
    movs = []
    for r in rows:
        if not r or not str(r[0]).strip().isdigit():
            continue
        r = (r + [""] * 8)[:8]
        movs.append({"nro": r[0], "fecha": r[1], "trans": r[2], "documento": r[3],
                     "ruc": r[4], "oficina": r[5], "cargo": _num(r[6]), "abono": _num(r[7])})
    # Aviso si el archivo trae contenido pero NO se parseó ninguna fila (formato cambió del banco)
    if len(rows) > 6 and not movs:
        print(f"⚠️ ADVERTENCIA: {os.path.basename(path)} tiene {len(rows)} filas pero 0 movimientos "
              f"parseados — ¿cambió el formato del export BN?", file=sys.stderr)
    abonos = round(sum(m["abono"] for m in movs), 2)
    cargos = round(sum(m["cargo"] for m in movs), 2)
    return {"movimientos": movs, "n": len(movs), "abonos": abonos,
            "cargos": cargos, "neto": round(abonos + cargos, 2)}


def archivar_movimientos(empresa: str, anio: int, mes: int, downloads: str | None = None) -> list[str]:
    """
    Mueve el `Movimientos.xls` más reciente de ~/Downloads a la carpeta EECC_BN de la empresa, y
    además escribe una copia `.csv` legible (encabezado con cuenta/periodo + filas). Devuelve rutas.
    """
    info = EMPRESAS[empresa]
    dl = downloads or os.path.join(os.path.expanduser("~"), "Downloads")
    cand = sorted(glob.glob(os.path.join(dl, "Movimientos*.xls")), key=os.path.getmtime, reverse=True)
    if not cand:
        raise FileNotFoundError(f"No se encontró Movimientos*.xls en {dl} (¿se exportó del BN?)")
    src = cand[0]
    # ⚠️ El export 'Movimientos.xls' del BN NO contiene la cuenta ni el RUC de la empresa (solo la
    # tabla de movimientos; la columna 'RUC' es la del cliente). No hay forma fiable de validar la
    # empresa desde el contenido → la salvaguarda es NO sobrescribir un destino existente (abajo) y
    # confiar en --empresa. El usuario debe descargar la cuenta correcta.
    dest_dir = os.path.join(ROOT, info["eecc"])
    os.makedirs(dest_dir, exist_ok=True)
    # Convención BN_{ALT,MER}_AAAA-MM_movimientos.* — la que leen los integradores y sources.py
    base = f"BN_{info['suf']}_{anio}-{mes:02d}_movimientos"

    # CSV legible
    rep = parse_movimientos_bn(src)
    csv_path = os.path.join(dest_dir, base + ".csv")
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["Cuenta", f"{info['cuenta']} - {empresa}"])
        w.writerow(["Periodo", f"{anio}-{mes:02d}"])
        w.writerow([])
        w.writerow(["Nro", "Fecha", "Trans", "Documento", "RUC", "Oficina", "Cargo", "Abono"])
        for m in rep["movimientos"]:
            w.writerow([m["nro"], m["fecha"], m["trans"], m["documento"], m["ruc"],
                        m["oficina"], m["cargo"] or "", m["abono"] or ""])
    # Mover el .xls original — no sobrescribir si ya existe (evita pisar el de la otra empresa/mes)
    xls_path = os.path.join(dest_dir, base + ".xls")
    if os.path.exists(xls_path):
        raise FileExistsError(f"Ya existe {os.path.basename(xls_path)} — no se sobrescribe. "
                              f"Renómbralo/muévelo antes si de verdad quieres reemplazarlo.")
    shutil.move(src, xls_path)
    return [csv_path, xls_path]


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Archivar/parsear movimientos BN descargados")
    ap.add_argument("--empresa", required=True, choices=list(EMPRESAS))
    ap.add_argument("--anio", type=int, default=2026)
    ap.add_argument("--mes", type=int, required=True)
    ap.add_argument("--solo-parsear", metavar="RUTA",
                    help="No archiva; sólo parsea el archivo indicado y muestra el resumen")
    a = ap.parse_args()
    if a.solo_parsear:
        rep = parse_movimientos_bn(a.solo_parsear)
        print(f"{rep['n']} movimientos · abonos +{rep['abonos']:,.2f} · "
              f"cargos {rep['cargos']:,.2f} · neto {rep['neto']:,.2f}")
    else:
        for p in archivar_movimientos(a.empresa, a.anio, a.mes):
            print("✅", p)
