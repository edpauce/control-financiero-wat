#!/usr/bin/env python3
"""
tools/scrape/bcp_empresas.py — Movimientos (EECC) de Telecrédito BCP por empresa.

FLUJO VERIFICADO (2026-06-22, vía Claude in Chrome — login + filtrado asistidos por el usuario).
Telecrédito es por RUC (una sesión por empresa) y exige token digital → el LOGIN siempre lo hace el
usuario. La navegación es parcialmente automatizable, pero el FORMULARIO de filtro NO acepta entrada
sintética de forma fiable (ver gotcha #2) → el FILTRADO se co-conduce con el usuario. La parte
determinista (archivar + parsear los movimientos con sus referencias) está implementada abajo.

── Flujo (Chrome MCP, asistido) ─────────────────────────────────────────────
1. El USUARIO abre Telecrédito (https://www.tlcbcp.com) desde su acceso de confianza e inicia sesión
   con clave + token. (La IA nunca ingresa credenciales ni token.)
2. Ir a Saldos y Movimientos → pestaña "Cuentas".
   URL directa útil: https://www.tlcbcp.com/#/h/saldos-movimientos/v2/cuentas
3. En la fila de la cuenta operativa, clic en el ícono de RELOJ (= "Ver movimientos históricos").
   ⚠️ El ícono de OJO es "Ver detalle"; el RELOJ es históricos. Verificar que la página resultante
   muestre la cuenta correcta ("Corriente Soles 193-...").
4. Filtro: Desde / Hasta en formato DD/MM/YYYY (máx. 31 días; sólo últimos 180 días disponibles),
   Tipo de movimiento = "Todas" (incluye ingresos y egresos). Clic "Buscar".
   ⚠️ GOTCHA #2: el formulario Angular sólo valida con PULSACIONES REALES. La escritura/selección
   automatizada se ve en pantalla pero NO persiste en el modelo → Buscar queda deshabilitado o
   reinicia. Por eso este paso lo hace el usuario. (Detectar Buscar habilitado vía DOM no basta.)
5. Exportar → CSV/Excel. El histórico NO trae el día de hoy (corta al día anterior); para HOY usar
   la pestaña "Detalle de cuenta" → "últimos movimientos" y exportar también.
   ⚠️ GOTCHA #3: los screenshots de esta página se congelan a veces (CDP timeout) → usar read_page/find.

── Archivos que produce BCP (en ~/Downloads) ────────────────────────────────
- Histórico:  `<sufijo>_movimientos_historicos.csv|.txt`  (latin-1, separador coma) — CON referencias.
  Columnas: Fecha, Fecha valuta, Descripción operación, Monto, Saldo, Operación - Número, Referencia2.
- Hoy:        `<sufijo>_ultimos_movimientos.xlsx` (hoja "AccountDetail") — incluye el día actual.
  Columnas: Fecha, Fecha valuta(vacía), Descripción, Monto, Sucursal-agencia, Nº operación, Usuario.
- `<sufijo>` = los 3 dígitos antes del verificador: ALTAMIRA=011 (…-0-11), MERAK=021 (…-0-21).

Refs útiles para conciliar: "Operación - Número" (todas) + "Referencia2" (glosa, sólo transferencias).
Egresos a otra cuenta aparecen como "A 193 1000101 0" (nº de cuenta destino) en la descripción.
"""
from __future__ import annotations
import csv
import glob
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from lib import empresas as E       # noqa: E402

PORTAL = "bcp_empresas"
ENV_PREFIX = "BCP_{ALT,MER}"
TELECREDITO_URL = "https://www.tlcbcp.com"

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

# Mapa empresa → (cuenta, sufijo de archivo, carpeta EECC relativa a la raíz).
# La cuenta sale de lib/empresas.py (fuente única; antes estaba duplicada aquí).
EMPRESAS = {
    "ALTAMIRA": {"cuenta": E.CUENTA_BCP["ALTAMIRA"], "sufijo": "011", "abrev": "ALT",
                 "eecc": "01. ALTAMIRA/01. EECC"},
    "MERAK":    {"cuenta": E.CUENTA_BCP["MERAK"], "sufijo": "021", "abrev": "MER",
                 "eecc": "02. MERAK/01. EECC"},
}


def _num(x) -> float:
    try:
        return float(str(x).replace(",", "").strip())
    except (ValueError, AttributeError):
        return 0.0


def _downloads(d: str | None = None) -> str:
    return d or os.path.join(os.path.expanduser("~"), "Downloads")


def _descargas(dl: str, base: str) -> list[str]:
    """Archivos `<base>*` de la carpeta de descargas, el más reciente (mtime) primero.
    El glob es `base*` y no `base.*`: el navegador numera las repeticiones como
    `<base> (1).csv` (segunda descarga del día) y ese —el más nuevo— era justo el que no se
    encontraba. Se ignoran descargas a medias (.crdownload/.part)."""
    cand = [f for f in glob.glob(os.path.join(dl, base + "*"))
            if not f.lower().endswith((".crdownload", ".part", ".download"))]
    return sorted(cand, key=os.path.getmtime, reverse=True)


def archivar_movimientos(empresa: str, anio: int, mes: int, downloads: str | None = None) -> list[str]:
    """
    Mueve los archivos de movimientos BCP de la empresa desde ~/Downloads a su carpeta EECC, con
    nombres ordenados. Busca el histórico (<sufijo>_movimientos_historicos*) y el de últimos
    movimientos (<sufijo>_ultimos_movimientos*), el más reciente por mtime. Devuelve las rutas destino.
    """
    info = EMPRESAS[empresa]
    suf, ab = info["sufijo"], info["abrev"]
    dl = _downloads(downloads)
    dest_dir = os.path.join(ROOT, info["eecc"], str(anio))
    os.makedirs(dest_dir, exist_ok=True)

    movidos = []
    # Histórico (csv o txt) → convención EECC_{ALT,MER}_AAAA-MM_movimientos.csv (la leen los integradores)
    hist = _descargas(dl, f"{suf}_movimientos_historicos")
    if hist:
        destino = os.path.join(dest_dir, f"EECC_{ab}_{anio}-{mes:02d}_movimientos.csv")
        shutil.move(hist[0], destino)
        movidos.append(destino)
    # Últimos movimientos (incluye hoy) — BCP exporta .csv o .xlsx según la pantalla
    ult = _descargas(dl, f"{suf}_ultimos_movimientos")
    if ult:
        ext = os.path.splitext(ult[0])[1].lower()
        destino = os.path.join(dest_dir, f"EECC_{ab}_{anio}-{mes:02d}_UltimosMov{ext}")
        shutil.move(ult[0], destino)
        movidos.append(destino)
    if not movidos:
        raise FileNotFoundError(
            f"No se encontraron archivos {suf}_movimientos_historicos* / {suf}_ultimos_movimientos* en {dl}")
    return movidos


def _fecha_txt(v) -> str | None:
    """'DD/MM/AAAA' desde texto o desde un datetime (Excel guarda fechas, no cadenas)."""
    from datetime import date as _d, datetime as _dt
    if isinstance(v, (_dt, _d)):
        return v.strftime("%d/%m/%Y")
    s = str(v or "").strip()
    return s if re.match(r"\d{1,2}/\d{1,2}/\d{4}", s) else None


def _delimitador(texto: str) -> str:
    """El BCP exporta .csv con coma y .txt con pipe. Se detecta en la línea de encabezado."""
    linea = next((l for l in texto.splitlines() if "Fecha" in l), texto[:400])
    return max(",;|\t", key=linea.count)


def _filas(path: str):
    """(cuenta, encabezado, filas) del histórico BCP, sea .csv/.txt (latin-1) o .xlsx."""
    if path.lower().endswith(".xlsx"):
        import openpyxl
        ws = openpyxl.load_workbook(path, data_only=True).active
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    else:
        text = open(path, "rb").read().decode("latin-1", errors="replace")
        rows = list(csv.reader(text.splitlines(), delimiter=_delimitador(text)))
    def _c(r, i):
        return str(r[i]).strip() if r and len(r) > i and r[i] is not None else ""
    cuenta = next((_c(r, 1) for r in rows if _c(r, 0) == "Cuenta"), "")
    hi = next((i for i, r in enumerate(rows) if _c(r, 0) == "Fecha"), None)
    if hi is None:
        return cuenta, [], rows
    return cuenta, [_c(rows[hi], i) for i in range(len(rows[hi]))], rows[hi + 1:]


def parse_movimientos(path: str) -> dict:
    """
    Parsea el histórico BCP a movimientos estructurados con sus referencias.
    Acepta .csv/.txt (latin-1, coma) y .xlsx — el .xlsx trae columnas extra (Sucursal, Hora,
    Usuario, UTC), así que las columnas se resuelven POR NOMBRE, nunca por posición.
    Devuelve {cuenta, movimientos:[{fecha,desc,monto,saldo,operacion,referencia}], ingresos,
    egresos, neto, n}.
    """
    cuenta, hdr, filas = _filas(path)
    idx = {}
    for clave, etiqueta in (("fecha", "fecha"), ("desc", "descripci"), ("monto", "monto"),
                            ("saldo", "saldo"), ("operacion", "operación - número"),
                            ("referencia", "referencia")):
        idx[clave] = next((i for i, h in enumerate(hdr)
                           if etiqueta in h.lower() and "valuta" not in h.lower()), None)
    movs = []
    for r in filas:
        f = _fecha_txt(r[idx["fecha"]]) if idx["fecha"] is not None and len(r) > idx["fecha"] else None
        if not f:
            continue
        def g(k):
            i = idx[k]
            return r[i] if i is not None and len(r) > i and r[i] is not None else ""
        movs.append({"fecha": f, "desc": str(g("desc")).strip(), "monto": _num(g("monto")),
                     "saldo": _num(g("saldo")), "operacion": str(g("operacion")).strip(),
                     "referencia": str(g("referencia")).strip()})
    if filas and not movs:
        print(f"⚠️ ADVERTENCIA: {os.path.basename(path)} tiene {len(filas)} filas tras el "
              f"encabezado pero 0 movimientos parseados — ¿cambió el formato del export BCP?",
              file=sys.stderr)
    ing = sum(m["monto"] for m in movs if m["monto"] > 0)
    egr = sum(m["monto"] for m in movs if m["monto"] < 0)
    return {"cuenta": cuenta, "movimientos": movs, "n": len(movs),
            "ingresos": round(ing, 2), "egresos": round(egr, 2), "neto": round(ing + egr, 2)}


# Nombre histórico: el histórico ya no es solo CSV, pero muchos tools lo llaman así.
parse_movimientos_csv = parse_movimientos


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Archivar/parsear movimientos BCP descargados")
    ap.add_argument("--empresa", required=True, choices=list(EMPRESAS))
    ap.add_argument("--anio", type=int, default=2026)
    ap.add_argument("--mes", type=int, required=True)
    ap.add_argument("--solo-parsear", metavar="RUTA",
                    help="No archiva; sólo parsea el CSV indicado y muestra el resumen")
    a = ap.parse_args()
    if a.solo_parsear:
        rep = parse_movimientos_csv(a.solo_parsear)
        print(f"{rep['cuenta']} · {rep['n']} movimientos · "
              f"ingresos +{rep['ingresos']:,.2f} · egresos {rep['egresos']:,.2f} · neto {rep['neto']:,.2f}")
    else:
        for p in archivar_movimientos(a.empresa, a.anio, a.mes):
            print("✅", p)
