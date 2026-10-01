"""
tools/lib/sources.py — Fuente única de verdad para localizar y leer los archivos del proyecto.

Todo tool del proyecto importa de aquí. NO se usan rutas fijas con fecha en ningún lado:
`latest(clave)` siempre devuelve el archivo más reciente por glob (misma convención que el
dashboard histórico). El esquema de hojas/columnas está documentado en references/fuentes-datos.md.

Uso:
    from lib import sources
    wb = sources.load("conciliacion")          # openpyxl workbook (data_only)
    ruta = sources.latest("detracciones")      # ruta del archivo más reciente
    sources.ROOT                               # raíz del proyecto
"""
from __future__ import annotations
import glob
import os
import re
from datetime import date, datetime

# ─── Raíz del proyecto ───────────────────────────────────────────────────────
# Este archivo vive en  <ROOT>/tools/lib/sources.py  → ROOT son dos niveles arriba.
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


def _resolve_root() -> str:
    """Permite override por env var EMPRESA_ROOT (útil en Cowork / rutas montadas)."""
    env = os.environ.get("EMPRESA_ROOT")
    if env and os.path.isdir(env):
        return os.path.abspath(env)
    return ROOT


# ─── Catálogo de fuentes (globs relativos a la raíz) ─────────────────────────
PATHS: dict[str, str] = {
    # El año va en glob: fijarlo a 2026 dejaría al pipeline leyendo archivos viejos en enero 2027.
    "conciliacion": "03. CONCILIACIONES/GLOBAL/CONCILIACION_*_REV_*.xlsx",
    "detracciones": "04. FINANZAS/10. DETRACCIONES/SEGUIMIENTO_DETRACCIONES_*.xlsx",
    "flujo":        "04. FINANZAS/03. FLUJO DE CAJA/FLUJO_CAJA_*.xlsx",
    "cobros":       "04. FINANZAS/08. COBROS/COBROS_*.xlsx",
    "proveedores":  "04. FINANZAS/09. PROVEEDORES/PROVEEDORES_*.xlsx",
    "creditos":     "04. FINANZAS/01. CREDITOS/CREDITOS_*.xlsx",
    "mapeo":        "08. HERRAMIENTAS/MAPEO_CONTABLE.xlsx",
    "directorio":   "08. HERRAMIENTAS/DIRECTORIO_CUENTAS.xlsx",
    # ─ Fuentes crudas descargadas de portales (las vigila tools/auditar_fuentes.py) ─
    "eecc_alt":     "01. ALTAMIRA/01. EECC/*/EECC_ALT_*",
    "eecc_mer":     "02. MERAK/01. EECC/*/EECC_MER_*",
    "bn_alt":       "01. ALTAMIRA/03. DETRACCIONES/*/01. EECC_BN/BN_ALT_*",
    "bn_mer":       "02. MERAK/03. DETRACCIONES/*/01. EECC_BN/BN_MER_*",
    "efact_alt":    "01. ALTAMIRA/02. FACTURACION/*/Emitidas/FACT-EMI_ALT_*_Reporte.xlsx",
    "efact_mer":    "02. MERAK/02. FACTURACION/*/Emitidas/FACT-EMI_MER_*_Reporte.xlsx",
}

# Carpeta física del archivo tributario (no es Excel; se recorre por subcarpetas)
TRIBUTARIO_DIR = "05. TRIBUTARIO"


def _latest_glob(pattern: str) -> str | None:
    """Archivo más reciente que coincide con el glob. Ignora temporales de Excel."""
    matches = [
        f for f in glob.glob(pattern, recursive=True)
        if not os.path.basename(f).startswith(("~$", ".~lock"))
    ]
    return max(matches, key=os.path.getmtime) if matches else None


def latest(key_or_glob: str) -> str | None:
    """
    Ruta del archivo más reciente para una clave de PATHS, o para un glob arbitrario
    (relativo a la raíz o absoluto).
    """
    root = _resolve_root()
    if key_or_glob in PATHS:
        pattern = os.path.join(root, PATHS[key_or_glob])
    elif os.path.isabs(key_or_glob):
        pattern = key_or_glob
    else:
        pattern = os.path.join(root, key_or_glob)
    return _latest_glob(pattern)


def load(key: str, data_only: bool = True):
    """Abre con openpyxl el archivo más reciente de la clave dada. Lanza si no existe."""
    import openpyxl
    ruta = latest(key)
    if not ruta:
        raise FileNotFoundError(f"No se encontró ninguna fuente para '{key}' "
                                f"(glob: {PATHS.get(key, key)})")
    return openpyxl.load_workbook(ruta, data_only=data_only)


def all_sources() -> dict[str, str | None]:
    """Devuelve {clave: ruta_resuelta_o_None} para diagnóstico."""
    return {k: latest(k) for k in PATHS}


# ─── Helpers de parseo (compartidos por todos los tools) ─────────────────────
def num(v) -> float:
    """Convierte a float de forma tolerante. None/''/fórmulas → 0.0."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if not s or s.startswith("="):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date(v):
    """Devuelve datetime.date o None a partir de date/datetime/str (DD/MM/YYYY o ISO)."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:19] if " " in s else s[:10], fmt).date()
        except ValueError:
            pass
    return None


def fmt_money(v, simbolo="S/ ") -> str:
    """'S/ 1,234.56' — negativos entre paréntesis."""
    if v is None or v == 0:
        return f"{simbolo}0"
    s = f"{abs(v):,.2f}"
    return f"({simbolo}{s})" if v < 0 else f"{simbolo}{s}"


# Constantes de meses (orden contable)
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


if __name__ == "__main__":
    # Diagnóstico rápido: ¿resuelven todas las fuentes?
    print(f"EMPRESA_ROOT = {_resolve_root()}\n")
    ok = True
    for k, ruta in all_sources().items():
        flag = "✅" if ruta else "❌"
        if not ruta:
            ok = False
        print(f"  {flag} {k:14s} {os.path.basename(ruta) if ruta else 'NO ENCONTRADO'}")
    print("\n", "✅ Todas las fuentes resuelven" if ok else "❌ Faltan fuentes")
