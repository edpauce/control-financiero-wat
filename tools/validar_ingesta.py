#!/usr/bin/env python3
"""
tools/validar_ingesta.py — P1a: compuerta de ingesta del cierre (gate mecánico, fatal).

Valida el CONTENIDO de cada fuente cruda antes de que el pipeline calcule un solo número.
Complementa a auditar_fuentes.py, que solo juzga frescura por el NOMBRE del archivo — y el
nombre miente: BN_MER_2026-03.pdf contenía enero, EECC_*_2026-01.xlsx contenía abril.

Cada chequeo nace de una falla silenciosa REAL que ya costó dinero o meses de ceguera:
  a. parseable / filas>0 ......... el .txt de mayo ALT parseó a 0 movimientos por el pipe
  b. fechas del periodo .......... archivos cuyo contenido es de otro mes que el del nombre
  c. cobertura hasta el corte .... el histórico bajado el día 30 llegaba al 29 (−12 movs, −S/3,300)
  d. conteo plausible ............ ±40% vs el mes anterior (SOLO AVISO: un mes atípico es legítimo)
  e. exceso sobre el corte ....... con --corte 15/MM, filas POSTERIORES al corte son FATAL: el
                                   pipeline no recorta, integraría la 2ª quincena etiquetada "al 15"

Política de severidad (calibrada contra los meses ya cerrados — criterio de aceptación del plan):
  · Corrupción (no parsea, filas>0 fallido, mes equivocado) ...... FATAL siempre.
  · Ausencia/cobertura de BCP histórico y efact ................. FATAL; --permitir-incompleto
    la baja a AVISO (cierres provisionales a mitad de mes).
  · BN movimientos y constancias SUNAT ausentes ................. AVISO (mayo cerró sin BN csv;
    SUNAT puede legítimamente no tener pagos). Presentes pero corruptos → FATAL.

Reporta TODOS los problemas de una pasada (no aborta en el primero) y sale con código 1 si hay
algún FATAL. Corre como paso 0 de integrar_mes.py, ANTES de tocar cualquier maestro.

Uso:  python3 tools/validar_ingesta.py --mes 2026-07 [--permitir-incompleto]
"""
from __future__ import annotations

import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import cierre as CI                 # noqa: E402
from scrape import bcp_empresas as BCP       # noqa: E402

FATAL, AVISO = "❌ FATAL", "⚠️ AVISO"

# Variación de conteo de filas vs el mes anterior que se considera plausible (criterio del plan).
TOLERANCIA_CONTEO = 0.40


def _fecha_mes(v) -> str | None:
    """'AAAA-MM' desde una celda de fecha en cualquiera de los formatos de las fuentes:
    'DD/MM/AAAA' (BCP), '2026.06.24' (BN), datetime (efact/xlsx)."""
    import datetime as _dt
    if isinstance(v, (_dt.datetime, _dt.date)):
        return f"{v.year:04d}-{v.month:02d}"
    s = str(v or "").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}"
    m = re.match(r"^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    return None


def _meses_de(fechas: list) -> dict[str, int]:
    from collections import Counter
    return dict(Counter(m for m in (_fecha_mes(f) for f in fechas) if m))


def _fkey_flex(v) -> tuple | None:
    """Clave (aaaa,mm,dd) desde datetime (efact/xlsx) o 'DD/MM/AAAA' (BCP). None si no parsea."""
    import datetime as _dt
    if isinstance(v, (_dt.datetime, _dt.date)):
        return (f"{v.year:04d}", f"{v.month:02d}", f"{v.day:02d}")
    s = str(v or "").strip()
    if re.match(r"^\d{1,2}/\d{1,2}/\d{4}", s):
        return CI.fkey(s)
    return None


def _chequear_exceso_corte(fechas: list, cfg, corte: tuple, etiqueta: str,
                           hallazgos: list) -> None:
    """FATAL si hay filas POSTERIORES a un corte quincenal declarado. El pipeline NO recorta:
    integraría la segunda quincena en un cierre etiquetado 'al 15' — hay que re-descargar el
    rango exacto. Corrupción de rango, así que nunca se degrada con --permitir-incompleto."""
    if int(cfg.CORTE[:2]) >= CI.ultimo_dia(cfg.ANIO, cfg.MES):
        return                                    # corte = fin de mes: no aplica
    tarde = [k for k in (_fkey_flex(f) for f in fechas) if k and k > corte]
    if tarde:
        a, m, d = max(tarde)
        hallazgos.append((FATAL, etiqueta,
                          f"{len(tarde)} fila(s) POSTERIORES al corte {cfg.CORTE} (hasta el "
                          f"{d}/{m}/{a}) — el pipeline no recorta al corte: re-descarga con "
                          f"rango exacto 01/{cfg.MES:02d}–{cfg.CORTE}"))


def _chequear_periodo(fechas: list, periodo: str, etiqueta: str, hallazgos: list,
                      permitir_fuera: bool = False) -> None:
    """FATAL si el contenido no pertenece al periodo. El caso real: nombre dice un mes,
    las filas son de otro."""
    meses = _meses_de(fechas)
    if not meses:
        hallazgos.append((FATAL, etiqueta, "no se pudo leer ninguna fecha del contenido"))
        return
    if periodo not in meses:
        real = max(meses, key=meses.get)
        hallazgos.append((FATAL, etiqueta,
                          f"el contenido es de {real} ({meses[real]} filas), no de {periodo} — "
                          f"el nombre del archivo miente"))
    elif not permitir_fuera:
        fuera = {m: n for m, n in meses.items() if m != periodo}
        if fuera:
            hallazgos.append((AVISO, etiqueta,
                              f"{sum(fuera.values())} filas de otros meses {sorted(fuera)} "
                              f"mezcladas con {periodo}"))


def _conteo_previo(cfg, empresa: str) -> int | None:
    """Filas del histórico BCP del mes anterior, si existe (para la plausibilidad ±40%)."""
    d, ab = CI.DIR_EMP[empresa], CI.ABREV_EMP[empresa]
    for anio in (cfg.PREV_PERIODO[:4],):
        f = CI.first(f"{d}/01. EECC/{anio}/EECC_{ab}_{cfg.PREV_PERIODO}_movimientos.*")
        if f:
            try:
                return BCP.parse_movimientos(f)["n"]
            except Exception:
                return None
    return None


# ─── P1b: heurística de identidad (WARN-ONLY el primer ciclo) ────────────────
# El archivo descargado declara INTERNAMENTE de quién es (cuenta BCP en la cabecera, cuenta BN,
# RUC del proveedor en las constancias). Si eso no cuadra con la carpeta/empresa donde se
# archivó, alguien guardó la descarga de una empresa como si fuera de la otra — la salvaguarda
# que efact no puede dar (su reporte no trae el RUC del emisor). AVISO por defecto hasta
# calibrar un ciclo de cierre; --identidad-fatal lo promueve (plan CEO: P1b warn-only).
def _digitos(s) -> str:
    return re.sub(r"\D", "", str(s or ""))


def _de_quien(cuenta_o_ruc: str, tabla: dict) -> str | None:
    """¿A qué empresa pertenece esta cuenta/RUC según los maestros? None si a ninguna."""
    dg = _digitos(cuenta_o_ruc)
    if not dg:
        return None
    for emp, val in tabla.items():
        v = _digitos(val)
        if v and (v in dg or dg in v):
            return emp
    return None


def identidad_bcp(cuenta_declarada: str, emp: str, nivel: str, hallazgos: list) -> None:
    from lib import empresas as E
    dueno = _de_quien(cuenta_declarada, E.CUENTA_BCP)
    if dueno is None and _digitos(cuenta_declarada):
        hallazgos.append((nivel, f"{CI.ABREV_EMP[emp]} · EECC BCP",
                          f"identidad: el archivo declara la cuenta {cuenta_declarada!r}, que no "
                          f"es de ninguna empresa del grupo"))
    elif dueno and dueno != emp:
        hallazgos.append((nivel, f"{CI.ABREV_EMP[emp]} · EECC BCP",
                          f"identidad: el archivo declara la cuenta de {dueno} "
                          f"({cuenta_declarada!r}) pero está archivado como {emp} — "
                          f"¿se guardó la descarga de la otra empresa?"))


def identidad_bn(path: str, emp: str, nivel: str, hallazgos: list) -> None:
    from lib import empresas as E
    import csv as _csv
    try:
        texto = open(path, "rb").read().decode("latin-1", "replace")
    except OSError:
        return
    primera = next(_csv.reader(texto.splitlines()), [])
    cuenta = primera[1] if len(primera) > 1 and str(primera[0]).strip() == "Cuenta" else ""
    dueno = _de_quien(cuenta, E.CUENTA_BN)
    if dueno and dueno != emp:
        hallazgos.append((nivel, f"{CI.ABREV_EMP[emp]} · BN detracciones",
                          f"identidad: la cuenta BN declarada ({cuenta.strip()!r}) es de {dueno}, "
                          f"no de {emp}"))


def identidad_sunat(path: str, emp: str, nivel: str, hallazgos: list) -> None:
    from lib import empresas as E
    import csv as _csv
    try:
        texto = open(path, "rb").read().decode("latin-1", "replace")
    except OSError:
        return
    rows = list(_csv.reader(texto.splitlines()))
    if len(rows) < 2:
        return
    hdr = rows[0]
    iR = next((i for i, h in enumerate(hdr) if "ruc proveedor" in h.lower()), None)
    if iR is None:
        return
    rucs = {r[iR].strip() for r in rows[1:] if len(r) > iR and r[iR].strip()}
    ajenos = {rc for rc in rucs if rc != E.RUC[emp]}
    if ajenos:
        de = {rc: (_de_quien(rc, E.RUC) or "desconocido") for rc in ajenos}
        hallazgos.append((nivel, f"{CI.ABREV_EMP[emp]} · constancias SUNAT",
                          f"identidad: el CSV trae constancias con RUC de proveedor ajeno a "
                          f"{emp}: " + ", ".join(f"{rc} ({q})" for rc, q in de.items())))


def validar(cfg, permitir_incompleto: bool = False,
            identidad_fatal: bool = False) -> list[tuple[str, str, str]]:
    """Corre todos los chequeos y devuelve [(nivel, fuente, mensaje)]. No imprime ni sale."""
    R = CI.rutas(cfg)
    corte = CI.fkey(f"{cfg.CORTE}/{cfg.ANIO}")
    hallazgos: list[tuple[str, str, str]] = []
    nivel_ausencia = AVISO if permitir_incompleto else FATAL
    nivel_id = FATAL if identidad_fatal else AVISO

    for emp in ("ALTAMIRA", "MERAK"):
        ab = CI.ABREV_EMP[emp]

        # ── BCP histórico: la columna vertebral del cierre ──
        et = f"{ab} · EECC BCP"
        f = CI.first(R["bcp"][emp])
        if not f:
            hallazgos.append((nivel_ausencia, et,
                              f"falta EECC_{ab}_{cfg.PERIODO}_movimientos.* — sin él no hay cierre"))
            continue
        try:
            rep = BCP.parse_movimientos(f)
        except SystemExit as e:
            hallazgos.append((FATAL, et, f"{os.path.basename(f)}: {e}"))
            continue
        identidad_bcp(rep.get("cuenta", ""), emp, nivel_id, hallazgos)
        movs = rep["movimientos"]
        if not movs:
            hallazgos.append((FATAL, et,
                              f"{os.path.basename(f)}: 0 movimientos parseados — formato roto "
                              f"(la clase de bug del .txt con pipe)"))
            continue
        _chequear_periodo([m["fecha"] for m in movs], cfg.PERIODO, et, hallazgos)
        _chequear_exceso_corte([m["fecha"] for m in movs], cfg, corte, et, hallazgos)

        # cobertura: histórico + UltimosMov (misma fusión que hace el pipeline)
        ultimo = max(CI.fkey(m["fecha"]) for m in movs)
        ult = CI.first(R["bcp_ult"][emp])
        if ult:
            try:
                extra = [u for u in CI.leer_ultimos(ult) if CI.fkey(u["fecha"]) > ultimo]
                if extra:
                    ultimo = max(CI.fkey(u["fecha"]) for u in extra)
            except SystemExit as e:
                hallazgos.append((FATAL, f"{ab} · UltimosMov", str(e)))
        if ultimo < corte:
            d, m, a = ultimo[2], ultimo[1], ultimo[0]
            hallazgos.append((nivel_ausencia, et,
                              f"los movimientos llegan al {d}/{m}/{a}, no al {cfg.CORTE} — el BCP "
                              f"publica el último día recién al día siguiente; re-descarga con el "
                              f"mes cerrado (así se perdieron 12 movs y S/3,300 en junio)"))

        # plausibilidad de conteo (SOLO AVISO): un mes con la mitad o el doble de movimientos
        # que el anterior no es necesariamente un error, pero merece un vistazo. Con corte
        # quincenal (--corte 15/07) se PRORRATEA: media quincena se compara contra medio mes.
        prev = _conteo_previo(cfg, emp)
        if prev and prev > 0:
            dd = int(cfg.CORTE[:2])
            fraccion = dd / CI.ultimo_dia(cfg.ANIO, cfg.MES)
            esperado = prev * fraccion
            var = (rep["n"] - esperado) / esperado
            if abs(var) > TOLERANCIA_CONTEO:
                base = (f"{prev} del mes anterior completo" if fraccion >= 0.999 else
                        f"~{esperado:.0f} prorrateados al {cfg.CORTE} ({prev} el mes anterior)")
                hallazgos.append((AVISO, et,
                                  f"{rep['n']} movimientos vs {base} ({var:+.0%}) — fuera de "
                                  f"±{TOLERANCIA_CONTEO:.0%}, verifica que el export esté completo"))

        # ── efact emitidas: si falta, integrar_cobros integra CERO facturas en silencio ──
        et = f"{ab} · efact emitidas"
        f = CI.first(R["efact"][emp])
        if not f:
            hallazgos.append((nivel_ausencia, et,
                              f"falta FACT-EMI_{ab}_{cfg.PERIODO}_Reporte.xlsx — sin él el cierre "
                              f"integra 0 facturas nuevas SIN AVISAR"))
        else:
            fechas, n = _leer_fechas_efact(f)
            if n == 0:
                hallazgos.append((AVISO, et,
                                  f"{os.path.basename(f)}: 0 comprobantes — válido solo si no se "
                                  f"emitió nada en {cfg.MES_NOMBRE}"))
            else:
                _chequear_periodo(fechas, cfg.PERIODO, et, hallazgos)
                _chequear_exceso_corte(fechas, cfg, corte, et, hallazgos)

        # ── BN movimientos: ausente = AVISO (mayo cerró sin él); corrupto = FATAL ──
        et = f"{ab} · BN detracciones"
        f = CI.first(R["bn"][emp])
        bn_file = f
        if not f:
            hallazgos.append((AVISO, et,
                              f"sin BN_{ab}_{cfg.PERIODO}_movimientos.* — las detracciones del mes "
                              f"no se validarán contra la cuenta BN"))
        elif f.lower().endswith((".csv", ".xls", ".xlsx", ".txt")):
            fechas = _leer_fechas_planas(f)
            if not fechas:
                hallazgos.append((FATAL, et,
                                  f"{os.path.basename(f)}: no se pudo leer ninguna fecha — "
                                  f"formato roto"))
            else:
                # el BN arrastra saldos: filas de otros meses al inicio son normales
                _chequear_periodo(fechas, cfg.PERIODO, et, hallazgos, permitir_fuera=True)
                identidad_bn(f, emp, nivel_id, hallazgos)

        # ── Constancias SUNAT: puede legítimamente no haber pagos ──
        et = f"{ab} · constancias SUNAT"
        f = CI.first(R["sunat"][emp])
        if not f:
            n_ab = _abonos_bn(bn_file) if bn_file else 0
            if n_ab:
                hallazgos.append((nivel_ausencia, et,
                                  f"sin DETR-SUNAT_{ab}_{cfg.PERIODO}.csv, pero el extracto del BN "
                                  f"trae {n_ab} abono(s) en {cfg.MES_NOMBRE}: hubo depósitos y no "
                                  f"hay constancias para cruzarlos (descárgalas de SUNAT → "
                                  f"03. DETRACCIONES/{cfg.ANIO}/02. Constancias_SUNAT/)"))
            else:
                hallazgos.append((AVISO, et,
                                  f"sin DETR-SUNAT_{ab}_{cfg.PERIODO}.csv — el BN no muestra abonos "
                                  f"en {cfg.MES_NOMBRE}, así que no haber pagos es coherente"))
        else:
            identidad_sunat(f, emp, nivel_id, hallazgos)

    return hallazgos


def _leer_fechas_efact(path: str) -> tuple[list, int]:
    """Fechas de emisión del reporte efact (lector mínimo, solo para validar)."""
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True, read_only=True).active
    filas = [list(r) for r in ws.iter_rows(values_only=True)]
    hdr = next((i for i, r in enumerate(filas)
                if r and any("Comprobante" in str(c or "") for c in r)), None)
    if hdr is None:
        return [], 0
    iF = next((k for k, h in enumerate(filas[hdr]) if "Emisi" in str(h or "")), None)
    datos = [r for r in filas[hdr + 1:] if r and any(c is not None for c in r)]
    fechas = [r[iF] for r in datos if iF is not None and len(r) > iF and r[iF] is not None]
    return fechas, len(datos)


def _abonos_bn(path: str) -> int:
    """Cuántos ABONOS trae el extracto del BN. 0 si no se puede leer.

    Sirve para una sola decisión: si el BN muestra depósitos del mes, entonces las constancias
    SUNAT NO son opcionales. Antes la compuerta dejaba pasar su ausencia con "válido si no hubo
    pagos" sin mirar si los hubo; el 01/09 agosto tenía 14 abonos entre las dos empresas y cero
    constancias, y eso habría entrado al cierre como si nada.
    """
    try:
        filas = []
        if path.lower().endswith(".xls"):
            import xlrd
            sh = xlrd.open_workbook(path).sheet_by_index(0)
            filas = [[str(sh.cell_value(i, j)).strip() for j in range(sh.ncols)]
                     for i in range(sh.nrows)]
        else:
            import csv
            filas = list(csv.reader(io.open(path, encoding="utf-8-sig", errors="replace")))
        n = 0
        for r in filas:
            if len(r) < 8 or not r[0].strip().rstrip(".").isdigit():
                continue
            v = r[7].replace(",", "").strip()
            try:
                if float(v) != 0:
                    n += 1
            except ValueError:
                pass
        return n
    except Exception:
        return 0


def _leer_fechas_planas(path: str) -> list:
    """Todas las celdas con pinta de fecha de un csv/xls/xlsx (para validar el BN)."""
    celdas = []
    if path.lower().endswith((".xls", ".xlsx")):
        try:
            import openpyxl
            ws = openpyxl.load_workbook(path, data_only=True, read_only=True).active
            for r in ws.iter_rows(values_only=True):
                celdas.extend(r or ())
        except Exception:
            try:
                import xlrd
                wb = xlrd.open_workbook(path)
                sh = wb.sheet_by_index(0)
                celdas = [sh.cell_value(r, c) for r in range(sh.nrows) for c in range(sh.ncols)]
            except Exception:
                return []
    else:
        import csv
        texto = open(path, "rb").read().decode("latin-1", "replace")
        deli = max(",;|\t", key=texto[:600].count)
        for row in csv.reader(texto.splitlines(), delimiter=deli):
            celdas.extend(row)
    return [c for c in celdas if _fecha_mes(c)]


def main():
    # config_minima, no cargar_config: la compuerta corre ANTES de que exista la config del mes
    # (bajas archivos → validas → recién armas tools/config/{mes}.py con SALDO_BN y overrides).
    cfg = CI.config_minima(CI.arg_mes())
    permitir = "--permitir-incompleto" in sys.argv
    id_fatal = "--identidad-fatal" in sys.argv
    print(f"🛡️ COMPUERTA DE INGESTA · {cfg.MES_NOMBRE} {cfg.ANIO} (corte {cfg.CORTE})"
          + ("  · modo provisional" if permitir else "")
          + ("  · identidad FATAL" if id_fatal else ""))
    hallazgos = validar(cfg, permitir, identidad_fatal=id_fatal)
    if not hallazgos:
        print("\n✅ Todas las fuentes válidas: contenido del periodo, completas hasta el corte.")
        return
    fatales = [h for h in hallazgos if h[0] == FATAL]
    print()
    for nivel, fuente, msg in sorted(hallazgos, key=lambda h: h[0] != FATAL):
        print(f"  {nivel}  {fuente}")
        print(f"          {msg}")
    print(f"\n{len(fatales)} fatal(es) · {len(hallazgos) - len(fatales)} aviso(s)")
    if fatales:
        sys.exit("\n❌ Compuerta cerrada: corrige los archivos y vuelve a correr. "
                 "Ningún maestro fue tocado.")
    print("✅ Compuerta abierta con avisos — revísalos antes de confiar en el cierre.")


if __name__ == "__main__":
    main()
