"""
P1a — Compuerta de ingesta. Cada test reproduce una falla silenciosa:
el histórico corto (faltaba el último día), el archivo cuyo nombre miente (BN 'marzo'
que contenía enero), y el formato roto que parseaba a 0 movimientos (el .txt con pipe).
La compuerta convierte cada una en falla ruidosa ANTES de tocar un maestro.
"""
import importlib.util
import os

import openpyxl
import pytest

from lib import cierre as CI

_ruta = os.path.join(os.path.dirname(__file__), "..", "tools", "validar_ingesta.py")
_spec = importlib.util.spec_from_file_location("validar_ingesta", _ruta)
VI = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(VI)


# ─── Fábrica de un ROOT sintético con fuentes de un mes ─────────────────────
CAB = ["Fecha", "Fecha valuta", "Descripción operación", "Monto", "Saldo",
       "Operación - Número", "Referencia2"]


CUENTAS = {"ALT": "193-1000101-0-11 - ALTAMIRA CONSULTING",
           "MER": "193-2000201-0-21 - MERAK CERTIFICACIONES"}


def _csv_bcp(ruta, filas, cuenta="193-2000201-0-21 - MERAK"):
    lineas = [f'"Cuenta","{cuenta}"', '"Moneda","Soles"', "",
              ",".join(f'"{c}"' for c in CAB)]
    lineas += [",".join(f'"{x}"' for x in f) for f in filas]
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    with open(ruta, "w", encoding="latin-1") as fh:
        fh.write("\n".join(lineas))


def _efact(ruta, fechas):
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Razón Social", "RUC", "Tipo Doc.", "N° Comprobante", "F. Emisión",
               "F. Vencimiento", "Total en Soles"])
    for i, f in enumerate(fechas, 1):
        ws.append([f"CLIENTE {i} S.A.C.", "20100000001", "Factura", f"F001-{1000+i}",
                   f, "", 1000.0])
    wb.save(ruta)


def _fila(fecha, monto, saldo):
    return [fecha, "", "DESC", f"{monto:.2f}", f"{saldo:.2f}", "00000001", ""]


def _mes_completo(root, periodo="2026-06", ult_dia="30/06/2026"):
    """Fuentes válidas del mes para ambas empresas."""
    a = periodo[:4]
    for d, ab in (("01. ALTAMIRA", "ALT"), ("02. MERAK", "MER")):
        _csv_bcp(os.path.join(root, d, "01. EECC", a, f"EECC_{ab}_{periodo}_movimientos.csv"),
                 [_fila(ult_dia, -1.0, 99.0), _fila(f"01/{periodo[5:]}/{a}", 100.0, 100.0)],
                 cuenta=CUENTAS[ab])
        _efact(os.path.join(root, d, "02. FACTURACION", a, "Emitidas",
                            f"FACT-EMI_{ab}_{periodo}_Reporte.xlsx"),
               [f"15/{periodo[5:]}/{a}"])


@pytest.fixture
def root(tmp_path, monkeypatch):
    monkeypatch.setattr(CI, "ROOT", str(tmp_path))
    monkeypatch.setattr(VI.CI, "ROOT", str(tmp_path))
    return str(tmp_path)


def _fatales(hallazgos):
    return [h for h in hallazgos if h[0] == VI.FATAL]


# ─── Camino feliz ────────────────────────────────────────────────────────────
def test_mes_completo_abre_la_compuerta(root):
    _mes_completo(root)
    h = VI.validar(CI.config_minima("2026-06"))
    assert not _fatales(h)
    # BN y SUNAT ausentes son avisos, nunca fatales (un mes puede cerrar así)
    assert all("BN" in f or "SUNAT" in f for _, f, _ in h)


# ─── Clase de bug 1: el histórico corto (faltaba el último día) ────
def test_cobertura_corta_cierra_la_compuerta(root):
    _mes_completo(root, ult_dia="29/06/2026")     # llega al 29, corte es 30
    h = VI.validar(CI.config_minima("2026-06"))
    cortos = [m for n, f, m in _fatales(h) if "29/06/2026" in m and "EECC BCP" in f]
    assert len(cortos) == 2                        # ambas empresas


def test_permitir_incompleto_baja_cobertura_a_aviso(root):
    _mes_completo(root, ult_dia="29/06/2026")
    h = VI.validar(CI.config_minima("2026-06"), permitir_incompleto=True)
    assert not _fatales(h)
    assert any("29/06/2026" in m for _, _, m in h)


# ─── Clase de bug 2: el nombre miente (BN 'marzo' contenía enero) ───────────
def test_contenido_de_otro_mes_cierra_la_compuerta(root):
    _mes_completo(root)
    # el archivo se llama junio pero su contenido es ENERO
    _csv_bcp(os.path.join(root, "02. MERAK", "01. EECC", "2026",
                          "EECC_MER_2026-06_movimientos.csv"),
             [_fila("19/01/2026", -1.0, 99.0), _fila("05/01/2026", 100.0, 100.0)])
    h = VI.validar(CI.config_minima("2026-06"))
    mentiras = [m for n, f, m in _fatales(h) if "miente" in m]
    assert len(mentiras) == 1
    assert "2026-01" in mentiras[0]


def test_el_nombre_miente_es_fatal_incluso_en_modo_provisional(root):
    """Corrupción nunca se degrada: un mes equivocado no es un cierre provisional."""
    _mes_completo(root)
    _csv_bcp(os.path.join(root, "02. MERAK", "01. EECC", "2026",
                          "EECC_MER_2026-06_movimientos.csv"),
             [_fila("19/01/2026", -1.0, 99.0)])
    h = VI.validar(CI.config_minima("2026-06"), permitir_incompleto=True)
    assert any("miente" in m for _, _, m in _fatales(h))


# ─── Clase de bug 3: formato roto → 0 movimientos (el .txt con pipe) ────────
def test_cero_movimientos_parseados_cierra_la_compuerta(root):
    _mes_completo(root)
    ruta = os.path.join(root, "02. MERAK", "01. EECC", "2026",
                        "EECC_MER_2026-06_movimientos.csv")
    with open(ruta, "w", encoding="latin-1") as fh:      # encabezado sin filas legibles
        fh.write('"Cuenta","193-2000201-0-21 - MERAK"\n"Fecha","Monto"\n"BASURA","x"\n')
    h = VI.validar(CI.config_minima("2026-06"))
    assert any("0 movimientos" in m or "formato roto" in m for _, _, m in _fatales(h))


# ─── Ausencias ───────────────────────────────────────────────────────────────
def test_falta_efact_es_fatal_porque_integraria_cero_facturas_en_silencio(root):
    _mes_completo(root)
    os.remove(os.path.join(root, "02. MERAK", "02. FACTURACION", "2026", "Emitidas",
                           "FACT-EMI_MER_2026-06_Reporte.xlsx"))
    h = VI.validar(CI.config_minima("2026-06"))
    assert any("efact" in f for _, f, _ in _fatales(h))


def test_falta_bcp_es_fatal(root):
    _mes_completo(root)
    os.remove(os.path.join(root, "01. ALTAMIRA", "01. EECC", "2026",
                           "EECC_ALT_2026-06_movimientos.csv"))
    h = VI.validar(CI.config_minima("2026-06"))
    assert any("no hay cierre" in m for _, _, m in _fatales(h))


# ─── Plausibilidad de conteo: SOLO aviso, nunca fatal ────────────────────────
def test_conteo_fuera_de_tolerancia_es_aviso_no_fatal(root):
    _mes_completo(root, periodo="2026-06")
    # mes anterior con 10 movimientos; junio tiene 2 → −80%, fuera de ±40%
    _csv_bcp(os.path.join(root, "02. MERAK", "01. EECC", "2026",
                          "EECC_MER_2026-05_movimientos.csv"),
             [_fila(f"{d:02d}/05/2026", 10.0, 10.0 * d) for d in range(1, 11)])
    h = VI.validar(CI.config_minima("2026-06"))
    avisos = [m for n, f, m in h if n == VI.AVISO and "mes anterior" in m]
    assert avisos and "MER" not in str(_fatales(h))


# ─── config_minima: el calendario sin exigir config del mes ─────────────────
def test_config_minima_deriva_el_calendario():
    c = CI.config_minima("2026-01")
    assert (c.PREV_PERIODO, c.PREV_NOMBRE, c.CORTE) == ("2025-12", "Diciembre", "31/01")


def test_config_minima_rechaza_periodos_invalidos():
    with pytest.raises(SystemExit):
        CI.config_minima("junio")
    with pytest.raises(SystemExit):
        CI.config_minima("2026-13")


# ─── P1b: heurística de identidad (warn-only el primer ciclo) ────────────────
def _csv_bcp_cuenta(ruta, cuenta, filas):
    """Como _csv_bcp pero con la cuenta declarada configurable."""
    lineas = [f'"Cuenta","{cuenta}"', '"Moneda","Soles"', "",
              ",".join(f'"{c}"' for c in CAB)]
    lineas += [",".join(f'"{x}"' for x in f) for f in filas]
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    with open(ruta, "w", encoding="latin-1") as fh:
        fh.write("\n".join(lineas))


def test_eecc_de_la_otra_empresa_se_detecta_como_aviso(root):
    """El EECC de MERAK guardado en la carpeta de ALTAMIRA: el contenido declara la cuenta
    193-2000201 y eso delata el error de archivado. AVISO (warn-only), no fatal."""
    _mes_completo(root)
    _csv_bcp_cuenta(os.path.join(root, "01. ALTAMIRA", "01. EECC", "2026",
                                 "EECC_ALT_2026-06_movimientos.csv"),
                    "193-2000201-0-21 - MERAK CERTIFICACIONES",
                    [_fila("30/06/2026", -1.0, 99.0), _fila("01/06/2026", 100.0, 100.0)])
    h = VI.validar(CI.config_minima("2026-06"))
    ids = [m for n, f, m in h if "identidad" in m and n == VI.AVISO]
    assert any("MERAK" in m and "ALTAMIRA" in m for m in ids)
    assert not _fatales(h)                       # warn-only: la compuerta sigue abierta


def test_cuenta_desconocida_tambien_avisa(root):
    _mes_completo(root)
    _csv_bcp_cuenta(os.path.join(root, "02. MERAK", "01. EECC", "2026",
                                 "EECC_MER_2026-06_movimientos.csv"),
                    "191-9999999-0-55 - OTRA EMPRESA SAC",
                    [_fila("30/06/2026", -1.0, 99.0)])
    h = VI.validar(CI.config_minima("2026-06"))
    assert any("ninguna empresa del grupo" in m for _, _, m in h)


def test_identidad_fatal_promueve_a_fatal(root):
    """Tras calibrar un ciclo, --identidad-fatal convierte el aviso en compuerta cerrada."""
    _mes_completo(root)
    _csv_bcp_cuenta(os.path.join(root, "01. ALTAMIRA", "01. EECC", "2026",
                                 "EECC_ALT_2026-06_movimientos.csv"),
                    "193-2000201-0-21 - MERAK",
                    [_fila("30/06/2026", -1.0, 99.0)])
    h = VI.validar(CI.config_minima("2026-06"), identidad_fatal=True)
    assert any("identidad" in m for _, _, m in _fatales(h))


def test_bn_de_la_otra_empresa_avisa(root):
    _mes_completo(root)
    d = os.path.join(root, "02. MERAK", "03. DETRACCIONES", "2026", "01. EECC_BN")
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, "BN_MER_2026-06_movimientos.csv"), "w", encoding="latin-1") as fh:
        fh.write("Cuenta,00-000-100101 - ALTAMIRA\nPeriodo,2026-06\n\n"
                 "Nro,Fecha,Trans,Documento,RUC,Oficina,Cargo,Abono\n"
                 "01,2026.06.24,NOT.ABON,1,20900000055,1234,,100.00\n")
    h = VI.validar(CI.config_minima("2026-06"))
    assert any("cuenta BN declarada" in m and "ALTAMIRA" in m for _, _, m in h)


def test_sunat_con_ruc_ajeno_avisa(root):
    """Constancias cuyo RUC de proveedor no es el de la empresa: o el CSV es de la otra
    empresa, o SUNAT devolvió algo raro. Aviso con el dueño real si se reconoce."""
    _mes_completo(root)
    d = os.path.join(root, "02. MERAK", "03. DETRACCIONES", "2026", "02. Constancias_SUNAT")
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, "DETR-SUNAT_MER_2026-06.csv"), "w", encoding="latin-1") as fh:
        fh.write("Numero Constancia,RUC Proveedor,Monto Deposito,Fecha Pago\n"
                 "300000001,20900000011,100.00,15/06/2026\n")     # RUC de ALTAMIRA
    h = VI.validar(CI.config_minima("2026-06"))
    assert any("RUC de proveedor ajeno" in m and "ALTAMIRA" in m for _, _, m in h)


def test_identidad_correcta_no_avisa(root):
    """Los archivos bien archivados no ganan ningún aviso de identidad (regresión sintética
    del criterio del plan: cero falsos rechazos)."""
    _mes_completo(root)
    h = VI.validar(CI.config_minima("2026-06"))
    assert not [m for _, _, m in h if "identidad" in m]


# ─── Corte quincenal: datos POSTERIORES al corte declarado son FATAL ─────────
# El pipeline NO recorta al corte: si el EECC bajado el día 20 llega al 20, el "cierre al 15"
# integraría la segunda quincena en silencio. La compuerta lo convierte en falla ruidosa.
def _cfg_quincenal(periodo="2026-06"):
    cfg = CI.config_minima(periodo)
    cfg.CORTE = f"15/{periodo[5:]}"
    return cfg


def test_quincenal_con_movimientos_despues_del_corte_cierra_la_compuerta(root):
    _mes_completo(root)                            # los EECC llegan al 30/06, corte es 15/06
    h = VI.validar(_cfg_quincenal())
    exceso = [m for n, f, m in _fatales(h) if "POSTERIORES al corte" in m and "EECC BCP" in f]
    assert len(exceso) == 2                        # ambas empresas
    assert "re-descarga" in exceso[0]


def test_quincenal_exceso_es_fatal_incluso_en_modo_provisional(root):
    """--permitir-incompleto perdona lo que FALTA, nunca lo que SOBRA (corrupción de rango)."""
    _mes_completo(root)
    h = VI.validar(_cfg_quincenal(), permitir_incompleto=True)
    assert any("POSTERIORES al corte" in m for _, _, m in _fatales(h))


def test_quincenal_con_rango_exacto_abre_la_compuerta(root):
    _mes_completo(root, ult_dia="15/06/2026")      # descarga 01–15: ni falta ni sobra
    h = VI.validar(_cfg_quincenal())
    assert not _fatales(h)


def test_quincenal_efact_con_factura_posterior_al_corte_es_fatal(root):
    _mes_completo(root, ult_dia="15/06/2026")
    _efact(os.path.join(root, "02. MERAK", "02. FACTURACION", "2026", "Emitidas",
                        "FACT-EMI_MER_2026-06_Reporte.xlsx"),
           ["10/06/2026", "20/06/2026"])           # la segunda se emitió tras el corte
    h = VI.validar(_cfg_quincenal())
    assert any("POSTERIORES al corte" in m and "efact" in f for _, f, m in _fatales(h))


def test_corte_fin_de_mes_no_aplica_el_chequeo_de_exceso(root):
    """Sin --corte (o con corte = último día) el chequeo ni se evalúa: cero falsos rechazos
    en los cierres mensuales normales."""
    _mes_completo(root)
    h = VI.validar(CI.config_minima("2026-06"))
    assert not any("POSTERIORES al corte" in m for _, _, m in h)
