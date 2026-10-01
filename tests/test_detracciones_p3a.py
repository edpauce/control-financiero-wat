"""
P3a — Claves de match en detracciones: constancia, RUC y fecha dejan de tirarse a la basura,
y el join se vuelve de DOBLE VÍA: una constancia cuyo comprobante no casa con ninguna factura
(un cliente consignó E001-155 en vez de E001-1552) ya no desaparece en silencio — sale como
huérfana con sus claves, la materia prima del fallback de P3b.
"""
import importlib.util
import os

import pytest

from lib import cierre as CI

_ruta = os.path.join(os.path.dirname(__file__), "..", "tools", "detalle_detracciones.py")
_spec = importlib.util.spec_from_file_location("detalle_detracciones", _ruta)
DD = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(DD)

CAB_SUNAT = ("Tipo de Cuenta,Numero de Cuenta,Numero Constancia,Periodo Tributario,"
             "RUC Proveedor,Nombre Proveedor,Tipo de Documento Adquiriente,"
             "Numero de Documento Adquiriente,Nombre/Razon Social del Adquiriente,"
             "Fecha Pago,Monto Deposito,Tipo Bien,Tipo Operacion,Tipo de Comprobante,"
             "Serie de Comprobante,Numero de Comprobante,Numero de pago de Detracciones,")


def _csv_sunat(tmp_path, filas):
    d = tmp_path / "02. MERAK" / "03. DETRACCIONES" / "2026" / "02. Constancias_SUNAT"
    d.mkdir(parents=True, exist_ok=True)
    (d / "DETR-SUNAT_MER_2026-06.csv").write_text(
        CAB_SUNAT + "\n" + "\n".join(filas), encoding="latin-1")


def _csv_bn(tmp_path, filas):
    d = tmp_path / "02. MERAK" / "03. DETRACCIONES" / "2026" / "01. EECC_BN"
    d.mkdir(parents=True, exist_ok=True)
    (d / "BN_MER_2026-06_movimientos.csv").write_text(
        "Cuenta,00-000-200201 - MERAK\nPeriodo,2026-06\n\n"
        "Nro,Fecha,Trans,Documento,RUC,Oficina,Cargo,Abono\n" + "\n".join(filas),
        encoding="latin-1")


# Una constancia que consigna E001-155 cuando la factura es E001-1552.
MAL_CONSIGNADA = ("Cuenta de Detracciones Convencional,00000200201,300000011,202606,20900000021,"
             "MERAK CERTIFICACIONES S.A.C. - MERA,06,20900000055,CLIENTE DE EJEMPLO "
             "UNO S.A.,23/06/2026,1998.00,022,01,01,E001,00000155,,")
RADITEL = ("Cuenta de Detracciones Convencional,00000200201,300000022,202606,20900000021,"
          "MERAK CERTIFICACIONES S.A.C. - MERA,06,20900000068,RADITEL PERU S.A.,"
          "25/06/2026,456.00,037,01,01,F001,00001069,,")


@pytest.fixture
def root(tmp_path, monkeypatch):
    monkeypatch.setattr(CI, "ROOT", str(tmp_path))
    monkeypatch.setattr(DD.CI, "ROOT", str(tmp_path))
    return tmp_path


def test_constancias_conservan_todas_las_claves(root):
    _csv_sunat(root, [MAL_CONSIGNADA, RADITEL])
    por_comp, todas = DD.constancias_sunat(CI.config_minima("2026-06"), "MERAK")
    assert len(todas) == 2
    c = por_comp["E001-155"][0]
    assert c["constancia"] == "300000011"
    assert c["ruc_adq"] == "20900000055"          # el adquiriente
    assert c["fecha"] == "23/06/2026"
    assert c["monto"] == 1998.0
    assert c["periodo"] == "202606"


def test_el_join_inverso_saca_la_huerfana(root):
    """El comprobante consignado (E001-155) no casa con la factura real (E001-1552): la
    constancia debe salir como huérfana con sus claves, NUNCA asignarse sola (eso es P3b,
    y solo como REVISAR)."""
    _csv_sunat(root, [MAL_CONSIGNADA, RADITEL])
    por_comp, todas = DD.constancias_sunat(CI.config_minima("2026-06"), "MERAK")
    facturas_del_mes = {"E001-1552", "F001-1069"}     # lo que emitió el mes
    usadas = {id(c) for comp in facturas_del_mes for c in por_comp.get(comp, [])}
    huerfanas = [c for c in todas if id(c) not in usadas]
    assert len(huerfanas) == 1
    assert huerfanas[0]["constancia"] == "300000011"
    assert huerfanas[0]["ruc_adq"] == "20900000055"    # claves listas para P3b


def test_compatibilidad_sunat_declaradas(root):
    _csv_sunat(root, [MAL_CONSIGNADA, RADITEL])
    d = DD.sunat_declaradas(CI.config_minima("2026-06"), "MERAK")
    assert d == {"E001-155": 1998.0, "F001-1069": 456.0}


# ─── Candado hoja 09: un PENDIENTE con constancia existente es un falso pendiente ───
YA_DEPOSITADA = ("Cuenta de Detracciones Convencional,00000200201,300000033,202606,20900000021,"
             "MERAK CERTIFICACIONES S.A.C. - MERA,06,20900000073,SERVICIOS INTEGRALES Y "
             "DIGITALES,24/06/2026,960.00,037,01,01,F001,00001071,,")


def _wb_hoja09(filas):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "09 Pendientes Auto-Det"
    ws.append(["Empresa", "Mes", "Fecha", "Cliente", "Comprobante", "Monto", "Det", "Estado"])
    for f in filas:
        ws.append(f)
    return wb


def test_candado_hoja09_detecta_pendiente_con_constancia(root):
    """F001-1071 figura en hoja 09 como PENDIENTE, pero el cliente ya la había depositado
    el 24/06 (constancia 300000033). El candado la delata; no cambia el estado."""
    _csv_sunat(root, [YA_DEPOSITADA])
    wb = _wb_hoja09([
        ["MERAK", "Jun 2026", "23/06/2026", "TECNOSER", "F001-1071", 8000, 960, "PENDIENTE"],
        ["MERAK", "Mar 2026", "23/03/2026", "TECOMSA", "F001-1017", 9000, 1080, "PENDIENTE"],
    ])
    avisos = DD.candado_hoja09(wb, CI.config_minima("2026-06"))
    assert len(avisos) == 1
    assert "F001-1071" in avisos[0] and "300000033" in avisos[0]
    assert wb["09 Pendientes Auto-Det"].cell(2, 8).value == "PENDIENTE"   # solo avisa


def test_candado_hoja09_ignora_depositadas_y_validar(root):
    _csv_sunat(root, [YA_DEPOSITADA])
    wb = _wb_hoja09([
        ["MERAK", "Jun 2026", "23/06/2026", "TECNOSER", "F001-1071", 8000, 960,
         "DEPOSITADO (cliente, 24/06)"],
        ["MERAK", "May 2026", "23/05/2026", "J&R ROJAS", "F001-1054", 9000, 1080,
         "VALIDAR CONTADOR"],
    ])
    assert DD.candado_hoja09(wb, CI.config_minima("2026-06")) == []


def test_bn_abonos_lee_el_formato_real(root):
    _csv_bn(root, ["01,2026.06.24,NOT.ABON,10000001,20900000055,1234,,1998.00",
                   "02,2026.06.20,NOT.CARG,1612,,1234,960.00,",      # cargo: se ignora
                   "03,2026.06.26,NOT.ABON,10000002,,1234,,300.00"])  # RUC vacío: válido
    ab = DD.bn_abonos(CI.config_minima("2026-06"), "MERAK")
    assert len(ab) == 2
    assert ab[0] == {"fecha": "2026.06.24", "ruc": "20900000055",
                     "doc": "10000001", "monto": 1998.0}


def test_bn_confirma_por_ruc_y_monto():
    c = {"ruc_adq": "20900000055", "monto": 1998.0}
    abonos = [{"fecha": "2026.06.24", "ruc": "20900000055", "doc": "x", "monto": 1998.0}]
    assert DD._bn_confirma(c, abonos)
    assert not DD._bn_confirma({"ruc_adq": "20999999999", "monto": 1998.0}, abonos)
    assert not DD._bn_confirma({"ruc_adq": "20900000055", "monto": 1997.0}, abonos)
    # RUC vacío en el BN: el monto exacto basta (el BN a veces no lo trae)
    assert DD._bn_confirma(c, [{"fecha": "x", "ruc": "", "doc": "x", "monto": 1998.0}])


def test_bn_ausente_devuelve_vacio_no_revienta(root):
    assert DD.bn_abonos(CI.config_minima("2026-05"), "MERAK") == []


# ─── P3b: fallback RUC+monto+fecha — propone, NUNCA cierra ──────────────────
def _fila(comp, esp, dep=0.0):
    return ["202606", "18/06/2026", comp, "CLIENTE", esp / 0.12, esp, dep, "", "neto",
            "🟡 Cliente deposita", "pagó neto"]


def _meta(ruc, esp, fcobro, dep=0.0):
    return {"ruc": ruc, "esp": esp, "dep": dep, "fcobro": fcobro}


def _h(constancia, comp, ruc, monto, fecha):
    return {"constancia": constancia, "comp": comp, "ruc_adq": ruc, "adquiriente": "X",
            "fecha": fecha, "monto": monto, "periodo": "202606"}


def test_candidata_unica_marca_revisar_sin_cerrar():
    """La constancia consigna E001-155; la factura real es E001-1552. RUC+monto(redondeo
    SUNAT)+fecha de cobro cuadran → 🟠 propuesta visible, pero Detr Depositada NO se toca."""
    filas = [_fila("E001-1552", 1998.28)]
    meta = [_meta("20900000055", 1998.28, "23/06/2026")]
    hs = [_h("300000011", "E001-155", "20900000055", 1998.00, "23/06/2026")]
    filas, hs = DD.fallback_p3b(filas, meta, hs)
    assert filas[0][9] == "🟠 Constancia por confirmar"
    assert "300000011" in filas[0][7] and filas[0][7].startswith("¿")
    assert filas[0][6] == 0.0                      # NUNCA auto-cierra
    assert "candidata única de E001-1552" in hs[0]["p3b"]


def test_dos_facturas_candidatas_es_ambiguo_y_no_propone_nada():
    filas = [_fila("F001-1", 500.0), _fila("F001-2", 500.0)]
    meta = [_meta("20111111111", 500.0, "20/06/2026"), _meta("20111111111", 500.0, "21/06/2026")]
    hs = [_h("300000001", "F001-9", "20111111111", 500.00, "20/06/2026")]
    filas, hs = DD.fallback_p3b(filas, meta, hs)
    assert "AMBIGUA (2 factura(s)" in hs[0]["p3b"]
    for f in filas:
        assert f[9] != "🟠 Constancia por confirmar"   # nada propuesto
        assert f[7] == ""


def test_fecha_fuera_de_ventana_no_casa():
    filas = [_fila("F001-1", 500.0)]
    meta = [_meta("20111111111", 500.0, "10/06/2026")]
    hs = [_h("300000001", "F001-9", "20111111111", 500.00, "20/06/2026")]   # 10 días
    filas, hs = DD.fallback_p3b(filas, meta, hs)
    assert "p3b" not in hs[0] and filas[0][9] == "🟡 Cliente deposita"


def test_sin_fecha_de_cobro_no_casa():
    """Conservador: sin ancla temporal no hay fallback (el plan exige las TRES claves)."""
    filas = [_fila("F001-1", 500.0)]
    meta = [_meta("20111111111", 500.0, None)]
    hs = [_h("300000001", "F001-9", "20111111111", 500.00, "20/06/2026")]
    filas, hs = DD.fallback_p3b(filas, meta, hs)
    assert "p3b" not in hs[0]


def test_ruc_distinto_no_casa():
    filas = [_fila("F001-1", 500.0)]
    meta = [_meta("20999999999", 500.0, "20/06/2026")]
    hs = [_h("300000001", "F001-9", "20111111111", 500.00, "20/06/2026")]
    filas, hs = DD.fallback_p3b(filas, meta, hs)
    assert "p3b" not in hs[0]


def test_tolerancia_es_redondeo_sunat_no_mas():
    """±1.00 cubre el redondeo al sol entero; S/2 de diferencia ya no casa."""
    filas = [_fila("F001-1", 500.75)]
    meta = [_meta("20111111111", 500.75, "20/06/2026")]
    hs = [_h("3001", "F001-9", "20111111111", 502.75, "20/06/2026")]
    filas, hs = DD.fallback_p3b(filas, meta, hs)
    assert "p3b" not in hs[0]


def test_factura_con_constancia_exacta_no_participa_del_fallback():
    """dep>0 = ya amarrada por el join exacto: el fallback no debe proponerle nada más."""
    filas = [_fila("F001-1", 500.0, dep=500.0)]
    meta = [_meta("20111111111", 500.0, "20/06/2026", dep=500.0)]
    hs = [_h("300000001", "F001-9", "20111111111", 500.00, "20/06/2026")]
    filas, hs = DD.fallback_p3b(filas, meta, hs)
    assert "p3b" not in hs[0]
