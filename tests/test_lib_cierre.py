"""
Tests de las funciones puras de lib/ — las que deciden montos, fechas y matching de facturas.
Ninguno toca los Excel reales del usuario: todo es sintético o en tmp_path.
"""
import datetime

import openpyxl
import pytest

from lib import cierre as CI
from lib import empresas as E


# ─── Umbral SPOT (S/700) ─────────────────────────────────────────────────────
@pytest.mark.parametrize("total,esperado", [
    (0, 0.0),
    (699.99, 0.0),
    (700, 0.0),        # exactamente en el umbral NO se detrae
    (700.01, 84.0),
    (1000, 120.0),
])
def test_detr_esperada_umbral(total, esperado):
    assert E.detr_esperada(total) == esperado


# ─── Normalización de comprobantes ───────────────────────────────────────────
@pytest.mark.parametrize("crudo,esperado", [
    ("F001-00001067", "F001-1067"),
    ("F001-1067", "F001-1067"),
    ("E001-138", "E001-138"),
    ("PAGO 20601234567 F001-9", "F001-9"),
    (None, None),
    ("texto sin comprobante", None),
])
def test_comp_norm(crudo, esperado):
    assert CI.comp_norm(crudo) == esperado


# ─── Periodo previo (borde de año) ───────────────────────────────────────────
@pytest.mark.parametrize("periodo,prev_periodo,prev_nombre", [
    ("2026-01", "2025-12", "Diciembre"),
    ("2026-07", "2026-06", "Junio"),
    ("2026-12", "2026-11", "Noviembre"),
])
def test_cargar_config_periodo_previo(tmp_path, monkeypatch, periodo, prev_periodo, prev_nombre):
    monkeypatch.setattr(CI, "CONFIG_DIR", str(tmp_path))
    (tmp_path / f"{periodo}.py").write_text("")
    cfg = CI.cargar_config(periodo)
    assert (cfg.PREV_PERIODO, cfg.PREV_NOMBRE) == (prev_periodo, prev_nombre)


def test_cargar_config_periodo_invalido_sale():
    with pytest.raises(SystemExit):
        CI.cargar_config("junio-2026")


def test_cargar_config_sin_archivo_sale(tmp_path, monkeypatch):
    monkeypatch.setattr(CI, "CONFIG_DIR", str(tmp_path))
    with pytest.raises(SystemExit):
        CI.cargar_config("2099-01")


def test_fecha_aging_deriva_del_ultimo_dia(tmp_path, monkeypatch):
    monkeypatch.setattr(CI, "CONFIG_DIR", str(tmp_path))
    (tmp_path / "2026-02.py").write_text("")
    assert CI.cargar_config("2026-02").FECHA_AGING == datetime.date(2026, 2, 28)


# ─── Orden de fechas: la clave que fusiona histórico + últimos movimientos ───
def test_fkey_ordena_por_fecha_no_lexicograficamente():
    # Sin normalizar a 2 dígitos, ('2026','6','1') > ('2026','06','30') y el 1/jun se trataría
    # como más nuevo que el 30/jun → doble conteo de cobros.
    assert CI.fkey("1/6/2026") < CI.fkey("30/06/2026")
    assert CI.fkey("01/06/2026") == CI.fkey("1/6/2026")
    assert CI.fkey("31/12/2025") < CI.fkey("01/01/2026")


def test_periodo_de():
    assert CI.periodo_de(2026, "Junio") == "2026-06"
    assert CI.periodo_de(2027, "Enero") == "2027-01"


# ─── Ajuste de flujo: lleva el año en la llave ───────────────────────────────
def test_ajuste_flujo_no_se_reaplica_en_otro_anio():
    assert CI.AJUSTE_FLUJO.get(("MERAK", "2026-06")) == 50000.0
    assert CI.AJUSTE_FLUJO.get(("MERAK", "2027-06")) is None


# ─── leer_ultimos: csv, xlsx-texto, xlsx-datetime, layout roto ───────────────
CABECERA = ["Fecha", "Fecha valuta", "Descripción", "Monto", "Sucursal", "Nº operación"]
FILA = ["30/06/2026", "30/06/2026", "DE KALPA ENGINEERING", 3300.0, "", "01482933"]


def _xlsx(tmp_path, filas, nombre="ult.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(CABECERA)
    for f in filas:
        ws.append(f)
    ruta = tmp_path / nombre
    wb.save(ruta)
    return str(ruta)


def test_leer_ultimos_csv(tmp_path):
    ruta = tmp_path / "ult.csv"
    ruta.write_text(",".join(CABECERA) + "\n" + ",".join(str(x) for x in FILA) + "\n",
                    encoding="latin-1")
    assert CI.leer_ultimos(str(ruta)) == [
        {"fecha": "30/06/2026", "desc": "DE KALPA ENGINEERING",
         "monto": 3300.0, "operacion": "01482933"}]


def test_leer_ultimos_xlsx_fecha_texto(tmp_path):
    movs = CI.leer_ultimos(_xlsx(tmp_path, [FILA]))
    assert movs == [{"fecha": "30/06/2026", "desc": "DE KALPA ENGINEERING",
                     "monto": 3300.0, "operacion": "01482933"}]


def test_leer_ultimos_xlsx_fecha_datetime(tmp_path):
    """Excel guarda las fechas como datetime. Devolver [] en silencio borraría los últimos
    días del mes del cierre (abonos incluidos)."""
    fila = [datetime.datetime(2026, 6, 30), datetime.datetime(2026, 6, 30),
            "DE KALPA ENGINEERING", 3300.0, "", "01482933"]
    movs = CI.leer_ultimos(_xlsx(tmp_path, [fila], "dt.xlsx"))
    assert len(movs) == 1
    assert movs[0]["fecha"] == "30/06/2026"
    assert movs[0]["monto"] == 3300.0


def test_leer_ultimos_grita_si_hay_filas_pero_no_parsea(tmp_path):
    ruta = _xlsx(tmp_path, [["BASURA", "", "", "", "", ""]], "roto.xlsx")
    with pytest.raises(SystemExit):
        CI.leer_ultimos(ruta)


def test_leer_ultimos_archivo_vacio_no_grita(tmp_path):
    assert CI.leer_ultimos(_xlsx(tmp_path, [], "vacio.xlsx")) == []


# ─── arg_mes ─────────────────────────────────────────────────────────────────
def test_arg_mes_sin_valor_sale(monkeypatch):
    monkeypatch.setattr("sys.argv", ["tool.py", "--mes"])
    with pytest.raises(SystemExit):
        CI.arg_mes()


def test_arg_mes_lee_el_valor(monkeypatch):
    monkeypatch.setattr("sys.argv", ["tool.py", "--mes", "2026-07"])
    assert CI.arg_mes() == "2026-07"
