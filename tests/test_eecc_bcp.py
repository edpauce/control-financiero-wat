"""
Lectura del histórico BCP y cobertura del mes.

El export .xlsx trae columnas que el .csv no tiene (Sucursal, Hora, Usuario, UTC), así que las
columnas se resuelven por NOMBRE de encabezado. Y el BCP publica el último día del mes recién al
día siguiente: descargar el histórico el día 30 deja el mes incompleto sin avisar.
"""
import datetime

import openpyxl
import pytest

from lib import cierre as CI
from scrape import bcp_empresas as BCP

CAB_CSV = ["Fecha", "Fecha valuta", "Descripción operación", "Monto", "Saldo",
           "Operación - Número", "Referencia2"]
CAB_XLSX = ["Fecha", "Fecha valuta", "Descripción operación", "Monto", "Saldo",
            "Sucursal - agencia", "Operación - Número", "Operación - Hora", "Usuario",
            "UTC", "Referencia2"]
FILA_CSV = ["30/06/2026", "", "DE KALPA ENGINEERING E", "3300.00", "8100.50", "02826417", "F001-9"]
FILA_XLSX = ["30/06/2026", "", "DE KALPA ENGINEERING E", 3300.0, 8100.50, "193-000",
             "02826417", "20:21:00", "BATCH", "0909", "F001-9"]

ESPERADO = {"fecha": "30/06/2026", "desc": "DE KALPA ENGINEERING E", "monto": 3300.0,
            "saldo": 8100.50, "operacion": "02826417", "referencia": "F001-9"}


def _csv(tmp_path, cab=CAB_CSV, filas=(FILA_CSV,)):
    ruta = tmp_path / "hist.csv"
    lineas = ["Cuenta,193-2000201-0-21 - MERAK", "Moneda,Soles", "", ",".join(cab)]
    lineas += [",".join(str(x) for x in f) for f in filas]
    ruta.write_text("\n".join(lineas), encoding="latin-1")
    return str(ruta)


def _xlsx(tmp_path, cab=CAB_XLSX, filas=(FILA_XLSX,), nombre="hist.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Cuenta", "193-2000201-0-21 - MERAK"])
    ws.append(["Moneda", "Soles"])
    ws.append([])
    ws.append(cab)
    for f in filas:
        ws.append(list(f))
    ruta = tmp_path / nombre
    wb.save(ruta)
    return str(ruta)


def test_csv_y_xlsx_dan_el_mismo_movimiento_pese_a_columnas_distintas(tmp_path):
    """El xlsx mete 'Sucursal - agencia' entre Saldo y Operación: leer por posición desplazaría
    la operación y la referencia."""
    a = BCP.parse_movimientos(_csv(tmp_path))["movimientos"]
    b = BCP.parse_movimientos(_xlsx(tmp_path))["movimientos"]
    assert a == b == [ESPERADO]


def test_xlsx_con_celdas_datetime(tmp_path):
    """Excel guarda las fechas como datetime, no como texto."""
    fila = list(FILA_XLSX)
    fila[0] = datetime.datetime(2026, 6, 30)
    movs = BCP.parse_movimientos(_xlsx(tmp_path, filas=(fila,), nombre="dt.xlsx"))["movimientos"]
    assert movs == [ESPERADO]


def test_cuenta_y_totales(tmp_path):
    rep = BCP.parse_movimientos(_xlsx(tmp_path))
    assert rep["cuenta"].startswith("193-2000201")
    assert (rep["ingresos"], rep["egresos"], rep["n"]) == (3300.0, 0.0, 1)


def test_no_confunde_fecha_valuta_con_fecha(tmp_path):
    rep = BCP.parse_movimientos(_csv(tmp_path))
    assert rep["movimientos"][0]["fecha"] == "30/06/2026"


# ─── Cobertura del mes ──────────────────────────────────────────────────────
class _Cfg:
    PERIODO, ANIO, CORTE = "2026-06", 2026, "30/06"


def test_verificar_cobertura_aborta_si_falta_el_ultimo_dia(monkeypatch):
    """El histórico se bajó el día 30 y llegaba solo al 29.
    Desaparecieron 12 movimientos, uno de ellos un abono de S/3,300."""
    monkeypatch.setattr("sys.argv", ["tool.py", "--mes", "2026-06"])
    movs = [{"fecha": "29/06/2026"}, {"fecha": "01/06/2026"}]
    with pytest.raises(SystemExit) as e:
        CI.verificar_cobertura(movs, _Cfg, "MERAK")
    assert "29/06/2026" in str(e.value) and "30/06" in str(e.value)


def test_verificar_cobertura_pasa_si_llega_al_corte(monkeypatch):
    monkeypatch.setattr("sys.argv", ["tool.py", "--mes", "2026-06"])
    CI.verificar_cobertura([{"fecha": "30/06/2026"}], _Cfg, "MERAK")


def test_verificar_cobertura_se_puede_forzar(monkeypatch):
    monkeypatch.setattr("sys.argv", ["tool.py", "--mes", "2026-06", "--permitir-incompleto"])
    CI.verificar_cobertura([{"fecha": "29/06/2026"}], _Cfg, "MERAK")


def test_verificar_cobertura_aborta_sin_movimientos(monkeypatch):
    monkeypatch.setattr("sys.argv", ["tool.py"])
    with pytest.raises(SystemExit):
        CI.verificar_cobertura([], _Cfg, "ALTAMIRA")


# ─── Delimitador: el BCP exporta .csv con coma y .txt con pipe ──────────────
def test_txt_con_pipe(tmp_path):
    """El export .txt del BCP usa '|'. Con el delimitador fijo en coma, el archivo parseaba a
    cero movimientos: un EECC llevaba meses ilegible."""
    ruta = tmp_path / "hist.txt"
    ruta.write_text(
        '"Cuenta"|"193-1000101-0-11 - ALTAMIRA CONSULTING SAC"|\n'
        '"Moneda"|"Soles"|\n\n'
        '"Fecha"|"Fecha valuta"|"Descripción operación"|"Monto"|"Saldo"|"Operación - Número"|"Referencia2"|\n'
        '"31/05/2026"|""|"IMPUESTO ITF"|"-0.10"|"2,500.00"|"00000000"|"Glosa libre"|\n',
        encoding="latin-1")
    rep = BCP.parse_movimientos(str(ruta))
    assert rep["n"] == 1
    assert rep["movimientos"][0] == {"fecha": "31/05/2026", "desc": "IMPUESTO ITF", "monto": -0.10,
                                     "saldo": 2500.0, "operacion": "00000000",
                                     "referencia": "Glosa libre"}


def test_delimitador_se_detecta_en_el_encabezado():
    assert BCP._delimitador('"Cuenta"|"x"\n"Fecha"|"Monto"|"Saldo"|') == "|"
    assert BCP._delimitador("Cuenta,x\nFecha,Monto,Saldo,") == ","
