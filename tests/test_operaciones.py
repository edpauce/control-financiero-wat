"""La regla del número de operación: unir las dos patas de un traspaso interno.

Lo que estos tests sostienen, en orden de lo que costaría equivocarse:

1. **Nunca sumar monedas distintas.** El par …307144 saca S/55.00 y mete US$15.97. Si alguien
   suma las patas, el informe reporta una brecha falsa de 39.03 cada vez que haya un traspaso a
   dólares — y una brecha falsa recurrente es exactamente lo que enseña a ignorar el informe.
2. **No cerrar por sufijo a secas.** Seis dígitos colisionan. Un par exige cuentas distintas y
   cercanía en el tiempo, y lo que no pasa se DEVUELVE como descartado, no se tira.
3. **Los tres prefijos son el mismo número.** `01…` carga, `02…` abona, `00…` la constancia.
"""
import datetime as dt
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))

from lib import operaciones as OP          # noqa: E402
from lib import empresas as EMP            # noqa: E402

F = dt.date(2026, 9, 3)


def mov(cuenta, monto, op, moneda="PEN", empresa="MERAK", fecha=F, glosa="DE OTRA CUENTA"):
    return {"cuenta": cuenta, "empresa": empresa, "moneda": moneda, "fecha": fecha,
            "monto": monto, "glosa": glosa, "op": op}


# ─── El sufijo ───────────────────────────────────────────────────────────────
@pytest.mark.parametrize("op", ["01482615", "02482615", "00482615", "482615"])
def test_los_tres_prefijos_son_el_mismo_numero(op):
    """Los dos primeros dígitos son el CANAL. Fue el hallazgo que habilitó todo esto."""
    assert OP.sufijo(op) == "482615"


def test_el_batch_no_es_una_operacion():
    """ITF, portes y mantenimiento llevan 00000000: agruparlos juntaría el mes en un montón."""
    assert OP.sufijo("00000000") is None
    assert OP.sufijo("") is None and OP.sufijo(None) is None and OP.sufijo("123") is None


# ─── La regla 1: monedas ─────────────────────────────────────────────────────
def test_moneda_mixta_no_suma_y_reporta_la_tasa():
    """El caso real del 01/09: S/55.00 → US$15.97 es un cambio, no un descuadre."""
    pares, _ = OP.emparejar([mov("193-2000201-0-21", -55.00, "01307144"),
                             mov("193-2000203-1-23", 15.97, "02307144", moneda="USD")])
    assert len(pares) == 1
    p = pares[0]
    assert p["estado"] == OP.CAMBIO
    assert p["neto"] is None, "sumar soles con dólares inventa una brecha de 39.03"
    assert p["tasa"] == 3.444


def test_misma_moneda_si_se_neutraliza():
    pares, _ = OP.emparejar([mov("193-1000101-0-11", -215.00, "01482615", empresa="ALTAMIRA"),
                             mov("193-2000201-0-21", 215.00, "02482615")])
    assert pares[0]["estado"] == OP.NEUTRO and pares[0]["neto"] == 0.0


def test_misma_moneda_que_no_cuadra_es_descuadre_no_silencio():
    pares, _ = OP.emparejar([mov("193-1000101-0-11", -215.00, "01482615", empresa="ALTAMIRA"),
                             mov("193-2000201-0-21", 200.00, "02482615")])
    assert pares[0]["estado"] == OP.DESCUADRE and pares[0]["neto"] == -15.0


# ─── La regla 2: no cerrar por sufijo a secas ────────────────────────────────
def test_dos_patas_en_la_misma_cuenta_no_son_un_par():
    pares, desc = OP.emparejar([mov("193-2000201-0-21", -100.0, "01111111"),
                                mov("193-2000201-0-21", 100.0, "02111111")])
    assert pares == [] and len(desc) == 1 and desc[0]["motivo"] == "una sola cuenta"


def test_sufijo_colisionado_se_descarta_por_distancia():
    """Dos operaciones sin relación pueden compartir seis dígitos; la fecha las separa."""
    lejos = F + dt.timedelta(days=OP.DIAS_MAXIMOS + 1)
    pares, desc = OP.emparejar([mov("193-2000201-0-21", -100.0, "01111111"),
                                mov("193-1000101-0-11", 100.0, "02111111",
                                    empresa="ALTAMIRA", fecha=lejos)])
    assert pares == []
    assert len(desc) == 1 and "colisionado" in desc[0]["motivo"]


def test_lo_descartado_se_devuelve_nunca_se_tira():
    """Un descarte invisible es un hueco que nadie audita."""
    _, desc = OP.emparejar([mov("193-2000201-0-21", -1.0, "01111111"),
                            mov("193-2000201-0-21", 1.0, "02111111")])
    assert desc and desc[0]["patas"], "el descarte debe traer las patas para poder mirarlas"


# ─── Clasificación y uso ─────────────────────────────────────────────────────
def test_distingue_intercia_de_traspaso_propio():
    propio = OP.emparejar([mov("193-2000201-0-21", -60.24, "01640381"),
                           mov("193-2000202-0-22", 60.24, "02640381")])[0][0]
    inter = OP.emparejar([mov("193-1000101-0-11", -2.34, "01519023", empresa="ALTAMIRA"),
                          mov("193-2000201-0-21", 2.34, "02519023")])[0][0]
    assert propio["tipo"] == OP.PROPIO and inter["tipo"] == OP.INTERCIA


def test_internos_saca_el_abono_anonimo_de_la_pila():
    """El motivo de todo: un «DE OTRA CUENTA» que es nuestro deja de ser «por identificar»."""
    movs = [mov("193-1000101-0-11", -215.00, "01482615", empresa="ALTAMIRA"),
            mov("193-2000201-0-21", 215.00, "02482615"),
            mov("193-2000201-0-21", 12500.00, "02715206", glosa="DE CLIENTE EJEMPLO SAC")]
    llaves = OP.internos(movs)
    assert ("193-2000201-0-21", F, 215.00, "02482615") in llaves
    assert ("193-2000201-0-21", F, 12500.00, "02715206") not in llaves, \
        "un cobro real de un tercero NO puede quedar marcado como movimiento interno"


def test_un_movimiento_suelto_no_genera_par():
    assert OP.emparejar([mov("193-2000201-0-21", 215.00, "02482615")]) == ([], [])


# ─── El registro de cuentas ──────────────────────────────────────────────────
def test_las_seis_cuentas_del_grupo_estan_en_el_codigo():
    """Vivían como prosa en el archivo del conciliador; una regla que depende de una lista
    solo vale si la lista está en el código."""
    todas = [c for ctas in EMP.CUENTAS_GRUPO.values() for c in ctas]
    assert len(todas) == 6
    assert {c[0] for c in EMP.CUENTAS_GRUPO["ALTAMIRA"]} == {
        "193-1000101-0-11", "193-1000102-0-12", "193-10001031-1-13"}
    assert EMP.CUENTA_BCP["MERAK"] == EMP.CUENTAS_GRUPO["MERAK"][0][0], \
        "la operativa del registro debe ser la misma que CUENTA_BCP"


@pytest.mark.parametrize("texto,esperado", [
    ("193-2000201-0-21 - MERAK CERTIFICACIONES SAC", "MERAK"),   # cabecera del extracto
    ("A 193 2000202 0", "MERAK"),                                # glosa de cargo
    ("A 193 1000102 0", "ALTAMIRA"),
    ("A 193 3000301 8", None),                                   # cuenta de un tercero
    ("", None),
])
def test_reconoce_cuenta_propia_en_las_dos_formas(texto, esperado):
    r = EMP.cuenta_del_grupo(texto)
    assert (r[0] if r else None) == esperado


# ─── El lector: columnas por nombre, no por posición ─────────────────────────
def test_las_columnas_se_ubican_por_nombre_no_por_indice():
    """El .csv trae 7 columnas y el .xlsx 11: la operación cambia de posición entre envases.
    Fijar el índice leía la columna equivocada EN SILENCIO, que es peor que no leer."""
    import cruzar_operaciones as C
    xlsx = ["Fecha", "Fecha valuta", "Descripción operación", "Monto", "Saldo",
            "Sucursal - agencia", "Operación - Número", "Operación - Hora", "Usuario", "UTC", "Referencia2"]
    csv_ = ["Fecha", "Fecha valuta", "Descripci?n operaci?n", "Monto", "Saldo",
            "Operaci?n - N?mero", "Referencia2"]
    assert C._indices(xlsx) == {"fecha": 0, "desc": 2, "monto": 3, "op": 6}
    assert C._indices(csv_) == {"fecha": 0, "desc": 2, "monto": 3, "op": 5}
    assert C._indices(["algo", "otra cosa"]) is None, "formato desconocido: no adivinar"


def test_la_hora_no_se_confunde_con_el_numero_de_operacion():
    import cruzar_operaciones as C
    ix = C._indices(["Fecha", "Descripción operación", "Monto", "Operación - Hora", "Operación - Número"])
    assert ix["op"] == 4
