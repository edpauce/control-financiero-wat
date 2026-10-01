"""
quien_deposito(): la autoría de una detracción abonada en el BN se decide por la salida
`PAGO DETRAC` del BCP o por el rol en la constancia SUNAT — nunca por los códigos del BN
(regla desmentida por la evidencia).
"""
import pytest

from lib import empresas as E


def test_salida_pago_detrac_en_bcp_es_nuestra_y_manda():
    # Abonos "1234" con RUC del cliente, pero PAGO DETRAC en BCP → nuestra.
    assert E.quien_deposito(True, None) == E.DEPOSITO_EMPRESA
    # Aunque la constancia nombre al cliente como adquiriente (siempre lo hace), la salida BCP manda.
    assert E.quien_deposito(True, "adquiriente") == E.DEPOSITO_EMPRESA


@pytest.mark.parametrize("rol,esperado", [
    ("proveedor", E.DEPOSITO_EMPRESA),
    ("Proveedor", E.DEPOSITO_EMPRESA),
    ("  PROVEEDOR  ", E.DEPOSITO_EMPRESA),
    ("adquiriente", E.DEPOSITO_CLIENTE),
    ("Adquirente", E.DEPOSITO_CLIENTE),
])
def test_sin_salida_bcp_decide_el_rol_en_la_constancia(rol, esperado):
    assert E.quien_deposito(False, rol) == esperado


def test_sin_evidencia_no_se_concluye_por_descarte():
    # ENT.EFEC en ventanilla: sin PAGO DETRAC ni constancia → 🟠, jamás "cliente".
    assert E.quien_deposito(False, None) == E.DEPOSITO_POR_CONFIRMAR
    assert E.quien_deposito(False, "") == E.DEPOSITO_POR_CONFIRMAR
    assert E.quien_deposito(False, "otro") == E.DEPOSITO_POR_CONFIRMAR
    # Y el glosario legado ya no etiqueta depositantes.
    assert not {"deposito_cliente", "auto_detraccion"} & set(E.CODIGOS_BN.values())
