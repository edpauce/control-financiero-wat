"""El plazo de 10 días de la Ley de Contrataciones solo aplica a clientes del Estado.

El riesgo real no es dejar fuera a una entidad, sino meter a un privado: en una cartera pueden
convivir «EMPRESA MUNICIPAL DE AGUA VALLEGRANDE» (estatal) y «VALLEGRANDE TOURS» (agencia de
viajes). Un patrón demasiado corto los confunde y le pondría 10 días de plazo a un cliente
privado, que es como inventar una deuda vencida.

Todos los nombres son ficticios.
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tools"))
from lib import clientes as CL


ESTATALES = [
    "EMPRESA MUNICIPAL DE AGUA VALLEGRANDE S.A.",
    "Empresa Municipal De Agua Vallegra",              # truncada, como la guarda el maestro
    "MUNICIPALIDAD DISTRITAL DE EJEMPLO",
    "GOBIERNO REGIONAL DE EJEMPLO",
    "UNIVERSIDAD NACIONAL DE EJEMPLO",
]
PRIVADOS = [
    "VALLEGRANDE TOURS E.I.R.L.",                      # el falso positivo que motiva el test
    "INGENIERIA Y REDES S.A.C.",
    "CLUB SOCIAL EJEMPLO",
    "MAPSA S.A.C.",
    "RUMI CONCRETOS S.A.C.",
    "",
]


def test_entidades_del_estado_se_reconocen():
    for c in ESTATALES:
        assert CL.es_estatal(c), c


def test_privados_no_se_confunden_con_el_estado():
    for c in PRIVADOS:
        assert not CL.es_estatal(c), c


def test_plazo_solo_para_el_estado():
    assert CL.plazo_dias("GOBIERNO REGIONAL DE EJEMPLO") == CL.PLAZO_ESTADO_DIAS == 10
    assert CL.plazo_dias("MAPSA S.A.C.") is None


def test_tildes_y_mayusculas_no_alteran_el_resultado():
    assert CL.es_estatal("municipalidad distrital de ejemplo")
    assert CL.es_estatal("UNIVERSIDAD NACIONAL DE EJEMPLO — Facultad de Ingeniería")
