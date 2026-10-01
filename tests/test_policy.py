"""
El gate de escritura: un check por test, más las garantías estructurales.

Lo que se está protegiendo aquí no es una función: es que las reglas de CLAUDE.md dejen de ser
prompts. Cada test nombra el caso real que motiva su check, para que dentro de seis meses se
entienda por qué existe y nadie lo borre por "redundante".
"""
import json
import os
from datetime import date

import pytest

from lib import policy as P

CONCILIACION = "03. CONCILIACIONES/GLOBAL/CONCILIACION_2026_REV_2026-08-31.xlsx"
COBROS = "04. FINANZAS/08. COBROS/COBROS_2026-08.xlsx"
DETRACCIONES = "04. FINANZAS/10. DETRACCIONES/SEGUIMIENTO_DETRACCIONES_2026.xlsx"
EECC_MER = "02. MERAK/01. EECC/2026/EECC_MER_2026-08_movimientos.xlsx"


@pytest.fixture
def raiz(tmp_path, monkeypatch):
    """Raíz aislada: los tests no tocan el repo real ni su log de auditoría."""
    monkeypatch.setattr(P, "RAIZ", tmp_path)
    return tmp_path


def maestro(raiz, rel: str, con_bak: bool = False):
    """Crea el archivo (y opcionalmente su .bak de hoy) dentro de la raíz de prueba."""
    p = raiz / rel
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text("xlsx de mentira", encoding="utf-8")
    if con_bak:
        (raiz / f"{rel}.{date.today().isoformat()}.bak").write_text("bak", encoding="utf-8")
    return p


def accion(raiz, **kw):
    base = dict(tool="t", empresa="MERAK", capacidad="cobros.escribir", periodo="2026-08",
                aplicar=True, evidencia="EVIDENCIA_2026-08.csv f12")
    base.update(kw)
    return P.Accion(**base)


def cerrar_mes(raiz, periodo: str):
    cfg = raiz / "tools" / "config" / f"{periodo}.py"
    cfg.parent.mkdir(parents=True, exist_ok=True)
    cfg.write_text("CERRADO = True\n", encoding="utf-8")


# ── Check 1: entidad ─────────────────────────────────────────────────────────
def test_1_entidad_no_se_cruza_sin_capacidad(raiz):
    """Un tool que dice ser de MERAK no escribe un maestro que vive en la carpeta de ALTAMIRA."""
    m = maestro(raiz, "01. ALTAMIRA/02. FACTURACION/2026/Emitidas/FACT-EMI_ALT_2026-08_Reporte.xlsx",
                con_bak=True)
    d = P.evaluar(accion(raiz, maestro=m, empresa="MERAK"))
    assert d.estado == P.DENEGAR
    assert any("entidad" in m_ for m_ in d.motivos)


def test_1_intercia_si_cruza_pero_lo_deja_escrito(raiz):
    """El cruce entre empresas existe (una paga por cuenta de la otra): se permite y se anota."""
    m = maestro(raiz, "01. ALTAMIRA/01. EECC/2026/x.xlsx", con_bak=True)
    d = P.evaluar(accion(raiz, maestro=m, empresa="MERAK", capacidad="intercia"))
    assert not any("entidad" in x for x in d.motivos)
    assert any("cruce entre empresas" in h for h in d.hallazgos)


# ── Check 2: capacidad ↔ maestro (la frontera real) ──────────────────────────
def test_2_cobros_no_escribe_detracciones(raiz):
    """La regla de coordinación de CLAUDE.md, hecha código: cada agente en su maestro."""
    m = maestro(raiz, DETRACCIONES, con_bak=True)
    d = P.evaluar(accion(raiz, maestro=m, empresa=None, capacidad="cobros.escribir"))
    assert d.estado == P.DENEGAR
    assert any("capacidad" in x for x in d.motivos)


def test_2_capacidad_correcta_permite(raiz):
    m = maestro(raiz, COBROS, con_bak=True)
    assert P.evaluar(accion(raiz, maestro=m, empresa=None)).estado == P.PERMITIR


def test_2_un_insumo_del_banco_no_se_escribe_nunca(raiz):
    """Los EECC son lo que el banco dijo. Escribirlos no es una operación, es un error."""
    m = maestro(raiz, EECC_MER, con_bak=True)
    d = P.evaluar(accion(raiz, maestro=m, capacidad="conciliacion.escribir"))
    assert d.estado == P.DENEGAR
    assert any("INSUMO" in x for x in d.motivos)


def test_2_maestro_desconocido_se_deniega(raiz):
    """Deny by default: un archivo que el mapa no cubre no se escribe 'por si acaso'."""
    m = maestro(raiz, "04. FINANZAS/08. COBROS/INVENTADO.xlsx", con_bak=True)
    assert P.evaluar(accion(raiz, maestro=m)).estado == P.DENEGAR


# ── Check 3: periodo abierto ─────────────────────────────────────────────────
def test_3_mes_cerrado_deniega(raiz):
    """Generaliza cierre.exigir_abierto(), que hoy solo llaman 3 tools de 60."""
    cerrar_mes(raiz, "2026-07")
    m = maestro(raiz, COBROS, con_bak=True)
    d = P.evaluar(accion(raiz, maestro=m, periodo="2026-07", empresa=None))
    assert d.estado == P.DENEGAR
    assert any("CERRADO" in x for x in d.motivos)


def test_3_forzar_no_permite_en_silencio_sino_que_pide_aprobacion(raiz):
    """--forzar convierte deny en 'requiere aprobación' CON motivo escrito; no en permitir."""
    cerrar_mes(raiz, "2026-07")
    m = maestro(raiz, COBROS, con_bak=True)
    d = P.evaluar(accion(raiz, maestro=m, periodo="2026-07", empresa=None, forzar=True))
    assert d.estado == P.APROBACION
    assert any("--forzar" in x for x in d.motivos)


def test_3_sin_config_el_mes_esta_abierto(raiz):
    m = maestro(raiz, COBROS, con_bak=True)
    assert P.evaluar(accion(raiz, maestro=m, periodo="2026-12", empresa=None)).estado == P.PERMITIR


# ── Check 4: respaldo previo ─────────────────────────────────────────────────
def test_4_sin_bak_del_dia_no_se_escribe(raiz):
    """El .bak es lo que salvó CREDITOS cuando insert_rows lo rompió seis días."""
    m = maestro(raiz, COBROS, con_bak=False)
    d = P.evaluar(accion(raiz, maestro=m, empresa=None))
    assert d.estado == P.DENEGAR
    assert any("respaldar" in x for x in d.motivos)


def test_4_acepta_la_otra_convencion_de_respaldo(raiz):
    """Hay tools que respaldan como `<maestro>.BACKUP-AAAA-MM-DD.xlsx`. Es un respaldo
    igual de válido; atar el check a un solo nombre denegaría escrituras bien protegidas."""
    m = maestro(raiz, "08. HERRAMIENTAS/DIRECTORIO_CUENTAS.xlsx")
    (raiz / "08. HERRAMIENTAS" /
     f"DIRECTORIO_CUENTAS.BACKUP-{date.today().isoformat()}.xlsx").write_text("b")
    d = P.evaluar(accion(raiz, maestro=m, empresa=None, capacidad="referencia.escribir",
                         evidencia=None))
    assert d.estado == P.PERMITIR, d.motivos


def test_4_un_archivo_nuevo_no_exige_respaldo(raiz):
    """integrar_conciliacion escribe una revisión NUEVA fechada en vez de sobrescribir. Es el
    patrón más seguro de los dos; exigirle un .bak de un archivo que aún no existe lo castigaría."""
    nuevo = raiz / "03. CONCILIACIONES/GLOBAL/CONCILIACION_2026_REV_2026-09-30.xlsx"
    nuevo.parent.mkdir(parents=True, exist_ok=True)          # el archivo NO se crea a propósito
    d = P.evaluar(accion(raiz, maestro=nuevo, empresa=None,
                         capacidad="conciliacion.escribir", evidencia=None))
    assert d.estado == P.PERMITIR, d.motivos
    assert any("archivo nuevo" in h for h in d.hallazgos)


def test_4_el_respaldo_solo_se_exige_al_escribir(raiz):
    """Sin --aplicar no hay escritura, así que exigir .bak sería ruido."""
    m = maestro(raiz, COBROS, con_bak=False)
    d = P.evaluar(accion(raiz, maestro=m, empresa=None, aplicar=False))
    assert d.estado == P.SIMULACION
    assert not d.motivos


# ── Check 5: umbral de monto ─────────────────────────────────────────────────
def test_5_monto_grande_pide_aprobacion(raiz):
    m = maestro(raiz, COBROS, con_bak=True)
    d = P.evaluar(accion(raiz, maestro=m, empresa=None, monto=12_500.0))
    assert d.estado == P.APROBACION


def test_5_el_abono_de_4400_pasa_por_debajo_del_umbral(raiz):
    """CASO VIVO. Un abono mal atribuido de S/4,400 queda BAJO los S/5,000 y el
    check 5 no lo ve. Lo atrapa el check 6, porque aquel cruce no citaba fuente. Si algún día se
    baja UMBRAL_MOVIMIENTO, este test lo dirá — y si se sube, dirá que el 6 queda solo."""
    m = maestro(raiz, COBROS, con_bak=True)
    solo_monto = P.evaluar(accion(raiz, maestro=m, empresa=None, monto=4_400.0))
    assert solo_monto.estado == P.PERMITIR, "el umbral por movimiento no ve los S/4,400"
    sin_fuente = P.evaluar(accion(raiz, maestro=m, empresa=None, monto=4_400.0, evidencia=None))
    assert sin_fuente.estado == P.APROBACION, "el check 6 es el que sí lo atrapa"


def test_5_total_de_corrida(raiz):
    m = maestro(raiz, COBROS, con_bak=True)
    assert P.evaluar(accion(raiz, maestro=m, empresa=None, total=55_000.00)).estado == P.APROBACION


# ── Check 6: evidencia citada ────────────────────────────────────────────────
def test_6_cambiar_un_cobro_sin_fuente_pide_aprobacion(raiz):
    m = maestro(raiz, COBROS, con_bak=True)
    d = P.evaluar(accion(raiz, maestro=m, empresa=None, evidencia=None))
    assert d.estado == P.APROBACION
    assert any("evidencia" in x for x in d.motivos)


def test_6_la_conciliacion_no_exige_evidencia(raiz):
    """Clasificar un movimiento no cambia el estado de un cobro: el check 6 no aplica ahí."""
    m = maestro(raiz, CONCILIACION, con_bak=True)
    d = P.evaluar(accion(raiz, maestro=m, empresa=None,
                         capacidad="conciliacion.escribir", evidencia=None))
    assert d.estado == P.PERMITIR


# ── Check 7: dry-run por defecto ─────────────────────────────────────────────
def test_7_sin_aplicar_es_simulacion(raiz):
    m = maestro(raiz, COBROS, con_bak=True)
    assert P.evaluar(accion(raiz, maestro=m, empresa=None, aplicar=False)).estado == P.SIMULACION


def test_7_exigir_devuelve_false_en_simulacion(raiz, capsys):
    m = maestro(raiz, COBROS, con_bak=True)
    assert P.exigir(P.evaluar(accion(raiz, maestro=m, empresa=None, aplicar=False))) is False


# ── Garantías estructurales ──────────────────────────────────────────────────
def test_el_mapa_cubre_todas_las_claves_de_sources():
    """Si sources.py gana un maestro y nadie le asigna capacidad, es un agujero. Falla nombrándolo."""
    from lib import sources as S
    faltan = sorted(set(S.PATHS) - set(P.CAPACIDAD_DE_MAESTRO))
    assert not faltan, (f"claves de sources.PATHS sin capacidad en policy: {faltan}. "
                        f"Asígnalas en CAPACIDAD_DE_MAESTRO (None si es un insumo).")


def test_el_mapa_no_inventa_claves():
    from lib import sources as S
    sobran = sorted(set(P.CAPACIDAD_DE_MAESTRO) - set(S.PATHS))
    assert not sobran, f"policy declara claves que sources.PATHS no tiene: {sobran}"


def test_evaluar_no_escribe_nada(raiz):
    """La decisión es pura: se puede testear sin efectos. Ni el log se toca."""
    m = maestro(raiz, COBROS, con_bak=True)
    antes = {p for p in raiz.rglob("*")}
    P.evaluar(accion(raiz, maestro=m, empresa=None))
    assert {p for p in raiz.rglob("*")} == antes
    assert not P.ruta_log().exists()


def test_el_log_solo_crece(raiz):
    """Append-only: la primera línea sigue idéntica tras la segunda escritura."""
    m = maestro(raiz, COBROS, con_bak=True)
    P.registrar(P.evaluar(accion(raiz, maestro=m, empresa=None, monto=1.0)))
    primera = P.ruta_log().read_text(encoding="utf-8").splitlines()[0]
    P.registrar(P.evaluar(accion(raiz, maestro=m, empresa=None, monto=2.0)))
    lineas = P.ruta_log().read_text(encoding="utf-8").splitlines()
    assert len(lineas) == 2
    assert lineas[0] == primera
    assert json.loads(lineas[1])["monto"] == 2.0


def test_el_evento_lleva_lo_que_hace_falta_para_auditar(raiz):
    m = maestro(raiz, COBROS, con_bak=True)
    ev = P.evaluar(accion(raiz, maestro=m, empresa=None, monto=4_400.0)).evento
    for campo in ("ts", "tool", "agente", "empresa", "maestro", "capacidad", "periodo",
                  "aplicar", "monto", "evidencia", "estado", "motivos", "bak"):
        assert campo in ev, f"el evento no lleva «{campo}»"
    assert ev["bak"] is not None, "el .bak debe quedar citado en el evento"


def test_todo_intento_se_registra_incluso_el_denegado(raiz):
    """Un deny sin rastro no sirve para aprender."""
    m = maestro(raiz, DETRACCIONES, con_bak=True)
    dec = P.evaluar(accion(raiz, maestro=m, empresa=None, capacidad="cobros.escribir"))
    with pytest.raises(SystemExit):
        P.exigir(dec)
    assert json.loads(P.ruta_log().read_text(encoding="utf-8").splitlines()[0])["estado"] == P.DENEGAR


def test_el_auditor_no_tiene_ninguna_capacidad_de_escritura():
    """No se le bloquea por llamarse auditor (eso sería confiar en la identidad declarada):
    se le bloquea porque ninguna capacidad suya cubre ningún maestro."""
    duenos = {c for c in P.CAPACIDAD_DE_MAESTRO.values() if c}
    assert not any(c.startswith("auditor") for c in duenos)


def test_aprobacion_sin_terminal_se_deniega(raiz):
    """Pedir permiso a nadie no es pedir permiso: launchd y sub-agentes no aprueban solos."""
    m = maestro(raiz, COBROS, con_bak=True)
    dec = P.evaluar(accion(raiz, maestro=m, empresa=None, monto=12_500.0))
    with pytest.raises(SystemExit):
        P.exigir(dec, interactivo=False)


def test_un_deny_no_se_degrada_a_aprobacion(raiz):
    """Si algo está denegado, que además pida aprobación no lo ablanda."""
    m = maestro(raiz, COBROS, con_bak=False)          # sin .bak → deny
    d = P.evaluar(accion(raiz, maestro=m, empresa=None, monto=99_999.0))   # y monto alto
    assert d.estado == P.DENEGAR


def test_un_test_no_ensucia_el_log_de_produccion(monkeypatch, tmp_path):
    """CASO VIVO: un test ejecutó un tool con --forzar y su evento llegó al jsonl real.
    Un rastro de auditoría con eventos que nunca ocurrieron no sirve para auditar."""
    real = P.ruta_log()                      # RAIZ sin tocar = el repo de verdad
    antes = real.read_text(encoding="utf-8") if real.exists() else None
    P.registrar(P.evaluar(P.Accion(tool="test", capacidad="intercia", aplicar=True)))
    ahora = real.read_text(encoding="utf-8") if real.exists() else None
    assert ahora == antes, "registrar() escribió en el log de producción durante un test"


def test_un_archivo_fuera_del_repo_no_lo_gobierna_el_gate(raiz, tmp_path_factory):
    """El gate protege los maestros DEL REPO. Un fixture de test o una copia suelta no lo son, y
    denegarlos no protegería nada: solo estorbaría para probar los tools. Con EMPRESA_ROOT
    apuntando a un sandbox, ese sandbox SÍ queda gateado (es lo que hace la raíz)."""
    fuera = tmp_path_factory.mktemp("otro") / "COBROS_2026-08.xlsx"
    fuera.write_text("copia suelta", encoding="utf-8")
    d = P.evaluar(accion(raiz, maestro=fuera, empresa=None, evidencia=None))
    assert d.estado == P.PERMITIR
    assert any("fuera de" in h for h in d.hallazgos)
