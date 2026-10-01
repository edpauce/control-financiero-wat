"""
El test que sostiene el gate en el tiempo.

Sin esto el gate se erosiona en tres meses: alguien escribe un tool nuevo, hace `wb.save()` sobre
un maestro y no se acuerda de `policy`. Nada falla, y la protección queda con un agujero que solo
se descubre el día que importa.

La regla: **todo tool que escriba un maestro debe pasar por el gate.** Las excepciones existen —
hay tools que escriben informes, paneles o archivos derivados que no son maestros — pero cada una
va nombrada aquí abajo CON SU MOTIVO. Una exención sin motivo escrito es una deuda escondida.
"""
import ast
import os
import re

import pytest

TOOLS = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "tools"))

# ── Exenciones ───────────────────────────────────────────────────────────────
# Cada entrada: por qué este archivo escribe algo sin pasar por el gate. Si un día un exento
# empieza a tocar un maestro de verdad, hay que quitarlo de aquí y cablearle el gate.
EXENTOS: dict[str, str] = {
    # El gate mismo y su utilería.
    "lib/policy.py": "es el gate",
    "lib/cierre.py": "respaldar() COPIA el maestro; el gate exige justamente que esto se llame antes",
    "lib/sellos.py": "sella dentro de un workbook que su llamador ya tiene abierto y gateado",
    # Escriben SU PROPIA salida (informes), no un maestro.
    "auditar_fuentes.py": "solo informa",
    "cruzar_tributos.py": "escribe un informe de cruce",
    # ⚠️ Esta exención depende de una CONVENCIÓN DE USO, no de una propiedad del archivo: vale
    # mientras nadie lo corra suelto. Si algún día se ejecuta fuera del cierre, escribiría hojas
    # dentro del SEGUIMIENTO sin gate. Riesgo bajo hoy; queda anotado para no olvidarlo.
    "detalle_detracciones.py": "escribe hojas de DETALLE dentro del SEGUIMIENTO; su llamador"
                               " ya registra la corrida — vale mientras no se corra suelto",
}

# Señales de que un archivo ESCRIBE algo persistente.
ESCRITURA = re.compile(r"\.save\(|\bto_excel\(|open\([^)]*['\"][wa]b?['\"]|\.write_text\(")
# Señales de que ese archivo escribe un MAESTRO (y no un informe suelto).
MAESTRO = re.compile(r"S\.latest\(|sources\.latest\(|base_master\(|latest\(['\"]")


def _tools():
    for raiz, _, ficheros in os.walk(TOOLS):
        if "__pycache__" in raiz or "/scrape" in raiz.replace(os.sep, "/"):
            continue
        for f in sorted(ficheros):
            if f.endswith(".py") and f != "__init__.py":
                ruta = os.path.join(raiz, f)
                yield os.path.relpath(ruta, TOOLS).replace(os.sep, "/"), ruta


def _lee(ruta):
    with open(ruta, encoding="utf-8") as fh:
        return fh.read()


def test_todo_tool_que_escribe_un_maestro_pasa_por_el_gate():
    """Si esto falla nombrando un archivo, o le cableas el gate o lo declaras exento con motivo."""
    culpables = []
    for rel, ruta in _tools():
        if rel in EXENTOS or os.path.basename(rel) in EXENTOS:
            continue
        src = _lee(ruta)
        if ESCRITURA.search(src) and MAESTRO.search(src) and "policy" not in src:
            culpables.append(rel)
    assert not culpables, (
        "estos tools escriben un maestro sin pasar por el gate:\n  "
        + "\n  ".join(culpables)
        + "\n\nCabléales `from lib import policy as POL` + POL.exigir(POL.evaluar(...)) antes de "
          "guardar, o decláralos en EXENTOS de este test CON EL MOTIVO."
    )


def test_las_exenciones_apuntan_a_archivos_que_existen():
    """Una exención de un archivo borrado es basura que oculta el siguiente agujero."""
    reales = {rel for rel, _ in _tools()} | {os.path.basename(rel) for rel, _ in _tools()}
    fantasmas = sorted(set(EXENTOS) - reales)
    assert not fantasmas, f"exenciones de archivos que ya no existen: {fantasmas}"


def test_toda_exencion_lleva_motivo():
    vacias = sorted(k for k, v in EXENTOS.items() if not v or len(v.strip()) < 10)
    assert not vacias, f"exenciones sin motivo escrito: {vacias}"


# Un motivo que confiesa ser temporal no es un motivo: es deuda con permiso de residencia.
# Un tool estuvo exento con «pendiente de cablear en fase 2» mientras escribía un maestro
# sin gate, y el test lo dejó pasar porque solo exigía QUE hubiera motivo.
CONFESIONES = ("pendiente", "por cablear", "todavía", "todavia", "por ahora", "temporal", "TODO",
               "más adelante", "mas adelante", "provisional")


def test_ninguna_exencion_confiesa_ser_temporal():
    deudas = sorted(f"{k}: «{v}»" for k, v in EXENTOS.items()
                    if any(c.lower() in v.lower() for c in CONFESIONES))
    assert not deudas, ("estas exenciones dicen que falta hacer el trabajo — cabléales el gate "
                        "en vez de dejarlas exentas:\n  " + "\n  ".join(deudas))


def test_los_tools_gateados_construyen_una_accion_valida():
    """Un `POL.exigir(...)` con la capacidad mal escrita denegaría siempre, y en un tool que solo
    corre en el cierre eso se descubre el peor día. Se valida estáticamente."""
    from lib import policy as P

    validas = {c for c in P.CAPACIDAD_DE_MAESTRO.values() if c} | set(P.CAPACIDADES_INTERCIA)
    malas = []
    for rel, ruta in _tools():
        src = _lee(ruta)
        if "POL.Accion(" not in src:
            continue
        for arbol in ast.walk(ast.parse(src)):
            if not (isinstance(arbol, ast.Call) and isinstance(arbol.func, ast.Attribute)
                    and arbol.func.attr == "Accion"):
                continue
            for kw in arbol.keywords:
                if kw.arg == "capacidad" and isinstance(kw.value, ast.Constant):
                    if kw.value.value not in validas:
                        malas.append(f"{rel}: capacidad «{kw.value.value}»")
    assert not malas, ("capacidades que el mapa no reconoce (denegarían siempre):\n  "
                       + "\n  ".join(malas))


def test_el_gate_va_antes_de_guardar_no_despues():
    """Un gate después del `save` no protege nada: el archivo ya se escribió."""
    tarde = []
    for rel, ruta in _tools():
        src = _lee(ruta)
        if "POL.exigir(" not in src:
            continue
        lineas = src.splitlines()
        primer_gate = next((i for i, l in enumerate(lineas) if "POL.exigir(" in l), None)
        # el primer save que escribe un maestro, ignorando los de módulos exentos
        primer_save = next((i for i, l in enumerate(lineas)
                            if ".save(" in l and not l.strip().startswith("#")), None)
        if primer_gate is not None and primer_save is not None and primer_save < primer_gate:
            tarde.append(f"{rel}: save en la línea {primer_save + 1}, gate en la {primer_gate + 1}")
    assert not tarde, "el gate llega tarde:\n  " + "\n  ".join(tarde)


# ── Que el check 5 no quede apagado donde importa ────────────────────────────
# Los umbrales (S/5,000 y S/20,000) solo vigilan si el tool DECLARA cuánto mueve. Durante un día
# entero el check estuvo probado y apagado: 17 de 18 tools llamaban al gate con monto=None, entre
# ellos todas las integraciones que meten extractos y facturas al maestro — justo donde se coló
# el abono de S/4,400. Estos son los que sí mueven importes y deben declararlos.
MUEVEN_DINERO = {"incorporar_rvie.py", "ingerir_sire_txt.py"}
# Los demás gateados no mueven importes: sellan, formatean, corrigen fechas, clasifican.
# Declararles un monto inventado sería peor que no declararlo.


def test_las_integraciones_declaran_cuanto_mueven():
    """Sin monto=, el check 5 no dispara nunca. Un umbral que no vigila nada es peor que ninguno:
    da la sensación de que hay un control donde no lo hay."""
    mudos = []
    for rel, ruta in _tools():
        nombre = os.path.basename(rel)
        if nombre not in MUEVEN_DINERO:
            continue
        src = _lee(ruta)
        if "POL.Accion(" in src and "monto=" not in src:
            mudos.append(rel)
    assert not mudos, ("estas integraciones mueven dinero al maestro y no declaran monto=, "
                       "así que el check 5 (umbral) no las vigila:\n  " + "\n  ".join(mudos))


def test_los_que_declaran_monto_estan_en_la_lista():
    """Si un tool nuevo empieza a declarar monto, entra en la lista — así no se pierde de vista."""
    fuera = []
    for rel, ruta in _tools():
        nombre = os.path.basename(rel)
        if nombre in MUEVEN_DINERO or nombre in EXENTOS or rel in EXENTOS:
            continue
        if "monto=" in _lee(ruta) and "POL.Accion(" in _lee(ruta):
            fuera.append(rel)
    assert not fuera, f"declaran monto pero no están en MUEVEN_DINERO: {fuera}"
