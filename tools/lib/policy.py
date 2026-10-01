#!/usr/bin/env python3
"""
tools/lib/policy.py — Gate DETERMINISTA de escritura sobre los maestros.

En una línea: las reglas de
seguridad de este repo eran texto en CLAUDE.md («el auditor jamás edita nada», «lo diario NO
escribe en los maestros»). Son buenas reglas y son prompts: un agente distraído las viola y nada
lo detecta. Esto las vuelve código.

CUATRO PRINCIPIOS QUE NO SE NEGOCIAN
  1. Sin LLM en la decisión. Python puro, determinista, sin red, sin dependencias nuevas.
  2. Deny by default en lo que ya es regla escrita: si CLAUDE.md dice que no se hace, se deniega,
     no se advierte.
  3. La frontera real es el mapa maestro↔capacidad, NO la identidad declarada. Nada obliga a un
     LLM a decir la verdad sobre quién es; sí se puede saber qué ARCHIVO está a punto de escribir.
  4. Todo intento se registra, permitido o no. Un deny sin rastro no sirve para aprender.

USO
    from lib import policy as P

    acc = P.Accion(tool="casar_cobros_huerfanos", maestro=ruta, empresa="MERAK",
                   capacidad="cobros.escribir", periodo="2026-08",
                   aplicar=("--aplicar" in sys.argv), monto=4400.0,
                   evidencia="EVIDENCIA_2026-08.csv fila 12")
    if P.exigir(P.evaluar(acc)):
        ...  # escribir

`evaluar()` es PURA: decide y no escribe nada (ni el log), para poder testear la decisión sin
matar el proceso de test. `exigir()` es quien registra el evento y corta el proceso.

LO QUE ESTA PIEZA NO HACE
No valida que el CRITERIO sea correcto — a qué factura va un abono anónimo sigue siendo juicio
del sub-agente con la evidencia delante, y así debe ser (automatizarlo produjo cruces
equivocados). El gate garantiza que la escritura sea legítima, respaldada, dentro del periodo,
dentro del dominio y registrada.
"""
from __future__ import annotations

import fnmatch
import json
import os
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

RAIZ = Path(os.environ.get("EMPRESA_ROOT",
                           Path(__file__).resolve().parent.parent.parent))

# ── Estados ──────────────────────────────────────────────────────────────────
PERMITIR = "permitir"
DENEGAR = "denegar"
APROBACION = "requiere_aprobacion"
SIMULACION = "simulacion"          # check 7: sin --aplicar no se escribe, nunca

# ── Umbrales (check 5) ───────────────────────────────────────────────────────
# Confirmados por gerencia. Conviene saber qué se está aceptando al fijarlos aquí:
# un abono mal atribuido de S/4,400 queda JUSTO DEBAJO de los 5,000, así que el check 5
# no lo habría visto. No es un descuido del umbral: en ese caso el que atrapa el error es el
# check 6 (evidencia citada), porque aquel cruce no tenía fuente. Los dos checks se reparten el
# trabajo — el 5 vigila el tamaño, el 6 vigila el sustento — y `test_policy` congela ese reparto
# (test_5_el_abono_de_4400_pasa_por_debajo_del_umbral). Si algún día se mueve UMBRAL_MOVIMIENTO,
# ese test dirá qué deja de cubrir el 5 y qué queda a cargo del 6.
UMBRAL_MOVIMIENTO = 5_000.0
UMBRAL_CORRIDA = 20_000.0

# ── Mapa capacidad ↔ maestro (la frontera real) ──────────────────────────────
# Deriva de la tabla de sub-agentes de CLAUDE.md. Un tool de cobros no escribe el SEGUIMIENTO de
# detracciones ni aunque se lo pidan. La clave es la de `sources.PATHS`; el valor, la capacidad
# que autoriza a escribirlo. `None` = NADIE lo escribe (son insumos del banco/OSE: si un tool
# intenta escribir un EECC, es un error, no una operación).
CAPACIDAD_DE_MAESTRO: dict[str, str | None] = {
    "conciliacion": "conciliacion.escribir",   # conciliador
    "cobros":       "cobros.escribir",         # cobrador
    "detracciones": "detracciones.escribir",   # detractor
    "proveedores":  "proveedores.escribir",    # pagador
    "flujo":        "flujo.escribir",          # tesorero
    "creditos":     "creditos.escribir",       # tesorero (deuda) — ausente en la propuesta
    "mapeo":        "referencia.escribir",     # conciliador — ausente en la propuesta
    "directorio":   "referencia.escribir",     # conciliador — ausente en la propuesta
    # Insumos: se leen, jamás se escriben.
    "eecc_alt": None, "eecc_mer": None,
    "bn_alt": None, "bn_mer": None,
    "efact_alt": None, "efact_mer": None,
}

# Capacidades que además pueden cruzar la frontera de empresa (check 1).
CAPACIDADES_INTERCIA = {"intercia", "conciliacion.escribir", "flujo.escribir"}

# El auditor NO figura aquí a propósito: no tiene ninguna capacidad de escritura. No se le
# bloquea "por ser el auditor" (eso sería confiar en la identidad declarada); se le bloquea
# porque ninguna capacidad suya cubre ningún maestro.
AGENTE_SOLO_LECTURA = {"auditor"}

CARPETA_EMPRESA = {"01. ALTAMIRA": "ALTAMIRA", "02. MERAK": "MERAK"}


# ── Tipos ────────────────────────────────────────────────────────────────────
@dataclass
class Accion:
    """Lo que un tool está a punto de hacer. `maestro` es la ruta del archivo a escribir."""
    tool: str
    maestro: str | os.PathLike | None = None
    empresa: str | None = None            # None = el tool toca ambas
    capacidad: str | None = None
    periodo: str | None = None            # "AAAA-MM"
    aplicar: bool = False
    monto: float | None = None            # el mayor movimiento individual de la corrida
    total: float | None = None            # suma de la corrida
    evidencia: str | None = None
    forzar: bool = False
    agente: str | None = None             # default: env EMPRESA_AGENTE, o "orquestador"

    def __post_init__(self):
        if self.agente is None:
            self.agente = os.environ.get("EMPRESA_AGENTE", "orquestador")


@dataclass
class Decision:
    estado: str
    motivos: list[str] = field(default_factory=list)      # por qué se deniega / se pide aprobación
    hallazgos: list[str] = field(default_factory=list)    # avisos que no cambian el estado
    evento: dict = field(default_factory=dict)

    @property
    def permitida(self) -> bool:
        return self.estado == PERMITIR


# ── Utilidades ───────────────────────────────────────────────────────────────
def clave_de_maestro(ruta) -> str | None:
    """Qué clave de `sources.PATHS` describe este archivo. None si ninguna lo cubre."""
    if ruta is None:
        return None
    from lib import sources as S

    p = Path(ruta)
    try:
        rel = p.resolve().relative_to(RAIZ.resolve()).as_posix()
    except (ValueError, OSError):
        rel = p.as_posix()
    for clave, patron in S.PATHS.items():
        if fnmatch.fnmatch(rel, patron) or fnmatch.fnmatch(p.name, Path(patron).name):
            return clave
    return None


def empresa_de_ruta(ruta) -> str | None:
    """La empresa dueña del archivo según su carpeta. None = maestro de grupo."""
    if ruta is None:
        return None
    partes = Path(ruta).as_posix().split("/")
    for parte in partes:
        if parte in CARPETA_EMPRESA:
            return CARPETA_EMPRESA[parte]
    return None


def periodo_de(ruta) -> str | None:
    """AAAA-MM leído del nombre del maestro (COBROS_2026-08.xlsx). None si no lo lleva."""
    import re as _re
    m = _re.search(r"(\d{4})-(\d{2})", Path(ruta).name) if ruta else None
    return f"{m.group(1)}-{m.group(2)}" if m else None


def _dentro_del_repo(ruta) -> bool:
    """¿El archivo vive bajo la raíz del repo (la real, o la que fije EMPRESA_ROOT)?"""
    try:
        Path(ruta).resolve().relative_to(RAIZ.resolve())
        return True
    except (ValueError, OSError):
        return False


def _mes_cerrado(periodo: str | None) -> bool:
    """¿tools/config/{periodo}.py declara CERRADO = True? Sin config, el mes está abierto."""
    if not periodo:
        return False
    cfg = RAIZ / "tools" / "config" / f"{periodo}.py"
    if not cfg.exists():
        return False
    try:
        for linea in cfg.read_text(encoding="utf-8").splitlines():
            limpia = linea.split("#")[0].strip().replace(" ", "")
            if limpia.startswith("CERRADO="):
                return limpia.split("=", 1)[1].strip().lower() == "true"
    except OSError:
        return False
    return False


def _bak_de_hoy(ruta) -> Path | None:
    """Cualquier respaldo de HOY junto al maestro, sea cual sea la convención del tool.

    Lo que el check 4 exige es que exista una copia previa del día, no un nombre concreto:
    `cierre.respaldar()` deja `<maestro>.AAAA-MM-DD.bak`, pero otros tools usan
    `<maestro sin .xlsx>.BACKUP-AAAA-MM-DD.xlsx` y ambas cumplen el mismo propósito. Atarse a
    un solo patrón haría que el gate denegara escrituras perfectamente respaldadas.
    """
    if ruta is None:
        return None
    src = Path(ruta)
    hoy = date.today().isoformat()
    candidatos = [
        Path(f"{src}.{hoy}.bak"),                                   # cierre.respaldar()
        src.with_name(f"{src.stem}.BACKUP-{hoy}{src.suffix}"),      # convención alternativa
    ]
    for c in candidatos:
        if c.exists():
            return c
    # y cualquier otro archivo hermano que lleve la fecha de hoy y huela a respaldo
    try:
        for h in src.parent.glob(f"{src.stem}*"):
            if hoy in h.name and h != src and (".bak" in h.name.lower()
                                               or "backup" in h.name.lower()):
                return h
    except OSError:
        pass
    return None


# ── Los 7 checks ─────────────────────────────────────────────────────────────
def evaluar(acc: Accion) -> Decision:
    """Decide. NO escribe nada — ni el maestro, ni el log. Ver `exigir()`."""
    motivos: list[str] = []
    hallazgos: list[str] = []
    estado = PERMITIR
    clave = clave_de_maestro(acc.maestro)

    # LÍMITE DEL GATE, deliberado: gobierna las escrituras sobre los maestros DEL REPO. Un archivo
    # fuera de la raíz no es un maestro — es una copia, un fixture de test, un sandbox — y
    # denegarlo no protegería nada: solo volvería el gate un estorbo para probar los tools. La
    # raíz se resuelve con EMPRESA_ROOT, así que un sandbox declarado SÍ queda gateado.
    if acc.maestro is not None and not _dentro_del_repo(acc.maestro):
        return Decision(
            estado=PERMITIR if acc.aplicar else SIMULACION, motivos=[],
            hallazgos=[f"«{Path(acc.maestro).name}» está fuera de {RAIZ.name}: el gate no lo "
                       f"gobierna"],
            evento=_evento(acc, PERMITIR if acc.aplicar else SIMULACION, [], [], clave))

    def denegar(msg: str):
        nonlocal estado
        estado = DENEGAR
        motivos.append(msg)

    def aprobar(msg: str):
        nonlocal estado
        if estado != DENEGAR:          # un deny nunca se degrada a "pide aprobación"
            estado = APROBACION
        motivos.append(msg)

    # 1 — Entidad: el maestro pertenece a la empresa declarada.
    duena = empresa_de_ruta(acc.maestro)
    if duena and acc.empresa and duena != acc.empresa:
        if acc.capacidad in CAPACIDADES_INTERCIA:
            hallazgos.append(f"cruce entre empresas: {acc.empresa} escribe un maestro de {duena} "
                             f"(permitido por capacidad «{acc.capacidad}»)")
        else:
            denegar(f"entidad: el maestro es de {duena} y la acción declara {acc.empresa}; "
                    f"la capacidad «{acc.capacidad}» no cruza empresas")

    # 2 — Capacidad ↔ maestro: mapa duro. La frontera real.
    if acc.maestro is not None:
        if clave is None:
            denegar(f"maestro fuera del mapa: «{Path(acc.maestro).name}» no lo cubre ninguna "
                    f"clave de sources.PATHS — si es un maestro nuevo, primero decláralo")
        else:
            permitida = CAPACIDAD_DE_MAESTRO.get(clave, ...)
            if permitida is ...:
                denegar(f"«{clave}» no tiene capacidad asignada en policy.CAPACIDAD_DE_MAESTRO "
                        f"— un maestro sin dueño es un agujero, asígnalo")
            elif permitida is None:
                denegar(f"«{clave}» es un INSUMO (extracto del banco / reporte del OSE): se lee, "
                        f"no se escribe")
            elif acc.capacidad != permitida:
                denegar(f"capacidad: escribir «{clave}» exige «{permitida}» y la acción declara "
                        f"«{acc.capacidad}»")

    # 3 — Periodo abierto. Generaliza cierre.exigir_abierto() a todo tool, no solo a 3 de 60.
    if _mes_cerrado(acc.periodo):
        if acc.forzar:
            aprobar(f"periodo {acc.periodo} CERRADO y se pasó --forzar: regenerarlo pisa el "
                    f"formato manual de los maestros; requiere aprobación explícita")
        else:
            denegar(f"periodo {acc.periodo} está CERRADO (tools/config/{acc.periodo}.py). "
                    f"Las correcciones van en sitio, editando celdas. Con --forzar se puede "
                    f"pedir aprobación")

    # 4 — Respaldo previo. Solo exigible cuando de verdad se va a escribir Y hay algo que
    # perder: un archivo que aún no existe no tiene trabajo previo que destruir. Los tools que
    # escriben una revisión NUEVA fechada (integrar_conciliacion → CONCILIACION_..._REV_fecha)
    # usan un patrón más seguro que sobrescribir en sitio; exigirles un .bak sería castigarlo.
    if acc.aplicar and acc.maestro is not None and estado != DENEGAR:
        existe = Path(acc.maestro).exists()
        if existe and _bak_de_hoy(acc.maestro) is None:
            denegar(f"sin respaldo: no existe el .bak de hoy de «{Path(acc.maestro).name}». "
                    f"Llama a cierre.respaldar(ruta) ANTES de escribir")
        elif not existe:
            hallazgos.append(f"archivo nuevo: «{Path(acc.maestro).name}» no existía; sin respaldo "
                             f"que exigir")

    # 5 — Umbral de monto.
    if acc.monto is not None and acc.monto > UMBRAL_MOVIMIENTO:
        aprobar(f"monto: un movimiento de S/{acc.monto:,.2f} supera el umbral de "
                f"S/{UMBRAL_MOVIMIENTO:,.2f}")
    if acc.total is not None and acc.total > UMBRAL_CORRIDA:
        aprobar(f"monto: la corrida mueve S/{acc.total:,.2f}, sobre el umbral de "
                f"S/{UMBRAL_CORRIDA:,.2f}")

    # 6 — Evidencia citada. Cambiar el estado de un cobro o un pago sin fuente es lo que produce
    # cruces equivocados: el número parecía calzar y nadie podía rastrear por qué.
    if acc.aplicar and clave in ("cobros", "proveedores", "detracciones") and not acc.evidencia:
        aprobar(f"sin evidencia citada: escribir «{clave}» cambia el estado de un cobro o un pago; "
                f"cita la fuente (constancia, EVIDENCIA_*.csv, op del banco)")

    # 7 — Dry-run por defecto. Sin --aplicar no se escribe, nunca.
    #
    # Gana sobre «requiere aprobación», y por eso: una simulación NO ESCRIBE, así que pedir
    # permiso para ella no protege nada — y bloquea el preview justo en los movimientos grandes,
    # que son los únicos que uno querría ver antes de decidir. El efecto real de exigirlo era
    # que la única forma de mirar una corrida de S/60,000 fuera autorizarla.
    #
    # NO gana sobre DENEGAR: si la acción se va a denegar, la simulación debe decirlo. Un
    # preview que calla la denegación manda a preparar un trabajo que nunca va a entrar.
    # Los motivos se conservan: la simulación informa QUÉ aprobación haría falta.
    if not acc.aplicar and estado != DENEGAR:
        estado = SIMULACION

    return Decision(estado=estado, motivos=motivos, hallazgos=hallazgos,
                    evento=_evento(acc, estado, motivos, hallazgos, clave))


# ── Evento de auditoría ──────────────────────────────────────────────────────
def ruta_log(cuando: date | None = None) -> Path:
    d = cuando or date.today()
    return RAIZ / "08. HERRAMIENTAS" / "auditoria" / f"eventos-{d:%Y-%m}.jsonl"


def _evento(acc: Accion, estado, motivos, hallazgos, clave) -> dict:
    bak = _bak_de_hoy(acc.maestro)
    return {
        "ts": datetime.now().isoformat(timespec="seconds"),
        "tool": acc.tool,
        "agente": acc.agente,
        "empresa": acc.empresa,
        "maestro": Path(acc.maestro).name if acc.maestro else None,
        "clave": clave,
        "capacidad": acc.capacidad,
        "periodo": acc.periodo,
        "aplicar": acc.aplicar,
        "monto": acc.monto,
        "total": acc.total,
        "evidencia": acc.evidencia,
        "estado": estado,
        "motivos": list(motivos),
        "hallazgos": list(hallazgos),
        "bak": bak.name if bak else None,
    }


def _es_repo_real() -> bool:
    """¿RAIZ apunta al repo de verdad, o a una raíz de prueba?"""
    return RAIZ.resolve() == Path(__file__).resolve().parent.parent.parent


def registrar(dec: Decision, cuando: date | None = None) -> Path:
    """Anexa el evento al jsonl del mes. APPEND-ONLY: jamás se reescribe ni se ordena.

    Bajo pytest NO escribe si la raíz es el repo real: un test que ejercita un tool
    con --forzar ensuciaría el rastro de auditoría con eventos que nunca
    ocurrieron, y un log de auditoría con ficción dentro no sirve para auditar. Los tests del
    propio gate sí escriben, porque apuntan RAIZ a un tmp_path.
    """
    if os.environ.get("PYTEST_CURRENT_TEST") and _es_repo_real():
        return ruta_log(cuando)
    ruta = ruta_log(cuando)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with open(ruta, "a", encoding="utf-8") as fh:
        fh.write(json.dumps(dec.evento, ensure_ascii=False) + "\n")
    return ruta


# ── El que corta ─────────────────────────────────────────────────────────────
def exigir(dec: Decision, interactivo: bool | None = None) -> bool:
    """Registra el evento y aplica la decisión.

    Devuelve True si el tool puede escribir. Aborta el proceso si la acción se deniega.
    En `requiere_aprobacion` pide confirmación por terminal; si no hay terminal (launchd,
    sub-agente, CI) DENIEGA — pedir permiso a nadie no es pedir permiso.
    """
    registrar(dec)

    if dec.estado == DENEGAR:
        detalle = "\n".join(f"   · {m}" for m in dec.motivos)
        sys.exit(f"⛔ policy: escritura DENEGADA en {dec.evento.get('tool')}\n{detalle}\n"
                 f"   (registrado en {ruta_log().relative_to(RAIZ)})")

    for h in dec.hallazgos:
        print(f"⚠️  policy: {h}")

    if dec.estado == APROBACION:
        detalle = "\n".join(f"   · {m}" for m in dec.motivos)
        print(f"🟠 policy: esta escritura REQUIERE APROBACIÓN\n{detalle}")
        if interactivo is None:
            interactivo = sys.stdin.isatty()
        if not interactivo:
            sys.exit("⛔ policy: requiere aprobación y no hay terminal para pedirla. "
                     "Re-córrelo de forma interactiva.")
        if input("   ¿Autorizas la escritura? [escribe SI] ").strip().upper() != "SI":
            sys.exit("⛔ policy: no autorizada.")
        return True

    if dec.estado == SIMULACION:
        print(f"🔍 policy: simulación (sin --aplicar) — no se escribe nada.")
        return False

    return True
