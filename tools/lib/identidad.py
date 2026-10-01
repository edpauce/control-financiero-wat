"""
tools/lib/identidad.py — ¿Esta glosa del banco nombra a este cliente?

POR QUÉ NO SIRVE UNA LISTA DE PALABRAS GENÉRICAS
Tuve una (`SERVICIOS`, `INVERSIONES`, `EMPRESA`…) y necesitaba un parche por cada familia
nueva: se escapó «EMPRESA DE TRANSPORTES», que calza con varios clientes distintos.

POR QUÉ TAMPOCO SIRVE «SI CALZA CON DOS CLIENTES, NO NOMBRA A NADIE» POR SUBCADENA
La idea es correcta; medirla por subcadena común la voltea, porque las cadenas largas que
comparten dos empresas son justamente las palabras corporativas. Medido sobre el maestro
de clientes: **más de la mitad de los abonos nombrados perdían el nombre.** «DE J K L INGENIEROS E»
chocaba con otros tres por `INGENIER`, deshaciendo el arreglo que acababa de hacerse.

LO QUE SÍ SE DERIVA DE LOS DATOS: LA GENERICIDAD ES DEL TOKEN, NO DEL DUEÑO
Un token que aparece en el nombre de dos o más clientes distintos no aporta identidad —
SERVICIOS, INGENIEROS, SOCIEDAD, PERU, TRANSPORTES salen solos del maestro, sin escribirlos.
La coincidencia tiene que venir de un token que SÍ aporta identidad. Así «DE SERVICIOS
LOGISTICO» sigue nombrando a Servicios Logísticos aunque Servicios Generales Mateo exista.

EL UNIVERSO ES EL MAESTRO ENTERO, NO LOS PENDIENTES DEL DÍA
Si la tabla se derivara de quién tiene facturas abiertas hoy, la misma glosa nombraría a JKL el
lunes y a nadie el martes. La identidad de un nombre es una propiedad de la cartera, no del
calendario.

DOS CLIENTES SON EL MISMO SI COMPARTEN RUC
«NOVA PRISMA S.A.C.» y «NOVA PRISMA SOCIEDAD ANONIMA» son uno solo; contarlos como dos hacía
genérico su propio nombre. El RUC manda cuando el maestro lo trae — es la misma regla que ya
se aplicaba a los clientes que comparten una razón social genérica.
"""
from __future__ import annotations

import re
from collections import Counter

MINIMO_TOKEN = 3        # letras para que un token cuente
MINIMO_PREFIJO = 4      # letras compartidas para dar dos tokens por el mismo


def tokens(texto: str) -> list[str]:
    """Palabras de ≥3 letras, en mayúscula, uniendo las corridas de letras sueltas.

    El banco escribe «J K L INGENIEROS» y el maestro «JKL INGENIEROS E.I.R.L.». Sin unir
    «J K L» en «JKL», el abono de JKL no nombra a JKL — que es el caso que más duele, porque
    tiene el nombre delante.
    """
    palabras = re.findall(r"[A-ZÑ]+", str(texto or "").upper())
    out: list[str] = []
    corrida: list[str] = []
    for w in palabras:
        if len(w) == 1:
            corrida.append(w)
            continue
        if corrida:
            out.append("".join(corrida))
            corrida = []
        out.append(w)
    if corrida:
        out.append("".join(corrida))
    return [w for w in out if len(w) >= MINIMO_TOKEN]


def casan(a: str, b: str) -> bool:
    """Dos tokens son el mismo si son iguales, o si el corto prefija al largo con ≥4 letras.

    El prefijo tolera el truncado del banco (30 caracteres) y los plurales (LOGISTICO ↔
    LOGISTICOS). La igualdad va aparte y sin mínimo: exigir 4 letras también para los iguales
    hacía que «LOG» no casara con «LOG», y LOG FLETES perdía su único token distintivo.
    """
    if a == b:
        return True
    corto, largo = (a, b) if len(a) <= len(b) else (b, a)
    return len(corto) >= MINIMO_PREFIJO and largo.startswith(corto)


def indice(clientes) -> Counter:
    """{token: nº de clientes DISTINTOS que lo llevan}. `clientes` es (clave, nombre); la clave
    es el RUC cuando se conoce, para que dos razones sociales del mismo RUC cuenten una vez."""
    por_token: dict[str, set] = {}
    for clave, nombre in clientes:
        for t in set(tokens(nombre)):
            por_token.setdefault(t, set()).add(clave)
    # un token también carga con los que lo prefijan: LOGISTICO y LOGISTICOS son el mismo
    idx = Counter()
    for t, claves in por_token.items():
        juntas = set(claves)
        for otro, otras in por_token.items():
            if otro != t and casan(t, otro):
                juntas |= otras
        idx[t] = len(juntas)
    return idx


def en_cartera(token: str, idx: Counter) -> bool:
    """¿Este token existe en algún nombre de cliente?

    Es la condición que faltaba. Sin ella `identifica()` devolvía True por DEFECTO para
    cualquier palabra desconocida, y así OTRA y CUENTA «identificaban»: «DE OTRA CUENTA» —una
    glosa del BCP, no un pagador— se clasificaba como si nombrara a alguien. Un token que no
    aparece en ninguna parte de la cartera no es identidad, es ruido.
    """
    return token in idx or any(casan(t, token) for t in idx)


def identifica(token: str, idx: Counter) -> bool:
    """¿Este token señala a UN cliente de la cartera? Tiene que ESTAR en ella y llevar a uno."""
    if not en_cartera(token, idx):
        return False
    n = idx.get(token)
    if n is None:
        n = max((v for t, v in idx.items() if casan(t, token)), default=1)
    return n <= 1


ANONIMA, NOMBRA = "anonima", "nombra"


def _alinea(glosa_tk: list[str], cliente_tk: list[str]) -> bool:
    """¿La glosa empieza donde empieza el nombre del cliente?

    El banco escribe la razón social del pagador desde el principio y la corta a 30 caracteres.
    Por eso el primer token tiene que coincidir: es lo que separa «DE HIDROSUR SAC» (que ES
    HIDROSUR S.A.C.) de «DE MINING COMPANY SERV» (que comparte una palabra con CONDOR MINING
    PERU y no es ella).
    """
    return bool(glosa_tk) and bool(cliente_tk) and casan(glosa_tk[0], cliente_tk[0])


def clasificar(glosa: str, clientes, idx: Counter | None = None) -> tuple[str, set]:
    """`(ANONIMA, ∅)` o `(NOMBRA, {claves})` — y la diferencia es la que decide.

    Tres situaciones, tres consecuencias distintas, y confundirlas es lo que hacía que una
    glosa perdiera el nombre:

      · **Ningún token de la glosa aparece en la cartera** → NOMBRA a alguien de fuera. Aquí
        caen el tercero real («DE TRAMO SAC»), la empresa hermana («DE ALTAMIRA CONSULTING») y
        también los literales del banco («DE OTRA CUENTA»), que no son un pagador. No se puede
        distinguir uno de otro sin un padrón de alias, así que los tres quedan BLOQUEADOS para
        cerrar y marcados. Es conservador a propósito: el costo es no cerrar algo que sí era
        del cliente; el de equivocarse al revés es cerrar la factura de otro.
      · **Con tokens en la cartera pero todos genéricos** («DE SERVICIOS» a secas,
        «DE SOCIEDAD INDUSTRIAL») → ANÓNIMA. Decide el importe, como siempre.
      · **Con token identificativo** → NOMBRA. Y la atribución la hace el PRIMER token
        identificativo, no cualquiera: el banco escribe el nombre del pagador al principio.
        Sin eso, «DE MINING COMPANY SERV» repartía identidad entre CONDOR MINING y EVEREST HT
        COMPANY por una palabra suelta, y con un solo choque habría atribuido el abono a
        Everest y podría haber cerrado una factura suya.
    """
    clientes = list(clientes)
    if idx is None:
        idx = indice(clientes)
    tk = tokens(glosa)
    gl = [t for t in tk if identifica(t, idx)]
    if gl:
        # DOS condiciones, no una. Compartir un token identificativo no basta: «DE MINING
        # COMPANY SERV» comparte MINING con CONDOR MINING PERU y con eso solo se le habría
        # atribuido el abono —y cerrado una factura suya por una palabra—. La glosa del banco
        # es el nombre del pagador TRUNCADO, así que tiene que ALINEAR DESDE EL PRINCIPIO:
        # «DE HIDROSUR SAC» empieza donde empieza HIDROSUR S.A.C.; «MINING» no empieza donde
        # empieza CONDOR. Eso distingue un nombre de una coincidencia de palabra.
        return NOMBRA, {clave for clave, nombre in clientes
                        if _alinea(tk, tokens(nombre))
                        and any(casan(g, c) for g in gl for c in tokens(nombre))}
    # Sin ningún token que identifique. Lo que decide ahora es si queda alguna palabra que la
    # cartera NO conozca: «DE TRAMO SAC» solo comparte SAC (genérico) y TRAMO no le suena a
    # nadie — eso es un pagador de fuera. «DE SOCIEDAD INDUSTRIAL» en cambio es todo palabras
    # que la cartera usa y ninguna distingue: no sabemos quién es, pero tampoco hay señal de
    # que sea ajeno, así que decide el importe.
    if any(not en_cartera(t, idx) for t in tk):
        return NOMBRA, set()                     # nombra a alguien de fuera, hermana o banco
    return ANONIMA, set()


def duenos(glosa: str, clientes, idx: Counter | None = None) -> set:
    """Las CLAVES de los clientes a los que esta glosa puede estar nombrando.

    Exige compartir al menos un token que identifique. Si no hay ninguno, la glosa no nombra a
    nadie y el llamador decide por importe, como con «TRANSF.BCO.BBVA».
    """
    clientes = list(clientes)
    if idx is None:
        idx = indice(clientes)
    gl = [t for t in tokens(glosa) if identifica(t, idx)]
    if not gl:
        return set()
    out = set()
    for clave, nombre in clientes:
        ct = tokens(nombre)
        if any(casan(g, c) for g in gl for c in ct):
            out.add(clave)
    return out
