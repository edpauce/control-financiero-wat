"""
tools/lib/empresas.py — Datos maestros de las empresas del grupo.

Fuente única para cuentas, RUC y reglas de detracciones (SPOT). Los valores de este archivo
son FICTICIOS: ilustran la forma de los datos, no describen a ninguna empresa real.
"""
from __future__ import annotations

EMPRESAS = ["ALTAMIRA", "MERAK"]

# Nombres completos
NOMBRE = {
    "ALTAMIRA": "ALTAMIRA CONSULTING S.A.C.",
    "MERAK": "MERAK CERTIFICACIONES S.A.C.",
}

# Cuenta corriente BCP (soles)
CUENTA_BCP = {
    "ALTAMIRA": "193-1000101-0-11",
    "MERAK":    "193-2000201-0-21",
}

# TODAS las cuentas BCP del grupo — no solo la operativa. Mientras las secundarias existían solo
# como prosa en la documentación, sus movimientos se clasificaban como PAGO PROVEEDOR = gasto:
# el código no tenía con qué saber que la contraparte era uno mismo. Una regla que depende de
# una lista solo vale si la lista está en el código.
# (cuenta, fragmento, moneda, papel). El FRAGMENTO es lo que el extracto escribe en
# «A 193 1000102 0».
CUENTAS_GRUPO = {
    "ALTAMIRA": [("193-1000101-0-11",  "1000101",  "PEN", "operativa"),
                 ("193-1000102-0-12",  "1000102",  "PEN", "secundaria"),
                 ("193-10001031-1-13", "10001031", "USD", "ahorro dólares")],
    "MERAK":    [("193-2000201-0-21",  "2000201",  "PEN", "operativa"),
                 ("193-2000202-0-22",  "2000202",  "PEN", "secundaria"),
                 ("193-2000203-1-23",  "2000203",  "USD", "dólares")],
}


def cuenta_del_grupo(texto: str):
    """`(empresa, cuenta, moneda, papel)` si `texto` nombra una cuenta NUESTRA; si no, None.

    Sirve para las dos puntas: el número completo de la cabecera del extracto
    («193-2000201-0-21 - MERAK …») y el fragmento suelto de una glosa de cargo
    («A 193 2000202 0»). Por eso compara contra el fragmento, que es lo común a ambas.
    """
    plano = "".join(ch for ch in str(texto or "") if ch.isdigit())
    if not plano:
        return None
    for emp, ctas in CUENTAS_GRUPO.items():
        for cta, frag, moneda, papel in ctas:
            if frag in plano:
                return emp, cta, moneda, papel
    return None


# Fragmento de la cuenta usado para detectar transferencias inter-compañía en el EECC
REF_INTERCIA = {
    "ALTAMIRA": "1000101",   # un cargo de MERAK con esta ref va hacia ALTAMIRA
    "MERAK":    "2000201",   # un cargo de ALTAMIRA con esta ref va hacia MERAK
}

# Cuenta de detracciones (Banco de la Nación)
CUENTA_BN = {
    "ALTAMIRA": "00-000-100101",
    "MERAK":    "00-000-200201",
}

RUC = {
    "ALTAMIRA": "20900000011",
    "MERAK":    "20900000021",
}
# Último dígito del RUC → define el cronograma de vencimiento SUNAT
ULTIMO_DIGITO_RUC = {
    "ALTAMIRA": 1,
    "MERAK":    1,
}
# Régimen tributario (RMT = Régimen MYPE Tributario)
REGIMEN = {
    "ALTAMIRA": "RMT",
    "MERAK":    "RMT",
}

# ─── Reglas de detracciones (SPOT) ───────────────────────────────────────────
TASA_SPOT = 0.12    # 12% general sobre el total del comprobante (servicios sujetos a SPOT)
UMBRAL_SPOT = 700   # NO aplica a operaciones <= S/ 700, exoneradas, exportación, ni NC.


def detr_esperada(total: float) -> float:
    """Detracción esperada (12%) para un comprobante, 0 si no supera el umbral SPOT."""
    return round(total * TASA_SPOT, 2) if total > UMBRAL_SPOT else 0.0

# DEPRECATED — NO usar para decidir quién depositó una detracción.
# Los códigos del EECC BN describen el CANAL/OFICINA del movimiento, no al depositante: hubo
# abonos "1234" que resultaron ser auto-detracciones nuestras (salida PAGO DETRAC en BCP el día
# anterior) y depósitos "78788888"/"EE 0456" que resultaron ser del cliente. Se conserva solo
# como glosario de canales para leer el EECC. Para la autoría usar quien_deposito() más abajo.
CODIGOS_BN = {
    "VA 1234":  "canal_oficina_1234",   # ventanilla/oficina 1234 — NO implica depósito del cliente
    "78788888": "canal_virtual",        # canal virtual SUNAT-BCP — NO implica auto-detracción
    "EE 0456":  "efectivo_agencia",     # efectivo en agencia — NO implica auto-detracción
}

DEPOSITO_EMPRESA = "empresa"            # auto-detracción: la pagamos nosotros
DEPOSITO_CLIENTE = "cliente"            # el cliente retuvo el 12% y lo depositó él
DEPOSITO_POR_CONFIRMAR = "por_confirmar"  # 🟠 sin evidencia suficiente: no cerrar


def quien_deposito(salida_bcp: bool, rol_constancia: str | None) -> str:
    """Quién depositó una detracción abonada en el BN: 'empresa', 'cliente' o 'por_confirmar'.

    Los códigos del EECC BN (VA 1234 / 78788888 / EE 0456) NO dicen quién depositó. Lo único que
    discrimina, en este orden:
      1. `salida_bcp`: existe el cargo `PAGO DETRAC <op>` en la conciliación BCP → la pagamos
         NOSOTROS por el canal virtual SUNAT-BCP (postea al BN 1-2 días después). Es el test
         definitivo; manda sobre cualquier otra lectura.
      2. `rol_constancia`: rol de NUESTRA empresa en la constancia SUNAT — 'proveedor' (la depositó
         la empresa, con nuestro Usuario SOL) o 'adquiriente' (la depositó el cliente).
    Sin salida BCP y sin constancia no se concluye nada: un depósito en efectivo (ENT.EFEC) hecho
    por nosotros tampoco deja rastro en el BCP, así que "no hay PAGO DETRAC" no prueba que fue el
    cliente → 'por_confirmar' (🟠), nunca 'cliente' por descarte. Y ojo: un extracto BCP parcial
    tampoco prueba ausencia de la salida (puede cortar antes de la fecha del cargo).
    """
    if salida_bcp:
        return DEPOSITO_EMPRESA
    rol = (rol_constancia or "").strip().lower()
    if rol.startswith("prov"):
        return DEPOSITO_EMPRESA
    if rol.startswith("adq"):
        return DEPOSITO_CLIENTE
    return DEPOSITO_POR_CONFIRMAR


def empresa_por_ruc_ref(ref: str) -> str | None:
    """Dada una referencia BCP, dice hacia qué empresa apunta (o None)."""
    for emp, frag in REF_INTERCIA.items():
        if frag in str(ref or ""):
            return emp
    return None
