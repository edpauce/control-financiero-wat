"""
tools/config/2026-07.py — Plantilla de la config de un cierre mensual (valores de ejemplo).

Aquí vive SOLO lo específico del mes: saldos al corte y las decisiones que una persona tomó
sobre movimientos concretos. Todo campo ausente toma su default en lib/cierre.py (_DEFAULTS).
"""

CERRADO = False          # True → el mes lleva formato manual: no se regenera sin --forzar
CORTE = None             # None → último día del mes; "15/07" para un cierre quincenal
REV_FECHA = "2026-07-31"

# Saldos de las cuentas de detracciones (Banco de la Nación) al corte.
SALDO_BN = {"MERAK": 40000.00, "ALTAMIRA": 1500.00}

# Reclasificaciones por número de operación bancaria, con su observación.
OP_OVERRIDES = {
    "01000001": "PAGO PROVEEDOR / TERCERO",
}
OP_OBS = {
    "01000001": "Certificación anual — cuota 1 de 2",
}
OP_CONTRAPARTE = {}

# Reglas por glosa propias del mes; se fusionan con GLOSA_PERMANENTES de glosario.py.
GLOSA_OVERRIDES = []

# Cobros confirmados con evidencia: comprobante → (fecha, monto).
CONFIRMADOS = {}
COBRADO_MANUAL = {}
ANULAR_MANUAL = []

AUTODEP_DOCS = set()     # documentos BN que son auto-depósito de la empresa
NOTA_09 = ""
