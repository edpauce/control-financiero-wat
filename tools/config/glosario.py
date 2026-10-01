"""
tools/config/glosario.py — Reglas de glosa PERMANENTES (valores de ejemplo).

(palabras clave, categoría, observación) que aplican a TODO mes. lib.cierre.cargar_config las
fusiona con las GLOSA_OVERRIDES del mes; una regla del mes que comparta palabra clave manda
sobre la permanente.
"""

GLOSA_PERMANENTES = [
    (("TIPO DE CAMBIO", "T/C"), "OTROS GASTOS", "Operación de tipo de cambio (compra/venta USD)"),
    (("MANTENIMIENTO",), "OTROS GASTOS", "Comisión de mantenimiento de cuenta"),
]
