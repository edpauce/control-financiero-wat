#!/usr/bin/env python3
"""
tools/sellar_mes.py — Sella (o re-sella) la integridad de los meses del maestro de
conciliación más reciente. Edita SOLO la hoja oculta '🔒 Sellos', en sitio.

Sellar = "estos números del banco son los definitivos". La auditoría y cada cierre
verifican los sellos y gritan si un mes sellado cambió de números. Re-correr este tool
es el acto CONSCIENTE de aceptar un cambio legítimo (p.ej. una recategorización aprobada
que tocó montos — no debería pasar — o un mes re-cerrado con fuentes corregidas).

Uso:  python3 tools/sellar_mes.py                # sella todos los meses con filas
      python3 tools/sellar_mes.py --mes 2026-07  # sella solo julio
"""
import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl                      # noqa: E402
from lib import cierre as CI         # noqa: E402
from lib import policy as POL
from lib import sellos as SE         # noqa: E402
from lib import sources as S         # noqa: E402


def main():
    meses = None
    if "--mes" in sys.argv:
        periodo = sys.argv[sys.argv.index("--mes") + 1]
        meses = [SE.MESES_ES[int(periodo[5:7]) - 1]]
    ruta = S.latest("conciliacion")
    if not ruta:
        sys.exit("❌ No hay maestro de conciliación.")
    CI.respaldar(ruta)
    wb = openpyxl.load_workbook(ruta)
    previos = SE.verificar(wb)
    sellados = SE.sellar(wb, meses, cuando=datetime.date.today().isoformat())
    if not POL.exigir(POL.evaluar(POL.Accion(
            tool="sellar_mes", maestro=ruta, empresa=None, capacidad="conciliacion.escribir",
            periodo=POL.periodo_de(ruta), aplicar=True, evidencia="huella sha256 de los movimientos del banco",
            forzar="--forzar" in sys.argv))):
        sys.exit("(no se escribió)")
    wb.save(ruta)
    if previos:
        print("⚠️ Cambios aceptados al re-sellar (eran violaciones):")
        for v in previos:
            print("   ", v.split(" — ")[0])
    print(f"🔒 Sellados {len(sellados)} (empresa × mes) en {os.path.basename(ruta)}:")
    for emp, mes, sha in sellados:
        print(f"    {emp[:3]} {mes:<10} {sha}")


if __name__ == "__main__":
    main()
