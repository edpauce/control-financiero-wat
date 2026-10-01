"""
Respaldo antes de escribir un maestro en sitio (lib.cierre.respaldar).

La regla: un slot por etiqueta o por día, y NUNCA se pisa. El primer respaldo del día es el
estado previo a cualquier corrida — el que hace falta cuando la segunda corrida rompe algo.
"""
import datetime

from lib import cierre as CI


def test_respalda_una_vez_por_dia(tmp_path):
    maestro = tmp_path / "COBROS_2026-08.xlsx"
    maestro.write_bytes(b"version 1")
    hoy = datetime.date.today().isoformat()

    bak = CI.respaldar(maestro)
    assert bak == tmp_path / f"COBROS_2026-08.xlsx.{hoy}.bak"
    assert bak.read_bytes() == b"version 1"

    # segunda corrida el mismo día: el maestro ya cambió, pero el .bak conserva el estado previo
    maestro.write_bytes(b"version 2 (tras la primera corrida)")
    bak2 = CI.respaldar(maestro)
    assert bak2 == bak
    assert bak.read_bytes() == b"version 1", "el slot del día no se pisa"
    assert sorted(p.name for p in tmp_path.iterdir()) == [
        "COBROS_2026-08.xlsx", f"COBROS_2026-08.xlsx.{hoy}.bak"]


def test_respalda_por_etiqueta(tmp_path):
    base = tmp_path / "SEGUIMIENTO_DETRACCIONES_2026.xlsx"
    base.write_bytes(b"base")
    bak = CI.respaldar(base, "2026-07")
    assert bak.name == "SEGUIMIENTO_DETRACCIONES_2026.xlsx.2026-07.bak"
    assert bak.read_bytes() == b"base"
    base.write_bytes(b"con julio")
    assert CI.respaldar(base, "2026-07").read_bytes() == b"base"     # mismo slot, no se pisa
    assert CI.respaldar(base, "2026-08").read_bytes() == b"con julio"  # otro periodo, otro slot


def test_acepta_ruta_como_texto(tmp_path):
    maestro = tmp_path / "CREDITOS_2026.xlsx"
    maestro.write_bytes(b"x")
    bak = CI.respaldar(str(maestro), "prueba")
    assert bak.exists() and bak.parent == tmp_path
