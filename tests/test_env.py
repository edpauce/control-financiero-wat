"""
lib/env.cargar_env(): parser dotenv compartido. La regresión que cubre: los parsers
viejos cortaban el valor en el primer `#`, así que una contraseña con `#` llegaba truncada
al portal.
"""
from lib import env as ENV


def test_password_con_numeral_no_se_trunca(tmp_path):
    f = tmp_path / ".env"
    f.write_text("EFACT_ALT_USER=alt@empresa.pe\nEFACT_ALT_PASS=abc#123#x\n", encoding="utf-8")
    vals = ENV.cargar_env(str(f))
    assert vals["EFACT_ALT_PASS"] == "abc#123#x"
    assert vals["EFACT_ALT_USER"] == "alt@empresa.pe"


def test_comentario_en_linea_solo_tras_espacio(tmp_path):
    f = tmp_path / ".env"
    f.write_text(
        "A=valor # comentario\n"
        "B=      # p.ej. 20900000011USUARIO1\n"     # así vienen las plantillas de .env.example
        "C=\"con # dentro\"\n"
        "D='p#a ss'\n"
        "export E=exportada\n"
        "# F=comentada\n"
        "linea sin igual\n",
        encoding="utf-8")
    assert ENV.cargar_env(str(f)) == {"A": "valor", "B": "", "C": "con # dentro",
                                      "D": "p#a ss", "E": "exportada"}


def test_archivo_inexistente_devuelve_vacio(tmp_path):
    assert ENV.cargar_env(str(tmp_path / "no_existe.env")) == {}
