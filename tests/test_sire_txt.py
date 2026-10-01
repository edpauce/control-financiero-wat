"""
Parser de TXT SIRE (11/08) — el contador declara el RVIE/RCE en formato pipe; es la fuente
oficial (crédito fiscal en compras). Mapea por NOMBRE de columna porque RVIE y RCE difieren.
"""
import importlib.util
import os

_ruta = os.path.join(os.path.dirname(__file__), "..", "tools", "lib", "sire_txt.py")
_spec = importlib.util.spec_from_file_location("sire_txt", _ruta)
ST = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ST)

# encabezado RCE (compras) recortado a lo esencial + una fila
RCE_HDR = ("RUC|Apellidos y Nombres o Razón social|Periodo|CAR SUNAT|Fecha de emisión|"
           "Fecha Vcto/Pago|Tipo CP/Doc.|Serie del CDP|Año|Nro CP o Doc. Nro Inicial (Rango)|"
           "Nro Final (Rango)|Tipo Doc Identidad|Nro Doc Identidad|Apellidos Nombres/ Razón  Social|"
           "BI Gravado DG|IGV / IPM DG|BI Gravado DGNG|IGV / IPM DGNG|BI Gravado DNG|IGV / IPM DNG|"
           "Valor Adq. NG|ISC|ICBPER|Otros Trib/ Cargos|Total CP|Moneda|Tipo de Cambio")
RCE_FILA = ("20900000011|ALTAMIRA CONSULTING S.A.C.|202606|CAR|25/06/2026||01|F008||146235||6|"
            "20900000043|ANDINA SEGUROS S.A.|149.00|26.82|0|0|0|0|0|0|0|0|175.82|PEN|1.000")
RCE_NC = ("20900000011|ALTAMIRA|202606|CAR|26/06/2026||07|F008||146236||6|20900000043|"
          "ANDINA SEGUROS S.A.|-49.00|-8.82|0|0|0|0|0|0|0|0|-57.82|PEN|1.000")

# encabezado RVIE (ventas) — sin 'Año', con 'BI Gravada' y 'Valor Facturado Exportación'
RVIE_HDR = ("Ruc|Razon Social|Periodo|CAR SUNAT|Fecha de emisión|Fecha Vcto/Pago|Tipo CP/Doc.|"
            "Serie del CDP|Nro CP o Doc. Nro Inicial (Rango)|Nro Final (Rango)|Tipo Doc Identidad|"
            "Nro Doc Identidad|Apellidos Nombres/ Razón Social|Valor Facturado Exportación|"
            "BI Gravada|Dscto BI|IGV / IPM|Dscto IGV / IPM|Mto Exonerado|Mto Inafecto|ISC|"
            "BI Grav IVAP|IVAP|ICBPER|Otros Tributos|Total CP|Moneda")
RVIE_FILA = ("20900000011|ALTAMIRA|202606|CAR|25/06/2026|30/06/2026|01|F001|1797||6|20900000036|"
             "INSTITUTO DE FORMACION|0|10169.49|0|1830.51|0|0|0|0|0|0|0|0|12000|PEN")


def _escribir(tmp_path, hdr, *filas):
    p = tmp_path / "sire.txt"
    p.write_text(hdr + "\n" + "\n".join(filas) + "\n", encoding="latin-1")
    return str(p)


def test_detecta_rce_y_parsea_proveedor_igv(tmp_path):
    libro, regs = ST.parse(_escribir(tmp_path, RCE_HDR, RCE_FILA))
    assert libro == "RCE"
    r = regs[0]
    assert r["comp"] == "F008-146235"
    assert r["ruc"] == "20900000043"                     # el PROVEEDOR, no ALTAMIRA
    assert r["contraparte"].startswith("ANDINA SEGUROS")
    assert r["igv"] == 26.82 and r["total"] == 175.82
    assert not r["es_nc"]


def test_detecta_rvie_y_no_confunde_titular_con_cliente(tmp_path):
    libro, regs = ST.parse(_escribir(tmp_path, RVIE_HDR, RVIE_FILA))
    assert libro == "RVIE"
    r = regs[0]
    assert r["comp"] == "F001-1797"
    assert r["ruc"] == "20900000036"                     # el CLIENTE
    assert r["contraparte"].startswith("INSTITUTO")
    assert r["igv"] == 1830.51 and r["total"] == 12000.0


def test_totales_resta_notas_de_credito(tmp_path):
    _, regs = ST.parse(_escribir(tmp_path, RCE_HDR, RCE_FILA, RCE_NC))
    t = ST.totales(regs)
    assert t["n"] == 2
    assert t["igv"] == round(26.82 - 8.82, 2)            # la NC resta
    assert t["total"] == round(175.82 - 57.82, 2)


def test_txt_vacio_no_revienta(tmp_path):
    p = tmp_path / "v.txt"; p.write_text("", encoding="latin-1")
    assert ST.parse(str(p)) == ("DESCONOCIDO", [])
