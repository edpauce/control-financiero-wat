"""
Tests de lib/comprobantes.py — parsers de comprobantes recibidos (XML UBL y RHE PDF).
Todo sintético: fixtures UBL mínimos construidos inline, texto de RHE imitando el orden
DESORDENADO que pypdf extrae de los PDFs reales de SUNAT.
"""
import datetime

import pytest

from lib import comprobantes as CP

CBC = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"
CAC = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"


def _party(tag, ruc, nombre):
    return f"""<cac:{tag}>
      <cac:Party>
        <cac:PartyIdentification><cbc:ID schemeID="6">{ruc}</cbc:ID></cac:PartyIdentification>
        <cac:PartyLegalEntity><cbc:RegistrationName>{nombre}</cbc:RegistrationName></cac:PartyLegalEntity>
      </cac:Party>
    </cac:{tag}>"""


def factura_xml(emisor="20100047218", emisor_nom="BANCO DE CREDITO DEL PERU",
                adquirente="20900000011", adq_nom="ALTAMIRA CONSULTING S.A.C.",
                total="44.00", igv="6.71", subtotal="37.29", moneda="PEN",
                serie="FI01-20000581", fecha="2026-07-07", tipo_code="01",
                concepto="CONCEPTO DE PAGO:PORTES"):
    """Factura UBL 2.1 mínima con la trampa real: <Signature> también trae un ID."""
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
         xmlns:cbc="{CBC}" xmlns:cac="{CAC}">
  <cbc:UBLVersionID>2.1</cbc:UBLVersionID>
  <cbc:ID>{serie}</cbc:ID>
  <cbc:IssueDate>{fecha}</cbc:IssueDate>
  <cbc:DueDate>2026-06-01</cbc:DueDate>
  <cbc:InvoiceTypeCode>{tipo_code}</cbc:InvoiceTypeCode>
  <cbc:DocumentCurrencyCode>{moneda}</cbc:DocumentCurrencyCode>
  <cac:Signature>
    <cbc:ID>SignatureID</cbc:ID>
    <cac:SignatoryParty>
      <cac:PartyIdentification><cbc:ID>99999999999</cbc:ID></cac:PartyIdentification>
    </cac:SignatoryParty>
  </cac:Signature>
  {_party("AccountingSupplierParty", emisor, emisor_nom)}
  {_party("AccountingCustomerParty", adquirente, adq_nom)}
  <cac:PaymentTerms><cbc:ID>FormaPago</cbc:ID><cbc:PaymentMeansID>Contado</cbc:PaymentMeansID></cac:PaymentTerms>
  <cac:TaxTotal><cbc:TaxAmount currencyID="{moneda}">{igv}</cbc:TaxAmount></cac:TaxTotal>
  <cac:LegalMonetaryTotal>
    <cbc:LineExtensionAmount currencyID="{moneda}">{subtotal}</cbc:LineExtensionAmount>
    <cbc:PayableAmount currencyID="{moneda}">{total}</cbc:PayableAmount>
  </cac:LegalMonetaryTotal>
  <cac:InvoiceLine>
    <cbc:ID>001</cbc:ID>
    <cac:TaxTotal><cbc:TaxAmount currencyID="{moneda}">0.00</cbc:TaxAmount></cac:TaxTotal>
    <cac:Item><cbc:Description>{concepto}</cbc:Description></cac:Item>
  </cac:InvoiceLine>
</Invoice>"""


NC_XML = f"""<?xml version="1.0" encoding="UTF-8"?>
<CreditNote xmlns="urn:oasis:names:specification:ubl:schema:xsd:CreditNote-2"
            xmlns:cbc="{CBC}" xmlns:cac="{CAC}">
  <cbc:ID>F001-00000160</cbc:ID>
  <cbc:IssueDate>2026-06-12</cbc:IssueDate>
  <cbc:DocumentCurrencyCode>PEN</cbc:DocumentCurrencyCode>
  {_party("AccountingSupplierParty", "20900000021", "MERAK CERTIFICACIONES S.A.C.")}
  {_party("AccountingCustomerParty", "20512345678", "CLIENTE SAC")}
  <cac:TaxTotal><cbc:TaxAmount currencyID="PEN">1372.88</cbc:TaxAmount></cac:TaxTotal>
  <cac:LegalMonetaryTotal>
    <cbc:LineExtensionAmount currencyID="PEN">7627.12</cbc:LineExtensionAmount>
    <cbc:PayableAmount currencyID="PEN">9000.00</cbc:PayableAmount>
  </cac:LegalMonetaryTotal>
</CreditNote>"""


# ─── parse_ubl ───────────────────────────────────────────────────────────────
def test_parse_ubl_factura_completa():
    d = CP.parse_ubl(factura_xml())
    assert d["tipo"] == "Factura"
    assert d["serie_numero"] == "FI01-20000581"
    assert d["fecha_emision"] == datetime.date(2026, 7, 7)
    assert d["fecha_venc"] == datetime.date(2026, 6, 1)
    assert d["moneda"] == "PEN"
    assert d["emisor_ruc"] == "20100047218"
    assert d["emisor_nombre"] == "BANCO DE CREDITO DEL PERU"
    assert d["adquirente_ruc"] == "20900000011"
    assert d["total"] == 44.00
    assert d["igv"] == 6.71
    assert d["subtotal"] == 37.29
    assert d["concepto"] == "PORTES"          # el prefijo 'CONCEPTO DE PAGO:' se limpia
    assert d["forma_pago"] == "CONTADO"


def test_parse_ubl_no_confunde_id_de_la_firma():
    """El <Signature> trae su propio ID y PartyIdentification: no deben pisar nada."""
    d = CP.parse_ubl(factura_xml())
    assert d["serie_numero"] != "SignatureID"
    assert d["emisor_ruc"] != "99999999999"


def test_parse_ubl_usd():
    d = CP.parse_ubl(factura_xml(moneda="USD", total="19.51", subtotal="19.51", igv="0.00"))
    assert d["moneda"] == "USD"
    assert d["total"] == 19.51


def test_parse_ubl_nota_credito():
    d = CP.parse_ubl(NC_XML)
    assert d["tipo"] == "Nota de Crédito"
    assert d["emisor_ruc"] == "20900000021"
    assert d["total"] == 9000.00


def test_parse_ubl_boleta():
    d = CP.parse_ubl(factura_xml(tipo_code="03"))
    assert d["tipo"] == "Boleta"


def test_parse_ubl_acepta_bytes():
    d = CP.parse_ubl(factura_xml().encode("utf-8"))
    assert d["serie_numero"] == "FI01-20000581"


# ─── clasificar ──────────────────────────────────────────────────────────────
def test_clasificar_cxp_por_adquirente():
    d = CP.parse_ubl(factura_xml(adquirente="20900000011"))
    assert CP.clasificar(d) == ("cxp", "ALTAMIRA")
    d = CP.parse_ubl(factura_xml(adquirente="20900000021"))
    assert CP.clasificar(d) == ("cxp", "MERAK")


def test_clasificar_emitido_propio_gana_al_adquirente():
    """Si el EMISOR es del grupo es una VENTA nuestra — aunque el adquirente fuera del grupo."""
    d = CP.parse_ubl(factura_xml(emisor="20900000021", adquirente="20900000011"))
    assert CP.clasificar(d) == ("emitido_propio", "MERAK")


def test_clasificar_ajeno():
    """Llega al mismo buzón pero el adquirente no es del grupo."""
    d = CP.parse_ubl(factura_xml(adquirente="20999999999", adq_nom="OTRA EMPRESA SAC"))
    assert CP.clasificar(d) == ("ajeno", None)


# ─── parse_rhe_texto ─────────────────────────────────────────────────────────
RHE_TEXTO = """DEL ARTÍCULO 33 DE LA LEY DEL IMPUESTO A LA RENTA
Recibí de:
R.U.C.
RECIBO POR HONORARIOS ELECTRONICO
Nro:
10000000014
-E001 10
PEREZ GARCIA MARIA ELENA
MZA. A LOTE. 1 URB. LOS EJEMPLOS
Total Neto Recibido:
ALTAMIRA CONSULTING S.A.C. - ALTAMIRA S.A.C.
RUC 20900000011
 QUINIENTOS NOVENTA Y OCHO Y 40/100 SOLES
SERVICIOS DE APOYO EN ACTIVIDADES COMERCIALES
07 Agosto 2026
598.40
(0.00)
598.40
SOLES
%) IR:8
Fecha de emisión
AL CONTADOForma de Pago:"""


def test_parse_rhe_texto_campos():
    d = CP.parse_rhe_texto(RHE_TEXTO, "RHE10000000014E00110.pdf")
    assert d["por_digitar"] is False
    assert d["emisor_ruc"] == "10000000014"
    assert d["adquirente_ruc"] == "20900000011"
    assert d["serie_numero"] == "E001-10"
    assert d["fecha_emision"] == datetime.date(2026, 8, 7)
    assert d["total"] == 598.40
    assert d["retencion"] == 0.0
    assert d["neto"] == 598.40
    assert d["igv"] == 0.0
    assert "Garcia" in d["emisor_nombre"]
    assert CP.clasificar(d) == ("cxp", "ALTAMIRA")


def test_parse_rhe_texto_con_retencion():
    t = RHE_TEXTO.replace("598.40\n(0.00)\n598.40", "650.00\n(52.00)\n598.00")
    d = CP.parse_rhe_texto(t)
    assert d["total"] == 650.00
    assert d["retencion"] == 52.00
    assert d["neto"] == 598.00


def test_parse_rhe_texto_ilegible_es_por_digitar():
    """PDF sin texto usable → por_digitar, con CERO montos inventados."""
    d = CP.parse_rhe_texto("", "RHE_escaneado.pdf")
    assert d["por_digitar"] is True
    assert d["total"] == 0.0
    assert d["archivo"] == "RHE_escaneado.pdf"


def test_parse_rhe_texto_sin_monto_es_por_digitar():
    d = CP.parse_rhe_texto("RECIBO POR HONORARIOS\n10000000014\nE001-99")
    assert d["por_digitar"] is True
