"""
tools/lib/comprobantes.py — Parsers de comprobantes RECIBIDOS (el insumo del módulo CxP).

Dos fuentes, dos parsers, ambos PUROS (sin I/O de escritura, testeables con fixtures):

  · parse_ubl(path|bytes) ... factura/boleta/NC/ND electrónica en XML UBL 2.1 (estándar SUNAT).
    Namespace-agnóstico (matching por local-name): los emisores usan prefijos distintos
    (n2:, cbc:, sin prefijo) y el mismo parser debe leerlos todos. OJO: el bloque <Signature>
    también trae PartyIdentification/ID — el RUC del emisor se busca SOLO dentro de
    AccountingSupplierParty (bug clásico: agarrar el ID de la firma).

  · parse_rhe_pdf(path) ..... recibo por honorarios electrónico en PDF (SUNAT no manda XML
    de RHE por correo). Best-effort con pypdf: el texto sale DESORDENADO (el layout PDF no
    respeta el orden visual), así que se pesca por regex campo a campo. Si falta algo
    esencial (RUC emisor o monto) devuelve {"por_digitar": True, ...} — NUNCA se inventa.

  · clasificar(doc) ......... decide el destino: CxP de ALTAMIRA/MERAK (por RUC ADQUIRENTE),
    emitido propio (RUC EMISOR nuestro → es una VENTA, no CxP), o ajeno.
"""
from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
from datetime import date

from . import empresas as EMPX

# RUC → empresa del grupo (fuente única: lib/empresas.py)
RUC_GRUPO = {v: k for k, v in EMPX.RUC.items()}

TIPO_DOC = {"01": "Factura", "03": "Boleta", "07": "Nota de Crédito", "08": "Nota de Débito"}
_MESES_TXT = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
              "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
              "noviembre": 11, "diciembre": 12}


# ─── helpers namespace-agnósticos ────────────────────────────────────────────
def _ln(el) -> str:
    return el.tag.split("}")[-1]


def _hijo(el, nombre):
    """Primer hijo DIRECTO con ese local-name."""
    for c in el:
        if _ln(c) == nombre:
            return c
    return None


def _busca(el, nombre):
    """Primer descendiente (a cualquier profundidad) con ese local-name."""
    for c in el.iter():
        if _ln(c) == nombre:
            return c
    return None


def _texto(el) -> str:
    return (el.text or "").strip() if el is not None else ""


def _num(s) -> float:
    try:
        return float(str(s).replace(",", ""))
    except (ValueError, TypeError):
        return 0.0


def _fecha_iso(s):
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", str(s or ""))
    return date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def _ruc_de_party(party_el) -> tuple[str, str]:
    """(ruc, nombre) desde un AccountingSupplierParty/AccountingCustomerParty."""
    ruc, nombre = "", ""
    if party_el is None:
        return ruc, nombre
    # RUC: PartyIdentification/ID con schemeID=6, o el primer ID de 11 dígitos
    for el in party_el.iter():
        if _ln(el) == "ID":
            v = _texto(el)
            scheme = next((a for k, a in el.attrib.items() if k.split("}")[-1] == "schemeID"), "")
            if scheme == "6" or re.fullmatch(r"[12]0\d{9}", v):
                ruc = v
                break
    if not ruc:  # fallback: CustomerAssignedAccountID (formato SAP del BCP)
        v = _texto(_busca(party_el, "CustomerAssignedAccountID"))
        if re.fullmatch(r"[12]0\d{9}", v):
            ruc = v
    nombre = _texto(_busca(party_el, "RegistrationName")) or _texto(_busca(party_el, "Name"))
    return ruc, re.sub(r"\s+", " ", nombre).strip()


# ─── Parser UBL ──────────────────────────────────────────────────────────────
def parse_ubl(src) -> dict:
    """
    Parsea un XML UBL (Invoice/CreditNote/DebitNote) a un dict plano:
      {tipo, serie_numero, fecha_emision, fecha_venc, moneda, emisor_ruc, emisor_nombre,
       adquirente_ruc, adquirente_nombre, subtotal, igv, total, concepto, forma_pago}
    Lanza ET.ParseError si el XML está corrupto.
    """
    if isinstance(src, bytes):
        root = ET.fromstring(src)
    elif isinstance(src, str) and src.lstrip().startswith("<"):
        root = ET.fromstring(src)
    else:
        root = ET.parse(src).getroot()

    raiz = _ln(root)
    tipo_code = _texto(_hijo(root, "InvoiceTypeCode"))
    if raiz == "CreditNote":
        tipo = "Nota de Crédito"
    elif raiz == "DebitNote":
        tipo = "Nota de Débito"
    else:
        tipo = TIPO_DOC.get(tipo_code, "Factura")

    # el ID del documento es hijo DIRECTO de la raíz (el de <Signature> no)
    serie = _texto(_hijo(root, "ID"))
    moneda = _texto(_hijo(root, "DocumentCurrencyCode")) or "PEN"

    sup_ruc, sup_nom = _ruc_de_party(_hijo(root, "AccountingSupplierParty"))
    cli_ruc, cli_nom = _ruc_de_party(_hijo(root, "AccountingCustomerParty"))

    # totales: SOLO los hijos directos de la raíz (cada línea repite TaxTotal)
    lmt = _hijo(root, "LegalMonetaryTotal") or _hijo(root, "RequestedMonetaryTotal")
    total = _num(_texto(_busca(lmt, "PayableAmount"))) if lmt is not None else 0.0
    subtotal = _num(_texto(_busca(lmt, "LineExtensionAmount"))) if lmt is not None else 0.0
    tax = _hijo(root, "TaxTotal")
    igv = _num(_texto(_hijo(tax, "TaxAmount"))) if tax is not None else 0.0

    # concepto: descripción de la primera línea ('CONCEPTO DE PAGO:PORTES' → 'PORTES')
    linea = (_hijo(root, "InvoiceLine") or _hijo(root, "CreditNoteLine")
             or _hijo(root, "DebitNoteLine"))
    concepto = ""
    if linea is not None:
        item = _busca(linea, "Item")
        concepto = _texto(_busca(item, "Description")) if item is not None else ""
        concepto = re.sub(r"^CONCEPTO DE PAGO\s*:\s*", "", concepto, flags=re.I).strip()

    forma = ""
    pt = _hijo(root, "PaymentTerms")
    if pt is not None:
        forma = _texto(_busca(pt, "PaymentMeansID")).upper()

    return {"tipo": tipo, "serie_numero": serie,
            "fecha_emision": _fecha_iso(_texto(_hijo(root, "IssueDate"))),
            "fecha_venc": _fecha_iso(_texto(_hijo(root, "DueDate"))),
            "moneda": moneda,
            "emisor_ruc": sup_ruc, "emisor_nombre": sup_nom,
            "adquirente_ruc": cli_ruc, "adquirente_nombre": cli_nom,
            "subtotal": round(subtotal, 2), "igv": round(igv, 2), "total": round(total, 2),
            "concepto": concepto[:80], "forma_pago": forma}


# ─── Parser RHE (PDF) ────────────────────────────────────────────────────────
def parse_rhe_texto(texto: str, nombre_archivo: str = "") -> dict:
    """
    Extrae un RHE del TEXTO de su PDF (separado del I/O para testear sin PDFs reales).
    Campos esenciales: RUC emisor (10...) y total. Si faltan → {"por_digitar": True}.
    """
    t = texto or ""
    out = {"tipo": "RHE", "moneda": "PEN", "archivo": nombre_archivo}

    m = re.search(r"\b(10\d{9})\b", t)
    out["emisor_ruc"] = m.group(1) if m else ""
    m = re.search(r"\b(20\d{9})\b", t)
    out["adquirente_ruc"] = m.group(1) if m else ""

    # serie: 'E001 - 110', '-E001 10', 'E001-00000110'
    m = re.search(r"E(\d{3})\s*-?\s*0*(\d+)\b", t)
    out["serie_numero"] = f"E{m.group(1)}-{int(m.group(2))}" if m else ""

    # fecha de emisión: '07 Agosto 2026' o '07/08/2026'
    fecha = None
    m = re.search(r"(\d{1,2})\s+(?:de\s+)?([A-Za-zé]+)\s+(?:de\s+)?(\d{4})", t)
    if m and m.group(2).lower() in _MESES_TXT:
        fecha = date(int(m.group(3)), _MESES_TXT[m.group(2).lower()], int(m.group(1)))
    else:
        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", t)
        if m:
            fecha = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    out["fecha_emision"] = fecha

    # montos: total por honorarios = mayor decimal FUERA de paréntesis; retención = el
    # que va entre paréntesis (el PDF imprime '(52.00)'); neto = total - retención.
    decs = [_num(x) for x in re.findall(r"(?<![(\d])(\d[\d,]*\.\d{2})\b", t)]
    ret = [_num(x) for x in re.findall(r"\((\d[\d,]*\.\d{2})\)", t)]
    out["retencion"] = max(ret) if ret else 0.0
    out["total"] = max(decs) if decs else 0.0
    out["neto"] = round(out["total"] - out["retencion"], 2)
    out["igv"] = 0.0          # RHE no lleva IGV
    out["subtotal"] = out["total"]

    # emisor: línea de MAYÚSCULAS cercana a la serie (heurística; si falla queda vacío)
    nom = ""
    lineas = [ln.strip() for ln in t.splitlines()]
    for i, ln in enumerate(lineas):
        if re.search(r"E\d{3}\s*-?\s*\d+", ln):
            for cand in lineas[i + 1: i + 4]:
                if re.fullmatch(r"[A-ZÑÁÉÍÓÚ' ]{8,60}", cand) and len(cand.split()) >= 2:
                    nom = cand.title()
                    break
            break
    out["emisor_nombre"] = nom

    m = re.search(r"Por concepto de\b(.*?)(?:\n[A-Z]|\Z)", t, re.S)
    conc = re.search(r"SERVICIOS?[^\n]{0,120}", t)
    out["concepto"] = (conc.group(0).strip()[:80] if conc else "")

    out["fecha_venc"] = None
    out["forma_pago"] = "CONTADO" if re.search(r"CONTADO", t, re.I) else ""
    out["por_digitar"] = not (out["emisor_ruc"] and out["total"] > 0)
    return out


def parse_rhe_pdf(path: str) -> dict:
    """parse_rhe_texto sobre el texto de un PDF real. PDF ilegible → por_digitar."""
    try:
        from pypdf import PdfReader
        texto = "\n".join((p.extract_text() or "") for p in PdfReader(path).pages)
    except Exception:
        texto = ""
    return parse_rhe_texto(texto, os.path.basename(path))


# ─── Clasificación ───────────────────────────────────────────────────────────
def clasificar(doc: dict) -> tuple[str, str | None]:
    """
    → ("emitido_propio", empresa) si el EMISOR es del grupo (es venta nuestra, no CxP);
      ("cxp", empresa)            si el ADQUIRENTE es del grupo;
      ("ajeno", None)             si ninguno de los dos (un tercero que comparte el buzón).
    El orden importa: primero emisor (jamás integrar ventas como compras).
    """
    emp = RUC_GRUPO.get(str(doc.get("emisor_ruc", "")))
    if emp:
        return "emitido_propio", emp
    emp = RUC_GRUPO.get(str(doc.get("adquirente_ruc", "")))
    if emp:
        return "cxp", emp
    return "ajeno", None
