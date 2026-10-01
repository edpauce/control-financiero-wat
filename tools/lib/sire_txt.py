"""
tools/lib/sire_txt.py — Parser de los TXT del SIRE de SUNAT (RVIE ventas / RCE compras).

Formato pipe (|) con encabezado; el contador los declara cada mes y son la fuente OFICIAL
(lo que da crédito fiscal en compras / lo declarado en ventas). Verificado: el IGV que
suma el RCE de un mes coincide con la constancia de SUNAT.

Mapea por NOMBRE de columna (no índice fijo) porque RVIE y RCE tienen layouts distintos:
RCE trae 'Año' y 'BI Gravado DG/DGNG/DNG'; RVIE trae 'Valor Facturado Exportación' y
'BI Gravada'. La contraparte es el PROVEEDOR en compras y el CLIENTE en ventas — en ambos
vive en 'Nro Doc Identidad' + 'Apellidos Nombres/ Razón Social' (distinta del titular, que
es 'Apellidos y Nombres o Razón social' en la col 1).
"""
from __future__ import annotations

import re
import unicodedata


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def _col(hdr: list[str], *needles: str, evitar: str = "") -> int:
    """Índice de la primera columna cuyo encabezado contiene TODOS los 'needles' y NO el
    'evitar' (para distinguir 'Razón social' del titular vs de la contraparte)."""
    ev = _norm(evitar)
    for i, h in enumerate(hdr):
        hn = _norm(h)
        if all(_norm(n) in hn for n in needles) and (not ev or ev not in hn):
            return i
    return -1


def _num(x) -> float:
    try:
        return round(float(str(x).replace(",", "").strip() or 0), 2)
    except ValueError:
        return 0.0


def detectar_libro(hdr: list[str]) -> str:
    """'RCE' (compras) · 'RVIE' (ventas) · 'DESCONOCIDO'."""
    txt = " ".join(_norm(h) for h in hdr)
    if "bi gravado dg" in txt or "valor adq" in txt:
        return "RCE"
    if "valor facturado exportacion" in txt or ("bi gravada" in txt and "dg" not in txt):
        return "RVIE"
    return "DESCONOCIDO"


def parse(path: str) -> tuple[str, list[dict]]:
    """
    Devuelve (libro, registros). Cada registro:
      {comp, tipo_cp, fecha, ruc, contraparte, bi, igv, total, moneda, tc, es_nc}
    comp = 'SERIE-NRO' (ej. 'F008-146235'). es_nc True si Tipo CP = 07 (nota de crédito).
    """
    # los TXT del SIRE vienen en UTF-8 (con BOM). Leerlos como latin-1 desarma los acentos de
    # los encabezados ('Fecha de emisión'→'emisin') y rompe el mapeo por nombre; utf-8-sig con
    # fallback a latin-1 por si algún export viene distinto.
    try:
        texto = open(path, encoding="utf-8-sig").read()
    except UnicodeDecodeError:
        texto = open(path, encoding="latin-1").read()
    lineas = [l for l in texto.splitlines() if l.strip()]
    if not lineas:
        return "DESCONOCIDO", []
    hdr = lineas[0].split("|")
    libro = detectar_libro(hdr)
    i_tipo = _col(hdr, "tipo cp")
    i_ser = _col(hdr, "serie del cdp")
    i_nro = _col(hdr, "nro", "inicial")
    i_fec = _col(hdr, "fecha de emision")
    # contraparte: 'Nro Doc Identidad' y el nombre de al lado (evita el titular de la col 1)
    i_ruc = _col(hdr, "nro doc identidad")
    i_nom = _col(hdr, "apellidos", "razon", evitar="o razon social")   # el de la contraparte
    if i_nom < 0:
        i_nom = i_ruc + 1 if i_ruc >= 0 else -1
    i_bi = _col(hdr, "bi grav")            # BI Gravada (RVIE) o BI Gravado DG (RCE): el primero
    i_igv = _col(hdr, "igv")               # IGV / IPM (primer match = el principal)
    i_tot = _col(hdr, "total cp")
    i_mon = _col(hdr, "moneda")
    i_tc = _col(hdr, "tipo de cambio")
    # Documento modificado: en una NOTA DE CRÉDITO dice QUÉ comprobante anula/rectifica.
    # Sin esto hay que adivinar por monto+cliente, y con facturas gemelas se falla: una NC
    # que anulaba una factura se había aplicado a su gemela del mismo importe.
    i_mser = _col(hdr, "serie cp modificado")
    i_mnro = _col(hdr, "nro cp modificado")

    regs = []
    for l in lineas[1:]:
        c = l.split("|")
        g = lambda i: c[i] if 0 <= i < len(c) else ""
        serie, nro = g(i_ser).strip(), g(i_nro).strip()
        mser, mnro = g(i_mser).strip(), g(i_mnro).strip()
        regs.append({
            "comp": f"{serie}-{nro}" if serie else nro,
            "tipo_cp": g(i_tipo).strip(),
            "es_nc": g(i_tipo).strip() == "07",
            # comprobante que la NC modifica ('' si no aplica). Normalizado igual que `comp`.
            "modifica": f"{mser}-{mnro}" if mser and mnro else "",
            "fecha": g(i_fec).strip(),
            "ruc": g(i_ruc).strip(),
            "contraparte": g(i_nom).strip(),
            "bi": _num(g(i_bi)),
            "igv": _num(g(i_igv)),
            "total": _num(g(i_tot)),
            "moneda": g(i_mon).strip() or "PEN",
            "tc": _num(g(i_tc)) or 1.0,
        })
    return libro, regs


def totales(regs: list[dict]) -> dict:
    """Totales SUMANDO tal cual: el TXT del SIRE ya trae las NC (tipo 07) con importe NEGATIVO,
    así que sumar directo resta la nota de crédito (no aplicar otro cambio de signo — sería
    doble negación). 'es_nc' queda solo como etiqueta para presentación."""
    return {"n": len(regs),
            "igv": round(sum(r["igv"] for r in regs), 2),
            "total": round(sum(r["total"] for r in regs), 2),
            "bi": round(sum(r["bi"] for r in regs), 2)}
