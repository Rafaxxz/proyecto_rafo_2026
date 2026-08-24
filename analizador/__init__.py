"""Nucleo del analizador: extraccion de PDFs SEACE y consultas OSCE / DIGESA.

Aqui no hay interfaz. La cara visible es la web estatica de web/sitio, que
dispara web/consultar.py dentro de GitHub Actions y este paquete hace el trabajo.
"""
import re

import pdfplumber

from . import buenapro, consulta, contratos, digesa, utiles  # noqa: F401


def clasificar_pdf(ruta):
    """Mira las dos primeras paginas y decide si es contrato o buena pro."""
    try:
        with pdfplumber.open(ruta) as pdf:
            texto = ''
            for pagina in pdf.pages[:2]:
                texto += (pagina.extract_text() or '') + '\n'
    except Exception:
        return 'desconocido', ''

    t = texto.upper()
    if 'OTORGAMIENTO DE BUENA PRO' in t or ('ENTIDAD CONVOCANTE' in t and 'VALOR REFERENCIAL' in t):
        return 'buena_pro', texto
    if re.search(r'CONTRATO\s+N[°ºO]', t) or 'CLAUSULA' in t or 'CLÁUSULA' in t:
        return 'contrato', texto
    if 'ENTIDAD CONVOCANTE' in t or 'SE@CE' in t:
        return 'buena_pro', texto
    if 'CONTRATISTA' in t or re.search(r'RUC\s*[N°ºO]*\s*[:\.]?\s*\d{11}', t):
        return 'contrato', texto
    return 'desconocido', texto


def procesar_ruta(ruta, nombre, salida):
    """Clasifica un PDF y lo manda a su modulo, acumulando en `salida`."""
    tipo, _ = clasificar_pdf(ruta)
    if tipo == 'contrato':
        salida['contratos'].append(contratos.procesar_archivo(ruta, nombre))
    elif tipo == 'buena_pro':
        salida['buena_pro'].append(buenapro.procesar_archivo(ruta, nombre))
    else:
        salida['desconocidos'].append({'archivo': nombre})
