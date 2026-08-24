"""Consulta directa por RUC, sin PDF de por medio.

Devuelve lo mismo que sale al procesar una buena pro (OSCE + DIGESA) y sabe
envolverlo con la forma que esperan los generadores de Excel y PDF.
"""
import re

from . import buenapro, digesa


def rucs(texto):
    """RUCs de un texto libre (separados por coma, espacio o salto de linea)."""
    vistos = []
    for ruc in re.findall(r'\d{11}', texto or ''):
        if ruc not in vistos:
            vistos.append(ruc)
    return vistos


def consultar_ruc(ruc):
    osce = buenapro.consultar_osce(ruc)
    dig = digesa.consultar(ruc)
    return {'ruc': ruc, 'osce': osce, 'digesa': dig,
            'razon_social': osce.get('razon_social', '')}


def como_resultado(consultas):
    """Envuelve las consultas con la forma que usan generar_excel/generar_pdf_general."""
    ganadores = [{
        'ruc': c.get('ruc', ''),
        'razon_social': c.get('razon_social', ''),
        'monto': '',
        'osce': c.get('osce', {}),
        'digesa': c.get('digesa', {}),
    } for c in consultas]
    return [{
        'archivo': 'Consulta por RUC',
        'entidad': '',
        'nomenclatura': '',
        'descripcion': '',
        'ganadores': ganadores,
        'ganadores_detalle': ganadores,
    }]
