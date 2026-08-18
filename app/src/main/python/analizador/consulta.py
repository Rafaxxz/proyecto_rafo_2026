"""Tercera interfaz: consulta directa por RUC, sin PDF de por medio.

Escribes uno o varios RUC y devuelve lo mismo que sale al procesar una buena pro
(OSCE + DIGESA). Los exportes reusan el Excel y el PDF general de buenapro.py:
se arma la misma estructura de "ganadores" que esos generadores esperan.
"""
import re

from flask import Blueprint, jsonify, render_template, request, send_file

from . import buenapro, digesa, utiles

bp = Blueprint('consulta', __name__, url_prefix='/ruc')


def _rucs(texto):
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


def _como_resultado(consultas):
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


@bp.route('/')
def index():
    return render_template('ruc.html')


@bp.route('/buscar', methods=['POST'])
def buscar():
    body = request.json or {}
    ruc = re.sub(r'\D', '', body.get('ruc') or '')
    if len(ruc) != 11:
        return jsonify({'error': 'El RUC debe tener 11 digitos.'}), 400
    return jsonify(consultar_ruc(ruc))


@bp.route('/buscar_lote', methods=['POST'])
def buscar_lote():
    body = request.json or {}
    rucs = _rucs(body.get('rucs') or '')
    if not rucs:
        return jsonify({'error': 'No encontre ningun RUC de 11 digitos.'}), 400
    return jsonify({'resultados': [consultar_ruc(r) for r in rucs]})


@bp.route('/exportar', methods=['POST'])
def exportar():
    consultas = (request.json or {}).get('datos', [])
    excel = buenapro.generar_excel(_como_resultado(consultas))
    return send_file(
        excel,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=utiles.nombre_archivo('consulta_ruc', 'xlsx'))


@bp.route('/exportar_pdf', methods=['POST'])
def exportar_pdf():
    consultas = (request.json or {}).get('datos', [])
    pdf = buenapro.generar_pdf_general(_como_resultado(consultas))
    return send_file(
        pdf,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=utiles.nombre_archivo('consulta_ruc', 'pdf'))
