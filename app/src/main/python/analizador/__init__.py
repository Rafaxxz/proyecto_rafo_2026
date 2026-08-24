import os
import re
import tempfile
import threading

import pdfplumber
from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from . import buenapro, consulta, contratos

BASE = os.path.dirname(os.path.abspath(__file__))
HOME = os.environ.get('HOME', tempfile.gettempdir())
UPLOAD_FOLDER = os.path.join(HOME, 'uploads')
INBOX = os.path.join(HOME, 'inbox')

_cola = []
_lock = threading.Lock()


def encolar(path):
    with _lock:
        if path not in _cola:
            _cola.append(path)


def clasificar_pdf(ruta):
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
    """Clasifica un PDF y lo manda a su modulo. Lo usa la app y la web."""
    tipo, _ = clasificar_pdf(ruta)
    if tipo == 'contrato':
        salida['contratos'].append(contratos.procesar_archivo(ruta, nombre))
    elif tipo == 'buena_pro':
        salida['buena_pro'].append(buenapro.procesar_archivo(ruta, nombre))
    else:
        salida['desconocidos'].append({'archivo': nombre})


def create_app():
    app = Flask(__name__, template_folder=os.path.join(BASE, 'templates'))
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
    app.register_blueprint(contratos.bp)
    app.register_blueprint(buenapro.bp)
    app.register_blueprint(consulta.bp)
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(INBOX, exist_ok=True)

    @app.route('/')
    def portada():
        return render_template('portada.html')

    @app.route('/auto/procesar', methods=['POST'])
    def auto_procesar():
        if 'pdfs' not in request.files:
            return jsonify({'error': 'No se enviaron archivos'}), 400
        salida = {'contratos': [], 'buena_pro': [], 'desconocidos': []}
        for archivo in request.files.getlist('pdfs'):
            if not archivo or not archivo.filename.lower().endswith('.pdf'):
                continue
            nombre = secure_filename(archivo.filename)
            ruta = os.path.join(UPLOAD_FOLDER, nombre)
            archivo.save(ruta)
            try:
                procesar_ruta(ruta, nombre, salida)
            finally:
                if os.path.exists(ruta):
                    os.remove(ruta)
        return jsonify(salida)

    @app.route('/auto/pendientes')
    def auto_pendientes():
        with _lock:
            return jsonify({'archivos': [re.sub(r'^\d+_', '', os.path.basename(p)) for p in _cola]})

    @app.route('/auto/archivo/<int:idx>')
    def auto_archivo(idx):
        with _lock:
            if idx < 0 or idx >= len(_cola):
                return 'no existe', 404
            ruta = _cola[idx]
        return send_file(ruta, mimetype='application/pdf')

    @app.route('/auto/limpiar', methods=['POST'])
    def auto_limpiar():
        with _lock:
            rutas = list(_cola)
            _cola.clear()
        for r in rutas:
            try:
                os.remove(r)
            except Exception:
                pass
        return jsonify({'ok': True})

    @app.route('/auto/procesar_pendientes', methods=['POST'])
    def auto_procesar_pendientes():
        with _lock:
            rutas = list(_cola)
            _cola.clear()
        salida = {'contratos': [], 'buena_pro': [], 'desconocidos': []}
        for ruta in rutas:
            try:
                procesar_ruta(ruta, os.path.basename(ruta), salida)
            finally:
                if os.path.exists(ruta):
                    os.remove(ruta)
        return jsonify(salida)

    return app
