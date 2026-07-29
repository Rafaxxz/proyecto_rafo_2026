from flask import Blueprint, render_template, request, jsonify, send_file
import pdfplumber
import re
import os
import io
import shutil
import sys
import time
import openpyxl
import requests
import urllib3
from bs4 import BeautifulSoup
from openpyxl.styles import Font, PatternFill, Alignment
from werkzeug.utils import secure_filename
from fpdf import FPDF

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

try:
    import pytesseract
except ImportError:
    pytesseract = None

if load_dotenv:
    load_dotenv()

bp = Blueprint('contratos', __name__, url_prefix='/contratos')
import tempfile
UPLOAD_FOLDER = os.path.join(os.environ.get('HOME', tempfile.gettempdir()), 'uploads')

DIGESA_URL = 'https://consultas-digesa.minsa.gob.pe/ConsultaWebRS/Consultas/Consulta_Registro_Sanitario.aspx'
BASE = 'https://eap.osce.gob.pe/perfilprov-bus/1.0/ficha'

HEADERS = {
    'User-Agent'     : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122 Safari/537.36',
    'Accept'         : 'application/json, text/plain, */*',
    'Accept-Language': 'es-PE,es;q=0.9',
    'Referer'        : 'https://apps.osce.gob.pe/perfilprov-ui/buscar',
    'Origin'         : 'https://apps.osce.gob.pe',
}

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def normalizar_texto(valor):
    return ' '.join((valor or '').split()).strip()

def extraer_telefonos_desde_texto(texto):
    patron = r'(?:\+?51\s*)?(?:\(?\d{2,3}\)?[\s-]*)?\d{3,4}[\s-]?\d{3,4}'
    telefonos = []
    vistos = set()
    for raw in re.findall(patron, texto):
        digitos = re.sub(r'\D', '', raw)
        if len(digitos) < 7:
            continue
        valor = normalizar_texto(raw)
        if digitos not in vistos:
            vistos.add(digitos)
            telefonos.append(valor)
    return telefonos

def extraer_hidden_webforms(soup):
    payload = {}
    for hidden in soup.select('input[type="hidden"][name]'):
        payload[hidden.get('name')] = hidden.get('value', '')
    return payload

def construir_payload_ruc_webforms(soup, ruc):
    form = soup.select_one('form#aspnetForm')
    if not form:
        raise ValueError('No se encontró el formulario ASP.NET en DIGESA.')
    payload = extraer_hidden_webforms(soup)
    for inp in form.select('input[name]'):
        nombre = inp.get('name')
        tipo = (inp.get('type') or '').lower()
        if nombre in payload:
            continue
        if tipo in ('checkbox', 'radio') and not inp.has_attr('checked'):
            continue
        payload[nombre] = inp.get('value', '')
    for sel in form.select('select[name]'):
        nombre = sel.get('name')
        option = sel.select_one('option[selected]') or sel.select_one('option')
        payload[nombre] = option.get('value', '') if option else ''
    payload['ctl00$ContentPlaceHolder1$TabContainer1$ClientState'] = '{"ActiveTabIndex":1,"TabState":[true,true,true,true,true,true]}'
    payload['ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$TextBox_ConsultaRUC'] = ruc
    payload['ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$Button_ConsultaRUC'] = 'Buscar'
    payload['__EVENTTARGET'] = 'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$Button_ConsultaRUC'
    payload['__EVENTARGUMENT'] = ''
    return payload

def mapear_indices_columnas(headers):
    mapeo = {'registro': -1, 'producto': -1, 'estado': -1, 'fecha': -1}
    for i, raw in enumerate(headers):
        h = normalizar_texto(raw).lower()
        if mapeo['registro'] == -1 and 'registro' in h and 'expediente' not in h: mapeo['registro'] = i
        if mapeo['producto'] == -1 and 'producto' in h: mapeo['producto'] = i
        if mapeo['estado'] == -1 and 'estado' in h: mapeo['estado'] = i
        if mapeo['fecha'] == -1 and ('emision' in h or 'emisi' in h or 'fecha' in h): mapeo['fecha'] = i
    return mapeo

def extraer_registros_desde_tabla(table):
    rows = table.select('tr')
    if not rows:
        return []
    headers = [normalizar_texto(th.get_text(' ', strip=True)) for th in rows[0].select('th')]
    indices = mapear_indices_columnas(headers)
    registros = []
    for row in rows[1:]:
        celdas = [normalizar_texto(td.get_text(' ', strip=True)) for td in row.select('td')]
        if not celdas:
            continue
        texto_fila = ' | '.join(celdas)
        numero  = celdas[indices['registro']] if 0 <= indices['registro'] < len(celdas) else ''
        producto = celdas[indices['producto']] if 0 <= indices['producto'] < len(celdas) else ''
        estado  = celdas[indices['estado']]   if 0 <= indices['estado']   < len(celdas) else ''
        fecha   = celdas[indices['fecha']]    if 0 <= indices['fecha']    < len(celdas) else ''
        if not numero:
            m = re.search(r'\b[A-Z]{1,3}[\-/]?\d{4,}\b', texto_fila)
            if m: numero = m.group(0)
        if not fecha:
            m = re.search(r'\b\d{1,2}[\-/]\d{1,2}[\-/]\d{2,4}\b', texto_fila)
            if m: fecha = m.group(0)
        if not estado:
            m = re.search(r'\b(Vigente|Vencido|Por Vencer|Cancelado|Suspendido(?:\s+Parcialmente)?|Reinscripci[oó]n en tramite|Baja de Registro Sanitario)\b', texto_fila, re.IGNORECASE)
            if m: estado = m.group(1)
        if not producto:
            candidato = ''
            for celda in celdas:
                if not celda or re.search(r'\d{11}', celda): continue
                if (numero and numero in celda) or (fecha and fecha in celda) or (estado and estado.lower() in celda.lower()): continue
                if len(celda) > len(candidato): candidato = celda
            producto = candidato
        if numero or producto or estado or fecha:
            registros.append({'numero_registro': numero, 'producto': producto, 'estado': estado, 'fecha_emision': fecha})
    return registros

def _digesa_headers_base():
    return {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122.0.0.0 Safari/537.36',
        'Accept': '*/*',
        'Accept-Language': 'es-ES,es;q=0.9,en-GB;q=0.7,en-US;q=0.6,es-PE;q=0.5',
        'Cache-Control': 'no-cache',
        'Referer': DIGESA_URL,
        'Origin': 'https://consultas-digesa.minsa.gob.pe',
    }

def _digesa_payload_base(soup, ruc, anio):
    """Payload base con todos los campos que manda el browser"""
    def val(id_):
        el = soup.find('input', {'id': id_})
        return el['value'] if el else ''

    return {
        'ctl00$ContentPlaceHolder1$ScriptManager1': (
            'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$UpdatePanel8'
            '|ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$Button_ConsultaRUC'
        ),
        'ctl00_ContentPlaceHolder1_TabContainer1_ClientState': '{"ActiveTabIndex":1,"TabState":[true,true,true,true,true,true]}',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaEmpresa$TextBox_ConsultaEmpresa': '',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaEmpresa$ddlEstado_Empresa': '%',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaEmpresa$ddlAñoEmision_Empresa': '2026',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$TextBox_ConsultaRUC': ruc,
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$ddlEstado_RUC': '%',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$ddlAñoEmision_RUC': anio,
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaProducto$TextBox_ConsultaProducto': '',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaProducto$ddlEstado_Producto': '%',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaProducto$ddlAñoEmision_Producto': '2026',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaCertificado$TextBox_ConsultaCertificado': '',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaExpediente$TextBox_NumeroExp': '',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaExpediente$DropDownList_AnoExp': '2026',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaExpediente$DropDownList_TupaExp': '30',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_Departamento$DDL_Departamento': '0',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_Departamento$ddlEstado_Departamento': '%',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_Departamento$ddlAñoEmision_Departamento': '2026',
        'ctl00$ContentPlaceHolder1$HiddenField_ParamBusqueda': ruc,
        'ctl00$ContentPlaceHolder1$HiddenField_ParamFlag': '4',
        'ctl00$ContentPlaceHolder1$HiddenField_ParamCodigo': '',
        '__EVENTTARGET': '',
        '__EVENTARGUMENT': '',
        '__VIEWSTATE': val('__VIEWSTATE'),
        '__VIEWSTATEGENERATOR': val('__VIEWSTATEGENERATOR'),
        '__VIEWSTATEENCRYPTED': '',
        '__EVENTVALIDATION': val('__EVENTVALIDATION'),
        '__ASYNCPOST': 'true',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$Button_ConsultaRUC': 'Buscar',
    }

def _extraer_vs(soup2_text):
    """Extrae VIEWSTATE, VIEWSTATEGENERATOR y EVENTVALIDATION de la respuesta UpdatePanel"""
    soup = BeautifulSoup(soup2_text, 'html.parser')
    def val(id_):
        el = soup.find('input', {'id': id_})
        return el['value'] if el else ''
    vs = val('__VIEWSTATE')
    vsg = val('__VIEWSTATEGENERATOR') or '60B26698'
    ev = val('__EVENTVALIDATION')
    if not vs:
        m = re.search(r'\|hiddenField\|__VIEWSTATE\|(.+?)\|', soup2_text)
        if m: vs = m.group(1)
    if not ev:
        m = re.search(r'\|hiddenField\|__EVENTVALIDATION\|(.+?)\|', soup2_text)
        if m: ev = m.group(1)
    return vs, vsg, ev

def consultar_digesa_por_ruc(ruc):
    ruc = re.sub(r'\D', '', ruc or '')
    vacio = {'ruc': ruc, 'registros': [], 'telefonos': [], 'rep_legal': '',
             'tiene_digesa': False, 'tipo': 'Revendedor', 'direccion': ''}
    if len(ruc) != 11:
        return vacio

    try:
        hb = _digesa_headers_base()
        hp = {**hb, 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
              'X-MicrosoftAjax': 'Delta=true', 'X-Requested-With': 'XMLHttpRequest'}

        session = requests.Session()
        resp_get = session.get(DIGESA_URL, headers={**hb, 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'}, timeout=30)
        soup_ini = BeautifulSoup(resp_get.text, 'html.parser')

        registros_todos = []
        telefonos = []
        rep_legal = ''
        direccion = ''

        for anio in ['2026', '2025', '2024', '2023', '2022']:
            try:
                # ── PASO 1: Buscar por RUC ──
                payload1 = _digesa_payload_base(soup_ini, ruc, anio)
                r1 = session.post(DIGESA_URL, data=payload1, headers=hp, timeout=60)
                c1 = r1.text

                if 'No se encontraron resultados' in c1:
                    continue

                soup1 = BeautifulSoup(c1, 'html.parser')
                tabla = soup1.find('table', id='ctl00_ContentPlaceHolder1_GridView1')
                if not tabla:
                    continue

                filas = tabla.find_all('tr')[1:]
                if not filas:
                    continue

                vs1, vsg1, ev1 = _extraer_vs(c1)

                hf_busqueda = ''
                hf_flag = ''
                hf_codigo = ''
                el = soup1.find('input', {'id': 'ctl00_ContentPlaceHolder1_HiddenField_ParamBusqueda'})
                if el: hf_busqueda = el.get('value', '')
                el = soup1.find('input', {'id': 'ctl00_ContentPlaceHolder1_HiddenField_ParamFlag'})
                if el: hf_flag = el.get('value', '')
                el = soup1.find('input', {'id': 'ctl00_ContentPlaceHolder1_HiddenField_ParamCodigo'})
                if el: hf_codigo = el.get('value', '')

                # ── PASO 2: Click en primer registro para obtener detalle ──
                primer_link = filas[0].find('a', id=lambda x: x and 'LinkButton_RegSan' in str(x))
                if primer_link and vs1:
                    link_name = re.sub(r'ctl00_', 'ctl00$', primer_link.get('id', ''))
                    link_name = link_name.replace('_ContentPlaceHolder1_', '$ContentPlaceHolder1$')
                    link_name = link_name.replace('_GridView1_', '$GridView1$')
                    link_name = link_name.replace('_ctl0', '$ctl0')
                    link_name = link_name.replace('_LinkButton_RegSan', '$LinkButton_RegSan')

                    payload2 = {
                        'ctl00$ContentPlaceHolder1$ScriptManager1': (
                            'ctl00$ContentPlaceHolder1$UpdatePanel1'
                            '|' + link_name
                        ),
                        'ctl00_ContentPlaceHolder1_TabContainer1_ClientState': '{"ActiveTabIndex":1,"TabState":[true,true,true,true,true,true]}',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaEmpresa$TextBox_ConsultaEmpresa': '',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaEmpresa$ddlEstado_Empresa': '%',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaEmpresa$ddlAñoEmision_Empresa': '2026',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$TextBox_ConsultaRUC': ruc,
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$ddlEstado_RUC': '%',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$ddlAñoEmision_RUC': anio,
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaProducto$TextBox_ConsultaProducto': '',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaProducto$ddlEstado_Producto': '%',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaProducto$ddlAñoEmision_Producto': '2026',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaCertificado$TextBox_ConsultaCertificado': '',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaExpediente$TextBox_NumeroExp': '',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaExpediente$DropDownList_AnoExp': '2026',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaExpediente$DropDownList_TupaExp': '30',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_Departamento$DDL_Departamento': '0',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_Departamento$ddlEstado_Departamento': '%',
                        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_Departamento$ddlAñoEmision_Departamento': '2026',
                        'ctl00$ContentPlaceHolder1$HiddenField_ParamBusqueda': hf_busqueda or ruc,
                        'ctl00$ContentPlaceHolder1$HiddenField_ParamFlag': hf_flag or '4',
                        'ctl00$ContentPlaceHolder1$HiddenField_ParamCodigo': hf_codigo,
                        '__EVENTTARGET': link_name,
                        '__EVENTARGUMENT': '',
                        '__VIEWSTATE': vs1,
                        '__VIEWSTATEGENERATOR': vsg1,
                        '__VIEWSTATEENCRYPTED': '',
                        '__EVENTVALIDATION': ev1,
                        '__ASYNCPOST': 'true',
                    }

                    r2 = session.post(DIGESA_URL, data=payload2, headers=hp, timeout=60)
                    soup2 = BeautifulSoup(r2.text, 'html.parser')

                    if not telefonos:
                        tel = soup2.find('span', id=lambda x: x and 'Label23' in str(x))
                        if tel:
                            t = normalizar_texto(tel.get_text())
                            if t: telefonos.append(t)

                    if not rep_legal:
                        rep = soup2.find('span', id=lambda x: x and 'Label24' in str(x))
                        if rep:
                            rep_legal = normalizar_texto(rep.get_text())

                    if not direccion:
                        dir_el = soup2.find('span', id=lambda x: x and 'Label19' in str(x))
                        if dir_el:
                            direccion = normalizar_texto(dir_el.get_text())

                for fila in filas:
                    celdas = fila.find_all('td')
                    if len(celdas) < 8:
                        continue
                    reg_el = celdas[0].find('a') or celdas[0]
                    numero = normalizar_texto(reg_el.get_text())
                    prod_span = celdas[3].find('span')
                    producto_full = normalizar_texto(prod_span.get_text() if prod_span else celdas[3].get_text())
                    prod_corto = producto_full.split('Denominación')[0].strip()
                    prod_corto = re.sub(r'^Producto\s+\d+\s*:', '', prod_corto).strip()
                    fecha_emision = normalizar_texto(celdas[4].get_text())
                    fecha_venc    = normalizar_texto(celdas[6].get_text())
                    empresa       = normalizar_texto(celdas[7].get_text()) if len(celdas) > 7 else ''
                    dir_celda     = normalizar_texto(celdas[8].get_text()) if len(celdas) > 8 else ''

                    if numero:
                        registros_todos.append({
                            'numero': numero,
                            'producto': prod_corto[:150],
                            'producto_completo': producto_full[:500],
                            'fecha_emision': fecha_emision,
                            'fecha_vencimiento': fecha_venc,
                            'empresa': empresa,
                            'direccion_registro': dir_celda,
                            'anio': anio,
                        })

            except Exception as e:
                print(f'[DIGESA] Error año {anio}: {e}')
                continue

        vistos = set()
        registros_unicos = []
        for r in registros_todos:
            if r['numero'] not in vistos:
                vistos.add(r['numero'])
                registros_unicos.append(r)

        tiene = len(registros_unicos) > 0
        tipo  = 'Fabricante / Distribuidor' if tiene else 'Revendedor'

        return {
            'ruc': ruc,
            'registros': registros_unicos,
            'telefonos': telefonos,
            'rep_legal': rep_legal,
            'direccion': direccion,
            'tiene_digesa': tiene,
            'tipo': tipo,
        }

    except Exception as e:
        print(f'[DIGESA ERROR] {ruc}: {e}')
        return vacio



#  MÓDULO OSCE
# ─────────────────────────────────────────────
def _n(v):
    return ' '.join(str(v or '').split()).strip()

def consultar_osce_por_ruc(ruc: str) -> dict:
    ruc = re.sub(r'\D', '', ruc)
    vacio = {
        'ruc': ruc, 'razon_social': '', 'emails': [],
        'telefonos_osce': [], 'celulares_osce': [],
        'direccion': '', 'estado_rnp': '',
        'capacidad_contratacion': '', 'es_apto': False
    }
    if len(ruc) != 11:
        return vacio

    try:
        data = {}
        for intento in range(2):
            try:
                r = requests.get(
                    f'{BASE}/{ruc}',
                    headers=HEADERS,
                    timeout=15,
                    verify=False
                )
                if r.status_code != 200:
                    print(f'[OSCE] status {r.status_code} para {ruc}')
                    if intento == 0:
                        time.sleep(0.35)
                        continue
                    return vacio
                data = r.json()
                break
            except Exception:
                if intento == 0:
                    time.sleep(0.35)
                    continue
                raise

        prov = data.get('proveedorT01') or {}
        if not prov:
            return vacio

        telefonos = prov.get('telefonos') or []
        emails    = prov.get('emails') or []

        celulares = [t for t in telefonos if t.startswith('9') and len(re.sub(r'\D','',t)) == 9]
        fijos     = [t for t in telefonos if t not in celulares]

        return {
            'ruc'                   : prov.get('numRuc', ruc),
            'razon_social'          : _n(prov.get('nomRzsProv', '')),
            'emails'                : [e.lower() for e in emails],
            'telefonos_osce'        : fijos,
            'celulares_osce'        : celulares,
            'direccion'             : '',
            'estado_rnp'            : 'Habilitado' if prov.get('esHabilitado') else 'No habilitado',
            'es_apto'               : prov.get('esAptoContratar', False),
            'capacidad_contratacion': '',
        }
    except Exception as e:
        print(f'[OSCE ERROR] {ruc}: {e}')
        return vacio

# ─────────────────────────────────────────────
#  Funciones para Consorcio y PDF
# ─────────────────────────────────────────────

def extraer_rucs_consorcio(texto_miembros):
    """Extrae RUCs de los miembros del consorcio"""
    rucs = []
    if not texto_miembros:
        return rucs
    for linea in texto_miembros.split('\n'):
        linea = linea.strip()
        if linea:
            ruc_match = re.search(r'(\d{11})', linea)
            if ruc_match:
                rucs.append(ruc_match.group(1))
    return rucs

def consultar_all_osce(ruc_principal, miembros_consorcio_texto):
    """Consulta OSCE para el RUC principal y todos los miembros del consorcio"""
    resultados = []
    rucs_consultados = set()
    
    # Consultar RUC principal
    if ruc_principal:
        resultados.append(consultar_osce_por_ruc(ruc_principal))
        rucs_consultados.add(ruc_principal)
    
    # Consultar RUCs del consorcio
    rucs_consorcio = extraer_rucs_consorcio(miembros_consorcio_texto)
    for ruc in rucs_consorcio:
        if ruc not in rucs_consultados:
            resultados.append(consultar_osce_por_ruc(ruc))
            rucs_consultados.add(ruc)
    
    return resultados

def generar_pdf_reporte(datos_completos):
    """Genera un PDF con formato profesional de los datos extraidos"""
    from fpdf import FPDF, XPos, YPos
    
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 15, "REPORTE DE LICITACIÓN", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    # Información Principal
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 10, "INFORMACIÓN PRINCIPAL", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)
    
    pdf.set_font("Helvetica", "", 10)
    datos_principales = [
        ("Archivo", datos_completos.get('archivo', 'N/A')),
        ("Número de Contrato", datos_completos.get('numero_contrato', 'N/A')),
        ("Zona", datos_completos.get('zona', 'N/A')),
    ]
    
    for label, valor in datos_principales:
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(50, 7, f"{label}:")
        pdf.set_font("Helvetica", "", 9)
        valor_texto = str(valor)[:100]
        pdf.multi_cell(0, 7, valor_texto, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    # Información del Ganador
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 10, "CONTRATISTA GANADOR", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)
    
    pdf.set_font("Helvetica", "", 10)
    ruc_principal = datos_completos.get('ruc', '')
    contratista = datos_completos.get('contratista', '')
    rep_legal = datos_completos.get('representante_legal', '')
    
    datos_ganador = [
        ("RUC", ruc_principal),
        ("Razón Social", contratista),
        ("Representante Legal", rep_legal),
        ("DNI Representante", datos_completos.get('dni_representante', 'N/A')),
    ]
    
    for label, valor in datos_ganador:
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(50, 7, f"{label}:")
        pdf.set_font("Helvetica", "", 9)
        valor_texto = str(valor)[:100]
        pdf.multi_cell(0, 7, valor_texto, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    # Información OSCE del Ganador
    if ruc_principal:
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 10, "INFORMACIÓN OSCE - GANADOR", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)
        
        pdf.set_font("Helvetica", "", 9)
        osce_info = [
            ("Razón Social OSCE", datos_completos.get('osce_razon_social', 'N/A')),
            ("Estado RNP", datos_completos.get('osce_estado_rnp', 'N/A')),
            ("Apto para contratar", "Sí" if datos_completos.get('osce_es_apto') else "No"),
            ("Teléfono", ', '.join(datos_completos.get('osce_telefonos', ['N/A']))),
            ("Celular", ', '.join(datos_completos.get('osce_celulares', ['N/A']))),
            ("Email", ', '.join(datos_completos.get('osce_emails', ['N/A']))),
        ]
        
        for label, valor in osce_info:
            pdf.set_font("Helvetica", "B", 8.5)
            pdf.cell(45, 6, f"{label}:")
            pdf.set_font("Helvetica", "", 8.5)
            valor_texto = str(valor)[:120]
            pdf.multi_cell(0, 6, valor_texto, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    # Miembros del Consorcio
    miembros = datos_completos.get('miembros_consorcio', '').strip()
    if miembros:
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 10, "MIEMBROS DEL CONSORCIO", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)
        
        pdf.set_font("Helvetica", "", 9)
        miembros_list = miembros.split('\n')
        for i, miembro in enumerate(miembros_list, 1):
            pdf.cell(10, 6, f"{i}.")
            pdf.multi_cell(0, 6, miembro.strip(), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        # Agregar info OSCE de miembros si está disponible
        if 'osce_consorcio' in datos_completos and datos_completos['osce_consorcio']:
            pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font("Helvetica", "B", 11)
            pdf.cell(0, 9, "INFORMACIÓN OSCE - MIEMBROS", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)
            
            for osce_miembro in datos_completos['osce_consorcio']:
                pdf.set_font("Helvetica", "B", 8)
                ruc_miembro = osce_miembro.get('ruc', 'N/A')
                razon_miembro = osce_miembro.get('razon_social', 'N/A')
                pdf.cell(0, 6, f"RUC: {ruc_miembro} - {razon_miembro}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                pdf.set_font("Helvetica", "", 7.5)
                pdf.cell(10, 5, "")
                pdf.cell(40, 5, "Estado RNP:")
                pdf.cell(0, 5, osce_miembro.get('estado_rnp', 'N/A'), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                pdf.cell(10, 5, "")
                pdf.cell(40, 5, "Apto:")
                pdf.cell(0, 5, "Sí" if osce_miembro.get('es_apto') else "No", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                telefonos = ', '.join(osce_miembro.get('telefonos_osce', ['N/A']))
                pdf.cell(10, 5, "")
                pdf.cell(40, 5, "Teléfono:")
                pdf.multi_cell(0, 5, telefonos, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                celulares = ', '.join(osce_miembro.get('celulares_osce', ['N/A']))
                pdf.cell(10, 5, "")
                pdf.cell(40, 5, "Celular:")
                pdf.multi_cell(0, 5, celulares, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    output = io.BytesIO()
    pdf.output(output)
    output.seek(0)
    return output

# ─────────────────────────────────────────────
#  OCR
# ─────────────────────────────────────────────

def configurar_ocr():
    if pytesseract is None:
        return False, f"OCR no disponible ({sys.executable}): instala pytesseract."
    tesseract_cmd = os.getenv('TESSERACT_CMD', '').strip()
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    else:
        ruta = shutil.which('tesseract')
        if ruta:
            pytesseract.pytesseract.tesseract_cmd = ruta
        else:
            for p in ['C:/Program Files/Tesseract-OCR/tesseract.exe', 'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe']:
                if os.path.exists(p):
                    pytesseract.pytesseract.tesseract_cmd = p
                    break
    try:
        v = pytesseract.get_tesseract_version()
        return True, f"OCR habilitado: Tesseract {v}."
    except Exception:
        return False, "pytesseract instalado pero Tesseract no encontrado. Configura TESSERACT_CMD."

def extraer_texto_ocr_pagina(pagina):
    if pytesseract is None:
        return ''
    try:
        imagen = pagina.to_image(resolution=300).original
        return (pytesseract.image_to_string(imagen, config='--oem 1 --psm 6 -l spa+eng') or '').strip()
    except Exception:
        return ''


# ─────────────────────────────────────────────
#  Extracción PDF
# ─────────────────────────────────────────────

def _limpiar_nombre_empresa(nombre):
    n = ' '.join(nombre.split()).strip(' ,.;:-')
    n = re.sub(r'^(?:LA\s+EMPRESA|EMPRESA|EL\s+CONSORCIO|CONSORCIO)\s+', '', n, flags=re.IGNORECASE)
    n = re.sub(r'\s+(?:REPRESENTADA\s+POR|CON\s+DOMICILIO|IDENTIFICADA\s+CON).*$', '', n, flags=re.IGNORECASE)
    return n.strip(' ,.;:-')

def extraer_miembros_consorcio(texto):
    if not re.search(r'\bconsorcio\b', texto, re.IGNORECASE):
        return ''
    miembros = []
    vistos = set()
    patrones = [
        r'([A-ZÁÉÍÓÚÑ0-9&\-\.,\s]{4,}?)\s+(?:con\s+)?RUC\s*[N°º]*\s*[:\.]?\s*(\d{11})',
        r'RUC\s*[N°º]*\s*[:\.]?\s*(\d{11})\s*(?:,|;|-)?\s*(?:de\s+la\s+empresa\s+|de\s+)?([A-ZÁÉÍÓÚÑ0-9&\-\.,\s]{4,})'
    ]
    for patron in patrones:
        for match in re.findall(patron, texto, re.IGNORECASE):
            nombre, ruc = match if patron == patrones[0] else (match[1], match[0])
            nombre = _limpiar_nombre_empresa(nombre)
            if not nombre or 'CONTRATISTA' in nombre.upper() or 'CONSORCIO' in nombre.upper():
                continue
            clave = (ruc, nombre.upper())
            if clave not in vistos:
                vistos.add(clave)
                miembros.append(f'{ruc} - {nombre}')
    return '\n'.join(miembros)

def extraer_datos_contrato(pdf_path, nombre_archivo):
    datos = {
        'archivo': nombre_archivo, 'numero_contrato': '', 'zona': '',
        'ruc': '', 'contratista': '', 'miembros_consorcio': '',
        'digesa_registros': [], 'digesa_productos': [], 'digesa_telefonos': [],
        'osce_razon_social': '', 'osce_emails': [], 'osce_telefonos': [],
        'osce_celulares': [], 'osce_direccion': '', 'osce_estado_rnp': '',
        'osce_capacidad': '', 'osce_es_apto': False,
        'ocr_aplicado': False, 'representante_legal': '', 'dni_representante': '', 'texto_raw': ''
    }
    try:
        with pdfplumber.open(pdf_path) as pdf:
            texto = ''
            for pagina in pdf.pages[:2]:
                texto_pagina = (pagina.extract_text() or '').strip()
                if len(texto_pagina) < 20:
                    texto_ocr = extraer_texto_ocr_pagina(pagina)
                    if texto_ocr:
                        datos['ocr_aplicado'] = True
                        texto_pagina = f"{texto_pagina}\n{texto_ocr}" if texto_pagina else texto_ocr
                if texto_pagina:
                    texto += texto_pagina + '\n'
            datos['texto_raw'] = texto

            m = re.search(r'CONTRATO\s+N[°º]\s*([^\n]+)', texto, re.IGNORECASE)
            if m:
                datos['numero_contrato'] = m.group(1).strip()
                zm = re.search(r'CGC[-\s]+(.+?)(?:/|$)', datos['numero_contrato'], re.IGNORECASE)
                if zm: datos['zona'] = zm.group(1).strip()

            m = re.search(r'RUC\s*[N°º]*\s*[:\.]?\s*(\d{11})', texto, re.IGNORECASE)
            if m: datos['ruc'] = m.group(1)

            m = re.search(r'de la otra parte\s+(.+?)\s+con\s+(?:código de\s+)?CONTRATISTA', texto, re.IGNORECASE | re.DOTALL)
            if m: datos['contratista'] = ' '.join(m.group(1).split()).strip()

            if not datos['contratista']:
                matches = re.findall(r'([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑA-Za-z\s\.]+(?:S\.A\.C\.|S\.A\.|E\.I\.R\.L\.|S\.R\.L\.|SAC|EIRL|SRL))', texto)
                if matches: datos['contratista'] = matches[-1].strip()

            m = re.search(r'(?:representante legal|Apoderado Legal)[^\n]*?(?:EL\(LA\)\s*)?(?:SEÑOR\(A\)|SEÑOR|SEÑORA)?\s+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑA-Za-z\s]+?)(?:,\s*con\s*DNI|con\s*DNI)', texto, re.IGNORECASE)
            if m: datos['representante_legal'] = ' '.join(m.group(1).split()).strip()

            matches_dni = re.findall(r'DNI\s*[N°º]*\s*[:\.]?\s*(\d{8})', texto, re.IGNORECASE)
            if matches_dni: datos['dni_representante'] = matches_dni[-1]

            datos['miembros_consorcio'] = extraer_miembros_consorcio(texto)
    except Exception as e:
        datos['error'] = str(e)
    return datos


# ─────────────────────────────────────────────
#  Excel
# ─────────────────────────────────────────────

def generar_excel(lista_datos):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contratos"

    encabezados = [
        'N°', 'Archivo PDF', 'Número de Contrato', 'Zona', 'RUC',
        'Contratista (Ganador)', 'Miembros del Consorcio (RUC - Nombre)',
        'Representante Legal', 'DNI Representante',
        'Registros DIGESA', 'Productos DIGESA', 'Teléfonos DIGESA',
        'Razón Social OSCE', 'Email OSCE', 'Teléfonos OSCE', 'Celulares OSCE',
        'Estado RNP', 'Apto para contratar'
    ]

    fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    font_header = Font(color="FFFFFF", bold=True, size=11)

    for col, titulo in enumerate(encabezados, 1):
        celda = ws.cell(row=1, column=col, value=titulo)
        celda.fill = fill_header
        celda.font = font_header
        celda.alignment = Alignment(horizontal='center', vertical='center')

    ws.row_dimensions[1].height = 20
    fill_par = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")

    for i, datos in enumerate(lista_datos, 1):
        digesa_regs = datos.get('digesa_registros', [])
        digesa_texto = '\n'.join(
            f"{r.get('numero_registro','')} | {r.get('estado','')} | {r.get('fecha_emision','')}"
            for r in digesa_regs
        ) if digesa_regs else ''

        valores = [
            i,
            datos.get('archivo', ''),
            datos.get('numero_contrato', ''),
            datos.get('zona', ''),
            datos.get('ruc', ''),
            datos.get('contratista', ''),
            datos.get('miembros_consorcio', ''),
            datos.get('representante_legal', ''),
            datos.get('dni_representante', ''),
            digesa_texto,
            '\n'.join(datos.get('digesa_productos', [])),
            '\n'.join(datos.get('digesa_telefonos', [])),
            datos.get('osce_razon_social', ''),
            '\n'.join(datos.get('osce_emails', [])),
            '\n'.join(datos.get('osce_telefonos', [])),
            '\n'.join(datos.get('osce_celulares', [])),
            datos.get('osce_estado_rnp', ''),
            'Sí' if datos.get('osce_es_apto') else 'No',
        ]

        for col, val in enumerate(valores, 1):
            celda = ws.cell(row=i + 1, column=col, value=val)
            if i % 2 == 0:
                celda.fill = fill_par
            celda.alignment = Alignment(vertical='center', wrap_text=True)

    anchos = [5, 30, 50, 25, 15, 40, 55, 35, 15, 35, 40, 20, 40, 30, 20, 20, 20, 15]
    for col, ancho in enumerate(anchos, 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = ancho

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output



def procesar_archivo(ruta, nombre):
    datos = extraer_datos_contrato(ruta, nombre)
    ruc = (datos.get('ruc') or '').strip()
    if re.fullmatch(r'\d{11}', ruc):
        osce = consultar_osce_por_ruc(ruc)
        datos['osce_razon_social'] = osce.get('razon_social', '')
        datos['osce_emails'] = osce.get('emails', [])
        datos['osce_telefonos'] = osce.get('telefonos_osce', [])
        datos['osce_celulares'] = osce.get('celulares_osce', [])
        datos['osce_direccion'] = osce.get('direccion', '')
        datos['osce_estado_rnp'] = osce.get('estado_rnp', '')
        datos['osce_capacidad'] = osce.get('capacidad_contratacion', '')
        datos['osce_es_apto'] = osce.get('es_apto', False)
        digesa = consultar_digesa(ruc)
        datos['digesa_registros'] = digesa.get('registros', [])
        datos['digesa_telefonos'] = digesa.get('telefonos', [])
        datos['digesa_rep_legal'] = digesa.get('rep_legal', '')
        datos['digesa_direccion'] = digesa.get('direccion', '')
        datos['digesa_tiene'] = digesa.get('tiene_digesa', False)
        datos['digesa_tipo'] = digesa.get('tipo', 'Revendedor')
        datos['digesa_productos'] = list(dict.fromkeys(
            r.get('producto', '') for r in digesa.get('registros', []) if r.get('producto', '')
        ))
        miembros_texto = datos.get('miembros_consorcio', '')
        if miembros_texto:
            datos['osce_consorcio'] = consultar_all_osce(ruc, miembros_texto)[1:]
        else:
            datos['osce_consorcio'] = []
    return datos


# ─────────────────────────────────────────────
#  Rutas Flask
# ─────────────────────────────────────────────

@bp.route('/')
def index():
    return render_template('contratos.html')

@bp.route('/procesar', methods=['POST'])
def procesar():
    if 'pdfs' not in request.files:
        return jsonify({'error': 'No se enviaron archivos'}), 400
    archivos = request.files.getlist('pdfs')
    resultados = []
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    for archivo in archivos:
        if archivo and allowed_file(archivo.filename):
            nombre = secure_filename(archivo.filename)
            ruta = os.path.join(UPLOAD_FOLDER, nombre)
            archivo.save(ruta)
            resultados.append(procesar_archivo(ruta, nombre))
            os.remove(ruta)
    return jsonify({'resultados': resultados})

@bp.route('/exportar', methods=['POST'])
def exportar():
    datos = request.json.get('datos', [])
    excel_file = generar_excel(datos)
    return send_file(
        excel_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='contratos_ganadores.xlsx'
    )

@bp.route('/exportar_pdf', methods=['POST'])
def exportar_pdf():
    """Exporta un registro individual a PDF"""
    datos = request.json.get('datos', {})
    if not datos:
        return jsonify({'error': 'No hay datos para exportar'}), 400
    
    try:
        pdf_file = generar_pdf_reporte(datos)
        archivo_nombre = f"reporte_{datos.get('ruc', 'unknown')}.pdf"
        return send_file(
            pdf_file,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=archivo_nombre
        )
    except Exception as e:
        return jsonify({'error': f'Error generando PDF: {str(e)}'}), 500

@bp.route('/exportar_pdf_lote', methods=['POST'])
def exportar_pdf_lote():
    """Exporta múltiples registros en un PDF consolidado con toda la información"""
    from fpdf import FPDF, XPos, YPos
    
    datos_lista = request.json.get('datos', [])
    if not datos_lista:
        return jsonify({'error': 'No hay datos para exportar'}), 400
    
    try:
        pdf = FPDF()
        
        for idx, datos in enumerate(datos_lista):
            if idx > 0:
                pdf.add_page()
            else:
                pdf.add_page()
            
            # ENCABEZADO
            pdf.set_font("Helvetica", "B", 14)
            pdf.cell(0, 12, f"REPORTE DE LICITACIÓN #{idx + 1}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_line_width(0.5)
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            pdf.cell(0, 2, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            
            # INFORMACIÓN PRINCIPAL
            pdf.set_font("Helvetica", "B", 11)
            pdf.cell(0, 8, "INFORMACIÓN PRINCIPAL DEL CONTRATO", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)
            
            pdf.set_font("Helvetica", "", 9)
            campos_principal = [
                ("Archivo PDF", datos.get('archivo', 'N/A')),
                ("Número de Contrato", datos.get('numero_contrato', 'N/A')),
                ("Zona", datos.get('zona', 'N/A')),
            ]
            
            for label, valor in campos_principal:
                pdf.set_font("Helvetica", "B", 8.5)
                pdf.cell(40, 5, f"{label}:")
                pdf.set_font("Helvetica", "", 8.5)
                pdf.multi_cell(0, 5, str(valor)[:80], new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            
            pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            
            # CONTRATISTA GANADOR
            pdf.set_font("Helvetica", "B", 11)
            pdf.cell(0, 8, "CONTRATISTA GANADOR", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)
            
            pdf.set_font("Helvetica", "", 9)
            campos_ganador = [
                ("RUC", datos.get('ruc', 'N/A')),
                ("Razón Social", datos.get('contratista', 'N/A')),
                ("Representante Legal", datos.get('representante_legal', 'N/A')),
                ("DNI Representante", datos.get('dni_representante', 'N/A')),
            ]
            
            for label, valor in campos_ganador:
                pdf.set_font("Helvetica", "B", 8.5)
                pdf.cell(40, 5, f"{label}:")
                pdf.set_font("Helvetica", "", 8.5)
                pdf.multi_cell(0, 5, str(valor)[:80], new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            
            pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            
            # INFORMACIÓN OSCE DEL GANADOR
            if datos.get('ruc'):
                pdf.set_font("Helvetica", "B", 11)
                pdf.cell(0, 8, "INFORMACIÓN OSCE - EMPRESA GANADORA", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)
                
                pdf.set_font("Helvetica", "", 9)
                campos_osce = [
                    ("Razón Social OSCE", datos.get('osce_razon_social', 'N/A')),
                    ("Estado RNP", datos.get('osce_estado_rnp', 'N/A')),
                    ("Apto para contratar", "Sí" if datos.get('osce_es_apto') else "No"),
                    ("Teléfono(s)", ', '.join(datos.get('osce_telefonos', ['N/A']))),
                    ("Celular(es)", ', '.join(datos.get('osce_celulares', ['N/A']))),
                    ("Email(s)", ', '.join(datos.get('osce_emails', ['N/A']))),
                ]
                
                for label, valor in campos_osce:
                    pdf.set_font("Helvetica", "B", 8.5)
                    pdf.cell(40, 5, f"{label}:")
                    pdf.set_font("Helvetica", "", 8.5)
                    valor_str = str(valor)[:100]
                    pdf.multi_cell(0, 5, valor_str, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            
            # MIEMBROS DEL CONSORCIO
            miembros_texto = datos.get('miembros_consorcio', '').strip()
            if miembros_texto:
                pdf.set_font("Helvetica", "B", 11)
                pdf.cell(0, 8, "MIEMBROS DEL CONSORCIO", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)
                
                miembros_list = miembros_texto.split('\n')
                
                # Mostrar lista de miembros
                pdf.set_font("Helvetica", "", 9)
                for i, miembro_txt in enumerate(miembros_list, 1):
                    miembro_txt = miembro_txt.strip()
                    if miembro_txt:
                        # Extraer RUC del miembro
                        ruc_match = re.search(r'(\d{11})', miembro_txt)
                        if ruc_match:
                            ruc_miembro = ruc_match.group(1)
                            pdf.set_font("Helvetica", "B", 8.5)
                            pdf.cell(10, 6, f"{i}.")
                            pdf.set_font("Helvetica", "", 8.5)
                            pdf.multi_cell(0, 6, miembro_txt, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        else:
                            pdf.cell(10, 6, f"{i}.")
                            pdf.multi_cell(0, 6, miembro_txt, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                # INFORMACIÓN OSCE DE MIEMBROS DEL CONSORCIO
                if 'osce_consorcio' in datos and datos['osce_consorcio']:
                    pdf.cell(0, 4, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf.set_font("Helvetica", "B", 11)
                    pdf.cell(0, 8, "INFORMACIÓN OSCE - MIEMBROS DEL CONSORCIO", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)
                    
                    for idx_miembro, osce_miembro in enumerate(datos['osce_consorcio'], 1):
                        pdf.set_font("Helvetica", "B", 9)
                        ruc_miembro = osce_miembro.get('ruc', 'N/A')
                        razon_miembro = osce_miembro.get('razon_social', 'N/A')
                        pdf.cell(0, 7, f"Miembro {idx_miembro}: RUC {ruc_miembro}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        
                        pdf.set_font("Helvetica", "", 8)
                        info_miembro = [
                            ("Razón Social", razon_miembro),
                            ("Estado RNP", osce_miembro.get('estado_rnp', 'N/A')),
                            ("Apto para contratar", "Sí" if osce_miembro.get('es_apto') else "No"),
                            ("Teléfono", ', '.join(osce_miembro.get('telefonos_osce', ['N/A']))),
                            ("Celular", ', '.join(osce_miembro.get('celulares_osce', ['N/A']))),
                            ("Email", ', '.join(osce_miembro.get('emails', ['N/A']))),
                        ]
                        
                        for lbl, val in info_miembro:
                            pdf.set_font("Helvetica", "B", 7.5)
                            pdf.cell(15, 4, "")
                            pdf.cell(35, 4, f"{lbl}:")
                            pdf.set_font("Helvetica", "", 7.5)
                            val_str = str(val)[:90]
                            pdf.multi_cell(0, 4, val_str, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        
                        pdf.cell(0, 2, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            
            # INFORMACIÓN DIGESA (si existe)
            if datos.get('digesa_registros'):
                pdf.set_font("Helvetica", "B", 11)
                pdf.cell(0, 8, "REGISTROS SANITARIOS (DIGESA)", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)
                
                pdf.set_font("Helvetica", "", 8)
                for reg in datos.get('digesa_registros', [])[:3]:  # Mostrar primeros 3
                    pdf.cell(0, 5, f"- Producto: {reg.get('producto', 'N/A')[:60]}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf.cell(5, 4, "")
                    pdf.multi_cell(0, 4, f"Estado: {reg.get('estado', 'N/A')} | Emision: {reg.get('fecha_emision', 'N/A')}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                if len(datos.get('digesa_registros', [])) > 3:
                    pdf.cell(0, 4, f"... y {len(datos.get('digesa_registros', [])) - 3} más", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        output = io.BytesIO()
        pdf.output(output)
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='contratos_lote.pdf'
        )
    except Exception as e:
        print(f"Error generando PDF lote: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Error generando PDF: {str(e)}'}), 500


# ─────────────────────────────────────────────
#  Consulta DIGESA por RUC (alias)
# ─────────────────────────────────────────────

def consultar_digesa(ruc):
    """Alias para consultar_digesa_por_ruc - convierte datos a formato compatible"""
    datos = consultar_digesa_por_ruc(ruc)
    # El formato ya es compatible, solo se devuelve directamente
    return datos


@bp.route('/consultar_digesa', methods=['POST'])
def consultar_digesa_endpoint():
    body = request.json or {}
    ruc = (body.get('ruc') or '').strip()
    if not ruc:
        return jsonify({'error': 'Debes enviar un RUC.'}), 400
    try:
        return jsonify(consultar_digesa(ruc))
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except requests.HTTPError as e:
        return jsonify({'error': f'Error HTTP DIGESA: {e}'}), 502
    except Exception as e:
        return jsonify({'error': f'No se pudo consultar DIGESA: {e}'}), 500

@bp.route('/consultar_osce', methods=['POST'])
def consultar_osce_endpoint():
    body = request.json or {}
    ruc = (body.get('ruc') or '').strip()
    if not ruc:
        return jsonify({'error': 'Debes enviar un RUC.'}), 400

    try:
        return jsonify(consultar_osce_por_ruc(ruc))
    except Exception as e:
        return jsonify({'error': str(e)}), 200
