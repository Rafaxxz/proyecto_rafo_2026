"""Cliente DIGESA (consulta de registros sanitarios por RUC).

Antes esta logica estaba duplicada en contratos.py y buenapro.py. Vive aqui para
que los dos modulos consulten igual y los arreglos se hagan en un solo lugar.

Flujo del sitio (ASP.NET WebForms + UpdatePanel):
  1. GET a la pagina para tomar VIEWSTATE/EVENTVALIDATION y los anios del combo.
  2. POST "Buscar" por RUC y anio de emision -> GridView1 con 10 registros por pagina.
  3. POST Page$N para recorrer las paginas siguientes del GridView.
  4. POST sobre un LinkButton de registro -> ficha con telefono, rep. legal y direccion.
"""
import re
from datetime import datetime

import requests
from bs4 import BeautifulSoup

URL = 'https://consultas-digesa.minsa.gob.pe/ConsultaWebRS/Consultas/Consulta_Registro_Sanitario.aspx'

ANIOS_ATRAS = 5    # cuantos anios de emision se revisan hacia atras
MAX_PAGINAS = 3    # paginas del GridView por anio (el sitio pagina de 10 en 10)


MESES = {
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
    'julio': 7, 'agosto': 8, 'septiembre': 9, 'setiembre': 9, 'octubre': 10,
    'noviembre': 11, 'diciembre': 12,
    # El sitio mezcla idiomas: la fecha de emision suele venir con el mes en ingles.
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
}


def _n(v):
    return ' '.join((v or '').split()).strip()


def _fecha(texto):
    """"7 de January del 2026" -> "07/01/2026". Lo que no calce se devuelve igual."""
    t = _n(texto)
    m = re.match(r'^(\d{1,2})\s+de\s+([A-Za-zÁÉÍÓÚáéíóú]+)\s+del?\s+(\d{4})$', t)
    if not m:
        return t
    mes = MESES.get(m.group(2).lower())
    return f'{int(m.group(1)):02d}/{mes:02d}/{m.group(3)}' if mes else t


def _headers():
    return {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122.0.0.0 Safari/537.36',
        'Accept': '*/*',
        'Accept-Language': 'es-ES,es;q=0.9,en-GB;q=0.7,en-US;q=0.6,es-PE;q=0.5',
        'Cache-Control': 'no-cache',
        'Referer': URL,
        'Origin': 'https://consultas-digesa.minsa.gob.pe',
    }


def _headers_post():
    return {**_headers(),
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'X-MicrosoftAjax': 'Delta=true',
            'X-Requested-With': 'XMLHttpRequest'}


def _valor(soup, id_):
    el = soup.find('input', {'id': id_})
    return el.get('value', '') if el else ''


def anios_disponibles(soup):
    """Anios del combo "Ano de emision" del sitio, del mas nuevo al mas viejo.

    Se leen del HTML en vez de fijarlos en el codigo: asi la app sigue
    consultando el anio en curso sin tener que recompilarla cada enero.
    """
    combo = soup.find('select', id=lambda x: x and 'Emision_RUC' in str(x))
    anios = []
    if combo:
        for op in combo.find_all('option'):
            v = _n(op.get('value'))
            if re.fullmatch(r'\d{4}', v):
                anios.append(v)
    if not anios:
        actual = datetime.now().year
        anios = [str(a) for a in range(actual, actual - ANIOS_ATRAS, -1)]
    anios.sort(reverse=True)
    return anios[:ANIOS_ATRAS]


def _payload_busqueda(soup, ruc, anio, tope):
    """Payload del boton Buscar con todos los campos que manda el navegador."""
    return {
        'ctl00$ContentPlaceHolder1$ScriptManager1': (
            'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$UpdatePanel8'
            '|ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$Button_ConsultaRUC'
        ),
        'ctl00_ContentPlaceHolder1_TabContainer1_ClientState': '{"ActiveTabIndex":1,"TabState":[true,true,true,true,true,true]}',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaEmpresa$TextBox_ConsultaEmpresa': '',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaEmpresa$ddlEstado_Empresa': '%',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaEmpresa$ddlAñoEmision_Empresa': tope,
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$TextBox_ConsultaRUC': ruc,
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$ddlEstado_RUC': '%',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$ddlAñoEmision_RUC': anio,
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaProducto$TextBox_ConsultaProducto': '',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaProducto$ddlEstado_Producto': '%',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaProducto$ddlAñoEmision_Producto': tope,
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaCertificado$TextBox_ConsultaCertificado': '',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaExpediente$TextBox_NumeroExp': '',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaExpediente$DropDownList_AnoExp': tope,
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaExpediente$DropDownList_TupaExp': '30',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_Departamento$DDL_Departamento': '0',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_Departamento$ddlEstado_Departamento': '%',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_Departamento$ddlAñoEmision_Departamento': tope,
        'ctl00$ContentPlaceHolder1$HiddenField_ParamBusqueda': ruc,
        'ctl00$ContentPlaceHolder1$HiddenField_ParamFlag': '4',
        'ctl00$ContentPlaceHolder1$HiddenField_ParamCodigo': '',
        '__EVENTTARGET': '',
        '__EVENTARGUMENT': '',
        '__VIEWSTATE': _valor(soup, '__VIEWSTATE'),
        '__VIEWSTATEGENERATOR': _valor(soup, '__VIEWSTATEGENERATOR'),
        '__VIEWSTATEENCRYPTED': '',
        '__EVENTVALIDATION': _valor(soup, '__EVENTVALIDATION'),
        '__ASYNCPOST': 'true',
        'ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$Button_ConsultaRUC': 'Buscar',
    }


def _extraer_vs(html):
    """VIEWSTATE, VIEWSTATEGENERATOR y EVENTVALIDATION de una respuesta UpdatePanel."""
    soup = BeautifulSoup(html, 'html.parser')
    vs = _valor(soup, '__VIEWSTATE')
    vsg = _valor(soup, '__VIEWSTATEGENERATOR') or '60B26698'
    ev = _valor(soup, '__EVENTVALIDATION')
    # La respuesta del UpdatePanel viene delimitada por pipes, no siempre como HTML.
    if not vs:
        m = re.search(r'\|hiddenField\|__VIEWSTATE\|(.+?)\|', html)
        if m:
            vs = m.group(1)
    if not ev:
        m = re.search(r'\|hiddenField\|__EVENTVALIDATION\|(.+?)\|', html)
        if m:
            ev = m.group(1)
    return vs, vsg, ev


def _grid(html):
    soup = BeautifulSoup(html, 'html.parser')
    return soup, soup.find('table', id='ctl00_ContentPlaceHolder1_GridView1')


def _filas_datos(tabla):
    """Filas de registros del GridView, sin encabezado ni fila de paginacion.

    La fila del paginador lleva una tabla anidada; sus <td> (uno por numero de
    pagina) se colaban como registros falsos con numero "1", "2", "3"...
    """
    if not tabla:
        return []
    filas = []
    for fila in tabla.find_all('tr'):
        if fila.find_parent('table') is not tabla:
            continue          # fila de la tabla anidada del paginador
        if fila.find('table'):
            continue          # la fila que contiene al paginador
        celdas = fila.find_all('td', recursive=False)
        if len(celdas) >= 8:
            filas.append((fila, celdas))
    return filas


def _parsear_fila(celdas, anio):
    reg_el = celdas[0].find('a') or celdas[0]
    numero = _n(reg_el.get_text())
    # Un registro sanitario es alfanumerico (ej. C2000821N/NKAISA); descartamos
    # restos de la interfaz que solo traen digitos.
    if len(numero) < 5 or numero.isdigit():
        return None
    prod_span = celdas[3].find('span')
    producto_full = _n(prod_span.get_text() if prod_span else celdas[3].get_text())
    prod_corto = producto_full.split('Denominación')[0].strip()
    prod_corto = re.sub(r'^Producto\s+\d+\s*:', '', prod_corto).strip()
    return {
        'numero': numero,
        'producto': prod_corto[:150],
        'producto_completo': producto_full[:500],
        'fecha_emision': _fecha(celdas[4].get_text()),
        'fecha_vencimiento': _fecha(celdas[6].get_text()),
        'empresa': _n(celdas[7].get_text()) if len(celdas) > 7 else '',
        'direccion_registro': _n(celdas[8].get_text()) if len(celdas) > 8 else '',
        'anio': anio,
    }


def _payload_pagina(base, html, numero_pagina):
    """Postback del paginador del GridView (Page$N)."""
    vs, vsg, ev = _extraer_vs(html)
    if not vs:
        return None
    p = dict(base)
    p.pop('ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$Button_ConsultaRUC', None)
    p['ctl00$ContentPlaceHolder1$ScriptManager1'] = (
        'ctl00$ContentPlaceHolder1$UpdatePanel1|ctl00$ContentPlaceHolder1$GridView1')
    p['__EVENTTARGET'] = 'ctl00$ContentPlaceHolder1$GridView1'
    p['__EVENTARGUMENT'] = f'Page${numero_pagina}'
    p['__VIEWSTATE'] = vs
    p['__VIEWSTATEGENERATOR'] = vsg
    p['__EVENTVALIDATION'] = ev
    return p


def _payload_detalle(base, html, soup_grid, link_id, hidden):
    """Postback del LinkButton de un registro para abrir su ficha."""
    vs, vsg, ev = _extraer_vs(html)
    if not vs:
        return None
    nombre = re.sub(r'ctl00_', 'ctl00$', link_id)
    nombre = nombre.replace('_ContentPlaceHolder1_', '$ContentPlaceHolder1$')
    nombre = nombre.replace('_GridView1_', '$GridView1$')
    nombre = nombre.replace('_ctl0', '$ctl0')
    nombre = nombre.replace('_LinkButton_RegSan', '$LinkButton_RegSan')

    p = dict(base)
    p.pop('ctl00$ContentPlaceHolder1$TabContainer1$TabPanel_ConsultaRUC$Button_ConsultaRUC', None)
    p['ctl00$ContentPlaceHolder1$ScriptManager1'] = f'ctl00$ContentPlaceHolder1$UpdatePanel1|{nombre}'
    p['ctl00$ContentPlaceHolder1$HiddenField_ParamBusqueda'] = hidden.get('busqueda') or base.get(
        'ctl00$ContentPlaceHolder1$HiddenField_ParamBusqueda', '')
    p['ctl00$ContentPlaceHolder1$HiddenField_ParamFlag'] = hidden.get('flag') or '4'
    p['ctl00$ContentPlaceHolder1$HiddenField_ParamCodigo'] = hidden.get('codigo', '')
    p['__EVENTTARGET'] = nombre
    p['__EVENTARGUMENT'] = ''
    p['__VIEWSTATE'] = vs
    p['__VIEWSTATEGENERATOR'] = vsg
    p['__EVENTVALIDATION'] = ev
    return p


def _hidden_actualizados(soup):
    def v(sufijo):
        el = soup.find('input', {'id': f'ctl00_ContentPlaceHolder1_HiddenField_Param{sufijo}'})
        return el.get('value', '') if el else ''
    return {'busqueda': v('Busqueda'), 'flag': v('Flag'), 'codigo': v('Codigo')}


def _ficha(session, html):
    """Telefono, representante legal y direccion de la ficha de un registro."""
    soup = BeautifulSoup(html, 'html.parser')

    def etiqueta(nombre):
        el = soup.find('span', id=lambda x: x and nombre in str(x))
        return _n(el.get_text()) if el else ''

    return {'telefono': etiqueta('Label23'),
            'rep_legal': etiqueta('Label24'),
            'direccion': etiqueta('Label19')}


def consultar(ruc):
    """Devuelve registros sanitarios, telefono, rep. legal y direccion de un RUC."""
    ruc = re.sub(r'\D', '', ruc or '')
    vacio = {'ruc': ruc, 'registros': [], 'telefonos': [], 'rep_legal': '',
             'tiene_digesa': False, 'tipo': 'Revendedor', 'direccion': ''}
    if len(ruc) != 11:
        return vacio

    try:
        hp = _headers_post()
        session = requests.Session()
        inicio = session.get(URL, headers={**_headers(), 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'}, timeout=30)
        soup_ini = BeautifulSoup(inicio.text, 'html.parser')

        anios = anios_disponibles(soup_ini)
        tope = anios[0] if anios else str(datetime.now().year)

        registros = []
        telefonos = []
        rep_legal = ''
        direccion = ''

        for anio in anios:
            try:
                base = _payload_busqueda(soup_ini, ruc, anio, tope)
                r = session.post(URL, data=base, headers=hp, timeout=60)
                html = r.text
                if 'No se encontraron resultados' in html:
                    continue

                soup_grid, tabla = _grid(html)
                filas = _filas_datos(tabla)
                if not filas:
                    continue

                pagina = 1
                while True:
                    for _, celdas in filas:
                        reg = _parsear_fila(celdas, anio)
                        if reg:
                            registros.append(reg)
                    if pagina >= MAX_PAGINAS or len(filas) < 10:
                        break        # ultima pagina o tope alcanzado
                    payload = _payload_pagina(base, html, pagina + 1)
                    if not payload:
                        break
                    r = session.post(URL, data=payload, headers=hp, timeout=60)
                    html = r.text
                    soup_grid, tabla = _grid(html)
                    siguientes = _filas_datos(tabla)
                    if not siguientes:
                        break
                    filas = siguientes
                    pagina += 1

                # Ficha del primer registro visible: telefono, rep. legal y direccion
                # son datos de la empresa, con abrir uno alcanza.
                if not (telefonos and rep_legal and direccion) and filas:
                    link = filas[0][0].find('a', id=lambda x: x and 'LinkButton_RegSan' in str(x))
                    if link and link.get('id'):
                        payload = _payload_detalle(base, html, soup_grid, link['id'],
                                                   _hidden_actualizados(soup_grid))
                        if payload:
                            det = session.post(URL, data=payload, headers=hp, timeout=60)
                            ficha = _ficha(session, det.text)
                            if ficha['telefono'] and not telefonos:
                                telefonos.append(ficha['telefono'])
                            rep_legal = rep_legal or ficha['rep_legal']
                            direccion = direccion or ficha['direccion']

            except Exception as e:
                print(f'[DIGESA] Error año {anio}: {e}')
                continue

        vistos = set()
        unicos = []
        for r in registros:
            if r['numero'] not in vistos:
                vistos.add(r['numero'])
                unicos.append(r)

        tiene = len(unicos) > 0
        return {
            'ruc': ruc,
            'registros': unicos,
            'telefonos': telefonos,
            'rep_legal': rep_legal,
            'direccion': direccion,
            'tiene_digesa': tiene,
            'tipo': 'Fabricante / Distribuidor' if tiene else 'Revendedor',
        }

    except Exception as e:
        print(f'[DIGESA ERROR] {ruc}: {e}')
        return vacio
