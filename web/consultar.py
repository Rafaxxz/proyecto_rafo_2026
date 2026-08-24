"""Motor de la version web: corre dentro de GitHub Actions.

La web de GitHub Pages es estatica y no puede consultar OSCE ni DIGESA por su
cuenta (ninguno de los dos manda cabeceras CORS, y DIGESA ademas necesita
cookies de sesion). Asi que la pagina dispara este script como workflow y
despues lee el resultado que queda publicado en la rama de resultados.

Usa exactamente los mismos modulos que el APK: aqui no se reimplementa nada.

Uso:
    python web/consultar.py --id abc123 --rucs "20100055237, 20611555548"
    python web/consultar.py --id abc123 --pdfs entradas/abc123
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(RAIZ, 'app', 'src', 'main', 'python'))

from analizador import buenapro, consulta, contratos, procesar_ruta, utiles   # noqa: E402


def rucs_de(texto):
    vistos = []
    for ruc in re.findall(r'\d{11}', texto or ''):
        if ruc not in vistos:
            vistos.append(ruc)
    return vistos


def pdfs_de(carpeta):
    if not carpeta or not os.path.isdir(carpeta):
        return []
    return sorted(os.path.join(carpeta, f) for f in os.listdir(carpeta)
                  if f.lower().endswith('.pdf'))


def escribir(ruta, contenido):
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    modo = 'wb' if isinstance(contenido, (bytes, bytearray)) else 'w'
    with open(ruta, modo, **({} if modo == 'wb' else {'encoding': 'utf-8'})) as f:
        f.write(contenido)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--id', required=True, help='identificador de la consulta')
    ap.add_argument('--rucs', default='', help='RUCs sueltos separados por lo que sea')
    ap.add_argument('--pdfs', default='', help='carpeta con los PDF subidos')
    ap.add_argument('--salida', default='publicar', help='carpeta donde dejar todo')
    args = ap.parse_args()

    ident = re.sub(r'[^A-Za-z0-9_-]', '', args.id)[:40] or 'consulta'
    salida = {'contratos': [], 'buena_pro': [], 'desconocidos': []}
    consultas = []

    rucs = rucs_de(args.rucs)
    for i, ruc in enumerate(rucs, 1):
        print(f'[{i}/{len(rucs)}] consultando RUC {ruc}...', flush=True)
        consultas.append(consulta.consultar_ruc(ruc))

    archivos = pdfs_de(args.pdfs)
    for i, ruta in enumerate(archivos, 1):
        nombre = os.path.basename(ruta)
        print(f'[{i}/{len(archivos)}] procesando PDF {nombre}...', flush=True)
        try:
            procesar_ruta(ruta, nombre, salida)
        except Exception as e:
            print(f'   fallo: {e}', flush=True)
            salida['desconocidos'].append({'archivo': nombre, 'error': str(e)})

    # Los exportes salen de los mismos generadores del APK.
    exportes = {}

    def exportar(clave, base, datos, generadores):
        exportes[clave] = {}
        for etiqueta, generar, ext in generadores:
            nombre = utiles.nombre_archivo(base, ext)
            escribir(os.path.join(args.salida, 'archivos', ident, nombre), generar(datos).read())
            exportes[clave][etiqueta] = nombre

    GEN_BUENA_PRO = (('excel', buenapro.generar_excel, 'xlsx'),
                     ('pdf', buenapro.generar_pdf_general, 'pdf'))

    if consultas:
        exportar('rucs', 'consulta_ruc', consulta.como_resultado(consultas), GEN_BUENA_PRO)
    if salida['buena_pro']:
        exportar('buena_pro', 'analisis_buena_pro', salida['buena_pro'], GEN_BUENA_PRO)
    if salida['contratos']:
        exportar('contratos', 'contratos_ganadores', salida['contratos'],
                 (('excel', contratos.generar_excel, 'xlsx'),))

    resultado = {
        'id': ident,
        'generado': datetime.now().isoformat(timespec='seconds'),
        'rucs': consultas,
        'contratos': salida['contratos'],
        'buena_pro': salida['buena_pro'],
        'desconocidos': salida['desconocidos'],
        'exportes': exportes,
    }
    destino = os.path.join(args.salida, 'resultados', f'{ident}.json')
    escribir(destino, json.dumps(resultado, ensure_ascii=False))
    print(f'listo: {destino} '
          f'({len(consultas)} RUC, {len(salida["contratos"])} contratos, '
          f'{len(salida["buena_pro"])} buena pro)', flush=True)


if __name__ == '__main__':
    main()
