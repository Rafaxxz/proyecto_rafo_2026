# Analizador SEACE (web)

Consulta proveedores del Estado desde el navegador: RNP de OSCE y registros
sanitarios de DIGESA, escribiendo un RUC o subiendo los PDF de contratos y
buenas pro. Los resultados se descargan en Excel y PDF.

No hay que instalar nada: se abre desde el celular o la PC.

## Como esta armado

GitHub Pages solo sirve archivos estaticos, y el navegador no puede consultar
OSCE ni DIGESA por su cuenta: ninguno de los dos manda la cabecera
`Access-Control-Allow-Origin`, y DIGESA ademas necesita cookies de sesion y
POSTs con VIEWSTATE. Por eso el Python corre en un runner de GitHub Actions:

    navegador -> Pages (web/sitio) -> dispara "Consulta web" -> el runner corre
    web/consultar.py -> deja el resultado en la rama "resultados" -> la web lo
    recoge, lo pinta y borra los archivos

No hay base de datos ni historial: el runner no tiene IP publica, asi que la
unica forma de devolver algo es dejarlo en un sitio del que la web lo recoja.
La rama `resultados` es ese buzon y se vacia sola. La web se lleva el Excel y
el PDF a la memoria del navegador, borra los archivos del repositorio y desde
ahi los descarga el usuario. Si alguien cierra la pestana a media consulta, el
propio workflow purga lo que quede pasado un dia.

    index.html           la web entera (una sola pagina, sin dependencias)
    .nojekyll            que Pages sirva el HTML tal cual, sin procesarlo
    analizador/          el nucleo: scraping, extraccion de PDFs y exportes
      digesa.py          cliente DIGESA (registros sanitarios por RUC)
      buenapro.py        PDFs de buena pro + OSCE + Excel y PDF de resultados
      contratos.py       PDFs de contrato + consorcio + Excel y PDF por lote
      consulta.py        consulta por RUC suelto
      utiles.py          nombres de archivo con dia, fecha y hora
    web/
      consultar.py       lo que ejecuta el runner
      requirements.txt   lo que instala el runner
    .github/workflows/
      consulta-web.yml   el motor que consulta y publica resultados

`index.html` vive en la raiz a proposito: asi Pages lo sirve directamente desde
la rama, sin workflow de publicacion. Antes se publicaba con un workflow y el
Jekyll de GitHub competia con el, pisando la app con el README.

## Puesta en marcha (una sola vez)

1. **Activar Pages**: Settings del repo > Pages > Build and deployment >
   Source: **Deploy from a branch**, rama `main`, carpeta `/ (root)`.
   Queda en `https://<usuario>.github.io/<repo>/` y se actualiza solo con cada
   push. (El workflow no puede activarlo por ti: crear el sitio pide rango de
   admin y el token de Actions no lo tiene.)

2. **Crear la clave de acceso**: Settings de tu cuenta (no del repo) >
   Developer settings > Personal access tokens > Fine-grained tokens >
   Generate new token.
   - Repository access: solo este repositorio.
   - Permissions > Repository: **Actions: Read and write** y
     **Contents: Read and write**. Nada mas.
   - Expiration: cuando venza, se genera otra.

3. Abre la web, pega la clave y listo: queda guardada en ese dispositivo.
   Para sacarla, "Cambiar la clave de este dispositivo".

## Como se usa

- **Por RUC**: escribes uno de 11 digitos, o pegas una lista y los consulta en
  fila.
- **Por PDF**: subes contratos y buenas pro juntos; se clasifican solos y cada
  uno pasa por su flujo.
- **Descargas**: Excel y PDF de cada tanda, con dia, fecha y hora en el nombre
  (`consulta_ruc_lunes-24-08-2026_10-17-10.xlsx`), asi nunca se pisan entre si.

## Probar sin pasar por GitHub

El motor corre igual en tu PC:

    pip install -r web/requirements.txt
    python web/consultar.py --id prueba --rucs "20100055237"
    python web/consultar.py --id prueba --pdfs carpeta_con_pdfs

Deja el JSON y los exportes en `publicar/`. Para ver la web en local:
`python -m http.server` dentro de `web/sitio` (igual necesita la clave, porque
habla con la API de GitHub).

## Lo que hay que saber

- Cada consulta tarda **1-2 minutos**: GitHub tiene que encender una maquina,
  instalar dependencias (quedan cacheadas) y recien ahi consultar.
- Los resultados pasan unos segundos por la rama `resultados` de un repo
  **publico** antes de que la web los borre. Son datos de registros publicos,
  pero durante ese rato cualquiera que de con el id podria leerlos.
- Las descargas viven en la memoria del navegador: si recargas la pagina antes
  de bajar el Excel o el PDF, hay que volver a consultar.
- La clave es un token real: si se filtra, alguien podria lanzar workflows en
  este repo. Se revoca desde donde se creo.
- Las consultas salen desde la IP del runner de GitHub, no desde tu internet.
  Si DIGESA bloquea esa IP, el resultado sale vacio.
- DIGESA se recorre hasta 5 anios de emision y 3 paginas por anio: una empresa
  con cientos de registros muestra los mas recientes hasta ese tope
  (`MAX_PAGINAS` en `analizador/digesa.py`).
- OCR esta desactivado: los PDF escaneados sin texto no se pueden leer.
- En repos publicos los minutos de Actions no se cobran.

## Historia

La primera version era un APK de Android (Flask + Chaquopy dentro de un
WebView). Sigue disponible en el tag `apk-final` por si alguna vez hace falta:

    git checkout apk-final
