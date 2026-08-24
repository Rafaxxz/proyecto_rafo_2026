# Analizador SEACE

El mismo analizador en dos formatos, compartiendo el mismo codigo Python:

- **APK**: Flask + los modulos (contratos y buena pro) corriendo en Android con
  Chaquopy dentro de un WebView. Al abrir o compartir un PDF desde cualquier app,
  aparece "Analizador SEACE" en Abrir con / Compartir y el archivo se procesa solo.
- **Web**: una pagina en GitHub Pages que usa GitHub Actions como motor.
  Ver [Version web](#version-web-github-pages--actions) mas abajo.

## Requisitos
- Android Studio (Hedgehog o mas nuevo) con JDK 17
- Python 3.12 instalado en tu PC (Chaquopy lo usa para empaquetar; si tienes otro
  3.12 en otra ruta, agrega en app/build.gradle.kts dentro de chaquopy.defaultConfig:
  buildPython("C:/ruta/a/python3.12/python.exe"))
- Celular arm64 (cualquier equipo de los ultimos ~7 anos)

## Compilar
1. Abre esta carpeta en Android Studio y deja que sincronice Gradle
   (la primera vez descarga SDK, Chaquopy y las wheels; tarda).
2. Build > Build App Bundle(s) / APK(s) > Build APK(s).
3. El APK queda en app/build/outputs/apk/debug/app-debug.apk.
   Pasalo al celular e instalalo (activa "instalar apps desconocidas").

## Uso
- Abrir la app normal: portada con modo automatico + modulos /contratos, /buenapro
  y /ruc (consulta por RUC sin PDF: escribes el RUC y salen OSCE y DIGESA; tambien
  acepta una lista de RUCs pegada de golpe).
- Abrir un PDF desde Descargas/WhatsApp/Drive > Abrir con > Analizador SEACE:
  se encola, la app se abre y lo clasifica y procesa solo.
- Compartir varios PDFs a la vez tambien funciona (SEND_MULTIPLE).
- Los exportes Excel/PDF se guardan en la carpeta Descargas del celular con el dia,
  la fecha y la hora en el nombre (ej. analisis_buena_pro_general_lunes-17-08-2026_08-30-45.pdf),
  asi nunca chocan entre si ni con archivos de instalaciones anteriores.

## Notas
- Las consultas DIGESA viven en analizador/digesa.py (un solo cliente para
  contratos y buena pro). Recorre hasta 5 anios de emision y 3 paginas del
  GridView por anio: si una empresa tiene cientos de registros sanitarios,
  se listan los mas recientes hasta ese tope (MAX_PAGINAS en ese archivo).
- Las consultas DIGESA/OSCE salen por el internet del celular (datos o wifi).
  Si DIGESA bloquea tu IP movil, veras el resultado vacio igual que en PC.
- OCR (pytesseract) queda desactivado en Android: no hay Tesseract en el celular.
  Los PDFs escaneados sin texto no se podran leer.
- pdfplumber va fijado en 0.9.0 y bs4 usa html.parser: las versiones nuevas
  arrastran librerias nativas sin wheel para Android (pypdfium2, lxml).
- Si Gradle se queja de que falta una wheel para Python 3.12, cambia
  version = "3.11" en app/build.gradle.kts y vuelve a sincronizar.

## Compilar sin Android Studio (GitHub Actions)
1. Crea un repo en GitHub y sube esta carpeta completa (incluye .github/).
2. En el repo: pestana Actions > workflow "Compilar APK" > Run workflow
   (o simplemente haz push, se dispara solo).
3. Cuando termine (~10-15 min la primera vez), entra al run y descarga
   el artifact "AnalizadorSeace-apk": ahi esta app-debug.apk.
4. Pasa el APK al celular e instalalo.

## Version web (GitHub Pages + Actions)

Sin instalar nada: se abre en el navegador del celular o de la PC.

### Por que no es una web normal
GitHub Pages solo sirve archivos estaticos, y el navegador no puede consultar
OSCE ni DIGESA por su cuenta: ninguno de los dos manda la cabecera
`Access-Control-Allow-Origin`, y DIGESA ademas necesita cookies de sesion y
POSTs con VIEWSTATE. Por eso el Python corre en un runner de GitHub Actions:

    celular -> Pages (web/sitio) -> dispara "Consulta web" -> el runner corre
    web/consultar.py -> publica en la rama "resultados" -> la web lo pinta

Es el mismo codigo del APK: `web/consultar.py` importa `analizador` tal cual,
no reimplementa ni el scraping ni los exportes.

### Puesta en marcha (una sola vez)
1. Activa Pages una sola vez: Settings del repo > Pages > Build and deployment >
   Source: **GitHub Actions**. (El workflow no puede activarlo solo: crear el
   sitio pide rango de admin y el token de Actions no lo tiene.)
   Despues, Actions > "Publicar web" > Run workflow. Queda en
   `https://<usuario>.github.io/<repo>/` y de ahi en adelante se actualiza
   solo con cada push que toque `web/sitio`.
2. Crea la clave de acceso: Settings de tu cuenta (no del repo) >
   Developer settings > Personal access tokens > Fine-grained tokens >
   Generate new token.
   - Repository access: solo este repositorio.
   - Permissions > Repository: **Actions: Read and write** y
     **Contents: Read and write**. Nada mas.
   - Expiration: lo que prefieras; cuando venza, se genera otro.
3. Abre la web, pega la clave y listo: queda guardada en ese dispositivo.
   Para sacarla, "Cambiar la clave de este dispositivo".

### Como se usa
- Consulta por RUC (uno o una lista pegada) igual que en la app.
- Subida de PDFs: se suben a la rama "resultados" y el runner los clasifica y
  procesa como el modo automatico del celular.
- Excel y PDF se descargan desde la misma pagina, con dia/fecha/hora en el nombre.

### Lo que hay que saber
- Cada consulta tarda **1-2 minutos**: GitHub tiene que encender una maquina,
  instalar dependencias (quedan cacheadas) y recien ahi consultar. No es
  instantaneo como el APK, que consulta directo.
- Los resultados quedan en la rama `resultados` de un repo **publico**: son
  datos de registros publicos, pero cualquiera puede leerlos si da con el id.
- La clave es un token real: si se filtra, alguien podria lanzar workflows en
  este repo. Se revoca desde la misma pantalla donde se creo.
- Las consultas salen desde la IP del runner de GitHub, no desde tu internet.
- En repos publicos los minutos de Actions no se cobran.
