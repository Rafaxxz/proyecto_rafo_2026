# Analizador SEACE - APK (todo dentro del celular)

Flask + tus dos modulos (contratos y buena pro) corriendo en Android con Chaquopy,
dentro de un WebView. Al abrir o compartir un PDF desde cualquier app, aparece
"Analizador SEACE" en Abrir con / Compartir y el archivo se procesa solo.

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
- Abrir la app normal: portada con modo automatico + modulos /contratos y /buenapro.
- Abrir un PDF desde Descargas/WhatsApp/Drive > Abrir con > Analizador SEACE:
  se encola, la app se abre y lo clasifica y procesa solo.
- Compartir varios PDFs a la vez tambien funciona (SEND_MULTIPLE).
- Los exportes Excel/PDF se guardan en la carpeta Descargas del celular.

## Notas
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
