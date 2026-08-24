# Rama de resultados

Esta rama la escribe sola el workflow "Consulta web". No edites nada a mano y
no la mezcles con main: aqui no vive codigo.

- `entradas/<id>/` : los PDF que sube la web antes de procesarlos.
- `resultados/<id>.json` : lo que encontro la consulta, que la web lee y pinta.
- `archivos/<id>/` : el Excel y el PDF listos para descargar.

El `<id>` lo genera la web al azar en cada consulta. Como el repositorio es
publico, cualquiera que adivine un id puede leer ese resultado: son datos de
registros publicos (RUC, OSCE, DIGESA), pero tenlo presente.

Si algun dia pesa mucho, se puede borrar entera y el workflow la vuelve a crear.
