# Buzon de resultados

Esta rama es un buzon temporal, no un archivo historico. La escribe sola el
workflow "Consulta web": no edites nada a mano y no la mezcles con main.

- `entradas/<id>/` : PDFs que sube la web; el workflow los borra al terminar.
- `resultados/<id>.json` : lo que encontro la consulta.
- `archivos/<id>/` : el Excel y el PDF de esa consulta.

Flujo normal: la web recoge todo, se lo lleva a la memoria del navegador y
enseguida borra los archivos de aqui. Si alguien cierra la pestana a media
consulta, el workflow purga lo que quede pasado un dia (el `<id>` empieza con
un sello de tiempo UTC justamente para eso).

O sea que en condiciones normales esta rama esta casi vacia. Si algun dia pesa
de mas, se puede borrar entera: el workflow la vuelve a crear.
