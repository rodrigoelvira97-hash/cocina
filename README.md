# Cocina

Calendario de comidas y cenas y biblioteca de recetas.

Aplicación estática para GitHub Pages, adaptada a teléfonos y ordenadores.
En el móvil abre Hoy; el calendario mensual se presenta en tarjetas por día.
Puede añadirse a la pantalla de inicio y consultar el contenido ya cargado sin conexión.

## Publicación

Settings → Pages → Deploy from a branch → main → / (root).

## Datos

Incluye una copia de consulta del catálogo de recetas. Los documentos maestros permanecen en OneDrive. Los enlaces privados se omiten de esta publicación.

El calendario inicial es reproducible con la misma edición del catálogo. Los cambios manuales y nuevas generaciones se guardan en cada navegador; esta versión todavía no sincroniza esos cambios entre dispositivos.

## Actualizaciones

El paquete se regenera desde el proyecto Cocina con `python scripts/preparar-web-github.py`. Después, publicar los cambios con GitHub Desktop.
