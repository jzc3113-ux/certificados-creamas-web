# Generador web de certificados Crea+

Herramienta web estática para generar certificados y constancias desde un Excel y una plantilla PNG limpia. Está pensada para publicarse en GitHub Pages y funcionar sin backend, sin Python y sin PowerShell.

## Privacidad y alcance

- El Excel, la plantilla PNG, la fuente opcional y los PDFs se procesan localmente en el navegador.
- La aplicación no sube archivos a servidores y no guarda datos reales en el repositorio.
- El repositorio no debe incluir bases reales, certificados generados ni plantillas con firmas reales.
- Si se agregan ejemplos en el futuro, deben ser ficticios y sin información sensible.

## Archivos principales

- `index.html`: estructura de la interfaz y carga de librerías por CDN.
- `styles.css`: diseño visual responsive inspirado en una herramienta simple de sorteo.
- `app.js`: lectura del Excel, previsualización, edición visual, guardado de configuración, generación de PDFs y descarga ZIP.

## Librerías usadas por CDN

- [SheetJS / xlsx](https://sheetjs.com/) para leer Excel.
- [jsPDF](https://github.com/parallax/jsPDF) para generar cada certificado PDF.
- [JSZip](https://stuk.github.io/jszip/) para empaquetar varios PDFs.
- [FileSaver.js](https://github.com/eligrey/FileSaver.js/) para descargas locales.

## Uso

1. Abre `index.html` en el navegador o visita la versión publicada en GitHub Pages.
2. Sube un Excel (`.xlsx`, `.xls` o `.csv`) con encabezados en la primera fila.
3. Sube una plantilla PNG limpia.
4. Opcionalmente, sube una fuente `.ttf` o `.otf`.
5. Selecciona las columnas para:
   - `nombre_completo`
   - `texto_certificado` o `texto_base`
   - `texto_gracias` (opcional)
   - `fecha`
   - `codigo_certificado`
6. Revisa la vista previa.
7. Arrastra los campos en el canvas o ajusta manualmente X, Y, tamaño, color, ancho máximo, visibilidad, mayúsculas y alineación.
8. Guarda la configuración en el navegador o descárgala como JSON para reutilizarla.
9. Genera los certificados.
10. Descarga el ZIP con todos los PDFs.

## Configuración JSON

La configuración exportada solo contiene posiciones, estilos y mapeo de columnas. No incluye Excel, PNG, fuente ni PDFs generados.

## Publicación en GitHub Pages

Este proyecto es completamente estático. Para publicarlo:

1. Sube los archivos del repositorio a GitHub.
2. En **Settings → Pages**, selecciona la rama y carpeta raíz del proyecto.
3. Guarda la configuración y espera a que GitHub Pages publique el sitio.

## MVP incluido

La versión inicial permite trabajar con una plantilla PNG subida por el usuario, una hoja Excel, los campos `nombre_completo`, `texto_certificado`, `fecha`, `codigo_certificado`, vista previa editable y descarga de certificados PDF en ZIP.
