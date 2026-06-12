# Generador web de certificados Crea+

Herramienta web estática para generar certificados y constancias desde un Excel simple y una plantilla PNG limpia. Está pensada para publicarse en GitHub Pages y funcionar sin backend, sin Python y sin PowerShell.

## Flujo del MVP

La herramienta trabaja con un flujo directo para usuarios no técnicos:

1. Un Excel representa un solo tipo de certificado o constancia.
2. Una plantilla PNG representa el diseño limpio de ese certificado.
3. Los textos variables vienen directamente desde columnas del Excel.
4. El JSON es opcional y solo sirve para guardar o reutilizar posiciones y estilos.

Columnas recomendadas del Excel:

- `nombre_completo`
- `texto_certificado`
- `texto_gracias`
- `fecha`
- `codigo_certificado`
- `correo`

Si el Excel tiene una sola hoja, la app la usa automáticamente. Si tiene varias hojas, muestra un selector simple de hoja, sin cruces ni lógica avanzada entre hojas.

## Privacidad y alcance

- El Excel, la plantilla PNG, las fuentes opcionales, los JSON de configuración y los PDFs se procesan localmente en el navegador.
- La aplicación no sube archivos a servidores y no guarda datos reales en el repositorio.
- El repositorio no debe incluir bases reales, certificados generados, fuentes oficiales ni plantillas con firmas reales.
- Si se agregan ejemplos en el futuro, deben ser ficticios y sin información sensible.

## Archivos principales

- `index.html`: estructura de la interfaz, flujo principal y carga de librerías por CDN.
- `styles.css`: diseño visual responsive con botones grandes y pasos claros.
- `app.js`: lectura del Excel, selección de hoja, mapeo de columnas, carga de varias fuentes, previsualización, edición visual, guardado de configuración, generación de PDFs y descarga ZIP.

## Librerías usadas por CDN

- [SheetJS / xlsx](https://sheetjs.com/) para leer Excel.
- [jsPDF](https://github.com/parallax/jsPDF) para generar cada certificado PDF.
- [JSZip](https://stuk.github.io/jszip/) para empaquetar varios PDFs.
- [FileSaver.js](https://github.com/eligrey/FileSaver.js/) para descargas locales.

## Uso

1. Abre `index.html` en el navegador o visita la versión publicada en GitHub Pages.
2. Sube un Excel (`.xlsx`, `.xls` o `.csv`) con encabezados en la primera fila.
3. Si el Excel tiene varias hojas, elige la hoja correcta en el selector.
4. Sube una plantilla PNG limpia.
5. Opcionalmente, sube una o varias fuentes `.ttf` o `.otf`.
6. Revisa o cambia el mapeo de columnas. Cualquier campo opcional puede quedar como “No usar”.
7. Revisa la vista previa.
8. Arrastra los campos en el canvas o ajusta manualmente X, Y, tamaño, color, ancho máximo, visibilidad, mayúsculas, alineación y fuente por campo.
9. Opcionalmente, guarda la configuración en el navegador o descárgala como JSON para reutilizarla.
10. Genera los certificados.
11. Descarga el ZIP con todos los PDFs.

## Fuentes

Si no subes fuentes, la app usa una fuente segura del navegador y Helvetica como respaldo en PDF. Si subes fuentes, puedes elegir una fuente distinta por cada campo editable:

- `nombre_completo`
- `texto_certificado`
- `texto_gracias`
- `fecha`
- `codigo_certificado`

Si una fuente falla al previsualizarse o al generarse en PDF, la app muestra un aviso claro, usa una fuente de respaldo y continúa la generación.

## Configuración JSON

El JSON es una opción avanzada. La configuración exportada solo contiene posiciones, estilos, fuente asignada por campo y mapeo de columnas. No incluye Excel, PNG, archivos de fuente ni PDFs generados.

## Publicación en GitHub Pages

Este proyecto es completamente estático. Para publicarlo:

1. Sube los archivos del repositorio a GitHub.
2. En **Settings → Pages**, selecciona la rama y carpeta raíz del proyecto.
3. Guarda la configuración y espera a que GitHub Pages publique el sitio.
