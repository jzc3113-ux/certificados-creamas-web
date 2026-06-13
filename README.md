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
7. Revisa la vista previa pequeña para ajustar posiciones.
8. Usa **Ver vista previa en grande** para revisar el certificado completo antes de generar todos los PDFs.
9. Arrastra los campos en el canvas o ajusta manualmente X, Y, tamaño, color, ancho máximo, interletrado, espacio inferior, visibilidad, mayúsculas, alineación y fuente por campo.
10. Opcionalmente, guarda la configuración en el navegador o descárgala como JSON para reutilizarla.
11. Genera los certificados.
12. Descarga el ZIP con todos los PDFs.
13. Si corriges una posición, color, tamaño, fuente o ancho después de generar, la app marca el ZIP anterior como desactualizado y puedes generar un nuevo ZIP sin refrescar la página.

## Fuentes

Si no subes fuentes, la app usa fuentes seguras del navegador y fuentes integradas de jsPDF como respaldo. Si subes varias fuentes, cada campo puede usar una fuente distinta desde el selector **Fuente** del panel de edición.

Recomendación usual:

- `nombre_completo`: una fuente condensada o de titular, por ejemplo nombres parecidos a Nort, Headline, Cond o NortHeadline.
- `texto_certificado`, `texto_gracias`, `fecha` y `codigo_certificado`: una fuente legible para cuerpo, por ejemplo nombres parecidos a Atkinson o Hyperlegible.

La app intenta asignar automáticamente esas fuentes por nombre cuando las detecta. Siempre puedes cambiar la fuente manualmente por campo.

Campos con selector de fuente:

- `nombre_completo`
- `texto_certificado`
- `texto_gracias`
- `fecha`
- `codigo_certificado`

Si una fuente falla al previsualizarse o al generarse en PDF, la app muestra un aviso claro, usa una fuente de respaldo y continúa la generación.

## Vista previa, regeneración y lotes

- **Ver vista previa en grande** abre un modal con fondo oscuro, el certificado centrado, navegación entre registros y scroll cuando el diseño no entra completo en pantalla.
- **Generar vista previa** actualiza la vista pequeña y la vista grande si está abierta.
- El PDF se genera desde el mismo canvas renderizado que usa la vista previa, insertando la imagen completa del certificado en jsPDF. Por eso posiciones, tamaños, saltos de línea, interletrado, espacios inferiores, fuentes y negritas coinciden con lo que ves antes de descargar.
- Después de generar un ZIP, puedes seguir moviendo campos o cambiando estilos. Si haces cambios, el ZIP queda marcado como desactualizado y debes generar certificados otra vez para obtener un ZIP actualizado.
- **Nuevo lote** limpia certificados generados, ZIP anterior, mensajes e índice de vista previa, pero conserva plantilla PNG, fuentes y configuración de campos para cargar otro Excel similar.
- **Reiniciar todo** limpia Excel, PNG, fuentes, JSON, configuración, vista previa y ZIP después de confirmar la acción.

## Nombres largos y negritas

El campo `nombre_completo` tiene autoajuste por defecto: reduce el tamaño hasta el mínimo configurado y, si aún no entra, parte el texto en máximo dos líneas. Si el nombre sigue siendo demasiado largo para el ancho disponible, la app muestra una advertencia para reducir tamaño o aumentar ancho máximo. Este comportamiento se aplica en la vista previa pequeña, la vista previa grande y el PDF.

Para poner una parte del texto en negrita, escribe el texto entre doble asterisco en el Excel. Ejemplo: `Por su participación en **Potencia tu perfil profesional**`. Esto funciona en textos largos como `texto_certificado` y `texto_gracias`, respeta saltos de línea reales y también secuencias `\n` escritas dentro de la celda.

Cada campo incluye controles finos de acabado tipográfico: `letterSpacing` para interletrado y `marginBottom` para empujar el siguiente bloque visual hacia abajo sin mover manualmente todos los campos. Por defecto, `nombre_completo` incluye un pequeño espacio inferior para separar mejor el nombre del contenido siguiente.

## Configuración JSON

El JSON es una opción avanzada. La configuración exportada solo contiene posiciones, estilos, interletrado, espacio inferior, fuente asignada por campo y mapeo de columnas. No incluye Excel, PNG, archivos de fuente ni PDFs generados.

La exportación incluye la fuente por campo en `fields` y también en `campos`, por ejemplo:

```json
{
  "campos": {
    "nombre_completo": {
      "font": "NortHeadlineCond"
    },
    "texto_certificado": {
      "font": "AtkinsonHyperlegible"
    }
  }
}
```

Al cargar un JSON, si la fuente guardada ya fue subida por el usuario, se aplica automáticamente. Si no está disponible, la app avisa y usa una fuente de respaldo hasta que subas esa fuente.

## Publicación en GitHub Pages

Este proyecto es completamente estático. Para publicarlo:

1. Sube los archivos del repositorio a GitHub.
2. En **Settings → Pages**, selecciona la rama y carpeta raíz del proyecto.
3. Guarda la configuración y espera a que GitHub Pages publique el sitio.
