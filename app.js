const CERTIFICATE_FIELDS = [
  { key: "nombre_completo", label: "Nombre completo", aliases: ["nombre_completo", "nombre", "nombres", "participante"] },
  { key: "texto_certificado", label: "Texto certificado", aliases: ["texto_certificado", "texto_base", "texto", "cuerpo"] },
  { key: "texto_gracias", label: "Texto gracias", aliases: ["texto_gracias", "gracias", "agradecimiento"] },
  { key: "fecha", label: "Fecha", aliases: ["fecha", "date"] },
  { key: "codigo_certificado", label: "Código certificado", aliases: ["codigo_certificado", "codigo", "código", "code"] },
];

const COLUMN_FIELDS = [
  ...CERTIFICATE_FIELDS,
  { key: "correo", label: "Correo", aliases: ["correo", "email", "e_mail", "mail"] },
];

const DEFAULT_FONT_ID = "browser-default";
const DEFAULT_FONT = {
  id: DEFAULT_FONT_ID,
  label: "Fuente segura del navegador",
  family: "Arial",
  pdfName: "helvetica",
  binary: "",
  objectUrl: "",
  loaded: true,
  error: "",
};

const DEFAULT_CONFIG = {
  columnMap: {
    nombre_completo: "",
    texto_certificado: "",
    texto_gracias: "",
    fecha: "",
    codigo_certificado: "",
    correo: "",
  },
  fields: {
    nombre_completo: { x: 600, y: 330, size: 46, color: "#243047", maxWidth: 840, visible: true, uppercase: true, align: "center", fontId: DEFAULT_FONT_ID },
    texto_certificado: { x: 600, y: 430, size: 25, color: "#243047", maxWidth: 860, visible: true, uppercase: false, align: "center", fontId: DEFAULT_FONT_ID },
    texto_gracias: { x: 600, y: 520, size: 21, color: "#697386", maxWidth: 760, visible: true, uppercase: false, align: "center", fontId: DEFAULT_FONT_ID },
    fecha: { x: 600, y: 635, size: 24, color: "#243047", maxWidth: 600, visible: true, uppercase: false, align: "center", fontId: DEFAULT_FONT_ID },
    codigo_certificado: { x: 1010, y: 780, size: 16, color: "#243047", maxWidth: 280, visible: true, uppercase: false, align: "right", fontId: DEFAULT_FONT_ID },
  },
};

const STORAGE_KEY = "creamas-certificados-config-v2";
const state = {
  workbook: null,
  sheetNames: [],
  selectedSheetName: "",
  rows: [],
  headers: [],
  templateImage: null,
  templateDataUrl: "",
  templateSize: { width: 1200, height: 850 },
  fonts: [structuredClone(DEFAULT_FONT)],
  fontWarnings: [],
  config: structuredClone(DEFAULT_CONFIG),
  selectedField: "nombre_completo",
  selectedRowIndex: 0,
  drag: null,
  zipBlob: null,
};

const els = {};

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  restoreConfig();
  bindEvents();
  renderSheetSelector();
  renderMappingControls();
  renderFieldTabs();
  renderFieldEditor();
  drawPreview();
  updateStatus();
});

function cacheElements() {
  Object.assign(els, {
    excelInput: document.getElementById("excelInput"),
    templateInput: document.getElementById("templateInput"),
    fontInput: document.getElementById("fontInput"),
    configInput: document.getElementById("configInput"),
    fileStatus: document.getElementById("fileStatus"),
    sheetPicker: document.getElementById("sheetPicker"),
    sheetSelect: document.getElementById("sheetSelect"),
    mappingGrid: document.getElementById("mappingGrid"),
    mappingTemplate: document.getElementById("mappingTemplate"),
    previewCanvas: document.getElementById("previewCanvas"),
    fieldTabs: document.getElementById("fieldTabs"),
    fieldEditor: document.getElementById("fieldEditor"),
    previewBtn: document.getElementById("previewBtn"),
    saveConfigBtn: document.getElementById("saveConfigBtn"),
    downloadConfigBtn: document.getElementById("downloadConfigBtn"),
    generateBtn: document.getElementById("generateBtn"),
    downloadZipBtn: document.getElementById("downloadZipBtn"),
    progressBar: document.getElementById("progressBar"),
    generationStatus: document.getElementById("generationStatus"),
    prevRecordBtn: document.getElementById("prevRecordBtn"),
    nextRecordBtn: document.getElementById("nextRecordBtn"),
    recordIndicator: document.getElementById("recordIndicator"),
  });
}

function bindEvents() {
  els.excelInput.addEventListener("change", handleExcelUpload);
  els.templateInput.addEventListener("change", handleTemplateUpload);
  els.fontInput.addEventListener("change", handleFontUpload);
  els.configInput.addEventListener("change", handleConfigUpload);
  els.sheetSelect.addEventListener("change", () => applyWorkbookSheet(els.sheetSelect.value));
  els.previewBtn.addEventListener("click", drawPreview);
  els.saveConfigBtn.addEventListener("click", saveConfigToBrowser);
  els.downloadConfigBtn.addEventListener("click", downloadConfigJson);
  els.generateBtn.addEventListener("click", generateCertificates);
  els.downloadZipBtn.addEventListener("click", downloadZip);
  els.prevRecordBtn.addEventListener("click", () => moveRecord(-1));
  els.nextRecordBtn.addEventListener("click", () => moveRecord(1));

  els.previewCanvas.addEventListener("pointerdown", startDrag);
  els.previewCanvas.addEventListener("pointermove", dragField);
  els.previewCanvas.addEventListener("pointerup", endDrag);
  els.previewCanvas.addEventListener("pointerleave", endDrag);
}

async function handleExcelUpload(event) {
  const [file] = event.target.files;
  if (!file) return;
  setGenerationStatus("Leyendo Excel simple...");
  const buffer = await file.arrayBuffer();
  state.workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  state.sheetNames = state.workbook.SheetNames || [];
  const firstSheetName = state.sheetNames[0] || "";
  applyWorkbookSheet(firstSheetName);
  renderSheetSelector();
  const sheetMessage = state.sheetNames.length > 1 ? ` Selecciona la hoja correcta si no es “${firstSheetName}”.` : "";
  setGenerationStatus(`Excel listo: ${state.rows.length} fila(s) en “${firstSheetName}”.${sheetMessage}`);
}

function applyWorkbookSheet(sheetName) {
  if (!state.workbook || !sheetName) return;
  const sheet = state.workbook.Sheets[sheetName];
  state.selectedSheetName = sheetName;
  state.rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  state.headers = Object.keys(state.rows[0] || {});
  state.selectedRowIndex = 0;
  autoMapColumns();
  renderMappingControls();
  updateStatus();
  drawPreview();
}

function renderSheetSelector() {
  if (!els.sheetPicker || !els.sheetSelect) return;
  if (state.sheetNames.length <= 1) {
    els.sheetPicker.hidden = true;
    els.sheetSelect.innerHTML = "";
    return;
  }
  els.sheetPicker.hidden = false;
  els.sheetSelect.innerHTML = state.sheetNames.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
  els.sheetSelect.value = state.selectedSheetName;
}

async function handleTemplateUpload(event) {
  const [file] = event.target.files;
  if (!file) return;
  if (file.type && file.type !== "image/png") {
    setGenerationStatus("La plantilla debe ser PNG.", true);
    return;
  }
  state.templateDataUrl = await readAsDataUrl(file);
  state.templateImage = await loadImage(state.templateDataUrl);
  state.templateSize = { width: state.templateImage.naturalWidth, height: state.templateImage.naturalHeight };
  els.previewCanvas.width = state.templateSize.width;
  els.previewCanvas.height = state.templateSize.height;
  updateStatus();
  drawPreview();
  setGenerationStatus(`Plantilla lista: ${state.templateSize.width} × ${state.templateSize.height}px.`);
}

async function handleFontUpload(event) {
  const files = Array.from(event.target.files || []);
  if (!files.length) return;

  const uploadedFonts = [];
  for (const file of files) {
    const font = await createFontFromFile(file);
    uploadedFonts.push(font);
    state.fonts.push(font);
  }

  autoAssignUploadedFonts(uploadedFonts);
  renderFieldEditor();
  updateStatus();
  drawPreview();

  const failed = uploadedFonts.filter((font) => !font.loaded);
  if (failed.length) {
    setGenerationStatus(`Se cargaron ${uploadedFonts.length - failed.length} fuente(s). ${failed.length} fuente(s) no pudieron previsualizarse y usarán respaldo si fallan en PDF.`, true);
  } else {
    setGenerationStatus(`Fuente(s) cargada(s): ${uploadedFonts.map((font) => font.label).join(", ")}. Puedes elegir una fuente por campo.`);
  }
}

async function createFontFromFile(file) {
  const buffer = await file.arrayBuffer();
  const binary = arrayBufferToBinary(buffer);
  const objectUrl = URL.createObjectURL(file);
  const label = file.name.replace(/\.(ttf|otf)$/i, "");
  const id = uniqueFontId(`font-${normalizeHeader(label) || Date.now()}`);
  const font = {
    id,
    label,
    family: `CreaMasFont_${id.replace(/[^a-zA-Z0-9_]/g, "_")}`,
    pdfName: `CreaMasPdf_${id.replace(/[^a-zA-Z0-9_]/g, "_")}`,
    binary,
    objectUrl,
    loaded: false,
    error: "",
  };

  try {
    const fontFace = new FontFace(font.family, `url(${objectUrl})`);
    await fontFace.load();
    document.fonts.add(fontFace);
    font.loaded = true;
  } catch (error) {
    console.warn(`No se pudo previsualizar la fuente ${file.name}`, error);
    font.error = "No se pudo previsualizar en el navegador.";
  }

  return font;
}

function autoAssignUploadedFonts(uploadedFonts) {
  if (!uploadedFonts.length) return;
  const firstUsable = uploadedFonts[0].id;
  const secondUsable = uploadedFonts[1]?.id || firstUsable;

  if (state.config.fields.nombre_completo.fontId === DEFAULT_FONT_ID) {
    state.config.fields.nombre_completo.fontId = firstUsable;
  }

  ["texto_certificado", "texto_gracias", "fecha", "codigo_certificado"].forEach((key) => {
    if (state.config.fields[key].fontId === DEFAULT_FONT_ID) state.config.fields[key].fontId = secondUsable;
  });
}

async function handleConfigUpload(event) {
  const [file] = event.target.files;
  if (!file) return;
  try {
    const imported = JSON.parse(await file.text());
    applyConfig(imported);
    renderMappingControls();
    renderFieldTabs();
    renderFieldEditor();
    drawPreview();
    setGenerationStatus("Configuración JSON cargada. Recuerda que el JSON solo guarda posiciones, estilos, fuentes asignadas por nombre interno y columnas; no trae Excel ni PNG.");
  } catch (error) {
    console.error(error);
    setGenerationStatus("No se pudo leer el JSON de configuración.", true);
  }
}

function autoMapColumns() {
  for (const field of COLUMN_FIELDS) {
    const currentColumn = state.config.columnMap[field.key];
    if (currentColumn && state.headers.includes(currentColumn)) continue;
    const match = state.headers.find((header) => field.aliases.includes(normalizeHeader(header)));
    state.config.columnMap[field.key] = match || "";
  }
}

function renderMappingControls() {
  if (!els.mappingGrid || !els.mappingTemplate) return;
  els.mappingGrid.innerHTML = "";
  COLUMN_FIELDS.forEach((field) => {
    const node = els.mappingTemplate.content.firstElementChild.cloneNode(true);
    const label = node.querySelector("span");
    const select = node.querySelector("select");
    label.textContent = field.label;
    select.innerHTML = `<option value="">-- No usar --</option>${state.headers.map((header) => `<option value="${escapeHtml(header)}">${escapeHtml(header)}</option>`).join("")}`;
    select.value = state.config.columnMap[field.key] || "";
    select.addEventListener("change", () => {
      state.config.columnMap[field.key] = select.value;
      drawPreview();
    });
    els.mappingGrid.appendChild(node);
  });
}

function renderFieldTabs() {
  els.fieldTabs.innerHTML = "";
  CERTIFICATE_FIELDS.forEach((field) => {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = field.label;
    button.className = field.key === state.selectedField ? "active" : "";
    button.addEventListener("click", () => {
      state.selectedField = field.key;
      renderFieldTabs();
      renderFieldEditor();
      drawPreview();
    });
    els.fieldTabs.appendChild(button);
  });
}

function renderFieldEditor() {
  const config = state.config.fields[state.selectedField];
  const fieldLabel = CERTIFICATE_FIELDS.find((field) => field.key === state.selectedField)?.label || state.selectedField;
  els.fieldEditor.innerHTML = `
    <label><span>Campo</span><input value="${escapeHtml(fieldLabel)}" disabled></label>
    <label><span>Fuente</span><select data-prop="fontId">${fontOptionsHtml()}</select></label>
    <div class="inline">
      ${numberInput("X", "x", config.x)}
      ${numberInput("Y", "y", config.y)}
    </div>
    <div class="inline">
      ${numberInput("Tamaño", "size", config.size)}
      ${numberInput("Ancho máximo", "maxWidth", config.maxWidth)}
    </div>
    <label><span>Color</span><input data-prop="color" type="color" value="${config.color}"></label>
    <label><span>Alineación</span><select data-prop="align">
      <option value="left">Izquierda</option>
      <option value="center">Centro</option>
      <option value="right">Derecha</option>
    </select></label>
    <label class="checkbox-row"><input data-prop="visible" type="checkbox" ${config.visible ? "checked" : ""}> Visible</label>
    <label class="checkbox-row"><input data-prop="uppercase" type="checkbox" ${config.uppercase ? "checked" : ""}> Mayúsculas</label>
  `;
  els.fieldEditor.querySelector('[data-prop="fontId"]').value = fontExists(config.fontId) ? config.fontId : DEFAULT_FONT_ID;
  els.fieldEditor.querySelector('[data-prop="align"]').value = config.align;
  els.fieldEditor.querySelectorAll("[data-prop]").forEach((input) => {
    input.addEventListener("input", updateSelectedFieldFromInput);
    input.addEventListener("change", updateSelectedFieldFromInput);
  });
}

function fontOptionsHtml() {
  return state.fonts.map((font) => `<option value="${escapeHtml(font.id)}">${escapeHtml(font.label)}${font.error ? " (respaldo si falla)" : ""}</option>`).join("");
}

function numberInput(label, prop, value) {
  return `<label><span>${label}</span><input data-prop="${prop}" type="number" step="1" value="${Number(value).toFixed(0)}"></label>`;
}

function updateSelectedFieldFromInput(event) {
  const input = event.target;
  const prop = input.dataset.prop;
  const config = state.config.fields[state.selectedField];
  if (input.type === "checkbox") config[prop] = input.checked;
  else if (input.type === "number") config[prop] = Number(input.value || 0);
  else config[prop] = input.value;
  drawPreview();
}

function drawPreview() {
  const ctx = els.previewCanvas.getContext("2d");
  const { width, height } = els.previewCanvas;
  ctx.clearRect(0, 0, width, height);

  if (state.templateImage) {
    ctx.drawImage(state.templateImage, 0, 0, width, height);
  } else {
    drawPlaceholderTemplate(ctx, width, height);
  }

  const record = getCurrentRecord();
  CERTIFICATE_FIELDS.forEach((field) => drawFieldOnCanvas(ctx, field.key, record));
  updateRecordIndicator();
}

function drawPlaceholderTemplate(ctx, width, height) {
  ctx.fillStyle = "#fffaf0";
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = "#ffbd32";
  ctx.lineWidth = 12;
  ctx.strokeRect(40, 40, width - 80, height - 80);
  ctx.fillStyle = "#d99000";
  ctx.font = "bold 42px sans-serif";
  ctx.textAlign = "center";
  ctx.fillText("Sube una plantilla PNG", width / 2, 120);
}

function drawFieldOnCanvas(ctx, key, record) {
  const fieldConfig = state.config.fields[key];
  if (!fieldConfig.visible) return;
  const text = getFieldText(key, record) || sampleText(key);
  const displayText = fieldConfig.uppercase ? text.toUpperCase() : text;
  const font = getFontById(fieldConfig.fontId);
  ctx.save();
  ctx.fillStyle = fieldConfig.color;
  ctx.font = `${key === "nombre_completo" ? "700" : "400"} ${fieldConfig.size}px "${font.family}", Arial, Helvetica, sans-serif`;
  ctx.textAlign = fieldConfig.align;
  ctx.textBaseline = "top";
  wrapCanvasText(ctx, displayText, fieldConfig.x, fieldConfig.y, fieldConfig.maxWidth, fieldConfig.size * 1.22);

  if (key === state.selectedField) {
    ctx.strokeStyle = "rgba(46, 184, 114, 0.85)";
    ctx.lineWidth = 2;
    const startX = fieldConfig.align === "center" ? fieldConfig.x - fieldConfig.maxWidth / 2 : fieldConfig.align === "right" ? fieldConfig.x - fieldConfig.maxWidth : fieldConfig.x;
    ctx.strokeRect(startX - 8, fieldConfig.y - 8, fieldConfig.maxWidth + 16, fieldConfig.size * 1.5 + 16);
  }
  ctx.restore();
}

function wrapCanvasText(ctx, text, x, y, maxWidth, lineHeight) {
  const paragraphs = String(text).split(/\n+/);
  let currentY = y;
  paragraphs.forEach((paragraph) => {
    const words = paragraph.split(/\s+/).filter(Boolean);
    let line = "";
    words.forEach((word) => {
      const testLine = line ? `${line} ${word}` : word;
      if (ctx.measureText(testLine).width > maxWidth && line) {
        ctx.fillText(line, x, currentY);
        line = word;
        currentY += lineHeight;
      } else {
        line = testLine;
      }
    });
    ctx.fillText(line, x, currentY);
    currentY += lineHeight;
  });
}

function getCanvasPoint(event) {
  const rect = els.previewCanvas.getBoundingClientRect();
  return {
    x: (event.clientX - rect.left) * (els.previewCanvas.width / rect.width),
    y: (event.clientY - rect.top) * (els.previewCanvas.height / rect.height),
  };
}

function startDrag(event) {
  const point = getCanvasPoint(event);
  const hit = hitTestField(point);
  if (!hit) return;
  state.selectedField = hit;
  const config = state.config.fields[hit];
  state.drag = { key: hit, offsetX: point.x - config.x, offsetY: point.y - config.y };
  els.previewCanvas.setPointerCapture(event.pointerId);
  els.previewCanvas.classList.add("dragging");
  renderFieldTabs();
  renderFieldEditor();
  drawPreview();
}

function dragField(event) {
  if (!state.drag) return;
  const point = getCanvasPoint(event);
  const config = state.config.fields[state.drag.key];
  config.x = Math.round(point.x - state.drag.offsetX);
  config.y = Math.round(point.y - state.drag.offsetY);
  renderFieldEditor();
  drawPreview();
}

function endDrag() {
  state.drag = null;
  els.previewCanvas.classList.remove("dragging");
}

function hitTestField(point) {
  return [...CERTIFICATE_FIELDS].reverse().find((field) => {
    const config = state.config.fields[field.key];
    if (!config.visible) return false;
    const x = config.align === "center" ? config.x - config.maxWidth / 2 : config.align === "right" ? config.x - config.maxWidth : config.x;
    const y = config.y;
    return point.x >= x - 12 && point.x <= x + config.maxWidth + 12 && point.y >= y - 12 && point.y <= y + Math.max(config.size * 2, 48);
  })?.key;
}

async function generateCertificates() {
  if (!validateBeforeGenerate()) return;

  els.generateBtn.disabled = true;
  els.downloadZipBtn.disabled = true;
  state.zipBlob = null;
  els.progressBar.value = 0;
  const zip = new JSZip();
  state.fontWarnings = [];

  try {
    for (let index = 0; index < state.rows.length; index += 1) {
      const record = state.rows[index];
      const pdfBlob = await createPdfForRecord(record);
      const name = sanitizeFileName(getFieldText("nombre_completo", record) || `certificado-${index + 1}`);
      const code = sanitizeFileName(getFieldText("codigo_certificado", record));
      zip.file(`${String(index + 1).padStart(3, "0")}-${name}${code ? `-${code}` : ""}.pdf`, pdfBlob);
      els.progressBar.value = Math.round(((index + 1) / state.rows.length) * 80);
      setGenerationStatus(`Generando PDF ${index + 1} de ${state.rows.length}...`);
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    setGenerationStatus("Comprimiendo ZIP...");
    state.zipBlob = await zip.generateAsync({ type: "blob" }, (metadata) => {
      els.progressBar.value = 80 + Math.round(metadata.percent * 0.2);
    });
    els.progressBar.value = 100;
    els.downloadZipBtn.disabled = false;
    const warningText = state.fontWarnings.length ? ` Aviso: ${state.fontWarnings.join(" ")}` : "";
    setGenerationStatus(`ZIP listo con ${state.rows.length} certificado(s).${warningText}`, Boolean(state.fontWarnings.length));
  } catch (error) {
    console.error(error);
    setGenerationStatus("Ocurrió un error al generar los certificados.", true);
  } finally {
    els.generateBtn.disabled = false;
  }
}

async function createPdfForRecord(record) {
  const { jsPDF } = window.jspdf;
  const { width, height } = state.templateSize;
  const orientation = width >= height ? "landscape" : "portrait";
  const pdf = new jsPDF({ orientation, unit: "px", format: [width, height], compress: true });
  pdf.addImage(state.templateDataUrl, "PNG", 0, 0, width, height);
  registerFontsInPdf(pdf);
  CERTIFICATE_FIELDS.forEach((field) => addFieldToPdf(pdf, field.key, record));
  return pdf.output("blob");
}

function registerFontsInPdf(pdf) {
  state.fonts.forEach((font) => {
    if (!font.binary) return;
    try {
      const fileName = `${font.pdfName}.ttf`;
      pdf.addFileToVFS(fileName, btoa(font.binary));
      pdf.addFont(fileName, font.pdfName, "normal");
      pdf.addFont(fileName, font.pdfName, "bold");
    } catch (error) {
      console.warn(`No se pudo registrar la fuente ${font.label} en el PDF`, error);
      font.error = "No se pudo usar en PDF.";
      addFontWarning(`La fuente “${font.label}” falló en PDF y se usó Helvetica.`);
    }
  });
}

function addFieldToPdf(pdf, key, record) {
  const config = state.config.fields[key];
  if (!config.visible) return;
  const text = getFieldText(key, record);
  if (!text) return;
  const displayText = config.uppercase ? text.toUpperCase() : text;
  const font = getFontById(config.fontId);
  const fontStyle = key === "nombre_completo" ? "bold" : "normal";
  const fontName = getPdfFontName(font);

  try {
    pdf.setFont(fontName, fontStyle);
  } catch (error) {
    console.warn(`No se pudo aplicar la fuente ${font.label}; se usará Helvetica`, error);
    addFontWarning(`La fuente “${font.label}” no se pudo aplicar y se usó Helvetica.`);
    pdf.setFont("helvetica", fontStyle);
  }

  pdf.setFontSize(config.size);
  pdf.setTextColor(config.color);

  const lines = pdf.splitTextToSize(displayText, config.maxWidth);
  const lineHeight = config.size * 1.15;
  lines.forEach((line, index) => {
    pdf.text(line, config.x, config.y + index * lineHeight, { align: config.align, baseline: "top" });
  });
}

function getPdfFontName(font) {
  return font?.binary && !font.error ? font.pdfName : "helvetica";
}

function addFontWarning(message) {
  if (!state.fontWarnings.includes(message)) state.fontWarnings.push(message);
}

function validateBeforeGenerate() {
  if (!state.rows.length) {
    setGenerationStatus("Primero sube un Excel con datos.", true);
    return false;
  }
  if (!state.templateDataUrl) {
    setGenerationStatus("Primero sube una plantilla PNG.", true);
    return false;
  }
  const missing = ["nombre_completo", "texto_certificado", "fecha", "codigo_certificado"].filter((key) => !state.config.columnMap[key]);
  if (missing.length) {
    setGenerationStatus(`Faltan columnas obligatorias: ${missing.join(", ")}. Puedes seleccionar “No usar” solo para campos opcionales como texto_gracias o correo.`, true);
    return false;
  }
  return true;
}

function downloadZip() {
  if (!state.zipBlob) return;
  saveAs(state.zipBlob, `certificados-creamas-${new Date().toISOString().slice(0, 10)}.zip`);
}

function saveConfigToBrowser() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(exportableConfig()));
  setGenerationStatus("Configuración guardada en este navegador. No se guardaron Excel, PNG, fuentes ni PDFs.");
}

function restoreConfig() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
    if (saved) applyConfig(saved);
  } catch (error) {
    console.warn("No se pudo restaurar la configuración", error);
  }
}

function downloadConfigJson() {
  const blob = new Blob([JSON.stringify(exportableConfig(), null, 2)], { type: "application/json;charset=utf-8" });
  saveAs(blob, `configuracion-certificados-creamas-${new Date().toISOString().slice(0, 10)}.json`);
}

function exportableConfig() {
  return {
    version: 2,
    note: "No incluye Excel, plantilla PNG, fuentes ni certificados. Solo posiciones, estilos, fuente asignada por campo y mapeo de columnas.",
    columnMap: state.config.columnMap,
    fields: state.config.fields,
  };
}

function applyConfig(imported) {
  state.config = structuredClone(DEFAULT_CONFIG);
  if (imported?.columnMap) state.config.columnMap = { ...state.config.columnMap, ...imported.columnMap };
  if (imported?.fields) {
    Object.entries(imported.fields).forEach(([key, value]) => {
      if (state.config.fields[key]) state.config.fields[key] = { ...state.config.fields[key], ...value };
    });
  }
  Object.values(state.config.fields).forEach((field) => {
    if (!field.fontId) field.fontId = DEFAULT_FONT_ID;
  });
}

function moveRecord(direction) {
  if (!state.rows.length) return;
  state.selectedRowIndex = (state.selectedRowIndex + direction + state.rows.length) % state.rows.length;
  drawPreview();
}

function getCurrentRecord() {
  return state.rows[state.selectedRowIndex] || {};
}

function getFieldText(key, record) {
  const column = state.config.columnMap[key];
  return column ? String(record[column] ?? "").trim() : "";
}

function sampleText(key) {
  const samples = {
    nombre_completo: "NOMBRE COMPLETO",
    texto_certificado: "Por su participación destacada en el taller de Crea+.",
    texto_gracias: "Gracias por ser parte de esta experiencia.",
    fecha: "12 de junio de 2026",
    codigo_certificado: "CERT-0001",
  };
  return samples[key] || "";
}

function updateRecordIndicator() {
  els.recordIndicator.textContent = state.rows.length ? `${state.selectedRowIndex + 1} / ${state.rows.length}` : "Sin datos";
}

function updateStatus() {
  const lines = [
    `Excel: ${state.rows.length ? `${state.rows.length} fila(s), ${state.headers.length} columna(s)${state.selectedSheetName ? ` · hoja “${state.selectedSheetName}”` : ""}` : "pendiente"}`,
    `Plantilla PNG: ${state.templateDataUrl ? `${state.templateSize.width} × ${state.templateSize.height}px` : "pendiente"}`,
    `Fuentes: ${state.fonts.length > 1 ? `${state.fonts.length - 1} subida(s) + respaldo del navegador` : "respaldo seguro del navegador"}`,
    "JSON: opcional, solo para recuperar posiciones y estilos anteriores.",
    "Privacidad: Excel, PNG, fuentes, JSON y PDFs se procesan localmente; no se guardan en GitHub.",
  ];
  els.fileStatus.innerHTML = lines.map((line) => `<div>${escapeHtml(line)}</div>`).join("");
}

function setGenerationStatus(message, isError = false) {
  els.generationStatus.textContent = message;
  els.generationStatus.style.color = isError ? "#b42318" : "";
}

function getFontById(id) {
  return state.fonts.find((font) => font.id === id) || state.fonts[0] || DEFAULT_FONT;
}

function fontExists(id) {
  return state.fonts.some((font) => font.id === id);
}

function uniqueFontId(baseId) {
  let id = baseId;
  let suffix = 2;
  while (fontExists(id)) {
    id = `${baseId}-${suffix}`;
    suffix += 1;
  }
  return id;
}

function normalizeHeader(header) {
  return String(header).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s-]+/g, "_");
}

function sanitizeFileName(value) {
  return String(value || "certificado").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9_-]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 90) || "certificado";
}

function readAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = src;
  });
}

function arrayBufferToBinary(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return binary;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char]));
}
