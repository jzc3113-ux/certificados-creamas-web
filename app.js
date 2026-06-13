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

const DEFAULT_FONT_ID = "browser-helvetica";
const FALLBACK_FONTS = [
  { id: DEFAULT_FONT_ID, label: "Respaldo navegador: Arial / Helvetica", family: "Arial", pdfName: "helvetica", binary: "", objectUrl: "", loaded: true, error: "", fallback: true },
  { id: "browser-times", label: "Respaldo navegador: Times", family: "Times New Roman", pdfName: "times", binary: "", objectUrl: "", loaded: true, error: "", fallback: true },
  { id: "browser-courier", label: "Respaldo navegador: Courier", family: "Courier New", pdfName: "courier", binary: "", objectUrl: "", loaded: true, error: "", fallback: true },
];
const DEFAULT_FONT = FALLBACK_FONTS[0];
const TITLE_FONT_PATTERN = /(nort|headline|cond|northeadline)/i;
const BODY_FONT_PATTERN = /(atkinson|hyperlegible)/i;

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
    nombre_completo: { x: 600, y: 300, size: 86, color: "#243047", maxWidth: 900, visible: true, uppercase: true, align: "center", fontId: DEFAULT_FONT_ID, autoFit: true, minSize: 70, maxLines: 2, overflowMode: "shrink-wrap", letterSpacing: 0, marginBottom: 14 },
    texto_certificado: { x: 600, y: 430, size: 25, color: "#243047", maxWidth: 860, visible: true, uppercase: false, align: "center", fontId: DEFAULT_FONT_ID, autoFit: false, minSize: 18, maxLines: 6, overflowMode: "wrap", letterSpacing: 0, marginBottom: 0 },
    texto_gracias: { x: 600, y: 520, size: 21, color: "#697386", maxWidth: 760, visible: true, uppercase: false, align: "center", fontId: DEFAULT_FONT_ID, autoFit: false, minSize: 16, maxLines: 4, overflowMode: "wrap", letterSpacing: 0, marginBottom: 0 },
    fecha: { x: 600, y: 635, size: 24, color: "#243047", maxWidth: 600, visible: true, uppercase: false, align: "center", fontId: DEFAULT_FONT_ID, autoFit: false, minSize: 16, maxLines: 2, overflowMode: "wrap", letterSpacing: 0, marginBottom: 0 },
    codigo_certificado: { x: 1010, y: 780, size: 16, color: "#243047", maxWidth: 280, visible: true, uppercase: false, align: "right", fontId: DEFAULT_FONT_ID, autoFit: false, minSize: 12, maxLines: 2, overflowMode: "wrap", letterSpacing: 0, marginBottom: 0 },
  },
};

const STORAGE_KEY = "creamas-certificados-config-v3";
const state = {
  workbook: null,
  sheetNames: [],
  selectedSheetName: "",
  rows: [],
  headers: [],
  templateImage: null,
  templateDataUrl: "",
  templateSize: { width: 1200, height: 850 },
  fonts: structuredClone(FALLBACK_FONTS),
  unavailableFonts: [],
  fontWarnings: [],
  config: structuredClone(DEFAULT_CONFIG),
  selectedField: "nombre_completo",
  selectedRowIndex: 0,
  drag: null,
  zipBlob: null,
  zipStale: false,
  layoutWarnings: [],
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
    largePreviewBtn: document.getElementById("largePreviewBtn"),
    largePreviewBtnActions: document.getElementById("largePreviewBtnActions"),
    previewModal: document.getElementById("previewModal"),
    closePreviewBtn: document.getElementById("closePreviewBtn"),
    largePreviewCanvas: document.getElementById("largePreviewCanvas"),
    largePrevRecordBtn: document.getElementById("largePrevRecordBtn"),
    largeNextRecordBtn: document.getElementById("largeNextRecordBtn"),
    largeRecordIndicator: document.getElementById("largeRecordIndicator"),
    saveConfigBtn: document.getElementById("saveConfigBtn"),
    downloadConfigBtn: document.getElementById("downloadConfigBtn"),
    generateBtn: document.getElementById("generateBtn"),
    downloadZipBtn: document.getElementById("downloadZipBtn"),
    newBatchBtn: document.getElementById("newBatchBtn"),
    resetAllBtn: document.getElementById("resetAllBtn"),
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
  els.sheetSelect.addEventListener("change", () => {
    markZipOutdated();
    applyWorkbookSheet(els.sheetSelect.value);
  });
  els.previewBtn.addEventListener("click", updatePreviewButton);
  els.largePreviewBtn.addEventListener("click", openLargePreview);
  els.largePreviewBtnActions.addEventListener("click", openLargePreview);
  els.closePreviewBtn.addEventListener("click", closeLargePreview);
  els.largePrevRecordBtn.addEventListener("click", () => moveRecord(-1));
  els.largeNextRecordBtn.addEventListener("click", () => moveRecord(1));
  els.previewModal.addEventListener("click", (event) => {
    if (event.target === els.previewModal) closeLargePreview();
  });
  els.saveConfigBtn.addEventListener("click", saveConfigToBrowser);
  els.downloadConfigBtn.addEventListener("click", downloadConfigJson);
  els.generateBtn.addEventListener("click", generateCertificates);
  els.downloadZipBtn.addEventListener("click", downloadZip);
  els.newBatchBtn.addEventListener("click", startNewBatch);
  els.resetAllBtn.addEventListener("click", resetAll);
  els.prevRecordBtn.addEventListener("click", () => moveRecord(-1));
  els.nextRecordBtn.addEventListener("click", () => moveRecord(1));

  els.previewCanvas.addEventListener("pointerdown", startDrag);
  els.previewCanvas.addEventListener("pointermove", dragField);
  els.previewCanvas.addEventListener("pointerup", endDrag);
  els.previewCanvas.addEventListener("pointerleave", endDrag);
}

function markZipOutdated() {
  if (!state.zipBlob) return;
  state.zipStale = true;
  setGenerationStatus("Hay cambios pendientes. Genera nuevamente los certificados para actualizar el ZIP.", true);
}

function clearGeneratedArtifacts({ keepMessage = false } = {}) {
  state.zipBlob = null;
  state.zipStale = false;
  els.downloadZipBtn.disabled = true;
  els.progressBar.value = 0;
  if (!keepMessage) setGenerationStatus("Listo para generar certificados.");
}

function clearExcelData() {
  state.workbook = null;
  state.sheetNames = [];
  state.selectedSheetName = "";
  state.rows = [];
  state.headers = [];
  state.selectedRowIndex = 0;
}

function startNewBatch() {
  clearExcelData();
  clearGeneratedArtifacts({ keepMessage: true });
  if (els.excelInput) els.excelInput.value = "";
  renderSheetSelector();
  renderMappingControls();
  drawPreview();
  closeLargePreview();
  updateStatus();
  setGenerationStatus("Nuevo lote listo. Se conservaron plantilla, fuentes, posiciones y estilos; sube otro Excel para generar un nuevo ZIP.");
}

function resetAll() {
  const confirmed = window.confirm("¿Seguro que deseas reiniciar todo? Se limpiarán los archivos cargados y la configuración actual.");
  if (!confirmed) return;
  state.fonts.filter((font) => font.objectUrl).forEach((font) => URL.revokeObjectURL(font.objectUrl));
  Object.assign(state, {
    workbook: null,
    sheetNames: [],
    selectedSheetName: "",
    rows: [],
    headers: [],
    templateImage: null,
    templateDataUrl: "",
    templateSize: { width: 1200, height: 850 },
    fonts: structuredClone(FALLBACK_FONTS),
    unavailableFonts: [],
    fontWarnings: [],
    config: structuredClone(DEFAULT_CONFIG),
    selectedField: "nombre_completo",
    selectedRowIndex: 0,
    drag: null,
    zipBlob: null,
    zipStale: false,
    layoutWarnings: [],
  });
  [els.excelInput, els.templateInput, els.fontInput, els.configInput].forEach((input) => {
    if (input) input.value = "";
  });
  localStorage.removeItem(STORAGE_KEY);
  els.previewCanvas.width = state.templateSize.width;
  els.previewCanvas.height = state.templateSize.height;
  els.largePreviewCanvas.width = state.templateSize.width;
  els.largePreviewCanvas.height = state.templateSize.height;
  renderSheetSelector();
  renderMappingControls();
  renderFieldTabs();
  renderFieldEditor();
  closeLargePreview();
  clearGeneratedArtifacts({ keepMessage: true });
  drawPreview();
  updateStatus();
  setGenerationStatus("Todo se reinició. Los archivos y la configuración actual se limpiaron del navegador.");
}

async function handleExcelUpload(event) {
  const [file] = event.target.files;
  if (!file) return;
  markZipOutdated();
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
  markZipOutdated();
  if (file.type && file.type !== "image/png") {
    setGenerationStatus("La plantilla debe ser PNG.", true);
    return;
  }
  state.templateDataUrl = await readAsDataUrl(file);
  state.templateImage = await loadImage(state.templateDataUrl);
  state.templateSize = { width: state.templateImage.naturalWidth, height: state.templateImage.naturalHeight };
  els.previewCanvas.width = state.templateSize.width;
  els.previewCanvas.height = state.templateSize.height;
  els.largePreviewCanvas.width = state.templateSize.width;
  els.largePreviewCanvas.height = state.templateSize.height;
  updateStatus();
  drawPreview();
  setGenerationStatus(`Plantilla lista: ${state.templateSize.width} × ${state.templateSize.height}px.`);
}

async function handleFontUpload(event) {
  const files = Array.from(event.target.files || []);
  if (!files.length) return;
  markZipOutdated();

  const uploadedFonts = [];
  for (const file of files) {
    const font = await createFontFromFile(file);
    uploadedFonts.push(font);
    state.fonts.push(font);
  }

  resolveSavedFontAssignments();
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
  const titleFont = findFontByPattern(TITLE_FONT_PATTERN, uploadedFonts) || null;
  const bodyFont = findFontByPattern(BODY_FONT_PATTERN, uploadedFonts) || null;

  if (titleFont && state.config.fields.nombre_completo.fontId === DEFAULT_FONT_ID) {
    state.config.fields.nombre_completo.fontId = titleFont.id;
    state.config.fields.nombre_completo.font = titleFont.label;
  }

  if (bodyFont) {
    ["texto_certificado", "texto_gracias", "fecha", "codigo_certificado"].forEach((key) => {
      if (state.config.fields[key].fontId === DEFAULT_FONT_ID) {
        state.config.fields[key].fontId = bodyFont.id;
        state.config.fields[key].font = bodyFont.label;
      }
    });
  }
}

function findFontByPattern(pattern, fonts = state.fonts) {
  return fonts.find((font) => pattern.test(font.label) || pattern.test(font.id));
}

async function handleConfigUpload(event) {
  const [file] = event.target.files;
  if (!file) return;
  try {
    markZipOutdated();
    const imported = JSON.parse(await file.text());
    applyConfig(imported);
    renderMappingControls();
    renderFieldTabs();
    renderFieldEditor();
    drawPreview();
    const missingFonts = state.unavailableFonts.length ? ` Fuentes no disponibles: ${state.unavailableFonts.join(", ")}. Se usó respaldo hasta que subas esas fuentes.` : "";
    setGenerationStatus(`Configuración JSON cargada. Recuerda que el JSON solo guarda posiciones, estilos, fuentes asignadas y columnas; no trae Excel ni PNG.${missingFonts}`, Boolean(state.unavailableFonts.length));
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
      markZipOutdated();
      drawPreview();
      drawLargePreview();
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
    <div class="inline">
      ${numberInput("Tamaño mínimo", "minSize", config.minSize ?? Math.max(8, config.size - 8))}
      ${numberInput("Máx. líneas", "maxLines", config.maxLines ?? 4)}
    </div>
    <div class="inline">
      ${numberInput("Interletrado", "letterSpacing", config.letterSpacing ?? 0, "0.1")}
      ${numberInput("Espacio inferior", "marginBottom", config.marginBottom ?? 0)}
    </div>
    <label><span>Modo overflow</span><select data-prop="overflowMode">
      <option value="shrink-wrap">Reducir y partir</option>
      <option value="shrink">Reducir</option>
      <option value="wrap">Partir líneas</option>
    </select></label>
    <label><span>Color</span><input data-prop="color" type="color" value="${config.color}"></label>
    <label><span>Alineación</span><select data-prop="align">
      <option value="left">Izquierda</option>
      <option value="center">Centro</option>
      <option value="right">Derecha</option>
    </select></label>
    <label class="checkbox-row"><input data-prop="visible" type="checkbox" ${config.visible ? "checked" : ""}> Visible</label>
    <label class="checkbox-row"><input data-prop="uppercase" type="checkbox" ${config.uppercase ? "checked" : ""}> Mayúsculas</label>
    <label class="checkbox-row"><input data-prop="autoFit" type="checkbox" ${config.autoFit ? "checked" : ""}> Autoajuste</label>
  `;
  els.fieldEditor.querySelector('[data-prop="fontId"]').value = fontExists(config.fontId) ? config.fontId : DEFAULT_FONT_ID;
  els.fieldEditor.querySelector('[data-prop="align"]').value = config.align;
  els.fieldEditor.querySelector('[data-prop="overflowMode"]').value = config.overflowMode || "wrap";
  els.fieldEditor.querySelectorAll("[data-prop]").forEach((input) => {
    input.addEventListener("input", updateSelectedFieldFromInput);
    input.addEventListener("change", updateSelectedFieldFromInput);
  });
}

function fontOptionsHtml() {
  return state.fonts.map((font) => `<option value="${escapeHtml(font.id)}">${escapeHtml(font.label)}${font.error ? " (respaldo si falla)" : ""}</option>`).join("");
}

function numberInput(label, prop, value, step = "1") {
  const numericValue = Number(value || 0);
  const formattedValue = step === "1" ? numericValue.toFixed(0) : String(numericValue);
  return `<label><span>${label}</span><input data-prop="${prop}" type="number" step="${step}" value="${formattedValue}"></label>`;
}

function updateSelectedFieldFromInput(event) {
  const input = event.target;
  const prop = input.dataset.prop;
  const config = state.config.fields[state.selectedField];
  if (input.type === "checkbox") config[prop] = input.checked;
  else if (input.type === "number") config[prop] = Number(input.value || 0);
  else config[prop] = input.value;
  if (prop === "fontId") config.font = getFontById(config.fontId).label;
  markZipOutdated();
  drawPreview();
  drawLargePreview();
}

function updatePreviewButton() {
  drawPreview();
  drawLargePreview();
  if (!state.rows.length || !state.templateDataUrl) {
    setGenerationStatus("Sube primero un Excel y una plantilla PNG para generar la vista previa.", true);
    return;
  }
  const warning = state.layoutWarnings.length ? ` ${state.layoutWarnings.join(" ")}` : "";
  setGenerationStatus(`Vista previa actualizada.${warning}`, Boolean(state.layoutWarnings.length));
}

function drawPreview() {
  renderCertificateToCanvas(getCurrentRecord(), els.previewCanvas, { showSelection: true });
  updateRecordIndicator();
}

function drawLargePreview() {
  if (!els.largePreviewCanvas || els.previewModal.hidden) return;
  renderCertificateToCanvas(getCurrentRecord(), els.largePreviewCanvas, { showSelection: false });
  updateRecordIndicator();
}

function renderCertificateToCanvas(record, canvas, { showSelection = false, resetWarnings = true } = {}) {
  const ctx = canvas.getContext("2d");
  const { width, height } = canvas;
  ctx.clearRect(0, 0, width, height);
  if (resetWarnings) state.layoutWarnings = [];

  if (state.templateImage) {
    ctx.drawImage(state.templateImage, 0, 0, width, height);
  } else {
    drawPlaceholderTemplate(ctx, width, height);
  }

  let yOffset = 0;
  CERTIFICATE_FIELDS.forEach((field, index) => {
    const originalY = state.config.fields[field.key].y;
    const yOverride = index === 0 ? originalY : originalY + yOffset;
    const result = drawFieldOnCanvas(ctx, field.key, record, { showSelection, yOverride });
    if (!result) return;
    const nextField = CERTIFICATE_FIELDS[index + 1];
    if (!nextField) return;
    const desiredGap = Number(state.config.fields[field.key].marginBottom || 0);
    yOffset += desiredGap;
    const nextFieldY = state.config.fields[nextField.key].y;
    const neededOffset = Math.max(0, result.bottom - (nextFieldY + yOffset));
    yOffset += neededOffset;
  });
  return canvas;
}


function openLargePreview() {
  drawPreview();
  if (!state.rows.length || !state.templateDataUrl) {
    setGenerationStatus("Sube primero un Excel y una plantilla PNG para generar la vista previa.", true);
    return;
  }
  els.largePreviewCanvas.width = state.templateSize.width;
  els.largePreviewCanvas.height = state.templateSize.height;
  els.previewModal.hidden = false;
  document.body.classList.add("modal-open");
  drawLargePreview();
  const warning = state.layoutWarnings.length ? ` ${state.layoutWarnings.join(" ")}` : "";
  setGenerationStatus(`Vista previa grande actualizada.${warning}`, Boolean(state.layoutWarnings.length));
}

function closeLargePreview() {
  els.previewModal.hidden = true;
  document.body.classList.remove("modal-open");
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

function drawFieldOnCanvas(ctx, key, record, { showSelection = true, yOverride = null } = {}) {
  const fieldConfig = { ...state.config.fields[key] };
  if (!fieldConfig.visible) return null;
  if (yOverride !== null) fieldConfig.y = yOverride;
  const rawText = getFieldText(key, record) || (state.rows.length ? "" : sampleText(key));
  if (!rawText) return null;
  const font = getFontById(fieldConfig.fontId);
  const layout = layoutCanvasField(ctx, key, rawText, fieldConfig, font);
  ctx.save();
  ctx.fillStyle = fieldConfig.color;
  ctx.textAlign = "left";
  ctx.textBaseline = "top";
  renderRichCanvasLines(ctx, layout.lines, fieldConfig, font, layout.size, key);

  if (showSelection && key === state.selectedField) {
    ctx.strokeStyle = "rgba(46, 184, 114, 0.85)";
    ctx.lineWidth = 2;
    const startX = alignedStartX(fieldConfig.x, fieldConfig.maxWidth, fieldConfig.align);
    ctx.strokeRect(startX - 8, fieldConfig.y - 8, fieldConfig.maxWidth + 16, layout.height + 16);
  }
  ctx.restore();

  if (layout.warning) addLayoutWarning(layout.warning);
  return { bottom: fieldConfig.y + layout.height, height: layout.height, lines: layout.lines, warning: layout.warning };
}

function layoutCanvasField(ctx, key, rawText, config, font) {
  const segments = parseRichText(rawText, config.uppercase);
  const baseSize = Number(config.size || 16);
  const minSize = Math.min(baseSize, Number(config.minSize || baseSize));
  const maxLines = Math.max(1, Number(config.maxLines || 99));
  const mode = config.overflowMode || (config.autoFit ? "shrink-wrap" : "wrap");
  const letterSpacing = Number(config.letterSpacing || 0);
  const measure = (text, bold, size) => {
    setCanvasFont(ctx, font, size, bold || key === "nombre_completo");
    return measureTextWithLetterSpacing(ctx, text, letterSpacing);
  };

  if (config.autoFit && mode.includes("shrink")) {
    for (let size = baseSize; size >= minSize; size -= 1) {
      const oneLine = buildRichLines(segments, (text, bold) => measure(text, bold, size), config.maxWidth, 1);
      if (!oneLine.truncated && oneLine.lines.length <= 1) return canvasLayoutResult(oneLine.lines, size, config);
    }
  }

  const wrapSize = config.autoFit && mode === "shrink-wrap" ? minSize : baseSize;
  const wrapped = buildRichLines(segments, (text, bold) => measure(text, bold, wrapSize), config.maxWidth, mode === "shrink" ? 1 : maxLines);
  const result = canvasLayoutResult(wrapped.lines, wrapSize, config);
  if (wrapped.truncated && key === "nombre_completo") {
    result.warning = "El nombre es demasiado largo para el espacio disponible. Reduce tamaño o aumenta ancho máximo.";
  }
  return result;
}

function canvasLayoutResult(lines, size, config) {
  const lineHeight = size * 1.18;
  return { lines, size, lineHeight, height: Math.max(lineHeight, lines.length * lineHeight), warning: "" };
}

function renderRichCanvasLines(ctx, lines, config, font, size, key) {
  const lineHeight = size * 1.18;
  lines.forEach((line, lineIndex) => {
    let currentX = alignedStartX(config.x, line.width, config.align);
    const y = config.y + lineIndex * lineHeight;
    line.segments.forEach((segment) => {
      const bold = segment.bold || key === "nombre_completo";
      setCanvasFont(ctx, font, size, bold);
      drawTextWithLetterSpacing(ctx, segment.text, currentX, y, Number(config.letterSpacing || 0));
      currentX += measureTextWithLetterSpacing(ctx, segment.text, Number(config.letterSpacing || 0));
    });
  });
}

function setCanvasFont(ctx, font, size, bold = false) {
  ctx.font = `${bold ? "700" : "400"} ${size}px "${font.family}", Arial, Helvetica, sans-serif`;
}

function measureTextWithLetterSpacing(ctx, text, letterSpacing = 0) {
  const characters = Array.from(String(text));
  if (!characters.length) return 0;
  const textWidth = characters.reduce((total, character) => total + ctx.measureText(character).width, 0);
  return textWidth + Math.max(0, characters.length - 1) * letterSpacing;
}

function drawTextWithLetterSpacing(ctx, text, x, y, letterSpacing = 0) {
  if (!letterSpacing) {
    ctx.fillText(text, x, y);
    return;
  }
  let currentX = x;
  Array.from(String(text)).forEach((character) => {
    ctx.fillText(character, currentX, y);
    currentX += ctx.measureText(character).width + letterSpacing;
  });
}

function alignedStartX(anchorX, width, align = "left") {
  if (align === "center") return anchorX - width / 2;
  if (align === "right") return anchorX - width;
  return anchorX;
}

function parseRichText(text, uppercase = false) {
  const normalized = String(text).replace(/\\n/g, "\n");
  return normalized.split("**").map((part, index) => ({
    text: uppercase ? part.toUpperCase() : part,
    bold: index % 2 === 1,
  })).filter((segment) => segment.text.length);
}

function buildRichLines(segments, measure, maxWidth, maxLines = Infinity) {
  const lines = [];
  let current = emptyRichLine();
  let truncated = false;

  const pushLine = () => {
    trimRichLine(current);
    if (current.segments.length || !lines.length) lines.push(current);
    current = emptyRichLine();
    if (lines.length >= maxLines) truncated = true;
  };

  outer: for (const segment of segments) {
    const tokens = segment.text.split(/(\n|\s+)/).filter((token) => token.length);
    for (const token of tokens) {
      if (truncated) break outer;
      if (token === "\n") {
        pushLine();
        continue;
      }
      const isSpace = /^\s+$/.test(token);
      const text = isSpace ? " " : token;
      if (isSpace && !current.segments.length) continue;
      const width = measure(text, segment.bold);
      if (!isSpace && current.segments.length && current.width + width > maxWidth) {
        pushLine();
        if (truncated) break outer;
      }
      current.segments.push({ text, bold: segment.bold, width });
      current.width += width;
    }
  }
  if (!truncated && current.segments.length) pushLine();
  return { lines: lines.slice(0, maxLines), truncated };
}

function emptyRichLine() {
  return { segments: [], width: 0 };
}

function trimRichLine(line) {
  while (line.segments.length && /^\s+$/.test(line.segments[line.segments.length - 1].text)) {
    const removed = line.segments.pop();
    line.width = Math.max(0, line.width - (removed.width || 0));
  }
  while (line.segments.length && /^\s+$/.test(line.segments[0].text)) {
    const removed = line.segments.shift();
    line.width = Math.max(0, line.width - (removed.width || 0));
  }
}

function addLayoutWarning(message) {
  if (!state.layoutWarnings.includes(message)) state.layoutWarnings.push(message);
}

function wrapCanvasText(ctx, text, x, y, maxWidth, lineHeight) {
  const layout = buildRichLines(parseRichText(text), (token) => ctx.measureText(token).width, maxWidth, Infinity);
  layout.lines.forEach((line, index) => ctx.fillText(line.segments.map((segment) => segment.text).join(""), x, y + index * lineHeight));
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
  drawLargePreview();
}

function dragField(event) {
  if (!state.drag) return;
  const point = getCanvasPoint(event);
  const config = state.config.fields[state.drag.key];
  config.x = Math.round(point.x - state.drag.offsetX);
  config.y = Math.round(point.y - state.drag.offsetY);
  renderFieldEditor();
  drawPreview();
  drawLargePreview();
}

function endDrag() {
  if (state.drag) markZipOutdated();
  state.drag = null;
  els.previewCanvas.classList.remove("dragging");
}

function hitTestField(point) {
  return [...CERTIFICATE_FIELDS].reverse().find((field) => {
    const config = state.config.fields[field.key];
    if (!config.visible) return false;
    const x = config.align === "center" ? config.x - config.maxWidth / 2 : config.align === "right" ? config.x - config.maxWidth : config.x;
    const y = config.y;
    const estimatedHeight = Math.max(config.size * (config.maxLines || 2) * 1.2, 48);
    return point.x >= x - 12 && point.x <= x + config.maxWidth + 12 && point.y >= y - 12 && point.y <= y + estimatedHeight;
  })?.key;
}

async function generateCertificates() {
  if (!validateBeforeGenerate()) return;

  els.generateBtn.disabled = true;
  els.downloadZipBtn.disabled = true;
  state.zipBlob = null;
  state.zipStale = false;
  els.progressBar.value = 0;
  const zip = new JSZip();
  state.fontWarnings = [];
  state.layoutWarnings = [];

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
    state.zipStale = false;
    els.downloadZipBtn.disabled = false;
    const warnings = [...state.fontWarnings, ...state.layoutWarnings];
    const warningText = warnings.length ? ` Aviso: ${warnings.join(" ")}` : "";
    setGenerationStatus(`ZIP listo con ${state.rows.length} certificado(s).${warningText}`, Boolean(warnings.length));
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
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  if (document.fonts?.ready) await document.fonts.ready;
  renderCertificateToCanvas(record, canvas, { showSelection: false, resetWarnings: false });
  const imageData = canvas.toDataURL("image/png");
  pdf.addImage(imageData, "PNG", 0, 0, width, height);
  return pdf.output("blob");
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
  if (state.zipStale) setGenerationStatus("Descargando el último ZIP generado. Hay cambios pendientes; genera nuevamente para actualizarlo.", true);
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
  const fields = structuredClone(state.config.fields);
  Object.entries(fields).forEach(([key, field]) => {
    field.font = getFontById(field.fontId).label;
    fields[key] = field;
  });
  const campos = Object.fromEntries(Object.entries(fields).map(([key, field]) => [key, { ...field, font: field.font }]));
  return {
    version: 3,
    note: "No incluye Excel, plantilla PNG, fuentes ni certificados. Solo posiciones, estilos, fuente asignada por campo y mapeo de columnas.",
    columnMap: state.config.columnMap,
    fields,
    campos,
  };
}

function applyConfig(imported) {
  state.config = structuredClone(DEFAULT_CONFIG);
  state.unavailableFonts = [];
  const importedFields = imported?.fields || imported?.campos || {};
  if (imported?.columnMap) state.config.columnMap = { ...state.config.columnMap, ...imported.columnMap };
  Object.entries(importedFields).forEach(([key, value]) => {
    if (!state.config.fields[key]) return;
    state.config.fields[key] = { ...state.config.fields[key], ...value };
    const savedFont = value.font || value.fontName || value.fontFamily || value.fontId;
    if (savedFont) applySavedFontToField(key, savedFont);
  });
  Object.values(state.config.fields).forEach((field) => {
    if (!field.fontId || !fontExists(field.fontId)) field.fontId = DEFAULT_FONT_ID;
    field.font = getFontById(field.fontId).label;
  });
}

function applySavedFontToField(key, savedFont) {
  const font = findFontBySavedName(savedFont);
  if (font) {
    state.config.fields[key].fontId = font.id;
    state.config.fields[key].font = font.label;
    delete state.config.fields[key].pendingFont;
    return;
  }
  state.config.fields[key].fontId = DEFAULT_FONT_ID;
  state.config.fields[key].font = getFontById(DEFAULT_FONT_ID).label;
  state.config.fields[key].pendingFont = savedFont;
  if (!state.unavailableFonts.includes(savedFont)) state.unavailableFonts.push(savedFont);
}

function resolveSavedFontAssignments() {
  Object.entries(state.config.fields).forEach(([key, field]) => {
    if (!field.pendingFont) return;
    const font = findFontBySavedName(field.pendingFont);
    if (!font) return;
    field.fontId = font.id;
    field.font = font.label;
    state.unavailableFonts = state.unavailableFonts.filter((name) => name !== field.pendingFont);
    delete field.pendingFont;
  });
}

function findFontBySavedName(savedFont) {
  const normalizedSaved = normalizeFontName(savedFont);
  return state.fonts.find((font) => {
    const candidates = [font.id, font.label, font.family, font.pdfName].map(normalizeFontName);
    return candidates.includes(normalizedSaved);
  });
}

function moveRecord(direction) {
  if (!state.rows.length) return;
  state.selectedRowIndex = (state.selectedRowIndex + direction + state.rows.length) % state.rows.length;
  drawPreview();
  drawLargePreview();
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
  const text = state.rows.length ? `${state.selectedRowIndex + 1} / ${state.rows.length}` : "Sin datos";
  els.recordIndicator.textContent = text;
  if (els.largeRecordIndicator) els.largeRecordIndicator.textContent = text;
}

function updateStatus() {
  const lines = [
    `Excel: ${state.rows.length ? `${state.rows.length} fila(s), ${state.headers.length} columna(s)${state.selectedSheetName ? ` · hoja “${state.selectedSheetName}”` : ""}` : "pendiente"}`,
    `Plantilla PNG: ${state.templateDataUrl ? `${state.templateSize.width} × ${state.templateSize.height}px` : "pendiente"}`,
    `Fuentes: ${uploadedFontCount() ? `${uploadedFontCount()} subida(s) + respaldos del navegador` : "respaldos seguros del navegador"}`,
    "JSON: opcional, solo para recuperar posiciones y estilos anteriores.",
    "Privacidad: Excel, PNG, fuentes, JSON y PDFs se procesan localmente; no se guardan en GitHub.",
  ];
  els.fileStatus.innerHTML = lines.map((line) => `<div>${escapeHtml(line)}</div>`).join("");
}

function setGenerationStatus(message, isError = false) {
  els.generationStatus.textContent = message;
  els.generationStatus.style.color = isError ? "#b42318" : "";
}

function uploadedFontCount() {
  return state.fonts.filter((font) => !font.fallback).length;
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

function normalizeFontName(name) {
  return normalizeHeader(name).replace(/^font_?/, "");
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
