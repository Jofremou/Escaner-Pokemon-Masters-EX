"use strict";


const SHEETS_WEB_APP_URL =
  " https://script.google.com/macros/s/AKfycbw2kJ6tzPQ-TUMUVsyfQ6eqK2998-EC5K_amGzRBrvYdQWDow_AnjXi8RItH6QiNS_4/exec ";

const SHEETS_SYNC_TOKEN =

/**
 * pairs debe contener objetos con esta estructura:
 * [
 *   {
 *     name: "Red (Classic) & Snorlax",
 *     moveLevel: 3,
 *     ex: true,
 *     exRole: true
 *   }
 * ]
 */
async function sincronizarConGoogleSheets(player, pairs) {
  if (!player) {
    throw new Error("Selecciona un jugador.");
  }

  if (!Array.isArray(pairs) || pairs.length === 0) {
    throw new Error("No hay parejas para sincronizar.");
  }

  const payload = {
    token: SHEETS_SYNC_TOKEN,
    player: player,
    pairs: pairs.map(pair => ({
      name: pair.name,
      moveLevel: Number(pair.moveLevel),
      ex: pair.ex,
      exRole: pair.exRole
    }))
  };

  /*
   * Al abrir el HTML desde file://, esta petición
   * usa no-cors y text/plain para evitar una petición
   * CORS de preflight.
   *
   * La respuesta no será legible desde el navegador:
   * que fetch termine no demuestra que el servidor
   * haya procesado correctamente la sincronización.
   */
  await fetch(SHEETS_WEB_APP_URL, {
    method: "POST",
    mode: "no-cors",
    headers: {
      "Content-Type": "text/plain;charset=utf-8"
    },
    body: JSON.stringify(payload)
  });

  alert(
    "Solicitud enviada. Comprueba Google Sheets " +
    "para confirmar que la sincronización se ha completado."
  );
}


const MAX_CARDS = 1000;
const MOVE_TYPES = [
  "Normal",
  "Fire",
  "Water",
  "Electric",
  "Grass",
  "Ice",
  "Fighting",
  "Poison",
  "Ground",
  "Flying",
  "Psychic",
  "Bug",
  "Rock",
  "Ghost",
  "Dragon",
  "Dark",
  "Steel",
  "Fairy",
  "Stellar",
];
const MOVE_TYPE_SAMPLE_SIZE = 16;
const ROLE_SAMPLE_SIZE = 20;
const COPY_BADGE_SAMPLE_SIZE = 24;
const COPY_BADGE_SHAPE_BORDER = 2;
const PAIR_ROLES = [
  { role: "Strike", label: "Atacante", icon: "strike" },
  { role: "Tech", label: "Sabotaje", icon: "tech" },
  { role: "Support", label: "Support", icon: "support" },
  { role: "Sprint", label: "Velocidad", icon: "sprint" },
  { role: "Field", label: "Entorno", icon: "field" },
  { role: "Multi", label: "Multi", icon: "multi" },
];
// Sync Pairs Tracker has no Stellar badge asset, so use a masked sample from the supplied Terapagos card.
const STELLAR_TYPE_ICON_DATA_URL =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAwAAAAMCAYAAABWdVznAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAIjSURBVChTBcHPT1IBAAfw7x/TqUuXLnWrS22tk85l5tKsZhQRNtPSlqIDRdCWPwaoGCQPoUe91EAL0UDMIrFEFI0s9Ymm0/KFbpbbt88He9ku7O13YyGlR3ikEiGpCpMDVYj77+FrSIOYmIehhhK8bNBh8NkrYPevC0nZCJe/CP5INdNLIpfn3UxPtXM7aebujJ6rlnKIJZfgrtABSwcC2oaLUefNpRirZ0YZ5972e67Ni1QW3DxYkvjL3sjAlauwFaiBt1t25JhPUy0UsX3UwPGExPikyI9BJ1MhN+Wx18w4ejh6V0fLhXJAWPPgVN1FPnjjZFSRORgb4YdEhJvbK1RW1/jJE+Sf6AJ/D8xwsMIO2DfGcFyjpsoVpH/9kJ3ReXrjs4xndplOZznimOJmWOF6/xYFrQ+wroRxtOgWz+n6aBr9yWpvlE+GJxiYlTkUStFu9DPmXea71gRNhT2ATZ7GCW0Lj+Q287wuSHvyH6XvCrvDSTr6J5n4nKXXMsfKfCdLz7QAXTtzUEkRHlO18azRT9vyIV9s7tOZ+MaBqRQXV7LsdSR5Pd9CdaEVeKoswvojgzyrxJPVFhb6IlQHJnhDCFDb6mGNQWBxSScL8jqgrwkC959bUePzoFYaRmmPxJwWF3PMAvOb3Lz8yMFrd6xU3fTB3PgFvQ4ZKDM+hLZJB41Bj9v1zdAYbNAYe1Fm8qC2YwiPuyJwOWWI7h30CRv4D19Jc8W55leUAAAAAElFTkSuQmCC";
const MOVE_TYPE_LABELS = {
  Normal: "Normal",
  Fire: "Fuego",
  Water: "Agua",
  Electric: "Eléctrico",
  Grass: "Planta",
  Ice: "Hielo",
  Fighting: "Lucha",
  Poison: "Veneno",
  Ground: "Tierra",
  Flying: "Volador",
  Psychic: "Psíquico",
  Bug: "Bicho",
  Rock: "Roca",
  Ghost: "Fantasma",
  Dragon: "Dragón",
  Dark: "Siniestro",
  Steel: "Acero",
  Fairy: "Hada",
  Stellar: "Astral",
};
const imageInput = document.querySelector("#image-file");
const imageName = document.querySelector("#image-name");
const imageWorkspace = document.querySelector("#image-workspace");
const sourceImage = document.querySelector("#source-image");
const cropOverlay = document.querySelector("#crop-overlay");
const cardsList = document.querySelector("#cards-list");
const cropStatus = document.querySelector("#crop-status");
const reviewCount = document.querySelector("#review-count");
const exportButton = document.querySelector("#export-button");
const recognizeButton = document.querySelector("#recognize-button");
const suggestionsExportButton = document.querySelector("#suggestions-export-button");
const catalogStatus = document.querySelector("#catalog-status");
const analysisStatus = document.querySelector("#analysis-status");
const errorMessage = document.querySelector("#app-error");
const gridInputs = [
  "#columns",
  "#rows",
  "#crop-x",
  "#crop-y",
  "#crop-width",
  "#crop-height",
].map((selector) => document.querySelector(selector));

let loadedImage = null;
let cards = [];
let renderScheduled = false;
let catalogPromise = null;
let catalogRecords = [];
let excludedCropIndexes = new Set();
let excludedCropSignature = "";
let moveTypeTemplatesPromise = null;
let pairRoleTemplatesPromise = null;
let copyBadgeTemplatesPromise = null;
let lastCropSignature = "";
let catalogImageUrls = new Set();
let ocrWorkerPromise = null;

imageInput.addEventListener("change", handleImageSelection);
gridInputs.forEach((input) => input.addEventListener("change", scheduleGridRender));
exportButton.addEventListener("click", exportCsv);
recognizeButton.addEventListener("click", analyzeAndRecognizeCards);
suggestionsExportButton.addEventListener("click", exportSuggestionsCsv);

function showError(message) {
  errorMessage.textContent = message;
  errorMessage.hidden = !message;
}

function moveTypeLabel(type) {
  return MOVE_TYPE_LABELS[type] || type;
}

function pairRoleLabel(role) {
  return PAIR_ROLES.find((entry) => entry.role === role)?.label || role;
}

function normalizePairRole(role) {
  if (role === "Strike (Physical)" || role === "Strike (Special)") return "Strike";
  return PAIR_ROLES.some((entry) => entry.role === role) ? role : "";
}

function readGrid() {
  const [columnsInput, rowsInput, xInput, yInput, widthInput, heightInput] = gridInputs;
  const grid = {
    columns: Number(columnsInput.value),
    rows: Number(rowsInput.value),
    x: Number(xInput.value),
    y: Number(yInput.value),
    width: Number(widthInput.value),
    height: Number(heightInput.value),
  };

  if (
    !Number.isInteger(grid.columns) ||
    !Number.isInteger(grid.rows) ||
    grid.columns < 1 ||
    grid.rows < 1 ||
    grid.columns > 20 ||
    grid.rows > 50 ||
    grid.columns * grid.rows > MAX_CARDS
  ) {
    throw new Error(`La cuadrícula debe tener entre 1 y ${MAX_CARDS} tarjetas.`);
  }

  if (
    ![grid.x, grid.y, grid.width, grid.height].every(Number.isFinite) ||
    grid.x < 0 ||
    grid.y < 0 ||
    grid.width <= 0 ||
    grid.height <= 0 ||
    grid.x + grid.width > 100 ||
    grid.y + grid.height > 100
  ) {
    throw new Error("El área de recorte debe quedar dentro de la imagen.");
  }

  return grid;
}

function detectGridStartY(image, grid) {
  const canvas = document.createElement("canvas");
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;

  const context = canvas.getContext("2d", {
    willReadFrequently: true
  });

  if (!context) return null;

  context.drawImage(image, 0, 0);

  const imageWidth = canvas.width;
  const imageHeight = canvas.height;
  const pixels = context.getImageData(
    0, 0, imageWidth, imageHeight
  ).data;

  const startX = Math.floor(imageWidth * grid.x / 100);
  const endX = Math.floor(
    imageWidth * (grid.x + grid.width) / 100
  );

  const rowHeight =
    imageHeight * grid.height / 100 / grid.rows;

  // Energia dels canvis visuals entre línies horitzontals.
  const edgeScores = new Float32Array(imageHeight);

  for (let y = 1; y < imageHeight; y++) {
    let score = 0;
    let samples = 0;

    for (let x = startX; x < endX; x += 2) {
      const above = (y - 1) * imageWidth * 4 + x * 4;
      const below = y * imageWidth * 4 + x * 4;

      const lumAbove =
        0.299 * pixels[above] +
        0.587 * pixels[above + 1] +
        0.114 * pixels[above + 2];

      const lumBelow =
        0.299 * pixels[below] +
        0.587 * pixels[below + 1] +
        0.114 * pixels[below + 2];

      score += Math.abs(lumBelow - lumAbove);
      samples++;
    }

    edgeScores[y] = samples ? score / samples : 0;
  }

  const minY = 0;
  const maxY = Math.min(
    Math.round(imageHeight * 0.10),
    Math.floor(imageHeight * (100 - grid.height) / 100)
  );

  const rowSpacing = rowHeight;
  let bestY = null;
  let bestScore = -Infinity;

  for (let y = minY; y <= maxY; y++) {
    let score = 0;

    // Comprova si els límits de les files coincideixen
    // amb canvis visuals horitzontals repetits.
    for (let row = 0; row <= grid.rows; row++) {
      const boundary = Math.round(y + row * rowSpacing);
      let strongestEdge = 0;

      for (let offset = -2; offset <= 2; offset++) {
        const sampleY = boundary + offset;

        if (sampleY >= 1 && sampleY < imageHeight) {
          strongestEdge = Math.max(
            strongestEdge,
            edgeScores[sampleY]
          );
        }
      }

      score += strongestEdge;
    }

    if (score > bestScore) {
      bestScore = score;
      bestY = y;
    }
  }

  if (bestY === null) return null;

  return Number((bestY / imageHeight * 100).toFixed(2));
}

async function handleImageSelection(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  showError("");
  if (!file.type.startsWith("image/")) {
    showError("El archivo seleccionado no parece ser una imagen.");
    imageInput.value = "";
    return;
  }

  try {
    const previousUrl = loadedImage?.url;
    const imageUrl = URL.createObjectURL(file);
    const image = new Image();
    image.src = imageUrl;
    await image.decode();

    loadedImage = { image, url: imageUrl };
    if (previousUrl) URL.revokeObjectURL(previousUrl);
    cards = [];
    lastCropSignature = "";
    excludedCropIndexes.clear();
    excludedCropSignature = "";
    if (image.naturalWidth <= 250 && image.naturalHeight <= 250) {
      setGridValue("#columns", 1);
      setGridValue("#rows", 1);
      setGridValue("#crop-x", 0);
      setGridValue("#crop-y", 0);
      setGridValue("#crop-width", 100);
      setGridValue("#crop-height", 100);
    } else {
      setGridValue("#columns", 4);
      setGridValue("#rows", 7);
      setGridValue("#crop-x", 7.5);
      setGridValue("#crop-y", 13.4);
      setGridValue("#crop-width", 86);
      setGridValue("#crop-height", 68.4);

      const detectedY = detectGridStartY(image, {
        columns: 4,
        rows: 7,
        x: 7.5,
        width: 86,
        height: 68.4
      });

      if (detectedY !== null) {
        setGridValue("#crop-y", detectedY);
      }
    }
    sourceImage.src = imageUrl;
    imageName.textContent = `${file.name} · ${image.naturalWidth} × ${image.naturalHeight} px`;
    imageWorkspace.hidden = false;
    recognizeButton.disabled = false;
    catalogStatus.textContent = "";
    analysisStatus.textContent = "";
    renderGrid();
  } catch (error) {
    showError(`No se pudo abrir la imagen: ${error.message}`);
    imageInput.value = "";
  }
}

function setGridValue(selector, value) {
  const input = document.querySelector(selector);
  input.value = String(value);
}

function scheduleGridRender() {
  if (!loadedImage || renderScheduled) return;
  renderScheduled = true;
  requestAnimationFrame(() => {
    renderScheduled = false;
    renderGrid();
  });
}

function renderGrid() {
  if (!loadedImage) return;

  try {
    showError("");

    // Leer la configuración actual de la cuadrícula.
    const grid = readGrid();

    // Dibujar la cuadrícula con la coordenada Y actual.
    drawOverlay(grid);

    // Generar los recortes.
    buildCrops(grid);

    // Actualizar el estado de los recortes.
    cropStatus.textContent =
      `${cards.length} tarjetas incluidas · ` +
      `área ${grid.width}% × ${grid.height}%`;

    // Actualizar la interfaz.
    renderCards();
    updateSummary();

  } catch (error) {
    cards = [];
    cropOverlay.replaceChildren();
    cardsList.replaceChildren();
    cropStatus.textContent = "";

    reviewCount.textContent =
      "Corrige los ajustes de recorte para continuar.";

    exportButton.disabled = true;
    recognizeButton.disabled = true;
    suggestionsExportButton.disabled = true;

    showError(error.message);
  }
}

function drawOverlay(grid) {
  const fragment = document.createDocumentFragment();
  const cellWidth = grid.width / grid.columns;
  const cellHeight = grid.height / grid.rows;

  for (let index = 0; index < grid.columns * grid.rows; index += 1) {
    if (excludedCropIndexes.has(index)) continue;
    const column = index % grid.columns;
    const row = Math.floor(index / grid.columns);
    const cell = document.createElement("span");
    cell.className = "crop-cell";
    cell.style.left = `${grid.x + column * cellWidth}%`;
    cell.style.top = `${grid.y + row * cellHeight}%`;
    cell.style.width = `${cellWidth}%`;
    cell.style.height = `${cellHeight}%`;
    fragment.append(cell);
  }

  cropOverlay.replaceChildren(fragment);
}

function buildCrops(grid) {
  const signature = [
    loadedImage.image.naturalWidth,
    loadedImage.image.naturalHeight,
    grid.columns,
    grid.rows,
    grid.x,
    grid.y,
    grid.width,
    grid.height,
  ].join(":");
  if (signature !== lastCropSignature) {
    excludedCropIndexes = new Set();
    excludedCropSignature = signature;
  }
  const oldCards = signature === lastCropSignature
    ? new Map(cards.map((card) => [card.cropIndex, card]))
    : new Map();
  const cellWidth = (loadedImage.image.naturalWidth * grid.width) / 100 / grid.columns;
  const cellHeight = (loadedImage.image.naturalHeight * grid.height) / 100 / grid.rows;
  const cropWidth = cellWidth * 0.92;
  const cropHeight = cellHeight * 0.94;
  const startX = (loadedImage.image.naturalWidth * grid.x) / 100;
  const startY = (loadedImage.image.naturalHeight * grid.y) / 100;
  const nextCards = [];

  for (let index = 0; index < grid.columns * grid.rows; index += 1) {
    const column = index % grid.columns;
    const row = Math.floor(index / grid.columns);
    const x = startX + column * cellWidth + (cellWidth - cropWidth) / 2;
    const y = startY + row * cellHeight + (cellHeight - cropHeight) / 2;
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(cropWidth));
    canvas.height = Math.max(1, Math.round(cropHeight));

    const context = canvas.getContext("2d");
    if (!context) throw new Error("Este navegador no permite recortar la imagen.");
    context.drawImage(
      loadedImage.image,
      x,
      y,
      cropWidth,
      cropHeight,
      0,
      0,
      canvas.width,
      canvas.height,
    );

    nextCards.push({
      cropIndex: index,
      imageUrl: canvas.toDataURL("image/jpeg", 0.9),
      pair: oldCards.get(index)?.pair || "",
      pokemon: oldCards.get(index)?.pokemon || "",
      copies: oldCards.get(index)?.copies || "",
      ex6: oldCards.get(index)?.ex6 || "",
      ex6Source: oldCards.get(index)?.ex6Source || "",
      subrole: oldCards.get(index)?.subrole || "",
      subroleSource: oldCards.get(index)?.subroleSource || "",
      exRoleAvailable: oldCards.get(index)?.exRoleAvailable,
      hasMaxLabel: oldCards.get(index)?.hasMaxLabel,
      matchManuallyEdited: oldCards.get(index)?.matchManuallyEdited || false,
      selectedMatchPath: oldCards.get(index)?.selectedMatchPath || "",
      selectedMatchDistance: oldCards.get(index)?.selectedMatchDistance,
      selectedMatchSource: oldCards.get(index)?.selectedMatchSource || "",
      analysisNote: oldCards.get(index)?.analysisNote || "",
      suggestions: oldCards.get(index)?.suggestions || [],
      detectedMoveTypes: oldCards.get(index)?.detectedMoveTypes || [],
      moveTypeDetectionAvailable: oldCards.get(index)?.moveTypeDetectionAvailable || false,
      detectedPairRole: oldCards.get(index)?.detectedPairRole || "",
      pairRoleDetectionAvailable: oldCards.get(index)?.pairRoleDetectionAvailable || false,
    });
  }

  lastCropSignature = signature;
  cards = nextCards;
}

function createField(card, labelText, key, type, options) {
  const label = document.createElement("label");
  label.append(document.createTextNode(labelText));

  let control;
  if (type === "select") {
    control = document.createElement("select");
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Seleccionar…";
    control.append(placeholder);

    options.forEach(([value, text]) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = text;
      control.append(option);
    });
  } else {
    control = document.createElement("input");
    control.type = type;
    if (type === "number") {
      control.min = "1";
      control.max = key === "copies" ? "10" : "200";
      control.step = "1";
      control.inputMode = "numeric";
    }
  }

  control.value = card[key] || "";
  control.addEventListener("input", () => {
    card[key] = control.value;
    if (key === "ex6") card.ex6Source = "manual";
    if (key === "subrole") card.subroleSource = "manual";
    if (key === "pair" || key === "pokemon") {
      card.matchManuallyEdited = true;
      card.selectedMatchPath = "";
      card.selectedMatchDistance = undefined;
      card.selectedMatchSource = "";
      const article = control.closest(".unit-card");
      article?.querySelectorAll(".candidate-selected").forEach((button) => {
        button.classList.remove("candidate-selected");
        button.setAttribute("aria-pressed", "false");
      });
      article?.querySelector(".auto-match-note")?.remove();
      card.exRoleAvailable = undefined;
    }
    updateSummary();
    actualizarEstadoBotonSync();
  });
  label.append(control);
  return label;
}

function renderCards() {
  const fragment = document.createDocumentFragment();

  cards.forEach((card) => {
    const article = document.createElement("article");
    article.className = "unit-card";
    article.dataset.index = String(card.cropIndex);

    const preview = document.createElement("div");
    preview.className = "unit-preview";
    const image = document.createElement("img");
    image.src = card.imageUrl;
    image.alt = `Recorte de la tarjeta ${card.cropIndex + 1}`;
    image.loading = "lazy";
    const number = document.createElement("span");
    number.className = "unit-number";
    number.textContent = `Tarjeta ${card.cropIndex + 1}`;
    preview.append(image, number);

    const fields = document.createElement("div");
    fields.className = "card-fields";
    const pairField = createField(card, "Personaje / compi *", "pair", "text");
    pairField.classList.add("pair-field");
    fields.append(
      pairField,
      createField(card, "Pokémon", "pokemon", "text"),
      createField(card, "Copias (1–10) *", "copies", "number"),
      createField(card, "6EX *", "ex6", "select", [
        ["sí", "Sí"],
        ["no", "No"],
      ]),
      createField(card, "Subrol *", "subrole", "select", [
        ["sin_desbloquear", "Sin desbloquear"],
        ["max", "MAX"],
      ]),
    );

    const cardActions = document.createElement("div");
    cardActions.className = "card-actions";
    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "delete-card-button";
    deleteButton.textContent = "Eliminar compi";
    deleteButton.setAttribute("aria-label", `Eliminar tarjeta ${card.cropIndex + 1}`);
    deleteButton.addEventListener("click", () => {
      if (excludedCropSignature !== lastCropSignature) return;
      excludedCropIndexes.add(card.cropIndex);
      cards = cards.filter((entry) => entry.cropIndex !== card.cropIndex);
      cropStatus.textContent = `${cards.length} tarjetas incluidas`;
      renderCards();
      updateSummary();
    });
    cardActions.append(deleteButton);

    if (card.analysisNote) {
      const analysisNote = document.createElement("p");
      analysisNote.className = "auto-analysis-note";
      analysisNote.textContent = card.analysisNote;
      fields.append(analysisNote);
    }
    if (card.suggestions?.length) {
      const moveTypeNote = document.createElement("p");
      moveTypeNote.className = "auto-analysis-note";
      moveTypeNote.textContent = card.moveTypeDetectionAvailable
        ? card.detectedMoveTypes.length
          ? `Tipos detectados a la derecha: ${card.detectedMoveTypes.map(moveTypeLabel).join(", ")}. ` +
            "Las coincidencias del catálogo ayudan a subir candidatos; los tipos no coincidentes no se penalizan porque sus etiquetas pueden estar incompletas."
          : "No se detectaron tipos con suficiente parecido en los iconos de la derecha."
        : "No se pudieron cargar las plantillas de tipo; no se usaron para ordenar.";
      fields.append(moveTypeNote);
    }
    if (card.suggestions?.length) {
      const roleNote = document.createElement("p");
      roleNote.className = "auto-analysis-note";
      roleNote.textContent = card.pairRoleDetectionAvailable
        ? card.detectedPairRole
          ? `Rol detectado bajo las estrellas: ${pairRoleLabel(card.detectedPairRole)}. ` +
            "Se usa para priorizar compis del mismo rol."
          : "No se detectó un rol con confianza suficiente bajo las estrellas."
        : "No se pudieron cargar las plantillas de rol; no se usaron para ordenar.";
      fields.append(roleNote);
    }
    if (card.selectedMatchPath) {
      const matchNote = document.createElement("p");
      matchNote.className = "auto-match-note";
      matchNote.textContent = card.selectedMatchSource === "manual"
        ? `Propuesta seleccionada (d=${card.selectedMatchDistance.toFixed(1)}).`
        : `Candidato aplicado automáticamente (d=${card.selectedMatchDistance.toFixed(1)}). ` +
          "Revisa el nombre y el Pokémon; puedes elegir otra propuesta.";
      fields.append(matchNote);
    }

    article.append(preview, fields, cardActions);
    if (card.suggestions?.length) {
      const suggestions = document.createElement("div");
      suggestions.className = "catalog-suggestions";
      const note = document.createElement("span");
      note.className = "suggestion-label";
      note.textContent =
        "Candidato automático arriba · elige otra propuesta si la imagen coincide mejor:";
      suggestions.append(note);

      card.suggestions.slice(0, 5).forEach((suggestion) => {
        const button = document.createElement("button");
        button.className = "candidate-button";
        if (suggestion.path === card.selectedMatchPath) {
          button.classList.add("candidate-selected");
          button.setAttribute("aria-pressed", "true");
        } else {
          button.setAttribute("aria-pressed", "false");
        }
        button.type = "button";
        button.title = `Usar ${suggestion.name} + ${suggestion.pokemon}`;
        const image = document.createElement("img");
        image.src = suggestion.imageUrl;
        image.alt = "";
        image.loading = "lazy";
        const copy = document.createElement("span");
        copy.className = "candidate-copy";
        const name = document.createElement("strong");
        name.textContent = [
          suggestion.name,
          suggestion.trainerAlt,
          suggestion.variant,
        ].filter(Boolean).join(" · ");
        const detail = document.createElement("span");
        const typeDetail = suggestion.detectedMoveTypes.length
          ? ` · tipos ${suggestion.sharedTypes}/${suggestion.detectedMoveTypes.length}` +
            (suggestion.typeBonus ? ` (+${suggestion.typeBonus.toFixed(1)})` : "")
          : "";
        const roleDetail = suggestion.detectedPairRole && suggestion.role
          ? ` · rol ${pairRoleLabel(suggestion.role)}` +
            (suggestion.roleScore > 0
              ? ` (+${suggestion.roleScore.toFixed(1)})`
              : ` (${pairRoleLabel(suggestion.detectedPairRole)} no coincide)`)
          : "";
        detail.textContent =
          `${suggestion.pokemon} · d=${suggestion.distance.toFixed(1)}${typeDetail}${roleDetail}`;
        button.title =
          `${[suggestion.name, suggestion.trainerAlt, suggestion.variant]
            .filter(Boolean)
            .join(" · ")} · ` +
          `${suggestion.pokemon || "Pokémon sin asociar"} · ` +
          `distancia RGB ${suggestion.distance.toFixed(2)} · ` +
          `rol catálogo: ${pairRoleLabel(suggestion.role) || "sin datos"} · ` +
          `tipos detectados: ${suggestion.detectedMoveTypes.map(moveTypeLabel).join(", ") || "ninguno"} · ` +
          `tipos del catálogo: ${suggestion.moveTypes.map(moveTypeLabel).join(", ") || "sin datos"}`;
        copy.append(name, detail);
        button.append(image, copy);
        button.addEventListener("click", () => {
          applyMatch(card, suggestion, true);
          renderCards();
          updateSummary();
        });
        suggestions.append(button);
      });

      fields.append(suggestions);
    }
    fragment.append(article);
  });

  cardsList.replaceChildren(fragment);
}

function applyMatch(card, suggestion, manual) {
  if (manual || !card.matchManuallyEdited) {
    card.pair = [suggestion.name, suggestion.trainerAlt].filter(Boolean).join(" - ");
    card.pokemon = suggestion.pokemon;
    card.selectedMatchPath = suggestion.path;
    card.selectedMatchDistance = suggestion.distance;
    card.selectedMatchSource = manual ? "manual" : "auto";
  }
  card.exRoleAvailable = suggestion.hasEXRole;
  if (manual) card.matchManuallyEdited = true;

  if (card.subroleSource === "auto") {
    card.subrole = "";
    card.subroleSource = "";
  }
  if (!card.subrole && card.hasMaxLabel !== undefined) {
    applyMaxDetection(card, card.hasMaxLabel);
  }
}

function applyMaxDetection(card, hasMax) {
  card.hasMaxLabel = hasMax;
  if (card.subroleSource === "manual") return false;
  card.subrole = hasMax ? "max" : "sin_desbloquear";
  card.subroleSource = "auto";
  return true;
}

function isCardComplete(card) {
  const copies = Number(card.copies);
  return (
    card.pair.trim() !== "" &&
    Number.isInteger(copies) &&
    copies >= 1 &&
    copies <= 10 &&
    (card.ex6 === "sí" || card.ex6 === "no") &&
    card.subrole !== ""
  );
}

function updateSummary() {
  const complete = cards.filter(isCardComplete).length;
  reviewCount.textContent = cards.length
    ? `${complete} de ${cards.length} tarjetas revisadas · los campos con * son obligatorios`
    : loadedImage
      ? "No quedan compis en la lista. Ajusta el recorte o carga otra captura."
      : "Carga una captura para empezar.";
  exportButton.disabled = cards.length === 0 || complete !== cards.length;
  recognizeButton.disabled = cards.length === 0;
  suggestionsExportButton.disabled = !cards.some((card) => card.suggestions?.length);
  actualizarEstadoBotonSync();
}

const TRACKER_DATA_URL =
  "https://raw.githubusercontent.com/pomasters/SyncPairsTracker/main/js/syncpairs.json";
const TRACKER_SITE_ROOT = "https://pomasters.github.io/SyncPairsTracker/";
const TESSERACT_SCRIPT =
  "https://cdn.jsdelivr.net/npm/tesseract.js@6.0.1/dist/tesseract.min.js";

async function fetchTrackerData() {
  const response = await fetch(TRACKER_DATA_URL);
  if (!response.ok) {
    throw new Error(`Sync Pairs Tracker respondió ${response.status} al cargar el catálogo.`);
  }

  return response.json();
}

function trackerImageUrl(path) {
  if (typeof path !== "string" || !/^icons\/[^/]+\.png$/.test(path)) {
    throw new Error("El catálogo contiene una ruta de icono no válida.");
  }
  return new URL(path, TRACKER_SITE_ROOT).href;
}

function imageDescriptor(bitmap, crop, size = 48) {
  const canvas = document.createElement("canvas");
  canvas.width = size;
  canvas.height = size;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) throw new Error("Este navegador no permite comparar imágenes.");

  const x = bitmap.width * crop.x;
  const y = bitmap.height * crop.y;
  const width = bitmap.width * crop.width;
  const height = bitmap.height * crop.height;
  context.drawImage(bitmap, x, y, width, height, 0, 0, size, size);
  const pixels = context.getImageData(0, 0, size, size).data;
  const rgb = [];
  const alpha = [];

  for (let index = 0; index < pixels.length; index += 4) {
    rgb.push(pixels[index], pixels[index + 1], pixels[index + 2]);
    alpha.push(pixels[index + 3] / 255);
  }

  return { rgb, alpha };
}

function imageDistance(left, right) {
  let total = 0;
  let weight = 0;

  for (let index = 0; index < left.rgb.length; index += 3) {
    const alpha = right.alpha[index / 3];
    if (alpha < 0.15) continue;

    const cellDistance =
      (Math.abs(left.rgb[index] - right.rgb[index]) +
        Math.abs(left.rgb[index + 1] - right.rgb[index + 1]) +
        Math.abs(left.rgb[index + 2] - right.rgb[index + 2])) /
      3;
    total += cellDistance * alpha;
    weight += alpha;
  }

  if (!weight) throw new Error("El catálogo contiene un retrato vacío.");
  return (total / weight / 255) * 100;
}

async function getMoveTypeTemplates() {
  if (moveTypeTemplatesPromise) return moveTypeTemplatesPromise;

  moveTypeTemplatesPromise = Promise.all(
    MOVE_TYPES.map(async (type) => {
      const imageUrl = type === "Stellar"
        ? STELLAR_TYPE_ICON_DATA_URL
        : `${TRACKER_SITE_ROOT}images/type_${type.toLowerCase()}.png`;
      try {
        const response = await fetch(imageUrl);
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const bitmap = await createImageBitmap(await response.blob());
        const descriptor = imageDescriptor(
          bitmap,
          { x: 0, y: 0, width: 1, height: 1 },
          MOVE_TYPE_SAMPLE_SIZE,
        );
        bitmap.close();
        const samples = [];
        for (let index = 0; index < descriptor.alpha.length; index += 1) {
          const weight = descriptor.alpha[index];
          if (weight < 0.15) continue;
          const pixel = index * 3;
          samples.push({
            pixel,
            weight,
            red: descriptor.rgb[pixel],
            green: descriptor.rgb[pixel + 1],
            blue: descriptor.rgb[pixel + 2],
          });
        }
        return { type, samples };
      } catch (error) {
        return { type, error: error.message };
      }
    }),
  ).then((results) => {
    const templates = results.filter((result) => result.samples?.length);
    const failed = results.filter((result) => result.error);
    if (!templates.length) {
      throw new Error(
        `No se pudieron cargar los iconos de tipo del catálogo: ${failed
          .slice(0, 3)
          .map((result) => `${result.type} (${result.error})`)
          .join(", ")}`,
      );
    }
    return { templates, failed };
  }).catch((error) => {
    moveTypeTemplatesPromise = null;
    throw error;
  });

  return moveTypeTemplatesPromise;
}

async function getPairRoleTemplates() {
  if (pairRoleTemplatesPromise) return pairRoleTemplatesPromise;

  pairRoleTemplatesPromise = Promise.all(
    PAIR_ROLES.map(async ({ role, icon }) => {
      const imageUrl = `${TRACKER_SITE_ROOT}images/role_${icon}.png`;
      try {
        const response = await fetch(imageUrl);
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const bitmap = await createImageBitmap(await response.blob());
        const descriptor = imageDescriptor(
          bitmap,
          { x: 0, y: 0, width: 1, height: 1 },
          ROLE_SAMPLE_SIZE,
        );
        bitmap.close();
        const samples = [];
        for (let index = 0; index < descriptor.alpha.length; index += 1) {
          const weight = descriptor.alpha[index];
          if (weight < 0.15) continue;
          const pixel = index * 3;
          samples.push({
            pixel,
            weight,
            red: descriptor.rgb[pixel],
            green: descriptor.rgb[pixel + 1],
            blue: descriptor.rgb[pixel + 2],
          });
        }
        return { role, samples };
      } catch (error) {
        return { role, error: error.message };
      }
    }),
  ).then((results) => {
    const templates = results.filter((result) => result.samples?.length);
    const failed = results.filter((result) => result.error);
    if (!templates.length) {
      throw new Error(
        `No se pudieron cargar los iconos de rol del catálogo: ${failed
          .map((result) => `${result.role} (${result.error})`)
          .join(", ")}`,
      );
    }
    return { templates, failed };
  }).catch((error) => {
    pairRoleTemplatesPromise = null;
    throw error;
  });

  return pairRoleTemplatesPromise;
}

async function getCopyBadgeTemplates() {
  if (copyBadgeTemplatesPromise) return copyBadgeTemplatesPromise;

  copyBadgeTemplatesPromise = Promise.all(
    Array.from({ length: 5 }, (_, index) => index + 1).flatMap((level) => [
      { level, excelsior: false, filename: `${level}.png` },
      { level, excelsior: true, filename: `${level}_2.png` },
    ]).map(async ({ level, excelsior, filename }) => {
      const imageUrl = `${TRACKER_SITE_ROOT}images/${filename}`;
      try {
        const response = await fetch(imageUrl);
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const bitmap = await createImageBitmap(await response.blob());
        const descriptor = imageDescriptor(
          bitmap,
          { x: 0, y: 0, width: 1, height: 1 },
          COPY_BADGE_SAMPLE_SIZE,
        );
        bitmap.close();
        const samples = [];
        for (let index = 0; index < descriptor.alpha.length; index += 1) {
          const weight = descriptor.alpha[index];
          if (weight < 0.15) continue;
          const pixel = index * 3;
          samples.push({
            pixel,
            weight,
            red: descriptor.rgb[pixel],
            green: descriptor.rgb[pixel + 1],
            blue: descriptor.rgb[pixel + 2],
          });
        }
        const shapeSamples = samples.filter((sample) => {
          const sampleIndex = sample.pixel / 3;
          const x = sampleIndex % COPY_BADGE_SAMPLE_SIZE;
          const y = Math.floor(sampleIndex / COPY_BADGE_SAMPLE_SIZE);
          return (
            x < COPY_BADGE_SHAPE_BORDER ||
            y < COPY_BADGE_SHAPE_BORDER ||
            x >= COPY_BADGE_SAMPLE_SIZE - COPY_BADGE_SHAPE_BORDER ||
            y >= COPY_BADGE_SAMPLE_SIZE - COPY_BADGE_SHAPE_BORDER
          );
        });
        const digitSamples = samples.filter((sample) => {
          const sampleIndex = sample.pixel / 3;
          const x = sampleIndex % COPY_BADGE_SAMPLE_SIZE;
          const y = Math.floor(sampleIndex / COPY_BADGE_SAMPLE_SIZE);
          return x >= 7 && x < 17 && y >= 6 && y < 19;
        });
        return { level, excelsior, samples, shapeSamples, digitSamples };
      } catch (error) {
        return { level, excelsior, error: error.message };
      }
    }),
  ).then((results) => {
    const templates = results.filter((result) => result.samples?.length);
    const failed = results.filter((result) => result.error);
    if (!templates.length) {
      throw new Error(
        `No se pudieron cargar las plantillas de copias: ${failed
          .slice(0, 4)
          .map((result) => `${result.level}${result.excelsior ? " (Excelsior)" : ""} (${result.error})`)
          .join(", ")}`,
      );
    }
    return { templates, failed };
  }).catch((error) => {
    copyBadgeTemplatesPromise = null;
    throw error;
  });

  return copyBadgeTemplatesPromise;
}

function sampleImageRegion(pixels, imageWidth, x, y, size, sampleSize) {
  const samples = new Uint8ClampedArray(sampleSize * sampleSize * 3);
  for (let row = 0; row < sampleSize; row += 1) {
    for (let column = 0; column < sampleSize; column += 1) {
      const sourceX = Math.min(
        imageWidth - 1,
        Math.floor(x + ((column + 0.5) * size) / sampleSize),
      );
      const sourceY = Math.floor(y + ((row + 0.5) * size) / sampleSize);
      const sourcePixel = (sourceY * imageWidth + sourceX) * 4;
      const samplePixel = (row * sampleSize + column) * 3;
      samples[samplePixel] = pixels[sourcePixel];
      samples[samplePixel + 1] = pixels[sourcePixel + 1];
      samples[samplePixel + 2] = pixels[sourcePixel + 2];
    }
  }
  return samples;
}

function moveTypeDistance(samples, template) {
  let total = 0;
  let weight = 0;
  template.samples.forEach((sample) => {
    total +=
      (Math.abs(samples[sample.pixel] - sample.red) +
        Math.abs(samples[sample.pixel + 1] - sample.green) +
        Math.abs(samples[sample.pixel + 2] - sample.blue)) *
      sample.weight;
    weight += sample.weight;
  });
  return weight ? (total / (weight * 3 * 255)) * 100 : 100;
}

function detectMoveTypes(image, templates) {
  const sampleSize = MOVE_TYPE_SAMPLE_SIZE;
  const canvas = document.createElement("canvas");
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) throw new Error("Este navegador no permite leer los iconos de tipo.");
  context.drawImage(image, 0, 0);
  const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
  const badgeSlots = [
    { x: 0.81, y: 0.26 },
    { x: 0.81, y: 0.43 },
  ];
  const sizeFractions = [0.12, 0.14, 0.16, 0.18];

  return badgeSlots.flatMap((slot) => {
    let best = null;
    sizeFractions.forEach((fraction) => {
      const size = Math.max(8, Math.round(canvas.width * fraction));
      for (let centerY = slot.y - 0.035; centerY <= slot.y + 0.035; centerY += 0.01) {
        for (let centerX = slot.x - 0.045; centerX <= slot.x + 0.045; centerX += 0.01) {
          const x = Math.round(canvas.width * centerX - size / 2);
          const y = Math.round(canvas.height * centerY - size / 2);
          if (
            x < 0 ||
            y < 0 ||
            x + size > canvas.width ||
            y + size > canvas.height
          ) {
            continue;
          }
          const regionSamples = sampleImageRegion(
            pixels,
            canvas.width,
            x,
            y,
            size,
            sampleSize,
          );
          templates.forEach((template) => {
            const distance = moveTypeDistance(regionSamples, template);
            if (!best || distance < best.distance) {
              best = {
                type: template.type,
                distance,
                x: x + size / 2,
                y: y + size / 2,
                size,
              };
            }
          });
        }
      }
    });

    return best && best.distance <= 12.5 ? [best] : [];
  });
}

function detectPairRole(image, templates) {
  const sampleSize = ROLE_SAMPLE_SIZE;
  const canvas = document.createElement("canvas");
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) throw new Error("Este navegador no permite leer los iconos de rol.");
  context.drawImage(image, 0, 0);
  const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
  const slot = { x: 0.18, y: 0.35 };
  const sizeFractions = [0.1, 0.12, 0.14, 0.16, 0.18];
  const matches = templates.map((template) => {
    let bestDistance = Infinity;
    sizeFractions.forEach((fraction) => {
      const size = Math.max(8, Math.round(canvas.width * fraction));
      for (let centerY = slot.y - 0.09; centerY <= slot.y + 0.09; centerY += 0.01) {
        for (let centerX = slot.x - 0.15; centerX <= slot.x + 0.15; centerX += 0.01) {
          const x = Math.round(canvas.width * centerX - size / 2);
          const y = Math.round(canvas.height * centerY - size / 2);
          if (
            x < 0 ||
            y < 0 ||
            x + size > canvas.width ||
            y + size > canvas.height
          ) {
            continue;
          }
          const samples = sampleImageRegion(
            pixels,
            canvas.width,
            x,
            y,
            size,
            sampleSize,
          );
          bestDistance = Math.min(bestDistance, moveTypeDistance(samples, template));
        }
      }
    });
    return { role: template.role, distance: bestDistance };
  }).sort((left, right) => left.distance - right.distance);

  const [best, runnerUp] = matches;
  const margin = runnerUp ? runnerUp.distance - best.distance : Infinity;
  return best && best.distance <= 22 && margin >= 1.2
    ? { role: best.role, distance: best.distance, margin }
    : null;
}

async function getCatalog() {
  if (catalogPromise) return catalogPromise;

  catalogStatus.textContent = "Descargando datos de Sync Pairs Tracker…";
  catalogPromise = (async () => {
    const data = await fetchTrackerData();
    if (!Array.isArray(data?.SYNCPAIRS)) {
      throw new Error("El formato de Sync Pairs Tracker no contiene SYNCPAIRS.");
    }

    const entries = data.SYNCPAIRS.flatMap((pair) => {
      if (
        typeof pair.trainerName !== "string" ||
        typeof pair.pokemonName !== "string" ||
        !Array.isArray(pair.images)
      ) {
        return [];
      }

      return pair.images
        .filter((path) => typeof path === "string" &&
          (path.endsWith("_5.png") || path.endsWith("_EX.png")))
        .map((path) => ({
          name: pair.trainerName,
          trainerAlt: pair.trainerAlt || "",
          pokemon: pair.pokemonName,
          role: normalizePairRole(pair.syncPairRole),
          moveTypes: Array.isArray(pair.tags)
            ? [...new Set(
              pair.tags
                .filter((tag) => typeof tag === "string" && tag.startsWith("MoveType"))
                .map((tag) => tag.slice("MoveType".length)),
            )]
            : [],
          variant: path.endsWith("_EX.png") ? "EX" : "5★",
          hasEXRole: Boolean(pair.syncPairRoleEX),
          path,
          imageUrl: trackerImageUrl(path),
        }));
    });
    if (!entries.length) {
      throw new Error("No se encontraron iconos 5★ o EX en el catálogo.");
    }

    let nextIndex = 0;
    let processedCount = 0;
    let loadedCount = 0;
    const failed = [];
    const worker = async () => {
      while (true) {
        const index = nextIndex;
        nextIndex += 1;
        if (index >= entries.length) return;

        const entry = entries[index];
        try {
          const response = await fetch(entry.imageUrl);
          if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
          }
          const imageBlob = await response.blob();
          const bitmap = await createImageBitmap(imageBlob);
          entry.descriptor = imageDescriptor(bitmap, { x: 0, y: 0, width: 1, height: 1 });
          bitmap.close();
          entry.imageObjectUrl = URL.createObjectURL(imageBlob);
          catalogImageUrls.add(entry.imageObjectUrl);
          loadedCount += 1;
        } catch (error) {
          failed.push(`${entry.path} (${error.message})`);
        }

        processedCount += 1;
        if (processedCount % 25 === 0 || processedCount === entries.length) {
          catalogStatus.textContent =
            `Cargando iconos del catálogo: ${processedCount}/${entries.length}…`;
        }
      }
    };

    await Promise.all(Array.from({ length: 12 }, worker));
    catalogRecords = entries.filter((entry) => entry.descriptor && entry.imageObjectUrl);
    if (!catalogRecords.length) {
      throw new Error("No se pudo cargar ningún icono de Sync Pairs Tracker.");
    }

    catalogStatus.textContent =
      `${catalogRecords.length} iconos 5★/EX cargados desde Sync Pairs Tracker.` +
      (failed.length
        ? ` Fallaron ${failed.length}; por ejemplo: ${failed.slice(0, 2).join(", ")}.`
        : "");
    return catalogRecords;
  })().catch((error) => {
    catalogImageUrls.forEach((url) => URL.revokeObjectURL(url));
    catalogImageUrls.clear();
    catalogRecords = [];
    catalogPromise = null;
    throw error;
  });

  return catalogPromise;
}

async function analyzeAndRecognizeCards() {
  if (!cards.length) return;

  recognizeButton.disabled = true;
  suggestionsExportButton.disabled = true;
  showError("");
  catalogStatus.textContent = "";

  try {
    await analyzeCardDetails();
    await matchCardsWithCatalog();
  } finally {
    updateSummary();
  }
}

async function matchCardsWithCatalog() {
  if (!cards.length) return;

  suggestionsExportButton.disabled = true;

  try {
    const catalog = await getCatalog();
    let typeTemplates = [];
    let typeTemplateWarning = "";
    try {
      const loadedTypes = await getMoveTypeTemplates();
      typeTemplates = loadedTypes.templates;
      if (loadedTypes.failed.length) {
        typeTemplateWarning =
          ` No se cargaron ${loadedTypes.failed.length} iconos de tipo; ` +
          "la detección puede estar incompleta.";
      }
    } catch (error) {
      typeTemplateWarning = ` No se pudieron cargar los iconos de tipo: ${error.message}`;
    }
    let pairRoleTemplates = [];
    let pairRoleTemplateWarning = "";
    try {
      const loadedRoles = await getPairRoleTemplates();
      pairRoleTemplates = loadedRoles.templates;
      if (loadedRoles.failed.length) {
        pairRoleTemplateWarning =
          ` No se cargaron ${loadedRoles.failed.length} iconos de rol; ` +
          "la detección puede estar incompleta.";
      }
    } catch (error) {
      pairRoleTemplateWarning = ` No se pudieron cargar los iconos de rol: ${error.message}`;
    }
    const matches = cards.map((card) => {
      const target = new Image();
      target.src = card.imageUrl;
      return target.decode().then(() => {
        const descriptor = imageDescriptor(target, { x: 0, y: 0, width: 1, height: 1 });
        const moveTypeDetections = typeTemplates.length
          ? detectMoveTypes(target, typeTemplates)
          : [];
        const detectedMoveTypes = moveTypeDetections.map((detection) => detection.type);
        const pairRoleDetection = pairRoleTemplates.length
          ? detectPairRole(target, pairRoleTemplates)
          : null;
        const detectedPairRole = pairRoleDetection?.role || "";
        const expectedVariant = card.ex6 === "sí" ? "EX" : card.ex6 === "no" ? "5★" : "";
        const variantCatalog = expectedVariant
          ? catalog.filter((entry) => entry.variant === expectedVariant)
          : [];
        const matchingCatalog = variantCatalog.length ? variantCatalog : catalog;

        const variants = matchingCatalog
          .map((entry) => ({
            ...entry,
            distance: imageDistance(descriptor, entry.descriptor),
          }))
          .map((entry) => {
            const sharedTypes = detectedMoveTypes.filter((type) =>
              entry.moveTypes.includes(type),
            ).length;
            const typeBonus = detectedMoveTypes.length && entry.moveTypes.length
              ? (sharedTypes / detectedMoveTypes.length) * 6
              : 0;
            const roleMatch = Boolean(
              detectedPairRole && entry.role && detectedPairRole === entry.role,
            );
            const roleScore = detectedPairRole && entry.role
              ? roleMatch ? 8 : -4
              : 0;
            return {
              ...entry,
              detectedMoveTypes,
              detectedPairRole,
              sharedTypes,
              typeBonus,
              roleMatch,
              roleScore,
              rankDistance: entry.distance - typeBonus - roleScore,
            };
          })
          .sort((left, right) => left.rankDistance - right.rankDistance);
        const bestByPair = new Map();
        variants.forEach((entry) => {
          const pairKey = `${entry.name}\u0000${entry.trainerAlt}\u0000${entry.pokemon}`;
          if (!bestByPair.has(pairKey)) bestByPair.set(pairKey, entry);
        });

        return [...bestByPair.values()]
          .sort((left, right) => left.rankDistance - right.rankDistance)
          .slice(0, 5)
          .map((entry) => ({
            name: entry.name,
            trainerAlt: entry.trainerAlt,
            pokemon: entry.pokemon,
            variant: entry.variant,
            hasEXRole: entry.hasEXRole,
            path: entry.path,
            imageUrl: entry.imageObjectUrl,
            distance: entry.distance,
            rankDistance: entry.rankDistance,
            moveTypes: entry.moveTypes,
            detectedMoveTypes: entry.detectedMoveTypes,
            sharedTypes: entry.sharedTypes,
            typeBonus: entry.typeBonus,
            role: entry.role,
            detectedPairRole: entry.detectedPairRole,
            roleMatch: entry.roleMatch,
            roleScore: entry.roleScore,
          }));
      });
    });

    const suggestions = await Promise.all(matches);
    let variantFallbacks = 0;
    let autoFilledCards = 0;
    let manuallyEditedCards = 0;
    cards.forEach((card, index) => {
      card.suggestions = suggestions[index];
      card.detectedMoveTypes = suggestions[index][0]?.detectedMoveTypes || [];
      card.moveTypeDetectionAvailable = typeTemplates.length > 0;
      card.detectedPairRole = suggestions[index][0]?.detectedPairRole || "";
      card.pairRoleDetectionAvailable = pairRoleTemplates.length > 0;
      const expectedVariant = card.ex6 === "sí" ? "EX" : card.ex6 === "no" ? "5★" : "";
      if (expectedVariant && !catalog.some((entry) => entry.variant === expectedVariant)) {
        variantFallbacks += 1;
      }
      if (suggestions[index].length) {
        if (card.matchManuallyEdited) {
          manuallyEditedCards += 1;
        } else {
          applyMatch(card, suggestions[index][0], false);
          autoFilledCards += 1;
        }
      }
    });
    renderCards();
    updateSummary();
    catalogStatus.textContent =
      `Se compararon los recortes con iconos del catálogo de 48 × 48 píxeles. ` +
      "La estrella 6EX de cada tarjeta limitó individualmente la búsqueda a iconos EX o 5★. " +
      `${cards.filter((card) => card.detectedMoveTypes.length).length} tarjetas con tipos de movimiento detectados en el lateral derecho; ` +
      "cada tipo coincidente en el catálogo da una bonificación al orden (máximo 6 puntos); los tipos no coincidentes no se penalizan porque las etiquetas pueden estar incompletas. " +
      `${cards.filter((card) => card.detectedPairRole).length} roles detectados bajo las estrellas; una coincidencia de rol suma 8 puntos y una discrepancia resta 4 al ranking. ` +
      `Se aplicó automáticamente la propuesta mejor puntuada por RGB, tipos y rol a ${autoFilledCards} tarjetas` +
      (manuallyEditedCards ? `; se respetaron las correcciones manuales en ${manuallyEditedCards}` : "") +
      " (d RGB=0 sería idéntico; ninguna distancia es una probabilidad)." +
      typeTemplateWarning +
      pairRoleTemplateWarning +
      (variantFallbacks
        ? ` En ${variantFallbacks} tarjetas no estaba disponible la variante esperada y se usó el catálogo completo.`
        : "");
    analysisStatus.textContent +=
      " Revisa las miniaturas y corrige los campos si hace falta.";
  } catch (error) {
    showError(`No se pudo comparar con Sync Pairs Tracker: ${error.message}`);
    catalogStatus.textContent = "Revisa la conexión con GitHub y vuelve a intentarlo.";
  }
}

function detectEXMarker(image) {
  const size = 48;
  const badge = { x: 0.32, y: 0.16, width: 0.14, height: 0.14 };
  const canvas = document.createElement("canvas");
  canvas.width = size;
  canvas.height = size;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) throw new Error("No se pudo analizar la estrella EX.");

  const badgeX = Math.round(size * badge.x);
  const badgeY = Math.round(size * badge.y);
  const badgeWidth = Math.round(size * badge.width);
  const badgeHeight = Math.round(size * badge.height);
  context.drawImage(
    image,
    image.naturalWidth * badge.x,
    image.naturalHeight * badge.y,
    image.naturalWidth * badge.width,
    image.naturalHeight * badge.height,
    badgeX,
    badgeY,
    badgeWidth,
    badgeHeight,
  );

  const pixels = context.getImageData(badgeX, badgeY, badgeWidth, badgeHeight).data;
  let coolPixels = 0;
  const coolHueBins = Array(5).fill(0);
  let analyzedPixels = 0;
  for (let index = 0; index < pixels.length; index += 4) {
    const red = pixels[index] / 255;
    const green = pixels[index + 1] / 255;
    const blue = pixels[index + 2] / 255;
    const max = Math.max(red, green, blue);
    const min = Math.min(red, green, blue);
    const saturation = max ? (max - min) / max : 0;
    if (max < 0.24 || saturation < 0.3) continue;

    analyzedPixels += 1;
    const delta = max - min;
    let hue;
    if (max === red) hue = 60 * (((green - blue) / delta) % 6);
    else if (max === green) hue = 60 * ((blue - red) / delta + 2);
    else hue = 60 * ((red - green) / delta + 4);
    hue = (hue + 360) % 360;

    if (hue >= 170 && hue < 320) {
      coolPixels += 1;
      coolHueBins[Math.min(4, Math.floor((hue - 170) / 30))] += 1;
    }
  }

  if (analyzedPixels < 5) {
    return { value: "no", reason: "estrella dorada o ausente (no 6EX)" };
  }
  const coolRatio = coolPixels / analyzedPixels;
  const distinctRainbowBands = coolHueBins.filter(
    (count) => count / analyzedPixels >= 0.03,
  ).length;

  if (coolRatio >= 0.18 && distinctRainbowBands >= 2) {
    return { value: "sí", reason: "estrella EX arcoíris" };
  }
  return { value: "no", reason: "estrella dorada o ausente (no 6EX)" };
}

function loadTesseract() {
  if (window.Tesseract) return Promise.resolve(window.Tesseract);
  if (!ocrWorkerPromise) {
    ocrWorkerPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = TESSERACT_SCRIPT;
      script.onload = () => {
        if (window.Tesseract) {
          resolve(window.Tesseract);
        } else {
          reject(new Error("Tesseract.js se cargó, pero no está disponible."));
        }
      };
      script.onerror = () => reject(new Error("No se pudo cargar Tesseract.js desde jsDelivr."));
      document.head.append(script);
    }).catch((error) => {
      ocrWorkerPromise = null;
      throw error;
    });
  }
  return ocrWorkerPromise;
}

async function getOcrWorker() {
  if (!ocrWorkerPromise || !window.Tesseract) {
    ocrWorkerPromise = loadTesseract();
  }
  const tesseract = await ocrWorkerPromise;
  if (!getOcrWorker.worker) {
    getOcrWorker.worker = await tesseract.createWorker("eng", 1);
  }
  return getOcrWorker.worker;
}

function makeOCRCrop(image, rectangle, scale) {
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(image.naturalWidth * rectangle.width * scale));
  canvas.height = Math.max(1, Math.round(image.naturalHeight * rectangle.height * scale));
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Este navegador no permite preparar la imagen para OCR.");
  context.imageSmoothingEnabled = false;
  context.drawImage(
    image,
    image.naturalWidth * rectangle.x,
    image.naturalHeight * rectangle.y,
    image.naturalWidth * rectangle.width,
    image.naturalHeight * rectangle.height,
    0,
    0,
    canvas.width,
    canvas.height,
  );
  return canvas;
}

async function detectExcelsiorBadge(image, templates, level, ocrReadingCount = 0) {
  const levelTemplates = templates.filter((template) => template.level === level);
  if (!levelTemplates.length) return null;

  const canvas = document.createElement("canvas");
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) throw new Error("Este navegador no permite leer la insignia de copias.");
  context.drawImage(image, 0, 0);
  const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
  const matches = levelTemplates.map((template) => {
    const shapeTemplate = { ...template, samples: template.shapeSamples };
    let colorDistance = Infinity;
    let shapeDistance = Infinity;
    let combinedDistance = Infinity;
    for (const sizeFraction of [0.16, 0.18, 0.2, 0.22, 0.24]) {
      const size = Math.max(8, Math.round(canvas.width * sizeFraction));
      for (let centerY = 0.76; centerY <= 0.88; centerY += 0.01) {
        for (let centerX = 0.11; centerX <= 0.19; centerX += 0.01) {
          const x = Math.round(canvas.width * centerX - size / 2);
          const y = Math.round(canvas.height * centerY - size / 2);
          if (
            x < 0 ||
            y < 0 ||
            x + size > canvas.width ||
            y + size > canvas.height
          ) {
            continue;
          }
          const samples = sampleImageRegion(
            pixels,
            canvas.width,
            x,
            y,
            size,
            COPY_BADGE_SAMPLE_SIZE,
          );
          const colorMatch = moveTypeDistance(samples, template);
          const shapeMatch = moveTypeDistance(samples, shapeTemplate);
          colorDistance = Math.min(colorDistance, colorMatch);
          shapeDistance = Math.min(shapeDistance, shapeMatch);
          combinedDistance = Math.min(
            combinedDistance,
            (colorMatch + 2 * shapeMatch) / 3,
          );
        }
      }
    }
    return { excelsior: template.excelsior, colorDistance, shapeDistance, combinedDistance };
  });
  const bestByType = [false, true]
    .map((excelsior) => ({
      excelsior,
      colorDistance: Math.min(...matches
        .filter((match) => match.excelsior === excelsior)
        .map((match) => match.colorDistance)),
      shapeDistance: Math.min(...matches
        .filter((match) => match.excelsior === excelsior)
        .map((match) => match.shapeDistance)),
      combinedDistance: Math.min(...matches
        .filter((match) => match.excelsior === excelsior)
        .map((match) => match.combinedDistance)),
    }));
  const normal = bestByType.find((match) => !match.excelsior);
  const excelsior = bestByType.find((match) => match.excelsior);
  const colorAdvantage = normal.colorDistance - excelsior.colorDistance;
  if (colorAdvantage >= 3 || (ocrReadingCount >= 3 && colorAdvantage >= 0.5)) {
    return true;
  }
  const ranked = bestByType.sort(
    (left, right) => left.combinedDistance - right.combinedDistance,
  );
  const [best, runnerUp] = ranked;
  return best && best.combinedDistance <= 32 &&
    runnerUp.combinedDistance - best.combinedDistance >= 1
    ? best.excelsior
    : null;
}

function detectCopyBadgeLevel(image, templates) {
  const canvas = document.createElement("canvas");
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) throw new Error("Este navegador no permite leer el número de copias.");
  context.drawImage(image, 0, 0);
  const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
  const bestByLevel = new Map();

  for (const template of templates) {
    const digitTemplate = { ...template, samples: template.digitSamples };
    let distance = Infinity;
    for (const sizeFraction of [0.18, 0.2, 0.22, 0.24, 0.26, 0.28, 0.3]) {
      const size = Math.max(8, Math.round(canvas.width * sizeFraction));
      for (let centerY = 0.76; centerY <= 0.9; centerY += 0.01) {
        for (let centerX = 0.08; centerX <= 0.22; centerX += 0.01) {
          const x = Math.round(canvas.width * centerX - size / 2);
          const y = Math.round(canvas.height * centerY - size / 2);
          if (
            x < 0 ||
            y < 0 ||
            x + size > canvas.width ||
            y + size > canvas.height
          ) {
            continue;
          }
          const samples = sampleImageRegion(
            pixels,
            canvas.width,
            x,
            y,
            size,
            COPY_BADGE_SAMPLE_SIZE,
          );
          distance = Math.min(distance, moveTypeDistance(samples, digitTemplate));
        }
      }
    }
    bestByLevel.set(
      template.level,
      Math.min(bestByLevel.get(template.level) ?? Infinity, distance),
    );
  }

  const ranked = [...bestByLevel.entries()].sort((left, right) => left[1] - right[1]);
  const [best] = ranked;
  return best && best[1] <= 34 ? best[0] : null;
}

async function recognizeCopies(worker, image, templates) {
  if (!templates.length) return null;
  await worker.setParameters({
    tessedit_char_whitelist: "012345",
    tessedit_pageseg_mode: "10",
  });
  const regions = [
    { x: 0.02, y: 0.68, width: 0.27, height: 0.3 },
    { x: 0.04, y: 0.7, width: 0.2, height: 0.22 },
    { x: 0.02, y: 0.72, width: 0.25, height: 0.25 },
  ];
  const readings = [];
  for (const region of regions) {
    const crop = makeOCRCrop(image, region, 10);
    const { data } = await worker.recognize(crop);
    const match = data.text.match(/[1-5]/);
    if (match) readings.push({ level: Number(match[0]), confidence: data.confidence });
  }

  const counts = new Map();
  readings.forEach((reading) => {
    const current = counts.get(reading.level) || { count: 0, confidence: 0 };
    current.count += 1;
    current.confidence = Math.max(current.confidence, reading.confidence);
    counts.set(reading.level, current);
  });
  const rankedVotes = [...counts.entries()].sort(
    (left, right) => right[1].count - left[1].count ||
      right[1].confidence - left[1].confidence,
  );
  const [level, vote] = rankedVotes[0] || [null, { count: 0, confidence: 0 }];
  const confidenceLeader = [...counts.entries()].sort(
    (left, right) => right[1].confidence - left[1].confidence,
  )[0];
  const confidenceRunnerUp = [...counts.entries()].sort(
    (left, right) => right[1].confidence - left[1].confidence,
  )[1];
  const hasClearSingleRead = Boolean(
    confidenceLeader &&
    confidenceLeader[1].confidence >= 50 &&
    confidenceLeader[1].confidence - (confidenceRunnerUp?.[1].confidence || 0) >= 20,
  );
  const hasMajorityRead = vote.count >= 2 || vote.confidence >= 70;
  const hasStrongOcrVote = hasMajorityRead || hasClearSingleRead;
  const detectedLevel = hasStrongOcrVote
    ? hasMajorityRead ? level : confidenceLeader[0]
    : detectCopyBadgeLevel(image, templates);
  if (!detectedLevel) return null;

  let excelsior = false;
  let badgeTypeConfidence = "OCR";
  if (templates.length) {
    const detected = await detectExcelsiorBadge(
      image,
      templates,
      detectedLevel,
      hasMajorityRead ? vote.count : 0,
    );
    if (detected === null) {
      badgeTypeConfidence = "ambiguous";
    } else {
      excelsior = detected;
      badgeTypeConfidence = "insignia";
    }
  }
  const value = excelsior ? detectedLevel + 5 : detectedLevel;
  return {
    value: String(value),
    readingCount: hasMajorityRead ? vote.count : hasClearSingleRead ? 1 : 0,
    readingTotal: regions.length,
    usedTemplate: !hasStrongOcrVote,
    badgeTypeConfidence,
    excelsior,
  };
}

async function recognizeMax(worker, image) {
  const regions = [
    { x: 0.05, y: 0.82, width: 0.9, height: 0.16 },
    { x: 0, y: 0.78, width: 0.95, height: 0.22 },
  ];
  for (const region of regions) {
    const crop = makeOCRCrop(image, region, 5);
    for (const pageSegMode of ["11", "10"]) {
      await worker.setParameters({
        tessedit_char_whitelist: "MAX",
        tessedit_pageseg_mode: pageSegMode,
      });
      const { data } = await worker.recognize(crop);
      const text = data.text.replace(/[^MAX]/gi, "").toUpperCase();
      if (text === "MAX") return true;
    }
  }
  return false;
}

async function analyzeCardDetails() {
  if (!cards.length) return;
  showError("");
  analysisStatus.textContent = "Analizando la insignia EX de cada tarjeta…";

  try {
    let copiesDetected = 0;
    let exYesDetected = 0;
    let exNoDetected = 0;
    let subrolesDetected = 0;
    let uncertainCopies = 0;
    let ambiguousCopyBadges = 0;
    const errors = [];
    const images = [];
    const notesByCard = cards.map(() => []);

    for (const [index, card] of cards.entries()) {
      const image = new Image();
      image.src = card.imageUrl;
      await image.decode();
      images.push(image);
      const notes = notesByCard[index];

      if (card.subroleSource === "auto") {
        card.subrole = "";
        card.subroleSource = "";
      }
      card.hasMaxLabel = undefined;

      const ex6 = detectEXMarker(image);
      if (card.ex6Source === "auto") {
        card.ex6 = "";
        card.ex6Source = "";
      }
      if (!card.ex6) {
        if (ex6) {
          card.ex6 = ex6.value;
          card.ex6Source = "auto";
          if (ex6.value === "sí") exYesDetected += 1;
          else exNoDetected += 1;
          notes.push(`6EX ${ex6.value} por ${ex6.reason} (revisar)`);
        } else {
          notes.push("6EX no concluyente");
        }
      }
      card.analysisNote = notes.join(" · ");
    }

    renderCards();
    updateSummary();

    analysisStatus.textContent = "Insignias EX analizadas. Cargando OCR para copias y MAX…";
    let worker = null;
    try {
      worker = await getOcrWorker();
    } catch (error) {
      errors.push(`OCR no disponible: ${error.message}`);
    }
    let copyBadgeTemplates = [];
    if (worker) {
      try {
        const loadedBadges = await getCopyBadgeTemplates();
        copyBadgeTemplates = loadedBadges.templates;
        if (loadedBadges.failed.length) {
          errors.push(`no se cargaron ${loadedBadges.failed.length} plantillas de insignias de copias`);
        }
      } catch (error) {
        errors.push(`plantillas de copias no disponibles: ${error.message}`);
      }
    }

    for (const [index, card] of cards.entries()) {
      const image = images[index];
      const notes = notesByCard[index];

      if (!worker) {
        if (!card.copies) {
          uncertainCopies += 1;
          notes.push("copias pendientes: OCR no disponible");
        }
        if (applyMaxDetection(card, false)) subrolesDetected += 1;
        notes.push("MAX no detectado: marcado como sin desbloquear (OCR no disponible; revisar)");
        card.analysisNote = notes.join(" · ");
        continue;
      }

      try {
        const copies = await recognizeCopies(worker, image, copyBadgeTemplates);
        if (!card.copies) {
          if (copies) {
            card.copies = copies.value;
            copiesDetected += 1;
            if (copies.badgeTypeConfidence === "ambiguous") {
              ambiguousCopyBadges += 1;
            }
            const copySource = copies.badgeTypeConfidence === "ambiguous"
              ? "tipo de insignia ambiguo"
              : copies.excelsior
                ? "insignia dentada Excelsior"
                : "insignia hexagonal";
            const copyMethod = copies.usedTemplate
              ? "plantilla visual"
              : `OCR ${copies.readingCount}/${copies.readingTotal} lecturas`;
            notes.push(
              copies.badgeTypeConfidence === "ambiguous"
                ? `copias ${copies.value} provisionales (si es Excelsior, sumar 5; ${copySource}; ${copyMethod}; revisar)`
                : `copias ${copies.value} (${copySource}; ${copyMethod}; revisar)`,
            );
          } else {
            uncertainCopies += 1;
            notes.push("copias no reconocidas o insignia ambigua; revisar manualmente");
          }
        }
      } catch (error) {
        errors.push(`tarjeta ${index + 1}, copias: ${error.message}`);
        notes.push("falló la lectura de copias");
      }

      try {
        const hasMax = await recognizeMax(worker, image);
        if (applyMaxDetection(card, hasMax)) subrolesDetected += 1;
        notes.push(
          hasMax
            ? "MAX detectado por OCR; revisar"
            : "MAX no detectado; marcado como sin desbloquear; revisar",
        );
      } catch (error) {
        errors.push(`tarjeta ${index + 1}, MAX: ${error.message}`);
        if (applyMaxDetection(card, false)) subrolesDetected += 1;
        notes.push("falló la lectura de MAX; marcado como sin desbloquear; revisar");
      }

      card.analysisNote = notes.join(" · ");
    }

    renderCards();
    updateSummary();
    analysisStatus.textContent =
      `Análisis terminado: ${copiesDetected} valores de copias; estrella 6EX clasificada en ` +
      `${exYesDetected + exNoDetected} tarjetas (${exYesDetected} sí, ${exNoDetected} no); ` +
      `${subrolesDetected} subroles rellenados. ` +
      `Insignias ambiguas con valor provisional: ${ambiguousCopyBadges}. ` +
      `No concluyentes: ${uncertainCopies} lecturas de copias. ` +
      "Comprueba siempre los campos detectados." +
      (errors.length ? ` Errores OCR: ${errors.slice(0, 3).join("; ")}.` : "");
  } catch (error) {
    showError(`No se pudo analizar la tarjeta: ${error.message}`);
    analysisStatus.textContent = "Revisa la conexión y vuelve a intentar la detección.";
  }
}

function csvCell(value) {
  const safeValue = String(value).replace(/"/g, '""');
  return `"${safeValue}"`;
}

function exportCsv() {
  if (!cards.length || !cards.every(isCardComplete)) {
    showError("Revisa los campos obligatorios de todas las tarjetas antes de exportar.");
    return;
  }

  const rows = [
    ["Personaje / compi", "Pokémon", "Copias", "6EX", "Subrol"],
    ...cards.map((card) => [
      card.pair.trim(),
      card.pokemon.trim(),
      card.copies,
      card.ex6 === "sí" ? "Sí" : "No",
      {
        sin_desbloquear: "Sin desbloquear",
        max: "MAX",
      }[card.subrole],
    ]),
  ];
  const csv = `\uFEFF${rows.map((row) => row.map(csvCell).join(";")).join("\r\n")}`;
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = downloadUrl;
  link.download = "compis-pokemon-masters.csv";
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
}

function exportSuggestionsCsv() {
  if (!cards.some((card) => card.suggestions?.length)) return;

  const rows = [
    [
      "Tarjeta",
      "Posición propuesta",
      "Entrenador candidato",
      "Traje / variante del entrenador",
      "Pokémon candidato",
      "Tipos detectados",
      "Tipos del catálogo",
      "Tipos detectados coincidentes",
      "Bonificación por tipos (máximo 6)",
      "Rol detectado",
      "Rol del catálogo",
      "Ajuste por rol",
      "Variante del icono",
      "Ruta del icono",
      "Distancia RGB media normalizada 0-100 (menor es más parecido; no es confianza)",
      "Distancia combinada usada para ordenar (RGB menos bonificaciones de tipo y rol)",
      "Estado de revisión",
    ],
  ];

  cards.forEach((card) => {
    (card.suggestions || []).forEach((suggestion, suggestionIndex) => {
      rows.push([
        card.cropIndex + 1,
        suggestionIndex + 1,
        suggestion.name,
        suggestion.trainerAlt,
        suggestion.pokemon,
        suggestion.detectedMoveTypes.map(moveTypeLabel).join(", "),
        suggestion.moveTypes.map(moveTypeLabel).join(", "),
        suggestion.detectedMoveTypes.length
          ? `${suggestion.sharedTypes}/${suggestion.detectedMoveTypes.length}`
          : "Sin detección",
        suggestion.typeBonus.toFixed(2),
        pairRoleLabel(suggestion.detectedPairRole),
        pairRoleLabel(suggestion.role),
        suggestion.roleScore.toFixed(2),
        suggestion.variant,
        suggestion.path,
        suggestion.distance.toFixed(2),
        suggestion.rankDistance.toFixed(2),
        "Pendiente de confirmar",
      ]);
    });
  });

  const csv = `\uFEFF${rows.map((row) => row.map(csvCell).join(";")).join("\r\n")}`;
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = downloadUrl;
  link.download = "resultado-reconocimiento-pixel-syncpairs-tracker.csv";
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
}

window.addEventListener("pagehide", () => {
  if (loadedImage?.url) URL.revokeObjectURL(loadedImage.url);
  catalogImageUrls.forEach((url) => URL.revokeObjectURL(url));
});


function scannerJsonp(params) {
  return new Promise((resolve, reject) => {
    const callbackName =
      "scannerCallback_" + Date.now() + "_" +
      Math.random().toString(36).slice(2);

    const script = document.createElement("script");
    const url = new URL(SHEETS_WEB_APP_URL);

    url.searchParams.set("callback", callbackName);
    Object.entries(params).forEach(([key, value]) => {
      url.searchParams.set(key, String(value));
    });

    let timeoutId;

    const cleanup = () => {
      clearTimeout(timeoutId);
      delete window[callbackName];
      script.remove();
    };

    window[callbackName] = (data) => {
      cleanup();
      resolve(data);
    };

    script.onerror = () => {
      cleanup();
      reject(new Error(
        "No se pudo contactar con Apps Script. Comprueba la URL y la implementación."
      ));
    };

    timeoutId = setTimeout(() => {
      cleanup();
      reject(new Error(
        "Se agotó el tiempo de espera de Apps Script."
      ));
    }, 15000);

    script.src = url.toString();
    document.head.append(script);
  });
}

async function cargarJugadoresSheets() {
  const select = document.getElementById("sync-player");
  const syncButton = document.getElementById("sync-button");
  const reloadButton = document.getElementById("reload-players-button");
  const status = document.getElementById("sync-status");

  if (!select || !status) return;

  // Evita lanzar varias peticiones simultáneas.
  if (reloadButton?.disabled) return;

  const previousPlayer = select.value;

  if (reloadButton) {
    reloadButton.disabled = true;
    reloadButton.textContent = "↻ Cargando jugadores...";
  }

  if (syncButton) {
    syncButton.disabled = true;
  }

  status.textContent = "Cargando la lista de jugadores...";

  try {
    const response = await scannerJsonp({
      action: "members",
      token: SHEETS_SYNC_TOKEN
    });

    if (!response.success) {
      throw new Error(
        response.error || "No se pudieron cargar los jugadores."
      );
    }

    if (!Array.isArray(response.members)) {
      throw new Error(
        "La respuesta del servidor no contiene una lista válida de jugadores."
      );
    }

    // Construimos la lista antes de sustituir la anterior.
    const fragment = document.createDocumentFragment();

    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Selecciona un jugador...";
    fragment.append(placeholder);

    response.members.forEach((member) => {
      const option = document.createElement("option");

      option.value = member;
      option.textContent = member;

      fragment.append(option);
    });

    // Actualizamos el selector únicamente cuando la respuesta es válida.
    select.replaceChildren(fragment);

    // Conservamos la selección anterior si ese jugador sigue existiendo.
    if (
      previousPlayer &&
      response.members.includes(previousPlayer)
    ) {
      select.value = previousPlayer;
    }

    if (response.members.length > 0) {
      status.textContent =
        `Lista de jugadores cargada: ${response.members.length} jugadores.`;
    } else {
      status.textContent =
        "No se han encontrado jugadores en Gym Members!A1:A20.";
    }

  } catch (error) {
    status.textContent =
      "Error al cargar jugadores: " +
      error.message +
      " Puedes volver a intentarlo con «Recargar jugadores».";

  } finally {
    if (reloadButton) {
      reloadButton.disabled = false;
      reloadButton.textContent = "↻ Recargar jugadores";
    }

    actualizarEstadoBotonSync();
  }
}

function actualizarEstadoBotonSync() {
  const select = document.getElementById("sync-player");
  const button = document.getElementById("sync-button");

  if (!select || !button) return;

  const complete = cards.length > 0 && cards.every(isCardComplete);
  button.disabled = !complete || !select.value;
}

async function esperarResultadoSync(requestId, timeoutMs = 60000) {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    const response = await scannerJsonp({
      action: "status",
      token: SHEETS_SYNC_TOKEN,
      requestId
    });

    if (!response.success) {
      throw new Error(response.error || "No se pudo consultar el resultado.");
    }

    if (response.state === "done") return response.result;

    if (response.state === "error") {
      throw new Error(response.error || "La sincronización ha fallado.");
    }

    await new Promise(resolve => setTimeout(resolve, 1500));
  }

  throw new Error(
    "No se ha recibido confirmación a tiempo. Comprueba Google Sheets antes de volver a sincronizar."
  );
}



async function sincronizarEscanerConSheets() {
  const select = document.getElementById("sync-player");
  const button = document.getElementById("sync-button");
  const status = document.getElementById("sync-status");

  if (!cards.length || !cards.every(isCardComplete)) {
    status.textContent = "Revisa todas las tarjetas antes de sincronizar.";
    return;
  }

  const player = select.value;
  if (!player) {
    status.textContent = "Selecciona un jugador.";
    return;
  }

  const requestId =
    "sync_" + Date.now() + "_" +
    Math.random().toString(36).slice(2, 10);

  const pairs = cards.map((card) => {
    const copies = Number(card.copies);
    const characterText = String(card.pair || "").trim();
    const pokemon = String(card.pokemon || "").trim();

    // Convierte "Larry - Special Costume 2" en
    // "Larry (Special Costume 2)".
    const match = characterText.match(/^(.+?)\s+-\s+(.+)$/);

    const character = match
      ? `${match[1].trim()} (${match[2].trim()})`
      : characterText;

    // Construye el nombre completo que espera Google Sheets.
    const fullName = pokemon
      ? `${character} & ${pokemon}`
      : character;

    return {
      name: fullName,
      moveLevel: copies,
      ex: card.ex6 === "sí" ? "Si" : "No",
      exRole: card.subrole === "max" ? "Si" : "No"
    };
  });

  const payload = {
    token: SHEETS_SYNC_TOKEN,
    requestId,
    player,
    pairs
  };

  button.disabled = true;
  status.textContent = "Enviando datos a Google Sheets...";

  try {
    await fetch(SHEETS_WEB_APP_URL, {
      method: "POST",
      mode: "no-cors",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify(payload)
    });

    status.textContent = "Solicitud enviada. Esperando confirmación...";

    const result = await esperarResultadoSync(requestId);

    status.textContent =
      `Sincronización completada para ${player}. ` +
      `Nuevas: ${result.added}; actualizadas: ${result.updated}; ` +
      `sin cambios: ${result.unchanged}.`;

  } catch (error) {
    status.textContent =
      "No se ha podido confirmar la sincronización: " + error.message;
  } finally {
    actualizarEstadoBotonSync();
  }
}

document.getElementById("sync-player")
  .addEventListener("change", actualizarEstadoBotonSync);

document.getElementById("sync-button")
  .addEventListener("click", sincronizarEscanerConSheets);

document.getElementById("reload-players-button")
  .addEventListener("click", cargarJugadoresSheets);

// Carga la lista automáticamente al iniciar la aplicación.
cargarJugadoresSheets();

/**
 * Convierte un píxel RGBA en un valor de escala de grises.
 */
function grayscaleAt(data, index) {
  return (
    0.299 * data[index] +
    0.587 * data[index + 1] +
    0.114 * data[index + 2]
  );
}

/**
 * Calcula la correlación normalizada entre dos regiones.
 *
 * Devuelve un valor entre -1 y 1:
 *   1  = coincidencia lineal muy alta
 *   0  = ausencia de correlación lineal
 *  -1  = correlación inversa
 *
 * La plantilla y la región candidata deben tener
 * las mismas dimensiones.
 */
function normalizedCorrelation(template, candidate) {
  const a = template.data;
  const b = candidate.data;

  if (
    template.width !== candidate.width ||
    template.height !== candidate.height ||
    a.length !== b.length
  ) {
    throw new Error(
      "La plantilla y la región candidata deben tener el mismo tamaño."
    );
  }

  const pixelCount = template.width * template.height;

  if (pixelCount < 2) {
    return -1;
  }

  let sumA = 0;
  let sumB = 0;

  for (let i = 0; i < a.length; i += 4) {
    sumA += grayscaleAt(a, i);
    sumB += grayscaleAt(b, i);
  }

  const meanA = sumA / pixelCount;
  const meanB = sumB / pixelCount;

  let covariance = 0;
  let varianceA = 0;
  let varianceB = 0;

  for (let i = 0; i < a.length; i += 4) {
    const da = grayscaleAt(a, i) - meanA;
    const db = grayscaleAt(b, i) - meanB;

    covariance += da * db;
    varianceA += da * da;
    varianceB += db * db;
  }

  const denominator = Math.sqrt(varianceA * varianceB);

  // Una región uniforme no proporciona suficiente información
  // para realizar una comparación fiable.
  if (denominator < 1e-8) {
    return -1;
  }

  return covariance / denominator;
}

/**
 * Busca una plantilla en distintas posiciones verticales.
 *
 * @param {HTMLImageElement} image Imagen original.
 * @param {Object} template Plantilla:
 *   {
 *     data: Uint8ClampedArray,
 *     width: number,
 *     height: number
 *   }
 *
 * @param {number} templateX Coordenada X de la plantilla.
 * @param {number} offsetY Desplazamiento vertical de la plantilla
 *                         respecto al borde superior de la tarjeta.
 * @param {number} minTopY Y mínima posible del borde superior.
 * @param {number} maxTopY Y máxima posible del borde superior.
 * @param {number} step Incremento vertical de la búsqueda.
 *
 * @returns {{
 *   y: number,
 *   score: number
 * } | null}
 */
function findFirstRowY(
  image,
  template,
  templateX,
  offsetY,
  minTopY,
  maxTopY,
  step = 1
) {
  const imageWidth = image.naturalWidth;
  const imageHeight = image.naturalHeight;

  if (
    !template ||
    !template.data ||
    !Number.isInteger(template.width) ||
    !Number.isInteger(template.height) ||
    template.width < 2 ||
    template.height < 2
  ) {
    throw new Error("La plantilla no es válida.");
  }

  if (
    !Number.isFinite(templateX) ||
    !Number.isFinite(offsetY) ||
    !Number.isFinite(minTopY) ||
    !Number.isFinite(maxTopY) ||
    !Number.isInteger(step) ||
    step < 1
  ) {
    throw new Error("Los parámetros de búsqueda no son válidos.");
  }

  if (
    templateX < 0 ||
    templateX + template.width > imageWidth
  ) {
    throw new Error(
      "La plantilla queda fuera del ancho de la imagen."
    );
  }

  const searchCanvas = document.createElement("canvas");

  searchCanvas.width = imageWidth;
  searchCanvas.height = imageHeight;

  const ctx = searchCanvas.getContext("2d", {
    willReadFrequently: true
  });

  if (!ctx) {
    throw new Error("No se pudo crear el contexto del canvas.");
  }

  ctx.drawImage(image, 0, 0);

  // Limita la búsqueda para que la plantilla nunca salga
  // de los límites verticales de la imagen.
  const firstTop = Math.max(
    0,
    Math.ceil(minTopY)
  );

  const lastTop = Math.min(
    imageHeight - template.height - offsetY,
    Math.floor(maxTopY)
  );

  if (firstTop > lastTop) {
    return null;
  }

  let bestMatch = null;

  for (
    let topY = firstTop;
    topY <= lastTop;
    topY += step
  ) {
    const patchY = Math.round(topY + offsetY);

    if (
      patchY < 0 ||
      patchY + template.height > imageHeight
    ) {
      continue;
    }

    const candidate = ctx.getImageData(
      templateX,
      patchY,
      template.width,
      template.height
    );

    const score = normalizedCorrelation(
      template,
      candidate
    );

    if (
      bestMatch === null ||
      score > bestMatch.score
    ) {
      bestMatch = {
        y: topY,
        score
      };
    }
  }

  return bestMatch;
}

/**
 * Actualiza grid.y utilizando una plantilla de referencia.
 *
 * templateConfig debe contener:
 *   data
 *   width
 *   height
 *   x
 *   offsetY
 *
 * Los límites de búsqueda se expresan en píxeles.
 */
function autoDetectGridY(
  image,
  grid,
  templateConfig,
  options = {}
) {
  const {
    minY = 150,
    maxY = 400,
    minScore = 0.80,
    step = 1
  } = options;

  const result = findFirstRowY(
    image,
    templateConfig,
    templateConfig.x,
    templateConfig.offsetY,
    minY,
    maxY,
    step
  );

  if (!result) {
    throw new Error(
      "No se encontró ninguna posición válida para la primera fila."
    );
  }

  if (result.score < minScore) {
    throw new Error(
      `No se ha encontrado una coincidencia suficientemente fiable. ` +
      `Puntuación obtenida: ${result.score.toFixed(3)}.`
    );
  }

  // Convierte los píxeles a porcentaje de la altura original.
  const detectedYPercent =
    (result.y / image.naturalHeight) * 100;

  // Solo cambia Y. Se conservan X, ancho, alto, filas y columnas.
  grid.y = detectedYPercent;

  return {
    yPixels: result.y,
    yPercent: detectedYPercent,
    score: result.score
  };
}

