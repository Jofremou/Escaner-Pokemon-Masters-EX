
/**
 * Sincroniza las parejas reconocidas por el escáner local
 * con la hoja del jugador seleccionado.
 *
 * Encabezados en la fila 2.
 * Datos desde la fila 3.
 *
 * Campos que modifica:
 * Player, Sync Pair, Superawakening, Move Level, EX
 * y la columna EX Role situada más a la derecha.
 */

const SCANNER_CONFIG = {
  HEADER_ROW: 2,
  FIRST_DATA_ROW: 3,

  // Sustituye este valor por una cadena privada tuya.
  // Usa la misma cadena en el HTML local.
  TOKEN:
};



function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = String(params.callback || "");
  const token = String(params.token || "");
  const action = String(params.action || "");

  if (!/^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)*$/.test(callback)) {
    return ContentService.createTextOutput("Callback no válido");
  }

  let result;

  try {
    if (token !== SCANNER_CONFIG.TOKEN) {
      throw new Error("Clave de sincronización incorrecta.");
    }

    if (action === "members") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const membersSheet = ss.getSheetByName("Gym Members");

      if (!membersSheet) {
        throw new Error('No existe la hoja "Gym Members".');
      }

      const members = membersSheet
        .getRange("A1:A20")
        .getDisplayValues()
        .flat()
        .map(name => name.trim())
        .filter(Boolean);

      result = { success: true, members: [...new Set(members)] };

    } else if (action === "status") {
      const requestId = String(params.requestId || "");

      if (!/^[a-zA-Z0-9_-]{1,100}$/.test(requestId)) {
        throw new Error("Identificador de solicitud no válido.");
      }

      const key = "scanner_status_" + requestId;
      const properties = PropertiesService.getScriptProperties();
      const saved = properties.getProperty(key);

      if (!saved) {
        result = { success: true, state: "pending" };
      } else {
        result = JSON.parse(saved);

        if (result.state === "done" || result.state === "error") {
          properties.deleteProperty(key);
        }
      }
    } else {
      throw new Error("Acción desconocida.");
    }
  } catch (error) {
    result = { success: false, error: error.message };
  }

  return ContentService.createTextOutput(
    callback + "(" + JSON.stringify(result) + ");"
  ).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  let requestId = "";

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No se han recibido datos.");
    }

    const payload = JSON.parse(e.postData.contents);
    requestId = String(payload.requestId || "");

    if (!/^[a-zA-Z0-9_-]{1,100}$/.test(requestId)) {
      throw new Error("Identificador de solicitud no válido.");
    }

    if (payload.token !== SCANNER_CONFIG.TOKEN) {
      throw new Error("Clave de sincronización incorrecta.");
    }

    const result = sincronizarCompis_(payload);

    guardarEstadoSincronizacion_(requestId, {
      success: true,
      state: "done",
      result: result
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, result: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    if (requestId) {
      guardarEstadoSincronizacion_(requestId, {
        success: false,
        state: "error",
        error: error.message
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function guardarEstadoSincronizacion_(requestId, status) {
  PropertiesService.getScriptProperties().setProperty(
    "scanner_status_" + requestId,
    JSON.stringify(status)
  );
}



/**
 * Busca la pestaña del jugador, compara las parejas
 * y modifica únicamente los campos administrados
 * por el escáner.
 *
 * Formato esperado:
 * {
 *   player: "Jofre",
 *   pairs: [
 *     {
 *       name: "Red (Classic) & Snorlax",
 *       moveLevel: 3,
 *       ex: true,
 *       exRole: true
 *     }
 *   ]
 * }
 */

function sincronizarCompis_(payload) {
  const player = String(payload.player || "").trim();
  const pairs = payload.pairs;

  if (!player) {
    throw new Error("Falta el nombre del jugador.");
  }

  if (!Array.isArray(pairs)) {
    throw new Error("La lista de parejas no es válida.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(player);

  if (!sheet) {
    throw new Error('No existe una pestaña llamada "' + player + '".');
  }

  if (!isGymMemberSheet_(ss, player)) {
    throw new Error('"' + player + '" no figura como miembro del gimnasio.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const lastColumn = sheet.getLastColumn();

    if (lastColumn < 1) {
      throw new Error("La hoja no contiene encabezados.");
    }

    const headers = sheet
      .getRange(
        SCANNER_CONFIG.HEADER_ROW,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(value => String(value).trim());

    /*
     * Asigna las columnas por encabezado.
     * Si un encabezado aparece varias veces,
     * conserva la posición de la última aparición.
     */
    const columns = {};

    headers.forEach((header, index) => {
      if (header) {
        columns[header] = index + 1;
      }
    });

    const requiredHeaders = [
      "Player",
      "Sync Pair",
      "Move Level",
      "EX",
      "EX Role"
    ];

    const missingHeaders = requiredHeaders.filter(
      header => !columns[header]
    );

    if (missingHeaders.length > 0) {
      throw new Error(
        "Faltan encabezados obligatorios: " +
        missingHeaders.join(", ")
      );
    }

    const playerCol = columns["Player"];
    const pairCol = columns["Sync Pair"];
    const moveLevelCol = columns["Move Level"];
    const exCol = columns["EX"];
    const exRoleCol = columns["EX Role"];

    const firstRow = SCANNER_CONFIG.FIRST_DATA_ROW;

    /*
     * Elimina duplicados de Sync Pair.
     * Conserva la primera aparición y elimina
     * las filas repetidas posteriores.
     *
     * La comparación ignora mayúsculas, minúsculas
     * y espacios sobrantes.
     */
    let lastRow = sheet.getLastRow();
    const rowsToDelete = [];

    if (lastRow >= firstRow) {
      const names = sheet.getRange(
        firstRow,
        pairCol,
        lastRow - firstRow + 1,
        1
      ).getDisplayValues();

      const seenNames = new Set();

      names.forEach((row, index) => {
        const name = String(row[0]).trim();

        if (!name) return;

        const normalizedName = normalizarNombreCompi_(name);

        if (seenNames.has(normalizedName)) {
          rowsToDelete.push(firstRow + index);
        } else {
          seenNames.add(normalizedName);
        }
      });
    }

    // Esborra de baix a dalt per no desplaçar les files pendents.
    rowsToDelete.reverse().forEach(rowNumber => {
      sheet.deleteRow(rowNumber);
    });

    if (rowsToDelete.length > 0) {
      SpreadsheetApp.flush();

      Logger.log(
        "S'han eliminat " + rowsToDelete.length +
        " files duplicades de la pestanya " + player
      );
    }

    // Reconstrueix el mapa després d'eliminar els duplicats.
    lastRow = sheet.getLastRow();
    const existingPairs = new Map();

    if (lastRow >= firstRow) {
      const names = sheet.getRange(
        firstRow,
        pairCol,
        lastRow - firstRow + 1,
        1
      ).getDisplayValues();

      names.forEach((row, index) => {
        const name = String(row[0]).trim();

        if (!name) return;

        existingPairs.set(
          normalizarNombreCompi_(name),
          firstRow + index
        );
      });
    }

    let updated = 0;
    let added = 0;
    let unchanged = 0;

    const seenInPayload = new Set();

    pairs.forEach(pair => {

    const rawName = String(pair.name || "").trim();

    if (!rawName) {
      throw new Error("Se ha recibido una pareja sin nombre.");
    }

    // Primero comprueba el nombre contra las opciones del desplegable.
    const name = obtenerNombreDesplegable_(
      sheet,
      pairCol,
      firstRow,
      rawName
    );

    // Utiliza la misma normalización para todas las comparaciones.
    const normalizedName = normalizarNombreCompi_(name);


      if (seenInPayload.has(normalizedName)) {
        throw new Error(
          'El escáner ha enviado la pareja "' +
          name + '" más de una vez.'
        );
      }

      seenInPayload.add(normalizedName);

      const moveLevel = Number(pair.moveLevel);

      if (
        !Number.isInteger(moveLevel) ||
        moveLevel < 1 ||
        moveLevel > 10
      ) {
        throw new Error(
          'Nivel de movimientos no válido para "' + name + '".'
        );
      }

      const ex = normalizarSiNo_(pair.ex, "EX", name);
      const exRole = normalizarSiNo_(pair.exRole, "EX Role", name);


      const valuesToWrite = [
        [pairCol, name],
        [playerCol, player],
        [moveLevelCol, moveLevel],
        [exCol, ex],
        [exRoleCol, exRole]
      ];

      Logger.log(JSON.stringify({
        player: player,
        name: name,
        moveLevel: moveLevel,
        ex: ex,
        exRole: exRole,
        columns: {
          player: playerCol,
          pair: pairCol,
          moveLevel: moveLevelCol,
          ex: exCol,
          exRole: exRoleCol
        }
      }));
      
      let rowNumber = existingPairs.get(normalizedName);
      const isNew = rowNumber === undefined;

      if (isNew) {
        const lastRowForPairs = Math.max(sheet.getLastRow(), firstRow);
        const pairNames = sheet.getRange(
          firstRow,
          pairCol,
          lastRowForPairs - firstRow + 1,
          1
        ).getDisplayValues().flat();

        const emptyIndex = pairNames.findIndex(
          value => String(value).trim() === ""
        );

        rowNumber = emptyIndex !== -1
          ? firstRow + emptyIndex
          : lastRowForPairs + 1;
      }

      let changed = false;

      valuesToWrite.forEach(([column, value]) => {
        const cell = sheet.getRange(rowNumber, column);
        const current = cell.getValue();

        if (!valoresIguales_(current, value)) {
          cell.setValue(value);
          changed = true;
        }
      });

      if (isNew) {
        existingPairs.set(normalizedName, rowNumber);
        added++;
      } else if (changed) {
        updated++;
      } else {
        unchanged++;
      }
    });

    // Actualiza el roster si se han modificado o añadido parejas,
    // o si se han eliminado duplicados.
    if (added > 0 || updated > 0 || rowsToDelete.length > 0) {
      SpreadsheetApp.flush();
      updateRoster(player);
      SpreadsheetApp.flush();
    }

    return {
      player: player,
      added: added,
      updated: updated,
      unchanged: unchanged,
      duplicatesRemoved: rowsToDelete.length,
      debug: {
        playerColumn: playerCol,
        pairColumn: pairCol,
        moveLevelColumn: moveLevelCol,
        exColumn: exCol,
        exRoleColumn: exRoleCol,
        lastPair: pairs.length ? pairs[pairs.length - 1] : null
      }
    };

  } finally {
    lock.releaseLock();
  }
}



/**
 * Normaliza valores booleanos y textos del escáner
 * a los valores que utiliza la hoja.
 *
 * EX y EX Role: "Si" / "No".
 */
function normalizarSiNo_(value, field, pairName) {
  if (typeof value === "boolean") {
    return value ? "Si" : "No";
  }

  const normalized = String(value).trim().toLowerCase();

  if (["si", "sí", "yes", "true", "1"].includes(normalized)) {
    return "Si";
  }

  if (["no", "false", "0"].includes(normalized)) {
    return "No";
  }

  throw new Error(
    'Valor de "' + field + '" no válido para "' + pairName + '".'
  );
}


/**
 * Compara los valores sin distinguir mayúsculas,
 * ni diferencias entre números y números en texto.
 */
function valoresIguales_(a, b) {
  if (typeof b === "number") {
    return a !== "" && Number(a) === b;
  }

  return String(a ?? "").trim().toLowerCase() ===
         String(b ?? "").trim().toLowerCase();
}


function normalizarNombreCompi_(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase();
}


function obtenerNombreDesplegable_(sheet, pairCol, row, name) {
  const cell = sheet.getRange(row, pairCol);
  const validation = cell.getDataValidation();

  if (!validation) {
    return name;
  }

  const criteriaType = validation.getCriteriaType();
  const criteriaValues = validation.getCriteriaValues();
  let options = [];

  if (
    criteriaType ===
    SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST
  ) {
    options = criteriaValues[0] || [];

  } else if (
    criteriaType ===
    SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE
  ) {
    const sourceRange = criteriaValues[0];

    options = sourceRange
      .getDisplayValues()
      .flat()
      .filter(value => String(value).trim() !== "");

  } else {
    // Altres tipus de validació: deixa que Google Sheets
    // apliqui les seves pròpies regles.
    return name;
  }

  const normalizedName = normalizarNombreCompi_(name);

  const match = options.find(
    option =>
      normalizarNombreCompi_(option) === normalizedName
  );

  if (match === undefined) {
    const sourceDescription =
      criteriaType ===
      SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE
        ? criteriaValues[0].getSheet().getName() +
          "!" + criteriaValues[0].getA1Notation()
        : "la llista manual";

    throw new Error(
      'La parella "' + name +
      '" no existeix entre les opcions permeses de Sync Pair (' +
      sourceDescription + ")."
    );
  }

  return String(match);
}

function diagnosticarDesplegableJofre() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Jofre");

  [3, 126, 127].forEach(row => {
    const cell = sheet.getRange(row, 2);
    const validation = cell.getDataValidation();

    Logger.log(JSON.stringify({
      celda: cell.getA1Notation(),
      valor: cell.getDisplayValue(),
      tipoValidacion: validation
        ? String(validation.getCriteriaType())
        : "Sin validación",
      criterios: validation
        ? validation.getCriteriaValues().map(value =>
            value && value.getA1Notation
              ? value.getA1Notation()
              : String(value)
          )
        : []
    }));
  });
}

