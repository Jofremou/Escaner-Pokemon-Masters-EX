/**
 * Comprueba si una hoja corresponde a un miembro del gym.
 *
 * Los nombres válidos están en Gym Members!A1:A20.
 */
function isGymMemberSheet_(ss, sheetName) {

  const membersSheet = ss.getSheetByName("Gym Members");

  if (!membersSheet) return false;

  const members = membersSheet
    .getRange(1, 1, 20, 1)
    .getDisplayValues()
    .flat()
    .map(member => member.trim())
    .filter(member => member !== "");

  return members.includes(sheetName.trim());
}


/**
 * ACTIVADOR INSTALABLE: Al editar
 *
 * Solo actúa cuando se edita una hoja cuyo nombre
 * coincide con un miembro incluido en Gym Members!A1:A20.
 */
function rosterOnEdit(e) {

  if (!e || !e.range || !e.source) return;

  const ss = e.source;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  /*
   * Primera comprobación:
   * si la hoja no pertenece a un miembro, no hacer nada.
   */
  if (!isGymMemberSheet_(ss, sheetName)) return;


  // Ignorar filas anteriores a la fila 3.
  if (e.range.getLastRow() < 3) return;


  const firstEditedColumn = e.range.getColumn();
  const lastEditedColumn = e.range.getLastColumn();

  // Columnas manuales: A, B, K, L y M.
  const manualColumns = [1, 2, 11, 12, 13];

  const affectsManualColumn = manualColumns.some(column =>
    column >= firstEditedColumn &&
    column <= lastEditedColumn
  );

  // No actualizar el Roster por cambios en otras columnas.
  if (!affectsManualColumn) return;


  Logger.log("Actualizando roster desde: " + sheetName);

  SpreadsheetApp.flush();

  updateRoster(sheetName);
}


/**
 * ACTIVADOR INSTALABLE: Al cambiar
 *
 * Solo actúa cuando se elimina una fila completa
 * en una hoja cuyo nombre coincide con un miembro del gym.
 */
function onChange(e) {

  if (!e || !e.source || !e.changeType) return;


  /*
   * Las ediciones normales las gestiona rosterOnEdit.
   * Este activador solo gestiona filas eliminadas.
   */
  if (e.changeType !== "REMOVE_ROW") return;


  const ss = e.source;
  const sheet = ss.getActiveSheet();

  if (!sheet) return;

  const sheetName = sheet.getName();


  /*
   * Si la hoja activa no pertenece a un miembro,
   * no hacer absolutamente nada.
   */
  if (!isGymMemberSheet_(ss, sheetName)) return;


  Logger.log("Fila eliminada en: " + sheetName);

  SpreadsheetApp.flush();

  updateRoster(sheetName);
}


/**
 * Actualiza en Roster únicamente las unidades del jugador indicado.
 *
 * Incluye un DocumentLock para impedir que dos activadores
 * escriban simultáneamente y creen duplicados.
 */
function updateRoster(playerName) {

  const PLAYER_START_ROW = 3;
  const ROSTER_START_ROW = 2;
  const NUM_COLUMNS = 16;

  const lock = LockService.getDocumentLock();
  const lockAcquired = lock.tryLock(30000);

  if (!lockAcquired) {
    Logger.log(
      "No se pudo obtener el bloqueo para actualizar: " + playerName
    );
    return;
  }

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const playerSheet = ss.getSheetByName(playerName);
    const rosterSheet = ss.getSheetByName("Roster");

    if (!playerSheet) {
      Logger.log("No existe la hoja del jugador: " + playerName);
      return;
    }

    if (!rosterSheet) {
      Logger.log('No existe la hoja "Roster".');
      return;
    }

    Logger.log("Actualizando roster desde: " + playerName);

    // ==================================================
    // Leer los datos actuales de la hoja del jugador
    // Las unidades empiezan en la fila 3
    // ==================================================

    const playerLastRow = playerSheet.getLastRow();

    let playerRows = [];

    if (playerLastRow >= PLAYER_START_ROW) {

      playerRows = playerSheet
        .getRange(
          PLAYER_START_ROW,
          1,
          playerLastRow - PLAYER_START_ROW + 1,
          NUM_COLUMNS
        )
        .getDisplayValues()
        .filter(row =>
          row[0].trim() !== "" &&
          row[1].trim() !== "" &&
          row[10].trim() !== "" &&
          row[11].trim() !== "" &&
          row[12].trim() !== ""
        );
    }

    // ==================================================
    // Crear mapa de unidades del jugador
    // Clave: jugador | unidad
    // ==================================================

    const playerMap = new Map();

    playerRows.forEach(row => {

      const player = row[0].trim();
      const syncPair = row[1].trim();

      const key = player + "|" + syncPair;

      if (playerMap.has(key)) {
        Logger.log(
          "Unidad duplicada en la hoja de " +
          playerName +
          ": " +
          syncPair
        );
      }

      playerMap.set(key, row);
    });

    // ==================================================
    // Leer el Roster
    // Los datos empiezan en la fila 2
    // ==================================================

    const rosterLastRow = rosterSheet.getLastRow();

    let rosterRows = [];

    if (rosterLastRow >= ROSTER_START_ROW) {

      rosterRows = rosterSheet
        .getRange(
          ROSTER_START_ROW,
          1,
          rosterLastRow - ROSTER_START_ROW + 1,
          NUM_COLUMNS
        )
        .getDisplayValues();
    }

    const rowsToDelete = [];
    const rowsToUpdate = [];
    const processedRosterKeys = new Set();

    // ==================================================
    // Revisar las filas existentes del jugador
    // ==================================================

    rosterRows.forEach((row, index) => {

      /*
       * index 0 corresponde a la fila 2 del Roster.
       */
      const rosterRowNumber = index + ROSTER_START_ROW;

      const rosterPlayer = row[0].trim();

      // Solo modificar las filas del jugador actual.
      if (rosterPlayer !== playerName.trim()) return;

      const rosterSyncPair = row[1].trim();
      const key = rosterPlayer + "|" + rosterSyncPair;

      /*
       * La unidad ya no existe en la hoja del jugador.
       */
      if (!playerMap.has(key)) {

        rowsToDelete.push(rosterRowNumber);
        return;
      }

      /*
       * La unidad ya apareció anteriormente en el Roster.
       * Esta fila es un duplicado.
       */
      if (processedRosterKeys.has(key)) {

        rowsToDelete.push(rosterRowNumber);

        Logger.log(
          "Duplicado eliminado de Roster: " +
          playerName +
          " - " +
          rosterSyncPair
        );

        return;
      }

      const newRow = playerMap.get(key);

      const rowHasChanges = newRow.some(
        (value, columnIndex) => value !== row[columnIndex]
      );

      if (rowHasChanges) {

        rowsToUpdate.push({
          rowNumber: rosterRowNumber,
          values: newRow
        });
      }

      processedRosterKeys.add(key);

      /*
       * La unidad ya existe en el Roster y no debe añadirse
       * posteriormente como una fila nueva.
       */
      playerMap.delete(key);
    });

    // ==================================================
    // Actualizar filas modificadas
    // ==================================================

    rowsToUpdate.forEach(item => {

      rosterSheet
        .getRange(
          item.rowNumber,
          1,
          1,
          NUM_COLUMNS
        )
        .setValues([item.values]);
    });

    // ==================================================
    // Eliminar filas borradas o duplicadas
    // Siempre de abajo hacia arriba
    // ==================================================

    rowsToDelete
      .sort((a, b) => b - a)
      .forEach(rowNumber => {
        rosterSheet.deleteRow(rowNumber);
      });

    // ==================================================
    // Añadir las nuevas unidades
    // ==================================================

    const rowsToAdd = Array.from(playerMap.values());

    if (rowsToAdd.length > 0) {

      /*
       * Si el Roster está vacío, empieza en la fila 2.
       * En caso contrario, añade después de la última fila.
       */
      const firstAvailableRow = Math.max(
        ROSTER_START_ROW,
        rosterSheet.getLastRow() + 1
      );

      rosterSheet
        .getRange(
          firstAvailableRow,
          1,
          rowsToAdd.length,
          NUM_COLUMNS
        )
        .setValues(rowsToAdd);
    }

    SpreadsheetApp.flush();

    Logger.log(
      "Roster actualizado correctamente para: " +
      playerName +
      ". Nuevas: " +
      rowsToAdd.length +
      ". Eliminadas: " +
      rowsToDelete.length +
      ". Actualizadas: " +
      rowsToUpdate.length
    );

  } catch (error) {

    Logger.log(
      "Error actualizando el roster de " +
      playerName +
      ": " +
      error.message
    );

    throw error;

  } finally {

    lock.releaseLock();
  }
}

function writeRoster() {

  const lock = LockService.getDocumentLock();
  let lockAcquired = false;

  try {

    lockAcquired = lock.tryLock(30000);

    if (!lockAcquired) {
      Logger.log("No se pudo obtener el bloqueo para reconstruir el Roster.");
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const originSheet = ss.getActiveSheet();

    const rosterSheet = ss.getSheetByName("Roster");
    const membersSheet = ss.getSheetByName("Gym Members");
    const metaSheet = ss.getSheetByName("Meta");

    if (!rosterSheet) {
      throw new Error('No se encontró la hoja "Roster".');
    }

    if (!membersSheet) {
      throw new Error('No se encontró la hoja "Gym Members".');
    }

    if (!metaSheet) {
      throw new Error('No se encontró la hoja "Meta".');
    }

    SpreadsheetApp.flush();


    // ==========================================
    // Normalizar nombres
    // ==========================================

    const normalizeText = value =>
      String(value || "")
        .trim()
        .toLowerCase();


    // ==========================================
    // Leer unidades meta desde Meta!A2:A
    // ==========================================

    const metaUnits = new Set();
    const metaLastRow = metaSheet.getLastRow();

    if (metaLastRow >= 2) {

      metaSheet
        .getRange(
          2,
          1,
          metaLastRow - 1,
          1
        )
        .getDisplayValues()
        .flat()
        .map(normalizeText)
        .filter(syncPair => syncPair !== "")
        .forEach(syncPair => metaUnits.add(syncPair));
    }


    // ==========================================
    // Leer miembros desde Gym Members!A1:A20
    // ==========================================

    const members = membersSheet
      .getRange(1, 1, 20, 1)
      .getDisplayValues()
      .flat()
      .map(member => member.trim())
      .filter(member => member !== "");


    // ==========================================
    // Recopilar datos de todos los jugadores
    // ==========================================

    const allRosterRows = [];

    /*
     * Impide añadir dos veces la misma combinación:
     * jugador + unidad.
     */
    const processedKeys = new Set();


    members.forEach(playerName => {

      const playerSheet = ss.getSheetByName(playerName);

      if (!playerSheet) {

        Logger.log(
          "No se encontró la hoja del jugador: " + playerName
        );

        return;
      }

      const lastRow = playerSheet.getLastRow();

      if (lastRow < 3) return;


      const playerRows = playerSheet
        .getRange(
          3,
          1,
          lastRow - 2,
          16
        )
        .getValues();


      playerRows.forEach((row, index) => {

        const validRow =
          String(row[0]).trim() !== "" &&
          String(row[1]).trim() !== "" &&
          String(row[10]).trim() !== "" &&
          String(row[11]).trim() !== "" &&
          String(row[12]).trim() !== "";

        if (!validRow) return;


        const player = normalizeText(row[0]);
        const syncPair = normalizeText(row[1]);

        const key = player + "|" + syncPair;


        // Evitar duplicados
        if (processedKeys.has(key)) {

          Logger.log(
            "Fila duplicada ignorada en " +
            playerName +
            ", fila " +
            (index + 3) +
            ": " +
            row[1]
          );

          return;
        }

        processedKeys.add(key);


        // ==========================================
        // Calcular columna O
        // ==========================================

        const scoreN = Number(row[13]);

        if (Number.isFinite(scoreN)) {

          const adjustedScore = metaUnits.has(syncPair)
            ? scoreN
            : scoreN * 0.4;

          row[14] = Math.round(adjustedScore * 100) / 100;

        } else {

          row[14] = "";
        }


        allRosterRows.push(row);
      });
    });


    // ==========================================
    // Limpiar Roster desde la fila 2
    // ==========================================

    const rosterLastRow = rosterSheet.getLastRow();

    if (rosterLastRow >= 2) {

      rosterSheet
        .getRange(
          2,
          1,
          rosterLastRow - 1,
          16
        )
        .clearContent();
    }


    // ==========================================
    // Escribir Roster desde la fila 2
    // ==========================================

    if (allRosterRows.length > 0) {

      rosterSheet
        .getRange(
          2,
          1,
          allRosterRows.length,
          16
        )
        .setValues(allRosterRows);
    }

    SpreadsheetApp.flush();


    // Mantener la hoja desde la que se pulsó el botón
    if (originSheet) {
      originSheet.activate();
    }


    ss.toast(
      "El Roster se ha actualizado correctamente. " +
      allRosterRows.length +
      " unidades procesadas.",
      "Actualización completada",
      5
    );


    Logger.log(
      "Roster reconstruido correctamente. " +
      "Total de unidades: " +
      allRosterRows.length
    );

  } catch (error) {

    Logger.log(
      "Error reconstruyendo el Roster: " +
      error.message
    );

    throw error;

  } finally {

    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

/**
 * =====================================================================
 * PASSIO BATTLE GYM - ASIGNACIÓN MULTIALGORITMO SELECCIONABLE
 * =====================================================================
 *
 * Genera las tablas Global, Física, Especial, diagnóstico del catálogo,
 * Seteadores y alternativas. La lógica se elige mediante el desplegable
 * situado en Asignacion Optima!B2.
 *
 * PERFILES DISPONIBLES
 * --------------------
 * - Cobertura equilibrada (Max-Min).
 * - Principales más potentes.
 * - Debilidad global del catálogo.
 * - Híbrido completo.
 *
 * FUERZA DE LAS UNIDADES
 * ----------------------
 *
 * - Roster!N contiene la fuerza actual y ya incorpora el multiplicador
 *   de Meta!C mediante la fórmula de Strength Score.
 * - Roster!P contiene el máximo potencial de la unidad.
 * - Este script NO vuelve a aplicar Meta!C ni un multiplicador numérico
 *   adicional por banner. Evita así ponderar dos veces Arc Suit, Master EX,
 *   Fair EX, Master, Fair, Seasonal, etc.
 * - Roster!H sí se utiliza para ORDENAR las unidades NO Meta que rellenan
 *   huecos cuando un jugador posee menos unidades Meta que K. Primero se
 *   eligen los banners de mayor prioridad y, dentro del mismo banner, la
 *   unidad con mayor Roster!N.
 *
 * DEBILIDAD GLOBAL DEL TIPO
 * -------------------------
 *
 * Para Global, Física y Especial se calcula por separado:
 *
 *   potencia relativa = suma Roster!P Meta del tipo / mayor suma de la modalidad
 *   escasez relativa  = unidades Meta del tipo / mayor número de la modalidad
 *
 *   debilidad global =
 *     70 % × (1 - potencia relativa)
 *   + 30 % × (1 - escasez relativa)
 *
 * Cuanto mayor sea la debilidad, más importante es dedicar a ese tipo
 * un Principal difícil de sustituir. Esto permite priorizar naturalmente
 * catálogos como Hielo físico, Hada físico o Eléctrico físico cuando su
 * techo real es bajo, sin fijar jugadores o tipos por nombre.
 *
 * PRINCIPALES
 * -----------
 *
 * - Deben tener EX Zone/Weather para el tipo.
 * - Se seleccionan antes que los apoyos.
 * - Los jugadores globalmente fuertes y los especialistas reciben
 *   prioridad para permanecer en el grupo de Principales.
 * - El coste de oportunidad se recalcula con los jugadores disponibles.
 * - Las pérdidas de los tipos débiles se amplifican mediante su índice de
 *   debilidad global.
 * - Se añade una solución semilla obtenida mediante min-cost max-flow,
 *   que optimiza globalmente una función aditiva de fuerza, debilidad y
 *   sustitución. Después, el beam search explora otras combinaciones y la
 *   evaluación max-min selecciona la mejor solución completa encontrada.
 *
 * ROLES
 * -----
 *
 * - 8 Principales únicos.
 * - 8 Apoyos 1 únicos entre los jugadores restantes.
 * - Los 4 restantes son Apoyo 2 y aparecen exactamente dos veces.
 * - Principal nunca baja a Apoyo 1 o Apoyo 2.
 * - Apoyo 1 nunca baja a Apoyo 2.
 * - Los intercambios finales solo se realizan dentro del mismo rol.
 *
 * OBJETIVO FINAL
 * --------------
 *
 * 1. Mantener jugadores prioritarios como Principales.
 * 2. Minimizar la peor pérdida ponderada de Principal en tipos débiles.
 * 3. Maximizar la calidad global del grupo de Principales.
 * 4. Maximizar la peor cobertura final del trío.
 * 5. Maximizar Principal + Apoyo 1.
 * 6. Reducir dependencia de Apoyo 2.
 * 7. Mejorar respaldo, cobertura Meta y suma total.
 */
function generarAsignacionOptima() {
  const CONFIG = {
    SPREADSHEET_ID: "",

    GUIDE_SHEET: "Guia Del Evento",
    MEMBERS_SHEET: "Gym Members",
    ROSTER_SHEET: "Roster",
    META_SHEET: "Meta",
    SYNC_PAIRS_SHEET: "SyncPairs",
    OUTPUT_SHEET: "Asignacion Optima",

    /*
     * Selector de algoritmo en Asignacion Optima!B2.
     * El texto visible se conserva aunque la hoja se limpie y regenere.
     */
    ALGORITHM_SELECTOR_CELL: "B2",
    DEFAULT_ALGORITHM_MODE: "HYBRID_FULL",
    ALGORITHM_MODE: "HYBRID_FULL",
    SELECTED_ALGORITHM_LABEL: "Híbrido completo",

    ALGORITHM_OPTIONS: [
      {
        id: "BALANCED_COVERAGE",
        label: "Cobertura equilibrada (Max-Min)",
        description:
          "Prioriza el peor porcentaje final del trío y después equilibra Principales, P+A1 y respaldo."
      },
      {
        id: "STRONG_PRINCIPALS",
        label: "Principales más potentes",
        description:
          "Prioriza que los jugadores de mayor fuerza general y Meta formen el grupo de Principales."
      },
      {
        id: "GLOBAL_WEAKNESS",
        label: "Debilidad global del catálogo",
        description:
          "Concentra mejores jugadores en tipos con pocas unidades Meta o un techo máximo bajo."
      },
      {
        id: "HYBRID_FULL",
        label: "Híbrido completo",
        description:
          "Combina Principales potentes, debilidad global, coste de sustitución y cobertura Max-Min."
      }
    ],

    /* Banderas configuradas automáticamente por el perfil elegido. */
    USE_GLOBAL_WEAKNESS: true,
    USE_PRINCIPAL_PRIORITY: true,
    USE_PREFERRED_PRINCIPALS: true,
    USE_MANDATORY_PRINCIPALS: true,

    PRINCIPAL_ALTERNATIVES_PER_TYPE: 3,

    // Gym Members.
    MEMBERS_START_ROW: 1,
    MEMBERS_COLUMN: 1,

    // Guia Del Evento.
    TYPE_ROW: 6,
    RANKING_START_ROW: 8,
    FIRST_TYPE_COLUMN: 2,
    NUMBER_OF_TYPES: 8,
    COLUMNS_PER_TYPE: 2,

    // Meta.
    META_START_ROW: 2,
    META_SYNC_PAIR_COLUMN: 1,
    META_TYPE_COLUMN: 2,

    // SyncPairs.
    SYNC_PAIRS_SYNC_PAIR_COLUMN: 1,
    SYNC_PAIRS_STYLE_COLUMN: 4,

    // Roster.
    ROSTER_START_ROW: 2,
    ROSTER_PLAYER_COLUMN: 1,
    ROSTER_SYNC_PAIR_COLUMN: 2,
    ROSTER_TYPE_COLUMN: 3,
    ROSTER_STYLE_COLUMN: 6,
    ROSTER_BANNER_COLUMN: 8,
    ROSTER_EX_ZONE_COLUMN: 9,
    ROSTER_REBUFF_COLUMN: 10,
    ROSTER_SCORE_COLUMN: 14,
    ROSTER_MAX_SCORE_COLUMN: 16,

    YES_VALUE: "Yes",
    PHYSICAL_VALUE: "Physical",
    SPECIAL_VALUE: "Special",
    HYBRID_VALUE: "Hybrid",

    /*
     * Prioridad de unidades NO Meta utilizadas para rellenar huecos.
     *
     * IMPORTANTE:
     * - No es un multiplicador de puntos.
     * - Roster!N sigue aportando el 100 % de la puntuación de la unidad.
     * - Esta lista solo decide QUÉ unidad no Meta ocupa primero una plaza.
     * - Dentro del mismo banner se prioriza la unidad con mayor Roster!N.
     * - Cualquier banner no listado queda al final.
     *
     * Los textos coinciden con los valores reales de Roster!H.
     */
    NON_META_BANNER_PRIORITY: [
      "Arc Suit Fair Scout",
      "EX Master Fair Scout",
      "EX Fair Scout",
      "Master Fair Scout",
      "Poké Fair Scout",
      "Seasonal Scout",
      "Special Costume Scout",
      "Variety Scout",
      "Mix Scout",
      "Gym Scout",
      "Spotlight Scout / General Pool",
      "Main Story: PML Arc",
      "Legendary Adventures",
      "Battle Points Exchange",
      "Event Reward",
      "Trainer Lodge Exchange",
      "Training Ticket Exchange"
    ],

    // Atacantes.
    PLAYERS_PER_TYPE: 3,
    PRINCIPAL_COUNT: 8,
    SUPPORT1_COUNT: 8,
    SUPPORT2_ASSIGNMENTS_PER_PLAYER: 2,

    /*
     * Tanto Meta como no Meta usan directamente Roster!N.
     * Las no Meta solo rellenan las plazas libres hasta K, pero no se
     * aplica ningún multiplicador adicional en este script.
     */

    /* Búsqueda amplia de combinaciones. */
    PRINCIPAL_BEAM_WIDTH: 5000,
    PRINCIPAL_FINALISTS: 54,
    SUPPORT1_BEAM_WIDTH: 2200,
    SUPPORT1_FINALISTS_PER_PRINCIPAL: 14,
    BEAM_MAX_CANDIDATES_PER_TYPE: 20,
    MODE_MAX_RUNTIME_MS: 90000,

    /* Prioridad de los jugadores más valiosos como Principales. */
    PREFERRED_PRINCIPAL_COUNT: 8,
    MAX_MANDATORY_PRINCIPALS: 6,

    /*
     * Preferencias de rol. IMPORTANTE:
     * - No se usan como restricciones duras de factibilidad.
     * - El algoritmo intenta mantener jugadores fuertes fuera de A2 y
     *   mantener Principal Meta >= Apoyos, pero nunca debe quedarse sin
     *   solución por estas preferencias.
     */
    PROTECTED_PLAYERS_FROM_SUPPORT2: 12,
    ENFORCE_SUPPORT2_PROTECTION: false,
    ENFORCE_META_HIERARCHY: false,
    ENABLE_FEASIBILITY_RESCUE: true,
    SPECIALIST_TOP_RANK: 2,
    SPECIALIST_REPLACEMENT_GAP: 0.12,


    /*
     * Debilidad global del catálogo. Roster!P y el número de unidades
     * Meta ya reflejan la calidad aplicada por Meta!C; no se multiplica
     * de nuevo por el banner.
     */
    GLOBAL_WEAKNESS_POWER_WEIGHT: 0.70,
    GLOBAL_WEAKNESS_COUNT_WEIGHT: 0.30,
    GLOBAL_WEAKNESS_PRIORITY_BOOST: 1.80,
    GLOBAL_WEAKNESS_PRINCIPAL_WEIGHT: 2.20,
    GLOBAL_WEAKNESS_SUPPORT_WEIGHT: 1.20,
    CRITICAL_GLOBAL_WEAK_TYPE_COUNT: 3,

    /* Alternativas completas conservadas para auditoría. */
    ALTERNATIVE_SOLUTIONS_TO_KEEP: 8,
    ALTERNATIVE_TYPES_TO_SHOW: 3,

    /* Optimización final, siempre dentro del mismo rol. */
    LOCAL_OPTIMIZATION_MAX_PASSES: 20,
    LOCAL_OPTIMIZATION_ENABLE_THREE_CYCLES: true,

    /* Bloqueo estricto de grupos. */
    LOCK_ROLE_GROUPS: true,
    ALLOW_PRINCIPAL_TO_SUPPORT2: false,
    ALLOW_SUPPORT1_TO_SUPPORT2: false,
    ALLOW_SUPPORT2_POOL_REPLACEMENT: false,
    ALLOW_SUPPORT1_PRINCIPAL_PROMOTION: false,

    /*
     * Un Principal por debajo de este porcentaje del máximo individual
     * sigue siendo válido, pero se penaliza como desempate y se avisa.
     */
    PRINCIPAL_LOW_STRENGTH_THRESHOLD: 0.30,

    /*
     * Confianza aproximada del roster:
     * unidades registradas por el jugador / mayor roster registrado.
     *
     * Solo se usa como último desempate y para mostrar avisos.
     */
    ROSTER_CONFIDENCE_WARNING_THRESHOLD: 0.60,

    // Seteadores.
    SETTERS_PER_TYPE: 3,
    TOP_PLAYERS_EXCLUDED_FROM_SETTERS: 8,
    MAX_SETTER_1_ASSIGNMENTS: 3,

    NORMALIZATION_SCALE: 1000000,
    COMPARISON_EPSILON: 1e-9
  };

  const spreadsheet = obtenerSpreadsheet_(CONFIG);

  /*
   * Guarda la hoja desde la que se pulsó el botón.
   * Al finalizar se vuelve a activar esta misma hoja.
   */
  const originSheet = spreadsheet.getActiveSheet();

  const selectedAlgorithm =
    prepararYLeerSelectorAlgoritmo_(
      spreadsheet,
      CONFIG
    );

  CONFIG.ALGORITHM_MODE =
    selectedAlgorithm.id;

  CONFIG.SELECTED_ALGORITHM_LABEL =
    selectedAlgorithm.label;

  aplicarPerfilAlgoritmo_(
    CONFIG,
    selectedAlgorithm.id
  );

  const guideSheet = spreadsheet.getSheetByName(CONFIG.GUIDE_SHEET);
  const membersSheet = spreadsheet.getSheetByName(CONFIG.MEMBERS_SHEET);
  const rosterSheet = spreadsheet.getSheetByName(CONFIG.ROSTER_SHEET);
  const metaSheet = spreadsheet.getSheetByName(CONFIG.META_SHEET);
  const syncPairsSheet = spreadsheet.getSheetByName(CONFIG.SYNC_PAIRS_SHEET);

  if (!guideSheet) {
    throw new Error(`No se ha encontrado la hoja "${CONFIG.GUIDE_SHEET}".`);
  }
  if (!membersSheet) {
    throw new Error(`No se ha encontrado la hoja "${CONFIG.MEMBERS_SHEET}".`);
  }
  if (!rosterSheet) {
    throw new Error(`No se ha encontrado la hoja "${CONFIG.ROSTER_SHEET}".`);
  }
  if (!metaSheet) {
    throw new Error(`No se ha encontrado la hoja "${CONFIG.META_SHEET}".`);
  }
  if (!syncPairsSheet) {
    throw new Error(`No se ha encontrado la hoja "${CONFIG.SYNC_PAIRS_SHEET}".`);
  }

  SpreadsheetApp.flush();

  const players = obtenerJugadores_(membersSheet, CONFIG);
  validarCapacidadInicial_(players, CONFIG);

  const guideData = leerClasificaciones_(
    guideSheet,
    players,
    CONFIG
  );

  const metaData = leerDatosMeta_(
    metaSheet,
    syncPairsSheet,
    rosterSheet,
    guideData.types,
    CONFIG
  );

  const rosterData = leerDatosRoster_(
    rosterSheet,
    players,
    guideData.types,
    metaData,
    CONFIG
  );

  const globalDataset = crearDatasetModo_(
    "GLOBAL",
    "Fuerza global",
    rosterData.globalScores,
    rosterData.globalMetaScores,
    metaData.maximumGlobalScoreByType,
    metaData.globalMetaCountByType,
    metaData.globalBannerSummaryByType,
    guideData.types,
    rosterData.globalEligibility,
    rosterData.exZoneEligibility,
    rosterData.rosterConfidenceByPlayer,
    rosterData.registeredUnitCountByPlayer,
    CONFIG
  );

  const physicalDataset = crearDatasetModo_(
    "PHYSICAL",
    "Fuerza física",
    rosterData.physicalScores,
    rosterData.physicalMetaScores,
    metaData.maximumPhysicalScoreByType,
    metaData.physicalMetaCountByType,
    metaData.physicalBannerSummaryByType,
    guideData.types,
    rosterData.physicalEligibility,
    combinarElegibilidadPrincipal_(
      rosterData.physicalEligibility,
      rosterData.exZoneEligibility
    ),
    rosterData.rosterConfidenceByPlayer,
    rosterData.registeredUnitCountByPlayer,
    CONFIG
  );

  const specialDataset = crearDatasetModo_(
    "SPECIAL",
    "Fuerza especial",
    rosterData.specialScores,
    rosterData.specialMetaScores,
    metaData.maximumSpecialScoreByType,
    metaData.specialMetaCountByType,
    metaData.specialBannerSummaryByType,
    guideData.types,
    rosterData.specialEligibility,
    combinarElegibilidadPrincipal_(
      rosterData.specialEligibility,
      rosterData.exZoneEligibility
    ),
    rosterData.rosterConfidenceByPlayer,
    rosterData.registeredUnitCountByPlayer,
    CONFIG
  );

  const globalResult = resolverModoCompleto_(
    globalDataset,
    players,
    CONFIG
  );

  const physicalResult = resolverModoCompleto_(
    physicalDataset,
    players,
    CONFIG
  );

  const specialResult = resolverModoCompleto_(
    specialDataset,
    players,
    CONFIG
  );

  /*
   * Los seteadores se siguen calculando respecto
   * a la tabla global.
   */
  const setterAssignments = generarAsignacionSetters_(
    players,
    globalDataset.types,
    globalDataset.normalizedScores,
    globalResult.solution,
    rosterData.rebuffUnitsByType,
    CONFIG
  );

  escribirResultadoCompleto_(
    spreadsheet,
    players,
    globalResult,
    physicalResult,
    specialResult,
    setterAssignments,
    CONFIG
  );

  const message =
    `Asignaciones generadas con "${CONFIG.SELECTED_ALGORITHM_LABEL}" ` +
    `en "${CONFIG.OUTPUT_SHEET}".`;

  SpreadsheetApp.flush();

  if (originSheet) {
    originSheet.activate();
  }

  SpreadsheetApp.flush();

  spreadsheet.toast(
    message,
    "Asignación óptima",
    8
  );

  console.log(message);
  return message;
}



/**
 * Crea la hoja de salida si no existe, prepara el desplegable y devuelve
 * el perfil seleccionado. La celda elegida se vuelve a escribir al final,
 * por lo que no se pierde al limpiar los resultados anteriores.
 */
function prepararYLeerSelectorAlgoritmo_(
  spreadsheet,
  config
) {
  let sheet =
    spreadsheet.getSheetByName(
      config.OUTPUT_SHEET
    );

  if (!sheet) {
    sheet =
      spreadsheet.insertSheet(
        config.OUTPUT_SHEET
      );
  }

  const selector =
    sheet.getRange(
      config.ALGORITHM_SELECTOR_CELL
    );

  const options =
    config.ALGORITHM_OPTIONS;

  const labels =
    options.map(
      option =>
        option.label
    );

  const defaultOption =
    options.find(
      option =>
        option.id ===
        config.DEFAULT_ALGORITHM_MODE
    ) || options[0];

  const currentValue =
    String(
      selector.getDisplayValue() || ""
    ).trim();

  const selected =
    options.find(
      option =>
        option.label ===
        currentValue
    ) || defaultOption;

  selector
    .setDataValidation(
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          labels,
          true
        )
        .setAllowInvalid(false)
        .setHelpText(
          "Selecciona la lógica y vuelve a ejecutar generarAsignacionOptima."
        )
        .build()
    )
    .setValue(
      selected.label
    );

  return selected;
}


/**
 * Activa un perfil de optimización. Todos usan el mismo lector de datos,
 * las mismas restricciones de roles y la misma salida; cambia el objetivo
 * de búsqueda y la forma de podar las combinaciones.
 */
function aplicarPerfilAlgoritmo_(
  config,
  mode
) {
  /* Valores comunes. */
  config.GLOBAL_WEAKNESS_POWER_WEIGHT = 0.70;
  config.GLOBAL_WEAKNESS_COUNT_WEIGHT = 0.30;
  config.GLOBAL_WEAKNESS_PRIORITY_BOOST = 1.80;
  config.CRITICAL_GLOBAL_WEAK_TYPE_COUNT = 3;

  config.PREFERRED_PRINCIPAL_COUNT = 8;
  config.MAX_MANDATORY_PRINCIPALS = 6;

  config.USE_GLOBAL_WEAKNESS = true;
  config.USE_PRINCIPAL_PRIORITY = true;
  config.USE_PREFERRED_PRINCIPALS = true;
  config.USE_MANDATORY_PRINCIPALS = true;

  switch (mode) {
    case "BALANCED_COVERAGE":
      config.USE_GLOBAL_WEAKNESS = false;
      config.USE_PRINCIPAL_PRIORITY = false;
      config.USE_PREFERRED_PRINCIPALS = false;
      config.USE_MANDATORY_PRINCIPALS = false;

      config.GLOBAL_WEAKNESS_PRIORITY_BOOST = 0;
      config.CRITICAL_GLOBAL_WEAK_TYPE_COUNT = 0;
      config.PREFERRED_PRINCIPAL_COUNT = 0;
      config.MAX_MANDATORY_PRINCIPALS = 0;
      break;

    case "STRONG_PRINCIPALS":
      config.USE_GLOBAL_WEAKNESS = false;
      config.USE_PRINCIPAL_PRIORITY = true;
      config.USE_PREFERRED_PRINCIPALS = true;
      config.USE_MANDATORY_PRINCIPALS = true;

      config.GLOBAL_WEAKNESS_PRIORITY_BOOST = 0;
      config.CRITICAL_GLOBAL_WEAK_TYPE_COUNT = 0;
      config.PREFERRED_PRINCIPAL_COUNT = 8;
      config.MAX_MANDATORY_PRINCIPALS = 6;
      break;

    case "GLOBAL_WEAKNESS":
      config.USE_GLOBAL_WEAKNESS = true;
      config.USE_PRINCIPAL_PRIORITY = true;
      config.USE_PREFERRED_PRINCIPALS = false;
      config.USE_MANDATORY_PRINCIPALS = false;

      config.PREFERRED_PRINCIPAL_COUNT = 0;
      config.MAX_MANDATORY_PRINCIPALS = 0;
      break;

    case "HYBRID_FULL":
    default:
      /* Mantiene todos los componentes activos. */
      break;
  }
}


function obtenerOpcionAlgoritmo_(
  config,
  mode
) {
  return (
    config.ALGORITHM_OPTIONS.find(
      option =>
        option.id === mode
    ) ||
    config.ALGORITHM_OPTIONS[0]
  );
}


function escribirPanelAlgoritmo_(
  sheet,
  config
) {
  const option =
    obtenerOpcionAlgoritmo_(
      config,
      config.ALGORITHM_MODE
    );

  sheet
    .getRange("A1:D1")
    .merge()
    .setValue(
      "CONFIGURACIÓN DE ASIGNACIÓN ÓPTIMA"
    )
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground("#1f4e78")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  sheet
    .getRange("A2")
    .setValue("ALGORITMO")
    .setFontWeight("bold")
    .setBackground("#d9eaf7");

  const selector =
    sheet.getRange(
      config.ALGORITHM_SELECTOR_CELL
    );

  selector
    .setDataValidation(
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          config.ALGORITHM_OPTIONS.map(
            item => item.label
          ),
          true
        )
        .setAllowInvalid(false)
        .setHelpText(
          "Selecciona el algoritmo y vuelve a ejecutar el botón."
        )
        .build()
    )
    .setValue(option.label)
    .setBackground("#fff2cc")
    .setFontWeight("bold");

  sheet
    .getRange("C2:D2")
    .merge()
    .setValue(option.description)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet
    .getRange("A3:D3")
    .merge()
    .setValue(
      "Cambia B2 y ejecuta de nuevo generarAsignacionOptima. " +
      "Las restricciones de EX, roles bloqueados y dos apariciones de Apoyo 2 se mantienen en todos los perfiles."
    )
    .setFontStyle("italic")
    .setFontColor("#666666")
    .setWrap(true);
}


/**
 * Resuelve una modalidad completa mediante tres fases:
 * Principal -> Apoyo 1 -> Apoyo 2.
 */
function resolverModoCompleto_(
  dataset,
  players,
  config
) {
  const allPlayerIndexes =
    players.map(
      (
        player,
        playerIndex
      ) =>
        playerIndex
    );

  let compactSolution = null;
  let datasetUsed = dataset;
  let configUsed = config;

  try {
    compactSolution =
      resolverAsignacionSecuencial_(
        datasetUsed,
        allPlayerIndexes,
        configUsed
      );
  } catch (error) {
    const errorMessage =
      String(
        error && error.message
          ? error.message
          : error
      );

    const isFeasibilityError =
      errorMessage.includes(
        "No se ha encontrado una combinación válida de Principales"
      ) ||
      errorMessage.includes(
        "No se ha encontrado ninguna solución completa válida"
      );

    if (
      !config.ENABLE_FEASIBILITY_RESCUE ||
      !isFeasibilityError
    ) {
      throw error;
    }

    /*
     * RESCATE DE FACTIBILIDAD
     * -----------------------
     * Si un perfil estratégico se queda sin solución completa, se vuelve
     * a intentar sin Principales obligatorios y sin restricciones duras
     * de jerarquía/protección. Las preferencias siguen interviniendo en
     * los comparadores, pero ya no pueden hacer imposible la asignación.
     */
    console.warn(
      `${dataset.label}: activando rescate de factibilidad tras: ` +
      errorMessage
    );

    configUsed = {
      ...config,
      USE_MANDATORY_PRINCIPALS: false,
      MAX_MANDATORY_PRINCIPALS: 0,
      ENFORCE_SUPPORT2_PROTECTION: false,
      ENFORCE_META_HIERARCHY: false
    };

    datasetUsed = {
      ...dataset,
      mandatoryPrincipalPlayers: new Set(),
      protectedSupport2Players: new Set()
    };

    compactSolution =
      resolverAsignacionSecuencial_(
        datasetUsed,
        allPlayerIndexes,
        configUsed
      );

    compactSolution.feasibilityRescueUsed = true;

    if (compactSolution.multiStartSummary) {
      compactSolution.multiStartSummary.feasibilityRescueUsed = true;
    }
  }

  const solution =
    construirSolucionFinal_(
      compactSolution,
      players,
      datasetUsed.types,
      datasetUsed.rawScores,
      datasetUsed.normalizedScores
    );

  validarSolucionSecuencial_(
    solution,
    datasetUsed,
    players,
    configUsed
  );

  const alternatives =
    (compactSolution.alternativeSolutions || [])
      .map(
        alternative => ({
          solution:
            construirSolucionFinal_(
              alternative,
              players,
              datasetUsed.types,
              datasetUsed.rawScores,
              datasetUsed.normalizedScores
            ),

          evaluation:
            evaluarSolucionJerarquica_(
              alternative,
              datasetUsed,
              configUsed
            )
        })
      );

  console.log(
    `${datasetUsed.label} [${datasetUsed.algorithmMode}] - Principales: ` +
    compactSolution.principalsByType
      .map(
        playerIndex =>
          players[playerIndex]
      )
      .join(", ")
  );

  return {
    dataset: datasetUsed,
    solution,

    evaluation:
      evaluarSolucionJerarquica_(
        compactSolution,
        datasetUsed,
        configUsed
      ),

    stageEvaluations:
      compactSolution.stageEvaluations,

    alternatives,

    multiStartSummary:
      compactSolution.multiStartSummary || null
  };
}


function crearDatasetModo_(
  mode,
  label,
  rawScores,
  metaScores,
  maximumIndividualScores,
  metaCountByType,
  bannerSummaryByType,
  baseTypes,
  assignmentEligibility,
  principalEligibility,
  rosterConfidenceByPlayer,
  registeredUnitCountByPlayer,
  config
) {
  const effectiveAssignmentEligibility =
    assignmentEligibility.map(
      (
        playerEligibility,
        playerIndex
      ) =>
        playerEligibility.map(
          (
            isEligible,
            typeIndex
          ) =>
            isEligible &&
            rawScores[
              playerIndex
            ][typeIndex] > 0
        )
    );

  const effectivePrincipalEligibility =
    principalEligibility.map(
      (
        playerEligibility,
        playerIndex
      ) =>
        playerEligibility.map(
          (
            isEligible,
            typeIndex
          ) =>
            isEligible &&
            effectiveAssignmentEligibility[
              playerIndex
            ][typeIndex]
        )
    );

  const maximumCatalogPower =
    Math.max(
      1,
      ...maximumIndividualScores
    );

  const maximumMetaCount =
    Math.max(
      1,
      ...metaCountByType
    );

  const powerWeight =
    Number(
      config.GLOBAL_WEAKNESS_POWER_WEIGHT
    ) || 0.70;

  const countWeight =
    Number(
      config.GLOBAL_WEAKNESS_COUNT_WEIGHT
    ) || 0.30;

  const weightTotal =
    Math.max(
      0.000001,
      powerWeight +
      countWeight
    );

  const types =
    baseTypes.map(
      (
        baseType,
        typeIndex
      ) => {
        const maximumIndividualScore =
          maximumIndividualScores[
            typeIndex
          ];

        const metaUnitCount =
          metaCountByType[
            typeIndex
          ] || 0;

        const catalogPowerRelative =
          maximumIndividualScore /
          maximumCatalogPower;

        const metaCountRelative =
          metaUnitCount /
          maximumMetaCount;

        const globalWeakness =
          Math.max(
            0,
            Math.min(
              1,
              (
                powerWeight *
                  (
                    1 -
                    catalogPowerRelative
                  ) +
                countWeight *
                  (
                    1 -
                    metaCountRelative
                  )
              ) /
              weightTotal
            )
          );

        return {
          name:
            baseType.name,

          maximumIndividualScore,

          maximumAssignedTeamScore:
            maximumIndividualScore *
            config.PLAYERS_PER_TYPE,

          idealScore:
            maximumIndividualScore *
            config.PLAYERS_PER_TYPE,

          metaUnitCount,
          catalogPower:
            maximumIndividualScore,
          catalogPowerRelative,
          metaCountRelative,
          globalWeakness,

          optimizationWeakness:
            config.USE_GLOBAL_WEAKNESS
              ? globalWeakness
              : 0,

          weaknessWeight:
            1 +
            (
              config.USE_GLOBAL_WEAKNESS
                ? globalWeakness
                : 0
            ) *
            config.GLOBAL_WEAKNESS_PRIORITY_BOOST,

          bannerSummary:
            bannerSummaryByType[
              typeIndex
            ] ||
            "Sin datos"
        };
      }
    );

  const weaknessOrder =
    types
      .map(
        (
          type,
          typeIndex
        ) => ({
          typeIndex,
          globalWeakness:
            type.globalWeakness
        })
      )
      .sort(
        (
          first,
          second
        ) =>
          second.globalWeakness -
          first.globalWeakness
      );

  weaknessOrder.forEach(
    (
      item,
      rank
    ) => {
      types[
        item.typeIndex
      ].globalWeaknessRank =
        rank + 1;

      types[
        item.typeIndex
      ].isCriticalGlobalWeakType =
        config.USE_GLOBAL_WEAKNESS &&
        rank <
        config.CRITICAL_GLOBAL_WEAK_TYPE_COUNT;
    }
  );

  const errors = [];

  types.forEach(
    (
      type,
      typeIndex
    ) => {
      const eligiblePlayers =
        effectiveAssignmentEligibility.reduce(
          (
            total,
            playerEligibility
          ) =>
            total +
            (
              playerEligibility[
                typeIndex
              ]
                ? 1
                : 0
            ),
          0
        );

      const principalCandidates =
        effectivePrincipalEligibility.reduce(
          (
            total,
            playerEligibility
          ) =>
            total +
            (
              playerEligibility[
                typeIndex
              ]
                ? 1
                : 0
            ),
          0
        );

      if (
        eligiblePlayers <
        config.PLAYERS_PER_TYPE
      ) {
        errors.push(
          `${type.name}: solo hay ${eligiblePlayers} jugadores con ` +
          `puntuación válida en ${label}.`
        );
      }

      if (
        principalCandidates === 0
      ) {
        errors.push(
          `${type.name}: no hay ningún candidato a Principal con EX.`
        );
      }

      if (
        type.maximumIndividualScore <= 0
      ) {
        errors.push(
          `${type.name}: el máximo individual de ${label} es 0.`
        );
      }
    }
  );

  if (
    errors.length > 0
  ) {
    throw new Error(
      `No se puede generar la tabla "${label}":\n\n` +
      errors.join("\n")
    );
  }

  const normalizedScores =
    normalizarPuntuaciones_(
      rawScores,
      types,
      config
    );

  const normalizedMetaScores =
    normalizarPuntuaciones_(
      metaScores,
      types,
      config
    );

  const bestPrincipalCoverageByType =
    types.map(
      (
        type,
        typeIndex
      ) =>
        normalizedScores.reduce(
          (
            best,
            playerScores,
            playerIndex
          ) =>
            effectivePrincipalEligibility[
              playerIndex
            ][typeIndex]
              ? Math.max(
                  best,
                  playerScores[
                    typeIndex
                  ] /
                  config.NORMALIZATION_SCALE
                )
              : best,
          0
        )
    );

  const specialistData =
    calcularEspecialistas_(
      normalizedScores,
      normalizedMetaScores,
      effectivePrincipalEligibility,
      types,
      config
    );

  const playerRoleValues =
    calcularValoresGeneralesJugadores_(
      normalizedScores,
      normalizedMetaScores,
      effectivePrincipalEligibility,
      rosterConfidenceByPlayer,
      specialistData.scoreByPlayer,
      types,
      config
    );

  const principalPriorityByPlayer =
    Array(
      rawScores.length
    ).fill(0);

  playerRoleValues.forEach(
    value => {
      principalPriorityByPlayer[
        value.playerIndex
      ] =
        value.principalPriority;
    }
  );

  const preferredPrincipalPlayers =
    new Set(
      config.USE_PREFERRED_PRINCIPALS
        ? playerRoleValues
            .slice(
              0,
              Math.min(
                config.PREFERRED_PRINCIPAL_COUNT,
                playerRoleValues.length
              )
            )
            .map(
              value =>
                value.playerIndex
            )
        : []
    );

  const protectedSupport2Players =
    construirProteccionSupport2_(
      playerRoleValues,
      specialistData.protectedPlayers,
      rawScores.length,
      config
    );

  const dataset = {
    mode,
    label,
    algorithmMode:
      config.ALGORITHM_MODE,
    types,

    rawScores,
    metaScores,
    normalizedScores,
    normalizedMetaScores,

    assignmentEligibility:
      effectiveAssignmentEligibility,

    principalEligibility:
      effectivePrincipalEligibility,

    rosterConfidenceByPlayer:
      rosterConfidenceByPlayer.slice(),

    registeredUnitCountByPlayer:
      registeredUnitCountByPlayer.slice(),

    rosterConfidenceWarningThreshold:
      config.ROSTER_CONFIDENCE_WARNING_THRESHOLD,

    playerRoleValues,
    principalPriorityByPlayer,
    preferredPrincipalPlayers,
    protectedSupport2Players,
    specialistData,
    bestPrincipalCoverageByType,

    globalWeaknessOrder:
      weaknessOrder.map(
        item =>
          item.typeIndex
      )
  };

  dataset.mandatoryPrincipalPlayers =
    config.USE_MANDATORY_PRINCIPALS
      ? seleccionarPrincipalesObligatorios_(
          dataset,
          config
        )
      : new Set();

  return dataset;
}





function calcularValoresGeneralesJugadores_(
  normalizedScores,
  normalizedMetaScores,
  principalEligibility,
  rosterConfidenceByPlayer,
  specialistScoreByPlayer,
  types,
  config
) {
  const scale =
    config.NORMALIZATION_SCALE;

  return normalizedScores
    .map(
      (
        playerScores,
        playerIndex
      ) => {
        const bestTotal =
          playerScores
            .slice()
            .sort(
              (a, b) =>
                b - a
            )
            .slice(
              0,
              3
            )
            .reduce(
              (
                total,
                score
              ) =>
                total + score,
              0
            );

        const bestMeta =
          normalizedMetaScores[
            playerIndex
          ]
            .slice()
            .sort(
              (a, b) =>
                b - a
            )
            .slice(
              0,
              3
            )
            .reduce(
              (
                total,
                score
              ) =>
                total + score,
              0
            );

        const exTypeCount =
          principalEligibility[
            playerIndex
          ].reduce(
            (
              total,
              eligible
            ) =>
              total +
              (
                eligible
                  ? 1
                  : 0
              ),
            0
          );

        const confidence =
          rosterConfidenceByPlayer[
            playerIndex
          ] || 0;

        const specialistScore =
          specialistScoreByPlayer[
            playerIndex
          ] || 0;

        const weakTypePotential =
          types.reduce(
            (
              total,
              type,
              typeIndex
            ) =>
              total +
              (
                principalEligibility[
                  playerIndex
                ][typeIndex]
                  ? normalizedScores[
                      playerIndex
                    ][typeIndex] *
                    type.optimizationWeakness *
                    0.35
                  : 0
              ),
            0
          );

        const generalValue =
          bestMeta +
          bestTotal * 0.35 +
          exTypeCount * scale * 0.08 +
          confidence * scale * 0.03;

        const principalPriority =
          config.USE_PRINCIPAL_PRIORITY
            ? (
                generalValue +
                specialistScore +
                weakTypePotential
              )
            : 0;

        return {
          playerIndex,
          bestMeta,
          bestTotal,
          exTypeCount,
          confidence,
          specialistScore,
          weakTypePotential,
          generalValue,
          principalPriority
        };
      }
    )
    .sort(
      (
        first,
        second
      ) => {
        if (
          first.principalPriority !==
          second.principalPriority
        ) {
          return (
            second.principalPriority -
            first.principalPriority
          );
        }

        if (
          first.bestMeta !==
          second.bestMeta
        ) {
          return (
            second.bestMeta -
            first.bestMeta
          );
        }

        return (
          first.playerIndex -
          second.playerIndex
        );
      }
    );
}





function leerDatosMeta_(
  metaSheet,
  syncPairsSheet,
  rosterSheet,
  types,
  config
) {
  const typeIndexByName =
    new Map();

  types.forEach(
    (
      type,
      typeIndex
    ) => {
      typeIndexByName.set(
        normalizarTexto_(
          type.name
        ),
        typeIndex
      );
    }
  );

  const styleBySyncPair =
    leerEstilosSyncPairs_(
      syncPairsSheet,
      config
    );

  const rosterReferenceBySyncPair =
    leerReferenciasMaximasRoster_(
      rosterSheet,
      config
    );

  const metaPairsByType =
    Array.from(
      {
        length:
          types.length
      },
      () =>
        new Set()
    );

  const metaUnitsByType =
    Array.from(
      {
        length:
          types.length
      },
      () =>
        new Map()
    );

  const lastRow =
    metaSheet.getLastRow();

  if (
    lastRow >=
    config.META_START_ROW
  ) {
    const lastRequiredColumn =
      Math.max(
        config.META_SYNC_PAIR_COLUMN,
        config.META_TYPE_COLUMN
      );

    const values =
      metaSheet
        .getRange(
          config.META_START_ROW,
          1,
          lastRow -
            config.META_START_ROW +
            1,
          lastRequiredColumn
        )
        .getValues();

    values.forEach(row => {
      const syncPair =
        limpiarTexto_(
          row[
            config.META_SYNC_PAIR_COLUMN -
              1
          ]
        );

      const typeName =
        limpiarTexto_(
          row[
            config.META_TYPE_COLUMN -
              1
          ]
        );

      if (
        !syncPair ||
        !typeName
      ) {
        return;
      }

      const typeIndex =
        typeIndexByName.get(
          normalizarTexto_(
            typeName
          )
        );

      if (
        typeIndex == null
      ) {
        return;
      }

      const syncPairKey =
        normalizarTexto_(
          syncPair
        );

      const rosterReference =
        rosterReferenceBySyncPair.get(
          syncPairKey
        );

      const style =
        styleBySyncPair.get(
          syncPairKey
        ) ||
        (
          rosterReference
            ? rosterReference.style
            : ""
        );

      const banner =
        rosterReference &&
        rosterReference.banner
          ? rosterReference.banner
          : "Sin banner";

      /*
       * Roster!P ya contiene el máximo potencial de la unidad y la
       * fórmula de puntuación ya aplica Meta!C. No se multiplica otra vez.
       */
      const maximumUnitScore =
        rosterReference &&
        rosterReference.maximumPotentialScore > 0
          ? rosterReference.maximumPotentialScore
          : 0;

      const normalizedStyle =
        normalizarTexto_(
          style
        );

      const isPhysical =
        normalizedStyle ===
          normalizarTexto_(
            config.PHYSICAL_VALUE
          ) ||
        normalizedStyle ===
          normalizarTexto_(
            config.HYBRID_VALUE
          );

      const isSpecial =
        normalizedStyle ===
          normalizarTexto_(
            config.SPECIAL_VALUE
          ) ||
        normalizedStyle ===
          normalizarTexto_(
            config.HYBRID_VALUE
          );

      metaPairsByType[
        typeIndex
      ].add(
        syncPairKey
      );

      const previous =
        metaUnitsByType[
          typeIndex
        ].get(
          syncPairKey
        );

      if (
        !previous ||
        maximumUnitScore >
          previous.maximumUnitScore
      ) {
        metaUnitsByType[
          typeIndex
        ].set(
          syncPairKey,
          {
            syncPair,
            style:
              normalizedStyle,
            banner,
            isPhysical,
            isSpecial,
            maximumUnitScore
          }
        );
      }
    });
  }

  const globalMetaCountByType = [];
  const physicalMetaCountByType = [];
  const specialMetaCountByType = [];

  const maximumGlobalScoreByType = [];
  const maximumPhysicalScoreByType = [];
  const maximumSpecialScoreByType = [];

  const globalBannerSummaryByType = [];
  const physicalBannerSummaryByType = [];
  const specialBannerSummaryByType = [];

  const errors = [];

  metaUnitsByType.forEach(
    (
      unitMap,
      typeIndex
    ) => {
      const units =
        Array.from(
          unitMap.values()
        );

      const globalUnits =
        units;

      const physicalUnits =
        units.filter(
          unit =>
            unit.isPhysical
        );

      const specialUnits =
        units.filter(
          unit =>
            unit.isSpecial
        );

      globalMetaCountByType[
        typeIndex
      ] =
        globalUnits.length;

      physicalMetaCountByType[
        typeIndex
      ] =
        physicalUnits.length;

      specialMetaCountByType[
        typeIndex
      ] =
        specialUnits.length;

      maximumGlobalScoreByType[
        typeIndex
      ] =
        globalUnits.reduce(
          (
            total,
            unit
          ) =>
            total +
            unit.maximumUnitScore,
          0
        );

      maximumPhysicalScoreByType[
        typeIndex
      ] =
        physicalUnits.reduce(
          (
            total,
            unit
          ) =>
            total +
            unit.maximumUnitScore,
          0
        );

      maximumSpecialScoreByType[
        typeIndex
      ] =
        specialUnits.reduce(
          (
            total,
            unit
          ) =>
            total +
            unit.maximumUnitScore,
          0
        );

      globalBannerSummaryByType[
        typeIndex
      ] =
        resumirBannersUnidades_(
          globalUnits
        );

      physicalBannerSummaryByType[
        typeIndex
      ] =
        resumirBannersUnidades_(
          physicalUnits
        );

      specialBannerSummaryByType[
        typeIndex
      ] =
        resumirBannersUnidades_(
          specialUnits
        );

      if (
        globalUnits.length === 0
      ) {
        errors.push(
          `${types[typeIndex].name}: no hay unidades Meta.`
        );
      }

      if (
        physicalUnits.length === 0
      ) {
        errors.push(
          `${types[typeIndex].name}: no hay unidades Meta ` +
          `Physical/Hybrid según SyncPairs!D.`
        );
      }

      if (
        specialUnits.length === 0
      ) {
        errors.push(
          `${types[typeIndex].name}: no hay unidades Meta ` +
          `Special/Hybrid según SyncPairs!D.`
        );
      }

      units.forEach(unit => {
        if (
          !unit.style
        ) {
          errors.push(
            `${types[typeIndex].name} / ${unit.syncPair}: no se ha ` +
            `encontrado estilo en SyncPairs!D ni en Roster!F.`
          );
        }

        if (
          unit.maximumUnitScore <= 0
        ) {
          errors.push(
            `${types[typeIndex].name} / ${unit.syncPair}: no tiene ` +
            `puntuación máxima válida. Completa Roster!P.`
          );
        }
      });
    }
  );

  if (
    errors.length > 0
  ) {
    throw new Error(
      "No se pueden calcular los máximos teóricos por tipo:\n\n" +
      errors.slice(
        0,
        60
      ).join("\n")
    );
  }

  return {
    metaPairsByType,
    metaUnitsByType,

    /*
     * Fuente autoritativa de Physical / Special / Hybrid.
     * Se reutiliza también al leer Roster para evitar discrepancias
     * entre Roster!F y SyncPairs!D.
     */
    styleBySyncPair,

    globalMetaCountByType,
    physicalMetaCountByType,
    specialMetaCountByType,

    maximumGlobalScoreByType,
    maximumPhysicalScoreByType,
    maximumSpecialScoreByType,

    globalBannerSummaryByType,
    physicalBannerSummaryByType,
    specialBannerSummaryByType
  };
}





/**
 * Lee SyncPairs!A:D y devuelve el estilo de cada Sync Pair.
 */
function resumirBannersUnidades_(
  units
) {
  if (
    !units ||
    units.length === 0
  ) {
    return "Sin unidades";
  }

  const counts =
    new Map();

  units.forEach(unit => {
    const banner =
      limpiarTexto_(
        unit.banner
      ) ||
      "Sin banner";

    counts.set(
      banner,
      (
        counts.get(
          banner
        ) || 0
      ) + 1
    );
  });

  return Array.from(
    counts.entries()
  )
    .sort(
      (
        first,
        second
      ) => {
        if (
          first[1] !==
          second[1]
        ) {
          return (
            second[1] -
            first[1]
          );
        }

        return first[0].localeCompare(
          second[0]
        );
      }
    )
    .map(
      entry =>
        `${entry[1]}× ${entry[0]}`
    )
    .join(" · ");
}


function leerEstilosSyncPairs_(
  syncPairsSheet,
  config
) {
  const styleBySyncPair =
    new Map();

  const lastRow =
    syncPairsSheet.getLastRow();

  if (
    lastRow < 2
  ) {
    return styleBySyncPair;
  }

  const lastRequiredColumn =
    Math.max(
      config.SYNC_PAIRS_SYNC_PAIR_COLUMN,
      config.SYNC_PAIRS_STYLE_COLUMN
    );

  const values =
    syncPairsSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastRequiredColumn
      )
      .getDisplayValues();

  values.forEach(row => {
    const syncPair =
      limpiarTexto_(
        row[
          config.SYNC_PAIRS_SYNC_PAIR_COLUMN -
            1
        ]
      );

    const style =
      limpiarTexto_(
        row[
          config.SYNC_PAIRS_STYLE_COLUMN -
            1
        ]
      );

    if (
      !syncPair ||
      !style
    ) {
      return;
    }

    styleBySyncPair.set(
      normalizarTexto_(
        syncPair
      ),
      style
    );
  });

  return styleBySyncPair;
}


/**
 * Consolida por Sync Pair:
 *
 * - Mayor Max Potential Score de Roster!P.
 * - Estilo de Roster!F como respaldo.
 */
function leerReferenciasMaximasRoster_(
  rosterSheet,
  config
) {
  const referenceBySyncPair =
    new Map();

  const lastRow =
    rosterSheet.getLastRow();

  if (
    lastRow <
    config.ROSTER_START_ROW
  ) {
    return referenceBySyncPair;
  }

  const lastRequiredColumn =
    Math.max(
      config.ROSTER_SYNC_PAIR_COLUMN,
      config.ROSTER_STYLE_COLUMN,
      config.ROSTER_BANNER_COLUMN,
      config.ROSTER_MAX_SCORE_COLUMN
    );

  const values =
    rosterSheet
      .getRange(
        config.ROSTER_START_ROW,
        1,
        lastRow -
          config.ROSTER_START_ROW +
          1,
        lastRequiredColumn
      )
      .getValues();

  values.forEach(row => {
    const syncPair =
      limpiarTexto_(
        row[
          config.ROSTER_SYNC_PAIR_COLUMN -
            1
        ]
      );

    if (!syncPair) {
      return;
    }

    const key =
      normalizarTexto_(
        syncPair
      );

    const style =
      limpiarTexto_(
        row[
          config.ROSTER_STYLE_COLUMN -
            1
        ]
      );

    const banner =
      limpiarTexto_(
        row[
          config.ROSTER_BANNER_COLUMN -
            1
        ]
      );

    const maximumPotentialScore =
      Math.max(
        0,
        convertirNumero_(
          row[
            config.ROSTER_MAX_SCORE_COLUMN -
              1
          ]
        )
      );

    const previous =
      referenceBySyncPair.get(
        key
      ) || {
        style: "",
        banner: "",
        maximumPotentialScore: 0
      };

    referenceBySyncPair.set(
      key,
      {
        style:
          previous.style ||
          style,

        banner:
          previous.banner ||
          banner,

        maximumPotentialScore:
          Math.max(
            previous.maximumPotentialScore,
            maximumPotentialScore
          )
      }
    );
  });

  return referenceBySyncPair;
}





/**
 * Lee Roster y construye las tres matrices de potencia efectiva.
 *
 * Para cada jugador, tipo y modalidad:
 *
 * 1. Se identifican las unidades Meta.
 * 2. Se seleccionan como máximo K unidades Meta y cuentan al 100 %.
 * 3. Si faltan huecos hasta K, se rellenan con unidades no Meta de la
 *    modalidad priorizadas por Roster!H (banner fuerte -> banner débil).
 * 4. Dentro del mismo banner, se elige primero la mayor Roster!N.
 * 5. Las unidades no Meta seleccionadas cuentan al 100 % de Roster!N.
 *
 * Un mismo Sync Pair solo puede contribuir una vez por jugador y tipo.
 * Si aparece duplicado en Roster, se conserva su puntuación más alta.
 */
function leerDatosRoster_(
  rosterSheet,
  players,
  types,
  metaData,
  config
) {
  const playerIndexByName =
    new Map();

  const typeIndexByName =
    new Map();

  players.forEach(
    (
      player,
      playerIndex
    ) => {
      playerIndexByName.set(
        normalizarTexto_(
          player
        ),
        playerIndex
      );
    }
  );

  types.forEach(
    (
      type,
      typeIndex
    ) => {
      typeIndexByName.set(
        normalizarTexto_(
          type.name
        ),
        typeIndex
      );
    }
  );

  const globalEligibility =
    crearMatrizBooleana_(
      players.length,
      types.length,
      false
    );

  const physicalEligibility =
    crearMatrizBooleana_(
      players.length,
      types.length,
      false
    );

  const specialEligibility =
    crearMatrizBooleana_(
      players.length,
      types.length,
      false
    );

  const exZoneEligibility =
    crearMatrizBooleana_(
      players.length,
      types.length,
      false
    );

  /*
   * Cada celda contiene un Map:
   * syncPair normalizado -> unidad consolidada.
   */
  const unitsByPlayerAndType =
    Array.from(
      {
        length:
          players.length
      },
      () =>
        Array.from(
          {
            length:
              types.length
          },
          () =>
            new Map()
        )
    );

  const rebuffUnitsByType =
    Array.from(
      {
        length:
          types.length
      },
      () =>
        new Map()
    );

  /*
   * Unidades únicas registradas por miembro.
   * Se utiliza únicamente para estimar la confianza de los datos.
   */
  const uniqueUnitsByPlayer =
    Array.from(
      {
        length:
          players.length
      },
      () =>
        new Set()
    );

  const lastRow =
    rosterSheet.getLastRow();

  if (
    lastRow >=
    config.ROSTER_START_ROW
  ) {
    const lastRequiredColumn =
      Math.max(
        config.ROSTER_PLAYER_COLUMN,
        config.ROSTER_SYNC_PAIR_COLUMN,
        config.ROSTER_TYPE_COLUMN,
        config.ROSTER_STYLE_COLUMN,
        config.ROSTER_BANNER_COLUMN,
        config.ROSTER_EX_ZONE_COLUMN,
        config.ROSTER_REBUFF_COLUMN,
        config.ROSTER_SCORE_COLUMN
      );

    const values =
      rosterSheet
        .getRange(
          config.ROSTER_START_ROW,
          1,
          lastRow -
            config.ROSTER_START_ROW +
            1,
          lastRequiredColumn
        )
        .getValues();

    values.forEach(row => {
      const player =
        limpiarTexto_(
          row[
            config.ROSTER_PLAYER_COLUMN -
              1
          ]
        );

      const syncPair =
        limpiarTexto_(
          row[
            config.ROSTER_SYNC_PAIR_COLUMN -
              1
          ]
        );

      const typeName =
        limpiarTexto_(
          row[
            config.ROSTER_TYPE_COLUMN -
              1
          ]
        );

      if (
        !player ||
        !syncPair ||
        !typeName
      ) {
        return;
      }

      const playerIndex =
        playerIndexByName.get(
          normalizarTexto_(
            player
          )
        );

      const typeIndex =
        typeIndexByName.get(
          normalizarTexto_(
            typeName
          )
        );

      if (
        playerIndex == null ||
        typeIndex == null
      ) {
        return;
      }

      const syncPairKey =
        normalizarTexto_(
          syncPair
        );

      uniqueUnitsByPlayer[
        playerIndex
      ].add(
        syncPairKey
      );

      /*
       * SyncPairs!D es la fuente autoritativa para Physical/Special/Hybrid.
       * Roster!F se mantiene únicamente como respaldo si la unidad no existe
       * en SyncPairs.
       */
      const rosterStyle =
        limpiarTexto_(
          row[
            config.ROSTER_STYLE_COLUMN -
              1
          ]
        );

      const syncPairStyle =
        metaData.styleBySyncPair &&
        metaData.styleBySyncPair.has(
          syncPairKey
        )
          ? metaData.styleBySyncPair.get(
              syncPairKey
            )
          : "";

      const style =
        normalizarTexto_(
          syncPairStyle ||
          rosterStyle
        );

      const banner =
        limpiarTexto_(
          row[
            config.ROSTER_BANNER_COLUMN -
              1
          ]
        );

      const exZone =
        limpiarTexto_(
          row[
            config.ROSTER_EX_ZONE_COLUMN -
              1
          ]
        );

      const rebuff =
        limpiarTexto_(
          row[
            config.ROSTER_REBUFF_COLUMN -
              1
          ]
        );

      const unitScore =
        Math.max(
          0,
          convertirNumero_(
            row[
              config.ROSTER_SCORE_COLUMN -
                1
            ]
          )
        );

      const isPhysical =
        style ===
          normalizarTexto_(
            config.PHYSICAL_VALUE
          ) ||
        style ===
          normalizarTexto_(
            config.HYBRID_VALUE
          );

      const isSpecial =
        style ===
          normalizarTexto_(
            config.SPECIAL_VALUE
          ) ||
        style ===
          normalizarTexto_(
            config.HYBRID_VALUE
          );

      const isMeta =
        metaData
          .metaPairsByType[
            typeIndex
          ]
          .has(
            syncPairKey
          );

      globalEligibility[
        playerIndex
      ][typeIndex] = true;

      if (isPhysical) {
        physicalEligibility[
          playerIndex
        ][typeIndex] = true;
      }

      if (isSpecial) {
        specialEligibility[
          playerIndex
        ][typeIndex] = true;
      }

      if (
        esValorActivo_(
          exZone,
          config.YES_VALUE
        )
      ) {
        exZoneEligibility[
          playerIndex
        ][typeIndex] = true;
      }

      const unitMap =
        unitsByPlayerAndType[
          playerIndex
        ][typeIndex];

      const previousUnit =
        unitMap.get(
          syncPairKey
        );

      if (
        !previousUnit ||
        unitScore >
          previousUnit.score
      ) {
        unitMap.set(
          syncPairKey,
          {
            syncPair,
            score:
              unitScore,
            banner:
              banner ||
              (
                previousUnit
                  ? previousUnit.banner
                  : ""
              ),
            bannerPriority:
              obtenerPrioridadBanner_(
                banner ||
                (
                  previousUnit
                    ? previousUnit.banner
                    : ""
                ),
                config
              ),
            isMeta,
            isPhysical,
            isSpecial
          }
        );
      }

      if (
        esValorActivo_(
          rebuff,
          config.YES_VALUE
        )
      ) {
        const playersForType =
          rebuffUnitsByType[
            typeIndex
          ];

        if (
          !playersForType.has(
            playerIndex
          )
        ) {
          playersForType.set(
            playerIndex,
            new Map()
          );
        }

        playersForType
          .get(
            playerIndex
          )
          .set(
            syncPairKey,
            syncPair
          );
      }
    });
  }

  const globalScores =
    crearMatrizNumerica_(
      players.length,
      types.length,
      0
    );

  const physicalScores =
    crearMatrizNumerica_(
      players.length,
      types.length,
      0
    );

  const specialScores =
    crearMatrizNumerica_(
      players.length,
      types.length,
      0
    );

  const globalMetaScores =
    crearMatrizNumerica_(
      players.length,
      types.length,
      0
    );

  const physicalMetaScores =
    crearMatrizNumerica_(
      players.length,
      types.length,
      0
    );

  const specialMetaScores =
    crearMatrizNumerica_(
      players.length,
      types.length,
      0
    );

  const scoreBreakdown = {
    global:
      crearMatrizDeObjetos_(
        players.length,
        types.length
      ),

    physical:
      crearMatrizDeObjetos_(
        players.length,
        types.length
      ),

    special:
      crearMatrizDeObjetos_(
        players.length,
        types.length
      )
  };

  for (
    let playerIndex = 0;
    playerIndex <
      players.length;
    playerIndex++
  ) {
    for (
      let typeIndex = 0;
      typeIndex <
        types.length;
      typeIndex++
    ) {
      const units =
        Array.from(
          unitsByPlayerAndType[
            playerIndex
          ][typeIndex]
          .values()
        );

      const globalK =
        metaData
          .globalMetaCountByType[
            typeIndex
          ];

      const physicalK =
        metaData
          .physicalMetaCountByType[
            typeIndex
          ];

      const specialK =
        metaData
          .specialMetaCountByType[
            typeIndex
          ];

      const globalResult =
        calcularPuntuacionEfectiva_(
          units,
          globalK,
          () => true,
          config
        );

      const physicalResult =
        calcularPuntuacionEfectiva_(
          units,
          physicalK,
          unit =>
            unit.isPhysical,
          config
        );

      const specialResult =
        calcularPuntuacionEfectiva_(
          units,
          specialK,
          unit =>
            unit.isSpecial,
          config
        );

      globalScores[
        playerIndex
      ][typeIndex] =
        globalResult.totalScore;

      physicalScores[
        playerIndex
      ][typeIndex] =
        physicalResult.totalScore;

      specialScores[
        playerIndex
      ][typeIndex] =
        specialResult.totalScore;

      globalMetaScores[
        playerIndex
      ][typeIndex] =
        globalResult.metaScore;

      physicalMetaScores[
        playerIndex
      ][typeIndex] =
        physicalResult.metaScore;

      specialMetaScores[
        playerIndex
      ][typeIndex] =
        specialResult.metaScore;

      scoreBreakdown.global[
        playerIndex
      ][typeIndex] =
        globalResult;

      scoreBreakdown.physical[
        playerIndex
      ][typeIndex] =
        physicalResult;

      scoreBreakdown.special[
        playerIndex
      ][typeIndex] =
        specialResult;
    }
  }

  const registeredUnitCountByPlayer =
    uniqueUnitsByPlayer.map(
      unitSet =>
        unitSet.size
    );

  const maximumRegisteredUnits =
    Math.max(
      1,
      ...registeredUnitCountByPlayer
    );

  /*
   * No representa la colección real del jugador.
   * Solo mide cuánto ha rellenado respecto al roster más completo
   * registrado en el gimnasio.
   */
  const rosterConfidenceByPlayer =
    registeredUnitCountByPlayer.map(
      unitCount =>
        unitCount /
        maximumRegisteredUnits
    );

  return {
    globalEligibility,
    physicalEligibility,
    specialEligibility,
    exZoneEligibility,

    globalScores,
    physicalScores,
    specialScores,

    globalMetaScores,
    physicalMetaScores,
    specialMetaScores,

    scoreBreakdown,
    rebuffUnitsByType,

    registeredUnitCountByPlayer,
    rosterConfidenceByPlayer
  };
}



/**
 * Normaliza el texto de un banner de forma más robusta que normalizarTexto_:
 * además elimina acentos para que "Poké" y "Poke" sean equivalentes.
 */
function normalizarClaveBanner_(
  value
) {
  return limpiarTexto_(
    value
  )
    .toLowerCase()
    .normalize("NFD")
    .replace(
      /[\u0300-\u036f]/g,
      ""
    );
}


/**
 * Devuelve una prioridad numérica:
 * 0 = banner de mayor prioridad.
 * Cuanto mayor sea el número, menor prioridad.
 *
 * Los banners desconocidos o vacíos se colocan al final de forma segura.
 */
function obtenerPrioridadBanner_(
  banner,
  config
) {
  const priorityList =
    Array.isArray(
      config.NON_META_BANNER_PRIORITY
    )
      ? config.NON_META_BANNER_PRIORITY
      : [];

  const normalizedBanner =
    normalizarClaveBanner_(
      banner
    );

  if (!normalizedBanner) {
    return (
      priorityList.length +
      1000
    );
  }

  for (
    let index = 0;
    index < priorityList.length;
    index++
  ) {
    if (
      normalizarClaveBanner_(
        priorityList[index]
      ) ===
      normalizedBanner
    ) {
      return index;
    }
  }

  return (
    priorityList.length +
    100
  );
}


/**
 * Orden de relleno de unidades NO Meta:
 *
 * 1. Banner de mayor prioridad.
 * 2. Mayor Strength Score actual (Roster!N).
 * 3. Nombre de Sync Pair para desempate determinista.
 */
function compararUnidadesNoMeta_(
  first,
  second,
  config
) {
  const firstPriority =
    Number.isFinite(
      first.bannerPriority
    )
      ? first.bannerPriority
      : obtenerPrioridadBanner_(
          first.banner,
          config
        );

  const secondPriority =
    Number.isFinite(
      second.bannerPriority
    )
      ? second.bannerPriority
      : obtenerPrioridadBanner_(
          second.banner,
          config
        );

  if (
    firstPriority !==
    secondPriority
  ) {
    return (
      firstPriority -
      secondPriority
    );
  }

  if (
    first.score !==
    second.score
  ) {
    return (
      second.score -
      first.score
    );
  }

  return String(
    first.syncPair || ""
  ).localeCompare(
    String(
      second.syncPair || ""
    )
  );
}


/**
 * Calcula la potencia efectiva de un conjunto de unidades.
 *
 * - Meta: usa el 100 % de Roster!N y siempre tiene prioridad.
 * - No Meta: solo rellena huecos hasta K y usa también el 100 % de Roster!N.
 * - Las no Meta se eligen por prioridad de Roster!H:
 *     banner más fuerte -> banner más débil.
 * - Dentro del mismo banner gana la unidad con mayor Roster!N.
 * - No se aplica ningún multiplicador adicional.
 * - Máximo total de unidades usadas: K.
 */
function calcularPuntuacionEfectiva_(
  units,
  k,
  modeFilter,
  config
) {
  const validUnits =
    units.filter(
      modeFilter
    );

  const metaUnits =
    validUnits
      .filter(
        unit =>
          unit.isMeta
      )
      .sort(
        (
          first,
          second
        ) =>
          second.score -
          first.score
      )
      .slice(
        0,
        k
      );

  const missingSlots =
    Math.max(
      0,
      k -
        metaUnits.length
    );

  const nonMetaUnits =
    validUnits
      .filter(
        unit =>
          !unit.isMeta
      )
      .sort(
        (
          first,
          second
        ) =>
          compararUnidadesNoMeta_(
            first,
            second,
            config
          )
      )
      .slice(
        0,
        missingSlots
      );

  const metaScore =
    metaUnits.reduce(
      (
        total,
        unit
      ) =>
        total +
        unit.score,
      0
    );

  const nonMetaRawScore =
    nonMetaUnits.reduce(
      (
        total,
        unit
      ) =>
        total +
        unit.score,
      0
    );

  /*
   * Roster!N ya contiene la puntuación que debe aportar la unidad.
   * Por tanto, una unidad no Meta de relleno se suma directamente,
   * sin volver a multiplicarla por 0,40 ni por ningún otro valor.
   */
  const nonMetaScore =
    nonMetaRawScore;

  return {
    totalScore:
      metaScore +
      nonMetaScore,

    metaScore,
    nonMetaRawScore,
    nonMetaScore,

    metaCount:
      metaUnits.length,

    nonMetaCount:
      nonMetaUnits.length,

    slots:
      k,

    metaSyncPairs:
      metaUnits.map(
        unit =>
          unit.syncPair
      ),

    nonMetaSyncPairs:
      nonMetaUnits.map(
        unit =>
          unit.syncPair
      ),

    /*
     * Útil para auditoría: permite comprobar exactamente qué rellenos
     * se eligieron y por qué banner.
     */
    nonMetaDetails:
      nonMetaUnits.map(
        unit => ({
          syncPair:
            unit.syncPair,
          banner:
            unit.banner ||
            "Sin banner",
          bannerPriority:
            obtenerPrioridadBanner_(
              unit.banner,
              config
            ),
          score:
            unit.score
        })
      )
  };
}


function crearMatrizNumerica_(
  rows,
  columns,
  initialValue
) {
  return Array.from(
    {
      length:
        rows
    },
    () =>
      Array(
        columns
      ).fill(
        initialValue
      )
  );
}


function crearMatrizDeObjetos_(
  rows,
  columns
) {
  return Array.from(
    {
      length:
        rows
    },
    () =>
      Array(
        columns
      ).fill(null)
  );
}


/**
 * Principal físico/especial:
 * tiene unidades válidas del estilo Y alguna EX Zone del mismo tipo.
 * Pueden ser filas/unidades diferentes.
 */
function combinarElegibilidadPrincipal_(
  styleEligibility,
  exZoneEligibility
) {
  return styleEligibility.map((playerEligibility, playerIndex) =>
    playerEligibility.map((isEligible, typeIndex) =>
      isEligible && exZoneEligibility[playerIndex][typeIndex]
    )
  );
}


/**
 * Valida que los grupos fijos del modo puedan cubrir las plazas.
 */


/**
 * ================================================================
 * ASIGNACIÓN SECUENCIAL
 * ================================================================
 */
function resolverAsignacionSecuencial_(
  dataset,
  allPlayerIndexes,
  config
) {
  const deadline =
    Date.now() +
    Math.max(
      20000,
      config.MODE_MAX_RUNTIME_MS
    );

  const principalStates =
    generarPrincipalesBeam_(
      dataset,
      allPlayerIndexes,
      config,
      deadline
    );

  if (
    principalStates.length === 0
  ) {
    throw new Error(
      `No se ha encontrado una combinación válida de Principales ` +
      `para "${dataset.label}".`
    );
  }

  let bestSolution = null;
  let bestEvaluation = null;
  let topSolutions = [];

  let principalChecked = 0;
  let support1Checked = 0;
  let support2Checked = 0;

  for (
    const principalState
    of principalStates
  ) {
    if (
      Date.now() >= deadline
    ) {
      break;
    }

    principalChecked++;

    const support1States =
      generarSupport1Beam_(
        dataset,
        allPlayerIndexes,
        principalState,
        config,
        deadline
      );

    for (
      const support1State
      of support1States
    ) {
      if (
        Date.now() >= deadline
      ) {
        break;
      }

      support1Checked++;

      const support2Result =
        resolverSupport2Exacto_(
          dataset,
          allPlayerIndexes,
          principalState,
          support1State,
          config,
          deadline
        );

      support2Checked +=
        support2Result.patternsChecked;

      for (
        let candidateSolution
        of support2Result.solutions
      ) {
        candidateSolution =
          optimizarRolesBloqueados_(
            candidateSolution,
            dataset,
            config
          );

        if (
          !esSolucionCompactaValida_(
            candidateSolution,
            dataset,
            config
          )
        ) {
          continue;
        }

        const evaluation =
          evaluarSolucionJerarquica_(
            candidateSolution,
            dataset,
            config
          );

        topSolutions =
          insertarSolucionOrdenada_(
            topSolutions,
            candidateSolution,
            evaluation,
            dataset,
            config
          );

        if (
          !bestSolution ||
          esMejorEvaluacionJerarquica_(
            evaluation,
            bestEvaluation,
            config
          )
        ) {
          bestSolution =
            clonarSolucionCompactaCompleta_(
              candidateSolution
            );

          bestEvaluation =
            evaluation;
        }
      }
    }
  }

  if (!bestSolution) {
    throw new Error(
      `No se ha encontrado ninguna solución completa válida para ` +
      `"${dataset.label}" incluso tras aplicar las reglas de factibilidad.`
    );
  }

  bestSolution.alternativeSolutions =
    topSolutions
      .filter(
        item =>
          item.evaluation.signature !==
          bestEvaluation.signature
      )
      .slice(
        0,
        config.ALTERNATIVE_SOLUTIONS_TO_KEEP
      )
      .map(
        item =>
          clonarSolucionCompactaCompleta_(
            item.solution
          )
      );

  bestSolution.multiStartSummary = {
    feasibilityRescueUsed:
      Boolean(bestSolution.feasibilityRescueUsed),
    principalChecked,
    support1Checked,
    support2Checked,
    principalFinalists:
      principalStates.length,
    mandatoryPrincipals:
      Array.from(
        dataset.mandatoryPrincipalPlayers
      ).length,
    exactGlobalSeed:
      principalStates.some(
        state =>
          state.exactGlobalSeed
      )
  };

  bestSolution.greedySteps = {
    principal: [],
    support1: [],
    support2: []
  };

  bestSolution.stageEvaluations =
    crearEvaluacionesEtapas_(
      bestSolution,
      dataset,
      config
    );

  console.log(
    `${dataset.label}: ${principalChecked} grupos de Principales, ` +
    `${support1Checked} grupos de Apoyo 1 y ${support2Checked} ` +
    `repartos de Apoyo 2 evaluados.`
  );

  return bestSolution;
}



/**
 * Construye el grupo protegido frente a Apoyo 2 sin crear una restricción
 * imposible.
 *
 * Antes se añadían los TOP N jugadores y después TODOS los especialistas.
 * Si la unión superaba las plazas disponibles fuera de Apoyo 2, ninguna
 * solución podía ser válida. Ahora:
 *
 * 1. Se priorizan especialistas críticos.
 * 2. Se completa con los jugadores de mayor valor general.
 * 3. Se limita el total al número máximo de jugadores que pueden quedar
 *    fuera de Apoyo 2.
 */
function construirProteccionSupport2_(
  playerRoleValues,
  specialistProtectedPlayers,
  playerCount,
  config
) {
  const support2PlayerCount =
    config.NUMBER_OF_TYPES /
    config.SUPPORT2_ASSIGNMENTS_PER_PLAYER;

  const maximumProtected =
    Math.max(
      0,
      playerCount -
      support2PlayerCount
    );

  const baseProtectedCount =
    Math.max(
      0,
      Math.min(
        config.PROTECTED_PLAYERS_FROM_SUPPORT2,
        maximumProtected,
        playerRoleValues.length
      )
    );

  const rankByPlayer =
    new Map();

  playerRoleValues.forEach(
    (
      value,
      rank
    ) => {
      rankByPlayer.set(
        value.playerIndex,
        rank
      );
    }
  );

  const specialists =
    Array.from(
      specialistProtectedPlayers || []
    )
      .sort(
        (
          first,
          second
        ) =>
          (
            rankByPlayer.get(first) ??
            Number.MAX_SAFE_INTEGER
          ) -
          (
            rankByPlayer.get(second) ??
            Number.MAX_SAFE_INTEGER
          )
      );

  const selected =
    [];

  const selectedSet =
    new Set();

  function addPlayer(
    playerIndex
  ) {
    if (
      selected.length >=
        maximumProtected ||
      selectedSet.has(
        playerIndex
      )
    ) {
      return;
    }

    selected.push(
      playerIndex
    );
    selectedSet.add(
      playerIndex
    );
  }

  /*
   * Especialistas primero, porque son precisamente los jugadores cuya
   * pérdida en determinados tipos es más difícil de sustituir.
   */
  specialists.forEach(
    addPlayer
  );

  /*
   * Después se garantiza la protección base de los jugadores fuertes.
   * Si ya no quedan plazas protegibles, se detiene de forma segura.
   */
  playerRoleValues
    .slice(
      0,
      baseProtectedCount
    )
    .forEach(
      value =>
        addPlayer(
          value.playerIndex
        )
    );

  /*
   * Si los especialistas han ocupado pocas plazas, seguimos recorriendo
   * el ranking general hasta alcanzar como máximo el objetivo configurado,
   * sin superar nunca el límite matemático.
   */
  for (
    const value
    of playerRoleValues
  ) {
    if (
      selected.length >=
        Math.min(
          maximumProtected,
          Math.max(
            baseProtectedCount,
            specialists.length
          )
        )
    ) {
      break;
    }

    addPlayer(
      value.playerIndex
    );
  }

  return new Set(
    selected
  );
}


function calcularEspecialistas_(
  normalizedScores,
  normalizedMetaScores,
  principalEligibility,
  types,
  config
) {
  const playerCount =
    normalizedScores.length;

  const scoreByPlayer =
    Array(
      playerCount
    ).fill(0);

  const protectedPlayers =
    new Set();

  const weakTypeIndexes =
    config.USE_GLOBAL_WEAKNESS
      ? types
      .map(
        (
          type,
          typeIndex
        ) => ({
          typeIndex,
          weakness:
            type.optimizationWeakness
        })
      )
      .sort(
        (
          first,
          second
        ) =>
          second.weakness -
          first.weakness
      )
      .slice(
        0,
        config.CRITICAL_GLOBAL_WEAK_TYPE_COUNT
      )
      .map(
        item =>
          item.typeIndex
      )
      : [];

  for (
    let typeIndex = 0;
    typeIndex <
      normalizedScores[0].length;
    typeIndex++
  ) {
    const type =
      types[
        typeIndex
      ];

    const candidates =
      normalizedScores
        .map(
          (
            playerScores,
            playerIndex
          ) => ({
            playerIndex,
            total:
              playerScores[
                typeIndex
              ],
            meta:
              normalizedMetaScores[
                playerIndex
              ][typeIndex]
          })
        )
        .filter(
          item =>
            principalEligibility[
              item.playerIndex
            ][typeIndex]
        )
        .sort(
          (
            first,
            second
          ) => {
            if (
              first.total !==
              second.total
            ) {
              return (
                second.total -
                first.total
              );
            }

            return (
              second.meta -
              first.meta
            );
          }
        );

    candidates
      .slice(
        0,
        config.SPECIALIST_TOP_RANK
      )
      .forEach(
        (
          item,
          rank
        ) => {
          const rankBonus =
            (
              config.SPECIALIST_TOP_RANK -
              rank
            ) *
            config.NORMALIZATION_SCALE *
            0.45 *
            type.weaknessWeight;

          const weakBonus =
            weakTypeIndexes.includes(
              typeIndex
            )
              ? config.NORMALIZATION_SCALE *
                0.65 *
                type.weaknessWeight
              : 0;

          scoreByPlayer[
            item.playerIndex
          ] +=
            rankBonus +
            weakBonus;
        }
      );

    if (
      candidates.length > 0
    ) {
      const best =
        candidates[0];

      const second =
        candidates[1];

      const gap =
        second
          ? best.total -
            second.total
          : best.total;

      const scarce =
        candidates.length <= 2;

      const critical =
        type.isCriticalGlobalWeakType;

      if (
        scarce ||
        critical ||
        gap >=
          config.SPECIALIST_REPLACEMENT_GAP *
          config.NORMALIZATION_SCALE
      ) {
        protectedPlayers.add(
          best.playerIndex
        );

        scoreByPlayer[
          best.playerIndex
        ] +=
          gap *
            1.75 *
            type.weaknessWeight +
          config.NORMALIZATION_SCALE *
            0.55 *
            type.weaknessWeight;
      }
    }
  }

  return {
    scoreByPlayer,
    protectedPlayers,
    weakTypeIndexes
  };
}





function seleccionarPrincipalesObligatorios_(
  dataset,
  config
) {
  const mandatory = [];

  for (
    const playerValue
    of dataset.playerRoleValues
  ) {
    if (
      mandatory.length >=
      config.MAX_MANDATORY_PRINCIPALS
    ) {
      break;
    }

    const candidate =
      mandatory.concat([
        playerValue.playerIndex
      ]);

    if (
      existeAsignacionPrincipalesConObligatorios_(
        candidate,
        dataset
      )
    ) {
      mandatory.push(
        playerValue.playerIndex
      );
    }
  }

  return new Set(
    mandatory
  );
}


function existeAsignacionPrincipalesConObligatorios_(
  mandatoryPlayers,
  dataset
) {
  const allPlayers =
    dataset.rawScores.map(
      (
        row,
        playerIndex
      ) =>
        playerIndex
    );

  const mandatorySet =
    new Set(
      mandatoryPlayers
    );

  const orderedMandatory =
    mandatoryPlayers
      .slice()
      .sort(
        (
          firstPlayer,
          secondPlayer
        ) =>
          contarTiposElegibles_(
            dataset.principalEligibility[
              firstPlayer
            ]
          ) -
          contarTiposElegibles_(
            dataset.principalEligibility[
              secondPlayer
            ]
          )
      );

  function visit(
    mandatoryIndex,
    usedTypes
  ) {
    if (
      mandatoryIndex ===
      orderedMandatory.length
    ) {
      const remainingTypes =
        dataset.types
          .map(
            (
              type,
              typeIndex
            ) =>
              typeIndex
          )
          .filter(
            typeIndex =>
              !usedTypes.has(
                typeIndex
              )
          );

      const remainingPlayers =
        allPlayers.filter(
          playerIndex =>
            !mandatorySet.has(
              playerIndex
            )
        );

      return esFasePendienteFactible_(
        remainingPlayers,
        remainingTypes,
        () => 1,
        dataset.principalEligibility
      );
    }

    const playerIndex =
      orderedMandatory[
        mandatoryIndex
      ];

    for (
      let typeIndex = 0;
      typeIndex <
        dataset.types.length;
      typeIndex++
    ) {
      if (
        usedTypes.has(
          typeIndex
        ) ||
        !dataset.principalEligibility[
          playerIndex
        ][typeIndex]
      ) {
        continue;
      }

      usedTypes.add(
        typeIndex
      );

      if (
        visit(
          mandatoryIndex + 1,
          usedTypes
        )
      ) {
        return true;
      }

      usedTypes.delete(
        typeIndex
      );
    }

    return false;
  }

  return visit(
    0,
    new Set()
  );
}


function contarTiposElegibles_(
  eligibilityRow
) {
  return eligibilityRow.reduce(
    (
      total,
      eligible
    ) =>
      total +
      (
        eligible
          ? 1
          : 0
      ),
    0
  );
}


function generarSemillaPrincipalGlobal_(
  dataset,
  allPlayerIndexes,
  config
) {
  const capacities =
    allPlayerIndexes.map(
      () => 1
    );

  const result =
    resolverAsignacionFlujo_(
      allPlayerIndexes,
      capacities,
      dataset.types.length,
      (
        playerIndex,
        typeIndex
      ) =>
        dataset.principalEligibility[
          playerIndex
        ][typeIndex],
      (
        playerIndex,
        typeIndex
      ) => {
        const type =
          dataset.types[
            typeIndex
          ];

        const coverage =
          coberturaIndividualTotal_(
            dataset,
            playerIndex,
            typeIndex
          );

        const meta =
          coberturaIndividualMeta_(
            dataset,
            playerIndex,
            typeIndex
          );

        const regret =
          Math.max(
            0,
            dataset.bestPrincipalCoverageByType[
              typeIndex
            ] -
            coverage
          );

        const mandatoryBonus =
          dataset.mandatoryPrincipalPlayers.has(
            playerIndex
          )
            ? 10000000000
            : 0;

        const preferredBonus =
          dataset.preferredPrincipalPlayers.has(
            playerIndex
          )
            ? 100000000
            : 0;

        return (
          mandatoryBonus +
          preferredBonus +
          coverage *
            type.weaknessWeight *
            10000000 +
          meta *
            type.optimizationWeakness *
            2500000 +
          dataset.principalPriorityByPlayer[
            playerIndex
          ] *
            0.01 -
          regret *
            type.weaknessWeight *
            12000000
        );
      }
    );

  if (!result) {
    return null;
  }

  const assignments =
    result.byType.slice();

  const usedPlayers =
    new Set(
      assignments
    );

  if (
    !Array.from(
      dataset.mandatoryPrincipalPlayers
    ).every(
      playerIndex =>
        usedPlayers.has(
          playerIndex
        )
    )
  ) {
    return null;
  }

  const state = {
    assignments,
    usedPlayers,

    principalPrioritySum:
      assignments.reduce(
        (
          total,
          playerIndex
        ) =>
          total +
          dataset.principalPriorityByPlayer[
            playerIndex
          ],
        0
      ),

    opportunityCost: 0,

    preferredCount:
      assignments.filter(
        playerIndex =>
          dataset.preferredPrincipalPlayers.has(
            playerIndex
          )
      ).length,

    exactGlobalSeed: true
  };

  state.principalCoverageProjection =
    assignments.map(
      (
        playerIndex,
        typeIndex
      ) =>
        coberturaIndividualTotal_(
          dataset,
          playerIndex,
          typeIndex
        )
    );

  state.weightedPrincipalDeficitProjection =
    crearVectorDeficitPonderado_(
      state.principalCoverageProjection,
      dataset
    );

  state.supportProjection =
    proyectarApoyo1Unico_(
      state,
      dataset,
      allPlayerIndexes,
      config
    );

  if (
    !state.supportProjection
  ) {
    return null;
  }

  state.weightedPairDeficitProjection =
    crearVectorDeficitPonderado_(
      state.supportProjection.pairCoverages,
      dataset
    );

  return state;
}


function generarPrincipalesBeam_(
  dataset,
  allPlayerIndexes,
  config,
  deadline
) {
  let beam = [{
    assignments:
      Array(
        dataset.types.length
      ).fill(null),

    usedPlayers:
      new Set(),

    principalPrioritySum: 0,
    opportunityCost: 0,
    preferredCount: 0
  }];

  for (
    let depth = 0;
    depth <
      dataset.types.length;
    depth++
  ) {
    if (
      Date.now() >= deadline
    ) {
      break;
    }

    const expanded = [];

    for (
      const state
      of beam
    ) {
      const typeIndex =
        seleccionarSiguienteTipoPrincipal_(
          state,
          dataset,
          allPlayerIndexes
        );

      if (
        typeIndex == null
      ) {
        continue;
      }

      const remainingMandatory =
        Array.from(
          dataset.mandatoryPrincipalPlayers
        ).filter(
          playerIndex =>
            !state.usedPlayers.has(
              playerIndex
            )
        );

      let candidates =
        allPlayerIndexes
          .filter(
            playerIndex =>
              !state.usedPlayers.has(
                playerIndex
              ) &&
              dataset.principalEligibility[
                playerIndex
              ][typeIndex]
          );

      const slotsAfter =
        dataset.types.length -
        depth -
        1;

      if (
        remainingMandatory.length >
        slotsAfter
      ) {
        candidates =
          candidates.filter(
            playerIndex =>
              dataset.mandatoryPrincipalPlayers.has(
                playerIndex
              )
          );
      }

      candidates.sort(
        (
          firstPlayer,
          secondPlayer
        ) =>
          compararCandidatosPrincipal_(
            firstPlayer,
            secondPlayer,
            typeIndex,
            state,
            dataset,
            allPlayerIndexes
          )
      );

      candidates =
        candidates.slice(
          0,
          config.BEAM_MAX_CANDIDATES_PER_TYPE
        );

      for (
        const playerIndex
        of candidates
      ) {
        const assignments =
          state.assignments.slice();

        assignments[
          typeIndex
        ] =
          playerIndex;

        const usedPlayers =
          new Set(
            state.usedPlayers
          );

        usedPlayers.add(
          playerIndex
        );

        const remainingTypes =
          dataset.types
            .map(
              (
                type,
                index
              ) =>
                index
            )
            .filter(
              index =>
                assignments[
                  index
                ] == null
            );

        const remainingPlayers =
          allPlayerIndexes.filter(
            index =>
              !usedPlayers.has(
                index
              )
          );

        if (
          !esEstadoPrincipalFactible_(
            assignments,
            usedPlayers,
            remainingPlayers,
            remainingTypes,
            dataset
          )
        ) {
          continue;
        }

        const opportunityCost =
          state.opportunityCost +
          calcularCosteOportunidadDinamico_(
            playerIndex,
            typeIndex,
            remainingPlayers,
            remainingTypes,
            dataset,
            dataset.principalEligibility
          );

        expanded.push({
          assignments,
          usedPlayers,

          principalPrioritySum:
            state.principalPrioritySum +
            dataset.principalPriorityByPlayer[
              playerIndex
            ],

          opportunityCost,

          preferredCount:
            state.preferredCount +
            (
              dataset.preferredPrincipalPlayers.has(
                playerIndex
              )
                ? 1
                : 0
            )
        });
      }
    }

    if (
      expanded.length === 0
    ) {
      return [];
    }

    expanded.forEach(
      state => {
        state.principalCoverageProjection =
          proyectarCoberturasPrincipales_(
            state,
            dataset,
            allPlayerIndexes
          );

        state.weightedPrincipalDeficitProjection =
          crearVectorDeficitPonderado_(
            state.principalCoverageProjection,
            dataset
          );
      }
    );

    expanded.sort(
      compararEstadosPrincipalesParciales_
    );

    beam =
      seleccionarEstadosDiversos_(
        expanded,
        config.PRINCIPAL_BEAM_WIDTH,
        state =>
          Array.from(
            state.usedPlayers
          )
            .sort(
              (a, b) =>
                a - b
            )
            .join(",")
      );
  }

  const complete =
    beam
      .filter(
        state =>
          state.assignments.every(
            value =>
              value != null
          ) &&
          Array.from(
            dataset.mandatoryPrincipalPlayers
          ).every(
            playerIndex =>
              state.usedPlayers.has(
                playerIndex
              )
          )
      );

  complete.forEach(
    state => {
      state.supportProjection =
        proyectarApoyo1Unico_(
          state,
          dataset,
          allPlayerIndexes,
          config
        );

      if (
        state.supportProjection
      ) {
        state.weightedPairDeficitProjection =
          crearVectorDeficitPonderado_(
            state.supportProjection.pairCoverages,
            dataset
          );
      }
    }
  );

  const exactSeed =
    generarSemillaPrincipalGlobal_(
      dataset,
      allPlayerIndexes,
      config
    );

  if (exactSeed) {
    const seedSignature =
      crearFirmaAsignaciones_(
        exactSeed.assignments
      );

    if (
      !complete.some(
        state =>
          crearFirmaAsignaciones_(
            state.assignments
          ) ===
          seedSignature
      )
    ) {
      complete.push(
        exactSeed
      );
    }
  }

  return complete
    .filter(
      state =>
        state.supportProjection != null
    )
    .sort(
      compararEstadosPrincipalesCompletos_
    )
    .slice(
      0,
      config.PRINCIPAL_FINALISTS
    );
}


function seleccionarSiguienteTipoPrincipal_(
  state,
  dataset,
  allPlayerIndexes
) {
  const unassignedTypes =
    dataset.types
      .map(
        (
          type,
          typeIndex
        ) =>
          typeIndex
      )
      .filter(
        typeIndex =>
          state.assignments[
            typeIndex
          ] == null
      );

  if (
    unassignedTypes.length === 0
  ) {
    return null;
  }

  return unassignedTypes
    .map(
      typeIndex => {
        const type =
          dataset.types[
            typeIndex
          ];

        const candidates =
          allPlayerIndexes.filter(
            playerIndex =>
              !state.usedPlayers.has(
                playerIndex
              ) &&
              dataset.principalEligibility[
                playerIndex
              ][typeIndex]
          );

        const mandatoryCandidates =
          candidates.filter(
            playerIndex =>
              dataset.mandatoryPrincipalPlayers.has(
                playerIndex
              )
          ).length;

        const bestCoverage =
          candidates.reduce(
            (
              best,
              playerIndex
            ) =>
              Math.max(
                best,
                coberturaIndividualTotal_(
                  dataset,
                  playerIndex,
                  typeIndex
                )
              ),
            0
          );

        const availableAverage =
          candidates.reduce(
            (
              total,
              playerIndex
            ) =>
              total +
              coberturaIndividualTotal_(
                dataset,
                playerIndex,
                typeIndex
              ),
            0
          ) /
          Math.max(
            1,
            candidates.length
          );

        const urgency =
          type.optimizationWeakness * 2.0 +
          (
            1 -
            bestCoverage
          ) * 0.45 +
          (
            1 -
            availableAverage
          ) * 0.30 +
          1 /
          Math.max(
            1,
            candidates.length
          );

        return {
          typeIndex,
          candidateCount:
            candidates.length,
          mandatoryCandidates,
          bestCoverage,
          globalWeakness:
            type.optimizationWeakness,
          critical:
            type.isCriticalGlobalWeakType,
          urgency
        };
      }
    )
    .sort(
      (
        first,
        second
      ) => {
        if (
          first.critical !==
          second.critical
        ) {
          return second.critical
            ? 1
            : -1;
        }

        if (
          first.urgency !==
          second.urgency
        ) {
          return (
            second.urgency -
            first.urgency
          );
        }

        if (
          first.candidateCount !==
          second.candidateCount
        ) {
          return (
            first.candidateCount -
            second.candidateCount
          );
        }

        if (
          first.mandatoryCandidates !==
          second.mandatoryCandidates
        ) {
          return (
            second.mandatoryCandidates -
            first.mandatoryCandidates
          );
        }

        return (
          first.typeIndex -
          second.typeIndex
        );
      }
    )[0]
    .typeIndex;
}





function compararCandidatosPrincipal_(
  firstPlayer,
  secondPlayer,
  typeIndex,
  state,
  dataset,
  allPlayerIndexes
) {
  const type =
    dataset.types[
      typeIndex
    ];

  const firstMandatory =
    dataset.mandatoryPrincipalPlayers.has(
      firstPlayer
    );

  const secondMandatory =
    dataset.mandatoryPrincipalPlayers.has(
      secondPlayer
    );

  if (
    firstMandatory !==
    secondMandatory
  ) {
    return secondMandatory
      ? 1
      : -1;
  }

  const firstCoverage =
    coberturaIndividualTotal_(
      dataset,
      firstPlayer,
      typeIndex
    );

  const secondCoverage =
    coberturaIndividualTotal_(
      dataset,
      secondPlayer,
      typeIndex
    );

  const bestAvailableCoverage =
    allPlayerIndexes.reduce(
      (
        best,
        playerIndex
      ) => {
        if (
          state.usedPlayers.has(
            playerIndex
          ) ||
          !dataset.principalEligibility[
            playerIndex
          ][typeIndex]
        ) {
          return best;
        }

        return Math.max(
          best,
          coberturaIndividualTotal_(
            dataset,
            playerIndex,
            typeIndex
          )
        );
      },
      0
    );

  const firstRegret =
    Math.max(
      0,
      bestAvailableCoverage -
      firstCoverage
    ) *
    type.weaknessWeight;

  const secondRegret =
    Math.max(
      0,
      bestAvailableCoverage -
      secondCoverage
    ) *
    type.weaknessWeight;

  if (
    firstRegret !==
    secondRegret
  ) {
    return (
      firstRegret -
      secondRegret
    );
  }

  const firstMeta =
    coberturaIndividualMeta_(
      dataset,
      firstPlayer,
      typeIndex
    );

  const secondMeta =
    coberturaIndividualMeta_(
      dataset,
      secondPlayer,
      typeIndex
    );

  const firstTypeUtility =
    firstCoverage *
      type.weaknessWeight +
    firstMeta *
      type.optimizationWeakness *
      0.35;

  const secondTypeUtility =
    secondCoverage *
      type.weaknessWeight +
    secondMeta *
      type.optimizationWeakness *
      0.35;

  if (
    firstTypeUtility !==
    secondTypeUtility
  ) {
    return (
      secondTypeUtility -
      firstTypeUtility
    );
  }

  const firstPreferred =
    dataset.preferredPrincipalPlayers.has(
      firstPlayer
    );

  const secondPreferred =
    dataset.preferredPrincipalPlayers.has(
      secondPlayer
    );

  if (
    firstPreferred !==
    secondPreferred
  ) {
    return secondPreferred
      ? 1
      : -1;
  }

  const priorityDifference =
    dataset.principalPriorityByPlayer[
      secondPlayer
    ] -
    dataset.principalPriorityByPlayer[
      firstPlayer
    ];

  if (
    priorityDifference !== 0
  ) {
    return priorityDifference;
  }

  return (
    firstPlayer -
    secondPlayer
  );
}





function esEstadoPrincipalFactible_(
  assignments,
  usedPlayers,
  remainingPlayers,
  remainingTypes,
  dataset
) {
  if (
    !esFasePendienteFactible_(
      remainingPlayers,
      remainingTypes,
      () => 1,
      dataset.principalEligibility
    )
  ) {
    return false;
  }

  const unassignedMandatory =
    Array.from(
      dataset.mandatoryPrincipalPlayers
    ).filter(
      playerIndex =>
        !usedPlayers.has(
          playerIndex
        )
    );

  if (
    unassignedMandatory.length >
    remainingTypes.length
  ) {
    return false;
  }

  return puedenUbicarseJugadoresEnTipos_(
    unassignedMandatory,
    remainingTypes,
    dataset.principalEligibility
  );
}


function puedenUbicarseJugadoresEnTipos_(
  players,
  types,
  eligibility
) {
  const orderedPlayers =
    players
      .slice()
      .sort(
        (
          firstPlayer,
          secondPlayer
        ) =>
          types.filter(
            typeIndex =>
              eligibility[
                firstPlayer
              ][typeIndex]
          ).length -
          types.filter(
            typeIndex =>
              eligibility[
                secondPlayer
              ][typeIndex]
          ).length
      );

  function visit(
    playerIndex,
    usedTypes
  ) {
    if (
      playerIndex ===
      orderedPlayers.length
    ) {
      return true;
    }

    const player =
      orderedPlayers[
        playerIndex
      ];

    for (
      const typeIndex
      of types
    ) {
      if (
        usedTypes.has(
          typeIndex
        ) ||
        !eligibility[
          player
        ][typeIndex]
      ) {
        continue;
      }

      usedTypes.add(
        typeIndex
      );

      if (
        visit(
          playerIndex + 1,
          usedTypes
        )
      ) {
        return true;
      }

      usedTypes.delete(
        typeIndex
      );
    }

    return false;
  }

  return visit(
    0,
    new Set()
  );
}


function proyectarCoberturasPrincipales_(
  state,
  dataset,
  allPlayerIndexes
) {
  return dataset.types.map(
    (
      type,
      typeIndex
    ) => {
      const assigned =
        state.assignments[
          typeIndex
        ];

      if (
        assigned != null
      ) {
        return coberturaIndividualTotal_(
          dataset,
          assigned,
          typeIndex
        );
      }

      return allPlayerIndexes.reduce(
        (
          best,
          playerIndex
        ) => {
          if (
            state.usedPlayers.has(
              playerIndex
            ) ||
            !dataset.principalEligibility[
              playerIndex
            ][typeIndex]
          ) {
            return best;
          }

          return Math.max(
            best,
            coberturaIndividualTotal_(
              dataset,
              playerIndex,
              typeIndex
            )
          );
        },
        0
      );
    }
  );
}


function proyectarApoyo1Unico_(
  principalState,
  dataset,
  allPlayerIndexes,
  config
) {
  const candidates =
    allPlayerIndexes.filter(
      playerIndex =>
        !principalState.usedPlayers.has(
          playerIndex
        )
    );

  const forcedSupportPlayers =
    candidates.filter(
      playerIndex =>
        dataset.protectedSupport2Players.has(
          playerIndex
        )
    );

  if (
    config.ENFORCE_SUPPORT2_PROTECTION &&
    forcedSupportPlayers.length >
    dataset.types.length
  ) {
    return null;
  }

  const flow =
    resolverAsignacionFlujo_(
      candidates,
      candidates.map(
        () => 1
      ),
      dataset.types.length,
      (
        playerIndex,
        typeIndex
      ) =>
        esApoyoValidoParaPrincipal_(
          playerIndex,
          principalState.assignments[
            typeIndex
          ],
          typeIndex,
          dataset,
          config
        ),
      (
        playerIndex,
        typeIndex
      ) => {
        const forcedBonus =
          dataset.protectedSupport2Players.has(
            playerIndex
          )
            ? 1000000000000
            : 0;

        return (
          forcedBonus +
          Math.round(
            coberturaParejaTotal_(
              dataset,
              principalState.assignments[
                typeIndex
              ],
              playerIndex,
              typeIndex
            ) *
            100000000
          )
        );
      }
    );

  if (!flow) {
    return null;
  }

  const usedSupport =
    new Set(
      flow.byType
    );

  const remaining =
    candidates.filter(
      playerIndex =>
        !usedSupport.has(
          playerIndex
        )
    );

  if (
    remaining.length !==
      dataset.types.length /
      config.SUPPORT2_ASSIGNMENTS_PER_PLAYER ||
    (
      config.ENFORCE_SUPPORT2_PROTECTION &&
      remaining.some(
        playerIndex =>
          dataset.protectedSupport2Players.has(
            playerIndex
          )
      )
    )
  ) {
    return null;
  }

  return {
    byType:
      flow.byType,

    pairCoverages:
      flow.byType.map(
        (
          playerIndex,
          typeIndex
        ) =>
          coberturaParejaTotal_(
            dataset,
            principalState.assignments[
              typeIndex
            ],
            playerIndex,
            typeIndex
          )
      )
  };
}


function crearVectorDeficitPonderado_(
  coverages,
  dataset
) {
  return coverages
    .map(
      (
        coverage,
        typeIndex
      ) =>
        Math.max(
          0,
          1 -
          coverage
        ) *
        dataset.types[
          typeIndex
        ].weaknessWeight
    )
    .sort(
      (
        first,
        second
      ) =>
        second -
        first
    );
}


function compararVectoresMenorEsMejor_(
  first,
  second
) {
  for (
    let index = 0;
    index <
      Math.min(
        first.length,
        second.length
      );
    index++
  ) {
    if (
      first[index] !==
      second[index]
    ) {
      return first[index] <
        second[index]
        ? 1
        : -1;
    }
  }

  return 0;
}


function compararEstadosPrincipalesParciales_(
  first,
  second
) {
  const deficitComparison =
    compararVectoresMenorEsMejor_(
      first.weightedPrincipalDeficitProjection,
      second.weightedPrincipalDeficitProjection
    );

  if (
    deficitComparison !== 0
  ) {
    return -deficitComparison;
  }

  if (
    first.preferredCount !==
    second.preferredCount
  ) {
    return (
      second.preferredCount -
      first.preferredCount
    );
  }

  if (
    first.principalPrioritySum !==
    second.principalPrioritySum
  ) {
    return (
      second.principalPrioritySum -
      first.principalPrioritySum
    );
  }

  const coverageComparison =
    compararVectoresCobertura_(
      first.principalCoverageProjection,
      second.principalCoverageProjection
    );

  if (
    coverageComparison !== 0
  ) {
    return -coverageComparison;
  }

  if (
    first.opportunityCost !==
    second.opportunityCost
  ) {
    return (
      first.opportunityCost -
      second.opportunityCost
    );
  }

  return crearFirmaAsignaciones_(
    first.assignments
  ).localeCompare(
    crearFirmaAsignaciones_(
      second.assignments
    )
  );
}





function compararEstadosPrincipalesCompletos_(
  first,
  second
) {
  const principalDeficitComparison =
    compararVectoresMenorEsMejor_(
      first.weightedPrincipalDeficitProjection,
      second.weightedPrincipalDeficitProjection
    );

  if (
    principalDeficitComparison !== 0
  ) {
    return -principalDeficitComparison;
  }

  if (
    first.preferredCount !==
    second.preferredCount
  ) {
    return (
      second.preferredCount -
      first.preferredCount
    );
  }

  if (
    first.principalPrioritySum !==
    second.principalPrioritySum
  ) {
    return (
      second.principalPrioritySum -
      first.principalPrioritySum
    );
  }

  const pairDeficitComparison =
    compararVectoresMenorEsMejor_(
      first.weightedPairDeficitProjection,
      second.weightedPairDeficitProjection
    );

  if (
    pairDeficitComparison !== 0
  ) {
    return -pairDeficitComparison;
  }

  const principalComparison =
    compararVectoresCobertura_(
      first.principalCoverageProjection,
      second.principalCoverageProjection
    );

  if (
    principalComparison !== 0
  ) {
    return -principalComparison;
  }

  return (
    first.opportunityCost -
    second.opportunityCost
  );
}





function generarSupport1Beam_(
  dataset,
  allPlayerIndexes,
  principalState,
  config,
  deadline
) {
  const availablePlayers =
    allPlayerIndexes.filter(
      playerIndex =>
        !principalState.usedPlayers.has(
          playerIndex
        )
    );

  let beam = [{
    assignments:
      Array(
        dataset.types.length
      ).fill(null),

    usedPlayers:
      new Set(),

    opportunityCost: 0
  }];

  for (
    let depth = 0;
    depth <
      dataset.types.length;
    depth++
  ) {
    if (
      Date.now() >= deadline
    ) {
      break;
    }

    const expanded = [];

    for (
      const state
      of beam
    ) {
      const typeIndex =
        seleccionarSiguienteTipoSupport1_(
          state,
          principalState,
          dataset,
          availablePlayers,
          config
        );

      if (
        typeIndex == null
      ) {
        continue;
      }

      let candidates =
        availablePlayers.filter(
          playerIndex =>
            !state.usedPlayers.has(
              playerIndex
            ) &&
            esApoyoValidoParaPrincipal_(
              playerIndex,
              principalState.assignments[
                typeIndex
              ],
              typeIndex,
              dataset,
              config
            )
        );

      candidates.sort(
        (
          firstPlayer,
          secondPlayer
        ) =>
          compararCandidatosSupport1_(
            firstPlayer,
            secondPlayer,
            typeIndex,
            principalState,
            state,
            dataset
          )
      );

      candidates =
        candidates.slice(
          0,
          config.BEAM_MAX_CANDIDATES_PER_TYPE
        );

      for (
        const playerIndex
        of candidates
      ) {
        const assignments =
          state.assignments.slice();

        assignments[
          typeIndex
        ] =
          playerIndex;

        const usedPlayers =
          new Set(
            state.usedPlayers
          );

        usedPlayers.add(
          playerIndex
        );

        const remainingTypes =
          dataset.types
            .map(
              (
                type,
                index
              ) =>
                index
            )
            .filter(
              index =>
                assignments[
                  index
                ] == null
            );

        const remainingPlayers =
          availablePlayers.filter(
            index =>
              !usedPlayers.has(
                index
              )
          );

        const protectedRemaining =
          remainingPlayers.filter(
            index =>
              dataset.protectedSupport2Players.has(
                index
              )
          );

        if (
          config.ENFORCE_SUPPORT2_PROTECTION &&
          protectedRemaining.length >
          remainingTypes.length
        ) {
          continue;
        }

        if (
          !esFasePendienteFactible_(
            remainingPlayers,
            remainingTypes,
            () => 1,
            crearElegibilidadSupport1_(
              principalState.assignments,
              dataset,
              config
            )
          )
        ) {
          continue;
        }

        const opportunityCost =
          state.opportunityCost +
          calcularCosteOportunidadDinamico_(
            playerIndex,
            typeIndex,
            remainingPlayers,
            remainingTypes,
            dataset,
            crearElegibilidadSupport1_(
              principalState.assignments,
              dataset,
              config
            )
          );

        const nextState = {
          assignments,
          usedPlayers,
          opportunityCost
        };

        nextState.pairCoverageProjection =
          proyectarCoberturasPareja_(
            nextState,
            principalState,
            dataset,
            availablePlayers,
            config
          );

        nextState.weightedPairDeficitProjection =
          crearVectorDeficitPonderado_(
            nextState.pairCoverageProjection,
            dataset
          );

        expanded.push(
          nextState
        );
      }
    }

    if (
      expanded.length === 0
    ) {
      return [];
    }

    expanded.sort(
      compararEstadosSupport1_
    );

    beam =
      seleccionarEstadosDiversos_(
        expanded,
        config.SUPPORT1_BEAM_WIDTH,
        state =>
          Array.from(
            state.usedPlayers
          )
            .sort(
              (a, b) =>
                a - b
            )
            .join(",")
      );
  }

  return beam
    .filter(
      state => {
        if (
          !state.assignments.every(
            value =>
              value != null
          )
        ) {
          return false;
        }

        const remaining =
          availablePlayers.filter(
            playerIndex =>
              !state.usedPlayers.has(
                playerIndex
              )
          );

        return (
          remaining.length ===
            dataset.types.length /
            config.SUPPORT2_ASSIGNMENTS_PER_PLAYER &&
          (
            !config.ENFORCE_SUPPORT2_PROTECTION ||
            remaining.every(
              playerIndex =>
                !dataset.protectedSupport2Players.has(
                  playerIndex
                )
            )
          )
        );
      }
    )
    .sort(
      compararEstadosSupport1_
    )
    .slice(
      0,
      config.SUPPORT1_FINALISTS_PER_PRINCIPAL
    );
}


function seleccionarSiguienteTipoSupport1_(
  state,
  principalState,
  dataset,
  availablePlayers,
  config
) {
  const unassigned =
    dataset.types
      .map(
        (
          type,
          typeIndex
        ) =>
          typeIndex
      )
      .filter(
        typeIndex =>
          state.assignments[
            typeIndex
          ] == null
      );

  if (
    unassigned.length === 0
  ) {
    return null;
  }

  return unassigned
    .map(
      typeIndex => {
        const principal =
          principalState.assignments[
            typeIndex
          ];

        const candidates =
          availablePlayers.filter(
            playerIndex =>
              !state.usedPlayers.has(
                playerIndex
              ) &&
              esApoyoValidoParaPrincipal_(
                playerIndex,
                principal,
                typeIndex,
                dataset,
                config
              )
          );

        const principalCoverage =
          coberturaIndividualTotal_(
            dataset,
            principal,
            typeIndex
          );

        const type =
          dataset.types[
            typeIndex
          ];

        const urgency =
          (
            1 -
            principalCoverage
          ) *
          type.weaknessWeight +
          type.optimizationWeakness * 0.40 +
          1 /
          Math.max(
            1,
            candidates.length
          );

        return {
          typeIndex,
          candidateCount:
            candidates.length,
          principalCoverage,
          urgency
        };
      }
    )
    .sort(
      (
        first,
        second
      ) => {
        if (
          first.urgency !==
          second.urgency
        ) {
          return (
            second.urgency -
            first.urgency
          );
        }

        if (
          first.candidateCount !==
          second.candidateCount
        ) {
          return (
            first.candidateCount -
            second.candidateCount
          );
        }

        return (
          first.typeIndex -
          second.typeIndex
        );
      }
    )[0]
    .typeIndex;
}





function compararCandidatosSupport1_(
  firstPlayer,
  secondPlayer,
  typeIndex,
  principalState,
  supportState,
  dataset
) {
  const firstProtected =
    dataset.protectedSupport2Players.has(
      firstPlayer
    );

  const secondProtected =
    dataset.protectedSupport2Players.has(
      secondPlayer
    );

  if (
    firstProtected !==
    secondProtected
  ) {
    return secondProtected
      ? 1
      : -1;
  }

  const principal =
    principalState.assignments[
      typeIndex
    ];

  const type =
    dataset.types[
      typeIndex
    ];

  const firstPair =
    coberturaParejaTotal_(
      dataset,
      principal,
      firstPlayer,
      typeIndex
    );

  const secondPair =
    coberturaParejaTotal_(
      dataset,
      principal,
      secondPlayer,
      typeIndex
    );

  const firstWeighted =
    firstPair *
    type.weaknessWeight;

  const secondWeighted =
    secondPair *
    type.weaknessWeight;

  if (
    firstWeighted !==
    secondWeighted
  ) {
    return (
      secondWeighted -
      firstWeighted
    );
  }

  const metaDifference =
    dataset.metaScores[
      secondPlayer
    ][typeIndex] -
    dataset.metaScores[
      firstPlayer
    ][typeIndex];

  if (
    metaDifference !== 0
  ) {
    return metaDifference;
  }

  return (
    firstPlayer -
    secondPlayer
  );
}





function proyectarCoberturasPareja_(
  supportState,
  principalState,
  dataset,
  availablePlayers,
  config
) {
  return dataset.types.map(
    (
      type,
      typeIndex
    ) => {
      const assigned =
        supportState.assignments[
          typeIndex
        ];

      const principal =
        principalState.assignments[
          typeIndex
        ];

      if (
        assigned != null
      ) {
        return coberturaParejaTotal_(
          dataset,
          principal,
          assigned,
          typeIndex
        );
      }

      return availablePlayers.reduce(
        (
          best,
          playerIndex
        ) => {
          if (
            supportState.usedPlayers.has(
              playerIndex
            ) ||
            !esApoyoValidoParaPrincipal_(
              playerIndex,
              principal,
              typeIndex,
              dataset,
              config
            )
          ) {
            return best;
          }

          return Math.max(
            best,
            coberturaParejaTotal_(
              dataset,
              principal,
              playerIndex,
              typeIndex
            )
          );
        },
        0
      );
    }
  );
}


function compararEstadosSupport1_(
  first,
  second
) {
  const weightedComparison =
    compararVectoresMenorEsMejor_(
      first.weightedPairDeficitProjection,
      second.weightedPairDeficitProjection
    );

  if (
    weightedComparison !== 0
  ) {
    return -weightedComparison;
  }

  const comparison =
    compararVectoresCobertura_(
      first.pairCoverageProjection,
      second.pairCoverageProjection
    );

  if (
    comparison !== 0
  ) {
    return -comparison;
  }

  if (
    first.opportunityCost !==
    second.opportunityCost
  ) {
    return (
      first.opportunityCost -
      second.opportunityCost
    );
  }

  return crearFirmaAsignaciones_(
    first.assignments
  ).localeCompare(
    crearFirmaAsignaciones_(
      second.assignments
    )
  );
}





function resolverSupport2Exacto_(
  dataset,
  allPlayerIndexes,
  principalState,
  support1State,
  config,
  deadline
) {
  const used =
    new Set([
      ...principalState.usedPlayers,
      ...support1State.usedPlayers
    ]);

  const support2Players =
    allPlayerIndexes.filter(
      playerIndex =>
        !used.has(
          playerIndex
        )
    );

  const expectedSupport2Players =
    dataset.types.length /
    config.SUPPORT2_ASSIGNMENTS_PER_PLAYER;

  if (
    support2Players.length !==
      expectedSupport2Players ||
    (
      config.ENFORCE_SUPPORT2_PROTECTION &&
      support2Players.some(
        playerIndex =>
          dataset.protectedSupport2Players.has(
            playerIndex
          )
      )
    )
  ) {
    return {
      solutions: [],
      patternsChecked: 0
    };
  }

  const typeOrder =
    dataset.types
      .map(
        (
          type,
          typeIndex
        ) => ({
          typeIndex,
          pairCoverage:
            coberturaParejaTotal_(
              dataset,
              principalState.assignments[
                typeIndex
              ],
              support1State.assignments[
                typeIndex
              ],
              typeIndex
            )
        })
      )
      .sort(
        (
          first,
          second
        ) =>
          first.pairCoverage -
          second.pairCoverage
      )
      .map(
        item =>
          item.typeIndex
      );

  const counts =
    Array(
      support2Players.length
    ).fill(0);

  const support2ByType =
    Array(
      dataset.types.length
    ).fill(null);

  let patternsChecked = 0;
  let top = [];

  function visit(
    depth
  ) {
    /*
     * No se corta una enumeración de Apoyo 2 a mitad.
     *
     * El bucle exterior ya comprueba el deadline antes de llamar a esta
     * función. Una vez empezada una combinación Principal + Apoyo 1,
     * sus repartos de Apoyo 2 se evalúan completos para no favorecer
     * accidentalmente a los patrones visitados primero.
     */
    if (
      depth ===
      typeOrder.length
    ) {
      if (
        counts.some(
          count =>
            count !==
            config.SUPPORT2_ASSIGNMENTS_PER_PLAYER
        )
      ) {
        return;
      }

      patternsChecked++;

      const solution = {
        principalsByType:
          principalState.assignments.slice(),

        support1ByType:
          support1State.assignments.slice(),

        support2ByType:
          support2ByType.slice(),

        principalPrioritySum:
          principalState.principalPrioritySum,

        preferredPrincipalCount:
          principalState.preferredCount,

        opportunityCost:
          principalState.opportunityCost +
          support1State.opportunityCost
      };

      if (
        !esSolucionCompactaValida_(
          solution,
          dataset,
          config
        )
      ) {
        return;
      }

      const evaluation =
        evaluarSolucionJerarquica_(
          solution,
          dataset,
          config
        );

      top =
        insertarSolucionOrdenada_(
          top,
          solution,
          evaluation,
          dataset,
          {
            ...config,
            ALTERNATIVE_SOLUTIONS_TO_KEEP: 3
          }
        );

      return;
    }

    const typeIndex =
      typeOrder[
        depth
      ];

    const principal =
      principalState.assignments[
        typeIndex
      ];

    for (
      let localIndex = 0;
      localIndex <
        support2Players.length;
      localIndex++
    ) {
      if (
        counts[
          localIndex
        ] >=
        config.SUPPORT2_ASSIGNMENTS_PER_PLAYER
      ) {
        continue;
      }

      const playerIndex =
        support2Players[
          localIndex
        ];

      if (
        !esApoyoValidoParaPrincipal_(
          playerIndex,
          principal,
          typeIndex,
          dataset,
          config
        )
      ) {
        continue;
      }

      counts[
        localIndex
      ]++;

      support2ByType[
        typeIndex
      ] =
        playerIndex;

      visit(
        depth + 1
      );

      support2ByType[
        typeIndex
      ] =
        null;

      counts[
        localIndex
      ]--;
    }
  }

  visit(0);

  return {
    solutions:
      top.map(
        item =>
          item.solution
      ),
    patternsChecked
  };
}


function optimizarRolesBloqueados_(
  originalSolution,
  dataset,
  config
) {
  const solution =
    clonarSolucionCompactaCompleta_(
      originalSolution
    );

  let currentEvaluation =
    evaluarSolucionJerarquica_(
      solution,
      dataset,
      config
    );

  for (
    let pass = 0;
    pass <
      config.LOCAL_OPTIMIZATION_MAX_PASSES;
    pass++
  ) {
    let bestMove = null;

    const roles = [
      {
        key:
          "principalsByType",
        eligibility:
          dataset.principalEligibility
      },
      {
        key:
          "support1ByType",
        eligibility:
          dataset.assignmentEligibility
      },
      {
        key:
          "support2ByType",
        eligibility:
          dataset.assignmentEligibility
      }
    ];

    for (
      const role
      of roles
    ) {
      const assignments =
        solution[
          role.key
        ];

      for (
        let firstType = 0;
        firstType <
          assignments.length - 1;
        firstType++
      ) {
        for (
          let secondType =
            firstType + 1;
          secondType <
            assignments.length;
          secondType++
        ) {
          const firstPlayer =
            assignments[
              firstType
            ];

          const secondPlayer =
            assignments[
              secondType
            ];

          if (
            firstPlayer ===
            secondPlayer ||
            !role.eligibility[
              firstPlayer
            ][secondType] ||
            !role.eligibility[
              secondPlayer
            ][firstType]
          ) {
            continue;
          }

          assignments[
            firstType
          ] =
            secondPlayer;

          assignments[
            secondType
          ] =
            firstPlayer;

          const valid =
            esSolucionCompactaValida_(
              solution,
              dataset,
              config
            );

          const evaluation =
            valid
              ? evaluarSolucionJerarquica_(
                  solution,
                  dataset,
                  config
                )
              : null;

          assignments[
            firstType
          ] =
            firstPlayer;

          assignments[
            secondType
          ] =
            secondPlayer;

          if (
            evaluation &&
            esMejorEvaluacionJerarquica_(
              evaluation,
              currentEvaluation,
              config
            ) &&
            (
              !bestMove ||
              esMejorEvaluacionJerarquica_(
                evaluation,
                bestMove.evaluation,
                config
              )
            )
          ) {
            bestMove = {
              roleKey:
                role.key,
              typeIndexes: [
                firstType,
                secondType
              ],
              players: [
                secondPlayer,
                firstPlayer
              ],
              evaluation
            };
          }
        }
      }

      if (
        config.LOCAL_OPTIMIZATION_ENABLE_THREE_CYCLES
      ) {
        for (
          let firstType = 0;
          firstType <
            assignments.length - 2;
          firstType++
        ) {
          for (
            let secondType =
              firstType + 1;
            secondType <
              assignments.length - 1;
            secondType++
          ) {
            for (
              let thirdType =
                secondType + 1;
              thirdType <
                assignments.length;
              thirdType++
            ) {
              const original = [
                assignments[
                  firstType
                ],
                assignments[
                  secondType
                ],
                assignments[
                  thirdType
                ]
              ];

              if (
                new Set(
                  original
                ).size < 3
              ) {
                continue;
              }

              const rotations = [
                [
                  original[2],
                  original[0],
                  original[1]
                ],
                [
                  original[1],
                  original[2],
                  original[0]
                ]
              ];

              for (
                const rotated
                of rotations
              ) {
                if (
                  !role.eligibility[
                    rotated[0]
                  ][firstType] ||
                  !role.eligibility[
                    rotated[1]
                  ][secondType] ||
                  !role.eligibility[
                    rotated[2]
                  ][thirdType]
                ) {
                  continue;
                }

                assignments[
                  firstType
                ] =
                  rotated[0];

                assignments[
                  secondType
                ] =
                  rotated[1];

                assignments[
                  thirdType
                ] =
                  rotated[2];

                const valid =
                  esSolucionCompactaValida_(
                    solution,
                    dataset,
                    config
                  );

                const evaluation =
                  valid
                    ? evaluarSolucionJerarquica_(
                        solution,
                        dataset,
                        config
                      )
                    : null;

                assignments[
                  firstType
                ] =
                  original[0];

                assignments[
                  secondType
                ] =
                  original[1];

                assignments[
                  thirdType
                ] =
                  original[2];

                if (
                  evaluation &&
                  esMejorEvaluacionJerarquica_(
                    evaluation,
                    currentEvaluation,
                    config
                  ) &&
                  (
                    !bestMove ||
                    esMejorEvaluacionJerarquica_(
                      evaluation,
                      bestMove.evaluation,
                      config
                    )
                  )
                ) {
                  bestMove = {
                    roleKey:
                      role.key,
                    typeIndexes: [
                      firstType,
                      secondType,
                      thirdType
                    ],
                    players:
                      rotated.slice(),
                    evaluation
                  };
                }
              }
            }
          }
        }
      }
    }

    if (!bestMove) {
      break;
    }

    bestMove.typeIndexes.forEach(
      (
        typeIndex,
        localIndex
      ) => {
        solution[
          bestMove.roleKey
        ][typeIndex] =
          bestMove.players[
            localIndex
          ];
      }
    );

    currentEvaluation =
      bestMove.evaluation;
  }

  return solution;
}


function esSolucionCompactaValida_(
  solution,
  dataset,
  config
) {
  const principalSet =
    new Set(
      solution.principalsByType
    );

  const support1Set =
    new Set(
      solution.support1ByType
    );

  const support2Count =
    new Map();

  if (
    principalSet.size !==
      config.PRINCIPAL_COUNT ||
    support1Set.size !==
      config.SUPPORT1_COUNT
  ) {
    return false;
  }

  for (
    const playerIndex
    of principalSet
  ) {
    if (
      support1Set.has(
        playerIndex
      )
    ) {
      return false;
    }
  }

  for (
    let typeIndex = 0;
    typeIndex <
      dataset.types.length;
    typeIndex++
  ) {
    const principal =
      solution.principalsByType[
        typeIndex
      ];

    const support1 =
      solution.support1ByType[
        typeIndex
      ];

    const support2 =
      solution.support2ByType[
        typeIndex
      ];

    if (
      principal == null ||
      support1 == null ||
      support2 == null ||
      !dataset.principalEligibility[
        principal
      ][typeIndex] ||
      !esApoyoValidoParaPrincipal_(
        support1,
        principal,
        typeIndex,
        dataset,
        config
      ) ||
      !esApoyoValidoParaPrincipal_(
        support2,
        principal,
        typeIndex,
        dataset,
        config
      )
    ) {
      return false;
    }

    if (
      principal === support1 ||
      principal === support2 ||
      support1 === support2
    ) {
      return false;
    }

    if (
      principalSet.has(
        support2
      ) ||
      support1Set.has(
        support2
      )
    ) {
      return false;
    }

    support2Count.set(
      support2,
      (
        support2Count.get(
          support2
        ) || 0
      ) + 1
    );
  }

  if (
    support2Count.size !==
    dataset.rawScores.length -
      config.PRINCIPAL_COUNT -
      config.SUPPORT1_COUNT
  ) {
    return false;
  }

  for (
    const count
    of support2Count.values()
  ) {
    if (
      count !==
      config.SUPPORT2_ASSIGNMENTS_PER_PLAYER
    ) {
      return false;
    }
  }

  if (
    Array.from(
      dataset.mandatoryPrincipalPlayers
    ).some(
      playerIndex =>
        !principalSet.has(
          playerIndex
        )
    )
  ) {
    return false;
  }

  return true;
}


function esApoyoValidoParaPrincipal_(
  supportPlayer,
  principalPlayer,
  typeIndex,
  dataset,
  config
) {
  if (
    !dataset.assignmentEligibility[
      supportPlayer
    ][typeIndex]
  ) {
    return false;
  }

  if (
    config &&
    config.ENFORCE_META_HIERARCHY
  ) {
    return (
      dataset.metaScores[
        supportPlayer
      ][typeIndex] <=
      dataset.metaScores[
        principalPlayer
      ][typeIndex]
    );
  }

  return true;
}


function crearElegibilidadSupport1_(
  principalsByType,
  dataset,
  config
) {
  return dataset.assignmentEligibility.map(
    (
      playerEligibility,
      playerIndex
    ) =>
      playerEligibility.map(
        (
          eligible,
          typeIndex
        ) => {
          if (!eligible) {
            return false;
          }

          if (
            config &&
            config.ENFORCE_META_HIERARCHY
          ) {
            return (
              dataset.metaScores[
                playerIndex
              ][typeIndex] <=
              dataset.metaScores[
                principalsByType[
                  typeIndex
                ]
              ][typeIndex]
            );
          }

          return true;
        }
      )
  );
}


function calcularCosteOportunidadDinamico_(
  playerIndex,
  assignedType,
  remainingPlayers,
  remainingTypes,
  dataset,
  eligibility
) {
  const assignedTypeInfo =
    dataset.types[
      assignedType
    ];

  const assignedCoverage =
    coberturaIndividualTotal_(
      dataset,
      playerIndex,
      assignedType
    );

  const bestAssignedAlternative =
    [
      playerIndex,
      ...remainingPlayers
    ].reduce(
      (
        best,
        candidate
      ) => {
        if (
          !eligibility[
            candidate
          ][assignedType]
        ) {
          return best;
        }

        return Math.max(
          best,
          coberturaIndividualTotal_(
            dataset,
            candidate,
            assignedType
          )
        );
      },
      0
    );

  let cost =
    Math.max(
      0,
      bestAssignedAlternative -
      assignedCoverage
    ) *
    assignedTypeInfo.weaknessWeight *
    dataset.types.length;

  for (
    const typeIndex
    of remainingTypes
  ) {
    if (
      !eligibility[
        playerIndex
      ][typeIndex]
    ) {
      continue;
    }

    const typeInfo =
      dataset.types[
        typeIndex
      ];

    const playerCoverage =
      coberturaIndividualTotal_(
        dataset,
        playerIndex,
        typeIndex
      );

    const bestAlternative =
      remainingPlayers.reduce(
        (
          best,
          alternativePlayer
        ) => {
          if (
            !eligibility[
              alternativePlayer
            ][typeIndex]
          ) {
            return best;
          }

          return Math.max(
            best,
            coberturaIndividualTotal_(
              dataset,
              alternativePlayer,
              typeIndex
            )
          );
        },
        0
      );

    cost +=
      Math.max(
        0,
        playerCoverage -
        bestAlternative
      ) *
      typeInfo.weaknessWeight;
  }

  return cost;
}





function evaluarSolucionCompacta_(
  solution,
  dataset
) {
  return evaluarSolucionJerarquica_(
    solution,
    dataset,
    {
      NORMALIZATION_SCALE: 1000000
    }
  );
}


function evaluarSolucionJerarquica_(
  solution,
  dataset,
  config
) {
  const principalTotalCoverages = [];
  const principalMetaCoverages = [];
  const pairTotalCoverages = [];
  const pairMetaCoverages = [];
  const finalTotalCoverages = [];
  const finalMetaCoverages = [];
  const backupTotalCoverages = [];
  const dependencies = [];

  const weightedPrincipalDeficits = [];
  const weightedPairDeficits = [];
  const weightedFinalDeficits = [];
  const weightedBackupDeficits = [];

  let preferredPrincipalCount = 0;
  let principalPrioritySum = 0;
  let totalFinalCoverage = 0;
  let totalMetaCoverage = 0;
  let totalPairCoverage = 0;
  let totalBackupCoverage = 0;
  let totalDependency = 0;
  let totalWeightedWeakTypeCoverage = 0;

  for (
    let typeIndex = 0;
    typeIndex <
      dataset.types.length;
    typeIndex++
  ) {
    const principal =
      solution.principalsByType[
        typeIndex
      ];

    const support1 =
      solution.support1ByType[
        typeIndex
      ];

    const support2 =
      solution.support2ByType[
        typeIndex
      ];

    const type =
      dataset.types[
        typeIndex
      ];

    const principalTotal =
      coberturaIndividualTotal_(
        dataset,
        principal,
        typeIndex
      );

    const principalMeta =
      coberturaIndividualMeta_(
        dataset,
        principal,
        typeIndex
      );

    const pairTotal =
      coberturaParejaTotal_(
        dataset,
        principal,
        support1,
        typeIndex
      );

    const pairMeta =
      coberturaParejaMeta_(
        dataset,
        principal,
        support1,
        typeIndex
      );

    const finalTotal =
      coberturaTrioTotal_(
        dataset,
        principal,
        support1,
        support2,
        typeIndex
      );

    const finalMeta =
      coberturaTrioMeta_(
        dataset,
        principal,
        support1,
        support2,
        typeIndex
      );

    const backup =
      coberturaRespaldoTotal_(
        dataset,
        principal,
        support1,
        support2,
        typeIndex
      );

    const dependency =
      Math.max(
        0,
        finalTotal -
        pairTotal
      );

    principalTotalCoverages.push(
      principalTotal
    );
    principalMetaCoverages.push(
      principalMeta
    );
    pairTotalCoverages.push(
      pairTotal
    );
    pairMetaCoverages.push(
      pairMeta
    );
    finalTotalCoverages.push(
      finalTotal
    );
    finalMetaCoverages.push(
      finalMeta
    );
    backupTotalCoverages.push(
      backup
    );
    dependencies.push(
      dependency
    );

    weightedPrincipalDeficits.push(
      Math.max(
        0,
        1 -
        principalTotal
      ) *
      type.weaknessWeight
    );

    weightedPairDeficits.push(
      Math.max(
        0,
        1 -
        pairTotal
      ) *
      type.weaknessWeight
    );

    weightedFinalDeficits.push(
      Math.max(
        0,
        1 -
        finalTotal
      ) *
      type.weaknessWeight
    );

    weightedBackupDeficits.push(
      Math.max(
        0,
        1 -
        backup
      ) *
      type.weaknessWeight
    );

    if (
      dataset.preferredPrincipalPlayers.has(
        principal
      )
    ) {
      preferredPrincipalCount++;
    }

    principalPrioritySum +=
      dataset.principalPriorityByPlayer[
        principal
      ];

    totalFinalCoverage +=
      finalTotal;
    totalMetaCoverage +=
      finalMeta;
    totalPairCoverage +=
      pairTotal;
    totalBackupCoverage +=
      backup;
    totalDependency +=
      dependency;

    totalWeightedWeakTypeCoverage +=
      principalTotal *
      type.weaknessWeight;
  }

  const opportunityCost =
    solution.opportunityCost ||
    calcularCosteOportunidadFinal_(
      solution,
      dataset
    );

  return {
    preferredPrincipalCount,
    principalPrioritySum,

    principalTotalCoverages,
    principalMetaCoverages,
    pairTotalCoverages,
    pairMetaCoverages,
    finalTotalCoverages,
    finalMetaCoverages,
    backupTotalCoverages,
    dependencies,

    sortedPrincipalTotalCoverages:
      ordenarCoberturas_(
        principalTotalCoverages
      ),

    sortedPrincipalMetaCoverages:
      ordenarCoberturas_(
        principalMetaCoverages
      ),

    sortedFinalTotalCoverages:
      ordenarCoberturas_(
        finalTotalCoverages
      ),

    sortedPairTotalCoverages:
      ordenarCoberturas_(
        pairTotalCoverages
      ),

    sortedBackupTotalCoverages:
      ordenarCoberturas_(
        backupTotalCoverages
      ),

    sortedFinalMetaCoverages:
      ordenarCoberturas_(
        finalMetaCoverages
      ),

    sortedDependenciesDescending:
      dependencies
        .slice()
        .sort(
          (
            first,
            second
          ) =>
            second -
            first
        ),

    sortedWeightedPrincipalDeficitsDescending:
      weightedPrincipalDeficits
        .slice()
        .sort(
          (
            first,
            second
          ) =>
            second -
            first
        ),

    sortedWeightedPairDeficitsDescending:
      weightedPairDeficits
        .slice()
        .sort(
          (
            first,
            second
          ) =>
            second -
            first
        ),

    sortedWeightedFinalDeficitsDescending:
      weightedFinalDeficits
        .slice()
        .sort(
          (
            first,
            second
          ) =>
            second -
            first
        ),

    sortedWeightedBackupDeficitsDescending:
      weightedBackupDeficits
        .slice()
        .sort(
          (
            first,
            second
          ) =>
            second -
            first
        ),

    totalFinalCoverage,
    totalMetaCoverage,
    totalPairCoverage,
    totalBackupCoverage,
    totalDependency,
    totalWeightedWeakTypeCoverage,
    opportunityCost,

    signature:
      solution.principalsByType.join("-") +
      "|" +
      solution.support1ByType.join("-") +
      "|" +
      solution.support2ByType.join("-")
  };
}





function esMejorEvaluacion_(
  candidate,
  currentBest,
  config
) {
  return esMejorEvaluacionJerarquica_(
    candidate,
    currentBest,
    config
  );
}


function esMejorEvaluacionJerarquica_(
  candidate,
  currentBest,
  config
) {
  if (!currentBest) {
    return true;
  }

  const mode =
    config.ALGORITHM_MODE ||
    "HYBRID_FULL";

  if (
    mode ===
    "BALANCED_COVERAGE"
  ) {
    return esMejorCoberturaEquilibrada_(
      candidate,
      currentBest
    );
  }

  if (
    mode ===
    "STRONG_PRINCIPALS"
  ) {
    return esMejorPrincipalesPotentes_(
      candidate,
      currentBest
    );
  }

  if (
    mode ===
    "GLOBAL_WEAKNESS"
  ) {
    return esMejorDebilidadGlobal_(
      candidate,
      currentBest
    );
  }

  return esMejorHibridoCompleto_(
    candidate,
    currentBest
  );
}


function esMejorCoberturaEquilibrada_(
  candidate,
  currentBest
) {
  const vectors = [
    "sortedFinalTotalCoverages",
    "sortedPrincipalTotalCoverages",
    "sortedPairTotalCoverages",
    "sortedBackupTotalCoverages",
    "sortedFinalMetaCoverages",
    "sortedPrincipalMetaCoverages"
  ];

  for (const field of vectors) {
    const comparison =
      compararArraysLexicograficamente_(
        candidate[field],
        currentBest[field]
      );

    if (comparison !== 0) {
      return comparison > 0;
    }
  }

  const dependencyComparison =
    compararArraysLexicograficamenteMenor_(
      candidate.sortedDependenciesDescending,
      currentBest.sortedDependenciesDescending
    );

  if (dependencyComparison !== 0) {
    return dependencyComparison > 0;
  }

  return desempatarEvaluaciones_(
    candidate,
    currentBest
  );
}


function esMejorPrincipalesPotentes_(
  candidate,
  currentBest
) {
  if (
    candidate.preferredPrincipalCount !==
    currentBest.preferredPrincipalCount
  ) {
    return (
      candidate.preferredPrincipalCount >
      currentBest.preferredPrincipalCount
    );
  }

  if (
    candidate.principalPrioritySum !==
    currentBest.principalPrioritySum
  ) {
    return (
      candidate.principalPrioritySum >
      currentBest.principalPrioritySum
    );
  }

  const vectors = [
    "sortedPrincipalTotalCoverages",
    "sortedPrincipalMetaCoverages",
    "sortedFinalTotalCoverages",
    "sortedPairTotalCoverages",
    "sortedBackupTotalCoverages",
    "sortedFinalMetaCoverages"
  ];

  for (const field of vectors) {
    const comparison =
      compararArraysLexicograficamente_(
        candidate[field],
        currentBest[field]
      );

    if (comparison !== 0) {
      return comparison > 0;
    }
  }

  return desempatarEvaluaciones_(
    candidate,
    currentBest
  );
}


function esMejorDebilidadGlobal_(
  candidate,
  currentBest
) {
  const lowerVectors = [
    "sortedWeightedPrincipalDeficitsDescending",
    "sortedWeightedFinalDeficitsDescending",
    "sortedWeightedPairDeficitsDescending",
    "sortedWeightedBackupDeficitsDescending"
  ];

  for (const field of lowerVectors) {
    const comparison =
      compararArraysLexicograficamenteMenor_(
        candidate[field],
        currentBest[field]
      );

    if (comparison !== 0) {
      return comparison > 0;
    }
  }

  const higherVectors = [
    "sortedFinalTotalCoverages",
    "sortedPrincipalTotalCoverages",
    "sortedPairTotalCoverages",
    "sortedBackupTotalCoverages",
    "sortedFinalMetaCoverages"
  ];

  for (const field of higherVectors) {
    const comparison =
      compararArraysLexicograficamente_(
        candidate[field],
        currentBest[field]
      );

    if (comparison !== 0) {
      return comparison > 0;
    }
  }

  return desempatarEvaluaciones_(
    candidate,
    currentBest
  );
}


function esMejorHibridoCompleto_(
  candidate,
  currentBest
) {
  if (
    candidate.preferredPrincipalCount !==
    currentBest.preferredPrincipalCount
  ) {
    return (
      candidate.preferredPrincipalCount >
      currentBest.preferredPrincipalCount
    );
  }

  const weakPrincipalComparison =
    compararArraysLexicograficamenteMenor_(
      candidate.sortedWeightedPrincipalDeficitsDescending,
      currentBest.sortedWeightedPrincipalDeficitsDescending
    );

  if (weakPrincipalComparison !== 0) {
    return weakPrincipalComparison > 0;
  }

  if (
    candidate.principalPrioritySum !==
    currentBest.principalPrioritySum
  ) {
    return (
      candidate.principalPrioritySum >
      currentBest.principalPrioritySum
    );
  }

  const weakFinalComparison =
    compararArraysLexicograficamenteMenor_(
      candidate.sortedWeightedFinalDeficitsDescending,
      currentBest.sortedWeightedFinalDeficitsDescending
    );

  if (weakFinalComparison !== 0) {
    return weakFinalComparison > 0;
  }

  const higherVectors = [
    "sortedFinalTotalCoverages",
    "sortedPrincipalTotalCoverages",
    "sortedPairTotalCoverages",
    "sortedPrincipalMetaCoverages"
  ];

  for (const field of higherVectors) {
    const comparison =
      compararArraysLexicograficamente_(
        candidate[field],
        currentBest[field]
      );

    if (comparison !== 0) {
      return comparison > 0;
    }
  }

  const lowerVectors = [
    "sortedWeightedPairDeficitsDescending",
    "sortedDependenciesDescending",
    "sortedWeightedBackupDeficitsDescending"
  ];

  for (const field of lowerVectors) {
    const comparison =
      compararArraysLexicograficamenteMenor_(
        candidate[field],
        currentBest[field]
      );

    if (comparison !== 0) {
      return comparison > 0;
    }
  }

  const remainingVectors = [
    "sortedBackupTotalCoverages",
    "sortedFinalMetaCoverages"
  ];

  for (const field of remainingVectors) {
    const comparison =
      compararArraysLexicograficamente_(
        candidate[field],
        currentBest[field]
      );

    if (comparison !== 0) {
      return comparison > 0;
    }
  }

  return desempatarEvaluaciones_(
    candidate,
    currentBest
  );
}


function desempatarEvaluaciones_(
  candidate,
  currentBest
) {
  const higherFields = [
    "totalFinalCoverage",
    "totalPairCoverage",
    "totalBackupCoverage",
    "totalMetaCoverage",
    "totalWeightedWeakTypeCoverage"
  ];

  for (const field of higherFields) {
    if (candidate[field] !== currentBest[field]) {
      return candidate[field] > currentBest[field];
    }
  }

  if (
    candidate.totalDependency !==
    currentBest.totalDependency
  ) {
    return (
      candidate.totalDependency <
      currentBest.totalDependency
    );
  }

  if (
    candidate.opportunityCost !==
    currentBest.opportunityCost
  ) {
    return (
      candidate.opportunityCost <
      currentBest.opportunityCost
    );
  }

  return (
    candidate.signature <
    currentBest.signature
  );
}



function compararArraysLexicograficamente_(
  first,
  second
) {
  for (
    let index = 0;
    index <
      Math.min(
        first.length,
        second.length
      );
    index++
  ) {
    if (
      first[index] !==
      second[index]
    ) {
      return (
        first[index] >
        second[index]
          ? 1
          : -1
      );
    }
  }

  return 0;
}


function compararArraysLexicograficamenteMenor_(
  first,
  second
) {
  for (
    let index = 0;
    index <
      Math.min(
        first.length,
        second.length
      );
    index++
  ) {
    if (
      first[index] !==
      second[index]
    ) {
      return (
        first[index] <
        second[index]
          ? 1
          : -1
      );
    }
  }

  return 0;
}


function crearEvaluacionesEtapas_(
  solution,
  dataset,
  config
) {
  const evaluation =
    evaluarSolucionJerarquica_(
      solution,
      dataset,
      config
    );

  return {
    principal: {
      totalCoverageScaled:
        evaluation.principalTotalCoverages.slice(),
      metaCoverageScaled:
        evaluation.principalMetaCoverages.slice()
    },

    support1: {
      totalCoverageScaled:
        evaluation.pairTotalCoverages.slice(),
      metaCoverageScaled:
        evaluation.pairMetaCoverages.slice()
    },

    support2: {
      totalCoverageScaled:
        evaluation.finalTotalCoverages.slice(),
      metaCoverageScaled:
        evaluation.finalMetaCoverages.slice(),
      backupCoverageScaled:
        evaluation.backupTotalCoverages.slice(),
      dependencyScaled:
        evaluation.dependencies.slice()
    }
  };
}


function calcularCosteOportunidadFinal_(
  solution,
  dataset
) {
  let cost = 0;

  const principalSet =
    new Set(
      solution.principalsByType
    );

  for (
    let assignedType = 0;
    assignedType <
      dataset.types.length;
    assignedType++
  ) {
    const playerIndex =
      solution.principalsByType[
        assignedType
      ];

    for (
      let alternativeType = 0;
      alternativeType <
        dataset.types.length;
      alternativeType++
    ) {
      if (
        alternativeType ===
        assignedType ||
        !dataset.principalEligibility[
          playerIndex
        ][alternativeType]
      ) {
        continue;
      }

      const currentAlternativePrincipal =
        solution.principalsByType[
          alternativeType
        ];

      const gain =
        coberturaIndividualTotal_(
          dataset,
          playerIndex,
          alternativeType
        ) -
        coberturaIndividualTotal_(
          dataset,
          currentAlternativePrincipal,
          alternativeType
        );

      cost +=
        Math.max(
          0,
          gain
        );
    }
  }

  return cost;
}


function insertarSolucionOrdenada_(
  current,
  solution,
  evaluation,
  dataset,
  config
) {
  if (
    current.some(
      item =>
        item.evaluation.signature ===
        evaluation.signature
    )
  ) {
    return current;
  }

  const next =
    current.concat([{
      solution:
        clonarSolucionCompactaCompleta_(
          solution
        ),
      evaluation
    }]);

  next.sort(
    (
      first,
      second
    ) => {
      if (
        esMejorEvaluacionJerarquica_(
          first.evaluation,
          second.evaluation,
          config
        )
      ) {
        return -1;
      }

      if (
        esMejorEvaluacionJerarquica_(
          second.evaluation,
          first.evaluation,
          config
        )
      ) {
        return 1;
      }

      return 0;
    }
  );

  return next.slice(
    0,
    config.ALTERNATIVE_SOLUTIONS_TO_KEEP +
      1
  );
}


function coberturaIndividualTotal_(
  dataset,
  playerIndex,
  typeIndex
) {
  return limitarCobertura_(
    dataset.rawScores[
      playerIndex
    ][typeIndex] /
    dataset.types[
      typeIndex
    ].maximumIndividualScore
  );
}


function coberturaIndividualMeta_(
  dataset,
  playerIndex,
  typeIndex
) {
  return limitarCobertura_(
    dataset.metaScores[
      playerIndex
    ][typeIndex] /
    dataset.types[
      typeIndex
    ].maximumIndividualScore
  );
}


function coberturaParejaTotal_(
  dataset,
  firstPlayer,
  secondPlayer,
  typeIndex
) {
  return limitarCobertura_(
    (
      dataset.rawScores[
        firstPlayer
      ][typeIndex] +
      dataset.rawScores[
        secondPlayer
      ][typeIndex]
    ) /
    (
      dataset.types[
        typeIndex
      ].maximumIndividualScore *
      2
    )
  );
}


function coberturaParejaMeta_(
  dataset,
  firstPlayer,
  secondPlayer,
  typeIndex
) {
  return limitarCobertura_(
    (
      dataset.metaScores[
        firstPlayer
      ][typeIndex] +
      dataset.metaScores[
        secondPlayer
      ][typeIndex]
    ) /
    (
      dataset.types[
        typeIndex
      ].maximumIndividualScore *
      2
    )
  );
}


function coberturaTrioTotal_(
  dataset,
  principal,
  support1,
  support2,
  typeIndex
) {
  return limitarCobertura_(
    (
      dataset.rawScores[
        principal
      ][typeIndex] +
      dataset.rawScores[
        support1
      ][typeIndex] +
      dataset.rawScores[
        support2
      ][typeIndex]
    ) /
    (
      dataset.types[
        typeIndex
      ].maximumIndividualScore *
      3
    )
  );
}


function coberturaTrioMeta_(
  dataset,
  principal,
  support1,
  support2,
  typeIndex
) {
  return limitarCobertura_(
    (
      dataset.metaScores[
        principal
      ][typeIndex] +
      dataset.metaScores[
        support1
      ][typeIndex] +
      dataset.metaScores[
        support2
      ][typeIndex]
    ) /
    (
      dataset.types[
        typeIndex
      ].maximumIndividualScore *
      3
    )
  );
}


function coberturaRespaldoTotal_(
  dataset,
  principal,
  support1,
  support2,
  typeIndex
) {
  const maximum =
    dataset.types[
      typeIndex
    ].maximumIndividualScore *
    2;

  return limitarCobertura_(
    Math.min(
      dataset.rawScores[
        principal
      ][typeIndex] +
        dataset.rawScores[
          support1
        ][typeIndex],

      dataset.rawScores[
        principal
      ][typeIndex] +
        dataset.rawScores[
          support2
        ][typeIndex],

      dataset.rawScores[
        support1
      ][typeIndex] +
        dataset.rawScores[
          support2
        ][typeIndex]
    ) /
    maximum
  );
}


function limitarCobertura_(
  value
) {
  if (
    !Number.isFinite(
      value
    )
  ) {
    return 0;
  }

  return Math.min(
    1,
    Math.max(
      0,
      value
    )
  );
}


function ordenarCoberturas_(
  values
) {
  return values
    .slice()
    .sort(
      (
        first,
        second
      ) =>
        first -
        second
    );
}


function compararVectoresCobertura_(
  first,
  second
) {
  return compararArraysLexicograficamente_(
    ordenarCoberturas_(
      first
    ),
    ordenarCoberturas_(
      second
    )
  );
}


function seleccionarEstadosDiversos_(
  states,
  width,
  keyFn
) {
  const selected = [];
  const seen =
    new Set();

  for (
    const state
    of states
  ) {
    const key =
      keyFn(
        state
      );

    if (
      seen.has(
        key
      )
    ) {
      continue;
    }

    selected.push(
      state
    );
    seen.add(
      key
    );

    if (
      selected.length >= width
    ) {
      return selected;
    }
  }

  for (
    const state
    of states
  ) {
    if (
      selected.includes(
        state
      )
    ) {
      continue;
    }

    selected.push(
      state
    );

    if (
      selected.length >= width
    ) {
      break;
    }
  }

  return selected;
}


function crearFirmaAsignaciones_(
  assignments
) {
  return assignments
    .map(
      value =>
        value == null
          ? "_"
          : value
    )
    .join("-");
}


function esFasePendienteFactible_(
  availablePlayers,
  pendingTypes,
  capacityFn,
  eligibilityMatrix
) {
  if (
    pendingTypes.length === 0
  ) {
    return true;
  }

  const candidates =
    availablePlayers.filter(
      playerIndex =>
        capacityFn(
          playerIndex
        ) > 0
    );

  const capacities =
    candidates.map(
      playerIndex =>
        capacityFn(
          playerIndex
        )
    );

  const result =
    resolverAsignacionFlujo_(
      candidates,
      capacities,
      pendingTypes.length,
      (
        playerIndex,
        localTypeIndex
      ) =>
        eligibilityMatrix[
          playerIndex
        ][
          pendingTypes[
            localTypeIndex
          ]
        ],
      () => 0
    );

  return Boolean(
    result
  );
}



function resolverAsignacionFlujo_(
  candidatePlayerIndexes,
  capacities,
  numberOfTypes,
  eligibilityFn,
  utilityFn
) {
  const requiredFlow = numberOfTypes;
  const totalCapacity = capacities.reduce(
    (total, value) => total + value,
    0
  );

  if (totalCapacity < requiredFlow) {
    return null;
  }

  const source = 0;
  const firstPlayerNode = 1;
  const firstTypeNode = firstPlayerNode + candidatePlayerIndexes.length;
  const sink = firstTypeNode + numberOfTypes;

  const graph = new MinCostFlow_(sink + 1);

  const edgeReferences = Array.from(
    { length: candidatePlayerIndexes.length },
    () => Array(numberOfTypes).fill(null)
  );

  candidatePlayerIndexes.forEach((playerIndex, localIndex) => {
    const playerNode = firstPlayerNode + localIndex;

    graph.addEdge(
      source,
      playerNode,
      capacities[localIndex],
      0
    );

    for (let typeIndex = 0; typeIndex < numberOfTypes; typeIndex++) {
      if (!eligibilityFn(playerIndex, typeIndex)) {
        continue;
      }

      edgeReferences[localIndex][typeIndex] = graph.addEdge(
        playerNode,
        firstTypeNode + typeIndex,
        1,
        -Math.round(utilityFn(playerIndex, typeIndex))
      );
    }
  });

  for (let typeIndex = 0; typeIndex < numberOfTypes; typeIndex++) {
    graph.addEdge(firstTypeNode + typeIndex, sink, 1, 0);
  }

  const result = graph.minCostMaxFlow(
    source,
    sink,
    requiredFlow
  );

  if (result.flow !== requiredFlow) {
    return null;
  }

  const byType = Array(numberOfTypes).fill(null);

  candidatePlayerIndexes.forEach((playerIndex, localIndex) => {
    for (let typeIndex = 0; typeIndex < numberOfTypes; typeIndex++) {
      const reference = edgeReferences[localIndex][typeIndex];

      if (!reference) {
        continue;
      }

      const edge = graph.graph[reference.from][reference.edgeIndex];
      const usedFlow = edge.initialCapacity - edge.capacity;

      if (usedFlow > 0) {
        byType[typeIndex] = playerIndex;
      }
    }
  });

  if (byType.some(value => value == null)) {
    return null;
  }

  return {
    byType,
    cost: result.cost
  };
}


/** Construye objetos finales de asignación. */

function construirSolucionFinal_(
  compactSolution,
  players,
  types,
  rawScores,
  normalizedScores
) {
  const assignmentsByType = Array.from(
    { length: types.length },
    () => []
  );

  const assignmentsByPlayer = Array.from(
    { length: players.length },
    () => []
  );

  for (let typeIndex = 0; typeIndex < types.length; typeIndex++) {
    const principalIndex = compactSolution.principalsByType[typeIndex];
    const support1Index = compactSolution.support1ByType[typeIndex];
    const support2Index = compactSolution.support2ByType[typeIndex];

    const principal = crearAsignacion_(
      principalIndex,
      typeIndex,
      "PRINCIPAL",
      1,
      players,
      types,
      rawScores,
      normalizedScores
    );

    const support1 = crearAsignacion_(
      support1Index,
      typeIndex,
      "APOYO_1",
      2,
      players,
      types,
      rawScores,
      normalizedScores
    );

    const support2 = crearAsignacion_(
      support2Index,
      typeIndex,
      "APOYO_2",
      3,
      players,
      types,
      rawScores,
      normalizedScores
    );

    assignmentsByType[typeIndex] = [principal, support1, support2];
    assignmentsByPlayer[principalIndex].push(principal);
    assignmentsByPlayer[support1Index].push(support1);
    assignmentsByPlayer[support2Index].push(support2);
  }

  return {
    assignmentsByType,
    assignmentsByPlayer
  };
}


function crearAsignacion_(
  playerIndex,
  typeIndex,
  role,
  position,
  players,
  types,
  rawScores,
  normalizedScores
) {
  return {
    playerIndex,
    typeIndex,
    player: players[playerIndex],
    type: types[typeIndex].name,
    score: rawScores[playerIndex][typeIndex],
    normalizedScore: normalizedScores[playerIndex][typeIndex],
    role,
    position
  };
}


/** Valida los grupos y restricciones finales. */


/**
 * Valida la distribución secuencial:
 *
 * - 8 Principales distintos.
 * - 8 Apoyos 1 distintos.
 * - 4 Apoyos 2, dos usos cada uno.
 * - Grupos disjuntos.
 * - Elegibilidad de modalidad y EX Zone.
 */
function validarSolucionSecuencial_(
  solution,
  dataset,
  players,
  config
) {
  const errors = [];

  const principalUsed =
    new Set();

  const support1Used =
    new Set();

  const support2Count =
    new Map();

  solution.assignmentsByType.forEach(
    (
      assignments,
      typeIndex
    ) => {
      const typeName =
        dataset.types[
          typeIndex
        ].name;

      if (
        assignments.length !==
        config.PLAYERS_PER_TYPE
      ) {
        errors.push(
          `${typeName} no tiene exactamente ` +
          `${config.PLAYERS_PER_TYPE} atacantes.`
        );

        return;
      }

      const playersInsideType =
        new Set(
          assignments.map(
            assignment =>
              assignment.playerIndex
          )
        );

      if (
        playersInsideType.size !==
        assignments.length
      ) {
        errors.push(
          `Hay un jugador repetido dentro de ${typeName}.`
        );
      }

      const principal =
        assignments.find(
          assignment =>
            assignment.role ===
            "PRINCIPAL"
        );

      const support1 =
        assignments.find(
          assignment =>
            assignment.role ===
            "APOYO_1"
        );

      const support2 =
        assignments.find(
          assignment =>
            assignment.role ===
            "APOYO_2"
        );

      if (!principal) {
        errors.push(
          `${typeName} no tiene Principal.`
        );
      } else {
        if (
          !dataset.principalEligibility[
            principal.playerIndex
          ][typeIndex]
        ) {
          errors.push(
            `"${principal.player}" no cumple EX Zone/Weather o ` +
            `la modalidad necesaria para ser Principal de ${typeName}.`
          );
        }

        if (
          principalUsed.has(
            principal.playerIndex
          )
        ) {
          errors.push(
            `"${principal.player}" aparece como Principal más de una vez.`
          );
        }

        principalUsed.add(
          principal.playerIndex
        );
      }

      if (!support1) {
        errors.push(
          `${typeName} no tiene Apoyo 1.`
        );
      } else {
        if (
          !dataset.assignmentEligibility[
            support1.playerIndex
          ][typeIndex]
        ) {
          errors.push(
            `"${support1.player}" no tiene fuerza válida para ` +
            `${typeName} en ${dataset.label}.`
          );
        }

        if (
          support1Used.has(
            support1.playerIndex
          )
        ) {
          errors.push(
            `"${support1.player}" aparece como Apoyo 1 más de una vez.`
          );
        }

        support1Used.add(
          support1.playerIndex
        );
      }

      if (!support2) {
        errors.push(
          `${typeName} no tiene Apoyo 2.`
        );
      } else {
        if (
          !dataset.assignmentEligibility[
            support2.playerIndex
          ][typeIndex]
        ) {
          errors.push(
            `"${support2.player}" no tiene fuerza válida para ` +
            `${typeName} en ${dataset.label}.`
          );
        }

        support2Count.set(
          support2.playerIndex,

          (
            support2Count.get(
              support2.playerIndex
            ) || 0
          ) + 1
        );
      }

      /*
       * La jerarquía Meta es una preferencia estratégica por defecto, no
       * una restricción dura. Solo se valida como error cuando se activa
       * explícitamente ENFORCE_META_HIERARCHY.
       */
      if (
        config.ENFORCE_META_HIERARCHY &&
        principal &&
        support1 &&
        dataset.metaScores[
          support1.playerIndex
        ][typeIndex] >
        dataset.metaScores[
          principal.playerIndex
        ][typeIndex]
      ) {
        errors.push(
          `${typeName}: Apoyo 1 tiene más puntos Meta que el Principal.`
        );
      }

      if (
        config.ENFORCE_META_HIERARCHY &&
        principal &&
        support2 &&
        dataset.metaScores[
          support2.playerIndex
        ][typeIndex] >
        dataset.metaScores[
          principal.playerIndex
        ][typeIndex]
      ) {
        errors.push(
          `${typeName}: Apoyo 2 tiene más puntos Meta que el Principal.`
        );
      }
    }
  );

  if (
    principalUsed.size !==
    config.PRINCIPAL_COUNT
  ) {
    errors.push(
      `Se esperaban ${config.PRINCIPAL_COUNT} Principales distintos, ` +
      `pero se han encontrado ${principalUsed.size}.`
    );
  }

  dataset.mandatoryPrincipalPlayers.forEach(
    playerIndex => {
      if (
        !principalUsed.has(
          playerIndex
        )
      ) {
        errors.push(
          `"${players[playerIndex]}" era Principal prioritario obligatorio ` +
          `y no aparece en el grupo de Principales.`
        );
      }
    }
  );

  if (
    support1Used.size !==
    config.SUPPORT1_COUNT
  ) {
    errors.push(
      `Se esperaban ${config.SUPPORT1_COUNT} Apoyos 1 distintos, ` +
      `pero se han encontrado ${support1Used.size}.`
    );
  }

  principalUsed.forEach(
    playerIndex => {
      if (
        support1Used.has(
          playerIndex
        ) ||
        support2Count.has(
          playerIndex
        )
      ) {
        errors.push(
          `"${players[playerIndex]}" aparece en más de un grupo de atacantes.`
        );
      }
    }
  );

  support1Used.forEach(
    playerIndex => {
      if (
        support2Count.has(
          playerIndex
        )
      ) {
        errors.push(
          `"${players[playerIndex]}" aparece como Apoyo 1 y Apoyo 2.`
        );
      }
    }
  );

  const expectedSupport2Players =
    players.length -
    config.PRINCIPAL_COUNT -
    config.SUPPORT1_COUNT;

  if (
    support2Count.size !==
    expectedSupport2Players
  ) {
    errors.push(
      `Se esperaban ${expectedSupport2Players} jugadores distintos ` +
      `en Apoyo 2, pero se han encontrado ${support2Count.size}.`
    );
  }

  support2Count.forEach(
    (
      count,
      playerIndex
    ) => {
      if (
        count !==
        config
          .SUPPORT2_ASSIGNMENTS_PER_PLAYER
      ) {
        errors.push(
          `"${players[playerIndex]}" aparece ${count} veces como ` +
          `Apoyo 2, pero debería aparecer exactamente ` +
          `${config.SUPPORT2_ASSIGNMENTS_PER_PLAYER} veces.`
        );
      }
    }
  );

  solution.assignmentsByPlayer.forEach(
    (
      assignments,
      playerIndex
    ) => {
      if (
        assignments.length === 0
      ) {
        errors.push(
          `"${players[playerIndex]}" no aparece en ninguna asignación.`
        );
      }
    }
  );

  if (
    errors.length > 0
  ) {
    throw new Error(
      `La asignación "${dataset.label}" no es válida:\n\n` +
      errors.join("\n")
    );
  }
}

function generarAsignacionSetters_(
  players,
  types,
  normalizedScores,
  globalSolution,
  rebuffUnitsByType,
  config
) {
  const powerByPlayer = calcularPotenciaGlobal_(
    normalizedScores,
    config
  );

  const ranking = players
    .map((player, playerIndex) => ({
      player,
      playerIndex,
      globalPower: powerByPlayer[playerIndex]
    }))
    .sort((a, b) => {
      if (b.globalPower !== a.globalPower) {
        return b.globalPower - a.globalPower;
      }
      return a.player.localeCompare(b.player);
    });

  const excludedTopPlayers = new Set(
    ranking
      .slice(
        0,
        Math.min(
          config.TOP_PLAYERS_EXCLUDED_FROM_SETTERS,
          players.length
        )
      )
      .map(item => item.playerIndex)
  );

  const candidatesByType = types.map((type, typeIndex) => {
    const attackersForType = new Set(
      globalSolution.assignmentsByType[typeIndex].map(
        assignment => assignment.playerIndex
      )
    );

    const candidates = [];

    rebuffUnitsByType[typeIndex].forEach((unitMap, playerIndex) => {
      if (
        excludedTopPlayers.has(playerIndex) ||
        attackersForType.has(playerIndex)
      ) {
        return;
      }

      candidates.push({
        playerIndex,
        player: players[playerIndex],
        globalPower: powerByPlayer[playerIndex],
        normalizedTypePower:
          normalizedScores[playerIndex][typeIndex] /
          config.NORMALIZATION_SCALE,
        units: Array.from(unitMap.values())
      });
    });

    candidates.sort(compararSettersPorDebilidad_);
    return candidates;
  });

  const assignmentsByType = Array.from(
    { length: types.length },
    () => Array(config.SETTERS_PER_TYPE).fill(null)
  );

  const setter1CountByPlayer = Array(players.length).fill(0);
  const totalSetterCountByPlayer = Array(players.length).fill(0);

  const typeProcessingOrder = types
    .map((type, typeIndex) => ({
      typeIndex,
      candidateCount: candidatesByType[typeIndex].length
    }))
    .sort((a, b) => {
      if (a.candidateCount !== b.candidateCount) {
        return a.candidateCount - b.candidateCount;
      }
      return a.typeIndex - b.typeIndex;
    });

  // Seteador 1: débil y máximo 3 veces.
  typeProcessingOrder.forEach(({ typeIndex }) => {
    const available = candidatesByType[typeIndex]
      .filter(
        candidate =>
          setter1CountByPlayer[candidate.playerIndex] <
          config.MAX_SETTER_1_ASSIGNMENTS
      )
      .sort((a, b) => {
        const weakness = compararSettersPorDebilidad_(a, b);
        if (weakness !== 0) {
          return weakness;
        }

        const usageDifference =
          setter1CountByPlayer[a.playerIndex] -
          setter1CountByPlayer[b.playerIndex];

        if (usageDifference !== 0) {
          return usageDifference;
        }

        return a.player.localeCompare(b.player);
      });

    const selected = available.length > 0 ? available[0] : null;
    assignmentsByType[typeIndex][0] = selected;

    if (selected) {
      setter1CountByPlayer[selected.playerIndex]++;
      totalSetterCountByPlayer[selected.playerIndex]++;
    }
  });

  // Seteadores 2 y 3.
  for (let typeIndex = 0; typeIndex < types.length; typeIndex++) {
    const selectedPlayers = new Set();
    const setter1 = assignmentsByType[typeIndex][0];

    if (setter1) {
      selectedPlayers.add(setter1.playerIndex);
    }

    for (
      let position = 1;
      position < config.SETTERS_PER_TYPE;
      position++
    ) {
      const available = candidatesByType[typeIndex]
        .filter(
          candidate => !selectedPlayers.has(candidate.playerIndex)
        )
        .sort((a, b) => {
          const weakness = compararSettersPorDebilidad_(a, b);
          if (weakness !== 0) {
            return weakness;
          }

          const usageDifference =
            totalSetterCountByPlayer[a.playerIndex] -
            totalSetterCountByPlayer[b.playerIndex];

          if (usageDifference !== 0) {
            return usageDifference;
          }

          return a.player.localeCompare(b.player);
        });

      const selected = available.length > 0 ? available[0] : null;
      assignmentsByType[typeIndex][position] = selected;

      if (selected) {
        selectedPlayers.add(selected.playerIndex);
        totalSetterCountByPlayer[selected.playerIndex]++;
      }
    }
  }

  return {
    assignmentsByType,
    setter1CountByPlayer,
    totalSetterCountByPlayer
  };
}


function compararSettersPorDebilidad_(a, b) {
  if (a.globalPower !== b.globalPower) {
    return a.globalPower - b.globalPower;
  }
  if (a.normalizedTypePower !== b.normalizedTypePower) {
    return a.normalizedTypePower - b.normalizedTypePower;
  }
  if (b.units.length !== a.units.length) {
    return b.units.length - a.units.length;
  }
  return a.player.localeCompare(b.player);
}


/** Escribe las 3 tablas de atacantes y la de seteadores. */
function escribirResultadoCompleto_(
  spreadsheet,
  players,
  globalResult,
  physicalResult,
  specialResult,
  setterAssignments,
  config
) {
  let outputSheet =
    spreadsheet.getSheetByName(
      config.OUTPUT_SHEET
    );

  if (!outputSheet) {
    outputSheet =
      spreadsheet.insertSheet(
        config.OUTPUT_SHEET
      );
  } else {
    outputSheet
      .getDataRange()
      .breakApart();
    outputSheet.clear();
    outputSheet
      .clearConditionalFormatRules();
  }

  const generatedAt =
    Utilities.formatDate(
      new Date(),
      spreadsheet
        .getSpreadsheetTimeZone(),
      "dd/MM/yyyy HH:mm:ss"
    );

  escribirPanelAlgoritmo_(
    outputSheet,
    config
  );

  const coverageRanges = [];
  let nextRow = 5;

  const modeResults = [
    globalResult,
    physicalResult,
    specialResult
  ];

  nextRow = escribirTablaAtacantes_(
    outputSheet,
    nextRow,
    "ASIGNACIÓN DE FUERZA GLOBAL",
    `${config.SELECTED_ALGORITHM_LABEL} · Meta y no Meta contribuyen`,
    globalResult,
    generatedAt,
    coverageRanges
  );

  nextRow += 2;

  nextRow = escribirTablaAtacantes_(
    outputSheet,
    nextRow,
    "ASIGNACIÓN DE FUERZA FÍSICA",
    `Physical/Hybrid · ${config.SELECTED_ALGORITHM_LABEL}`,
    physicalResult,
    generatedAt,
    coverageRanges
  );

  nextRow += 2;

  nextRow = escribirTablaAtacantes_(
    outputSheet,
    nextRow,
    "ASIGNACIÓN DE FUERZA ESPECIAL",
    `Special/Hybrid · ${config.SELECTED_ALGORITHM_LABEL}`,
    specialResult,
    generatedAt,
    coverageRanges
  );

  nextRow += 2;

  nextRow = escribirTablaDebilidadGlobal_(
    outputSheet,
    nextRow,
    modeResults,
    generatedAt
  );

  nextRow += 2;

  nextRow = escribirTablaSetters_(
    outputSheet,
    nextRow,
    globalResult.dataset.types,
    setterAssignments,
    generatedAt,
    config
  );

  nextRow += 2;

  escribirTablaAlternativasPrincipal_(
    outputSheet,
    nextRow,
    players,
    modeResults,
    generatedAt,
    config
  );

  outputSheet.setConditionalFormatRules([
    SpreadsheetApp
      .newConditionalFormatRule()
      .setGradientMinpoint(
        "#f8696b"
      )
      .setGradientMidpointWithValue(
        "#ffeb84",
        SpreadsheetApp.InterpolationType.NUMBER,
        "0.75"
      )
      .setGradientMaxpoint(
        "#63be7b"
      )
      .setRanges(
        coverageRanges
      )
      .build()
  ]);

  outputSheet.setFrozenRows(2);

  const widths = [
    120,
    175,
    85,
    85,
    105,
    95,
    175,
    85,
    85,
    105,
    175,
    85,
    85,
    95,
    105,
    105,
    390
  ];

  widths.forEach(
    (
      width,
      index
    ) =>
      outputSheet.setColumnWidth(
        index + 1,
        width
      )
  );

  outputSheet.autoResizeRows(
    1,
    outputSheet.getLastRow()
  );
}





function escribirTablaAtacantes_(
  sheet,
  startRow,
  title,
  subtitle,
  result,
  generatedAt,
  coverageRanges
) {
  const headers = [
    "TIPO",
    "PRINCIPAL",
    "PUNTOS",
    "META",
    "COB. PRINCIPAL",
    "DÉFICIT",
    "APOYO 1",
    "PUNTOS",
    "META",
    "COB. P+A1",
    "APOYO 2",
    "PUNTOS",
    "META",
    "TOTAL",
    "COB. TRÍO",
    "RESPALDO",
    "AVISOS"
  ];

  const rows =
    result.solution.assignmentsByType.map(
      (
        assignments,
        typeIndex
      ) => {
        const principal =
          assignments.find(
            item =>
              item.role ===
              "PRINCIPAL"
          );

        const support1 =
          assignments.find(
            item =>
              item.role ===
              "APOYO_1"
          );

        const support2 =
          assignments.find(
            item =>
              item.role ===
              "APOYO_2"
          );

        const typeInfo =
          result.dataset.types[
            typeIndex
          ];

        const maximum =
          typeInfo.maximumIndividualScore;

        const principalMeta =
          result.dataset.metaScores[
            principal.playerIndex
          ][typeIndex];

        const support1Meta =
          result.dataset.metaScores[
            support1.playerIndex
          ][typeIndex];

        const support2Meta =
          result.dataset.metaScores[
            support2.playerIndex
          ][typeIndex];

        const pairTotal =
          principal.score +
          support1.score;

        const total =
          pairTotal +
          support2.score;

        const backup =
          Math.min(
            principal.score +
              support1.score,
            principal.score +
              support2.score,
            support1.score +
              support2.score
          );

        const principalCoverage =
          maximum > 0
            ? Math.min(
                1,
                principal.score /
                  maximum
              )
            : 0;

        const deficit =
          Math.max(
            0,
            maximum -
            principal.score
          );

        const pairCoverage =
          maximum > 0
            ? Math.min(
                1,
                pairTotal /
                  (
                    maximum * 2
                  )
              )
            : 0;

        const trioCoverage =
          maximum > 0
            ? Math.min(
                1,
                total /
                  (
                    maximum * 3
                  )
              )
            : 0;

        const backupCoverage =
          maximum > 0
            ? Math.min(
                1,
                backup /
                  (
                    maximum * 2
                  )
              )
            : 0;

        const warnings =
          crearAvisosAtacantes_(
            principal,
            support1,
            support2,
            typeIndex,
            result.dataset
          );

        return [
          typeInfo.name,
          principal.player,
          principal.score,
          principalMeta,
          principalCoverage,
          deficit,
          support1.player,
          support1.score,
          support1Meta,
          pairCoverage,
          support2.player,
          support2.score,
          support2Meta,
          total,
          trioCoverage,
          backupCoverage,
          warnings
        ];
      }
    );

  const titleRow = startRow;
  const subtitleRow = startRow + 1;
  const generatedRow = startRow + 2;
  const headerRow = startRow + 3;
  const dataStartRow = startRow + 4;

  sheet.getRange(
    titleRow,
    1,
    1,
    headers.length
  )
    .merge()
    .setValue(title)
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#1f4e78")
    .setFontColor("#ffffff");

  sheet.getRange(
    subtitleRow,
    1,
    1,
    headers.length
  )
    .merge()
    .setValue(subtitle)
    .setHorizontalAlignment("center")
    .setFontStyle("italic")
    .setBackground("#d9eaf7");

  sheet.getRange(
    generatedRow,
    1,
    1,
    headers.length
  )
    .merge()
    .setValue(
      `Generado: ${generatedAt}`
    )
    .setHorizontalAlignment("center")
    .setFontColor("#666666");

  sheet.getRange(
    headerRow,
    1,
    1,
    headers.length
  )
    .setValues([
      headers
    ])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#5b9bd5")
    .setFontColor("#ffffff");

  sheet.getRange(
    dataStartRow,
    1,
    rows.length,
    headers.length
  )
    .setValues(rows)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true
    );

  [
    3,
    4,
    6,
    8,
    9,
    12,
    13,
    14
  ].forEach(
    column =>
      sheet.getRange(
        dataStartRow,
        column,
        rows.length,
        1
      )
        .setNumberFormat(
          "0.00"
        )
  );

  [
    5,
    10,
    15,
    16
  ].forEach(
    column => {
      const range =
        sheet.getRange(
          dataStartRow,
          column,
          rows.length,
          1
        );

      range.setNumberFormat(
        "0.0%"
      );

      coverageRanges.push(
        range
      );
    }
  );

  sheet.getRange(
    dataStartRow,
    2,
    rows.length,
    5
  ).setBackground(
    "#fff2cc"
  );

  sheet.getRange(
    dataStartRow,
    7,
    rows.length,
    4
  ).setBackground(
    "#e2f0d9"
  );

  sheet.getRange(
    dataStartRow,
    11,
    rows.length,
    3
  ).setBackground(
    "#fce4d6"
  );

  sheet.getRange(
    dataStartRow,
    17,
    rows.length,
    1
  ).setWrap(true);

  return (
    dataStartRow +
    rows.length
  );
}





function crearAvisosAtacantes_(
  principal,
  support1,
  support2,
  typeIndex,
  dataset
) {
  const warnings = [];

  const principalMeta =
    dataset.metaScores[
      principal.playerIndex
    ][typeIndex];

  const support1Meta =
    dataset.metaScores[
      support1.playerIndex
    ][typeIndex];

  const support2Meta =
    dataset.metaScores[
      support2.playerIndex
    ][typeIndex];

  if (
    support1.score >
      principal.score &&
    support1Meta <=
      principalMeta
  ) {
    warnings.push(
      "Apoyo 1 tiene más puntos totales, pero el Principal tiene igual o más Meta"
    );
  }

  if (
    support2.score >
      principal.score &&
    support2Meta <=
      principalMeta
  ) {
    warnings.push(
      "Apoyo 2 tiene más puntos totales, pero el Principal tiene igual o más Meta"
    );
  }

  [
    principal,
    support1,
    support2
  ].forEach(
    assignment => {
      const confidence =
        dataset.rosterConfidenceByPlayer[
          assignment.playerIndex
        ];

      if (
        confidence <
        dataset.rosterConfidenceWarningThreshold
      ) {
        warnings.push(
          `${assignment.player}: roster ${(confidence * 100).toFixed(0)} %`
        );
      }
    }
  );

  const type =
    dataset.types[
      typeIndex
    ];

  if (
    type.isCriticalGlobalWeakType
  ) {
    warnings.push(
      `Catálogo globalmente débil #${type.globalWeaknessRank}: ` +
      `${(type.globalWeakness * 100).toFixed(0)} %`
    );
  }

  return warnings.join(
    " · "
  );
}


function escribirTablaDebilidadGlobal_(
  sheet,
  startRow,
  modeResults,
  generatedAt
) {
  const headers = [
    "TABLA",
    "TIPO",
    "UNIDADES META",
    "POTENCIA CATÁLOGO",
    "POTENCIA REL.",
    "ESCASEZ REL.",
    "DEBILIDAD GLOBAL",
    "RANGO DEBILIDAD",
    "BANNERS META"
  ];

  const rows = [];

  modeResults.forEach(result => {
    result.dataset.types
      .slice()
      .sort(
        (
          first,
          second
        ) =>
          second.globalWeakness -
          first.globalWeakness
      )
      .forEach(type => {
        rows.push([
          result.dataset.label,
          type.name,
          type.metaUnitCount,
          type.catalogPower,
          type.catalogPowerRelative,
          type.metaCountRelative,
          type.globalWeakness,
          type.globalWeaknessRank,
          type.bannerSummary
        ]);
      });
  });

  const titleRow = startRow;
  const subtitleRow = startRow + 1;
  const generatedRow = startRow + 2;
  const headerRow = startRow + 3;
  const dataStartRow = startRow + 4;

  sheet
    .getRange(
      titleRow,
      1,
      1,
      headers.length
    )
    .merge()
    .setValue(
      "DEBILIDAD GLOBAL DEL CATÁLOGO META"
    )
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#548235")
    .setFontColor("#ffffff");

  sheet
    .getRange(
      subtitleRow,
      1,
      1,
      headers.length
    )
    .merge()
    .setValue(
      "70 % potencia máxima Roster!P + 30 % escasez de unidades · Meta!C no se vuelve a aplicar"
    )
    .setHorizontalAlignment("center")
    .setFontStyle("italic")
    .setBackground("#e2f0d9");

  sheet
    .getRange(
      generatedRow,
      1,
      1,
      headers.length
    )
    .merge()
    .setValue(
      `Generado: ${generatedAt}`
    )
    .setHorizontalAlignment("center")
    .setFontColor("#666666");

  sheet
    .getRange(
      headerRow,
      1,
      1,
      headers.length
    )
    .setValues([
      headers
    ])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#70ad47")
    .setFontColor("#ffffff");

  sheet
    .getRange(
      dataStartRow,
      1,
      rows.length,
      headers.length
    )
    .setValues(rows)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true
    );

  sheet
    .getRange(
      dataStartRow,
      4,
      rows.length,
      1
    )
    .setNumberFormat("0.00");

  sheet
    .getRange(
      dataStartRow,
      5,
      rows.length,
      3
    )
    .setNumberFormat("0.0%");

  sheet
    .getRange(
      dataStartRow,
      9,
      rows.length,
      1
    )
    .setWrap(true);

  return (
    dataStartRow +
    rows.length
  );
}


function escribirTablaSetters_(
  sheet,
  startRow,
  types,
  setterAssignments,
  generatedAt,
  config
) {
  const headers = [
    "TIPO",
    "SETEADOR 1",
    "SETEADOR 2",
    "SETEADOR 3"
  ];

  const rows = setterAssignments.assignmentsByType.map(
    (setters, typeIndex) => [
      types[typeIndex].name,
      setters[0] ? setters[0].player : "SIN CANDIDATO",
      setters[1] ? setters[1].player : "SIN CANDIDATO",
      setters[2] ? setters[2].player : "SIN CANDIDATO"
    ]
  );

  const titleRow = startRow;
  const subtitleRow = startRow + 1;
  const generatedRow = startRow + 2;
  const headerRow = startRow + 3;
  const dataStartRow = startRow + 4;

  sheet
    .getRange(titleRow, 1, 1, 4)
    .merge()
    .setValue("ASIGNACIÓN DE SETEADORES")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#c55a11")
    .setFontColor("#ffffff");

  sheet
    .getRange(subtitleRow, 1, 1, 4)
    .merge()
    .setValue(
      "Tres jugadores débiles por tipo · " +
      `${config.TOP_PLAYERS_EXCLUDED_FROM_SETTERS} jugadores top excluidos · ` +
      `Seteador 1 máximo ${config.MAX_SETTER_1_ASSIGNMENTS} veces`
    )
    .setHorizontalAlignment("center")
    .setFontStyle("italic")
    .setBackground("#fce4d6");

  sheet
    .getRange(generatedRow, 1, 1, 4)
    .merge()
    .setValue(`Datos de Roster actualizados: ${generatedAt}`)
    .setHorizontalAlignment("center")
    .setFontColor("#666666");

  sheet
    .getRange(headerRow, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#ed7d31")
    .setFontColor("#ffffff");

  sheet
    .getRange(dataStartRow, 1, rows.length, headers.length)
    .setValues(rows)
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment("center");

  sheet.getRange(dataStartRow, 2, rows.length, 3).setBackground("#fce4d6");

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    for (let column = 2; column <= 4; column++) {
      if (rows[rowIndex][column - 1] === "SIN CANDIDATO") {
        sheet
          .getRange(dataStartRow + rowIndex, column)
          .setBackground("#f4cccc")
          .setFontColor("#9c0006")
          .setFontWeight("bold");
      }
    }
  }

  return (
    dataStartRow +
    rows.length
  );
}


/**
 * Escribe todos los ajustes automáticos y las sugerencias no aplicadas.
 */
function escribirResumenBusqueda_(
  sheet,
  startRow,
  modeResults,
  generatedAt,
  config
) {
  const headers = [
    "TABLA",
    "PRIORITARIOS EN PRINCIPAL",
    "PRINCIPAL MÍN.",
    "PEOR COB. TRÍO",
    "COB. MEDIA",
    "PEOR P+A1",
    "PEOR RESPALDO",
    "MÁX. DEP. A2",
    "A2 PROTEGIDOS",
    "GRUPOS P",
    "GRUPOS A1",
    "REPARTOS A2"
  ];

  const rows =
    modeResults.map(
      result => {
        const evaluation =
          result.evaluation;

        const support2Players =
          new Set(
            result.solution.assignmentsByType.map(
              assignments =>
                assignments.find(
                  item =>
                    item.role ===
                    "APOYO_2"
                ).playerIndex
            )
          );

        const protectedSupport2 =
          Array.from(
            support2Players
          ).filter(
            playerIndex =>
              result.dataset.protectedSupport2Players.has(
                playerIndex
              )
          ).length;

        const search =
          result.multiStartSummary || {};

        return [
          result.dataset.label,
          evaluation.preferredPrincipalCount,
          Math.min(
            ...evaluation.principalTotalCoverages
          ),
          Math.min(
            ...evaluation.finalTotalCoverages
          ),
          evaluation.totalFinalCoverage /
            evaluation.finalTotalCoverages.length,
          Math.min(
            ...evaluation.pairTotalCoverages
          ),
          Math.min(
            ...evaluation.backupTotalCoverages
          ),
          Math.max(
            ...evaluation.dependencies
          ),
          protectedSupport2,
          search.principalChecked || 0,
          search.support1Checked || 0,
          search.support2Checked || 0
        ];
      }
    );

  const titleRow = startRow;
  const subtitleRow = startRow + 1;
  const generatedRow = startRow + 2;
  const headerRow = startRow + 3;
  const dataStartRow = startRow + 4;

  sheet.getRange(
    titleRow,
    1,
    1,
    headers.length
  )
    .merge()
    .setValue(
      "RESUMEN DE OPTIMIZACIÓN"
    )
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#7030a0")
    .setFontColor("#ffffff");

  sheet.getRange(
    subtitleRow,
    1,
    1,
    headers.length
  )
    .merge()
    .setValue(
      "Principales prioritarios · coste de oportunidad dinámico · max-min final"
    )
    .setHorizontalAlignment("center")
    .setFontStyle("italic")
    .setBackground("#e4dfec");

  sheet.getRange(
    generatedRow,
    1,
    1,
    headers.length
  )
    .merge()
    .setValue(
      `Generado: ${generatedAt}`
    )
    .setHorizontalAlignment("center")
    .setFontColor("#666666");

  sheet.getRange(
    headerRow,
    1,
    1,
    headers.length
  )
    .setValues([
      headers
    ])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#8064a2")
    .setFontColor("#ffffff");

  sheet.getRange(
    dataStartRow,
    1,
    rows.length,
    headers.length
  )
    .setValues(rows)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true
    )
    .setHorizontalAlignment("center");

  sheet.getRange(
    dataStartRow,
    3,
    rows.length,
    6
  ).setNumberFormat(
    "0.0%"
  );

  return (
    dataStartRow +
    rows.length
  );
}


function escribirTablaAlternativasPrincipal_(
  sheet,
  startRow,
  players,
  modeResults,
  generatedAt,
  config
) {
  const headers = [
    "TABLA",
    "TIPO",
    "PRINCIPAL ACTUAL",
    "ALTERNATIVA",
    "ROL ACTUAL",
    "PUNTOS",
    "META",
    "COBERTURA",
    "DIF. PUNTOS",
    "DIF. META",
    "OBSERVACIÓN"
  ];

  const rows = [];

  modeResults.forEach(
    result => {
      result.dataset.types.forEach(
        (
          typeInfo,
          typeIndex
        ) => {
          const principal =
            result.solution
              .assignmentsByType[
                typeIndex
              ]
              .find(
                assignment =>
                  assignment.role ===
                  "PRINCIPAL"
              );

          const candidates =
            players
              .map(
                (
                  player,
                  playerIndex
                ) => ({
                  player,
                  playerIndex
                })
              )
              .filter(
                candidate =>
                  candidate.playerIndex !==
                    principal.playerIndex &&
                  result.dataset
                    .principalEligibility[
                      candidate.playerIndex
                    ][typeIndex]
              )
              .map(
                candidate => {
                  const score =
                    result.dataset.rawScores[
                      candidate.playerIndex
                    ][typeIndex];

                  const metaScore =
                    result.dataset.metaScores[
                      candidate.playerIndex
                    ][typeIndex];

                  const currentRole =
                    obtenerRolActualAlternativa_(
                      result,
                      candidate.playerIndex
                    );

                  return {
                    player:
                      candidate.player,

                    playerIndex:
                      candidate.playerIndex,

                    score,
                    metaScore,

                    coverage:
                      typeInfo
                        .maximumIndividualScore > 0
                        ? Math.min(
                            1,
                            Math.max(
                              0,
                              score /
                                typeInfo
                                  .maximumIndividualScore
                            )
                          )
                        : 0,

                    currentRole
                  };
                }
              )
              .sort(
                (
                  first,
                  second
                ) => {
                  if (
                    first.metaScore !==
                    second.metaScore
                  ) {
                    return (
                      second.metaScore -
                      first.metaScore
                    );
                  }

                  if (
                    first.score !==
                    second.score
                  ) {
                    return (
                      second.score -
                      first.score
                    );
                  }

                  return (
                    first.playerIndex -
                    second.playerIndex
                  );
                }
              )
              .slice(
                0,
                Math.max(
                  1,
                  config
                    .PRINCIPAL_ALTERNATIVES_PER_TYPE ||
                    3
                )
              );

          candidates.forEach(
            candidate => {
              rows.push([
                result.dataset.label,
                typeInfo.name,
                principal.player,
                candidate.player,
                candidate.currentRole.label,
                candidate.score,
                candidate.metaScore,
                candidate.coverage,
                candidate.score -
                  principal.score,
                candidate.metaScore -
                  result.dataset.metaScores[
                    principal.playerIndex
                  ][typeIndex],
                crearObservacionAlternativaPrincipal_(
                  result,
                  typeIndex,
                  principal.playerIndex,
                  candidate.playerIndex,
                  candidate.currentRole
                )
              ]);
            }
          );

          if (
            candidates.length === 0
          ) {
            rows.push([
              result.dataset.label,
              typeInfo.name,
              principal.player,
              "SIN ALTERNATIVA",
              "",
              0,
              0,
              0,
              0,
              0,
              "No hay otro jugador con EX Zone/Weather y fuerza válida para este tipo."
            ]);
          }
        }
      );
    }
  );

  const titleRow = startRow;
  const subtitleRow = startRow + 1;
  const generatedRow = startRow + 2;
  const headerRow = startRow + 3;
  const dataStartRow = startRow + 4;

  sheet
    .getRange(
      titleRow,
      1,
      1,
      headers.length
    )
    .merge()
    .setValue(
      "ALTERNATIVAS A PRINCIPAL"
    )
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#548235")
    .setFontColor("#ffffff");

  sheet
    .getRange(
      subtitleRow,
      1,
      1,
      headers.length
    )
    .merge()
    .setValue(
      `Hasta ${config.PRINCIPAL_ALTERNATIVES_PER_TYPE || 3} candidatos con EX por tipo · tabla informativa`
    )
    .setHorizontalAlignment("center")
    .setFontStyle("italic")
    .setBackground("#e2f0d9");

  sheet
    .getRange(
      generatedRow,
      1,
      1,
      headers.length
    )
    .merge()
    .setValue(
      `Generado: ${generatedAt}`
    )
    .setHorizontalAlignment("center")
    .setFontColor("#666666");

  sheet
    .getRange(
      headerRow,
      1,
      1,
      headers.length
    )
    .setValues([
      headers
    ])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#70ad47")
    .setFontColor("#ffffff");

  sheet
    .getRange(
      dataStartRow,
      1,
      rows.length,
      headers.length
    )
    .setValues(rows)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true
    );

  [
    6,
    7,
    9,
    10
  ].forEach(
    column =>
      sheet
        .getRange(
          dataStartRow,
          column,
          rows.length,
          1
        )
        .setNumberFormat(
          "0.00"
        )
  );

  sheet
    .getRange(
      dataStartRow,
      8,
      rows.length,
      1
    )
    .setNumberFormat(
      "0.0%"
    );

  sheet
    .getRange(
      dataStartRow,
      11,
      rows.length,
      1
    )
    .setWrap(true);

  sheet
    .getRange(
      dataStartRow,
      3,
      rows.length,
      2
    )
    .setBackground(
      "#e2f0d9"
    );

  return (
    dataStartRow +
    rows.length
  );
}


function obtenerRolActualAlternativa_(
  result,
  playerIndex
) {
  const assignments =
    result.solution
      .assignmentsByPlayer[
        playerIndex
      ] || [];

  if (
    assignments.length === 0
  ) {
    return {
      role: "SIN_ASIGNAR",
      typeIndexes: [],
      label: "SIN ASIGNAR"
    };
  }

  const role =
    assignments[0].role;

  const roleLabel =
    role === "PRINCIPAL"
      ? "PRINCIPAL"
      : role === "APOYO_1"
        ? "APOYO 1"
        : "APOYO 2";

  return {
    role,
    typeIndexes:
      assignments.map(
        assignment =>
          assignment.typeIndex
      ),

    label:
      `${roleLabel}: ` +
      assignments
        .map(
          assignment =>
            assignment.type
        )
        .join(", ")
  };
}


function crearObservacionAlternativaPrincipal_(
  result,
  targetTypeIndex,
  currentPrincipalIndex,
  alternativePlayerIndex,
  currentRole
) {
  if (
    currentRole.role ===
    "PRINCIPAL"
  ) {
    const sourceTypeIndex =
      currentRole.typeIndexes[0];

    const reciprocalSwapIsValid =
      result.dataset
        .principalEligibility[
          currentPrincipalIndex
        ][sourceTypeIndex];

    if (
      reciprocalSwapIsValid
    ) {
      return (
        `Intercambio directo posible con el Principal de ` +
        `${result.dataset.types[sourceTypeIndex].name}.`
      );
    }

    return (
      `Puede ser Principal de este tipo, pero el Principal actual no ` +
      `puede cubrir directamente ${result.dataset.types[sourceTypeIndex].name}; ` +
      `requiere reajustar varios Principales.`
    );
  }

  if (
    currentRole.role ===
    "APOYO_1"
  ) {
    return (
      "Candidato válido con EX; usarlo como Principal requiere recalcular los grupos de Principal y Apoyo 1."
    );
  }

  if (
    currentRole.role ===
    "APOYO_2"
  ) {
    return (
      "Candidato válido con EX; usarlo como Principal requiere recalcular los tres grupos."
    );
  }

  return (
    "Candidato válido con EX; requiere recalcular la asignación."
  );
}


function escribirTablaAlternativas_(
  sheet,
  startRow,
  modeResults,
  generatedAt,
  config
) {
  const headers = [
    "TABLA",
    "TIPO",
    "COB. ACTUAL",
    "COB. ALTERNATIVA",
    "DIFERENCIA",
    "PRINCIPAL ACTUAL",
    "PRINCIPAL ALTERNATIVO",
    "TRÍO ACTUAL",
    "TRÍO ALTERNATIVO",
    "MAYOR BAJADA EN OTRO TIPO",
    "OBSERVACIÓN"
  ];

  const rows = [];

  modeResults.forEach(
    result => {
      const weakestTypes =
        result.evaluation.finalTotalCoverages
          .map(
            (
              coverage,
              typeIndex
            ) => ({
              coverage,
              typeIndex
            })
          )
          .sort(
            (
              first,
              second
            ) =>
              first.coverage -
              second.coverage
          )
          .slice(
            0,
            config.ALTERNATIVE_TYPES_TO_SHOW
          );

      weakestTypes.forEach(
        weakType => {
          let bestAlternative = null;

          result.alternatives.forEach(
            alternative => {
              const alternativeCoverage =
                alternative.evaluation.finalTotalCoverages[
                  weakType.typeIndex
                ];

              if (
                alternativeCoverage <=
                weakType.coverage
              ) {
                return;
              }

              const drops =
                result.evaluation.finalTotalCoverages.map(
                  (
                    currentCoverage,
                    typeIndex
                  ) =>
                    currentCoverage -
                    alternative.evaluation.finalTotalCoverages[
                      typeIndex
                    ]
                );

              const maximumDrop =
                Math.max(
                  0,
                  ...drops
                );

              if (
                !bestAlternative ||
                alternativeCoverage >
                  bestAlternative.coverage ||
                (
                  alternativeCoverage ===
                    bestAlternative.coverage &&
                  maximumDrop <
                    bestAlternative.maximumDrop
                )
              ) {
                bestAlternative = {
                  alternative,
                  coverage:
                    alternativeCoverage,
                  maximumDrop
                };
              }
            }
          );

          if (!bestAlternative) {
            rows.push([
              result.dataset.label,
              result.dataset.types[
                weakType.typeIndex
              ].name,
              weakType.coverage,
              weakType.coverage,
              0,
              obtenerNombreRol_(
                result.solution,
                weakType.typeIndex,
                "PRINCIPAL"
              ),
              "—",
              obtenerTrioTexto_(
                result.solution,
                weakType.typeIndex
              ),
              "—",
              0,
              "No se encontró otra solución completa que mejore este tipo"
            ]);

            return;
          }

          rows.push([
            result.dataset.label,
            result.dataset.types[
              weakType.typeIndex
            ].name,
            weakType.coverage,
            bestAlternative.coverage,
            bestAlternative.coverage -
              weakType.coverage,
            obtenerNombreRol_(
              result.solution,
              weakType.typeIndex,
              "PRINCIPAL"
            ),
            obtenerNombreRol_(
              bestAlternative.alternative.solution,
              weakType.typeIndex,
              "PRINCIPAL"
            ),
            obtenerTrioTexto_(
              result.solution,
              weakType.typeIndex
            ),
            obtenerTrioTexto_(
              bestAlternative.alternative.solution,
              weakType.typeIndex
            ),
            bestAlternative.maximumDrop,
            "Alternativa completa válida; no se aplica automáticamente"
          ]);
        }
      );
    }
  );

  const titleRow = startRow;
  const subtitleRow = startRow + 1;
  const generatedRow = startRow + 2;
  const headerRow = startRow + 3;
  const dataStartRow = startRow + 4;

  sheet.getRange(
    titleRow,
    1,
    1,
    headers.length
  )
    .merge()
    .setValue(
      "ALTERNATIVAS CERCANAS ENCONTRADAS"
    )
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#375623")
    .setFontColor("#ffffff");

  sheet.getRange(
    subtitleRow,
    1,
    1,
    headers.length
  )
    .merge()
    .setValue(
      "Otras soluciones completas que respetan EX y roles bloqueados"
    )
    .setHorizontalAlignment("center")
    .setFontStyle("italic")
    .setBackground("#e2f0d9");

  sheet.getRange(
    generatedRow,
    1,
    1,
    headers.length
  )
    .merge()
    .setValue(
      `Generado: ${generatedAt}`
    )
    .setHorizontalAlignment("center")
    .setFontColor("#666666");

  sheet.getRange(
    headerRow,
    1,
    1,
    headers.length
  )
    .setValues([
      headers
    ])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#70ad47")
    .setFontColor("#ffffff");

  sheet.getRange(
    dataStartRow,
    1,
    rows.length,
    headers.length
  )
    .setValues(rows)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true
    );

  sheet.getRange(
    dataStartRow,
    3,
    rows.length,
    3
  ).setNumberFormat(
    "0.0%"
  );

  sheet.getRange(
    dataStartRow,
    10,
    rows.length,
    1
  ).setNumberFormat(
    "0.0%"
  );

  sheet.getRange(
    dataStartRow,
    8,
    rows.length,
    4
  ).setWrap(true);
}


function obtenerNombreRol_(
  solution,
  typeIndex,
  role
) {
  const assignment =
    solution.assignmentsByType[
      typeIndex
    ].find(
      item =>
        item.role === role
    );

  return assignment
    ? assignment.player
    : "—";
}


function obtenerTrioTexto_(
  solution,
  typeIndex
) {
  return solution.assignmentsByType[
    typeIndex
  ]
    .map(
      assignment =>
        assignment.player
    )
    .join(" / ");
}


function obtenerSpreadsheet_(config) {
  if (config.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(config.SPREADSHEET_ID);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      "No se ha podido acceder al Google Sheets. " +
      "Vincula el Apps Script al archivo o configura SPREADSHEET_ID."
    );
  }

  return spreadsheet;
}


/** Valida que existan exactamente 20 jugadores. */
function validarCapacidadInicial_(players, config) {
  if (
    config.PRINCIPAL_COUNT !==
      config.NUMBER_OF_TYPES ||
    config.SUPPORT1_COUNT !==
      config.NUMBER_OF_TYPES
  ) {
    throw new Error(
      "La configuración de roles no es válida: debe existir exactamente " +
      "un Principal y un Apoyo 1 por cada tipo."
    );
  }

  if (
    config.SUPPORT2_ASSIGNMENTS_PER_PLAYER <= 0 ||
    config.NUMBER_OF_TYPES %
      config.SUPPORT2_ASSIGNMENTS_PER_PLAYER !== 0
  ) {
    throw new Error(
      "La configuración de Apoyo 2 no es válida: " +
      "NUMBER_OF_TYPES debe ser divisible exactamente entre " +
      "SUPPORT2_ASSIGNMENTS_PER_PLAYER."
    );
  }

  const support2Players =
    config.NUMBER_OF_TYPES /
    config.SUPPORT2_ASSIGNMENTS_PER_PLAYER;

  const expectedPlayers =
    config.PRINCIPAL_COUNT +
    config.SUPPORT1_COUNT +
    support2Players;

  if (players.length !== expectedPlayers) {
    throw new Error(
      `Esta implementación necesita exactamente ${expectedPlayers} jugadores.\n\n` +
      `${config.PRINCIPAL_COUNT} Principales\n` +
      `${config.SUPPORT1_COUNT} Apoyos 1\n` +
      `${support2Players} jugadores de Apoyo 2\n\n` +
      `Actualmente hay ${players.length}.`
    );
  }
}


/** Lee Gym Members. */
function obtenerJugadores_(sheet, config) {
  const lastRow = sheet.getLastRow();

  if (lastRow < config.MEMBERS_START_ROW) {
    return [];
  }

  const values = sheet
    .getRange(
      config.MEMBERS_START_ROW,
      config.MEMBERS_COLUMN,
      lastRow - config.MEMBERS_START_ROW + 1,
      1
    )
    .getDisplayValues()
    .flat();

  const players = [];
  const seen = new Set();

  values.forEach(value => {
    const player = limpiarTexto_(value);

    if (!player) {
      return;
    }

    const key = normalizarTexto_(player);

    if (seen.has(key)) {
      throw new Error(
        `El jugador "${player}" aparece repetido en Gym Members.`
      );
    }

    seen.add(key);
    players.push(player);
  });

  return players;
}


/** Lee los rankings globales de Guia Del Evento. */
function leerClasificaciones_(sheet, players, config) {
  const playerIndexByName = new Map();

  players.forEach((player, index) => {
    playerIndexByName.set(normalizarTexto_(player), index);
  });

  const scores = Array.from(
    { length: players.length },
    () => Array(config.NUMBER_OF_TYPES).fill(0)
  );

  const types = [];
  const errors = [];

  for (
    let typeIndex = 0;
    typeIndex < config.NUMBER_OF_TYPES;
    typeIndex++
  ) {
    const startColumn =
      config.FIRST_TYPE_COLUMN +
      typeIndex * config.COLUMNS_PER_TYPE;

    const typeCell = sheet.getRange(config.TYPE_ROW, startColumn);
    const typeName = limpiarTexto_(typeCell.getDisplayValue());

    if (!typeName) {
      errors.push(
        `No se ha encontrado el tipo en ${typeCell.getA1Notation()}.`
      );
      types.push({ name: `Tipo ${typeIndex + 1}` });
      continue;
    }

    const ranking = sheet
      .getRange(
        config.RANKING_START_ROW,
        startColumn,
        players.length,
        2
      )
      .getValues();

    const found = new Set();

    ranking.forEach(row => {
      const playerName = limpiarTexto_(row[0]);

      if (!playerName) {
        return;
      }

      const playerKey = normalizarTexto_(playerName);

      if (!playerIndexByName.has(playerKey)) {
        errors.push(
          `El jugador "${playerName}" aparece en ${typeName}, ` +
          `pero no está en Gym Members.`
        );
        return;
      }

      if (found.has(playerKey)) {
        errors.push(
          `El jugador "${playerName}" aparece más de una vez en ${typeName}.`
        );
        return;
      }

      found.add(playerKey);
      scores[playerIndexByName.get(playerKey)][typeIndex] =
        convertirNumero_(row[1]);
    });

    players.forEach(player => {
      if (!found.has(normalizarTexto_(player))) {
        errors.push(
          `El jugador "${player}" no aparece en la clasificación de ${typeName}.`
        );
      }
    });

    types.push({ name: typeName });
  }

  if (errors.length > 0) {
    throw new Error(
      "Se han encontrado problemas en Guia Del Evento:\n\n" +
      errors.slice(0, 50).join("\n")
    );
  }

  return {
    types,
    scores
  };
}


/**
 * Normaliza la puntuación de cada jugador contra el máximo
 * individual teórico de su tipo y modalidad.
 */
function normalizarPuntuaciones_(scores, types, config) {
  return scores.map(playerScores =>
    playerScores.map((score, typeIndex) => {
      const maximumIndividualScore =
        types[typeIndex]
          .maximumIndividualScore;

      if (
        maximumIndividualScore <= 0
      ) {
        return 0;
      }

      return Math.round(
        Math.min(
          1,
          score /
            maximumIndividualScore
        ) *
        config.NORMALIZATION_SCALE
      );
    })
  );
}


/** Minimum Cost Maximum Flow. */
class MinCostFlow_ {
  constructor(numberOfNodes) {
    this.graph = Array.from(
      { length: numberOfNodes },
      () => []
    );
  }

  addEdge(from, to, capacity, cost) {
    const forwardIndex = this.graph[from].length;
    const reverseIndex = this.graph[to].length;

    this.graph[from].push({
      to,
      reverseIndex,
      capacity,
      initialCapacity: capacity,
      cost
    });

    this.graph[to].push({
      to: from,
      reverseIndex: forwardIndex,
      capacity: 0,
      initialCapacity: 0,
      cost: -cost
    });

    return {
      from,
      edgeIndex: forwardIndex
    };
  }

  minCostMaxFlow(source, sink, requiredFlow) {
    let totalFlow = 0;
    let totalCost = 0;
    const numberOfNodes = this.graph.length;

    while (totalFlow < requiredFlow) {
      const distance = Array(numberOfNodes).fill(Infinity);
      const previousNode = Array(numberOfNodes).fill(-1);
      const previousEdge = Array(numberOfNodes).fill(-1);
      const inQueue = Array(numberOfNodes).fill(false);
      const queue = [source];
      let queuePosition = 0;

      distance[source] = 0;
      inQueue[source] = true;

      while (queuePosition < queue.length) {
        const currentNode = queue[queuePosition++];
        inQueue[currentNode] = false;

        for (
          let edgeIndex = 0;
          edgeIndex < this.graph[currentNode].length;
          edgeIndex++
        ) {
          const edge = this.graph[currentNode][edgeIndex];

          if (edge.capacity <= 0) {
            continue;
          }

          const newDistance = distance[currentNode] + edge.cost;

          if (newDistance < distance[edge.to]) {
            distance[edge.to] = newDistance;
            previousNode[edge.to] = currentNode;
            previousEdge[edge.to] = edgeIndex;

            if (!inQueue[edge.to]) {
              queue.push(edge.to);
              inQueue[edge.to] = true;
            }
          }
        }
      }

      if (previousNode[sink] === -1) {
        break;
      }

      let flowToAdd = requiredFlow - totalFlow;
      let currentNode = sink;

      while (currentNode !== source) {
        const parentNode = previousNode[currentNode];
        const edgeIndex = previousEdge[currentNode];
        const edge = this.graph[parentNode][edgeIndex];

        flowToAdd = Math.min(flowToAdd, edge.capacity);
        currentNode = parentNode;
      }

      currentNode = sink;

      while (currentNode !== source) {
        const parentNode = previousNode[currentNode];
        const edgeIndex = previousEdge[currentNode];
        const edge = this.graph[parentNode][edgeIndex];
        const reverseEdge =
          this.graph[currentNode][edge.reverseIndex];

        edge.capacity -= flowToAdd;
        reverseEdge.capacity += flowToAdd;
        currentNode = parentNode;
      }

      totalFlow += flowToAdd;
      totalCost += flowToAdd * distance[sink];
    }

    return {
      flow: totalFlow,
      cost: totalCost
    };
  }
}


/** Potencia media normalizada de un jugador en los 8 tipos. */
function calcularPotenciaGlobal_(normalizedScores, config) {
  return normalizedScores.map(playerScores => {
    if (playerScores.length === 0) {
      return 0;
    }

    return (
      playerScores.reduce((total, score) => total + score, 0) /
      playerScores.length /
      config.NORMALIZATION_SCALE
    );
  });
}


function crearMatrizBooleana_(rows, columns, initialValue) {
  return Array.from(
    { length: rows },
    () => Array(columns).fill(initialValue)
  );
}


function crearMatrizDeArrays_(rows, columns) {
  return Array.from(
    { length: rows },
    () => Array.from({ length: columns }, () => [])
  );
}


function clonarSolucionCompactaCompleta_(
  solution
) {
  return {
    principalsByType:
      solution
        .principalsByType
        .slice(),

    support1ByType:
      solution
        .support1ByType
        .slice(),

    support2ByType:
      solution
        .support2ByType
        .slice(),

    greedySteps:
      solution.greedySteps
        ? {
            principal:
              (
                solution
                  .greedySteps
                  .principal || []
              ).slice(),

            support1:
              (
                solution
                  .greedySteps
                  .support1 || []
              ).slice(),

            support2:
              (
                solution
                  .greedySteps
                  .support2 || []
              ).slice()
          }
        : {
            principal: [],
            support1: [],
            support2: []
          },

    rebalancingActions:
      (
        solution
          .rebalancingActions || []
      ).slice(),

    multiStartSummary:
      solution.multiStartSummary ||
      null
  };
}


function clonarSolucionCompacta_(solution) {
  return {
    principalsByType: solution.principalsByType.slice(),
    support1ByType: solution.support1ByType.slice(),
    support2ByType: solution.support2ByType.slice()
  };
}


function esValorActivo_(value, configuredYesValue) {
  const normalizedValue = normalizarTexto_(value);

  return new Set([
    normalizarTexto_(configuredYesValue),
    "yes",
    "si",
    "sí",
    "true",
    "1",
    "x"
  ]).has(normalizedValue);
}


function limpiarTexto_(value) {
  return String(value == null ? "" : value)
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}


function normalizarTexto_(value) {
  return limpiarTexto_(value).toLowerCase();
}


function convertirNumero_(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  let text = limpiarTexto_(value);

  if (!text) {
    return 0;
  }

  if (text.includes(",")) {
    text = text.replace(/\./g, "").replace(",", ".");
  }

  const result = Number(text);
  return Number.isFinite(result) ? result : 0;
}

function generarRankingFuerzaTipos() {

  const CONFIG = {
    MEMBERS_SHEET: "Gym Members",
    ROSTER_SHEET: "Roster",
    META_SHEET: "Meta",
    SYNC_PAIRS_SHEET: "SyncPairs",
    OUTPUT_SHEET: "Type Strength",

    MEMBERS_RANGE: "A1:A20",

    // true: los tipos más débiles aparecen primero.
    // false: los tipos más fuertes aparecen primero.
    WEAKEST_FIRST: true,

    /*
     * COLUMNAS DE ROSTER
     *
     * A = Player
     * B = Sync Pair
     * C = Type
     * D = Role / categoría física o especial
     * N = Strength Score
     * P = Max Potential Score
     */
    ROSTER_COLUMNS: {
      PLAYER: 1,
      SYNC_PAIR: 2,
      TYPE: 3,
      CATEGORY: 4,
      CURRENT_SCORE: 14,
      MAX_SCORE: 16
    },

    /*
     * COLUMNAS DE META
     *
     * A = Sync Pair
     * B = Type
     */
    META_COLUMNS: {
      SYNC_PAIR: 1,
      TYPE: 2
    },

    /*
     * COLUMNAS DE SYNC PAIRS
     *
     * A = Sync Pair
     * C = Role / categoría física o especial
     *
     * Esta hoja se utiliza para determinar si una unidad
     * es física o especial, incluso si ningún miembro la tiene.
     */
    SYNC_PAIR_COLUMNS: {
      SYNC_PAIR: 1,
      CATEGORY: 4
    }
  };


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalSheet = ss.getActiveSheet();
  const originalRange = ss.getActiveRange();

  const membersSheet = ss.getSheetByName(CONFIG.MEMBERS_SHEET);
  const rosterSheet = ss.getSheetByName(CONFIG.ROSTER_SHEET);
  const metaSheet = ss.getSheetByName(CONFIG.META_SHEET);
  const syncPairsSheet = ss.getSheetByName(CONFIG.SYNC_PAIRS_SHEET);

  if (!membersSheet) {
    throw new Error(`No existe la hoja "${CONFIG.MEMBERS_SHEET}".`);
  }

  if (!rosterSheet) {
    throw new Error(`No existe la hoja "${CONFIG.ROSTER_SHEET}".`);
  }

  if (!metaSheet) {
    throw new Error(`No existe la hoja "${CONFIG.META_SHEET}".`);
  }


  // ============================================================
  // 1. MIEMBROS DEL GIMNASIO
  // ============================================================

  const members = membersSheet
    .getRange(CONFIG.MEMBERS_RANGE)
    .getDisplayValues()
    .flat()
    .map(value => String(value).trim())
    .filter(value => value !== "");

  if (members.length === 0) {
    throw new Error(
      `No hay miembros definidos en ${CONFIG.MEMBERS_SHEET}!${CONFIG.MEMBERS_RANGE}.`
    );
  }

  const memberKeys = members.map(rankingTipos_normalize_);
  const validMembers = new Set(memberKeys);


  // ============================================================
  // 2. DATOS DEL ROSTER
  // ============================================================

  const rosterLastRow = rosterSheet.getLastRow();

  const rosterData = rosterLastRow >= 2
    ? rosterSheet
        .getRange(
          2,
          1,
          rosterLastRow - 1,
          CONFIG.ROSTER_COLUMNS.MAX_SCORE
        )
        .getValues()
    : [];


  // ============================================================
  // 3. MAPA DE CATEGORÍA FÍSICA / ESPECIAL
  // ============================================================

  const categoryByPair = new Map();


  /*
   * Primero se obtiene la categoría desde SyncPairs.
   * Esto permite conocer la categoría de una unidad meta
   * aunque ningún miembro del gimnasio la tenga.
   */
  if (syncPairsSheet && syncPairsSheet.getLastRow() >= 2) {

    const syncLastRow = syncPairsSheet.getLastRow();

    const syncData = syncPairsSheet
      .getRange(
        2,
        1,
        syncLastRow - 1,
        Math.max(
          CONFIG.SYNC_PAIR_COLUMNS.SYNC_PAIR,
          CONFIG.SYNC_PAIR_COLUMNS.CATEGORY
        )
      )
      .getDisplayValues();

    syncData.forEach(row => {

      const pair = row[
        CONFIG.SYNC_PAIR_COLUMNS.SYNC_PAIR - 1
      ];

      const category = row[
        CONFIG.SYNC_PAIR_COLUMNS.CATEGORY - 1
      ];

      const pairKey = rankingTipos_normalize_(pair);

      if (pairKey && category) {
        categoryByPair.set(pairKey, String(category).trim());
      }
    });
  }


  /*
   * Si la categoría no estaba en SyncPairs,
   * se intenta recuperar desde Roster.
   */
  rosterData.forEach(row => {

    const pair = row[
      CONFIG.ROSTER_COLUMNS.SYNC_PAIR - 1
    ];

    const category = row[
      CONFIG.ROSTER_COLUMNS.CATEGORY - 1
    ];

    const pairKey = rankingTipos_normalize_(pair);

    if (
      pairKey &&
      category &&
      !categoryByPair.has(pairKey)
    ) {
      categoryByPair.set(pairKey, String(category).trim());
    }
  });


  // ============================================================
  // 4. UNIDADES META POR TIPO
  // ============================================================

  const metaLastRow = metaSheet.getLastRow();

  const metaData = metaLastRow >= 2
    ? metaSheet
        .getRange(
          2,
          1,
          metaLastRow - 1,
          Math.max(
            CONFIG.META_COLUMNS.SYNC_PAIR,
            CONFIG.META_COLUMNS.TYPE
          )
        )
        .getDisplayValues()
    : [];

  /*
   * Estructura:
   *
   * metaPairsByType:
   *   typeKey -> Map(pairKey -> nombre de la unidad)
   */
  const metaPairsByType = new Map();
  const typeNames = new Map();

  metaData.forEach(row => {

    const pair = String(
      row[CONFIG.META_COLUMNS.SYNC_PAIR - 1] || ""
    ).trim();

    const type = String(
      row[CONFIG.META_COLUMNS.TYPE - 1] || ""
    ).trim();

    if (!pair || !type) return;

    const pairKey = rankingTipos_normalize_(pair);
    const typeKey = rankingTipos_normalize_(type);

    typeNames.set(typeKey, type);

    if (!metaPairsByType.has(typeKey)) {
      metaPairsByType.set(typeKey, new Map());
    }

    metaPairsByType
      .get(typeKey)
      .set(pairKey, pair);
  });


  if (metaPairsByType.size === 0) {
    throw new Error(
      `No se han encontrado unidades meta en la hoja "${CONFIG.META_SHEET}".`
    );
  }


  // ============================================================
  // 5. ELIMINAR POSIBLES DUPLICADOS DEL ROSTER
  // ============================================================

  /*
   * Si por algún motivo existe la misma unidad más de una vez
   * para un jugador, solamente se conserva la fila con mayor
   * Strength Score.
   */
  const bestRosterEntry = new Map();

  rosterData.forEach(row => {

    const player = row[
      CONFIG.ROSTER_COLUMNS.PLAYER - 1
    ];

    const pair = row[
      CONFIG.ROSTER_COLUMNS.SYNC_PAIR - 1
    ];

    const type = row[
      CONFIG.ROSTER_COLUMNS.TYPE - 1
    ];

    const category = row[
      CONFIG.ROSTER_COLUMNS.CATEGORY - 1
    ];

    const currentScore = rankingTipos_toNumber_(
      row[CONFIG.ROSTER_COLUMNS.CURRENT_SCORE - 1]
    );

    const maxScore = rankingTipos_toNumber_(
      row[CONFIG.ROSTER_COLUMNS.MAX_SCORE - 1]
    );

    const playerKey = rankingTipos_normalize_(player);
    const pairKey = rankingTipos_normalize_(pair);
    const typeKey = rankingTipos_normalize_(type);

    if (
      !playerKey ||
      !pairKey ||
      !typeKey ||
      !validMembers.has(playerKey)
    ) {
      return;
    }

    const uniqueKey = `${playerKey}|||${pairKey}`;

    const currentEntry = {
      playerKey,
      pairKey,
      typeKey,
      typeName: String(type).trim(),
      category: String(category || "").trim(),
      currentScore,
      maxScore
    };

    const previousEntry = bestRosterEntry.get(uniqueKey);

    if (
      !previousEntry ||
      currentScore > previousEntry.currentScore
    ) {
      bestRosterEntry.set(uniqueKey, currentEntry);
    }
  });


  // ============================================================
  // 6. MÁXIMO POTENCIAL DE CADA UNIDAD
  // ============================================================

  /*
   * Como P representa el máximo teórico de la unidad,
   * debe ser igual para todos los miembros.
   *
   * Se utiliza el valor más alto encontrado por seguridad.
   */
  const maxPotentialByPair = new Map();

  bestRosterEntry.forEach(entry => {

    if (entry.maxScore <= 0) return;

    const previousMax =
      maxPotentialByPair.get(entry.pairKey) || 0;

    if (entry.maxScore > previousMax) {
      maxPotentialByPair.set(
        entry.pairKey,
        entry.maxScore
      );
    }
  });


  // ============================================================
  // 7. AGRUPAR PUNTUACIONES POR JUGADOR, TIPO Y CATEGORÍA
  // ============================================================

  /*
   * scoreGroups:
   *
   * jugador | tipo | global
   * jugador | tipo | physical
   * jugador | tipo | special
   */
  const scoreGroups = new Map();

  bestRosterEntry.forEach(entry => {

    rankingTipos_addScore_(
      scoreGroups,
      entry.playerKey,
      entry.typeKey,
      "global",
      entry.currentScore
    );

    const categoryText =
      categoryByPair.get(entry.pairKey) || "";

    const category =
      rankingTipos_getCategory_(categoryText);

    if (
      category === "physical" ||
      category === "hybrid"
    ) {
      rankingTipos_addScore_(
        scoreGroups,
        entry.playerKey,
        entry.typeKey,
        "physical",
        entry.currentScore
      );
    }

    if (
      category === "special" ||
      category === "hybrid"
    ) {
      rankingTipos_addScore_(
        scoreGroups,
        entry.playerKey,
        entry.typeKey,
        "special",
        entry.currentScore
      );
    }
  });


  // ============================================================
  // 8. CALCULAR LOS TRES RANKINGS
  // ============================================================

  const globalRanking = rankingTipos_buildRanking_({
    category: "global",
    members: memberKeys,
    metaPairsByType,
    typeNames,
    categoryByPair,
    maxPotentialByPair,
    scoreGroups,
    weakestFirst: CONFIG.WEAKEST_FIRST
  });

  const physicalRanking = rankingTipos_buildRanking_({
    category: "physical",
    members: memberKeys,
    metaPairsByType,
    typeNames,
    categoryByPair,
    maxPotentialByPair,
    scoreGroups,
    weakestFirst: CONFIG.WEAKEST_FIRST
  });

  const specialRanking = rankingTipos_buildRanking_({
    category: "special",
    members: memberKeys,
    metaPairsByType,
    typeNames,
    categoryByPair,
    maxPotentialByPair,
    scoreGroups,
    weakestFirst: CONFIG.WEAKEST_FIRST
  });


  // ============================================================
  // 9. CREAR / LIMPIAR HOJA DE RESULTADOS
  // ============================================================

  let outputSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET);

  if (!outputSheet) {
    outputSheet = ss.insertSheet(CONFIG.OUTPUT_SHEET);
  }

  const requiredColumns = 26;
  const requiredRows = 40;

  if (outputSheet.getMaxColumns() < requiredColumns) {
    outputSheet.insertColumnsAfter(
      outputSheet.getMaxColumns(),
      requiredColumns - outputSheet.getMaxColumns()
    );
  }

  if (outputSheet.getMaxRows() < requiredRows) {
    outputSheet.insertRowsAfter(
      outputSheet.getMaxRows(),
      requiredRows - outputSheet.getMaxRows()
    );
  }

  outputSheet
    .getRange(1, 1, requiredRows, requiredColumns)
    .breakApart();

  outputSheet.clear();
  outputSheet.setConditionalFormatRules([]);
  outputSheet.setFrozenRows(2);
  outputSheet.setHiddenGridlines(true);


  // ============================================================
  // 10. ESCRIBIR LAS TRES TABLAS
  // ============================================================

  rankingTipos_writeTable_(
    outputSheet,
    1,
    "FUERZA GLOBAL",
    globalRanking,
    "#1F4E78"
  );

  rankingTipos_writeTable_(
    outputSheet,
    10,
    "FUERZA FÍSICA",
    physicalRanking,
    "#843C0C"
  );

  rankingTipos_writeTable_(
    outputSheet,
    19,
    "FUERZA ESPECIAL",
    specialRanking,
    "#7030A0"
  );


  // ============================================================
  // 11. INFORMACIÓN Y ADVERTENCIAS
  // ============================================================

  const maximumTableLength = Math.max(
    globalRanking.length,
    physicalRanking.length,
    specialRanking.length
  );

  let noteRow = maximumTableLength + 5;

  const orderText = CONFIG.WEAKEST_FIRST
    ? "Los tipos aparecen ordenados del más débil al más fuerte."
    : "Los tipos aparecen ordenados del más fuerte al más débil.";

  const explanation =
    `Cálculo: cada miembro aporta sus mejores unidades hasta alcanzar ` +
    `el número de unidades meta de ese tipo y categoría. ` +
    `El máximo representa que los ${members.length} miembros tienen ` +
    `todas las unidades meta maximizadas. ${orderText}`;

  outputSheet
    .getRange(noteRow, 1, 1, 26)
    .merge()
    .setValue(explanation)
    .setWrap(true)
    .setFontStyle("italic")
    .setBackground("#EAF2F8")
    .setVerticalAlignment("middle");

  outputSheet.setRowHeight(noteRow, 45);

  noteRow += 2;


  const missingMaximums = new Set();

  metaPairsByType.forEach(pairMap => {
    pairMap.forEach((pairName, pairKey) => {
      if (!maxPotentialByPair.has(pairKey)) {
        missingMaximums.add(pairName);
      }
    });
  });

  if (missingMaximums.size > 0) {

    outputSheet
      .getRange(noteRow, 1, 1, 26)
      .merge()
      .setValue(
        "⚠ No se ha encontrado Max Potential Score en Roster!P " +
        "para estas unidades meta: " +
        Array.from(missingMaximums).join(", ")
      )
      .setWrap(true)
      .setBackground("#FFF2CC")
      .setFontColor("#9C6500");

    outputSheet.setRowHeight(noteRow, 55);

    noteRow += 2;
  }


  const uncategorizedMeta = new Set();

  metaPairsByType.forEach(pairMap => {

    pairMap.forEach((pairName, pairKey) => {

      const category = rankingTipos_getCategory_(
        categoryByPair.get(pairKey) || ""
      );

      if (category === "other") {
        uncategorizedMeta.add(pairName);
      }
    });
  });

  if (uncategorizedMeta.size > 0) {

    outputSheet
      .getRange(noteRow, 1, 1, 26)
      .merge()
      .setValue(
        "⚠ Estas unidades meta no contienen Physical/Físico o " +
        "Special/Especial en la columna de categoría y, por tanto, " +
        "solo aparecen en la tabla global: " +
        Array.from(uncategorizedMeta).join(", ")
      )
      .setWrap(true)
      .setBackground("#FCE4D6")
      .setFontColor("#9C0006");

    outputSheet.setRowHeight(noteRow, 65);
  }


  outputSheet
    .getRange(1, 1)
    .setNote(
      `Actualizado: ${Utilities.formatDate(
        new Date(),
        ss.getSpreadsheetTimeZone(),
        "dd/MM/yyyy HH:mm:ss"
      )}`
    );

  SpreadsheetApp.flush();

  // Volver a la hoja desde la que se pulsó el botón
  originalSheet.activate();

  if (
    originalRange &&
    originalRange.getSheet().getSheetId() === originalSheet.getSheetId()
  ) {
    originalSheet.setActiveRange(originalRange);
  }

  SpreadsheetApp.flush();

  ss.toast(
    "Ranking de fuerza por tipos actualizado correctamente.",
    "Ranking de tipos",
    5
  );
}


// ============================================================================
// CONSTRUIR UN RANKING
// ============================================================================

function rankingTipos_buildRanking_(options) {

  const {
    category,
    members,
    metaPairsByType,
    typeNames,
    categoryByPair,
    maxPotentialByPair,
    scoreGroups,
    weakestFirst
  } = options;

  const results = [];

  metaPairsByType.forEach((pairMap, typeKey) => {

    const allMetaPairs = Array.from(pairMap.entries());

    let selectedMetaPairs;

    if (category === "global") {

      selectedMetaPairs = allMetaPairs;

    } else {

      selectedMetaPairs = allMetaPairs.filter(([pairKey]) => {

        const pairCategory = rankingTipos_getCategory_(
          categoryByPair.get(pairKey) || ""
        );

        if (category === "physical") {
          return (
            pairCategory === "physical" ||
            pairCategory === "hybrid"
          );
        }

        if (category === "special") {
          return (
            pairCategory === "special" ||
            pairCategory === "hybrid"
          );
        }

        return false;
      });
    }


    /*
     * Número máximo de unidades que puede aportar cada miembro.
     *
     * Global:
     * número total de unidades meta del tipo.
     *
     * Física:
     * número de unidades meta físicas del tipo.
     *
     * Especial:
     * número de unidades meta especiales del tipo.
     */
    const metaCount = selectedMetaPairs.length;

    let actualStrength = 0;

    if (metaCount > 0) {

      members.forEach(memberKey => {

        const groupKey =
          `${memberKey}|||${typeKey}|||${category}`;

        const scores = scoreGroups.get(groupKey) || [];

        const bestScores = scores
          .slice()
          .sort((a, b) => b - a)
          .slice(0, metaCount);

        actualStrength += bestScores.reduce(
          (total, score) => total + score,
          0
        );
      });
    }


    const missingMaxPairs = [];

    let maximumPerMember = 0;

    selectedMetaPairs.forEach(([pairKey, pairName]) => {

      const maximum = maxPotentialByPair.get(pairKey);

      if (maximum === undefined || maximum <= 0) {
        missingMaxPairs.push(pairName);
      } else {
        maximumPerMember += maximum;
      }
    });


    let maximumStrength = null;
    let percentage = null;
    let deficit = null;
    let status = "OK";

    if (metaCount === 0) {

      status = category === "physical"
        ? "Sin unidades meta físicas"
        : category === "special"
          ? "Sin unidades meta especiales"
          : "Sin unidades meta";

    } else if (missingMaxPairs.length > 0) {

      status =
        `Falta máximo de ${missingMaxPairs.length} meta`;

    } else {

      maximumStrength =
        maximumPerMember * members.length;

      if (maximumStrength > 0) {

        percentage =
          actualStrength / maximumStrength;

        deficit =
          maximumStrength - actualStrength;

        if (actualStrength > maximumStrength) {
          status = "Revisar: supera el máximo";
        }
      }
    }


    results.push({
      type: typeNames.get(typeKey) || typeKey,
      actualStrength,
      maximumStrength,
      percentage,
      metaCount,
      deficit,
      status
    });
  });


  results.sort((a, b) => {

    const aValid = a.percentage !== null;
    const bValid = b.percentage !== null;

    /*
     * Los resultados incompletos aparecen siempre al final.
     */
    if (aValid && !bValid) return -1;
    if (!aValid && bValid) return 1;

    if (!aValid && !bValid) {
      return a.type.localeCompare(b.type);
    }

    if (weakestFirst) {

      if (a.percentage !== b.percentage) {
        return a.percentage - b.percentage;
      }

      return a.actualStrength - b.actualStrength;

    } else {

      if (a.percentage !== b.percentage) {
        return b.percentage - a.percentage;
      }

      return b.actualStrength - a.actualStrength;
    }
  });

  return results;
}


// ============================================================================
// ESCRIBIR Y FORMATEAR UNA TABLA
// ============================================================================

function rankingTipos_writeTable_(
  sheet,
  startColumn,
  title,
  ranking,
  titleColor
) {

  const tableWidth = 8;
  const titleRow = 1;
  const headerRow = 2;
  const firstDataRow = 3;

  const titleRange = sheet.getRange(
    titleRow,
    startColumn,
    1,
    tableWidth
  );

  titleRange
    .merge()
    .setValue(title)
    .setBackground(titleColor)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(13)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setRowHeight(titleRow, 28);


  const headers = [[
    "Prioridad",
    "Tipo",
    "Fuerza actual",
    "Fuerza máxima",
    "% del máximo",
    "Nº meta",
    "Déficit",
    "Estado"
  ]];

  const headerRange = sheet.getRange(
    headerRow,
    startColumn,
    1,
    tableWidth
  );

  headerRange
    .setValues(headers)
    .setBackground("#D9E2F3")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#A6A6A6",
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet.setRowHeight(headerRow, 36);


  if (ranking.length === 0) return;


  const values = ranking.map((entry, index) => [
    index + 1,
    entry.type,
    entry.actualStrength,
    entry.maximumStrength === null
      ? ""
      : entry.maximumStrength,
    entry.percentage === null
      ? ""
      : entry.percentage,
    entry.metaCount,
    entry.deficit === null
      ? ""
      : entry.deficit,
    entry.status
  ]);

  const dataRange = sheet.getRange(
    firstDataRow,
    startColumn,
    values.length,
    tableWidth
  );

  dataRange
    .setValues(values)
    .setVerticalAlignment("middle")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#D9D9D9",
      SpreadsheetApp.BorderStyle.SOLID
    );


  sheet
    .getRange(
      firstDataRow,
      startColumn,
      values.length,
      1
    )
    .setHorizontalAlignment("center");

  sheet
    .getRange(
      firstDataRow,
      startColumn + 2,
      values.length,
      2
    )
    .setNumberFormat("#,##0.00");

  sheet
    .getRange(
      firstDataRow,
      startColumn + 4,
      values.length,
      1
    )
    .setNumberFormat("0.0%")
    .setHorizontalAlignment("center");

  sheet
    .getRange(
      firstDataRow,
      startColumn + 5,
      values.length,
      1
    )
    .setNumberFormat("0")
    .setHorizontalAlignment("center");

  sheet
    .getRange(
      firstDataRow,
      startColumn + 6,
      values.length,
      1
    )
    .setNumberFormat("#,##0.00");

  sheet
    .getRange(
      firstDataRow,
      startColumn + 7,
      values.length,
      1
    )
    .setWrap(true);


  /*
   * Colorear el porcentaje:
   *
   * Rojo: menos del 50%
   * Naranja: menos del 70%
   * Amarillo: menos del 85%
   * Verde: 85% o más
   */
  const percentageBackgrounds = ranking.map(entry => {

    if (entry.percentage === null) {
      return ["#E7E6E6"];
    }

    if (entry.percentage < 0.50) {
      return ["#F4CCCC"];
    }

    if (entry.percentage < 0.70) {
      return ["#FCE5CD"];
    }

    if (entry.percentage < 0.85) {
      return ["#FFF2CC"];
    }

    return ["#D9EAD3"];
  });

  sheet
    .getRange(
      firstDataRow,
      startColumn + 4,
      ranking.length,
      1
    )
    .setBackgrounds(percentageBackgrounds)
    .setFontWeight("bold");


  const statusBackgrounds = ranking.map(entry => [
    entry.status === "OK"
      ? "#FFFFFF"
      : "#FFF2CC"
  ]);

  sheet
    .getRange(
      firstDataRow,
      startColumn + 7,
      ranking.length,
      1
    )
    .setBackgrounds(statusBackgrounds);


  // Anchuras de las columnas
  sheet.setColumnWidth(startColumn, 72);
  sheet.setColumnWidth(startColumn + 1, 105);
  sheet.setColumnWidth(startColumn + 2, 110);
  sheet.setColumnWidth(startColumn + 3, 110);
  sheet.setColumnWidth(startColumn + 4, 105);
  sheet.setColumnWidth(startColumn + 5, 75);
  sheet.setColumnWidth(startColumn + 6, 100);
  sheet.setColumnWidth(startColumn + 7, 150);
}


// ============================================================================
// AÑADIR PUNTUACIÓN A UN GRUPO
// ============================================================================

function rankingTipos_addScore_(
  groups,
  playerKey,
  typeKey,
  category,
  score
) {

  const key =
    `${playerKey}|||${typeKey}|||${category}`;

  if (!groups.has(key)) {
    groups.set(key, []);
  }

  groups.get(key).push(score);
}


// ============================================================================
// CLASIFICAR COMO FÍSICO O ESPECIAL
// ============================================================================

function rankingTipos_getCategory_(value) {

  const normalized = rankingTipos_normalize_(value);

  if (
    normalized.includes("hybrid") ||
    normalized.includes("hibrido") ||
    normalized.includes("hibrida")
  ) {
    return "hybrid";
  }

  if (
    normalized.includes("physical") ||
    normalized.includes("fisico") ||
    normalized.includes("fisica")
  ) {
    return "physical";
  }

  if (
    normalized.includes("special") ||
    normalized.includes("especial")
  ) {
    return "special";
  }

  return "other";
}


// ============================================================================
// NORMALIZAR TEXTOS
// ============================================================================

function rankingTipos_normalize_(value) {

  return String(value === null || value === undefined ? "" : value)
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}


// ============================================================================
// CONVERTIR VALOR EN NÚMERO
// ============================================================================

function rankingTipos_toNumber_(value) {

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  const normalized = String(value || "")
    .trim()
    .replace(/\s/g, "")
    .replace(",", ".");

  const number = Number(normalized);

  return Number.isFinite(number) ? number : 0;
}