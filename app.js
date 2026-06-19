/**
 * Application Controller — Office Script Visualizer & Mapper (Gold Edition)
 *
 * Wires ASTEngine + LineageGraph to the DOM.
 * All code analysis delegated to the AST engine — zero regex for logic.
 */

// ─── GLOBAL STATE ─────────────────────────────────────────────────────
const state = {
  activeTab: 'visualizer',
  selectedScript: '',
  currentCode: '',
  sourceFile: null,         // Parsed AST (ts.SourceFile)
  configSAP: null,          // Extracted CONFIG_SAP data, or null
  legacyData: null,         // Extracted legacy headers/sheets, or null
  parsedMetadata: {
    sheets: [],
    columns: [],
    startRow: 0,
    hasStaticRow: false
  },
  columnMapping: {},
  simulation: {
    activeSheet: 'Hoja Activa',
    sheetsData: {},
    running: false,
    step: 0,
    intervalId: null
  }
};

// ─── DOM ELEMENTS ─────────────────────────────────────────────────────
const elements = {
  tabs: document.querySelectorAll('.tab-link'),
  tabContents: document.querySelectorAll('.tab-content'),
  presetList: document.getElementById('preset-list'),
  codeEditor: document.getElementById('code-editor'),
  dissectPanel: document.getElementById('dissect-panel'),
  statSheets: document.getElementById('stat-sheets'),
  statCols: document.getElementById('stat-cols'),
  statStartRow: document.getElementById('stat-start-row'),
  statComplexity: document.getElementById('stat-complexity'),
  statEngine: document.getElementById('stat-engine'),
  mapperList: document.getElementById('mapper-list'),
  btnApplyMapping: document.getElementById('btn-apply-mapping'),
  migrationBanner: document.getElementById('migration-banner'),
  btnMigrate: document.getElementById('btn-migrate'),
  suggestionsGrid: document.getElementById('suggestions-grid'),
  lineageCanvas: document.getElementById('lineage-canvas'),
  simulationGrid: document.getElementById('simulation-grid'),
  btnRunSim: document.getElementById('btn-run-sim'),
  btnResetSim: document.getElementById('btn-reset-sim'),
  simStatus: document.getElementById('sim-status'),
  simSheetTabs: document.getElementById('sim-sheet-tabs'),
  toastOverlay: document.getElementById('toast-overlay')
};

// ─── INITIALIZE ───────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  initPresets();
  initTabNavigation();
  initEditorEvents();
  initSimulation();
  initMigrationButton();
  initMappingButton();
});

// ─── TOAST NOTIFICATIONS ──────────────────────────────────────────────
function showToast(message, type) {
  type = type || 'success';
  const toast = document.createElement('div');
  toast.className = 'toast ' + type;
  toast.innerHTML = '<span style="color:var(--gold)">✦</span> <span>' + message + '</span>';
  elements.toastOverlay.appendChild(toast);
  setTimeout(() => toast.remove(), 3500);
}

// ─── PRESETS ──────────────────────────────────────────────────────────
function initPresets() {
  elements.presetList.innerHTML = '';
  const fileNames = Object.keys(PRESET_SCRIPTS);

  fileNames.forEach((fileName, index) => {
    const li = document.createElement('li');
    li.className = 'preset-item' + (index === 0 ? ' active' : '');
    li.textContent = fileName;
    li.title = fileName;
    li.addEventListener('click', () => selectPreset(fileName, li));
    elements.presetList.appendChild(li);
  });

  if (fileNames.length > 0) loadScript(fileNames[0]);
}

function selectPreset(fileName, element) {
  document.querySelectorAll('.preset-item').forEach(i => i.classList.remove('active'));
  element.classList.add('active');
  loadScript(fileName);
  showToast('Cargado: ' + fileName);
}

function loadScript(fileName) {
  state.selectedScript = fileName;
  state.currentCode = PRESET_SCRIPTS[fileName] || '';
  elements.codeEditor.value = state.currentCode;
  state.columnMapping = {};
  analyzeScript();
  resetSimulation();
}

// ─── TAB NAVIGATION ──────────────────────────────────────────────────
function initTabNavigation() {
  elements.tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const target = tab.dataset.tab;
      state.activeTab = target;
      elements.tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      elements.tabContents.forEach(c => {
        c.classList.toggle('active', c.id === target + '-tab');
      });
      if (target === 'simulation') renderSimulationSheet();
      if (target === 'lineage') renderLineage();
    });
  });
}

// ─── EDITOR ──────────────────────────────────────────────────────────
function initEditorEvents() {
  elements.codeEditor.addEventListener('input', e => {
    state.currentCode = e.target.value;
    analyzeScript();
  });
}

// ─── CORE ANALYSIS (AST-Based) ───────────────────────────────────────

/**
 * Master analysis function.  Delegates entirely to ASTEngine.
 * Called on every script load and every editor keystroke (debounced by browser).
 */
function analyzeScript() {
  const code = state.currentCode;
  if (!code.trim()) {
    state.sourceFile = null;
    state.configSAP = null;
    state.legacyData = null;
    state.parsedMetadata = { sheets: [], columns: [], startRow: 0, hasStaticRow: false };
    refreshAllPanels();
    return;
  }

  try {
    const sf = ASTEngine.parse(code);
    state.sourceFile = sf;

    // Try CONFIG_SAP first, fall back to legacy extraction
    const config = ASTEngine.extractConfigSAP(sf);
    state.configSAP = config;

    if (config && config.data) {
      // CONFIG_SAP mode: extract metadata from the config object
      const cd = config.data;
      const sheets = [];
      if (cd.sheets) {
        for (const key in cd.sheets) {
          if (cd.sheets[key].name) sheets.push(cd.sheets[key].name);
        }
      }

      const columns = [];
      if (cd.columns) {
        for (const key in cd.columns) {
          if (cd.columns[key].header) columns.push(cd.columns[key].header);
        }
      }

      const exec = cd.execution || {};
      state.parsedMetadata = {
        sheets: sheets.length > 0 ? sheets : ['Hoja Activa'],
        columns,
        startRow: exec.startRow === 'dynamic' ? 'Dinámico (Post-Encabezados)' : exec.startRow || 1,
        hasStaticRow: typeof exec.startRow === 'number' && exec.startRow > 1
      };
      state.legacyData = null;
    } else {
      // Legacy mode: extract via AST tree walk
      const legacy = ASTEngine.extractLegacyHeaders(sf);
      state.legacyData = legacy;

      const sheets = legacy.sheets.map(s => s.value);
      if (sheets.length === 0) sheets.push('Hoja Activa');

      const columns = legacy.headers.map(h => h.value);

      // Detect static row from AST
      const execParams = detectLegacyExecParams(sf);

      state.parsedMetadata = {
        sheets,
        columns,
        startRow: execParams.startRow,
        hasStaticRow: execParams.hasStaticRow
      };
    }
  } catch (err) {
    console.error('ASTEngine parse error:', err);
    state.sourceFile = null;
    state.configSAP = null;
    state.legacyData = null;
  }

  refreshAllPanels();
}

function detectLegacyExecParams(sf) {
  let startRow = 'Dinámico (Post-Encabezados)';
  let hasStaticRow = false;

  // Use ASTEngine walker to find for-loop patterns
  ASTEngine.walk(sf, node => {
    if (node.kind === ts.SyntaxKind.ForStatement && node.initializer &&
        node.initializer.kind === ts.SyntaxKind.VariableDeclarationList) {
      const decls = node.initializer.declarations;
      if (decls.length === 1 && decls[0].initializer &&
          decls[0].initializer.kind === ts.SyntaxKind.NumericLiteral) {
        const val = Number(decls[0].initializer.text);
        if (val > 1) {
          const cond = node.condition ? node.condition.getText(sf) : '';
          if (cond.indexOf('Math.min') === -1) {
            startRow = val + 1; // display as 1-indexed row
            hasStaticRow = true;
          }
        }
      }
    }
    // Dynamic detection: headerRowCurrent + 1
    if (node.kind === ts.SyntaxKind.BinaryExpression) {
      const txt = node.getText(sf);
      if (txt.indexOf('headerRowCurrent') !== -1 || txt.indexOf('filaDeEncabezados') !== -1) {
        startRow = 'Dinámico (Post-Encabezados)';
        hasStaticRow = false;
      }
    }
  });

  return { startRow, hasStaticRow };
}

function refreshAllPanels() {
  updateStats();
  renderDissection();
  renderColumnMapper();
  renderSuggestions();
  updateMigrationBanner();
  updateSuggestionBadge();

  if (state.activeTab === 'simulation') {
    initSimulationData();
    renderSimulationSheet();
  }
  if (state.activeTab === 'lineage') renderLineage();
}

// ─── SIDEBAR STATS ───────────────────────────────────────────────────
function updateStats() {
  const pm = state.parsedMetadata;
  elements.statSheets.textContent = pm.sheets.join(', ');
  elements.statCols.textContent = pm.columns.length > 0 ? pm.columns.length + ' Detectadas' : 'Ninguna';
  elements.statStartRow.textContent = typeof pm.startRow === 'number' ? 'Fila ' + pm.startRow : pm.startRow;

  let complexity = 'Baja';
  const code = state.currentCode;
  if (code.indexOf('Map') !== -1 || code.indexOf('agrupacion') !== -1) complexity = 'Alta';
  else if (code.indexOf('highlight') !== -1 || code.indexOf('indexOf') !== -1) complexity = 'Media';
  elements.statComplexity.textContent = complexity;

  // Engine mode indicator
  if (elements.statEngine) {
    elements.statEngine.textContent = state.configSAP ? 'CONFIG_SAP (AST)' : 'Legacy (AST)';
    elements.statEngine.style.color = state.configSAP ? 'var(--success)' : 'var(--warning)';
  }
}

// ─── DISSECTION PANEL (AST-Based) ────────────────────────────────────
function renderDissection() {
  elements.dissectPanel.innerHTML = '';

  if (!state.sourceFile) {
    elements.dissectPanel.innerHTML = '<div class="dissect-empty">Carga un script para analizarlo.</div>';
    return;
  }

  const blocks = ASTEngine.dissectBlocks(state.sourceFile);

  if (blocks.length === 0) {
    elements.dissectPanel.innerHTML = '<div class="dissect-empty">No se detectó una función main() en el código.</div>';
    return;
  }

  blocks.forEach(block => {
    const card = document.createElement('div');
    card.className = 'dissect-card dissect-type-' + block.type;
    card.innerHTML =
      '<div class="dissect-step">' + block.step + '</div>' +
      '<div class="dissect-title">' + block.title + '</div>' +
      '<div class="dissect-desc">' + block.description + '</div>' +
      '<div class="dissect-lines">Líneas ' + block.startLine + '–' + block.endLine + '</div>';

    card.addEventListener('click', () => {
      document.querySelectorAll('.dissect-card').forEach(c => c.classList.remove('highlighted'));
      card.classList.add('highlighted');
      scrollEditorToLine(block.startLine);
    });

    elements.dissectPanel.appendChild(card);
  });
}

function scrollEditorToLine(lineNum) {
  const textarea = elements.codeEditor;
  const text = textarea.value;
  const lines = text.split('\n');

  let charIndex = 0;
  for (let i = 0; i < lineNum - 1 && i < lines.length; i++) {
    charIndex += lines[i].length + 1;
  }

  textarea.focus();
  textarea.setSelectionRange(charIndex, charIndex + (lines[lineNum - 1] || '').length);
  const lineHeight = 18;
  textarea.scrollTop = (lineNum - 3) * lineHeight;
}

// ─── COLUMN MAPPER (AST-Based) ──────────────────────────────────────
function renderColumnMapper() {
  elements.mapperList.innerHTML = '';

  if (state.configSAP && state.configSAP.data && state.configSAP.data.columns) {
    // CONFIG_SAP mode: show each column with its config path
    const cols = state.configSAP.data.columns;
    const keys = Object.keys(cols);

    if (keys.length === 0) {
      showMapperEmpty();
      return;
    }

    elements.btnApplyMapping.style.display = 'block';

    keys.forEach(key => {
      const col = cols[key];
      const header = col.header || '';
      if (!state.columnMapping[key]) state.columnMapping[key] = header;

      const div = document.createElement('div');
      div.className = 'mapping-row';
      div.innerHTML =
        '<div class="mapping-info">' +
          '<span class="col-key">' + key + '</span>' +
          '<span class="col-label">' + header + '</span>' +
        '</div>' +
        '<input type="text" class="col-input" value="' + escapeAttr(state.columnMapping[key]) + '" data-key="' + key + '" data-path="columns.' + key + '.header">';

      div.querySelector('input').addEventListener('input', e => {
        state.columnMapping[e.target.dataset.key] = e.target.value;
      });

      elements.mapperList.appendChild(div);
    });
  } else if (state.legacyData && state.legacyData.headers.length > 0) {
    // Legacy mode: show extracted headers
    elements.btnApplyMapping.style.display = 'none'; // can't safely mutate legacy

    state.legacyData.headers.forEach(h => {
      const div = document.createElement('div');
      div.className = 'mapping-row mapping-legacy';
      div.innerHTML =
        '<span class="col-label">' + h.value + '</span>' +
        '<span class="col-context">' + Array.from(h.contexts).join(', ') + '</span>' +
        '<span class="col-line">L' + h.locations[0].line + '</span>';

      elements.mapperList.appendChild(div);
    });

    // Show message about migrating
    const note = document.createElement('div');
    note.className = 'mapper-note';
    note.innerHTML = '⚠️ Este script usa encabezados dispersos. Migra a CONFIG_SAP para habilitar el remapeo quirúrgico.';
    elements.mapperList.appendChild(note);
  } else {
    showMapperEmpty();
  }
}

function showMapperEmpty() {
  elements.mapperList.innerHTML =
    '<div class="mapper-empty">No se detectaron encabezados explícitos para mapear en este script.</div>';
  elements.btnApplyMapping.style.display = 'none';
}

function escapeAttr(str) {
  return str.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ─── MAPPING APPLICATION (Surgical AST Mutation) ─────────────────────
function initMappingButton() {
  if (elements.btnApplyMapping) {
    elements.btnApplyMapping.addEventListener('click', applyASTMapping);
  }
}

function applyASTMapping() {
  if (!state.configSAP) {
    showToast('Solo se puede remapear en scripts con CONFIG_SAP.', 'warning');
    return;
  }

  const mutations = [];
  const cols = state.configSAP.data.columns;

  for (const key in state.columnMapping) {
    const newVal = state.columnMapping[key];
    const oldVal = cols[key] ? cols[key].header : '';
    if (newVal !== oldVal) {
      mutations.push({ path: 'columns.' + key + '.header', newValue: newVal });
    }
  }

  if (mutations.length === 0) {
    showToast('No se realizaron cambios en el mapeo.', 'warning');
    return;
  }

  const result = ASTEngine.mutateConfigBatch(state.currentCode, mutations);

  if (result.success) {
    state.currentCode = result.source;
    elements.codeEditor.value = result.source;
    analyzeScript();
    showToast('Mutación quirúrgica aplicada: ' + result.diffs.length + ' encabezado(s). Lógica intacta ✓');
  } else {
    showToast('Error en mutación: ' + result.error, 'danger');
  }
}

// ─── MIGRATION BANNER ────────────────────────────────────────────────
function initMigrationButton() {
  if (elements.btnMigrate) {
    elements.btnMigrate.addEventListener('click', handleMigration);
  }
}

function updateMigrationBanner() {
  if (!elements.migrationBanner) return;

  if (state.configSAP) {
    elements.migrationBanner.className = 'migration-banner config-present';
    elements.migrationBanner.innerHTML =
      '<span class="migration-badge success">✓ CONFIG_SAP</span>' +
      '<span>Motor AST quirúrgico activo. Las mutaciones están aisladas al bloque de configuración.</span>';
  } else if (state.legacyData && (state.legacyData.headers.length > 0 || state.legacyData.sheets.length > 0)) {
    elements.migrationBanner.className = 'migration-banner legacy-present';
    elements.migrationBanner.innerHTML =
      '<span class="migration-badge warning">⚠ Legacy</span>' +
      '<span>Script sin CONFIG_SAP. Los encabezados están dispersos en el código.</span>' +
      '<button class="btn btn-primary btn-sm" id="btn-migrate">Migrar a CONFIG_SAP</button>';
    // Re-attach event
    document.getElementById('btn-migrate').addEventListener('click', handleMigration);
  } else {
    elements.migrationBanner.className = 'migration-banner';
    elements.migrationBanner.innerHTML = '';
  }
}

function handleMigration() {
  if (!state.currentCode.trim()) return;

  const result = ASTEngine.migrateToConfigPattern(state.currentCode);

  if (result.alreadyMigrated) {
    showToast('Este script ya tiene CONFIG_SAP.', 'warning');
    return;
  }

  if (result.success) {
    state.currentCode = result.source;
    elements.codeEditor.value = result.source;
    analyzeScript();

    const s = result.stats;
    showToast(
      'Migración completada: ' + s.headersDetected + ' encabezados, ' +
      s.sheetsDetected + ' hojas, ' + s.totalEdits + ' referencias actualizadas.'
    );
  } else {
    showToast('Error en migración: ' + result.error, 'danger');
    if (result.warnings && result.warnings.length > 0) {
      console.warn('Migration warnings:', result.warnings);
    }
  }
}

// ─── SUGGESTIONS (AST-Based) ─────────────────────────────────────────
function renderSuggestions() {
  elements.suggestionsGrid.innerHTML = '';

  if (state.parsedMetadata.hasStaticRow) {
    const row = state.parsedMetadata.startRow;
    const card = document.createElement('div');
    card.className = 'suggestion-card';
    card.innerHTML =
      '<div class="suggestion-header">' +
        '<div class="suggestion-title">Optimización: Fila de Inicio Estática (Fila ' + row + ')</div>' +
        '<span class="suggestion-badge warning">Sugerencia</span>' +
      '</div>' +
      '<div class="suggestion-desc">' +
        'El script inicia la lectura fijamente en la fila ' + row + '. Si SAP modifica la estructura del reporte, el script omitirá registros o leerá filas incorrectas.' +
      '</div>' +
      '<div class="diff-preview">' +
        '<div class="diff-title">Código Propuesto</div>' +
        '<div class="diff-lines">' +
          '<span class="diff-del">- for (let i = ' + (row - 1) + '; i &lt; valores.length; i++) {</span>\n' +
          '<span class="diff-add">+ // Buscar encabezados dinámicamente\n+ for (let i = headerRowCurrent + 1; i &lt; valores.length; i++) {</span>' +
        '</div>' +
      '</div>' +
      '<button class="btn btn-primary" id="btn-apply-sug-row">Aplicar Corrección Dinámica</button>';

    card.querySelector('#btn-apply-sug-row').addEventListener('click', applyDynamicRowSuggestion);
    elements.suggestionsGrid.appendChild(card);
  }

  // Migration suggestion for legacy scripts
  if (!state.configSAP && state.legacyData && state.legacyData.headers.length > 0) {
    const card = document.createElement('div');
    card.className = 'suggestion-card';
    card.innerHTML =
      '<div class="suggestion-header">' +
        '<div class="suggestion-title">Arquitectura: Migrar a CONFIG_SAP</div>' +
        '<span class="suggestion-badge warning">Recomendado</span>' +
      '</div>' +
      '<div class="suggestion-desc">' +
        'Este script tiene ' + state.legacyData.headers.length + ' encabezados dispersos como string literals. ' +
        'Migrar a CONFIG_SAP permite mutación quirúrgica sin riesgo de corromper la lógica de negocio.' +
      '</div>' +
      '<button class="btn btn-primary" id="btn-sug-migrate">Migrar Ahora</button>';

    card.querySelector('#btn-sug-migrate').addEventListener('click', handleMigration);
    elements.suggestionsGrid.appendChild(card);
  }

  if (!state.parsedMetadata.hasStaticRow && state.configSAP) {
    const card = document.createElement('div');
    card.className = 'suggestion-card success';
    card.innerHTML =
      '<div class="suggestion-header">' +
        '<div class="suggestion-title">¡Estructura Óptima Detectada!</div>' +
        '<span class="suggestion-badge success" style="background:rgba(63,185,80,0.1);color:var(--success);border:1px solid rgba(63,185,80,0.2)">Óptimo</span>' +
      '</div>' +
      '<div class="suggestion-desc">' +
        'Este script usa CONFIG_SAP y búsquedas de encabezado dinámicas. Alta tolerancia ante variaciones en reportes SAP.' +
      '</div>';
    elements.suggestionsGrid.appendChild(card);
  }
}

function updateSuggestionBadge() {
  const badge = document.getElementById('sug-badge');
  if (!badge) return;
  let count = 0;
  if (state.parsedMetadata.hasStaticRow) count++;
  if (!state.configSAP && state.legacyData && state.legacyData.headers.length > 0) count++;
  badge.textContent = count;
  badge.style.display = count > 0 ? '' : 'none';
}

function applyDynamicRowSuggestion() {
  // This is the ONE place we do a targeted text replacement,
  // because the suggestion modifies the loop structure itself (not CONFIG_SAP).
  // We still use the AST to locate the exact for-statement.
  if (!state.sourceFile) return;

  let found = false;
  ASTEngine.walk(state.sourceFile, node => {
    if (found) return false;
    if (node.kind === ts.SyntaxKind.ForStatement && node.initializer &&
        node.initializer.kind === ts.SyntaxKind.VariableDeclarationList) {
      const decls = node.initializer.declarations;
      if (decls.length !== 1 || !decls[0].initializer) return;
      if (decls[0].initializer.kind !== ts.SyntaxKind.NumericLiteral) return;
      const val = Number(decls[0].initializer.text);
      if (val <= 1) return;

      // Verify this is a data loop, not a header scan
      const cond = node.condition ? node.condition.getText(state.sourceFile) : '';
      if (cond.indexOf('Math.min') !== -1) return;

      // Get positions for surgical replacement of the initializer value
      const numNode = decls[0].initializer;
      const start = numNode.getStart(state.sourceFile);
      const end = numNode.getEnd();

      let code = state.currentCode;

      // Inject headerRowCurrent if not present
      if (code.indexOf('headerRowCurrent') === -1 && code.indexOf('filaDeEncabezados') === -1) {
        // Find getValues() call to inject after
        const valuesIdx = code.indexOf('.getValues()');
        if (valuesIdx !== -1) {
          const lineEnd = code.indexOf('\n', valuesIdx);
          if (lineEnd !== -1) {
            const injection = '\n    let headerRowCurrent = 0; // Fila de encabezados (búsqueda dinámica)';
            code = code.slice(0, lineEnd) + injection + code.slice(lineEnd);
            // Adjust positions since we inserted text before the for loop
            const shift = injection.length;
            state.currentCode = code;
            // Re-parse with the injection
            const sf2 = ASTEngine.parse(code);
            // Now find the for loop again and replace the numeric literal
            let replaced = false;
            ASTEngine.walk(sf2, n2 => {
              if (replaced) return false;
              if (n2.kind === ts.SyntaxKind.ForStatement && n2.initializer &&
                  n2.initializer.kind === ts.SyntaxKind.VariableDeclarationList) {
                const d2 = n2.initializer.declarations;
                if (d2.length === 1 && d2[0].initializer &&
                    d2[0].initializer.kind === ts.SyntaxKind.NumericLiteral &&
                    Number(d2[0].initializer.text) > 1) {
                  const c2 = n2.condition ? n2.condition.getText(sf2) : '';
                  if (c2.indexOf('Math.min') === -1) {
                    const s2 = d2[0].initializer.getStart(sf2);
                    const e2 = d2[0].initializer.getEnd();
                    code = code.slice(0, s2) + 'headerRowCurrent + 1' + code.slice(e2);
                    replaced = true;
                  }
                }
              }
            });
            if (replaced) {
              state.currentCode = code;
              elements.codeEditor.value = code;
              analyzeScript();
              showToast('Corrección de inicio dinámico aplicada via AST.');
              found = true;
              return false;
            }
          }
        }
      } else {
        // headerRowCurrent already exists — just replace the number
        code = code.slice(0, start) + 'headerRowCurrent + 1' + code.slice(end);
        state.currentCode = code;
        elements.codeEditor.value = code;
        analyzeScript();
        showToast('Corrección de inicio dinámico aplicada via AST.');
        found = true;
        return false;
      }
    }
  });

  if (!found) {
    showToast('No se encontró un bucle con fila estática para corregir.', 'danger');
  }
}

// ─── LINEAGE GRAPH ───────────────────────────────────────────────────
function renderLineage() {
  if (!elements.lineageCanvas || !state.sourceFile) return;

  try {
    const graphData = ASTEngine.analyzeLineage(state.sourceFile);
    LineageGraph.render(elements.lineageCanvas, graphData);
  } catch (err) {
    console.error('Lineage analysis error:', err);
  }
}

// ─── SIMULATION ENGINE ──────────────────────────────────────────────
// (Preserved from v1 — no architectural changes needed)

function initSimulation() {
  elements.btnRunSim.addEventListener('click', runSimulation);
  elements.btnResetSim.addEventListener('click', resetSimulation);
  initSimulationData();
}

function initSimulationData() {
  const sheets = state.parsedMetadata.sheets;
  const columns = state.parsedMetadata.columns.length > 0
    ? state.parsedMetadata.columns
    : ['Nº doc.', 'Referencia', 'Asignación', 'Importe en ML', 'IVA 16%'];

  state.simulation.sheetsData = {};

  sheets.forEach(sheetName => {
    const rows = [];
    if (sheetName === 'IVA COBRADO') {
      const cols = ['Asignación', 'Importe en ML', 'Bco.prp.', 'Diferencia'];
      for (let i = 0; i < 6; i++) {
        rows.push({
          'Asignación': 'FACT-2026-1020' + i,
          'Importe en ML': '$ ' + (15000 + i * 2300).toLocaleString('en-US') + '.00',
          'Bco.prp.': 'BANORTE', 'Diferencia': '$ 0.00'
        });
      }
      state.simulation.sheetsData[sheetName] = { columns: cols, rows };
    } else if (sheetName === 'por factura') {
      const cols = ['St', 'Referencia', 'Asignación', 'Nº doc.', 'Clase', 'Fecha doc.', 'Importe en ML'];
      for (let i = 0; i < 6; i++) {
        rows.push({
          'St': '🔴', 'Referencia': 'REF1020' + i, 'Asignación': 'FACT-2026-1020' + i,
          'Nº doc.': '10005432' + i, 'Clase': 'RV', 'Fecha doc.': '1' + i + '/06/2026',
          'Importe en ML': '$ ' + (12000 + i * 1500).toLocaleString('en-US') + '.00'
        });
      }
      state.simulation.sheetsData[sheetName] = { columns: cols, rows };
    } else if (sheetName === 'todas las partidas' || sheetName === 'Sheet1') {
      const cols = ['Nº doc.', 'Doc.comp.', 'Compens.', 'Importe en ML', 'Texto'];
      for (let i = 0; i < 6; i++) {
        rows.push({
          'Nº doc.': '10005432' + i, 'Doc.comp.': '9005423' + i,
          'Compens.': '1' + i + '/06/2026',
          'Importe en ML': '$ ' + (12000 + i * 1500).toLocaleString('en-US') + '.00',
          'Texto': 'Pago Factura 1020' + i
        });
      }
      state.simulation.sheetsData[sheetName] = { columns: cols, rows };
    } else {
      const cols = [...columns];
      if (cols.indexOf('Nº doc.') === -1) cols.unshift('Nº doc.');
      if (cols.indexOf('Asignación') === -1) cols.push('Asignación');
      if (cols.indexOf('Importe en ML') === -1) cols.push('Importe en ML');
      for (let i = 0; i < 6; i++) {
        const row = {};
        cols.forEach(col => {
          if (col === 'Nº doc.') row[col] = '10005432' + i;
          else if (col === 'Asignación') row[col] = 'FACT-2026-1020' + i;
          else if (col === 'Importe en ML') row[col] = '$ ' + (12000 + i * 1500).toLocaleString('en-US') + '.00';
          else if (col === 'IVA 16%') row[col] = '$ ' + (1920 + i * 240).toLocaleString('en-US') + '.00';
          else if (col === 'Referencia') row[col] = 'REF1020' + i;
          else if (col === 'POLIZA' || col === 'FECHA' || col === 'IVA') row[col] = '';
          else row[col] = 'Val-' + i;
        });
        rows.push(row);
      }
      state.simulation.sheetsData[sheetName] = { columns: cols, rows };
    }
  });

  state.simulation.activeSheet = sheets[0] || 'Hoja Activa';
  renderSheetTabs();
}

function renderSheetTabs() {
  elements.simSheetTabs.innerHTML = '';
  Object.keys(state.simulation.sheetsData).forEach(sheet => {
    const btn = document.createElement('button');
    btn.className = 'tab-link' + (sheet === state.simulation.activeSheet ? ' active' : '');
    btn.innerHTML = '<span style="color:var(--gold)">⊞</span> ' + sheet;
    btn.addEventListener('click', () => {
      state.simulation.activeSheet = sheet;
      renderSheetTabs();
      renderSimulationSheet();
    });
    elements.simSheetTabs.appendChild(btn);
  });
}

function renderSimulationSheet(highlightedRow, scanningCols, activeCell) {
  highlightedRow = highlightedRow !== undefined ? highlightedRow : -1;
  scanningCols = scanningCols || [];
  activeCell = activeCell || null;

  const data = state.simulation.sheetsData[state.simulation.activeSheet];
  if (!data) return;

  elements.simulationGrid.innerHTML = '';

  const thead = document.createElement('thead');
  const hRow = document.createElement('tr');
  data.columns.forEach(col => {
    const th = document.createElement('th');
    th.textContent = col;
    if (scanningCols.indexOf(col) !== -1) th.className = 'header-active';
    hRow.appendChild(th);
  });
  thead.appendChild(hRow);
  elements.simulationGrid.appendChild(thead);

  const tbody = document.createElement('tbody');
  data.rows.forEach((row, ri) => {
    const tr = document.createElement('tr');
    if (ri === highlightedRow) tr.className = 'highlight-yellow';
    data.columns.forEach(col => {
      const td = document.createElement('td');
      td.textContent = row[col] || '';
      if (activeCell && activeCell.row === ri && activeCell.col === col) td.className = 'cell-active';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  elements.simulationGrid.appendChild(tbody);
}

function runSimulation() {
  if (state.simulation.running) return;
  state.simulation.running = true;
  elements.btnRunSim.disabled = true;
  elements.btnRunSim.style.opacity = '0.5';
  elements.simStatus.textContent = 'Simulando ejecución: Inicializando workbook...';

  let step = 0;
  const activeSheet = Object.keys(state.simulation.sheetsData)[0];
  const targetSheet = Object.keys(state.simulation.sheetsData)[1] || 'IVA COBRADO';

  state.simulation.intervalId = setInterval(() => {
    step++;
    if (step === 1) {
      elements.simStatus.textContent = 'Escaneando encabezados en "' + activeSheet + '"...';
      state.simulation.activeSheet = activeSheet;
      renderSheetTabs();
      const cols = state.simulation.sheetsData[activeSheet].columns;
      renderSimulationSheet(-1, cols.slice(0, 3));
    } else if (step === 2) {
      if (state.simulation.sheetsData[targetSheet]) {
        elements.simStatus.textContent = 'Escaneando encabezados en "' + targetSheet + '"...';
        state.simulation.activeSheet = targetSheet;
        renderSheetTabs();
        renderSimulationSheet(-1, state.simulation.sheetsData[targetSheet].columns.slice(0, 3));
      } else {
        elements.simStatus.textContent = 'Mapeo completado. Iniciando barrido...';
      }
    } else if (step >= 3 && step <= 8) {
      const ri = step - 3;
      state.simulation.activeSheet = activeSheet;
      renderSheetTabs();
      elements.simStatus.textContent = 'Procesando fila ' + (ri + 1) + '...';

      const name = state.selectedScript.toLowerCase();
      const rowData = state.simulation.sheetsData[activeSheet].rows[ri];

      if (name.indexOf('colores') !== -1) {
        renderSimulationSheet(ri, [], { row: ri, col: 'Nº doc.' });
      } else if (name.indexOf('conciliacion') !== -1 || name.indexOf('auxiliar') !== -1) {
        const cols = state.simulation.sheetsData[activeSheet].columns;
        if (cols.indexOf('POLIZA') === -1) cols.push('POLIZA', 'FECHA', 'IVA');
        rowData['POLIZA'] = 'PE-100234-' + ri;
        rowData['FECHA'] = '1' + ri + '/06/2026';
        rowData['IVA'] = '$ ' + (350 * (ri + 1)) + '.00';
        renderSimulationSheet(ri, [], { row: ri, col: 'POLIZA' });
      } else {
        renderSimulationSheet(ri, [], { row: ri, col: 'Asignación' });
      }
    } else {
      clearInterval(state.simulation.intervalId);
      state.simulation.running = false;
      elements.btnRunSim.disabled = false;
      elements.btnRunSim.style.opacity = '1';
      elements.simStatus.textContent = 'Simulación finalizada.';

      state.simulation.activeSheet = activeSheet;
      renderSheetTabs();

      const name = state.selectedScript.toLowerCase();
      if (name.indexOf('colores') !== -1) {
        const tbody = elements.simulationGrid.querySelector('tbody');
        if (tbody) {
          const rows = tbody.querySelectorAll('tr');
          if (rows[1]) rows[1].className = 'highlight-yellow';
          if (rows[3]) rows[3].className = 'highlight-yellow';
          if (rows[5]) rows[5].className = 'highlight-yellow';
        }
        showToast('¡Simulación completada! 3 filas resaltadas en amarillo.');
      } else if (name.indexOf('conciliacion') !== -1) {
        renderSimulationSheet();
        showToast('¡Simulación completada! Columnas POLIZA, FECHA, IVA calculadas.');
      } else {
        renderSimulationSheet();
        showToast('¡Simulación completada!');
      }
    }
  }, 1500);
}

function resetSimulation() {
  if (state.simulation.intervalId) clearInterval(state.simulation.intervalId);
  state.simulation.running = false;
  elements.btnRunSim.disabled = false;
  elements.btnRunSim.style.opacity = '1';
  elements.simStatus.textContent = 'Listo para simular la ejecución.';
  initSimulationData();
  renderSimulationSheet();
}
