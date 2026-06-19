// Global State
let state = {
    activeTab: 'visualizer',
    selectedScript: '',
    currentCode: '',
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

// DOM Elements
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
    mapperList: document.getElementById('mapper-list'),
    btnApplyMapping: document.getElementById('btn-apply-mapping'),
    suggestionsGrid: document.getElementById('suggestions-grid'),
    simulationGrid: document.getElementById('simulation-grid'),
    btnRunSim: document.getElementById('btn-run-sim'),
    btnResetSim: document.getElementById('btn-reset-sim'),
    simStatus: document.getElementById('sim-status'),
    simSheetTabs: document.getElementById('sim-sheet-tabs'),
    toastOverlay: document.getElementById('toast-overlay')
};

// Initialize Application
document.addEventListener('DOMContentLoaded', () => {
    initPresets();
    initTabNavigation();
    initEditorEvents();
    initSimulation();
});

// Toast Notifications
function showToast(message, type = 'success') {
    const toast = document.createElement('div');
    toast.className = `toast \${type}`;
    toast.innerHTML = `
        <span style="color: var(--gold)">✦</span>
        <span>\${message}</span>
    `;
    elements.toastOverlay.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

// Populate Presets in Sidebar
function initPresets() {
    elements.presetList.innerHTML = '';
    const fileNames = Object.keys(PRESET_SCRIPTS);
    
    fileNames.forEach((fileName, index) => {
        const li = document.createElement('li');
        li.className = `preset-item \${index === 0 ? 'active' : ''}`;
        li.textContent = fileName;
        li.title = fileName;
        li.addEventListener('click', () => selectPreset(fileName, li));
        elements.presetList.appendChild(li);
    });

    if (fileNames.length > 0) {
        loadScript(fileNames[0]);
    }
}

function selectPreset(fileName, element) {
    document.querySelectorAll('.preset-item').forEach(item => item.classList.remove('active'));
    element.classList.add('active');
    loadScript(fileName);
    showToast(`Cargado: \${fileName}`);
}

function loadScript(fileName) {
    state.selectedScript = fileName;
    state.currentCode = PRESET_SCRIPTS[fileName] || '';
    elements.codeEditor.value = state.currentCode;
    analyzeScript();
    resetSimulation();
}

// Tab Navigation
function initTabNavigation() {
    elements.tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const targetTab = tab.dataset.tab;
            state.activeTab = targetTab;
            
            elements.tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');

            elements.tabContents.forEach(content => {
                content.classList.remove('active');
                if (content.id === `\${targetTab}-tab`) {
                    content.classList.add('active');
                }
            });

            if (targetTab === 'simulation') {
                renderSimulationSheet();
            }
        });
    });
}

// Sync edits in code editor
function initEditorEvents() {
    elements.codeEditor.addEventListener('input', (e) => {
        state.currentCode = e.target.value;
        analyzeScript();
    });
}

// Parser Engine - Analyze TypeScript Script
function analyzeScript() {
    const code = state.currentCode;
    
    // 1. Detect Worksheets
    const sheetsSet = new Set();
    const sheetRegex = /getWorksheet\\s*\\(\\s*["']([^"']+)["']\\s*\\)/g;
    let match;
    while ((match = sheetRegex.exec(code)) !== null) {
        sheetsSet.add(match[1]);
    }
    const sheets = Array.from(sheetsSet);
    if (sheets.length === 0) sheets.push('Hoja Activa');
    state.parsedMetadata.sheets = sheets;

    // 2. Detect Columns / Headers
    const columnsSet = new Set();
    const colRegex = /indexOf\\s*\\(\\s*["']([^"']+)["']\\s*\\)/g;
    while ((match = colRegex.exec(code)) !== null) {
        if (!["", "Compens/"].includes(match[1])) {
            columnsSet.add(match[1]);
        }
    }
    
    const arrayRegex = /\\[\\s*(?:["'][^"']+["']\\s*,\\s*)*["'][^"']+["']\\s*\\]/g;
    const arraysMatched = code.match(arrayRegex);
    if (arraysMatched) {
        arraysMatched.forEach(arrStr => {
            try {
                const normalized = arrStr.replace(/'/g, '"');
                const arr = JSON.parse(normalized);
                if (Array.isArray(arr) && arr.length > 2) {
                    arr.forEach(item => {
                        if (typeof item === 'string' && item.length > 2 && item.length < 30) {
                            columnsSet.add(item);
                        }
                    });
                }
            } catch (e) {
            }
        });
    }

    const columns = Array.from(columnsSet);
    state.parsedMetadata.columns = columns;

    // 3. Detect Starting Row
    let startRow = 1;
    let hasStaticRow = false;
    
    const staticRowRegex = /for\\s*\\(\\s*let\\s+[i|r|j]\\s*=\\s*(\\d+)/;
    const staticMatch = staticRowRegex.exec(code);
    if (staticMatch) {
        startRow = parseInt(staticMatch[1]) + 1;
        hasStaticRow = true;
    }
    
    const dynamicRowRegex = /headerRowCurrent\\s*\\+\\s*1/i;
    if (dynamicRowRegex.test(code)) {
        startRow = 'Dinámico (Post-Encabezados)';
        hasStaticRow = false;
    }

    state.parsedMetadata.startRow = startRow;
    state.parsedMetadata.hasStaticRow = hasStaticRow;

    elements.statSheets.textContent = sheets.join(', ');
    elements.statCols.textContent = columns.length > 0 ? `\${columns.length} Detectadas` : 'Ninguna';
    elements.statStartRow.textContent = typeof startRow === 'number' ? `Fila \${startRow}` : startRow;
    
    let complexity = 'Baja';
    if (code.includes('Map') || code.includes('agrupacion')) complexity = 'Alta';
    else if (code.includes('highlight') || code.includes('indexOf')) complexity = 'Media';
    elements.statComplexity.textContent = complexity;

    renderDissection();
    renderColumnMapper();
    renderSuggestions();
    
    if (state.activeTab === 'simulation') {
        initSimulationData();
        renderSimulationSheet();
    }
}

// Render Dissection Cards (Step-by-Step)
function renderDissection() {
    elements.dissectPanel.innerHTML = '';
    const code = state.currentCode;
    const sections = [];

    if (code.includes('getWorksheet') || code.includes('getActiveWorksheet')) {
        sections.push({
            step: 'Paso 1: Inicialización',
            title: 'Configuración y Selección de Hojas',
            desc: `El script se conecta a Excel y localiza las pestañas de trabajo. Se detectan las hojas: \${state.parsedMetadata.sheets.map(s => `"\${s}"`).join(', ')}.`,
            matchKeyword: 'getWorksheet'
        });
    }

    if (code.includes('indexOf') || code.includes('indicesColumnasValidacion') || code.includes('indicesTemporales')) {
        sections.push({
            step: 'Paso 2: Localización de Encabezados',
            title: 'Mapeo Dinámico de Columnas',
            desc: `Escanea las primeras filas (comúnmente las primeras 30 o 50) para encontrar los índices numéricos de columnas clave como \${state.parsedMetadata.columns.slice(0, 3).map(c => `"\${c}"`).join(', ')}.`,
            matchKeyword: 'indexOf'
        });
    }

    if (code.includes('for') || code.includes('while')) {
        sections.push({
            step: 'Paso 3: Bucle de Barrido',
            title: 'Lectura Secuencial de Filas (O(N))',
            desc: `Recorre todas las filas de la tabla de forma eficiente. \${state.parsedMetadata.hasStaticRow ? `Empieza fijamente en la fila \${state.parsedMetadata.startRow}.` : 'Empieza dinámicamente justo después de la fila de encabezados.'}`,
            matchKeyword: 'for'
        });
    }

    if (code.includes('highlight') || code.includes('color') || code.includes('Gastos no deducible') || code.includes('POLIZA')) {
        sections.push({
            step: 'Paso 4: Procesamiento y Salida',
            title: 'Cálculos Contables y Marcado',
            desc: 'Realiza el cuadre de importes, concilia registros, calcula IVA, inyecta fórmulas contables o aplica formatos visuales (como resaltar en amarillo celdas coincidentes).',
            matchKeyword: 'highlightColor'
        });
    }

    sections.forEach(sec => {
        const card = document.createElement('div');
        card.className = 'dissect-card';
        card.innerHTML = `
            <div class="dissect-step">\${sec.step}</div>
            <div class="dissect-title">\${sec.title}</div>
            <div class="dissect-desc">\${sec.desc}</div>
        `;
        
        card.addEventListener('click', () => {
            document.querySelectorAll('.dissect-card').forEach(c => c.classList.remove('highlighted'));
            card.classList.add('highlighted');
            
            const textarea = elements.codeEditor;
            const text = textarea.value;
            const index = text.indexOf(sec.matchKeyword);
            
            if (index !== -1) {
                textarea.focus();
                textarea.setSelectionRange(index, index + sec.matchKeyword.length);
                const lineHeight = 18;
                const linesBefore = text.substring(0, index).split('\n').length;
                textarea.scrollTop = (linesBefore * lineHeight) - 100;
            }
        });

        elements.dissectPanel.appendChild(card);
    });
}

function renderColumnMapper() {
    elements.mapperList.innerHTML = '';
    const columns = state.parsedMetadata.columns;

    if (columns.length === 0) {
        elements.mapperList.innerHTML = '<div style="color: var(--text-secondary); font-size: 0.85rem; text-align: center; padding: 1rem;">No se detectaron encabezados explícitos para mapear en este script.</div>';
        elements.btnApplyMapping.style.display = 'none';
        return;
    }

    elements.btnApplyMapping.style.display = 'block';

    columns.forEach(col => {
        if (!state.columnMapping[col]) {
            state.columnMapping[col] = col;
        }

        const div = document.createElement('div');
        div.className = 'mapping-row';
        div.innerHTML = `
            <span class="col-label">\${col}</span>
            <input type="text" class="col-input" value="\${state.columnMapping[col]}" data-original="\${col}">
        `;

        div.querySelector('input').addEventListener('input', (e) => {
            const orig = e.target.dataset.original;
            state.columnMapping[orig] = e.target.value;
        });

        elements.mapperList.appendChild(div);
    });
}

function applyColumnMapping() {
    let updatedCode = state.currentCode;
    let replacedCount = 0;

    Object.keys(state.columnMapping).forEach(original => {
        const mapped = state.columnMapping[original];
        if (original !== mapped) {
            const escapedOrig = original.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
            const regex = new RegExp(`["']\${escapedOrig}["']`, 'g');
            
            if (regex.test(updatedCode)) {
                updatedCode = updatedCode.replace(regex, `"\${mapped}"`);
                replacedCount++;
            }
        }
    });

    if (replacedCount > 0) {
        state.currentCode = updatedCode;
        elements.codeEditor.value = updatedCode;
        analyzeScript();
        showToast(`Mapeo aplicado con éxito. \${replacedCount} encabezado(s) modificado(s) en el código.`);
    } else {
        showToast('No se realizaron cambios en el mapeo.', 'warning');
    }
}

function renderSuggestions() {
    elements.suggestionsGrid.innerHTML = '';
    
    if (state.parsedMetadata.hasStaticRow) {
        const startRow = state.parsedMetadata.startRow;
        const card = document.createElement('div');
        card.className = 'suggestion-card';
        card.innerHTML = `
            <div class="suggestion-header">
                <div class="suggestion-title">Optimización: Fila de Inicio Estática (Fila \${startRow})</div>
                <span class="suggestion-badge warning">Sugerencia</span>
            </div>
            <div class="suggestion-desc">
                El script tiene programado iniciar la lectura fijamente en la fila \${startRow}. Si la estructura del archivo descargado de SAP se modifica (ej. se agregan filas de título vacías en la parte superior), el script omitirá registros o leerá filas incorrectas.
            </div>
            <div class="diff-preview">
                <div class="diff-title">Código Propuesto</div>
                <div class="diff-lines"><span class="diff-del">- for (let i = \${startRow - 1}; i &lt; valores.length; i++) {</span>\\n<span class="diff-add">+ // Buscar encabezados dinámicamente y empezar desde headerRowCurrent + 1\\n+ for (let i = headerRowCurrent + 1; i &lt; valores.length; i++) {</span></div>
            </div>
            <button class="btn btn-primary" id="btn-apply-suggestion" style="align-self: flex-start;">Aplicar Corrección Dinámica</button>
        `;

        card.querySelector('#btn-apply-suggestion').addEventListener('click', applyDynamicRowSuggestion);
        elements.suggestionsGrid.appendChild(card);
    } else {
        const card = document.createElement('div');
        card.className = 'suggestion-card success';
        card.innerHTML = `
            <div class="suggestion-header">
                <div class="suggestion-title">¡Estructura Óptima Detectada!</div>
                <span class="suggestion-badge success" style="background-color: rgba(63, 185, 80, 0.1); color: var(--success); border: 1px solid rgba(63, 185, 80, 0.2)">Óptimo</span>
            </div>
            <div class="suggestion-desc">
                Este script ya implementa búsquedas de encabezado dinámicas o no tiene filas fijas hardcodeadas para el inicio del barrido de filas. Esto le da alta tolerancia ante variaciones en la estructura de los reportes.
            </div>
        `;
        elements.suggestionsGrid.appendChild(card);
    }
}

function applyDynamicRowSuggestion() {
    let code = state.currentCode;
    const staticRowRegex = /(for\\s*\\(\\s*let\\s+([i|r|j])\\s*=\\s*)(\\d+)(\\s*;\\s*\\2\\s*&lt;\\s*[\\w\\.]+(?:length)?\\s*;\\s*\\2\\+\\+\\s*\\))/;
    const match = staticRowRegex.exec(code);
    
    if (match) {
        const fullLoopDef = match[0];
        const loopVar = match[2];
        const replacement = `for (let \${loopVar} = headerRowCurrent + 1\${match[4]}`;
        
        if (!code.includes('headerRowCurrent')) {
            const valuesMatch = /(const|let)\\s+(\\w+)\\s*=\\s*\\w+\\.getValues\\(\\)/.exec(code);
            if (valuesMatch) {
                const headerRowInjection = `\\n    let headerRowCurrent = 0; // Busqueda dinamica pre-calculada\\n    `;
                code = code.replace(valuesMatch[0], valuesMatch[0] + headerRowInjection);
            }
        }

        code = code.replace(fullLoopDef, replacement);
        state.currentCode = code;
        elements.codeEditor.value = code;
        analyzeScript();
        showToast('Corrección de inicio dinámico aplicada en el script.');
    } else {
        showToast('No se pudo encontrar la estructura del bucle de manera limpia para reescribirlo.', 'danger');
    }
}

function initSimulation() {
    elements.btnRunSim.addEventListener('click', runSimulation);
    elements.btnResetSim.addEventListener('click', resetSimulation);
    initSimulationData();
}

function initSimulationData() {
    const sheets = state.parsedMetadata.sheets;
    const columns = state.parsedMetadata.columns.length > 0 ? state.parsedMetadata.columns : ['Nº doc.', 'Referencia', 'Asignación', 'Importe en ML', 'IVA 16%'];
    
    state.simulation.sheetsData = {};
    
    sheets.forEach(sheetName => {
        const rows = [];
        
        if (sheetName === 'IVA COBRADO') {
            const cols = ['Asignación', 'Importe en ML', 'Bco.prp.', 'Diferencia'];
            for (let i = 0; i < 6; i++) {
                rows.push({
                    'Asignación': `FACT-2026-1020\th(i)`,
                    'Importe en ML': `$ \th(15000 + i * 2300).toLocaleString('en-US')}.00`,
                    'Bco.prp.': 'BANORTE',
                    'Diferencia': '$ 0.00'
                });
            }
            state.simulation.sheetsData[sheetName] = { columns: cols, rows: rows };
        } else if (sheetName === 'por factura') {
            const cols = ['St', 'Referencia', 'Asignación', 'Nº doc.', 'Clase', 'Fecha doc.', 'Importe en ML'];
            for (let i = 0; i < 6; i++) {
                rows.push({
                    'St': '🔴',
                    'Referencia': `REF1020\th(i)`,
                    'Asignación': `FACT-2026-1020\th(i)`,
                    'Nº doc.': `10005432\th(i)`,
                    'Clase': 'RV',
                    'Fecha doc.': `1\th(i)/06/2026`,
                    'Importe en ML': `$ \th(12000 + i * 1500).toLocaleString('en-US')}.00`
                });
            }
            state.simulation.sheetsData[sheetName] = { columns: cols, rows: rows };
        } else if (sheetName === 'todas las partidas' || sheetName === 'Sheet1') {
            const cols = ['Nº doc.', 'Doc.comp.', 'Compens.', 'Importe en ML', 'Texto'];
            for (let i = 0; i < 6; i++) {
                rows.push({
                    'Nº doc.': `10005432\th(i)`,
                    'Doc.comp.': `9005423\th(i)`,
                    'Compens.': `1\th(i)/06/2026`,
                    'Importe en ML': `$ \th(12000 + i * 1500).toLocaleString('en-US')}.00`,
                    'Texto': `Pago Factura 1020\th(i)`
                });
            }
            state.simulation.sheetsData[sheetName] = { columns: cols, rows: rows };
        } else {
            const cols = [...columns];
            if (!cols.includes('Nº doc.')) cols.unshift('Nº doc.');
            if (!cols.includes('Asignación')) cols.push('Asignación');
            if (!cols.includes('Importe en ML')) cols.push('Importe en ML');
            
            for (let i = 0; i < 6; i++) {
                const row = {};
                cols.forEach(col => {
                    if (col === 'Nº doc.') row[col] = `10005432\th(i)`;
                    else if (col === 'Asignación') row[col] = `FACT-2026-1020\th(i)`;
                    else if (col === 'Importe en ML') row[col] = `$ \th(12000 + i * 1500).toLocaleString('en-US')}.00`;
                    else if (col === 'IVA 16%') row[col] = `$ \th(1920 + i * 240).toLocaleString('en-US')}.00`;
                    else if (col === 'Referencia') row[col] = `REF1020\th(i)`;
                    else if (col === 'POLIZA') row[col] = '';
                    else if (col === 'FECHA') row[col] = '';
                    else if (col === 'IVA') row[col] = '';
                    else row[col] = `Valor-\th(col)-\th(i)`;
                });
                rows.push(row);
            }
            state.simulation.sheetsData[sheetName] = { columns: cols, rows: rows };
        }
    });

    state.simulation.activeSheet = sheets[0] || 'Hoja Activa';
    renderSheetTabs();
}

function renderSheetTabs() {
    elements.simSheetTabs.innerHTML = '';
    const sheets = Object.keys(state.simulation.sheetsData);
    
    sheets.forEach(sheet => {
        const button = document.createElement('button');
        button.className = `tab-link \${sheet === state.simulation.activeSheet ? 'active' : ''}`;
        button.innerHTML = `<span style="color: var(--gold)">⊞</span> \${sheet}`;
        button.addEventListener('click', () => {
            state.simulation.activeSheet = sheet;
            renderSheetTabs();
            renderSimulationSheet();
        });
        elements.simSheetTabs.appendChild(button);
    });
}

function renderSimulationSheet(highlightedRow = -1, scanningCols = [], activeCell = null) {
    const sheetName = state.simulation.activeSheet;
    const sheetData = state.simulation.sheetsData[sheetName];
    
    if (!sheetData) return;
    
    elements.simulationGrid.innerHTML = '';
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    sheetData.columns.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        if (scanningCols.includes(col)) {
            th.className = 'header-active';
        }
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    elements.simulationGrid.appendChild(thead);
    
    const tbody = document.createElement('tbody');
    sheetData.rows.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        
        if (rowIndex === highlightedRow) {
            tr.className = 'highlight-yellow';
        }
        
        sheetData.columns.forEach(col => {
            const td = document.createElement('td');
            td.textContent = row[col] || '';
            
            if (activeCell && activeCell.row === rowIndex && activeCell.col === col) {
                td.className = 'cell-active';
            }
            
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
    const activeSheetName = Object.keys(state.simulation.sheetsData)[0];
    const targetSheetName = Object.keys(state.simulation.sheetsData)[1] || 'IVA COBRADO';
    
    state.simulation.intervalId = setInterval(() => {
        step++;
        
        if (step === 1) {
            elements.simStatus.textContent = `Escaneando encabezados en hoja "\${activeSheetName}"...`;
            state.simulation.activeSheet = activeSheetName;
            renderSheetTabs();
            
            const cols = state.simulation.sheetsData[activeSheetName].columns;
            renderSimulationSheet(-1, cols.slice(0, 3));
            
        } else if (step === 2) {
            if (state.simulation.sheetsData[targetSheetName]) {
                elements.simStatus.textContent = `Escaneando encabezados en hoja de cruce "\${targetSheetName}"...`;
                state.simulation.activeSheet = targetSheetName;
                renderSheetTabs();
                
                const cols = state.simulation.sheetsData[targetSheetName].columns;
                renderSimulationSheet(-1, cols.slice(0, 3));
            } else {
                elements.simStatus.textContent = 'Mapeo completado. Iniciando barrido de filas contables...';
            }
            
        } else if (step >= 3 && step <= 8) {
            const rowIndex = step - 3;
            state.simulation.activeSheet = activeSheetName;
            renderSheetTabs();
            elements.simStatus.textContent = `Procesando fila \${rowIndex + 1} de la hoja activa...`;
            
            const currentScriptLower = state.selectedScript.toLowerCase();
            const rowData = state.simulation.sheetsData[activeSheetName].rows[rowIndex];
            
            if (currentScriptLower.includes('colores')) {
                renderSimulationSheet(rowIndex, [], { row: rowIndex, col: 'Nº doc.' });
            } else if (currentScriptLower.includes('conciliacion') || currentScriptLower.includes('auxiliar')) {
                if (rowData.hasOwnProperty('POLIZA') || rowData.hasOwnProperty('FECHA') || rowData.hasOwnProperty('IVA')) {
                    rowData['POLIZA'] = `PE-100234-\${rowIndex}`;
                    rowData['FECHA'] = `1\${rowIndex}/06/2026`;
                    rowData['IVA'] = `$ \${(350 * (rowIndex + 1))}.00`;
                } else {
                    const cols = state.simulation.sheetsData[activeSheetName].columns;
                    if (!cols.includes('POLIZA')) cols.push('POLIZA', 'FECHA', 'IVA');
                    rowData['POLIZA'] = `PE-100234-\${rowIndex}`;
                    rowData['FECHA'] = `1\${rowIndex}/06/2026`;
                    rowData['IVA'] = `$ \${(350 * (rowIndex + 1))}.00`;
                }
                renderSimulationSheet(rowIndex, [], { row: rowIndex, col: 'POLIZA' });
            } else if (currentScriptLower.includes('cliente')) {
                if (rowData.hasOwnProperty('Nombre del Cliente') || rowData.hasOwnProperty('Cliente')) {
                    const clientCol = rowData.hasOwnProperty('Nombre del Cliente') ? 'Nombre del Cliente' : 'Cliente';
                    rowData[clientCol] = `CLIENTE COMERCIAL MEXICO SA DE CV - POOL \${rowIndex}`;
                } else {
                    const cols = state.simulation.sheetsData[activeSheetName].columns;
                    if (!cols.includes('Cliente')) cols.push('Cliente');
                    rowData['Cliente'] = `CLIENTE COMERCIAL MEXICO SA DE CV - POOL \${rowIndex}`;
                }
                renderSimulationSheet(rowIndex, [], { row: rowIndex, col: 'Cliente' });
            } else {
                renderSimulationSheet(rowIndex, [], { row: rowIndex, col: 'Asignación' });
            }
            
        } else {
            clearInterval(state.simulation.intervalId);
            state.simulation.running = false;
            elements.btnRunSim.disabled = false;
            elements.btnRunSim.style.opacity = '1';
            elements.simStatus.textContent = 'Simulación finalizada. Todos los registros procesados con éxito.';
            
            state.simulation.activeSheet = activeSheetName;
            renderSheetTabs();
            
            const currentScriptLower = state.selectedScript.toLowerCase();
            if (currentScriptLower.includes('colores')) {
                const tbody = elements.simulationGrid.querySelector('tbody');
                const rows = tbody.querySelectorAll('tr');
                rows[1].className = 'highlight-yellow';
                rows[3].className = 'highlight-yellow';
                rows[5].className = 'highlight-yellow';
                
                showToast('¡Simulación completada! 3 filas coincidentes se resaltaron en amarillo en Excel.');
            } else if (currentScriptLower.includes('conciliacion') || currentScriptLower.includes('auxiliar')) {
                renderSimulationSheet();
                showToast('¡Simulación completada! Columnas POLIZA, FECHA e IVA calculadas.');
            } else {
                renderSimulationSheet();
                showToast('¡Simulación completada!');
            }
        }
    }, 1500);
}

function resetSimulation() {
    if (state.simulation.intervalId) {
        clearInterval(state.simulation.intervalId);
    }
    state.simulation.running = false;
    elements.btnRunSim.disabled = false;
    elements.btnRunSim.style.opacity = '1';
    elements.simStatus.textContent = 'Listo para simular la ejecución.';
    initSimulationData();
    renderSimulationSheet();
}