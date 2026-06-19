/**
 * AST Engine for Office Script Analysis
 *
 * Structural code analysis using the TypeScript Compiler API.
 * Zero regex for logic extraction. All mutations are AST-guided and surgical.
 *
 * Dependencies: TypeScript compiler (global `ts` object from CDN)
 */
const ASTEngine = (() => {
  'use strict';

  // ─── GUARDS ─────────────────────────────────────────────────────────
  function requireTS() {
    if (typeof ts === 'undefined') {
      throw new Error(
        'TypeScript compiler not loaded. Include typescript.js before ast-engine.js.'
      );
    }
  }

  // ─── CORE PARSING ───────────────────────────────────────────────────
  /** Parses TypeScript source code into a full AST with parent pointers. */
  function parse(source) {
    requireTS();
    return ts.createSourceFile(
      'script.ts', source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS
    );
  }

  // ─── AST TRAVERSAL UTILITIES ────────────────────────────────────────
  function walk(node, visitor) {
    const result = visitor(node);
    if (result === false) return;
    ts.forEachChild(node, child => walk(child, visitor));
  }

  function findNodes(root, predicate) {
    const results = [];
    walk(root, node => { if (predicate(node)) results.push(node); });
    return results;
  }

  function findFirst(root, predicate) {
    let found = null;
    walk(root, node => {
      if (found) return false;
      if (predicate(node)) { found = node; return false; }
    });
    return found;
  }

  // Kind checks — intentionally verbose for readability
  function isStringLiteral(n) { return n.kind === ts.SyntaxKind.StringLiteral; }
  function isNumericLiteral(n) { return n.kind === ts.SyntaxKind.NumericLiteral; }
  function isIdentifier(n) { return n.kind === ts.SyntaxKind.Identifier; }
  function isObjectLiteral(n) { return n.kind === ts.SyntaxKind.ObjectLiteralExpression; }
  function isPropertyAssignment(n) { return n.kind === ts.SyntaxKind.PropertyAssignment; }
  function isCallExpression(n) { return n.kind === ts.SyntaxKind.CallExpression; }
  function isPropertyAccess(n) { return n.kind === ts.SyntaxKind.PropertyAccessExpression; }
  function isVariableDeclaration(n) { return n.kind === ts.SyntaxKind.VariableDeclaration; }
  function isVariableStatement(n) { return n.kind === ts.SyntaxKind.VariableStatement; }
  function isBinaryExpression(n) { return n.kind === ts.SyntaxKind.BinaryExpression; }
  function isForStatement(n) { return n.kind === ts.SyntaxKind.ForStatement; }
  function isArrayLiteral(n) { return n.kind === ts.SyntaxKind.ArrayLiteralExpression; }
  function isFunctionDeclaration(n) { return n.kind === ts.SyntaxKind.FunctionDeclaration; }
  function isElementAccess(n) { return n.kind === ts.SyntaxKind.ElementAccessExpression; }

  /**
   * Resolves a PropertyAccessExpression chain to a dot-separated path.
   * e.g. CONFIG_SAP.columns.nDoc.header → "CONFIG_SAP.columns.nDoc.header"
   */
  function resolvePropertyChain(node, sf) {
    if (isIdentifier(node)) return node.text;
    if (isPropertyAccess(node)) {
      return resolvePropertyChain(node.expression, sf) + '.' + node.name.text;
    }
    return node.getText(sf);
  }

  /** Gets the 1-indexed line number for a position in the source file. */
  function getLineNumber(sf, pos) {
    return sf.getLineAndCharacterOfPosition(pos).line + 1;
  }

  // ─── OBJECT LITERAL ↔ JS CONVERSION ────────────────────────────────

  /** Recursively converts an AST ObjectLiteralExpression to a plain JS object. */
  function objectLiteralToJS(node, sf) {
    if (isObjectLiteral(node)) {
      const result = {};
      for (const prop of node.properties) {
        if (isPropertyAssignment(prop)) {
          result[prop.name.getText(sf)] = objectLiteralToJS(prop.initializer, sf);
        }
      }
      return result;
    }
    if (isStringLiteral(node)) return node.text;
    if (isNumericLiteral(node)) return Number(node.text);
    if (isArrayLiteral(node)) return node.elements.map(el => objectLiteralToJS(el, sf));
    if (node.kind === ts.SyntaxKind.TrueKeyword) return true;
    if (node.kind === ts.SyntaxKind.FalseKeyword) return false;
    if (node.kind === ts.SyntaxKind.NullKeyword) return null;
    return node.getText(sf);
  }

  /** Generates a CONFIG_SAP source code block from structured data. */
  function generateConfigBlock(sheets, columns, execution) {
    const I = '    ';  // 4-space indent matching Office Script style
    const I2 = I + I;
    const I3 = I2 + I;
    const lines = [];

    lines.push(I + 'const CONFIG_SAP = {');
    lines.push(I2 + 'version: "2.0",');
    lines.push('');

    // Sheets
    lines.push(I2 + 'sheets: {');
    const sk = Object.keys(sheets);
    sk.forEach((key, i) => {
      const s = sheets[key];
      const esc = s.name.replace(/"/g, '\\"');
      lines.push(I3 + key + ': { name: "' + esc + '", role: "' + s.role + '" }' + (i < sk.length - 1 ? ',' : ''));
    });
    lines.push(I2 + '},');
    lines.push('');

    // Columns
    lines.push(I2 + 'columns: {');
    const ck = Object.keys(columns);
    ck.forEach((key, i) => {
      const c = columns[key];
      const esc = c.header.replace(/"/g, '\\"');
      const desc = (c.description || '').replace(/"/g, '\\"');
      lines.push(I3 + key + ': { header: "' + esc + '", description: "' + desc + '" }' + (i < ck.length - 1 ? ',' : ''));
    });
    lines.push(I2 + '},');
    lines.push('');

    // Execution
    const exec = execution || {};
    const sr = typeof exec.startRow === 'number' ? String(exec.startRow) : '"dynamic"';
    lines.push(I2 + 'execution: {');
    lines.push(I3 + 'headerSearchDepth: ' + (exec.headerSearchDepth || 30) + ',');
    lines.push(I3 + 'startRow: ' + sr + ',');
    lines.push(I3 + 'minHeaderMatches: ' + (exec.minHeaderMatches || 2));
    lines.push(I2 + '}');

    lines.push(I + '};');
    return lines.join('\n');
  }

  // ─── CONFIG_SAP EXTRACTION ──────────────────────────────────────────

  /**
   * Locates and extracts the CONFIG_SAP declaration from a parsed source file.
   * Returns null if not found.  Also builds a nodeMap for surgical path-based access.
   */
  function extractConfigSAP(sf) {
    const decl = findFirst(sf, node =>
      isVariableDeclaration(node) &&
      node.name && node.name.getText(sf) === 'CONFIG_SAP' &&
      node.initializer && isObjectLiteral(node.initializer)
    );
    if (!decl) return null;

    const configNode = decl.initializer;
    const data = objectLiteralToJS(configNode, sf);

    // Build a map of dotted paths → AST nodes for direct surgical access
    const nodeMap = {};
    (function buildMap(obj, prefix) {
      if (!isObjectLiteral(obj)) return;
      for (const prop of obj.properties) {
        if (!isPropertyAssignment(prop)) continue;
        const key = prop.name.getText(sf);
        const path = prefix ? prefix + '.' + key : key;
        nodeMap[path] = prop.initializer;
        if (isObjectLiteral(prop.initializer)) buildMap(prop.initializer, path);
      }
    })(configNode, '');

    // Walk up to the VariableStatement wrapper
    let stmt = decl;
    while (stmt && !isVariableStatement(stmt)) stmt = stmt.parent;

    return {
      node: configNode,
      declaration: decl,
      statement: stmt,
      startPos: (stmt || decl).getStart(sf),
      endPos: (stmt || decl).getEnd(),
      data,
      nodeMap
    };
  }

  // ─── LEGACY HEADER EXTRACTION (Zero Regex) ─────────────────────────

  /**
   * Heuristic filter: determines whether a string literal is plausibly
   * an Excel column header vs. a color code, format string, or error message.
   * Uses character analysis only — no regular expressions.
   */
  function isPlausibleHeader(str) {
    if (!str || str.length < 2 || str.length > 35) return false;
    const ch0 = str.charAt(0);
    // Hex color codes: #RRGGBB
    if (ch0 === '#' && str.length === 7) return false;
    // Number format strings: #,##0  or  #,##0.00;[Red]…
    if (ch0 === '#' && str.indexOf(',') !== -1) return false;
    // Pure numeric
    if (str.length < 8 && !isNaN(Number(str)) && str.trim().length > 0) return false;
    // Excel formulas
    if (ch0 === '=') return false;
    // R1C1 references
    if (str.indexOf('RC[') !== -1 || str.indexOf('RC[-') !== -1) return false;
    // Column range references: A:A, N:N
    if (str.length <= 3 && str.indexOf(':') !== -1) return false;
    // Known format strings
    const formats = ['dd-mmm-yy', 'General', 'dd/mm/yyyy', 'hh:mm:ss'];
    for (let f = 0; f < formats.length; f++) {
      if (str === formats[f]) return false;
    }
    // Number format patterns (high ratio of #, 0, ;, [ chars)
    const fmtChars = '#0;[]';
    let fmtCount = 0;
    for (let i = 0; i < str.length; i++) {
      if (fmtChars.indexOf(str[i]) !== -1) fmtCount++;
    }
    if (str.length > 3 && fmtCount > str.length * 0.4) return false;
    // Very long strings with spaces → error messages / descriptions
    if (str.length > 30) return false;
    return true;
  }

  function recordHeader(map, node, sf, context) {
    const text = node.text;
    if (!map.has(text)) map.set(text, { value: text, locations: [], contexts: new Set() });
    const entry = map.get(text);
    entry.locations.push({
      start: node.getStart(sf), end: node.getEnd(),
      line: getLineNumber(sf, node.getStart(sf)), context
    });
    entry.contexts.add(context);
  }

  function recordSheet(map, node, sf) {
    const text = node.text;
    if (!map.has(text)) map.set(text, { value: text, locations: [] });
    map.get(text).locations.push({
      start: node.getStart(sf), end: node.getEnd(),
      line: getLineNumber(sf, node.getStart(sf))
    });
  }

  /**
   * For scripts WITHOUT CONFIG_SAP: walks the AST and extracts all string
   * literals that appear in header-search contexts.
   *
   * Comments are ignored by design — the AST stores them as trivia, not nodes.
   * Variable names are Identifier nodes, not StringLiteral — also excluded.
   */
  function extractLegacyHeaders(sf) {
    const headers = new Map();
    const sheets = new Map();
    const headerArrays = [];

    walk(sf, node => {
      // ── 1. CallExpression: row.indexOf("Header"), arr.findIndex(…), arr.includes(…) ──
      if (isCallExpression(node) && isPropertyAccess(node.expression)) {
        const method = node.expression.name.text;

        if (method === 'indexOf' && node.arguments.length > 0) {
          const arg = node.arguments[0];
          if (isStringLiteral(arg) && isPlausibleHeader(arg.text)) {
            recordHeader(headers, arg, sf, 'indexOf');
          }
        }

        if (method === 'findIndex' && node.arguments.length > 0) {
          // Walk into the arrow/function body for string literals
          walk(node.arguments[0], inner => {
            if (isStringLiteral(inner) && isPlausibleHeader(inner.text)) {
              recordHeader(headers, inner, sf, 'findIndex');
            }
          });
        }

        if (method === 'includes' && node.arguments.length > 0) {
          const arg = node.arguments[0];
          if (isStringLiteral(arg) && isPlausibleHeader(arg.text)) {
            recordHeader(headers, arg, sf, 'includes');
          }
        }

        if (method === 'getWorksheet' && node.arguments.length > 0) {
          const arg = node.arguments[0];
          if (isStringLiteral(arg)) recordSheet(sheets, arg, sf);
        }
      }

      // ── 2. ArrayLiteral with ≥ 2 plausible header strings ──
      if (isArrayLiteral(node)) {
        const stringEls = node.elements.filter(el => isStringLiteral(el));
        if (stringEls.length >= 2) {
          const plausible = stringEls.filter(el => isPlausibleHeader(el.text));
          if (plausible.length >= 2) {
            headerArrays.push({
              node, headers: plausible.map(el => el.text),
              line: getLineNumber(sf, node.getStart(sf))
            });
            for (const el of plausible) recordHeader(headers, el, sf, 'arrayLiteral');
          }
        }
      }

      // ── 3. Equality comparisons: t === "Header" ──
      if (isBinaryExpression(node)) {
        const op = node.operatorToken.kind;
        if (op === ts.SyntaxKind.EqualsEqualsEqualsToken || op === ts.SyntaxKind.EqualsEqualsToken) {
          for (const side of [node.left, node.right]) {
            if (isStringLiteral(side) && isPlausibleHeader(side.text)) {
              recordHeader(headers, side, sf, 'equality');
            }
          }
        }
      }
    });

    // Convert Maps to arrays
    const hArr = []; headers.forEach(v => hArr.push(v));
    const sArr = []; sheets.forEach(v => sArr.push(v));
    return { headers: hArr, sheets: sArr, headerArrays };
  }

  // ─── SURGICAL MUTATION ──────────────────────────────────────────────

  /**
   * Mutates a single value inside CONFIG_SAP by its dotted AST path.
   *
   * @param {string} source   — Full source code
   * @param {string} path     — e.g. "columns.nDoc.header"
   * @param {string} newValue — Unquoted replacement string
   * @returns {{ success: boolean, source: string, error?: string, diff?: object }}
   */
  function mutateConfigValue(source, path, newValue) {
    const sf = parse(source);
    const config = extractConfigSAP(sf);
    if (!config) {
      return { success: false, source, error: 'CONFIG_SAP no encontrado en el código fuente.' };
    }

    const targetNode = config.nodeMap[path];
    if (!targetNode) {
      return { success: false, source, error: 'Ruta "' + path + '" no existe en CONFIG_SAP.' };
    }
    if (!isStringLiteral(targetNode)) {
      return { success: false, source, error: 'El nodo en "' + path + '" no es un literal de cadena.' };
    }

    const start = targetNode.getStart(sf);
    const end = targetNode.getEnd();
    const oldValue = targetNode.text;
    const q = source[start]; // preserve original quote character

    // Escape the new value for the quote type used
    let escaped = '';
    for (let i = 0; i < newValue.length; i++) {
      const ch = newValue[i];
      if (ch === q) { escaped += '\\' + q; }
      else if (ch === '\\') { escaped += '\\\\'; }
      else { escaped += ch; }
    }

    const newSource = source.slice(0, start) + q + escaped + q + source.slice(end);

    // Validate: re-parse and check for diagnostics
    const validation = parse(newSource);
    const diags = validation.parseDiagnostics || [];
    if (diags.length > 0) {
      const msg = diags[0].messageText;
      return { success: false, source, error: 'Sintaxis inválida: ' + (typeof msg === 'string' ? msg : msg.messageText) };
    }

    // Verify business logic immutability
    const bizBefore = source.slice(0, config.startPos) + source.slice(config.endPos);
    const newConfig = extractConfigSAP(validation);
    if (newConfig) {
      const bizAfter = newSource.slice(0, newConfig.startPos) + newSource.slice(newConfig.endPos);
      if (bizBefore !== bizAfter) {
        return { success: false, source, error: 'CRITICAL: La mutación alteró código fuera de CONFIG_SAP.' };
      }
    }

    return {
      success: true, source: newSource,
      diff: { path, oldValue, newValue, line: getLineNumber(sf, start) }
    };
  }

  /**
   * Applies multiple CONFIG_SAP mutations atomically.
   * All-or-nothing: if any mutation fails, the original source is returned.
   */
  function mutateConfigBatch(source, mutations) {
    let current = source;
    const diffs = [];
    for (const m of mutations) {
      const result = mutateConfigValue(current, m.path, m.newValue);
      if (!result.success) return { success: false, source, error: result.error, failedAt: m.path };
      current = result.source;
      diffs.push(result.diff);
    }
    return { success: true, source: current, diffs };
  }

  // ─── MIGRATION ENGINE ──────────────────────────────────────────────

  /**
   * Converts a header string to a valid camelCase JS identifier.
   * No regex — uses manual Unicode decomposition and character classification.
   */
  function headerToKey(header) {
    const norm = header.normalize('NFD');
    let cleaned = '';
    for (let i = 0; i < norm.length; i++) {
      const code = norm.charCodeAt(i);
      // Skip combining diacritical marks (U+0300..U+036F)
      if (code >= 0x0300 && code <= 0x036F) continue;
      // Skip ordinal indicators (º U+00BA, ª U+00AA)
      if (code === 0x00BA || code === 0x00AA) continue;
      const ch = norm[i];
      if ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9')) {
        cleaned += ch;
      } else if (cleaned.length > 0 && cleaned[cleaned.length - 1] !== ' ') {
        cleaned += ' '; // word boundary
      }
    }
    const words = cleaned.trim().split(' ').filter(w => w.length > 0);
    if (words.length === 0) return 'col';
    let key = '';
    for (let i = 0; i < words.length; i++) {
      let w = words[i].toLowerCase();
      if (i > 0) w = w.charAt(0).toUpperCase() + w.slice(1);
      key += w;
    }
    if (key.charAt(0) >= '0' && key.charAt(0) <= '9') {
      key = 'col' + key.charAt(0).toUpperCase() + key.slice(1);
    }
    return key;
  }

  /**
   * Determines the semantic context of a StringLiteral in the AST.
   * Used to decide which references are safe for automatic replacement.
   */
  function determineStringContext(stringNode, sf) {
    let current = stringNode.parent;
    while (current) {
      if (isCallExpression(current) && isPropertyAccess(current.expression)) {
        return current.expression.name.text; // 'indexOf','includes','setValue',etc.
      }
      if (isArrayLiteral(current)) return 'arrayLiteral';
      if (isBinaryExpression(current)) {
        const op = current.operatorToken.kind;
        if (op === ts.SyntaxKind.EqualsEqualsEqualsToken || op === ts.SyntaxKind.EqualsEqualsToken) {
          return 'equality';
        }
      }
      if (current.kind === ts.SyntaxKind.Block || current.kind === ts.SyntaxKind.SourceFile) break;
      current = current.parent;
    }
    return 'unknown';
  }

  /**
   * Migrates a legacy script to the CONFIG_SAP pattern.
   *
   * 1. Parses and extracts all headers/sheets via AST
   * 2. Generates CONFIG_SAP block
   * 3. Inserts it after function main() {
   * 4. Replaces SAFE references (indexOf, findIndex, arrayLiteral, getWorksheet)
   * 5. Validates the result
   */
  function migrateToConfigPattern(source) {
    const sf = parse(source);
    if (extractConfigSAP(sf)) {
      return { success: true, source, alreadyMigrated: true, stats: { edits: 0 } };
    }

    const legacy = extractLegacyHeaders(sf);
    if (legacy.headers.length === 0 && legacy.sheets.length === 0) {
      return { success: false, source, error: 'No se detectaron encabezados o hojas para migrar.' };
    }

    // Build column entries
    const columnEntries = {};
    const headerKeyMap = {};
    const usedKeys = new Set();

    for (const h of legacy.headers) {
      let key = headerToKey(h.value);
      let candidate = key;
      let ctr = 2;
      while (usedKeys.has(candidate)) { candidate = key + ctr; ctr++; }
      key = candidate;
      usedKeys.add(key);
      columnEntries[key] = { header: h.value, description: '' };
      headerKeyMap[h.value] = key;
    }

    // Build sheet entries
    const sheetEntries = {};
    const sheetKeyMap = {};
    const roles = ['source', 'target', 'auxiliary', 'output', 'staging'];
    let ri = 0;
    for (const s of legacy.sheets) {
      let key = headerToKey(s.value);
      if (usedKeys.has(key)) key = key + 'Sheet';
      usedKeys.add(key);
      sheetEntries[key] = { name: s.value, role: roles[ri] || 'auxiliary' };
      sheetKeyMap[s.value] = key;
      ri++;
    }

    const execution = detectExecutionParams(sf);
    const configBlock = generateConfigBlock(sheetEntries, columnEntries, execution);

    // Find insertion point: after `function main(...) {`
    const mainFunc = findFirst(sf, node =>
      isFunctionDeclaration(node) && node.name && node.name.text === 'main' && node.body
    );
    if (!mainFunc) {
      return { success: false, source, error: 'No se encontró function main() en el código fuente.' };
    }

    const bodyOpenBrace = mainFunc.body.getStart(sf);
    const insertionPoint = bodyOpenBrace + 1;

    // Phase 1: Insert CONFIG_SAP
    let result = source.slice(0, insertionPoint) + '\n' + configBlock + '\n' + source.slice(insertionPoint);

    // Phase 2: Re-parse and replace safe references
    const sf2 = parse(result);
    const config2 = extractConfigSAP(sf2);
    if (!config2) {
      return { success: false, source, error: 'Error interno: CONFIG_SAP generado no es parseable.' };
    }

    const cfgStart = config2.startPos;
    const cfgEnd = config2.endPos;

    // Contexts considered safe for automatic replacement
    const safeHeaderContexts = { indexOf: 1, findIndex: 1, arrayLiteral: 1, includes: 1 };

    const edits = [];
    walk(sf2, node => {
      if (!isStringLiteral(node)) return;
      const nStart = node.getStart(sf2);
      const nEnd = node.getEnd();
      if (nStart >= cfgStart && nEnd <= cfgEnd) return; // inside CONFIG_SAP

      const text = node.text;
      const ctx = determineStringContext(node, sf2);

      if (sheetKeyMap[text] !== undefined && ctx === 'getWorksheet') {
        edits.push({ start: nStart, end: nEnd, replacement: 'CONFIG_SAP.sheets.' + sheetKeyMap[text] + '.name' });
      } else if (headerKeyMap[text] !== undefined && safeHeaderContexts[ctx]) {
        edits.push({ start: nStart, end: nEnd, replacement: 'CONFIG_SAP.columns.' + headerKeyMap[text] + '.header' });
      }
    });

    // Apply edits from end to start (preserves positions)
    edits.sort((a, b) => b.start - a.start);
    for (const edit of edits) {
      result = result.slice(0, edit.start) + edit.replacement + result.slice(edit.end);
    }

    // Validate
    const sfFinal = parse(result);
    const diags = sfFinal.parseDiagnostics || [];
    return {
      success: diags.length === 0,
      source: result,
      warnings: diags.map(d => typeof d.messageText === 'string' ? d.messageText : d.messageText.messageText),
      stats: {
        headersDetected: Object.keys(headerKeyMap).length,
        sheetsDetected: Object.keys(sheetKeyMap).length,
        totalEdits: edits.length,
        configInserted: true
      }
    };
  }

  /**
   * Detects execution parameters from legacy script AST.
   */
  function detectExecutionParams(sf) {
    const params = { headerSearchDepth: 30, startRow: 'dynamic', minHeaderMatches: 2 };

    walk(sf, node => {
      // Math.min(30, ...) → headerSearchDepth
      if (isCallExpression(node) && isPropertyAccess(node.expression)) {
        const chain = resolvePropertyChain(node.expression, sf);
        if (chain === 'Math.min' && node.arguments.length >= 2 && isNumericLiteral(node.arguments[0])) {
          params.headerSearchDepth = Number(node.arguments[0].text);
        }
      }

      // for (let i = 10; ...) → static startRow
      if (isForStatement(node) && node.initializer &&
          node.initializer.kind === ts.SyntaxKind.VariableDeclarationList) {
        const decls = node.initializer.declarations;
        if (decls.length === 1 && decls[0].initializer && isNumericLiteral(decls[0].initializer)) {
          const val = Number(decls[0].initializer.text);
          if (val > 1) {
            // Distinguish header scan loops (have Math.min in condition) from data loops
            const cond = node.condition ? node.condition.getText(sf) : '';
            if (cond.indexOf('Math.min') === -1) params.startRow = val;
          }
        }
      }

      // coincidencias >= N → minHeaderMatches
      if (isBinaryExpression(node) &&
          node.operatorToken.kind === ts.SyntaxKind.GreaterThanEqualsToken) {
        if (isIdentifier(node.left) && node.left.text === 'coincidencias' && isNumericLiteral(node.right)) {
          params.minHeaderMatches = Number(node.right.text);
        }
      }
    });

    return params;
  }

  // ─── DATA LINEAGE ANALYSIS ─────────────────────────────────────────

  function addNode(nodes, idSet, id, type, label, line) {
    if (idSet.has(id)) return;
    idSet.add(id);
    nodes.push({ id, type, label, line: line || null });
  }

  /** Walk up to find the variable that a value is assigned to. */
  function findAssignmentTarget(node, sf) {
    let current = node.parent;
    while (current) {
      if (isVariableDeclaration(current) && current.name) return current.name.getText(sf);
      if (isBinaryExpression(current) &&
          current.operatorToken.kind === ts.SyntaxKind.EqualsToken &&
          isIdentifier(current.left)) return current.left.text;
      if (current.kind === ts.SyntaxKind.Block || current.kind === ts.SyntaxKind.SourceFile) break;
      current = current.parent;
    }
    return null;
  }

  /**
   * Traces how column values flow through the script.
   *
   * Phase 1: Column acquisition (indexOf → index variable)
   * Phase 2: Array access (values[i][colVar] → data variable)
   * Phase 3: Transforms & comparisons on tracked variables
   * Phase 4: Output operations (setValue, setColor, etc.)
   * Phase 5: Control flow (if conditions that gate outputs)
   *
   * @returns {{ nodes: Array, edges: Array }}
   */
  function analyzeLineage(sf) {
    const nodes = [];
    const edges = [];
    const nodeIds = new Set();

    // ── Phase 1: Column Acquisitions ────────────────────────
    const acquisitions = [];

    walk(sf, node => {
      if (!isCallExpression(node) || !isPropertyAccess(node.expression)) return;
      const method = node.expression.name.text;
      if (method !== 'indexOf' && method !== 'findIndex') return;

      let headerName = null;

      if (method === 'indexOf' && node.arguments.length > 0) {
        const arg = node.arguments[0];
        if (isStringLiteral(arg)) {
          headerName = arg.text;
        } else if (isPropertyAccess(arg)) {
          const chain = resolvePropertyChain(arg, sf);
          const parts = chain.split('.');
          if (parts[0] === 'CONFIG_SAP' && parts[1] === 'columns' && parts[3] === 'header') {
            headerName = parts[2];
          }
        }
      }

      if (!headerName || (!isPlausibleHeader(headerName) && headerName.length > 20)) return;

      // Find assignment target
      let indexVar = null;
      const parent = node.parent;
      if (parent && isVariableDeclaration(parent) && parent.name) {
        indexVar = parent.name.getText(sf);
      } else if (parent && isBinaryExpression(parent) &&
                 parent.operatorToken.kind === ts.SyntaxKind.EqualsToken &&
                 isIdentifier(parent.left)) {
        indexVar = parent.left.text;
      }

      if (indexVar) {
        acquisitions.push({ header: headerName, indexVar, line: getLineNumber(sf, node.getStart(sf)) });
      }
    });

    for (const acq of acquisitions) {
      const colId = 'col:' + acq.header;
      const varId = 'var:' + acq.indexVar;
      addNode(nodes, nodeIds, colId, 'column', acq.header);
      addNode(nodes, nodeIds, varId, 'variable', acq.indexVar, acq.line);
      edges.push({ from: colId, to: varId, label: 'indexOf()', type: 'acquisition' });
    }

    // ── Phase 2: Array Access ────────────────────────────────
    const dataReads = [];

    walk(sf, node => {
      if (!isElementAccess(node)) return;
      const arg = node.argumentExpression;
      if (!arg || !isIdentifier(arg)) return;

      const acq = acquisitions.find(a => a.indexVar === arg.text);
      if (!acq) return;

      const resultVar = findAssignmentTarget(node, sf);
      let arrayName = '';
      if (isElementAccess(node.expression) && isIdentifier(node.expression.expression)) {
        arrayName = node.expression.expression.text;
      }

      if (resultVar) {
        dataReads.push({ indexVar: arg.text, header: acq.header, dataVar: resultVar, sourceArray: arrayName });
        const varId = 'var:' + arg.text;
        const readId = 'var:' + resultVar;
        addNode(nodes, nodeIds, readId, 'variable', resultVar, getLineNumber(sf, node.getStart(sf)));
        const label = arrayName ? arrayName + '[·][' + arg.text + ']' : '[·][' + arg.text + ']';
        edges.push({ from: varId, to: readId, label, type: 'access' });
      }
    });

    // ── Phase 3: Transforms & Comparisons ────────────────────
    const trackedVars = new Set([
      ...acquisitions.map(a => a.indexVar),
      ...dataReads.map(d => d.dataVar)
    ]);

    walk(sf, node => {
      if (!isCallExpression(node) || !isPropertyAccess(node.expression)) return;
      const method = node.expression.name.text;
      const obj = node.expression.expression;

      // String transforms
      const transforms = ['toUpperCase', 'toLowerCase', 'trim', 'toString', 'substring', 'replace'];
      if (transforms.indexOf(method) !== -1 && isIdentifier(obj) && trackedVars.has(obj.text)) {
        const resultVar = findAssignmentTarget(node, sf);
        if (resultVar && resultVar !== obj.text) {
          addNode(nodes, nodeIds, 'var:' + resultVar, 'transform', resultVar, getLineNumber(sf, node.getStart(sf)));
          edges.push({ from: 'var:' + obj.text, to: 'var:' + resultVar, label: '.' + method + '()', type: 'transform' });
          trackedVars.add(resultVar);
        }
      }

      // includes() comparisons
      if (method === 'includes' && node.arguments.length > 0) {
        const objName = isIdentifier(obj) ? obj.text : null;
        const argNode = node.arguments[0];
        const argName = isIdentifier(argNode) ? argNode.text : null;

        if ((objName && trackedVars.has(objName)) || (argName && trackedVars.has(argName))) {
          const resultVar = findAssignmentTarget(node, sf);
          if (resultVar) {
            const cmpId = 'var:' + resultVar;
            addNode(nodes, nodeIds, cmpId, 'comparison', resultVar, getLineNumber(sf, node.getStart(sf)));
            if (objName && trackedVars.has(objName)) edges.push({ from: 'var:' + objName, to: cmpId, label: '.includes()', type: 'comparison' });
            if (argName && trackedVars.has(argName)) edges.push({ from: 'var:' + argName, to: cmpId, label: '.includes()', type: 'comparison' });
            trackedVars.add(resultVar);
          }
        }
      }

      // Output operations: setValue, setColor, setFormula, setValues
      const outputMethods = ['setColor', 'setValue', 'setFormula', 'setValues', 'setNumberFormat'];
      if (outputMethods.indexOf(method) !== -1) {
        const usedVars = [];
        walk(node, inner => {
          if (isIdentifier(inner) && trackedVars.has(inner.text) && usedVars.indexOf(inner.text) === -1) {
            usedVars.push(inner.text);
          }
        });
        // Also check the expression chain (e.g., getCell(i, colIva))
        walk(node.expression, inner => {
          if (isIdentifier(inner) && trackedVars.has(inner.text) && usedVars.indexOf(inner.text) === -1) {
            usedVars.push(inner.text);
          }
        });

        if (usedVars.length > 0) {
          let label = method;
          if (node.arguments.length > 0 && isStringLiteral(node.arguments[0])) {
            label += '("' + node.arguments[0].text + '")';
          }
          const outId = 'out:' + method + ':L' + getLineNumber(sf, node.getStart(sf));
          addNode(nodes, nodeIds, outId, 'output', label, getLineNumber(sf, node.getStart(sf)));
          for (const v of usedVars) {
            edges.push({ from: 'var:' + v, to: outId, label: method, type: 'output' });
          }
        }
      }
    });

    // ── Phase 4: Control Flow (if statements gating outputs) ──
    walk(sf, node => {
      if (node.kind !== ts.SyntaxKind.IfStatement || !node.expression) return;

      const condVars = [];
      walk(node.expression, inner => {
        if (isIdentifier(inner) && trackedVars.has(inner.text) && condVars.indexOf(inner.text) === -1) {
          condVars.push(inner.text);
        }
      });

      if (condVars.length > 0 && node.thenStatement) {
        walk(node.thenStatement, inner => {
          if (isCallExpression(inner) && isPropertyAccess(inner.expression)) {
            const m = inner.expression.name.text;
            if (['setColor', 'setValue', 'setFormula'].indexOf(m) !== -1) {
              const outId = 'out:' + m + ':L' + getLineNumber(sf, inner.getStart(sf));
              if (nodeIds.has(outId)) {
                for (const cv of condVars) {
                  const exists = edges.some(e => e.from === 'var:' + cv && e.to === outId && e.type === 'control');
                  if (!exists) edges.push({ from: 'var:' + cv, to: outId, label: 'if (·)', type: 'control' });
                }
              }
            }
          }
        });
      }
    });

    return { nodes, edges };
  }

  // ─── BLOCK DISSECTION ──────────────────────────────────────────────

  function containsCall(node, sf, methodNames) {
    let found = false;
    walk(node, n => {
      if (found) return false;
      if (isCallExpression(n) && isPropertyAccess(n.expression) &&
          methodNames.indexOf(n.expression.name.text) !== -1) {
        found = true; return false;
      }
    });
    return found;
  }

  function isHeaderSearchLoop(forStmt, sf) {
    if (forStmt.condition) {
      if (forStmt.condition.getText(sf).indexOf('Math.min') !== -1) return true;
    }
    if (forStmt.statement && containsCall(forStmt.statement, sf, ['indexOf'])) {
      let hasBreak = false;
      walk(forStmt.statement, n => { if (n.kind === ts.SyntaxKind.BreakStatement) hasBreak = true; });
      if (hasBreak) return true;
    }
    return false;
  }

  function isGuard(stmt, sf) {
    if (stmt.kind === ts.SyntaxKind.IfStatement && stmt.thenStatement) {
      const t = stmt.thenStatement.getText(sf);
      return t.indexOf('return') !== -1 && t.length < 40;
    }
    return false;
  }

  function getVarStatementName(stmt, sf) {
    if (!isVariableStatement(stmt)) return null;
    const dl = stmt.declarationList;
    if (dl && dl.declarations.length > 0) return dl.declarations[0].name.getText(sf);
    return null;
  }

  /**
   * Segments the function body into semantic blocks for UI dissection.
   */
  function dissectBlocks(sf) {
    const blocks = [];
    let stepNum = 0;

    const mainFunc = findFirst(sf, node =>
      isFunctionDeclaration(node) && node.name && node.name.text === 'main' && node.body
    );
    if (!mainFunc || !mainFunc.body) return blocks;

    const stmts = mainFunc.body.statements;
    const hasConfig = !!extractConfigSAP(sf);
    let i = 0;

    while (i < stmts.length) {
      const stmt = stmts[i];
      const startLine = getLineNumber(sf, stmt.getStart(sf));
      const endLine = getLineNumber(sf, stmt.getEnd());

      // ── CONFIG_SAP ──
      if (hasConfig && isVariableStatement(stmt) && getVarStatementName(stmt, sf) === 'CONFIG_SAP') {
        stepNum++;
        blocks.push({
          id: 'config', step: 'Paso ' + stepNum + ': Configuración',
          title: 'Objeto CONFIG_SAP',
          description: 'Bloque de configuración aislado. Este es el ÚNICO punto que se modifica al cambiar layouts de SAP. La lógica de negocio está protegida.',
          startLine, endLine, type: 'config'
        });
        i++; continue;
      }

      // ── Sheet acquisition ──
      if (containsCall(stmt, sf, ['getWorksheet', 'getActiveWorksheet', 'getWorksheets'])) {
        let end = i;
        for (let j = i + 1; j < stmts.length; j++) {
          if (containsCall(stmts[j], sf, ['getWorksheet', 'getActiveWorksheet', 'getUsedRange', 'getValues']) || isGuard(stmts[j], sf)) end = j;
          else break;
        }
        stepNum++;
        blocks.push({
          id: 'init-' + stepNum, step: 'Paso ' + stepNum + ': Inicialización',
          title: 'Conexión a Hojas de Trabajo',
          description: 'Localiza las pestañas de Excel, obtiene los rangos con datos y carga los valores en memoria.',
          startLine, endLine: getLineNumber(sf, stmts[end].getEnd()), type: 'initialization'
        });
        i = end + 1; continue;
      }

      // ── Header scanning loop ──
      if (isForStatement(stmt) && isHeaderSearchLoop(stmt, sf)) {
        stepNum++;
        blocks.push({
          id: 'headers-' + stepNum, step: 'Paso ' + stepNum + ': Localización de Encabezados',
          title: 'Escaneo Dinámico de Columnas',
          description: 'Recorre las primeras filas buscando los nombres de columna esperados. Calcula el índice numérico de cada columna para el procesamiento posterior.',
          startLine, endLine, type: 'headerScan'
        });
        i++; continue;
      }

      // ── Main data loop ──
      if (isForStatement(stmt) && !isHeaderSearchLoop(stmt, sf)) {
        const hasOutput = containsCall(stmt, sf, ['setColor', 'setValue', 'setFormula', 'setValues', 'setNumberFormat']);
        stepNum++;
        blocks.push({
          id: 'process-' + stepNum,
          step: 'Paso ' + stepNum + (hasOutput ? ': Procesamiento y Salida' : ': Barrido de Datos'),
          title: hasOutput ? 'Bucle Principal con Escritura' : 'Lectura Secuencial de Filas',
          description: hasOutput
            ? 'Recorre las filas de datos, aplica la lógica de negocio y escribe resultados (valores, fórmulas, formatos) en celdas de Excel.'
            : 'Recorre las filas del rango para lectura y análisis.',
          startLine, endLine, type: hasOutput ? 'processing' : 'dataLoop'
        });
        i++; continue;
      }

      // ── Fallback: group consecutive non-classified statements ──
      let end = i;
      for (let j = i + 1; j < stmts.length; j++) {
        if (!isForStatement(stmts[j]) && !containsCall(stmts[j], sf, ['getWorksheet', 'getActiveWorksheet'])) end = j;
        else break;
      }
      if (end >= i) {
        stepNum++;
        blocks.push({
          id: 'logic-' + stepNum, step: 'Paso ' + stepNum + ': Lógica Auxiliar',
          title: 'Procesamiento Intermedio',
          description: 'Declaraciones de variables, validaciones, y preparación de datos.',
          startLine, endLine: getLineNumber(sf, stmts[end].getEnd()), type: 'auxiliary'
        });
      }
      i = end + 1;
    }

    return blocks;
  }

  // ─── PUBLIC API ─────────────────────────────────────────────────────
  return {
    parse,
    extractConfigSAP,
    extractLegacyHeaders,
    mutateConfigValue,
    mutateConfigBatch,
    migrateToConfigPattern,
    analyzeLineage,
    dissectBlocks,
    headerToKey,
    generateConfigBlock,
    isPlausibleHeader,
    resolvePropertyChain,
    getLineNumber,
    walk,
    findNodes
  };
})();
