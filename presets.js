const PRESET_SCRIPTS = {
  "COBRANZA_COLORES (CONFIG_SAP).txt": `    /**
     * Script con CONFIG_SAP para buscar "Nº doc." en "Asignación" y colorear filas hijas.
     * Mapeado quirúrgicamente mediante AST.
     */
    function main(workbook: ExcelScript.Workbook) {
        const CONFIG_SAP = {
            version: "2.0",
            sheets: {
                current: { name: "Hoja Activa", role: "source" },
                target: { name: "IVA COBRADO", role: "target" }
            },
            columns: {
                nDoc: { header: "Nº doc.", description: "Número de documento SAP" },
                iva: { header: "IVA 16%", description: "Importe del IVA" },
                asignacion: { header: "Asignación", description: "Columna de Asignación" },
                importe: { header: "Importe en ML", description: "Importe en Moneda Local" },
                bcoTarget: { header: "Bco.prp.", description: "Banco propio en destino" }
            },
            execution: {
                headerSearchDepth: 30,
                startRow: "dynamic",
                minHeaderMatches: 2
            }
        };

        const currentSheet = workbook.getActiveWorksheet();
        const targetSheet = workbook.getWorksheet(CONFIG_SAP.sheets.target.name);

        if (!targetSheet) {
            throw new Error("No se encontró la pestaña '" + CONFIG_SAP.sheets.target.name + "'.");
        }

        const currentRange = currentSheet.getUsedRange();
        const targetRange = targetSheet.getUsedRange();

        if (!currentRange || !targetRange) return;

        const currentValues = currentRange.getValues() as string[][];
        const targetValues = targetRange.getValues() as string[][];

        // Buscar encabezados en la hoja actual
        let colNDoc = -1;
        let colIva = -1;
        let headerRowCurrent = -1;

        for (let i = 0; i < Math.min(CONFIG_SAP.execution.headerSearchDepth, currentValues.length); i++) {
            const row = currentValues[i].map(h => h?.toString().trim() || "");
            
            const idxDoc = row.indexOf(CONFIG_SAP.columns.nDoc.header);
            const idxIva = row.indexOf(CONFIG_SAP.columns.iva.header);
            
            if (idxDoc !== -1) colNDoc = idxDoc;
            if (idxIva !== -1) colIva = idxIva;
            
            if (colNDoc !== -1 && colIva !== -1) {
                headerRowCurrent = i;
                break;
            }
        }

        // Buscar encabezados en la hoja destino
        let colAsignacion = -1;
        let colImporte = -1;
        let colBcoTarget = -1;
        let headerRowTarget = -1;

        for (let i = 0; i < Math.min(CONFIG_SAP.execution.headerSearchDepth, targetValues.length); i++) {
            const row = targetValues[i].map(h => h?.toString().trim() || "");
            
            const idxAsig = row.indexOf(CONFIG_SAP.columns.asignacion.header);
            const idxImp = row.indexOf(CONFIG_SAP.columns.importe.header);
            const idxBco = row.indexOf(CONFIG_SAP.columns.bcoTarget.header);
            
            if (idxAsig !== -1) colAsignacion = idxAsig;
            if (idxImp !== -1) colImporte = idxImp;
            if (idxBco !== -1) colBcoTarget = idxBco;

            if (colAsignacion !== -1 && colImporte !== -1) {
                headerRowTarget = i;
                break;
            }
        }

        if (colNDoc === -1 || colIva === -1 || colAsignacion === -1 || colImporte === -1) {
            throw new Error("Faltan encabezados esenciales en el libro de Excel.");
        }

        const highlightColor = "#FFFF00";

        for (let i = headerRowCurrent + 1; i < currentValues.length; i++) {
            const nDoc = currentValues[i][colNDoc].toString().trim();
            if (!nDoc) continue;

            const nDocUpper = nDoc.toUpperCase();
            let hasMatch = false;

            for (let j = headerRowTarget + 1; j < targetValues.length; j++) {
                const asignacion = targetValues[j][colAsignacion]?.toString().trim();
                if (!asignacion) continue;

                const asigUpper = asignacion.toUpperCase();
                if (asigUpper.includes(nDocUpper) || nDocUpper.includes(asigUpper)) {
                    hasMatch = true;
                    break;
                }
            }

            if (hasMatch) {
                currentSheet.getCell(i, colIva).getFormat().getFill().setColor(highlightColor);
            }
        }
    }`,

  "COBRANZA_COLORES V2.txt": `    /**
    * Script reformado (MÁS DINÁMICO) para buscar "Nº doc." en "Asignación" y colorear filas hijas.
    */
    function main(workbook: ExcelScript.Workbook) {
        const currentSheet = workbook.getActiveWorksheet();
        const targetSheet = workbook.getWorksheet("IVA COBRADO");

        if (!targetSheet) {
            throw new Error("No se encontró la pestaña 'IVA COBRADO'.");
        }

        const currentRange = currentSheet.getUsedRange();
        const targetRange = targetSheet.getUsedRange();

        if (!currentRange || !targetRange) return;

        const currentValues = currentRange.getValues() as string[][];
        const targetValues = targetRange.getValues() as string[][];

        // Buscar encabezados en la hoja actual
        let colNDoc = -1;
        let colIva = -1;
        let headerRowCurrent = -1;

        for (let i = 0; i < Math.min(30, currentValues.length); i++) {
            const row = currentValues[i].map(h => h?.toString().trim() || "");
            
            const idxDoc = row.indexOf("Nº doc.");
            const idxIva = row.indexOf("IVA 16%");
            
            if (idxDoc !== -1) colNDoc = idxDoc;
            if (idxIva !== -1) colIva = idxIva;
            
            if (colNDoc !== -1 && colIva !== -1) {
                headerRowCurrent = i;
                break;
            }
        }

        // Buscar encabezados en la hoja destino
        let colAsignacion = -1;
        let colImporte = -1;
        let colBcoTarget = -1;
        let headerRowTarget = -1;

        for (let i = 0; i < Math.min(30, targetValues.length); i++) {
            const row = targetValues[i].map(h => h?.toString().trim() || "");
            
            const idxAsig = row.indexOf("Asignación");
            const idxImp = row.indexOf("Importe en ML");
            const idxBco = row.indexOf("Bco.prp.");
            
            if (idxAsig !== -1) colAsignacion = idxAsig;
            if (idxImp !== -1) colImporte = idxImp;
            if (idxBco !== -1) colBcoTarget = idxBco;

            if (colAsignacion !== -1 && colImporte !== -1) {
                headerRowTarget = i;
                break;
            }
        }

        if (colNDoc === -1 || colIva === -1 || colAsignacion === -1 || colImporte === -1) {
            throw new Error("Faltan encabezados (Nº doc., IVA 16%, Asignación o Importe en ML).");
        }

        // Color universal para los resultados (Amarillo)
        const highlightColor = "#FFFF00";

        // Función auxiliar para extraer números verdaderos ignorando símbolos como $, MXN, espacios, etc.
        function parseDinamicNumber(val: string): number {
            const cleanStr = val.replace(/[^0-9\.\-]/g, "");
            return parseFloat(cleanStr);
        }

        // Recorremos la hoja actual
        for (let i = headerRowCurrent + 1; i < currentValues.length; i++) {
            const nDoc = currentValues[i][colNDoc].toString().trim();
            
            if (!nDoc) continue; // Solo procesa filas "padre"

            const nDocUpper = nDoc.toUpperCase();

            // Buscamos coincidencia dinámica en la "Asignación"
            let hasMatch = false;
            for (let j = headerRowTarget + 1; j < targetValues.length; j++) {
                const asignacion = targetValues[j][colAsignacion]?.toString().trim();
                if (!asignacion) continue;

                const asigUpper = asignacion.toUpperCase();

                if (asigUpper.includes(nDocUpper) || nDocUpper.includes(asigUpper)) {
                    hasMatch = true;
                    break;
                }
            }

            // Si hay coincidencia, coloreamos la fila hija de la hoja actual
            if (hasMatch) {
                currentSheet.getCell(i, colIva).getFormat().getFill().setColor(highlightColor);
            }
        }
    }`,

  "CONCILIACION ENTRE AUXILIAR Y PAPEL DE TRABAJO DAVID.txt": `/**
 * SISTEMA DE CONCILIACIÓN DINÁMICA E INTELIGENTE
 * 
 * Este script procesa el auxiliar actual:
 * 1. Identifica los encabezados clave (Asignación, Nº doc., Importe en ML).
 * 2. Agrega los encabezados "POLIZA", "FECHA", "IVA".
 * 3. Llenado dinámico de la columna "POLIZA" evaluando la "Referencia":
 *    - Si empieza con "G", copia el valor de "Nº doc."
 *    - Si no, extrae los primeros 10 caracteres de "Referencia" (=IZQUIERDA).
 */
function main(workbook: ExcelScript.Workbook) {
    const auxiliarSheet = workbook.getActiveWorksheet();
    const nombreAuxiliar = auxiliarSheet.getName();
    const usedRange = auxiliarSheet.getUsedRange();
    
    if (!usedRange) {
        console.log("La hoja auxiliar está vacía.");
        return;
    }

    const values = usedRange.getValues() as string[][];
    const targetHeaders = ["Asignación", "Nº doc.", "Importe en ML"];
    
    let filaDeEncabezados = -1;
    let ultimaColumna = -1;
    
    let colReferencia = -1;
    let colNumDoc = -1;

    for (let r = 0; r < values.length; r++) {
        const row = values[r].map(item => item ? item.toString().trim() : "");
        
        let coincidencias = 0;
        for (const target of targetHeaders) {
            if (row.includes(target)) coincidencias++;
        }
        
        if (coincidencias >= 2) {
            filaDeEncabezados = r;
            
            for (let c = 0; c < row.length; c++) {
                if (row[c] !== "") {
                    ultimaColumna = c;
                }
            }
            
            colNumDoc = row.indexOf("Nº doc.");
            colReferencia = row.indexOf("Referencia");
            
            break; 
        }
    }

    if (filaDeEncabezados === -1) {
        console.log("Error: No se ubicaron los encabezados básicos (Asignación, Nº doc., etc).");
        return;
    }
    
    if (colNumDoc === -1 || colReferencia === -1) {
        console.log("Advertencia: Falta la columna 'Nº doc.' o 'Referencia'. No se puede automatizar la columna POLIZA.");
        return;
    }

    const nuevosEncabezados = ["POLIZA", "FECHA", "IVA"];
    const colPoliza = ultimaColumna + 1;
    
    for (let i = 0; i < nuevosEncabezados.length; i++) {
        const celdaNueva = usedRange.getCell(filaDeEncabezados, colPoliza + i);
        celdaNueva.setValue(nuevosEncabezados[i]);
        celdaNueva.getFormat().getFont().setBold(true);
        celdaNueva.getFormat().getFill().setColor("#D9E1F2");
    }
    
    let primeraFilaAuxiliar = -1;
    for (let r = filaDeEncabezados + 1; r < values.length; r++) {
        const row = values[r];
        const refVal = row[colReferencia] ? row[colReferencia].toString().trim() : "";
        const docVal = row[colNumDoc] ? row[colNumDoc].toString().trim() : "";
        
        if (!refVal && !docVal) continue;
        
        if (primeraFilaAuxiliar === -1) {
            primeraFilaAuxiliar = r;
        }
        
        let polizaFinal = "";
        
        if (refVal.toUpperCase().includes("IVA")) {
            polizaFinal = "";
        } else if (refVal.toUpperCase().startsWith("G")) {
            polizaFinal = docVal;
        } else {
            polizaFinal = refVal.substring(0, 10);
        }
        
        usedRange.getCell(r, colPoliza).setValue(polizaFinal);
    }
    
    let addressPrimeraPoliza = "";
    if (primeraFilaAuxiliar !== -1) {
        addressPrimeraPoliza = auxiliarSheet.getCell(primeraFilaAuxiliar, colPoliza).getAddress();
    }
    
    let detalleSheet = workbook.getWorksheet("Detalle de conciliación");
    if (!detalleSheet) {
        detalleSheet = workbook.addWorksheet("Detalle de conciliación");
    } else {
        detalleSheet.getUsedRange()?.clear();
    }
    
    let filaActual = 2;
    const todasLasHojas = workbook.getWorksheets();
    const headersDetalle = ["POLIZA", "IVA", "FECHA", "", "", "IVA", "FECHA"];
    
    for (const hoja of todasLasHojas) {
        if (hoja.getName().toUpperCase().includes("ZN")) {
            const zRange = hoja.getUsedRange();
            if (!zRange) continue;
            
            const zValues = zRange.getValues() as string[][];
            let zHeaderRow = -1;
            let colIva16 = -1;
            let colIva8 = -1;
            let colDoctoZN = -1;
            let colFechaPago = -1;
            
            for (let r = 0; r < Math.min(30, zValues.length); r++) {
                const row = zValues[r].map(v => v ? v.toString().trim().toUpperCase() : "");
                
                const idx16 = row.indexOf("IVA AL 16%");
                const idx8 = row.indexOf("IVA AL 8%");
                const idxDocto = row.indexOf("DOCTO CONTABLE");
                const idxFecha = row.indexOf("FECHA DE PAGO");
                
                if (idx16 !== -1) colIva16 = idx16;
                if (idx8 !== -1) colIva8 = idx8;
                if (idxDocto !== -1) colDoctoZN = idxDocto;
                if (idxFecha !== -1) colFechaPago = idxFecha;
                
                if (colDoctoZN !== -1 && (colIva16 !== -1 || colIva8 !== -1)) {
                    zHeaderRow = r;
                    break;
                }
            }
            
            if (zHeaderRow === -1 || colDoctoZN === -1) {
                continue;
            }
            
            const procesarBloque = (colIvaIndex: number, porcentajeTexto: string) => {
                if (colIvaIndex === -1) return;
                
                let datosAExtraer: {docto: string, iva: string|number, fecha: string|number}[] = [];
                
                for (let r = zHeaderRow + 1; r < zValues.length; r++) {
                    const ivaText = zValues[r][colIvaIndex] ? zValues[r][colIvaIndex].toString().replace(/[^0-9.-]/g, "") : "";
                    const ivaNum = parseFloat(ivaText);
                    
                    if (!isNaN(ivaNum) && ivaNum !== 0) {
                        const doctoStr = zValues[r][colDoctoZN] ? zValues[r][colDoctoZN].toString().trim() : "";
                        
                        if (doctoStr !== "") {
                            const ivaOriginal = zValues[r][colIvaIndex] ? zValues[r][colIvaIndex] : "";
                            const fechaOriginal = (colFechaPago !== -1 && zValues[r][colFechaPago]) ? zValues[r][colFechaPago] : "";
                            
                            datosAExtraer.push({
                                docto: doctoStr,
                                iva: ivaOriginal as string|number,
                                fecha: fechaOriginal as string|number
                            });
                        }
                    }
                }
                
                if (datosAExtraer.length > 0) {
                    const numRows = datosAExtraer.length;
                    const dataRowStart = filaActual + 3; 
                    const dataRowEnd = dataRowStart + numRows - 1;
                    const lookupRange = `'Detalle de conciliación'!$B$${dataRowStart}:$D$${dataRowEnd}`;
                    
                    if (filaActual - 1 >= 0) {
                        detalleSheet.getCell(filaActual - 1, 2).setValue("FECHA");
                        detalleSheet.getCell(filaActual - 1, 2).getFormat().getFont().setBold(true);
                        detalleSheet.getCell(filaActual - 1, 3).setFormula(`=VLOOKUP(${addressPrimeraPoliza}, ${lookupRange}, 3, FALSE)`);
                        detalleSheet.getCell(filaActual - 1, 3).getFormat().getFont().setBold(true);
                        detalleSheet.getCell(filaActual - 1, 3).setNumberFormat("dd-mmm-yy");
                    }
                    
                    detalleSheet.getCell(filaActual, 2).setValue("IVA");
                    detalleSheet.getCell(filaActual, 2).getFormat().getFont().setBold(true);
                    detalleSheet.getCell(filaActual, 3).setFormula(`=VLOOKUP(${addressPrimeraPoliza}, ${lookupRange}, 2, FALSE)`);
                    detalleSheet.getCell(filaActual, 3).setNumberFormat("#,##0.00");
                    detalleSheet.getCell(filaActual, 3).getFormat().getFont().setBold(true);
                    
                    detalleSheet.getCell(filaActual, 1).setValue(`${hoja.getName()} ${porcentajeTexto}`);
                    detalleSheet.getCell(filaActual, 1).getFormat().getFont().setBold(true);
                    
                    filaActual++;
                    
                    const headersRange = detalleSheet.getRangeByIndexes(filaActual, 1, 1, headersDetalle.length);
                    headersRange.setValues([headersDetalle]);
                    headersRange.getFormat().getFont().setBold(true);
                    headersRange.getFormat().getFill().setColor("#E2EFDA");
                    filaActual++;
                    
                    const blockArr: (string | number)[][] = [];
                    const formatArr: string[][] = [];
                    
                    for (let i = 0; i < numRows; i++) {
                        const data = datosAExtraer[i];
                        const excelRow = filaActual + i + 1;
                        
                        let doctoVal: string | number = data.docto;
                        if (!isNaN(Number(data.docto)) && data.docto.trim() !== "") {
                            doctoVal = Number(data.docto);
                        }
                        
                        blockArr.push([
                            doctoVal, 
                            data.iva, 
                            data.fecha, 
                            `=C${excelRow}-G${excelRow}`,
                            "", 
                            `=VLOOKUP(B${excelRow}, '${nombreAuxiliar}'!$B:$K, 10, FALSE)`,
                            `=VLOOKUP(B${excelRow}, '${nombreAuxiliar}'!$B:$I, 8, FALSE)`
                        ]);

                        formatArr.push([
                            "General",
                            "#,##0.00",
                            "dd-mmm-yy",
                            "#,##0.00;[Red]-#,##0.00",
                            "General", 
                            "#,##0.00",
                            "dd-mmm-yy"
                        ]);
                    }
                    
                    const dataRange = detalleSheet.getRangeByIndexes(filaActual, 1, numRows, 7);
                    dataRange.setNumberFormats(formatArr);
                    dataRange.setValues(blockArr);
                    
                    filaActual += numRows;
                    filaActual += 4;
                }
            };
            
            procesarBloque(colIva16, "16%");
            procesarBloque(colIva8, "8%");
        }
    }
    
    console.log("Proceso ejecutado: Múltiples bloques de Detalle generados dinámicamente.");
}`,

  "CONCILIACION_PROPUESTAS_PARA MI.txt": `/**
 * Macro para consolidar pagos de proveedores y conciliar contra corte bancario.
 * REESCRITO PARA OFFICE SCRIPTS (Excel Web/Desktop).
 */
function main(workbook: ExcelScript.Workbook): void {
    const sheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();
    const usedRange: ExcelScript.Range | undefined = sheet.getUsedRange();

    if (!usedRange) {
        console.log("La hoja activa está vacía.");
        return;
    }

    const data: (string | number | boolean)[][] = usedRange.getValues();
    if (data.length <= 1) return;

    const headers: string[] = data[0].map((h: string | number | boolean) =>
        String(h).trim().toUpperCase().replace(/\s+/g, ' ')
    );

    const idxFecha: number = headers.indexOf("FECHA DE PAGO");
    const idxProveedor: number = headers.indexOf("NOMBRE DEL PROVEEDOR");
    const idxObs: number = headers.indexOf("OBSERVACIONES");
    const idxTotal: number = headers.indexOf("NETO");

    const faltantes: string[] = [];
    if (idxFecha === -1) faltantes.push("FECHA DE PAGO");
    if (idxProveedor === -1) faltantes.push("NOMBRE DEL PROVEEDOR");
    if (idxTotal === -1) faltantes.push("NETO");
    if (idxObs === -1) faltantes.push("OBSERVACIONES");

    if (faltantes.length > 0) {
        throw new Error(\`Faltan columnas en la fila 1: [ \${faltantes.join(", ")} ].\`);
    }

    function parseDateToInteger(val: string | number | boolean): number {
        if (typeof val === 'number') {
            const dateInfo = new Date(Date.UTC(1899, 11, 30));
            dateInfo.setUTCDate(dateInfo.getUTCDate() + Math.floor(val));
            return (dateInfo.getUTCFullYear() * 10000) + ((dateInfo.getUTCMonth() + 1) * 100) + dateInfo.getUTCDate();
        }
        if (typeof val === 'string') {
            let s = val.replace(/[\.\-]/g, '/').trim();
            let parts = s.split('/');
            if (parts.length === 3) {
                let p0 = parseInt(parts[0], 10), p1 = parseInt(parts[1], 10), p2 = parseInt(parts[2], 10);
                let d = 0, m = 0, y = 0;
                if (p2 >= 1000) { y = p2; d = p0; m = p1; }
                else if (p0 >= 1000) { y = p0; m = p1; d = p2; }
                else { y = p2 + 2000; d = p0; m = p1; }
                return (y * 10000) + (m * 100) + d;
            }
        }
        return 0;
    }

    let maxDateInt: number = 0;
    for (let i = 1; i < data.length; i++) {
        const currentInt = parseDateToInteger(data[i][idxFecha]);
        if (currentInt > maxDateInt) maxDateInt = currentInt;
    }

    if (maxDateInt === 0) {
        console.log("No se encontraron fechas válidas.");
        return;
    }

    const totalesPorProveedor: { [proveedor: string]: number } = {};
    const ultimaFilaProveedor: { [proveedor: string]: number } = {};

    for (let i = 1; i < data.length; i++) {
        const currentInt = parseDateToInteger(data[i][idxFecha]);
        if (currentInt === maxDateInt) {
            const proveedor: string = String(data[i][idxProveedor]).trim();
            const totalVal: string | number | boolean = data[i][idxTotal];
            const total: number = typeof totalVal === 'number' ? totalVal : parseFloat(String(totalVal)) || 0;

            if (proveedor) {
                if (totalesPorProveedor[proveedor] === undefined) {
                    totalesPorProveedor[proveedor] = 0;
                }
                totalesPorProveedor[proveedor] += total;
                ultimaFilaProveedor[proveedor] = i;
            }
        }
    }

    const colDestinoIndex: number = idxObs + 1;
    const numRows: number = usedRange.getRowCount();

    if (numRows > 1) {
        sheet.getRangeByIndexes(1, colDestinoIndex, numRows - 1, 1).clear(ExcelScript.ClearApplyTo.contents);
    }

    sheet.getCell(0, colDestinoIndex).setValue("TOTAL CONSOLIDADO");

    const pagosConsolidadosArray: number[] = [];

    for (const proveedor in totalesPorProveedor) {
        const filaIndice: number = ultimaFilaProveedor[proveedor];
        const sumaTotal: number = Math.abs(totalesPorProveedor[proveedor]);

        const cell = sheet.getCell(filaIndice, colDestinoIndex);
        cell.setValue(sumaTotal);
        cell.setNumberFormat("#,##0.00");

        if (sumaTotal > 0.001) {
            pagosConsolidadosArray.push(sumaTotal);
        }
    }

    let diffSheet: ExcelScript.Worksheet | undefined = undefined;
    for (const ws of workbook.getWorksheets()) {
        if (ws.getName().trim().toUpperCase() === "REPORTE DIARIO") {
            diffSheet = ws;
            break;
        }
    }

    if (!diffSheet) {
        console.log("No se encontró la pestaña 'REPORTE DIARIO'. Saltando paso 2.");
        return;
    }

    const diffRange = diffSheet.getUsedRange();
    if (!diffRange) return;

    const diffData = diffRange.getValues();
    let headerRow = -1;

    const colCorte = 5;
    const colPagos = 6;
    const colDif = 7;

    for (let r = 0; r < diffData.length; r++) {
        const valF = String(diffData[r][colCorte]).replace(/\s+/g, '').toUpperCase();
        if (valF === "CORTEBANCARIO") {
            headerRow = r;
            break;
        }
    }

    if (headerRow === -1) {
        throw new Error("No se encontró el encabezado 'CORTE BANCARIO' en la columna F de la hoja 'Reporte diario'.");
    }

    let corteBancarioArray: number[] = [];
    for (let r = headerRow + 1; r < diffData.length; r++) {
        const val = diffData[r][colCorte];
        if (val !== "" && val !== null) {
            const num = typeof val === 'number' ? val : parseFloat(String(val)) || 0;
            if (num > 0.001) corteBancarioArray.push(num);
        }
    }

    corteBancarioArray.sort((a, b) => a - b);
    pagosConsolidadosArray.sort((a, b) => a - b);

    let i = 0;
    let j = 0;

    let alignedCorte: (number | string)[] = [];
    let alignedPagos: (number | string)[] = [];

    const epsilon = 0.01;

    while (i < corteBancarioArray.length || j < pagosConsolidadosArray.length) {
        if (i < corteBancarioArray.length && j < pagosConsolidadosArray.length) {
            let valCorte = corteBancarioArray[i];
            let valPago = pagosConsolidadosArray[j];

            if (Math.abs(valCorte - valPago) <= epsilon) {
                alignedCorte.push(valCorte);
                alignedPagos.push(valPago);
                i++; j++;
            } else if (valCorte < valPago) {
                alignedCorte.push(valCorte);
                alignedPagos.push("");
                i++;
            } else {
                alignedCorte.push("");
                alignedPagos.push(valPago);
                j++;
            }
        } else if (i < corteBancarioArray.length) {
            alignedCorte.push(corteBancarioArray[i]);
            alignedPagos.push("");
            i++;
        } else {
            alignedCorte.push("");
            alignedPagos.push(pagosConsolidadosArray[j]);
            j++;
        }
    }

    const lastRowToClear = Math.max(diffData.length, headerRow + alignedCorte.length + 5);
    if (lastRowToClear > headerRow + 1) {
        diffSheet.getRangeByIndexes(headerRow + 1, colCorte, lastRowToClear - headerRow, 3).clear(ExcelScript.ClearApplyTo.contents);
    }

    diffSheet.getCell(headerRow, colCorte).setValue("CORTE BANCARIO");
    diffSheet.getCell(headerRow, colPagos).setValue("PAGOS HECHOS");
    diffSheet.getCell(headerRow, colDif).setValue("DIFERENCIA");

    if (alignedCorte.length > 0) {
        const outCorte = alignedCorte.map(v => [v]);
        const outPagos = alignedPagos.map(v => [v]);

        diffSheet.getRangeByIndexes(headerRow + 1, colCorte, alignedCorte.length, 1).setValues(outCorte);
        diffSheet.getRangeByIndexes(headerRow + 1, colPagos, alignedPagos.length, 1).setValues(outPagos);
        diffSheet.getRangeByIndexes(headerRow + 1, colCorte, alignedCorte.length, 2).setNumberFormat("#,##0.00");

        const formulaRange = diffSheet.getRangeByIndexes(headerRow + 1, colDif, alignedCorte.length, 1);
        const formulas = alignedCorte.map(() => ["=N(RC[-2]) - N(RC[-1])"]);
        formulaRange.setFormulasR1C1(formulas);
        formulaRange.setNumberFormat("#,##0.00");
    }
}`,

  "ENCONTRAR FECHA REAL EN POR FACTURA DAVID.txt": `/**
 * SISTEMA DE CONCILIACIÓN DINÁMICA - FASE 1 y 2
 * Construcción paso a paso según los requerimientos indicados.
 */
function main(workbook: ExcelScript.Workbook) {
  const sheetFactura = workbook.getWorksheet("por factura");
  const sheetPartidas = workbook.getWorksheet("todas las partidas");

  if (!sheetFactura && !sheetPartidas) {
    console.log("No se encontraron las pestañas objetivo.");
    return;
  }

  const encabezadosClave = [
    "St", "Referencia", "Asignación", "Nº doc.", "Clase", 
    "Fecha doc.", "Doc.comp.", "Compens.", "IO", "II", 
    "Importe en ML", "ML", "Mon.", "Importe en MD", "Texto"
  ];

  function procesarEstructura(sheet: ExcelScript.Worksheet, esPorFactura: boolean) {
    if (!sheet) return null;
    const usedRange = sheet.getUsedRange();
    if (!usedRange) return null;

    const valores = usedRange.getValues() as string[][];
    const offsetRow = usedRange.getRowIndex();

    const filasDeEncabezados: { rangeIdx: number, absRow: number, indicesValores: { numDoc: number; clase: number; fechaDoc: number; docComp: number; } }[] = [];

    for (let r = 0; r < valores.length; r++) {
      const filaTextos = valores[r].map(c => c ? c.toString().trim() : "");
      let coincidencias = 0;
      for (const encabezado of encabezadosClave) {
        if (filaTextos.some(t => t === encabezado || t === "Compens/")) {
          coincidencias++;
        }
      }

      if (coincidencias >= 4) {
        let absRow = offsetRow + r;
        
        if (filasDeEncabezados.length === 0 || (absRow - filasDeEncabezados[filasDeEncabezados.length - 1].absRow > 4)) {
          const mapCols = {
            numDoc: filaTextos.findIndex(t => t === "Nº doc." || t === "N° doc." || t === "N doc."),
            clase: filaTextos.findIndex(t => t === "Clase"),
            fechaDoc: filaTextos.findIndex(t => t === "Fecha doc." || t === "Fecha doc/"),
            docComp: filaTextos.findIndex(t => t === "Doc.comp." || t === "Doc. comp.")
          };

          filasDeEncabezados.push({ rangeIdx: r, absRow: absRow, indicesValores: mapCols });
        }
      }
    }

    if (esPorFactura && filasDeEncabezados.length > 0) {
      const primerHeaderAbsIdx = filasDeEncabezados[0].absRow;
      const filaArribaAbsIndex = primerHeaderAbsIdx > 0 ? primerHeaderAbsIdx - 1 : 0;
      
      const columnaK = sheet.getRange("K:K");
      columnaK.insert(ExcelScript.InsertShiftDirection.right);

      const celdaFechaPago = sheet.getCell(filaArribaAbsIndex, 10);
      celdaFechaPago.setValue("Fecha de pago");
      celdaFechaPago.getFormat().getFont().setBold(true);

      const celdaComisiones = sheet.getCell(filaArribaAbsIndex, 11);
      celdaComisiones.setValue("Comisiones");
      celdaComisiones.getFormat().getFont().setBold(true);
    }

    return { sheet, valores, filasDeEncabezados, offsetRow };
  }

  const infoFactura = procesarEstructura(sheetFactura, true);
  const infoPartidas = procesarEstructura(sheetPartidas, false);

  if (infoFactura && infoPartidas && infoFactura.filasDeEncabezados.length > 0 && infoPartidas.filasDeEncabezados.length > 0) {
    const colMapPartidas = infoPartidas.filasDeEncabezados[0].indicesValores;
    const colMapFactura = infoFactura.filasDeEncabezados[0].indicesValores;

    if (colMapPartidas.clase !== -1 && colMapPartidas.numDoc !== -1 && colMapFactura.docComp !== -1) {
      
      for (let p = 0; p < infoPartidas.valores.length; p++) {
        const filaP = infoPartidas.valores[p];
        const clasePartida = filaP[colMapPartidas.clase] ? filaP[colMapPartidas.clase].toString().trim() : "";
        
        if (clasePartida === "DZ") {
          const numDocDZ = filaP[colMapPartidas.numDoc] ? filaP[colMapPartidas.numDoc].toString().trim() : "";
          const fechaDocDZ = filaP[colMapPartidas.fechaDoc] ? filaP[colMapPartidas.fechaDoc].toString().trim() : "";

          if (!numDocDZ) continue;

          let docCompsBuscados: string[] = [];

          for (let f = 0; f < infoFactura.valores.length; f++) {
            const f_numDoc = infoFactura.valores[f][colMapFactura.numDoc] ? infoFactura.valores[f][colMapFactura.numDoc].toString().trim() : "";
            const f_clase = infoFactura.valores[f][colMapFactura.clase] ? infoFactura.valores[f][colMapFactura.clase].toString().trim() : "";
            
            const isMatch = (f_numDoc === numDocDZ) || (f_numDoc !== "" && numDocDZ !== "" && Number(f_numDoc.replace(/,/g, "")) === Number(numDocDZ.replace(/,/g, "")));

            if (isMatch && f_clase === "DZ") {
              const docCompActual = infoFactura.valores[f][colMapFactura.docComp] ? infoFactura.valores[f][colMapFactura.docComp].toString().trim() : "";
              
              if (docCompActual.startsWith("1000")) {
                const absRowF = f + infoFactura.offsetRow;
                sheetFactura.getCell(absRowF, 10).setValue(docCompActual);
              } else if (docCompActual !== "") {
                docCompsBuscados.push(docCompActual);
              }
            }
          }

          for (const docComp of docCompsBuscados) {
            for (let f2 = 0; f2 < infoFactura.valores.length; f2++) {
              const f2_clase = infoFactura.valores[f2][colMapFactura.clase] ? infoFactura.valores[f2][colMapFactura.clase].toString().trim() : "";
              const f2_docComp = infoFactura.valores[f2][colMapFactura.docComp] ? infoFactura.valores[f2][colMapFactura.docComp].toString().trim() : "";
              
              if (f2_clase === "RV" && f2_docComp === docComp) {
                const absRowF2 = f2 + infoFactura.offsetRow;
                sheetFactura.getCell(absRowF2, 10).setValue(fechaDocDZ);
              }
            }
          }
        }
      }
      console.log("Cruce de datos DZ->RV con Fechas completado.");
    }
  }

  if (sheetFactura) {
    const finalUsedRange = sheetFactura.getUsedRange();
    if (finalUsedRange) {
      const finalVals = finalUsedRange.getValues() as string[][];
      const startR = finalUsedRange.getRowIndex();
      const startC = finalUsedRange.getColumnIndex();

      const relJ = 9 - startC;
      const relK = 10 - startC;

      for (let r = 0; r < finalVals.length; r++) {
        const absRow = startR + r;

        const convertirFecha = (relIdx: number, absCol: number) => {
          if (relIdx >= 0 && relIdx < finalVals[r].length) {
            const val = finalVals[r][relIdx] ? finalVals[r][relIdx].toString().trim() : "";
            const match = val.match(/^(\d{2})\.(\d{2})\.(\d{2,4})$/);
            if (match) {
              let dia = match[1];
              let mes = match[2];
              let anio = match[3];
              if (anio.length === 2) anio = "20" + anio;
              
              const isoDate = \`\${anio}-\${mes}-\${dia}\`;
              const celdaFecha = sheetFactura.getCell(absRow, absCol);
              celdaFecha.setValue(isoDate);
              celdaFecha.setNumberFormat("dd-mmm-yy");
            }
          }
        };

        const convertirANumero = (relIdx: number, absCol: number) => {
          if (relIdx >= 0 && relIdx < finalVals[r].length) {
            let strVal = finalVals[r][relIdx] ? finalVals[r][relIdx].toString().trim() : "";
            if (strVal !== "") {
              if (strVal.endsWith("-")) {
                strVal = "-" + strVal.substring(0, strVal.length - 1);
              }
              let saneado = strVal.replace(/,/g, ""); 
              let num = Number(saneado);
              if (!isNaN(num)) {
                sheetFactura.getCell(absRow, absCol).setValue(num);
              }
            }
          }
        };

        convertirFecha(relJ, 9);
        convertirFecha(relK, 10);
        convertirANumero(13 - startC, 13);
        convertirANumero(16 - startC, 16);
      }
    }

    sheetFactura.getRange("N:N").setNumberFormat("#,##0.00;[Red]-#,##0.00");
    sheetFactura.getRange("Q:Q").setNumberFormat("#,##0.00;[Red]-#,##0.00");
  }

  if (sheetPartidas) {
    const partidasUsedRange = sheetPartidas.getUsedRange();
    if (partidasUsedRange) {
      const pVals = partidasUsedRange.getValues() as string[][];
      const pStartR = partidasUsedRange.getRowIndex();
      const pStartC = partidasUsedRange.getColumnIndex();

      for (let r = 0; r < pVals.length; r++) {
        const absRow = pStartR + r;

        const convertirANumeroPartidas = (relIdx: number, absCol: number) => {
          if (relIdx >= 0 && relIdx < pVals[r].length) {
            let strVal = pVals[r][relIdx] ? pVals[r][relIdx].toString().trim() : "";
            if (strVal !== "") {
              if (strVal.endsWith("-")) {
                strVal = "-" + strVal.substring(0, strVal.length - 1);
              }
              let saneado = strVal.replace(/,/g, ""); 
              let num = Number(saneado);
              if (!isNaN(num)) {
                sheetPartidas.getCell(absRow, absCol).setValue(num);
              }
            }
          }
        };

        convertirANumeroPartidas(12 - pStartC, 12);
        convertirANumeroPartidas(15 - pStartC, 15);
      }
    }

    sheetPartidas.getRange("M:M").setNumberFormat("#,##0.00;[Red]-#,##0.00");
    sheetPartidas.getRange("P:P").setNumberFormat("#,##0.00;[Red]-#,##0.00");
  }
}`,

  "ENCONTRAR FECHA REAL EN POR FACTURA DAVID_V2.txt": `function main(workbook: ExcelScript.Workbook) {
  const sheetFactura = workbook.getWorksheet("por factura");
  const sheetPartidas = workbook.getWorksheet("todas las partidas");

  if (!sheetFactura && !sheetPartidas) {
    console.log("No se encontraron las pestañas objetivo.");
    return;
  }

  const encabezadosClave = [
    "St", "Referencia", "Asignación", "Nº doc.", "Clase", 
    "Fecha doc.", "Doc.comp.", "Compens.", "IO", "II", 
    "Importe en ML", "ML", "Mon.", "Importe en MD", "Texto"
  ];

  function procesarEstructura(sheet: ExcelScript.Worksheet, esPorFactura: boolean) {
    if (!sheet) return null;
    const usedRange = sheet.getUsedRange();
    if (!usedRange) return null;

    const valores = usedRange.getValues() as string[][];
    const offsetRow = usedRange.getRowIndex();

    const filasDeEncabezados: { rangeIdx: number, absRow: number, indicesValores: { numDoc: number; clase: number; fechaDoc: number; docComp: number; } }[] = [];

    for (let r = 0; r < valores.length; r++) {
      const filaTextos = valores[r].map(c => c ? c.toString().trim() : "");
      let coincidencias = 0;
      for (const encabezado of encabezadosClave) {
        if (filaTextos.some(t => t === encabezado || t === "Compens/")) {
          coincidencias++;
        }
      }

      if (coincidencias >= 4) {
        let absRow = offsetRow + r;
        
        if (filasDeEncabezados.length === 0 || (absRow - filasDeEncabezados[filasDeEncabezados.length - 1].absRow > 4)) {
          const mapCols = {
            numDoc: filaTextos.findIndex(t => t === "Nº doc." || t === "N° doc." || t === "N doc."),
            clase: filaTextos.findIndex(t => t === "Clase"),
            fechaDoc: filaTextos.findIndex(t => t === "Fecha doc." || t === "Fecha doc/"),
            docComp: filaTextos.findIndex(t => t === "Doc.comp." || t === "Doc. comp.")
          };

          filasDeEncabezados.push({ rangeIdx: r, absRow: absRow, indicesValores: mapCols });
        }
      }
    }

    if (esPorFactura && filasDeEncabezados.length > 0) {
      const primerHeaderAbsIdx = filasDeEncabezados[0].absRow;
      const filaArribaAbsIndex = primerHeaderAbsIdx > 0 ? primerHeaderAbsIdx - 1 : 0;
      
      const columnaK = sheet.getRange("K:K");
      columnaK.insert(ExcelScript.InsertShiftDirection.right);

      const celdaFechaPago = sheet.getCell(filaArribaAbsIndex, 10);
      celdaFechaPago.setValue("Fecha de pago");
      celdaFechaPago.getFormat().getFont().setBold(true);

      const celdaComisiones = sheet.getCell(filaArribaAbsIndex, 11);
      celdaComisiones.setValue("Comisiones");
      celdaComisiones.getFormat().getFont().setBold(true);
    }

    return { sheet, valores, filasDeEncabezados, offsetRow };
  }

  const infoFactura = procesarEstructura(sheetFactura, true);
  const infoPartidas = procesarEstructura(sheetPartidas, false);

  if (infoFactura && infoPartidas && infoFactura.filasDeEncabezados.length > 0 && infoPartidas.filasDeEncabezados.length > 0) {
    const colMapPartidas = infoPartidas.filasDeEncabezados[0].indicesValores;
    const colMapFactura = infoFactura.filasDeEncabezados[0].indicesValores;

    if (colMapPartidas.clase !== -1 && colMapPartidas.numDoc !== -1 && colMapFactura.docComp !== -1) {
      
      for (let p = 0; p < infoPartidas.valores.length; p++) {
        const filaP = infoPartidas.valores[p];
        const clasePartida = filaP[colMapPartidas.clase] ? filaP[colMapPartidas.clase].toString().trim() : "";
        
        if (clasePartida === "DZ") {
          const numDocDZ = filaP[colMapPartidas.numDoc] ? filaP[colMapPartidas.numDoc].toString().trim() : "";
          const fechaDocDZ = filaP[colMapPartidas.fechaDoc] ? filaP[colMapPartidas.fechaDoc].toString().trim() : "";

          if (!numDocDZ) continue;

          let docCompsBuscados: string[] = [];

          for (let f = 0; f < infoFactura.valores.length; f++) {
            const f_numDoc = infoFactura.valores[f][colMapFactura.numDoc] ? infoFactura.valores[f][colMapFactura.numDoc].toString().trim() : "";
            const f_clase = infoFactura.valores[f][colMapFactura.clase] ? infoFactura.valores[f][colMapFactura.clase].toString().trim() : "";
            
            const isMatch = (f_numDoc === numDocDZ) || (f_numDoc !== "" && numDocDZ !== "" && Number(f_numDoc.replace(/,/g, "")) === Number(numDocDZ.replace(/,/g, "")));

            if (isMatch && f_clase === "DZ") {
              const docCompActual = infoFactura.valores[f][colMapFactura.docComp] ? infoFactura.valores[f][colMapFactura.docComp].toString().trim() : "";
              
              if (docCompActual.startsWith("1000")) {
                const absRowF = f + infoFactura.offsetRow;
                sheetFactura.getCell(absRowF, 10).setValue(docCompActual);
              } else if (docCompActual !== "") {
                docCompsBuscados.push(docCompActual);
              }
            }
          }

          for (const docComp of docCompsBuscados) {
            for (let f2 = 0; f2 < infoFactura.valores.length; f2++) {
              const f2_clase = infoFactura.valores[f2][colMapFactura.clase] ? infoFactura.valores[f2][colMapFactura.clase].toString().trim() : "";
              const f2_docComp = infoFactura.valores[f2][colMapFactura.docComp] ? infoFactura.valores[f2][colMapFactura.docComp].toString().trim() : "";
              
              if (f2_clase === "RV" && f2_docComp === docComp) {
                const absRowF2 = f2 + infoFactura.offsetRow;
                sheetFactura.getCell(absRowF2, 10).setValue(fechaDocDZ);
              }
            }
          }
        }
      }
    }
  }

  if (sheetFactura) {
    const finalUsedRange = sheetFactura.getUsedRange();
    if (finalUsedRange) {
      const finalVals = finalUsedRange.getValues() as string[][];
      const startR = finalUsedRange.getRowIndex();
      const startC = finalUsedRange.getColumnIndex();

      const relJ = 9 - startC;
      const relK = 10 - startC;

      for (let r = 0; r < finalVals.length; r++) {
        const absRow = startR + r;

        const convertirFecha = (relIdx: number, absCol: number) => {
          if (relIdx >= 0 && relIdx < finalVals[r].length) {
            const val = finalVals[r][relIdx] ? finalVals[r][relIdx].toString().trim() : "";
            const match = val.match(/^(\d{2})\.(\d{2})\.(\d{2,4})$/);
            if (match) {
              let dia = match[1];
              let mes = match[2];
              let anio = match[3];
              if (anio.length === 2) anio = "20" + anio;
              
              const isoDate = \`\${anio}-\\mes-\\dia\`;
              const celdaFecha = sheetFactura.getCell(absRow, absCol);
              celdaFecha.setValue(isoDate);
              celdaFecha.setNumberFormat("dd-mmm-yy");
            }
          }
        };

        const convertirANumero = (relIdx: number, absCol: number) => {
          if (relIdx >= 0 && relIdx < finalVals[r].length) {
            let strVal = finalVals[r][relIdx] ? finalVals[r][relIdx].toString().trim() : "";
            if (strVal !== "") {
              if (strVal.endsWith("-")) {
                strVal = "-" + strVal.substring(0, strVal.length - 1);
              }
              let saneado = strVal.replace(/,/g, ""); 
              let num = Number(saneado);
              if (!isNaN(num)) {
                sheetFactura.getCell(absRow, absCol).setValue(num);
              }
            }
          }
        };

        convertirFecha(relJ, 9);
        convertirFecha(relK, 10);
        convertirANumero(13 - startC, 13);
        convertirANumero(16 - startC, 16);
      }
    }

    sheetFactura.getRange("N:N").setNumberFormat("#,##0.00;[Red]-#,##0.00");
    sheetFactura.getRange("Q:Q").setNumberFormat("#,##0.00;[Red]-#,##0.00");
  }

  if (sheetPartidas) {
    const partidasUsedRange = sheetPartidas.getUsedRange();
    if (partidasUsedRange) {
      const pVals = partidasUsedRange.getValues() as string[][];
      const pStartR = partidasUsedRange.getRowIndex();
      const pStartC = partidasUsedRange.getColumnIndex();

      for (let r = 0; r < pVals.length; r++) {
        const absRow = pStartR + r;

        const convertirANumeroPartidas = (relIdx: number, absCol: number) => {
          if (relIdx >= 0 && relIdx < pVals[r].length) {
            let strVal = pVals[r][relIdx] ? pVals[r][relIdx].toString().trim() : "";
            if (strVal !== "") {
              if (strVal.endsWith("-")) {
                strVal = "-" + strVal.substring(0, strVal.length - 1);
              }
              let saneado = strVal.replace(/,/g, ""); 
              let num = Number(saneado);
              if (!isNaN(num)) {
                sheetPartidas.getCell(absRow, absCol).setValue(num);
              }
            }
          }
        };

        convertirANumeroPartidas(12 - pStartC, 12);
        convertirANumeroPartidas(15 - pStartC, 15);
      }
    }

    sheetPartidas.getRange("M:M").setNumberFormat("#,##0.00;[Red]-#,##0.00");
    sheetPartidas.getRange("P:P").setNumberFormat("#,##0.00;[Red]-#,##0.00");
  }
}`,

  "TYPE SCRIPT_COBRANZA_COLORES_PARA MI.txt": `/**
 * Script para conciliar documentos e IVA entre la hoja activa y "IVA COBRADO"
 * Diseñado para Office Scripts en Excel en la Web o Excel de Escritorio moderno.
 */
function main(workbook: ExcelScript.Workbook) {
    const currentSheet = workbook.getActiveWorksheet();
    const targetSheet = workbook.getWorksheet("IVA COBRADO");

    if (!targetSheet) {
        throw new Error("Fallo de ejecución: No existe una pestaña llamada 'IVA COBRADO'.");
    }
    if (currentSheet.getName() === targetSheet.getName()) {
        throw new Error("Fallo de ejecución: No puedes ejecutar este script estando posicionado en la pestaña 'IVA COBRADO'.");
    }

    const currentRange = currentSheet.getUsedRange();
    const targetRange = targetSheet.getUsedRange();

    if (!currentRange || !targetRange) {
        console.log("Una de las hojas está vacía.");
        return;
    }

    const currentValues = currentRange.getValues() as string[][];
    const targetValues = targetRange.getValues() as string[][];

    let headerRowCurrent = -1;
    let colNDocCurrent = -1;
    let colIvaCurrent = -1;

    for (let i = 0; i < Math.min(30, currentValues.length); i++) {
        const row = currentValues[i].map(h => h?.toString().trim() || "");
        const cDoc = row.indexOf("Nº doc.");
        const cIva = row.indexOf("IVA 16%");

        if (cDoc !== -1 && cIva !== -1) {
            headerRowCurrent = i;
            colNDocCurrent = cDoc;
            colIvaCurrent = cIva;
            break;
        }
    }

    let headerRowTarget = -1;
    let colAsignacionTarget = -1;
    let colImporteTarget = -1;
    let colBcoTarget = -1;

    for (let i = 0; i < Math.min(30, targetValues.length); i++) {
        const row = targetValues[i].map(h => h?.toString().trim() || "");
        const cAsig = row.indexOf("Asignación");
        const cImp = row.indexOf("Importe en ML");
        const cBco = row.indexOf("Bco.prp.");

        if (cAsig !== -1 && cImp !== -1 && cBco !== -1) {
            headerRowTarget = i;
            colAsignacionTarget = cAsig;
            colImporteTarget = cImp;
            colBcoTarget = cBco;
            break;
        }
    }

    if (headerRowCurrent === -1 || colNDocCurrent === -1 || colIvaCurrent === -1) {
        throw new Error("Error en Hoja Actual: No se encontraron los encabezados 'Nº doc.' o 'IVA 16%' en las primeras 30 filas.");
    }
    if (headerRowTarget === -1 || colAsignacionTarget === -1 || colImporteTarget === -1 || colBcoTarget === -1) {
        throw new Error("Error en IVA COBRADO: No se encontraron los encabezados 'Asignación', 'Importe en ML' o 'Bco.prp.' en las primeras 30 filas.");
    }

    const colors = [
        "#ED7D31",
        "#FFFF00",
        "#92D050",
        "#00B050",
        "#00B0F0",
        "#7030A0",
        "#F4B084"
    ];

    const bcoValues = targetValues
        .slice(headerRowTarget + 1)
        .map(row => row[colBcoTarget]?.toString().trim() || "")
        .filter(val => val !== "");

    const uniqueSheets = Array.from(new Set(bcoValues));
    let currentSheetIndex = uniqueSheets.indexOf(currentSheet.getName());

    if (currentSheetIndex === -1) {
        currentSheetIndex = uniqueSheets.length;
    }

    if (currentSheetIndex >= colors.length) {
        throw new Error(\`Fallo de escalabilidad: Estás intentando procesar la hoja número \${currentSheetIndex + 1}, pero tu regla estricta solo permite 7 colores.\`);
    }

    const sheetColor = colors[currentSheetIndex];

    for (let i = headerRowCurrent + 1; i < currentValues.length; i++) {
        const nDoc = currentValues[i][colNDocCurrent]?.toString().trim();
        const ivaCurrentRaw = parseFloat(currentValues[i][colIvaCurrent]?.toString());

        if (!nDoc || isNaN(ivaCurrentRaw)) continue;

        for (let j = headerRowTarget + 1; j < targetValues.length; j++) {
            const asignacion = targetValues[j][colAsignacionTarget]?.toString().trim();
            const importeTargetRaw = parseFloat(targetValues[j][colImporteTarget]?.toString());

            if (!asignacion || isNaN(importeTargetRaw)) continue;

            if (asignacion.includes(nDoc)) {
                const absIva = Math.round(Math.abs(ivaCurrentRaw) * 100) / 100;
                const absImporte = Math.round(Math.abs(importeTargetRaw) * 100) / 100;
                const diferencia = Math.abs(absIva - absImporte);

                if (diferencia <= 1.00) {
                    targetSheet.getCell(j, colBcoTarget).setValue(currentSheet.getName());
                    currentSheet.getCell(i, colIvaCurrent).getFormat().getFill().setColor(sheetColor);
                    targetSheet.getCell(j, colImporteTarget).getFormat().getFill().setColor(sheetColor);
                    break;
                }
            }
        }
    }
}`,

  "TYPE SCRIPT_COBRANZA_PARA MI.txt": `/**
 * Script de Integración de Pagos SAP (Compensaciones)
 * Arquitectura basada en agrupación de Documentos de Compensación (Doc.comp.)
 * 100% Dinámico: Busca las columnas en la hoja 210 por sus encabezados.
 */
interface PartidaSAP {
    noDoc: string;
    importe: number;
    referencia: string;
    cliente: string;
}

function main(workbook: ExcelScript.Workbook) {
    let wsCobranza = workbook.getActiveWorksheet();
    let wsPartidas = workbook.getWorksheet("Sheet1"); 

    if (!wsPartidas) {
        console.log("ERROR CRÍTICO: No se encontró la hoja de partidas llamada 'Sheet1'.");
        return;
    }

    if (wsCobranza.getName() === wsPartidas.getName()) {
        console.log("ERROR CRÍTICO: Estás posicionado en 'Sheet1'. Ve a la hoja de Cobranza antes de ejecutar.");
        return;
    }

    const mapaDocACompensacion = new Map<string, string>();
    const agrupacionCompensaciones = new Map<string, PartidaSAP[]>();

    const rangoPartidas = wsPartidas.getUsedRange();
    if (!rangoPartidas) {
        console.log("ADVERTENCIA: La hoja 'Sheet1' está vacía.");
        return;
    }

    const valoresPartidas = rangoPartidas.getValues();
    let clienteActual = "";
    
    for (let i = 1; i < valoresPartidas.length; i++) {
        const valColA = String(valoresPartidas[i][0] || "").trim();
        if (valColA.toUpperCase().startsWith("NOMBRE")) {
            if (valColA.toUpperCase() === "NOMBRE") {
                let nameVal = String(valoresPartidas[i][1] || "").trim();
                if (!nameVal) nameVal = String(valoresPartidas[i][2] || "").trim();
                if (!nameVal) nameVal = String(valoresPartidas[i][3] || "").trim();
                clienteActual = nameVal;
            } else {
                clienteActual = valColA.replace(/^NOMBRE\\s*/i, "").trim();
            }
        }

        if (valoresPartidas[i].length <= 10) continue;
        const rawNoDoc = valoresPartidas[i][4];
        const rawReferencia = valoresPartidas[i][5];
        const rawDocComp = valoresPartidas[i][6];
        const rawImporte = valoresPartidas[i][10];

        const noDoc = limpiarIDRiguroso(rawNoDoc);
        const docComp = limpiarIDRiguroso(rawDocComp);

        const esNumeroSAP = /^\\d+$/; 
        
        if (esNumeroSAP.test(noDoc) && esNumeroSAP.test(docComp)) {
            const importeDesinfectado = String(rawImporte).replace(/,/g, '').replace(/ /g, '');
            const importe = Number(importeDesinfectado);
            
            if (!isNaN(importe)) {
                mapaDocACompensacion.set(noDoc, docComp);
                if (!agrupacionCompensaciones.has(docComp)) {
                    agrupacionCompensaciones.set(docComp, []);
                }
                agrupacionCompensaciones.get(docComp).push({
                    noDoc: noDoc,
                    importe: importe,
                    referencia: String(rawReferencia || "").trim(),
                    cliente: clienteActual
                });
            }
        }
    }

    const rangoCobranza = wsCobranza.getUsedRange();
    if (!rangoCobranza) {
        console.log("ADVERTENCIA: La hoja activa está vacía.");
        return;
    }

    const encabezadosCobranza = wsCobranza.getRange("A10:AZ10").getValues()[0];
    let idxNoDoc = -1, idxImporte = -1, idxImporteMD = -1, idxFactura = -1, idxCliente = -1;
    let idxTpCambio = -1, idxImporteML3 = -1, idxCobranza16 = -1;
    let idxIva16 = -1, idxTotalDep = -1;
    let idxIntereses = -1, idxTraspasos = -1, idxOtros = -1, idxDepOtros = -1;
    let idxDepCancel = -1, idxOtrosIetu = -1, idxRetIva = -1, idxCob0 = -1, idxCobExp = -1;
    
    for (let c = 0; c < encabezadosCobranza.length; c++) {
        const titulo = String(encabezadosCobranza[c]).trim().toUpperCase();
        if (titulo.includes("Nº DOC")) idxNoDoc = c;
        if (titulo === "IMPORTE") idxImporte = c;
        if (titulo === "IMPORTE EN MD") idxImporteMD = c;
        if (titulo === "FACTURA" || titulo === "REFERENCIA") idxFactura = c;
        if (titulo === "CLIENTE") idxCliente = c;
        if (titulo.includes("TP.CAMB.EF")) idxTpCambio = c;
        if (titulo.includes("IMPORTE EN ML3")) idxImporteML3 = c;
        if (titulo === "COBRANZA 16%") idxCobranza16 = c;
        if (titulo === "IVA 16%") idxIva16 = c;
        if (titulo === "TOTAL DEPOSITOS") idxTotalDep = c;
        if (titulo.includes("INTERESES")) idxIntereses = c;
        if (titulo.includes("TRASPASOS BANCARIOS")) idxTraspasos = c;
        if (titulo === "OTROS") idxOtros = c; 
        if (titulo.includes("DEP DE OTROS MESES")) idxDepOtros = c;
        if (titulo.includes("DEPOSITOS CANCEL")) idxDepCancel = c;
        if (titulo.includes("OTROS ING PARA IETU")) idxOtrosIetu = c;
        if (titulo.includes("RET DE IVA")) idxRetIva = c;
        if (titulo.includes("COBRANZA 0%")) idxCob0 = c;
        if (titulo.includes("COBRANZA EXP")) idxCobExp = c;
    }

    if (idxNoDoc === -1 || idxImporte === -1 || idxImporteMD === -1 || idxFactura === -1 || idxCliente === -1 || 
        idxTpCambio === -1 || idxImporteML3 === -1 || idxCobranza16 === -1 || idxIva16 === -1 || idxTotalDep === -1) {
        console.log("ERROR CRÍTICO: Faltan encabezados básicos.");
        return;
    }

    const letraImporte = obtenerLetraColumna(idxImporte);
    const letraImporteMD = obtenerLetraColumna(idxImporteMD);
    const letraTpCambio = obtenerLetraColumna(idxTpCambio);
    const letraImporteML3 = obtenerLetraColumna(idxImporteML3);
    const letraCobranza16 = obtenerLetraColumna(idxCobranza16);
    const letraIva16 = obtenerLetraColumna(idxIva16);
    
    const columnasSuma = [
        idxIntereses !== -1 ? obtenerLetraColumna(idxIntereses) : "",
        idxTraspasos !== -1 ? obtenerLetraColumna(idxTraspasos) : "",
        idxOtros !== -1 ? obtenerLetraColumna(idxOtros) : "",
        idxDepOtros !== -1 ? obtenerLetraColumna(idxDepOtros) : "",
        idxDepCancel !== -1 ? obtenerLetraColumna(idxDepCancel) : "",
        idxOtrosIetu !== -1 ? obtenerLetraColumna(idxOtrosIetu) : "",
        idxRetIva !== -1 ? obtenerLetraColumna(idxRetIva) : "",
        letraCobranza16,
        letraIva16,
        idxCob0 !== -1 ? obtenerLetraColumna(idxCob0) : "",
        idxCobExp !== -1 ? obtenerLetraColumna(idxCobExp) : ""
    ].filter(letra => letra !== "");

    const rowCountInicial = rangoCobranza.getRowCount();
    let limiteInferiorTabla = 10; 
    let vaciasConsecutivas = 0;
    
    for(let i = 10; i < rowCountInicial; i++) {
        let celdaVal = String(wsCobranza.getCell(i, idxNoDoc).getValue() || "").trim();
        if(celdaVal === "") {
            vaciasConsecutivas++;
            if(vaciasConsecutivas >= 3) {
                limiteInferiorTabla = i - 3;
                break;
            }
        } else {
            vaciasConsecutivas = 0;
            limiteInferiorTabla = i;
        }
    }
    
    if(limiteInferiorTabla < 10) limiteInferiorTabla = rowCountInicial - 1;
    
    let filasEliminadas = 0;
    for(let i = limiteInferiorTabla; i >= 10; i--) {
        let celdaVal = String(wsCobranza.getCell(i, idxNoDoc).getValue() || "").trim();
        if(celdaVal === "") {
            wsCobranza.getRangeByIndexes(i, 0, 1, 1).getEntireRow().delete(ExcelScript.DeleteShiftDirection.up);
            filasEliminadas++;
        }
    }

    if (filasEliminadas > 0) {
        limiteInferiorTabla -= filasEliminadas;
    }

    let contadorModificados = 0;
    for (let i = limiteInferiorTabla; i >= 10; i--) {
        const celdaNoDoc = wsCobranza.getCell(i, idxNoDoc);
        const targetNoDoc = limpiarIDRiguroso(celdaNoDoc.getValue());

        if (targetNoDoc && /^\\d+$/.test(targetNoDoc)) {
            const docCompAsociado = mapaDocACompensacion.get(targetNoDoc);
            
            if (docCompAsociado) {
                const partidasDelGrupo = agrupacionCompensaciones.get(docCompAsociado);
                let contrapartidas = partidasDelGrupo.filter(p => p.noDoc !== targetNoDoc);
                
                const valImporteMD = wsCobranza.getCell(i, idxImporteMD).getValue();
                const importeMDNum = parseFloat(String(valImporteMD)) || 0;
                
                let sumaSAP = 0;
                for (let p of contrapartidas) {
                    sumaSAP += p.importe;
                }
                
                let diferencia = Math.round((importeMDNum - sumaSAP) * 100) / 100;
                let difAbsoluta = Math.abs(diferencia);
                
                if (difAbsoluta > 0 && difAbsoluta <= 10) {
                    contrapartidas.push({
                        noDoc: "DIFERENCIA", 
                        importe: diferencia, 
                        referencia: "Gastos no deducible", 
                        cliente: contrapartidas.length > 0 ? contrapartidas[0].cliente : "" 
                    });
                }

                const rowExcelPadre = i + 1;
                
                let valTpCambio = wsCobranza.getCell(i, idxTpCambio).getValue();
                let tpCambioNum = parseFloat(String(valTpCambio));
                if (isNaN(tpCambioNum) || String(valTpCambio).trim() === "") {
                    tpCambioNum = 1;
                }
                const isTpCambioNot1 = (tpCambioNum !== 1);
                
                if (contrapartidas.length === 1) {
                    const celdaDestino = wsCobranza.getCell(i, idxImporte);
                    celdaDestino.setValue(contrapartidas[0].importe);
                    celdaDestino.setNumberFormat("#,##0.00"); 

                    wsCobranza.getCell(i, idxFactura).setValue(contrapartidas[0].referencia);
                    wsCobranza.getCell(i, idxCliente).setValue(contrapartidas[0].cliente);
                    
                    const celdaCuadre = wsCobranza.getCell(i, 0);
                    celdaCuadre.setFormulaLocal(\`=+\${letraImporteMD}\${rowExcelPadre}-\${letraImporte}\${rowExcelPadre}\`);
                    celdaCuadre.setNumberFormat("#,##0.00;[Red]-#,##0.00");
                    
                    const celdaCobranza16 = wsCobranza.getCell(i, idxCobranza16);
                    if (isTpCambioNot1) {
                        celdaCobranza16.setFormulaLocal(\`=+\${letraImporteML3}\${rowExcelPadre}/1.16\`);
                    } else {
                        celdaCobranza16.setFormulaLocal(\`=+\${letraImporte}\${rowExcelPadre}/1.16\`);
                    }
                    celdaCobranza16.setNumberFormat("#,##0.00");
                    
                    const celdaIva16 = wsCobranza.getCell(i, idxIva16);
                    celdaIva16.setFormulaLocal(\`=+\${letraCobranza16}\${rowExcelPadre}*0.16\`);
                    celdaIva16.setNumberFormat("#,##0.00");

                    const celdaTotalDep = wsCobranza.getCell(i, idxTotalDep);
                    const formulaSuma = "=SUM(" + columnasSuma.map(col => col + rowExcelPadre).join(",") + ")";
                    celdaTotalDep.setFormula(formulaSuma); 
                    celdaTotalDep.setNumberFormat("#,##0.00");

                    contadorModificados++;
                    
                } else if (contrapartidas.length > 1) {
                    wsCobranza.getCell(i, idxImporte).setValue("");
                    wsCobranza.getCell(i, idxFactura).setValue(""); 
                    wsCobranza.getCell(i, idxCliente).setValue(""); 
                    
                    if (isTpCambioNot1) {
                        wsCobranza.getCell(i, idxImporteML3).setValue("");
                    }
                    
                    const celdaCuadrePadre = wsCobranza.getCell(i, 0);
                    celdaCuadrePadre.setFormulaLocal(\`=+\${letraImporteMD}\${rowExcelPadre}-\${letraImporte}\${rowExcelPadre}\`);
                    celdaCuadrePadre.setNumberFormat("#,##0.00;[Red]-#,##0.00");

                    for (let j = 0; j < contrapartidas.length; j++) {
                        const indexInsercion = i + 1 + j;
                        const rowExcelHija = indexInsercion + 1;
                        
                        const filaNueva = wsCobranza.getRangeByIndexes(indexInsercion, 0, 1, 1).getEntireRow();
                        filaNueva.insert(ExcelScript.InsertShiftDirection.down);
                        
                        const celdaImporteNueva = wsCobranza.getCell(indexInsercion, idxImporte);
                        celdaImporteNueva.setValue(contrapartidas[j].importe);
                        celdaImporteNueva.setNumberFormat("#,##0.00");

                        wsCobranza.getCell(indexInsercion, idxFactura).setValue(contrapartidas[j].referencia);
                        wsCobranza.getCell(indexInsercion, idxCliente).setValue(contrapartidas[j].cliente);
                        
                        const celdaCuadreHija = wsCobranza.getCell(indexInsercion, 0);
                        celdaCuadreHija.setFormulaLocal(\`=+\${letraImporteMD}\${rowExcelHija}-\${letraImporte}\${rowExcelHija}\`);
                        celdaCuadreHija.setNumberFormat("#,##0.00;[Red]-#,##0.00");
                        
                        const celdaCob16Nueva = wsCobranza.getCell(indexInsercion, idxCobranza16);
                        
                        if (isTpCambioNot1) {
                            const celdaML3Nueva = wsCobranza.getCell(indexInsercion, idxImporteML3);
                            celdaML3Nueva.setFormulaLocal(\`=+\${letraImporte}\${rowExcelHija}*\${letraTpCambio}\${rowExcelPadre}\`);
                            celdaML3Nueva.setNumberFormat("#,##0.00");
                            celdaCob16Nueva.setFormulaLocal(\`=+\${letraImporteML3}\${rowExcelHija}/1.16\`);
                        } else {
                            celdaCob16Nueva.setFormulaLocal(\`=+\${letraImporte}\${rowExcelHija}/1.16\`);
                        }
                        
                        celdaCob16Nueva.setNumberFormat("#,##0.00");
                        
                        const celdaIva16Nueva = wsCobranza.getCell(indexInsercion, idxIva16);
                        celdaIva16Nueva.setFormulaLocal(\`=+\${letraCobranza16}\${rowExcelHija}*0.16\`);
                        celdaIva16Nueva.setNumberFormat("#,##0.00");

                        const celdaTotalDepNueva = wsCobranza.getCell(indexInsercion, idxTotalDep);
                        const formulaSumaHija = "=SUM(" + columnasSuma.map(col => col + rowExcelHija).join(",") + ")";
                        celdaTotalDepNueva.setFormula(formulaSumaHija);
                        celdaTotalDepNueva.setNumberFormat("#,##0.00");

                        contadorModificados++;
                    }
                }
            }
        }
    }

    const rangoPost = wsCobranza.getUsedRange();
    const idxUltimaFila = rangoPost.getLastCell().getRowIndex();
    const valoresPost = wsCobranza.getRange(\`A1:AZ\${idxUltimaFila + 1}\`).getValues();
    
    let ultimoTpCambioNum = 1;
    let ultimoTpCambioRaw = "";

    for (let i = 10; i <= idxUltimaFila; i++) {
        const rowExcel = i + 1;
        
        let colNoDocVal = String(valoresPost[i][idxNoDoc] || "").trim();
        let tieneNoDoc = (colNoDocVal !== "");
        let valImporteRaw = valoresPost[i][idxImporte];
        let valImporteStr = String(valImporteRaw || "").trim();
        let tieneImporte = (valImporteStr !== "");
        let facturaVal = String(valoresPost[i][idxFactura] || "").trim();
        let actualTpCambioRaw = valoresPost[i][idxTpCambio];

        if (tieneNoDoc) {
             let parsed = parseFloat(String(actualTpCambioRaw));
             if (!isNaN(parsed) && parsed !== 1 && String(actualTpCambioRaw).trim() !== "") {
                 ultimoTpCambioNum = parsed;
                 ultimoTpCambioRaw = actualTpCambioRaw;
             } else {
                 ultimoTpCambioNum = 1;
                 ultimoTpCambioRaw = "";
             }
        } else {
             if (tieneImporte && ultimoTpCambioNum !== 1) {
                 wsCobranza.getCell(i, idxTpCambio).setValue(ultimoTpCambioRaw);
                 actualTpCambioRaw = ultimoTpCambioRaw; 
             }
        }

        let parsedReal = parseFloat(String(actualTpCambioRaw));
        if (!isNaN(parsedReal) && parsedReal !== 1 && String(actualTpCambioRaw).trim() !== "") {
             if (tieneImporte && idxImporteML3 !== -1) {
                 const celdaML3 = wsCobranza.getCell(i, idxImporteML3);
                 celdaML3.setFormulaLocal(\`=+\${letraImporte}\${rowExcel}*\${letraTpCambio}\${rowExcel}\`);
                 celdaML3.setNumberFormat("#,##0.00");
             }
        }

        if (facturaVal === "Gastos no deducible") {
             wsCobranza.getCell(i, idxCobranza16).setValue("");
             
             if (idxOtros !== -1 && tieneImporte) {
                 wsCobranza.getCell(i, idxOtros).setValue(valImporteRaw);
                 wsCobranza.getCell(i, idxOtros).setNumberFormat("#,##0.00");
             }
        }
    }
}`,

  "TYPE SCRIPT_NOMBRE DEL CLIENTE PARA DAVID.txt": `/**
 * Script Optimizado de Extracción de Clientes (Single Pass)
 * Detecta dinámicamente las columnas clave y utiliza una regla de umbral (>= 3 columnas llenas).
 * INCLUYE CORRECCIÓN ALV: Ignora automáticamente las filas de encabezados repetidos de SAP.
 */
function main(workbook: ExcelScript.Workbook) {
    let wsActiva = workbook.getActiveWorksheet();
    let rangoTotal = wsActiva.getUsedRange();
    
    if (!rangoTotal) {
        console.log("ADVERTENCIA: La hoja activa está vacía.");
        return;
    }

    const valores = rangoTotal.getValues();
    const fragmentosBuscados = [
        "REFERENCIA", "ASIGNACI", "Nº DOC", "CLASE", 
        "FECHA DOC", "DOC.COMP", "COMPENS"
    ];
    
    let indicesColumnasValidacion: number[] = [];
    let filaDeEncabezados = -1;

    for (let i = 0; i < Math.min(valores.length, 50); i++) {
        let coincidenciasEnFila = 0;
        let indicesTemporales: number[] = [];

        for (let c = 0; c < valores[i].length; c++) {
            let textoCelda = String(valores[i][c]).toUpperCase().replace(/\\s+/g, " ").trim();
            let coincide = fragmentosBuscados.some(frag => textoCelda.includes(frag));
            
            if (coincide) {
                coincidenciasEnFila++;
                indicesTemporales.push(c);
            }
        }

        if (coincidenciasEnFila >= 3) {
            filaDeEncabezados = i;
            indicesColumnasValidacion = indicesTemporales;
            break;
        }
    }

    if (indicesColumnasValidacion.length === 0) {
        console.log("ERROR CRÍTICO: No se encontraron los encabezados objetivo.");
        return;
    }

    let clienteActual = "";
    let contadorModificados = 0;

    for (let i = 0; i < valores.length; i++) {
        const valColA = String(valores[i][0] || "").trim().toUpperCase();
        
        if (valColA.startsWith("NOMBRE")) {
            if (valColA === "NOMBRE") {
                let nameVal = String(valores[i][1] || "").trim();
                if (!nameVal) nameVal = String(valores[i][2] || "").trim();
                if (!nameVal) nameVal = String(valores[i][3] || "").trim();
                clienteActual = nameVal;
            } else {
                clienteActual = valColA.replace(/^NOMBRE\\s*/i, "").trim();
            }
            continue; 
        }

        let columnasConDatos = 0;
        let esFilaRepetidaEncabezado = false;
        
        for (let idx of indicesColumnasValidacion) {
            if (idx < valores[i].length) {
                let contenido = String(valores[i][idx] || "").toUpperCase().replace(/\\s+/g, " ").trim();
                
                if (contenido !== "") {
                    columnasConDatos++;
                    if (fragmentosBuscados.some(frag => contenido.includes(frag))) {
                        esFilaRepetidaEncabezado = true;
                    }
                }
            }
        }

        if (columnasConDatos >= 3 && !esFilaRepetidaEncabezado && i > filaDeEncabezados) {
            wsActiva.getCell(i, 0).setValue(clienteActual);
            contadorModificados++;
        }
    }

    console.log(\`✅ Proceso completado. Se insertó el cliente en \${contadorModificados} filas.\`);
}`
};
