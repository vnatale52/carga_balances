// backend/server.js (Versión con Formato Avanzado y Subtotales)

import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import xlsx from 'xlsx';
import readline from 'readline';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json()); 
app.use(express.static(path.join(__dirname, '../frontend')));

// --- FUNCIONES DE PROCESAMIENTO OPTIMIZADAS (Sin cambios respecto a la versión anterior) ---

async function procesarBalhist(filePath, filtros) {
    const isAllEntities = filtros.entidad.includes("0");
    const selectedEntitiesSet = isAllEntities ? null : new Set(filtros.entidad.map(Number));
    const fechaDesde = filtros.balhistDesde;
    const fechaHasta = filtros.balhistHasta;
    if (fechaDesde > fechaHasta) return [];
    const resultados = [];
    const fileStream = fs.createReadStream(filePath, { encoding: 'latin1' });
    const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity });
    for await (const linea of rl) {
        if (linea.trim() === '') continue;
        const [numEntidadStr, fechaBceStr, numCuentaStr, saldoStr] = linea.split('\t');
        if (!numEntidadStr || !fechaBceStr || !numCuentaStr || saldoStr === undefined) continue;
        const entidadActual = parseInt(numEntidadStr.replace(/"/g, ''), 10);
        const anio = fechaBceStr.replace(/"/g, '').substring(0, 4);
        const mes = fechaBceStr.replace(/"/g, '').substring(4, 6);
        const fechaComparable = `${anio}-${mes}`;
        const matchesDateRange = (fechaComparable >= fechaDesde && fechaComparable <= fechaHasta);
        const matchesEntity = isAllEntities || selectedEntitiesSet.has(entidadActual);
        if (matchesDateRange && matchesEntity) {
            resultados.push({
                num_entidad: entidadActual,
                fecha_bce: `${mes}-${anio}`,
                num_cuenta: parseInt(numCuentaStr.replace(/"/g, ''), 10),
                saldo: parseInt(saldoStr.trim(), 10)
            });
        }
    }
    return resultados;
}

async function procesarCuentas(filePath) {
    const cuentas = [];
    const fileStream = fs.createReadStream(filePath, { encoding: 'latin1' });
    const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity });
    for await (const linea of rl) {
        if (linea.trim() === '') continue;
        const [numCuenta, descripcion] = linea.split('\t');
        if (!numCuenta || !descripcion) continue;
        cuentas.push({
            num_cuenta: parseInt(numCuenta.replace(/"/g, ''), 10),
            descripcion_cuenta: descripcion.replace(/"/g, '').trim(),
        });
    }
    return cuentas;
}

async function procesarNomina(filePath) {
    const nomina = [];
    const fileStream = fs.createReadStream(filePath, { encoding: 'latin1' });
    const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity });
    for await (const linea of rl) {
        if (linea.trim() === '') continue;
        const [numEntidad, nombreEntidad, nombreCorto] = linea.split('\t');
        if (!numEntidad || !nombreEntidad) continue;
        nomina.push({
            num_entidad: parseInt(numEntidad.replace(/"/g, ''), 10),
            nombre_entidad: nombreEntidad.replace(/"/g, '').trim(),
            nombre_corto: (nombreCorto || '').replace(/"/g, '').trim()
        });
    }
    return nomina;
}

function procesarIndices(filePath) {
    try {
        const buffer = fs.readFileSync(filePath);
        const workbook = xlsx.read(buffer, { type: 'buffer', cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
        const indicesMap = new Map();
        for (const row of jsonData) {
            if (!row || row.length < 2) continue;
            const fechaValue = row[0];
            const indiceValue = row[1];
            if (!indiceValue || !(fechaValue instanceof Date) || isNaN(fechaValue)) continue;
            const anio = fechaValue.getFullYear();
            const mes = ('0' + (fechaValue.getMonth() + 1)).slice(-2);
            const fechaFormatoIndice = `${mes}-${anio}`;
            const indiceStr = String(indiceValue).replace(',', '.');
            indicesMap.set(fechaFormatoIndice, parseFloat(indiceStr));
        }
        return indicesMap;
    } catch (error) {
        console.error("Error crítico al leer o procesar el archivo Excel:", error);
        return new Map();
    }
}

function getMonthsInRange(start, end) {
    const startDate = new Date(`${start}-01T00:00:00`);
    const endDate = new Date(`${end}-01T00:00:00`);
    let currentDate = startDate;
    const months = [];
    while (currentDate <= endDate) {
        const month = ('0' + (currentDate.getMonth() + 1)).slice(-2);
        const year = currentDate.getFullYear();
        months.push(`${month}-${year}`);
        currentDate.setMonth(currentDate.getMonth() + 1);
    }
    return months;
}

// --- FUNCIÓN CENTRAL REFACTORIZADA PARA ORDENAMIENTO Y SUBTOTALES ---
function prepareDataForSheet(balancesDeEstaEntidad, cuentasMap, nominaMap, allMonths, indicesMap) {
    if (!balancesDeEstaEntidad || balancesDeEstaEntidad.length === 0) return [];
    
    const num_entidad = balancesDeEstaEntidad[0].num_entidad;
    const infoEntidad = nominaMap.get(num_entidad) || { nombre_entidad: 'Desconocido', num_entidad };

    const pivotedData = {};
    for (const balance of balancesDeEstaEntidad) {
        if (!pivotedData[balance.num_cuenta]) {
            const desc = (cuentasMap.get(balance.num_cuenta) || {}).descripcion_cuenta || 'No encontrada';
            pivotedData[balance.num_cuenta] = { desc, saldos: {} };
        }
        pivotedData[balance.num_cuenta].saldos[balance.fecha_bce] = balance.saldo;
    }

    const newHeaders = ['Entidad', 'Nombre Entidad', 'Cuenta', 'Descripción Cuenta'];
    const numericHeaders = [];
    allMonths.forEach(month => {
        const headersForMonth = [
            `${month} Saldo en moneda constante`,
            `${month} Saldo Histórico solo del mes`,
            `${month} Saldo Histórico acumulado al mes`,
            `${month} AXI mensual solo del mes`,
            `${month} AXI acumulado al mes`
        ];
        newHeaders.push(...headersForMonth);
        numericHeaders.push(...headersForMonth);
    });

    const firstRowContent = new Array(newHeaders.length).fill('');
    firstRowContent[0] = '<< Volver a la Tabla de Contenidos';
    firstRowContent[1] = "Cifras expresadas en miles de pesos argentinos – Elaborado en base a información publicada por el B.C.R.A y al Indice-FACPCE-Res.-JG-539-18. A los fines de esta aplicación, el ajuste por inflación está calculado – únicamente – para las cuentas de resultados, el cual no está calculado para los rubros no monetarios de las cuentas patrimoniales.";

    const axiCoefficients = allMonths.map((month, i) => {
        if (i === 0) return 0;
        const currentMonthIndex = indicesMap.get(month);
        const previousMonthIndex = indicesMap.get(allMonths[i - 1]);
        return (currentMonthIndex && previousMonthIndex) ? (currentMonthIndex / previousMonthIndex) - 1 : 0;
    });
    const axiRow = new Array(newHeaders.length).fill('');
    axiRow[3] = '% del Coeficiente AXI';
    allMonths.forEach((_, i) => {
        const colIndex = 4 + (i * 5) + 3; // Columna de AXI mensual
        axiRow[colIndex] = axiCoefficients[i];
    });

    const dataForSheet = [firstRowContent, axiRow, newHeaders];
    
    // 1. Separar cuentas de resultados y el resto
    const cuentasKeys = Object.keys(pivotedData).map(Number);
    const cuentasDeResultadosKeys = cuentasKeys.filter(c => c >= 500000 && c < 700000).sort((a, b) => a - b);
    const otrasCuentasKeys = cuentasKeys.filter(c => c < 500000 || c >= 700000).sort((a, b) => a - b);
    
    // Función auxiliar para procesar una lista de cuentas
    const processAccounts = (keys) => {
        const rows = [];
        for (const num_cuenta of keys) {
            const cuentaData = pivotedData[num_cuenta];
            const isAdjustable = (num_cuenta >= 500000 && num_cuenta < 700000);
            const rowObject = {
                'Entidad': infoEntidad.num_entidad,
                'Nombre Entidad': infoEntidad.nombre_entidad,
                'Cuenta': num_cuenta,
                'Descripción Cuenta': cuentaData.desc,
            };
            let saldoHistAcumuladoAnterior = 0, axiAcumuladoAnterior = 0, saldoMonedaConstanteAnterior = 0;
            allMonths.forEach((month, i) => {
                const saldoEnMonedaConstanteMes = cuentaData.saldos[month] || 0;
                let axiMensualMes = 0;
                if (isAdjustable) {
                    axiMensualMes = saldoMonedaConstanteAnterior * axiCoefficients[i];
                }
                const axiAcumuladoMes = axiAcumuladoAnterior + axiMensualMes;
                const saldoHistAcumuladoMes = saldoEnMonedaConstanteMes - axiAcumuladoMes;
                const saldoHistoricoMes = saldoHistAcumuladoMes - saldoHistAcumuladoAnterior;
                rowObject[`${month} Saldo en moneda constante`] = saldoEnMonedaConstanteMes;
                rowObject[`${month} Saldo Histórico solo del mes`] = saldoHistoricoMes;
                rowObject[`${month} Saldo Histórico acumulado al mes`] = saldoHistAcumuladoMes;
                rowObject[`${month} AXI mensual solo del mes`] = axiMensualMes;
                rowObject[`${month} AXI acumulado al mes`] = axiAcumuladoMes;
                saldoHistAcumuladoAnterior = saldoHistAcumuladoMes;
                axiAcumuladoAnterior = axiAcumuladoMes;
                saldoMonedaConstanteAnterior = saldoEnMonedaConstanteMes;
            });
            rows.push(newHeaders.map(header => rowObject[header] ?? ''));
        }
        return rows;
    };

    // 2. Procesar cuentas de resultados con subtotales
    if (cuentasDeResultadosKeys.length > 0) {
        let currentSubtotalGroup = null;
        let subtotalAccumulator = Object.fromEntries(numericHeaders.map(h => [h, 0]));
        const grandTotalAccumulator = Object.fromEntries(numericHeaders.map(h => [h, 0]));

        for (const num_cuenta of cuentasDeResultadosKeys) {
            const group = String(num_cuenta).substring(0, 2);
            if (currentSubtotalGroup !== group && currentSubtotalGroup !== null) {
                const subtotalRow = new Array(newHeaders.length).fill('');
                subtotalRow[3] = `Subtotal Cuentas ${currentSubtotalGroup}...`;
                numericHeaders.forEach(h => subtotalRow[newHeaders.indexOf(h)] = subtotalAccumulator[h]);
                dataForSheet.push(subtotalRow);
                subtotalAccumulator = Object.fromEntries(numericHeaders.map(h => [h, 0]));
            }
            currentSubtotalGroup = group;
            
            const [processedRowData] = processAccounts([num_cuenta]);
            dataForSheet.push(processedRowData);
            
            numericHeaders.forEach(h => {
                const value = processedRowData[newHeaders.indexOf(h)] || 0;
                subtotalAccumulator[h] += value;
                grandTotalAccumulator[h] += value;
            });
        }
        // Añadir el último subtotal
        const lastSubtotalRow = new Array(newHeaders.length).fill('');
        lastSubtotalRow[3] = `Subtotal Cuentas ${currentSubtotalGroup}...`;
        numericHeaders.forEach(h => lastSubtotalRow[newHeaders.indexOf(h)] = subtotalAccumulator[h]);
        dataForSheet.push(lastSubtotalRow);

        // Añadir el total general
        const grandTotalRow = new Array(newHeaders.length).fill('');
        grandTotalRow[3] = `Total Cuentas de Resultados`;
        numericHeaders.forEach(h => grandTotalRow[newHeaders.indexOf(h)] = grandTotalAccumulator[h]);
        dataForSheet.push(grandTotalRow);
    }
    
    // 3. Añadir un separador y el resto de las cuentas
    if (otrasCuentasKeys.length > 0 && cuentasDeResultadosKeys.length > 0) {
        dataForSheet.push(new Array(newHeaders.length).fill('')); // Fila en blanco como separador
    }
    dataForSheet.push(...processAccounts(otrasCuentasKeys));
    
    return dataForSheet;
}

// --- ENDPOINTS ---
app.get('/api/entidades', async (req, res) => {
    try {
        const nominaPath = path.join(__dirname, '../frontend/data/nomina.txt');
        if (!fs.existsSync(nominaPath)) return res.status(404).json({ message: 'Archivo nomina.txt no encontrado.' });
        res.json(await procesarNomina(nominaPath));
    } catch (error) {
        res.status(500).json({ message: 'Error interno al leer entidades.' });
    }
});

app.post('/generate-report', async (req, res) => {
    try {
        // ... (Procesamiento de datos inicial sin cambios) ...
        const filtros = req.body;
        const filePaths = { balhist: path.join(__dirname, '../frontend/data/balhist.txt'), cuentas: path.join(__dirname, '../frontend/data/cuentas.txt'), nomina: path.join(__dirname, '../frontend/data/nomina.txt'), indices: path.join(__dirname, '../frontend/data/indices.xlsx') };
        for (const key in filePaths) { if (!fs.existsSync(filePaths[key])) { return res.status(404).send(`Error: El archivo ${path.basename(filePaths[key])} no se encuentra.`); } }
        const datosBalhistFiltrados = await procesarBalhist(filePaths.balhist, filtros);
        if (datosBalhistFiltrados.length === 0) { return res.status(404).send('No se encontraron registros de balance con los filtros seleccionados.'); }
        const indicesMap = procesarIndices(filePaths.indices);
        if (indicesMap.size === 0) { return res.status(404).send('No se pudieron leer los datos del archivo indices.xlsx.'); }
        const cuentasData = await procesarCuentas(filePaths.cuentas);
        const nominaData = await procesarNomina(filePaths.nomina);
        const cuentasMap = new Map(cuentasData.map(c => [c.num_cuenta, c]));
        const nominaMap = new Map(nominaData.map(e => [e.num_entidad, e]));
        const workbook = xlsx.utils.book_new();
        const TOC_SHEET_NAME = 'Table of Contents';
        const allMonths = getMonthsInRange(filtros.balhistDesde, filtros.balhistHasta); 
        const balancesPorEntidad = new Map();
        for (const balance of datosBalhistFiltrados) {
            if (!balancesPorEntidad.has(balance.num_entidad)) balancesPorEntidad.set(balance.num_entidad, []);
            balancesPorEntidad.get(balance.num_entidad).push(balance);
        }
        const sortedEntityNumbers = Array.from(balancesPorEntidad.keys()).sort((a, b) => a - b);
        const tocSheetData = [['Hoja', 'Número de Entidad', 'Nombre de Entidad']];
        const sheetsToAppend = [];
        
        // --- SECCIÓN DE FORMATEO Y ESTILOS ---
        for (const num_entidad of sortedEntityNumbers) {
            const dataForSheet = prepareDataForSheet(balancesPorEntidad.get(num_entidad), cuentasMap, nominaMap, allMonths, indicesMap);
            if (dataForSheet.length <= 3) continue;

            const infoEntidad = nominaMap.get(num_entidad) || {};
            let sheetName = `${String(num_entidad).padStart(5, '0')} - ${infoEntidad.nombre_corto || infoEntidad.nombre_entidad || ''}`.trim().substring(0, 31).replace(/[\\/*?[\]]/g, '');
            tocSheetData.push([sheetName, num_entidad, infoEntidad.nombre_entidad || '']);
            const worksheet = xlsx.utils.aoa_to_sheet(dataForSheet, { skipHeader: true }); // skipHeader para manejarlo nosotros

            // --- Definición de Estilos ---
            const numberFormat2Decimals = '#.##0,00';
            const headerStyle = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "4F81BD" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
            const totalStyle = { font: { bold: true }, numFmt: numberFormat2Decimals };
            const subtotalStyle = { font: { bold: true, italic: true }, numFmt: numberFormat2Decimals, fill: { fgColor: { rgb: "F2F2F2" } } };
            const defaultNumericStyle = { numFmt: numberFormat2Decimals };
            const axiStyle = { numFmt: '0,0000' }; // 4 decimales para el coeficiente
            const disclaimerStyle = { font: { italic: true }, alignment: { wrapText: true, vertical: "center" } };

            // --- Aplicar Estilos y Formatos celda por celda ---
            const range = xlsx.utils.decode_range(worksheet['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_address = { c: C, r: R };
                    const cell_ref = xlsx.utils.encode_cell(cell_address);
                    if (!worksheet[cell_ref]) continue;

                    if (R === 2) { // Fila de Títulos
                        worksheet[cell_ref].s = headerStyle;
                    } else if (typeof worksheet[cell_ref].v === 'number') {
                        const cellValueText = dataForSheet[R] ? String(dataForSheet[R][3]) : "";
                        if (cellValueText.startsWith("Total")) {
                            worksheet[cell_ref].s = totalStyle;
                        } else if (cellValueText.startsWith("Subtotal")) {
                            worksheet[cell_ref].s = subtotalStyle;
                        } else if (R === 1 && C >= 4) { // Fila AXI
                             worksheet[cell_ref].s = axiStyle;
                        } else {
                            worksheet[cell_ref].s = defaultNumericStyle;
                        }
                    }
                }
            }
            
            // --- Formatos Especiales ---
            worksheet['A1'].l = { Target: `#'${TOC_SHEET_NAME}'!A1`, Tooltip: `Ir a ${TOC_SHEET_NAME}` };
            worksheet['B1'].s = disclaimerStyle;

            // --- Ancho de Columnas y Merges ---
            const colWidths = [ { wch: 10 }, { wch: 30 }, { wch: 12 }, { wch: 45 } ]; // Entidad, Nombre, Cuenta, Descripción
            allMonths.forEach(() => { colWidths.push({ wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }); });
            worksheet['!cols'] = colWidths;
            worksheet['!merges'] = [{ s: { r: 0, c: 1 }, e: { r: 0, c: 8 } }]; // Unir B1 a I1 para la leyenda
            worksheet['!rows'] = [{ hpt: 60 }, { hpt: 15 }, { hpt: 40 }]; // Altura para fila 1 (leyenda), 2 (AXI), 3 (Títulos)

            sheetsToAppend.push({ name: sheetName, sheet: worksheet });
        }
        
        if (sheetsToAppend.length === 0) { return res.status(404).send('No se generaron hojas. Verifique los datos de origen y filtros.'); }
        const tocWorksheet = xlsx.utils.aoa_to_sheet(tocSheetData);
        sheetsToAppend.forEach((sheetInfo, index) => {
            const cellAddress = `A${index + 2}`;
            if (tocWorksheet[cellAddress]) { tocWorksheet[cellAddress].l = { Target: `#'${sheetInfo.name}'!A1`, Tooltip: `Ir a la hoja ${sheetInfo.name}` }; }
        });
        tocWorksheet['!cols'] = [{ wch: 35 }, { wch: 15 }, { wch: 50 }];
        xlsx.utils.book_append_sheet(workbook, tocWorksheet, TOC_SHEET_NAME);
        sheetsToAppend.forEach(sheetInfo => { xlsx.utils.book_append_sheet(workbook, sheetInfo.sheet, sheetInfo.name); });
        
        const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        const nombreArchivo = `Reporte_Ajustado_Final_${filtros.balhistDesde}_a_${filtros.balhistHasta}.xlsx`;
        res.setHeader('Content-Disposition', `attachment; filename="${nombreArchivo}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.status(200).send(excelBuffer);
    } catch (processingError) {
        console.error("Error crítico durante el procesamiento:", processingError);
        res.status(500).send('Falló el proceso de la aplicación.');
    }
});

app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});