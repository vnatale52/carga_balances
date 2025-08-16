// backend/server.js (Versión con Formato Avanzado y Subtotales CORREGIDA)

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

// --- FUNCIONES DE PROCESAMIENTO OPTIMIZADAS (Sin cambios) ---
async function procesarBalhist(filePath, filtros) { /* ...código sin cambios... */ }
async function procesarCuentas(filePath) { /* ...código sin cambios... */ }
async function procesarNomina(filePath) { /* ...código sin cambios... */ }
function procesarIndices(filePath) { /* ...código sin cambios... */ }
function getMonthsInRange(start, end) { /* ...código sin cambios... */ }

// --- FUNCIÓN CENTRAL CON LÓGICA DE SUBTOTALES CORREGIDA ---
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
        const headersForMonth = [`${month} Saldo en moneda constante`, `${month} Saldo Histórico solo del mes`, `${month} Saldo Histórico acumulado al mes`, `${month} AXI mensual solo del mes`, `${month} AXI acumulado al mes`];
        newHeaders.push(...headersForMonth);
        numericHeaders.push(...headersForMonth);
    });

    const firstRowContent = new Array(newHeaders.length).fill(null);
    firstRowContent[0] = '<< Volver a la Tabla de Contenidos';
    firstRowContent[1] = "Cifras expresadas en miles de pesos argentinos – Elaborado en base a información publicada por el B.C.R.A y al Indice-FACPCE-Res.-JG-539-18. A los fines de esta aplicación, el ajuste por inflación está calculado – únicamente – para las cuentas de resultados, el cual no está calculado para los rubros no monetarios de las cuentas patrimoniales.";

    const axiCoefficients = allMonths.map((month, i) => {
        if (i === 0) return 0;
        const currentMonthIndex = indicesMap.get(month);
        const previousMonthIndex = indicesMap.get(allMonths[i - 1]);
        return (currentMonthIndex && previousMonthIndex) ? (currentMonthIndex / previousMonthIndex) - 1 : 0;
    });
    const axiRow = new Array(newHeaders.length).fill(null);
    axiRow[3] = '% del Coeficiente AXI';
    allMonths.forEach((_, i) => {
        axiRow[4 + (i * 5) + 3] = axiCoefficients[i];
    });

    const dataForSheet = [firstRowContent, axiRow, newHeaders];
    
    const cuentasKeys = Object.keys(pivotedData).map(Number);
    const cuentasDeResultadosKeys = cuentasKeys.filter(c => c >= 500000 && c < 700000).sort((a, b) => a - b);
    const otrasCuentasKeys = cuentasKeys.filter(c => c < 500000 || c >= 700000).sort((a, b) => a - b);
    
    const processAccountRow = (num_cuenta) => {
        const cuentaData = pivotedData[num_cuenta];
        const isAdjustable = (num_cuenta >= 500000 && num_cuenta < 700000);
        const rowObject = { 'Entidad': infoEntidad.num_entidad, 'Nombre Entidad': infoEntidad.nombre_entidad, 'Cuenta': num_cuenta, 'Descripción Cuenta': cuentaData.desc };
        let saldoHistAcumuladoAnterior = 0, axiAcumuladoAnterior = 0, saldoMonedaConstanteAnterior = 0;
        allMonths.forEach((month, i) => {
            const saldoEnMonedaConstanteMes = cuentaData.saldos[month] || 0;
            let axiMensualMes = isAdjustable ? saldoMonedaConstanteAnterior * axiCoefficients[i] : 0;
            const axiAcumuladoMes = axiAcumuladoAnterior + axiMensualMes;
            const saldoHistAcumuladoMes = saldoEnMonedaConstanteMes - axiAcumuladoMes;
            const saldoHistoricoMes = saldoHistAcumuladoMes - saldoHistAcumuladoAnterior;
            rowObject[`${month} Saldo en moneda constante`] = saldoEnMonedaConstanteMes; rowObject[`${month} Saldo Histórico solo del mes`] = saldoHistoricoMes; rowObject[`${month} Saldo Histórico acumulado al mes`] = saldoHistAcumuladoMes; rowObject[`${month} AXI mensual solo del mes`] = axiMensualMes; rowObject[`${month} AXI acumulado al mes`] = axiAcumuladoMes;
            saldoHistAcumuladoAnterior = saldoHistAcumuladoMes; axiAcumuladoAnterior = axiAcumuladoMes; saldoMonedaConstanteAnterior = saldoEnMonedaConstanteMes;
        });
        return newHeaders.map(header => rowObject[header] ?? null);
    };

    if (cuentasDeResultadosKeys.length > 0) {
        let subtotalAccumulator = Object.fromEntries(numericHeaders.map(h => [h, 0]));
        const grandTotalAccumulator = Object.fromEntries(numericHeaders.map(h => [h, 0]));
        let currentGroup = String(cuentasDeResultadosKeys[0]).substring(0, 2);

        for (let i = 0; i < cuentasDeResultadosKeys.length; i++) {
            const num_cuenta = cuentasDeResultadosKeys[i];
            const group = String(num_cuenta).substring(0, 2);

            if (group !== currentGroup) {
                const subtotalRow = new Array(newHeaders.length).fill(null);
                subtotalRow[3] = `Subtotal Cuentas ${currentGroup}...`;
                numericHeaders.forEach(h => subtotalRow[newHeaders.indexOf(h)] = subtotalAccumulator[h]);
                dataForSheet.push(subtotalRow);
                subtotalAccumulator = Object.fromEntries(numericHeaders.map(h => [h, 0]));
                currentGroup = group;
            }

            const processedRow = processAccountRow(num_cuenta);
            dataForSheet.push(processedRow);
            
            numericHeaders.forEach(h => {
                const value = processedRow[newHeaders.indexOf(h)] || 0;
                subtotalAccumulator[h] += value;
                grandTotalAccumulator[h] += value;
            });

            if (i === cuentasDeResultadosKeys.length - 1) {
                const lastSubtotalRow = new Array(newHeaders.length).fill(null);
                lastSubtotalRow[3] = `Subtotal Cuentas ${currentGroup}...`;
                numericHeaders.forEach(h => lastSubtotalRow[newHeaders.indexOf(h)] = subtotalAccumulator[h]);
                dataForSheet.push(lastSubtotalRow);
            }
        }

        const grandTotalRow = new Array(newHeaders.length).fill(null);
        grandTotalRow[3] = `Total Cuentas de Resultados`;
        numericHeaders.forEach(h => grandTotalRow[newHeaders.indexOf(h)] = grandTotalAccumulator[h]);
        dataForSheet.push(grandTotalRow);
    }
    
    if (otrasCuentasKeys.length > 0) {
        if (cuentasDeResultadosKeys.length > 0) dataForSheet.push(new Array(newHeaders.length).fill(null));
        otrasCuentasKeys.forEach(key => dataForSheet.push(processAccountRow(key)));
    }
    
    return dataForSheet;
}

// --- ENDPOINTS ---
app.get('/api/entidades', async (req, res) => { /* ...código sin cambios... */ });

app.post('/generate-report', async (req, res) => {
    try {
        const filtros = req.body;
        const filePaths = { balhist: path.join(__dirname, '../frontend/data/balhist.txt'), cuentas: path.join(__dirname, '../frontend/data/cuentas.txt'), nomina: path.join(__dirname, '../frontend/data/nomina.txt'), indices: path.join(__dirname, '../frontend/data/indices.xlsx') };
        for (const key in filePaths) { if (!fs.existsSync(filePaths[key])) { return res.status(404).send(`Error: El archivo ${path.basename(filePaths[key])} no se encuentra.`); } }
        
        const [datosBalhistFiltrados, cuentasData, nominaData] = await Promise.all([
            procesarBalhist(filePaths.balhist, filtros),
            procesarCuentas(filePaths.cuentas),
            procesarNomina(filePaths.nomina)
        ]);
        
        if (datosBalhistFiltrados.length === 0) { return res.status(404).send('No se encontraron registros de balance con los filtros seleccionados.'); }
        const indicesMap = procesarIndices(filePaths.indices);
        if (indicesMap.size === 0) { return res.status(404).send('No se pudieron leer los datos del archivo indices.xlsx.'); }
        
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
        
        for (const num_entidad of sortedEntityNumbers) {
            const dataForSheet = prepareDataForSheet(balancesPorEntidad.get(num_entidad), cuentasMap, nominaMap, allMonths, indicesMap);
            if (dataForSheet.length <= 3) continue;

            const infoEntidad = nominaMap.get(num_entidad) || {};
            let sheetName = `${String(num_entidad).padStart(5, '0')} - ${infoEntidad.nombre_corto || infoEntidad.nombre_entidad || ''}`.trim().substring(0, 31).replace(/[\\/*?[\]]/g, '');
            tocSheetData.push([sheetName, num_entidad, infoEntidad.nombre_entidad || '']);

            // CORRECCIÓN: Se elimina {skipHeader: true} para que la librería procese todas nuestras filas correctamente.
            const worksheet = xlsx.utils.aoa_to_sheet(dataForSheet);

            const numberFormat2Decimals = '#,##0.00'; // Formato universal para separador de miles y 2 decimales
            const headerStyle = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "4F81BD" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
            const totalStyle = { font: { bold: true }, numFmt: numberFormat2Decimals };
            const subtotalStyle = { font: { bold: true, italic: true }, numFmt: numberFormat2Decimals, fill: { fgColor: { rgb: "F2F2F2" } } };
            const axiStyle = { numFmt: '0.0000' };
            const disclaimerStyle = { font: { italic: true, sz: 9 }, alignment: { wrapText: true, vertical: "center" } };

            const range = xlsx.utils.decode_range(worksheet['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_ref = xlsx.utils.encode_cell({ c: C, r: R });
                    const cell = worksheet[cell_ref];
                    if (!cell) continue;

                    // Aplicar estilos basados en la fila y contenido
                    if (R === 2) { // Fila de Títulos (índice 2)
                        cell.s = headerStyle;
                    } else if (cell.t === 'n') { // Solo para celdas numéricas
                        const descCellValue = worksheet[xlsx.utils.encode_cell({c: 3, r: R})]?.v || "";
                        if (descCellValue.startsWith("Total")) {
                            cell.s = totalStyle;
                        } else if (descCellValue.startsWith("Subtotal")) {
                            cell.s = subtotalStyle;
                        } else if (R === 1) { // Fila AXI
                             cell.s = axiStyle;
                        } else {
                            cell.z = numberFormat2Decimals; // Usar .z es más simple para solo formato
                        }
                    }
                }
            }
            
            worksheet['A1'].l = { Target: `#'${TOC_SHEET_NAME}'!A1`, Tooltip: `Ir a ${TOC_SHEET_NAME}` };
            if (worksheet['B1']) worksheet['B1'].s = disclaimerStyle;

            const colWidths = [ { wch: 10 }, { wch: 30 }, { wch: 12 }, { wch: 45 } ];
            allMonths.forEach(() => { colWidths.push({ wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }); });
            worksheet['!cols'] = colWidths;
            worksheet['!merges'] = [{ s: { r: 0, c: 1 }, e: { r: 0, c: 8 } }];
            worksheet['!rows'] = [{ hpt: 60 }, null, { hpt: 40 }];

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

// Re-incluyo las funciones que se habían omitido por brevedad en la respuesta anterior
// (Asegúrate de que estas funciones estén presentes en tu archivo)
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