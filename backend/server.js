// backend/server.js (Versión Definitiva y Corregida - 100% Limpia)

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

// --- FUNCIONES DE PROCESAMIENTO ---
async function procesarCuentas(filePath) {
    const cuentas = [];
    const fileStream = fs.createReadStream(filePath, { encoding: 'latin1' });
    const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity });
    for await (const linea of rl) {
        if (linea.trim() === '') continue;
        const [numCuenta, descripcion] = linea.split('\t');
        if (!numCuenta || !descripcion) continue;
        cuentas.push({ num_cuenta: parseInt(numCuenta.replace(/"/g, ''), 10), descripcion_cuenta: descripcion.replace(/"/g, '').trim() });
    }
    return new Map(cuentas.map(c => [c.num_cuenta, c]));
}

async function procesarNomina(filePath) {
    const nomina = [];
    const fileStream = fs.createReadStream(filePath, { encoding: 'latin1' });
    const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity });
    for await (const linea of rl) {
        if (linea.trim() === '') continue;
        const [numEntidad, nombreEntidad, nombreCorto] = linea.split('\t');
        if (!numEntidad || !nombreEntidad) continue;
        nomina.push({ num_entidad: parseInt(numEntidad.replace(/"/g, ''), 10), nombre_entidad: nombreEntidad.replace(/"/g, '').trim(), nombre_corto: (nombreCorto || '').replace(/"/g, '').trim() });
    }
    return new Map(nomina.map(e => [e.num_entidad, e]));
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
    } catch (error) { console.error("Error al procesar indices.xlsx:", error); return new Map(); }
}

function getMonthsInRange(start, end) {
    const startDate = new Date(`${start}-01T00:00:00Z`);
    const endDate = new Date(`${end}-01T00:00:00Z`);
    let currentDate = startDate;
    const months = [];
    while (currentDate <= endDate) {
        const month = ('0' + (currentDate.getUTCMonth() + 1)).slice(-2);
        const year = currentDate.getUTCFullYear();
        months.push(`${month}-${year}`);
        currentDate.setUTCMonth(currentDate.getUTCMonth() + 1);
    }
    return months;
}

function prepareDataForSheet(balancesDeEstaEntidad, cuentasMap, nominaMap, allMonths, indicesMap, num_entidad) {
    if (!balancesDeEstaEntidad || balancesDeEstaEntidad.length === 0) return [];
    const infoEntidad = nominaMap.get(num_entidad) || { nombre_entidad: 'Desconocido', num_entidad };
    const pivotedData = {};
    for (const balance of balancesDeEstaEntidad) { if (!pivotedData[balance.num_cuenta]) { const desc = (cuentasMap.get(balance.num_cuenta) || {}).descripcion_cuenta || 'No encontrada'; pivotedData[balance.num_cuenta] = { desc, saldos: {} }; } pivotedData[balance.num_cuenta].saldos[balance.fecha_bce] = balance.saldo; }
    const newHeaders = ['Entidad', 'Nombre Entidad', 'Cuenta', 'Descripción Cuenta'];
    const numericHeaders = [];
    allMonths.forEach(month => { const headersForMonth = [`${month} Saldo en moneda constante`, `${month} Saldo Histórico solo del mes`, `${month} Saldo Histórico acumulado al mes`, `${month} AXI mensual solo del mes`, `${month} AXI acumulado al mes`]; newHeaders.push(...headersForMonth); numericHeaders.push(...headersForMonth); });
    const firstRowContent = new Array(newHeaders.length).fill(null);
    firstRowContent[0] = '<< Volver a la Tabla de Contenidos';
    firstRowContent[1] = "Cifras expresadas en miles de pesos argentinos – Elaborado en base a información publicada por el B.C.R.A y al Indice-FACPCE-Res.-JG-539-18. A los fines de esta aplicación, el ajuste por inflación está calculado – únicamente – para las cuentas de resultados, el cual no está calculado para los rubros no monetarios de las cuentas patrimoniales.";
    const axiCoefficients = allMonths.map((month, i) => { if (i === 0) return 0; const currentMonthIndex = indicesMap.get(month); const previousMonthIndex = indicesMap.get(allMonths[i - 1]); return (currentMonthIndex && previousMonthIndex) ? (currentMonthIndex / previousMonthIndex) - 1 : 0; });
    const axiRow = new Array(newHeaders.length).fill(null);
    axiRow[3] = '% del Coeficiente AXI';
    allMonths.forEach((_, i) => { axiRow[4 + (i * 5) + 3] = axiCoefficients[i]; });
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
                dataForSheet.push(new Array(newHeaders.length).fill(null)); 
                subtotalAccumulator = Object.fromEntries(numericHeaders.map(h => [h, 0]));
                currentGroup = group;
            }
            const processedRow = processAccountRow(num_cuenta);
            dataForSheet.push(processedRow);
            numericHeaders.forEach(h => { const value = processedRow[newHeaders.indexOf(h)] || 0; subtotalAccumulator[h] += value; grandTotalAccumulator[h] += value; });
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
    if (otrasCuentasKeys.length > 0) { if (cuentasDeResultadosKeys.length > 0) dataForSheet.push(new Array(newHeaders.length).fill(null)); otrasCuentasKeys.forEach(key => dataForSheet.push(processAccountRow(key))); }
    const emptyRow = new Array(newHeaders.length).fill(null);
    dataForSheet.push(emptyRow, emptyRow);
    dataForSheet.push(['Observaciones:']);
    dataForSheet.push(['Posibles causas que generen diferencias entre el AXI calculado en forma automática por esta app, con respecto al AXI real contabilizado por la entidad:']);
    dataForSheet.push(['- Ajustes contables con fecha valor, realizados a posteriori del cierre de la presentación al BCRA del respectivo balance mensual TXT y, por ende, que no hayan impactado realmente en el balance presentado ante el BCRA (pero en este caso el banco debiera haber realizado una nueva presentación ante el BCRA rectifcando el anterior balance).']);
    dataForSheet.push(['- En los casos en que el INDEC hubiere, a posteriori, rectificado o corregido o publicado un nuevo IPIM (y el banco hubiere utilizado el IPIM "provisorio" anteriormente publicado), ello podría generar diferencia en el AXI (debido a que esta app toma como dato para el cálculo del AXI, el balance TXT en moneda constante).']);
    dataForSheet.push(['- Causa real de diferencias: está App calcula (mediante "ingeniería matemática inversa") el AXI partiendo del saldo en moneda constante expresado en el miles de $, mientras que el banco realmente calcula el AXI partiendo del saldo histórico en CIFRAS COMPLETAS, lo cual es una fuente de pequeñas diferencias. Diferencia máxima estimada anual por simple redondeo a miles de $ : 500 (rendondeo) por 12 meses, igual a 6000 (en cifras completas), para cada cuenta contable de resultados.']);
    dataForSheet.push(['Saludos ... cuando pueda, seguimos ...']);
    return dataForSheet;
}

// --- ENDPOINTS ---
app.get('/api/entidades', async (req, res) => {
    try {
        const nominaPath = path.join(__dirname, '../frontend/data/nomina.txt');
        if (!fs.existsSync(nominaPath)) return res.status(404).json({ message: 'Archivo nomina.txt no encontrado.' });
        const nominaMap = await procesarNomina(nominaPath);
        res.json(Array.from(nominaMap.values()));
    } catch (error) { res.status(500).json({ message: 'Error interno al leer entidades.' }); }
});

app.post('/generate-report', async (req, res) => {
    try {
        console.log("Report generation started...");
        const filtros = req.body;
        const filePaths = { balhist: path.join(__dirname, '../frontend/data/balhist.txt'), cuentas: path.join(__dirname, '../frontend/data/cuentas.txt'), nomina: path.join(__dirname, '../frontend/data/nomina.txt'), indices: path.join(__dirname, '../frontend/data/indices.xlsx') };
        for (const key in filePaths) { if (!fs.existsSync(filePaths[key])) return res.status(404).send(`Error: El archivo ${path.basename(filePaths[key])} no se encuentra.`); }
        console.log("Loading lookup data (cuentas, nomina, indices)...");
        const [cuentasMap, nominaMap, indicesMap] = await Promise.all([ procesarCuentas(filePaths.cuentas), procesarNomina(filePaths.nomina), Promise.resolve(procesarIndices(filePaths.indices)) ]);
        console.log("Lookup data loaded.");
        const workbook = xlsx.utils.book_new();
        const TOC_SHEET_NAME = 'Table of Contents';
        const allMonths = getMonthsInRange(filtros.balhistDesde, filtros.balhistHasta);
        const tocSheetData = [['Hoja', 'Número de Entidad', 'Nombre de Entidad']];
        const balancesPorEntidad = new Map();
        const isAllEntities = filtros.entidad.includes("0");
        const selectedEntitiesSet = isAllEntities ? null : new Set(filtros.entidad.map(Number));
        console.log("Starting to stream and process balhist.txt...");
        const fileStream = fs.createReadStream(filePaths.balhist, { encoding: 'latin1' });
        const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity });
        for await (const linea of rl) {
            const [numEntidadStr, fechaBceStr, numCuentaStr, saldoStr] = linea.split('\t');
            if (!numEntidadStr || !fechaBceStr || !numCuentaStr || saldoStr === undefined) continue;
            const entidadActual = parseInt(numEntidadStr.replace(/"/g, ''), 10);
            if (!isAllEntities && !selectedEntitiesSet.has(entidadActual)) continue;
            const anio = fechaBceStr.replace(/"/g, '').substring(0, 4);
            const mes = fechaBceStr.replace(/"/g, '').substring(4, 6);
            const fechaComparable = `${anio}-${mes}`;
            if (fechaComparable >= filtros.balhistDesde && fechaComparable <= filtros.balhistHasta) {
                if (!balancesPorEntidad.has(entidadActual)) balancesPorEntidad.set(entidadActual, []);
                balancesPorEntidad.get(entidadActual).push({ fecha_bce: `${mes}-${anio}`, num_cuenta: parseInt(numCuentaStr.replace(/"/g, ''), 10), saldo: parseInt(saldoStr.trim(), 10) });
            }
        }
        console.log(`Finished processing balhist.txt. Found data for ${balancesPorEntidad.size} entities.`);
        if (balancesPorEntidad.size === 0) return res.status(404).send('No se encontraron registros de balance con los filtros seleccionados.');
        const sortedEntityNumbers = Array.from(balancesPorEntidad.keys()).sort((a, b) => a - b);
        for (const num_entidad of sortedEntityNumbers) {
            console.log(`Generating sheet for entity ${num_entidad}...`);
            const entityBalances = balancesPorEntidad.get(num_entidad);
            const dataForSheet = prepareDataForSheet(entityBalances, cuentasMap, nominaMap, allMonths, indicesMap, num_entidad);
            if (dataForSheet.length <= 3) { console.log(`Skipping sheet for entity ${num_entidad} due to no data.`); continue; }
            const infoEntidad = nominaMap.get(num_entidad) || {};
            let sheetName = `${String(num_entidad).padStart(5, '0')} - ${infoEntidad.nombre_corto || infoEntidad.nombre_entidad || ''}`.trim().substring(0, 31).replace(/[\\/*?[\]]/g, '');
            tocSheetData.push([sheetName, num_entidad, infoEntidad.nombre_entidad || '']);
            const worksheet = xlsx.utils.aoa_to_sheet(dataForSheet);

            const numberFormat2Decimals = '#,##0.00';
            const percentFormat4Decimals = '0.0000%';
            const headerStyle = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "4F81BD" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
            const disclaimerStyle = { font: { italic: true, sz: 9 }, alignment: { wrapText: true, vertical: "center" } };
            const obsTitleStyle = { font: { bold: true, sz: 12 } };
            const obsBodyStyle = { font: { sz: 10 }, alignment: { wrapText: true, vertical: "top" } };
            const totalNumericStyle = { font: { bold: true }, numFmt: numberFormat2Decimals, fill: { fgColor: { rgb: "FFFF00" } } };
            const subtotalNumericStyle = { font: { bold: true, italic: true }, numFmt: numberFormat2Decimals, fill: { fgColor: { rgb: "D3D3D3" } } };
            const totalTextStyle = { font: { bold: true }, fill: { fgColor: { rgb: "FFFF00" } } };
            const subtotalTextStyle = { font: { bold: true, italic: true }, fill: { fgColor: { rgb: "D3D3D3" } } };

            const range = xlsx.utils.decode_range(worksheet['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                const descCellValue = worksheet[xlsx.utils.encode_cell({c: 3, r: R})]?.v || "";
                const isTotalRow = descCellValue.startsWith("Total");
                const isSubtotalRow = descCellValue.startsWith("Subtotal");

                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_ref = xlsx.utils.encode_cell({ c: C, r: R });
                    const cell = worksheet[cell_ref];
                    if (!cell || !cell.v) continue;
                    
                    const cellValueStr = cell.v.toString();
                    if (R === 2) { cell.s = headerStyle; continue; }
                    if (cellValueStr.startsWith('Observaciones:')) { cell.s = obsTitleStyle; continue; } 
                    if (cellValueStr.startsWith('Posibles causas') || cellValueStr.startsWith('- ') || cellValueStr.startsWith('Saludos')) { cell.s = obsBodyStyle; continue; } 

                    if (isTotalRow) {
                        cell.s = (cell.t === 'n') ? totalNumericStyle : totalTextStyle;
                    } else if (isSubtotalRow) {
                        cell.s = (cell.t === 'n') ? subtotalNumericStyle : subtotalTextStyle;
                    } else if (cell.t === 'n') {
                        if (R === 1) { cell.z = percentFormat4Decimals; }
                        else { cell.z = numberFormat2Decimals; }
                    }
                }
            }
            if (worksheet['A1']) worksheet['A1'].l = { Target: `#'${TOC_SHEET_NAME}'!A1`, Tooltip: `Ir a ${TOC_SHEET_NAME}` };
            if (worksheet['B1']) worksheet['B1'].s = disclaimerStyle;
            const obsStartRow = dataForSheet.findIndex(row => typeof row[0] === 'string' && row[0].startsWith('Observaciones:'));
            if (obsStartRow !== -1) {
                if (!worksheet['!merges']) worksheet['!merges'] = [];
                for (let R = obsStartRow; R < dataForSheet.length; R++) { worksheet['!merges'].push({ s: { r: R, c: 0 }, e: { r: R, c: 8 } }); }
            }
            const colWidths = [ { wch: 10 }, { wch: 30 }, { wch: 12 }, { wch: 45 } ];
            allMonths.forEach(() => { colWidths.push({ wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }); });
            worksheet['!cols'] = colWidths;
            if (!worksheet['!merges']) worksheet['!merges'] = [];
            const b1MergeExists = worksheet['!merges'].some(m => m.s.r === 0 && m.s.c === 1);
            if (!b1MergeExists) { worksheet['!merges'].push({ s: { r: 0, c: 1 }, e: { r: 0, c: 8 } }); }
            worksheet['!rows'] = [{ hpt: 35 }, null, { hpt: 30 }];
            xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
        }
        console.log("All sheets generated. Finalizing workbook...");
        const tocWorksheet = xlsx.utils.aoa_to_sheet(tocSheetData);
        tocWorksheet['!cols'] = [{ wch: 35 }, { wch: 15 }, { wch: 50 }];
        tocSheetData.slice(1).forEach((row, index) => { const sheetName = row[0]; const cellAddress = `A${index + 2}`; if (tocWorksheet[cellAddress]) { tocWorksheet[cellAddress].l = { Target: `#'${sheetName}'!A1`, Tooltip: `Ir a la hoja ${sheetName}` }; } });
        xlsx.utils.book_append_sheet(workbook, tocWorksheet, TOC_SHEET_NAME);
        workbook.SheetNames.splice(workbook.SheetNames.indexOf(TOC_SHEET_NAME), 1);
        workbook.SheetNames.unshift(TOC_SHEET_NAME);
        const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        const nombreArchivo = `Reporte_Ajustado_Final_${filtros.balhistDesde}_a_${filtros.balhistHasta}.xlsx`;
        res.setHeader('Content-Disposition', `attachment; filename="${nombreArchivo}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        console.log("Sending file to client.");
        res.status(200).send(excelBuffer);
    } catch (processingError) {
        console.error("Error crítico durante el procesamiento:", processingError);
        res.status(500).send('Falló el proceso de la aplicación.');
    }
});

app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});