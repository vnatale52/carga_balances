// backend/server.js (Versión Final: Ordenamiento por Entidad/Fecha/Cuenta + Título Actualizado)

import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import xlsx from 'xlsx-js-style'; 
import readline from 'readline';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// --- MIDDLEWARE ---
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

// --- PROCESAR EL PLAN DE CUENTAS ---
function procesarPlanCuentas(filePath) {
    try {
        if (!fs.existsSync(filePath)) return { descMap: new Map(), rubroMap: new Map() };
        
        const buffer = fs.readFileSync(filePath);
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
        
        const descMap = new Map();
        const rubroMap = new Map();

        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length < 2) continue;
            
            const cuentaRaw = row[0]; 
            const descPlan = row[1]; 
            const rubroPlan = row[2]; 

            if (cuentaRaw) {
                const numCuenta = parseInt(String(cuentaRaw).replace(/\D/g, ''), 10);
                if (!isNaN(numCuenta)) {
                    descMap.set(numCuenta, String(descPlan || '').trim());
                    if (rubroPlan) {
                        rubroMap.set(numCuenta, String(rubroPlan).trim());
                    }
                }
            }
        }
        return { descMap, rubroMap };
    } catch (error) {
        console.error("Error al procesar archivo de Plan de Ctas:", error);
        return { descMap: new Map(), rubroMap: new Map() };
    }
}

// --- OBTENER DATOS CRUDOS DEL PLAN ---
function obtenerDatosRawPlan(filePath) {
    try {
        if (!fs.existsSync(filePath)) return [];
        const buffer = fs.readFileSync(filePath);
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        return xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    } catch (error) {
        console.error("Error leyendo raw plan:", error);
        return [];
    }
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
        console.error("Error al procesar indices.xlsx:", error);
        return new Map();
    }
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

// --- LÓGICA DE PREPARACIÓN DE DATOS ---
function prepareDataForSheet(balancesDeEstaEntidad, cuentasMap, nominaMap, planMap, allMonths, indicesMap, num_entidad) {
    if (!balancesDeEstaEntidad || balancesDeEstaEntidad.length === 0) return [];
    const infoEntidad = nominaMap.get(num_entidad) || { nombre_entidad: 'Desconocido', num_entidad };
    
    const pivotedData = {};
    for (const balance of balancesDeEstaEntidad) { 
        if (!pivotedData[balance.num_cuenta]) { 
            const desc = (cuentasMap.get(balance.num_cuenta) || {}).descripcion_cuenta || 'No encontrada'; 
            const descPlan = planMap.get(balance.num_cuenta) || '';
            pivotedData[balance.num_cuenta] = { desc, descPlan, saldos: {} }; 
        } 
        pivotedData[balance.num_cuenta].saldos[balance.fecha_bce] = balance.saldo; 
    }

    const newHeaders = ['Entidad', 'Nombre Entidad', 'Cuenta', 'Descripción Cuenta', 'Descripción de la cuenta según Plan de Cuentas del BCRA'];
    const numericHeaders = [];
    const axiRow = [null, null, null, null, '% del Coeficiente AXI']; 

    const axiCoefficients = allMonths.map((month, i) => { 
        if (i === 0) return 0; 
        const currentMonthIndex = indicesMap.get(month); 
        const previousMonthIndex = indicesMap.get(allMonths[i - 1]); 
        return (currentMonthIndex && previousMonthIndex) ? (currentMonthIndex / previousMonthIndex) - 1 : 0; 
    });
    
    allMonths.forEach((month, i) => {
        const [mes, anio] = month.split('-');

        if (mes === '01' && i > 0) {
            newHeaders.push(`Cuenta (${anio})`, `Descripción Cuenta (${anio})`);
            axiRow.push(null, null);
        }

        const headersForMonth = [
            `${month} Saldo en moneda constante`, 
            `${month} Saldo Histórico solo del mes`, 
            `${month} Saldo Histórico acumulado al mes`, 
            `${month} AXI mensual solo del mes`, 
            `${month} AXI acumulado al mes`
        ];
        newHeaders.push(...headersForMonth);
        numericHeaders.push(...headersForMonth);

        axiRow.push(null, null, null, axiCoefficients[i], null);

        if (mes === '12') {
            newHeaders.push(`${month}__Control_Incoherencia__`); 
            axiRow.push(null);

            const headersRefundicion = [
                `${month} Refundición del Saldo en moneda constante`,
                `${month} Refundición del Saldo Histórico solo del mes`,
                `${month} Refundición del Saldo Histórico acumulado al mes`,
                `${month} Refundición del AXI mensual solo del mes`,
                `${month} Refundición del AXI acumulado al mes`
            ];
            newHeaders.push(...headersRefundicion);
            numericHeaders.push(...headersRefundicion);
            axiRow.push(null, null, null, null, null);

            newHeaders.push(`${month}__SEPARATOR__`);
            axiRow.push(null);
        }
    });

    const firstRowContent = new Array(newHeaders.length).fill(null);
    firstRowContent[0] = '<<== Back to TOC';
    firstRowContent[1] = "Formatea esta hoja a tu gusto. Cifras expresadas en miles de pesos argentinos. Elaborado en base a información publicada por el B.C.R.A y al Indice-FACPCE-Res.-JG-539-18.   A los fines específicos de esta aplicación, el ajuste por inflación está calculado – únicamente – para las cuentas de resultados, es decir, no está calculado también para los rubros no monetarios de las cuentas patrimoniales (por ejemplo, Bienes de Uso, Intangibles y cuentas del Patrimonio Neto).";

    const dataForSheet = [firstRowContent, axiRow, newHeaders];
    const cuentasKeys = Object.keys(pivotedData).map(Number);
    const cuentasDeResultadosKeys = cuentasKeys.filter(c => c >= 500000 && c < 700000).sort((a, b) => a - b);
    const otrasCuentasKeys = cuentasKeys.filter(c => c < 500000 || c >= 700000).sort((a, b) => a - b);

    const processAccountRow = (num_cuenta) => {
        const cuentaData = pivotedData[num_cuenta];
        const isAdjustable = (num_cuenta >= 500000 && num_cuenta < 700000); 
        const isRecpam = String(num_cuenta).startsWith('62');

        const rowObject = { 
            'Entidad': infoEntidad.num_entidad, 
            'Nombre Entidad': infoEntidad.nombre_entidad, 
            'Cuenta': num_cuenta, 
            'Descripción Cuenta': cuentaData.desc,
            'Descripción de la cuenta según Plan de Cuentas del BCRA': cuentaData.descPlan
        };

        let saldoHistAcumuladoAnterior = 0;
        let axiAcumuladoAnterior = 0;
        let saldoMonedaConstanteAnterior = 0;

        allMonths.forEach((month, i) => {
            const [mes, anio] = month.split('-');

            if (mes === '01' && i > 0) {
                rowObject[`Cuenta (${anio})`] = num_cuenta;
                rowObject[`Descripción Cuenta (${anio})`] = cuentaData.desc;
            }

            const saldoEnMonedaConstanteMes = cuentaData.saldos[month] || 0;
            let axiMensualMes = 0;
            let axiAcumuladoMes = 0;
            let saldoHistAcumuladoMes = 0;
            let saldoHistoricoMes = 0;

            if (isRecpam) {
                // Solo moneda constante
            } else {
                axiMensualMes = isAdjustable ? Math.round((saldoMonedaConstanteAnterior * axiCoefficients[i]) * 100) / 100 : 0;
                
                axiAcumuladoMes = axiAcumuladoAnterior + axiMensualMes;
                saldoHistAcumuladoMes = saldoEnMonedaConstanteMes - axiAcumuladoMes;
                saldoHistoricoMes = saldoHistAcumuladoMes - saldoHistAcumuladoAnterior;
            }

            rowObject[`${month} Saldo en moneda constante`] = saldoEnMonedaConstanteMes; 
            rowObject[`${month} Saldo Histórico solo del mes`] = saldoHistoricoMes; 
            rowObject[`${month} Saldo Histórico acumulado al mes`] = saldoHistAcumuladoMes; 
            rowObject[`${month} AXI mensual solo del mes`] = axiMensualMes; 
            rowObject[`${month} AXI acumulado al mes`] = axiAcumuladoMes;

            if (isRecpam) {
                saldoHistAcumuladoAnterior = 0; 
                axiAcumuladoAnterior = 0; 
                saldoMonedaConstanteAnterior = saldoEnMonedaConstanteMes;
            } else {
                saldoHistAcumuladoAnterior = saldoHistAcumuladoMes; 
                axiAcumuladoAnterior = axiAcumuladoMes; 
                saldoMonedaConstanteAnterior = saldoEnMonedaConstanteMes;
            }

            if (mes === '12') {
                let mensajeIncoherencia = "";
                if (Math.abs(saldoEnMonedaConstanteMes) < 0.01) { 
                     if (Math.abs(saldoHistAcumuladoMes) > 0.01 || Math.abs(axiAcumuladoMes) > 0.01) {
                         mensajeIncoherencia = "Incoherencia"; 
                     }
                }
                rowObject[`${month}__Control_Incoherencia__`] = mensajeIncoherencia;

                rowObject[`${month} Refundición del Saldo en moneda constante`] = 0;
                rowObject[`${month} Refundición del Saldo Histórico solo del mes`] = 0;
                rowObject[`${month} Refundición del Saldo Histórico acumulado al mes`] = 0;
                rowObject[`${month} Refundición del AXI mensual solo del mes`] = 0;
                rowObject[`${month} Refundición del AXI acumulado al mes`] = 0;

                rowObject[`${month}__SEPARATOR__`] = null;

                if (isAdjustable) {
                    saldoHistAcumuladoAnterior = 0;
                    axiAcumuladoAnterior = 0;
                    saldoMonedaConstanteAnterior = 0;
                }
            }
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

    if (otrasCuentasKeys.length > 0) { 
        if (cuentasDeResultadosKeys.length > 0) dataForSheet.push(new Array(newHeaders.length).fill(null)); 
        otrasCuentasKeys.forEach(key => dataForSheet.push(processAccountRow(key))); 
    }
    
    const emptyRow = new Array(newHeaders.length).fill(null);
    dataForSheet.push(emptyRow, emptyRow);
    dataForSheet.push(['Observaciones del sistema:']);
    dataForSheet.push(['- Es imprescindible la descarga y lectura del archivo muy_importante_ajuste_por_inflacion.pdf']);

    return dataForSheet;
}

// --- ENDPOINTS ---
app.get('/api/entidades', async (req, res) => { try { const nominaPath = path.join(__dirname, '../frontend/data/nomina.txt'); if (!fs.existsSync(nominaPath)) return res.status(404).json({ message: 'Archivo nomina.txt no encontrado.' }); const nominaMap = await procesarNomina(nominaPath); res.json(Array.from(nominaMap.values())); } catch (error) { res.status(500).json({ message: 'Error interno al leer entidades.' }); } });

app.post('/generate-report', async (req, res) => {
    try {
        console.log("Report generation started...");
        const filtros = req.body;
        const filePaths = { 
            balhist: path.join(__dirname, '../frontend/data/balhist.txt'), 
            cuentas: path.join(__dirname, '../frontend/data/cuentas.txt'), 
            nomina: path.join(__dirname, '../frontend/data/nomina.txt'), 
            indices: path.join(__dirname, '../frontend/data/indices.xlsx'),
            plan: path.join(__dirname, '../frontend/data/Plan de Ctas de Resultados y su Rubro de Exposicion.xlsx')
        };
        
        for (const key in filePaths) { 
            if (!fs.existsSync(filePaths[key])) {
                return res.status(404).send(`Error: El archivo ${path.basename(filePaths[key])} no se encuentra.`); 
            }
        }
        
        console.log("Loading lookup data...");
        const [cuentasMap, nominaMap, indicesMap, { descMap, rubroMap }] = await Promise.all([ 
            procesarCuentas(filePaths.cuentas), 
            procesarNomina(filePaths.nomina), 
            Promise.resolve(procesarIndices(filePaths.indices)),
            Promise.resolve(procesarPlanCuentas(filePaths.plan))
        ]);
        
        const planRawData = obtenerDatosRawPlan(filePaths.plan);

        const workbook = xlsx.utils.book_new();
        const TOC_SHEET_NAME = 'Table of Contents';
        const allMonths = getMonthsInRange(filtros.balhistDesde, filtros.balhistHasta);
        const tocSheetData = [['Hoja', 'Número de Entidad', 'Nombre de Entidad']];
        const balancesPorEntidad = new Map();
        const isAllEntities = filtros.entidad.includes("0");
        const selectedEntitiesSet = isAllEntities ? null : new Set(filtros.entidad.map(Number));
        
        console.log("Processing balhist.txt...");
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
        
        if (balancesPorEntidad.size === 0) return res.status(404).send('No se encontraron registros.');
        const sortedEntityNumbers = Array.from(balancesPorEntidad.keys()).sort((a, b) => a - b);

        // CAMBIO: Header con Texto de Advertencia
        const decemberSummaryData = [[
            'Nombre Entidad', 
            'Entidad', 
            'Fecha', 
            'Cuenta', 
            'Descripción Cuenta', 
            'Descripción de la cuenta según Plan de Cuentas del BCRA', 
            'Saldo en moneda constante según Balance TXT del BCRA, en miles de $', 
            'Rubro de Exposición en Estado de Resultados - Tabla de Conversión RI-NIIF (si faltara agregar algún Rubro, habrá una diferencia; cuidado, esta App utiliza la versión publicada al 21-11-2025)',
            ''
        ]];

        // ESTILOS
        const defaultFont = { name: "Arial", sz: 9 }; 
        const allBorders = { top: { style: "thin", color: { auto: 1 } }, bottom: { style: "thin", color: { auto: 1 } }, left: { style: "thin", color: { auto: 1 } }, right: { style: "thin", color: { auto: 1 } } };
        
        const numberFormatInteger = '#,##0';     
        const numberFormatDecimal = '#,##0.00';  
        const percentFormat6Decimals = '0.000000%'; 

        const headerStyle = { font: { name: "Arial", sz: 8, bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "4F81BD" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: allBorders };
        const totalStyle = { font: { ...defaultFont, bold: true }, numFmt: numberFormatDecimal, fill: { fgColor: { rgb: "FFFF00" } }, border: allBorders };
        const subtotalStyle = { font: { ...defaultFont, bold: true, italic: true }, numFmt: numberFormatDecimal, fill: { fgColor: { rgb: "D3D3D3" } }, border: allBorders };
        const defaultCellStyle = { font: defaultFont, border: allBorders };
        
        const integerFormatStyle = { ...defaultCellStyle, numFmt: "0" };
        const integerMoneyStyle = { ...defaultCellStyle, numFmt: numberFormatInteger }; 
        const decimalFormatStyle = { ...defaultCellStyle, numFmt: numberFormatDecimal }; 
        const percent6Style = { ...defaultCellStyle, numFmt: percentFormat6Decimals }; 
        const errorTextStyle = { ...defaultCellStyle, font: { ...defaultFont, color: { rgb: "FF0000" }, bold: true }, alignment: { wrapText: true } };
        const obsTitleStyle = { font: { ...defaultFont, sz: 10, bold: true } };
        const obsBodyStyle = { font: defaultFont, alignment: { wrapText: true, vertical: "top" } };
        const disclaimerStyle = { font: { name: "Arial", sz: 10, bold: true, italic: true }, alignment: { horizontal: "centerAcross", vertical: "center" } };

        for (const num_entidad of sortedEntityNumbers) {
            const entityBalances = balancesPorEntidad.get(num_entidad);
            const infoEntidad = nominaMap.get(num_entidad) || {};

            const balancesDiciembre = entityBalances.filter(b => b.fecha_bce.startsWith('12-'));
            for (const bal of balancesDiciembre) {
                const descCuenta = (cuentasMap.get(bal.num_cuenta) || {}).descripcion_cuenta || 'No encontrada';
                const descPlan = descMap.get(bal.num_cuenta) || '';
                
                let rubro = '';
                if (bal.num_cuenta >= 500000 && bal.num_cuenta < 700000) {
                    rubro = rubroMap.get(bal.num_cuenta) || '';
                }

                decemberSummaryData.push([
                    infoEntidad.nombre_entidad || 'Desconocido',
                    num_entidad,
                    bal.fecha_bce, 
                    bal.num_cuenta,
                    descCuenta,
                    descPlan,
                    bal.saldo,
                    rubro,
                    null
                ]);
            }

            console.log(`Generating sheet for entity ${num_entidad}...`);
            const dataForSheet = prepareDataForSheet(entityBalances, cuentasMap, nominaMap, descMap, allMonths, indicesMap, num_entidad);
            if (dataForSheet.length <= 3) continue;
            
            let sheetName = `${String(num_entidad).padStart(5, '0')} - ${infoEntidad.nombre_corto || infoEntidad.nombre_entidad || ''}`.trim().substring(0, 31).replace(/[\\/*?[\]]/g, '');
            tocSheetData.push([sheetName, num_entidad, infoEntidad.nombre_entidad || '']);
            
            const worksheet = xlsx.utils.aoa_to_sheet(dataForSheet);

            worksheet['!views'] = [
                { state: 'frozen', xSplit: 3, ySplit: 3, topLeftCell: 'D4' }
            ];

            const range = xlsx.utils.decode_range(worksheet['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                if (!worksheet['!rows']) worksheet['!rows'] = [];
                if (R > 2) worksheet['!rows'][R] = { hpt: 12 }; 

                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_ref = xlsx.utils.encode_cell({ c: C, r: R });
                    const cell = worksheet[cell_ref];
                    if (!cell) continue;

                    cell.s = defaultCellStyle;
                    if (R === 0 && C === 1) {
                         cell.s = disclaimerStyle; 
                    } else if (R === 2) {
                         if (cell.v && cell.v.toString().endsWith('__Control_Incoherencia__')) {
                             cell.v = "Incoherencia en el AXI mostrado, porque el saldo en moneda constante a diciembre está en cero";
                             cell.s = { 
                                 ...headerStyle, 
                                 fill: { fgColor: { rgb: "FFFF00" } },
                                 font: { ...headerStyle.font, color: { rgb: "FF0000" } } 
                             };
                         } else if (cell.v && cell.v.toString().endsWith('__SEPARATOR__')) {
                             cell.v = "";
                             cell.s = { fill: { fgColor: { rgb: "FF0000" } } };
                         } else {
                             cell.s = headerStyle;
                         }
                    } else if (cell.v?.toString().startsWith('Observaciones') || cell.v?.toString().startsWith('- Es imprescindible')) {
                         cell.s = { ...disclaimerStyle, alignment: { horizontal: "centerAcross", vertical: "center" } };
                    } else if (cell.v?.toString().startsWith('- ')) {
                         cell.s = obsBodyStyle;
                    } else if (cell.v === 'Incoherencia') {
                         cell.s = errorTextStyle;
                    } else {
                         const originalHeader = dataForSheet[2][C];
                         if (originalHeader && originalHeader.endsWith('__SEPARATOR__')) {
                             cell.s = { fill: { fgColor: { rgb: "FF0000" } } };
                         } else if (cell.t === 'n') {
                             const descCellValue = worksheet[xlsx.utils.encode_cell({c: 3, r: R})]?.v || "";
                             const headerVal = worksheet[xlsx.utils.encode_cell({c: C, r: 2})]?.v || "";
                             const headerStr = String(headerVal).toLowerCase();
                             
                             if (descCellValue.startsWith("Total")) {
                                 cell.s = totalStyle;
                             } else if (descCellValue.startsWith("Subtotal")) {
                                 cell.s = subtotalStyle;
                             } else if (R === 1 && C > 4) { 
                                 cell.s = percent6Style; 
                             } else {
                                 if (headerStr.includes("cuenta") || headerStr.includes("entidad")) {
                                     cell.s = integerFormatStyle; 
                                 } else if (headerStr.includes("histórico") || headerStr.includes("axi")) {
                                     cell.s = decimalFormatStyle; 
                                 } else {
                                     cell.s = integerMoneyStyle; 
                                 }
                             }
                         }
                    }
                }
            }
            if (worksheet['A1']) worksheet['A1'].l = { Target: `#'${TOC_SHEET_NAME}'!A1`, Tooltip: `Ir a la hoja ${TOC_SHEET_NAME}` };
            
            const colWidths = [ { wch: 6 }, { wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 13 } ];
            
            allMonths.forEach((month, i) => { 
                const [mes, anio] = month.split('-');
                if (mes === '01' && i > 0) {
                    colWidths.push({ wch: 8 });  
                    colWidths.push({ wch: 12 }); 
                }
                colWidths.push({ wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }); 
                
                if(mes === '12') {
                    colWidths.push({ wch: 15 }); 
                    colWidths.push({ wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 6 }); 
                    colWidths.push({ wch: 3 }); 
                }
            });
            
            worksheet['!cols'] = colWidths;
            worksheet['!rows'][0] = { hpt: 17 }; 
            worksheet['!rows'][2] = { hpt: 60 }; 

            xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
        }

        if (decemberSummaryData.length > 1) { 
            console.log("Generating December Summary Sheet...");
            
            const summaryHeader = decemberSummaryData[0];
            const summaryRows = decemberSummaryData.slice(1);

            const totalsByYearAndEntityAndRubro = {};

            summaryRows.forEach(row => {
                const entityName = row[0]; // Nombre Entidad
                const dateStr = row[2]; 
                const year = dateStr ? dateStr.split('-')[1] : 'Unknown';
                const saldo = row[6] || 0; 
                const rubro = row[7];      
                
                if (rubro) {
                    if (!totalsByYearAndEntityAndRubro[year]) totalsByYearAndEntityAndRubro[year] = {};
                    if (!totalsByYearAndEntityAndRubro[year][entityName]) totalsByYearAndEntityAndRubro[year][entityName] = {};
                    if (!totalsByYearAndEntityAndRubro[year][entityName][rubro]) totalsByYearAndEntityAndRubro[year][entityName][rubro] = 0;
                    totalsByYearAndEntityAndRubro[year][entityName][rubro] += saldo;
                }
            });

            // CAMBIO: Ordenamiento solicitado (Entidad -> Fecha -> Cuenta)
            summaryRows.sort((a, b) => {
                // 1. Entidad (ID) - Índice 1
                if (a[1] !== b[1]) return a[1] - b[1];

                // 2. Fecha (Año) - Índice 2 ("12-2022")
                const yearA = parseInt(a[2].split('-')[1], 10);
                const yearB = parseInt(b[2].split('-')[1], 10);
                if (yearA !== yearB) return yearA - yearB;

                // 3. Cuenta (Número) - Índice 3
                return a[3] - b[3];
            });

            const finalSummaryData = [summaryHeader, ...summaryRows];
            const wsSummary = xlsx.utils.aoa_to_sheet(finalSummaryData);
            
            const totalsTable = [["Año", "Nombre Entidad", "Agrupado por Rubro de Exposión en el Estado de Resultados para permitir su rápido cotejo contra los EE.CC. de Publicación", "Total del Rubro en moneda constate"]];
            
            const sortedYears = Object.keys(totalsByYearAndEntityAndRubro).sort();

            sortedYears.forEach(year => {
                const entities = totalsByYearAndEntityAndRubro[year];
                const sortedEntities = Object.keys(entities).sort();

                sortedEntities.forEach(entity => {
                    const rubros = entities[entity];
                    let entityTotal = 0;
                    
                    Object.keys(rubros).sort().forEach(rubro => {
                        const total = rubros[rubro];
                        totalsTable.push([year, entity, rubro, total]);
                        entityTotal += total;
                    });
                    
                    // Resultado por Entidad en ese Año
                    totalsTable.push([year, entity, "Resultado del Ejercicio", entityTotal]);
                    totalsTable.push([null, null, null, null]); 
                });
            });

            xlsx.utils.sheet_add_aoa(wsSummary, totalsTable, { origin: "J1" });

            const summarySheetName = "Resumen Diciembre";
            const summaryRange = xlsx.utils.decode_range(wsSummary['!ref']);
            
            wsSummary['!cols'] = [
                {wch: 20}, {wch: 6}, {wch: 10}, {wch: 8}, {wch: 30}, {wch: 13}, {wch: 15}, {wch: 25}, 
                {wch: 2}, 
                {wch: 6}, {wch: 30}, {wch: 40}, {wch: 15} 
            ];
            if (!wsSummary['!rows']) wsSummary['!rows'] = [];
            wsSummary['!rows'][0] = { hpt: 30 };

            wsSummary['!views'] = [
                { state: 'frozen', xSplit: 4, ySplit: 1, topLeftCell: 'E2', activeCell: 'E2' }
            ];

            for (let R = summaryRange.s.r; R <= summaryRange.e.r; ++R) {
                for (let C = summaryRange.s.c; C <= summaryRange.e.c; ++C) {
                    const cell_ref = xlsx.utils.encode_cell({ c: C, r: R });
                    const cell = wsSummary[cell_ref];
                    if (!cell) continue;

                    if (R === 0) {
                        if (C === 8) {
                            cell.s = { fill: { fgColor: { rgb: "FF0000" } } };
                        } else {
                            cell.s = headerStyle;
                        }
                    } else {
                        if (C <= 7) {
                            if (C === 6) { 
                                cell.s = integerMoneyStyle; 
                            } else if (C === 1 || C === 3) { 
                                cell.s = integerFormatStyle;
                            } else {
                                cell.s = defaultCellStyle;
                            }
                        }
                        // TABLA LATERAL (Indices J=9, K=10, L=11, M=12)
                        else if (C === 9) { // Año
                             cell.s = integerFormatStyle;
                        }
                        else if (C === 10) { // Nombre Entidad
                             cell.s = defaultCellStyle;
                        }
                        else if (C === 11) { // Rubro Label
                            if (cell.v === "Resultado del Ejercicio") {
                                cell.s = totalStyle;
                            } else {
                                cell.s = defaultCellStyle;
                            }
                        } else if (C === 12) { // Total Monto
                            const labelCell = wsSummary[xlsx.utils.encode_cell({c: 11, r: R})];
                            if (labelCell && labelCell.v === "Resultado del Ejercicio") {
                                cell.s = totalStyle;
                            } else {
                                cell.s = integerMoneyStyle;
                            }
                        }
                    }
                }
            }
            xlsx.utils.book_append_sheet(workbook, wsSummary, summarySheetName);
            tocSheetData.push([summarySheetName, "Global", "Resumen de Saldos a Diciembre"]);
        }

        // --- Hoja Plan de Cuentas Completo ---
        if (planRawData && planRawData.length > 0) {
            console.log("Generating Full Plan Sheet...");
            const planSheetName = "Plan de Cuentas BCRA";
            const wsPlan = xlsx.utils.aoa_to_sheet(planRawData);
            
            const planRange = xlsx.utils.decode_range(wsPlan['!ref']);
            wsPlan['!cols'] = [{wch: 10}, {wch: 60}, {wch: 30}, {wch: 20}];
            
            for (let R = planRange.s.r; R <= planRange.e.r; ++R) {
                for (let C = planRange.s.c; C <= planRange.e.c; ++C) {
                    const cell_ref = xlsx.utils.encode_cell({ c: C, r: R });
                    const cell = wsPlan[cell_ref];
                    if (!cell) continue;

                    if (R === 0) {
                        cell.s = headerStyle;
                    } else {
                        cell.s = defaultCellStyle;
                    }
                }
            }
            
            xlsx.utils.book_append_sheet(workbook, wsPlan, planSheetName);
            tocSheetData.push([planSheetName, "Referencia", "Plan de Cuentas Completo"]);
        }

        const tocWorksheet = xlsx.utils.aoa_to_sheet(tocSheetData);
        tocWorksheet['!cols'] = [{ wch: 35 }, { wch: 15 }, { wch: 50 }];
        tocSheetData.slice(1).forEach((row, index) => { const sheetName = row[0]; const cellAddress = `A${index + 2}`; if (tocWorksheet[cellAddress]) { tocWorksheet[cellAddress].l = { Target: `#'${sheetName}'!A1`, Tooltip: `Ir a la hoja ${sheetName}` }; tocWorksheet[cellAddress].s = { font: { ...defaultFont, color: { rgb: "0000FF" }, underline: true } }; } });
        xlsx.utils.book_append_sheet(workbook, tocWorksheet, TOC_SHEET_NAME);
        workbook.SheetNames.splice(workbook.SheetNames.indexOf(TOC_SHEET_NAME), 1);
        workbook.SheetNames.unshift(TOC_SHEET_NAME);
        
        const excelBuffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = now.getFullYear();
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');
        const timestamp = `${year}-${month}-${day}_${hours}-${minutes}-${seconds}`;
        const nombreEntidadFile = filtros.nombreEntidad || "Entidad_Desconocida";
        const nombreArchivo = `Reporte_Pivoteado_${nombreEntidadFile}_${filtros.balhistDesde}_a_${filtros.balhistHasta}_Fecha_Generacion_${timestamp}.xlsx`;
                
        res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
        res.setHeader('Content-Disposition', `attachment; filename="${nombreArchivo}"`);
        res.status(200).send(excelBuffer);
    } catch (processingError) {
        console.error("Error crítico:", processingError);
        res.status(500).send('Falló el proceso de la aplicación.');
    }
});

app.listen(PORT, () => { console.log(`Servidor corriendo en http://localhost:${PORT}`); });