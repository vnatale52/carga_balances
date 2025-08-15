// backend/server.js (Versión con optimización de memoria en todas las lecturas de archivos)

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

// --- FUNCIONES DE PROCESAMIENTO OPTIMIZADAS ---

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
        const buffer = fs.readFileSync(filePath); // OK mantener esto, los .xlsx se leen como buffer y suelen ser pequeños
        const workbook = xlsx.read(buffer, { type: 'buffer', cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
        const indicesMap = new Map();
        if (!jsonData || jsonData.length === 0) return indicesMap;
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
    const months = [];
    let currentDate = startDate;
    while (currentDate <= endDate) {
        const month = ('0' + (currentDate.getMonth() + 1)).slice(-2);
        const year = currentDate.getFullYear();
        months.push(`${month}-${year}`);
        currentDate.setMonth(currentDate.getMonth() + 1);
    }
    return months;
}

function prepareDataForSheet(balancesDeEstaEntidad, cuentasMap, nominaMap, allMonths, indicesMap) {
    // ... (esta función no necesita cambios)
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
    allMonths.forEach(month => {
        newHeaders.push(`${month} Saldo en moneda constante`);
        newHeaders.push(`${month} Saldo Histórico solo del mes`);
        newHeaders.push(`${month} Saldo Histórico acumulado al mes`);
        newHeaders.push(`${month} AXI mensual solo del mes`);
        newHeaders.push(`${month} AXI acumulado al mes`);
    });
    const firstRowContent = new Array(newHeaders.length).fill('');
    firstRowContent[0] = '<< Volver a la Tabla de Contenidos';
    const axiCoefficients = allMonths.map((month, i) => {
        if (month.startsWith('01-')) return 0;
        if (i === 0) return 0;
        const currentMonthIndex = indicesMap.get(month);
        const previousMonthIndex = indicesMap.get(allMonths[i - 1]);
        return (currentMonthIndex && previousMonthIndex) ? (currentMonthIndex / previousMonthIndex) - 1 : 0;
    });
    const axiRow = new Array(newHeaders.length).fill('');
    axiRow[3] = '% del Coeficiente AXI';
    allMonths.forEach((month, i) => {
        const colIndex = 4 + (i * 5) + 3;
        axiRow[colIndex] = axiCoefficients[i];
    });
    const dataForSheet = [firstRowContent, axiRow, newHeaders];
    const sortedCuentas = Object.keys(pivotedData).sort((a, b) => parseInt(a) - parseInt(b));
    for (const num_cuenta_str of sortedCuentas) {
        const cuentaData = pivotedData[num_cuenta_str];
        const num_cuenta = parseInt(num_cuenta_str);
        const isAdjustable = (num_cuenta > 500000 && num_cuenta < 700000);
        const rowObject = {
            'Entidad': infoEntidad.num_entidad,
            'Nombre Entidad': infoEntidad.nombre_entidad,
            'Cuenta': num_cuenta,
            'Descripción Cuenta': cuentaData.desc,
        };
        let saldoHistAcumuladoAnterior = 0;
        let axiAcumuladoAnterior = 0;
        let saldoMonedaConstanteAnterior = 0;
        allMonths.forEach((month, i) => {
            const saldoEnMonedaConstanteMes = cuentaData.saldos[month] || 0;
            const coefAXI = axiCoefficients[i];
            let axiMensualMes = 0;
            if (isAdjustable) {
                axiMensualMes = saldoMonedaConstanteAnterior * coefAXI;
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
        const rowAsArray = newHeaders.map(header => rowObject[header] !== undefined ? rowObject[header] : '');
        dataForSheet.push(rowAsArray);
    }
    return dataForSheet;
}

// --- ENDPOINTS ---
app.get('/api/entidades', async (req, res) => { // <-- Hacemos este endpoint async también
    try {
        const nominaPath = path.join(__dirname, '../frontend/data/nomina.txt');
        if (!fs.existsSync(nominaPath)) return res.status(404).json({ message: 'Archivo nomina.txt no encontrado.' });
        // Usamos la nueva función async
        const nominaData = await procesarNomina(nominaPath);
        res.json(nominaData);
    } catch (error) {
        res.status(500).json({ message: 'Error interno al leer entidades.' });
    }
});

app.post('/generate-report', async (req, res) => {
    try {
        const filtros = req.body;
        const filePaths = {
            balhist: path.join(__dirname, '../frontend/data/balhist.txt'),
            cuentas: path.join(__dirname, '../frontend/data/cuentas.txt'),
            nomina: path.join(__dirname, '../frontend/data/nomina.txt'),
            indices: path.join(__dirname, '../frontend/data/indices.xlsx')
        };
        for (const key in filePaths) {
            if (!fs.existsSync(filePaths[key])) {
                return res.status(404).send(`Error: El archivo ${path.basename(filePaths[key])} no se encuentra.`);
            }
        }
        
        // Se llaman a todas las funciones de lectura de archivos con await
        const datosBalhistFiltrados = await procesarBalhist(filePaths.balhist, filtros);
        if (datosBalhistFiltrados.length === 0) {
            return res.status(404).send('No se encontraron registros de balance con los filtros seleccionados.');
        }

        const indicesMap = procesarIndices(filePaths.indices);
        if (indicesMap.size === 0) {
            return res.status(404).send('No se pudieron leer los datos del archivo indices.xlsx.');
        }

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
        for (const num_entidad of sortedEntityNumbers) {
            const dataForSheet = prepareDataForSheet(balancesPorEntidad.get(num_entidad), cuentasMap, nominaMap, allMonths, indicesMap);
            
            if (dataForSheet.length > 3) {
                const infoEntidad = nominaMap.get(num_entidad) || {};
                const formattedNum = String(num_entidad).padStart(5, '0');
                let sheetName = `${formattedNum} - ${infoEntidad.nombre_corto || infoEntidad.nombre_entidad || ''}`.trim();
                sheetName = sheetName.substring(0, 31).replace(/[\\/*?[\]]/g, '');
                tocSheetData.push([sheetName, num_entidad, infoEntidad.nombre_entidad || '']);
                const worksheet = xlsx.utils.aoa_to_sheet(dataForSheet);
                worksheet['A1'].l = { Target: `#'${TOC_SHEET_NAME}'!A1`, Tooltip: `Ir a ${TOC_SHEET_NAME}` };

                for (let i = 0; i < allMonths.length; ++i) {
                    const cellAddress = xlsx.utils.encode_cell({ r: 1, c: 4 + (i * 5) + 3 });
                    if (worksheet[cellAddress] && typeof worksheet[cellAddress].v === 'number') {
                        worksheet[cellAddress].t = 'n';
                        worksheet[cellAddress].z = '0.00%';
                    }
                }
                
                const colWidths = [ { wch: 15 }, { wch: 30 }, { wch: 10 }, { wch: 45 } ];
                allMonths.forEach(() => {
                    colWidths.push({ wch: 25 }); colWidths.push({ wch: 25 }); colWidths.push({ wch: 25 }); colWidths.push({ wch: 25 }); colWidths.push({ wch: 25 });
                });
                worksheet['!cols'] = colWidths;
                sheetsToAppend.push({ name: sheetName, sheet: worksheet });
            }
        }
        if (sheetsToAppend.length === 0) {
            return res.status(404).send('No se generaron hojas. Verifique los datos de origen y filtros.');
        }
        const tocWorksheet = xlsx.utils.aoa_to_sheet(tocSheetData);
        sheetsToAppend.forEach((sheetInfo, index) => {
            const cellAddress = `A${index + 2}`;
            if (tocWorksheet[cellAddress]) {
                tocWorksheet[cellAddress].l = { Target: `#'${sheetInfo.name}'!A1`, Tooltip: `Ir a la hoja ${sheetInfo.name}` };
            }
        });
        tocWorksheet['!cols'] = [{ wch: 35 }, { wch: 15 }, { wch: 50 }];
        xlsx.utils.book_append_sheet(workbook, tocWorksheet, TOC_SHEET_NAME);
        sheetsToAppend.forEach(sheetInfo => {
            xlsx.utils.book_append_sheet(workbook, sheetInfo.sheet, sheetInfo.name);
        });
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