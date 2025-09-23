// backend/server.js (Versión Final con Formatos de Número y Fuente Específicos)

import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import xlsx from 'xlsx';
import readline from 'readline';
import {
    procesarCuentas,
    procesarNomina,
    procesarIndices,
    getMonthsInRange,
    prepareDataForSheet
} from './utils.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const FRONTEND_PATH = '../frontend/data';

const app = express();
const PORT = process.env.PORT || 3000;

// --- MIDDLEWARE ---
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../frontend')));

// --- ENDPOINTS ---
app.get('/api/entidades', async (req, res) => {
    try {
        const nominaPath = path.join(__dirname, '../frontend/data/nomina.txt');
        if (!fs.existsSync(nominaPath)) {
            return res.status(404).json({ message: 'Archivo nomina.txt no encontrado.' });
        }
        const nominaMap = await procesarNomina(nominaPath);
        res.json(Array.from(nominaMap.values()));
    } catch (error) {
        console.log(error)
        res.status(500).json({ message: 'Error interno al leer entidades.' });
    }
});

app.post('/generate-report', async (req, res) => {
    try {
        console.log("Report generation started...");
        const filtros = req.body;
        const filePaths = {
            balhist: path.join(__dirname, `${FRONTEND_PATH}/balhist.txt`),
            cuentas: path.join(__dirname, `${FRONTEND_PATH}/cuentas.txt`),
            nomina: path.join(__dirname, `${FRONTEND_PATH}/nomina.txt`),
            indices: path.join(__dirname, `${FRONTEND_PATH}/indices.xlsx`)
        };
        for (const key in filePaths) {
            if (!fs.existsSync(filePaths[key])) {
                return res.status(404).send(`Error: El archivo ${path.basename(filePaths[key])} no se encuentra.`);
            }
        }
        
        console.log("Loading lookup data (cuentas, nomina, indices)...");
        const [cuentasMap, nominaMap, indicesMap] = await Promise.all([
             procesarCuentas(filePaths.cuentas),
             procesarNomina(filePaths.nomina),
             Promise.resolve(procesarIndices(filePaths.indices))
        ]);
        
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
                balancesPorEntidad.get(entidadActual).push({
                    fecha_bce: `${mes}-${anio}`,
                    num_cuenta: parseInt(numCuentaStr.replace(/"/g, ''), 10),
                    saldo: parseInt(saldoStr.trim(), 10)
                });
            }
        }
        
        console.log(`Finished processing balhist.txt. Found data for ${balancesPorEntidad.size} entities.`);
        if (balancesPorEntidad.size === 0) {
            return res.status(404).send('No se encontraron registros de balance con los filtros seleccionados.');
        }
        
        const sortedEntityNumbers = Array.from(balancesPorEntidad.keys()).sort((a, b) => a - b);

        // =================================================================================
        // INICIO: DEFINICIÓN DE ESTILOS DE EXCEL (Ajustes Finales)
        // =================================================================================
        
        // REQUERIMIENTO 4: Reduce el tamaño del font en cada página.
        const defaultFont = { name: "Arial", sz: 9 }; 
        
        const allBorders = {
            top: { style: "thin", color: { auto: 1 } },
            bottom: { style: "thin", color: { auto: 1 } },
            left: { style: "thin", color: { auto: 1 } },
            right: { style: "thin", color: { auto: 1 } },
        };
        
        // REQUERIMIENTO 2: -245457 sea mostrado como -245.457,00
        const numberFormatWithSeparators = '#.##0,00';
        
        // REQUERIMIENTO 1: 0,0469... sea mostrado como 4,694% 
        const percentFormatWith3Decimals = '0,000%'; 

        const headerStyle = {
            font: { ...defaultFont, bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "4F81BD" } },
            alignment: { horizontal: "center", vertical: "center", wrapText: true },
            border: allBorders
        };
        const totalStyle = {
            font: { ...defaultFont, bold: true },
            numFmt: numberFormatWithSeparators,
            fill: { fgColor: { rgb: "FFFF00" } },
            border: allBorders
        };
        const subtotalStyle = {
            font: { ...defaultFont, bold: true, italic: true },
            numFmt: numberFormatWithSeparators,
            fill: { fgColor: { rgb: "D3D3D3" } },
            border: allBorders
        };
        const defaultCellStyle = { font: defaultFont, border: allBorders };
        const integerFormatStyle = { ...defaultCellStyle, numFmt: "0" };
        const decimalFormatStyle = { ...defaultCellStyle, numFmt: numberFormatWithSeparators };
        const percentFormatStyle = { ...defaultCellStyle, numFmt: percentFormatWith3Decimals };
        
        const obsTitleStyle = { font: { ...defaultFont, sz: 10, bold: true } };
        const obsBodyStyle = { font: defaultFont, alignment: { wrapText: true, vertical: "top" } };
        const disclaimerStyle = {
            font: { ...defaultFont, sz: 8, italic: true },
            alignment: { wrapText: true, vertical: "center" }
        };
        // =================================================================================
        // FIN: DEFINICIÓN DE ESTILOS DE EXCEL
        // =================================================================================

        for (const num_entidad of sortedEntityNumbers) {
            console.log(`Generating sheet for entity ${num_entidad}...`);
            const entityBalances = balancesPorEntidad.get(num_entidad);
            const dataForSheet = prepareDataForSheet(
                entityBalances,
                cuentasMap,
                nominaMap,
                allMonths,
                indicesMap,
                num_entidad
            );
            if (dataForSheet.length <= 3) { continue; }
            
            const infoEntidad = nominaMap.get(num_entidad) || {};
            const nombre = infoEntidad.nombre_corto || infoEntidad.nombre_entidad || '';
            let sheetName = `${String(num_entidad).padStart(5, '0')} - ${nombre}`
                                .trim()
                                .substring(0, 31)
                                .replace(/[\\/*?[\]]/g, '');

            tocSheetData.push([sheetName, num_entidad, infoEntidad.nombre_entidad || '']);
            
            const worksheet = xlsx.utils.aoa_to_sheet(dataForSheet);
            
            const range = xlsx.utils.decode_range(worksheet['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                if (!worksheet['!rows']){
                    worksheet['!rows'] = [];
                }

                if (R > 2){
                    worksheet['!rows'][R] = { hpt: 12 }; 
                }

                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_ref = xlsx.utils.encode_cell({ c: C, r: R });
                    const cell = worksheet[cell_ref];
                    if (!cell) continue;

                    cell.s = defaultCellStyle;

                    if (R === 0 && C === 1) {
                        cell.s = disclaimerStyle;
                    } else if (R === 2) {
                        cell.s = headerStyle;
                    } else if (cell.v?.toString().startsWith('Observaciones')) {
                         cell.s = obsTitleStyle;
                    } else if (
                        cell.v?.toString().startsWith('Posibles causas') ||
                        cell.v?.toString().startsWith('- ') ||
                        cell.v?.toString().startsWith('Para cualquier')
                    ) {
                         cell.s = obsBodyStyle;
                    } else if (cell.t === 'n') {
                         const descCellValue = worksheet[xlsx.utils.encode_cell({c: 3, r: R})]?.v || "";
                         
                         if (descCellValue.startsWith("Total")) {
                             cell.s = totalStyle;
                         } else if (descCellValue.startsWith("Subtotal")) {
                             cell.s = subtotalStyle;
                         } else if (R === 1 && C > 3) {
                             cell.s = percentFormatStyle;
                         } else if (C === 0 || C === 2) {
                            cell.s = integerFormatStyle;
                         } else {
                            cell.s = decimalFormatStyle;
                         }
                    }
                }
            }

            if (worksheet['A1']) {
                worksheet['A1'].l = { 
                    Target: `#'${TOC_SHEET_NAME}'!A1`,
                    Tooltip: `Ir a la hoja ${TOC_SHEET_NAME}`
                };
            }
            
            const obsStartRow = dataForSheet.findIndex(row => {
                return (typeof row[0] === 'string') && row[0].startsWith('Observaciones:');
            });

            if (obsStartRow !== -1) {
                if (!worksheet['!merges']){
                    worksheet['!merges'] = [];
                }
                for (let R = obsStartRow; R < dataForSheet.length; R++) {
                    worksheet['!merges'].push({
                        s: { r: R, c: 0 },
                        e: { r: R, c: 8 }
                    });
                }
            }
            
            const colWidths = [
                { wch: 10 },
                { wch: 30 },
                { wch: 12 },
                { wch: 45 }
            ];

            allMonths.forEach(() => { 
                colWidths.push(
                    { wch: 18 },
                    { wch: 18 },
                    { wch: 18 },
                    { wch: 18 },
                    { wch: 18 }
                ); 
            });
            worksheet['!cols'] = colWidths;
            
            if (!worksheet['!merges']){
                worksheet['!merges'] = [];
            }
            worksheet['!merges'].push({
                s: { r: 0, c: 1 },
                e: { r: 0, c: 8 }
            });
            
            worksheet['!rows'][0] = { hpt: 17 }; 
            worksheet['!rows'][2] = { hpt: 22 }; 
            
            xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
        }

        console.log("All sheets generated. Finalizing workbook...");
        const tocWorksheet = xlsx.utils.aoa_to_sheet(tocSheetData);
        tocWorksheet['!cols'] = [{ wch: 35 }, { wch: 15 }, { wch: 50 }];

        tocSheetData.slice(1).forEach((row, index) => {
            const sheetName = row[0];
            const cellAddress = `A${index + 2}`;
            if (tocWorksheet[cellAddress]) {
                tocWorksheet[cellAddress].l = {
                    Target: `#'${sheetName}'!A1`,
                    Tooltip: `Ir a la hoja ${sheetName}`
                };
                tocWorksheet[cellAddress].s = {
                    font: { 
                        ...defaultFont, color: { rgb: "0000FF" },
                        underline: true
                    }
                };
            }
        });
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










