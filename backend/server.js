// backend/server.js (Versión Corregida y Mejorada)

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

// --- MIDDLEWARE ---
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../frontend')));

// --- FUNCIONES DE PROCESAMIENTO ---
async function procesarCuentas(filePath) { const cuentas = []; const fileStream = fs.createReadStream(filePath, { encoding: 'latin1' }); const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity }); for await (const linea of rl) { if (linea.trim() === '') continue; const [numCuenta, descripcion] = linea.split('\t'); if (!numCuenta || !descripcion) continue; cuentas.push({ num_cuenta: parseInt(numCuenta.replace(/"/g, ''), 10), descripcion_cuenta: descripcion.replace(/"/g, '').trim() }); } return new Map(cuentas.map(c => [c.num_cuenta, c])); }
async function procesarNomina(filePath) { const nomina = []; const fileStream = fs.createReadStream(filePath, { encoding: 'latin1' }); const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity }); for await (const linea of rl) { if (linea.trim() === '') continue; const [numEntidad, nombreEntidad, nombreCorto] = linea.split('\t'); if (!numEntidad || !nombreEntidad) continue; nomina.push({ num_entidad: parseInt(numEntidad.replace(/"/g, ''), 10), nombre_entidad: nombreEntidad.replace(/"/g, '').trim(), nombre_corto: (nombreCorto || '').replace(/"/g, '').trim() }); } return new Map(nomina.map(e => [e.num_entidad, e])); }
function procesarIndices(filePath) { try { const buffer = fs.readFileSync(filePath); const workbook = xlsx.read(buffer, { type: 'buffer', cellDates: true }); const sheetName = workbook.SheetNames[0]; const worksheet = workbook.Sheets[sheetName]; const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 }); const indicesMap = new Map(); for (const row of jsonData) { if (!row || row.length < 2) continue; const fechaValue = row[0]; const indiceValue = row[1]; if (!indiceValue || !(fechaValue instanceof Date) || isNaN(fechaValue)) continue; const anio = fechaValue.getFullYear(); const mes = ('0' + (fechaValue.getMonth() + 1)).slice(-2); const fechaFormatoIndice = `${mes}-${anio}`; const indiceStr = String(indiceValue).rep