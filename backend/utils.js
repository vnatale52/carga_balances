import fs from 'fs';
import xlsx from 'xlsx';
import readline from 'readline';
// --- FUNCIONES DE PROCESAMIENTO (Sin cambios en su lógica) ---
export async function procesarCuentas(filePath) {
    const cuentas = [];
    const fileStream = fs.createReadStream(filePath, { encoding: 'latin1' });
    const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity });
    for await (const linea of rl) {
        if (linea.trim() === '') continue;
        const [numCuenta, descripcion] = linea.split('\t');
        if (!numCuenta || !descripcion) continue;
        cuentas.push({
            num_cuenta: parseInt(numCuenta.replace(/"/g, ''), 10),
            descripcion_cuenta: descripcion.replace(/"/g, '').trim()
        });
    }
    return new Map(cuentas.map(c => [c.num_cuenta, c]));
}

export async function procesarNomina(filePath) {
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
    return new Map(nomina.map(e => [e.num_entidad, e]));
}

export function procesarIndices(filePath) {
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

export function getMonthsInRange(start, end) {
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

// --- LÓGICA DE PREPARACIÓN DE DATOS (Sin cambios en su lógica) ---
export function prepareDataForSheet(balancesDeEstaEntidad, cuentasMap, nominaMap, allMonths, indicesMap, num_entidad) {
    if (!balancesDeEstaEntidad || balancesDeEstaEntidad.length === 0) return [];
    const infoEntidad = nominaMap.get(num_entidad) || { nombre_entidad: 'Desconocido', num_entidad };
    const pivotedData = {};

    for (const balance of balancesDeEstaEntidad) {
        if (!pivotedData[balance.num_cuenta]){
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

    const firstRowContent = new Array(newHeaders.length).fill(null);
    firstRowContent[0] = '<<== Volver a la TOC';
    firstRowContent[1] = "Formatea esta hoja a tu gusto. Cifras expresadas en miles de pesos argentinos. Elaborado en base a información publicada por el B.C.R.A y al Indice-FACPCE-Res.-JG-539-18.   A los fines específicos de esta aplicación, el ajuste por inflación está calculado – únicamente – para las cuentas de resultados, es decir, no está calculado también para los rubros no monetarios de las cuentas patrimoniales (por ejemplo, Bienes de Uso, Intangibles y cuentas del Patrimonio Neto).";

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

    const cuentasDeResultadosKeys = cuentasKeys
                                        .filter(c => c >= 500000 && c < 700000)
                                        .sort((a, b) => a - b);
    const otrasCuentasKeys = cuentasKeys
                                .filter(c => c < 500000 || c >= 700000)
                                .sort((a, b) => a - b);
    
    const processAccountRow = (num_cuenta) => {
        const cuentaData = pivotedData[num_cuenta];
        const isAdjustable = (num_cuenta >= 500000 && num_cuenta < 700000);
        const rowObject = {
            'Entidad': infoEntidad.num_entidad,
            'Nombre Entidad': infoEntidad.nombre_entidad,
            'Cuenta': num_cuenta,
            'Descripción Cuenta': cuentaData.desc
        };
        let saldoHistAcumuladoAnterior = 0, axiAcumuladoAnterior = 0, saldoMonedaConstanteAnterior = 0;
        
        allMonths.forEach((month, i) => {
            const saldoEnMonedaConstanteMes = cuentaData.saldos[month] || 0;
            let axiMensualMes = isAdjustable ? saldoMonedaConstanteAnterior * axiCoefficients[i] : 0;
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
    const emptyRow = new Array(newHeaders.length).fill(null);
    dataForSheet.push(emptyRow, emptyRow);
    
    dataForSheet.push(['Observaciones, comentarios y errores a corregir al 02-09-2025 :']);
    dataForSheet.push([ 'MUY IMPORTANTE: Se debe elegir sólo un año completo y sólo uno, desde enero a diciembre del mismo año. Si has eligido más de un año, para las cuentas que han tenido saldo sólo en el primer año seleccionado y que en el año subsiguiente ya han dejado de tener saldo, la aplicación arrastrará un saldo incorrecto para los años en que han dejado de tener movimiento. También arrastra, incorrectamente, al año subsiguiente, el ajuste por inflación acumulado en el año inmediato anterior (porque no tiene definida la "refundición de cuentas"). Falta corregir tales errores en la lógica del programa. Seleccionando sólo un año calendario, tales errores NO se producirán, en absoluto.']);
    dataForSheet.push(['Para las cuentas que comienzan con 62, que son las cuentas dónde se expone el resultado monetario generado por el ajuste por inflación, estas cuentas - debido a su particularidad - no tienen saldo histórico, pero la aplicación muestra - incorrectamente - un saldo histórico. Falta corregir dicho error en la lógica del programa.']);
    dataForSheet.push(['Posibles causas que generan diferencias entre el Ajuste por Inflación (AXI) calculado en forma automática por esta app, con respecto al AXI real contabilizado por el banco. Estas causas inciden en la anticuación de partidas realizada por esta app para el cálculo del AXI:']);
    dataForSheet.push(['- Ajustes contables con fecha valor (que sería la principal causa), realizados a posteriori del cierre de la presentación al BCRA del respectivo balance mensual TXT y, por ende, que no hayan impactado realmente en el balance presentado ante el BCRA (pero en este caso el banco debiera haber realizado una nueva presentación ante el BCRA, rectificando el anterior balance).']);
    dataForSheet.push(['- En los casos en que el INDEC hubiere, a posteriori, rectificado o corregido o publicado un nuevo IPIM (y el banco hubiere utilizado el IPIM "provisorio" anteriormente publicado), ello podría generar diferencia en el AXI (debido a que esta app toma como dato para el cálculo del AXI, el balance TXT en moneda constante publicado por el BCRA).']);
    dataForSheet.push(['- Para las cuentas de ingresos cuyas descripciones comiencen con "Resultado por", en los casos en que el saldo mensual de tales cuentas de ingresos quede invertido (debido a la volatilidad de las cotizaciones), dicho saldo, por expresa norma del BCRA, debe ser reclasificado  a su correspondiente cuenta de egresos (por ejemplo, Resultado de Títulos ...) . En este caso, se produce una diferencia en el AXI calculado por esta aplicación, con respecto al AXI realmente contabilizado por el banco (pero debiera compensarse con la diferencia, a su vez, generada en la cuenta de destino de dicha reclasificación). Este tipo de reclasificaciones pueden producirse varias veces para la misma cuenta y dentro de un mismo ejercicio.']);
    dataForSheet.push(['- En cualquier reclasificación contable de cuentas de resultados, desde una cuenta de resultados a otra de resultados, dicha reclasificación debiera realizarse, reclasificándose - separadamente - por un lado, el saldo histórico y,  por otro lado,  el saldo del AXI. Si así no se hiciere, ello afectaría la anticuación de partidas realizada por esta app, que parte, simplemente, del saldo según Balance TXT publicado por el BCRA.']);
    dataForSheet.push(['- Para que el AXI calculado por esta aplicación coincida con el AXI real, contabilizado por el Banco,  debe definirse en esta aplicación, como rango de fechas, necesariamente, desde Enero a Diciembre. Si no fuera así, el AXI calculado por esta aplicación sería incompleto (debido a que no abarca el ejercicio completo).']);   
    dataForSheet.push(['- El total de diferencias que surjan al cierre de cada ejercicio contable (Diciembre),  entre  a) el  total del AXI calculado por esta aplicación para las cuentas de resultados, con respecto a  b) el  total del AXI real contabilizado por el Banco,  coincidirá, a su vez, con  c) el total del saldo histórico acumulado calculado por esta aplicación, con respecto a  d) el total histórico real contabilizado por el banco (ello es debido a la lógica matemática implementada en esta app, es decir, la diferencia que surge en AXI, se compensa en el Histórico y viceversa).']);  
    dataForSheet.push(['Causa real de diferencias: esta app calcula el AXI (mediante "ingeniería matemática inversa"), partiendo del saldo en moneda constante, expresado en el miles de $, mientras que el banco realmente calcula el AXI partiendo del saldo histórico en CIFRAS COMPLETAS, lo cual es una fuente de pequeñas diferencias. Diferencia máxima estimada anual por simple redondeo a miles de $ : 500 (rendondeo) por 12 meses, igual a 6000 (en cifras completas), para cada cuenta contable de resultados.']); 
    dataForSheet.push(['Dado que el server sólo permite subir archivos de datos hasta un cierto límite máximo de tamaño, el tamaño del archivo de balances txt ha sido reducido, estándo sólo disponibles los balances desde Enero 2022 en adelante. Si necesitaras años anteriores, solo tienes que avisarle a Vincenzo y él - volentieri - proveerá.']);
    dataForSheet.push(['Para cualquier comentario, sugerencia o indicación de un posible error, contacta a Vincenzo  en vnatale52@gmail.com.  Saludos ... and happy coding and calculating ...']);
    return dataForSheet;
}