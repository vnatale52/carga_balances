// frontend/scripts.js (Versión Corregida: Envío de nombre de Entidad al Backend)

document.addEventListener('DOMContentLoaded', async () => {
    const select = document.getElementById('entidadSelect');
    const statusDiv = document.getElementById('status');

    try {
        const response = await fetch("/api/entidades");
      
        if (!response.ok) throw new Error(`No se pudieron cargar las entidades (código: ${response.status})`);
        
        const entidades = await response.json();
        select.innerHTML = ''; 

        const allOption = document.createElement('option');
        allOption.value = "0"; 
        allOption.textContent = "00000 - Todas las Entidades (ignora otras selecciones; funciona solamente si se corre en localhost)";
        select.appendChild(allOption);

        entidades.sort((a, b) => a.num_entidad - b.num_entidad).forEach(entidad => {
            const option = document.createElement('option');
            option.value = entidad.num_entidad;
            const formattedNum = String(entidad.num_entidad).padStart(5, '0');
            option.textContent = `${formattedNum} - ${entidad.nombre_entidad}`;
            select.appendChild(option);
        });

        select.addEventListener('change', () => {
            const selectedOptions = Array.from(select.selectedOptions).map(opt => opt.value);
            if (selectedOptions.includes("0") && selectedOptions.length > 1) {
                Array.from(select.options).forEach(opt => {
                    opt.selected = (opt.value === "0");
                });
            }
        });

    } catch (error) {
        select.innerHTML = '<option value="" disabled selected>Error al cargar entidades</option>';
        statusDiv.textContent = `Error de red: ${error.message}`;
        statusDiv.style.color = 'red';
    }
});

document.getElementById('reportForm').addEventListener('submit', async function (event) {
    event.preventDefault();

    const statusDiv = document.getElementById('status');
    const select = document.getElementById('entidadSelect');
    const generateBtn = document.getElementById('generateReportBtn');
    
    const selectedEntidades = Array.from(select.selectedOptions).map(option => option.value);

    if (selectedEntidades.length === 0) {
        statusDiv.textContent = 'Error: Debes seleccionar al menos una entidad.';
        statusDiv.style.color = 'red';
        return;
    }

    const balhistDesde = document.getElementById('balhistDesdeInput').value;
    const balhistHasta = document.getElementById('balhistHastaInput').value;

    // --- INICIO CORRECCIÓN: Obtener y limpiar el nombre de la entidad ---
    let nombreEntidadClean = "Entidad_Desconocida";

    if (selectedEntidades.includes("0")) {
        nombreEntidadClean = "Todas_Las_Entidades";
    } else if (selectedEntidades.length > 1) {
        nombreEntidadClean = "Multiples_Entidades";
    } else {
        // Obtener el texto de la opción seleccionada (ej: "00011 - BANCO NACION")
        const selectedOption = Array.from(select.selectedOptions)[0];
        if (selectedOption) {
            let text = selectedOption.textContent;
            // Si tiene el formato "00000 - NOMBRE", nos quedamos con la parte del nombre
            if (text.includes(" - ")) {
                text = text.split(" - ")[1];
            }
            // Reemplazar espacios por guiones bajos y quitar caracteres raros para que sea seguro en el nombre del archivo
            nombreEntidadClean = text.trim().replace(/[^a-zA-Z0-9]/g, "_").replace(/_+/g, "_");
        }
    }
    // --- FIN CORRECCIÓN ---
    
    const filtros = {
        entidad: selectedEntidades,
        nombreEntidad: nombreEntidadClean, // Agregamos el nombre al payload
        balhistDesde,
        balhistHasta
    };

    generateBtn.disabled = true;
    generateBtn.textContent = 'Generando...';
    statusDiv.textContent = 'Procesando... Esto puede tardar varios minutos.';
    statusDiv.style.color = 'orange';
 
    try {
        const response = await fetch('/generate-report', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(filtros),
        });

        if (response.ok) {
            statusDiv.textContent = 'Proceso completado. Iniciando descarga...';
            statusDiv.style.color = 'green';

            const blob = await response.blob();
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            
            // --- LÓGICA PARA OBTENER EL NOMBRE DEL ARCHIVO ---
            let fileName = null;
            
            // 1. Intentar leer el header del servidor (Prioridad)
            const contentDisposition = response.headers.get('Content-Disposition');
            if (contentDisposition) {
                const match = contentDisposition.match(/filename="?([^"]+)"?/);
                if (match && match[1]) {
                    fileName = match[1];
                }
            }

            // 2. Fallback: Si no se pudo leer el header, generar nombre local
            if (!fileName) {
                const now = new Date();
                const year = now.getFullYear();
                const month = String(now.getMonth() + 1).padStart(2, '0');
                const day = String(now.getDate()).padStart(2, '0');
                const hours = String(now.getHours()).padStart(2, '0');
                const minutes = String(now.getMinutes()).padStart(2, '0');
                const seconds = String(now.getSeconds()).padStart(2, '0');
                const timestamp = `${year}-${month}-${day}_${hours}-${minutes}-${seconds}`;

                fileName = `Reporte_Pivoteado_${nombreEntidadClean}_${balhistDesde}_a_${balhistHasta}_${timestamp}.xlsx`;
            }

            link.download = fileName;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(link.href);

            statusDiv.textContent = '¡Reporte generado con éxito! Revisa tus descargas.';
            statusDiv.style.color = '#007bff';

            setTimeout(() => {
                statusDiv.textContent = '';
            }, 8000);

        } else {
            const errorText = await response.text();
            throw new Error(errorText || `Error del servidor: ${response.status}`);
        }

    } catch (error) {
        statusDiv.textContent = `Error: ${error.message}`;
        statusDiv.style.color = 'red';
        console.error('Detalle del error:', error);
    } finally {
        generateBtn.disabled = false;
        generateBtn.textContent = 'Generar Reporte Excel';
    }
});