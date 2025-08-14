// frontend/scripts.js

document.addEventListener('DOMContentLoaded', async () => {
    const select = document.getElementById('entidadSelect');
    const statusDiv = document.getElementById('status');

    try {
        const response = await fetch("https://carga-balances.onrender.com/data");   //   ✅ para Render quitado "api" "https://carga-balances.onrender.com/api/data"
      
        if (!response.ok) throw new Error('No se pudieron cargar las entidades.');
        
        const entidades = await response.json();
        select.innerHTML = ''; 

        // Añadir la opción "Todas las Entidades"
        const allOption = document.createElement('option');
        allOption.value = "0"; 
        allOption.textContent = "00000 - Todas las Entidades (ignora otras selecciones)";
        select.appendChild(allOption);

        entidades.sort((a, b) => a.num_entidad - b.num_entidad).forEach(entidad => {
            const option = document.createElement('option');
            option.value = entidad.num_entidad;
            const formattedNum = String(entidad.num_entidad).padStart(5, '0');
            option.textContent = `${formattedNum} - ${entidad.nombre_entidad}`;
            select.appendChild(option);
        });

        // Evento para gestionar la selección de "Todas las entidades"
        select.addEventListener('change', () => {
            const selectedOptions = Array.from(select.selectedOptions).map(opt => opt.value);
            if (selectedOptions.includes("0") && selectedOptions.length > 1) {
                // Si "Todas" está seleccionada, deseleccionar las demás
                Array.from(select.options).forEach(opt => {
                    opt.selected = (opt.value === "0");
                });
            }
        });

    } catch (error) {
        select.innerHTML = '<option value="" disabled selected>Error al cargar entidades</option>';
        statusDiv.textContent = `Error: ${error.message}`;
        statusDiv.style.color = 'red';
    }
});

document.getElementById('reportForm').addEventListener('submit', async function (event) {
    event.preventDefault();

    const statusDiv = document.getElementById('status');
    const select = document.getElementById('entidadSelect');

    // --- MODIFICADO: Obtener todas las opciones seleccionadas ---
    const selectedEntidades = Array.from(select.selectedOptions).map(option => option.value);

    if (selectedEntidades.length === 0) {
        statusDiv.textContent = 'Error: Debes seleccionar al menos una entidad.';
        statusDiv.style.color = 'red';
        return;
    }

    const balhistDesde = document.getElementById('balhistDesdeInput').value;
    const balhistHasta = document.getElementById('balhistHastaInput').value;
    
    // El backend ahora recibirá un array de entidades
    const filtros = {
        entidad: selectedEntidades, 
        balhistDesde,
        balhistHasta,
        indicesDesde: balhistDesde,
        indicesHasta: balhistHasta
    };

    statusDiv.textContent = 'Procesando...';
    statusDiv.style.color = 'orange';
 
    try {
        const response = await fetch('http://localhost:3000/generate-report', {
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
            
            // --- Nombre de archivo dinámico mejorado ---
            let nombreEntidad;
            if (selectedEntidades.includes('0') || selectedEntidades.length > 5) {
                nombreEntidad = "Multiples_Entidades";
            } else if (selectedEntidades.length > 1) {
                nombreEntidad = `Entidades_${selectedEntidades.join('_')}`;
            } else {
                nombreEntidad = `Entidad_${selectedEntidades[0]}`;
            }
            if (selectedEntidades.includes('0')) nombreEntidad = "Todas_Entidades";

            link.download = `Reporte_Pivoteado_${nombreEntidad}_${balhistDesde}_a_${balhistHasta}.xlsx`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(link.href);

        } else {
            const errorText = await response.text();
            throw new Error(errorText || `Error del servidor: ${response.status}`);
        }

    } catch (error) {
        statusDiv.textContent = `Error: ${error.message}`;
        statusDiv.style.color = 'red';
        console.error('Detalle del error:', error);
    }

});

