// Referencias HTML
const formulario = document.getElementById('form-registro');
const contenedorTabla = document.getElementById('contenedor-tabla');

// Datos
let estudiantesAtrasados = JSON.parse(localStorage.getItem('estudiantes_atrasados')) || [];

// Inicializar Fecha
document.getElementById('fecha-actual').innerText = new Date().toLocaleDateString('es-ES', {
    weekday: 'short', day: 'numeric', month: 'short'
});

function guardar() {
    localStorage.setItem('estudiantes_atrasados', JSON.stringify(estudiantesAtrasados));
}

function calcularEstadisticas() {
    document.getElementById('stat-total').innerText = estudiantesAtrasados.length;
    
    if (estudiantesAtrasados.length > 0) {
        document.getElementById('stat-ultimo').innerText = estudiantesAtrasados[estudiantesAtrasados.length - 1].hora;

        // Calcular motivo más común
        const conteo = {};
        estudiantesAtrasados.forEach(e => conteo[e.motivo] = (conteo[e.motivo] || 0) + 1);
        const motivoMasComun = Object.keys(conteo).reduce((a, b) => conteo[a] > conteo[b] ? a : b);
        document.getElementById('stat-motivo').innerText = motivoMasComun;
    } else {
        document.getElementById('stat-ultimo').innerText = "--:--";
        document.getElementById('stat-motivo').innerText = "N/A";
    }
}

function renderizarTabla() {
    calcularEstadisticas();

    if (estudiantesAtrasados.length === 0) {
        contenedorTabla.innerHTML = `
            <div class="text-center py-12 text-slate-500">
                <p class="text-4xl mb-2">📥</p>
                <p>No se han registrado atrasos el día de hoy.</p>
            </div>
        `;
        return;
    }

    let filas = estudiantesAtrasados.map((est, i) => `
        <tr class="border-b border-slate-700/50 hover:bg-slate-700/30 transition text-sm">
            <td class="p-3 font-mono text-slate-400">${i + 1}</td>
            <td class="p-3 font-semibold text-white">${est.nombre}</td>
            <td class="p-3 text-slate-300">${est.grado}</td>
            <td class="p-3">
                <span class="bg-slate-700 text-slate-300 text-xs px-2.5 py-1 rounded-md border border-slate-600">
                    ${est.motivo}
                </span>
            </td>
            <td class="p-3">
                <span class="bg-blue-500/20 text-blue-300 border border-blue-500/30 font-bold text-xs px-2.5 py-1 rounded-md">
                    ${est.hora}
                </span>
            </td>
        </tr>
    `).reverse().join('');

    contenedorTabla.innerHTML = `
        <div class="overflow-x-auto">
            <table class="w-full text-left">
                <thead class="bg-slate-900/50 text-slate-400 uppercase text-xs">
                    <tr>
                        <th class="p-3">#</th>
                        <th class="p-3">Estudiante</th>
                        <th class="p-3">Grado</th>
                        <th class="p-3">Motivo</th>
                        <th class="p-3">Hora</th>
                    </tr>
                </thead>
                <tbody>${filas}</tbody>
            </table>
        </div>
    `;
}

// Registro de formulario
formulario.addEventListener('submit', (e) => {
    e.preventDefault();

    const nombre = document.getElementById('nombre').value;
    const grado = document.getElementById('grado').value;
    const motivo = document.getElementById('motivo').value;
    const hora = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true });

    const nuevo = { nombre, grado, motivo, hora };
    estudiantesAtrasados.push(nuevo);
    guardar();
    formulario.reset();
    renderizarTabla();

    // Mostrar Pase de entrada modal
    mostrarModal(nuevo);
});

function mostrarModal(est) {
    document.getElementById('ticket-nombre').innerText = est.nombre;
    document.getElementById('ticket-grado').innerText = est.grado;
    document.getElementById('ticket-motivo').innerText = est.motivo;
    document.getElementById('ticket-hora').innerText = est.hora;
    document.getElementById('modal-pase').classList.remove('hidden');
}

function cerrarModal() {
    document.getElementById('modal-pase').classList.add('hidden');
}

function exportarCSV() {
    // 1. Obtener los datos actuales del arreglo o del almacenamiento
    let datos = [];
    if (typeof estudiantesAtrasados !== 'undefined' && estudiantesAtrasados.length > 0) {
        datos = estudiantesAtrasados;
    } else {
        datos = JSON.parse(localStorage.getItem('estudiantes_atrasados')) || [];
    }

    if (datos.length === 0) {
        alert("No hay registros almacenados para exportar.");
        return;
    }

    // 2. Definir estilos para Excel (Encabezado azul claro + bordes + centrado)
    const estiloEncabezado = {
        font: { bold: true, name: "Calibri", sz: 11 },
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: "DCE6F1" } }, // Color azul suave
        border: {
            top: { style: "thin", color: { rgb: "000000" } },
            bottom: { style: "thin", color: { rgb: "000000" } },
            left: { style: "thin", color: { rgb: "000000" } },
            right: { style: "thin", color: { rgb: "000000" } }
        }
    };

    const estiloCelda = {
        font: { name: "Calibri", sz: 11 },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
            top: { style: "thin", color: { rgb: "000000" } },
            bottom: { style: "thin", color: { rgb: "000000" } },
            left: { style: "thin", color: { rgb: "000000" } },
            right: { style: "thin", color: { rgb: "000000" } }
        }
    };

    // 3. Crear matriz de datos ordenados
    const encabezados = ["N°", "Estudiante", "Grado / Curso", "Motivo del Retraso", "Hora de Ingreso"];
    const filas = [
        encabezados.map(texto => ({ v: texto, s: estiloEncabezado }))
    ];

    datos.forEach((est, index) => {
        filas.push([
            { v: index + 1, s: estiloCelda },
            { v: est.nombre || est.estudiante || "", s: estiloCelda },
            { v: est.grado || "", s: estiloCelda },
            { v: est.motivo || "", s: estiloCelda },
            { v: est.hora || est.horaIngreso || "", s: estiloCelda }
        ]);
    });

    // 4. Crear hoja y definir anchos de columna exactos
    const hoja = XLSX.utils.aoa_to_sheet(filas);
    hoja['!cols'] = [
        { wch: 8 },  // N°
        { wch: 28 }, // Estudiante
        { wch: 22 }, // Grado / Curso
        { wch: 28 }, // Motivo del Retraso
        { wch: 20 }  // Hora de Ingreso
    ];

    // 5. Generar y descargar el archivo .xlsx
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Tardanzas");

    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(libro, `Registro_Tardanzas_${fecha}.xlsx`);
}
    
function limpiarRegistros() {
    if (confirm("¿Vaciar todos los registros?")) {
        estudiantesAtrasados = [];
        guardar();
        renderizarTabla();
    }
}

// Carga Inicial
renderizarTabla();
// ===== INICIO DE SESIÓN =====
const HASH_ACCESO = 'd9502f323eeec1973f827b45c665f4e5870aff9e65c22a5507fba39b258da81c';

async function calcularHash(texto) {
    const datos = new TextEncoder().encode(texto);
    const resultado = await crypto.subtle.digest('SHA-256', datos);
    return [...new Uint8Array(resultado)].map(b => b.toString(16).padStart(2, '0')).join('');
}

function mostrarApp() {
    document.getElementById('pantalla-login').classList.add('hidden');
    document.getElementById('app-principal').style.display = 'block';
}

document.getElementById('form-login').addEventListener('submit', async (e) => {
    e.preventDefault();
    const usuario = document.getElementById('login-usuario').value.trim();
    const clave = document.getElementById('login-clave').value;
    const hash = await calcularHash(usuario + ':' + clave);

    if (hash === HASH_ACCESO) {
        sessionStorage.setItem('sesion_activa', 'si');
        document.getElementById('login-error').innerText = '';
        mostrarApp();
    } else {
        document.getElementById('login-error').innerText = 'Usuario o contraseña incorrectos';
        document.getElementById('login-clave').value = '';
    }
});

function cerrarSesion() {
    sessionStorage.removeItem('sesion_activa');
    location.reload();
}

if (sessionStorage.getItem('sesion_activa') === 'si') mostrarApp();