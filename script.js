// script.js — versión corregida y estable

let inventario = JSON.parse(localStorage.getItem('inventario')) || [];
let codigoAFila = {};
let excelInicialCargado = false;
let quaggaIniciado = false;

// ---------- CARGA EXCEL INICIAL ---------- ESTO ES OPCIONAL; PARA INVENTARIOS CORTOS CON EXCEL YA AGREGADO PARA RECONOCIMIENTO DE CODIGOS; CRUCE AUTOMATICO
async function cargarExcelInicial() {
    try {
        console.log('');
        const url = '';
        const response = await fetch(url);
        console.log('Respuesta del fetch:', response.status);
        if (!response.ok) {
            throw new Error(`Error HTTP: ${response.status} - ${response.statusText}`);
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        // normalizar y cargar inventario
        inventario = data.map(item => {
            const codigoRaw = item.Codigo ?? item.codigo ?? item.Código ?? item.Cod ?? '';
            const codigo = ('' + codigoRaw).trim().toUpperCase();
            return { codigo, estado: 'pendiente' };
        }).filter(it => it.codigo); // quitar vacíos

        localStorage.setItem('inventario', JSON.stringify(inventario));
        actualizarMapeo();
        actualizarTabla();
        excelInicialCargado = true;
        setResult('Escaneo e inventario abierto para las bibliotecas del mundo. / Scanning and open inventory for the world’s libraries', 'blue');
    } catch (error) {
        console.error('', error);
        setResult('' + (error.message || error), 'orange');
        // continuar sin bloquear
    }
}

// ---------- UTILIDADES ----------
function setResult(text, color) {
    const el = document.getElementById('result');
    if (el) el.innerHTML = `<p style="color: ${color || 'black'};">${text}</p>`;
}

function guardarInventario() {
    localStorage.setItem('inventario', JSON.stringify(inventario));
}

function actualizarMapeo() {
    codigoAFila = {};
    inventario.forEach((item, index) => {
        const codigo = (item.codigo || '').toString().trim().toUpperCase();
        item.codigo = codigo;
        // si no tiene estado, dejar como pendiente
        item.estado = item.estado || 'pendiente';
        codigoAFila[codigo] = index;
    });
}

// ---------- QUAGGA (cámara / escaneo) ----------
function iniciarQuagga() {
    if (quaggaIniciado) {
        setResult('La cámara ya está iniciada. / The camera is already started.', 'green');
        return;
    }
    if (typeof Quagga === 'undefined') {
        setResult('Error: QuaggaJS no cargó. Verifica la conexión a internet. / Error: QuaggaJS failed to load. Please check your internet connection.', 'red');
        return;
    }

    Quagga.init({
        inputStream: {
            name: "Live",
            type: "LiveStream",
            target: document.querySelector('#interactive'),
            constraints: {
                width: { ideal: 640 },
                height: { ideal: 480 },
                facingMode: "environment"
            }
        },
        locator: { patchSize: "medium", halfSample: true },
        numOfWorkers: 2,
        decoder: { readers: ["code_128_reader", "ean_reader", "ean_8_reader", "code_39_reader", "upc_reader"] },
        locate: true
    }, function(err) {
        if (err) {
            console.error('Quagga init error:', err);
            setResult('Error: No se pudo acceder a la cámara. Permite permisos y toca "Iniciar Cámara" de nuevo. / Camera access failed. Please allow permissions and tap "Start Camera" again.', 'red');
            return;
        }
        Quagga.start();
        quaggaIniciado = true;
        const camEl = document.getElementById('camaraIndicador');
        if (camEl) camEl.style.display = 'block';
        setResult('Cámara iniciada. Escanea un código. / Camera started. Please scan a code.', 'green');
    });

    Quagga.onDetected(function(result) {
        try {
            let code = result.codeResult.code || '';
            code = code.toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
            // Filtro según la regla: inicia con 'B' y >=7 caracteres
            if (!code.startsWith('B') || code.length < 7) return;
            procesarCodigo(code);
        } catch (e) {
            console.warn('Error procesando resultado Quagga:', e);
        }
    });
}

// ---------- PROCESO DE CÓDIGO (marcar TODO como "encontrado") ----------
function procesarCodigo(codigo) {
    codigo = (codigo || '').toString().trim().toUpperCase();
    if (!codigo) return;

    if (codigoAFila[codigo] !== undefined) {
        // existe en inventario: marcar encontrado
        inventario[codigoAFila[codigo]].estado = 'encontrado';
    } else {
        // NO existe: agregar y marcar encontrado (ya no hay "nuevo")
        inventario.push({ codigo: codigo, estado: 'encontrado' });
        codigoAFila[codigo] = inventario.length - 1;
    }

    setResult(`✔ Código <span style="font-weight: bold; color: black; font-size: 32px;">${codigo}</span> marcado como encontrado / it has been marked as found`, 'green');
    guardarInventario();
    actualizarTabla();
}

// ---------- PROCESAR MANUAL ----------
function procesarManual() {
    const el = document.getElementById('codigoManual');
    if (!el) return;
    const codigo = el.value.trim().toUpperCase();
    if (!codigo) {
        alert('Ingresa un código válido. / Enter a valid code.');
        return;
    }
    procesarCodigo(codigo);
    el.value = '';
}

// ---------- TABLA UI ----------
function actualizarTabla() {
    const tbody = document.querySelector('#inventarioTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (!Array.isArray(inventario) || inventario.length === 0) {
        tbody.innerHTML = '<tr><td colspan="2" style="text-align:center; color:#666;">No hay datos</td></tr>';
        return;
    }

    inventario.forEach(item => {
        const clase = (item.estado === 'encontrado') ? 'verde' : (item.estado === 'morado' ? 'morado' : '');
        const estado = item.estado || '';
        const row = `<tr class="${clase}"><td>${item.codigo}</td><td>${estado}</td></tr>`;
        tbody.insertAdjacentHTML('beforeend', row);
    });
}

// ---------- DESCARGAR EXCEL ----------
function descargarExcel() {
    if (!inventario || inventario.length === 0) {
        alert('No hay datos para descargar.');
        return;
    }
    const ws = XLSX.utils.json_to_sheet(inventario);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Inventario");
    XLSX.writeFile(wb, "inventario_actualizado.xlsx");
}

// ---------- RESET ----------
function resetearInventario() {
    if (!confirm('¿Seguro que quieres resetear el inventario? Se perderán cambios no guardados. / Are you sure you want to reset the inventory? Unsaved changes will be lost.')) return;
    localStorage.removeItem('inventario');
    inventario = [];
    actualizarMapeo();
    actualizarTabla();
    // intentar recargar excel inicial
    cargarExcelInicial();
    setResult('Inventario reseteado. / Inventory has been reset', 'blue');
}

// ---------- MODAL GUIA ----------
function mostrarGuia() {
    const modal = document.getElementById('guiaModal');
    if (modal) modal.style.display = 'block';
}
function cerrarGuia() {
    const modal = document.getElementById('guiaModal');
    if (modal) modal.style.display = 'none';
}

// ---------- EVENT LISTENERS SEGUROS ----------
function safeAddListener(id, evt, fn) {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener(evt, fn);
}

// botones
safeAddListener('guiaBtn', 'click', mostrarGuia);
safeAddListener('guiaBtn', 'touchstart', mostrarGuia);
safeAddListener('closeGuia', 'click', cerrarGuia);
safeAddListener('closeGuia', 'touchstart', cerrarGuia);
safeAddListener('iniciarCamaraBtn', 'click', iniciarQuagga);
safeAddListener('iniciarCamaraBtn', 'touchstart', iniciarQuagga);
safeAddListener('procesarBtn', 'click', procesarManual);
safeAddListener('procesarBtn', 'touchstart', procesarManual);
safeAddListener('descargarBtn', 'click', descargarExcel);
safeAddListener('descargarBtn', 'touchstart', descargarExcel);
safeAddListener('resetBtn', 'click', resetearInventario);
safeAddListener('resetBtn', 'touchstart', resetearInventario);

// Cargar al inicio
(async () => {
    // primero cargar inventario guardado si existe
    try {
        // actualizar mapeo de lo guardado
        actualizarMapeo();
        // intentar cargar excel si no hay inventario guardado
        if (!inventario || inventario.length === 0) {
            await cargarExcelInicial();
        } else {
            actualizarTabla();
        }
    } catch (e) {
        console.warn('Error inicial:', e);
    }
})();

// ==========================================
// MODULO CRUCE DE INVENTARIOS (INDEPENDIENTE)
// ==========================================

// Mostrar / ocultar panel
function toggleCruceInventario() {
    const panel = document.getElementById("cruceInventarioPanel");
    if (!panel) return;
    panel.style.display = panel.style.display === "none" ? "block" : "none";
}

// Simulación de progreso (UX)
function simularProgreso() {
    const barra = document.getElementById("progresoCruce");
    const texto = document.getElementById("textoProgreso");
    let progreso = 0;

    return setInterval(() => {
        if (progreso < 90) {
            progreso += Math.random() * 7;
            barra.style.width = `${Math.min(progreso, 90)}%`;
            texto.textContent = `${Math.floor(progreso)}%`;
        }
    }, 300);
}

// Cruce principal
async function cruzarInventarioExcel() {
    const inv = document.getElementById("excelInventario").files[0];
    const esc = document.getElementById("excelEscaneo").files[0];

    if (!inv || !esc) {
        alert("Selecciona ambos archivos Excel");
        return;
    }

    const barra = document.getElementById("progresoCruce");
    const texto = document.getElementById("textoProgreso");

    barra.style.width = "0%";
    texto.textContent = "0%";

    const intervalo = simularProgreso();

    const formData = new FormData();
    formData.append("inventario", inv);
    formData.append("escaneo", esc);

    try {
        const response = await fetch("/api/cruzar-inventario", {
            method: "POST",
            body: formData
        });

        if (!response.ok) throw new Error();

        const blob = await response.blob();

        clearInterval(intervalo);
        barra.style.width = "100%";
        texto.textContent = "100%";

        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "inventario_cruzado.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();

    } catch (err) {
        clearInterval(intervalo);
        alert("Error durante el cruce de inventarios");
    }
}

// Listeners seguros (clic + móvil)
safeAddListener("toggleCruceBtn", "click", toggleCruceInventario);
safeAddListener("toggleCruceBtn", "touchstart", toggleCruceInventario);

safeAddListener("cruceExcelBtn", "click", cruzarInventarioExcel);
safeAddListener("cruceExcelBtn", "touchstart", cruzarInventarioExcel);


