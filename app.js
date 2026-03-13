const firebaseConfig = {
  apiKey: "AIzaSyCyp_5RXNbAzDbOQpupxphC_Y3KOoLIY2E",
  authDomain: "siga-c1830.firebaseapp.com",
  projectId: "siga-c1830",
  storageBucket: "siga-c1830.firebasestorage.app",
  messagingSenderId: "937887702037",
  appId: "1:937887702037:web:7909784c55f2ed3883821c",
  measurementId: "G-90QFPGDSEP"
};

// Inicialización Firebase Compat
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

// --- Configuración PFA ---
const HIERARCHY_ORDER = [
    "Comisario General", "Comisario Mayor", "Comisario Inspector", 
    "Comisario", "Subcomisario", "Principal", "Inspector", 
    "Subinspector", "Ayudante", "Suboficial Mayor", "Suboficial Auxiliar", 
    "Suboficial Escribiente", "Sargento 1ro", "Sargento", 
    "Cabo 1ro", "Cabo", "Agente"
];

const DIAS = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];

// --- Estado Global (Ahora sincronizado con Firestore) ---
let personal = [];
let licencias = [];
let guardias = {};
let serviciosExternos = [];
let usuarios_siga = {};
let currentUser = JSON.parse(localStorage.getItem('siga_session')) || null;

// --- Estado Global de Calendario ---
let currentMonth = new Date().getMonth();
let currentYear = new Date().getFullYear();

const MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
];

// --- Inicialización y Sincronización en Tiempo Real ---
document.addEventListener('DOMContentLoaded', () => {
    try {
        startClock();
        initNavigation();
        initHierarchySelect();
        initCalendarSelects();
        
        // Sincronización en Tiempo Real con Firestore
        syncData();
        
        initEventListeners();
        console.log("SIGA: Sistema Dashboard Institucional iniciado con Firebase Firestore.");
    } catch (error) {
        console.error("SIGA Error:", error);
    }
});

async function syncData() {
    // Carga inicial rápida de Usuarios (para evitar bloqueo en login)
    try {
        const userSnap = await db.collection("usuarios").get();
        userSnap.forEach(doc => {
            usuarios_siga[doc.id] = doc.data();
        });
    } catch (e) {
        console.error("Error en carga inicial:", e);
        if (e.code === 'permission-denied') {
            alert("⚠️ Error de Firebase: Permisos denegados.");
        }
    }

    // 1. Personal
    db.collection("personal").onSnapshot((snapshot) => {
        personal = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        sortPersonal();
        renderPersonalCards();
        renderAdminPersonal();
        updatePersonalSelects();
        renderHomeStats();
    });

    // 2. Usuarios
    db.collection("usuarios").onSnapshot((snapshot) => {
        usuarios_siga = {};
        snapshot.docs.forEach(doc => {
            usuarios_siga[doc.id] = doc.data();
        });

        // Check if admin user exists, if not, create it
        if (!usuarios_siga['admin']) {
            db.collection("usuarios").doc("admin").set({
                username: 'admin',
                password: 'siga2026',
                rol: 'Administrador',
                nombre: '',
                apellido: 'ADMINISTRADOR',
                jerarquia: '',
                legajo: ''
            });
            console.log("Admin user created.");
        }

        renderUsersTable();
        // Recargar sesión si el usuario cambió
        if (currentUser && usuarios_siga[currentUser.username]) {
            currentUser = { ...usuarios_siga[currentUser.username] };
            updateUIForRole();
        }
    });

    // 3. Licencias
    db.collection("licencias").onSnapshot((snapshot) => {
        licencias = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        renderLicencias();
        renderGuardGrid();
    });

    // 4. Guardias
    db.collection("guardias").onSnapshot((snapshot) => {
        guardias = {};
        snapshot.docs.forEach(doc => {
            guardias[doc.id] = doc.data().asignaciones;
        });
        renderGuardGrid();
        renderHomeStats();
    });

    // 5. Servicios Externos
    db.collection("servicios").onSnapshot((snapshot) => {
        serviciosExternos = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        renderServices();
    });

    // Verificar Auth después de la primera carga
    checkAuth();
}

// --- Exponer funciones al scope global (para onclick en HTML) ---
window.switchModule = switchModule;
window.changeMonth = changeMonth;
window.goToToday = goToToday;
window.jumpToDate = jumpToDate;
window.editPersonal = editPersonal;
window.deletePersonal = deletePersonal;
window.editUser = editUser;
window.deleteUser = deleteUser;
window.removeFromQA = removeFromQA;
window.removeStaffFromService = removeStaffFromService;
window.addStaffToService = addStaffToService;
window.deleteService = deleteService;
window.exportarExcelServicio = exportarExcelServicio;
window.deleteLicencia = deleteLicencia;

function initUsers() {
    // Ya no es necesario hardcodear aquí, se maneja en Firestore
}

function checkAuth() {
    const loginScreen = document.getElementById('login-screen');
    const appContainer = document.getElementById('app-container');
    const loginCard = document.getElementById('login-card');
    const loader = document.getElementById('loader');

    console.log("checkAuth invocado. currentUser:", currentUser);

    if (currentUser && currentUser.username) {
        loginScreen.style.display = 'none';
        appContainer.style.display = 'flex';
        updateUIForRole();
        renderHomeStats();
        switchModule('home');
    } else {
        // Asegurar que el formulario sea visible y el loader esté oculto
        loginScreen.style.display = 'flex';
        appContainer.style.display = 'none';
        if (loginCard) loginCard.style.display = 'block';
        if (loader) loader.style.display = 'none';
        
        // Limpiar campos para re-ingreso
        const userField = document.getElementById('login-user');
        if (userField) {
            userField.value = '';
            userField.focus();
        }
        const passField = document.getElementById('login-pass');
        if (passField) passField.value = '';
    }
}

function updateUIForRole() {
    if (!currentUser) return;

    // Nombre en Header
    const fullName = currentUser.username === 'admin' 
        ? 'ADMINISTRADOR' 
        : `${currentUser.jerarquia} ${currentUser.apellido}`.toUpperCase();
    document.getElementById('active-user-name').innerText = fullName;
    
    // Perfil en Modal
    document.getElementById('profile-full-name').innerText = currentUser.username === 'admin' 
        ? 'ADMINISTRADOR' 
        : `${currentUser.nombre} ${currentUser.apellido}`;
    document.getElementById('profile-role').innerText = currentUser.rol.toUpperCase();

    // Home Page Info
    document.getElementById('home-user-name').innerText = `BIENVENIDO, ${fullName}`;
    document.getElementById('home-user-role').innerText = currentUser.rol.toUpperCase();

    // Visibilidad por Rol
    if (currentUser.rol === 'Administrador') {
        document.body.classList.add('role-admin');
    } else {
        document.body.classList.remove('role-admin');
        // El Operador PUEDE entrar a admin-personal para cargar, pero no a configuración
        const activeModule = document.querySelector('.module.active');
        if (activeModule && activeModule.id === 'configuracion') {
            switchModule('personal');
        }
    }
}

function renderHomeStats() {
    // Total Personal
    document.getElementById('stat-personal-count').innerText = personal.length;

    // Asignaciones hoy
    const now = new Date();
    const dateKey = `${now.getFullYear()}-${(now.getMonth() + 1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}`;
    const todayGuards = guardias[dateKey] ? guardias[dateKey].length : 0;
    document.getElementById('stat-active-guards').innerText = todayGuards;
}

function switchModule(targetId) {
    const modules = document.querySelectorAll('.module');
    const navBtns = document.querySelectorAll('.nav-btn');
    
    navBtns.forEach(btn => {
        btn.classList.remove('active');
        if (btn.getAttribute('data-target') === targetId) btn.classList.add('active');
    });

    modules.forEach(m => {
        m.classList.remove('active');
        if (m.id === targetId) m.classList.add('active');
    });
}

function initCalendarSelects() {
    const mSelect = document.getElementById('cal-month');
    const ySelect = document.getElementById('cal-year');
    if (!mSelect || !ySelect) return;

    mSelect.innerHTML = MESES.map((m, i) => `<option value="${i}" ${i === currentMonth ? 'selected' : ''}>${m.toUpperCase()}</option>`).join('');
    
    let years = '';
    for (let i = currentYear - 5; i <= currentYear + 5; i++) {
        years += `<option value="${i}" ${i === currentYear ? 'selected' : ''}>${i}</option>`;
    }
    ySelect.innerHTML = years;
}

function changeMonth(delta) {
    currentMonth += delta;
    if (currentMonth < 0) { currentMonth = 11; currentYear--; }
    if (currentMonth > 11) { currentMonth = 0; currentYear++; }
    updateCalendarState();
}

function goToToday() {
    const today = new Date();
    currentMonth = today.getMonth();
    currentYear = today.getFullYear();
    updateCalendarState();
}

function jumpToDate() {
    currentMonth = parseInt(document.getElementById('cal-month').value);
    currentYear = parseInt(document.getElementById('cal-year').value);
    updateCalendarState();
}

function updateCalendarState() {
    initCalendarSelects();
    renderGuardGrid();
}

// --- Reloj Elegante ---
function startClock() {
    function update() {
        const now = new Date();
        const dateStr = now.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' });
        const dayStr = now.toLocaleDateString('es-AR', { weekday: 'short' }).replace('.', '').toUpperCase();
        const timeStr = now.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
        
        document.getElementById('current-date').innerText = `${dayStr} ${dateStr}`;
        document.getElementById('current-time').innerText = timeStr;
    }
    update();
    setInterval(update, 1000);
}

// --- Gestión de Personal (Visualización Premium) ---
function renderPersonalCards() {
    const grid = document.getElementById('personal-cards');
    const search = document.getElementById('search-personal')?.value.toLowerCase() || '';
    if (!grid) return;
    grid.innerHTML = '';

    personal.forEach((p, index) => {
        if (!p.nombre.toLowerCase().includes(search) && !p.jerarquia.toLowerCase().includes(search)) return;

        const card = document.createElement('div');
        card.className = 'person-card';
        card.onclick = () => showFicha(index);
        card.innerHTML = `
            <span class="rank-label">${p.jerarquia}</span>
            <h4>${p.nombre}</h4>
            <p style="font-size: 0.8rem; color: var(--text-secondary);">LP: ${p.legajo}</p>
        `;
        grid.appendChild(card);
    });
}

function showFicha(index) {
    const p = personal[index];
    const body = document.getElementById('ficha-body');
    body.innerHTML = `
        <div class="ficha-header" style="border-bottom: 2px solid var(--primary-blue); padding-bottom: 1rem; margin-bottom: 1.5rem;">
            <span class="rank-label">${p.jerarquia}</span>
            <h2 style="color: var(--primary-blue); margin-top: 0.5rem;">${p.nombre}</h2>
        </div>
        <div class="ficha-grid">
            <div class="ficha-item"><label>Legajo</label><p>${p.legajo}</p></div>
            <div class="ficha-item"><label>DNI</label><p>${p.dni}</p></div>
            <div class="ficha-item"><label>Correo Electrónico</label><p>${p.gmail || '-'}</p></div>
            <div class="ficha-item"><label>Obra Social</label><p>${p.os || '-'}</p></div>
            <div class="ficha-item"><label>CSJPFA</label><p>${p.csjpfa || '-'}</p></div>
            <div class="ficha-item"><label>Armamento</label><p>${p.armamento || 'SIN ASIGNAR'}</p></div>
            <div class="ficha-item full-width"><label>Domicilio</label><p>${p.domicilio}</p></div>
            <div class="ficha-item"><label>Teléfono</label><p>${p.telefono}</p></div>
            <div class="ficha-item"><label>Alternativo</label><p>${p.telAlternativo || '-'}</p></div>
        </div>
    `;
    document.getElementById('ficha-modal').style.display = 'block';
}

// --- Administración ---
function renderAdminPersonal() {
    const tbody = document.getElementById('admin-personal-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    personal.forEach((p, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><span class="rank-label" style="background:#f1f5f9; color:var(--primary-blue); font-size:0.6rem;">${p.jerarquia}</span></td>
            <td><b>${p.nombre}</b></td>
            <td>${p.legajo}</td>
            <td>${p.dni}</td>
            <td class="admin-only">
                <button class="btn btn-secondary btn-sm" onclick="editPersonal(${index})"><i class="ph ph-note-pencil"></i></button>
                <button class="btn btn-danger btn-sm" onclick="deletePersonal(${index})"><i class="ph ph-trash"></i></button>
            </td>
            <td class="operator-view" style="display: none;">-</td>
        `;
        tbody.appendChild(tr);
    });
    updatePersonalSelects();
}

// --- Gestión de Usuarios (Configuración) ---
function renderUsersTable() {
    const tbody = document.getElementById('users-table-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    Object.values(usuarios_siga).forEach(u => {
        const tr = document.createElement('tr');
        const displayName = u.username === 'admin' 
            ? 'ADMINISTRADOR' 
            : `${u.jerarquia} ${u.apellido}, ${u.nombre}`;
            
        tr.innerHTML = `
            <td><b>${u.username}</b></td>
            <td>${displayName}</td>
            <td><span class="rank-label" style="background:${u.rol==='Administrador'?'#ecfdf5':'#eff6ff'}; color:${u.rol==='Administrador'?'#10b981':'#3b82f6'}; border:none;">${u.rol.toUpperCase()}</span></td>
            <td>
                ${u.username !== 'admin' ? `
                <button class="btn btn-secondary btn-sm" onclick="editUser('${u.username}')"><i class="ph ph-note-pencil"></i></button>
                <button class="btn btn-danger btn-sm" onclick="deleteUser('${u.username}')"><i class="ph ph-trash"></i></button>
                ` : '<small>Protegido</small>'}
            </td>
        `;
        tbody.appendChild(tr);
    });
}

async function deleteUser(username) {
    if (username === 'admin') return;
    if (confirm(`¿Eliminar al usuario ${username}?`)) {
        try {
            await db.collection("usuarios").doc(username).delete();
        } catch (error) {
            console.error("Error al eliminar usuario:", error);
            alert("Error en la operación.");
        }
    }
}

function editUser(username) {
    const u = usuarios_siga[username];
    document.getElementById('user-edit-id').value = u.username;
    document.getElementById('u-jerarquia').value = u.jerarquia;
    document.getElementById('u-nombre').value = u.nombre;
    document.getElementById('u-apellido').value = u.apellido;
    document.getElementById('u-username').value = u.username;
    document.getElementById('u-username').disabled = true;
    document.getElementById('u-password').value = u.password;
    document.getElementById('u-rol').value = u.rol;
    
    document.getElementById('user-modal-title').innerText = 'Editar Usuario';
    document.getElementById('user-modal').style.display = 'block';
}

// --- Guardia (Calendario Mensual Profesional) ---
function renderGuardGrid() {
    const grid = document.getElementById('guard-grid');
    if (!grid) return;
    grid.innerHTML = '';

    const firstDay = new Date(currentYear, currentMonth, 1).getDay();
    const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
    const prevMonthDays = new Date(currentYear, currentMonth, 0).getDate();
    
    // Ajuste para Lunes como primer día (0=Dom, 1=Lun, ..., 6=Sab)
    let startOffset = firstDay === 0 ? 6 : firstDay - 1;

    // Días del mes anterior (relleno)
    for (let i = startOffset - 1; i >= 0; i--) {
        const d = prevMonthDays - i;
        grid.appendChild(createDayElement(d, currentMonth - 1, currentYear, true));
    }

    // Días del mes actual
    for (let i = 1; i <= daysInMonth; i++) {
        grid.appendChild(createDayElement(i, currentMonth, currentYear, false));
    }

    // Días del mes siguiente (relleno hasta 42 celdas para grilla perfecta 7x6)
    const totalCells = 42;
    const remaining = totalCells - grid.children.length;
    for (let i = 1; i <= remaining; i++) {
        grid.appendChild(createDayElement(i, currentMonth + 1, currentYear, true));
    }
}

function createDayElement(day, month, year, isOtherMonth) {
    const dateObj = new Date(year, month, day);
    const d = dateObj.getDate();
    const m = dateObj.getMonth();
    const y = dateObj.getFullYear();
    const dateKey = `${y}-${(m + 1).toString().padStart(2, '0')}-${d.toString().padStart(2, '0')}`;
    
    const div = document.createElement('div');
    div.className = `guard-day ${isOtherMonth ? 'other-month' : ''}`;
    
    const today = new Date();
    if (d === today.getDate() && m === today.getMonth() && y === today.getFullYear()) {
        div.classList.add('today');
    }

    div.onclick = () => openQuickAdd(dateKey);
    
    const dayAssignments = guardias[dateKey] || [];
    let listHtml = '';
    dayAssignments.forEach(nombre => {
        const p = personal.find(pers => pers.nombre === nombre);
        if (p) {
            const apellido = p.nombre.split(' ').pop();
            listHtml += `<span><b class="rank">${p.jerarquia.split(' ').map(s => s[0]).join('')}.</b> ${apellido}</span>`;
        }
    });

    div.innerHTML = `
        <div class="day-num">${d}</div>
        <div class="assignments-list">
            ${listHtml}
        </div>
    `;
    return div;
}

function openQuickAdd(dateKey) {
    const parts = dateKey.split('-');
    const displayDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
    document.getElementById('qa-title').innerText = `Asignación: ${displayDate}`;
    
    const select = document.getElementById('qa-select');
    const dayAssignments = guardias[dateKey] || [];
    
    let options = '<option value="">+ Seleccionar personal...</option>';
    personal.forEach(p => {
        if (dayAssignments.includes(p.nombre)) return;
        
        // Validación de Licencia por fecha específica
        const formattedDate = new Date(parts[0], parts[1]-1, parts[2]);
        const onLic = licencias.some(l => {
            const desde = new Date(l.desde + "T00:00:00");
            const hasta = new Date(l.hasta + "T23:59:59");
            return l.nombre === p.nombre && formattedDate >= desde && formattedDate <= hasta;
        });

        options += `<option value="${p.nombre}" ${onLic ? 'disabled' : ''}>
            ${p.jerarquia} ${p.nombre}${onLic ? ' (LICENCIA)' : ''}
        </option>`;
    });
    
    select.innerHTML = options;
    select.onchange = async (e) => {
        if (!e.target.value) return;
        if (!guardias[dateKey]) guardias[dateKey] = [];
        const nuevasAsignaciones = [...guardias[dateKey], e.target.value];
        
        try {
            await db.collection("guardias").doc(dateKey).set({ asignaciones: nuevasAsignaciones });
        } catch (error) {
            console.error("Error al asignar guardia:", error);
            alert("Error al guardar en la nube.");
        }
    };
    
    const listEl = document.getElementById('qa-list');
    listEl.innerHTML = '';
    dayAssignments.forEach((nombre, idx) => {
        const tag = document.createElement('div');
        tag.className = 'person-tag';
        tag.innerHTML = `<span>${nombre}</span><button class="remove-tag" onclick="removeFromQA('${dateKey}', ${idx})">×</button>`;
        listEl.appendChild(tag);
    });
    
    document.getElementById('quick-add-modal').style.display = 'block';
}

async function removeFromQA(dateKey, idx) {
    const nuevasAsignaciones = [...guardias[dateKey]];
    nuevasAsignaciones.splice(idx, 1);
    try {
        await db.collection("guardias").doc(dateKey).set({ asignaciones: nuevasAsignaciones });
        openQuickAdd(dateKey); // Refrescamos el modal
    } catch (error) {
        console.error("Error al quitar asignación:", error);
    }
}

// --- Servicios Operativos ---
function renderServices() {
    const container = document.getElementById('services-container');
    if (!container) return;
    container.innerHTML = '';

    if (serviciosExternos.length === 0) {
        container.innerHTML = '<div style="text-align:center; padding: 4rem; color: var(--text-secondary);">No hay servicios programados.</div>';
        return;
    }

    serviciosExternos.forEach((s, sIdx) => {
        const card = document.createElement('div');
        card.className = 'service-event-card';
        
        const staffHtml = s.asignados.map((nombre, pIdx) => {
            const p = personal.find(pers => pers.nombre === nombre);
            return p ? `
                <div class="person-tag small" style="flex-direction: column; align-items: flex-start; gap: 0.2rem; min-width: 180px;">
                    <div style="display:flex; width:100%; justify-content: space-between; align-items:center;">
                        <b>${p.jerarquia.substring(0,4)}.</b> ${p.nombre}
                        <button class="remove-tag" onclick="removeStaffFromService(${sIdx}, ${pIdx})">×</button>
                    </div>
                    <div style="font-size: 0.65rem; opacity: 0.8; display:flex; flex-direction:column;">
                        <span>LP: ${p.legajo} | DNI: ${p.dni}</span>
                        <span><i class="ph ph-envelope-simple"></i> ${p.gmail || '-'}</span>
                        <span><i class="ph ph-phone"></i> ${p.telefono}</span>
                    </div>
                </div>` : '';
        }).join('');

        card.innerHTML = `
            <div class="event-info">
                <h4>${s.motivo}</h4>
                <div class="event-details">
                    <div><i class="ph ph-calendar-blank"></i> ${s.fecha}</div>
                    <div><i class="ph ph-clock"></i> ${s.hora} HS</div>
                    <div><i class="ph ph-map-pin"></i> ${s.lugar}</div>
                </div>
                <button class="btn btn-danger btn-sm" style="margin-top:1.5rem;" onclick="deleteService(${sIdx})">Cancelar Evento</button>
            </div>
                <div class="event-staff">
                    <h5>Efectivos Comisionados</h5>
                    <div class="staff-list">${staffHtml}</div>
                    <select class="form-group" style="margin-top:1rem; padding:0.5rem;" onchange="addStaffToService(${sIdx}, this.value)">
                        <option value="">+ Asignar...</option>
                        ${personal.map(p => `<option value="${p.nombre}">${p.jerarquia} ${p.nombre}</option>`).join('')}
                    </select>
                    <button class="btn btn-secondary btn-sm" style="margin-top:1rem; width:100%; border-color: var(--primary-blue); color: var(--primary-blue);" onclick="exportarExcelServicio(${sIdx})">
                        <i class="ph ph-file-xls"></i> Exportar Planilla Oficial
                    </button>
                </div>
        `;
        container.appendChild(card);
    });
}

async function addStaffToService(sIdx, nombre) {
    if (!nombre) return;
    const s = serviciosExternos[sIdx];
    if (s.asignados.includes(nombre)) return;
    
    const nuevosAsignados = [...s.asignados, nombre];
    try {
        await db.collection("servicios").doc(s.id).update({ asignados: nuevosAsignados });
    } catch (error) {
        console.error("Error al asignar personal al servicio:", error);
    }
}

async function removeStaffFromService(sIdx, pIdx) {
    const s = serviciosExternos[sIdx];
    const nuevosAsignados = [...s.asignados];
    nuevosAsignados.splice(pIdx, 1);
    try {
        await db.collection("servicios").doc(s.id).update({ asignados: nuevosAsignados });
    } catch (error) {
        console.error("Error al quitar personal del servicio:", error);
    }
}

async function deleteService(idx) {
    const s = serviciosExternos[idx];
    if (confirm('¿Eliminar servicio operativo?')) {
        try {
            await db.collection("servicios").doc(s.id).delete();
        } catch (error) {
            console.error("Error al eliminar servicio:", error);
        }
    }
}

// --- Exportación Excel Premium de Servicios ---
function exportarExcelServicio(idx) {
    const s = serviciosExternos[idx];
    const fileName = `Planilla_Servicio_${s.motivo.replace(/\s+/g, '_')}.xlsx`;
    
    // 1. Preparar Personal (Ordenado por Jerarquía)
    const staff = s.asignados.map(nombre => personal.find(p => p.nombre === nombre)).filter(p => p);
    staff.sort((a, b) => HIERARCHY_ORDER.indexOf(a.jerarquia) - HIERARCHY_ORDER.indexOf(b.jerarquia));
    
    // 2. Definir Estilos Comunes
    const styleHeader = {
        fill: { fgColor: { rgb: "FFFF00" } }, // Amarillo
        font: { bold: true, color: { rgb: "000000" }, sz: 12 },
        alignment: { horizontal: "center", vertical: "center" },
        border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } }
    };
    
    const styleTitle = {
        font: { bold: true, sz: 11 },
        border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } }
    };

    const styleCell = {
        border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } },
        alignment: { vertical: "center" }
    };

    // 3. Crear Estructura de Datos para la Planilla
    const rows = [
        // Fila 1: Encabezado Institucional
        [{ v: "POLICIA FEDERAL ARGENTINA", s: styleHeader }, "", "", "", ""],
        ["", "", "", "", ""], // Espacio
        [{ v: "DIVISIÓN:", s: styleTitle }, { v: "CENTRO DE ENTRENAMIENTO TECNOLÓGICO", s: styleCell }, "", "", ""],
        [{ v: "MOTIVO:", s: styleTitle }, { v: s.motivo.toUpperCase(), s: styleCell }, "", "", ""],
        [{ v: "LUGAR:", s: styleTitle }, { v: s.lugar.toUpperCase(), s: styleCell }, "", "", ""],
        [{ v: "FECHA/HORA:", s: styleTitle }, { v: `${s.fecha} - ${s.hora} HS`, s: styleCell }, "", "", ""],
        ["", "", "", "", ""], // Espacio
        // Encabezados de Tabla
        [
            { v: "JERARQUÍA", s: styleHeader }, 
            { v: "NOMBRE Y APELLIDO", s: styleHeader }, 
            { v: "LP", s: styleHeader }, 
            { v: "DNI", s: styleHeader }, 
            { v: "TELÉFONO", s: styleHeader }
        ]
    ];

    // 4. Agregar Personal
    staff.forEach(p => {
        rows.push([
            { v: p.jerarquia, s: styleCell },
            { v: p.nombre, s: styleCell },
            { v: p.legajo, s: styleCell },
            { v: p.dni, s: styleCell },
            { v: p.telefono, s: styleCell }
        ]);
    });

    // 5. Crear Libro y Hoja
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(rows);

    // 6. Configuraciones Finales (Merge y AutoFit)
    ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, // Merge PFA Header
        { s: { r: 2, c: 1 }, e: { r: 2, c: 4 } }, // Merge División info
        { s: { r: 3, c: 1 }, e: { r: 3, c: 4 } }, // Merge Motivo info
        { s: { r: 4, c: 1 }, e: { r: 4, c: 4 } }, // Merge Lugar info
        { s: { r: 5, c: 1 }, e: { r: 5, c: 4 } }  // Merge Fecha info
    ];

    // Lógica Automática de AutoFit
    const colWidths = [15, 30, 10, 12, 15]; // Anchos base
    rows.forEach(row => {
        row.forEach((cell, i) => {
            const val = cell.v ? String(cell.v) : "";
            if (val.length + 2 > colWidths[i]) colWidths[i] = val.length + 2;
        });
    });
    ws['!cols'] = colWidths.map(w => ({ wch: w }));

    XLSX.utils.book_append_sheet(wb, ws, "Planilla_Servicio");
    XLSX.writeFile(wb, fileName);
}

// --- Licencias ---
function renderLicencias() {
    const tbody = document.getElementById('licencias-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    licencias.forEach((l, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><b>${l.nombre}</b></td>
            <td>${l.tipo}</td>
            <td>${l.desde}</td>
            <td>${l.hasta}</td>
            <td><span class="rank-label" style="background:#ecfdf5; color:#10b981; border:none;">ACTIVA</span></td>
            <td><button class="btn btn-danger btn-sm" onclick="deleteLicencia(${index})"><i class="ph ph-trash"></i></button></td>
        `;
        tbody.appendChild(tr);
    });
}

async function deleteLicencia(index) {
    const l = licencias[index];
    if (confirm('¿Eliminar registro de licencia?')) {
        try {
            await db.collection("licencias").doc(l.id).delete();
        } catch (error) {
            console.error("Error al eliminar licencia:", error);
            alert("Error en la operación.");
        }
    }
}

// --- Event Listeners ---
function initNavigation() {
    const navBtns = document.querySelectorAll('.nav-btn');
    const modules = document.querySelectorAll('.module');

    navBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const target = btn.getAttribute('data-target');
            navBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            modules.forEach(m => {
                m.classList.remove('active');
                if (m.id === target) m.classList.add('active');
            });
        });
    });
}

function initHierarchySelect() {
    const pSelects = [document.getElementById('p-jerarquia'), document.getElementById('u-jerarquia')];
    pSelects.forEach(sel => {
        if (sel) {
            sel.innerHTML = HIERARCHY_ORDER.map(h => `<option value="${h}">${h}</option>`).join('');
        }
    });
}

function updatePersonalSelects() {
    const lSelect = document.getElementById('l-personal');
    if (lSelect) {
        lSelect.innerHTML = personal.map(p => `<option value="${p.nombre}">${p.jerarquia} ${p.nombre}</option>`).join('');
    }
}

// --- Funciones de Guardado (Refactorizadas para Firestore) ---
function sortPersonal() {
    personal.sort((a, b) => HIERARCHY_ORDER.indexOf(a.jerarquia) - HIERARCHY_ORDER.indexOf(b.jerarquia));
}

async function deletePersonal(index) {
    const p = personal[index];
    if (confirm(`¿Dar de BAJA a ${p.nombre}?`)) {
        try {
            await db.collection("personal").doc(p.id).delete();
        } catch (error) {
            console.error("Error al eliminar personal:", error);
            alert("Error al eliminar los datos.");
        }
    }
}

function editPersonal(index) {
    const p = personal[index];
    document.getElementById('edit-index').value = p.id;
    document.getElementById('p-jerarquia').value = p.jerarquia;
    document.getElementById('p-nombre').value = p.nombre;
    document.getElementById('p-legajo').value = p.legajo;
    document.getElementById('p-dni').value = p.dni;
    document.getElementById('p-gmail').value = p.gmail || '';
    document.getElementById('p-os').value = p.os || '';
    document.getElementById('p-csjpfa').value = p.csjpfa || '';
    document.getElementById('p-arma').value = p.armamento || '';
    document.getElementById('p-domicilio').value = p.domicilio;
    document.getElementById('p-telefono').value = p.telefono;
    document.getElementById('p-tel-alt').value = p.telAlternativo || '';
    
    document.getElementById('modal-title').innerText = 'Modificar Efectivo';
    document.getElementById('personal-modal').style.display = 'block';
}

function initEventListeners() {
    document.getElementById('search-personal')?.addEventListener('input', renderPersonalCards);

    document.getElementById('add-personal-btn')?.addEventListener('click', () => {
        document.getElementById('personal-form').reset();
        document.getElementById('edit-index').value = '';
        document.getElementById('modal-title').innerText = 'Registro de Efectivo';
        document.getElementById('personal-modal').style.display = 'block';
    });

    document.getElementById('personal-form')?.addEventListener('submit', async (e) => {
        e.preventDefault();
        const editId = document.getElementById('edit-index').value; // Ahora guardamos el ID de Firestore si existe
        const legajo = document.getElementById('p-legajo').value;
        const dni = document.getElementById('p-dni').value;

        // Validación de Duplicidad (LP y DNI) - Ahora sobre el array local sincronizado
        const isDuplicate = personal.some((p) => {
            if (editId !== '' && p.id === editId) return false; 
            return p.legajo === legajo || p.dni === dni;
        });

        if (isDuplicate) {
            alert('⚠️ ERROR DE SEGURIDAD: El Legajo Personal (LP) o el DNI ya se encuentran registrados en el sistema. Por favor, verifique los datos.');
            return;
        }

        const data = {
            jerarquia: document.getElementById('p-jerarquia').value,
            nombre: document.getElementById('p-nombre').value.toUpperCase(),
            legajo: document.getElementById('p-legajo').value,
            dni: document.getElementById('p-dni').value,
            gmail: document.getElementById('p-gmail').value,
            os: document.getElementById('p-os').value,
            csjpfa: document.getElementById('p-csjpfa').value,
            armamento: document.getElementById('p-arma').value,
            domicilio: document.getElementById('p-domicilio').value,
            telefono: document.getElementById('p-telefono').value,
            telAlternativo: document.getElementById('p-tel-alt').value
        };

        try {
            if (editId === '') {
                await db.collection("personal").add(data);
            } else {
                await db.collection("personal").doc(editId).update(data);
            }
            document.getElementById('personal-modal').style.display = 'none';
        } catch (error) {
            console.error("Error al guardar personal:", error);
            alert("Error al guardar en la nube.");
        }
    });

    document.getElementById('add-servicio-btn')?.addEventListener('click', () => {
        document.getElementById('servicio-form').reset();
        document.getElementById('servicio-modal').style.display = 'block';
    });

    document.getElementById('servicio-form')?.addEventListener('submit', async (e) => {
        e.preventDefault();
        const data = {
            motivo: document.getElementById('s-motivo').value,
            fecha: document.getElementById('s-fecha').value,
            hora: document.getElementById('s-hora').value,
            lugar: document.getElementById('s-lugar').value,
            asignados: []
        };
        try {
            await db.collection("servicios").add(data);
            document.getElementById('servicio-modal').style.display = 'none';
        } catch (error) {
            console.error("Error al crear servicio:", error);
            alert("Error al guardar en la nube.");
        }
    });

    document.getElementById('add-licencia-btn')?.addEventListener('click', () => {
        document.getElementById('licencia-modal').style.display = 'block';
    });

    document.getElementById('licencia-form')?.addEventListener('submit', async (e) => {
        e.preventDefault();
        const data = {
            nombre: document.getElementById('l-personal').value,
            tipo: document.getElementById('l-tipo').value,
            desde: document.getElementById('l-desde').value,
            hasta: document.getElementById('l-hasta').value
        };
        try {
            await db.collection("licencias").add(data);
            document.getElementById('licencia-modal').style.display = 'none';
        } catch (error) {
            console.error("Error al registrar licencia:", error);
            alert("Error al guardar en la nube.");
        }
    });

    document.getElementById('close-ficha').onclick = () => document.getElementById('ficha-modal').style.display = 'none';
    document.getElementById('close-qa').onclick = () => document.getElementById('quick-add-modal').style.display = 'none';
    document.getElementById('close-modal').onclick = () => document.getElementById('personal-modal').style.display = 'none';
    document.getElementById('close-servicio-modal').onclick = () => document.getElementById('servicio-modal').style.display = 'none';
    document.getElementById('close-licencia-modal').onclick = () => document.getElementById('licencia-modal').style.display = 'none';

    // --- Login Logic ---
    document.getElementById('login-form')?.addEventListener('submit', async (e) => {
        e.preventDefault();
        const user = document.getElementById('login-user').value.toLowerCase().trim();
        const pass = document.getElementById('login-pass').value;
        
        const errorMsg = document.getElementById('login-error');
        const loginCard = document.getElementById('login-card');
        const loader = document.getElementById('loader');

        let userData = usuarios_siga[user];

        // Fallback: Si no está en el objeto local, buscar directamente en Firestore
        if (!userData) {
            console.log("No está en caché, buscando en Firestore...");
            try {
                const userDoc = await db.collection("usuarios").where("username", "==", user).get();
                if (!userDoc.empty) {
                    userData = userDoc.docs[0].data();
                    console.log("Usuario encontrado en Firestore (fallback)");
                }
            } catch (e) {
                console.error("Error en fallback de auth:", e);
                alert("⚠️ Error de conexión: " + e.message);
            }
        }

        if (userData) {
            if (userData.password === pass) {
                currentUser = userData;
                localStorage.setItem('siga_session', JSON.stringify(currentUser));
                errorMsg.style.display = 'none';
                
                // Efecto de Carga Institucional
                loginCard.style.display = 'none';
                loader.style.display = 'flex';

                setTimeout(() => {
                    loader.style.display = 'none';
                    checkAuth();
                }, 1000);
            } else {
                errorMsg.innerText = "⚠️ Contraseña incorrecta";
                errorMsg.style.display = 'block';
            }
        } else {
            errorMsg.innerText = "⚠️ El usuario no existe";
            errorMsg.style.display = 'block';
        }
    });

        if (confirm('¿Cerrar sesión del sistema?')) {
            currentUser = null;
            localStorage.removeItem('siga_session');
            // Volver al estado inicial de login
            checkAuth();
        }

    // --- Profile & Pass Change ---
    document.getElementById('user-profile-trigger')?.addEventListener('click', () => {
        document.getElementById('profile-modal').style.display = 'block';
    });

    document.getElementById('change-pass-form')?.addEventListener('submit', (e) => {
        e.preventDefault();
        const cur = document.getElementById('current-pass').value;
        const n1 = document.getElementById('new-pass').value;
        const n2 = document.getElementById('confirm-pass').value;

        if (cur !== currentUser.password) {
            alert('⚠️ La contraseña actual es incorrecta.');
            return;
        }
        if (n1 !== n2) {
            alert('⚠️ Las nuevas contraseñas no coinciden.');
            return;
        }

        usuarios_siga[currentUser.username].password = n1;
        currentUser.password = n1;
        localStorage.setItem('usuarios_siga', JSON.stringify(usuarios_siga));
        sessionStorage.setItem('siga_session', JSON.stringify(currentUser));
        
        alert('✅ Contraseña actualizada correctamente.');
        document.getElementById('profile-modal').style.display = 'none';
        e.target.reset();
    });

    // --- User Management ---
    document.getElementById('add-user-btn')?.addEventListener('click', () => {
        document.getElementById('user-form').reset();
        document.getElementById('u-username').disabled = false;
        document.getElementById('user-edit-id').value = '';
        document.getElementById('user-modal-title').innerText = 'Crear Nuevo Usuario';
        document.getElementById('user-modal').style.display = 'block';
    });

    document.getElementById('user-form')?.addEventListener('submit', async (e) => {
        e.preventDefault();
        const editId = document.getElementById('user-edit-id').value;
        const username = document.getElementById('u-username').value;
        
        if (!editId && usuarios_siga[username]) {
            alert('⚠️ El nombre de usuario ya existe.');
            return;
        }

        const userData = {
            username: username,
            password: document.getElementById('u-password').value,
            nombre: document.getElementById('u-nombre').value.toUpperCase(),
            apellido: document.getElementById('u-apellido').value.toUpperCase(),
            jerarquia: document.getElementById('u-jerarquia').value,
            rol: document.getElementById('u-rol').value
        };

        try {
            await db.collection("usuarios").doc(username).set(userData);
            document.getElementById('user-modal').style.display = 'none';
        } catch (error) {
            console.error("Error al guardar usuario:", error);
            alert("Error al guardar en la nube.");
        }
    });

    document.getElementById('close-profile').onclick = () => document.getElementById('profile-modal').style.display = 'none';
    document.getElementById('close-user-modal').onclick = () => document.getElementById('user-modal').style.display = 'none';

    window.onclick = (event) => {
        if (event.target.className === 'modal') event.target.style.display = 'none';
    };

    // --- Importación Excel ---
    document.getElementById('import-excel')?.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (evt) => {
            const data = evt.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

            rows.forEach(row => {
                const p = {
                    jerarquia: row.Jerarquia || row.JERARQUIA || 'Agente',
                    nombre: (row.Nombre || row.NOMBRE || row['Apellido y Nombre'] || '').toUpperCase(),
                    legajo: row.LP || row.Legajo || '',
                    dni: row.DNI || '',
                    gmail: row.Gmail || row.GMAIL || '',
                    os: row['Obra Social'] || row.OS || '',
                    csjpfa: row.CSJPFA || '',
                    armamento: row.Armamento || row.ARMA || ''
                };
                if (p.nombre) personal.push(p);
            });
            savePersonal();
            alert('Personal importado exitosamente.');
        };
        reader.readAsBinaryString(file);
    });

    // --- Exportación Guardia (Mensual) ---
    document.getElementById('export-guardia-btn')?.addEventListener('click', () => {
        const monthName = MESES[currentMonth].toUpperCase();
        const data = [[`CONTROL DE GUARDIA INSTITUCIONAL - ${monthName} ${currentYear}`]];
        data.push(['FECHA', 'JERARQUÍA', 'NOMBRE Y APELLIDO', 'LP', 'DNI']);

        const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
        
        for (let d = 1; d <= daysInMonth; d++) {
            const dateKey = `${currentYear}-${(currentMonth + 1).toString().padStart(2, '0')}-${d.toString().padStart(2, '0')}`;
            const displayDate = `${d}/${currentMonth + 1}/${currentYear}`;
            const asig = guardias[dateKey] || [];
            
            if (asig.length === 0) {
                data.push([displayDate, '-', 'SIN ASIGNAR', '-', '-']);
            } else {
                asig.forEach(n => {
                    const p = personal.find(pers => pers.nombre === n) || {};
                    data.push([displayDate, p.jerarquia || '-', n, p.legajo || '-', p.dni || '-']);
                });
            }
        }

        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Resumen_Mensual");
        XLSX.writeFile(wb, `SIGA_Guardia_${monthName}_${currentYear}.xlsx`);
    });

    // --- Exportación Resumen de Servicios ---
    document.getElementById('export-servicios-resumen-btn')?.addEventListener('click', () => {
        const data = [['RESUMEN GENERAL DE SERVICIOS OPERATIVOS EXTERNOS']];
        data.push(['FECHA', 'MOTIVO', 'LUGAR', 'HORARIO', 'PERSONAL ASIGNADO']);
        
        serviciosExternos.forEach(s => {
            data.push([s.fecha, s.motivo, s.lugar, s.hora, s.asignados.join(', ')]);
        });

        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Resumen_Servicios");
        XLSX.writeFile(wb, "SIGA_Resumen_Servicios_Externos.xlsx");
    });
}
