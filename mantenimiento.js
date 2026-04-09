const API_BASE_CANDIDATES = resolveApiBaseCandidates();
let ui = null;
let currentUser = null;
let currentMaintenance = null;
let updateInProgress = false;
let authTokenMemory = '';

document.addEventListener('DOMContentLoaded', () => {
    ui = {
        superAccessSection: document.getElementById('superAccessSection'),
        superPanel: document.getElementById('superPanel'),
        loginForm: document.getElementById('superLoginForm'),
        username: document.getElementById('superUsername'),
        password: document.getElementById('superPassword'),
        loginMessage: document.getElementById('superLoginMessage'),
        maintenanceState: document.getElementById('maintenanceState'),
        maintenanceUpdatedAt: document.getElementById('maintenanceUpdatedAt'),
        maintenanceUpdatedBy: document.getElementById('maintenanceUpdatedBy'),
        panelMessage: document.getElementById('maintenancePanelMessage'),
        enableBtn: document.getElementById('enableMaintenanceBtn'),
        disableBtn: document.getElementById('disableMaintenanceBtn'),
        refreshBtn: document.getElementById('refreshMaintenanceBtn'),
        logoutBtn: document.getElementById('superLogoutBtn')
    };

    if (!ui.loginForm) {
        return;
    }

    ui.loginForm.addEventListener('submit', (event) => {
        event.preventDefault();
        void handleSuperLogin();
    });

    ui.enableBtn?.addEventListener('click', () => {
        void updateMaintenanceMode(true);
    });

    ui.disableBtn?.addEventListener('click', () => {
        void updateMaintenanceMode(false);
    });

    ui.refreshBtn?.addEventListener('click', () => {
        void refreshMaintenanceStatus();
    });

    ui.logoutBtn?.addEventListener('click', () => {
        void logoutSuperSession();
    });

    resetMaintenanceView();
    void bootstrap();
});

async function bootstrap() {
    try {
        const payload = await apiRequest('/auth/me', { method: 'GET' });
        currentUser = payload?.user || null;
    } catch (error) {
        currentUser = null;
    }

    if (!isSuper(currentUser)) {
        showLoginSection();
        return;
    }

    showSuperPanel();
    await refreshMaintenanceStatus();
}

async function handleSuperLogin() {
    const username = normalizeText(ui?.username?.value).toLowerCase();
    const password = String(ui?.password?.value || '');

    if (!username || !password) {
        setMessage(ui?.loginMessage, 'Ingresa usuario y contrasena.', 'error');
        return;
    }

    setMessage(ui?.loginMessage, 'Validando acceso super...', '');
    setControlsBusy(true);

    try {
        const payload = await apiRequest('/auth/login', {
            method: 'POST',
            body: JSON.stringify({ username, password })
        });

        authTokenMemory = normalizeText(payload?.token || '');
        currentUser = payload?.user || null;
        if (!isSuper(currentUser)) {
            authTokenMemory = '';
            await safeLogout();
            currentUser = null;
            showLoginSection();
            setMessage(ui?.loginMessage, 'La cuenta no tiene rol SUPER.', 'error');
            return;
        }

        if (ui?.loginForm) {
            ui.loginForm.reset();
        }

        showSuperPanel();
        setMessage(ui?.loginMessage, 'Sesion super iniciada correctamente.', 'success');
        await refreshMaintenanceStatus();
    } catch (error) {
        currentUser = null;
        showLoginSection();
        setMessage(ui?.loginMessage, error.message || 'No fue posible iniciar sesion.', 'error');
    } finally {
        setControlsBusy(false);
    }
}

async function logoutSuperSession() {
    setControlsBusy(true);
    await safeLogout();
    authTokenMemory = '';
    currentUser = null;
    currentMaintenance = null;
    resetMaintenanceView();
    showLoginSection();
    setMessage(ui?.loginMessage, 'Sesion cerrada.', 'success');
    setControlsBusy(false);
}

async function refreshMaintenanceStatus() {
    if (!isSuper(currentUser)) {
        showLoginSection();
        setMessage(ui?.loginMessage, 'Solo el rol SUPER puede ver este panel.', 'error');
        return;
    }

    setControlsBusy(true);
    setMessage(ui?.panelMessage, 'Consultando estado actual...', '');

    try {
        const payload = await apiRequest('/system/mantenimiento', { method: 'GET' });
        currentMaintenance = payload?.maintenance || null;
        renderMaintenanceStatus();
        setMessage(ui?.panelMessage, 'Estado actualizado.', 'success');
    } catch (error) {
        setMessage(ui?.panelMessage, error.message || 'No fue posible leer el estado.', 'error');
    } finally {
        setControlsBusy(false);
    }
}

async function updateMaintenanceMode(enabled) {
    if (updateInProgress) {
        return;
    }

    if (!isSuper(currentUser)) {
        showLoginSection();
        setMessage(ui?.loginMessage, 'Solo el rol SUPER puede cambiar este estado.', 'error');
        return;
    }

    updateInProgress = true;
    setControlsBusy(true);
    syncToggleButtons();
    setMessage(ui?.panelMessage, enabled ? 'Activando mantencion...' : 'Volviendo a modo normal...', '');

    try {
        const payload = await apiRequest('/system/mantenimiento', {
            method: 'PUT',
            body: JSON.stringify({ enabled })
        });
        currentMaintenance = payload?.maintenance || { enabled };
        renderMaintenanceStatus();
        setMessage(ui?.panelMessage, normalizeText(payload?.message) || 'Estado actualizado.', 'success');
    } catch (error) {
        setMessage(ui?.panelMessage, error.message || 'No fue posible actualizar el estado.', 'error');
    } finally {
        updateInProgress = false;
        setControlsBusy(false);
        syncToggleButtons();
    }
}

function renderMaintenanceStatus() {
    const enabled = Boolean(currentMaintenance?.enabled);
    setText(ui?.maintenanceState, enabled ? 'ACTIVO' : 'INACTIVO');
    setText(ui?.maintenanceUpdatedBy, normalizeText(currentMaintenance?.updatedBy) || '-');
    setText(ui?.maintenanceUpdatedAt, formatDateTime(currentMaintenance?.updatedAt) || '-');
    syncToggleButtons();
}

function resetMaintenanceView() {
    setText(ui?.maintenanceState, 'SIN SESION');
    setText(ui?.maintenanceUpdatedBy, '-');
    setText(ui?.maintenanceUpdatedAt, '-');
    setMessage(ui?.panelMessage, '', '');
    syncToggleButtons();
}

function syncToggleButtons() {
    const enabled = Boolean(currentMaintenance?.enabled);
    if (ui?.enableBtn) {
        ui.enableBtn.disabled = updateInProgress || enabled;
    }
    if (ui?.disableBtn) {
        ui.disableBtn.disabled = updateInProgress || !enabled;
    }
    if (ui?.refreshBtn) {
        ui.refreshBtn.disabled = updateInProgress;
    }
}

function showLoginSection() {
    ui?.superAccessSection?.classList.remove('hidden');
    ui?.superPanel?.classList.add('hidden');
}

function showSuperPanel() {
    ui?.superAccessSection?.classList.add('hidden');
    ui?.superPanel?.classList.remove('hidden');
}

function setControlsBusy(isBusy) {
    const disabled = Boolean(isBusy);

    if (ui?.loginForm) {
        const controls = ui.loginForm.querySelectorAll('input, button');
        controls.forEach((control) => {
            control.disabled = disabled;
        });
    }

    if (ui?.logoutBtn) {
        ui.logoutBtn.disabled = disabled;
    }
    if (ui?.enableBtn) {
        ui.enableBtn.disabled = disabled || Boolean(currentMaintenance?.enabled) || updateInProgress;
    }
    if (ui?.disableBtn) {
        ui.disableBtn.disabled = disabled || !Boolean(currentMaintenance?.enabled) || updateInProgress;
    }
    if (ui?.refreshBtn) {
        ui.refreshBtn.disabled = disabled || updateInProgress;
    }
}

async function safeLogout() {
    try {
        await apiRequest('/auth/logout', { method: 'POST' });
    } catch (error) {
        return;
    }
}

async function apiRequest(path, options = {}) {
    const requestOptions = {
        method: options.method || 'GET',
        credentials: 'include',
        headers: {
            'Content-Type': 'application/json',
            ...(options.headers || {})
        }
    };

    if (authTokenMemory) {
        requestOptions.headers.Authorization = `Bearer ${authTokenMemory}`;
    }

    if (Object.prototype.hasOwnProperty.call(options, 'body')) {
        requestOptions.body = options.body;
    }

    let response = null;
    let payload = {};
    for (const apiBase of API_BASE_CANDIDATES) {
        try {
            response = await fetch(`${apiBase}${path}`, requestOptions);
            payload = await response.json().catch(() => ({}));
            break;
        } catch (error) {
            response = null;
        }
    }

    if (!response) {
        const fallbackLabel = API_BASE_CANDIDATES.join(', ');
        throw new Error(`No hay conexion con el servidor API. Se intento con: ${fallbackLabel}.`);
    }

    if (!response.ok) {
        const message = normalizeText(payload?.message) || `Error HTTP ${response.status}`;
        const error = new Error(message);
        error.status = response.status;
        error.payload = payload;
        throw error;
    }

    return payload;
}

function isSuper(user) {
    return normalizeText(user?.role).toUpperCase() === 'SUPER';
}

function formatDateTime(value) {
    if (!value) {
        return '';
    }

    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) {
        return normalizeText(value);
    }

    return parsed.toLocaleString('es-CL', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    });
}

function setMessage(element, text, type = '') {
    if (!element) {
        return;
    }

    element.textContent = String(text || '');
    element.className = 'maintenance-message';
    if (type) {
        element.classList.add(type);
    }
}

function setText(element, text) {
    if (!element) {
        return;
    }

    element.textContent = String(text || '');
}

function normalizeText(value) {
    return String(value || '').trim();
}

function resolveApiBaseCandidates() {
    const candidates = [];
    const appendCandidate = (value) => {
        const normalized = normalizeApiBase(value);
        if (!normalized || candidates.includes(normalized)) {
            return;
        }
        candidates.push(normalized);
    };

    appendCandidate(resolveApiBase());
    appendCandidate('/api');

    const host = String(window.location.hostname || '').trim().toLowerCase();
    const isLocalhost = host === 'localhost' || host === '127.0.0.1';
    if (isLocalhost) {
        appendCandidate(`http://${host}:3000/api`);
        if (host === 'localhost') {
            appendCandidate('http://127.0.0.1:3000/api');
        } else if (host === '127.0.0.1') {
            appendCandidate('http://localhost:3000/api');
        }
    }

    return candidates;
}

function resolveApiBase() {
    const host = String(window.location.hostname || '').trim().toLowerCase();
    const port = String(window.location.port || '').trim();
    const protocol = String(window.location.protocol || '').trim().toLowerCase();
    if (host && !host.includes(':') && port !== '3000' && protocol !== 'https:') {
        return `http://${host}:3000/api`;
    }

    return '/api';
}

function normalizeApiBase(value) {
    const text = String(value || '').trim();
    if (!text) {
        return '';
    }
    return text.replace(/\/+$/, '');
}
