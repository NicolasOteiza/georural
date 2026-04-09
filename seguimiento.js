const API_BASE_CANDIDATES = resolveApiBaseCandidates();
let ui = null;
let searchInProgress = false;

document.addEventListener('DOMContentLoaded', async () => {
    ui = {
        form: document.getElementById('publicProgressForm'),
        rol: document.getElementById('publicProgressRol'),
        searchBtn: document.getElementById('publicProgressSearchBtn'),
        loading: document.getElementById('publicProgressLoading'),
        message: document.getElementById('publicProgressMessage'),
        resultSection: document.getElementById('publicProgressResultSection'),
        resultRol: document.getElementById('publicProgressResultRol'),
        resultNombre: document.getElementById('publicProgressNombre'),
        resultCorreo: document.getElementById('publicProgressCorreo'),
        historyBody: document.getElementById('publicProgressHistoryBody'),
        loginRouteNoticeOverlay: document.getElementById('loginRouteNoticeOverlay'),
        loginRouteNoticeUrl: document.getElementById('loginRouteNoticeUrl'),
        loginRouteNoticeExpire: document.getElementById('loginRouteNoticeExpire'),
        loginRouteNoticeOpenBtn: document.getElementById('loginRouteNoticeOpenBtn'),
        loginRouteNoticeCloseBtn: document.getElementById('loginRouteNoticeCloseBtn')
    };

    if (!ui.form) {
        return;
    }
    if (!ui.searchBtn) {
        ui.searchBtn = ui.form.querySelector('button[type="submit"]');
    }

    const systemStatusPayload = await loadPublicSystemStatus();
    const redirectedToMaintenance = await enforceMaintenanceRedirectIfNeeded(systemStatusPayload);
    if (redirectedToMaintenance) {
        return;
    }

    applyLoginRouteNoticePopup(systemStatusPayload);

    setSearchBusyState(false);
    renderHistory([]);
    hideResultSection();
    clearMessage();

    ui.form.addEventListener('submit', (event) => {
        event.preventDefault();
        void handleSearchSubmit();
    });
    if (ui.loginRouteNoticeCloseBtn) {
        ui.loginRouteNoticeCloseBtn.addEventListener('click', hideLoginRouteNoticePopup);
    }
    if (ui.loginRouteNoticeOverlay) {
        ui.loginRouteNoticeOverlay.addEventListener('click', (event) => {
            if (event.target === ui.loginRouteNoticeOverlay) {
                hideLoginRouteNoticePopup();
            }
        });
    }
});

async function handleSearchSubmit() {
    if (searchInProgress) {
        return;
    }

    const rol = normalizeText(ui?.rol?.value);
    if (!rol) {
        setMessage('Ingresa el ROL para buscar el progreso.', 'error');
        hideResultSection();
        return;
    }

    await executeProgressSearch(rol);
}

async function executeProgressSearch(rol) {
    searchInProgress = true;
    setSearchBusyState(true);
    setMessage('Buscando progreso del ROL...', '');
    hideResultSection();

    try {
        const result = await fetchPublicProgressByRol(rol);
        showResultSection(result || {});
        renderHistory(Array.isArray(result?.historial) ? result.historial : []);
        setMessage('Progreso cargado correctamente.', 'success');
    } catch (error) {
        hideResultSection();
        setMessage(error.message || 'No fue posible buscar el progreso.', 'error');
    } finally {
        searchInProgress = false;
        setSearchBusyState(false);
    }
}

async function fetchPublicProgressByRol(rol) {
    const query = new URLSearchParams();
    query.set('rol', normalizeText(rol));
    return apiRequest(`/public/registros/progreso?${query.toString()}`);
}

async function loadPublicSystemStatus() {
    try {
        return await apiRequest('/public/system-status');
    } catch (error) {
        return null;
    }
}

function normalizeLoginNoticeUrl(value) {
    const raw = normalizeText(value);
    if (!raw) {
        return '';
    }
    if (/^https?:\/\//i.test(raw)) {
        return raw;
    }
    if (/^localhost(?:\:\d+)?\//i.test(raw) || /^127\.0\.0\.1(?:\:\d+)?\//i.test(raw)) {
        return `http://${raw}`;
    }
    if (/^[a-z0-9.-]+\.[a-z]{2,}(?:\/|$)/i.test(raw)) {
        return `https://${raw}`;
    }
    if (raw.startsWith('/')) {
        return raw;
    }
    return `/${raw.replace(/^\/+/, '')}`;
}

function resolveDisplayUrl(url) {
    const normalizedUrl = normalizeLoginNoticeUrl(url);
    if (!normalizedUrl) {
        return '';
    }
    if (/^https?:\/\//i.test(normalizedUrl)) {
        try {
            const parsed = new URL(normalizedUrl);
            return `${parsed.host}${parsed.pathname}${parsed.search}${parsed.hash}`;
        } catch (error) {
            return normalizedUrl.replace(/^https?:\/\//i, '');
        }
    }
    return `${window.location.host}${normalizedUrl}`;
}

function applyLoginRouteNoticePopup(statusPayload = null) {
    if (!ui?.loginRouteNoticeOverlay) {
        return;
    }

    const notice = statusPayload?.loginRouteNotice;
    if (!notice || !notice.active) {
        hideLoginRouteNoticePopup();
        return;
    }

    const targetUrl = normalizeLoginNoticeUrl(notice.url);
    if (!targetUrl) {
        hideLoginRouteNoticePopup();
        return;
    }

    const urlForDisplay = resolveDisplayUrl(targetUrl);
    if (ui.loginRouteNoticeUrl) {
        ui.loginRouteNoticeUrl.textContent = urlForDisplay || targetUrl;
    }
    if (ui.loginRouteNoticeOpenBtn) {
        ui.loginRouteNoticeOpenBtn.href = targetUrl;
    }
    if (ui.loginRouteNoticeExpire) {
        const daysRemaining = Number(notice?.daysRemaining || 0);
        if (daysRemaining > 0) {
            ui.loginRouteNoticeExpire.textContent = `Aviso vigente por ${daysRemaining} dia${daysRemaining === 1 ? '' : 's'} mas.`;
        } else if (notice?.expiresAt) {
            ui.loginRouteNoticeExpire.textContent = `Aviso vigente hasta ${formatDateTime(notice.expiresAt)}.`;
        } else {
            ui.loginRouteNoticeExpire.textContent = '';
        }
    }

    ui.loginRouteNoticeOverlay.classList.remove('hidden');
    ui.loginRouteNoticeOverlay.setAttribute('aria-hidden', 'false');
}

function hideLoginRouteNoticePopup() {
    if (!ui?.loginRouteNoticeOverlay) {
        return;
    }
    ui.loginRouteNoticeOverlay.classList.add('hidden');
    ui.loginRouteNoticeOverlay.setAttribute('aria-hidden', 'true');
}

async function enforceMaintenanceRedirectIfNeeded(statusPayload = null) {
    try {
        const payload = statusPayload && typeof statusPayload === 'object' ? statusPayload : await apiRequest('/public/system-status');
        const maintenanceEnabled = Boolean(payload?.maintenance?.enabled);
        if (!maintenanceEnabled) {
            return false;
        }

        try {
            const mePayload = await apiRequest('/auth/me');
            const role = normalizeText(mePayload?.user?.role).toUpperCase();
            if (role === 'SUPER') {
                return false;
            }
        } catch (error) {
            // Sin sesion valida; se aplica redireccion.
        }

        window.location.href = 'mantenimiento.html';
        return true;
    } catch (error) {
        return false;
    }
}

async function apiRequest(path) {
    const normalizedPath = String(path || '');
    let response = null;
    let payload = {};

    for (const apiBase of API_BASE_CANDIDATES) {
        try {
            response = await fetch(`${apiBase}${normalizedPath}`, {
                method: 'GET',
                credentials: 'include',
                headers: {
                    'Content-Type': 'application/json'
                }
            });
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
        const message = normalizeText(payload?.message) || 'No fue posible procesar la solicitud.';
        throw new Error(message);
    }

    return payload;
}

function showResultSection(result = {}) {
    if (!ui?.resultSection) {
        return;
    }

    setFieldValue(ui.resultRol, result.rol);
    setFieldValue(ui.resultNombre, result.nombre);
    setFieldValue(ui.resultCorreo, result.correo);
    ui.resultSection.classList.remove('hidden');
}

function hideResultSection() {
    if (!ui?.resultSection) {
        return;
    }

    setFieldValue(ui.resultRol, '');
    setFieldValue(ui.resultNombre, '');
    setFieldValue(ui.resultCorreo, '');
    renderHistory([]);
    ui.resultSection.classList.add('hidden');
}

function renderHistory(historial = []) {
    if (!ui?.historyBody) {
        return;
    }

    ui.historyBody.innerHTML = '';
    if (!Array.isArray(historial) || historial.length === 0) {
        ui.historyBody.innerHTML = '<tr><td colspan="2">Sin historial de progreso para mostrar.</td></tr>';
        return;
    }

    historial.forEach((item) => {
        const row = document.createElement('tr');

        const fechaCell = document.createElement('td');
        fechaCell.textContent = formatDateTime(item?.fecha);

        const comentarioCell = document.createElement('td');
        comentarioCell.textContent = String(item?.comentario || '');

        row.appendChild(fechaCell);
        row.appendChild(comentarioCell);
        ui.historyBody.appendChild(row);
    });
}

function formatDateTime(value) {
    const parsed = parseDateValue(value);
    if (!parsed) {
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

function parseDateValue(value) {
    if (!value) {
        return null;
    }

    if (value instanceof Date) {
        return Number.isNaN(value.getTime()) ? null : value;
    }

    const text = String(value || '').trim();
    if (!text) {
        return null;
    }

    const sqlDateTimeMatch = /^(\d{4})-(\d{2})-(\d{2})[\sT](\d{2}):(\d{2})(?::(\d{2}))?$/.exec(text);
    if (sqlDateTimeMatch) {
        const year = Number(sqlDateTimeMatch[1]);
        const month = Number(sqlDateTimeMatch[2]);
        const day = Number(sqlDateTimeMatch[3]);
        const hour = Number(sqlDateTimeMatch[4]);
        const minute = Number(sqlDateTimeMatch[5]);
        const second = Number(sqlDateTimeMatch[6] || '0');
        const localDate = new Date(year, month - 1, day, hour, minute, second);
        if (!Number.isNaN(localDate.getTime())) {
            return localDate;
        }
    }

    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
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
    const metaValue = normalizeApiBase(document.querySelector('meta[name="geo-rural-api-base"]')?.content || '');
    if (metaValue) {
        return metaValue;
    }

    const host = String(window.location.hostname || '').trim().toLowerCase();
    const port = String(window.location.port || '').trim();
    const isLocalhost = host === 'localhost' || host === '127.0.0.1';
    if (isLocalhost && port !== '3000') {
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

function setFieldValue(field, value) {
    if (!field) {
        return;
    }
    field.value = String(value || '');
}

function setMessage(text, type = '') {
    if (!ui?.message) {
        return;
    }

    ui.message.textContent = String(text || '');
    ui.message.className = 'seguimiento-message';
    if (type) {
        ui.message.classList.add(type);
    }
}

function clearMessage() {
    setMessage('', '');
}

function setSearchBusyState(isBusy) {
    const busy = Boolean(isBusy);

    if (ui?.rol) {
        ui.rol.disabled = busy;
    }

    if (ui?.searchBtn) {
        if (!ui.searchBtn.dataset.defaultLabel) {
            ui.searchBtn.dataset.defaultLabel = String(ui.searchBtn.textContent || '').trim() || 'BUSCAR PROGRESO';
        }
        ui.searchBtn.disabled = busy;
        ui.searchBtn.classList.toggle('is-loading', busy);
        ui.searchBtn.textContent = busy ? 'BUSCANDO...' : ui.searchBtn.dataset.defaultLabel;
    }

    if (ui?.loading) {
        ui.loading.classList.toggle('hidden', !busy);
        ui.loading.setAttribute('aria-hidden', busy ? 'false' : 'true');
    }
}

function normalizeText(value) {
    return String(value || '').trim();
}
