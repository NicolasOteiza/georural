const express = require('express');
const cors = require('cors');
const mysql = require('mysql2/promise');
const http = require('http');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const https = require('https');
const dotenv = require('dotenv');
const nodemailer = require('nodemailer');
const PDFDocument = require('pdfkit');
const { PDFDocument: PDFLibDocument, StandardFonts, rgb } = require('pdf-lib');

function normalizeEnvironmentName(value) {
    const normalized = String(value || '')
        .trim()
        .toLowerCase();

    if (!normalized) {
        return 'development';
    }

    if (normalized === 'prod') {
        return 'production';
    }
    if (normalized === 'dev') {
        return 'development';
    }

    return normalized.replace(/[^a-z0-9_-]/g, '') || 'development';
}

function readEnvironmentArg(argv = process.argv.slice(2)) {
    for (const arg of Array.isArray(argv) ? argv : []) {
        const item = String(arg || '').trim();
        if (!item) {
            continue;
        }

        if (item.startsWith('--env=')) {
            return item.slice('--env='.length);
        }

        if (item.startsWith('--app-env=')) {
            return item.slice('--app-env='.length);
        }
    }

    return '';
}

function loadEnvironmentConfig() {
    const cliEnvValue = readEnvironmentArg();
    const appEnv = normalizeEnvironmentName(cliEnvValue || process.env.APP_ENV || process.env.NODE_ENV || 'development');
    const candidateFiles = ['.env', `.env.${appEnv}`];
    const loadedFiles = [];

    for (const fileName of candidateFiles) {
        const filePath = path.resolve(process.cwd(), fileName);
        if (!fs.existsSync(filePath)) {
            continue;
        }

        const result = dotenv.config({ path: filePath, override: true });
        if (!result.error) {
            loadedFiles.push(filePath);
        }
    }

    process.env.APP_ENV = appEnv;
    process.env.NODE_ENV = appEnv === 'production' ? 'production' : 'development';

    return {
        appEnv,
        loadedFiles
    };
}

const ENV_CONFIG = loadEnvironmentConfig();
const APP_ENV = ENV_CONFIG.appEnv;
const IS_PRODUCTION_ENV = APP_ENV === 'production';

const DB_HOST = process.env.DB_HOST || '127.0.0.1';
const DB_PORT = Number(process.env.DB_PORT || 3306);
const DB_USER = process.env.DB_USER || 'root';
const DB_PASSWORD = process.env.DB_PASSWORD || '';
const DB_NAME = process.env.DB_NAME || 'geo_rural';
const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || '127.0.0.1';
const APP_URL = normalizeText(process.env.APP_URL);
const API_WRITE_KEY = normalizeText(process.env.API_WRITE_KEY);
const IS_LOCALHOST_BIND = HOST === '127.0.0.1' || HOST === 'localhost' || HOST === '::1';
const AUTH_SESSION_HOURS = parsePositiveInt(process.env.AUTH_SESSION_HOURS, 12);
const LOGIN_MAX_ATTEMPTS = parsePositiveInt(process.env.LOGIN_MAX_ATTEMPTS, 5);
const LOGIN_WINDOW_MINUTES = parsePositiveInt(process.env.LOGIN_WINDOW_MINUTES, 15);
const LOGIN_LOCK_MINUTES = parsePositiveInt(process.env.LOGIN_LOCK_MINUTES, 2);
const MAX_PASSWORD_LENGTH = parsePositiveInt(process.env.MAX_PASSWORD_LENGTH, 200);
const SESSION_CLEANUP_INTERVAL_MINUTES = parsePositiveInt(process.env.SESSION_CLEANUP_INTERVAL_MINUTES, 30);
const ACTIVE_SESSION_WINDOW_MINUTES = parsePositiveInt(process.env.ACTIVE_SESSION_WINDOW_MINUTES, 20);
const SESSION_COOKIE_NAME = 'geo_rural_session';
const SESSION_COOKIE_MAX_AGE_SECONDS = AUTH_SESSION_HOURS * 60 * 60;
const AUTH_COOKIE_DOMAIN = normalizeText(process.env.AUTH_COOKIE_DOMAIN || '');
const AUTH_COOKIE_SAME_SITE = parseSameSiteValue(process.env.AUTH_COOKIE_SAME_SITE, 'Lax');
const AUTH_COOKIE_SECURE_MODE = parseCookieSecureMode(process.env.AUTH_COOKIE_SECURE, 'auto');
const DEFAULT_USERS = parseDefaultUsers(process.env.DEFAULT_USERS || '');
const ROLE_ADMIN = 'ADMIN';
const ROLE_SECRETARIA = 'SECRETARIA';
const ROLE_OPERADOR = 'OPERADOR';
const ROLE_SUPERVISOR = 'SUPERVISOR';
const ROLE_SUPER = 'SUPER';
const PRIVILEGED_ROLES = Object.freeze([ROLE_ADMIN, ROLE_SUPER]);
const ASSIGNABLE_ROLES_BY_ADMIN = Object.freeze([ROLE_ADMIN, ROLE_SECRETARIA, ROLE_OPERADOR, ROLE_SUPERVISOR]);
const ASSIGNABLE_ROLES_BY_SUPER = Object.freeze([ROLE_SUPER, ROLE_ADMIN, ROLE_SECRETARIA, ROLE_OPERADOR, ROLE_SUPERVISOR]);
const SUPERUSER_USERNAME = normalizeText(process.env.SUPERUSER_USERNAME || 'superusuario').toLowerCase();
const SUPERUSER_PASSWORD = String(process.env.SUPERUSER_PASSWORD || '');
const SUPERUSER_NAME = normalizeText(process.env.SUPERUSER_NAME || 'Super Usuario') || 'Super Usuario';
const SUPERUSER_BRANCH = normalizeText(process.env.SUPERUSER_BRANCH || 'Casa Matriz') || 'Casa Matriz';
const GUEST_USERNAME = normalizeText(process.env.GUEST_USERNAME || 'invitado').toLowerCase() || 'invitado';
const GUEST_DISPLAY_NAME = normalizeText(process.env.GUEST_NAME || 'Invitado') || 'Invitado';
const GUEST_BRANCH_NAME = normalizeText(process.env.GUEST_BRANCH || 'Casa Matriz') || 'Casa Matriz';
const DEFAULT_CREATION_COMMENT = 'creacion y recepcion al iniciar un nuevo registro';
const SECRETARIA_NO_MOVIMIENTO_DAYS = parsePositiveInt(process.env.SECRETARIA_NO_MOVIMIENTO_DAYS, 8);
const FACTURA_SOLICITUD_ESTADO_PENDIENTE = 'PENDIENTE';
const FACTURA_SOLICITUD_ESTADO_ENVIADA = 'ENVIADA';
const FACTURA_SOLICITUD_ESTADO_ANULADA = 'ANULADA';
const SIMPLE_EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const HISTORY_DETAIL_MAX_CHANGES = parsePositiveInt(process.env.HISTORY_DETAIL_MAX_CHANGES, 4);
const HISTORY_SUMMARY_MAX_LABELS = parsePositiveInt(process.env.HISTORY_SUMMARY_MAX_LABELS, 5);
const UTM_SOURCE_URL = normalizeText(process.env.UTM_SOURCE_URL || 'https://mindicador.cl/api/utm');
const UTM_SOURCE_TIMEOUT_MS = parsePositiveInt(process.env.UTM_SOURCE_TIMEOUT_MS, 5000);
const UTM_SOURCE_CACHE_MINUTES = parsePositiveInt(process.env.UTM_SOURCE_CACHE_MINUTES, 30);
const UTM_SOURCE_ERROR_CACHE_MINUTES = parsePositiveInt(process.env.UTM_SOURCE_ERROR_CACHE_MINUTES, 3);
const COTIZACION_DOCUMENT_PATH = '/assets/Carta de Cotizacion Geo Rural.pdf';
const SMTP_HOST = normalizeText(process.env.SMTP_HOST || '');
const SMTP_PORT = parsePositiveInt(process.env.SMTP_PORT, 587);
const SMTP_USER = normalizeText(process.env.SMTP_USER || '');
const SMTP_PASSWORD = String(process.env.SMTP_PASSWORD || '');
const SMTP_SECURE = parseBooleanFlag(process.env.SMTP_SECURE, false);
const SMTP_FROM_EMAIL = normalizeText(process.env.SMTP_FROM_EMAIL || SMTP_USER).toLowerCase();
const SMTP_FROM_NAME = normalizeText(process.env.SMTP_FROM_NAME || 'Geo Rural');
const MAIL_TEMPLATE_MAX_LENGTH = parsePositiveInt(process.env.MAIL_TEMPLATE_MAX_LENGTH, 400000);
const MAIL_TEMPLATE_KEYS = Object.freeze(['cotizacionHtml', 'facturaSingleHtml', 'facturaPendingHtml']);
const SMTP_CONFIG_SINGLETON_ID = 1;
const SYSTEM_CONFIG_SINGLETON_ID = 1;
const LOGIN_ROUTE_NOTICE_DURATION_DAYS = parsePositiveInt(process.env.LOGIN_ROUTE_NOTICE_DURATION_DAYS, 15);
const LOGIN_ROUTE_NOTICE_MAX_URL_LENGTH = parsePositiveInt(process.env.LOGIN_ROUTE_NOTICE_MAX_URL_LENGTH, 255);
const LOGIN_ROUTE_NOTICE_DEV_URL = normalizeText(
    process.env.LOGIN_ROUTE_NOTICE_DEV_URL || 'http://localhost/geo_rural/registro/login/registro.html'
);
const LOGIN_ROUTE_NOTICE_PROD_URL = normalizeText(
    process.env.LOGIN_ROUTE_NOTICE_PROD_URL || 'https://mi-registro.cl/geo_rural/registro/login/registro.html'
);
const LOGIN_ROUTE_NOTICE_DEFAULT_URL = normalizeText(
    process.env.LOGIN_ROUTE_NOTICE_DEFAULT_URL || (IS_PRODUCTION_ENV ? LOGIN_ROUTE_NOTICE_PROD_URL : LOGIN_ROUTE_NOTICE_DEV_URL)
);
const LOGIN_ROUTE_NOTICE_DEFAULT_MESSAGE = normalizeText(
    process.env.LOGIN_ROUTE_NOTICE_DEFAULT_MESSAGE ||
        'Pedimos disculpas, pero por cambios de seguridad y para habilitar la vista del cliente se ha modificado la ruta de acceso a las cuentas de usuarios.'
);
const SMTP_ENV_CONFIG = Object.freeze({
    host: SMTP_HOST,
    port: SMTP_PORT,
    secure: SMTP_SECURE,
    user: SMTP_USER,
    password: SMTP_PASSWORD,
    fromEmail: SMTP_FROM_EMAIL,
    fromName: SMTP_FROM_NAME
});
const DEFAULT_COTIZACION_SERVICIOS = Object.freeze([
    {
        nombreServicio: 'Servicio estandar Geo Rural',
        valorUtm: 1
    }
]);
const COTIZACION_TEMPLATE_ROWS = Object.freeze([
    { lotes: '1 a 6', valorUtm: 5.33, y: 572.54 },
    { lotes: '7 a 15', valorUtm: 6.33, y: 546.86 },
    { lotes: '16 a 25', valorUtm: 7.33, y: 521.18 },
    { lotes: '26 a 35', valorUtm: 8.33, y: 495.5 },
    { lotes: '36 a 45', valorUtm: 9.33, y: 469.82 },
    { lotes: '46 a 55', valorUtm: 10.33, y: 443.9 },
    { lotes: '56 a 65', valorUtm: 11.33, y: 418.22 },
    { lotes: '66 a 75', valorUtm: 12.33, y: 392.54 },
    { lotes: '76 a 85', valorUtm: 13.33, y: 366.86 },
    { lotes: '86 a 95', valorUtm: 14.33, y: 341.18 },
    { lotes: '96 a 105', valorUtm: 15.33, y: 315.5 }
]);
const COTIZACION_TABLE_COLUMNS = Object.freeze([
    { key: 'lotes', title: 'TRAMO LOTES', width: 164, align: 'left' },
    { key: 'utm', title: 'VALOR UTM', width: 84, align: 'right' },
    { key: 'neto', title: 'NETO CLP', width: 84, align: 'right' },
    { key: 'iva', title: 'IVA CLP', width: 74, align: 'right' },
    { key: 'total', title: 'TOTAL CLP', width: 86, align: 'right' }
]);
const COTIZACION_TEMPLATE_CLEAR_AREAS = Object.freeze({
    utm: { x: 120, y: 634, width: 458, height: 48 },
    table: { x: 24, y: 260, width: 546, height: 360 }
});
const MONTH_NAMES_ES = Object.freeze([
    'enero',
    'febrero',
    'marzo',
    'abril',
    'mayo',
    'junio',
    'julio',
    'agosto',
    'septiembre',
    'octubre',
    'noviembre',
    'diciembre'
]);

const DEFAULT_CORS_ORIGINS = [APP_URL, `http://localhost:${PORT}`, `http://127.0.0.1:${PORT}`, 'http://localhost', 'http://127.0.0.1']
    .map((item) => normalizeText(item))
    .filter((item) => item.length > 0);
const ALLOWED_ORIGINS = parseCsv(process.env.CORS_ORIGINS, DEFAULT_CORS_ORIGINS);
const FRONTEND_FILES = [
    'index.html',
    'script.js',
    'styles.css',
    'seguimiento.html',
    'seguimiento.js',
    'seguimiento.css',
    'mantenimiento.html',
    'mantenimiento.css',
    'mantenimiento.js',
    'registro/login/registro.html',
    'logo.png',
    'logo.svg'
];
let runtimeBootstrapAdmin = null;
const loginAttempts = new Map();
let lastSessionCleanupAt = 0;
let utmSourceCache = {
    key: '',
    value: null,
    expiresAt: 0
};
let smtpTransporter = null;
let smtpTransporterKey = '';
let maintenanceModeCache = {
    enabled: false,
    updatedAt: '',
    updatedBy: ''
};
let loginRouteNoticeCache = {
    enabled: false,
    url: LOGIN_ROUTE_NOTICE_DEFAULT_URL,
    message: LOGIN_ROUTE_NOTICE_DEFAULT_MESSAGE,
    startsAt: '',
    expiresAt: '',
    updatedAt: '',
    updatedBy: ''
};

const app = express();
let pool;
const MAINTENANCE_ALLOWED_PAGE_PATHS = new Set([
    '/mantenimiento',
    '/mantenimiento/',
    '/mantenimiento.html',
    '/mantenimiento.css',
    '/mantenimiento.js',
    '/logo.png',
    '/logo.svg',
    '/favicon.ico'
]);
const MAINTENANCE_ALLOWED_API_PATHS = new Set([
    '/api/health',
    '/api/auth/login',
    '/api/auth/logout',
    '/api/auth/me',
    '/api/public/system-status',
    '/api/system/mantenimiento'
]);
const PUBLIC_FRONTEND_PATHS = new Set([
    '/',
    '/index.html',
    '/seguimiento',
    '/seguimiento.html',
    '/mantenimiento',
    '/mantenimiento/',
    '/mantenimiento.html',
    '/mantenimiento.css',
    '/mantenimiento.js',
    '/registro',
    '/registro/',
    '/registro/login',
    '/registro/login/',
    '/registro/login/registro.html',
    '/script.js',
    '/styles.css',
    '/seguimiento.js',
    '/seguimiento.css',
    '/logo.png',
    '/logo.svg',
    '/favicon.ico'
]);
const KNOWN_FRONTEND_PATHS = new Set(
    [
        '/',
        '/index.html',
        '/seguimiento',
        '/registro',
        '/registro/',
        '/registro/login',
        '/registro/login/',
        '/mantenimiento',
        '/mantenimiento/'
    ].concat(FRONTEND_FILES.map((fileName) => `/${fileName}`))
);
const LOGIN_REDIRECT_ENTRY_PATHS = new Set(['/registro', '/registro/login']);

app.use(
    cors({
        origin(origin, callback) {
            if (!origin) {
                return callback(null, true);
            }

            if (isAllowedOrigin(origin)) {
                return callback(null, true);
            }

            return callback(null, false);
        },
        methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
        allowedHeaders: ['Content-Type', 'X-API-Key', 'Authorization'],
        credentials: true,
        maxAge: 600
    })
);
app.use(express.json({ limit: '1mb' }));

function parsePositiveInt(value, fallback) {
    const parsed = Number(value);
    return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
}

function parseBooleanFlag(value, fallback = false) {
    const normalized = String(value || '')
        .trim()
        .toLowerCase();
    if (['1', 'true', 'yes', 'on', 'si', 's'].includes(normalized)) {
        return true;
    }
    if (['0', 'false', 'no', 'off', 'n'].includes(normalized)) {
        return false;
    }

    return Boolean(fallback);
}

function parseMaintenanceModeInput(value) {
    if (value === true) {
        return true;
    }
    if (value === false) {
        return false;
    }

    const normalized = String(value || '')
        .trim()
        .toLowerCase();
    if (['1', 'true', 'yes', 'on', 'si', 's'].includes(normalized)) {
        return true;
    }
    if (['0', 'false', 'no', 'off', 'n'].includes(normalized)) {
        return false;
    }

    return null;
}

function parseSameSiteValue(value, fallback = 'Lax') {
    const normalized = String(value || '')
        .trim()
        .toLowerCase();
    if (normalized === 'strict') {
        return 'Strict';
    }
    if (normalized === 'none') {
        return 'None';
    }
    if (normalized === 'lax') {
        return 'Lax';
    }

    return fallback;
}

function parseCookieSecureMode(value, fallback = 'auto') {
    const normalized = String(value || '')
        .trim()
        .toLowerCase();
    if (['always', 'on', 'true', '1', 'yes'].includes(normalized)) {
        return 'always';
    }
    if (['never', 'off', 'false', '0', 'no'].includes(normalized)) {
        return 'never';
    }
    if (normalized === 'auto') {
        return 'auto';
    }

    return fallback;
}

function getStartupHints(error) {
    if (!error || typeof error !== 'object') {
        return [];
    }

    if (error.code === 'ECONNREFUSED' && error.syscall === 'connect') {
        return [
            `No hay un servidor MySQL escuchando en ${DB_HOST}:${DB_PORT}.`,
            'Inicia MySQL desde XAMPP (boton Start en el modulo MySQL).',
            'Si usas otro host o puerto, ajusta DB_HOST y DB_PORT en .env.'
        ];
    }

    if (error.code === 'ER_ACCESS_DENIED_ERROR') {
        return [
            `MySQL rechazo las credenciales para ${DB_USER}@${DB_HOST}:${DB_PORT}.`,
            'Revisa DB_USER y DB_PASSWORD en .env.'
        ];
    }

    if (error.code === 'ER_DBACCESS_DENIED_ERROR' || error.code === 'ER_SPECIFIC_ACCESS_DENIED_ERROR') {
        return [
            `El usuario MySQL "${DB_USER}" no tiene permisos suficientes sobre la base "${DB_NAME}".`,
            'Otorga permisos sobre esa base o crea el esquema manualmente con database/schema.sql.'
        ];
    }

    if (error.code === 'ER_BAD_DB_ERROR') {
        return [
            `No se encontro la base "${DB_NAME}" y no se pudo crear automaticamente.`,
            'Verifica permisos del usuario MySQL o crea el esquema manualmente con database/schema.sql.'
        ];
    }

    if (error.code === 'EADDRINUSE' && error.syscall === 'listen') {
        return [
            `El puerto ${PORT} ya esta en uso por otro proceso.`,
            'Cierra el proceso que usa ese puerto o cambia PORT en .env.',
            'En Windows puedes usar: netstat -ano | findstr :3000'
        ];
    }

    if (error.code === 'EACCES' && error.syscall === 'listen') {
        return [
            `No hay permisos para abrir ${HOST}:${PORT}.`,
            'Ejecuta con permisos adecuados o usa un puerto permitido (por ejemplo >= 1024).'
        ];
    }

    return [];
}

function cleanDbIdentifier(value) {
    return String(value || '').replace(/[^a-zA-Z0-9_]/g, '');
}

function normalizeText(value) {
    return String(value || '').trim();
}

function escapeHtmlText(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function resolveErrorPageTitle(statusCode) {
    const safeCode = Number.isInteger(statusCode) ? statusCode : 500;
    if (safeCode === 401) {
        return 'ACCESO NO AUTORIZADO';
    }
    if (safeCode === 403) {
        return 'SOLO PERSONAL AUTORIZADO';
    }
    if (safeCode === 404) {
        return 'RUTA EXTRAVIADA';
    }
    if (safeCode === 503) {
        return 'SERVICIO TEMPORALMENTE NO DISPONIBLE';
    }
    return normalizeText(http.STATUS_CODES[safeCode]).toUpperCase() || `ERROR ${safeCode}`;
}

function resolveErrorPageDetail(statusCode, message = '') {
    const customMessage = normalizeText(message);
    if (customMessage) {
        return customMessage;
    }

    const safeCode = Number.isInteger(statusCode) ? statusCode : 500;
    if (safeCode === 401) {
        return 'Debes iniciar sesion para continuar.';
    }
    if (safeCode === 403) {
        return 'Esta zona es exclusiva para personal autorizado de Geo Rural.';
    }
    if (safeCode === 404) {
        return 'La ruta solicitada no existe o fue movida.';
    }
    if (safeCode === 503) {
        return 'El sistema no esta disponible por el momento. Intenta nuevamente en unos minutos.';
    }
    return 'No fue posible completar la solicitud.';
}

function buildErrorPageHtml({ statusCode = 500, message = '', requestPath = '' } = {}) {
    const safeCode = Number.isInteger(statusCode) && statusCode >= 100 ? statusCode : 500;
    const title = resolveErrorPageTitle(safeCode);
    const detail = resolveErrorPageDetail(safeCode, message);
    const pathText = normalizeText(requestPath);
    const pathHtml = pathText
        ? `<p class="error-path">Ruta: <code>${escapeHtmlText(pathText)}</code></p>`
        : '';

    return `<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${escapeHtmlText(String(safeCode))} | Geo Rural</title>
    <style>
        :root {
            color-scheme: light;
        }
        * {
            box-sizing: border-box;
        }
        body {
            margin: 0;
            min-height: 100vh;
            font-family: "Segoe UI", Tahoma, Arial, sans-serif;
            background: radial-gradient(circle at 15% 10%, #f7fbff 0%, #d7e3f7 45%, #b8cbe8 100%);
            color: #17345f;
            display: grid;
            place-items: center;
            padding: 18px;
        }
        .error-card {
            width: min(980px, 100%);
            border: 1px solid #91a8cc;
            border-radius: 16px;
            background: rgba(255, 255, 255, 0.78);
            box-shadow: 0 18px 44px rgba(20, 44, 86, 0.18);
            padding: 24px;
            text-align: center;
            backdrop-filter: blur(2px);
        }
        .error-code {
            margin: 0;
            font-size: clamp(72px, 17vw, 158px);
            line-height: 0.9;
            font-weight: 800;
            letter-spacing: 1px;
            color: #0f2e59;
            text-shadow: 0 2px 0 #ffffff;
        }
        .error-title {
            margin: 12px 0 6px;
            font-size: clamp(20px, 2.8vw, 32px);
            font-weight: 800;
            color: #1a3d72;
        }
        .error-detail {
            margin: 0;
            font-size: clamp(15px, 1.9vw, 19px);
            color: #284d7c;
            font-weight: 600;
        }
        .error-path {
            margin: 12px 0 0;
            color: #335880;
            font-size: 13px;
        }
        .error-path code {
            font-size: 12px;
            background: rgba(210, 223, 244, 0.7);
            border: 1px solid #9cb3d7;
            border-radius: 6px;
            padding: 2px 6px;
            color: #1c3e6f;
        }
        .scene {
            position: relative;
            margin: 18px auto 10px;
            width: min(760px, 100%);
            min-height: 250px;
            border-radius: 12px;
            border: 1px solid #93aacd;
            background: linear-gradient(180deg, rgba(239, 246, 255, 0.92), rgba(206, 221, 244, 0.96));
            overflow: hidden;
        }
        .scene::before {
            content: '';
            position: absolute;
            inset: auto 0 0;
            height: 74px;
            background: linear-gradient(180deg, rgba(132, 159, 199, 0.28), rgba(103, 129, 169, 0.55));
        }
        .door-zone {
            position: absolute;
            right: 10%;
            bottom: 58px;
            width: 220px;
            height: 182px;
            border-radius: 10px 10px 0 0;
            background: linear-gradient(180deg, #35557f 0%, #203c63 100%);
            box-shadow: 0 20px 30px rgba(24, 44, 76, 0.25);
        }
        .door {
            position: absolute;
            left: 50%;
            transform-origin: left center;
            transform: translateX(-50%);
            bottom: 0;
            width: 132px;
            height: 150px;
            border: 2px solid #19365e;
            border-radius: 8px 8px 0 0;
            background: linear-gradient(180deg, #2a4f7f 0%, #1b365d 100%);
            animation: door-open 2.6s ease-in-out infinite;
        }
        .door::after {
            content: '';
            position: absolute;
            right: 11px;
            top: 76px;
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: #d6e5f9;
            box-shadow: 0 0 0 2px rgba(23, 52, 95, 0.45);
        }
        .door-glow {
            position: absolute;
            left: 50%;
            bottom: 0;
            transform: translateX(-50%);
            width: 94px;
            height: 134px;
            border-radius: 8px 8px 0 0;
            background: radial-gradient(circle at center, rgba(194, 217, 250, 0.82), rgba(150, 180, 223, 0.2));
            animation: glow-pulse 2.6s ease-in-out infinite;
            pointer-events: none;
        }
        .door-sign {
            position: absolute;
            top: -12px;
            left: 50%;
            transform: translateX(-50%);
            width: 198px;
            border: 1px solid #cda754;
            border-radius: 7px;
            padding: 7px 10px;
            font-size: 11px;
            font-weight: 800;
            letter-spacing: 0.7px;
            color: #583f08;
            background: linear-gradient(180deg, #ffe59d 0%, #f7cc63 100%);
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.2);
            text-transform: uppercase;
        }
        .person {
            position: absolute;
            left: clamp(30px, 20vw, 200px);
            bottom: 56px;
            width: 72px;
            height: 130px;
            animation: person-walk 2.6s ease-in-out infinite;
        }
        .person-head {
            position: absolute;
            top: 0;
            left: 25px;
            width: 22px;
            height: 22px;
            border-radius: 50%;
            background: #112d50;
        }
        .person-body {
            position: absolute;
            top: 20px;
            left: 30px;
            width: 12px;
            height: 54px;
            border-radius: 10px;
            background: #16385f;
        }
        .person-arm {
            position: absolute;
            top: 28px;
            width: 30px;
            height: 6px;
            border-radius: 6px;
            background: #16385f;
            transform-origin: left center;
        }
        .person-arm.left {
            left: 8px;
            transform: rotate(24deg);
            animation: arm-left 0.9s ease-in-out infinite;
        }
        .person-arm.right {
            left: 34px;
            transform: rotate(-22deg);
            animation: arm-right 0.9s ease-in-out infinite;
        }
        .person-leg {
            position: absolute;
            bottom: 0;
            width: 10px;
            height: 56px;
            border-radius: 9px;
            background: #102a48;
            transform-origin: top center;
        }
        .person-leg.left {
            left: 22px;
            animation: leg-left 0.9s ease-in-out infinite;
        }
        .person-leg.right {
            left: 40px;
            animation: leg-right 0.9s ease-in-out infinite;
        }
        .file {
            position: absolute;
            left: 42px;
            bottom: 86px;
            width: 24px;
            height: 16px;
            border-radius: 3px;
            border: 1px solid #be952d;
            background: linear-gradient(180deg, #f6dc8e 0%, #e7ba4d 100%);
            box-shadow: 0 4px 10px rgba(23, 52, 95, 0.2);
            animation: file-fly 2.6s linear infinite;
        }
        .file::before {
            content: '';
            position: absolute;
            top: -4px;
            left: 3px;
            width: 10px;
            height: 4px;
            border-radius: 2px 2px 0 0;
            background: #ffe9b6;
            border: 1px solid #d4b362;
            border-bottom: none;
        }
        .file-2 {
            animation-delay: 0.86s;
        }
        .file-3 {
            animation-delay: 1.72s;
        }
        .back-link {
            display: inline-block;
            margin-top: 14px;
            border: 1px solid #6b89b5;
            border-radius: 9px;
            padding: 8px 14px;
            background: #254976;
            color: #ffffff;
            text-decoration: none;
            font-weight: 700;
            font-size: 14px;
        }
        .back-link:hover {
            background: #1d3b61;
        }
        @keyframes door-open {
            0%, 100% {
                transform: translateX(-50%) perspective(320px) rotateY(0deg);
            }
            40%, 70% {
                transform: translateX(-50%) perspective(320px) rotateY(-42deg);
            }
        }
        @keyframes glow-pulse {
            0%, 100% {
                opacity: 0.55;
            }
            50% {
                opacity: 0.95;
            }
        }
        @keyframes person-walk {
            0%, 100% {
                transform: translateX(0) scale(1);
                opacity: 1;
            }
            60% {
                transform: translateX(230px) scale(0.95);
                opacity: 0.95;
            }
            80% {
                transform: translateX(280px) scale(0.88);
                opacity: 0.72;
            }
        }
        @keyframes arm-left {
            0%, 100% {
                transform: rotate(24deg);
            }
            50% {
                transform: rotate(-12deg);
            }
        }
        @keyframes arm-right {
            0%, 100% {
                transform: rotate(-22deg);
            }
            50% {
                transform: rotate(12deg);
            }
        }
        @keyframes leg-left {
            0%, 100% {
                transform: rotate(12deg);
            }
            50% {
                transform: rotate(-10deg);
            }
        }
        @keyframes leg-right {
            0%, 100% {
                transform: rotate(-10deg);
            }
            50% {
                transform: rotate(12deg);
            }
        }
        @keyframes file-fly {
            0% {
                transform: translate(0, 0) rotate(-4deg);
                opacity: 0;
            }
            12% {
                opacity: 1;
            }
            80% {
                opacity: 1;
            }
            100% {
                transform: translate(360px, -16px) rotate(4deg);
                opacity: 0;
            }
        }
        @media (max-width: 720px) {
            .error-card {
                padding: 18px 14px;
            }
            .scene {
                min-height: 230px;
            }
            .door-zone {
                right: 5%;
                width: 180px;
                height: 166px;
            }
            .door-sign {
                width: 162px;
                font-size: 10px;
                padding: 6px 8px;
            }
            .door {
                width: 114px;
                height: 138px;
            }
            .person {
                left: 20px;
                transform: scale(0.9);
                transform-origin: left bottom;
            }
            @keyframes file-fly {
                0% {
                    transform: translate(0, 0) rotate(-4deg);
                    opacity: 0;
                }
                12% {
                    opacity: 1;
                }
                80% {
                    opacity: 1;
                }
                100% {
                    transform: translate(270px, -12px) rotate(4deg);
                    opacity: 0;
                }
            }
        }
    </style>
</head>
<body>
    <main class="error-card">
        <p class="error-code">${escapeHtmlText(String(safeCode))}</p>
        <h1 class="error-title">${escapeHtmlText(title)}</h1>
        <p class="error-detail">${escapeHtmlText(detail)}</p>
        ${pathHtml}
        <div class="scene" role="img" aria-label="Persona entrando a una puerta con letrero de solo personal autorizado.">
            <div class="door-zone">
                <div class="door-sign">SOLO PERSONAL AUTORIZADO</div>
                <div class="door-glow"></div>
                <div class="door"></div>
            </div>
            <div class="person">
                <span class="person-head"></span>
                <span class="person-body"></span>
                <span class="person-arm left"></span>
                <span class="person-arm right"></span>
                <span class="person-leg left"></span>
                <span class="person-leg right"></span>
            </div>
            <span class="file file-1"></span>
            <span class="file file-2"></span>
            <span class="file file-3"></span>
        </div>
        <a class="back-link" href="/index.html">Volver al inicio</a>
    </main>
</body>
</html>`;
}

function sendErrorPage(res, { statusCode = 500, message = '', requestPath = '' } = {}) {
    const safeCode = Number.isInteger(statusCode) && statusCode >= 100 ? statusCode : 500;
    const html = buildErrorPageHtml({
        statusCode: safeCode,
        message,
        requestPath
    });
    return res.status(safeCode).type('html; charset=utf-8').send(html);
}

function normalizeRut(value) {
    return String(value || '').replace(/[^0-9kK]/g, '').toLowerCase();
}

function normalizeRolForSearch(value) {
    const original = normalizeText(value);
    const normalized = String(original || '')
        .normalize('NFKC')
        .replace(/\u00A0/g, ' ')
        .toLowerCase();
    const canonical = normalized.replace(/\s+/g, '');
    const compact = canonical.replace(/[^0-9a-z]/g, '');

    return {
        original,
        canonical,
        compact
    };
}

function buildRolCanonicalSql(columnName = 'rol') {
    return `LOWER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(${columnName}, ' ', ''), CHAR(9), ''), CHAR(10), ''), CHAR(13), ''), CHAR(160), ''))`;
}

function buildRolCompactSql(columnName = 'rol') {
    const canonicalSql = buildRolCanonicalSql(columnName);
    return `REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(${canonicalSql}, '.', ''), '-', ''), '_', ''), '/', ''), ',', ''), ';', ''), ':', '')`;
}

function buildRolWhereClause(columnName = 'rol') {
    const canonicalSql = buildRolCanonicalSql(columnName);
    const compactSql = buildRolCompactSql(columnName);
    return `(${canonicalSql} = ? OR ${compactSql} = ?)`;
}

function calculateRutVerifierDigit(rutBodyDigits) {
    let sum = 0;
    let multiplier = 2;

    for (let index = rutBodyDigits.length - 1; index >= 0; index -= 1) {
        sum += Number(rutBodyDigits[index]) * multiplier;
        multiplier = multiplier === 7 ? 2 : multiplier + 1;
    }

    const remainder = 11 - (sum % 11);
    if (remainder === 11) {
        return '0';
    }
    if (remainder === 10) {
        return 'k';
    }

    return String(remainder);
}

function formatRut(rutBodyDigits, verifierDigit) {
    const normalizedBody = String(rutBodyDigits || '').replace(/^0+(?=\d)/, '');
    const safeBody = normalizedBody || '0';
    const bodyWithDots = safeBody.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return `${bodyWithDots}-${String(verifierDigit || '').toUpperCase()}`;
}

function normalizeAndFormatRut(value) {
    const normalized = normalizeRut(value);
    if (normalized.length < 2) {
        return '';
    }

    const rutBodyDigits = normalized.slice(0, -1);
    const verifierDigit = normalized.slice(-1);
    if (!/^\d+$/.test(rutBodyDigits)) {
        return '';
    }

    const expectedVerifier = calculateRutVerifierDigit(rutBodyDigits);
    if (verifierDigit !== expectedVerifier) {
        return '';
    }

    return formatRut(rutBodyDigits, verifierDigit);
}

function isValidEmail(value) {
    return SIMPLE_EMAIL_REGEX.test(String(value || '').trim());
}

function normalizeEstado(value) {
    const estado = String(value || '').trim().toLowerCase();

    if (!estado) {
        return '';
    }

    if (estado === 'facturada' || estado === 'activo' || estado === 'realizada' || estado === 'si') {
        return 'facturada';
    }

    if (estado === 'no_facturada' || estado === 'inactivo' || estado === 'no_realizada' || estado === 'no') {
        return 'no_facturada';
    }

    return estado;
}

function normalizeFacturaSolicitudEstado(value) {
    const estado = String(value || '').trim().toUpperCase();
    if (estado === FACTURA_SOLICITUD_ESTADO_PENDIENTE) {
        return FACTURA_SOLICITUD_ESTADO_PENDIENTE;
    }
    if (estado === FACTURA_SOLICITUD_ESTADO_ENVIADA) {
        return FACTURA_SOLICITUD_ESTADO_ENVIADA;
    }
    if (estado === FACTURA_SOLICITUD_ESTADO_ANULADA) {
        return FACTURA_SOLICITUD_ESTADO_ANULADA;
    }

    return '';
}

function normalizeUserRole(value, fallbackRole = ROLE_OPERADOR) {
    const role = String(value || '').trim().toUpperCase();
    if (role === ROLE_SUPER) {
        return ROLE_SUPER;
    }
    if (role === ROLE_ADMIN) {
        return ROLE_ADMIN;
    }
    if (role === ROLE_SECRETARIA) {
        return ROLE_SECRETARIA;
    }
    if (role === ROLE_OPERADOR) {
        return ROLE_OPERADOR;
    }
    if (role === ROLE_SUPERVISOR) {
        return ROLE_SUPERVISOR;
    }

    const fallback = String(fallbackRole || '').trim().toUpperCase();
    if (
        fallback === ROLE_SUPER ||
        fallback === ROLE_ADMIN ||
        fallback === ROLE_SECRETARIA ||
        fallback === ROLE_OPERADOR ||
        fallback === ROLE_SUPERVISOR
    ) {
        return fallback;
    }

    return ROLE_OPERADOR;
}

function parseCsv(value, fallback = []) {
    const entries = String(value || '')
        .split(',')
        .map((entry) => entry.trim())
        .filter((entry) => entry.length > 0);

    return entries.length > 0 ? entries : [...fallback];
}

function parseDefaultUsers(value) {
    const entries = String(value || '')
        .split(',')
        .map((entry) => entry.trim())
        .filter((entry) => entry.length > 0);

    const users = [];
    for (const entry of entries) {
        const [usernameRaw, passwordRaw, branchRaw, nameRaw, roleRaw] = entry
            .split('|')
            .map((part) => normalizeText(part));
        if (!usernameRaw || !passwordRaw || !branchRaw) {
            continue;
        }

        const normalizedRole = normalizeUserRole(roleRaw, ROLE_OPERADOR);

        users.push({
            username: usernameRaw.toLowerCase(),
            password: passwordRaw,
            sucursal: branchRaw,
            nombre: nameRaw || usernameRaw,
            role: normalizedRole
        });
    }

    return users;
}

function generateSecurePassword(length = 18) {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%*+-_';
    let password = '';

    while (password.length < length) {
        const randomByte = crypto.randomBytes(1)[0];
        password += chars[randomByte % chars.length];
    }

    if (!/[A-Z]/.test(password) || !/[a-z]/.test(password) || !/\d/.test(password) || !/[!@#$%*+\-_]/.test(password)) {
        return generateSecurePassword(length);
    }

    return password;
}

function resolveInitialUsers() {
    if (DEFAULT_USERS.length > 0) {
        return DEFAULT_USERS;
    }

    if (!runtimeBootstrapAdmin) {
        runtimeBootstrapAdmin = {
            username: 'admin',
            password: generateSecurePassword(20),
            sucursal: 'Casa Matriz',
            nombre: 'Administrador Temporal',
            role: ROLE_ADMIN
        };

        console.warn('DEFAULT_USERS vacio: se genero un ADMIN temporal aleatorio para el primer inicio.');
        console.warn(`Usuario temporal: ${runtimeBootstrapAdmin.username}`);
        console.warn(`Clave temporal generada: ${runtimeBootstrapAdmin.password}`);
        console.warn('Ingresa y cambia esta clave inmediatamente.');
    }

    return [runtimeBootstrapAdmin];
}

function parseNumIngreso(value) {
    const match = /^(\d+)-(\d{4})$/.exec(String(value || '').trim());
    if (!match) {
        return null;
    }

    return {
        correlativo: Number(match[1]),
        anio: Number(match[2])
    };
}

function isValidIsoDateOnly(value) {
    const text = normalizeText(value);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
        return false;
    }

    const parsed = new Date(`${text}T00:00:00Z`);
    if (Number.isNaN(parsed.getTime())) {
        return false;
    }

    return parsed.toISOString().slice(0, 10) === text;
}

function parseHistoryDateTimeInput(value) {
    const normalized = normalizeText(value).replace(' ', 'T');
    const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/.exec(normalized);
    if (!match) {
        return null;
    }

    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const hour = Number(match[4]);
    const minute = Number(match[5]);
    const second = Number(match[6] || '0');

    if (
        !Number.isInteger(year) ||
        !Number.isInteger(month) ||
        !Number.isInteger(day) ||
        !Number.isInteger(hour) ||
        !Number.isInteger(minute) ||
        !Number.isInteger(second)
    ) {
        return null;
    }

    if (month < 1 || month > 12 || day < 1 || day > 31 || hour < 0 || hour > 23 || minute < 0 || minute > 59 || second < 0 || second > 59) {
        return null;
    }

    const candidate = new Date(Date.UTC(year, month - 1, day, hour, minute, second));
    if (Number.isNaN(candidate.getTime())) {
        return null;
    }

    if (
        candidate.getUTCFullYear() !== year ||
        candidate.getUTCMonth() + 1 !== month ||
        candidate.getUTCDate() !== day ||
        candidate.getUTCHours() !== hour ||
        candidate.getUTCMinutes() !== minute ||
        candidate.getUTCSeconds() !== second
    ) {
        return null;
    }

    const safeSecond = String(second).padStart(2, '0');
    return `${match[1]}-${match[2]}-${match[3]} ${match[4]}:${match[5]}:${safeSecond}`;
}

function toApiDateTimeLocal(value) {
    if (!value) {
        return '';
    }

    if (value instanceof Date) {
        if (Number.isNaN(value.getTime())) {
            return '';
        }

        // Always include timezone in API payload to avoid client-side offsets.
        return value.toISOString();
    }

    const raw = String(value || '').trim();
    if (!raw) {
        return '';
    }

    const parsed = new Date(raw);
    if (Number.isNaN(parsed.getTime())) {
        return raw;
    }
    return parsed.toISOString();
}

function parseNumLotes(value) {
    if (value === '' || value === null || typeof value === 'undefined') {
        return { isValid: true, value: null };
    }

    const parsed = Number(value);
    const maxInt32 = 2147483647;

    if (!Number.isSafeInteger(parsed) || parsed < 1 || parsed > maxInt32) {
        return { isValid: false, value: null };
    }

    return { isValid: true, value: parsed };
}

function parseFacturaMonto(value) {
    if (value === '' || value === null || typeof value === 'undefined') {
        return { isValid: true, value: null };
    }

    const raw = String(value).trim();
    if (!raw) {
        return { isValid: true, value: null };
    }

    const compact = raw.replace(/\$/g, '').replace(/\s+/g, '').replace(/[^0-9,.-]/g, '');
    if (!compact) {
        return { isValid: false, value: null };
    }

    const hasDot = compact.includes('.');
    const hasComma = compact.includes(',');
    let normalized = compact;

    if (hasDot && hasComma) {
        const lastDot = compact.lastIndexOf('.');
        const lastComma = compact.lastIndexOf(',');
        if (lastDot > lastComma) {
            normalized = compact.replace(/,/g, '');
        } else {
            normalized = compact.replace(/\./g, '').replace(',', '.');
        }
    } else if (hasComma) {
        if (/^\d{1,3}(,\d{3})+$/.test(compact)) {
            normalized = compact.replace(/,/g, '');
        } else {
            normalized = compact.replace(',', '.');
        }
    } else if (hasDot) {
        if (/^\d{1,3}(\.\d{3})+$/.test(compact)) {
            normalized = compact.replace(/\./g, '');
        }
    }

    const parsed = Number(normalized);
    if (!Number.isFinite(parsed) || parsed <= 0) {
        return { isValid: false, value: null };
    }

    return { isValid: true, value: Number(parsed.toFixed(2)) };
}

function parseUtmValue(value) {
    const parsedMonto = parseFacturaMonto(value);
    if (!parsedMonto.isValid || parsedMonto.value === null) {
        return { isValid: false, value: null };
    }

    return { isValid: true, value: Number(parsedMonto.value.toFixed(2)) };
}

function getCurrentCalendarPeriod() {
    const now = new Date();
    return {
        anio: now.getFullYear(),
        mes: now.getMonth() + 1,
        dia: now.getDate()
    };
}

function httpGetJson(url, timeoutMs = UTM_SOURCE_TIMEOUT_MS) {
    return new Promise((resolve, reject) => {
        let settled = false;
        const finish = (handler, payload) => {
            if (settled) {
                return;
            }
            settled = true;
            handler(payload);
        };

        const request = https.get(
            url,
            {
                headers: {
                    Accept: 'application/json',
                    'User-Agent': 'geo-rural-utm/1.0'
                }
            },
            (response) => {
                const statusCode = Number(response.statusCode || 0);
                if (statusCode < 200 || statusCode >= 300) {
                    response.resume();
                    finish(reject, new Error(`UTM_SOURCE_HTTP_${statusCode || 'UNKNOWN'}`));
                    return;
                }

                let rawBody = '';
                response.setEncoding('utf8');
                response.on('data', (chunk) => {
                    rawBody += String(chunk || '');
                    if (rawBody.length > 1024 * 1024) {
                        request.destroy(new Error('UTM_SOURCE_RESPONSE_TOO_LARGE'));
                    }
                });

                response.on('end', () => {
                    try {
                        finish(resolve, JSON.parse(rawBody));
                    } catch (error) {
                        finish(reject, new Error('UTM_SOURCE_INVALID_JSON'));
                    }
                });
            }
        );

        request.setTimeout(Math.max(1000, Number(timeoutMs) || UTM_SOURCE_TIMEOUT_MS), () => {
            request.destroy(new Error('UTM_SOURCE_TIMEOUT'));
        });
        request.on('error', (error) => {
            finish(reject, error);
        });
    });
}

function findUtmValueByPeriodFromSource(payload, anio, mes) {
    if (!payload || !Array.isArray(payload.serie)) {
        return null;
    }

    for (const item of payload.serie) {
        if (!item || typeof item !== 'object') {
            continue;
        }

        const fecha = new Date(item.fecha);
        if (Number.isNaN(fecha.getTime())) {
            continue;
        }

        const entryAnio = fecha.getUTCFullYear();
        const entryMes = fecha.getUTCMonth() + 1;
        if (entryAnio !== anio || entryMes !== mes) {
            continue;
        }

        const parsedValue = Number(item.valor);
        if (!Number.isFinite(parsedValue) || parsedValue <= 0) {
            continue;
        }

        return Number(parsedValue.toFixed(2));
    }

    return null;
}

async function fetchUtmSuggestionForPeriod(anio, mes) {
    const safeAnio = Number(anio);
    const safeMes = Number(mes);
    if (!Number.isInteger(safeAnio) || !Number.isInteger(safeMes) || safeMes < 1 || safeMes > 12) {
        return null;
    }

    const cacheKey = `${safeAnio}-${String(safeMes).padStart(2, '0')}`;
    if (utmSourceCache.key === cacheKey && Date.now() < Number(utmSourceCache.expiresAt || 0)) {
        return utmSourceCache.value || null;
    }

    try {
        const payload = await httpGetJson(UTM_SOURCE_URL, UTM_SOURCE_TIMEOUT_MS);
        const value = findUtmValueByPeriodFromSource(payload, safeAnio, safeMes);
        const suggestion =
            Number.isFinite(value) && value > 0
                ? {
                      valor: value.toFixed(2),
                      fuente: 'mindicador.cl',
                      url: UTM_SOURCE_URL,
                      fetchedAt: new Date().toISOString()
                  }
                : null;

        utmSourceCache = {
            key: cacheKey,
            value: suggestion,
            expiresAt: Date.now() + (suggestion ? UTM_SOURCE_CACHE_MINUTES : UTM_SOURCE_ERROR_CACHE_MINUTES) * 60 * 1000
        };
        return suggestion;
    } catch (error) {
        console.warn(`No fue posible obtener UTM sugerida desde ${UTM_SOURCE_URL}: ${error.message}`);
        utmSourceCache = {
            key: cacheKey,
            value: null,
            expiresAt: Date.now() + UTM_SOURCE_ERROR_CACHE_MINUTES * 60 * 1000
        };
        return null;
    }
}

function parseFacturaSolicitudPayload(body = {}) {
    const rawPayload = body && typeof body === 'object' ? body : {};
    const destinationEmail = normalizeText(rawPayload.destinationEmail).toLowerCase();
    const payload = {
        numIngreso: normalizeText(rawPayload.numIngreso),
        nombreRazonSocial: normalizeText(rawPayload.nombreRazonSocial || rawPayload.nombreCliente),
        rut: normalizeText(rawPayload.rut || rawPayload.rutCliente),
        giro: normalizeText(rawPayload.giro),
        direccion: normalizeText(rawPayload.direccion),
        comuna: normalizeText(rawPayload.comuna),
        ciudad: normalizeText(rawPayload.ciudad),
        contacto: normalizeText(rawPayload.contacto || rawPayload.correoFactura),
        observacion: normalizeText(rawPayload.observacion || rawPayload.referencia),
        montoFacturarRaw: rawPayload.montoFacturar || rawPayload.monto || '',
        destinationEmail,
        estado: normalizeFacturaSolicitudEstado(rawPayload.estado)
    };

    if (!payload.numIngreso) {
        return { isValid: false, message: 'Debes indicar el NRO INGRESO de la factura.', payload: null };
    }

    if (!parseNumIngreso(payload.numIngreso)) {
        return { isValid: false, message: 'NRO INGRESO invalido. Usa formato 123-2026.', payload: null };
    }

    if (!payload.nombreRazonSocial || !payload.rut) {
        return { isValid: false, message: 'Completa nombre/razon social y RUT para la factura.', payload: null };
    }

    if (!payload.giro) {
        return { isValid: false, message: 'Debes completar el GIRO para la factura.', payload: null };
    }

    if (!payload.direccion) {
        return { isValid: false, message: 'Debes completar la DIRECCION para la factura.', payload: null };
    }

    if (!payload.observacion) {
        return { isValid: false, message: 'Debes completar la OBSERVACION para la factura.', payload: null };
    }

    if (!payload.estado) {
        return { isValid: false, message: 'Estado de solicitud invalido. Usa PENDIENTE, ENVIADA o ANULADA.', payload: null };
    }

    const formattedRut = normalizeAndFormatRut(payload.rut);
    if (!formattedRut) {
        return { isValid: false, message: 'RUT invalido. Usa formato 12.345.678-5.', payload: null };
    }
    payload.rut = formattedRut;

    const parsedMonto = parseFacturaMonto(payload.montoFacturarRaw);
    if (!parsedMonto.isValid || parsedMonto.value === null) {
        return { isValid: false, message: 'MONTO A FACTURAR invalido.', payload: null };
    }
    payload.montoFacturar = parsedMonto.value;

    if (payload.destinationEmail && !isValidEmail(payload.destinationEmail)) {
        return { isValid: false, message: 'El correo del contador no es valido.', payload: null };
    }

    if (payload.estado === FACTURA_SOLICITUD_ESTADO_ENVIADA && !payload.destinationEmail) {
        return {
            isValid: false,
            message: 'Debes indicar el correo del contador para registrar una solicitud ENVIADA.',
            payload: null
        };
    }

    const lengthChecks = [
        ['NRO INGRESO', payload.numIngreso, 20],
        ['NOMBRE / RAZON SOCIAL', payload.nombreRazonSocial, 255],
        ['RUT', payload.rut, 20],
        ['GIRO', payload.giro, 255],
        ['DIRECCION', payload.direccion, 255],
        ['COMUNA', payload.comuna, 120],
        ['CIUDAD', payload.ciudad, 120],
        ['CONTACTO', payload.contacto, 255],
        ['CORREO DESTINO', payload.destinationEmail, 255],
        ['ESTADO', payload.estado, 20]
    ];

    for (const [label, value, maxLength] of lengthChecks) {
        if (!value) {
            continue;
        }

        if (exceedsMaxLength(value, maxLength)) {
            return { isValid: false, message: `${label} excede el maximo permitido de ${maxLength} caracteres.`, payload: null };
        }
    }

    if (payload.observacion && exceedsMaxLength(payload.observacion, 1000)) {
        return { isValid: false, message: 'OBSERVACION excede el maximo permitido de 1000 caracteres.', payload: null };
    }

    return { isValid: true, message: '', payload };
}

function exceedsMaxLength(value, maxLength) {
    return String(value || '').length > Number(maxLength || 0);
}

function validateRegistroPayloadLengths(payload) {
    const checks = [
        ['NRO INGRESO', payload.numIngreso, 20],
        ['NRO INGRESO nuevo', payload.numIngresoNuevo, 20],
        ['NOMBRE O RAZON SOCIAL', payload.nombre, 255],
        ['RUT', payload.rut, 20],
        ['TELEFONO', payload.telefono, 40],
        ['CORREO', payload.correo, 255],
        ['REGION', payload.region, 120],
        ['COMUNA', payload.comuna, 120],
        ['NOMBRE PREDIO', payload.nombrePredio, 255],
        ['ROL', payload.rol, 120],
        ['ESTADO', payload.estado, 40],
        ['FACTURA NOMBRE / RAZON SOCIAL', payload.facturaNombreRazon, 255],
        ['FACTURA NUMERO', payload.facturaNumero, 80],
        ['FACTURA RUT', payload.facturaRut, 20],
        ['FACTURA GIRO', payload.facturaGiro, 255],
        ['FACTURA DIRECCION', payload.facturaDireccion, 255],
        ['FACTURA COMUNA', payload.facturaComuna, 120],
        ['FACTURA CIUDAD', payload.facturaCiudad, 120],
        ['FACTURA CONTACTO', payload.facturaContacto, 255]
    ];

    for (const [label, value, maxLength] of checks) {
        if (typeof value === 'undefined' || value === null || value === '') {
            continue;
        }

        if (exceedsMaxLength(value, maxLength)) {
            return `${label} excede el maximo permitido de ${maxLength} caracteres.`;
        }
    }

    if (payload.comentario && exceedsMaxLength(payload.comentario, 4000)) {
        return 'COMENTARIO excede el maximo permitido de 4000 caracteres.';
    }

    if (payload.facturaObservacion && exceedsMaxLength(payload.facturaObservacion, 1000)) {
        return 'FACTURA OBSERVACION excede el maximo permitido de 1000 caracteres.';
    }

    if (Array.isArray(payload.documentos) && payload.documentos.length > 100) {
        return 'DOCUMENTACION RECIBIDA excede el maximo permitido de 100 elementos.';
    }

    if (Array.isArray(payload.documentos)) {
        for (const doc of payload.documentos) {
            if (exceedsMaxLength(doc, 80)) {
                return 'Uno de los identificadores de DOCUMENTACION excede el maximo permitido de 80 caracteres.';
            }
        }
    }

    return '';
}

function safeJsonArray(value) {
    if (!Array.isArray(value)) {
        return [];
    }

    const cleanValues = value
        .map((item) => String(item || '').trim())
        .filter((item) => item.length > 0);

    return [...new Set(cleanValues)];
}

function parseStoredDocumentList(value) {
    if (!value) {
        return [];
    }

    if (Array.isArray(value)) {
        return safeJsonArray(value);
    }

    try {
        return safeJsonArray(JSON.parse(value));
    } catch (error) {
        return [];
    }
}

function hashToken(value) {
    return crypto.createHash('sha256').update(String(value || '')).digest('hex');
}

function createPasswordHash(password, saltHex = '') {
    const saltBuffer = saltHex ? Buffer.from(saltHex, 'hex') : crypto.randomBytes(16);
    const hashBuffer = crypto.pbkdf2Sync(String(password || ''), saltBuffer, 120000, 64, 'sha512');

    return {
        salt: saltBuffer.toString('hex'),
        hash: hashBuffer.toString('hex')
    };
}

function verifyPassword(password, saltHex, expectedHashHex) {
    if (!saltHex || !expectedHashHex) {
        return false;
    }

    const calculated = createPasswordHash(password, saltHex).hash;
    const left = Buffer.from(calculated, 'hex');
    const right = Buffer.from(expectedHashHex, 'hex');

    if (left.length !== right.length) {
        return false;
    }

    return crypto.timingSafeEqual(left, right);
}

function generateSessionToken() {
    return crypto.randomBytes(32).toString('hex');
}

function parseCookies(req) {
    const cookieHeader = req.get('cookie');
    if (!cookieHeader) {
        return {};
    }

    const pairs = cookieHeader.split(';');
    const parsed = {};
    for (const pair of pairs) {
        const index = pair.indexOf('=');
        if (index <= 0) {
            continue;
        }

        const key = normalizeText(pair.slice(0, index));
        const rawValue = pair.slice(index + 1);
        if (!key) {
            continue;
        }

        try {
            parsed[key] = decodeURIComponent(rawValue);
        } catch (error) {
            parsed[key] = rawValue;
        }
    }

    return parsed;
}

function buildSessionCookie(value, maxAgeSeconds) {
    const sameSiteRequiresSecure = AUTH_COOKIE_SAME_SITE === 'None';
    const shouldUseSecureCookie =
        sameSiteRequiresSecure || AUTH_COOKIE_SECURE_MODE === 'always' || (AUTH_COOKIE_SECURE_MODE === 'auto' && !IS_LOCALHOST_BIND);
    const attributes = [
        `${SESSION_COOKIE_NAME}=${encodeURIComponent(String(value || ''))}`,
        `Max-Age=${Math.max(0, Number(maxAgeSeconds) || 0)}`,
        'Path=/',
        'HttpOnly',
        `SameSite=${AUTH_COOKIE_SAME_SITE}`
    ];

    if (AUTH_COOKIE_DOMAIN) {
        attributes.push(`Domain=${AUTH_COOKIE_DOMAIN}`);
    }

    if (shouldUseSecureCookie) {
        attributes.push('Secure');
    }

    return attributes.join('; ');
}

function setSessionCookie(res, token) {
    res.append('Set-Cookie', buildSessionCookie(token, SESSION_COOKIE_MAX_AGE_SECONDS));
}

function clearSessionCookie(res) {
    const cookie = [
        buildSessionCookie('', 0),
        'Expires=Thu, 01 Jan 1970 00:00:00 GMT'
    ].join('; ');
    res.append('Set-Cookie', cookie);
}

function extractBearerToken(req) {
    const header = normalizeText(req.get('authorization'));
    const match = /^Bearer\s+(.+)$/i.exec(header);
    if (match) {
        return normalizeText(match[1]);
    }

    const cookies = parseCookies(req);
    return normalizeText(cookies[SESSION_COOKIE_NAME]);
}

function formatAuthUser(row) {
    const username = normalizeText(row.username).toLowerCase();

    return {
        id: row.user_id,
        username: row.username,
        nombre: row.nombre || row.username,
        sucursal: row.sucursal || '',
        role: normalizeUserRole(row.role, ROLE_OPERADOR),
        mustChangePassword: Number(row.must_change_password || 0) === 1,
        isGuest: username === GUEST_USERNAME
    };
}

function isGuestUser(user) {
    const username = normalizeText(user && user.username).toLowerCase();
    return Boolean(username) && username === GUEST_USERNAME;
}

function isGuestUsername(value) {
    return normalizeText(value).toLowerCase() === GUEST_USERNAME;
}

function isAdminUser(user) {
    return normalizeUserRole(user?.role, ROLE_OPERADOR) === ROLE_ADMIN;
}

function isSuperUser(user) {
    return normalizeUserRole(user?.role, ROLE_OPERADOR) === ROLE_SUPER;
}

function isPrivilegedRole(role) {
    return PRIVILEGED_ROLES.includes(normalizeUserRole(role, ROLE_OPERADOR));
}

function shouldForcePasswordChangeForRole(role) {
    return normalizeUserRole(role, ROLE_OPERADOR) !== ROLE_SUPER;
}

function isAdminOrSuperUser(user) {
    return isAdminUser(user) || isSuperUser(user);
}

function isSecretaryUser(user) {
    return normalizeUserRole(user?.role, ROLE_OPERADOR) === ROLE_SECRETARIA;
}

function isOperatorUser(user) {
    return normalizeUserRole(user?.role, ROLE_OPERADOR) === ROLE_OPERADOR;
}

function isSupervisorUser(user) {
    return normalizeUserRole(user?.role, ROLE_OPERADOR) === ROLE_SUPERVISOR;
}

function canUseSecretaryFeatures(user) {
    return isSecretaryUser(user) || isSuperUser(user);
}

function canUseInvoiceWorkflow(user) {
    return isOperatorUser(user) || canUseSecretaryFeatures(user);
}

function canUseOperatorDocumentAlerts(user) {
    return isOperatorUser(user) || isSuperUser(user);
}

function canReceiveWriteApiKey(user) {
    return Boolean(API_WRITE_KEY) && !isGuestUser(user);
}

function getAssignableRolesForUser(user) {
    return isSuperUser(user) ? ASSIGNABLE_ROLES_BY_SUPER : ASSIGNABLE_ROLES_BY_ADMIN;
}

function formatAssignableRolesLabel(user) {
    return getAssignableRolesForUser(user).join(', ');
}

function hasBackofficeRegistroAccess(user) {
    return isAdminOrSuperUser(user) || isSecretaryUser(user);
}

function getRegistroAccessScope(user, tableAlias = '') {
    // Registros globales: todos los roles autenticados pueden consultar sin filtro por sucursal.
    return {
        denied: false,
        clause: '',
        params: []
    };
}

function getFacturaSolicitudAccessScope(user, tableAlias = '') {
    if (hasBackofficeRegistroAccess(user)) {
        return {
            denied: false,
            clause: '',
            params: []
        };
    }

    const branch = normalizeText(user && user.sucursal);
    if (!branch) {
        return {
            denied: true,
            clause: '',
            params: []
        };
    }

    const prefix = tableAlias ? `${tableAlias}.` : '';
    return {
        denied: false,
        clause: `${prefix}created_by_sucursal = ?`,
        params: [branch]
    };
}

function getOperatorDocumentAlertsScope(user, tableAlias = '') {
    if (isSuperUser(user)) {
        return {
            denied: false,
            clause: '',
            params: []
        };
    }

    if (!isOperatorUser(user)) {
        return {
            denied: true,
            clause: '',
            params: []
        };
    }

    const branch = normalizeText(user && user.sucursal);
    if (!branch) {
        return {
            denied: true,
            clause: '',
            params: []
        };
    }

    const prefix = tableAlias ? `${tableAlias}.` : '';
    return {
        denied: false,
        clause: `(${prefix}sucursal = ? OR ${prefix}sucursal IS NULL OR TRIM(${prefix}sucursal) = '')`,
        params: [branch]
    };
}

function buildHistoryComment(action, providedComment, docsAdded = [], docsRemoved = []) {
    const cleanComment = normalizeText(providedComment);

    if (cleanComment) {
        return cleanComment;
    }

    if (action === 'CREACION') {
        return 'Registro creado.';
    }

    if (docsAdded.length > 0 || docsRemoved.length > 0) {
        return 'Documentacion actualizada.';
    }

    return 'Registro modificado.';
}

function formatHistoryTextValue(value) {
    const normalized = normalizeText(value);
    return normalized || 'vacio';
}

function formatHistoryEstadoValue(value) {
    const normalized = normalizeEstado(value);
    if (!normalized) {
        return 'sin estado';
    }

    if (normalized === 'facturada') {
        return 'facturado';
    }

    if (normalized === 'no_facturada') {
        return 'no facturado';
    }

    return normalized;
}

function formatHistoryNumLotesValue(value) {
    if (value === null || typeof value === 'undefined' || value === '') {
        return 'sin valor';
    }

    return String(value);
}

function formatHistoryFacturaMontoValue(value) {
    if (value === null || typeof value === 'undefined' || value === '') {
        return 'sin valor';
    }

    const parsed = Number(value);
    if (!Number.isFinite(parsed) || parsed <= 0) {
        return 'sin valor';
    }

    return `$ ${parsed.toLocaleString('es-CL', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
    })}`;
}

function formatHistoryChangeLabelSummary(labels, maxVisible = HISTORY_SUMMARY_MAX_LABELS) {
    const uniqueLabels = Array.from(
        new Set(
            (Array.isArray(labels) ? labels : [])
                .map((item) => normalizeText(item))
                .filter((item) => Boolean(item))
        )
    );

    if (uniqueLabels.length === 0) {
        return '';
    }

    const safeMax = Number.isInteger(maxVisible) && maxVisible > 0 ? maxVisible : 5;
    if (uniqueLabels.length <= safeMax) {
        return uniqueLabels.join(', ');
    }

    const visibleLabels = uniqueLabels.slice(0, safeMax);
    const remaining = uniqueLabels.length - safeMax;
    return `${visibleLabels.join(', ')} +${remaining} mas`;
}

function buildAutomaticModificationComment(existingRow, payload, docsAdded = [], docsRemoved = [], parsedNumLotesValue = null) {
    const changes = [];

    const appendTextChange = (label, beforeValue, afterValue, group = 'registro') => {
        const before = normalizeText(beforeValue);
        const after = normalizeText(afterValue);
        if (before === after) {
            return;
        }

        changes.push({
            group,
            label,
            detail: `modificacion de ${label}: ${formatHistoryTextValue(before)} -> ${formatHistoryTextValue(after)}`
        });
    };

    appendTextChange('nro ingreso', existingRow.num_ingreso, payload.numIngresoNuevo, 'registro');
    appendTextChange('nombre', existingRow.nombre, payload.nombre, 'registro');
    appendTextChange('rut', existingRow.rut, payload.rut, 'registro');
    appendTextChange('telefono', existingRow.telefono, payload.telefono, 'registro');
    appendTextChange('correo', existingRow.correo, payload.correo, 'registro');
    appendTextChange('region', existingRow.region, payload.region, 'registro');
    appendTextChange('comuna', existingRow.comuna, payload.comuna, 'registro');
    appendTextChange('nombre predio', existingRow.nombre_predio, payload.nombrePredio, 'registro');
    appendTextChange('rol', existingRow.rol, payload.rol, 'registro');
    appendTextChange('factura nombre/razon social', existingRow.factura_nombre_razon, payload.facturaNombreRazon, 'factura');
    appendTextChange('factura numero', existingRow.factura_numero, payload.facturaNumero, 'factura');
    appendTextChange('factura rut', existingRow.factura_rut, payload.facturaRut, 'factura');
    appendTextChange('factura giro', existingRow.factura_giro, payload.facturaGiro, 'factura');
    appendTextChange('factura direccion', existingRow.factura_direccion, payload.facturaDireccion, 'factura');
    appendTextChange('factura comuna', existingRow.factura_comuna, payload.facturaComuna, 'factura');
    appendTextChange('factura ciudad', existingRow.factura_ciudad, payload.facturaCiudad, 'factura');
    appendTextChange('factura contacto', existingRow.factura_contacto, payload.facturaContacto, 'factura');
    appendTextChange('factura observacion', existingRow.factura_observacion, payload.facturaObservacion, 'factura');

    const beforeLotes = existingRow.num_lotes === null || typeof existingRow.num_lotes === 'undefined' ? null : Number(existingRow.num_lotes);
    const afterLotes =
        parsedNumLotesValue === null || typeof parsedNumLotesValue === 'undefined'
            ? null
            : Number(parsedNumLotesValue);
    if (beforeLotes !== afterLotes) {
        changes.push({
            group: 'registro',
            label: 'nro de lotes',
            detail: `modificacion de nro de lotes: ${formatHistoryNumLotesValue(beforeLotes)} -> ${formatHistoryNumLotesValue(afterLotes)}`
        });
    }

    const beforeEstado = normalizeEstado(existingRow.estado);
    const afterEstado = normalizeEstado(payload.estado);
    if (beforeEstado !== afterEstado) {
        changes.push({
            group: 'registro',
            label: 'estado de factura',
            detail: `modificacion de estado de factura: ${formatHistoryEstadoValue(beforeEstado)} -> ${formatHistoryEstadoValue(afterEstado)}`
        });
    }

    const beforeFacturaMonto =
        existingRow.factura_monto === null || typeof existingRow.factura_monto === 'undefined'
            ? null
            : Number(existingRow.factura_monto);
    const afterFacturaMonto =
        payload.facturaMonto === null || typeof payload.facturaMonto === 'undefined'
            ? null
            : Number(payload.facturaMonto);
    if (beforeFacturaMonto !== afterFacturaMonto) {
        changes.push({
            group: 'factura',
            label: 'monto a facturar',
            detail: `modificacion de monto a facturar: ${formatHistoryFacturaMontoValue(beforeFacturaMonto)} -> ${formatHistoryFacturaMontoValue(afterFacturaMonto)}`
        });
    }

    if (docsAdded.length > 0) {
        changes.push({
            group: 'documentos',
            label: 'agregados',
            detail: `documentos agregados: ${docsAdded.join(', ')}`
        });
    }

    if (docsRemoved.length > 0) {
        changes.push({
            group: 'documentos',
            label: 'eliminados',
            detail: `documentos eliminados: ${docsRemoved.join(', ')}`
        });
    }

    if (changes.length === 0) {
        return 'modificacion sin cambios detectados.';
    }

    if (changes.length <= HISTORY_DETAIL_MAX_CHANGES) {
        return changes.map((item) => item.detail).join(' | ');
    }

    const registroLabels = changes
        .filter((item) => item.group === 'registro')
        .map((item) => item.label);
    const facturaLabels = changes
        .filter((item) => item.group === 'factura')
        .map((item) => normalizeText(item.label).replace(/^factura\s+/i, ''));
    const summaryParts = [];

    const registroSummary = formatHistoryChangeLabelSummary(registroLabels);
    if (registroSummary) {
        summaryParts.push(`registro: ${registroSummary}`);
    }

    const facturaSummary = formatHistoryChangeLabelSummary(facturaLabels);
    if (facturaSummary) {
        summaryParts.push(`factura: ${facturaSummary}`);
    }

    if (docsAdded.length > 0) {
        const addedPreview = docsAdded.slice(0, 3).join(', ');
        const addedSuffix = docsAdded.length > 3 ? ` +${docsAdded.length - 3} mas` : '';
        summaryParts.push(`documentos +${docsAdded.length}${addedPreview ? ` (${addedPreview}${addedSuffix})` : ''}`);
    }
    if (docsRemoved.length > 0) {
        const removedPreview = docsRemoved.slice(0, 3).join(', ');
        const removedSuffix = docsRemoved.length > 3 ? ` +${docsRemoved.length - 3} mas` : '';
        summaryParts.push(`documentos -${docsRemoved.length}${removedPreview ? ` (${removedPreview}${removedSuffix})` : ''}`);
    }

    const header = `modificacion resumida (${changes.length} cambios)`;
    return summaryParts.length > 0 ? `${header}: ${summaryParts.join(' | ')}` : header;
}

function toApiRow(row) {
    const parsedFacturaMonto = Number(row.factura_monto);
    const facturaMontoValue =
        row.factura_monto === null || typeof row.factura_monto === 'undefined' || !Number.isFinite(parsedFacturaMonto)
            ? ''
            : parsedFacturaMonto.toFixed(2);
    return {
        numIngreso: row.num_ingreso,
        nombre: row.nombre,
        rut: row.rut,
        telefono: row.telefono || '',
        correo: row.correo || '',
        region: row.region,
        comuna: row.comuna,
        nombrePredio: row.nombre_predio || '',
        rol: row.rol,
        numLotes: row.num_lotes ?? '',
        estado: normalizeEstado(row.estado),
        comentario: row.comentario || '',
        documentos: parseStoredDocumentList(row.documentos),
        factura: {
            nombreRazonSocial: row.factura_nombre_razon || '',
            numeroFactura: row.factura_numero || '',
            rut: row.factura_rut || '',
            giro: row.factura_giro || '',
            direccion: row.factura_direccion || '',
            comuna: row.factura_comuna || '',
            ciudad: row.factura_ciudad || '',
            contacto: row.factura_contacto || '',
            observacion: row.factura_observacion || '',
            montoFacturar: facturaMontoValue
        }
    };
}

function toFacturaSolicitudRow(row) {
    const parsedMonto = Number(row.monto_facturar);
    const montoValue =
        row.monto_facturar === null || typeof row.monto_facturar === 'undefined' || !Number.isFinite(parsedMonto)
            ? ''
            : parsedMonto.toFixed(2);

    const createdAt = row.created_at instanceof Date ? row.created_at.toISOString() : String(row.created_at);
    const sentAt =
        row.sent_at instanceof Date
            ? row.sent_at.toISOString()
            : row.sent_at
              ? String(row.sent_at)
              : '';

    return {
        id: Number(row.id),
        numIngreso: row.num_ingreso || '',
        nombreRazonSocial: row.nombre_razon_social || '',
        rut: row.rut || '',
        giro: row.giro || '',
        direccion: row.direccion || '',
        comuna: row.comuna || '',
        ciudad: row.ciudad || '',
        contacto: row.contacto || '',
        observacion: row.observacion || '',
        montoFacturar: montoValue,
        estado: normalizeFacturaSolicitudEstado(row.estado) || FACTURA_SOLICITUD_ESTADO_PENDIENTE,
        destinationEmail: row.destino_email || '',
        createdBy: row.created_by || '',
        sucursal: row.created_by_sucursal || '',
        sentAt,
        createdAt
    };
}

function toHistoryRow(row) {
    const createdAt = toApiDateTimeLocal(row.created_at);
    const usuario = row.sucursal ? `${row.usuario_nombre} (${row.sucursal})` : row.usuario_nombre;

    return {
        id: Number(row.id || 0),
        numIngreso: row.num_ingreso || '',
        accion: normalizeText(row.accion || ''),
        fecha: createdAt,
        comentario: row.comentario,
        usuario: usuario || 'Sin usuario'
    };
}

function toPublicProgressHistoryRow(row) {
    return {
        fecha: normalizeText(row?.fecha),
        comentario: String(row?.comentario || '')
    };
}

function toUtmMensualRow(row) {
    if (!row || typeof row !== 'object') {
        return null;
    }

    const parsedValue = Number(row.valor);
    return {
        anio: Number(row.anio),
        mes: Number(row.mes),
        valor: Number.isFinite(parsedValue) ? parsedValue.toFixed(2) : '',
        registradoPor: row.registrado_por_nombre || '',
        createdAt: row.created_at instanceof Date ? row.created_at.toISOString() : String(row.created_at || ''),
        updatedAt: row.updated_at instanceof Date ? row.updated_at.toISOString() : String(row.updated_at || '')
    };
}

function toCotizacionServicioResumen(row, utmValue) {
    if (!row || typeof row !== 'object') {
        return null;
    }

    const nombreServicio = normalizeText(row.nombre_servicio);
    const parsedValorUtm = Number(row.valor_utm);
    const parsedUtmValue = Number(utmValue);
    if (!nombreServicio || !Number.isFinite(parsedValorUtm) || parsedValorUtm <= 0 || !Number.isFinite(parsedUtmValue) || parsedUtmValue <= 0) {
        return null;
    }

    const valorUtm = Number(parsedValorUtm.toFixed(2));
    const valorPesos = Number((valorUtm * parsedUtmValue).toFixed(2));
    return {
        id: Number(row.id),
        nombreServicio,
        valorUtm: valorUtm.toFixed(2),
        valorPesos: valorPesos.toFixed(2)
    };
}

function formatCurrencyClp(value) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed) || parsed <= 0) {
        return '-';
    }

    return `$ ${parsed.toLocaleString('es-CL', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
    })}`;
}

function formatUtmUnits(value) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed) || parsed <= 0) {
        return '-';
    }

    return `${parsed.toLocaleString('es-CL', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
    })} UTM`;
}

function formatThousandsDots(value) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) {
        return '0';
    }

    return Math.round(parsed).toLocaleString('es-CL', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    });
}

function formatTemplateUtmUnits(value) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed) || parsed <= 0) {
        return '0,00';
    }

    return parsed.toLocaleString('es-CL', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function formatMonthlyPeriodLabel(anio, mes) {
    const safeYear = Number(anio);
    const safeMonth = Number(mes);
    const validMonth = Number.isInteger(safeMonth) && safeMonth >= 1 && safeMonth <= 12 ? safeMonth : 1;
    const validYear = Number.isInteger(safeYear) && safeYear >= 2000 ? safeYear : new Date().getFullYear();
    return `${(MONTH_NAMES_ES[validMonth - 1] || 'mes').toUpperCase()} ${validYear}`;
}

function normalizeCotizacionParcelamientoKey(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ');
}

function buildCotizacionParcelamientoRows(utmValue) {
    const parsedUtm = Number(utmValue);
    if (!Number.isFinite(parsedUtm) || parsedUtm <= 0) {
        return [];
    }

    return COTIZACION_TEMPLATE_ROWS.map((row) => {
        const lotes = normalizeText(row?.lotes);
        const valorUtm = Number(row?.valorUtm);
        if (!lotes || !Number.isFinite(valorUtm) || valorUtm <= 0) {
            return null;
        }

        const neto = Math.round(valorUtm * parsedUtm);
        const iva = Math.round(neto * 0.19);
        const total = neto + iva;
        return {
            key: normalizeCotizacionParcelamientoKey(lotes),
            lotes,
            valorUtm,
            neto,
            iva,
            total
        };
    }).filter(Boolean);
}

function buildCotizacionAttachmentFileName(summary) {
    const year = Number(summary?.periodo?.anio);
    const month = Number(summary?.periodo?.mes);
    const safeYear = Number.isInteger(year) && year >= 2000 ? year : new Date().getFullYear();
    const safeMonth = Number.isInteger(month) && month >= 1 && month <= 12 ? month : new Date().getMonth() + 1;
    return `Cotizacion_Geo_Rural_${safeYear}_${String(safeMonth).padStart(2, '0')}.pdf`;
}

function getCotizacionTemplateFileAbsolutePath() {
    const normalized = String(COTIZACION_DOCUMENT_PATH || '')
        .replace(/\\/g, '/')
        .replace(/^\/+/, '');
    return path.resolve(process.cwd(), normalized);
}

function resolveCotizacionEmailLogoAttachment() {
    const logoCandidates = [
        path.resolve(process.cwd(), 'logo.png'),
        path.resolve(process.cwd(), 'logo.jpg'),
        path.resolve(process.cwd(), 'logo.jpeg'),
        path.resolve(process.cwd(), 'assets', 'logo.png'),
        path.resolve(process.cwd(), 'assets', 'logo.jpg'),
        path.resolve(process.cwd(), 'assets', 'logo.jpeg')
    ];

    const logoPath = logoCandidates.find((candidatePath) => fs.existsSync(candidatePath));
    if (!logoPath) {
        return null;
    }

    return {
        filename: path.basename(logoPath),
        path: logoPath,
        cid: 'geo-rural-cotizacion-logo'
    };
}

function isPdfBuffer(buffer) {
    if (!Buffer.isBuffer(buffer) || buffer.length < 5) {
        return false;
    }

    return buffer.slice(0, 5).toString('ascii') === '%PDF-';
}

function drawTemplateTextCell(page, font, text, x, y, width, options = {}) {
    const safeText = String(text || '');
    const fontSize = Number.isFinite(Number(options.fontSize)) ? Number(options.fontSize) : 11;
    const align = options.align === 'right' ? 'right' : 'left';
    const cellPadding = Number.isFinite(Number(options.padding)) ? Number(options.padding) : 1.5;
    const clearBackground = options.clearBackground !== false;
    const clearHeight = Number.isFinite(Number(options.clearHeight)) ? Number(options.clearHeight) : fontSize + 3;
    const clearInsetX = Number.isFinite(Number(options.clearInsetX)) ? Number(options.clearInsetX) : 0.4;
    const clearOffsetY = Number.isFinite(Number(options.clearOffsetY)) ? Number(options.clearOffsetY) : 1.4;
    const textWidth = font.widthOfTextAtSize(safeText, fontSize);
    const drawX = align === 'right' ? Math.max(x + cellPadding, x + width - cellPadding - textWidth) : x + cellPadding;
    const clearX = x + clearInsetX;
    const clearY = y - clearOffsetY;
    const clearWidth = Math.max(2, width - clearInsetX * 2);

    if (clearBackground) {
        page.drawRectangle({
            x: clearX,
            y: clearY,
            width: clearWidth,
            height: clearHeight,
            color: rgb(1, 1, 1)
        });
    }
    page.drawText(safeText, {
        x: drawX,
        y,
        size: fontSize,
        font,
        color: rgb(0, 0, 0)
    });
}

function drawCotizacionUtmSummaryBox(page, fontBold, utmTitle) {
    const tableX = 52;
    const tableWidth = COTIZACION_TABLE_COLUMNS.reduce((sum, column) => sum + Number(column.width || 0), 0);
    const boxX = tableX;
    const boxY = 649.2;
    const boxWidth = tableWidth;
    const boxHeight = 18.6;
    const clearMarginLeft = 28;
    const clearMarginRight = 22;
    const clearMarginVertical = 6;

    // Limpieza de la zona completa para eliminar el recuadro/texto original del PDF base.
    page.drawRectangle({
        x: boxX - clearMarginLeft,
        y: boxY - clearMarginVertical,
        width: boxWidth + clearMarginLeft + clearMarginRight,
        height: boxHeight + clearMarginVertical * 2,
        color: rgb(1, 1, 1)
    });

    // Nuevo recuadro de UTM.
    page.drawRectangle({
        x: boxX,
        y: boxY,
        width: boxWidth,
        height: boxHeight,
        color: rgb(1, 1, 1),
        borderColor: rgb(0, 0, 0),
        borderWidth: 1
    });

    const fontSize = 11;
    const safeTitle = String(utmTitle || '');
    const textWidth = fontBold.widthOfTextAtSize(safeTitle, fontSize);
    const centeredX = boxX + (boxWidth - textWidth) / 2;
    const drawX = Number.isFinite(centeredX) ? Math.max(boxX + 6, centeredX) : boxX + 6;
    const drawY = boxY + (boxHeight - fontSize) / 2 + 0.6;

    page.drawText(safeTitle, {
        x: drawX,
        y: drawY,
        size: fontSize,
        font: fontBold,
        color: rgb(0, 0, 0)
    });
}

function drawCotizacionDynamicTable(page, font, fontBold, utmValue) {
    const firstRowY = Number(COTIZACION_TEMPLATE_ROWS[0]?.y || 572.54);
    const lastRowY = Number(COTIZACION_TEMPLATE_ROWS[COTIZACION_TEMPLATE_ROWS.length - 1]?.y || 315.5);
    const rowStep =
        COTIZACION_TEMPLATE_ROWS.length > 1
            ? Math.abs(Number(COTIZACION_TEMPLATE_ROWS[0].y) - Number(COTIZACION_TEMPLATE_ROWS[1].y))
            : 25.7;
    const rowHeight = Number.isFinite(rowStep) && rowStep > 0 ? rowStep : 25.7;

    const tableX = 52;
    const tableWidth = COTIZACION_TABLE_COLUMNS.reduce((sum, column) => sum + Number(column.width || 0), 0);
    const headerHeight = 22;
    const tableTopY = firstRowY + 21.5;
    const rowsHeight = rowHeight * COTIZACION_TEMPLATE_ROWS.length;
    const tableHeight = headerHeight + rowsHeight;
    const tableBottomY = tableTopY - tableHeight;
    const clearZone = COTIZACION_TEMPLATE_CLEAR_AREAS.table;

    // Se limpia el bloque completo para evitar montajes con valores impresos del PDF base.
    page.drawRectangle({
        x: clearZone.x,
        y: clearZone.y,
        width: clearZone.width,
        height: clearZone.height,
        color: rgb(1, 1, 1)
    });

    // Marco exterior.
    page.drawRectangle({
        x: tableX,
        y: tableBottomY,
        width: tableWidth,
        height: tableHeight,
        borderColor: rgb(0, 0, 0),
        borderWidth: 1
    });

    // Separador de cabecera.
    const headerSeparatorY = tableTopY - headerHeight;
    page.drawLine({
        start: { x: tableX, y: headerSeparatorY },
        end: { x: tableX + tableWidth, y: headerSeparatorY },
        thickness: 1,
        color: rgb(0, 0, 0)
    });

    // Lineas verticales y titulos.
    let columnX = tableX;
    COTIZACION_TABLE_COLUMNS.forEach((column, columnIndex) => {
        const currentWidth = Number(column.width || 0);
        if (columnIndex > 0) {
            page.drawLine({
                start: { x: columnX, y: tableBottomY },
                end: { x: columnX, y: tableTopY },
                thickness: 1,
                color: rgb(0, 0, 0)
            });
        }

        drawTemplateTextCell(page, fontBold, column.title, columnX + 1, tableTopY - 15.5, currentWidth - 2, {
            fontSize: 9,
            align: column.align === 'right' ? 'right' : 'left',
            clearBackground: false,
            padding: 4
        });
        columnX += currentWidth;
    });

    // Filas con datos calculados.
    COTIZACION_TEMPLATE_ROWS.forEach((row, index) => {
        const rowTopY = headerSeparatorY - index * rowHeight;
        const rowBottomY = rowTopY - rowHeight;
        if (index > 0) {
            page.drawLine({
                start: { x: tableX, y: rowTopY },
                end: { x: tableX + tableWidth, y: rowTopY },
                thickness: 0.8,
                color: rgb(0, 0, 0)
            });
        }

        const unitValue = Number(row.valorUtm);
        const netValue = Math.round(unitValue * utmValue);
        const ivaValue = Math.round(netValue * 0.19);
        const totalValue = netValue + ivaValue;
        const rowValues = {
            lotes: String(row.lotes || '-'),
            utm: formatTemplateUtmUnits(unitValue),
            neto: `$ ${formatThousandsDots(netValue)}`,
            iva: `$ ${formatThousandsDots(ivaValue)}`,
            total: `$ ${formatThousandsDots(totalValue)}`
        };

        let cellX = tableX;
        COTIZACION_TABLE_COLUMNS.forEach((column) => {
            const cellWidth = Number(column.width || 0);
            drawTemplateTextCell(page, font, rowValues[column.key] || '-', cellX + 1, rowBottomY + 8.2, cellWidth - 2, {
                fontSize: 10,
                align: column.align === 'right' ? 'right' : 'left',
                clearBackground: false,
                padding: 4
            });
            cellX += cellWidth;
        });
    });
}

async function buildCotizacionTemplatePdfBuffer(summary) {
    const utmValue = Number(summary?.utm?.valor);
    const year = Number(summary?.periodo?.anio);
    const month = Number(summary?.periodo?.mes);
    const safeYear = Number.isInteger(year) && year >= 2000 ? year : new Date().getFullYear();
    const safeMonth = Number.isInteger(month) && month >= 1 && month <= 12 ? month : new Date().getMonth() + 1;

    if (!Number.isFinite(utmValue) || utmValue <= 0) {
        throw new Error('UTM vigente invalida para actualizar la carta de cotizacion.');
    }

    const templateFilePath = getCotizacionTemplateFileAbsolutePath();
    if (!fs.existsSync(templateFilePath)) {
        throw new Error(`No se encontro el documento base de cotizacion: ${templateFilePath}`);
    }

    const templateBytes = fs.readFileSync(templateFilePath);
    const pdfDoc = await PDFLibDocument.load(templateBytes);
    const pages = pdfDoc.getPages();
    if (!pages.length) {
        throw new Error('El documento base de cotizacion no contiene paginas.');
    }

    const firstPage = pages[0];
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    const monthLabel = String(MONTH_NAMES_ES[safeMonth - 1] || 'mes').toUpperCase();
    const utmTitle = `VALOR U.T.M. MES DE ${monthLabel} ${safeYear} ($ ${formatThousandsDots(utmValue)})`;

    drawCotizacionUtmSummaryBox(firstPage, fontBold, utmTitle);

    drawCotizacionDynamicTable(firstPage, font, fontBold, utmValue);

    const bytes = await pdfDoc.save();
    return Buffer.from(bytes);
}

function buildCotizacionParcelamientoTableLines(parcelamientoRows = []) {
    const rows = Array.isArray(parcelamientoRows) ? parcelamientoRows : [];
    const lines = ['TRAMO LOTES | VALOR UTM | NETO CLP | IVA CLP | TOTAL CLP'];
    rows.forEach((row) => {
        lines.push(
            `${row.lotes} | ${formatTemplateUtmUnits(row.valorUtm)} | ${formatCurrencyClp(row.neto)} | ${formatCurrencyClp(
                row.iva
            )} | ${formatCurrencyClp(row.total)}`
        );
    });
    return lines;
}

function escapeHtmlForEmail(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function buildCotizacionParcelamientoTableHtml(parcelamientoRows = [], selectedParcelamiento = null) {
    const rows = Array.isArray(parcelamientoRows) ? parcelamientoRows : [];
    const selectedKey = selectedParcelamiento ? normalizeCotizacionParcelamientoKey(selectedParcelamiento.key || selectedParcelamiento.lotes) : '';
    const bodyRows = rows
        .map((row) => {
            const rowKey = normalizeCotizacionParcelamientoKey(row.key || row.lotes);
            const isSelected = selectedKey && rowKey === selectedKey;
            const rowStyle = isSelected ? 'background:#e8f5ff;font-weight:700;color:#1f74bf;' : '';
            return `<tr style="${rowStyle}"><td style="border:1px solid #d9deea;padding:6px 8px;">${escapeHtmlForEmail(
                row.lotes
            )}</td><td style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">${escapeHtmlForEmail(
                formatTemplateUtmUnits(row.valorUtm)
            )}</td><td style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">${escapeHtmlForEmail(
                formatCurrencyClp(row.neto)
            )}</td><td style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">${escapeHtmlForEmail(
                formatCurrencyClp(row.iva)
            )}</td><td style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">${escapeHtmlForEmail(
                formatCurrencyClp(row.total)
            )}</td></tr>`;
        })
        .join('');

    return `<table role="presentation" cellspacing="0" cellpadding="0" style="width:92%;max-width:680px;margin:0 auto;border-collapse:collapse;font-size:12px;color:#2d84d2;">
        <thead>
            <tr style="background:#e6f4ff;color:#1f74bf;">
                <th style="border:1px solid #d9deea;padding:6px 8px;text-align:left;">Tramo lotes</th>
                <th style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">Valor UTM</th>
                <th style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">Neto CLP</th>
                <th style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">IVA CLP</th>
                <th style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">Total CLP</th>
            </tr>
        </thead>
        <tbody>${bodyRows}</tbody>
    </table>`;
}

function buildCotizacionPostTableParagraphs(periodLabel) {
    const safePeriod = normalizeText(periodLabel) || 'periodo vigente';
    return [
        '2).- Alcance general de los valores informados:',
        'Los valores del cuadro son referenciales, estan expresados en pesos chilenos e incluyen IVA (19%).',
        `Corresponden al periodo ${safePeriod} y pueden variar si cambia el valor oficial de la UTM.`,
        '',
        '3).- Consideraciones para la evaluacion:',
        'El monto definitivo de verificacion se determina con los antecedentes tecnicos del proyecto y el tramo final de lotes.',
        'Si el proyecto cambia de tramo durante la revision, se aplicara el valor del tramo correspondiente.',
        '',
        '4).- Vigencia y contacto:',
        'La presente cotizacion tiene vigencia durante el mes calendario indicado.',
        'Para continuar con su proceso o resolver dudas, puede responder este correo o contactar su sucursal Geo Rural.'
    ];
}

function buildCotizacionFormalSignatureLines() {
    return [
        'Saludos cordiales,',
        '',
        '',
        '',
        '',
        'Pedro Antonio Gerardo Herrera Mendez',
        '',
        'Pedro Ignacio Albornoz Sateler',
        '',
        'GEO RURAL VERIFICACIONES LIMITADA',
        '',
        'Uno Norte N° 801, oficina 306, Talca'
    ];
}

function buildCotizacionEmailText(summary, clientName, referenciaCliente, authUser, options = {}) {
    const nowText = new Date().toLocaleString('es-CL', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    });
    const periodLabel = formatMonthlyPeriodLabel(summary?.periodo?.anio, summary?.periodo?.mes);
    const userName = normalizeText(authUser?.nombre || authUser?.username || 'Secretaria');
    const branch = normalizeText(authUser?.sucursal || 'Sin sucursal');
    const clientNameText = normalizeText(clientName);
    const referenciaText = normalizeText(referenciaCliente);
    const parcelamientoRows =
        Array.isArray(options?.parcelamientoRows) && options.parcelamientoRows.length > 0
            ? options.parcelamientoRows
            : buildCotizacionParcelamientoRows(summary?.utm?.valor);
    const selectedParcelamiento = options?.selectedParcelamiento || null;
    const tableLeadText = selectedParcelamiento
        ? `Adicionalmente, presentamos el siguiente cuadro para el tramo consultado (${selectedParcelamiento.lotes} lotes) en el periodo ${periodLabel}:`
        : `Adicionalmente, presentamos el siguiente cuadro de valores para el periodo ${periodLabel}:`;
    const lines = [
        clientNameText ? `Estimada/o ${clientNameText}:` : 'Estimada/o cliente:',
        '',
        'Junto con saludar, tenemos el agrado de dirigirnos a Ud. a fin de dar respuesta a su consulta.',
        `Fecha: ${nowText}`,
        `Periodo UTM: ${periodLabel}`,
        `UTM vigente: ${formatCurrencyClp(summary?.utm?.valor)}`,
        ''
    ];

    if (referenciaText) {
        lines.push(`Referencia cliente: ${referenciaText}`);
        lines.push('');
    }

    if (selectedParcelamiento) {
        lines.push('1).- Cotizacion por los servicios de verificacion:');
        lines.push(
            `En atencion a su consulta para el tramo de ${selectedParcelamiento.lotes} lotes, el valor referencial con IVA es ${formatCurrencyClp(
                selectedParcelamiento.total
            )}.`
        );
        lines.push('');
    }

    lines.push(
        tableLeadText,
        ...buildCotizacionParcelamientoTableLines(parcelamientoRows),
        '',
        ...buildCotizacionPostTableParagraphs(periodLabel),
        '',
        'Para mas informacion, se adjunta la carta de cotizacion en formato PDF con el detalle completo y la tabla actualizada al periodo correspondiente.',
        '',
        `Atendido por: ${userName} | Sucursal: ${branch}`,
        '',
        ...buildCotizacionFormalSignatureLines()
    );

    return lines.join('\n');
}

function buildCotizacionEmailHtml(summary, clientName, referenciaCliente, authUser, options = {}) {
    const nowText = new Date().toLocaleString('es-CL', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    });
    const periodLabel = formatMonthlyPeriodLabel(summary?.periodo?.anio, summary?.periodo?.mes);
    const userName = normalizeText(authUser?.nombre || authUser?.username || 'Secretaria');
    const branch = normalizeText(authUser?.sucursal || 'Sin sucursal');
    const clientNameText = normalizeText(clientName);
    const referenciaText = normalizeText(referenciaCliente);
    const parcelamientoRows =
        Array.isArray(options?.parcelamientoRows) && options.parcelamientoRows.length > 0
            ? options.parcelamientoRows
            : buildCotizacionParcelamientoRows(summary?.utm?.valor);
    const selectedParcelamiento = options?.selectedParcelamiento || null;
    const greeting = clientNameText ? `Estimada/o ${escapeHtmlForEmail(clientNameText)}:` : 'Estimada/o cliente:';
    const selectedBlock = selectedParcelamiento
        ? `<p style="margin:0 0 10px 0;"><strong>1).- Cotizacion por los servicios de verificacion:</strong></p>
           <p style="margin:0 0 14px 0;">En atencion a su consulta para el tramo de <strong>${escapeHtmlForEmail(
               selectedParcelamiento.lotes
           )} lotes</strong>, el valor referencial con IVA es <strong>${escapeHtmlForEmail(
               formatCurrencyClp(selectedParcelamiento.total)
           )}</strong>.</p>`
        : '';
    const referenceBlock = referenciaText
        ? `<p style="margin:0 0 12px 0;"><strong>Referencia cliente:</strong> ${escapeHtmlForEmail(referenciaText)}</p>`
        : '';
    const tableLeadText = selectedParcelamiento
        ? `Adicionalmente, presentamos el siguiente cuadro para el tramo consultado (${escapeHtmlForEmail(
              selectedParcelamiento.lotes
          )} lotes) en el periodo ${escapeHtmlForEmail(periodLabel)}:`
        : `Adicionalmente, presentamos el siguiente cuadro de valores para el periodo ${escapeHtmlForEmail(periodLabel)}:`;
    const signatureLines = buildCotizacionFormalSignatureLines();
    const signatureHtml = `<div style="margin-top:18px;">
        <p style="margin:0 0 16px 0;">${escapeHtmlForEmail(signatureLines[0])}</p>
        <p style="margin:0 0 10px 0;">${escapeHtmlForEmail(signatureLines[5])}</p>
        <p style="margin:0 0 10px 0;">${escapeHtmlForEmail(signatureLines[7])}</p>
        <p style="margin:0 0 10px 0;font-weight:700;color:#0f2f66;">${escapeHtmlForEmail(signatureLines[9])}</p>
        <p style="margin:0;">${escapeHtmlForEmail(signatureLines[11])}</p>
    </div>`;
    const logoCid = normalizeText(options?.logoCid);
    const logoHtml = logoCid
        ? `<div style="margin-top:18px;text-align:center;">
            <img src="cid:${escapeHtmlForEmail(logoCid)}" alt="Geo Rural" style="max-width:180px;width:46%;height:auto;">
        </div>`
        : '';
    const postTableParagraphsHtml = buildCotizacionPostTableParagraphs(periodLabel)
        .map((line) => {
            if (!line) {
                return '<div style="height:8px;"></div>';
            }
            const isSection = /^\d+\)\.-/.test(line);
            if (isSection) {
                return `<p style="margin:12px 0 6px 0;font-weight:700;color:#1f74bf;">${escapeHtmlForEmail(line)}</p>`;
            }
            return `<p style="margin:0 0 6px 0;color:#2d84d2;">${escapeHtmlForEmail(line)}</p>`;
        })
        .join('');
    const customTemplateHtml = normalizeMailTemplateValue(options?.customHtmlTemplate, MAIL_TEMPLATE_MAX_LENGTH);
    if (customTemplateHtml) {
        const replacements = {
            SALUDO: greeting,
            FECHA: escapeHtmlForEmail(nowText),
            PERIODO_UTM: escapeHtmlForEmail(periodLabel),
            UTM_VIGENTE: escapeHtmlForEmail(formatCurrencyClp(summary?.utm?.valor)),
            REFERENCIA_BLOQUE_HTML: referenceBlock,
            TRAMO_BLOQUE_HTML: selectedBlock,
            TABLA_VALORES_HTML: buildCotizacionParcelamientoTableHtml(parcelamientoRows, selectedParcelamiento),
            PARRAFOS_POST_TABLA_HTML: postTableParagraphsHtml,
            ATENCION_NOMBRE: escapeHtmlForEmail(userName),
            ATENCION_SUCURSAL: escapeHtmlForEmail(branch),
            FIRMA_HTML: signatureHtml,
            LOGO_HTML: logoHtml
        };
        return renderMailTemplateHtml(customTemplateHtml, replacements);
    }

    return `<div style="font-family:Arial,Helvetica,sans-serif;background:#eef8ff;padding:20px 10px;color:#2d84d2;line-height:1.45;font-size:14px;">
        <div style="max-width:780px;margin:0 auto;background:#ffffff;border:1px solid #d6e2ff;border-radius:14px;overflow:hidden;">
            <div style="background:linear-gradient(135deg,#4ea7f8 0%,#2f86d7 100%);padding:16px 22px;color:#ffffff;">
                <p style="margin:0;font-size:18px;font-weight:700;letter-spacing:0.2px;">Carta de Cotizacion Geo Rural</p>
                <p style="margin:6px 0 0 0;font-size:12px;opacity:0.95;">Periodo ${escapeHtmlForEmail(periodLabel)} | UTM vigente ${escapeHtmlForEmail(
                    formatCurrencyClp(summary?.utm?.valor)
                )}</p>
            </div>
            <div style="padding:20px 22px 18px 22px;">
                <p style="margin:0 0 12px 0;">${greeting}</p>
                <p style="margin:0 0 8px 0;">Junto con saludar, tenemos el agrado de dirigirnos a Ud. a fin de dar respuesta a su consulta.</p>
                <p style="margin:0 0 12px 0;color:#2d84d2;">Fecha: ${escapeHtmlForEmail(nowText)}<br>Periodo UTM: ${escapeHtmlForEmail(
                    periodLabel
                )}<br>UTM vigente: ${escapeHtmlForEmail(formatCurrencyClp(summary?.utm?.valor))}</p>
                ${referenceBlock}
                ${selectedBlock}
                <p style="margin:0 0 10px 0;">${tableLeadText}</p>
                ${buildCotizacionParcelamientoTableHtml(parcelamientoRows, selectedParcelamiento)}
                <div style="margin-top:14px;padding:12px 14px;background:#f9fbff;border:1px solid #e4ebfb;border-radius:10px;">
                    ${postTableParagraphsHtml}
                </div>
                <p style="margin:14px 0 8px 0;">Para mas informacion, se adjunta la carta de cotizacion en formato PDF con el detalle completo y la tabla actualizada al periodo correspondiente.</p>
                <div style="margin-top:12px;padding:10px 12px;background:#f5f8ff;border-left:4px solid #2f86d7;border-radius:6px;">
                    <p style="margin:0 0 4px 0;font-weight:700;color:#1f74bf;">Atencion Comercial</p>
                    <p style="margin:0;color:#2d84d2;">Atendido por: ${escapeHtmlForEmail(userName)}<br>Sucursal: ${escapeHtmlForEmail(branch)}</p>
                </div>
                ${signatureHtml}
                ${logoHtml}
            </div>
        </div>
    </div>`;
}

function buildCotizacionEmailBody(summary, clientName, referenciaCliente, authUser, options = {}) {
    return {
        text: buildCotizacionEmailText(summary, clientName, referenciaCliente, authUser, options),
        html: buildCotizacionEmailHtml(summary, clientName, referenciaCliente, authUser, options)
    };
}

function escapeRegExpTemplateToken(value) {
    return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function renderMailTemplateHtml(template, replacements = {}) {
    let html = String(template || '');
    const entries = replacements && typeof replacements === 'object' ? Object.entries(replacements) : [];
    entries.forEach(([token, rawValue]) => {
        const pattern = new RegExp(`{{\\s*${escapeRegExpTemplateToken(token)}\\s*}}`, 'gi');
        html = html.replace(pattern, String(rawValue ?? ''));
    });
    return html;
}

function buildDefaultCotizacionTemplateHtml() {
    return `<!doctype html>
<html lang="es">
<body style="margin:0;padding:0;">
    <div style="font-family:Arial,Helvetica,sans-serif;background:#eef8ff;padding:20px 10px;color:#2d84d2;line-height:1.45;font-size:14px;">
        <div style="max-width:780px;margin:0 auto;background:#ffffff;border:1px solid #d6e2ff;border-radius:14px;overflow:hidden;">
            <div style="background:linear-gradient(135deg,#4ea7f8 0%,#2f86d7 100%);padding:16px 22px;color:#ffffff;">
                <p style="margin:0;font-size:18px;font-weight:700;letter-spacing:0.2px;">Carta de Cotizacion Geo Rural</p>
                <p style="margin:6px 0 0 0;font-size:12px;opacity:0.95;">Periodo {{PERIODO_UTM}} | UTM vigente {{UTM_VIGENTE}}</p>
            </div>
            <div style="padding:20px 22px 18px 22px;">
                <p style="margin:0 0 12px 0;">{{SALUDO}}</p>
                <p style="margin:0 0 8px 0;">Junto con saludar, tenemos el agrado de dirigirnos a Ud. a fin de dar respuesta a su consulta.</p>
                <p style="margin:0 0 12px 0;color:#2d84d2;">Fecha: {{FECHA}}<br>Periodo UTM: {{PERIODO_UTM}}<br>UTM vigente: {{UTM_VIGENTE}}</p>
                {{REFERENCIA_BLOQUE_HTML}}
                {{TRAMO_BLOQUE_HTML}}
                <p style="margin:0 0 10px 0;">Adicionalmente, presentamos el siguiente cuadro para el periodo {{PERIODO_UTM}}:</p>
                {{TABLA_VALORES_HTML}}
                <div style="margin-top:14px;padding:12px 14px;background:#f9fbff;border:1px solid #e4ebfb;border-radius:10px;">
                    {{PARRAFOS_POST_TABLA_HTML}}
                </div>
                <p style="margin:14px 0 8px 0;">Para mas informacion, se adjunta la carta de cotizacion en formato PDF con el detalle completo y la tabla actualizada al periodo correspondiente.</p>
                <div style="margin-top:12px;padding:10px 12px;background:#f5f8ff;border-left:4px solid #2f86d7;border-radius:6px;">
                    <p style="margin:0 0 4px 0;font-weight:700;color:#1f74bf;">Atencion Comercial</p>
                    <p style="margin:0;color:#2d84d2;">Atendido por: {{ATENCION_NOMBRE}}<br>Sucursal: {{ATENCION_SUCURSAL}}</p>
                </div>
                {{FIRMA_HTML}}
                {{LOGO_HTML}}
            </div>
        </div>
    </div>
</body>
</html>`;
}

function buildDefaultFacturaSingleTemplateHtml() {
    return `<!doctype html>
<html lang="es">
<body style="margin:0;padding:16px;background:#eef3fb;font-family:Arial,Helvetica,sans-serif;color:#10213f;">
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="max-width:920px;margin:0 auto;background:#ffffff;border:1px solid #d2dced;border-radius:10px;overflow:hidden;">
        <tr>
            <td style="padding:16px 18px;background:#d7e4f7;border-bottom:1px solid #c2d4ee;">
                <h2 style="margin:0;font-size:20px;line-height:1.3;color:#163463;">{{TITULO}}</h2>
            </td>
        </tr>
        <tr>
            <td style="padding:14px 18px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
                    {{RESUMEN_FILAS_HTML}}
                </table>
                <div style="margin-top:2px;padding-top:8px;">
                    <h3 style="margin:0 0 10px;font-size:16px;color:#163463;">Detalle de factura</h3>
                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
                        {{DETALLE_FILAS_HTML}}
                    </table>
                </div>
            </td>
        </tr>
    </table>
</body>
</html>`;
}

function buildDefaultFacturaPendingTemplateHtml() {
    return `<!doctype html>
<html lang="es">
<body style="margin:0;padding:16px;background:#eef3fb;font-family:Arial,Helvetica,sans-serif;color:#10213f;">
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="max-width:920px;margin:0 auto;background:#ffffff;border:1px solid #d2dced;border-radius:10px;overflow:hidden;">
        <tr>
            <td style="padding:16px 18px;background:#d7e4f7;border-bottom:1px solid #c2d4ee;">
                <h2 style="margin:0;font-size:20px;line-height:1.3;color:#163463;">{{TITULO}}</h2>
            </td>
        </tr>
        <tr>
            <td style="padding:14px 18px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
                    {{RESUMEN_FILAS_HTML}}
                </table>
                <div style="margin-top:2px;padding-top:8px;">
                    <h3 style="margin:0 0 10px;font-size:16px;color:#163463;">Resumen de facturas pendientes</h3>
                    {{TABLA_RESUMEN_HTML}}
                    {{DETALLES_FACTURAS_HTML}}
                </div>
            </td>
        </tr>
    </table>
</body>
</html>`;
}

function getDefaultMailTemplates() {
    return {
        cotizacionHtml: buildDefaultCotizacionTemplateHtml(),
        facturaSingleHtml: buildDefaultFacturaSingleTemplateHtml(),
        facturaPendingHtml: buildDefaultFacturaPendingTemplateHtml()
    };
}

function normalizeMailTemplateValue(value, maxLength = MAIL_TEMPLATE_MAX_LENGTH) {
    const normalized = String(value || '').replace(/\u0000/g, '').trim();
    if (!normalized) {
        return '';
    }
    if (!Number.isInteger(maxLength) || maxLength <= 0) {
        return normalized;
    }
    return normalized.slice(0, maxLength);
}

function extractMailTemplatesFromRow(row) {
    const source = row && typeof row === 'object' ? row : {};
    return {
        cotizacionHtml: normalizeMailTemplateValue(source.cotizacion_html_template ?? source.cotizacionHtml),
        facturaSingleHtml: normalizeMailTemplateValue(source.factura_single_html_template ?? source.facturaSingleHtml),
        facturaPendingHtml: normalizeMailTemplateValue(source.factura_pending_html_template ?? source.facturaPendingHtml)
    };
}

function resolveMailTemplatesForUse(row) {
    const defaults = getDefaultMailTemplates();
    const stored = extractMailTemplatesFromRow(row);
    return {
        cotizacionHtml: stored.cotizacionHtml || defaults.cotizacionHtml,
        facturaSingleHtml: stored.facturaSingleHtml || defaults.facturaSingleHtml,
        facturaPendingHtml: stored.facturaPendingHtml || defaults.facturaPendingHtml
    };
}

function normalizeMailTemplatesPayload(rawTemplates) {
    const source = rawTemplates && typeof rawTemplates === 'object' ? rawTemplates : {};
    const templates = {};
    MAIL_TEMPLATE_KEYS.forEach((key) => {
        templates[key] = normalizeMailTemplateValue(source[key], MAIL_TEMPLATE_MAX_LENGTH);
    });
    return templates;
}

function buildCotizacionPdfBuffer(summary, referenciaCliente, authUser) {
    return new Promise((resolve, reject) => {
        const doc = new PDFDocument({
            size: 'A4',
            margin: 48
        });
        const chunks = [];
        doc.on('data', (chunk) => chunks.push(chunk));
        doc.on('end', () => resolve(Buffer.concat(chunks)));
        doc.on('error', reject);

        const pageBottom = () => doc.page.height - doc.page.margins.bottom;
        const periodLabel = formatMonthlyPeriodLabel(summary?.periodo?.anio, summary?.periodo?.mes);
        const nowText = new Date().toLocaleString('es-CL', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: false
        });
        const userName = normalizeText(authUser?.nombre || authUser?.username || 'Secretaria');
        const branch = normalizeText(authUser?.sucursal || 'Sin sucursal');

        doc.font('Helvetica-Bold').fontSize(18).text('Carta de Cotizacion Geo Rural', {
            align: 'center'
        });
        doc.moveDown(0.4);
        doc.font('Helvetica').fontSize(11).text(`Periodo UTM: ${periodLabel}`);
        doc.text(`Valor UTM vigente: ${formatCurrencyClp(summary?.utm?.valor)}`);
        doc.text(`Referencia cliente: ${normalizeText(referenciaCliente)}`);
        doc.text(`Fecha emision: ${nowText}`);
        doc.moveDown(0.6);

        const colWidths = [290, 90, 120];
        const startX = doc.page.margins.left;
        const totalWidth = colWidths.reduce((sum, value) => sum + value, 0);
        let y = doc.y;

        const ensureSpace = (heightNeeded) => {
            if (y + heightNeeded <= pageBottom()) {
                return;
            }

            doc.addPage();
            y = doc.page.margins.top;
        };

        const drawTableRow = (cells, options = {}) => {
            const isHeader = options.header === true;
            const alignments = options.alignments || ['left', 'right', 'right'];
            const cellHeights = cells.map((value, index) =>
                doc.heightOfString(String(value || ''), {
                    width: colWidths[index] - 10,
                    align: alignments[index] || 'left'
                })
            );
            const rowHeight = Math.max(24, ...cellHeights.map((height) => height + 8));
            ensureSpace(rowHeight);

            if (isHeader) {
                doc.rect(startX, y, totalWidth, rowHeight).fillOpacity(0.12).fillAndStroke('#2a3d6b', '#8fa6c9');
                doc.fillOpacity(1);
            }

            let x = startX;
            cells.forEach((value, index) => {
                doc.rect(x, y, colWidths[index], rowHeight).stroke('#8fa6c9');
                doc.font(isHeader ? 'Helvetica-Bold' : 'Helvetica')
                    .fontSize(10)
                    .fillColor('#2a3d6b')
                    .text(String(value || ''), x + 5, y + 4, {
                        width: colWidths[index] - 10,
                        align: alignments[index] || 'left'
                    });
                x += colWidths[index];
            });
            y += rowHeight;
        };

        drawTableRow(['Item servicio', 'Valor UTM', 'Valor CLP'], {
            header: true,
            alignments: ['left', 'center', 'center']
        });

        const servicios = Array.isArray(summary?.servicios) ? summary.servicios : [];
        servicios.forEach((item) => {
            drawTableRow(
                [item.nombreServicio || '-', formatUtmUnits(item.valorUtm), formatCurrencyClp(item.valorPesos)],
                {
                    alignments: ['left', 'right', 'right']
                }
            );
        });

        y += 12;
        ensureSpace(70);
        doc.font('Helvetica-Bold').fontSize(11).fillColor('#2a3d6b');
        doc.text(`Total estandar: ${formatUtmUnits(summary?.totalUtm)}`, startX, y);
        y += 18;
        doc.text(`Equivalencia referencial CLP: ${formatCurrencyClp(summary?.totalPesos)}`, startX, y);
        y += 26;
        doc.font('Helvetica').fontSize(10);
        doc.text(`Emitido por: ${userName}`, startX, y);
        y += 14;
        doc.text(`Sucursal: ${branch}`, startX, y);
        y += 14;
        doc.text('Documento generado automaticamente por sistema Geo Rural.', startX, y);

        doc.end();
    });
}

function isSmtpConfigReady(config) {
    const source = config && typeof config === 'object' ? config : {};
    const host = normalizeText(source.host);
    const fromEmail = normalizeText(source.fromEmail).toLowerCase();
    const port = Number(source.port);

    if (!host || !fromEmail || !Number.isInteger(port) || port <= 0 || port > 65535) {
        return false;
    }

    return isValidEmail(fromEmail);
}

function toSmtpRuntimeConfig(source) {
    const raw = source && typeof source === 'object' ? source : {};
    const host = normalizeText(raw.host ?? raw.smtp_host ?? raw.smtpHost);
    const port = parsePositiveInt(raw.port ?? raw.smtp_port ?? raw.smtpPort, 587);
    const secure = parseBooleanFlag(raw.secure ?? raw.smtp_secure ?? raw.smtpSecure, false);
    const user = normalizeText(raw.user ?? raw.smtp_user ?? raw.smtpUser);
    const password = String(raw.password ?? raw.smtp_password ?? raw.smtpPassword ?? '');
    const rawFromEmail = raw.fromEmail ?? raw.smtp_from_email ?? raw.smtpFromEmail ?? user;
    const rawFromName = raw.fromName ?? raw.smtp_from_name ?? raw.smtpFromName ?? 'Geo Rural';
    const fromEmail = normalizeText(rawFromEmail).toLowerCase();
    const fromName = normalizeText(rawFromName) || 'Geo Rural';

    return {
        host,
        port,
        secure,
        user,
        password,
        fromEmail,
        fromName
    };
}

function toSmtpPublicConfig(config, options = {}) {
    const source = toSmtpRuntimeConfig(config);
    const includePassword = options.includePassword === true;

    return {
        smtpHost: source.host,
        smtpPort: source.port,
        smtpSecure: source.secure,
        smtpUser: source.user,
        smtpPassword: includePassword ? source.password : '',
        smtpFromEmail: source.fromEmail,
        smtpFromName: source.fromName,
        hasPassword: Boolean(source.password)
    };
}

function buildSmtpTransportKey(config) {
    const source = toSmtpRuntimeConfig(config);
    return [source.host, source.port, source.secure ? '1' : '0', source.user, source.password].join('|');
}

function resetSmtpTransporterCache() {
    smtpTransporter = null;
    smtpTransporterKey = '';
}

function buildSmtpFromAddress(config) {
    const source = toSmtpRuntimeConfig(config);
    return source.fromName ? `${source.fromName} <${source.fromEmail}>` : source.fromEmail;
}

function normalizeMailSubjectLine(value) {
    return String(value || '')
        .replace(/[\r\n]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function normalizeMailTextContent(value) {
    return String(value || '')
        .replace(/\r\n/g, '\n')
        .replace(/\u0000/g, '')
        .trim();
}

function normalizeMailHtmlContent(value) {
    return String(value || '').replace(/\u0000/g, '').trim();
}

function getSmtpTransporter(config) {
    const runtimeConfig = toSmtpRuntimeConfig(config);
    if (!isSmtpConfigReady(runtimeConfig)) {
        return null;
    }

    const nextKey = buildSmtpTransportKey(runtimeConfig);
    if (smtpTransporter && smtpTransporterKey === nextKey) {
        return smtpTransporter;
    }

    const transportConfig = {
        host: runtimeConfig.host,
        port: runtimeConfig.port,
        secure: runtimeConfig.secure
    };
    if (runtimeConfig.user) {
        transportConfig.auth = {
            user: runtimeConfig.user,
            pass: runtimeConfig.password
        };
    }

    smtpTransporter = nodemailer.createTransport(transportConfig);
    smtpTransporterKey = nextKey;
    return smtpTransporter;
}

async function getStoredCorreoConfigurationRow() {
    const [rows] = await pool.execute(
        `
            SELECT
                id,
                smtp_host,
                smtp_port,
                smtp_secure,
                smtp_user,
                smtp_password,
                smtp_from_email,
                smtp_from_name,
                cotizacion_html_template,
                factura_single_html_template,
                factura_pending_html_template,
                updated_by_id,
                updated_by_nombre,
                created_at,
                updated_at
            FROM correo_configuracion
            WHERE id = ?
            LIMIT 1
        `,
        [SMTP_CONFIG_SINGLETON_ID]
    );

    return rows.length > 0 ? rows[0] : null;
}

async function resolveActiveSmtpConfiguration() {
    const storedRow = await getStoredCorreoConfigurationRow();
    if (storedRow) {
        const storedConfig = toSmtpRuntimeConfig(storedRow);
        if (isSmtpConfigReady(storedConfig)) {
            return {
                source: 'DATABASE',
                config: storedConfig,
                storedRow
            };
        }
    }

    const envConfig = toSmtpRuntimeConfig(SMTP_ENV_CONFIG);
    if (isSmtpConfigReady(envConfig)) {
        return {
            source: 'ENV',
            config: envConfig,
            storedRow
        };
    }

    return {
        source: 'NONE',
        config: null,
        storedRow
    };
}

async function buildCurrentCotizacionSummary() {
    const status = await getCurrentMonthUtmStatus();
    const utmValue = Number(status?.utm?.valor);
    if (!Number.isFinite(utmValue) || utmValue <= 0) {
        return {
            ok: false,
            reason: 'UTM_REQUIRED',
            status
        };
    }

    const [rows] = await pool.execute(
        `
            SELECT
                id,
                nombre_servicio,
                valor_utm,
                orden_visual
            FROM cotizacion_servicios_estandar
            WHERE activo = 1
            ORDER BY orden_visual ASC, id ASC
        `
    );

    const servicios = rows.map((row) => toCotizacionServicioResumen(row, utmValue)).filter(Boolean);
    const parcelamientoRows = buildCotizacionParcelamientoRows(utmValue);
    const totalUtm = servicios.reduce((sum, item) => sum + Number(item.valorUtm || 0), 0);
    const totalPesos = servicios.reduce((sum, item) => sum + Number(item.valorPesos || 0), 0);
    const documentUrl = encodeURI(COTIZACION_DOCUMENT_PATH);

    return {
        ok: true,
        summary: {
            periodo: {
                anio: Number(status.anio),
                mes: Number(status.mes),
                dia: Number(status.dia)
            },
            utm: {
                valor: Number(utmValue.toFixed(2)).toFixed(2),
                registradoPor: normalizeText(status?.utm?.registradoPor),
                updatedAt: normalizeText(status?.utm?.updatedAt)
            },
            parcelamiento: parcelamientoRows.map((row) => ({
                key: row.key,
                lotes: row.lotes,
                valorUtm: Number(row.valorUtm.toFixed(2)),
                neto: row.neto,
                iva: row.iva,
                total: row.total
            })),
            servicios,
            totalUtm: totalUtm.toFixed(2),
            totalPesos: totalPesos.toFixed(2),
            documento: {
                nombre: path.basename(COTIZACION_DOCUMENT_PATH),
                url: documentUrl
            },
            suggestedUtm: status?.suggestedUtm || null
        }
    };
}

async function getCurrentMonthUtmStatus() {
    const period = getCurrentCalendarPeriod();
    const [rows] = await pool.execute(
        `
            SELECT
                anio,
                mes,
                valor,
                registrado_por_nombre,
                created_at,
                updated_at
            FROM utm_mensual
            WHERE anio = ? AND mes = ?
            LIMIT 1
        `,
        [period.anio, period.mes]
    );

    const currentRow = rows.length > 0 ? rows[0] : null;
    const suggestedUtm = await fetchUtmSuggestionForPeriod(period.anio, period.mes);
    return {
        anio: period.anio,
        mes: period.mes,
        dia: period.dia,
        // Debe exigir captura solo cuando no existe UTM del periodo actual.
        required: !currentRow,
        existeRegistro: Boolean(currentRow),
        utm: currentRow ? toUtmMensualRow(currentRow) : null,
        suggestedUtm: suggestedUtm || null
    };
}

function requireWriteAccess(req, res, next) {
    if (isGuestUser(req.authUser)) {
        return res.status(403).json({ message: 'El modo invitado es solo lectura.' });
    }

    if (!API_WRITE_KEY) {
        return next();
    }

    const incomingKey = normalizeText(req.get('x-api-key'));
    if (incomingKey && incomingKey === API_WRITE_KEY) {
        return next();
    }

    return res.status(401).json({
        message: 'No autorizado para operaciones de escritura. Falta o es invalido el header X-API-Key.',
        code: 'WRITE_API_KEY_REQUIRED'
    });
}

function enforceAllowedOrigin(req, res, next) {
    const origin = normalizeText(req.get('origin'));
    if (!origin || isAllowedOrigin(origin)) {
        return next();
    }

    return res.status(403).json({ message: 'Origen no permitido.' });
}

function isAllowedOrigin(origin) {
    if (!origin) {
        return true;
    }

    if (ALLOWED_ORIGINS.includes(origin)) {
        return true;
    }

    if (isLocalhostOrigin(origin) && ALLOWED_ORIGINS.some((item) => isLocalhostOrigin(item))) {
        return true;
    }

    if (!IS_PRODUCTION_ENV && isPrivateNetworkOrigin(origin)) {
        return true;
    }

    return false;
}

function isLocalhostOrigin(value) {
    try {
        const parsed = new URL(value);
        return parsed.hostname === 'localhost' || parsed.hostname === '127.0.0.1';
    } catch (error) {
        return false;
    }
}

function isPrivateIpv4Host(hostname) {
    const match = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(String(hostname || '').trim());
    if (!match) {
        return false;
    }

    const octets = match.slice(1).map((part) => Number(part));
    if (octets.some((value) => !Number.isInteger(value) || value < 0 || value > 255)) {
        return false;
    }

    if (octets[0] === 10) {
        return true;
    }
    if (octets[0] === 127) {
        return true;
    }
    if (octets[0] === 192 && octets[1] === 168) {
        return true;
    }
    if (octets[0] === 172 && octets[1] >= 16 && octets[1] <= 31) {
        return true;
    }

    return false;
}

function isPrivateNetworkOrigin(value) {
    try {
        const parsed = new URL(value);
        const protocol = String(parsed.protocol || '').toLowerCase();
        if (protocol !== 'http:' && protocol !== 'https:') {
            return false;
        }

        const hostname = String(parsed.hostname || '').trim().toLowerCase();
        return hostname === 'localhost' || isPrivateIpv4Host(hostname);
    } catch (error) {
        return false;
    }
}

function getClientIp(req) {
    const forwarded = normalizeText(req.get('x-forwarded-for'));
    if (forwarded) {
        return normalizeText(forwarded.split(',')[0]);
    }

    return normalizeText(req.ip || req.socket?.remoteAddress || 'unknown');
}

function getLoginAttemptKey(req, username) {
    return `${getClientIp(req)}|${String(username || '').toLowerCase()}`;
}

function pruneLoginAttempts(now = Date.now()) {
    const maxAgeMs = Math.max(LOGIN_WINDOW_MINUTES, LOGIN_LOCK_MINUTES) * 60 * 1000 * 2;
    for (const [key, entry] of loginAttempts.entries()) {
        if (!entry) {
            loginAttempts.delete(key);
            continue;
        }

        if (entry.blockedUntil && entry.blockedUntil > now) {
            continue;
        }

        if (now - entry.windowStartedAt > maxAgeMs) {
            loginAttempts.delete(key);
        }
    }
}

function getLoginRateLimitState(req, username) {
    const now = Date.now();
    pruneLoginAttempts(now);
    const key = getLoginAttemptKey(req, username);
    const entry = loginAttempts.get(key);

    if (!entry) {
        return { blocked: false, retryAfterSeconds: 0 };
    }

    if (entry.blockedUntil && entry.blockedUntil > now) {
        return {
            blocked: true,
            retryAfterSeconds: Math.max(1, Math.ceil((entry.blockedUntil - now) / 1000))
        };
    }

    if (entry.blockedUntil && entry.blockedUntil <= now) {
        loginAttempts.delete(key);
    }

    return { blocked: false, retryAfterSeconds: 0 };
}

function registerFailedLoginAttempt(req, username) {
    const now = Date.now();
    const key = getLoginAttemptKey(req, username);
    const existing = loginAttempts.get(key);
    const windowMs = LOGIN_WINDOW_MINUTES * 60 * 1000;
    const lockMs = LOGIN_LOCK_MINUTES * 60 * 1000;

    let entry = existing;
    if (!entry || now - entry.windowStartedAt > windowMs) {
        entry = {
            attempts: 0,
            windowStartedAt: now,
            blockedUntil: 0
        };
    }

    entry.attempts += 1;
    if (entry.attempts >= LOGIN_MAX_ATTEMPTS) {
        entry.attempts = 0;
        entry.windowStartedAt = now;
        entry.blockedUntil = now + lockMs;
    }

    loginAttempts.set(key, entry);
}

function clearFailedLoginAttempts(req, username) {
    const key = getLoginAttemptKey(req, username);
    loginAttempts.delete(key);
}

function setMaintenanceModeCacheFromRow(row) {
    const source = row && typeof row === 'object' ? row : {};
    const enabledValue = Number(source.maintenance_mode || source.maintenanceMode || 0);
    const updatedAt = source.updated_at instanceof Date ? source.updated_at.toISOString() : normalizeText(source.updated_at || '');
    const updatedBy = normalizeText(source.updated_by_name || source.updatedBy || '');

    maintenanceModeCache = {
        enabled: enabledValue === 1,
        updatedAt,
        updatedBy
    };
}

function normalizeLoginRouteNoticeUrl(value) {
    const raw = normalizeText(value || LOGIN_ROUTE_NOTICE_DEFAULT_URL);
    if (!raw) {
        return normalizeText(LOGIN_ROUTE_NOTICE_DEFAULT_URL || '/geo_rural/registro/login/registro.html');
    }

    if (IS_PRODUCTION_ENV) {
        if (
            /^https?:\/\/localhost(?:\:\d+)?(?:\/|$)/i.test(raw) ||
            /^https?:\/\/127\.0\.0\.1(?:\:\d+)?(?:\/|$)/i.test(raw) ||
            /^localhost(?:\:\d+)?(?:\/|$)/i.test(raw) ||
            /^127\.0\.0\.1(?:\:\d+)?(?:\/|$)/i.test(raw) ||
            /^https?:\/\/m[iy]-registro\.cl(?:\/|$)/i.test(raw) ||
            /^m[iy]-registro\.cl(?:\/|$)/i.test(raw)
        ) {
            return normalizeText(LOGIN_ROUTE_NOTICE_PROD_URL || LOGIN_ROUTE_NOTICE_DEFAULT_URL || '/geo_rural/registro/login/registro.html');
        }
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

function toIsoStringOrEmpty(value) {
    if (!value) {
        return '';
    }

    if (value instanceof Date) {
        return Number.isNaN(value.getTime()) ? '' : value.toISOString();
    }

    const text = normalizeText(value);
    if (!text) {
        return '';
    }

    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? '' : parsed.toISOString();
}

function setLoginRouteNoticeCacheFromRow(row) {
    const source = row && typeof row === 'object' ? row : {};
    loginRouteNoticeCache = {
        enabled: Number(source.login_notice_enabled || source.loginNoticeEnabled || 0) === 1,
        url: normalizeLoginRouteNoticeUrl(source.login_notice_url || source.loginNoticeUrl || LOGIN_ROUTE_NOTICE_DEFAULT_URL),
        message: normalizeText(source.login_notice_message || source.loginNoticeMessage || LOGIN_ROUTE_NOTICE_DEFAULT_MESSAGE),
        startsAt: toIsoStringOrEmpty(source.login_notice_starts_at || source.loginNoticeStartsAt || ''),
        expiresAt: toIsoStringOrEmpty(source.login_notice_expires_at || source.loginNoticeExpiresAt || ''),
        updatedAt: toIsoStringOrEmpty(source.login_notice_updated_at || source.loginNoticeUpdatedAt || ''),
        updatedBy: normalizeText(source.login_notice_updated_by_name || source.loginNoticeUpdatedBy || '')
    };
}

function resolveLoginRouteNoticeStatusPayload() {
    const enabled = Boolean(loginRouteNoticeCache.enabled);
    const startsAt = normalizeText(loginRouteNoticeCache.startsAt || '');
    const expiresAt = normalizeText(loginRouteNoticeCache.expiresAt || '');
    const now = Date.now();
    const expiresAtMs = expiresAt ? new Date(expiresAt).getTime() : Number.NaN;
    const hasValidExpiration = Number.isFinite(expiresAtMs);
    const active = enabled && (!hasValidExpiration || expiresAtMs > now);
    const daysRemaining =
        active && hasValidExpiration ? Math.max(1, Math.ceil((expiresAtMs - now) / (24 * 60 * 60 * 1000))) : 0;

    return {
        enabled,
        active,
        daysRemaining,
        url: normalizeLoginRouteNoticeUrl(loginRouteNoticeCache.url || LOGIN_ROUTE_NOTICE_DEFAULT_URL),
        message: normalizeText(loginRouteNoticeCache.message || LOGIN_ROUTE_NOTICE_DEFAULT_MESSAGE),
        startsAt,
        expiresAt,
        updatedAt: normalizeText(loginRouteNoticeCache.updatedAt || ''),
        updatedBy: normalizeText(loginRouteNoticeCache.updatedBy || '')
    };
}

function getLoginRouteNoticePublicPayload() {
    const status = resolveLoginRouteNoticeStatusPayload();
    return {
        active: status.active,
        url: status.url,
        message: status.message,
        startsAt: status.startsAt,
        expiresAt: status.expiresAt,
        daysRemaining: status.daysRemaining
    };
}

function getLoginRouteNoticeAdminPayload() {
    return resolveLoginRouteNoticeStatusPayload();
}

function normalizeHostName(value) {
    const host = normalizeText(value || '').toLowerCase();
    if (!host) {
        return '';
    }

    const bracketIpv6Match = /^\[([^\]]+)\](?::\d+)?$/.exec(host);
    if (bracketIpv6Match) {
        return normalizeText(bracketIpv6Match[1]);
    }

    if (/^[^:]+:\d+$/.test(host)) {
        return host.replace(/:\d+$/, '');
    }

    return host;
}

function isLocalhostHostName(hostname) {
    const host = normalizeHostName(hostname);
    return host === 'localhost' || host === '127.0.0.1' || host === '::1';
}

function getRequestHostName(req) {
    const forwardedHost = normalizeText(req?.get?.('x-forwarded-host'));
    const rawHost = forwardedHost || normalizeText(req?.get?.('host'));
    if (!rawHost) {
        return '';
    }

    const firstHost = normalizeText(rawHost.split(',')[0]);
    return normalizeHostName(firstHost);
}

function getRequestProtocol(req) {
    const forwardedProto = normalizeText(req?.get?.('x-forwarded-proto'));
    if (forwardedProto) {
        const first = normalizeText(forwardedProto.split(',')[0]).toLowerCase();
        if (first === 'http' || first === 'https') {
            return first;
        }
    }

    const requestProto = normalizeText(req?.protocol).toLowerCase();
    if (requestProto === 'http' || requestProto === 'https') {
        return requestProto;
    }

    return 'https';
}

function normalizeComparablePathname(pathname) {
    const raw = normalizeText(pathname || '');
    if (!raw) {
        return '/';
    }

    const withoutQuery = raw.split('?')[0].split('#')[0];
    const withLeadingSlash = withoutQuery.startsWith('/') ? withoutQuery : `/${withoutQuery}`;
    const compacted = withLeadingSlash.replace(/\/{2,}/g, '/');
    if (compacted.length > 1 && compacted.endsWith('/')) {
        return compacted.slice(0, -1);
    }
    return compacted || '/';
}

function extractPathnameFromRedirectTarget(target) {
    const normalizedTarget = normalizeText(target || '');
    if (!normalizedTarget) {
        return '';
    }

    if (/^https?:\/\//i.test(normalizedTarget)) {
        try {
            return normalizeComparablePathname(new URL(normalizedTarget).pathname || '/');
        } catch (error) {
            return '';
        }
    }

    if (normalizedTarget.startsWith('/')) {
        return normalizeComparablePathname(normalizedTarget);
    }

    return '';
}

function isLoopbackRedirectTarget(target) {
    const value = normalizeText(target || '');
    if (!value) {
        return false;
    }

    return (
        /^https?:\/\/localhost(?:\:\d+)?(?:\/|$)/i.test(value) ||
        /^https?:\/\/127\.0\.0\.1(?:\:\d+)?(?:\/|$)/i.test(value) ||
        /^https?:\/\/\[::1\](?:\:\d+)?(?:\/|$)/i.test(value) ||
        /^localhost(?:\:\d+)?(?:\/|$)/i.test(value) ||
        /^127\.0\.0\.1(?:\:\d+)?(?:\/|$)/i.test(value) ||
        /^\[::1\](?:\:\d+)?(?:\/|$)/i.test(value)
    );
}

function resolveConfiguredLoginRedirectTarget(requestPathname = '', req = null) {
    const fallbackTarget = '/geo_rural/registro/login/registro.html';
    const configuredTarget = normalizeLoginRouteNoticeUrl(
        loginRouteNoticeCache.url || LOGIN_ROUTE_NOTICE_DEFAULT_URL || fallbackTarget
    );
    if (!configuredTarget) {
        return fallbackTarget;
    }

    const requestComparablePath = normalizeComparablePathname(requestPathname);
    const configuredPathname = extractPathnameFromRedirectTarget(configuredTarget);
    if (configuredPathname && LOGIN_REDIRECT_ENTRY_PATHS.has(configuredPathname)) {
        return fallbackTarget;
    }
    if (configuredPathname && configuredPathname === requestComparablePath) {
        return fallbackTarget;
    }

    const requestHostName = getRequestHostName(req);
    if (requestHostName && !isLocalhostHostName(requestHostName) && isLoopbackRedirectTarget(configuredTarget)) {
        const protocol = getRequestProtocol(req);
        return `${protocol}://${requestHostName}${fallbackTarget}`;
    }

    return configuredTarget;
}

function getMaintenanceModeStatusPayload() {
    return {
        enabled: Boolean(maintenanceModeCache.enabled),
        updatedAt: normalizeText(maintenanceModeCache.updatedAt || ''),
        updatedBy: normalizeText(maintenanceModeCache.updatedBy || '')
    };
}

function isMaintenanceModeEnabled() {
    return Boolean(maintenanceModeCache.enabled);
}

function shouldAutoDisableLoginRouteNotice() {
    const status = resolveLoginRouteNoticeStatusPayload();
    return status.enabled && !status.active;
}

function normalizeRequestPath(req) {
    const rawPath = normalizeText(req?.path || req?.originalUrl || '/');
    if (!rawPath) {
        return '/';
    }

    const queryIndex = rawPath.indexOf('?');
    return queryIndex >= 0 ? rawPath.slice(0, queryIndex) : rawPath;
}

function isApiRequestPath(pathname) {
    return String(pathname || '').startsWith('/api/');
}

function isPublicFrontendPath(pathname) {
    return PUBLIC_FRONTEND_PATHS.has(String(pathname || ''));
}

function isKnownFrontendPath(pathname) {
    const normalizedPath = String(pathname || '');
    if (!normalizedPath) {
        return false;
    }

    if (KNOWN_FRONTEND_PATHS.has(normalizedPath)) {
        return true;
    }

    if (normalizedPath === '/assets' || normalizedPath.startsWith('/assets/')) {
        return true;
    }

    return false;
}

function shouldProtectFrontendPath(pathname) {
    const normalizedPath = String(pathname || '');
    if (!isKnownFrontendPath(normalizedPath)) {
        return false;
    }
    return !isPublicFrontendPath(normalizedPath);
}

async function resolveMaintenanceSessionUser(req) {
    if (!pool) {
        return null;
    }

    const token = extractBearerToken(req);
    if (!token) {
        return null;
    }

    const session = await findSessionUserByToken(token);
    if (!session || !session.user) {
        return null;
    }

    req.authUser = session.user;
    req.authSessionId = session.sessionId;
    req.authTokenHash = session.tokenHash;
    return session.user;
}

async function enforceMaintenanceMode(req, res, next) {
    if (!isMaintenanceModeEnabled()) {
        return next();
    }

    if (req.method === 'OPTIONS') {
        return next();
    }

    const pathname = normalizeRequestPath(req);
    if (MAINTENANCE_ALLOWED_PAGE_PATHS.has(pathname)) {
        return next();
    }

    let sessionUser = null;
    try {
        sessionUser = await resolveMaintenanceSessionUser(req);
    } catch (error) {
        sessionUser = null;
    }

    if (isSuperUser(sessionUser)) {
        return next();
    }

    if (isApiRequestPath(pathname)) {
        if (MAINTENANCE_ALLOWED_API_PATHS.has(pathname)) {
            return next();
        }

        return res.status(503).json({
            message: 'El sitio se encuentra en mantencion. Intenta nuevamente mas tarde.',
            code: 'SITE_MAINTENANCE'
        });
    }

    return res.redirect('/mantenimiento.html');
}

async function enforceFrontendRouteProtection(req, res, next) {
    if (req.method !== 'GET' && req.method !== 'HEAD') {
        return next();
    }

    const pathname = normalizeRequestPath(req);
    if (isApiRequestPath(pathname)) {
        return next();
    }

    if (!shouldProtectFrontendPath(pathname)) {
        return next();
    }

    let sessionUser = null;
    try {
        sessionUser = await resolveMaintenanceSessionUser(req);
    } catch (error) {
        sessionUser = null;
    }

    if (sessionUser) {
        return next();
    }

    return res.redirect('/index.html');
}

async function cleanupExpiredSessionsIfNeeded(force = false) {
    const now = Date.now();
    const intervalMs = SESSION_CLEANUP_INTERVAL_MINUTES * 60 * 1000;
    if (!force && now - lastSessionCleanupAt < intervalMs) {
        return;
    }

    await pool.query('DELETE FROM sesiones WHERE expires_at <= NOW() OR revoked_at IS NOT NULL');
    lastSessionCleanupAt = now;
}

async function findSessionUserByToken(token) {
    const tokenHash = hashToken(token);
    const [rows] = await pool.execute(
        `
            SELECT
                s.id AS session_id,
                u.id AS user_id,
                u.username,
                u.nombre,
                u.sucursal,
                u.role,
                u.must_change_password,
                u.activo
            FROM sesiones s
            INNER JOIN usuarios u ON u.id = s.user_id
            WHERE s.token_hash = ?
              AND s.revoked_at IS NULL
              AND s.expires_at > NOW()
            LIMIT 1
        `,
        [tokenHash]
    );

    if (!rows.length || Number(rows[0].activo) !== 1) {
        return null;
    }

    await pool.execute(
        'UPDATE sesiones SET last_used_at = NOW(), expires_at = DATE_ADD(NOW(), INTERVAL ? HOUR) WHERE id = ?',
        [AUTH_SESSION_HOURS, rows[0].session_id]
    );

    return {
        sessionId: rows[0].session_id,
        tokenHash,
        user: formatAuthUser(rows[0])
    };
}

async function requireAuth(req, res, next) {
    try {
        const token = extractBearerToken(req);
        if (!token) {
            clearSessionCookie(res);
            console.info(`[Auth] Sesion rechazada: token ausente. IP=${getClientIp(req)}`);
            return res.status(401).json({ message: 'Debes iniciar sesion para usar el sistema.' });
        }

        const session = await findSessionUserByToken(token);
        if (!session) {
            clearSessionCookie(res);
            console.info(`[Auth] Sesion rechazada: token invalido o expirado. IP=${getClientIp(req)}`);
            return res.status(401).json({ message: 'Sesion invalida o expirada. Inicia sesion nuevamente.' });
        }

        req.authUser = session.user;
        req.authSessionId = session.sessionId;
        req.authTokenHash = session.tokenHash;
        setSessionCookie(res, token);
        return next();
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible validar la sesion.' });
    }
}

function requireAdmin(req, res, next) {
    if (!isAdminOrSuperUser(req.authUser)) {
        return res.status(403).json({ message: 'Solo un administrador o superusuario puede realizar esta accion.' });
    }

    return next();
}

function requireSuper(req, res, next) {
    if (!isSuperUser(req.authUser)) {
        return res.status(403).json({ message: 'Solo el superusuario puede realizar esta accion.' });
    }

    return next();
}

function requirePasswordChangeCompleted(req, res, next) {
    if (req.authUser && req.authUser.mustChangePassword) {
        return res.status(403).json({
            message: 'Debes cambiar tu clave temporal antes de continuar.',
            code: 'PASSWORD_CHANGE_REQUIRED'
        });
    }

    return next();
}

async function ensureColumn(tableName, columnName, definitionSql) {
    const safeTable = cleanDbIdentifier(tableName);
    const safeColumn = cleanDbIdentifier(columnName);
    const [columns] = await pool.query(`SHOW COLUMNS FROM \`${safeTable}\` LIKE '${safeColumn}'`);
    if (columns.length === 0) {
        await pool.query(`ALTER TABLE \`${safeTable}\` ADD COLUMN ${definitionSql}`);
    }
}

async function ensureUniqueCorrelativoIndex() {
    const [indexes] = await pool.query("SHOW INDEX FROM registros WHERE Key_name = 'uq_anio_correlativo'");
    if (indexes.length > 0) {
        return;
    }

    await pool.query('ALTER TABLE registros ADD UNIQUE KEY uq_anio_correlativo (anio, correlativo)');
}

async function ensureRegistrosAuditColumns() {
    await ensureColumn('registros', 'created_by', '`created_by` VARCHAR(120) NULL AFTER comentario');
    await ensureColumn('registros', 'created_by_sucursal', '`created_by_sucursal` VARCHAR(120) NULL AFTER created_by');
    await ensureColumn('registros', 'updated_by', '`updated_by` VARCHAR(120) NULL AFTER created_by_sucursal');
    await ensureColumn('registros', 'updated_by_sucursal', '`updated_by_sucursal` VARCHAR(120) NULL AFTER updated_by');
}

async function ensureRegistrosFacturaColumns() {
    await ensureColumn('registros', 'factura_nombre_razon', '`factura_nombre_razon` VARCHAR(255) NULL AFTER comentario');
    await ensureColumn('registros', 'factura_numero', '`factura_numero` VARCHAR(80) NULL AFTER factura_nombre_razon');
    await ensureColumn('registros', 'factura_rut', '`factura_rut` VARCHAR(20) NULL AFTER factura_numero');
    await ensureColumn('registros', 'factura_giro', '`factura_giro` VARCHAR(255) NULL AFTER factura_rut');
    await ensureColumn('registros', 'factura_direccion', '`factura_direccion` VARCHAR(255) NULL AFTER factura_giro');
    await ensureColumn('registros', 'factura_comuna', '`factura_comuna` VARCHAR(120) NULL AFTER factura_direccion');
    await ensureColumn('registros', 'factura_ciudad', '`factura_ciudad` VARCHAR(120) NULL AFTER factura_comuna');
    await ensureColumn('registros', 'factura_contacto', '`factura_contacto` VARCHAR(255) NULL AFTER factura_ciudad');
    await ensureColumn('registros', 'factura_observacion', '`factura_observacion` TEXT NULL AFTER factura_contacto');
    await ensureColumn('registros', 'factura_monto', '`factura_monto` DECIMAL(14,2) NULL AFTER factura_observacion');
}

async function ensureHistoryEditColumns() {
    await ensureColumn('registro_historial', 'editado', '`editado` TINYINT(1) NOT NULL DEFAULT 0 AFTER sucursal');
    await ensureColumn('registro_historial', 'editado_por', '`editado_por` VARCHAR(120) NULL AFTER editado');
    await ensureColumn('registro_historial', 'editado_por_sucursal', '`editado_por_sucursal` VARCHAR(120) NULL AFTER editado_por');
    await ensureColumn('registro_historial', 'editado_at', '`editado_at` DATETIME NULL AFTER editado_por_sucursal');
}

async function ensureUsersRoleColumn() {
    await ensureColumn('usuarios', 'role', "`role` VARCHAR(20) NOT NULL DEFAULT 'OPERADOR' AFTER sucursal");
    await pool.query("UPDATE usuarios SET role = UPPER(role)");
    await pool.query(
        "UPDATE usuarios SET role = 'OPERADOR' WHERE role IS NULL OR role = '' OR role NOT IN ('SUPER', 'ADMIN', 'SECRETARIA', 'OPERADOR', 'SUPERVISOR')"
    );
}

async function ensureUsersMustChangePasswordColumn() {
    await ensureColumn(
        'usuarios',
        'must_change_password',
        '`must_change_password` TINYINT(1) NOT NULL DEFAULT 0 AFTER password_hash'
    );
    await pool.query('UPDATE usuarios SET must_change_password = 0 WHERE must_change_password IS NULL');
}

async function ensureCorreoTemplateColumns() {
    await ensureColumn(
        'correo_configuracion',
        'cotizacion_html_template',
        '`cotizacion_html_template` LONGTEXT NULL AFTER smtp_from_name'
    );
    await ensureColumn(
        'correo_configuracion',
        'factura_single_html_template',
        '`factura_single_html_template` LONGTEXT NULL AFTER cotizacion_html_template'
    );
    await ensureColumn(
        'correo_configuracion',
        'factura_pending_html_template',
        '`factura_pending_html_template` LONGTEXT NULL AFTER factura_single_html_template'
    );
}

async function createBranchesTable() {
    await pool.query(`
        CREATE TABLE IF NOT EXISTS sucursales (
            id INT UNSIGNED NOT NULL AUTO_INCREMENT,
            nombre VARCHAR(120) NOT NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            UNIQUE KEY uq_sucursales_nombre (nombre)
        )
    `);
}

async function seedBranchesFromUsers() {
    const branchNames = new Set();

    DEFAULT_USERS.forEach((user) => {
        const branch = normalizeText(user.sucursal);
        if (branch) {
            branchNames.add(branch);
        }
    });

    const [userRows] = await pool.execute(
        `
            SELECT DISTINCT sucursal
            FROM usuarios
            WHERE sucursal IS NOT NULL AND TRIM(sucursal) <> ''
        `
    );
    userRows.forEach((row) => {
        const branch = normalizeText(row.sucursal);
        if (branch) {
            branchNames.add(branch);
        }
    });

    for (const branch of branchNames) {
        await pool.execute('INSERT INTO sucursales (nombre) VALUES (?) ON DUPLICATE KEY UPDATE nombre = nombre', [branch]);
    }
}

async function branchExists(branchName) {
    const cleanBranch = normalizeText(branchName);
    if (!cleanBranch) {
        return false;
    }

    const [rows] = await pool.execute('SELECT id FROM sucursales WHERE nombre = ? LIMIT 1', [cleanBranch]);
    return rows.length > 0;
}

async function createAuthTables() {
    await pool.query(`
        CREATE TABLE IF NOT EXISTS usuarios (
            id INT UNSIGNED NOT NULL AUTO_INCREMENT,
            username VARCHAR(80) NOT NULL,
            nombre VARCHAR(120) NOT NULL,
            sucursal VARCHAR(120) NOT NULL,
            role VARCHAR(20) NOT NULL DEFAULT 'OPERADOR',
            password_salt CHAR(32) NOT NULL,
            password_hash CHAR(128) NOT NULL,
            must_change_password TINYINT(1) NOT NULL DEFAULT 0,
            activo TINYINT(1) NOT NULL DEFAULT 1,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            UNIQUE KEY uq_usuarios_username (username)
        )
    `);

    await pool.query(`
        CREATE TABLE IF NOT EXISTS sesiones (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            user_id INT UNSIGNED NOT NULL,
            token_hash CHAR(64) NOT NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            last_used_at DATETIME NULL DEFAULT NULL,
            expires_at DATETIME NOT NULL,
            revoked_at DATETIME NULL DEFAULT NULL,
            PRIMARY KEY (id),
            UNIQUE KEY uq_sesiones_token_hash (token_hash),
            KEY idx_sesiones_user (user_id),
            KEY idx_sesiones_expires (expires_at),
            CONSTRAINT fk_sesiones_user_id FOREIGN KEY (user_id) REFERENCES usuarios(id)
                ON DELETE CASCADE
        )
    `);
}

async function createHistoryTable() {
    await pool.query(`
        CREATE TABLE IF NOT EXISTS registro_historial (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            registro_id INT UNSIGNED NOT NULL,
            num_ingreso VARCHAR(20) NOT NULL,
            accion VARCHAR(40) NOT NULL,
            comentario TEXT NOT NULL,
            usuario_nombre VARCHAR(120) NOT NULL,
            sucursal VARCHAR(120) NULL,
            editado TINYINT(1) NOT NULL DEFAULT 0,
            editado_por VARCHAR(120) NULL,
            editado_por_sucursal VARCHAR(120) NULL,
            editado_at DATETIME NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            KEY idx_historial_registro_fecha (registro_id, created_at),
            KEY idx_historial_num_ingreso (num_ingreso),
            CONSTRAINT fk_historial_registro_id FOREIGN KEY (registro_id) REFERENCES registros(id)
                ON DELETE CASCADE
        )
    `);
}

async function createFacturaSolicitudesTable() {
    await pool.query(`
        CREATE TABLE IF NOT EXISTS factura_solicitudes (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            num_ingreso VARCHAR(20) NULL,
            nombre_razon_social VARCHAR(255) NOT NULL,
            rut VARCHAR(20) NOT NULL,
            giro VARCHAR(255) NULL,
            direccion VARCHAR(255) NULL,
            comuna VARCHAR(120) NULL,
            ciudad VARCHAR(120) NULL,
            contacto VARCHAR(255) NULL,
            observacion TEXT NULL,
            monto_facturar DECIMAL(14,2) NULL,
            destino_email VARCHAR(255) NULL,
            estado VARCHAR(20) NOT NULL DEFAULT 'PENDIENTE',
            created_by VARCHAR(120) NULL,
            created_by_sucursal VARCHAR(120) NULL,
            updated_by VARCHAR(120) NULL,
            updated_by_sucursal VARCHAR(120) NULL,
            sent_at DATETIME NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            KEY idx_factura_solicitudes_estado (estado),
            KEY idx_factura_solicitudes_created_at (created_at),
            KEY idx_factura_solicitudes_num_ingreso (num_ingreso),
            KEY idx_factura_solicitudes_created_sucursal (created_by_sucursal)
        )
    `);
}

async function createRegistroDocumentAlertsTable() {
    await pool.query(`
        CREATE TABLE IF NOT EXISTS registro_documentos_alertas (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            registro_id INT UNSIGNED NOT NULL,
            num_ingreso VARCHAR(20) NOT NULL,
            sucursal VARCHAR(120) NULL,
            documentos_agregados JSON NOT NULL,
            created_by VARCHAR(120) NULL,
            revisado TINYINT(1) NOT NULL DEFAULT 0,
            revisado_por VARCHAR(120) NULL,
            revisado_por_sucursal VARCHAR(120) NULL,
            revisado_at DATETIME NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            KEY idx_doc_alertas_revisado_fecha (revisado, created_at),
            KEY idx_doc_alertas_sucursal (sucursal),
            KEY idx_doc_alertas_ingreso (num_ingreso),
            CONSTRAINT fk_doc_alertas_registro_id FOREIGN KEY (registro_id) REFERENCES registros(id)
                ON DELETE CASCADE
        )
    `);
}

async function createUtmMensualTable() {
    await pool.query(`
        CREATE TABLE IF NOT EXISTS utm_mensual (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            anio SMALLINT UNSIGNED NOT NULL,
            mes TINYINT UNSIGNED NOT NULL,
            valor DECIMAL(14,2) NOT NULL,
            registrado_por_id INT UNSIGNED NULL,
            registrado_por_nombre VARCHAR(120) NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            UNIQUE KEY uq_utm_mensual_periodo (anio, mes),
            KEY idx_utm_mensual_registrado_por_id (registrado_por_id)
        )
    `);
}

async function createCotizacionServiciosTable() {
    await pool.query(`
        CREATE TABLE IF NOT EXISTS cotizacion_servicios_estandar (
            id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
            nombre_servicio VARCHAR(255) NOT NULL,
            valor_utm DECIMAL(14,2) NOT NULL,
            orden_visual SMALLINT UNSIGNED NOT NULL DEFAULT 1,
            activo TINYINT(1) NOT NULL DEFAULT 1,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            KEY idx_cotizacion_servicios_activo_orden (activo, orden_visual)
        )
    `);
}

async function createCorreoConfiguracionTable() {
    await pool.query(`
        CREATE TABLE IF NOT EXISTS correo_configuracion (
            id TINYINT UNSIGNED NOT NULL,
            smtp_host VARCHAR(255) NOT NULL DEFAULT '',
            smtp_port SMALLINT UNSIGNED NOT NULL DEFAULT 587,
            smtp_secure TINYINT(1) NOT NULL DEFAULT 0,
            smtp_user VARCHAR(255) NULL,
            smtp_password VARCHAR(255) NULL,
            smtp_from_email VARCHAR(255) NOT NULL DEFAULT '',
            smtp_from_name VARCHAR(120) NULL,
            cotizacion_html_template LONGTEXT NULL,
            factura_single_html_template LONGTEXT NULL,
            factura_pending_html_template LONGTEXT NULL,
            updated_by_id INT UNSIGNED NULL,
            updated_by_nombre VARCHAR(120) NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            KEY idx_correo_config_updated_by (updated_by_id)
        )
    `);
}

async function seedCotizacionServiciosDefaults() {
    const [rows] = await pool.execute('SELECT COUNT(*) AS total FROM cotizacion_servicios_estandar');
    const total = Number(rows[0]?.total || 0);
    if (total > 0) {
        return;
    }

    let order = 1;
    for (const item of DEFAULT_COTIZACION_SERVICIOS) {
        const nombreServicio = normalizeText(item?.nombreServicio);
        const valorUtm = Number(item?.valorUtm);
        if (!nombreServicio || !Number.isFinite(valorUtm) || valorUtm <= 0) {
            continue;
        }

        await pool.execute(
            `
                INSERT INTO cotizacion_servicios_estandar (
                    nombre_servicio,
                    valor_utm,
                    orden_visual,
                    activo
                )
                VALUES (?, ?, ?, 1)
            `,
            [nombreServicio, Number(valorUtm.toFixed(2)), order]
        );
        order += 1;
    }
}

async function ensureFacturaSolicitudesUniqueIngresoIndex() {
    const [indexes] = await pool.query("SHOW INDEX FROM factura_solicitudes WHERE Key_name = 'uq_factura_solicitudes_num_ingreso'");
    if (indexes.length > 0) {
        return;
    }

    await pool.execute("UPDATE factura_solicitudes SET num_ingreso = NULL WHERE num_ingreso IS NULL OR TRIM(num_ingreso) = ''");

    const [duplicateIngresoRows] = await pool.execute(
        `
            SELECT num_ingreso
            FROM factura_solicitudes
            WHERE num_ingreso IS NOT NULL AND TRIM(num_ingreso) <> ''
            GROUP BY num_ingreso
            HAVING COUNT(*) > 1
        `
    );

    for (const duplicateRow of duplicateIngresoRows) {
        const numIngreso = normalizeText(duplicateRow.num_ingreso);
        if (!numIngreso) {
            continue;
        }

        const [rows] = await pool.execute(
            `
                SELECT id
                FROM factura_solicitudes
                WHERE num_ingreso = ?
                ORDER BY
                    (estado = ?) DESC,
                    COALESCE(sent_at, updated_at, created_at) DESC,
                    id DESC
            `,
            [numIngreso, FACTURA_SOLICITUD_ESTADO_ENVIADA]
        );

        if (rows.length <= 1) {
            continue;
        }

        const idsToDelete = rows.slice(1).map((row) => Number(row.id)).filter((id) => Number.isInteger(id) && id > 0);
        if (!idsToDelete.length) {
            continue;
        }

        const placeholders = idsToDelete.map(() => '?').join(', ');
        await pool.execute(`DELETE FROM factura_solicitudes WHERE id IN (${placeholders})`, idsToDelete);
    }

    await pool.query('ALTER TABLE factura_solicitudes ADD UNIQUE KEY uq_factura_solicitudes_num_ingreso (num_ingreso)');
}

async function seedDefaultUsers() {
    const [rows] = await pool.execute('SELECT COUNT(*) AS total FROM usuarios');
    const total = Number(rows[0].total || 0);

    if (total > 0) {
        return;
    }

    const initialUsers = resolveInitialUsers();

    for (const user of initialUsers) {
        const passwordData = createPasswordHash(user.password);
        await pool.execute(
            `
                INSERT INTO usuarios (username, nombre, sucursal, role, password_salt, password_hash, must_change_password)
                VALUES (?, ?, ?, ?, ?, ?, 0)
            `,
            [user.username, user.nombre, user.sucursal, user.role, passwordData.salt, passwordData.hash]
        );
    }

    console.warn(
        `Se crearon ${initialUsers.length} usuarios iniciales (${initialUsers
            .map((item) => `${item.username}@${item.sucursal}`)
            .join(', ')}).`
    );
    console.warn('Verifica y cambia las claves iniciales lo antes posible.');
}

function resolveSuperUserBootstrapPassword() {
    const normalizedPassword = String(SUPERUSER_PASSWORD || '');
    if (normalizedPassword.length >= 6 && !exceedsMaxLength(normalizedPassword, MAX_PASSWORD_LENGTH)) {
        return {
            password: normalizedPassword,
            generated: false
        };
    }

    return {
        password: generateSecurePassword(20),
        generated: true
    };
}

async function ensureSuperUserAccount() {
    if (!SUPERUSER_USERNAME) {
        throw new Error('SUPERUSER_USERNAME no puede estar vacio.');
    }

    if (!/^[a-z0-9._-]{3,40}$/.test(SUPERUSER_USERNAME)) {
        throw new Error('SUPERUSER_USERNAME invalido. Usa 3-40 caracteres (a-z, 0-9, ., _, -).');
    }

    if (isGuestUsername(SUPERUSER_USERNAME)) {
        throw new Error('SUPERUSER_USERNAME no puede ser el mismo usuario invitado.');
    }

    if (exceedsMaxLength(SUPERUSER_NAME, 120) || exceedsMaxLength(SUPERUSER_BRANCH, 120)) {
        throw new Error('SUPERUSER_NAME y SUPERUSER_BRANCH no pueden exceder 120 caracteres.');
    }

    let superAccountId = null;
    const [existingRows] = await pool.execute('SELECT id, role FROM usuarios WHERE username = ? LIMIT 1', [SUPERUSER_USERNAME]);
    if (existingRows.length > 0) {
        await pool.execute(
            `
                UPDATE usuarios
                SET nombre = ?, sucursal = ?, role = ?, must_change_password = 0, activo = 1, updated_at = NOW()
                WHERE id = ?
            `,
            [SUPERUSER_NAME, SUPERUSER_BRANCH, ROLE_SUPER, existingRows[0].id]
        );
        superAccountId = Number(existingRows[0].id);

        if (normalizeUserRole(existingRows[0].role, ROLE_OPERADOR) !== ROLE_SUPER) {
            console.warn(`Usuario ${SUPERUSER_USERNAME} promovido a SUPER automaticamente.`);
        }
    } else {
        const superPasswordData = resolveSuperUserBootstrapPassword();
        const passwordData = createPasswordHash(superPasswordData.password);
        const [insertResult] = await pool.execute(
            `
                INSERT INTO usuarios (username, nombre, sucursal, role, password_salt, password_hash, must_change_password, activo)
                VALUES (?, ?, ?, ?, ?, ?, 0, 1)
            `,
            [SUPERUSER_USERNAME, SUPERUSER_NAME, SUPERUSER_BRANCH, ROLE_SUPER, passwordData.salt, passwordData.hash]
        );
        superAccountId = Number(insertResult.insertId);

        if (superPasswordData.generated) {
            console.warn(
                `Se creo usuario SUPER por seguridad: ${SUPERUSER_USERNAME}. Clave temporal generada: ${superPasswordData.password}`
            );
            console.warn(
                'Define SUPERUSER_PASSWORD en .env para un valor permanente y cambialo inmediatamente en produccion.'
            );
        } else {
            console.warn(`Se creo usuario SUPER: ${SUPERUSER_USERNAME}.`);
        }
    }

    if (Number.isInteger(superAccountId) && superAccountId > 0) {
        const [demoteResult] = await pool.execute(
            `
                UPDATE usuarios
                SET role = ?, must_change_password = 0, updated_at = NOW()
                WHERE UPPER(role) = 'SUPER' AND id <> ?
            `,
            [ROLE_ADMIN, superAccountId]
        );

        const demotedCount = Number(demoteResult.affectedRows || 0);
        if (demotedCount > 0) {
            console.warn(
                `Se detectaron ${demotedCount} cuentas SUPER adicionales y se degradaron a ADMIN para mantener una sola cuenta SUPER.`
            );
        }
    }
}

async function countActivePrivilegedUsers() {
    const [rows] = await pool.execute(
        "SELECT COUNT(*) AS totalPrivileged FROM usuarios WHERE UPPER(role) IN ('ADMIN', 'SUPER') AND activo = 1"
    );
    return Number(rows[0]?.totalPrivileged || 0);
}

async function countSuperUsers(excludeUserId = null) {
    const hasExclusion = Number.isInteger(excludeUserId) && excludeUserId > 0;
    const [rows] = await pool.execute(
        `
            SELECT COUNT(*) AS totalSuper
            FROM usuarios
            WHERE UPPER(role) = 'SUPER'
              ${hasExclusion ? 'AND id <> ?' : ''}
        `,
        hasExclusion ? [excludeUserId] : []
    );

    return Number(rows[0]?.totalSuper || 0);
}

async function ensureAtLeastOnePrivilegedAccount() {
    if ((await countActivePrivilegedUsers()) > 0) {
        return;
    }

    const preferredPrivileged = resolveInitialUsers().find((item) => isPrivilegedRole(item.role));
    if (!preferredPrivileged) {
        throw new Error('No existe usuario ADMIN/SUPER de referencia. Define uno en DEFAULT_USERS o SUPERUSER_*.');
    }

    const preferredRole = isPrivilegedRole(preferredPrivileged.role) ? preferredPrivileged.role : ROLE_ADMIN;
    const [existingRows] = await pool.execute('SELECT id FROM usuarios WHERE username = ? LIMIT 1', [preferredPrivileged.username]);
    if (existingRows.length > 0) {
        await pool.execute('UPDATE usuarios SET role = ?, activo = 1 WHERE id = ?', [preferredRole, existingRows[0].id]);
        console.warn(`Usuario ${preferredPrivileged.username} promovido a ${preferredRole} automaticamente.`);
        return;
    }

    const passwordData = createPasswordHash(preferredPrivileged.password);
    const mustChangePassword = shouldForcePasswordChangeForRole(preferredRole) ? 1 : 0;
    await pool.execute(
        `
            INSERT INTO usuarios (username, nombre, sucursal, role, password_salt, password_hash, must_change_password, activo)
            VALUES (?, ?, ?, ?, ?, ?, ?, 1)
        `,
        [
            preferredPrivileged.username,
            preferredPrivileged.nombre,
            preferredPrivileged.sucursal,
            preferredRole,
            passwordData.salt,
            passwordData.hash,
            mustChangePassword
        ]
    );

    console.warn(`Se creo usuario ${preferredRole} por seguridad: ${preferredPrivileged.username}. Cambia su contrasena.`);
}

async function ensureGuestUserAccount() {
    const generatedPassword = createPasswordHash(generateSecurePassword(24));
    await pool.execute(
        `
            INSERT INTO usuarios (username, nombre, sucursal, role, password_salt, password_hash, must_change_password, activo)
            VALUES (?, ?, ?, ?, ?, ?, 0, 1)
            ON DUPLICATE KEY UPDATE
                nombre = VALUES(nombre),
                sucursal = VALUES(sucursal),
                role = VALUES(role),
                must_change_password = 0
        `,
        [
            GUEST_USERNAME,
            GUEST_DISPLAY_NAME,
            GUEST_BRANCH_NAME,
            ROLE_OPERADOR,
            generatedPassword.salt,
            generatedPassword.hash
        ]
    );

    const [rows] = await pool.execute(
        `
            SELECT
                id AS user_id,
                username,
                nombre,
                sucursal,
                role,
                must_change_password,
                activo
                FROM usuarios
                WHERE LOWER(username) = ?
                LIMIT 1
            `,
            [GUEST_USERNAME]
        );

    if (!rows.length) {
        throw new Error('No fue posible inicializar el usuario invitado.');
    }

    return rows[0];
}

async function reserveNextIngreso(connection, anio) {
    const [rows] = await connection.execute(
        'SELECT COALESCE(MAX(correlativo), 0) AS maxCorrelativo FROM registros WHERE anio = ? FOR UPDATE',
        [anio]
    );

    const nextCorrelativo = Number(rows[0].maxCorrelativo || 0) + 1;
    return {
        anio,
        correlativo: nextCorrelativo,
        numIngreso: `${nextCorrelativo}-${anio}`
    };
}

async function insertHistoryEntry(executor, historyPayload) {
    await executor.execute(
        `
            INSERT INTO registro_historial (
                registro_id,
                num_ingreso,
                accion,
                comentario,
                usuario_nombre,
                sucursal
            ) VALUES (?, ?, ?, ?, ?, ?)
        `,
        [
            historyPayload.registroId,
            historyPayload.numIngreso,
            historyPayload.accion,
            historyPayload.comentario,
            historyPayload.usuarioNombre,
            historyPayload.sucursal || null
        ]
    );
}

async function fetchHistoryByRegistroId(registroId) {
    const [rows] = await pool.execute(
        `
            SELECT
                id,
                num_ingreso,
                accion,
                created_at,
                comentario,
                usuario_nombre,
                sucursal,
                editado,
                editado_por,
                editado_por_sucursal,
                editado_at
            FROM registro_historial
            WHERE registro_id = ?
            ORDER BY created_at DESC,
                     CASE WHEN UPPER(accion) = 'COMENTARIO' THEN 0 ELSE 1 END ASC,
                     id DESC
        `,
        [registroId]
    );

    return rows.map(toHistoryRow);
}

async function createSystemConfigTable() {
    await pool.query(`
        CREATE TABLE IF NOT EXISTS sistema_configuracion (
            id TINYINT UNSIGNED NOT NULL,
            maintenance_mode TINYINT(1) NOT NULL DEFAULT 0,
            login_notice_enabled TINYINT(1) NOT NULL DEFAULT 0,
            login_notice_url VARCHAR(255) NOT NULL DEFAULT '/geo_rural/registro/login/registro.html',
            login_notice_message VARCHAR(600) NULL,
            login_notice_starts_at DATETIME NULL,
            login_notice_expires_at DATETIME NULL,
            login_notice_updated_by_id INT UNSIGNED NULL,
            login_notice_updated_by_name VARCHAR(120) NULL,
            login_notice_updated_at DATETIME NULL,
            updated_by_id INT UNSIGNED NULL,
            updated_by_name VARCHAR(120) NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            PRIMARY KEY (id)
        )
    `);
}

async function ensureSystemConfigColumns() {
    await ensureColumn('sistema_configuracion', 'login_notice_enabled', '`login_notice_enabled` TINYINT(1) NOT NULL DEFAULT 0 AFTER maintenance_mode');
    await ensureColumn(
        'sistema_configuracion',
        'login_notice_url',
        "`login_notice_url` VARCHAR(255) NOT NULL DEFAULT '/geo_rural/registro/login/registro.html' AFTER login_notice_enabled"
    );
    await ensureColumn(
        'sistema_configuracion',
        'login_notice_message',
        '`login_notice_message` VARCHAR(600) NULL AFTER login_notice_url'
    );
    await ensureColumn(
        'sistema_configuracion',
        'login_notice_starts_at',
        '`login_notice_starts_at` DATETIME NULL AFTER login_notice_message'
    );
    await ensureColumn(
        'sistema_configuracion',
        'login_notice_expires_at',
        '`login_notice_expires_at` DATETIME NULL AFTER login_notice_starts_at'
    );
    await ensureColumn(
        'sistema_configuracion',
        'login_notice_updated_by_id',
        '`login_notice_updated_by_id` INT UNSIGNED NULL AFTER login_notice_expires_at'
    );
    await ensureColumn(
        'sistema_configuracion',
        'login_notice_updated_by_name',
        '`login_notice_updated_by_name` VARCHAR(120) NULL AFTER login_notice_updated_by_id'
    );
    await ensureColumn(
        'sistema_configuracion',
        'login_notice_updated_at',
        '`login_notice_updated_at` DATETIME NULL AFTER login_notice_updated_by_name'
    );
}

async function ensureSystemConfigRow() {
    const defaultUrl = normalizeLoginRouteNoticeUrl(LOGIN_ROUTE_NOTICE_DEFAULT_URL);
    const defaultMessage = normalizeText(LOGIN_ROUTE_NOTICE_DEFAULT_MESSAGE);
    await pool.execute(
        `
            INSERT INTO sistema_configuracion (id, maintenance_mode, login_notice_url, login_notice_message)
            VALUES (?, 0, ?, ?)
            ON DUPLICATE KEY UPDATE
                login_notice_url = IF(
                    login_notice_url IS NULL OR login_notice_url = '' OR login_notice_url = '/registro/login/registro.html',
                    VALUES(login_notice_url),
                    login_notice_url
                ),
                login_notice_message = IF(
                    login_notice_message IS NULL OR login_notice_message = '',
                    VALUES(login_notice_message),
                    login_notice_message
                )
        `,
        [SYSTEM_CONFIG_SINGLETON_ID, defaultUrl, defaultMessage]
    );

    if (IS_PRODUCTION_ENV) {
        await pool.execute(
            `
                UPDATE sistema_configuracion
                SET login_notice_url = ?
                WHERE id = ?
                  AND (
                        login_notice_url IS NULL
                     OR TRIM(login_notice_url) = ''
                     OR LOWER(TRIM(login_notice_url)) = '/geo_rural/registro/login/registro.html'
                     OR LOWER(TRIM(login_notice_url)) = '/registro/login/registro.html'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'localhost/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE '127.0.0.1/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'http://localhost/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'https://localhost/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'http://127.0.0.1/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'https://127.0.0.1/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'mi-registro.cl/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'http://mi-registro.cl/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'https://mi-registro.cl/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'my-registro.cl/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'http://my-registro.cl/%'
                     OR LOWER(TRIM(login_notice_url)) LIKE 'https://my-registro.cl/%'
                  )
            `,
            [defaultUrl, SYSTEM_CONFIG_SINGLETON_ID]
        );
    }
}

async function loadMaintenanceModeCache() {
    const [rows] = await pool.execute(
        `
            SELECT
                maintenance_mode,
                updated_by_name,
                updated_at,
                login_notice_enabled,
                login_notice_url,
                login_notice_message,
                login_notice_starts_at,
                login_notice_expires_at,
                login_notice_updated_by_name,
                login_notice_updated_at
            FROM sistema_configuracion
            WHERE id = ?
            LIMIT 1
        `,
        [SYSTEM_CONFIG_SINGLETON_ID]
    );

    if (!rows.length) {
        maintenanceModeCache = {
            enabled: false,
            updatedAt: '',
            updatedBy: ''
        };
        loginRouteNoticeCache = {
            enabled: false,
            url: normalizeLoginRouteNoticeUrl(LOGIN_ROUTE_NOTICE_DEFAULT_URL),
            message: normalizeText(LOGIN_ROUTE_NOTICE_DEFAULT_MESSAGE),
            startsAt: '',
            expiresAt: '',
            updatedAt: '',
            updatedBy: ''
        };
        return;
    }

    setMaintenanceModeCacheFromRow(rows[0]);
    setLoginRouteNoticeCacheFromRow(rows[0]);

    if (shouldAutoDisableLoginRouteNotice()) {
        await pool.execute(
            `
                UPDATE sistema_configuracion
                SET login_notice_enabled = 0,
                    login_notice_updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                  AND login_notice_enabled = 1
            `,
            [SYSTEM_CONFIG_SINGLETON_ID]
        );
        loginRouteNoticeCache.enabled = false;
        loginRouteNoticeCache.updatedAt = new Date().toISOString();
    }
}

async function updateMaintenanceMode(enabled, actorUser = null) {
    const enabledFlag = enabled ? 1 : 0;
    const actorIdValue = Number(actorUser?.id);
    const updatedById = Number.isInteger(actorIdValue) && actorIdValue > 0 ? actorIdValue : null;
    const updatedByName = normalizeText(actorUser?.nombre || actorUser?.username || '') || null;

    await pool.execute(
        `
            UPDATE sistema_configuracion
            SET maintenance_mode = ?, updated_by_id = ?, updated_by_name = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        `,
        [enabledFlag, updatedById, updatedByName, SYSTEM_CONFIG_SINGLETON_ID]
    );

    await loadMaintenanceModeCache();
    return getMaintenanceModeStatusPayload();
}

async function updateLoginRouteNotice(enabled, actorUser = null, requestedUrl = '') {
    const actorIdValue = Number(actorUser?.id);
    const updatedById = Number.isInteger(actorIdValue) && actorIdValue > 0 ? actorIdValue : null;
    const updatedByName = normalizeText(actorUser?.nombre || actorUser?.username || '') || null;
    const safeUrl = normalizeLoginRouteNoticeUrl(
        requestedUrl || loginRouteNoticeCache.url || LOGIN_ROUTE_NOTICE_DEFAULT_URL
    );
    const safeMessage = normalizeText(loginRouteNoticeCache.message || LOGIN_ROUTE_NOTICE_DEFAULT_MESSAGE);

    if (exceedsMaxLength(safeUrl, LOGIN_ROUTE_NOTICE_MAX_URL_LENGTH)) {
        throw new Error(`La URL del aviso excede ${LOGIN_ROUTE_NOTICE_MAX_URL_LENGTH} caracteres.`);
    }

    if (enabled) {
        await pool.execute(
            `
                UPDATE sistema_configuracion
                SET login_notice_enabled = 1,
                    login_notice_url = ?,
                    login_notice_message = ?,
                    login_notice_starts_at = CURRENT_TIMESTAMP,
                    login_notice_expires_at = DATE_ADD(CURRENT_TIMESTAMP, INTERVAL ? DAY),
                    login_notice_updated_by_id = ?,
                    login_notice_updated_by_name = ?,
                    login_notice_updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            `,
            [safeUrl, safeMessage, LOGIN_ROUTE_NOTICE_DURATION_DAYS, updatedById, updatedByName, SYSTEM_CONFIG_SINGLETON_ID]
        );
    } else {
        await pool.execute(
            `
                UPDATE sistema_configuracion
                SET login_notice_enabled = 0,
                    login_notice_url = ?,
                    login_notice_message = ?,
                    login_notice_updated_by_id = ?,
                    login_notice_updated_by_name = ?,
                    login_notice_updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            `,
            [safeUrl, safeMessage, updatedById, updatedByName, SYSTEM_CONFIG_SINGLETON_ID]
        );
    }

    await loadMaintenanceModeCache();
    return getLoginRouteNoticeAdminPayload();
}

async function updateLoginRouteNoticeUrl(actorUser = null, requestedUrl = '') {
    const requestedValue = normalizeText(requestedUrl);
    if (!requestedValue) {
        throw new Error('Debes ingresar la URL de redireccion.');
    }

    const actorIdValue = Number(actorUser?.id);
    const updatedById = Number.isInteger(actorIdValue) && actorIdValue > 0 ? actorIdValue : null;
    const updatedByName = normalizeText(actorUser?.nombre || actorUser?.username || '') || null;
    const safeUrl = normalizeLoginRouteNoticeUrl(requestedValue);
    const safeMessage = normalizeText(loginRouteNoticeCache.message || LOGIN_ROUTE_NOTICE_DEFAULT_MESSAGE);

    if (exceedsMaxLength(safeUrl, LOGIN_ROUTE_NOTICE_MAX_URL_LENGTH)) {
        throw new Error(`La URL del aviso excede ${LOGIN_ROUTE_NOTICE_MAX_URL_LENGTH} caracteres.`);
    }

    await pool.execute(
        `
            UPDATE sistema_configuracion
            SET login_notice_url = ?,
                login_notice_message = ?,
                login_notice_updated_by_id = ?,
                login_notice_updated_by_name = ?,
                login_notice_updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        `,
        [safeUrl, safeMessage, updatedById, updatedByName, SYSTEM_CONFIG_SINGLETON_ID]
    );

    await loadMaintenanceModeCache();
    return getLoginRouteNoticeAdminPayload();
}

async function initDatabase() {
    const sanitizedDbName = cleanDbIdentifier(DB_NAME) || 'geo_rural';
    const bootstrapConnection = await mysql.createConnection({
        host: DB_HOST,
        port: DB_PORT,
        user: DB_USER,
        password: DB_PASSWORD
    });

    try {
        await bootstrapConnection.query(
            `CREATE DATABASE IF NOT EXISTS \`${sanitizedDbName}\` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci`
        );
    } catch (error) {
        const missingCreateDatabasePermission =
            error &&
            (error.code === 'ER_DBACCESS_DENIED_ERROR' || error.code === 'ER_SPECIFIC_ACCESS_DENIED_ERROR');

        if (!missingCreateDatabasePermission) {
            throw error;
        }

        console.warn(
            `No se pudo crear automaticamente la base "${sanitizedDbName}" por permisos insuficientes. Se intentara continuar con una base existente.`
        );
    } finally {
        await bootstrapConnection.end();
    }

    pool = mysql.createPool({
        host: DB_HOST,
        port: DB_PORT,
        user: DB_USER,
        password: DB_PASSWORD,
        database: sanitizedDbName,
        waitForConnections: true,
        connectionLimit: 10,
        queueLimit: 0
    });

    await pool.query(`
        CREATE TABLE IF NOT EXISTS registros (
            id INT UNSIGNED NOT NULL AUTO_INCREMENT,
            num_ingreso VARCHAR(20) NOT NULL,
            correlativo INT UNSIGNED NOT NULL,
            anio SMALLINT UNSIGNED NOT NULL,
            nombre VARCHAR(255) NOT NULL,
            rut VARCHAR(20) NOT NULL,
            telefono VARCHAR(40) NULL,
            correo VARCHAR(255) NULL,
            region VARCHAR(120) NOT NULL,
            comuna VARCHAR(120) NOT NULL,
            nombre_predio VARCHAR(255) NULL,
            rol VARCHAR(120) NOT NULL,
            num_lotes INT NULL,
            estado VARCHAR(40) NULL,
            comentario TEXT NULL,
            factura_nombre_razon VARCHAR(255) NULL,
            factura_numero VARCHAR(80) NULL,
            factura_rut VARCHAR(20) NULL,
            factura_giro VARCHAR(255) NULL,
            factura_direccion VARCHAR(255) NULL,
            factura_comuna VARCHAR(120) NULL,
            factura_ciudad VARCHAR(120) NULL,
            factura_contacto VARCHAR(255) NULL,
            factura_observacion TEXT NULL,
            factura_monto DECIMAL(14,2) NULL,
            created_by VARCHAR(120) NULL,
            created_by_sucursal VARCHAR(120) NULL,
            updated_by VARCHAR(120) NULL,
            updated_by_sucursal VARCHAR(120) NULL,
            documentos JSON NOT NULL,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            UNIQUE KEY uq_num_ingreso (num_ingreso),
            UNIQUE KEY uq_anio_correlativo (anio, correlativo),
            KEY idx_nombre (nombre),
            KEY idx_rut (rut),
            KEY idx_rol (rol)
        )
    `);

    try {
        await ensureUniqueCorrelativoIndex();
        await ensureRegistrosAuditColumns();
        await ensureRegistrosFacturaColumns();
        await createAuthTables();
        await cleanupExpiredSessionsIfNeeded(true);
        await ensureUsersRoleColumn();
        await ensureUsersMustChangePasswordColumn();
        await createBranchesTable();
        await createHistoryTable();
        await ensureHistoryEditColumns();
        await createFacturaSolicitudesTable();
        await createRegistroDocumentAlertsTable();
        await createUtmMensualTable();
        await createCotizacionServiciosTable();
        await createCorreoConfiguracionTable();
        await createSystemConfigTable();
        await ensureSystemConfigColumns();
        await ensureSystemConfigRow();
        await ensureCorreoTemplateColumns();
        await seedCotizacionServiciosDefaults();
        await ensureFacturaSolicitudesUniqueIngresoIndex();
        await seedDefaultUsers();
        await ensureSuperUserAccount();
        await ensureAtLeastOnePrivilegedAccount();
        await ensureGuestUserAccount();
        await seedBranchesFromUsers();
        await loadMaintenanceModeCache();
    } catch (error) {
        console.error('No fue posible terminar la inicializacion de la base:', error.message);
        throw error;
    }
}

app.get('/api/health', async (req, res) => {
    try {
        await pool.query('SELECT 1');
        res.json({ ok: true });
    } catch (error) {
        res.status(500).json({ ok: false, message: 'Error de conexion con base de datos.' });
    }
});

app.use('/api', enforceAllowedOrigin);
app.use(enforceMaintenanceMode);
app.use(enforceFrontendRouteProtection);

app.post('/api/auth/login', async (req, res) => {
    try {
        const username = normalizeText(req.body.username).toLowerCase();
        const password = String(req.body.password || '');

        if (!username || !password) {
            return res.status(400).json({ message: 'Ingresa usuario y contrasena.' });
        }

        if (exceedsMaxLength(username, 80) || exceedsMaxLength(password, MAX_PASSWORD_LENGTH)) {
            return res.status(400).json({ message: 'Credenciales con largo invalido.' });
        }

        const loginRateLimit = getLoginRateLimitState(req, username);
        if (loginRateLimit.blocked) {
            return res.status(429).json({
                message: `Demasiados intentos fallidos. Intenta nuevamente en ${loginRateLimit.retryAfterSeconds} segundos.`,
                code: 'LOGIN_RATE_LIMIT'
            });
        }

        await cleanupExpiredSessionsIfNeeded();

        const [rows] = await pool.execute(
            `
                SELECT
                    id AS user_id,
                    username,
                    nombre,
                    sucursal,
                    role,
                    must_change_password,
                    password_salt,
                    password_hash,
                    activo
                FROM usuarios
                WHERE LOWER(username) = ?
                LIMIT 1
            `,
            [username]
        );

        if (!rows.length || Number(rows[0].activo) !== 1) {
            registerFailedLoginAttempt(req, username);
            return res.status(401).json({ message: 'Usuario o contrasena incorrectos.' });
        }

        if (!verifyPassword(password, rows[0].password_salt, rows[0].password_hash)) {
            registerFailedLoginAttempt(req, username);
            return res.status(401).json({ message: 'Usuario o contrasena incorrectos.' });
        }

        clearFailedLoginAttempts(req, username);

        const token = generateSessionToken();
        await pool.execute(
            `
                INSERT INTO sesiones (user_id, token_hash, expires_at, last_used_at)
                VALUES (?, ?, DATE_ADD(NOW(), INTERVAL ? HOUR), NOW())
            `,
            [rows[0].user_id, hashToken(token), AUTH_SESSION_HOURS]
        );

        setSessionCookie(res, token);
        console.info(`[Auth] Sesion iniciada por ${rows[0].username}. IP=${getClientIp(req)}`);

        const authUser = formatAuthUser(rows[0]);
        const responsePayload = {
            token,
            expiresInHours: AUTH_SESSION_HOURS,
            user: authUser
        };

        if (canReceiveWriteApiKey(authUser)) {
            responsePayload.apiWriteKey = API_WRITE_KEY;
        }

        return res.json(responsePayload);
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible iniciar sesion.' });
    }
});

app.post('/api/auth/guest', async (req, res) => {
    try {
        await cleanupExpiredSessionsIfNeeded();
        const guestUser = await ensureGuestUserAccount();

        if (Number(guestUser.activo) !== 1) {
            clearSessionCookie(res);
            return res.status(403).json({ message: 'La cuenta de invitado esta deshabilitada por el administrador.' });
        }

        const token = generateSessionToken();
        await pool.execute(
            `
                INSERT INTO sesiones (user_id, token_hash, expires_at, last_used_at)
                VALUES (?, ?, DATE_ADD(NOW(), INTERVAL ? HOUR), NOW())
            `,
            [guestUser.user_id, hashToken(token), AUTH_SESSION_HOURS]
        );

        setSessionCookie(res, token);
        console.info(`[Auth] Sesion invitado iniciada por ${guestUser.username}. IP=${getClientIp(req)}`);

        const guestAuthUser = formatAuthUser(guestUser);
        return res.json({
            token,
            expiresInHours: AUTH_SESSION_HOURS,
            user: guestAuthUser
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible iniciar sesion como invitado.' });
    }
});

app.get('/api/auth/me', requireAuth, async (req, res) => {
    const responsePayload = { user: req.authUser };
    if (canReceiveWriteApiKey(req.authUser)) {
        responsePayload.apiWriteKey = API_WRITE_KEY;
    }
    return res.json(responsePayload);
});

app.post('/api/auth/logout', requireAuth, async (req, res) => {
    try {
        await pool.execute('UPDATE sesiones SET revoked_at = NOW() WHERE id = ?', [req.authSessionId]);
        clearSessionCookie(res);
        console.info(
            `[Auth] Sesion cerrada manualmente por ${normalizeText(req.authUser?.username || req.authUser?.nombre || 'desconocido')}. IP=${getClientIp(req)}`
        );
        return res.json({ message: 'Sesion cerrada correctamente.' });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible cerrar la sesion.' });
    }
});

app.post('/api/auth/change-password', requireAuth, async (req, res) => {
    try {
        const currentPassword = String(req.body.currentPassword || '');
        const newPassword = String(req.body.newPassword || '');

        if (!currentPassword || !newPassword) {
            return res.status(400).json({ message: 'Debes ingresar clave actual y nueva clave.' });
        }

        if (newPassword.length < 6) {
            return res.status(400).json({ message: 'La nueva clave debe tener al menos 6 caracteres.' });
        }

        if (exceedsMaxLength(currentPassword, MAX_PASSWORD_LENGTH) || exceedsMaxLength(newPassword, MAX_PASSWORD_LENGTH)) {
            return res.status(400).json({ message: `La clave no puede exceder ${MAX_PASSWORD_LENGTH} caracteres.` });
        }

        const [rows] = await pool.execute(
            `
                SELECT
                    id AS user_id,
                    username,
                    nombre,
                    sucursal,
                    role,
                    must_change_password,
                    password_salt,
                    password_hash,
                    activo
                FROM usuarios
                WHERE id = ?
                LIMIT 1
            `,
            [req.authUser.id]
        );

        if (!rows.length || Number(rows[0].activo) !== 1) {
            return res.status(401).json({ message: 'Usuario no disponible para cambiar clave.' });
        }

        if (!verifyPassword(currentPassword, rows[0].password_salt, rows[0].password_hash)) {
            return res.status(401).json({ message: 'La clave actual es incorrecta.' });
        }

        const newPasswordData = createPasswordHash(newPassword);
        await pool.execute(
            `
                UPDATE usuarios
                SET password_salt = ?, password_hash = ?, must_change_password = 0, updated_at = NOW()
                WHERE id = ?
            `,
            [newPasswordData.salt, newPasswordData.hash, req.authUser.id]
        );

        await pool.execute('UPDATE sesiones SET revoked_at = NOW() WHERE user_id = ? AND id <> ?', [
            req.authUser.id,
            req.authSessionId
        ]);

        const updatedUser = {
            ...rows[0],
            password_salt: newPasswordData.salt,
            password_hash: newPasswordData.hash,
            must_change_password: 0
        };

        return res.json({
            message: 'Clave actualizada correctamente.',
            user: formatAuthUser(updatedUser)
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible actualizar la clave.' });
    }
});

app.get('/api/public/registros/progreso', async (req, res) => {
    try {
        const rolSearch = normalizeRolForSearch(req.query.rol);
        const rol = rolSearch.original;

        if (!rol) {
            return res.status(400).json({ message: 'Ingresa el ROL para buscar el progreso.' });
        }

        if (exceedsMaxLength(rol, 120)) {
            return res.status(400).json({ message: 'El ROL excede el largo maximo permitido.' });
        }

        if (!rolSearch.compact) {
            return res.status(400).json({ message: 'Ingresa un ROL valido.' });
        }

        const [rows] = await pool.execute(
            `
                SELECT
                    id,
                    nombre,
                    correo,
                    nombre_predio,
                    region,
                    comuna,
                    rol
                FROM registros
                WHERE ${buildRolWhereClause('rol')}
                ORDER BY id DESC
                LIMIT 1
            `,
            [rolSearch.canonical, rolSearch.compact]
        );

        if (!rows.length) {
            return res.status(404).json({ message: 'No se encontro un registro con ese ROL.' });
        }

        const targetRow = rows[0];
        const historial = await fetchHistoryByRegistroId(targetRow.id);
        return res.json({
            nombre: normalizeText(targetRow.nombre),
            correo: normalizeText(targetRow.correo).toLowerCase(),
            nombrePredio: normalizeText(targetRow.nombre_predio),
            region: normalizeText(targetRow.region),
            comuna: normalizeText(targetRow.comuna),
            rol: normalizeText(targetRow.rol),
            historial: historial.map(toPublicProgressHistoryRow)
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible cargar el historial de progreso.' });
    }
});

app.get('/api/public/system-status', async (req, res) => {
    try {
        await loadMaintenanceModeCache();
        return res.json({
            maintenance: getMaintenanceModeStatusPayload(),
            loginRouteNotice: getLoginRouteNoticePublicPayload()
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible obtener el estado del sistema.' });
    }
});

app.use('/api', requireAuth);
app.use('/api', requirePasswordChangeCompleted);

app.get('/api/system/mantenimiento', requireSuper, async (req, res) => {
    try {
        await loadMaintenanceModeCache();
        return res.json({
            maintenance: getMaintenanceModeStatusPayload(),
            loginRouteNotice: getLoginRouteNoticeAdminPayload()
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible obtener el estado de mantencion.' });
    }
});

app.put('/api/system/mantenimiento', requireSuper, async (req, res) => {
    try {
        const enabled = parseMaintenanceModeInput(req.body?.enabled);
        if (enabled === null) {
            return res.status(400).json({ message: 'Debes enviar enabled en formato booleano.' });
        }

        const maintenance = await updateMaintenanceMode(enabled, req.authUser);
        console.info(
            `[System] Modo mantencion ${enabled ? 'activado' : 'desactivado'} por ${normalizeText(req.authUser?.username || req.authUser?.nombre || 'super')}.`
        );
        return res.json({
            message: enabled ? 'Modo mantencion activado.' : 'Modo mantencion desactivado.',
            maintenance,
            loginRouteNotice: getLoginRouteNoticeAdminPayload()
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible actualizar el modo de mantencion.' });
    }
});

app.put('/api/system/login-route-notice', requireSuper, async (req, res) => {
    try {
        const enabled = parseMaintenanceModeInput(req.body?.enabled);
        if (enabled === null) {
            return res.status(400).json({ message: 'Debes enviar enabled en formato booleano.' });
        }

        const requestedUrl = normalizeText(req.body?.url || '');
        const loginRouteNotice = await updateLoginRouteNotice(enabled, req.authUser, requestedUrl);
        console.info(
            `[System] Aviso de ruta login ${enabled ? 'activado' : 'desactivado'} por ${normalizeText(req.authUser?.username || req.authUser?.nombre || 'super')}.`
        );
        return res.json({
            message: enabled
                ? `Aviso activado por ${LOGIN_ROUTE_NOTICE_DURATION_DAYS} dias.`
                : 'Aviso de ruta login desactivado.',
            maintenance: getMaintenanceModeStatusPayload(),
            loginRouteNotice
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: error.message || 'No fue posible actualizar el aviso de ruta login.' });
    }
});

app.put('/api/system/login-route-notice/url', requireSuper, async (req, res) => {
    try {
        const requestedUrl = normalizeText(req.body?.url || '');
        if (!requestedUrl) {
            return res.status(400).json({ message: 'Debes ingresar la URL de redireccion.' });
        }

        const loginRouteNotice = await updateLoginRouteNoticeUrl(req.authUser, requestedUrl);
        console.info(
            `[System] URL de redireccion login actualizada por ${normalizeText(req.authUser?.username || req.authUser?.nombre || 'super')}.`
        );
        return res.json({
            message: 'URL de redireccion actualizada.',
            maintenance: getMaintenanceModeStatusPayload(),
            loginRouteNotice
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: error.message || 'No fue posible actualizar la URL de redireccion.' });
    }
});

app.get('/api/mail-templates', async (req, res) => {
    try {
        const storedRow = await getStoredCorreoConfigurationRow();
        return res.json({
            templates: resolveMailTemplatesForUse(storedRow),
            updatedAt: storedRow?.updated_at instanceof Date ? storedRow.updated_at.toISOString() : normalizeText(storedRow?.updated_at || ''),
            updatedBy: normalizeText(storedRow?.updated_by_nombre || '')
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible cargar las plantillas de correo.' });
    }
});

app.get('/api/utm/mes-actual', async (req, res) => {
    try {
        if (!canUseSecretaryFeatures(req.authUser)) {
            return res.status(403).json({ message: 'Solo secretaria o superusuario pueden consultar la UTM mensual.' });
        }

        const status = await getCurrentMonthUtmStatus();
        return res.json(status);
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible consultar la UTM mensual.' });
    }
});

app.post('/api/utm/mes-actual', async (req, res) => {
    try {
        if (!canUseSecretaryFeatures(req.authUser)) {
            return res.status(403).json({ message: 'Solo secretaria o superusuario pueden registrar la UTM mensual.' });
        }

        const parsedValue = parseUtmValue(req.body?.valor);
        const parsedConfirmation = parseUtmValue(req.body?.confirmacion);
        if (!parsedValue.isValid || parsedValue.value === null || !parsedConfirmation.isValid || parsedConfirmation.value === null) {
            return res.status(400).json({ message: 'Ingresa y confirma un valor UTM valido mayor a 0.' });
        }

        if (Number(parsedValue.value.toFixed(2)) !== Number(parsedConfirmation.value.toFixed(2))) {
            return res.status(400).json({ message: 'El monto UTM y su confirmacion no coinciden.' });
        }

        const period = getCurrentCalendarPeriod();
        const userId = Number(req.authUser?.id);
        const registeredById = Number.isInteger(userId) && userId > 0 ? userId : null;
        const registeredByName = normalizeText(req.authUser?.nombre || req.authUser?.username || '');

        await pool.execute(
            `
                INSERT INTO utm_mensual (
                    anio,
                    mes,
                    valor,
                    registrado_por_id,
                    registrado_por_nombre
                )
                VALUES (?, ?, ?, ?, ?)
                ON DUPLICATE KEY UPDATE
                    valor = VALUES(valor),
                    registrado_por_id = VALUES(registrado_por_id),
                    registrado_por_nombre = VALUES(registrado_por_nombre),
                    updated_at = CURRENT_TIMESTAMP
            `,
            [period.anio, period.mes, parsedValue.value, registeredById, registeredByName]
        );

        const status = await getCurrentMonthUtmStatus();
        return res.json({
            message: 'UTM mensual guardada correctamente.',
            status
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible guardar la UTM mensual.' });
    }
});

app.get('/api/cotizaciones/resumen', async (req, res) => {
    try {
        if (!canUseSecretaryFeatures(req.authUser)) {
            return res.status(403).json({ message: 'Solo secretaria o superusuario pueden preparar cotizaciones.' });
        }

        const summaryResult = await buildCurrentCotizacionSummary();
        if (!summaryResult.ok) {
            return res.status(409).json({
                message: 'Debes registrar la UTM mensual antes de preparar una cotizacion.',
                status: summaryResult.status || null
            });
        }

        return res.json(summaryResult.summary);
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible preparar el resumen de cotizacion.' });
    }
});

app.post('/api/cotizaciones/enviar', async (req, res) => {
    try {
        if (!canUseSecretaryFeatures(req.authUser)) {
            return res.status(403).json({ message: 'Solo secretaria o superusuario pueden enviar cotizaciones.' });
        }

        const smtpResolution = await resolveActiveSmtpConfiguration();
        if (!smtpResolution.config) {
            return res.status(503).json({
                message:
                    'Envio no disponible: configura el servicio de correo (SMTP) en el menu del superusuario o variables de entorno.'
            });
        }

        const destinationEmail = normalizeText(req.body?.destinationEmail).toLowerCase();
        const clientName = normalizeText(req.body?.clientName || req.body?.nombreCliente);
        const referenciaCliente = normalizeText(req.body?.referenciaCliente);
        const parcelamientoRangeInput = normalizeText(
            req.body?.parcelamientoRangeKey ?? req.body?.parcelamientoRange ?? req.body?.tramoParcelamiento
        );
        const parcelamientoRangeKey = normalizeCotizacionParcelamientoKey(parcelamientoRangeInput);
        if (!clientName) {
            return res.status(400).json({ message: 'Debes indicar el nombre del cliente.' });
        }
        if (!destinationEmail) {
            return res.status(400).json({ message: 'Debes indicar el correo del cliente.' });
        }
        if (exceedsMaxLength(clientName, 120)) {
            return res.status(400).json({ message: 'El nombre del cliente excede el maximo permitido de 120 caracteres.' });
        }
        if (!isValidEmail(destinationEmail)) {
            return res.status(400).json({ message: 'El correo del cliente no es valido.' });
        }
        if (exceedsMaxLength(destinationEmail, 255)) {
            return res.status(400).json({ message: 'El correo del cliente excede el maximo permitido de 255 caracteres.' });
        }
        if (referenciaCliente && exceedsMaxLength(referenciaCliente, 500)) {
            return res.status(400).json({ message: 'La referencia personalizada excede el maximo permitido de 500 caracteres.' });
        }
        if (parcelamientoRangeInput && exceedsMaxLength(parcelamientoRangeInput, 80)) {
            return res.status(400).json({ message: 'El rango de parcelamiento excede el maximo permitido de 80 caracteres.' });
        }

        const summaryResult = await buildCurrentCotizacionSummary();
        if (!summaryResult.ok) {
            return res.status(409).json({
                message: 'Debes registrar la UTM mensual antes de enviar la cotizacion.',
                status: summaryResult.status || null
            });
        }

        const summary = summaryResult.summary;
        if (!Array.isArray(summary?.servicios) || summary.servicios.length === 0) {
            return res.status(409).json({
                message: 'No hay servicios estandar activos para generar la cotizacion.'
            });
        }
        const parcelamientoRows =
            Array.isArray(summary?.parcelamiento) && summary.parcelamiento.length > 0
                ? summary.parcelamiento.map((row) => ({
                      key: normalizeCotizacionParcelamientoKey(row?.key || row?.lotes),
                      lotes: normalizeText(row?.lotes),
                      valorUtm: Number(row?.valorUtm),
                      neto: Number(row?.neto),
                      iva: Number(row?.iva),
                      total: Number(row?.total)
                  }))
                : buildCotizacionParcelamientoRows(summary?.utm?.valor);

        let selectedParcelamiento = null;
        if (parcelamientoRangeKey) {
            selectedParcelamiento = parcelamientoRows.find((row) => normalizeCotizacionParcelamientoKey(row?.key || row?.lotes) === parcelamientoRangeKey) || null;
            if (!selectedParcelamiento) {
                return res.status(400).json({ message: 'El rango de parcelamiento seleccionado no es valido.' });
            }
        }
        const emailParcelamientoRows = selectedParcelamiento ? [selectedParcelamiento] : parcelamientoRows;

        const attachmentBuffer = await buildCotizacionTemplatePdfBuffer(summary);
        const periodLabel = formatMonthlyPeriodLabel(summary?.periodo?.anio, summary?.periodo?.mes);
        const subject = `Cotizacion Geo Rural ${periodLabel}`;
        const transporter = getSmtpTransporter(smtpResolution.config);
        if (!transporter) {
            return res.status(503).json({
                message: 'Envio no disponible: no fue posible inicializar configuracion SMTP.'
            });
        }

        const fromAddress = buildSmtpFromAddress(smtpResolution.config);
        const mailTemplates = resolveMailTemplatesForUse(smtpResolution.storedRow);
        const emailBody = buildCotizacionEmailBody(summary, clientName, referenciaCliente, req.authUser, {
            parcelamientoRows: emailParcelamientoRows,
            selectedParcelamiento,
            customHtmlTemplate: mailTemplates.cotizacionHtml
        });
        await transporter.sendMail({
            from: fromAddress,
            to: destinationEmail,
            subject,
            text: emailBody.text,
            html: emailBody.html,
            attachments: [
                {
                    filename: buildCotizacionAttachmentFileName(summary),
                    content: attachmentBuffer,
                    contentType: 'application/pdf'
                }
            ]
        });

        return res.json({
            message: `Cotizacion enviada correctamente a ${destinationEmail}.`,
            destinationEmail,
            subject
        });
    } catch (error) {
        console.error('Error enviando cotizacion por correo:', error);
        return res.status(500).json({ message: 'No fue posible enviar la cotizacion por correo.' });
    }
});

app.get('/api/facturas/solicitudes', async (req, res) => {
    try {
        if (!canUseInvoiceWorkflow(req.authUser)) {
            return res.status(403).json({ message: 'Solo OPERADOR, SECRETARIA o SUPER pueden consultar solicitudes de factura.' });
        }

        const estadoQuery = normalizeText(req.query.estado).toUpperCase();
        const limitRaw = Number(req.query.limit || 300);
        const limit = Number.isInteger(limitRaw) ? Math.max(1, Math.min(limitRaw, 1000)) : 300;
        const estado =
            !estadoQuery || estadoQuery === 'TODAS' || estadoQuery === 'TODOS' || estadoQuery === 'ALL'
                ? ''
                : normalizeFacturaSolicitudEstado(estadoQuery);

        if (estadoQuery && !estado && !['TODAS', 'TODOS', 'ALL'].includes(estadoQuery)) {
            return res.status(400).json({ message: 'Estado invalido para buscar solicitudes. Usa PENDIENTE, ENVIADA o ANULADA.' });
        }

        const accessScope = getFacturaSolicitudAccessScope(req.authUser, 's');
        if (accessScope.denied) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para consultar solicitudes.' });
        }

        const where = [];
        const values = [];
        if (estado) {
            where.push('s.estado = ?');
            values.push(estado);
        }
        if (accessScope.clause) {
            where.push(accessScope.clause);
            values.push(...accessScope.params);
        }

        const whereSql = where.length > 0 ? `WHERE ${where.join(' AND ')}` : '';
        const [rows] = await pool.execute(
            `
                SELECT
                    s.id,
                    s.num_ingreso,
                    s.nombre_razon_social,
                    s.rut,
                    s.giro,
                    s.direccion,
                    s.comuna,
                    s.ciudad,
                    s.contacto,
                    s.observacion,
                    s.monto_facturar,
                    s.destino_email,
                    s.estado,
                    s.created_by,
                    s.created_by_sucursal,
                    s.sent_at,
                    s.created_at
                FROM factura_solicitudes s
                ${whereSql}
                ORDER BY s.created_at DESC, s.id DESC
                LIMIT ${limit}
            `,
            values
        );

        const solicitudes = rows.map(toFacturaSolicitudRow);
        return res.json({ solicitudes });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible listar las solicitudes de factura.' });
    }
});

app.get('/api/facturas/solicitudes/por-ingreso/:numIngreso', async (req, res) => {
    try {
        if (!canUseInvoiceWorkflow(req.authUser)) {
            return res.status(403).json({ message: 'Solo OPERADOR, SECRETARIA o SUPER pueden consultar solicitudes de factura.' });
        }

        const numIngreso = normalizeText(decodeURIComponent(req.params.numIngreso));
        if (!parseNumIngreso(numIngreso)) {
            return res.status(400).json({ message: 'NRO INGRESO invalido.' });
        }

        const accessScope = getFacturaSolicitudAccessScope(req.authUser, 's');
        if (accessScope.denied) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para consultar solicitudes.' });
        }

        const scopeWhere = accessScope.clause ? ` AND ${accessScope.clause}` : '';
        const [rows] = await pool.execute(
            `
                SELECT
                    s.id,
                    s.num_ingreso,
                    s.nombre_razon_social,
                    s.rut,
                    s.giro,
                    s.direccion,
                    s.comuna,
                    s.ciudad,
                    s.contacto,
                    s.observacion,
                    s.monto_facturar,
                    s.destino_email,
                    s.estado,
                    s.created_by,
                    s.created_by_sucursal,
                    s.sent_at,
                    s.created_at
                FROM factura_solicitudes s
                WHERE s.num_ingreso = ?${scopeWhere}
                ORDER BY
                    (s.estado = ?) DESC,
                    COALESCE(s.sent_at, s.updated_at, s.created_at) DESC,
                    s.id DESC
                LIMIT 1
            `,
            [numIngreso, ...accessScope.params, FACTURA_SOLICITUD_ESTADO_ENVIADA]
        );

        if (!rows.length) {
            return res.json({ solicitud: null });
        }

        return res.json({ solicitud: toFacturaSolicitudRow(rows[0]) });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible consultar la solicitud de factura.' });
    }
});

app.post('/api/facturas/solicitudes', requireWriteAccess, async (req, res) => {
    if (!canUseInvoiceWorkflow(req.authUser)) {
        return res.status(403).json({ message: 'Solo OPERADOR, SECRETARIA o SUPER pueden gestionar solicitudes de factura.' });
    }

    const connection = await pool.getConnection();
    try {
        const parsed = parseFacturaSolicitudPayload(req.body);
        if (!parsed.isValid) {
            return res.status(400).json({ message: parsed.message });
        }

        const payload = parsed.payload;
        const userName = req.authUser.nombre || req.authUser.username;
        const userBranch = normalizeText(req.authUser.sucursal) || null;
        const canManageManualStatus = isSuperUser(req.authUser);
        const accessScope = getFacturaSolicitudAccessScope(req.authUser);
        const registroScope = getRegistroAccessScope(req.authUser);

        if (accessScope.denied || (!hasBackofficeRegistroAccess(req.authUser) && !userBranch)) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para registrar solicitudes.' });
        }
        if (!canManageManualStatus && payload.estado === FACTURA_SOLICITUD_ESTADO_ANULADA) {
            return res.status(403).json({ message: 'Solo el superusuario puede marcar solicitudes como ANULADA.' });
        }

        await connection.beginTransaction();
        const syncFacturaDataIntoRegistro = async () => {
            if (registroScope.denied) {
                return;
            }

            const registroScopeWhere = registroScope.clause ? ` AND ${registroScope.clause}` : '';
            await connection.execute(
                `
                    UPDATE registros
                    SET
                        factura_nombre_razon = ?,
                        factura_rut = ?,
                        factura_giro = ?,
                        factura_direccion = ?,
                        factura_comuna = ?,
                        factura_ciudad = ?,
                        factura_contacto = ?,
                        factura_observacion = ?,
                        factura_monto = ?,
                        updated_by = ?,
                        updated_by_sucursal = ?,
                        updated_at = NOW()
                    WHERE num_ingreso = ?${registroScopeWhere}
                `,
                [
                    payload.nombreRazonSocial,
                    payload.rut,
                    payload.giro || null,
                    payload.direccion || null,
                    payload.comuna || null,
                    payload.ciudad || null,
                    payload.contacto || null,
                    payload.observacion || null,
                    payload.montoFacturar,
                    userName,
                    userBranch,
                    payload.numIngreso,
                    ...registroScope.params
                ]
            );
        };

        const scopeWhere = accessScope.clause ? ` AND ${accessScope.clause}` : '';
        const [existingRows] = await connection.execute(
            `
                SELECT
                    id,
                    estado,
                    destino_email,
                    sent_at
                FROM factura_solicitudes
                WHERE num_ingreso = ?${scopeWhere}
                LIMIT 1
                FOR UPDATE
            `,
            [payload.numIngreso, ...accessScope.params]
        );

        if (!existingRows.length) {
            const [insertResult] = await connection.execute(
                `
                    INSERT INTO factura_solicitudes (
                        num_ingreso,
                        nombre_razon_social,
                        rut,
                        giro,
                        direccion,
                        comuna,
                        ciudad,
                        contacto,
                        observacion,
                        monto_facturar,
                        destino_email,
                        estado,
                        created_by,
                        created_by_sucursal,
                        updated_by,
                        updated_by_sucursal,
                        sent_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                `,
                [
                    payload.numIngreso,
                    payload.nombreRazonSocial,
                    payload.rut,
                    payload.giro || null,
                    payload.direccion || null,
                    payload.comuna || null,
                    payload.ciudad || null,
                    payload.contacto || null,
                    payload.observacion || null,
                    payload.montoFacturar,
                    payload.destinationEmail || null,
                    payload.estado,
                    userName,
                    userBranch,
                    userName,
                    userBranch,
                    payload.estado === FACTURA_SOLICITUD_ESTADO_ENVIADA ? new Date() : null
                ]
            );

            await syncFacturaDataIntoRegistro();
            await connection.commit();
            return res.status(201).json({
                message: payload.estado === FACTURA_SOLICITUD_ESTADO_ENVIADA
                    ? 'Solicitud registrada como ENVIADA.'
                    : payload.estado === FACTURA_SOLICITUD_ESTADO_ANULADA
                      ? 'Solicitud registrada como ANULADA.'
                      : 'Solicitud registrada como PENDIENTE.',
                solicitudId: insertResult.insertId,
                estado: payload.estado
            });
        }

        const existingStatus = normalizeFacturaSolicitudEstado(existingRows[0].estado) || FACTURA_SOLICITUD_ESTADO_PENDIENTE;
        const blockDowngradeToPending =
            !canManageManualStatus &&
            existingStatus === FACTURA_SOLICITUD_ESTADO_ENVIADA &&
            payload.estado === FACTURA_SOLICITUD_ESTADO_PENDIENTE;
        const nextStatus = blockDowngradeToPending ? FACTURA_SOLICITUD_ESTADO_ENVIADA : payload.estado;
        const fallbackDestinationEmail = normalizeText(existingRows[0].destino_email).toLowerCase();
        const nextDestinationEmail =
            nextStatus === FACTURA_SOLICITUD_ESTADO_ENVIADA
                ? payload.destinationEmail || fallbackDestinationEmail || null
                : null;
        const existingSentAt = existingRows[0].sent_at ? new Date(existingRows[0].sent_at) : null;
        const nextSentAt = nextStatus === FACTURA_SOLICITUD_ESTADO_ENVIADA ? existingSentAt || new Date() : null;

        await connection.execute(
            `
                UPDATE factura_solicitudes
                SET
                    nombre_razon_social = ?,
                    rut = ?,
                    giro = ?,
                    direccion = ?,
                    comuna = ?,
                    ciudad = ?,
                    contacto = ?,
                    observacion = ?,
                    monto_facturar = ?,
                    destino_email = ?,
                    estado = ?,
                    updated_by = ?,
                    updated_by_sucursal = ?,
                    sent_at = ?,
                    updated_at = NOW()
                WHERE id = ?
            `,
            [
                payload.nombreRazonSocial,
                payload.rut,
                payload.giro || null,
                payload.direccion || null,
                payload.comuna || null,
                payload.ciudad || null,
                payload.contacto || null,
                payload.observacion || null,
                payload.montoFacturar,
                nextDestinationEmail,
                nextStatus,
                userName,
                userBranch,
                nextSentAt,
                existingRows[0].id
            ]
        );

        await syncFacturaDataIntoRegistro();
        await connection.commit();
        return res.json({
            message: blockDowngradeToPending
                ? 'La factura ya estaba ENVIADA. Se actualizaron sus datos sin cambiar el estado.'
                : nextStatus === FACTURA_SOLICITUD_ESTADO_ENVIADA
                  ? 'Solicitud actualizada en estado ENVIADA.'
                  : nextStatus === FACTURA_SOLICITUD_ESTADO_ANULADA
                    ? 'Solicitud actualizada en estado ANULADA.'
                    : 'Solicitud actualizada en estado PENDIENTE.',
            solicitudId: existingRows[0].id,
            estado: nextStatus
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        if (error && error.code === 'ER_DUP_ENTRY') {
            return res.status(409).json({ message: 'La solicitud de esta factura ya existe. Vuelve a cargar el estado.' });
        }
        return res.status(500).json({ message: 'No fue posible registrar la solicitud de factura.' });
    } finally {
        connection.release();
    }
});

app.post('/api/facturas/solicitudes/marcar-enviadas', requireWriteAccess, async (req, res) => {
    try {
        if (!canUseInvoiceWorkflow(req.authUser)) {
            return res.status(403).json({ message: 'Solo OPERADOR, SECRETARIA o SUPER pueden marcar solicitudes como enviadas.' });
        }

        const destinationEmail = normalizeText(req.body.destinationEmail).toLowerCase();
        if (!destinationEmail) {
            return res.status(400).json({ message: 'Debes ingresar el correo del contador.' });
        }
        if (!isValidEmail(destinationEmail)) {
            return res.status(400).json({ message: 'El correo del contador no es valido.' });
        }
        if (exceedsMaxLength(destinationEmail, 255)) {
            return res.status(400).json({ message: 'El correo del contador excede el maximo permitido de 255 caracteres.' });
        }

        const rawNumIngresos = Array.isArray(req.body?.numIngresos) ? req.body.numIngresos : [];
        const numIngresos = [];
        for (const item of rawNumIngresos) {
            const normalizedIngreso = normalizeText(item);
            if (!normalizedIngreso) {
                continue;
            }
            if (!parseNumIngreso(normalizedIngreso)) {
                return res.status(400).json({ message: 'Uno de los NRO INGRESO enviados es invalido. Usa formato 123-2026.' });
            }
            if (!numIngresos.includes(normalizedIngreso)) {
                numIngresos.push(normalizedIngreso);
            }
        }
        if (rawNumIngresos.length > 0 && numIngresos.length === 0) {
            return res.status(400).json({ message: 'Debes indicar al menos un NRO INGRESO valido para actualizar.' });
        }

        const userName = req.authUser.nombre || req.authUser.username;
        const userBranch = normalizeText(req.authUser.sucursal) || null;
        if (!hasBackofficeRegistroAccess(req.authUser) && !userBranch) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para enviar pendientes.' });
        }

        const accessScope = getFacturaSolicitudAccessScope(req.authUser);
        if (accessScope.denied) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para enviar pendientes.' });
        }

        const where = ['estado = ?'];
        const values = [FACTURA_SOLICITUD_ESTADO_PENDIENTE];
        if (accessScope.clause) {
            where.push(accessScope.clause);
            values.push(...accessScope.params);
        }
        if (numIngresos.length > 0) {
            where.push(`num_ingreso IN (${numIngresos.map(() => '?').join(', ')})`);
            values.push(...numIngresos);
        }

        const [updateResult] = await pool.execute(
            `
                UPDATE factura_solicitudes
                SET
                    estado = ?,
                    destino_email = ?,
                    sent_at = NOW(),
                    updated_by = ?,
                    updated_by_sucursal = ?,
                    updated_at = NOW()
                WHERE ${where.join(' AND ')}
            `,
            [FACTURA_SOLICITUD_ESTADO_ENVIADA, destinationEmail, userName, userBranch, ...values]
        );

        const updatedCount = Number(updateResult.affectedRows || 0);
        return res.json({
            message:
                updatedCount > 0
                    ? `Se marcaron ${updatedCount} solicitud(es) como ENVIADA.`
                    : 'No hay solicitudes pendientes para actualizar.',
            updatedCount
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible actualizar las solicitudes pendientes.' });
    }
});

app.post('/api/facturas/correo/enviar', requireWriteAccess, async (req, res) => {
    try {
        if (!canUseInvoiceWorkflow(req.authUser)) {
            return res.status(403).json({ message: 'Solo OPERADOR, SECRETARIA o SUPER pueden enviar facturas por correo.' });
        }

        const smtpResolution = await resolveActiveSmtpConfiguration();
        if (!smtpResolution.config) {
            return res.status(503).json({
                message: 'Envio no disponible: configura el servicio de correo (SMTP) en el menu del superusuario o variables de entorno.'
            });
        }

        const destinationEmail = normalizeText(req.body?.destinationEmail).toLowerCase();
        const subject = normalizeMailSubjectLine(req.body?.subject);
        const textBody = normalizeMailTextContent(req.body?.textBody ?? req.body?.text);
        const htmlBody = normalizeMailHtmlContent(req.body?.htmlBody ?? req.body?.html);

        if (!destinationEmail) {
            return res.status(400).json({ message: 'Debes ingresar el correo del contador.' });
        }
        if (!isValidEmail(destinationEmail)) {
            return res.status(400).json({ message: 'El correo del contador no es valido.' });
        }
        if (exceedsMaxLength(destinationEmail, 255)) {
            return res.status(400).json({ message: 'El correo del contador excede el maximo permitido de 255 caracteres.' });
        }
        if (!subject) {
            return res.status(400).json({ message: 'Debes indicar el asunto del correo.' });
        }
        if (exceedsMaxLength(subject, 200)) {
            return res.status(400).json({ message: 'El asunto del correo excede el maximo permitido de 200 caracteres.' });
        }
        if (!textBody && !htmlBody) {
            return res.status(400).json({ message: 'Debes indicar el contenido del correo a enviar.' });
        }
        if (exceedsMaxLength(textBody, 200000)) {
            return res.status(400).json({ message: 'El contenido de texto del correo excede el maximo permitido.' });
        }
        if (exceedsMaxLength(htmlBody, 400000)) {
            return res.status(400).json({ message: 'El contenido HTML del correo excede el maximo permitido.' });
        }

        const transporter = getSmtpTransporter(smtpResolution.config);
        if (!transporter) {
            return res.status(503).json({
                message: 'Envio no disponible: no fue posible inicializar configuracion SMTP.'
            });
        }

        await transporter.sendMail({
            from: buildSmtpFromAddress(smtpResolution.config),
            to: destinationEmail,
            subject,
            text: textBody || undefined,
            html: htmlBody || undefined
        });

        return res.json({
            message: `Correo enviado correctamente a ${destinationEmail}.`,
            destinationEmail,
            subject
        });
    } catch (error) {
        console.error('Error enviando factura por correo:', error);
        return res.status(500).json({ message: 'No fue posible enviar la factura por correo.' });
    }
});

app.get('/api/admin/sucursales', requireAdmin, async (req, res) => {
    try {
        const [rows] = await pool.execute(
            `
                SELECT id, nombre, created_at
                FROM sucursales
                ORDER BY nombre ASC
            `
        );

        const sucursales = rows.map((row) => ({
            id: row.id,
            nombre: row.nombre,
            createdAt: row.created_at instanceof Date ? row.created_at.toISOString() : String(row.created_at)
        }));

        return res.json({ sucursales });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible listar las sucursales.' });
    }
});

app.post('/api/admin/sucursales', requireAdmin, async (req, res) => {
    try {
        const nombre = normalizeText(req.body.nombre);
        if (!nombre) {
            return res.status(400).json({ message: 'Debes ingresar un nombre de sucursal.' });
        }

        if (exceedsMaxLength(nombre, 120)) {
            return res.status(400).json({ message: 'El nombre de sucursal no puede exceder 120 caracteres.' });
        }

        const [insertResult] = await pool.execute('INSERT INTO sucursales (nombre) VALUES (?)', [nombre]);
        return res.status(201).json({
            message: `Sucursal ${nombre} creada correctamente.`,
            sucursal: {
                id: insertResult.insertId,
                nombre
            }
        });
    } catch (error) {
        console.error(error);
        if (error && error.code === 'ER_DUP_ENTRY') {
            return res.status(409).json({ message: 'Esa sucursal ya existe.' });
        }

        return res.status(500).json({ message: 'No fue posible crear la sucursal.' });
    }
});

app.put('/api/admin/sucursales/:branchId', requireAdmin, async (req, res) => {
    const connection = await pool.getConnection();
    try {
        const branchId = Number(req.params.branchId);
        const nombre = normalizeText(req.body.nombre);

        if (!Number.isInteger(branchId) || branchId <= 0) {
            return res.status(400).json({ message: 'Identificador de sucursal invalido.' });
        }

        if (!nombre) {
            return res.status(400).json({ message: 'Debes ingresar un nombre de sucursal.' });
        }

        if (exceedsMaxLength(nombre, 120)) {
            return res.status(400).json({ message: 'El nombre de sucursal no puede exceder 120 caracteres.' });
        }

        await connection.beginTransaction();
        const [rows] = await connection.execute('SELECT id, nombre FROM sucursales WHERE id = ? FOR UPDATE', [branchId]);
        if (!rows.length) {
            await connection.rollback();
            return res.status(404).json({ message: 'No existe esa sucursal.' });
        }

        const previousName = normalizeText(rows[0].nombre);
        await connection.execute('UPDATE sucursales SET nombre = ?, updated_at = NOW() WHERE id = ?', [nombre, branchId]);

        if (previousName && previousName !== nombre) {
            await connection.execute('UPDATE usuarios SET sucursal = ? WHERE sucursal = ?', [nombre, previousName]);
            await connection.execute('UPDATE registros SET created_by_sucursal = ? WHERE created_by_sucursal = ?', [nombre, previousName]);
            await connection.execute('UPDATE registros SET updated_by_sucursal = ? WHERE updated_by_sucursal = ?', [nombre, previousName]);
            await connection.execute('UPDATE registro_historial SET sucursal = ? WHERE sucursal = ?', [nombre, previousName]);
            await connection.execute('UPDATE factura_solicitudes SET created_by_sucursal = ? WHERE created_by_sucursal = ?', [nombre, previousName]);
            await connection.execute('UPDATE factura_solicitudes SET updated_by_sucursal = ? WHERE updated_by_sucursal = ?', [nombre, previousName]);
        }

        await connection.commit();
        return res.json({
            message: `Sucursal ${nombre} actualizada correctamente.`,
            sucursal: {
                id: branchId,
                nombre
            }
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        if (error && error.code === 'ER_DUP_ENTRY') {
            return res.status(409).json({ message: 'Esa sucursal ya existe.' });
        }

        return res.status(500).json({ message: 'No fue posible actualizar la sucursal.' });
    } finally {
        connection.release();
    }
});

app.delete('/api/admin/sucursales/:branchId', requireAdmin, async (req, res) => {
    try {
        const branchId = Number(req.params.branchId);
        if (!Number.isInteger(branchId) || branchId <= 0) {
            return res.status(400).json({ message: 'Identificador de sucursal invalido.' });
        }

        const [rows] = await pool.execute('SELECT id, nombre FROM sucursales WHERE id = ? LIMIT 1', [branchId]);
        if (!rows.length) {
            return res.status(404).json({ message: 'No existe esa sucursal.' });
        }

        const branchName = normalizeText(rows[0].nombre);
        const [assignedRows] = await pool.execute('SELECT COUNT(*) AS total FROM usuarios WHERE sucursal = ?', [branchName]);
        if (Number(assignedRows[0].total || 0) > 0) {
            return res.status(400).json({
                message: 'No se puede eliminar la sucursal porque esta asignada a usuarios activos.'
            });
        }

        await pool.execute('DELETE FROM sucursales WHERE id = ?', [branchId]);
        return res.json({ message: `Sucursal ${branchName} eliminada.` });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible eliminar la sucursal.' });
    }
});

app.get('/api/admin/usuarios', requireAdmin, async (req, res) => {
    try {
        const [rows] = await pool.execute(
            `
                SELECT
                    id,
                    username,
                    nombre,
                    sucursal,
                    role,
                    must_change_password,
                    activo,
                    created_at
                FROM usuarios
                ORDER BY
                    CASE WHEN LOWER(username) = ? THEN 1 ELSE 0 END ASC,
                    id ASC
            `
            ,
            [GUEST_USERNAME]
        );

        const visibleRows = isSuperUser(req.authUser)
            ? rows
            : rows.filter(
                  (row) =>
                      normalizeUserRole(row.role, ROLE_OPERADOR) !== ROLE_SUPER &&
                      !isGuestUsername(row.username)
              );

        const usuarios = visibleRows.map((row) => {
            const isGuest = isGuestUsername(row.username);
            return {
                id: row.id,
                username: row.username,
                nombre: row.nombre,
                sucursal: row.sucursal,
                role: isGuest ? 'INVITADO' : normalizeUserRole(row.role, ROLE_OPERADOR),
                mustChangePassword: Number(row.must_change_password || 0) === 1,
                activo: Number(row.activo) === 1,
                isGuest,
                createdAt: row.created_at instanceof Date ? row.created_at.toISOString() : String(row.created_at)
            };
        });

        return res.json({ usuarios });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible listar los usuarios.' });
    }
});

app.post('/api/admin/usuarios', requireAdmin, async (req, res) => {
    try {
        const username = normalizeText(req.body.username).toLowerCase();
        const nombre = normalizeText(req.body.nombre);
        const sucursal = normalizeText(req.body.sucursal);
        const password = String(req.body.password || '');
        const roleRaw = normalizeText(req.body.role).toUpperCase();
        const role = normalizeUserRole(roleRaw || ROLE_OPERADOR, ROLE_OPERADOR);
        const assignableRoles = getAssignableRolesForUser(req.authUser);
        if (roleRaw && (role !== roleRaw || !assignableRoles.includes(role))) {
            return res.status(400).json({ message: `ROL invalido. Usa ${formatAssignableRolesLabel(req.authUser)}.` });
        }
        if (role === ROLE_ADMIN && !isSuperUser(req.authUser)) {
            return res.status(403).json({ message: 'Solo el superusuario puede crear usuarios con rol ADMIN.' });
        }

        if (!/^[a-z0-9._-]{3,40}$/.test(username)) {
            return res.status(400).json({ message: 'Usuario invalido. Usa 3-40 caracteres (a-z, 0-9, ., _, -).' });
        }

        if (!nombre || !sucursal || password.length < 6) {
            return res.status(400).json({ message: 'Completa nombre, sucursal y contrasena (minimo 6 caracteres).' });
        }

        if (exceedsMaxLength(nombre, 120) || exceedsMaxLength(sucursal, 120)) {
            return res.status(400).json({ message: 'Nombre y sucursal no pueden exceder 120 caracteres.' });
        }

        if (exceedsMaxLength(password, MAX_PASSWORD_LENGTH)) {
            return res.status(400).json({ message: `La contrasena no puede exceder ${MAX_PASSWORD_LENGTH} caracteres.` });
        }

        if (!(await branchExists(sucursal))) {
            return res.status(400).json({ message: 'La sucursal seleccionada no existe o no esta disponible.' });
        }

        if (role === ROLE_SUPER && (await countSuperUsers()) > 0) {
            return res.status(400).json({ message: 'Solo se permite una cuenta con rol SUPER en el sistema.' });
        }

        const mustChangePassword = shouldForcePasswordChangeForRole(role) ? 1 : 0;
        const passwordData = createPasswordHash(password);
        const [insertResult] = await pool.execute(
            `
                INSERT INTO usuarios (username, nombre, sucursal, role, password_salt, password_hash, must_change_password, activo)
                VALUES (?, ?, ?, ?, ?, ?, ?, 1)
            `,
            [username, nombre, sucursal, role, passwordData.salt, passwordData.hash, mustChangePassword]
        );

        return res.status(201).json({
            message: `Usuario ${username} creado correctamente.`,
            usuario: {
                id: insertResult.insertId,
                username,
                nombre,
                sucursal,
                role,
                mustChangePassword: mustChangePassword === 1,
                activo: true
            }
        });
    } catch (error) {
        console.error(error);
        if (error && error.code === 'ER_DUP_ENTRY') {
            return res.status(409).json({ message: 'Ese usuario ya existe.' });
        }

        return res.status(500).json({ message: 'No fue posible crear el usuario.' });
    }
});

app.put('/api/admin/usuarios/:userId', requireAdmin, async (req, res) => {
    try {
        const userId = Number(req.params.userId);
        const username = normalizeText(req.body.username).toLowerCase();
        const nombre = normalizeText(req.body.nombre);
        const sucursal = normalizeText(req.body.sucursal);
        const roleRaw = normalizeText(req.body.role).toUpperCase();
        const role = normalizeUserRole(roleRaw || ROLE_OPERADOR, ROLE_OPERADOR);
        const password = String(req.body.password || '');
        const assignableRoles = getAssignableRolesForUser(req.authUser);

        if (roleRaw && (role !== roleRaw || !assignableRoles.includes(role))) {
            return res.status(400).json({ message: `ROL invalido. Usa ${formatAssignableRolesLabel(req.authUser)}.` });
        }

        if (!Number.isInteger(userId) || userId <= 0) {
            return res.status(400).json({ message: 'Identificador de usuario invalido.' });
        }

        if (!/^[a-z0-9._-]{3,40}$/.test(username)) {
            return res.status(400).json({ message: 'Usuario invalido. Usa 3-40 caracteres (a-z, 0-9, ., _, -).' });
        }

        if (!nombre || !sucursal) {
            return res.status(400).json({ message: 'Completa nombre y sucursal.' });
        }

        if (exceedsMaxLength(nombre, 120) || exceedsMaxLength(sucursal, 120)) {
            return res.status(400).json({ message: 'Nombre y sucursal no pueden exceder 120 caracteres.' });
        }

        if (!(await branchExists(sucursal))) {
            return res.status(400).json({ message: 'La sucursal seleccionada no existe o no esta disponible.' });
        }

        if (password && password.length < 6) {
            return res.status(400).json({ message: 'Si envias una contrasena, debe tener al menos 6 caracteres.' });
        }

        if (password && exceedsMaxLength(password, MAX_PASSWORD_LENGTH)) {
            return res.status(400).json({ message: `La contrasena no puede exceder ${MAX_PASSWORD_LENGTH} caracteres.` });
        }

        const [targetRows] = await pool.execute(
            'SELECT id, username, role, must_change_password, activo FROM usuarios WHERE id = ? LIMIT 1',
            [userId]
        );
        if (!targetRows.length) {
            return res.status(404).json({ message: 'No existe ese usuario.' });
        }

        if (isGuestUsername(targetRows[0].username)) {
            return res
                .status(400)
                .json({ message: 'La cuenta invitado no se edita aqui. Usa su checkbox para habilitarla o deshabilitarla.' });
        }

        const targetUsername = normalizeText(targetRows[0].username).toLowerCase();
        if (targetUsername === SUPERUSER_USERNAME) {
            if (username !== SUPERUSER_USERNAME) {
                return res.status(400).json({ message: 'La cuenta SUPER principal no puede cambiar su usuario.' });
            }
            if (role !== ROLE_SUPER) {
                return res.status(400).json({ message: 'La cuenta SUPER principal no puede perder su rol SUPER.' });
            }
        }

        const currentRole = normalizeUserRole(targetRows[0].role, ROLE_OPERADOR);
        if (role === ROLE_ADMIN && !isSuperUser(req.authUser) && currentRole !== ROLE_ADMIN) {
            return res.status(403).json({ message: 'Solo el superusuario puede asignar el rol ADMIN.' });
        }
        if (currentRole === ROLE_SUPER && !isSuperUser(req.authUser)) {
            return res.status(403).json({ message: 'Solo el superusuario puede editar otra cuenta SUPER.' });
        }

        if (role === ROLE_SUPER && !isSuperUser(req.authUser)) {
            return res.status(403).json({ message: 'Solo el superusuario puede asignar el rol SUPER.' });
        }

        if (role === ROLE_SUPER && currentRole !== ROLE_SUPER && (await countSuperUsers(userId)) > 0) {
            return res.status(400).json({ message: 'Solo se permite una cuenta con rol SUPER en el sistema.' });
        }

        const targetIsActive = Number(targetRows[0].activo || 0) === 1;
        if (targetIsActive && isPrivilegedRole(currentRole) && !isPrivilegedRole(role)) {
            if ((await countActivePrivilegedUsers()) <= 1) {
                return res.status(400).json({ message: 'Debe existir al menos un usuario ADMIN o SUPER activo.' });
            }
        }

        let mustChangePassword = Number(targetRows[0].must_change_password || 0) === 1;
        if (role === ROLE_SUPER) {
            mustChangePassword = false;
        }

        if (password) {
            const passwordData = createPasswordHash(password);
            mustChangePassword = shouldForcePasswordChangeForRole(role);
            await pool.execute(
                `
                    UPDATE usuarios
                    SET
                        username = ?,
                        nombre = ?,
                        sucursal = ?,
                        role = ?,
                        password_salt = ?,
                        password_hash = ?,
                        must_change_password = ?,
                        updated_at = NOW()
                    WHERE id = ?
                `,
                [username, nombre, sucursal, role, passwordData.salt, passwordData.hash, mustChangePassword ? 1 : 0, userId]
            );

            if (Number(userId) === Number(req.authUser?.id)) {
                await pool.execute('UPDATE sesiones SET revoked_at = NOW() WHERE user_id = ? AND id <> ?', [
                    userId,
                    req.authSessionId
                ]);
            } else {
                await pool.execute('UPDATE sesiones SET revoked_at = NOW() WHERE user_id = ?', [userId]);
            }
        } else {
            if (isPrivilegedRole(role)) {
                await pool.execute(
                    `
                        UPDATE usuarios
                        SET username = ?, nombre = ?, sucursal = ?, role = ?, must_change_password = 0, updated_at = NOW()
                        WHERE id = ?
                    `,
                    [username, nombre, sucursal, role, userId]
                );
            } else {
                await pool.execute(
                    `
                        UPDATE usuarios
                        SET username = ?, nombre = ?, sucursal = ?, role = ?, must_change_password = ?, updated_at = NOW()
                        WHERE id = ?
                    `,
                    [username, nombre, sucursal, role, mustChangePassword ? 1 : 0, userId]
                );
            }
        }

        return res.json({
            message: `Usuario ${username} actualizado correctamente.`,
            usuario: {
                id: userId,
                username,
                nombre,
                sucursal,
                role,
                mustChangePassword
            }
        });
    } catch (error) {
        console.error(error);
        if (error && error.code === 'ER_DUP_ENTRY') {
            return res.status(409).json({ message: 'Ese usuario ya existe.' });
        }

        return res.status(500).json({ message: 'No fue posible actualizar el usuario.' });
    }
});

app.post('/api/admin/usuarios/:userId/guest-status', requireSuper, async (req, res) => {
    try {
        const userId = Number(req.params.userId);
        if (!Number.isInteger(userId) || userId <= 0) {
            return res.status(400).json({ message: 'Identificador de usuario invalido.' });
        }

        const activeRaw = req.body && req.body.activo;
        if (typeof activeRaw !== 'boolean') {
            return res.status(400).json({ message: 'Debes indicar el estado activo como true o false.' });
        }

        const [targetRows] = await pool.execute('SELECT id, username, activo FROM usuarios WHERE id = ? LIMIT 1', [userId]);
        if (!targetRows.length) {
            return res.status(404).json({ message: 'No existe ese usuario.' });
        }

        if (!isGuestUsername(targetRows[0].username)) {
            return res.status(400).json({ message: 'Solo se puede cambiar estado de la cuenta invitado.' });
        }

        const activeValue = activeRaw ? 1 : 0;
        await pool.execute('UPDATE usuarios SET activo = ?, updated_at = NOW() WHERE id = ?', [activeValue, userId]);

        if (!activeRaw) {
            await pool.execute('UPDATE sesiones SET revoked_at = NOW() WHERE user_id = ?', [userId]);
        }

        return res.json({
            message: activeRaw
                ? 'Cuenta invitado habilitada correctamente.'
                : 'Cuenta invitado deshabilitada correctamente.',
            usuario: {
                id: userId,
                username: targetRows[0].username,
                activo: activeRaw,
                isGuest: true
            }
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible actualizar el estado de la cuenta invitado.' });
    }
});

app.post('/api/admin/usuarios/:userId/reset-password', requireAdmin, async (req, res) => {
    try {
        const userId = Number(req.params.userId);
        const newPassword = String(req.body.newPassword || '');

        if (!Number.isInteger(userId) || userId <= 0) {
            return res.status(400).json({ message: 'Identificador de usuario invalido.' });
        }

        if (newPassword.length < 6) {
            return res.status(400).json({ message: 'La clave temporal debe tener al menos 6 caracteres.' });
        }

        if (exceedsMaxLength(newPassword, MAX_PASSWORD_LENGTH)) {
            return res.status(400).json({ message: `La clave no puede exceder ${MAX_PASSWORD_LENGTH} caracteres.` });
        }

        const [targetRows] = await pool.execute('SELECT id, username, role, activo FROM usuarios WHERE id = ? LIMIT 1', [userId]);
        if (!targetRows.length) {
            return res.status(404).json({ message: 'No existe ese usuario.' });
        }

        if (isGuestUsername(targetRows[0].username)) {
            return res
                .status(400)
                .json({ message: 'La cuenta invitado no usa reseteo de clave manual. Solo habilitala o deshabilitala.' });
        }

        const targetRole = normalizeUserRole(targetRows[0].role, ROLE_OPERADOR);
        if (targetRole === ROLE_SUPER && !isSuperUser(req.authUser)) {
            return res.status(403).json({ message: 'Solo el superusuario puede restablecer la clave de una cuenta SUPER.' });
        }

        const mustChangePassword = shouldForcePasswordChangeForRole(targetRole) ? 1 : 0;
        const passwordData = createPasswordHash(newPassword);
        await pool.execute(
            `
                UPDATE usuarios
                SET password_salt = ?, password_hash = ?, must_change_password = ?, activo = 1, updated_at = NOW()
                WHERE id = ?
            `,
            [passwordData.salt, passwordData.hash, mustChangePassword, userId]
        );

        await pool.execute('UPDATE sesiones SET revoked_at = NOW() WHERE user_id = ?', [userId]);

        const message =
            targetRole === ROLE_SUPER
                ? `Clave de cuenta SUPER actualizada para ${targetRows[0].username}.`
                : `Clave temporal restablecida para ${targetRows[0].username}. Debe cambiarla al ingresar.`;

        return res.json({
            message
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible restablecer la clave.' });
    }
});

app.delete('/api/admin/usuarios/:userId', requireAdmin, async (req, res) => {
    try {
        const userId = Number(req.params.userId);
        if (!Number.isInteger(userId) || userId <= 0) {
            return res.status(400).json({ message: 'Identificador de usuario invalido.' });
        }

        if (Number(userId) === Number(req.authUser?.id)) {
            return res.status(400).json({ message: 'No puedes eliminar tu propio usuario.' });
        }

        const [targetRows] = await pool.execute('SELECT id, username, role FROM usuarios WHERE id = ? LIMIT 1', [userId]);
        if (!targetRows.length) {
            return res.status(404).json({ message: 'No existe ese usuario.' });
        }

        if (isGuestUsername(targetRows[0].username)) {
            return res
                .status(400)
                .json({ message: 'La cuenta invitado no se elimina. Usa su checkbox para habilitarla o deshabilitarla.' });
        }

        if (normalizeText(targetRows[0].username).toLowerCase() === SUPERUSER_USERNAME) {
            return res.status(400).json({ message: 'La cuenta SUPER principal no se puede eliminar.' });
        }

        const targetRole = normalizeUserRole(targetRows[0].role, ROLE_OPERADOR);
        if (targetRole === ROLE_SUPER && !isSuperUser(req.authUser)) {
            return res.status(403).json({ message: 'Solo el superusuario puede eliminar una cuenta SUPER.' });
        }

        const targetIsActive = Number(targetRows[0].activo || 0) === 1;
        if (targetIsActive && isPrivilegedRole(targetRole)) {
            if ((await countActivePrivilegedUsers()) <= 1) {
                return res.status(400).json({ message: 'Debe existir al menos un usuario ADMIN o SUPER activo.' });
            }
        }

        await pool.execute('DELETE FROM usuarios WHERE id = ?', [userId]);
        return res.json({ message: `Usuario ${targetRows[0].username} eliminado.` });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible eliminar el usuario.' });
    }
});

app.get('/api/admin/correo-config', requireSuper, async (req, res) => {
    try {
        const activeConfigResult = await resolveActiveSmtpConfiguration();
        const storedRow = activeConfigResult.storedRow;
        const hasStoredConfig = Boolean(storedRow);
        const editableConfig = hasStoredConfig
            ? toSmtpPublicConfig(storedRow)
            : activeConfigResult.source === 'ENV'
              ? toSmtpPublicConfig(activeConfigResult.config || SMTP_ENV_CONFIG)
              : toSmtpPublicConfig({});

        const updatedAtRaw = storedRow?.updated_at || null;
        const updatedAtIso =
            updatedAtRaw instanceof Date
                ? updatedAtRaw.toISOString()
                : normalizeText(updatedAtRaw || '');

        return res.json({
            source: activeConfigResult.source,
            hasStoredConfig,
            isReady: Boolean(activeConfigResult.config),
            config: editableConfig,
            updatedAt: updatedAtIso,
            updatedBy: normalizeText(storedRow?.updated_by_nombre || '')
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible cargar la configuracion de correo.' });
    }
});

app.put('/api/admin/correo-config', requireSuper, async (req, res) => {
    try {
        const smtpHost = normalizeText(req.body?.smtpHost ?? req.body?.host);
        const smtpPort = parsePositiveInt(req.body?.smtpPort ?? req.body?.port, 0);
        const smtpSecure = parseBooleanFlag(req.body?.smtpSecure ?? req.body?.secure, false);
        const smtpUser = normalizeText(req.body?.smtpUser ?? req.body?.user);
        const smtpFromEmail = normalizeText(req.body?.smtpFromEmail ?? req.body?.fromEmail ?? smtpUser).toLowerCase();
        const smtpFromName = normalizeText(req.body?.smtpFromName ?? req.body?.fromName) || 'Geo Rural';
        const providedPassword = String(req.body?.smtpPassword ?? req.body?.password ?? '');
        const keepPassword = parseBooleanFlag(req.body?.keepPassword, true);

        if (!smtpHost) {
            return res.status(400).json({ message: 'Debes ingresar el servidor SMTP (host).' });
        }
        if (exceedsMaxLength(smtpHost, 255)) {
            return res.status(400).json({ message: 'El servidor SMTP no puede exceder 255 caracteres.' });
        }
        if (!smtpPort || smtpPort > 65535) {
            return res.status(400).json({ message: 'Debes ingresar un puerto SMTP valido (1-65535).' });
        }
        if (exceedsMaxLength(smtpUser, 255)) {
            return res.status(400).json({ message: 'El usuario SMTP no puede exceder 255 caracteres.' });
        }
        if (!smtpFromEmail) {
            return res.status(400).json({ message: 'Debes ingresar un correo remitente.' });
        }
        if (!isValidEmail(smtpFromEmail)) {
            return res.status(400).json({ message: 'El correo remitente no es valido.' });
        }
        if (exceedsMaxLength(smtpFromEmail, 255)) {
            return res.status(400).json({ message: 'El correo remitente no puede exceder 255 caracteres.' });
        }
        if (exceedsMaxLength(smtpFromName, 120)) {
            return res.status(400).json({ message: 'El nombre remitente no puede exceder 120 caracteres.' });
        }
        if (providedPassword && exceedsMaxLength(providedPassword, 255)) {
            return res.status(400).json({ message: 'La clave SMTP no puede exceder 255 caracteres.' });
        }

        const existingRow = await getStoredCorreoConfigurationRow();
        let smtpPassword = '';
        if (providedPassword) {
            smtpPassword = providedPassword;
        } else if (keepPassword) {
            smtpPassword = String(existingRow?.smtp_password || '');
        }

        const updatedById = Number(req.authUser?.id);
        const updatedByName = normalizeText(req.authUser?.nombre || req.authUser?.username || 'superusuario');

        await pool.execute(
            `
                INSERT INTO correo_configuracion (
                    id,
                    smtp_host,
                    smtp_port,
                    smtp_secure,
                    smtp_user,
                    smtp_password,
                    smtp_from_email,
                    smtp_from_name,
                    updated_by_id,
                    updated_by_nombre
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON DUPLICATE KEY UPDATE
                    smtp_host = VALUES(smtp_host),
                    smtp_port = VALUES(smtp_port),
                    smtp_secure = VALUES(smtp_secure),
                    smtp_user = VALUES(smtp_user),
                    smtp_password = VALUES(smtp_password),
                    smtp_from_email = VALUES(smtp_from_email),
                    smtp_from_name = VALUES(smtp_from_name),
                    updated_by_id = VALUES(updated_by_id),
                    updated_by_nombre = VALUES(updated_by_nombre),
                    updated_at = CURRENT_TIMESTAMP
            `,
            [
                SMTP_CONFIG_SINGLETON_ID,
                smtpHost,
                smtpPort,
                smtpSecure ? 1 : 0,
                smtpUser || null,
                smtpPassword || null,
                smtpFromEmail,
                smtpFromName,
                Number.isInteger(updatedById) && updatedById > 0 ? updatedById : null,
                updatedByName || null
            ]
        );

        resetSmtpTransporterCache();
        const savedRow = await getStoredCorreoConfigurationRow();
        const savedConfig = toSmtpRuntimeConfig(savedRow || {});
        const updatedAtRaw = savedRow?.updated_at || null;
        const updatedAtIso =
            updatedAtRaw instanceof Date
                ? updatedAtRaw.toISOString()
                : normalizeText(updatedAtRaw || '');

        return res.json({
            message: 'Configuracion de correo guardada correctamente.',
            source: 'DATABASE',
            isReady: isSmtpConfigReady(savedConfig),
            config: toSmtpPublicConfig(savedRow || savedConfig),
            updatedAt: updatedAtIso,
            updatedBy: normalizeText(savedRow?.updated_by_nombre || updatedByName)
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible guardar la configuracion de correo.' });
    }
});

app.get('/api/admin/correo-plantillas', requireSuper, async (req, res) => {
    try {
        const storedRow = await getStoredCorreoConfigurationRow();
        const defaults = getDefaultMailTemplates();
        const templates = resolveMailTemplatesForUse(storedRow);
        const updatedAtRaw = storedRow?.updated_at || null;
        const updatedAtIso =
            updatedAtRaw instanceof Date
                ? updatedAtRaw.toISOString()
                : normalizeText(updatedAtRaw || '');

        return res.json({
            templates,
            defaults,
            updatedAt: updatedAtIso,
            updatedBy: normalizeText(storedRow?.updated_by_nombre || '')
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible cargar las plantillas de correo.' });
    }
});

app.put('/api/admin/correo-plantillas', requireSuper, async (req, res) => {
    try {
        const rawTemplates = req.body?.templates ?? req.body;
        const templates = normalizeMailTemplatesPayload(rawTemplates);

        for (const key of MAIL_TEMPLATE_KEYS) {
            const rawValue = rawTemplates && typeof rawTemplates === 'object' ? rawTemplates[key] : '';
            if (rawValue && String(rawValue).length > MAIL_TEMPLATE_MAX_LENGTH) {
                return res.status(400).json({
                    message: `La plantilla ${key} excede el maximo permitido de ${MAIL_TEMPLATE_MAX_LENGTH} caracteres.`
                });
            }
        }

        const updatedById = Number(req.authUser?.id);
        const updatedByName = normalizeText(req.authUser?.nombre || req.authUser?.username || 'superusuario');
        const existingRow = await getStoredCorreoConfigurationRow();

        if (existingRow) {
            await pool.execute(
                `
                    UPDATE correo_configuracion
                    SET
                        cotizacion_html_template = ?,
                        factura_single_html_template = ?,
                        factura_pending_html_template = ?,
                        updated_by_id = ?,
                        updated_by_nombre = ?,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                `,
                [
                    templates.cotizacionHtml || null,
                    templates.facturaSingleHtml || null,
                    templates.facturaPendingHtml || null,
                    Number.isInteger(updatedById) && updatedById > 0 ? updatedById : null,
                    updatedByName || null,
                    SMTP_CONFIG_SINGLETON_ID
                ]
            );
        } else {
            await pool.execute(
                `
                    INSERT INTO correo_configuracion (
                        id,
                        cotizacion_html_template,
                        factura_single_html_template,
                        factura_pending_html_template,
                        updated_by_id,
                        updated_by_nombre
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                `,
                [
                    SMTP_CONFIG_SINGLETON_ID,
                    templates.cotizacionHtml || null,
                    templates.facturaSingleHtml || null,
                    templates.facturaPendingHtml || null,
                    Number.isInteger(updatedById) && updatedById > 0 ? updatedById : null,
                    updatedByName || null
                ]
            );
        }

        const savedRow = await getStoredCorreoConfigurationRow();
        const resolvedTemplates = resolveMailTemplatesForUse(savedRow);
        const updatedAtRaw = savedRow?.updated_at || null;
        const updatedAtIso =
            updatedAtRaw instanceof Date
                ? updatedAtRaw.toISOString()
                : normalizeText(updatedAtRaw || '');

        return res.json({
            message: 'Plantillas de correo guardadas correctamente.',
            templates: resolvedTemplates,
            defaults: getDefaultMailTemplates(),
            updatedAt: updatedAtIso,
            updatedBy: normalizeText(savedRow?.updated_by_nombre || updatedByName)
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible guardar las plantillas de correo.' });
    }
});

app.get('/api/admin/sesiones-activas', requireSuper, async (req, res) => {
    try {
        const requestedWindowMinutes = parsePositiveInt(req.query?.windowMinutes, ACTIVE_SESSION_WINDOW_MINUTES);
        const windowMinutes = Math.min(Math.max(requestedWindowMinutes, 5), 180);

        await cleanupExpiredSessionsIfNeeded();

        const [summaryRows] = await pool.execute(
            `
                SELECT
                    COUNT(*) AS sesiones_vigentes,
                    SUM(
                        CASE
                            WHEN COALESCE(s.last_used_at, s.created_at) >= DATE_SUB(NOW(), INTERVAL ? MINUTE) THEN 1
                            ELSE 0
                        END
                    ) AS equipos_activos,
                    COUNT(
                        DISTINCT CASE
                            WHEN COALESCE(s.last_used_at, s.created_at) >= DATE_SUB(NOW(), INTERVAL ? MINUTE) THEN s.user_id
                            ELSE NULL
                        END
                    ) AS usuarios_activos
                FROM sesiones s
                INNER JOIN usuarios u ON u.id = s.user_id
                WHERE s.revoked_at IS NULL
                  AND s.expires_at > NOW()
                  AND u.activo = 1
            `,
            [windowMinutes, windowMinutes]
        );

        const [activeRows] = await pool.execute(
            `
                SELECT
                    s.id AS session_id,
                    u.username,
                    u.nombre,
                    u.sucursal,
                    u.role,
                    COALESCE(s.last_used_at, s.created_at) AS activity_at,
                    s.expires_at
                FROM sesiones s
                INNER JOIN usuarios u ON u.id = s.user_id
                WHERE s.revoked_at IS NULL
                  AND s.expires_at > NOW()
                  AND u.activo = 1
                  AND COALESCE(s.last_used_at, s.created_at) >= DATE_SUB(NOW(), INTERVAL ? MINUTE)
                ORDER BY activity_at DESC, s.id DESC
                LIMIT 100
            `,
            [windowMinutes]
        );

        const summary = summaryRows[0] || {};
        const toIsoDateTime = (value) => (value instanceof Date ? value.toISOString() : normalizeText(value || ''));

        return res.json({
            windowMinutes,
            equiposActivos: Number(summary.equipos_activos || 0),
            usuariosActivos: Number(summary.usuarios_activos || 0),
            sesionesVigentes: Number(summary.sesiones_vigentes || 0),
            actualizadoEn: new Date().toISOString(),
            equipos: activeRows.map((row) => ({
                sessionId: Number(row.session_id || 0),
                username: normalizeText(row.username),
                nombre: normalizeText(row.nombre),
                sucursal: normalizeText(row.sucursal),
                role: normalizeUserRole(row.role, ROLE_OPERADOR),
                lastActivityAt: toIsoDateTime(row.activity_at),
                expiresAt: toIsoDateTime(row.expires_at)
            }))
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible obtener el estado de equipos activos.' });
    }
});

app.get('/api/admin/cotizaciones/documento', requireSuper, async (req, res) => {
    try {
        const absolutePath = getCotizacionTemplateFileAbsolutePath();
        if (!fs.existsSync(absolutePath)) {
            return res.status(404).json({ message: 'No se encontro el documento base de cotizacion.' });
        }

        const stats = fs.statSync(absolutePath);
        return res.json({
            fileName: path.basename(absolutePath),
            relativePath: COTIZACION_DOCUMENT_PATH,
            sizeBytes: Number(stats.size || 0),
            updatedAt: stats.mtime instanceof Date ? stats.mtime.toISOString() : ''
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible consultar el documento base de cotizacion.' });
    }
});

app.post('/api/admin/cotizaciones/documento', requireSuper, async (req, res) => {
    let tempPath = '';
    try {
        const maxFileBytes = 10 * 1024 * 1024;
        const rawFileName = normalizeText(req.body?.fileName);
        const rawFileBase64 = String(req.body?.fileBase64 || '')
            .trim()
            .replace(/^data:application\/pdf;base64,/i, '')
            .replace(/\s+/g, '');

        if (!rawFileBase64) {
            return res.status(400).json({ message: 'Debes seleccionar un archivo PDF para actualizar.' });
        }

        if (rawFileName && !/\.pdf$/i.test(rawFileName)) {
            return res.status(400).json({ message: 'Solo se permiten archivos PDF (.pdf).' });
        }

        const fileBuffer = Buffer.from(rawFileBase64, 'base64');
        if (!fileBuffer.length) {
            return res.status(400).json({ message: 'No fue posible leer el contenido del PDF seleccionado.' });
        }
        if (fileBuffer.length > maxFileBytes) {
            return res.status(413).json({ message: 'El archivo PDF excede el maximo permitido de 10 MB.' });
        }
        if (!isPdfBuffer(fileBuffer)) {
            return res.status(400).json({ message: 'El archivo seleccionado no corresponde a un PDF valido.' });
        }

        try {
            const pdfDoc = await PDFLibDocument.load(fileBuffer);
            if (pdfDoc.getPageCount() < 1) {
                return res.status(400).json({ message: 'El PDF seleccionado no contiene paginas.' });
            }
        } catch (pdfError) {
            return res.status(400).json({ message: 'No fue posible validar el archivo PDF seleccionado.' });
        }

        const targetPath = getCotizacionTemplateFileAbsolutePath();
        const targetDir = path.dirname(targetPath);
        if (!fs.existsSync(targetDir)) {
            fs.mkdirSync(targetDir, { recursive: true });
        }

        tempPath = `${targetPath}.tmp-${Date.now()}-${crypto.randomBytes(4).toString('hex')}`;
        fs.writeFileSync(tempPath, fileBuffer);
        fs.renameSync(tempPath, targetPath);
        tempPath = '';

        const stats = fs.statSync(targetPath);
        return res.json({
            message: `Documento base actualizado correctamente (${path.basename(targetPath)}).`,
            fileName: path.basename(targetPath),
            relativePath: COTIZACION_DOCUMENT_PATH,
            sizeBytes: Number(stats.size || 0),
            updatedAt: stats.mtime instanceof Date ? stats.mtime.toISOString() : ''
        });
    } catch (error) {
        if (tempPath && fs.existsSync(tempPath)) {
            try {
                fs.unlinkSync(tempPath);
            } catch (cleanupError) {
                // Sin accion: evitar romper respuesta principal.
            }
        }
        console.error(error);
        return res.status(500).json({ message: 'No fue posible actualizar el documento base de cotizacion.' });
    }
});

app.get('/api/next-ingreso', async (req, res) => {
    try {
        const year = new Date().getFullYear();
        const [rows] = await pool.execute(
            'SELECT COALESCE(MAX(correlativo), 0) AS maxCorrelativo FROM registros WHERE anio = ?',
            [year]
        );

        const next = Number(rows[0].maxCorrelativo || 0) + 1;
        res.json({ numIngreso: `${next}-${year}` });
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: 'No fue posible obtener el siguiente numero de ingreso.' });
    }
});

app.post('/api/registros', requireWriteAccess, async (req, res) => {
    const incomingComment = normalizeText(req.body.comentario);
    const rawRut = normalizeText(req.body.rut);
    const rawCorreo = normalizeText(req.body.correo).toLowerCase();
    const rawFactura = req.body && typeof req.body.factura === 'object' && req.body.factura ? req.body.factura : {};
    const payload = {
        numIngreso: normalizeText(req.body.numIngreso),
        nombre: normalizeText(req.body.nombre),
        rut: rawRut,
        telefono: normalizeText(req.body.telefono),
        correo: rawCorreo,
        region: normalizeText(req.body.region),
        comuna: normalizeText(req.body.comuna),
        nombrePredio: normalizeText(req.body.nombrePredio),
        rol: normalizeText(req.body.rol),
        numLotesRaw: req.body.numLotes,
        estado: normalizeEstado(req.body.estado),
        comentario: incomingComment || DEFAULT_CREATION_COMMENT,
        documentos: safeJsonArray(req.body.documentos),
        facturaNombreRazon: normalizeText(rawFactura.nombreRazonSocial),
        facturaNumero: normalizeText(rawFactura.numeroFactura || rawFactura.numero || rawFactura.numeroDeFactura),
        facturaRut: normalizeText(rawFactura.rut),
        facturaGiro: normalizeText(rawFactura.giro),
        facturaDireccion: normalizeText(rawFactura.direccion),
        facturaComuna: normalizeText(rawFactura.comuna),
        facturaCiudad: normalizeText(rawFactura.ciudad),
        facturaContacto: normalizeText(rawFactura.contacto),
        facturaObservacion: normalizeText(rawFactura.observacion),
        facturaMontoRaw: rawFactura.montoFacturar
    };

    const createLengthError = validateRegistroPayloadLengths(payload);
    if (createLengthError) {
        return res.status(400).json({ message: createLengthError });
    }

    const parsedFromClient = payload.numIngreso ? parseNumIngreso(payload.numIngreso) : null;
    if (payload.numIngreso && !parsedFromClient) {
        return res.status(400).json({ message: 'NRO INGRESO invalido. Usa formato 123-2026.' });
    }

    if (!payload.nombre || !payload.rut || !payload.region || !payload.comuna || !payload.rol) {
        return res.status(400).json({ message: 'Faltan campos obligatorios para guardar.' });
    }

    const formattedRut = normalizeAndFormatRut(payload.rut);
    if (!formattedRut) {
        return res.status(400).json({ message: 'RUT invalido. Usa formato 12.345.678-5.' });
    }
    payload.rut = formattedRut;

    if (payload.correo && !isValidEmail(payload.correo)) {
        return res.status(400).json({ message: 'Correo invalido. Usa formato nombre@proveedor.cl.' });
    }

    const parsedLotes = parseNumLotes(payload.numLotesRaw);
    if (!parsedLotes.isValid) {
        return res.status(400).json({ message: 'NRO DE LOTES debe ser un numero entero positivo.' });
    }

    const parsedFacturaMonto = parseFacturaMonto(payload.facturaMontoRaw);
    if (!parsedFacturaMonto.isValid) {
        return res.status(400).json({ message: 'MONTO A FACTURAR invalido.' });
    }
    payload.facturaMonto = parsedFacturaMonto.value;

    const authRole = normalizeUserRole(req.authUser?.role, ROLE_OPERADOR);
    if (authRole === ROLE_SUPERVISOR) {
        return res.status(403).json({
            message: 'Tu rol solo puede buscar registros y agregar comentarios.'
        });
    }
    if (authRole === ROLE_OPERADOR) {
        return res.status(403).json({
            message: 'Tu rol no puede crear registros. Solo puedes agregar comentarios y actualizar documentacion recibida.'
        });
    }

    const canEditFacturaNumero = isSecretaryUser(req.authUser) && !isGuestUser(req.authUser);
    if (!canEditFacturaNumero || payload.estado !== 'facturada') {
        payload.facturaNumero = '';
    }

    const userHasBackofficeAccess = hasBackofficeRegistroAccess(req.authUser);
    const userBranch = normalizeText(req.authUser.sucursal) || null;
    if (!userHasBackofficeAccess && !userBranch) {
        return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para operar registros.' });
    }

    const useManualIngreso = userHasBackofficeAccess && Boolean(parsedFromClient);
    const connection = await pool.getConnection();
    const userName = req.authUser.nombre || req.authUser.username;

    try {
        await connection.beginTransaction();
        const reservedIngreso = useManualIngreso
            ? {
                  numIngreso: payload.numIngreso,
                  correlativo: parsedFromClient.correlativo,
                  anio: parsedFromClient.anio
              }
            : await reserveNextIngreso(connection, new Date().getFullYear());

        const [insertResult] = await connection.execute(
            `
                INSERT INTO registros (
                    num_ingreso, correlativo, anio,
                    nombre, rut, telefono, correo,
                    region, comuna, nombre_predio,
                    rol, num_lotes, estado, comentario,
                    factura_nombre_razon, factura_numero, factura_rut, factura_giro,
                    factura_direccion, factura_comuna, factura_ciudad,
                    factura_contacto, factura_observacion, factura_monto,
                    created_by, created_by_sucursal,
                    updated_by, updated_by_sucursal,
                    documentos
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            `,
            [
                reservedIngreso.numIngreso,
                reservedIngreso.correlativo,
                reservedIngreso.anio,
                payload.nombre,
                payload.rut,
                payload.telefono || null,
                payload.correo || null,
                payload.region,
                payload.comuna,
                payload.nombrePredio || null,
                payload.rol,
                parsedLotes.value,
                payload.estado || null,
                payload.comentario || null,
                payload.facturaNombreRazon || null,
                payload.facturaNumero || null,
                payload.facturaRut || null,
                payload.facturaGiro || null,
                payload.facturaDireccion || null,
                payload.facturaComuna || null,
                payload.facturaCiudad || null,
                payload.facturaContacto || null,
                payload.facturaObservacion || null,
                payload.facturaMonto,
                userName,
                userBranch,
                userName,
                userBranch,
                JSON.stringify(payload.documentos)
            ]
        );

        const historyComment = buildHistoryComment('CREACION', payload.comentario, payload.documentos, []);
        await insertHistoryEntry(connection, {
            registroId: insertResult.insertId,
            numIngreso: reservedIngreso.numIngreso,
            accion: 'CREACION',
            comentario: historyComment,
            usuarioNombre: userName,
            sucursal: userBranch
        });

        await connection.commit();
        return res.status(201).json({
            message: `Registro ${reservedIngreso.numIngreso} guardado.`,
            numIngreso: reservedIngreso.numIngreso
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);

        if (error && error.code === 'ER_DUP_ENTRY') {
            return res.status(409).json({ message: 'El NRO INGRESO ya existe en la base de datos.' });
        }

        return res.status(500).json({ message: 'No fue posible guardar el registro.' });
    } finally {
        connection.release();
    }
});

app.put('/api/registros/:numIngreso', requireWriteAccess, async (req, res) => {
    const routeIngreso = normalizeText(decodeURIComponent(req.params.numIngreso));
    const parsedIngreso = parseNumIngreso(routeIngreso);

    if (!parsedIngreso) {
        return res.status(400).json({ message: 'NRO INGRESO invalido.' });
    }

    const rawRut = normalizeText(req.body.rut);
    const rawCorreo = normalizeText(req.body.correo).toLowerCase();
    const hasFacturaPayload = Boolean(req.body && typeof req.body.factura === 'object' && req.body.factura);
    const rawFactura = hasFacturaPayload ? req.body.factura : {};
    const payload = {
        numIngresoNuevo: normalizeText(req.body.numIngresoNuevo || req.body.numIngreso || routeIngreso),
        nombre: normalizeText(req.body.nombre),
        rut: rawRut,
        telefono: normalizeText(req.body.telefono),
        correo: rawCorreo,
        region: normalizeText(req.body.region),
        comuna: normalizeText(req.body.comuna),
        nombrePredio: normalizeText(req.body.nombrePredio),
        rol: normalizeText(req.body.rol),
        numLotesRaw: req.body.numLotes,
        estado: normalizeEstado(req.body.estado),
        comentario: normalizeText(req.body.comentario),
        documentos: safeJsonArray(req.body.documentos),
        facturaNombreRazon: normalizeText(rawFactura.nombreRazonSocial),
        facturaNumero: normalizeText(rawFactura.numeroFactura || rawFactura.numero || rawFactura.numeroDeFactura),
        facturaRut: normalizeText(rawFactura.rut),
        facturaGiro: normalizeText(rawFactura.giro),
        facturaDireccion: normalizeText(rawFactura.direccion),
        facturaComuna: normalizeText(rawFactura.comuna),
        facturaCiudad: normalizeText(rawFactura.ciudad),
        facturaContacto: normalizeText(rawFactura.contacto),
        facturaObservacion: normalizeText(rawFactura.observacion),
        facturaMontoRaw: rawFactura.montoFacturar
    };

    const updateLengthError = validateRegistroPayloadLengths(payload);
    if (updateLengthError) {
        return res.status(400).json({ message: updateLengthError });
    }

    const authRole = normalizeUserRole(req.authUser?.role, ROLE_OPERADOR);
    if (authRole === ROLE_SUPERVISOR) {
        return res.status(403).json({ message: 'Tu rol solo puede buscar registros y agregar comentarios.' });
    }
    const isOperatorLimitedEditor = authRole === ROLE_OPERADOR && !isGuestUser(req.authUser);
    if (isOperatorLimitedEditor) {
        // En OPERADOR, MODIFICAR solo aplica cambios de documentacion.
        payload.numIngresoNuevo = routeIngreso;
    }

    if (!isOperatorLimitedEditor) {
        if (!payload.nombre || !payload.rut || !payload.region || !payload.comuna || !payload.rol) {
            return res.status(400).json({ message: 'Faltan campos obligatorios para modificar.' });
        }

        const formattedRut = normalizeAndFormatRut(payload.rut);
        if (!formattedRut) {
            return res.status(400).json({ message: 'RUT invalido. Usa formato 12.345.678-5.' });
        }
        payload.rut = formattedRut;

        if (payload.correo && !isValidEmail(payload.correo)) {
            return res.status(400).json({ message: 'Correo invalido. Usa formato nombre@proveedor.cl.' });
        }
    }

    let parsedIngresoNuevo = parsedIngreso;
    if (!isOperatorLimitedEditor) {
        parsedIngresoNuevo = parseNumIngreso(payload.numIngresoNuevo);
        if (!parsedIngresoNuevo) {
            return res.status(400).json({ message: 'NRO INGRESO nuevo invalido. Usa formato 123-2026.' });
        }
    }

    const userHasBackofficeAccess = hasBackofficeRegistroAccess(req.authUser);
    if (!userHasBackofficeAccess && payload.numIngresoNuevo !== routeIngreso) {
        return res.status(403).json({ message: 'Solo ADMIN o SECRETARIA pueden cambiar el NRO INGRESO.' });
    }

    const userBranch = normalizeText(req.authUser.sucursal) || null;
    if (!userHasBackofficeAccess && !userBranch) {
        return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para operar registros.' });
    }

    const registroScope = getRegistroAccessScope(req.authUser);
    if (registroScope.denied) {
        return res.status(403).json({ message: 'No tienes permisos para operar registros sin sucursal asignada.' });
    }

    let parsedLotes = parseNumLotes(payload.numLotesRaw);
    if (isOperatorLimitedEditor) {
        parsedLotes = { isValid: true, value: null };
    }
    if (!parsedLotes.isValid) {
        return res.status(400).json({ message: 'NRO DE LOTES debe ser un numero entero positivo.' });
    }

    let parsedFacturaMonto = parseFacturaMonto(payload.facturaMontoRaw);
    if (isOperatorLimitedEditor) {
        parsedFacturaMonto = { isValid: true, value: null };
    }
    if (!parsedFacturaMonto.isValid) {
        return res.status(400).json({ message: 'MONTO A FACTURAR invalido.' });
    }
    payload.facturaMonto = parsedFacturaMonto.value;

    const connection = await pool.getConnection();
    const userName = req.authUser.nombre || req.authUser.username;

    try {
        await connection.beginTransaction();

        const recordScopeWhere = registroScope.clause ? ` AND ${registroScope.clause}` : '';
        const [existingRows] = await connection.execute(
            `
                SELECT
                    id,
                    num_ingreso,
                    nombre,
                    rut,
                    telefono,
                    correo,
                    region,
                    comuna,
                    nombre_predio,
                    rol,
                    num_lotes,
                    estado,
                    comentario,
                    factura_nombre_razon,
                    factura_numero,
                    factura_rut,
                    factura_giro,
                    factura_direccion,
                    factura_comuna,
                    factura_ciudad,
                    factura_contacto,
                    factura_observacion,
                    factura_monto,
                    documentos
                FROM registros
                WHERE num_ingreso = ?${recordScopeWhere}
                FOR UPDATE
            `,
            [routeIngreso, ...registroScope.params]
        );

        if (!existingRows.length) {
            await connection.rollback();
            return res.status(404).json({ message: 'No existe ese registro para modificar.' });
        }

        if (isOperatorLimitedEditor) {
            payload.numIngresoNuevo = routeIngreso;
            payload.nombre = normalizeText(existingRows[0].nombre);
            payload.rut = normalizeText(existingRows[0].rut);
            payload.telefono = normalizeText(existingRows[0].telefono);
            payload.correo = normalizeText(existingRows[0].correo).toLowerCase();
            payload.region = normalizeText(existingRows[0].region);
            payload.comuna = normalizeText(existingRows[0].comuna);
            payload.nombrePredio = normalizeText(existingRows[0].nombre_predio);
            payload.rol = normalizeText(existingRows[0].rol);
            payload.estado = normalizeEstado(existingRows[0].estado);
            payload.facturaNombreRazon = normalizeText(existingRows[0].factura_nombre_razon);
            payload.facturaNumero = normalizeText(existingRows[0].factura_numero);
            payload.facturaRut = normalizeText(existingRows[0].factura_rut);
            payload.facturaGiro = normalizeText(existingRows[0].factura_giro);
            payload.facturaDireccion = normalizeText(existingRows[0].factura_direccion);
            payload.facturaComuna = normalizeText(existingRows[0].factura_comuna);
            payload.facturaCiudad = normalizeText(existingRows[0].factura_ciudad);
            payload.facturaContacto = normalizeText(existingRows[0].factura_contacto);
            payload.facturaObservacion = normalizeText(existingRows[0].factura_observacion);
            payload.facturaMonto = existingRows[0].factura_monto === null ? null : Number(existingRows[0].factura_monto);
            parsedLotes.value = existingRows[0].num_lotes === null ? null : Number(existingRows[0].num_lotes);
        }

        if (!hasFacturaPayload) {
            payload.facturaNombreRazon = normalizeText(existingRows[0].factura_nombre_razon);
            payload.facturaNumero = normalizeText(existingRows[0].factura_numero);
            payload.facturaRut = normalizeText(existingRows[0].factura_rut);
            payload.facturaGiro = normalizeText(existingRows[0].factura_giro);
            payload.facturaDireccion = normalizeText(existingRows[0].factura_direccion);
            payload.facturaComuna = normalizeText(existingRows[0].factura_comuna);
            payload.facturaCiudad = normalizeText(existingRows[0].factura_ciudad);
            payload.facturaContacto = normalizeText(existingRows[0].factura_contacto);
            payload.facturaObservacion = normalizeText(existingRows[0].factura_observacion);
            payload.facturaMonto = existingRows[0].factura_monto === null ? null : Number(existingRows[0].factura_monto);
        }

        const canEditFacturaNumero = isSecretaryUser(req.authUser) && !isGuestUser(req.authUser);
        if (!canEditFacturaNumero) {
            payload.facturaNumero = normalizeText(existingRows[0].factura_numero);
        }
        if (payload.estado !== 'facturada') {
            payload.facturaNumero = '';
        }

        const docsBefore = parseStoredDocumentList(existingRows[0].documentos);
        const docsAfter = payload.documentos;
        const docsAdded = docsAfter.filter((doc) => !docsBefore.includes(doc));
        const docsRemoved = docsBefore.filter((doc) => !docsAfter.includes(doc));
        const existingComment = normalizeText(existingRows[0].comentario);
        // El comentario se modifica solo via /comentario para evitar mezclas con MODIFICAR.
        payload.comentario = existingComment;
        const autoGeneratedComment = buildAutomaticModificationComment(existingRows[0], payload, docsAdded, docsRemoved, parsedLotes.value);
        const hasStructuredChanges = autoGeneratedComment !== 'modificacion sin cambios detectados.';
        if (!hasStructuredChanges) {
            await connection.rollback();
            return res.json({
                message: 'No se detectaron cambios. No se registro modificacion.',
                numIngreso: routeIngreso,
                noChanges: true
            });
        }

        await connection.execute(
            `
                UPDATE registros
                SET
                    num_ingreso = ?,
                    correlativo = ?,
                    anio = ?,
                    nombre = ?,
                    rut = ?,
                    telefono = ?,
                    correo = ?,
                    region = ?,
                    comuna = ?,
                    nombre_predio = ?,
                    rol = ?,
                    num_lotes = ?,
                    estado = ?,
                    comentario = ?,
                    factura_nombre_razon = ?,
                    factura_numero = ?,
                    factura_rut = ?,
                    factura_giro = ?,
                    factura_direccion = ?,
                    factura_comuna = ?,
                    factura_ciudad = ?,
                    factura_contacto = ?,
                    factura_observacion = ?,
                    factura_monto = ?,
                    updated_by = ?,
                    updated_by_sucursal = ?,
                    documentos = ?
                WHERE num_ingreso = ?
            `,
            [
                payload.numIngresoNuevo,
                parsedIngresoNuevo.correlativo,
                parsedIngresoNuevo.anio,
                payload.nombre,
                payload.rut,
                payload.telefono || null,
                payload.correo || null,
                payload.region,
                payload.comuna,
                payload.nombrePredio || null,
                payload.rol,
                parsedLotes.value,
                payload.estado || null,
                payload.comentario || null,
                payload.facturaNombreRazon || null,
                payload.facturaNumero || null,
                payload.facturaRut || null,
                payload.facturaGiro || null,
                payload.facturaDireccion || null,
                payload.facturaComuna || null,
                payload.facturaCiudad || null,
                payload.facturaContacto || null,
                payload.facturaObservacion || null,
                payload.facturaMonto,
                userName,
                userBranch,
                JSON.stringify(docsAfter),
                routeIngreso
            ]
        );

        const historyComment = autoGeneratedComment || 'Registro modificado.';
        await insertHistoryEntry(connection, {
            registroId: existingRows[0].id,
            numIngreso: payload.numIngresoNuevo,
            accion: 'MODIFICACION',
            comentario: historyComment,
            usuarioNombre: userName,
            sucursal: userBranch
        });

        if (docsAdded.length > 0) {
            await connection.execute(
                `
                    INSERT INTO registro_documentos_alertas (
                        registro_id,
                        num_ingreso,
                        sucursal,
                        documentos_agregados,
                        created_by
                    )
                    VALUES (?, ?, ?, ?, ?)
                `,
                [existingRows[0].id, payload.numIngresoNuevo, userBranch, JSON.stringify(docsAdded), userName]
            );
        }

        await connection.commit();
        return res.json({ message: `Registro ${payload.numIngresoNuevo} modificado.`, numIngreso: payload.numIngresoNuevo });
    } catch (error) {
        await connection.rollback();
        console.error(error);

        if (error && error.code === 'ER_DUP_ENTRY') {
            return res.status(409).json({ message: 'El NRO INGRESO nuevo ya existe en la base de datos.' });
        }

        return res.status(500).json({ message: 'No fue posible modificar el registro.' });
    } finally {
        connection.release();
    }
});

app.post('/api/registros/:numIngreso/comentario', requireWriteAccess, async (req, res) => {
    const routeIngreso = normalizeText(decodeURIComponent(req.params.numIngreso));
    const parsedIngreso = parseNumIngreso(routeIngreso);
    if (!parsedIngreso) {
        return res.status(400).json({ message: 'NRO INGRESO invalido.' });
    }

    const comentario = normalizeText(req.body.comentario);
    if (!comentario) {
        return res.status(400).json({ message: 'Debes ingresar un comentario para guardar.' });
    }
    if (exceedsMaxLength(comentario, 4000)) {
        return res.status(400).json({ message: 'COMENTARIO excede el maximo permitido de 4000 caracteres.' });
    }

    const userHasBackofficeAccess = hasBackofficeRegistroAccess(req.authUser);
    const userBranch = normalizeText(req.authUser.sucursal) || null;
    if (!userHasBackofficeAccess && !userBranch) {
        return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para operar registros.' });
    }

    const registroScope = getRegistroAccessScope(req.authUser);
    if (registroScope.denied) {
        return res.status(403).json({ message: 'No tienes permisos para operar registros sin sucursal asignada.' });
    }

    const connection = await pool.getConnection();
    const userName = req.authUser.nombre || req.authUser.username;

    try {
        await connection.beginTransaction();

        const recordScopeWhere = registroScope.clause ? ` AND ${registroScope.clause}` : '';
        const [existingRows] = await connection.execute(
            `
                SELECT id, num_ingreso
                FROM registros
                WHERE num_ingreso = ?${recordScopeWhere}
                FOR UPDATE
            `,
            [routeIngreso, ...registroScope.params]
        );

        if (!existingRows.length) {
            await connection.rollback();
            return res.status(404).json({ message: 'No existe ese registro para guardar comentario.' });
        }

        await connection.execute(
            `
                UPDATE registros
                SET
                    comentario = ?,
                    updated_by = ?,
                    updated_by_sucursal = ?,
                    updated_at = NOW()
                WHERE id = ?
            `,
            [comentario, userName, userBranch, existingRows[0].id]
        );

        await insertHistoryEntry(connection, {
            registroId: existingRows[0].id,
            numIngreso: existingRows[0].num_ingreso,
            accion: 'COMENTARIO',
            comentario,
            usuarioNombre: userName,
            sucursal: userBranch
        });

        await connection.commit();
        return res.json({
            message: 'Comentario guardado correctamente.',
            numIngreso: existingRows[0].num_ingreso
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        return res.status(500).json({ message: 'No fue posible guardar el comentario.' });
    } finally {
        connection.release();
    }
});

app.put('/api/registros/:numIngreso/historial/:historyId', requireWriteAccess, requireAdmin, async (req, res) => {
    const routeIngreso = normalizeText(decodeURIComponent(req.params.numIngreso));
    const parsedIngreso = parseNumIngreso(routeIngreso);
    if (!parsedIngreso) {
        return res.status(400).json({ message: 'NRO INGRESO invalido.' });
    }

    const historyIdRaw = Number(req.params.historyId);
    const historyId = Number.isSafeInteger(historyIdRaw) ? historyIdRaw : NaN;
    if (!Number.isInteger(historyId) || historyId <= 0) {
        return res.status(400).json({ message: 'Identificador de comentario invalido.' });
    }

    const comentario = normalizeText(req.body.comentario);
    if (!comentario) {
        return res.status(400).json({ message: 'Debes ingresar un comentario para guardar.' });
    }
    if (exceedsMaxLength(comentario, 4000)) {
        return res.status(400).json({ message: 'COMENTARIO excede el maximo permitido de 4000 caracteres.' });
    }

    const fechaSql = parseHistoryDateTimeInput(req.body.fecha);
    if (!fechaSql) {
        return res.status(400).json({ message: 'Fecha de comentario invalida. Usa formato YYYY-MM-DDTHH:mm.' });
    }
    const parsedEditedDate = new Date(String(req.body.fecha || '').replace(' ', 'T'));
    if (Number.isNaN(parsedEditedDate.getTime())) {
        return res.status(400).json({ message: 'Fecha de comentario invalida.' });
    }
    if (parsedEditedDate.getFullYear() < 2000) {
        return res.status(400).json({ message: 'La fecha del comentario no puede ser anterior al ano 2000.' });
    }
    if (parsedEditedDate.getTime() > Date.now() + 5 * 60 * 1000) {
        return res.status(400).json({ message: 'La fecha del comentario no puede estar en el futuro.' });
    }

    const registroScope = getRegistroAccessScope(req.authUser);
    if (registroScope.denied) {
        return res.status(403).json({ message: 'No tienes permisos para operar registros sin sucursal asignada.' });
    }

    const connection = await pool.getConnection();
    const adminName = req.authUser.nombre || req.authUser.username;
    const adminBranch = normalizeText(req.authUser.sucursal) || null;

    try {
        await connection.beginTransaction();

        const recordScopeWhere = registroScope.clause ? ` AND ${registroScope.clause}` : '';
        const [recordRows] = await connection.execute(
            `
                SELECT id, num_ingreso
                FROM registros
                WHERE num_ingreso = ?${recordScopeWhere}
                FOR UPDATE
            `,
            [routeIngreso, ...registroScope.params]
        );
        if (!recordRows.length) {
            await connection.rollback();
            return res.status(404).json({ message: 'No existe ese registro para editar comentarios.' });
        }

        const registroId = Number(recordRows[0].id);
        const [historyRows] = await connection.execute(
            `
                SELECT id
                FROM registro_historial
                WHERE id = ? AND registro_id = ?
                FOR UPDATE
            `,
            [historyId, registroId]
        );
        if (!historyRows.length) {
            await connection.rollback();
            return res.status(404).json({ message: 'No existe ese comentario en el historial de progreso.' });
        }

        await connection.execute(
            `
                UPDATE registro_historial
                SET
                    comentario = ?,
                    created_at = ?,
                    editado = 0,
                    editado_por = NULL,
                    editado_por_sucursal = NULL,
                    editado_at = NULL
                WHERE id = ?
            `,
            [comentario, fechaSql, historyId]
        );

        const [latestCommentRows] = await connection.execute(
            `
                SELECT comentario
                FROM registro_historial
                WHERE registro_id = ? AND accion = 'COMENTARIO'
                ORDER BY created_at DESC, id DESC
                LIMIT 1
            `,
            [registroId]
        );
        if (latestCommentRows.length > 0) {
            await connection.execute(
                `
                    UPDATE registros
                    SET
                        comentario = ?,
                        updated_by = ?,
                        updated_by_sucursal = ?,
                        updated_at = NOW()
                    WHERE id = ?
                `,
                [normalizeText(latestCommentRows[0].comentario) || null, adminName, adminBranch, registroId]
            );
        }

        const [updatedRows] = await connection.execute(
            `
                SELECT
                    id,
                    num_ingreso,
                    accion,
                    created_at,
                    comentario,
                    usuario_nombre,
                    sucursal,
                    editado,
                    editado_por,
                    editado_por_sucursal,
                    editado_at
                FROM registro_historial
                WHERE id = ?
                LIMIT 1
            `,
            [historyId]
        );

        await connection.commit();
        return res.json({
            message: 'Comentario de historial de progreso actualizado correctamente.',
            historial: updatedRows.length > 0 ? toHistoryRow(updatedRows[0]) : null
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        return res.status(500).json({ message: 'No fue posible actualizar el comentario del historial de progreso.' });
    } finally {
        connection.release();
    }
});

app.delete('/api/registros/:numIngreso/historial/:historyId', requireWriteAccess, requireSuper, async (req, res) => {
    const routeIngreso = normalizeText(decodeURIComponent(req.params.numIngreso));
    const parsedIngreso = parseNumIngreso(routeIngreso);
    if (!parsedIngreso) {
        return res.status(400).json({ message: 'NRO INGRESO invalido.' });
    }

    const historyIdRaw = Number(req.params.historyId);
    const historyId = Number.isSafeInteger(historyIdRaw) ? historyIdRaw : NaN;
    if (!Number.isInteger(historyId) || historyId <= 0) {
        return res.status(400).json({ message: 'Identificador de comentario invalido.' });
    }

    const registroScope = getRegistroAccessScope(req.authUser);
    if (registroScope.denied) {
        return res.status(403).json({ message: 'No tienes permisos para operar registros sin sucursal asignada.' });
    }

    const connection = await pool.getConnection();
    const superName = req.authUser.nombre || req.authUser.username;
    const superBranch = normalizeText(req.authUser.sucursal) || null;

    try {
        await connection.beginTransaction();

        const recordScopeWhere = registroScope.clause ? ` AND ${registroScope.clause}` : '';
        const [recordRows] = await connection.execute(
            `
                SELECT id, num_ingreso
                FROM registros
                WHERE num_ingreso = ?${recordScopeWhere}
                FOR UPDATE
            `,
            [routeIngreso, ...registroScope.params]
        );
        if (!recordRows.length) {
            await connection.rollback();
            return res.status(404).json({ message: 'No existe ese registro para eliminar comentarios.' });
        }

        const registroId = Number(recordRows[0].id);
        const [historyRows] = await connection.execute(
            `
                SELECT id, accion
                FROM registro_historial
                WHERE id = ? AND registro_id = ?
                FOR UPDATE
            `,
            [historyId, registroId]
        );
        if (!historyRows.length) {
            await connection.rollback();
            return res.status(404).json({ message: 'No existe ese comentario en el historial de progreso.' });
        }

        await connection.execute('DELETE FROM registro_historial WHERE id = ? LIMIT 1', [historyId]);

        const [latestCommentRows] = await connection.execute(
            `
                SELECT comentario
                FROM registro_historial
                WHERE registro_id = ? AND accion = 'COMENTARIO'
                ORDER BY created_at DESC, id DESC
                LIMIT 1
            `,
            [registroId]
        );
        const updatedComment = latestCommentRows.length > 0 ? normalizeText(latestCommentRows[0].comentario) : '';
        await connection.execute(
            `
                UPDATE registros
                SET
                    comentario = ?,
                    updated_by = ?,
                    updated_by_sucursal = ?,
                    updated_at = NOW()
                WHERE id = ?
            `,
            [updatedComment || null, superName, superBranch, registroId]
        );

        await connection.commit();
        return res.json({
            message: 'Comentario eliminado correctamente del historial de progreso.'
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        return res.status(500).json({ message: 'No fue posible eliminar el comentario del historial de progreso.' });
    } finally {
        connection.release();
    }
});

app.delete('/api/registros/:numIngreso', requireWriteAccess, requireAdmin, async (req, res) => {
    const routeIngreso = normalizeText(decodeURIComponent(req.params.numIngreso));
    const parsedIngreso = parseNumIngreso(routeIngreso);

    if (!parsedIngreso) {
        return res.status(400).json({ message: 'NRO INGRESO invalido.' });
    }

    const connection = await pool.getConnection();

    try {
        await connection.beginTransaction();

        const [existingRows] = await connection.execute(
            'SELECT id FROM registros WHERE num_ingreso = ? FOR UPDATE',
            [routeIngreso]
        );

        if (!existingRows.length) {
            await connection.rollback();
            return res.status(404).json({ message: 'No existe ese registro para eliminar.' });
        }

        await connection.execute('DELETE FROM registro_historial WHERE registro_id = ?', [existingRows[0].id]);
        await connection.execute('DELETE FROM registros WHERE id = ?', [existingRows[0].id]);

        await connection.commit();
        return res.json({ message: `Registro ${routeIngreso} eliminado.` });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        return res.status(500).json({ message: 'No fue posible eliminar el registro.' });
    } finally {
        connection.release();
    }
});

app.get('/api/registros/sucursales-disponibles', async (req, res) => {
    try {
        const registroScope = getRegistroAccessScope(req.authUser);
        if (registroScope.denied) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para buscar registros.' });
        }

        const [catalogRows] = await pool.execute(
            `
                SELECT nombre
                FROM sucursales
                ORDER BY nombre ASC
            `
        );
        const [registroRows] = await pool.execute(
            `
                SELECT DISTINCT sucursal
                FROM (
                    SELECT created_by_sucursal AS sucursal FROM registros
                    UNION
                    SELECT updated_by_sucursal AS sucursal FROM registros
                ) AS registro_sucursales
                WHERE sucursal IS NOT NULL AND TRIM(sucursal) <> ''
            `
        );

        const sucursales = Array.from(
            new Set(
                [...catalogRows, ...registroRows]
                    .map((row) => normalizeText(row.nombre || row.sucursal))
                    .filter((value) => value.length > 0)
            )
        ).sort((left, right) => left.localeCompare(right, 'es'));

        return res.json({ sucursales });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible listar sucursales disponibles.' });
    }
});

app.get('/api/registros/seguimiento-sin-movimiento', async (req, res) => {
    try {
        if (!canUseSecretaryFeatures(req.authUser)) {
            return res.status(403).json({ message: 'Solo SECRETARIA o SUPER pueden consultar este seguimiento.' });
        }

        const registroScope = getRegistroAccessScope(req.authUser, 'r');
        if (registroScope.denied) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para consultar este seguimiento.' });
        }

        const requestedDays = parsePositiveInt(req.query.dias, SECRETARIA_NO_MOVIMIENTO_DAYS);
        const thresholdDays = Math.max(1, Math.min(requestedDays, 365));

        const where = ['hstats.total_movimientos = 1', 'hstats.primer_movimiento_at <= DATE_SUB(NOW(), INTERVAL ? DAY)'];
        const values = [thresholdDays];
        if (registroScope.clause) {
            where.push(registroScope.clause);
            values.push(...registroScope.params);
        }

        const [rows] = await pool.execute(
            `
                SELECT
                    r.num_ingreso,
                    r.nombre,
                    r.rut,
                    r.region,
                    r.comuna,
                    COALESCE(r.updated_by_sucursal, r.created_by_sucursal, '') AS sucursal,
                    hstats.primer_movimiento_at AS created_at,
                    hstats.total_movimientos,
                    TIMESTAMPDIFF(DAY, hstats.primer_movimiento_at, NOW()) AS dias_sin_movimiento
                FROM registros r
                INNER JOIN (
                    SELECT
                        registro_id,
                        COUNT(*) AS total_movimientos,
                        MIN(created_at) AS primer_movimiento_at
                    FROM registro_historial
                    GROUP BY registro_id
                ) hstats
                    ON hstats.registro_id = r.id
                WHERE ${where.join(' AND ')}
                ORDER BY dias_sin_movimiento DESC, hstats.primer_movimiento_at ASC, r.id ASC
                LIMIT 500
            `,
            values
        );

        const registros = rows.map((row) => ({
            numIngreso: row.num_ingreso,
            nombre: row.nombre || '',
            rut: row.rut || '',
            region: row.region || '',
            comuna: row.comuna || '',
            sucursal: normalizeText(row.sucursal),
            createdAt: row.created_at instanceof Date ? row.created_at.toISOString() : String(row.created_at || ''),
            totalMovimientos: Number(row.total_movimientos || 0),
            diasSinMovimiento: Number(row.dias_sin_movimiento || 0)
        }));

        return res.json({
            thresholdDays,
            registros
        });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible cargar el seguimiento de registros sin movimiento.' });
    }
});

app.get('/api/registros/documentos-alertas', async (req, res) => {
    try {
        if (!canUseOperatorDocumentAlerts(req.authUser)) {
            return res.status(403).json({ message: 'Solo OPERADOR o SUPER pueden consultar esta lista.' });
        }

        const alertScope = getOperatorDocumentAlertsScope(req.authUser, 'a');
        if (alertScope.denied) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para consultar esta lista.' });
        }

        const where = ['a.revisado = 0'];
        const values = [];
        if (alertScope.clause) {
            where.push(alertScope.clause);
            values.push(...alertScope.params);
        }

        const [rows] = await pool.execute(
            `
                SELECT
                    a.id,
                    a.num_ingreso,
                    a.sucursal,
                    a.documentos_agregados,
                    a.created_by,
                    a.created_at
                FROM registro_documentos_alertas a
                WHERE ${where.join(' AND ')}
                ORDER BY a.created_at DESC, a.id DESC
                LIMIT 500
            `,
            values
        );

        const alertas = rows.map((row) => ({
            id: Number(row.id),
            numIngreso: normalizeText(row.num_ingreso),
            sucursal: normalizeText(row.sucursal),
            createdBy: normalizeText(row.created_by),
            createdAt: row.created_at instanceof Date ? row.created_at.toISOString() : String(row.created_at || ''),
            documentosAgregados: parseStoredDocumentList(row.documentos_agregados)
        }));

        return res.json({ alertas });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible cargar alertas de documentos nuevos.' });
    }
});

app.post('/api/registros/documentos-alertas/:alertId/revisar', requireWriteAccess, async (req, res) => {
    const alertId = Number.parseInt(req.params.alertId, 10);
    if (!Number.isInteger(alertId) || alertId <= 0) {
        return res.status(400).json({ message: 'Identificador de alerta invalido.' });
    }

    if (!canUseOperatorDocumentAlerts(req.authUser)) {
        return res.status(403).json({ message: 'Solo OPERADOR o SUPER pueden revisar esta lista.' });
    }

    const alertScope = getOperatorDocumentAlertsScope(req.authUser);
    if (alertScope.denied) {
        return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para revisar esta lista.' });
    }

    const reviewedBy = req.authUser.nombre || req.authUser.username;
    const reviewedByBranch = normalizeText(req.authUser.sucursal) || null;
    const whereScopeClause = alertScope.clause ? ` AND ${alertScope.clause}` : '';

    try {
        const [updateResult] = await pool.execute(
            `
                UPDATE registro_documentos_alertas
                SET
                    revisado = 1,
                    revisado_por = ?,
                    revisado_por_sucursal = ?,
                    revisado_at = NOW()
                WHERE id = ?
                  AND revisado = 0${whereScopeClause}
            `,
            [reviewedBy, reviewedByBranch, alertId, ...alertScope.params]
        );

        if (Number(updateResult?.affectedRows || 0) === 0) {
            return res.status(404).json({ message: 'No se encontro ese item pendiente para revisar.' });
        }

        return res.json({ message: 'Item marcado como revisado.' });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible marcar el item como revisado.' });
    }
});

app.get('/api/registros/buscar-rango-fechas', async (req, res) => {
    try {
        const fechaDesde = normalizeText(req.query.fechaDesde);
        const fechaHasta = normalizeText(req.query.fechaHasta);
        const sucursal = normalizeText(req.query.sucursal);
        const region = normalizeText(req.query.region);
        const comuna = normalizeText(req.query.comuna);
        const registroScope = getRegistroAccessScope(req.authUser);

        if (registroScope.denied) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para buscar registros.' });
        }

        if (!isValidIsoDateOnly(fechaDesde) || !isValidIsoDateOnly(fechaHasta)) {
            return res.status(400).json({ message: 'Rango de fechas invalido. Usa formato YYYY-MM-DD.' });
        }

        if (fechaDesde > fechaHasta) {
            return res.status(400).json({ message: 'La fecha DESDE no puede ser mayor que HASTA.' });
        }

        if (sucursal && exceedsMaxLength(sucursal, 120)) {
            return res.status(400).json({ message: 'Sucursal excede el largo maximo permitido.' });
        }
        if (region && exceedsMaxLength(region, 120)) {
            return res.status(400).json({ message: 'Region excede el largo maximo permitido.' });
        }
        if (comuna && exceedsMaxLength(comuna, 120)) {
            return res.status(400).json({ message: 'Comuna excede el largo maximo permitido.' });
        }

        const where = ['DATE(created_at) BETWEEN ? AND ?'];
        const values = [fechaDesde, fechaHasta];
        if (sucursal) {
            where.push('(created_by_sucursal = ? OR updated_by_sucursal = ?)');
            values.push(sucursal, sucursal);
        }
        if (region) {
            where.push('region = ?');
            values.push(region);
        }
        if (comuna) {
            where.push('comuna = ?');
            values.push(comuna);
        }
        if (registroScope.clause) {
            where.push(registroScope.clause);
            values.push(...registroScope.params);
        }

        const [rows] = await pool.execute(
            `
                SELECT
                    num_ingreso,
                    nombre,
                    rut,
                    rol,
                    telefono,
                    correo,
                    region,
                    comuna,
                    estado,
                    created_by_sucursal,
                    updated_by_sucursal,
                    created_at
                FROM registros
                WHERE ${where.join(' AND ')}
                ORDER BY created_at DESC, id DESC
                LIMIT 300
            `,
            values
        );

        const registros = rows.map((row) => ({
            numIngreso: row.num_ingreso,
            nombre: row.nombre,
            rut: row.rut,
            rol: row.rol,
            telefono: normalizeText(row.telefono),
            correo: normalizeText(row.correo),
            region: row.region,
            comuna: row.comuna,
            sucursal: normalizeText(row.updated_by_sucursal || row.created_by_sucursal),
            estado: normalizeEstado(row.estado),
            createdAt: row.created_at instanceof Date ? row.created_at.toISOString() : String(row.created_at)
        }));

        return res.json({ registros });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible buscar registros por rango de fechas.' });
    }
});

app.get('/api/registros/buscar', async (req, res) => {
    try {
        const nombre = normalizeText(req.query.nombre);
        const rut = normalizeText(req.query.rut);
        const rol = normalizeText(req.query.rol);
        const rolSearch = normalizeRolForSearch(rol);
        const numIngreso = normalizeText(req.query.numIngreso);
        const registroScope = getRegistroAccessScope(req.authUser);

        if (registroScope.denied) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para buscar registros.' });
        }

        if (exceedsMaxLength(nombre, 255) || exceedsMaxLength(rut, 20) || exceedsMaxLength(rol, 120) || exceedsMaxLength(numIngreso, 20)) {
            return res.status(400).json({ message: 'Uno o mas criterios de busqueda exceden el largo maximo permitido.' });
        }

        if (numIngreso && !parseNumIngreso(numIngreso)) {
            return res.status(400).json({ message: 'NRO INGRESO invalido para busqueda exacta.' });
        }

        if (!nombre && !rut && !rol && !numIngreso) {
            return res.status(400).json({ message: 'Ingresa nombre, rut, rol o NRO INGRESO para buscar.' });
        }

        if (rol && !rolSearch.compact) {
            return res.status(400).json({ message: 'Ingresa un ROL valido para buscar.' });
        }

        const where = [];
        const values = [];

        if (numIngreso) {
            where.push('num_ingreso = ?');
            values.push(numIngreso);
        }

        if (nombre) {
            where.push('LOWER(nombre) LIKE LOWER(?)');
            values.push(`%${nombre}%`);
        }

        if (rut) {
            where.push("REPLACE(REPLACE(REPLACE(LOWER(rut), '.', ''), '-', ''), ' ', '') = ?");
            values.push(normalizeRut(rut));
        }

        if (rol) {
            where.push(buildRolWhereClause('rol'));
            values.push(rolSearch.canonical, rolSearch.compact);
        }

        if (registroScope.clause) {
            where.push(registroScope.clause);
            values.push(...registroScope.params);
        }

        const queryLimit = numIngreso ? 1 : 20;
        const query = `
            SELECT
                id,
                num_ingreso, nombre, rut, telefono, correo,
                region, comuna, nombre_predio, rol,
                num_lotes, estado, comentario,
                factura_nombre_razon, factura_numero, factura_rut, factura_giro,
                factura_direccion, factura_comuna, factura_ciudad,
                factura_contacto, factura_observacion, factura_monto,
                documentos
            FROM registros
            WHERE ${where.join(' AND ')}
            ORDER BY id DESC
            LIMIT ${queryLimit}
        `;

        const [rows] = await pool.execute(query, values);

        if (!rows.length) {
            return res.status(404).json({ message: 'No se encontro un registro con los datos ingresados.' });
        }

        if (rows.length > 1) {
            const coincidencias = rows.map((row) => ({
                numIngreso: row.num_ingreso,
                nombre: row.nombre,
                rut: row.rut,
                rol: row.rol,
                region: row.region,
                comuna: row.comuna
            }));

            return res.json({
                message: `Se encontraron ${rows.length} registros. Refina la busqueda o selecciona un NRO INGRESO exacto.`,
                coincidencias
            });
        }

        const response = toApiRow(rows[0]);
        if (isGuestUser(req.authUser)) {
            response.comentario = '';
            response.historial = [];
            return res.json(response);
        }

        response.historial = await fetchHistoryByRegistroId(rows[0].id);
        return res.json(response);
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible buscar el registro.' });
    }
});

app.get('/api/registros/:numIngreso/historial', async (req, res) => {
    try {
        const numIngreso = normalizeText(decodeURIComponent(req.params.numIngreso));
        const parsedIngreso = parseNumIngreso(numIngreso);
        const registroScope = getRegistroAccessScope(req.authUser);

        if (registroScope.denied) {
            return res.status(403).json({ message: 'Tu usuario no tiene sucursal asignada para consultar historial.' });
        }

        if (!parsedIngreso) {
            return res.status(400).json({ message: 'NRO INGRESO invalido.' });
        }

        const scopeWhere = registroScope.clause ? ` AND ${registroScope.clause}` : '';
        const [rows] = await pool.execute(
            `SELECT id FROM registros WHERE num_ingreso = ?${scopeWhere} LIMIT 1`,
            [numIngreso, ...registroScope.params]
        );
        if (!rows.length) {
            return res.status(404).json({ message: 'No existe ese registro.' });
        }

        if (isGuestUser(req.authUser)) {
            return res.json({ historial: [] });
        }

        const historial = await fetchHistoryByRegistroId(rows[0].id);
        return res.json({ historial });
    } catch (error) {
        console.error(error);
        return res.status(500).json({ message: 'No fue posible cargar el historial de progreso.' });
    }
});

app.use('/assets', express.static(path.join(__dirname, 'assets')));

app.get('/mantenimiento', (req, res) => {
    res.sendFile(path.join(__dirname, 'mantenimiento.html'));
});

app.get('/mantenimiento/', (req, res) => {
    res.sendFile(path.join(__dirname, 'mantenimiento.html'));
});

app.get('/seguimiento', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/index.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/registro', (req, res) => {
    res.redirect(resolveConfiguredLoginRedirectTarget(req.path, req));
});

app.get('/registro/', (req, res) => {
    res.redirect(resolveConfiguredLoginRedirectTarget(req.path, req));
});

app.get('/registro/login', (req, res) => {
    res.redirect(resolveConfiguredLoginRedirectTarget(req.path, req));
});

app.get('/registro/login/', (req, res) => {
    res.redirect(resolveConfiguredLoginRedirectTarget(req.path, req));
});

for (const fileName of FRONTEND_FILES) {
    if (fileName === 'index.html') {
        continue;
    }

    app.get(`/${fileName}`, (req, res) => {
        res.sendFile(path.join(__dirname, fileName));
    });
}

app.use((req, res, next) => {
    const pathname = normalizeRequestPath(req);
    if (isApiRequestPath(pathname)) {
        return res.status(404).json({ message: 'Ruta API no encontrada.' });
    }

    return res.redirect('/index.html');
});

app.use((err, req, res, next) => {
    const pathname = normalizeRequestPath(req);
    const isApiPath = isApiRequestPath(pathname);
    const isJsonParseError =
        err &&
        (err.type === 'entity.parse.failed' ||
            ((err instanceof SyntaxError || err.name === 'SyntaxError') &&
                (Number(err.status) === 400 || Number(err.statusCode) === 400) &&
                Object.prototype.hasOwnProperty.call(err, 'body')));

    if (isJsonParseError) {
        if (isApiPath) {
            return res.status(400).json({ message: 'JSON invalido en la solicitud.' });
        }
        return res.redirect('/index.html');
    }

    if (err && Number.isInteger(err.status) && err.status >= 400 && err.status < 500) {
        if (isApiPath) {
            return res.status(err.status).json({ message: err.message || 'Solicitud invalida.' });
        }
        return res.redirect('/index.html');
    }

    console.error(err);
    if (isApiPath) {
        return res.status(500).json({ message: 'Error interno del servidor.' });
    }
    return res.redirect('/index.html');
});

async function start() {
    try {
        if (!IS_LOCALHOST_BIND && !API_WRITE_KEY) {
            throw new Error('Debes configurar API_WRITE_KEY cuando HOST no es localhost.');
        }

        await initDatabase();
        await new Promise((resolve, reject) => {
            const server = app.listen(PORT, HOST, () => {
                resolve(server);
            });
            server.on('error', reject);
        });

        {
            console.log(`[Config] APP_ENV=${APP_ENV}`);
            if (ENV_CONFIG.loadedFiles.length > 0) {
                console.log(`[Config] Archivos de entorno cargados: ${ENV_CONFIG.loadedFiles.join(', ')}`);
            } else {
                console.log('[Config] No se encontraron archivos .env en el directorio actual.');
            }
            console.log(`Servidor activo en http://${HOST}:${PORT}`);
            console.log(`CORS habilitado para: ${ALLOWED_ORIGINS.join(', ')}`);
            if (!IS_PRODUCTION_ENV) {
                console.log('CORS en desarrollo: se permite acceso desde origenes de red privada (LAN).');
            }

            if (API_WRITE_KEY) {
                console.log('Proteccion de escritura activada (header: X-API-Key).');
            } else {
                console.warn('API_WRITE_KEY no configurada: operaciones POST/PUT sin token (solo recomendable en localhost).');
            }
        }
    } catch (error) {
        console.error('No fue posible iniciar el servidor:', error.message);
        for (const hint of getStartupHints(error)) {
            console.error(`Sugerencia: ${hint}`);
        }
        process.exit(1);
    }
}

start();
