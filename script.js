
const API_BASE_STORAGE = 'geo_rural_api_base';
let API_BASE = resolveApiBase();
const API_BASE_FALLBACKS = resolveApiFallbackBases(API_BASE);
const AUTH_TOKEN_STORAGE = 'geo_rural_auth_token';
const API_KEY_STORAGE = 'geo_rural_api_key';
const DOCUMENT_OPTIONS_STORAGE = 'geo_rural_document_options';
const DOCUMENT_OPTIONS_STORAGE_VERSION = 2;
const DOCUMENT_GROUP_IDS = ['grupo_1', 'grupo_2', 'grupo_3', 'grupo_4', 'grupo_5'];
const DEFAULT_DOCUMENT_GROUP_TITLES = [
    'Identificacion del cliente',
    'Documentos de propiedad',
    'Antecedentes tecnicos',
    'Soportes complementarios',
    'Documentos adicionales'
];
const OPERATOR_LAST_DEST_EMAIL_STORAGE = 'geo_rural_last_invoice_destination_email';
const SECRETARY_DAILY_ALERT_STORAGE_PREFIX = 'geo_rural_secretary_no_movement_seen_';
const SECRETARY_NO_MOVEMENT_DAYS = 8;
const APP_BUILD = '2026-04-10-01';
const MAIL_SEND_PROGRESS_MIN_MS = 1600;
const MAIL_SEND_PROGRESS_FINAL_MS = 900;
const REQUESTED_INVOICES_PAGE_SIZE = 5;
const MAIL_TEMPLATE_TYPES = Object.freeze(['cotizacionHtml', 'facturaSingleHtml', 'facturaPendingHtml']);
const ADMIN_PENDING_MAIL_PREVIEW_MIN_ITEMS = 1;
const ADMIN_PENDING_MAIL_PREVIEW_MAX_ITEMS = 10;
const ADMIN_PENDING_MAIL_PREVIEW_DEFAULT_ITEMS = 2;
const MONTH_NAMES_ES = [
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
];
const REGION_ROMAN_NUMERAL_LOOKUP = Object.freeze({
    'arica y parinacota': 'XV',
    tarapaca: 'I',
    antofagasta: 'II',
    atacama: 'III',
    coquimbo: 'IV',
    valparaiso: 'V',
    metropolitana: 'XIII',
    'metropolitana de santiago': 'XIII',
    'o higgins': 'VI',
    'libertador general bernardo o higgins': 'VI',
    maule: 'VII',
    nuble: 'XVI',
    biobio: 'VIII',
    araucania: 'IX',
    'los rios': 'XIV',
    'los lagos': 'X',
    aysen: 'XI',
    'aysen del general carlos ibanez del campo': 'XI',
    magallanes: 'XII',
    'magallanes y de la antartica chilena': 'XII'
});

const REGION_COMUNAS = {
    'Arica y Parinacota': ['Arica', 'Camarones', 'Putre', 'General Lagos'],
    Tarapaca: ['Iquique', 'Alto Hospicio', 'Pozo Almonte', 'Camina', 'Colchane', 'Huara', 'Pica'],
    Antofagasta: [
        'Antofagasta',
        'Mejillones',
        'Sierra Gorda',
        'Taltal',
        'Calama',
        'Ollague',
        'San Pedro de Atacama',
        'Tocopilla',
        'Maria Elena'
    ],
    Atacama: ['Copiapo', 'Caldera', 'Tierra Amarilla', 'Chanaral', 'Diego de Almagro', 'Vallenar', 'Alto del Carmen', 'Freirina', 'Huasco'],
    Coquimbo: [
        'La Serena',
        'Coquimbo',
        'Andacollo',
        'La Higuera',
        'Paiguano',
        'Vicuna',
        'Illapel',
        'Canela',
        'Los Vilos',
        'Salamanca',
        'Ovalle',
        'Combarbala',
        'Monte Patria',
        'Punitaqui',
        'Rio Hurtado'
    ],
    Valparaiso: [
        'Valparaiso',
        'Casablanca',
        'Concon',
        'Juan Fernandez',
        'Puchuncavi',
        'Quintero',
        'Vina del Mar',
        'Isla de Pascua',
        'Los Andes',
        'Calle Larga',
        'Rinconada',
        'San Esteban',
        'La Ligua',
        'Cabildo',
        'Papudo',
        'Petorca',
        'Zapallar',
        'Quillota',
        'Calera',
        'Hijuelas',
        'La Cruz',
        'Nogales',
        'San Antonio',
        'Algarrobo',
        'Cartagena',
        'El Quisco',
        'El Tabo',
        'Santo Domingo',
        'San Felipe',
        'Catemu',
        'Llaillay',
        'Panquehue',
        'Putaendo',
        'Santa Maria',
        'Quilpue',
        'Limache',
        'Olmue',
        'Villa Alemana'
    ],
    "O'Higgins": [
        'Rancagua',
        'Codegua',
        'Coinco',
        'Coltauco',
        'Donihue',
        'Graneros',
        'Las Cabras',
        'Machali',
        'Malloa',
        'Mostazal',
        'Olivar',
        'Peumo',
        'Pichidegua',
        'Quinta de Tilcoco',
        'Rengo',
        'Requinoa',
        'San Vicente',
        'Pichilemu',
        'La Estrella',
        'Litueche',
        'Marchihue',
        'Navidad',
        'Paredones',
        'San Fernando',
        'Chepica',
        'Chimbarongo',
        'Lolol',
        'Nancagua',
        'Palmilla',
        'Peralillo',
        'Placilla',
        'Pumanque',
        'Santa Cruz'
    ],
    Maule: [
        'Talca',
        'Constitucion',
        'Curepto',
        'Empedrado',
        'Maule',
        'Pelarco',
        'Pencahue',
        'Rio Claro',
        'San Clemente',
        'San Rafael',
        'Cauquenes',
        'Chanco',
        'Pelluhue',
        'Curico',
        'Hualane',
        'Licanten',
        'Molina',
        'Rauco',
        'Romeral',
        'Sagrada Familia',
        'Teno',
        'Vichuquen',
        'Linares',
        'Colbun',
        'Longavi',
        'Parral',
        'Retiro',
        'San Javier',
        'Villa Alegre',
        'Yerbas Buenas'
    ],
    Nuble: [
        'Cobquecura',
        'Coelemu',
        'Ninhue',
        'Portezuelo',
        'Quirihue',
        'Ranquil',
        'Trehuaco',
        'Bulnes',
        'Chillan Viejo',
        'Chillan',
        'El Carmen',
        'Pemuco',
        'Pinto',
        'Quillon',
        'San Ignacio',
        'Yungay',
        'Coihueco',
        'Niquen',
        'San Carlos',
        'San Fabian',
        'San Nicolas'
    ],
    Biobio: [
        'Concepcion',
        'Coronel',
        'Chiguayante',
        'Florida',
        'Hualqui',
        'Lota',
        'Penco',
        'San Pedro de la Paz',
        'Santa Juana',
        'Talcahuano',
        'Tome',
        'Hualpen',
        'Lebu',
        'Arauco',
        'Canete',
        'Contulmo',
        'Curanilahue',
        'Los Alamos',
        'Tirua',
        'Los Angeles',
        'Antuco',
        'Cabrero',
        'Laja',
        'Mulchen',
        'Nacimiento',
        'Negrete',
        'Quilaco',
        'Quilleco',
        'San Rosendo',
        'Santa Barbara',
        'Tucapel',
        'Yumbel',
        'Alto Biobio'
    ],
    Araucania: [
        'Temuco',
        'Carahue',
        'Cunco',
        'Curarrehue',
        'Freire',
        'Galvarino',
        'Gorbea',
        'Lautaro',
        'Loncoche',
        'Melipeuco',
        'Nueva Imperial',
        'Padre Las Casas',
        'Perquenco',
        'Pitrufquen',
        'Pucon',
        'Saavedra',
        'Teodoro Schmidt',
        'Tolten',
        'Vilcun',
        'Villarrica',
        'Cholchol',
        'Angol',
        'Collipulli',
        'Curacautin',
        'Ercilla',
        'Lonquimay',
        'Los Sauces',
        'Lumaco',
        'Puren',
        'Renaico',
        'Traiguen',
        'Victoria'
    ],
    'Los Rios': ['Valdivia', 'Corral', 'Lanco', 'Los Lagos', 'Mafil', 'Mariquina', 'Paillaco', 'Panguipulli', 'La Union', 'Futrono', 'Lago Ranco', 'Rio Bueno'],
    'Los Lagos': [
        'Puerto Montt',
        'Calbuco',
        'Cochamo',
        'Fresia',
        'Frutillar',
        'Los Muermos',
        'Llanquihue',
        'Maullin',
        'Puerto Varas',
        'Castro',
        'Ancud',
        'Chonchi',
        'Curaco de Velez',
        'Dalcahue',
        'Puqueldon',
        'Queilen',
        'Quellon',
        'Quemchi',
        'Quinchao',
        'Osorno',
        'Puerto Octay',
        'Purranque',
        'Puyehue',
        'Rio Negro',
        'San Juan de la Costa',
        'San Pablo',
        'Chaiten',
        'Futaleufu',
        'Hualaihue',
        'Palena'
    ],
    Aysen: ['Coyhaique', 'Lago Verde', 'Aysen', 'Cisnes', 'Guaitecas', 'Cochrane', "O'Higgins", 'Tortel', 'Chile Chico', 'Rio Ibanez'],
    Magallanes: [
        'Punta Arenas',
        'Laguna Blanca',
        'Rio Verde',
        'San Gregorio',
        'Cabo de Hornos',
        'Antartica',
        'Porvenir',
        'Primavera',
        'Timaukel',
        'Natales',
        'Torres del Paine'
    ],
    Metropolitana: [
        'Cerrillos',
        'Cerro Navia',
        'Conchali',
        'El Bosque',
        'Estacion Central',
        'Huechuraba',
        'Independencia',
        'La Cisterna',
        'La Florida',
        'La Granja',
        'La Pintana',
        'La Reina',
        'Las Condes',
        'Lo Barnechea',
        'Lo Espejo',
        'Lo Prado',
        'Macul',
        'Maipu',
        'Nunoa',
        'Pedro Aguirre Cerda',
        'Penalolen',
        'Providencia',
        'Pudahuel',
        'Quilicura',
        'Quinta Normal',
        'Recoleta',
        'Renca',
        'Santiago',
        'San Joaquin',
        'San Miguel',
        'San Ramon',
        'Vitacura',
        'Puente Alto',
        'Pirque',
        'San Jose de Maipo',
        'Colina',
        'Lampa',
        'Tiltil',
        'San Bernardo',
        'Buin',
        'Calera de Tango',
        'Paine',
        'Melipilla',
        'Alhue',
        'Curacavi',
        'Maria Pinto',
        'San Pedro',
        'Talagante',
        'El Monte',
        'Isla de Maipo',
        'Padre Hurtado',
        'Penaflor'
    ]
};

let ui = null;
let currentUser = null;
let loadedIngresoOriginal = '';
let suggestedIngresoForNewRecord = '';
let suppressResetHandler = false;
let adminIngresoMode = 'current';
let appViewMode = 'registro';
let editingAdminUserId = null;
let editingAdminUserRole = '';
let editingBranchId = null;
let editingDocumentValue = '';
let availableBranches = [];
let defaultDocumentGroups = [];
let defaultDocumentOptions = [];
let configuredDocumentGroups = [];
let configuredDocumentOptions = [];
let loadedRecordUnknownDocuments = [];
let loadedComentarioOriginal = '';
let loadedRegistroSnapshot = null;
let authTokenMemory = '';
let apiKeyMemory = '';
let operatorPendingInvoices = [];
let operatorDocumentAlerts = [];
let requestedInvoicesHistory = [];
let selectedRequestedInvoiceRequestId = '';
let requestedInvoicesCurrentPage = 1;
let currentInvoiceRequestStatus = '';
let invoiceModalMode = '';
let invoiceEmailPromptState = null;
let editingHistoryEntryId = null;
let editingHistoryIngreso = '';
let editingHistoryAction = '';
let secretaryNoMovementRecords = [];
let dateRangeMatches = [];
let selectedDateRangeIngreso = '';
let dateRangeCurrentPage = 1;
let dateRangePageSize = 6;
let dateRangeResizeTimerId = 0;
let searchMatchCandidates = [];
let selectedSearchMatchIngreso = '';
let searchMatchCurrentPage = 1;
let dateRangeAvailableBranches = [];
let pendingMonthlyUtmStatus = null;
let secretariaQuotationSummary = null;
let secretariaQuotationTemplateMetadata = null;
let adminEmailConfigLoaded = null;
let runtimeMailTemplates = null;
let adminMailTemplatesLoaded = null;
let adminMailTemplateDrafts = null;
let adminMailTemplateActiveType = 'cotizacionHtml';
let adminMailTemplatePendingPreviewCount = ADMIN_PENDING_MAIL_PREVIEW_DEFAULT_ITEMS;
let adminUsageStatsLoaded = null;
let adminMaintenanceStatusLoaded = null;
let adminMaintenanceUpdateInProgress = false;
let adminLoginRouteNoticeStatusLoaded = null;
let adminLoginRouteNoticeUpdateInProgress = false;
let mailSendProgressOpenedAt = 0;
let authLoginInProgress = false;
const DEFAULT_CREATION_COMMENT = 'creacion y recepcion al iniciar un nuevo registro';
const REGISTRO_REQUIRED_FIELD_IDS = ['nombre', 'rut', 'region', 'comuna', 'rol'];
const REGISTRO_INLINE_ERROR_FIELD_IDS = [...REGISTRO_REQUIRED_FIELD_IDS, 'correo', 'numLotes'];
const DATE_RANGE_MIN_PAGE_SIZE = 6;
const DATE_RANGE_FALLBACK_PAGE_SIZE = 10;
const DATE_RANGE_MAX_PAGE_SIZE = 40;
const SEARCH_MATCH_PAGE_SIZE = 12;
const FACTURA_NUMBER_MAX_LENGTH = 80;
const FACTURA_REQUIRED_FIELD_IDS = ['facturaNombreRazon', 'facturaRut', 'facturaGiro', 'facturaDireccion', 'facturaMonto'];
const FACTURA_INLINE_ERROR_FIELD_IDS = [...FACTURA_REQUIRED_FIELD_IDS, 'facturaComuna', 'facturaCiudad', 'facturaContacto', 'facturaNumero'];
const EXISTING_RECORD_SAVE_WARNING =
    'No se puede crear porque el registro ya existe. Si ha realizado cambios pulse en MODIFICAR para guardar los cambios.';
const PENDING_MODIFY_REQUIRED_WARNING =
    'Se detectaron cambios en el formulario o documentacion recibida. Favor guardar cambios pulsando el boton MODIFICAR.';

document.addEventListener('DOMContentLoaded', async () => {
    console.info(`[Geo Rural] Frontend build ${APP_BUILD}`);
    ui = {
        appTopPanels: document.getElementById('appTopPanels'),
        authCard: document.getElementById('authCard'),
        passwordChangeCard: document.getElementById('passwordChangeCard'),
        passwordChangeForm: document.getElementById('passwordChangeForm'),
        passwordChangeIntro: document.getElementById('passwordChangeIntro'),
        currentPassword: document.getElementById('currentPassword'),
        newPassword: document.getElementById('newPassword'),
        confirmPassword: document.getElementById('confirmPassword'),
        passwordChangeMessage: document.getElementById('passwordChangeMessage'),
        utmCaptureCard: document.getElementById('utmCaptureCard'),
        utmCaptureForm: document.getElementById('utmCaptureForm'),
        utmCaptureIntro: document.getElementById('utmCaptureIntro'),
        utmCapturePeriod: document.getElementById('utmCapturePeriod'),
        utmCaptureSource: document.getElementById('utmCaptureSource'),
        utmValue: document.getElementById('utmValue'),
        utmConfirmValue: document.getElementById('utmConfirmValue'),
        utmCaptureMessage: document.getElementById('utmCaptureMessage'),
        loginForm: document.getElementById('loginForm'),
        loginSubmitBtn: document.querySelector('#loginForm button[type="submit"]'),
        guestLoginBtn: document.getElementById('guestLoginBtn'),
        loginUsername: document.getElementById('loginUsername'),
        loginPassword: document.getElementById('loginPassword'),
        loginMessage: document.getElementById('loginMessage'),
        registroForm: document.getElementById('registroForm'),
        registroMainSection: document.getElementById('registroMainSection'),
        sessionUser: document.getElementById('sessionUser'),
        logoutBtn: document.getElementById('logoutBtn'),
        adminPanel: document.getElementById('adminPanel'),
        adminPanelTitle: document.getElementById('adminPanelTitle'),
        operatorBillingPanel: document.getElementById('operatorBillingPanel'),
        operatorEmitInvoiceBtn: document.getElementById('operatorEmitInvoiceBtn'),
        operatorQueueInvoiceBtn: document.getElementById('operatorQueueInvoiceBtn'),
        operatorSendPendingBtn: document.getElementById('operatorSendPendingBtn'),
        operatorPendingSummary: document.getElementById('operatorPendingSummary'),
        operatorInvoiceMessage: document.getElementById('operatorInvoiceMessage'),
        operatorPendingInvoicesBody: document.getElementById('operatorPendingInvoicesBody'),
        adminUsersSection: document.getElementById('adminUsersSection'),
        adminBranchesSection: document.getElementById('adminBranchesSection'),
        adminDocumentsSection: document.getElementById('adminDocumentsSection'),
        adminEmailSection: document.getElementById('adminEmailSection'),
        adminMailTemplatesSection: document.getElementById('adminMailTemplatesSection'),
        adminUsageSection: document.getElementById('adminUsageSection'),
        adminMaintenanceSection: document.getElementById('adminMaintenanceSection'),
        adminToggleUsersBtn: document.getElementById('adminToggleUsersBtn'),
        adminToggleBranchesBtn: document.getElementById('adminToggleBranchesBtn'),
        adminToggleDocumentsBtn: document.getElementById('adminToggleDocumentsBtn'),
        adminToggleEmailBtn: document.getElementById('adminToggleEmailBtn'),
        adminToggleMailTemplatesBtn: document.getElementById('adminToggleMailTemplatesBtn'),
        adminToggleUsageBtn: document.getElementById('adminToggleUsageBtn'),
        adminToggleMaintenanceBtn: document.getElementById('adminToggleMaintenanceBtn'),
        secretariaViewInvoicesBtn: document.getElementById('secretariaViewInvoicesBtn'),
        secretariaQuotationBtn: document.getElementById('secretariaQuotationBtn'),
        secretariaEditUtmBtn: document.getElementById('secretariaEditUtmBtn'),
        secretariaNoMovementBtn: document.getElementById('secretariaNoMovementBtn'),
        adminCreateUserBtn: document.getElementById('adminCreateUserBtn'),
        adminCancelEditBtn: document.getElementById('adminCancelEditBtn'),
        adminYearCurrentBtn: document.getElementById('adminYearCurrentBtn'),
        adminYearPastBtn: document.getElementById('adminYearPastBtn'),
        adminYearMessage: document.getElementById('adminYearMessage'),
        adminUserForm: document.getElementById('adminUserForm'),
        adminUsername: document.getElementById('adminUsername'),
        adminNombre: document.getElementById('adminNombre'),
        adminSucursal: document.getElementById('adminSucursal'),
        adminPassword: document.getElementById('adminPassword'),
        adminRole: document.getElementById('adminRole'),
        adminMessage: document.getElementById('adminMessage'),
        adminUsersBody: document.getElementById('adminUsersBody'),
        adminBranchForm: document.getElementById('adminBranchForm'),
        adminBranchName: document.getElementById('adminBranchName'),
        adminBranchSaveBtn: document.getElementById('adminBranchSaveBtn'),
        adminBranchCancelBtn: document.getElementById('adminBranchCancelBtn'),
        adminBranchMessage: document.getElementById('adminBranchMessage'),
        adminBranchesBody: document.getElementById('adminBranchesBody'),
        adminDocumentForm: document.getElementById('adminDocumentForm'),
        adminDocumentLabel: document.getElementById('adminDocumentLabel'),
        adminDocumentGroup: document.getElementById('adminDocumentGroup'),
        adminDocumentSaveBtn: document.getElementById('adminDocumentSaveBtn'),
        adminDocumentCancelBtn: document.getElementById('adminDocumentCancelBtn'),
        adminDocumentGroupTitles: document.getElementById('adminDocumentGroupTitles'),
        adminDocumentGroupSaveBtn: document.getElementById('adminDocumentGroupSaveBtn'),
        adminDocumentMessage: document.getElementById('adminDocumentMessage'),
        adminDocumentsBody: document.getElementById('adminDocumentsBody'),
        adminEmailForm: document.getElementById('adminEmailForm'),
        adminEmailHost: document.getElementById('adminEmailHost'),
        adminEmailPort: document.getElementById('adminEmailPort'),
        adminEmailSecure: document.getElementById('adminEmailSecure'),
        adminEmailUser: document.getElementById('adminEmailUser'),
        adminEmailPassword: document.getElementById('adminEmailPassword'),
        adminEmailKeepPassword: document.getElementById('adminEmailKeepPassword'),
        adminEmailFromEmail: document.getElementById('adminEmailFromEmail'),
        adminEmailFromName: document.getElementById('adminEmailFromName'),
        adminEmailLoadBtn: document.getElementById('adminEmailLoadBtn'),
        adminEmailSaveBtn: document.getElementById('adminEmailSaveBtn'),
        adminEmailSource: document.getElementById('adminEmailSource'),
        adminEmailMessage: document.getElementById('adminEmailMessage'),
        adminMailTemplateType: document.getElementById('adminMailTemplateType'),
        adminMailTemplateEditor: document.getElementById('adminMailTemplateEditor'),
        adminMailTemplatePreviewBtn: document.getElementById('adminMailTemplatePreviewBtn'),
        adminMailTemplateReloadBtn: document.getElementById('adminMailTemplateReloadBtn'),
        adminMailTemplateSaveBtn: document.getElementById('adminMailTemplateSaveBtn'),
        adminMailTemplateHints: document.getElementById('adminMailTemplateHints'),
        adminMailTemplatePreviewFrame: document.getElementById('adminMailTemplatePreviewFrame'),
        adminMailTemplatePendingControls: document.getElementById('adminMailTemplatePendingControls'),
        adminMailTemplatePendingLessBtn: document.getElementById('adminMailTemplatePendingLessBtn'),
        adminMailTemplatePendingMoreBtn: document.getElementById('adminMailTemplatePendingMoreBtn'),
        adminMailTemplatePendingCountLabel: document.getElementById('adminMailTemplatePendingCountLabel'),
        adminMailTemplateSource: document.getElementById('adminMailTemplateSource'),
        adminMailTemplateMessage: document.getElementById('adminMailTemplateMessage'),
        adminUsageSummary: document.getElementById('adminUsageSummary'),
        adminUsageRefreshBtn: document.getElementById('adminUsageRefreshBtn'),
        adminUsageBody: document.getElementById('adminUsageBody'),
        adminUsageMessage: document.getElementById('adminUsageMessage'),
        adminMaintenanceState: document.getElementById('adminMaintenanceState'),
        adminMaintenanceUpdatedAt: document.getElementById('adminMaintenanceUpdatedAt'),
        adminMaintenanceUpdatedBy: document.getElementById('adminMaintenanceUpdatedBy'),
        adminMaintenanceEnableBtn: document.getElementById('adminMaintenanceEnableBtn'),
        adminMaintenanceDisableBtn: document.getElementById('adminMaintenanceDisableBtn'),
        adminMaintenanceRefreshBtn: document.getElementById('adminMaintenanceRefreshBtn'),
        adminMaintenanceMessage: document.getElementById('adminMaintenanceMessage'),
        adminLoginNoticeState: document.getElementById('adminLoginNoticeState'),
        adminLoginNoticeUrl: document.getElementById('adminLoginNoticeUrl'),
        adminLoginNoticeExpiresAt: document.getElementById('adminLoginNoticeExpiresAt'),
        adminLoginNoticeUpdatedBy: document.getElementById('adminLoginNoticeUpdatedBy'),
        adminLoginNoticeEnableBtn: document.getElementById('adminLoginNoticeEnableBtn'),
        adminLoginNoticeDisableBtn: document.getElementById('adminLoginNoticeDisableBtn'),
        adminLoginNoticeRefreshBtn: document.getElementById('adminLoginNoticeRefreshBtn'),
        adminLoginNoticeMessage: document.getElementById('adminLoginNoticeMessage'),
        numIngreso: document.getElementById('numIngreso'),
        numIngresoHint: document.getElementById('numIngresoHint'),
        estado: document.getElementById('estado'),
        facturaNumeroGroup: document.getElementById('facturaNumeroGroup'),
        facturaNumero: document.getElementById('facturaNumero'),
        region: document.getElementById('region'),
        comuna: document.getElementById('comuna'),
        comentario: document.getElementById('comentario'),
        guardarComentarioBtn: document.getElementById('guardarComentarioBtn'),
        comentarioGroup: document.getElementById('comentarioGroup'),
        comentariosHistoryGroup: document.getElementById('comentariosHistoryGroup'),
        comentariosHistoryMainAnchor: document.getElementById('comentariosHistoryMainAnchor'),
        comentariosHistorySideAnchor: document.getElementById('comentariosHistorySideAnchor'),
        registroActionsSecretaryAnchor: document.getElementById('registroActionsSecretaryAnchor'),
        registroActionsRow: document.getElementById('registroActionsRow'),
        registroActionsDefaultAnchor: document.getElementById('registroActionsDefaultAnchor'),
        documentosContainer: document.getElementById('documentosContainer'),
        formMessage: document.getElementById('formMessage'),
        historialBody: document.getElementById('historialBody'),
        historialActionHeader: document.getElementById('historialActionHeader'),
        facturaDataSection: document.getElementById('facturaDataSection'),
        facturaNombreRazon: document.getElementById('facturaNombreRazon'),
        facturaRut: document.getElementById('facturaRut'),
        facturaGiro: document.getElementById('facturaGiro'),
        facturaDireccion: document.getElementById('facturaDireccion'),
        facturaComuna: document.getElementById('facturaComuna'),
        facturaCiudad: document.getElementById('facturaCiudad'),
        facturaContacto: document.getElementById('facturaContacto'),
        facturaObservacion: document.getElementById('facturaObservacion'),
        facturaMonto: document.getElementById('facturaMonto'),
        facturaEnviarContadorBtn: document.getElementById('facturaEnviarContadorBtn'),
        facturaAgregarListaBtn: document.getElementById('facturaAgregarListaBtn'),
        facturaEnviarPendientesBtn: document.getElementById('facturaEnviarPendientesBtn'),
        facturaFormMessage: document.getElementById('facturaFormMessage'),
        btnGuardar: document.querySelector('#registroForm .button-row .guardar'),
        btnModificar: document.querySelector('#registroForm .button-row .modificar'),
        btnBuscar: document.querySelector('#registroForm .button-row .buscar'),
        btnBuscarFecha: document.getElementById('buscarFechaBtn'),
        operatorDocAlertsBtn: document.getElementById('operatorDocAlertsBtn'),
        btnEliminar: document.querySelector('#registroForm .button-row .eliminar'),
        btnLimpiar: document.querySelector('#registroForm .button-row .limpiar'),
        invoiceModalOverlay: document.getElementById('invoiceModalOverlay'),
        invoiceModalForm: document.getElementById('invoiceModalForm'),
        invoiceModalTitle: document.getElementById('invoiceModalTitle'),
        invoiceModalSubmitBtn: document.getElementById('invoiceModalSubmitBtn'),
        invoiceModalCancelBtn: document.getElementById('invoiceModalCancelBtn'),
        invoiceModalMessage: document.getElementById('invoiceModalMessage'),
        invoiceNumIngreso: document.getElementById('invoiceNumIngreso'),
        invoiceNombreCliente: document.getElementById('invoiceNombreCliente'),
        invoiceRutCliente: document.getElementById('invoiceRutCliente'),
        invoiceCorreoFactura: document.getElementById('invoiceCorreoFactura'),
        invoiceDireccion: document.getElementById('invoiceDireccion'),
        invoiceGiro: document.getElementById('invoiceGiro'),
        invoiceComunaRegion: document.getElementById('invoiceComunaRegion'),
        invoiceReferencia: document.getElementById('invoiceReferencia'),
        secretariaInvoicesModalOverlay: document.getElementById('secretariaInvoicesModalOverlay'),
        secretariaInvoicesSearchName: document.getElementById('secretariaInvoicesSearchName'),
        secretariaInvoicesSearchDate: document.getElementById('secretariaInvoicesSearchDate'),
        secretariaInvoicesClearFiltersBtn: document.getElementById('secretariaInvoicesClearFiltersBtn'),
        secretariaInvoicesSummary: document.getElementById('secretariaInvoicesSummary'),
        secretariaInvoicesBody: document.getElementById('secretariaInvoicesBody'),
        secretariaInvoicesPagination: document.getElementById('secretariaInvoicesPagination'),
        secretariaInvoicesFirstPageBtn: document.getElementById('secretariaInvoicesFirstPageBtn'),
        secretariaInvoicesPrevPageBtn: document.getElementById('secretariaInvoicesPrevPageBtn'),
        secretariaInvoicesNextPageBtn: document.getElementById('secretariaInvoicesNextPageBtn'),
        secretariaInvoicesLastPageBtn: document.getElementById('secretariaInvoicesLastPageBtn'),
        secretariaInvoicesPageInfo: document.getElementById('secretariaInvoicesPageInfo'),
        secretariaInvoicesMessage: document.getElementById('secretariaInvoicesMessage'),
        secretariaInvoicesDetailBtn: document.getElementById('secretariaInvoicesDetailBtn'),
        secretariaInvoicesCloseBtn: document.getElementById('secretariaInvoicesCloseBtn'),
        secretariaInvoiceDetailModalOverlay: document.getElementById('secretariaInvoiceDetailModalOverlay'),
        secretariaInvoiceDetailRequestId: document.getElementById('secretariaInvoiceDetailRequestId'),
        secretariaInvoiceDetailNumIngreso: document.getElementById('secretariaInvoiceDetailNumIngreso'),
        secretariaInvoiceDetailEstado: document.getElementById('secretariaInvoiceDetailEstado'),
        secretariaInvoiceDetailNombre: document.getElementById('secretariaInvoiceDetailNombre'),
        secretariaInvoiceDetailRut: document.getElementById('secretariaInvoiceDetailRut'),
        secretariaInvoiceDetailMonto: document.getElementById('secretariaInvoiceDetailMonto'),
        secretariaInvoiceDetailGiro: document.getElementById('secretariaInvoiceDetailGiro'),
        secretariaInvoiceDetailDireccion: document.getElementById('secretariaInvoiceDetailDireccion'),
        secretariaInvoiceDetailComuna: document.getElementById('secretariaInvoiceDetailComuna'),
        secretariaInvoiceDetailCiudad: document.getElementById('secretariaInvoiceDetailCiudad'),
        secretariaInvoiceDetailContacto: document.getElementById('secretariaInvoiceDetailContacto'),
        secretariaInvoiceDetailDestinationEmail: document.getElementById('secretariaInvoiceDetailDestinationEmail'),
        secretariaInvoiceDetailObservacion: document.getElementById('secretariaInvoiceDetailObservacion'),
        secretariaInvoiceDetailMessage: document.getElementById('secretariaInvoiceDetailMessage'),
        secretariaInvoiceDetailSaveBtn: document.getElementById('secretariaInvoiceDetailSaveBtn'),
        secretariaInvoiceDetailResendBtn: document.getElementById('secretariaInvoiceDetailResendBtn'),
        secretariaInvoiceDetailCloseBtn: document.getElementById('secretariaInvoiceDetailCloseBtn'),
        secretariaQuotationModalOverlay: document.getElementById('secretariaQuotationModalOverlay'),
        secretariaQuotationForm: document.getElementById('secretariaQuotationForm'),
        secretariaQuotationClientName: document.getElementById('secretariaQuotationClientName'),
        secretariaQuotationClientEmail: document.getElementById('secretariaQuotationClientEmail'),
        secretariaQuotationPeriod: document.getElementById('secretariaQuotationPeriod'),
        secretariaQuotationUtmValue: document.getElementById('secretariaQuotationUtmValue'),
        secretariaQuotationReference: document.getElementById('secretariaQuotationReference'),
        secretariaQuotationParcelamientoRange: document.getElementById('secretariaQuotationParcelamientoRange'),
        secretariaQuotationSource: document.getElementById('secretariaQuotationSource'),
        secretariaQuotationItemsBody: document.getElementById('secretariaQuotationItemsBody'),
        secretariaQuotationTotals: document.getElementById('secretariaQuotationTotals'),
        secretariaQuotationMessage: document.getElementById('secretariaQuotationMessage'),
        secretariaQuotationRefreshBtn: document.getElementById('secretariaQuotationRefreshBtn'),
        secretariaQuotationReplaceTemplateBtn: document.getElementById('secretariaQuotationReplaceTemplateBtn'),
        secretariaQuotationSendBtn: document.getElementById('secretariaQuotationSendBtn'),
        secretariaQuotationCloseBtn: document.getElementById('secretariaQuotationCloseBtn'),
        secretariaQuotationTemplateInput: document.getElementById('secretariaQuotationTemplateInput'),
        secretariaQuotationTemplateInfo: document.getElementById('secretariaQuotationTemplateInfo'),
        mailSendProgressOverlay: document.getElementById('mailSendProgressOverlay'),
        mailSendProgressText: document.getElementById('mailSendProgressText'),
        mailSendProgressStatus: document.getElementById('mailSendProgressStatus'),
        invoiceEmailPromptOverlay: document.getElementById('invoiceEmailPromptOverlay'),
        invoiceEmailPromptInput: document.getElementById('invoiceEmailPromptInput'),
        invoiceEmailPromptMessage: document.getElementById('invoiceEmailPromptMessage'),
        invoiceEmailPromptCancelBtn: document.getElementById('invoiceEmailPromptCancelBtn'),
        invoiceEmailPromptSendBtn: document.getElementById('invoiceEmailPromptSendBtn'),
        historyEditModalOverlay: document.getElementById('historyEditModalOverlay'),
        historyEditForm: document.getElementById('historyEditForm'),
        historyEditFecha: document.getElementById('historyEditFecha'),
        historyEditComentario: document.getElementById('historyEditComentario'),
        historyEditMessage: document.getElementById('historyEditMessage'),
        historyEditDeleteBtn: document.getElementById('historyEditDeleteBtn'),
        historyEditCancelBtn: document.getElementById('historyEditCancelBtn'),
        historyEditSaveBtn: document.getElementById('historyEditSaveBtn'),
        secretaryNoMovementOverlay: document.getElementById('secretaryNoMovementOverlay'),
        secretaryNoMovementFilterForm: document.getElementById('secretaryNoMovementFilterForm'),
        secretaryNoMovementFilterRegion: document.getElementById('secretaryNoMovementFilterRegion'),
        secretaryNoMovementFilterComuna: document.getElementById('secretaryNoMovementFilterComuna'),
        secretaryNoMovementFilterSucursal: document.getElementById('secretaryNoMovementFilterSucursal'),
        secretaryNoMovementFilterFrom: document.getElementById('secretaryNoMovementFilterFrom'),
        secretaryNoMovementFilterTo: document.getElementById('secretaryNoMovementFilterTo'),
        secretaryNoMovementApplyBtn: document.getElementById('secretaryNoMovementApplyBtn'),
        secretaryNoMovementClearBtn: document.getElementById('secretaryNoMovementClearBtn'),
        secretaryNoMovementBody: document.getElementById('secretaryNoMovementBody'),
        secretaryNoMovementMessage: document.getElementById('secretaryNoMovementMessage'),
        secretaryNoMovementCloseBtn: document.getElementById('secretaryNoMovementCloseBtn'),
        operatorDocAlertsOverlay: document.getElementById('operatorDocAlertsOverlay'),
        operatorDocAlertsBody: document.getElementById('operatorDocAlertsBody'),
        operatorDocAlertsMessage: document.getElementById('operatorDocAlertsMessage'),
        operatorDocAlertsRefreshBtn: document.getElementById('operatorDocAlertsRefreshBtn'),
        operatorDocAlertsCloseBtn: document.getElementById('operatorDocAlertsCloseBtn'),
        dateRangeModalOverlay: document.getElementById('dateRangeModalOverlay'),
        dateRangeForm: document.getElementById('dateRangeForm'),
        dateRangeFrom: document.getElementById('dateRangeFrom'),
        dateRangeTo: document.getElementById('dateRangeTo'),
        dateRangeSucursal: document.getElementById('dateRangeSucursal'),
        dateRangeRegion: document.getElementById('dateRangeRegion'),
        dateRangeComuna: document.getElementById('dateRangeComuna'),
        dateRangeTableWrapper: document.getElementById('dateRangeTableWrapper'),
        dateRangeResultsBody: document.getElementById('dateRangeResultsBody'),
        dateRangeMessage: document.getElementById('dateRangeMessage'),
        dateRangeExportBtn: document.getElementById('dateRangeExportBtn'),
        dateRangeAcceptBtn: document.getElementById('dateRangeAcceptBtn'),
        dateRangeCancelBtn: document.getElementById('dateRangeCancelBtn'),
        dateRangeFirstBtn: document.getElementById('dateRangeFirstBtn'),
        dateRangePrevBtn: document.getElementById('dateRangePrevBtn'),
        dateRangePageInfo: document.getElementById('dateRangePageInfo'),
        dateRangeNextBtn: document.getElementById('dateRangeNextBtn'),
        dateRangeLastBtn: document.getElementById('dateRangeLastBtn'),
        searchMatchModalOverlay: document.getElementById('searchMatchModalOverlay'),
        searchMatchResultsBody: document.getElementById('searchMatchResultsBody'),
        searchMatchMessage: document.getElementById('searchMatchMessage'),
        searchMatchFirstBtn: document.getElementById('searchMatchFirstBtn'),
        searchMatchPrevBtn: document.getElementById('searchMatchPrevBtn'),
        searchMatchPageInfo: document.getElementById('searchMatchPageInfo'),
        searchMatchNextBtn: document.getElementById('searchMatchNextBtn'),
        searchMatchLastBtn: document.getElementById('searchMatchLastBtn'),
        searchMatchAcceptBtn: document.getElementById('searchMatchAcceptBtn'),
        searchMatchCancelBtn: document.getElementById('searchMatchCancelBtn')
    };

    // Fuerza estado inicial: solo login visible hasta validar sesion.
    if (ui.appTopPanels) {
        ui.appTopPanels.classList.add('hidden');
    }
    document.body.classList.remove('app-authenticated');
    ui.registroForm.classList.add('hidden');
    ui.authCard.classList.remove('hidden');
    ui.passwordChangeCard.classList.add('hidden');
    if (ui.utmCaptureCard) {
        ui.utmCaptureCard.classList.add('hidden');
    }
    ui.adminPanel.classList.add('hidden');
    setOperatorBillingVisibility(false);
    setFacturaContainerVisibility(false);
    syncFacturaNumeroVisibility({ clearOnHide: true });

    populateRegiones(ui.region, 'Seleccione region');
    populateRegiones(ui.dateRangeRegion, 'TODAS');
    populateComunas(ui.dateRangeComuna, '', '', 'TODAS');
    populateRegiones(ui.secretaryNoMovementFilterRegion, 'TODAS');
    populateComunas(ui.secretaryNoMovementFilterComuna, '', '', 'TODAS');
    populateSecretaryNoMovementBranchFilter([]);
    initializeSecretariaInvoiceDetailLocationSelectors();
    renderHistory([]);
    setFormMessage(ui.formMessage, '', '');
    setFormMessage(ui.loginMessage, '', '');
    setFormMessage(ui.passwordChangeMessage, '', '');
    setFormMessage(ui.utmCaptureMessage, '', '');
    if (ui.utmCaptureSource) {
        ui.utmCaptureSource.textContent = '';
    }
    setFormMessage(ui.adminMessage, '', '');
    setFormMessage(ui.adminBranchMessage, '', '');
    setFormMessage(ui.adminDocumentMessage, '', '');
    setFormMessage(ui.adminEmailMessage, '', '');
    setFormMessage(ui.adminMailTemplateMessage, '', '');
    setFormMessage(ui.adminUsageMessage, '', '');
    setFormMessage(ui.adminMaintenanceMessage, '', '');
    setFormMessage(ui.adminLoginNoticeMessage, '', '');
    if (ui.adminEmailSource) {
        ui.adminEmailSource.textContent = '';
    }
    if (ui.adminUsageSummary) {
        ui.adminUsageSummary.textContent = 'Sin datos de actividad disponibles.';
    }
    clearAdminMaintenanceState();
    setFormMessage(ui.operatorInvoiceMessage, '', '');
    setFormMessage(ui.facturaFormMessage, '', '');
    setFormMessage(ui.invoiceModalMessage, '', '');
    setFormMessage(ui.secretariaInvoicesMessage, '', '');
    setFormMessage(ui.secretariaInvoiceDetailMessage, '', '');
    setFormMessage(ui.secretariaQuotationMessage, '', '');
    if (ui.secretariaQuotationTemplateInfo) {
        ui.secretariaQuotationTemplateInfo.textContent = '';
    }
    setFormMessage(ui.mailSendProgressStatus, '', '');
    setFormMessage(ui.invoiceEmailPromptMessage, '', '');
    setFormMessage(ui.historyEditMessage, '', '');
    setFormMessage(ui.secretaryNoMovementMessage, '', '');
    setFormMessage(ui.operatorDocAlertsMessage, '', '');
    setFormMessage(ui.dateRangeMessage, '', '');
    setFormMessage(ui.searchMatchMessage, '', '');
    initializeDocumentOptions();
    initializeOperatorBillingFeatures();
    registerRegistroRequiredFieldValidationListeners();
    registerFacturaInlineValidationListeners();

    ui.region.addEventListener('change', () => {
        populateComunas(ui.comuna, ui.region.value);
    });
    if (ui.dateRangeRegion && ui.dateRangeComuna) {
        ui.dateRangeRegion.addEventListener('change', () => {
            populateComunas(ui.dateRangeComuna, ui.dateRangeRegion.value, '', 'TODAS');
        });
    }
    if (ui.secretaryNoMovementFilterRegion && ui.secretaryNoMovementFilterComuna) {
        ui.secretaryNoMovementFilterRegion.addEventListener('change', () => {
            populateComunas(
                ui.secretaryNoMovementFilterComuna,
                ui.secretaryNoMovementFilterRegion.value,
                '',
                'TODAS'
            );
            void applySecretaryNoMovementFilters();
        });
    }
    if (ui.estado) {
        ui.estado.addEventListener('change', handleEstadoFacturaChange);
    }

    ui.loginForm.addEventListener('submit', handleLoginSubmit);
    if (ui.guestLoginBtn) {
        ui.guestLoginBtn.addEventListener('click', () => void handleGuestLogin());
    }
    ui.passwordChangeForm.addEventListener('submit', handlePasswordChangeSubmit);
    if (ui.utmCaptureForm) {
        ui.utmCaptureForm.addEventListener('submit', handleUtmCaptureSubmit);
    }
    if (ui.utmValue) {
        ui.utmValue.addEventListener('input', () => {
            applyFacturaMontoInputFormatting(ui.utmValue);
        });
        ui.utmValue.addEventListener('blur', () => {
            applyFacturaMontoInputFormatting(ui.utmValue);
        });
    }
    if (ui.utmConfirmValue) {
        ui.utmConfirmValue.addEventListener('input', () => {
            applyFacturaMontoInputFormatting(ui.utmConfirmValue);
        });
        ui.utmConfirmValue.addEventListener('blur', () => {
            applyFacturaMontoInputFormatting(ui.utmConfirmValue);
        });
    }
    ui.logoutBtn.addEventListener('click', handleLogoutClick);
    ui.adminCreateUserBtn.addEventListener('click', () => void handleAdminCreateUser());
    if (ui.adminCancelEditBtn) {
        ui.adminCancelEditBtn.addEventListener('click', handleAdminCancelEdit);
    }
    ui.adminUserForm.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
            event.preventDefault();
            void handleAdminCreateUser();
        }
    });
    ui.adminToggleUsersBtn.addEventListener('click', handleAdminToggleUsers);
    if (ui.adminToggleBranchesBtn) {
        ui.adminToggleBranchesBtn.addEventListener('click', handleAdminToggleBranches);
    }
    if (ui.adminToggleDocumentsBtn) {
        ui.adminToggleDocumentsBtn.addEventListener('click', handleAdminToggleDocuments);
    }
    if (ui.adminToggleEmailBtn) {
        ui.adminToggleEmailBtn.addEventListener('click', () => void handleAdminToggleEmail());
    }
    if (ui.adminToggleMailTemplatesBtn) {
        ui.adminToggleMailTemplatesBtn.addEventListener('click', () => void handleAdminToggleMailTemplates());
    }
    if (ui.adminToggleUsageBtn) {
        ui.adminToggleUsageBtn.addEventListener('click', () => void handleAdminToggleUsage());
    }
    if (ui.adminToggleMaintenanceBtn) {
        ui.adminToggleMaintenanceBtn.addEventListener('click', () => void handleAdminToggleMaintenance());
    }
    if (ui.adminMaintenanceRefreshBtn) {
        ui.adminMaintenanceRefreshBtn.addEventListener('click', () => void loadAdminMaintenanceStatus());
    }
    if (ui.adminMaintenanceEnableBtn) {
        ui.adminMaintenanceEnableBtn.addEventListener('click', () => void updateAdminMaintenanceMode(true));
    }
    if (ui.adminMaintenanceDisableBtn) {
        ui.adminMaintenanceDisableBtn.addEventListener('click', () => void updateAdminMaintenanceMode(false));
    }
    if (ui.adminLoginNoticeRefreshBtn) {
        ui.adminLoginNoticeRefreshBtn.addEventListener('click', () => void updateAdminLoginRouteNoticeUrl());
    }
    if (ui.adminLoginNoticeEnableBtn) {
        ui.adminLoginNoticeEnableBtn.addEventListener('click', () => void updateAdminLoginRouteNotice(true));
    }
    if (ui.adminLoginNoticeDisableBtn) {
        ui.adminLoginNoticeDisableBtn.addEventListener('click', () => void updateAdminLoginRouteNotice(false));
    }
    if (ui.secretariaViewInvoicesBtn) {
        ui.secretariaViewInvoicesBtn.addEventListener('click', handleOpenSecretariaInvoicesModal);
    }
    if (ui.secretariaQuotationBtn) {
        ui.secretariaQuotationBtn.addEventListener('click', () => void handleOpenSecretariaQuotationModal());
    }
    if (ui.secretariaEditUtmBtn) {
        ui.secretariaEditUtmBtn.addEventListener('click', () => void handleOpenMonthlyUtmEdit());
    }
    if (ui.secretariaNoMovementBtn) {
        ui.secretariaNoMovementBtn.addEventListener('click', () => void handleOpenSecretaryNoMovementModal());
    }
    if (ui.operatorDocAlertsBtn) {
        ui.operatorDocAlertsBtn.addEventListener('click', () => void handleOpenOperatorDocumentAlertsModal());
    }
    ui.adminYearCurrentBtn.addEventListener('click', () => void handleAdminYearModeClick('current'));
    ui.adminYearPastBtn.addEventListener('click', () => void handleAdminYearModeClick('historical'));
    if (ui.adminBranchSaveBtn) {
        ui.adminBranchSaveBtn.addEventListener('click', () => void handleBranchSave());
    }
    if (ui.adminBranchCancelBtn) {
        ui.adminBranchCancelBtn.addEventListener('click', handleBranchCancelEdit);
    }
    if (ui.adminBranchForm) {
        ui.adminBranchForm.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                void handleBranchSave();
            }
        });
    }
    if (ui.adminDocumentSaveBtn) {
        ui.adminDocumentSaveBtn.addEventListener('click', handleAdminDocumentSave);
    }
    if (ui.adminDocumentCancelBtn) {
        ui.adminDocumentCancelBtn.addEventListener('click', handleAdminDocumentCancel);
    }
    if (ui.adminDocumentGroupSaveBtn) {
        ui.adminDocumentGroupSaveBtn.addEventListener('click', handleAdminDocumentGroupTitlesSave);
    }
    if (ui.adminDocumentForm) {
        ui.adminDocumentForm.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                const insideGroupEditor = Boolean(event.target?.closest?.('#adminDocumentGroupTitles'));
                if (insideGroupEditor) {
                    event.preventDefault();
                    void handleAdminDocumentGroupTitlesSave();
                    return;
                }
                event.preventDefault();
                void handleAdminDocumentSave();
            }
        });
    }
    if (ui.adminEmailLoadBtn) {
        ui.adminEmailLoadBtn.addEventListener('click', () => void loadAdminEmailConfig());
    }
    if (ui.adminEmailSaveBtn) {
        ui.adminEmailSaveBtn.addEventListener('click', () => void handleAdminEmailSave());
    }
    if (ui.adminMailTemplateType) {
        ui.adminMailTemplateType.addEventListener('change', () => {
            persistCurrentAdminTemplateEditorValue();
            adminMailTemplateActiveType = normalizeMailTemplateType(ui.adminMailTemplateType.value);
            renderAdminMailTemplateEditor(adminMailTemplateActiveType);
            renderAdminMailTemplatePreview(adminMailTemplateActiveType);
        });
    }
    if (ui.adminMailTemplatePreviewBtn) {
        ui.adminMailTemplatePreviewBtn.addEventListener('click', () => {
            persistCurrentAdminTemplateEditorValue();
            renderAdminMailTemplatePreview(adminMailTemplateActiveType);
        });
    }
    if (ui.adminMailTemplateReloadBtn) {
        ui.adminMailTemplateReloadBtn.addEventListener('click', () => void loadAdminMailTemplates());
    }
    if (ui.adminMailTemplateSaveBtn) {
        ui.adminMailTemplateSaveBtn.addEventListener('click', () => void handleAdminMailTemplatesSave());
    }
    if (ui.adminMailTemplateEditor) {
        ui.adminMailTemplateEditor.addEventListener('input', () => {
            persistCurrentAdminTemplateEditorValue();
        });
    }
    if (ui.adminMailTemplatePendingLessBtn) {
        ui.adminMailTemplatePendingLessBtn.addEventListener('click', () => {
            setAdminMailTemplatePendingPreviewCount(adminMailTemplatePendingPreviewCount - 1);
            renderAdminMailTemplatePreview(adminMailTemplateActiveType);
        });
    }
    if (ui.adminMailTemplatePendingMoreBtn) {
        ui.adminMailTemplatePendingMoreBtn.addEventListener('click', () => {
            setAdminMailTemplatePendingPreviewCount(adminMailTemplatePendingPreviewCount + 1);
            renderAdminMailTemplatePreview(adminMailTemplateActiveType);
        });
    }
    if (ui.adminUsageRefreshBtn) {
        ui.adminUsageRefreshBtn.addEventListener('click', () => void loadAdminUsageStats());
    }
    if (ui.adminEmailPassword && ui.adminEmailKeepPassword) {
        ui.adminEmailPassword.addEventListener('input', () => {
            if (ui.adminEmailPassword.value.trim()) {
                ui.adminEmailKeepPassword.checked = true;
            }
        });
    }

    if (ui.operatorEmitInvoiceBtn) {
        ui.operatorEmitInvoiceBtn.addEventListener('click', handleSendCurrentInvoiceToAccountant);
    }
    if (ui.operatorQueueInvoiceBtn) {
        ui.operatorQueueInvoiceBtn.addEventListener('click', handleAddCurrentInvoiceToPending);
    }
    if (ui.operatorSendPendingBtn) {
        ui.operatorSendPendingBtn.addEventListener('click', handleSendPendingInvoicesByEmail);
    }
    if (ui.facturaEnviarContadorBtn) {
        ui.facturaEnviarContadorBtn.addEventListener('click', handleSendCurrentInvoiceToAccountant);
    }
    if (ui.facturaAgregarListaBtn) {
        ui.facturaAgregarListaBtn.addEventListener('click', handleAddCurrentInvoiceToPending);
    }
    if (ui.facturaEnviarPendientesBtn) {
        ui.facturaEnviarPendientesBtn.addEventListener('click', handleSendPendingInvoicesByEmail);
    }
    registerInvoiceAutofillBindings();
    if (ui.invoiceEmailPromptCancelBtn) {
        ui.invoiceEmailPromptCancelBtn.addEventListener('click', cancelInvoiceEmailPrompt);
    }
    if (ui.invoiceEmailPromptSendBtn) {
        ui.invoiceEmailPromptSendBtn.addEventListener('click', confirmInvoiceEmailPrompt);
    }
    if (ui.invoiceEmailPromptInput) {
        ui.invoiceEmailPromptInput.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                confirmInvoiceEmailPrompt();
                return;
            }
            if (event.key === 'Escape') {
                event.preventDefault();
                cancelInvoiceEmailPrompt();
            }
        });
    }
    if (ui.invoiceEmailPromptOverlay) {
        ui.invoiceEmailPromptOverlay.addEventListener('click', (event) => {
            if (event.target === ui.invoiceEmailPromptOverlay) {
                cancelInvoiceEmailPrompt();
            }
        });
    }
    if (ui.historyEditCancelBtn) {
        ui.historyEditCancelBtn.addEventListener('click', closeHistoryEditModal);
    }
    if (ui.historyEditForm) {
        ui.historyEditForm.addEventListener('submit', (event) => void handleHistoryEditSubmit(event));
    }
    if (ui.historyEditDeleteBtn) {
        ui.historyEditDeleteBtn.addEventListener('click', () => void handleHistoryDeleteClick());
    }
    if (ui.historyEditModalOverlay) {
        ui.historyEditModalOverlay.addEventListener('click', (event) => {
            if (event.target === ui.historyEditModalOverlay) {
                closeHistoryEditModal();
            }
        });
    }
    if (ui.secretaryNoMovementCloseBtn) {
        ui.secretaryNoMovementCloseBtn.addEventListener('click', closeSecretaryNoMovementModal);
    }
    if (ui.secretaryNoMovementOverlay) {
        ui.secretaryNoMovementOverlay.addEventListener('click', (event) => {
            if (event.target === ui.secretaryNoMovementOverlay) {
                closeSecretaryNoMovementModal();
            }
        });
    }
    if (ui.operatorDocAlertsCloseBtn) {
        ui.operatorDocAlertsCloseBtn.addEventListener('click', closeOperatorDocumentAlertsModal);
    }
    if (ui.operatorDocAlertsRefreshBtn) {
        ui.operatorDocAlertsRefreshBtn.addEventListener('click', () => void refreshOperatorDocumentAlerts({ silent: false }));
    }
    if (ui.operatorDocAlertsOverlay) {
        ui.operatorDocAlertsOverlay.addEventListener('click', (event) => {
            if (event.target === ui.operatorDocAlertsOverlay) {
                closeOperatorDocumentAlertsModal();
            }
        });
    }
    if (ui.secretaryNoMovementFilterForm) {
        ui.secretaryNoMovementFilterForm.addEventListener('submit', (event) => {
            event.preventDefault();
            void applySecretaryNoMovementFilters();
        });
    }
    if (ui.secretaryNoMovementClearBtn) {
        ui.secretaryNoMovementClearBtn.addEventListener('click', () => {
            resetSecretaryNoMovementFilters();
            void applySecretaryNoMovementFilters();
        });
    }
    if (ui.secretaryNoMovementFilterSucursal) {
        ui.secretaryNoMovementFilterSucursal.addEventListener('change', () => {
            void applySecretaryNoMovementFilters();
        });
    }
    if (ui.secretaryNoMovementFilterComuna) {
        ui.secretaryNoMovementFilterComuna.addEventListener('change', () => {
            void applySecretaryNoMovementFilters();
        });
    }
    if (ui.secretaryNoMovementFilterFrom) {
        ui.secretaryNoMovementFilterFrom.addEventListener('change', () => {
            void applySecretaryNoMovementFilters();
        });
    }
    if (ui.secretaryNoMovementFilterTo) {
        ui.secretaryNoMovementFilterTo.addEventListener('change', () => {
            void applySecretaryNoMovementFilters();
        });
    }

    ui.btnGuardar.addEventListener('click', handleGuardar);
    ui.btnModificar.addEventListener('click', handleModificar);
    ui.btnBuscar.addEventListener('click', handleBuscar);
    if (ui.btnBuscarFecha) {
        ui.btnBuscarFecha.addEventListener('click', () => void handleOpenDateRangeModal());
    }
    if (ui.btnEliminar) {
        ui.btnEliminar.addEventListener('click', handleEliminarRegistro);
    }
    if (ui.dateRangeForm) {
        ui.dateRangeForm.addEventListener('submit', (event) => void handleDateRangeSearch(event));
    }
    if (ui.dateRangeAcceptBtn) {
        ui.dateRangeAcceptBtn.addEventListener('click', () => void handleDateRangeAcceptSelection());
    }
    if (ui.dateRangeExportBtn) {
        ui.dateRangeExportBtn.addEventListener('click', () => void handleDateRangeExportWorkbook());
    }
    if (ui.dateRangeCancelBtn) {
        ui.dateRangeCancelBtn.addEventListener('click', closeDateRangeModal);
    }
    if (ui.dateRangeFirstBtn) {
        ui.dateRangeFirstBtn.addEventListener('click', () => handleDateRangeGoToPage(1));
    }
    if (ui.dateRangePrevBtn) {
        ui.dateRangePrevBtn.addEventListener('click', () => handleDateRangeGoToPage(dateRangeCurrentPage - 1));
    }
    if (ui.dateRangeNextBtn) {
        ui.dateRangeNextBtn.addEventListener('click', () => handleDateRangeGoToPage(dateRangeCurrentPage + 1));
    }
    if (ui.dateRangeLastBtn) {
        ui.dateRangeLastBtn.addEventListener('click', () => {
            const pageSize = getDateRangePageSize();
            const totalPages = Math.max(1, Math.ceil((dateRangeMatches.length || 0) / pageSize));
            handleDateRangeGoToPage(totalPages);
        });
    }
    if (ui.dateRangeModalOverlay) {
        ui.dateRangeModalOverlay.addEventListener('click', (event) => {
            if (event.target === ui.dateRangeModalOverlay) {
                closeDateRangeModal();
            }
        });
    }
    window.addEventListener('resize', scheduleDateRangeLayoutRefresh);
    if (ui.searchMatchAcceptBtn) {
        ui.searchMatchAcceptBtn.addEventListener('click', () => void handleSearchMatchAcceptSelection());
    }
    if (ui.searchMatchFirstBtn) {
        ui.searchMatchFirstBtn.addEventListener('click', () => handleSearchMatchGoToPage(1));
    }
    if (ui.searchMatchPrevBtn) {
        ui.searchMatchPrevBtn.addEventListener('click', () => handleSearchMatchGoToPage(searchMatchCurrentPage - 1));
    }
    if (ui.searchMatchNextBtn) {
        ui.searchMatchNextBtn.addEventListener('click', () => handleSearchMatchGoToPage(searchMatchCurrentPage + 1));
    }
    if (ui.searchMatchLastBtn) {
        ui.searchMatchLastBtn.addEventListener('click', () => {
            const totalPages = Math.max(1, Math.ceil((searchMatchCandidates.length || 0) / SEARCH_MATCH_PAGE_SIZE));
            handleSearchMatchGoToPage(totalPages);
        });
    }
    if (ui.searchMatchCancelBtn) {
        ui.searchMatchCancelBtn.addEventListener('click', closeSearchMatchModal);
    }
    if (ui.searchMatchModalOverlay) {
        ui.searchMatchModalOverlay.addEventListener('click', (event) => {
            if (event.target === ui.searchMatchModalOverlay) {
                closeSearchMatchModal();
            }
        });
    }
    if (ui.secretariaInvoicesCloseBtn) {
        ui.secretariaInvoicesCloseBtn.addEventListener('click', closeSecretariaInvoicesModal);
    }
    if (ui.secretariaInvoicesSearchName) {
        ui.secretariaInvoicesSearchName.addEventListener('input', () => {
            requestedInvoicesCurrentPage = 1;
            renderRequestedInvoicesHistory();
        });
        ui.secretariaInvoicesSearchName.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                requestedInvoicesCurrentPage = 1;
                renderRequestedInvoicesHistory();
            }
        });
    }
    if (ui.secretariaInvoicesSearchDate) {
        ui.secretariaInvoicesSearchDate.addEventListener('change', () => {
            requestedInvoicesCurrentPage = 1;
            renderRequestedInvoicesHistory();
        });
    }
    if (ui.secretariaInvoicesClearFiltersBtn) {
        ui.secretariaInvoicesClearFiltersBtn.addEventListener('click', () => {
            resetSecretariaInvoicesFilters();
            requestedInvoicesCurrentPage = 1;
            renderRequestedInvoicesHistory();
        });
    }
    if (ui.secretariaInvoicesFirstPageBtn) {
        ui.secretariaInvoicesFirstPageBtn.addEventListener('click', () => {
            requestedInvoicesCurrentPage = 1;
            renderRequestedInvoicesHistory();
        });
    }
    if (ui.secretariaInvoicesPrevPageBtn) {
        ui.secretariaInvoicesPrevPageBtn.addEventListener('click', () => {
            requestedInvoicesCurrentPage = Math.max(1, requestedInvoicesCurrentPage - 1);
            renderRequestedInvoicesHistory();
        });
    }
    if (ui.secretariaInvoicesNextPageBtn) {
        ui.secretariaInvoicesNextPageBtn.addEventListener('click', () => {
            requestedInvoicesCurrentPage += 1;
            renderRequestedInvoicesHistory();
        });
    }
    if (ui.secretariaInvoicesLastPageBtn) {
        ui.secretariaInvoicesLastPageBtn.addEventListener('click', () => {
            const filteredItems = getFilteredRequestedInvoicesHistory();
            const totalPages = Math.max(1, Math.ceil(filteredItems.length / REQUESTED_INVOICES_PAGE_SIZE));
            requestedInvoicesCurrentPage = totalPages;
            renderRequestedInvoicesHistory();
        });
    }
    if (ui.secretariaInvoicesDetailBtn) {
        ui.secretariaInvoicesDetailBtn.addEventListener('click', () => void handleOpenSecretariaInvoiceDetailModal());
    }
    if (ui.secretariaInvoiceDetailCloseBtn) {
        ui.secretariaInvoiceDetailCloseBtn.addEventListener('click', closeSecretariaInvoiceDetailModal);
    }
    if (ui.secretariaInvoiceDetailSaveBtn) {
        ui.secretariaInvoiceDetailSaveBtn.addEventListener('click', () => void handleSaveSecretariaInvoiceDetail());
    }
    if (ui.secretariaInvoiceDetailResendBtn) {
        ui.secretariaInvoiceDetailResendBtn.addEventListener('click', () => void handleResendSecretariaInvoiceDetail());
    }
    if (ui.secretariaQuotationForm) {
        ui.secretariaQuotationForm.addEventListener('submit', (event) => void handleSendSecretariaQuotation(event));
    }
    if (ui.secretariaQuotationRefreshBtn) {
        ui.secretariaQuotationRefreshBtn.addEventListener('click', () => void refreshSecretariaQuotationSummary());
    }
    if (ui.secretariaQuotationReplaceTemplateBtn) {
        ui.secretariaQuotationReplaceTemplateBtn.addEventListener('click', handleSecretariaQuotationReplaceTemplateClick);
    }
    if (ui.secretariaQuotationTemplateInput) {
        ui.secretariaQuotationTemplateInput.addEventListener(
            'change',
            (event) => void handleSecretariaQuotationTemplateInputChange(event)
        );
    }
    if (ui.secretariaQuotationCloseBtn) {
        ui.secretariaQuotationCloseBtn.addEventListener('click', closeSecretariaQuotationModal);
    }
    if (ui.secretariaQuotationClientName) {
        ui.secretariaQuotationClientName.addEventListener('input', () => {
            enforceNombreTitleCase(ui.secretariaQuotationClientName);
        });
        ui.secretariaQuotationClientName.addEventListener('blur', () => {
            enforceNombreTitleCase(ui.secretariaQuotationClientName);
        });
    }
    if (ui.secretariaQuotationClientEmail) {
        ui.secretariaQuotationClientEmail.addEventListener('input', () => {
            applyEmailInputNormalization(ui.secretariaQuotationClientEmail);
        });
        ui.secretariaQuotationClientEmail.addEventListener('blur', () => {
            applyEmailInputNormalization(ui.secretariaQuotationClientEmail);
        });
    }
    if (ui.secretariaInvoicesModalOverlay) {
        ui.secretariaInvoicesModalOverlay.addEventListener('click', (event) => {
            if (event.target === ui.secretariaInvoicesModalOverlay) {
                closeSecretariaInvoicesModal();
            }
        });
    }
    if (ui.secretariaInvoiceDetailModalOverlay) {
        ui.secretariaInvoiceDetailModalOverlay.addEventListener('click', (event) => {
            if (event.target === ui.secretariaInvoiceDetailModalOverlay) {
                closeSecretariaInvoiceDetailModal();
            }
        });
    }
    if (ui.secretariaQuotationModalOverlay) {
        ui.secretariaQuotationModalOverlay.addEventListener('click', (event) => {
            if (event.target === ui.secretariaQuotationModalOverlay) {
                closeSecretariaQuotationModal();
            }
        });
    }

    const searchInputIds = ['numIngreso', 'nombre', 'rut', 'rol'];
    searchInputIds.forEach((inputId) => {
        const inputElement = document.getElementById(inputId);
        if (!inputElement) {
            return;
        }

        inputElement.addEventListener('keydown', (event) => {
            if (event.key !== 'Enter') {
                return;
            }

            event.preventDefault();
            void handleBuscar();
        });
    });

    if (ui.comentario) {
        ui.comentario.addEventListener('input', () => {
            enforceComentarioTitleCase(ui.comentario);
            updateQuickCommentButtonState();
        });
        ui.comentario.addEventListener('blur', () => {
            enforceComentarioTitleCase(ui.comentario);
            updateQuickCommentButtonState();
        });
    }
    if (ui.guardarComentarioBtn) {
        ui.guardarComentarioBtn.addEventListener('click', () => void handleGuardarSoloComentario());
    }
    const nombreInput = document.getElementById('nombre');
    if (nombreInput) {
        nombreInput.addEventListener('input', () => {
            enforceNombreTitleCase(nombreInput);
        });
        nombreInput.addEventListener('blur', () => {
            enforceNombreTitleCase(nombreInput);
        });
    }
    const rutInput = document.getElementById('rut');
    if (rutInput) {
        rutInput.addEventListener('input', () => {
            applyRutInputFormatting(rutInput);
        });
        rutInput.addEventListener('blur', () => {
            applyRutInputFormatting(rutInput);
        });
    }
    const correoInput = document.getElementById('correo');
    if (correoInput) {
        correoInput.addEventListener('blur', () => {
            applyEmailInputNormalization(correoInput);
        });
    }
    if (ui.adminEmailFromEmail) {
        ui.adminEmailFromEmail.addEventListener('blur', () => {
            applyEmailInputNormalization(ui.adminEmailFromEmail);
        });
    }
    if (ui.facturaMonto) {
        ui.facturaMonto.addEventListener('input', () => {
            applyFacturaMontoInputFormatting(ui.facturaMonto);
        });
        ui.facturaMonto.addEventListener('blur', () => {
            applyFacturaMontoInputFormatting(ui.facturaMonto);
        });
    }
    if (ui.secretariaInvoiceDetailMonto) {
        ui.secretariaInvoiceDetailMonto.addEventListener('input', () => {
            applyFacturaMontoInputFormatting(ui.secretariaInvoiceDetailMonto);
        });
        ui.secretariaInvoiceDetailMonto.addEventListener('blur', () => {
            applyFacturaMontoInputFormatting(ui.secretariaInvoiceDetailMonto);
        });
    }
    if (ui.secretariaInvoiceDetailRut) {
        ui.secretariaInvoiceDetailRut.addEventListener('input', () => {
            applyRutInputFormatting(ui.secretariaInvoiceDetailRut);
        });
        ui.secretariaInvoiceDetailRut.addEventListener('blur', () => {
            applyRutInputFormatting(ui.secretariaInvoiceDetailRut);
        });
    }

    const syncPendingLoadedChangesFromForm = () => {
        syncPendingLoadedChangesInterlock({ canSearch: true, canClear: true });
    };
    ui.registroForm.addEventListener('input', syncPendingLoadedChangesFromForm);
    ui.registroForm.addEventListener('change', syncPendingLoadedChangesFromForm);

    ui.registroForm.addEventListener('reset', (event) => {
        if (suppressResetHandler) {
            return;
        }
        requestAnimationFrame(async () => {
            populateComunas(ui.comuna, '');
            clearRegistroFieldErrors();
            if (canUseHistoricalIngresoMode(currentUser) && adminIngresoMode === 'historical') {
                ui.numIngreso.value = '';
                suggestedIngresoForNewRecord = '';
            } else {
                await loadNextIngreso(ui.numIngreso, ui.formMessage);
            }
            renderHistory([]);
            loadedIngresoOriginal = '';
            clearLoadedRegistroSnapshot();
            loadedRecordUnknownDocuments = [];
            loadedComentarioOriginal = '';
            resetFacturaForm();
            syncFacturaNumeroVisibility({ clearOnHide: true });
            updateRegistroActionButtons();
        });
    });

    updateRegistroActionButtons();
    syncAdminRoleOptionsForCurrentUser();
    syncSecretariaQuotationTemplateControls();
    const redirectedToMaintenance = await enforceMaintenanceRedirectIfNeeded();
    if (redirectedToMaintenance) {
        return;
    }
    await bootstrapSession();
});

function setAuthLoginBusyState(isBusy, mode = 'credentials') {
    const busy = Boolean(isBusy);
    const loginMode = mode === 'guest' ? 'guest' : 'credentials';
    const submitBtn = ui?.loginSubmitBtn;
    const guestBtn = ui?.guestLoginBtn;

    if (ui?.loginUsername) {
        ui.loginUsername.disabled = busy;
    }
    if (ui?.loginPassword) {
        ui.loginPassword.disabled = busy;
    }

    if (submitBtn) {
        if (!submitBtn.dataset.defaultLabel) {
            submitBtn.dataset.defaultLabel = String(submitBtn.textContent || '').trim() || 'INGRESAR';
        }
        submitBtn.disabled = busy;
        submitBtn.classList.toggle('is-loading', busy && loginMode === 'credentials');
        submitBtn.setAttribute('aria-busy', busy && loginMode === 'credentials' ? 'true' : 'false');
        submitBtn.textContent = busy && loginMode === 'credentials' ? 'VALIDANDO...' : submitBtn.dataset.defaultLabel;
    }

    if (guestBtn) {
        if (!guestBtn.dataset.defaultLabel) {
            guestBtn.dataset.defaultLabel = String(guestBtn.textContent || '').trim() || 'INGRESAR COMO INVITADO';
        }
        guestBtn.disabled = busy;
        guestBtn.classList.toggle('is-loading', busy && loginMode === 'guest');
        guestBtn.setAttribute('aria-busy', busy && loginMode === 'guest' ? 'true' : 'false');
        guestBtn.textContent = busy && loginMode === 'guest' ? 'INGRESANDO...' : guestBtn.dataset.defaultLabel;
    }
}

async function handleLoginSubmit(event) {
    event.preventDefault();

    if (authLoginInProgress) {
        return;
    }

    const username = ui.loginUsername.value.trim();
    const password = ui.loginPassword.value;

    if (!username || !password) {
        setFormMessage(ui.loginMessage, 'Ingresa usuario y contrasena.', 'error');
        return;
    }

    authLoginInProgress = true;
    setAuthLoginBusyState(true, 'credentials');
    setFormMessage(ui.loginMessage, 'Validando credenciales...', '');

    try {
        const result = await apiRequest('/auth/login', {
            method: 'POST',
            body: JSON.stringify({ username, password }),
            skipAuthRedirect: true
        });

        if (!result.user) {
            throw new Error('Respuesta de login invalida.');
        }

        setAuthToken(result.token || '');
        setApiKey(result.apiWriteKey || '');
        ui.loginForm.reset();
        setFormMessage(ui.loginMessage, '', '');
        await activateAuthenticatedUser(result.user, { passwordMessage: '' });
    } catch (error) {
        if (isAuthFailureError(error)) {
            clearAuthSession();
            currentUser = null;
            setFormMessage(ui.loginMessage, 'Usuario o contrasena incorrectos.', 'error');
            return;
        }

        const fallbackMessage = normalizeText(error?.message) || 'No fue posible iniciar sesion.';
        setFormMessage(ui.loginMessage, fallbackMessage, 'error');
    } finally {
        authLoginInProgress = false;
        setAuthLoginBusyState(false);
    }
}

async function handleGuestLogin() {
    if (authLoginInProgress) {
        return;
    }

    authLoginInProgress = true;
    setAuthLoginBusyState(true, 'guest');
    setFormMessage(ui.loginMessage, 'Validando acceso invitado...', '');

    try {
        const result = await apiRequest('/auth/guest', {
            method: 'POST',
            skipAuthRedirect: true
        });

        if (!result.user) {
            throw new Error('Respuesta de login invitado invalida.');
        }

        setAuthToken(result.token || '');
        setApiKey(result.apiWriteKey || '');
        ui.loginForm.reset();
        setFormMessage(ui.loginMessage, '', '');
        await activateAuthenticatedUser(result.user, { passwordMessage: '' });
    } catch (error) {
        if (isAuthFailureError(error)) {
            clearAuthSession();
            currentUser = null;
        }
        setFormMessage(ui.loginMessage, error.message, 'error');
    } finally {
        authLoginInProgress = false;
        setAuthLoginBusyState(false);
    }
}

async function handleLogoutClick() {
    const authToken = getAuthToken();
    if (authToken) {
        try {
            await apiRequest('/auth/logout', {
                method: 'POST',
                skipAuthRedirect: true
            });
        } catch (error) {
            // Ignorar errores de red/autorizacion al cerrar sesion localmente.
        }
    }

    clearAuthSession();
    currentUser = null;
    suppressResetHandler = true;
    ui.registroForm.reset();
    suppressResetHandler = false;
    renderHistory([]);
    showAuthCard('Sesion cerrada.');
}

async function bootstrapSession() {
    const token = getAuthToken();
    if (!token) {
        clearAuthSession();
        showAuthCard('Inicia sesion para usar el sistema.');
        return;
    }

    try {
        const data = await apiRequest('/auth/me', { skipAuthRedirect: true });
        if (!data.user) {
            throw new Error('Sesion invalida.');
        }

        setApiKey(data.apiWriteKey || '');
        await activateAuthenticatedUser(data.user, {
            passwordMessage: 'Debes cambiar tu clave para usar el sistema.'
        });
    } catch (error) {
        if (isAuthFailureError(error)) {
            clearAuthSession();
            showAuthCard('Inicia sesion para usar el sistema.');
            return;
        }

        showAuthCard(error.message || 'No fue posible restaurar la sesion en este momento.');
    }
}

async function activateAuthenticatedUser(user, options = {}) {
    const passwordMessage = options.passwordMessage || '';
    if (!user) {
        throw new Error('Sesion invalida.');
    }

    currentUser = user;
    if (isPasswordChangeRequired(currentUser)) {
        showPasswordChangeCard(currentUser, passwordMessage);
        return;
    }

    const monthlyUtmStatus = await fetchMonthlyUtmStatusIfNeeded(currentUser);
    if (monthlyUtmStatus && monthlyUtmStatus.required) {
        showMonthlyUtmCaptureCard(currentUser, monthlyUtmStatus);
        return;
    }

    pendingMonthlyUtmStatus = monthlyUtmStatus || null;
    showAppForUser(currentUser);
    await loadRuntimeMailTemplates({ showErrors: false });
    await loadNextIngreso(ui.numIngreso, ui.formMessage);
    renderHistory([]);
    void maybeShowSecretaryNoMovementAlert(currentUser);
}

async function enforceMaintenanceRedirectIfNeeded() {
    try {
        const statusPayload = await apiRequest('/public/system-status', {
            skipAuthRedirect: true,
            skipPasswordRedirect: true
        });
        const maintenanceEnabled = Boolean(statusPayload?.maintenance?.enabled);
        if (!maintenanceEnabled) {
            return false;
        }

        try {
            const mePayload = await apiRequest('/auth/me', {
                skipAuthRedirect: true,
                skipPasswordRedirect: true
            });
            if (isSuperUser(mePayload?.user)) {
                return false;
            }
        } catch (error) {
            // Sin sesion valida; se aplica redireccion.
        }

        window.location.href = '../../mantenimiento.html';
        return true;
    } catch (error) {
        return false;
    }
}

async function fetchMonthlyUtmStatusIfNeeded(user) {
    if (!canUseSecretaryFeatures(user) || isGuestUser(user)) {
        pendingMonthlyUtmStatus = null;
        return null;
    }

    const status = await apiRequest('/utm/mes-actual', {
        skipAuthRedirect: true,
        skipPasswordRedirect: true
    });
    pendingMonthlyUtmStatus = status || null;
    return status || null;
}

async function handleOpenMonthlyUtmEdit() {
    if (!currentUser || !canUseSecretaryFeatures(currentUser)) {
        setFormMessage(ui.formMessage, 'Solo SECRETARIA o SUPER pueden editar la UTM mensual.', 'error');
        return;
    }

    try {
        const status = await fetchMonthlyUtmStatusIfNeeded(currentUser);
        showMonthlyUtmCaptureCard(
            currentUser,
            status || pendingMonthlyUtmStatus || null,
            '',
            { manualEdit: true }
        );
    } catch (error) {
        setFormMessage(ui.formMessage, error.message || 'No fue posible cargar la UTM mensual.', 'error');
    }
}

function showAppForUser(user) {
    const userRole = normalizeUserRole(user?.role);
    const userRoleLabel = isGuestUser(user) ? 'INVITADO' : userRole;
    const userCanAccessAdminConsole = canAccessAdminConsole(user);
    const userCanManageCatalogs = canManageAdminCatalogs(user);

    if (ui.appTopPanels) {
        ui.appTopPanels.classList.remove('hidden');
    }
    document.body.classList.add('app-authenticated');
    applyRoleLayoutClass(user);
    applyRoleActionButtonsPlacement(user);
    applyCommentsHistoryPlacement(user);
    ui.authCard.classList.add('hidden');
    ui.passwordChangeCard.classList.add('hidden');
    if (ui.utmCaptureCard) {
        ui.utmCaptureCard.classList.add('hidden');
    }
    ui.registroForm.classList.remove('hidden');
    ui.sessionUser.textContent = `Usuario: ${user.nombre} | Sucursal: ${user.sucursal || 'Sin sucursal'} | Rol: ${userRoleLabel}`;
    appViewMode = 'registro';
    setRegistroMainVisibility(true);
    exitAdminEditMode();
    exitBranchEditMode();
    configureIngresoFieldByRole(user);
    loadedComentarioOriginal = '';
    resetFacturaForm();
    syncFacturaFromMainForm();
    setFacturaContainerVisibility(canUseFacturaContainer(user));
    setComentarioVisibilityByUser(user);
    setGuestReadOnlyUiState(user);
    syncFacturaNumeroVisibility({ clearOnHide: false });
    syncOperatorDocumentAlertsButtonVisibility(user);
    void loadDateRangeBranches({ preserveSelection: false });
    setFormMessage(ui.facturaFormMessage, '', '');
    setFormMessage(ui.utmCaptureMessage, '', '');
    setFormMessage(ui.secretariaQuotationMessage, '', '');
    syncSecretariaInvoiceDetailStatusControl();
    syncSecretariaQuotationTemplateControls();
    if (ui.utmCaptureSource) {
        ui.utmCaptureSource.textContent = '';
    }
    if (!canUseSecretaryFeatures(user)) {
        secretariaQuotationSummary = null;
        secretariaQuotationTemplateMetadata = null;
        applySecretariaQuotationSummary(null);
    }
    updateRegistroActionButtons();
    if (userCanAccessAdminConsole) {
        ui.adminPanel.classList.remove('hidden');
        if (ui.adminPanelTitle) {
            if (isSuperUser(user)) {
                ui.adminPanelTitle.textContent = 'MENU SUPERUSUARIO';
            } else if (isSecretaryUser(user)) {
                ui.adminPanelTitle.textContent = 'MENU SECRETARIA';
            } else {
                ui.adminPanelTitle.textContent = 'MENU ADMINISTRADOR';
            }
        }
        if (ui.adminToggleUsersBtn) {
            ui.adminToggleUsersBtn.classList.toggle('hidden', !userCanManageCatalogs);
        }
        if (ui.adminToggleBranchesBtn) {
            ui.adminToggleBranchesBtn.classList.toggle('hidden', !userCanManageCatalogs);
        }
        if (ui.adminToggleDocumentsBtn) {
            ui.adminToggleDocumentsBtn.classList.toggle('hidden', !userCanManageCatalogs);
        }
        if (ui.adminToggleEmailBtn) {
            ui.adminToggleEmailBtn.classList.toggle('hidden', !isSuperUser(user));
        }
        if (ui.adminToggleMailTemplatesBtn) {
            ui.adminToggleMailTemplatesBtn.classList.toggle('hidden', !isSuperUser(user));
        }
        if (ui.adminToggleUsageBtn) {
            ui.adminToggleUsageBtn.classList.toggle('hidden', !isSuperUser(user));
        }
        if (ui.adminToggleMaintenanceBtn) {
            ui.adminToggleMaintenanceBtn.classList.toggle('hidden', !isSuperUser(user));
        }
        if (ui.secretariaViewInvoicesBtn) {
            ui.secretariaViewInvoicesBtn.classList.toggle('hidden', !canUseSecretaryFeatures(user));
        }
        if (ui.secretariaQuotationBtn) {
            ui.secretariaQuotationBtn.classList.toggle('hidden', !canUseSecretaryFeatures(user));
        }
        if (ui.secretariaEditUtmBtn) {
            ui.secretariaEditUtmBtn.classList.toggle('hidden', !canUseSecretaryFeatures(user));
        }
        if (ui.secretariaNoMovementBtn) {
            ui.secretariaNoMovementBtn.classList.toggle('hidden', !canUseSecretaryFeatures(user));
        }
        if (!canUseSecretaryFeatures(user)) {
            closeSecretariaInvoicesModal();
            closeSecretariaQuotationModal();
        }

        setOperatorBillingVisibility(canUseOperatorBilling(user));
        setAdminUsersVisibility(false);
        setAdminBranchesVisibility(false);
        setAdminDocumentsVisibility(false);
        setAdminEmailVisibility(false);
        setAdminMailTemplatesVisibility(false);
        setAdminUsageVisibility(false);
        setAdminMaintenanceVisibility(false);
        adminIngresoMode = 'current';
        void setAdminIngresoMode('current', { refreshIngreso: false });

        if (userCanManageCatalogs) {
            void loadAdminUsers();
            void loadBranches();
            renderAdminDocuments(configuredDocumentOptions);
            exitDocumentEditMode();
            if (isSuperUser(user)) {
                void loadAdminEmailConfig();
                void loadAdminMailTemplates();
                void loadAdminUsageStats();
                void loadAdminMaintenanceStatus({ silent: true });
            } else {
                clearAdminEmailConfigState();
                clearAdminMailTemplatesState();
                clearAdminUsageStatsState();
                clearAdminMaintenanceState();
            }
        } else {
            renderAdminUsers([]);
            renderBranches([]);
            renderAdminDocuments([]);
            populateAdminSucursalSelect([]);
            exitDocumentEditMode();
            clearAdminEmailConfigState();
            clearAdminMailTemplatesState();
            clearAdminUsageStatsState();
            clearAdminMaintenanceState();
            setFormMessage(ui.adminMessage, '', '');
            setFormMessage(ui.adminBranchMessage, '', '');
            setFormMessage(ui.adminDocumentMessage, '', '');
            setFormMessage(ui.adminEmailMessage, '', '');
            setFormMessage(ui.adminMailTemplateMessage, '', '');
            setFormMessage(ui.adminUsageMessage, '', '');
            setFormMessage(ui.adminMaintenanceMessage, '', '');
            setFormMessage(ui.adminLoginNoticeMessage, '', '');
        }
    } else {
        ui.adminPanel.classList.add('hidden');
        closeSecretariaInvoicesModal();
        closeSecretariaQuotationModal();
        if (ui.adminPanelTitle) {
            ui.adminPanelTitle.textContent = 'MENU ADMINISTRADOR';
        }
        if (ui.adminToggleUsersBtn) {
            ui.adminToggleUsersBtn.classList.remove('hidden');
        }
        if (ui.adminToggleBranchesBtn) {
            ui.adminToggleBranchesBtn.classList.remove('hidden');
        }
        if (ui.adminToggleDocumentsBtn) {
            ui.adminToggleDocumentsBtn.classList.remove('hidden');
        }
        if (ui.adminToggleEmailBtn) {
            ui.adminToggleEmailBtn.classList.add('hidden');
        }
        if (ui.adminToggleMailTemplatesBtn) {
            ui.adminToggleMailTemplatesBtn.classList.add('hidden');
        }
        if (ui.adminToggleUsageBtn) {
            ui.adminToggleUsageBtn.classList.add('hidden');
        }
        if (ui.adminToggleMaintenanceBtn) {
            ui.adminToggleMaintenanceBtn.classList.add('hidden');
        }
        if (ui.secretariaViewInvoicesBtn) {
            ui.secretariaViewInvoicesBtn.classList.add('hidden');
        }
        if (ui.secretariaQuotationBtn) {
            ui.secretariaQuotationBtn.classList.add('hidden');
        }
        if (ui.secretariaEditUtmBtn) {
            ui.secretariaEditUtmBtn.classList.add('hidden');
        }
        if (ui.secretariaNoMovementBtn) {
            ui.secretariaNoMovementBtn.classList.add('hidden');
        }

        setOperatorBillingVisibility(canUseOperatorBilling(user));
        setAdminUsersVisibility(false);
        setAdminBranchesVisibility(false);
        setAdminDocumentsVisibility(false);
        setAdminEmailVisibility(false);
        setAdminMailTemplatesVisibility(false);
        setAdminUsageVisibility(false);
        setAdminMaintenanceVisibility(false);
        renderAdminUsers([]);
        renderBranches([]);
        renderAdminDocuments([]);
        populateAdminSucursalSelect([]);
        clearAdminEmailConfigState();
        clearAdminMailTemplatesState();
        clearAdminUsageStatsState();
        clearAdminMaintenanceState();
        exitDocumentEditMode();
        if (canUseOperatorBilling(user)) {
            renderOperatorPendingInvoices();
            setFormMessage(ui.operatorInvoiceMessage, '', '');
        }
    }

    void refreshOperatorDocumentAlerts({ silent: true });
    syncAdminRoleOptionsForCurrentUser();
    void refreshInvoiceRequestsCollections();
}

function showAuthCard(message = '') {
    authLoginInProgress = false;
    setAuthLoginBusyState(false);
    if (ui.appTopPanels) {
        ui.appTopPanels.classList.add('hidden');
    }
    document.body.classList.remove('app-authenticated');
    applyRoleLayoutClass(null);
    applyRoleActionButtonsPlacement(null);
    applyCommentsHistoryPlacement(null);
    ui.registroForm.classList.add('hidden');
    ui.passwordChangeCard.classList.add('hidden');
    if (ui.utmCaptureCard) {
        ui.utmCaptureCard.classList.add('hidden');
    }
    ui.authCard.classList.remove('hidden');
    ui.passwordChangeForm.reset();
    if (ui.utmCaptureForm) {
        ui.utmCaptureForm.reset();
    }
    ui.sessionUser.textContent = 'Sesion no iniciada';
    if (message) {
        setFormMessage(ui.loginMessage, message, 'error');
    }
    ui.adminPanel.classList.add('hidden');
    if (ui.adminPanelTitle) {
        ui.adminPanelTitle.textContent = 'MENU ADMINISTRADOR';
    }
    if (ui.adminToggleUsersBtn) {
        ui.adminToggleUsersBtn.classList.remove('hidden');
    }
    if (ui.adminToggleBranchesBtn) {
        ui.adminToggleBranchesBtn.classList.remove('hidden');
    }
    if (ui.adminToggleDocumentsBtn) {
        ui.adminToggleDocumentsBtn.classList.remove('hidden');
    }
    if (ui.adminToggleEmailBtn) {
        ui.adminToggleEmailBtn.classList.add('hidden');
    }
    if (ui.adminToggleMailTemplatesBtn) {
        ui.adminToggleMailTemplatesBtn.classList.add('hidden');
    }
    if (ui.adminToggleUsageBtn) {
        ui.adminToggleUsageBtn.classList.add('hidden');
    }
    if (ui.adminToggleMaintenanceBtn) {
        ui.adminToggleMaintenanceBtn.classList.add('hidden');
    }
    if (ui.secretariaViewInvoicesBtn) {
        ui.secretariaViewInvoicesBtn.classList.add('hidden');
    }
    if (ui.secretariaQuotationBtn) {
        ui.secretariaQuotationBtn.classList.add('hidden');
    }
    if (ui.secretariaEditUtmBtn) {
        ui.secretariaEditUtmBtn.classList.add('hidden');
    }
    if (ui.secretariaNoMovementBtn) {
        ui.secretariaNoMovementBtn.classList.add('hidden');
    }
    if (ui.operatorDocAlertsBtn) {
        ui.operatorDocAlertsBtn.classList.add('hidden');
        ui.operatorDocAlertsBtn.classList.remove('has-pending');
        ui.operatorDocAlertsBtn.textContent = 'DOCS NUEVOS';
    }
    setOperatorBillingVisibility(false);
    closeDateRangeModal();
    closeSearchMatchModal();
    closeSecretariaInvoicesModal();
    closeSecretariaQuotationModal();
    closeInvoiceEmailPrompt();
    closeHistoryEditModal();
    closeSecretaryNoMovementModal();
    closeOperatorDocumentAlertsModal();
    if (ui.mailSendProgressOverlay) {
        ui.mailSendProgressOverlay.classList.add('hidden');
    }
    if (ui.mailSendProgressText) {
        ui.mailSendProgressText.textContent = 'Enviando correo...';
    }
    setFormMessage(ui.mailSendProgressStatus, '', '');
    clearAdminEmailConfigState();
    syncAdminRoleOptionsForCurrentUser();
    syncSecretariaInvoiceDetailStatusControl();
    exitAdminEditMode();
    exitBranchEditMode();
    exitDocumentEditMode();
    setAdminUsersVisibility(false);
    setAdminBranchesVisibility(false);
    setAdminDocumentsVisibility(false);
    setAdminEmailVisibility(false);
    setAdminMailTemplatesVisibility(false);
    setAdminUsageVisibility(false);
    setAdminMaintenanceVisibility(false);
    if (ui.adminEmailSection) {
        ui.adminEmailSection.classList.add('hidden');
    }
    if (ui.adminMailTemplatesSection) {
        ui.adminMailTemplatesSection.classList.add('hidden');
    }
    if (ui.adminUsageSection) {
        ui.adminUsageSection.classList.add('hidden');
    }
    if (ui.adminMaintenanceSection) {
        ui.adminMaintenanceSection.classList.add('hidden');
    }
    clearAdminMailTemplatesState();
    clearAdminUsageStatsState();
    clearAdminMaintenanceState();
    setAdminYearMessage('');
    renderAdminUsers([]);
    renderBranches([]);
    renderAdminDocuments([]);
    populateAdminSucursalSelect([]);
    availableBranches = [];
    dateRangeAvailableBranches = [];
    setFormMessage(ui.adminBranchMessage, '', '');
    setFormMessage(ui.adminDocumentMessage, '', '');
    setFormMessage(ui.adminEmailMessage, '', '');
    setFormMessage(ui.adminMailTemplateMessage, '', '');
    setFormMessage(ui.adminUsageMessage, '', '');
    setFormMessage(ui.adminMaintenanceMessage, '', '');
    setFormMessage(ui.adminLoginNoticeMessage, '', '');
    adminIngresoMode = 'current';
    appViewMode = 'registro';
    loadedRecordUnknownDocuments = [];
    loadedComentarioOriginal = '';
    setComentarioVisibilityByUser(null);
    setGuestReadOnlyUiState(null);
    resetFacturaForm();
    setFacturaContainerVisibility(false);
    setRegistroMainVisibility(true);
    configureIngresoFieldByRole(null);
    updateRegistroActionButtons();
    operatorPendingInvoices = [];
    operatorDocumentAlerts = [];
    requestedInvoicesHistory = [];
    selectedRequestedInvoiceRequestId = '';
    requestedInvoicesCurrentPage = 1;
    pendingMonthlyUtmStatus = null;
    secretariaQuotationSummary = null;
    secretariaQuotationTemplateMetadata = null;
    setFormMessage(ui.utmCaptureMessage, '', '');
    setFormMessage(ui.secretariaQuotationMessage, '', '');
    syncSecretariaQuotationTemplateControls();
    if (ui.utmCaptureSource) {
        ui.utmCaptureSource.textContent = '';
    }
    renderOperatorPendingInvoices();
    renderRequestedInvoicesHistory();
    updateSecretariaInvoiceDetailButtonState();
}

function handleSessionExpired(message) {
    clearAuthSession();
    currentUser = null;
    if (!ui) {
        return;
    }

    suppressResetHandler = true;
   ui.registroForm.reset();
    suppressResetHandler = false;
    renderHistory([]);
    loadedIngresoOriginal = '';
    clearLoadedRegistroSnapshot();
    loadedRecordUnknownDocuments = [];
    loadedComentarioOriginal = '';
    resetFacturaForm();
    closeDateRangeModal();
    closeSearchMatchModal();
    updateRegistroActionButtons();
    showAuthCard(message || 'Sesion expirada. Vuelve a iniciar sesion.');
}

function isAdminUser(user) {
    return normalizeUserRole(user?.role) === 'ADMIN';
}

function isSuperUser(user) {
    return normalizeUserRole(user?.role) === 'SUPER';
}

function isSecretaryUser(user) {
    return normalizeUserRole(user?.role) === 'SECRETARIA';
}

function isOperatorUser(user) {
    return normalizeUserRole(user?.role) === 'OPERADOR';
}

function isSupervisorUser(user) {
    return normalizeUserRole(user?.role) === 'SUPERVISOR';
}

function canUseSecretaryFeatures(user) {
    return isSecretaryUser(user) || isSuperUser(user);
}

function isGuestUser(user) {
    if (!user || typeof user !== 'object') {
        return false;
    }

    if (Boolean(user.isGuest)) {
        return true;
    }

    return String(user.username || '').trim().toLowerCase() === 'invitado';
}

function canAccessAdminConsole(user) {
    return isAdminUser(user) || canUseSecretaryFeatures(user);
}

function canManageAdminCatalogs(user) {
    return isAdminUser(user) || isSuperUser(user);
}

function canUseHistoricalIngresoMode(user) {
    return canAccessAdminConsole(user);
}

function canUseOperatorBilling(user) {
    return isOperatorUser(user) || isSuperUser(user);
}

function canUseOperatorDocumentAlerts(user) {
    return canUseOperatorBilling(user) && !isGuestUser(user);
}

function canUseFacturaContainer(user) {
    return canUseSecretaryFeatures(user);
}

function canUseFacturaNumeroField(user) {
    return isSecretaryUser(user) && !isGuestUser(user);
}

function syncFacturaNumeroVisibility(options = {}) {
    if (!ui || !ui.facturaNumeroGroup || !ui.facturaNumero) {
        return;
    }

    const clearOnHide = options.clearOnHide === true;
    const normalizedEstado = normalizeEstadoValue(ui.estado?.value);
    const shouldShow = canUseFacturaNumeroField(currentUser) && normalizedEstado === 'facturada';

    if (!shouldShow && clearOnHide) {
        ui.facturaNumero.value = '';
        clearRegistroFieldError('facturaNumero');
    }

    ui.facturaNumeroGroup.classList.toggle('hidden', !shouldShow);
    ui.facturaNumero.disabled = !shouldShow;
}

function handleEstadoFacturaChange() {
    syncFacturaNumeroVisibility({ clearOnHide: true });
}

function canUseInvoiceWorkflow(user) {
    return canUseOperatorBilling(user) || canUseFacturaContainer(user);
}

function canCreateRegistros(user) {
    return isAdminUser(user) || canUseSecretaryFeatures(user);
}

function canEditRegistros(user) {
    return !isGuestUser(user) && !isSupervisorUser(user);
}

function canCommentRegistros(user) {
    return !isGuestUser(user);
}

function normalizeUserRole(roleValue) {
    const role = String(roleValue || '').trim().toUpperCase();
    if (role === 'SUPER') {
        return 'SUPER';
    }
    if (role === 'ADMIN') {
        return 'ADMIN';
    }
    if (role === 'SECRETARIA') {
        return 'SECRETARIA';
    }
    if (role === 'SUPERVISOR') {
        return 'SUPERVISOR';
    }
    return 'OPERADOR';
}

function applyRoleLayoutClass(user) {
    const shouldUseSecretaryLayout = canUseSecretaryFeatures(user);
    document.body.classList.toggle('layout-secretaria', shouldUseSecretaryLayout);
}

function applyRoleActionButtonsPlacement(user) {
    if (!ui || !ui.registroActionsRow) {
        return;
    }

    if (ui.registroActionsSecretaryAnchor) {
        const topAnchorParent = ui.registroActionsSecretaryAnchor.parentNode;
        if (!topAnchorParent) {
            return;
        }

        if (ui.registroActionsRow.parentNode === topAnchorParent && ui.registroActionsRow.nextElementSibling === ui.registroActionsSecretaryAnchor) {
            return;
        }

        topAnchorParent.insertBefore(ui.registroActionsRow, ui.registroActionsSecretaryAnchor);
        return;
    }

    if (!ui.registroActionsDefaultAnchor) {
        return;
    }

    const defaultParent = ui.registroActionsDefaultAnchor.parentNode;
    if (!defaultParent) {
        return;
    }

    if (ui.registroActionsRow.parentNode === defaultParent && ui.registroActionsRow.nextElementSibling === ui.registroActionsDefaultAnchor) {
        return;
    }

    defaultParent.insertBefore(ui.registroActionsRow, ui.registroActionsDefaultAnchor);
}

function shouldPlaceCommentsHistoryOnSidePanel(user) {
    if (!user || typeof user !== 'object') {
        return false;
    }

    return isGuestUser(user) || isOperatorUser(user) || isAdminUser(user) || isSupervisorUser(user);
}

function applyCommentsHistoryPlacement(user) {
    if (!ui || !ui.comentariosHistoryGroup) {
        return;
    }

    const placeOnSidePanel = shouldPlaceCommentsHistoryOnSidePanel(user);
    const targetAnchor = placeOnSidePanel ? ui.comentariosHistorySideAnchor : ui.comentariosHistoryMainAnchor;
    if (!targetAnchor || !targetAnchor.parentNode) {
        return;
    }

    const targetParent = targetAnchor.parentNode;
    if (
        ui.comentariosHistoryGroup.parentNode === targetParent &&
        ui.comentariosHistoryGroup.nextElementSibling === targetAnchor
    ) {
        return;
    }

    targetParent.insertBefore(ui.comentariosHistoryGroup, targetAnchor);
}

function setComentarioVisibilityByUser(user) {
    const shouldShow = !isGuestUser(user);

    if (ui && ui.comentarioGroup) {
        ui.comentarioGroup.classList.toggle('hidden', !shouldShow);
    }
    if (ui && ui.comentariosHistoryGroup) {
        ui.comentariosHistoryGroup.classList.toggle('hidden', !shouldShow);
    }

    if (!shouldShow) {
        if (ui && ui.comentario) {
            ui.comentario.value = '';
        }
        loadedComentarioOriginal = '';
        renderHistory([]);
    }
}

function setGuestReadOnlyUiState(user) {
    const isGuest = isGuestUser(user);
    const isSupervisorRestricted = isSupervisorUser(user) && !isGuest;
    const isOperatorRestricted = isOperatorUser(user) && !isGuest;
    const hasLoadedRecord = Boolean(loadedIngresoOriginal);
    const actionButtons = [
        ui?.operatorEmitInvoiceBtn,
        ui?.operatorQueueInvoiceBtn,
        ui?.operatorSendPendingBtn,
        ui?.facturaEnviarContadorBtn,
        ui?.facturaAgregarListaBtn,
        ui?.facturaEnviarPendientesBtn
    ];

    actionButtons.forEach((button) => {
        if (!button) {
            return;
        }
        button.disabled = isGuest || !canUseInvoiceWorkflow(user);
    });
    applyInvoiceSendButtonsState();

    if (!ui || !ui.registroForm) {
        return;
    }

    const guestEditableFieldIds = hasLoadedRecord ? new Set() : new Set(['nombre', 'rut', 'rol']);
    const supervisorEditableFieldIds = hasLoadedRecord ? new Set(['comentario']) : new Set(['nombre', 'rut', 'rol']);
    const operatorSearchFieldIds = new Set(['numIngreso', 'nombre', 'rut', 'rol']);
    const controls = ui.registroForm.querySelectorAll('input, select, textarea');
    controls.forEach((control) => {
        const type = String(control.type || '').toLowerCase();
        if (type === 'button' || type === 'submit' || type === 'reset') {
            return;
        }

        if (!isGuest) {
            if (!isSupervisorRestricted && !isOperatorRestricted) {
                control.disabled = false;
                return;
            }

            const controlId = String(control.id || '');
            const isDocumentCheckbox = type === 'checkbox' && String(control.name || '') === 'documentos';
            if (isSupervisorRestricted) {
                const canEditComentario = controlId === 'comentario' && hasLoadedRecord;
                const canUseSearchField = supervisorEditableFieldIds.has(controlId) && !hasLoadedRecord;
                control.disabled = !(canEditComentario || canUseSearchField);
                return;
            }
            const canEditComentario = controlId === 'comentario' && hasLoadedRecord;
            const canEditDocuments = isDocumentCheckbox && hasLoadedRecord;
            const canUseSearchField = operatorSearchFieldIds.has(controlId) && !hasLoadedRecord;

            control.disabled = !(canEditComentario || canEditDocuments || canUseSearchField);
            return;
        }

        const controlId = String(control.id || '');
        control.disabled = !guestEditableFieldIds.has(controlId);
    });

    if ((isGuest || isSupervisorRestricted) && ui.numIngreso) {
        ui.numIngreso.readOnly = true;
    }
    if (isOperatorRestricted && ui.numIngreso) {
        ui.numIngreso.readOnly = hasLoadedRecord;
    }

    if (!isGuest && ui.comuna && ui.region) {
        const selectedComuna = ui.comuna.value;
        populateComunas(ui.comuna, ui.region.value, selectedComuna);
    }
    syncFacturaNumeroVisibility({ clearOnHide: false });
}

function isPasswordChangeRequired(user) {
    return Boolean(user && user.mustChangePassword);
}

function canUseQuickCommentSave(user) {
    return canCommentRegistros(user);
}

function updateQuickCommentButtonState() {
    if (!ui || !ui.guardarComentarioBtn) {
        return;
    }

    const canUseFeature = canUseQuickCommentSave(currentUser);
    const hasLoadedRecord = Boolean(loadedIngresoOriginal);
    const hasTypedComment = Boolean(formatComentarioTitleCase(ui.comentario?.value || '').trim());

    ui.guardarComentarioBtn.classList.toggle('hidden', !canUseFeature);
    ui.guardarComentarioBtn.disabled = !(canUseFeature && hasLoadedRecord && hasTypedComment);
}

function hasAnyRegistroMainInputValue() {
    const inputIds = ['nombre', 'rut', 'telefono', 'correo', 'region', 'comuna', 'nombrePredio', 'rol', 'numLotes', 'estado', 'comentario'];
    const hasTextValue = inputIds.some((inputId) => {
        const node = document.getElementById(inputId);
        if (!node) {
            return false;
        }
        return String(node.value || '').trim().length > 0;
    });

    if (hasTextValue) {
        return true;
    }

    if (!ui || !ui.documentosContainer) {
        return false;
    }

    return Boolean(ui.documentosContainer.querySelector('input[name="documentos"]:checked'));
}

function normalizeRegistroSnapshotText(value) {
    return String(value || '')
        .replace(/\s+/g, ' ')
        .trim();
}

function buildCurrentRegistroSnapshot() {
    const selectedDocuments = Array.from(ui?.documentosContainer?.querySelectorAll('input[name="documentos"]:checked') || [])
        .map((input) => String(input.value || '').trim())
        .filter((value) => value.length > 0);
    const retainedUnknownDocs = Array.isArray(loadedRecordUnknownDocuments) ? loadedRecordUnknownDocuments : [];
    const documentos = Array.from(new Set([...selectedDocuments, ...retainedUnknownDocs]))
        .map((value) => String(value || '').trim())
        .filter((value) => value.length > 0)
        .sort((left, right) => left.localeCompare(right, 'es'));

    return {
        numIngreso: normalizeRegistroSnapshotText(document.getElementById('numIngreso')?.value),
        nombre: normalizeRegistroSnapshotText(document.getElementById('nombre')?.value),
        rut: normalizeRegistroSnapshotText(document.getElementById('rut')?.value),
        telefono: normalizeRegistroSnapshotText(document.getElementById('telefono')?.value),
        correo: normalizeRegistroSnapshotText(document.getElementById('correo')?.value).toLowerCase(),
        region: normalizeRegistroSnapshotText(document.getElementById('region')?.value),
        comuna: normalizeRegistroSnapshotText(document.getElementById('comuna')?.value),
        nombrePredio: normalizeRegistroSnapshotText(document.getElementById('nombrePredio')?.value),
        rol: normalizeRegistroSnapshotText(document.getElementById('rol')?.value),
        numLotes: normalizeRegistroSnapshotText(document.getElementById('numLotes')?.value),
        estado: normalizeEstadoValue(document.getElementById('estado')?.value),
        factura: {
            nombreRazonSocial: normalizeRegistroSnapshotText(ui?.facturaNombreRazon?.value),
            numeroFactura: normalizeRegistroSnapshotText(ui?.facturaNumero?.value),
            rut: normalizeRegistroSnapshotText(ui?.facturaRut?.value),
            giro: normalizeRegistroSnapshotText(ui?.facturaGiro?.value),
            direccion: normalizeRegistroSnapshotText(ui?.facturaDireccion?.value),
            comuna: normalizeRegistroSnapshotText(ui?.facturaComuna?.value),
            ciudad: normalizeRegistroSnapshotText(ui?.facturaCiudad?.value),
            contacto: normalizeRegistroSnapshotText(ui?.facturaContacto?.value),
            observacion: normalizeRegistroSnapshotText(ui?.facturaObservacion?.value),
            montoFacturar: normalizeRegistroSnapshotText(ui?.facturaMonto?.value)
        },
        documentos
    };
}

function saveLoadedRegistroSnapshot() {
    loadedRegistroSnapshot = buildCurrentRegistroSnapshot();
}

function syncLoadedFacturaSnapshotWithCurrentForm() {
    if (!loadedIngresoOriginal || !loadedRegistroSnapshot) {
        return false;
    }
    const currentSnapshot = buildCurrentRegistroSnapshot();
    loadedRegistroSnapshot = {
        ...loadedRegistroSnapshot,
        factura: currentSnapshot.factura
    };
    return true;
}

function clearLoadedRegistroSnapshot() {
    loadedRegistroSnapshot = null;
}

function hasPendingLoadedRegistroChanges() {
    if (!loadedIngresoOriginal || !loadedRegistroSnapshot) {
        return false;
    }

    const currentSnapshot = buildCurrentRegistroSnapshot();
    return JSON.stringify(currentSnapshot) !== JSON.stringify(loadedRegistroSnapshot);
}

function syncPendingLoadedChangesInterlock({ canSearch = true, canClear = true, forceWarningMessage = false } = {}) {
    const hasPendingLoadedChanges = Boolean(loadedIngresoOriginal) && hasPendingLoadedRegistroChanges();

    if (ui?.btnBuscar) {
        ui.btnBuscar.disabled = !(canSearch && !hasPendingLoadedChanges);
    }
    if (ui?.btnBuscarFecha) {
        ui.btnBuscarFecha.disabled = !(canSearch && !hasPendingLoadedChanges);
    }
    if (ui?.btnLimpiar) {
        ui.btnLimpiar.disabled = !canClear;
    }

    const currentFormMessage = String(ui?.formMessage?.textContent || '').trim();
    if (hasPendingLoadedChanges) {
        if (forceWarningMessage || currentFormMessage !== PENDING_MODIFY_REQUIRED_WARNING) {
            setFormMessage(ui?.formMessage, PENDING_MODIFY_REQUIRED_WARNING, 'error');
        }
    } else if (currentFormMessage === PENDING_MODIFY_REQUIRED_WARNING) {
        setFormMessage(ui?.formMessage, '', '');
    }

    return hasPendingLoadedChanges;
}

function showPendingModifyRequiredWarningPopup() {
    window.alert(PENDING_MODIFY_REQUIRED_WARNING);
}

function updateRegistroActionButtons() {
    if (!ui || !ui.btnGuardar || !ui.btnBuscar || !ui.btnModificar || !ui.btnLimpiar) {
        return;
    }

    const hasLoadedRecord = Boolean(loadedIngresoOriginal);
    const hasFormData = hasAnyRegistroMainInputValue();
    const userCanDelete = canManageAdminCatalogs(currentUser);
    const userCanCreate = canCreateRegistros(currentUser);
    const userCanModify = canEditRegistros(currentUser);
    const canSave = userCanCreate;
    const canSearch = true;
    const canClear = true;
    const canModify = userCanModify && hasLoadedRecord && hasFormData;
    const canDelete = hasLoadedRecord && userCanDelete;

    ui.btnGuardar.classList.toggle('hidden', !userCanCreate || hasLoadedRecord);
    ui.btnModificar.classList.toggle('hidden', !userCanModify);
    ui.btnGuardar.disabled = !canSave;
    ui.btnModificar.disabled = !canModify;
    if (ui.btnEliminar) {
        ui.btnEliminar.classList.toggle('hidden', !userCanDelete);
        ui.btnEliminar.disabled = !canDelete;
    }
    ui.btnLimpiar.disabled = !canClear;
    setGuestReadOnlyUiState(currentUser);
    syncPendingLoadedChangesInterlock({ canSearch, canClear });
    updateQuickCommentButtonState();
}

function showExistingRecordSaveWarningPopup() {
    window.alert(EXISTING_RECORD_SAVE_WARNING);
}

function initializeOperatorBillingFeatures() {
    operatorPendingInvoices = [];
    requestedInvoicesHistory = [];
    selectedRequestedInvoiceRequestId = '';
    requestedInvoicesCurrentPage = 1;
    renderOperatorPendingInvoices();
    renderRequestedInvoicesHistory();
    updateSecretariaInvoiceDetailButtonState();
    syncFacturaFromMainForm();
}

async function refreshInvoiceRequestsCollections() {
    if (!currentUser) {
        return;
    }

    if (canUseOperatorBilling(currentUser) || canUseFacturaContainer(currentUser)) {
        await refreshOperatorPendingInvoices();
    }

    if (canUseSecretaryFeatures(currentUser)) {
        await refreshRequestedInvoicesHistory();
    }
}

function setOperatorBillingVisibility(isVisible) {
    if (!ui || !ui.operatorBillingPanel) {
        return;
    }

    ui.operatorBillingPanel.classList.toggle('hidden', !isVisible);
    if (!isVisible) {
        closeInvoiceModal();
    }
}

function setFacturaContainerVisibility(isVisible) {
    if (!ui || !ui.facturaDataSection) {
        return;
    }

    ui.facturaDataSection.classList.toggle('hidden', !isVisible);
}

function normalizeInvoiceText(value, maxLength = 255) {
    const normalized = String(value || '')
        .replace(/\s+/g, ' ')
        .trim();

    if (!maxLength || maxLength <= 0) {
        return normalized;
    }

    return normalized.slice(0, maxLength);
}

function isValidEmail(value) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function normalizeEmailValue(value) {
    return String(value || '').trim().toLowerCase();
}

function normalizeRutValue(value) {
    return String(value || '').replace(/[^0-9kK]/g, '').toUpperCase();
}

function normalizeRutTypingValue(value) {
    const uppercaseValue = String(value || '').toUpperCase();
    const digitsOnly = uppercaseValue.replace(/[^0-9]/g, '');
    const includesVerifierK = uppercaseValue.includes('K');

    if (includesVerifierK) {
        const rutBody = digitsOnly.slice(0, 8);
        return rutBody ? `${rutBody}K` : 'K';
    }

    return digitsOnly.slice(0, 9);
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
        return 'K';
    }

    return String(remainder);
}

function formatRutBodyWithDots(rutBodyDigits, verifierDigit) {
    const normalizedBody = String(rutBodyDigits || '').replace(/^0+(?=\d)/, '');
    const safeBody = normalizedBody || '0';
    const bodyWithDots = safeBody.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return `${bodyWithDots}-${String(verifierDigit || '').toUpperCase()}`;
}

function formatRutInputProgressive(value) {
    const normalized = normalizeRutTypingValue(value);
    if (!normalized) {
        return '';
    }

    if (normalized.length === 1) {
        return normalized;
    }

    const rutBody = normalized.slice(0, -1);
    const verifierDigit = normalized.slice(-1);
    const bodyWithDots = rutBody.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return `${bodyWithDots}-${verifierDigit}`;
}

function parseChileanRut(value) {
    const normalized = normalizeRutValue(value);
    if (normalized.length < 2) {
        return { isValid: false, normalized: '', formatted: '' };
    }

    const rutBodyDigits = normalized.slice(0, -1);
    const verifierDigit = normalized.slice(-1);
    if (!/^\d+$/.test(rutBodyDigits)) {
        return { isValid: false, normalized, formatted: '' };
    }

    const expectedVerifier = calculateRutVerifierDigit(rutBodyDigits);
    const isValid = verifierDigit === expectedVerifier;
    return {
        isValid,
        normalized,
        formatted: isValid ? formatRutBodyWithDots(rutBodyDigits, verifierDigit) : ''
    };
}

function applyRutInputFormatting(inputElement) {
    if (!inputElement) {
        return;
    }

    const rawValue = String(inputElement.value || '');
    const cursorStart = typeof inputElement.selectionStart === 'number' ? inputElement.selectionStart : null;
    if (!rawValue) {
        inputElement.value = '';
        return;
    }

    const formattedValue = formatRutInputProgressive(rawValue);
    inputElement.value = formattedValue;

    if (cursorStart === null || document.activeElement !== inputElement) {
        return;
    }

    const valueBeforeCursor = rawValue.slice(0, cursorStart);
    const formattedBeforeCursor = formatRutInputProgressive(valueBeforeCursor);
    const nextCursor = Math.min(formattedBeforeCursor.length, formattedValue.length);
    inputElement.setSelectionRange(nextCursor, nextCursor);
}

function applyEmailInputNormalization(inputElement) {
    if (!inputElement) {
        return;
    }

    inputElement.value = normalizeEmailValue(inputElement.value);
}

function validateAndNormalizeRutCorreo(data) {
    const parsedRut = parseChileanRut(data.rut);
    if (!parsedRut.isValid) {
        return {
            error: 'RUT invalido. Usa formato 12.345.678-5.',
            errorField: 'rut',
            data: null
        };
    }

    const normalizedCorreo = normalizeEmailValue(data.correo);
    if (normalizedCorreo && !isValidEmail(normalizedCorreo)) {
        return {
            error: 'Correo invalido. Usa formato nombre@proveedor.cl.',
            errorField: 'correo',
            data: null
        };
    }

    return {
        error: '',
        errorField: '',
        data: {
            ...data,
            rut: parsedRut.formatted,
            correo: normalizedCorreo
        }
    };
}

function generateInvoiceDraftId() {
    return `inv-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
}

function closeInvoiceModal() {
    if (ui && ui.invoiceModalOverlay) {
        ui.invoiceModalOverlay.classList.add('hidden');
    }
    invoiceModalMode = '';
    setFormMessage(ui.invoiceModalMessage, '', '');
}

function closeSecretariaInvoicesModal() {
    if (ui && ui.secretariaInvoicesModalOverlay) {
        ui.secretariaInvoicesModalOverlay.classList.add('hidden');
    }
    closeSecretariaInvoiceDetailModal();
    selectedRequestedInvoiceRequestId = '';
    requestedInvoicesCurrentPage = 1;
    updateSecretariaInvoiceDetailButtonState();
    setFormMessage(ui?.secretariaInvoicesMessage, '', '');
}

function closeSecretariaInvoiceDetailModal() {
    if (ui && ui.secretariaInvoiceDetailModalOverlay) {
        ui.secretariaInvoiceDetailModalOverlay.classList.add('hidden');
    }
    setFormMessage(ui?.secretariaInvoiceDetailMessage, '', '');
}

function initializeSecretariaInvoiceDetailLocationSelectors() {
    if (!ui || !ui.secretariaInvoiceDetailCiudad || !ui.secretariaInvoiceDetailComuna) {
        return;
    }

    populateRegiones(ui.secretariaInvoiceDetailCiudad, 'Seleccione ciudad');
    populateComunas(ui.secretariaInvoiceDetailComuna, '');
    ui.secretariaInvoiceDetailCiudad.addEventListener('change', () => {
        populateComunas(ui.secretariaInvoiceDetailComuna, ui.secretariaInvoiceDetailCiudad.value);
    });
}

function updateSecretariaInvoiceDetailButtonState() {
    if (!ui || !ui.secretariaInvoicesDetailBtn) {
        return;
    }

    const hasSelection = Boolean(selectedRequestedInvoiceRequestId);
    ui.secretariaInvoicesDetailBtn.disabled = !hasSelection;
}

function syncSecretariaInvoiceDetailStatusControl() {
    if (!ui || !ui.secretariaInvoiceDetailEstado) {
        return;
    }

    const canManageStatus = isSuperUser(currentUser);
    ui.secretariaInvoiceDetailEstado.disabled = !canManageStatus;
}

function resetSecretariaInvoicesFilters() {
    if (!ui) {
        return;
    }

    if (ui.secretariaInvoicesSearchName) {
        ui.secretariaInvoicesSearchName.value = '';
    }
    if (ui.secretariaInvoicesSearchDate) {
        ui.secretariaInvoicesSearchDate.value = '';
    }
    requestedInvoicesCurrentPage = 1;
}

function getSecretariaInvoicesFilterValues() {
    const nameQuery = normalizeInvoiceText(ui?.secretariaInvoicesSearchName?.value, 255).toLocaleLowerCase('es-CL');
    const dateQuery = String(ui?.secretariaInvoicesSearchDate?.value || '').trim();
    return {
        nameQuery,
        dateQuery,
        hasFilters: Boolean(nameQuery || dateQuery)
    };
}

function getFilteredRequestedInvoicesHistory() {
    const source = Array.isArray(requestedInvoicesHistory) ? requestedInvoicesHistory : [];
    const filters = getSecretariaInvoicesFilterValues();
    if (!filters.hasFilters) {
        return source;
    }

    return source.filter((item) => {
        const itemName = normalizeInvoiceText(item?.nombreRazonSocial, 255).toLocaleLowerCase('es-CL');
        const matchesName = !filters.nameQuery || itemName.includes(filters.nameQuery);
        let matchesDate = true;
        if (filters.dateQuery) {
            const parsedDate = new Date(item?.requestedAt || item?.createdAt || '');
            const itemDate = Number.isNaN(parsedDate.getTime()) ? '' : toLocalDateInputValue(parsedDate);
            matchesDate = itemDate === filters.dateQuery;
        }
        return matchesName && matchesDate;
    });
}

function getRequestedInvoicesPaginationInfo(filteredItems) {
    const totalItems = Array.isArray(filteredItems) ? filteredItems.length : 0;
    const totalPages = Math.max(1, Math.ceil(totalItems / REQUESTED_INVOICES_PAGE_SIZE));
    requestedInvoicesCurrentPage = Math.min(Math.max(1, requestedInvoicesCurrentPage), totalPages);
    const startIndex = totalItems > 0 ? (requestedInvoicesCurrentPage - 1) * REQUESTED_INVOICES_PAGE_SIZE : 0;
    const endIndexExclusive = totalItems > 0 ? Math.min(startIndex + REQUESTED_INVOICES_PAGE_SIZE, totalItems) : 0;
    const pageItems = totalItems > 0 ? filteredItems.slice(startIndex, endIndexExclusive) : [];

    return {
        totalItems,
        totalPages,
        currentPage: requestedInvoicesCurrentPage,
        startIndex,
        endIndexExclusive,
        startItemNumber: totalItems > 0 ? startIndex + 1 : 0,
        endItemNumber: endIndexExclusive,
        pageItems
    };
}

function renderRequestedInvoicesPaginationControls(paginationInfo) {
    if (
        !ui ||
        !ui.secretariaInvoicesPagination ||
        !ui.secretariaInvoicesFirstPageBtn ||
        !ui.secretariaInvoicesPrevPageBtn ||
        !ui.secretariaInvoicesNextPageBtn ||
        !ui.secretariaInvoicesLastPageBtn
    ) {
        return;
    }

    const info = paginationInfo || {
        totalItems: 0,
        totalPages: 1,
        currentPage: 1,
        startItemNumber: 0,
        endItemNumber: 0
    };
    const hasItems = Number(info.totalItems) > 0;
    const hasMultiplePages = Number(info.totalItems) > REQUESTED_INVOICES_PAGE_SIZE;
    ui.secretariaInvoicesPagination.classList.toggle('hidden', !hasMultiplePages);

    ui.secretariaInvoicesFirstPageBtn.disabled = !hasItems || info.currentPage <= 1;
    ui.secretariaInvoicesPrevPageBtn.disabled = !hasItems || info.currentPage <= 1;
    ui.secretariaInvoicesNextPageBtn.disabled = !hasItems || info.currentPage >= info.totalPages;
    ui.secretariaInvoicesLastPageBtn.disabled = !hasItems || info.currentPage >= info.totalPages;
    if (ui.secretariaInvoicesPageInfo) {
        ui.secretariaInvoicesPageInfo.textContent = hasItems
            ? `Grupo ${info.currentPage} de ${info.totalPages} | Mostrando ${info.startItemNumber}-${info.endItemNumber} de ${info.totalItems}`
            : 'Grupo 0 de 0';
    }
}

function setSelectedRequestedInvoiceRequest(requestId) {
    selectedRequestedInvoiceRequestId = normalizeInvoiceText(requestId, 80);
    if (selectedRequestedInvoiceRequestId) {
        const filteredItems = getFilteredRequestedInvoicesHistory();
        const selectedIndex = filteredItems.findIndex(
            (item) => String(item.requestId || '') === selectedRequestedInvoiceRequestId
        );
        if (selectedIndex >= 0) {
            requestedInvoicesCurrentPage = Math.floor(selectedIndex / REQUESTED_INVOICES_PAGE_SIZE) + 1;
        }
    }
    updateSecretariaInvoiceDetailButtonState();
    renderRequestedInvoicesHistory();
}

function getSelectedRequestedInvoiceRequest() {
    if (!selectedRequestedInvoiceRequestId) {
        return null;
    }

    const selected = requestedInvoicesHistory.find((item) => String(item.requestId || '') === selectedRequestedInvoiceRequestId);
    return selected || null;
}

function applySecretariaInvoiceDetailForm(item) {
    if (!ui || !item) {
        return;
    }

    if (ui.secretariaInvoiceDetailRequestId) {
        ui.secretariaInvoiceDetailRequestId.value = String(item.requestId || '');
    }
    if (ui.secretariaInvoiceDetailNumIngreso) {
        ui.secretariaInvoiceDetailNumIngreso.value = item.numIngreso || '';
    }
    syncSecretariaInvoiceDetailStatusControl();
    if (ui.secretariaInvoiceDetailEstado) {
        const normalizedStatus = normalizeInvoiceRequestStatus(item.estado) || 'PENDIENTE';
        ui.secretariaInvoiceDetailEstado.value = normalizedStatus;
    }
    if (ui.secretariaInvoiceDetailNombre) {
        ui.secretariaInvoiceDetailNombre.value = item.nombreRazonSocial || '';
    }
    if (ui.secretariaInvoiceDetailRut) {
        ui.secretariaInvoiceDetailRut.value = item.rut || '';
    }
    if (ui.secretariaInvoiceDetailMonto) {
        const normalizedAmount = normalizeInvoiceAmount(item.montoFacturar);
        ui.secretariaInvoiceDetailMonto.value = normalizedAmount ? formatInvoiceAmountForInput(normalizedAmount) : '';
    }
    if (ui.secretariaInvoiceDetailGiro) {
        ui.secretariaInvoiceDetailGiro.value = item.giro || '';
    }
    if (ui.secretariaInvoiceDetailDireccion) {
        ui.secretariaInvoiceDetailDireccion.value = item.direccion || '';
    }
    const selectedRegion = normalizeInvoiceText(item.ciudad, 120);
    const selectedComuna = normalizeInvoiceText(item.comuna, 120);
    if (ui.secretariaInvoiceDetailCiudad) {
        populateRegiones(ui.secretariaInvoiceDetailCiudad, 'Seleccione ciudad');
        setSelectValueWithFallback(ui.secretariaInvoiceDetailCiudad, selectedRegion);
    }
    if (ui.secretariaInvoiceDetailComuna) {
        populateComunas(ui.secretariaInvoiceDetailComuna, selectedRegion, selectedComuna);
        if (selectedComuna && ui.secretariaInvoiceDetailComuna.value !== selectedComuna) {
            setSelectValueWithFallback(ui.secretariaInvoiceDetailComuna, selectedComuna);
            ui.secretariaInvoiceDetailComuna.disabled = false;
        }
    }
    if (ui.secretariaInvoiceDetailContacto) {
        ui.secretariaInvoiceDetailContacto.value = item.contacto || '';
    }
    if (ui.secretariaInvoiceDetailDestinationEmail) {
        ui.secretariaInvoiceDetailDestinationEmail.value = item.destinationEmail || '';
    }
    if (ui.secretariaInvoiceDetailObservacion) {
        ui.secretariaInvoiceDetailObservacion.value = item.observacion || '';
    }
}

function collectSecretariaInvoiceDetailDraft() {
    const selected = getSelectedRequestedInvoiceRequest();
    const selectedStatus = normalizeInvoiceRequestStatus(ui?.secretariaInvoiceDetailEstado?.value);
    const fallbackStatus = normalizeInvoiceRequestStatus(selected?.estado) || 'PENDIENTE';
    return {
        id: selected?.id || '',
        requestId: selected?.requestId || '',
        createdAt: selected?.requestedAt || new Date().toISOString(),
        createdBy: selected?.requestedBy || '',
        sucursal: selected?.requestedBySucursal || '',
        numIngreso: normalizeInvoiceText(ui?.secretariaInvoiceDetailNumIngreso?.value, 20),
        nombreRazonSocial: normalizeInvoiceText(ui?.secretariaInvoiceDetailNombre?.value, 255),
        rut: normalizeInvoiceText(ui?.secretariaInvoiceDetailRut?.value, 20),
        giro: normalizeInvoiceText(ui?.secretariaInvoiceDetailGiro?.value, 255),
        direccion: normalizeInvoiceText(ui?.secretariaInvoiceDetailDireccion?.value, 255),
        comuna: normalizeInvoiceText(ui?.secretariaInvoiceDetailComuna?.value, 120),
        ciudad: normalizeInvoiceText(ui?.secretariaInvoiceDetailCiudad?.value, 120),
        contacto: normalizeInvoiceText(ui?.secretariaInvoiceDetailContacto?.value, 255),
        observacion: normalizeInvoiceText(ui?.secretariaInvoiceDetailObservacion?.value, 1000),
        montoFacturar: normalizeInvoiceAmount(ui?.secretariaInvoiceDetailMonto?.value),
        destinationEmail: normalizeInvoiceText(ui?.secretariaInvoiceDetailDestinationEmail?.value, 255).toLowerCase(),
        estado: selectedStatus || fallbackStatus
    };
}

async function persistSecretariaInvoiceDetail(options = {}) {
    const forceSent = options.forceSent === true;
    const statusOverride = normalizeInvoiceRequestStatus(options.estado);
    const targetEmail = normalizeInvoiceText(options.destinationEmail, 255).toLowerCase();
    const draft = collectSecretariaInvoiceDetailDraft();
    const parsed = validateInvoiceDraft(draft, { requireAmount: true });
    if (parsed.error) {
        return { error: parsed.error, data: null, response: null };
    }

    const canManageStatus = isSuperUser(currentUser);
    const requestedStatus = statusOverride || normalizeInvoiceRequestStatus(draft.estado) || 'PENDIENTE';
    const finalState = forceSent ? 'ENVIADA' : canManageStatus ? requestedStatus : 'PENDIENTE';
    const finalDestinationEmail = finalState === 'ENVIADA' ? targetEmail || draft.destinationEmail : '';
    if (finalState === 'ENVIADA' && !finalDestinationEmail) {
        return { error: 'Debes indicar el correo del contador para guardar en estado ENVIADA.', data: null, response: null };
    }

    try {
        const response = await registerInvoiceRequestInDatabase(parsed.data, {
            estado: finalState,
            destinationEmail: finalDestinationEmail
        });
        return {
            error: '',
            data: {
                ...parsed.data,
                estado: finalState,
                destinationEmail: finalDestinationEmail
            },
            response
        };
    } catch (error) {
        return { error: error.message, data: null, response: null };
    }
}

async function handleOpenSecretariaInvoiceDetailModal() {
    if (!canUseSecretaryFeatures(currentUser) || !ui || !ui.secretariaInvoiceDetailModalOverlay) {
        return;
    }

    const selected = getSelectedRequestedInvoiceRequest();
    if (!selected) {
        setFormMessage(ui.secretariaInvoicesMessage, 'Selecciona una factura para ver detalle.', 'error');
        return;
    }

    applySecretariaInvoiceDetailForm(selected);
    setFormMessage(ui.secretariaInvoiceDetailMessage, '', '');
    ui.secretariaInvoiceDetailModalOverlay.classList.remove('hidden');
}

async function handleSaveSecretariaInvoiceDetail() {
    if (!canUseSecretaryFeatures(currentUser) || !ui) {
        return;
    }

    const selected = getSelectedRequestedInvoiceRequest();
    if (!selected) {
        setFormMessage(ui.secretariaInvoicesMessage, 'Selecciona una factura para editar.', 'error');
        return;
    }

    const saveResult = await persistSecretariaInvoiceDetail();
    if (saveResult.error) {
        setFormMessage(ui.secretariaInvoiceDetailMessage, saveResult.error, 'error');
        return;
    }

    await refreshRequestedInvoicesHistory({
        showErrors: true,
        messageElement: ui.secretariaInvoiceDetailMessage
    });
    await refreshOperatorPendingInvoices();
    const updatedById = getSelectedRequestedInvoiceRequest();
    const updatedByIngreso = requestedInvoicesHistory.find((item) => item.numIngreso === saveResult.data.numIngreso) || null;
    const updated = updatedById || updatedByIngreso;
    if (updated) {
        setSelectedRequestedInvoiceRequest(updated.requestId);
        applySecretariaInvoiceDetailForm(updated);
    }
    setFormMessage(
        ui.secretariaInvoicesMessage,
        saveResult.response?.message || 'Datos de la factura actualizados correctamente.',
        'success'
    );
    closeSecretariaInvoiceDetailModal();
}

async function handleResendSecretariaInvoiceDetail() {
    if (!canUseSecretaryFeatures(currentUser) || !ui) {
        return;
    }

    const currentEmail = normalizeInvoiceText(ui.secretariaInvoiceDetailDestinationEmail?.value, 255).toLowerCase();
    const destinationEmail = await promptAccountantEmail(currentEmail);
    if (destinationEmail === null) {
        return;
    }
    if (destinationEmail === '') {
        setFormMessage(ui.secretariaInvoiceDetailMessage, 'Debes ingresar el correo del contador.', 'error');
        return;
    }
    if (destinationEmail === 'INVALID_EMAIL') {
        setFormMessage(ui.secretariaInvoiceDetailMessage, 'El correo del contador no es valido.', 'error');
        return;
    }
    if (ui.secretariaInvoiceDetailDestinationEmail) {
        ui.secretariaInvoiceDetailDestinationEmail.value = destinationEmail;
    }

    const saveResult = await persistSecretariaInvoiceDetail({
        forceSent: true,
        destinationEmail
    });
    if (saveResult.error) {
        setFormMessage(ui.secretariaInvoiceDetailMessage, saveResult.error, 'error');
        return;
    }

    const subject = buildMailSubject(
        `Solicitud factura ${saveResult.data.numIngreso || ''} ${saveResult.data.nombreRazonSocial || ''}`.trim()
    );
    showMailSendProgress('Enviando factura por correo...');
    try {
        await sendInvoiceEmailViaSmtp({
            destinationEmail,
            subject,
            textBody: buildSingleInvoiceEmailBody(saveResult.data),
            htmlBody: buildSingleInvoiceEmailHtml(saveResult.data)
        });
        await hideMailSendProgress('Correo enviado correctamente.', 'success');
    } catch (error) {
        await hideMailSendProgress('No fue posible enviar el correo.', 'error', {
            minDurationMs: MAIL_SEND_PROGRESS_MIN_MS,
            finalDurationMs: 1200
        });
        setFormMessage(ui.secretariaInvoiceDetailMessage, error.message, 'error');
        return;
    }

    await refreshRequestedInvoicesHistory({
        showErrors: true,
        messageElement: ui.secretariaInvoiceDetailMessage
    });
    await refreshOperatorPendingInvoices();
    const updated = requestedInvoicesHistory.find((item) => item.numIngreso === saveResult.data.numIngreso) || null;
    if (updated) {
        setSelectedRequestedInvoiceRequest(updated.requestId);
        applySecretariaInvoiceDetailForm(updated);
    }

    setFormMessage(
        ui.secretariaInvoiceDetailMessage,
        saveResult.response?.message || `Solicitud reenviada al contador (${destinationEmail}).`,
        'success'
    );
}

async function handleOpenSecretariaInvoicesModal() {
    if (!canUseSecretaryFeatures(currentUser) || !ui || !ui.secretariaInvoicesModalOverlay) {
        return;
    }

    closeSecretariaQuotationModal();
    resetSecretariaInvoicesFilters();
    selectedRequestedInvoiceRequestId = '';
    requestedInvoicesCurrentPage = 1;
    updateSecretariaInvoiceDetailButtonState();
    renderRequestedInvoicesHistory();
    setFormMessage(ui.secretariaInvoicesMessage, '', '');
    ui.secretariaInvoicesModalOverlay.classList.remove('hidden');
    await refreshRequestedInvoicesHistory({
        showErrors: true,
        messageElement: ui.secretariaInvoicesMessage
    });
}

function closeSecretariaQuotationModal() {
    if (ui && ui.secretariaQuotationModalOverlay) {
        ui.secretariaQuotationModalOverlay.classList.add('hidden');
    }
    if (ui && ui.secretariaQuotationTemplateInput) {
        ui.secretariaQuotationTemplateInput.value = '';
    }
    setFormMessage(ui?.secretariaQuotationMessage, '', '');
}

function syncSecretariaQuotationTemplateControls() {
    if (!ui) {
        return;
    }

    const canManageTemplate = isSuperUser(currentUser);
    if (ui.secretariaQuotationReplaceTemplateBtn) {
        ui.secretariaQuotationReplaceTemplateBtn.classList.toggle('hidden', !canManageTemplate);
        ui.secretariaQuotationReplaceTemplateBtn.disabled = false;
    }
    if (ui.secretariaQuotationTemplateInput) {
        ui.secretariaQuotationTemplateInput.value = '';
    }
    if (!canManageTemplate) {
        secretariaQuotationTemplateMetadata = null;
        if (ui.secretariaQuotationTemplateInfo) {
            ui.secretariaQuotationTemplateInfo.textContent = '';
        }
    }
}

function formatSecretariaQuotationFileSize(sizeBytes) {
    const parsed = Number(sizeBytes);
    if (!Number.isFinite(parsed) || parsed <= 0) {
        return '';
    }
    if (parsed < 1024) {
        return `${Math.round(parsed)} B`;
    }
    if (parsed < 1024 * 1024) {
        return `${(parsed / 1024).toFixed(1)} KB`;
    }
    return `${(parsed / (1024 * 1024)).toFixed(2)} MB`;
}

function applySecretariaQuotationTemplateMetadata(metadata) {
    secretariaQuotationTemplateMetadata = metadata && typeof metadata === 'object' ? metadata : null;
    if (!ui || !ui.secretariaQuotationTemplateInfo) {
        return;
    }
    if (!isSuperUser(currentUser)) {
        ui.secretariaQuotationTemplateInfo.textContent = '';
        return;
    }

    const fileName = normalizeInvoiceText(secretariaQuotationTemplateMetadata?.fileName, 255);
    if (!fileName) {
        ui.secretariaQuotationTemplateInfo.textContent = 'No se encontro un PDF base cargado.';
        return;
    }

    const updatedAtRaw = normalizeInvoiceText(secretariaQuotationTemplateMetadata?.updatedAt, 80);
    const updatedAtText = updatedAtRaw ? formatDateTime(updatedAtRaw) : '';
    const details = [`PDF base actual: ${fileName}`];
    const sizeLabel = formatSecretariaQuotationFileSize(secretariaQuotationTemplateMetadata?.sizeBytes);
    if (sizeLabel) {
        details.push(`Tamano: ${sizeLabel}`);
    }
    if (updatedAtText) {
        details.push(`Actualizado: ${updatedAtText}`);
    }

    ui.secretariaQuotationTemplateInfo.textContent = details.join(' | ');
}

async function loadSecretariaQuotationTemplateMetadata(options = {}) {
    if (!ui) {
        return null;
    }
    if (!isSuperUser(currentUser)) {
        applySecretariaQuotationTemplateMetadata(null);
        return null;
    }

    const showErrors = options.showErrors === true;
    try {
        const result = await apiRequest('/admin/cotizaciones/documento');
        applySecretariaQuotationTemplateMetadata(result);
        return result;
    } catch (error) {
        applySecretariaQuotationTemplateMetadata(null);
        if (showErrors) {
            setFormMessage(ui.secretariaQuotationMessage, error.message, 'error');
        } else if (ui.secretariaQuotationTemplateInfo) {
            ui.secretariaQuotationTemplateInfo.textContent = 'No fue posible consultar el PDF base de cotizacion.';
        }
        return null;
    }
}

function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
        if (!(file instanceof File)) {
            reject(createApiError('Debes seleccionar un archivo PDF valido.'));
            return;
        }

        const reader = new FileReader();
        reader.onload = () => {
            resolve(String(reader.result || ''));
        };
        reader.onerror = () => {
            reject(createApiError('No fue posible leer el archivo PDF seleccionado.'));
        };
        reader.readAsDataURL(file);
    });
}

function handleSecretariaQuotationReplaceTemplateClick() {
    if (!isSuperUser(currentUser) || !ui || !ui.secretariaQuotationTemplateInput) {
        return;
    }

    ui.secretariaQuotationTemplateInput.value = '';
    setFormMessage(ui.secretariaQuotationMessage, '', '');
    ui.secretariaQuotationTemplateInput.click();
}

async function handleSecretariaQuotationTemplateInputChange(event) {
    if (!isSuperUser(currentUser) || !ui) {
        return;
    }

    const inputElement = event?.target;
    const selectedFile = inputElement?.files?.[0];
    if (!selectedFile) {
        return;
    }

    const fileName = String(selectedFile.name || '').trim();
    const maxFileBytes = 10 * 1024 * 1024;
    if (!/\.pdf$/i.test(fileName)) {
        setFormMessage(ui.secretariaQuotationMessage, 'Solo puedes seleccionar archivos PDF (.pdf).', 'error');
        if (inputElement) {
            inputElement.value = '';
        }
        return;
    }
    if (Number(selectedFile.size || 0) > maxFileBytes) {
        setFormMessage(ui.secretariaQuotationMessage, 'El PDF excede el tamano maximo permitido de 10 MB.', 'error');
        if (inputElement) {
            inputElement.value = '';
        }
        return;
    }

    if (ui.secretariaQuotationReplaceTemplateBtn) {
        ui.secretariaQuotationReplaceTemplateBtn.disabled = true;
    }
    setFormMessage(ui.secretariaQuotationMessage, 'Actualizando PDF base de cotizacion...', 'success');
    try {
        const fileBase64 = await readFileAsDataUrl(selectedFile);
        const result = await apiRequest('/admin/cotizaciones/documento', {
            method: 'POST',
            body: JSON.stringify({
                fileName,
                fileBase64
            })
        });
        applySecretariaQuotationTemplateMetadata(result);
        setFormMessage(
            ui.secretariaQuotationMessage,
            result?.message || 'Documento base actualizado correctamente.',
            'success'
        );
    } catch (error) {
        setFormMessage(ui.secretariaQuotationMessage, error.message, 'error');
    } finally {
        if (inputElement) {
            inputElement.value = '';
        }
        if (ui.secretariaQuotationReplaceTemplateBtn) {
            ui.secretariaQuotationReplaceTemplateBtn.disabled = false;
        }
    }
}

function waitForMs(durationMs) {
    const safeDuration = Number.isFinite(Number(durationMs)) ? Math.max(0, Number(durationMs)) : 0;
    return new Promise((resolve) => {
        window.setTimeout(resolve, safeDuration);
    });
}

function showMailSendProgress(startMessage = 'Enviando correo...') {
    if (!ui || !ui.mailSendProgressOverlay) {
        return;
    }

    mailSendProgressOpenedAt = Date.now();
    if (ui.mailSendProgressText) {
        ui.mailSendProgressText.textContent = startMessage;
    }
    if (ui.mailSendProgressStatus) {
        ui.mailSendProgressStatus.textContent = '';
        ui.mailSendProgressStatus.classList.remove('success', 'error');
    }
    ui.mailSendProgressOverlay.classList.remove('hidden');
}

async function hideMailSendProgress(finalMessage = '', statusType = 'success', options = {}) {
    if (!ui || !ui.mailSendProgressOverlay) {
        return;
    }

    const minDurationMs = Number.isFinite(Number(options.minDurationMs))
        ? Math.max(0, Number(options.minDurationMs))
        : MAIL_SEND_PROGRESS_MIN_MS;
    const finalDurationMs = Number.isFinite(Number(options.finalDurationMs))
        ? Math.max(0, Number(options.finalDurationMs))
        : MAIL_SEND_PROGRESS_FINAL_MS;

    const elapsedMs = Date.now() - mailSendProgressOpenedAt;
    const remainingMs = Math.max(0, minDurationMs - elapsedMs);
    if (remainingMs > 0) {
        await waitForMs(remainingMs);
    }

    if (ui.mailSendProgressStatus) {
        ui.mailSendProgressStatus.textContent = finalMessage || '';
        ui.mailSendProgressStatus.classList.remove('success', 'error');
        if (statusType === 'error') {
            ui.mailSendProgressStatus.classList.add('error');
        } else if (statusType === 'success') {
            ui.mailSendProgressStatus.classList.add('success');
        }
    }

    if (finalDurationMs > 0) {
        await waitForMs(finalDurationMs);
    }

    ui.mailSendProgressOverlay.classList.add('hidden');
    if (ui.mailSendProgressText) {
        ui.mailSendProgressText.textContent = 'Enviando correo...';
    }
    if (ui.mailSendProgressStatus) {
        ui.mailSendProgressStatus.textContent = '';
        ui.mailSendProgressStatus.classList.remove('success', 'error');
    }
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

function renderSecretariaQuotationItems(servicios) {
    if (!ui || !ui.secretariaQuotationItemsBody) {
        return;
    }

    ui.secretariaQuotationItemsBody.innerHTML = '';
    if (!Array.isArray(servicios) || servicios.length === 0) {
        ui.secretariaQuotationItemsBody.innerHTML = '<tr><td colspan="3">Sin valores de cotizacion para mostrar.</td></tr>';
        return;
    }

    servicios.forEach((item) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${escapeHtml(item.nombreServicio || '-')}</td>
            <td>${escapeHtml(formatUtmUnits(item.valorUtm))}</td>
            <td>${escapeHtml(formatInvoiceAmount(item.valorPesos))}</td>
        `;
        ui.secretariaQuotationItemsBody.appendChild(row);
    });
}

function renderSecretariaQuotationParcelamientoRanges(parcelamientoRows, selectedKey = '') {
    if (!ui || !ui.secretariaQuotationParcelamientoRange) {
        return;
    }

    const select = ui.secretariaQuotationParcelamientoRange;
    const currentValue = normalizeInvoiceText(selectedKey || select.value, 80);
    select.innerHTML = '';

    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Sin tramo especifico (enviar tabla completa)';
    select.appendChild(defaultOption);

    if (!Array.isArray(parcelamientoRows) || parcelamientoRows.length === 0) {
        select.value = '';
        return;
    }

    parcelamientoRows.forEach((item) => {
        const key = normalizeInvoiceText(item?.key || item?.lotes, 80);
        if (!key) {
            return;
        }
        const lotes = normalizeInvoiceText(item?.lotes, 80) || key;
        const total = formatInvoiceAmount(item?.total);
        const option = document.createElement('option');
        option.value = key;
        option.textContent = `${lotes} lotes (${total} total con IVA)`;
        select.appendChild(option);
    });

    const canKeepSelection = Array.from(select.options).some((option) => option.value === currentValue);
    select.value = canKeepSelection ? currentValue : '';
}

function applySecretariaQuotationSummary(summary) {
    secretariaQuotationSummary = summary || null;
    if (!ui) {
        return;
    }

    if (ui.secretariaQuotationPeriod) {
        if (summary?.periodo) {
            ui.secretariaQuotationPeriod.value = formatMonthlyUtmPeriodLabel({
                anio: summary.periodo.anio,
                mes: summary.periodo.mes
            });
        } else {
            ui.secretariaQuotationPeriod.value = '';
        }
    }
    if (ui.secretariaQuotationUtmValue) {
        ui.secretariaQuotationUtmValue.value = summary?.utm?.valor ? formatInvoiceAmount(summary.utm.valor) : '';
    }
    if (ui.secretariaQuotationSource) {
        const updatedText = normalizeInvoiceText(summary?.utm?.updatedAt, 60);
        const registeredBy = normalizeInvoiceText(summary?.utm?.registradoPor, 120);
        const details = [];
        if (registeredBy) {
            details.push(`Registrado por: ${registeredBy}`);
        }
        if (updatedText) {
            details.push(`Actualizado: ${updatedText}`);
        }
        ui.secretariaQuotationSource.textContent = details.join(' | ');
    }
    if (ui.secretariaQuotationTotals) {
        if (summary) {
            ui.secretariaQuotationTotals.textContent = `Total estandar: ${formatUtmUnits(
                summary.totalUtm
            )} | Equivalencia CLP: ${formatInvoiceAmount(summary.totalPesos)}`;
        } else {
            ui.secretariaQuotationTotals.textContent = '';
        }
    }
    renderSecretariaQuotationParcelamientoRanges(summary?.parcelamiento || [], ui.secretariaQuotationParcelamientoRange?.value || '');
    renderSecretariaQuotationItems(summary?.servicios || []);
}

async function refreshSecretariaQuotationSummary(options = {}) {
    const showErrors = options.showErrors !== false;
    if (!canUseSecretaryFeatures(currentUser) || !ui) {
        return null;
    }

    try {
        const summary = await apiRequest('/cotizaciones/resumen');
        applySecretariaQuotationSummary(summary);
        setFormMessage(ui.secretariaQuotationMessage, '', '');
        return summary;
    } catch (error) {
        secretariaQuotationSummary = null;
        applySecretariaQuotationSummary(null);
        if (showErrors) {
            setFormMessage(ui.secretariaQuotationMessage, error.message, 'error');
        }
        return null;
    }
}

async function handleOpenSecretariaQuotationModal() {
    if (!canUseSecretaryFeatures(currentUser) || !ui || !ui.secretariaQuotationModalOverlay) {
        return;
    }

    closeSecretariaInvoicesModal();
    setFormMessage(ui.secretariaQuotationMessage, '', '');
    ui.secretariaQuotationModalOverlay.classList.remove('hidden');
    syncSecretariaQuotationTemplateControls();
    if (ui.secretariaQuotationClientName) {
        ui.secretariaQuotationClientName.value = '';
    }
    if (ui.secretariaQuotationClientEmail) {
        ui.secretariaQuotationClientEmail.value = '';
    }
    if (ui.secretariaQuotationReference) {
        ui.secretariaQuotationReference.value = '';
    }
    if (ui.secretariaQuotationParcelamientoRange) {
        ui.secretariaQuotationParcelamientoRange.value = '';
    }

    const loadTasks = [refreshSecretariaQuotationSummary({ showErrors: true })];
    if (isSuperUser(currentUser)) {
        loadTasks.push(loadSecretariaQuotationTemplateMetadata({ showErrors: false }));
    }
    await Promise.all(loadTasks);
    if (ui.secretariaQuotationClientName) {
        requestAnimationFrame(() => {
            ui.secretariaQuotationClientName.focus();
        });
    } else if (ui.secretariaQuotationClientEmail) {
        requestAnimationFrame(() => {
            ui.secretariaQuotationClientEmail.focus();
        });
    }
}

async function handleSendSecretariaQuotation(event) {
    if (event) {
        event.preventDefault();
    }

    if (!canUseSecretaryFeatures(currentUser) || !ui) {
        return;
    }

    const clientNameRaw = normalizeInvoiceText(ui.secretariaQuotationClientName?.value, 120);
    const clientName = formatNombreRazonSocialTitleCase(clientNameRaw);
    if (ui.secretariaQuotationClientName) {
        ui.secretariaQuotationClientName.value = clientName;
    }
    const destinationEmail = normalizeInvoiceText(ui.secretariaQuotationClientEmail?.value, 255).toLowerCase();
    const referenciaCliente = normalizeInvoiceText(ui.secretariaQuotationReference?.value, 500);
    const parcelamientoRangeKey = normalizeInvoiceText(ui.secretariaQuotationParcelamientoRange?.value, 80);
    if (!clientName) {
        setFormMessage(ui.secretariaQuotationMessage, 'Debes ingresar el nombre del cliente.', 'error');
        return;
    }
    if (!destinationEmail) {
        setFormMessage(ui.secretariaQuotationMessage, 'Debes ingresar el correo del cliente.', 'error');
        return;
    }
    if (!isValidEmail(destinationEmail)) {
        setFormMessage(ui.secretariaQuotationMessage, 'El correo del cliente no es valido.', 'error');
        return;
    }

    let summary = secretariaQuotationSummary;
    if (!summary) {
        summary = await refreshSecretariaQuotationSummary({ showErrors: true });
    }
    if (!summary) {
        return;
    }
    if (!Array.isArray(summary.servicios) || summary.servicios.length === 0) {
        setFormMessage(ui.secretariaQuotationMessage, 'No hay valores estandar configurados para cotizacion.', 'error');
        return;
    }

    const sendButton = ui.secretariaQuotationSendBtn;
    if (sendButton) {
        sendButton.disabled = true;
    }
    setFormMessage(ui.secretariaQuotationMessage, 'Enviando cotizacion por correo...', 'success');
    showMailSendProgress('Enviando cotizacion por correo...');

    try {
        const buildPayload = (referenceValue) => ({
            clientName,
            destinationEmail,
            referenciaCliente: referenceValue,
            parcelamientoRangeKey
        });
        let result;
        try {
            result = await apiRequest('/cotizaciones/enviar', {
                method: 'POST',
                body: JSON.stringify(buildPayload(referenciaCliente))
            });
        } catch (sendError) {
            const normalizedMessage = String(sendError?.message || '').toLowerCase();
            const needsLegacyReferenceFallback =
                !referenciaCliente &&
                normalizedMessage.includes('referencia personalizada') &&
                normalizedMessage.includes('cliente');
            if (!needsLegacyReferenceFallback) {
                throw sendError;
            }

            // Compatibilidad: algunos despliegues antiguos exigen referencia no vacia.
            result = await apiRequest('/cotizaciones/enviar', {
                method: 'POST',
                body: JSON.stringify(buildPayload(clientName))
            });
        }
        await hideMailSendProgress('Correo enviado correctamente.', 'success');
        setFormMessage(
            ui.secretariaQuotationMessage,
            result?.message || `Cotizacion enviada correctamente a ${destinationEmail}.`,
            'success'
        );
    } catch (error) {
        await hideMailSendProgress('No fue posible enviar el correo.', 'error', {
            minDurationMs: MAIL_SEND_PROGRESS_MIN_MS,
            finalDurationMs: 1200
        });
        setFormMessage(ui.secretariaQuotationMessage, error.message, 'error');
    } finally {
        if (sendButton) {
            sendButton.disabled = false;
        }
    }
}

function toLocalDateInputValue(date) {
    const target = date instanceof Date ? date : new Date(date);
    const timezoneOffsetMs = target.getTimezoneOffset() * 60000;
    return new Date(target.getTime() - timezoneOffsetMs).toISOString().slice(0, 10);
}

function isValidDateInputValue(value) {
    const text = String(value || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
        return false;
    }

    const parsed = new Date(`${text}T00:00:00Z`);
    if (Number.isNaN(parsed.getTime())) {
        return false;
    }

    return parsed.toISOString().slice(0, 10) === text;
}

function populateDateRangeSucursalSelect(sucursales, selectedValue = '') {
    if (!ui || !ui.dateRangeSucursal) {
        return;
    }

    const normalizedOptions = Array.from(
        new Set(
            (sucursales || [])
                .map((item) => String(item || '').trim())
                .filter((item) => item.length > 0)
        )
    ).sort((left, right) => left.localeCompare(right, 'es'));

    ui.dateRangeSucursal.innerHTML = '<option value="">TODAS</option>';
    normalizedOptions.forEach((branchName) => {
        const option = document.createElement('option');
        option.value = branchName;
        option.textContent = branchName;
        ui.dateRangeSucursal.appendChild(option);
    });

    if (selectedValue && normalizedOptions.includes(selectedValue)) {
        ui.dateRangeSucursal.value = selectedValue;
        return;
    }

    ui.dateRangeSucursal.value = '';
}

async function loadDateRangeBranches(options = {}) {
    const preserveSelection = options.preserveSelection !== false;
    const previousSelection = preserveSelection && ui?.dateRangeSucursal ? String(ui.dateRangeSucursal.value || '').trim() : '';

    try {
        const data = await apiRequest('/registros/sucursales-disponibles');
        dateRangeAvailableBranches = Array.isArray(data.sucursales) ? data.sucursales : [];
        populateDateRangeSucursalSelect(dateRangeAvailableBranches, previousSelection);
    } catch (error) {
        if (!Array.isArray(dateRangeAvailableBranches)) {
            dateRangeAvailableBranches = [];
        }
        populateDateRangeSucursalSelect(dateRangeAvailableBranches, previousSelection);
    }
}

function closeDateRangeModal(options = {}) {
    if (!ui || !ui.dateRangeModalOverlay) {
        return;
    }

    if (dateRangeResizeTimerId) {
        window.clearTimeout(dateRangeResizeTimerId);
        dateRangeResizeTimerId = 0;
    }

    const preserveResults = Boolean(options.preserveResults);
    ui.dateRangeModalOverlay.classList.add('hidden');
    setFormMessage(ui.dateRangeMessage, '', '');
    if (!preserveResults) {
        selectedDateRangeIngreso = '';
        dateRangeMatches = [];
        dateRangeCurrentPage = 1;
        dateRangePageSize = DATE_RANGE_FALLBACK_PAGE_SIZE;
        if (ui && ui.dateRangeSucursal) {
            ui.dateRangeSucursal.value = '';
        }
        if (ui && ui.dateRangeRegion) {
            ui.dateRangeRegion.value = '';
        }
        if (ui && ui.dateRangeComuna) {
            populateComunas(ui.dateRangeComuna, '', '', 'TODAS');
        }
        renderDateRangeResults([]);
        updateDateRangeExportButtonState([]);
    }
}

function updateDateRangeExportButtonState(registros) {
    if (!ui || !ui.dateRangeExportBtn) {
        return;
    }

    const hasResults = Array.isArray(registros) && registros.length > 0;
    ui.dateRangeExportBtn.classList.toggle('hidden', !hasResults);
    ui.dateRangeExportBtn.disabled = !hasResults;
}

function getDateRangePageSize() {
    const safePageSize = Number.isInteger(dateRangePageSize) ? dateRangePageSize : 0;
    const normalized = safePageSize > 0 ? safePageSize : DATE_RANGE_FALLBACK_PAGE_SIZE;
    return Math.min(DATE_RANGE_MAX_PAGE_SIZE, Math.max(DATE_RANGE_MIN_PAGE_SIZE, normalized));
}

function recalculateDateRangePageSize() {
    if (!ui || !ui.dateRangeTableWrapper) {
        dateRangePageSize = DATE_RANGE_FALLBACK_PAGE_SIZE;
        return;
    }

    const wrapperHeight = ui.dateRangeTableWrapper.clientHeight;
    if (!Number.isFinite(wrapperHeight) || wrapperHeight <= 0) {
        dateRangePageSize = DATE_RANGE_FALLBACK_PAGE_SIZE;
        return;
    }

    const tableHeader = ui.dateRangeTableWrapper.querySelector('thead');
    const tableHeaderHeight = tableHeader ? tableHeader.getBoundingClientRect().height : 0;
    const sampleRow = ui.dateRangeTableWrapper.querySelector('#dateRangeResultsBody tr');
    const sampleRowHeight = sampleRow ? sampleRow.getBoundingClientRect().height : 0;
    const rowHeight = sampleRowHeight > 0 ? sampleRowHeight : 38;
    const availableBodyHeight = Math.max(wrapperHeight - tableHeaderHeight - 2, rowHeight);
    const computedRows = Math.floor(availableBodyHeight / rowHeight);
    const normalizedRows = computedRows > 0 ? computedRows : DATE_RANGE_FALLBACK_PAGE_SIZE;
    dateRangePageSize = Math.min(DATE_RANGE_MAX_PAGE_SIZE, Math.max(DATE_RANGE_MIN_PAGE_SIZE, normalizedRows));
}

function scheduleDateRangeLayoutRefresh() {
    if (dateRangeResizeTimerId) {
        window.clearTimeout(dateRangeResizeTimerId);
    }

    dateRangeResizeTimerId = window.setTimeout(() => {
        dateRangeResizeTimerId = 0;
        if (!ui || !ui.dateRangeModalOverlay || ui.dateRangeModalOverlay.classList.contains('hidden')) {
            return;
        }

        const previousPageSize = getDateRangePageSize();
        recalculateDateRangePageSize();
        if (previousPageSize !== getDateRangePageSize() && Array.isArray(dateRangeMatches) && dateRangeMatches.length > 0) {
            renderDateRangeResults(dateRangeMatches);
        }
    }, 120);
}

function updateDateRangePaginationControls(totalItems) {
    if (!ui) {
        return;
    }

    const safeTotalItems = Number.isInteger(totalItems) && totalItems > 0 ? totalItems : 0;
    const pageSize = getDateRangePageSize();
    const totalPages = safeTotalItems > 0 ? Math.max(1, Math.ceil(safeTotalItems / pageSize)) : 1;
    dateRangeCurrentPage = Math.min(Math.max(dateRangeCurrentPage, 1), totalPages);
    const isEmpty = safeTotalItems === 0;
    const isFirstPage = dateRangeCurrentPage <= 1;
    const isLastPage = dateRangeCurrentPage >= totalPages;

    if (ui.dateRangePageInfo) {
        ui.dateRangePageInfo.textContent = isEmpty ? '0/0' : `${dateRangeCurrentPage}/${totalPages}`;
    }
    if (ui.dateRangeFirstBtn) {
        ui.dateRangeFirstBtn.disabled = isEmpty || isFirstPage;
    }
    if (ui.dateRangePrevBtn) {
        ui.dateRangePrevBtn.disabled = isEmpty || isFirstPage;
    }
    if (ui.dateRangeNextBtn) {
        ui.dateRangeNextBtn.disabled = isEmpty || isLastPage;
    }
    if (ui.dateRangeLastBtn) {
        ui.dateRangeLastBtn.disabled = isEmpty || isLastPage;
    }
}

function handleDateRangeGoToPage(targetPage) {
    if (!Array.isArray(dateRangeMatches) || dateRangeMatches.length === 0) {
        return;
    }

    const pageSize = getDateRangePageSize();
    const totalPages = Math.max(1, Math.ceil(dateRangeMatches.length / pageSize));
    const safeTargetPage = Math.min(Math.max(Number(targetPage) || 1, 1), totalPages);
    if (safeTargetPage === dateRangeCurrentPage) {
        return;
    }

    dateRangeCurrentPage = safeTargetPage;
    renderDateRangeResults(dateRangeMatches);
}

function renderDateRangeResults(registros) {
    if (!ui || !ui.dateRangeResultsBody) {
        return;
    }

    const rows = Array.isArray(registros) ? registros : [];
    recalculateDateRangePageSize();
    const pageSize = getDateRangePageSize();
    updateDateRangeExportButtonState(rows);
    updateDateRangePaginationControls(rows.length);
    ui.dateRangeResultsBody.innerHTML = '';
    if (rows.length === 0) {
        ui.dateRangeResultsBody.innerHTML = '<tr><td colspan="8">Sin resultados para mostrar.</td></tr>';
        return;
    }

    const startIndex = (dateRangeCurrentPage - 1) * pageSize;
    const pageRows = rows.slice(startIndex, startIndex + pageSize);

    pageRows.forEach((registro, pageIndex) => {
        const index = startIndex + pageIndex;
        const row = document.createElement('tr');
        const radioId = `dateRangeSel-${index}`;

        row.innerHTML = `
            <td><input type="radio" id="${radioId}" name="dateRangeSelection" value="${escapeHtml(registro.numIngreso || '')}"></td>
            <td><label for="${radioId}">${escapeHtml(registro.numIngreso || '')}</label></td>
            <td>${escapeHtml(formatDateTime(registro.createdAt))}</td>
            <td>${escapeHtml(registro.nombre || '')}</td>
            <td>${escapeHtml(registro.rut || '')}</td>
            <td>${escapeHtml(registro.rol || '')}</td>
            <td>${escapeHtml(registro.region || '')}</td>
            <td>${escapeHtml(registro.comuna || '')}</td>
        `;

        const radio = row.querySelector(`input[name="dateRangeSelection"]`);
        if (radio) {
            radio.checked = selectedDateRangeIngreso === String(registro.numIngreso || '');
            radio.addEventListener('change', () => {
                selectedDateRangeIngreso = String(registro.numIngreso || '');
            });
        }

        row.addEventListener('click', () => {
            if (!radio) {
                return;
            }
            radio.checked = true;
            selectedDateRangeIngreso = String(registro.numIngreso || '');
        });

        ui.dateRangeResultsBody.appendChild(row);
    });
}

function closeSearchMatchModal(options = {}) {
    if (!ui || !ui.searchMatchModalOverlay) {
        return;
    }

    const preserveResults = Boolean(options.preserveResults);
    ui.searchMatchModalOverlay.classList.add('hidden');
    setFormMessage(ui.searchMatchMessage, '', '');

    if (!preserveResults) {
        selectedSearchMatchIngreso = '';
        searchMatchCandidates = [];
        searchMatchCurrentPage = 1;
        renderSearchMatchResults([]);
    }
}

function updateSearchMatchPaginationControls(totalItems) {
    if (!ui) {
        return;
    }

    const safeTotalItems = Number.isInteger(totalItems) && totalItems > 0 ? totalItems : 0;
    const totalPages = safeTotalItems > 0 ? Math.max(1, Math.ceil(safeTotalItems / SEARCH_MATCH_PAGE_SIZE)) : 1;
    searchMatchCurrentPage = Math.min(Math.max(searchMatchCurrentPage, 1), totalPages);
    const isEmpty = safeTotalItems === 0;
    const isFirstPage = searchMatchCurrentPage <= 1;
    const isLastPage = searchMatchCurrentPage >= totalPages;

    if (ui.searchMatchPageInfo) {
        ui.searchMatchPageInfo.textContent = isEmpty ? '0/0' : `${searchMatchCurrentPage}/${totalPages}`;
    }
    if (ui.searchMatchFirstBtn) {
        ui.searchMatchFirstBtn.disabled = isEmpty || isFirstPage;
    }
    if (ui.searchMatchPrevBtn) {
        ui.searchMatchPrevBtn.disabled = isEmpty || isFirstPage;
    }
    if (ui.searchMatchNextBtn) {
        ui.searchMatchNextBtn.disabled = isEmpty || isLastPage;
    }
    if (ui.searchMatchLastBtn) {
        ui.searchMatchLastBtn.disabled = isEmpty || isLastPage;
    }
}

function handleSearchMatchGoToPage(targetPage) {
    if (!Array.isArray(searchMatchCandidates) || searchMatchCandidates.length === 0) {
        return;
    }

    const totalPages = Math.max(1, Math.ceil(searchMatchCandidates.length / SEARCH_MATCH_PAGE_SIZE));
    const safeTargetPage = Math.min(Math.max(Number(targetPage) || 1, 1), totalPages);
    if (safeTargetPage === searchMatchCurrentPage) {
        return;
    }

    searchMatchCurrentPage = safeTargetPage;
    renderSearchMatchResults(searchMatchCandidates);
}

function renderSearchMatchResults(coincidencias) {
    if (!ui || !ui.searchMatchResultsBody) {
        return;
    }

    const rows = Array.isArray(coincidencias) ? coincidencias : [];
    updateSearchMatchPaginationControls(rows.length);
    ui.searchMatchResultsBody.innerHTML = '';
    if (rows.length === 0) {
        ui.searchMatchResultsBody.innerHTML = '<tr><td colspan="4">Sin coincidencias para mostrar.</td></tr>';
        return;
    }

    const startIndex = (searchMatchCurrentPage - 1) * SEARCH_MATCH_PAGE_SIZE;
    const pageRows = rows.slice(startIndex, startIndex + SEARCH_MATCH_PAGE_SIZE);

    pageRows.forEach((item, pageIndex) => {
        const index = startIndex + pageIndex;
        const numIngreso = String(item?.numIngreso || '').trim();
        const row = document.createElement('tr');
        row.classList.add('search-match-row');
        if (selectedSearchMatchIngreso === numIngreso) {
            row.classList.add('is-selected');
        }

        const radioId = `searchMatchSel-${index}`;
        row.innerHTML = `
            <td><input type="radio" id="${radioId}" name="searchMatchSelection" value="${escapeHtml(numIngreso)}"></td>
            <td><label for="${radioId}">${escapeHtml(numIngreso || '-')}</label></td>
            <td>${escapeHtml(item?.rol || '')}</td>
            <td>${escapeHtml(item?.nombre || '')}</td>
        `;

        const radio = row.querySelector(`input[name="searchMatchSelection"]`);
        const handleRowSelection = () => {
            selectedSearchMatchIngreso = numIngreso;
            renderSearchMatchResults(searchMatchCandidates);
        };

        if (radio) {
            radio.checked = selectedSearchMatchIngreso === numIngreso;
            radio.addEventListener('change', handleRowSelection);
        }

        row.addEventListener('click', () => {
            if (!radio) {
                return;
            }
            radio.checked = true;
            handleRowSelection();
        });

        ui.searchMatchResultsBody.appendChild(row);
    });
}

function openSearchMatchModal(coincidencias, message = '') {
    if (!ui || !ui.searchMatchModalOverlay) {
        return false;
    }

    const normalizedMatches = Array.isArray(coincidencias)
        ? coincidencias.filter((item) => String(item?.numIngreso || '').trim().length > 0)
        : [];
    if (normalizedMatches.length === 0) {
        return false;
    }

    searchMatchCandidates = normalizedMatches;
    searchMatchCurrentPage = 1;
    selectedSearchMatchIngreso = String(normalizedMatches[0].numIngreso || '').trim();
    renderSearchMatchResults(searchMatchCandidates);
    setFormMessage(
        ui.searchMatchMessage,
        normalizeText(message) || `Se encontraron ${normalizedMatches.length} coincidencias. Selecciona el registro correcto.`,
        ''
    );
    ui.searchMatchModalOverlay.classList.remove('hidden');
    return true;
}

async function handleSearchMatchAcceptSelection() {
    if (!selectedSearchMatchIngreso) {
        setFormMessage(ui.searchMatchMessage, 'Selecciona un registro para cargar.', 'error');
        return;
    }

    if (loadedIngresoOriginal && hasPendingLoadedRegistroChanges()) {
        showPendingModifyRequiredWarningPopup();
        return;
    }

    try {
        await loadRegistroByIngreso(selectedSearchMatchIngreso, 'seleccion de coincidencias');
        closeSearchMatchModal();
    } catch (error) {
        setFormMessage(ui.searchMatchMessage, error.message, 'error');
    }
}

async function loadRegistroByIngreso(numIngreso, sourceLabel = 'base de datos') {
    const normalizedIngreso = String(numIngreso || '').trim();
    if (!normalizedIngreso) {
        throw new Error('NRO INGRESO invalido para cargar el registro.');
    }

    const match = await apiRequest(`/registros/buscar?numIngreso=${encodeURIComponent(normalizedIngreso)}`);
    fillForm(ui.registroForm, match, ui.comuna);
    loadedIngresoOriginal = match.numIngreso || '';
    saveLoadedRegistroSnapshot();
    await syncCurrentInvoiceRequestStatusByIngreso(loadedIngresoOriginal);
    renderHistory(isGuestUser(currentUser) ? [] : match.historial || []);
    updateRegistroActionButtons();
    if (!isGuestUser(currentUser)) {
        showExistingRecordSaveWarningPopup();
    }
    if (normalizeInvoiceRequestStatus(currentInvoiceRequestStatus) === 'ENVIADA') {
        setFormMessage(ui.facturaFormMessage, 'Esta factura ya fue enviada al contador.', 'success');
    } else {
        setFormMessage(ui.facturaFormMessage, '', '');
    }

    const message = isGuestUser(currentUser)
        ? `Registro ${match.numIngreso} cargado en modo solo lectura.`
        : `Registro ${match.numIngreso} cargado desde ${sourceLabel}.`;
    setFormMessage(ui.formMessage, message, 'success');

    return match;
}

async function handleOpenDateRangeModal() {
    if (!ui || !ui.dateRangeModalOverlay || !ui.dateRangeFrom || !ui.dateRangeTo) {
        return;
    }

    if (loadedIngresoOriginal && hasPendingLoadedRegistroChanges()) {
        showPendingModifyRequiredWarningPopup();
        return;
    }

    const today = toLocalDateInputValue(new Date());
    const lastThirtyDays = toLocalDateInputValue(new Date(Date.now() - 29 * 24 * 60 * 60 * 1000));
    await loadDateRangeBranches({ preserveSelection: false });
    ui.dateRangeFrom.value = ui.dateRangeFrom.value || lastThirtyDays;
    ui.dateRangeTo.value = ui.dateRangeTo.value || today;
    if (ui.dateRangeSucursal) {
        ui.dateRangeSucursal.value = '';
    }
    if (ui.dateRangeRegion) {
        ui.dateRangeRegion.value = '';
    }
    if (ui.dateRangeComuna) {
        populateComunas(ui.dateRangeComuna, '', '', 'TODAS');
    }
    setFormMessage(ui.dateRangeMessage, '', '');
    selectedDateRangeIngreso = '';
    dateRangeMatches = [];
    dateRangeCurrentPage = 1;
    dateRangePageSize = DATE_RANGE_FALLBACK_PAGE_SIZE;
    renderDateRangeResults([]);
    ui.dateRangeModalOverlay.classList.remove('hidden');
    scheduleDateRangeLayoutRefresh();
}

async function handleDateRangeSearch(event) {
    event.preventDefault();

    if (!ui || !ui.dateRangeFrom || !ui.dateRangeTo) {
        return;
    }

    const fechaDesde = String(ui.dateRangeFrom.value || '').trim();
    const fechaHasta = String(ui.dateRangeTo.value || '').trim();
    const sucursal = String(ui.dateRangeSucursal?.value || '').trim();
    const region = String(ui.dateRangeRegion?.value || '').trim();
    const comuna = String(ui.dateRangeComuna?.value || '').trim();
    if (!isValidDateInputValue(fechaDesde) || !isValidDateInputValue(fechaHasta)) {
        setFormMessage(ui.dateRangeMessage, 'Ingresa un rango de fechas valido.', 'error');
        return;
    }

    if (fechaDesde > fechaHasta) {
        setFormMessage(ui.dateRangeMessage, 'La fecha DESDE no puede ser mayor que HASTA.', 'error');
        return;
    }

    try {
        const query = new URLSearchParams({
            fechaDesde,
            fechaHasta
        });
        if (sucursal) {
            query.set('sucursal', sucursal);
        }
        if (region) {
            query.set('region', region);
        }
        if (comuna) {
            query.set('comuna', comuna);
        }
        const result = await apiRequest(`/registros/buscar-rango-fechas?${query.toString()}`);
        const registros = Array.isArray(result.registros) ? result.registros : [];
        dateRangeMatches = registros;
        dateRangeCurrentPage = 1;
        selectedDateRangeIngreso = registros.length === 1 ? String(registros[0].numIngreso || '') : '';
        renderDateRangeResults(registros);
        scheduleDateRangeLayoutRefresh();
        setFormMessage(
            ui.dateRangeMessage,
            registros.length > 0
                ? `Se encontraron ${registros.length} registros en el rango seleccionado.`
                : 'No se encontraron registros para ese rango de fechas.',
            registros.length > 0 ? 'success' : 'error'
        );
    } catch (error) {
        setFormMessage(ui.dateRangeMessage, error.message, 'error');
    }
}

function formatEstadoForExport(value) {
    const normalized = normalizeEstadoValue(value);
    if (normalized === 'facturada') {
        return 'FACTURADA';
    }
    if (normalized === 'no_facturada') {
        return 'NO FACTURADA';
    }
    return String(normalized || '').toUpperCase();
}

function normalizeRegionLookupKey(value) {
    const normalized = String(value || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .replace(/\([^)]*\)/g, ' ')
        .replace(/\bregion\b/g, ' ')
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();
    return normalized.replace(/^de\s+/, '').replace(/^del\s+/, '');
}

function resolveRegionRomanNumeral(regionValue) {
    const normalizedRegion = normalizeRegionLookupKey(regionValue);
    if (!normalizedRegion) {
        return '';
    }
    return REGION_ROMAN_NUMERAL_LOOKUP[normalizedRegion] || '';
}

function formatNumIngresoForDateRangeExport(numIngreso, regionValue) {
    const normalizedIngreso = normalizeText(numIngreso);
    if (!normalizedIngreso) {
        return '';
    }

    const romanNumeral = resolveRegionRomanNumeral(regionValue);
    if (!romanNumeral) {
        return normalizedIngreso;
    }

    const expectedSuffix = `-${romanNumeral}`;
    if (normalizedIngreso.toUpperCase().endsWith(expectedSuffix)) {
        return normalizedIngreso;
    }

    return `${normalizedIngreso}${expectedSuffix}`;
}

function buildDateRangeExportRows(registros) {
    if (!Array.isArray(registros)) {
        return [];
    }

    return registros.map((registro, index) => [
        String(index + 1),
        formatNumIngresoForDateRangeExport(registro?.numIngreso, registro?.region),
        formatDateTime(registro?.createdAt),
        normalizeText(registro?.nombre),
        normalizeText(registro?.rut),
        normalizeText(registro?.rol),
        normalizeText(registro?.telefono),
        normalizeText(registro?.correo),
        normalizeText(registro?.region),
        normalizeText(registro?.comuna),
        normalizeText(registro?.sucursal),
        formatEstadoForExport(registro?.estado)
    ]);
}

function buildDateRangeExportFileName(extension = 'xlsx') {
    const safeExtension = String(extension || '').trim().toLowerCase() === 'xls' ? 'xls' : 'xlsx';
    const from = String(ui?.dateRangeFrom?.value || '').trim().replace(/-/g, '');
    const to = String(ui?.dateRangeTo?.value || '').trim().replace(/-/g, '');
    if (from && to) {
        return `registros_${from}_${to}.${safeExtension}`;
    }
    const now = new Date();
    const fallback = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    return `registros_${fallback}.${safeExtension}`;
}

function triggerBrowserDownload(linkHref, fileName) {
    const link = document.createElement('a');
    link.href = linkHref;
    link.download = fileName;
    link.rel = 'noopener';
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function downloadBlobFile(blob, fileName) {
    const browserNavigator = window.navigator || {};
    if (typeof browserNavigator.msSaveOrOpenBlob === 'function') {
        browserNavigator.msSaveOrOpenBlob(blob, fileName);
        return;
    }

    const urlApi = window.URL || window.webkitURL;
    if (!urlApi || typeof urlApi.createObjectURL !== 'function') {
        throw new Error('El navegador no soporta descarga de archivos.');
    }

    const objectUrl = urlApi.createObjectURL(blob);
    try {
        triggerBrowserDownload(objectUrl, fileName);
    } finally {
        urlApi.revokeObjectURL(objectUrl);
    }
}

async function exportDateRangeWorkbookAsXlsxTable(registros) {
    const excelJsLib = window && window.ExcelJS;
    if (!excelJsLib || typeof excelJsLib.Workbook !== 'function') {
        throw new Error('No se pudo cargar la libreria de tabla Excel.');
    }

    const headers = ['#', 'NRO INGRESO', 'FECHA REGISTRO', 'NOMBRE', 'RUT', 'ROL', 'TELEFONO', 'CORREO', 'REGION', 'COMUNA', 'SUCURSAL', 'ESTADO'];
    const rows = buildDateRangeExportRows(registros).map((rowValues) => rowValues.map((value) => String(value ?? '')));
    const columnWidths = [6, 18, 20, 40, 16, 16, 15, 35, 20, 20, 20, 14];

    const workbook = new excelJsLib.Workbook();
    workbook.creator = 'Geo Rural';
    workbook.created = new Date();
    const worksheet = workbook.addWorksheet('Registros', {
        views: [{ state: 'frozen', ySplit: 1 }]
    });

    worksheet.columns = headers.map((name, index) => ({
        header: name,
        key: `c${index + 1}`,
        width: columnWidths[index]
    }));

    worksheet.addTable({
        name: 'TablaRegistros',
        ref: 'A1',
        headerRow: true,
        totalsRow: false,
        style: {
            theme: 'TableStyleMedium2',
            showRowStripes: true
        },
        columns: headers.map((name) => ({ name })),
        rows
    });

    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell) => {
            cell.alignment = { vertical: 'middle' };
            if (rowNumber === 1) {
                cell.font = { bold: true };
            }
        });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    downloadBlobFile(blob, buildDateRangeExportFileName('xlsx'));
}

function exportDateRangeWorkbookAsXls(registros) {
    const xlsxLib = window && window.XLSX;
    if (!xlsxLib || !xlsxLib.utils || typeof xlsxLib.writeFile !== 'function') {
        throw new Error('No se pudo cargar la libreria de exportacion Excel.');
    }

    const headers = ['#', 'NRO INGRESO', 'FECHA REGISTRO', 'NOMBRE', 'RUT', 'ROL', 'TELEFONO', 'CORREO', 'REGION', 'COMUNA', 'SUCURSAL', 'ESTADO'];
    const rows = buildDateRangeExportRows(registros);
    const worksheetData = [headers, ...rows];
    const worksheet = xlsxLib.utils.aoa_to_sheet(worksheetData);

    worksheet['!cols'] = [
        { wch: 6 },
        { wch: 18 },
        { wch: 20 },
        { wch: 40 },
        { wch: 16 },
        { wch: 16 },
        { wch: 15 },
        { wch: 35 },
        { wch: 20 },
        { wch: 20 },
        { wch: 20 },
        { wch: 14 }
    ];

    const lastCol = xlsxLib.utils.encode_col(headers.length - 1);
    const lastRow = rows.length + 1;
    worksheet['!autofilter'] = { ref: `A1:${lastCol}${lastRow}` };

    const workbook = xlsxLib.utils.book_new();
    xlsxLib.utils.book_append_sheet(workbook, worksheet, 'Registros');
    xlsxLib.writeFile(workbook, buildDateRangeExportFileName('xls'), {
        bookType: 'xls',
        compression: true
    });
}

async function handleDateRangeExportWorkbook() {
    const registros = Array.isArray(dateRangeMatches) ? dateRangeMatches : [];
    if (registros.length === 0) {
        setFormMessage(ui?.dateRangeMessage, 'No hay resultados para exportar.', 'error');
        return;
    }

    try {
        const canExportAsExcelTable = Boolean(window && window.ExcelJS && typeof window.ExcelJS.Workbook === 'function');
        if (canExportAsExcelTable) {
            await exportDateRangeWorkbookAsXlsxTable(registros);
        } else {
            exportDateRangeWorkbookAsXls(registros);
        }
        setFormMessage(ui?.dateRangeMessage, `Se exportaron ${registros.length} registros en Excel.`, 'success');
    } catch (error) {
        const reason = normalizeText(error?.message || '');
        console.error('[Excel Export] Error al exportar rango de fechas:', error);
        setFormMessage(
            ui?.dateRangeMessage,
            reason ? `No fue posible generar el archivo Excel: ${reason}` : 'No fue posible generar el archivo Excel.',
            'error'
        );
    }
}

async function handleDateRangeAcceptSelection() {
    if (!selectedDateRangeIngreso) {
        setFormMessage(ui.dateRangeMessage, 'Selecciona un registro para cargar.', 'error');
        return;
    }

    if (loadedIngresoOriginal && hasPendingLoadedRegistroChanges()) {
        showPendingModifyRequiredWarningPopup();
        return;
    }

    try {
        await loadRegistroByIngreso(selectedDateRangeIngreso, 'busqueda por fecha');
        closeDateRangeModal();
    } catch (error) {
        setFormMessage(ui.dateRangeMessage, error.message, 'error');
    }
}

function getSecretaryNoMovementStorageKey(user) {
    const userId = Number(user?.id);
    if (Number.isInteger(userId) && userId > 0) {
        return `${SECRETARY_DAILY_ALERT_STORAGE_PREFIX}${userId}`;
    }

    const username = normalizeInvoiceText(user?.username, 80).toLowerCase() || 'anonimo';
    return `${SECRETARY_DAILY_ALERT_STORAGE_PREFIX}${username}`;
}

function getSecretaryNoMovementSeenDate(user) {
    const storageKey = getSecretaryNoMovementStorageKey(user);
    try {
        return String(window.localStorage.getItem(storageKey) || '').trim();
    } catch (error) {
        return '';
    }
}

function markSecretaryNoMovementSeenToday(user) {
    const storageKey = getSecretaryNoMovementStorageKey(user);
    const todayKey = toLocalDateInputValue(new Date());
    try {
        window.localStorage.setItem(storageKey, todayKey);
    } catch (error) {
        // Sin accion: no bloquear el flujo por storage.
    }
}

function getSecretaryNoMovementRecordDateKey(value) {
    if (!value) {
        return '';
    }

    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) {
        return '';
    }

    return toLocalDateInputValue(parsed);
}

function populateSecretaryNoMovementBranchFilter(records, selectedValue = '') {
    if (!ui?.secretaryNoMovementFilterSucursal) {
        return;
    }

    const normalizedOptions = Array.from(
        new Set(
            (Array.isArray(records) ? records : [])
                .map((item) => normalizeInvoiceText(item?.sucursal, 120))
                .filter((item) => Boolean(item))
        )
    ).sort((left, right) => left.localeCompare(right, 'es'));

    ui.secretaryNoMovementFilterSucursal.innerHTML = '<option value="">TODAS</option>';
    normalizedOptions.forEach((branchName) => {
        const option = document.createElement('option');
        option.value = branchName;
        option.textContent = branchName;
        ui.secretaryNoMovementFilterSucursal.appendChild(option);
    });

    if (selectedValue && normalizedOptions.includes(selectedValue)) {
        ui.secretaryNoMovementFilterSucursal.value = selectedValue;
        return;
    }

    ui.secretaryNoMovementFilterSucursal.value = '';
}

function resetSecretaryNoMovementFilters() {
    if (!ui) {
        return;
    }

    if (ui.secretaryNoMovementFilterRegion) {
        ui.secretaryNoMovementFilterRegion.value = '';
    }
    if (ui.secretaryNoMovementFilterComuna) {
        populateComunas(ui.secretaryNoMovementFilterComuna, '', '', 'TODAS');
    }
    if (ui.secretaryNoMovementFilterFrom) {
        ui.secretaryNoMovementFilterFrom.value = '';
    }
    if (ui.secretaryNoMovementFilterTo) {
        ui.secretaryNoMovementFilterTo.value = '';
    }
    populateSecretaryNoMovementBranchFilter(secretaryNoMovementRecords);
}

function renderSecretaryNoMovementRows(records) {
    if (!ui?.secretaryNoMovementBody) {
        return;
    }

    ui.secretaryNoMovementBody.innerHTML = '';
    if (!Array.isArray(records) || records.length === 0) {
        ui.secretaryNoMovementBody.innerHTML = '<tr><td colspan="8">Sin registros pendientes para mostrar.</td></tr>';
        return;
    }

    records.forEach((item) => {
        const days = Number(item?.diasSinMovimiento);
        const daysLabel = Number.isFinite(days) ? String(Math.max(0, Math.floor(days))) : '';
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${escapeHtml(item?.numIngreso || '')}</td>
            <td>${escapeHtml(formatDateTime(item?.createdAt || ''))}</td>
            <td>${escapeHtml(daysLabel)}</td>
            <td>${escapeHtml(item?.nombre || '')}</td>
            <td>${escapeHtml(item?.rut || '')}</td>
            <td>${escapeHtml(item?.sucursal || '')}</td>
            <td>${escapeHtml(item?.region || '')}</td>
            <td>${escapeHtml(item?.comuna || '')}</td>
        `;
        ui.secretaryNoMovementBody.appendChild(row);
    });
}

function applySecretaryNoMovementFilters() {
    if (!ui) {
        return false;
    }

    const sourceRecords = Array.isArray(secretaryNoMovementRecords) ? secretaryNoMovementRecords : [];
    const region = normalizeInvoiceText(ui.secretaryNoMovementFilterRegion?.value, 120);
    const comuna = normalizeInvoiceText(ui.secretaryNoMovementFilterComuna?.value, 120);
    const sucursal = normalizeInvoiceText(ui.secretaryNoMovementFilterSucursal?.value, 120);
    const fechaDesde = String(ui.secretaryNoMovementFilterFrom?.value || '').trim();
    const fechaHasta = String(ui.secretaryNoMovementFilterTo?.value || '').trim();

    if (fechaDesde && !isValidDateInputValue(fechaDesde)) {
        setFormMessage(ui.secretaryNoMovementMessage, 'La fecha DESDE es invalida.', 'error');
        return false;
    }
    if (fechaHasta && !isValidDateInputValue(fechaHasta)) {
        setFormMessage(ui.secretaryNoMovementMessage, 'La fecha HASTA es invalida.', 'error');
        return false;
    }
    if (fechaDesde && fechaHasta && fechaDesde > fechaHasta) {
        setFormMessage(ui.secretaryNoMovementMessage, 'La fecha DESDE no puede ser mayor que HASTA.', 'error');
        return false;
    }

    const filteredRecords = sourceRecords.filter((item) => {
        const itemRegion = normalizeInvoiceText(item?.region, 120);
        const itemComuna = normalizeInvoiceText(item?.comuna, 120);
        const itemSucursal = normalizeInvoiceText(item?.sucursal, 120);
        const itemDate = getSecretaryNoMovementRecordDateKey(item?.createdAt);

        if (region && itemRegion !== region) {
            return false;
        }
        if (comuna && itemComuna !== comuna) {
            return false;
        }
        if (sucursal && itemSucursal !== sucursal) {
            return false;
        }
        if (fechaDesde && (!itemDate || itemDate < fechaDesde)) {
            return false;
        }
        if (fechaHasta && (!itemDate || itemDate > fechaHasta)) {
            return false;
        }

        return true;
    });

    renderSecretaryNoMovementRows(filteredRecords);
    if (sourceRecords.length === 0) {
        setFormMessage(ui.secretaryNoMovementMessage, 'No hay registros con 8 o mas dias y comentario unico.', 'success');
        return true;
    }

    if (filteredRecords.length === 0) {
        setFormMessage(
            ui.secretaryNoMovementMessage,
            `No hay registros que coincidan con los filtros. Pendientes totales: ${sourceRecords.length}.`,
            'error'
        );
        return true;
    }

    setFormMessage(
        ui.secretaryNoMovementMessage,
        `Mostrando ${filteredRecords.length} de ${sourceRecords.length} registros pendientes.`,
        'error'
    );
    return true;
}

function closeSecretaryNoMovementModal() {
    if (ui?.secretaryNoMovementOverlay) {
        ui.secretaryNoMovementOverlay.classList.add('hidden');
    }
}

async function fetchSecretaryNoMovementRecords() {
    const query = new URLSearchParams({
        dias: String(SECRETARY_NO_MOVEMENT_DAYS)
    });
    const result = await apiRequest(`/registros/seguimiento-sin-movimiento?${query.toString()}`);
    const records = Array.isArray(result?.registros) ? result.registros : [];
    secretaryNoMovementRecords = records;
    return records;
}

function openSecretaryNoMovementModal() {
    if (!ui?.secretaryNoMovementOverlay) {
        return;
    }

    resetSecretaryNoMovementFilters();
    applySecretaryNoMovementFilters();
    ui.secretaryNoMovementOverlay.classList.remove('hidden');
    if (ui.secretaryNoMovementCloseBtn) {
        window.requestAnimationFrame(() => {
            ui.secretaryNoMovementCloseBtn.focus();
        });
    }
}

async function handleOpenSecretaryNoMovementModal() {
    if (!canUseSecretaryFeatures(currentUser)) {
        setFormMessage(ui.formMessage, 'Solo SECRETARIA o SUPER pueden ver este seguimiento.', 'error');
        return;
    }

    try {
        await fetchSecretaryNoMovementRecords();
        openSecretaryNoMovementModal();
    } catch (error) {
        setFormMessage(ui.formMessage, error.message, 'error');
    }
}

async function maybeShowSecretaryNoMovementAlert(user) {
    if (!isSecretaryUser(user) || isGuestUser(user)) {
        return;
    }

    const todayKey = toLocalDateInputValue(new Date());
    if (getSecretaryNoMovementSeenDate(user) === todayKey) {
        return;
    }

    try {
        await fetchSecretaryNoMovementRecords();
        markSecretaryNoMovementSeenToday(user);
        if (!secretaryNoMovementRecords.length) {
            return;
        }
        openSecretaryNoMovementModal();
    } catch (error) {
        // Evitar bloquear el ingreso por errores en esta alerta.
        console.warn('[Geo Rural] No fue posible cargar seguimiento de registros sin movimiento.', error);
    }
}

function getDocumentOptionLabelByValue(value) {
    const normalizedValue = normalizeDocumentOptionValue(value);
    if (!normalizedValue) {
        return '';
    }

    const option = Array.isArray(configuredDocumentOptions)
        ? configuredDocumentOptions.find((item) => normalizeDocumentOptionValue(item?.value) === normalizedValue)
        : null;
    if (option && option.label) {
        return normalizeDocumentOptionLabel(option.label);
    }

    return normalizeDocumentOptionLabel(String(value || '').replace(/_/g, ' '));
}

function sanitizeOperatorDocumentAlertRecord(item) {
    const alertId = Number(item?.id);
    const docs = Array.isArray(item?.documentosAgregados) ? item.documentosAgregados : [];
    const documentosAgregados = docs
        .map((doc) => normalizeDocumentOptionValue(doc))
        .filter((doc) => doc.length > 0);
    return {
        id: Number.isInteger(alertId) && alertId > 0 ? alertId : 0,
        numIngreso: normalizeInvoiceText(item?.numIngreso, 20),
        sucursal: normalizeInvoiceText(item?.sucursal, 120),
        createdAt: String(item?.createdAt || ''),
        documentosAgregados
    };
}

function updateOperatorDocumentAlertsButtonState() {
    if (!ui?.operatorDocAlertsBtn) {
        return;
    }

    const pendingCount = Array.isArray(operatorDocumentAlerts) ? operatorDocumentAlerts.length : 0;
    ui.operatorDocAlertsBtn.textContent = pendingCount > 0 ? `DOCS NUEVOS (${pendingCount})` : 'DOCS NUEVOS';
    ui.operatorDocAlertsBtn.classList.toggle('has-pending', pendingCount > 0);
}

function syncOperatorDocumentAlertsButtonVisibility(user) {
    if (!ui?.operatorDocAlertsBtn) {
        return;
    }

    const canUseAlerts = canUseOperatorDocumentAlerts(user);
    ui.operatorDocAlertsBtn.classList.toggle('hidden', !canUseAlerts);
    if (!canUseAlerts) {
        operatorDocumentAlerts = [];
        updateOperatorDocumentAlertsButtonState();
        closeOperatorDocumentAlertsModal();
        return;
    }

    updateOperatorDocumentAlertsButtonState();
}

function renderOperatorDocumentAlertsRows(records) {
    if (!ui?.operatorDocAlertsBody) {
        return;
    }

    ui.operatorDocAlertsBody.innerHTML = '';
    if (!Array.isArray(records) || records.length === 0) {
        ui.operatorDocAlertsBody.innerHTML = '<tr><td colspan="5">Sin documentos nuevos por revisar.</td></tr>';
        return;
    }

    records.forEach((item) => {
        const row = document.createElement('tr');
        const docs = Array.isArray(item?.documentosAgregados) ? item.documentosAgregados : [];
        const docsListHtml =
            docs.length > 0
                ? `<ul class="operator-doc-alerts-doc-list">${docs
                      .map((doc) => `<li>${escapeHtml(getDocumentOptionLabelByValue(doc) || doc)}</li>`)
                      .join('')}</ul>`
                : '<span>Sin detalle.</span>';

        row.innerHTML = `
            <td>${escapeHtml(item?.numIngreso || '')}</td>
            <td>${escapeHtml(formatDateTime(item?.createdAt || ''))}</td>
            <td>${escapeHtml(item?.sucursal || '')}</td>
            <td>${docsListHtml}</td>
            <td></td>
        `;

        const actionCell = row.lastElementChild;
        const reviewLabel = document.createElement('label');
        reviewLabel.className = 'operator-doc-alerts-review-label';
        const reviewInput = document.createElement('input');
        reviewInput.type = 'checkbox';
        reviewInput.addEventListener('change', () => {
            if (!reviewInput.checked) {
                return;
            }
            void handleMarkOperatorDocumentAlertReviewed(Number(item?.id), reviewInput);
        });
        reviewLabel.appendChild(reviewInput);
        reviewLabel.appendChild(document.createTextNode('Revisado'));
        actionCell.appendChild(reviewLabel);

        ui.operatorDocAlertsBody.appendChild(row);
    });
}

function closeOperatorDocumentAlertsModal() {
    if (ui?.operatorDocAlertsOverlay) {
        ui.operatorDocAlertsOverlay.classList.add('hidden');
    }
}

function openOperatorDocumentAlertsModal() {
    if (!ui?.operatorDocAlertsOverlay) {
        return;
    }

    renderOperatorDocumentAlertsRows(operatorDocumentAlerts);
    if (operatorDocumentAlerts.length === 0) {
        setFormMessage(ui.operatorDocAlertsMessage, 'Sin documentos nuevos por revisar.', 'success');
    } else {
        setFormMessage(
            ui.operatorDocAlertsMessage,
            `Pendientes por revisar: ${operatorDocumentAlerts.length}.`,
            'success'
        );
    }
    ui.operatorDocAlertsOverlay.classList.remove('hidden');
    if (ui.operatorDocAlertsCloseBtn) {
        window.requestAnimationFrame(() => {
            ui.operatorDocAlertsCloseBtn.focus();
        });
    }
}

async function fetchOperatorDocumentAlerts() {
    const result = await apiRequest('/registros/documentos-alertas');
    const records = Array.isArray(result?.alertas) ? result.alertas : [];
    operatorDocumentAlerts = records
        .map((item) => sanitizeOperatorDocumentAlertRecord(item))
        .filter((item) => item.id > 0);
    updateOperatorDocumentAlertsButtonState();
    return operatorDocumentAlerts;
}

async function refreshOperatorDocumentAlerts({ silent = true } = {}) {
    if (!canUseOperatorDocumentAlerts(currentUser)) {
        operatorDocumentAlerts = [];
        updateOperatorDocumentAlertsButtonState();
        renderOperatorDocumentAlertsRows([]);
        return [];
    }

    try {
        await fetchOperatorDocumentAlerts();
        renderOperatorDocumentAlertsRows(operatorDocumentAlerts);
        if (!silent) {
            if (operatorDocumentAlerts.length === 0) {
                setFormMessage(ui?.operatorDocAlertsMessage, 'Sin documentos nuevos por revisar.', 'success');
            } else {
                setFormMessage(
                    ui?.operatorDocAlertsMessage,
                    `Pendientes por revisar: ${operatorDocumentAlerts.length}.`,
                    'success'
                );
            }
        }
        return operatorDocumentAlerts;
    } catch (error) {
        if (silent) {
            console.warn('[Geo Rural] No fue posible actualizar alertas de documentos para operador.', error);
            return [];
        }
        setFormMessage(ui?.operatorDocAlertsMessage, error.message, 'error');
        throw error;
    }
}

async function handleMarkOperatorDocumentAlertReviewed(alertId, inputElement) {
    if (!canUseOperatorDocumentAlerts(currentUser)) {
        setFormMessage(ui?.operatorDocAlertsMessage, 'Solo OPERADOR o SUPER pueden revisar esta lista.', 'error');
        if (inputElement) {
            inputElement.checked = false;
        }
        return;
    }

    if (!Number.isInteger(alertId) || alertId <= 0) {
        setFormMessage(ui?.operatorDocAlertsMessage, 'No fue posible identificar el item seleccionado.', 'error');
        if (inputElement) {
            inputElement.checked = false;
        }
        return;
    }

    if (inputElement) {
        inputElement.disabled = true;
    }

    try {
        const result = await apiRequest(`/registros/documentos-alertas/${encodeURIComponent(alertId)}/revisar`, {
            method: 'POST',
            body: JSON.stringify({})
        });
        operatorDocumentAlerts = operatorDocumentAlerts.filter((item) => Number(item?.id) !== alertId);
        updateOperatorDocumentAlertsButtonState();
        renderOperatorDocumentAlertsRows(operatorDocumentAlerts);
        if (operatorDocumentAlerts.length === 0) {
            setFormMessage(ui?.operatorDocAlertsMessage, 'Sin documentos nuevos por revisar.', 'success');
        } else {
            setFormMessage(
                ui?.operatorDocAlertsMessage,
                result?.message || `Item revisado. Pendientes: ${operatorDocumentAlerts.length}.`,
                'success'
            );
        }
    } catch (error) {
        if (inputElement) {
            inputElement.checked = false;
        }
        setFormMessage(ui?.operatorDocAlertsMessage, error.message, 'error');
    } finally {
        if (inputElement) {
            inputElement.disabled = false;
        }
    }
}

async function handleOpenOperatorDocumentAlertsModal() {
    if (!canUseOperatorDocumentAlerts(currentUser)) {
        setFormMessage(ui?.formMessage, 'Solo OPERADOR o SUPER pueden ver esta lista.', 'error');
        return;
    }

    try {
        await refreshOperatorDocumentAlerts({ silent: false });
        openOperatorDocumentAlertsModal();
    } catch (error) {
        setFormMessage(ui?.formMessage, error.message, 'error');
    }
}

function parseInvoiceAmountNumber(value) {
    if (typeof value === 'number') {
        return Number.isFinite(value) ? value : NaN;
    }

    const raw = String(value || '').trim();
    if (!raw) {
        return NaN;
    }

    const compact = raw.replace(/\$/g, '').replace(/\s+/g, '').replace(/[^0-9,.-]/g, '');
    if (!compact) {
        return NaN;
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
        } else if (/^\d+,\d{1,2}$/.test(compact)) {
            normalized = compact.replace(',', '.');
        } else {
            normalized = compact.replace(/,/g, '');
        }
    } else if (hasDot) {
        if (/^\d{1,3}(\.\d{3})+$/.test(compact)) {
            normalized = compact.replace(/\./g, '');
        } else if (/^\d+\.\d{1,2}$/.test(compact)) {
            normalized = compact;
        } else {
            normalized = compact.replace(/\./g, '');
        }
    }

    const parsed = Number(normalized);
    if (!Number.isFinite(parsed) || parsed < 0) {
        return NaN;
    }

    return parsed;
}

function parseInvoiceAmountDigitsValue(value) {
    const digitsOnly = String(value || '').replace(/\D/g, '');
    if (!digitsOnly) {
        return NaN;
    }

    const parsed = Number(digitsOnly);
    if (!Number.isFinite(parsed) || parsed < 0) {
        return NaN;
    }

    return parsed;
}

function formatInvoiceAmountForInput(value) {
    const parsed = parseInvoiceAmountNumber(value);
    if (!Number.isFinite(parsed) || parsed < 0) {
        return '';
    }

    const rounded = Math.round(parsed);
    return `$ ${rounded.toLocaleString('es-CL')}`;
}

function applyFacturaMontoInputFormatting(inputElement) {
    if (!inputElement) {
        return;
    }

    const rawValue = String(inputElement.value || '');
    if (!rawValue.trim()) {
        inputElement.value = '';
        return;
    }

    const cursorStart = typeof inputElement.selectionStart === 'number' ? inputElement.selectionStart : null;
    const parsedDigitsValue = parseInvoiceAmountDigitsValue(rawValue);
    const formattedValue = Number.isFinite(parsedDigitsValue) ? `$ ${parsedDigitsValue.toLocaleString('es-CL')}` : '';
    inputElement.value = formattedValue;

    if (cursorStart === null || document.activeElement !== inputElement) {
        return;
    }

    const valueBeforeCursor = rawValue.slice(0, cursorStart);
    const parsedDigitsBeforeCursor = parseInvoiceAmountDigitsValue(valueBeforeCursor);
    const formattedBeforeCursor = Number.isFinite(parsedDigitsBeforeCursor)
        ? `$ ${parsedDigitsBeforeCursor.toLocaleString('es-CL')}`
        : '';
    const nextCursor = Math.min(formattedBeforeCursor.length, formattedValue.length);
    inputElement.setSelectionRange(nextCursor, nextCursor);
}

function normalizeInvoiceAmount(value) {
    const parsed = parseInvoiceAmountNumber(value);
    if (!Number.isFinite(parsed) || parsed <= 0) {
        return '';
    }

    return Math.round(parsed).toFixed(2);
}

function formatInvoiceAmount(value) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed) || parsed <= 0) {
        return '-';
    }

    return `$ ${parsed.toLocaleString('es-CL', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
    })}`;
}

function buildInvoiceComunaRegionLabel(comuna, ciudad) {
    return [normalizeInvoiceText(comuna, 120), normalizeInvoiceText(ciudad, 120)].filter(Boolean).join(' / ');
}

function sanitizePendingInvoiceDraft(rawItem) {
    if (!rawItem || typeof rawItem !== 'object') {
        return null;
    }

    const draft = {
        id: normalizeInvoiceText(rawItem.id, 80) || generateInvoiceDraftId(),
        createdAt: normalizeInvoiceText(rawItem.createdAt, 40) || new Date().toISOString(),
        createdBy: normalizeInvoiceText(rawItem.createdBy, 120),
        sucursal: normalizeInvoiceText(rawItem.sucursal, 120),
        numIngreso: normalizeInvoiceText(rawItem.numIngreso, 20),
        nombreRazonSocial: normalizeInvoiceText(rawItem.nombreRazonSocial || rawItem.nombreCliente, 255),
        rut: normalizeInvoiceText(rawItem.rut || rawItem.rutCliente, 20),
        giro: normalizeInvoiceText(rawItem.giro, 255),
        direccion: normalizeInvoiceText(rawItem.direccion, 255),
        comuna: normalizeInvoiceText(rawItem.comuna, 120),
        ciudad: normalizeInvoiceText(rawItem.ciudad, 120),
        contacto: normalizeInvoiceText(rawItem.contacto || rawItem.correoFactura, 255),
        observacion: normalizeInvoiceText(rawItem.observacion || rawItem.referencia, 1000),
        montoFacturar: normalizeInvoiceAmount(rawItem.montoFacturar || rawItem.monto || '')
    };

    if (!draft.nombreRazonSocial || !draft.rut) {
        return null;
    }

    if (draft.numIngreso && !/^\d+-\d{4}$/.test(draft.numIngreso)) {
        draft.numIngreso = '';
    }

    return draft;
}

function normalizeInvoiceRequestStatus(value) {
    const status = String(value || '').trim().toUpperCase();
    if (status === 'PENDIENTE') {
        return 'PENDIENTE';
    }
    if (status === 'ENVIADA') {
        return 'ENVIADA';
    }
    if (status === 'ANULADA') {
        return 'ANULADA';
    }

    return '';
}

function formatInvoiceRequestStatus(status) {
    const normalized = normalizeInvoiceRequestStatus(status);
    if (normalized === 'ENVIADA') {
        return 'Enviada';
    }
    if (normalized === 'PENDIENTE') {
        return 'Pendiente';
    }
    if (normalized === 'ANULADA') {
        return 'Anulada';
    }

    return 'Sin estado';
}

function applyInvoiceSendButtonsState() {
    const isGuest = isGuestUser(currentUser);
    const canUseWorkflow = canUseInvoiceWorkflow(currentUser);
    const isSent = normalizeInvoiceRequestStatus(currentInvoiceRequestStatus) === 'ENVIADA';
    const disableSend = isGuest || isSent || !canUseWorkflow;
    const disableQueue = isGuest || isSent || !canUseWorkflow;
    const sendButtons = [ui?.operatorEmitInvoiceBtn, ui?.facturaEnviarContadorBtn];
    const queueButtons = [ui?.operatorQueueInvoiceBtn, ui?.facturaAgregarListaBtn];

    sendButtons.forEach((button) => {
        if (!button) {
            return;
        }

        button.disabled = disableSend;
        if (!canUseWorkflow && !isGuest) {
            button.title = 'Tu rol no tiene permisos para gestionar facturas.';
            return;
        }
        if (isSent && !isGuest) {
            button.title = 'Esta factura ya fue enviada al contador.';
            return;
        }

        button.removeAttribute('title');
    });

    queueButtons.forEach((button) => {
        if (!button) {
            return;
        }

        button.disabled = disableQueue;
        if (!canUseWorkflow && !isGuest) {
            button.title = 'Tu rol no tiene permisos para gestionar facturas.';
            return;
        }
        if (isSent && !isGuest) {
            button.title = 'Esta factura ya fue enviada al contador y no puede volver a lista pendiente.';
            return;
        }

        button.removeAttribute('title');
    });
}

function setCurrentInvoiceRequestStatus(status) {
    currentInvoiceRequestStatus = normalizeInvoiceRequestStatus(status);
    applyInvoiceSendButtonsState();
}

function clearCurrentInvoiceRequestStatus() {
    setCurrentInvoiceRequestStatus('');
}

async function syncCurrentInvoiceRequestStatusByIngreso(numIngreso, options = {}) {
    const normalizedIngreso = String(numIngreso || '').trim();
    if (!canUseInvoiceWorkflow(currentUser)) {
        clearCurrentInvoiceRequestStatus();
        return;
    }
    if (!normalizedIngreso || !/^\d+-\d{4}$/.test(normalizedIngreso)) {
        clearCurrentInvoiceRequestStatus();
        return;
    }

    const showErrors = options.showErrors === true;
    const targetMessage = options.messageElement || ui?.facturaFormMessage || ui?.operatorInvoiceMessage;

    try {
        const result = await apiRequest(`/facturas/solicitudes/por-ingreso/${encodeURIComponent(normalizedIngreso)}`);
        const solicitud = result && result.solicitud ? result.solicitud : null;
        const status = normalizeInvoiceRequestStatus(solicitud?.estado);
        setCurrentInvoiceRequestStatus(status);
    } catch (error) {
        clearCurrentInvoiceRequestStatus();
        if (showErrors) {
            setFormMessage(targetMessage, error.message, 'error');
        }
    }
}

function sanitizeRequestedInvoiceHistoryEntry(rawItem) {
    const draft = sanitizePendingInvoiceDraft(rawItem);
    if (!draft) {
        return null;
    }

    const status = normalizeInvoiceRequestStatus(rawItem?.estado) || 'PENDIENTE';
    const destinationEmail = normalizeInvoiceText(rawItem?.destinationEmail, 255).toLowerCase();
    const requestedAt = normalizeInvoiceText(rawItem?.createdAt || rawItem?.requestedAt, 40) || new Date().toISOString();
    return {
        ...draft,
        requestId: normalizeInvoiceText(rawItem?.id, 80) || `req-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
        requestedAt,
        estado: status,
        destinationEmail,
        sentAt: normalizeInvoiceText(rawItem?.sentAt, 40),
        requestedBy: normalizeInvoiceText(rawItem?.createdBy || rawItem?.requestedBy, 120) || draft.createdBy,
        requestedBySucursal: normalizeInvoiceText(rawItem?.sucursal || rawItem?.requestedBySucursal, 120) || draft.sucursal
    };
}

function buildInvoiceRequestsQuery(estado = '', limit = 300) {
    const query = new URLSearchParams();
    const normalizedStatus = normalizeInvoiceRequestStatus(estado);
    if (normalizedStatus) {
        query.set('estado', normalizedStatus);
    }
    const safeLimit = Number.isInteger(Number(limit)) ? Math.max(1, Math.min(Number(limit), 1000)) : 300;
    query.set('limit', String(safeLimit));
    return query.toString();
}

async function fetchInvoiceRequests(estado = '', limit = 300) {
    const query = buildInvoiceRequestsQuery(estado, limit);
    const result = await apiRequest(`/facturas/solicitudes?${query}`);
    if (!Array.isArray(result.solicitudes)) {
        return [];
    }

    return result.solicitudes.map(sanitizeRequestedInvoiceHistoryEntry).filter((item) => Boolean(item));
}

async function refreshOperatorPendingInvoices(options = {}) {
    if (!canUseInvoiceWorkflow(currentUser)) {
        operatorPendingInvoices = [];
        renderOperatorPendingInvoices();
        return false;
    }

    const showErrors = options.showErrors === true;
    const targetMessage = options.messageElement || ui?.operatorInvoiceMessage || ui?.facturaFormMessage;

    try {
        const pendingInvoices = await fetchInvoiceRequests('PENDIENTE', 500);
        operatorPendingInvoices = pendingInvoices.map((item) => sanitizePendingInvoiceDraft(item)).filter((item) => Boolean(item));
        renderOperatorPendingInvoices();
        return true;
    } catch (error) {
        if (showErrors) {
            setFormMessage(targetMessage, error.message, 'error');
        }
        return false;
    }
}

async function refreshRequestedInvoicesHistory(options = {}) {
    if (!canUseSecretaryFeatures(currentUser)) {
        requestedInvoicesHistory = [];
        selectedRequestedInvoiceRequestId = '';
        updateSecretariaInvoiceDetailButtonState();
        renderRequestedInvoicesHistory();
        return false;
    }

    const showErrors = options.showErrors === true;
    const targetMessage = options.messageElement || ui?.secretariaInvoicesMessage;
    const previousSelectedId = selectedRequestedInvoiceRequestId;

    try {
        requestedInvoicesHistory = await fetchInvoiceRequests('', 500);
        if (previousSelectedId && requestedInvoicesHistory.some((item) => String(item.requestId || '') === previousSelectedId)) {
            selectedRequestedInvoiceRequestId = previousSelectedId;
        } else {
            selectedRequestedInvoiceRequestId = '';
        }
        updateSecretariaInvoiceDetailButtonState();
        renderRequestedInvoicesHistory();
        return true;
    } catch (error) {
        selectedRequestedInvoiceRequestId = '';
        updateSecretariaInvoiceDetailButtonState();
        if (showErrors) {
            setFormMessage(targetMessage, error.message, 'error');
        }
        return false;
    }
}

function renderRequestedInvoicesHistory() {
    if (!ui || !ui.secretariaInvoicesBody || !ui.secretariaInvoicesSummary) {
        return;
    }

    const historyItems = Array.isArray(requestedInvoicesHistory) ? requestedInvoicesHistory : [];
    const filteredItems = getFilteredRequestedInvoicesHistory();
    const filters = getSecretariaInvoicesFilterValues();
    const historyCount = historyItems.length;
    const shownCount = filteredItems.length;
    const pendingCount = filteredItems.filter((item) => normalizeInvoiceRequestStatus(item.estado) === 'PENDIENTE').length;
    const pagination = getRequestedInvoicesPaginationInfo(filteredItems);
    const pageItems = pagination.pageItems;
    ui.secretariaInvoicesSummary.textContent = filters.hasFilters
        ? `Solicitudes registradas: ${historyCount} | Coinciden: ${shownCount} | Pendientes: ${pendingCount} | Grupo ${pagination.currentPage}/${pagination.totalPages}`
        : `Solicitudes registradas: ${historyCount} | Pendientes: ${pendingCount} | Grupo ${pagination.currentPage}/${pagination.totalPages}`;
    ui.secretariaInvoicesBody.innerHTML = '';

    if (!shownCount) {
        ui.secretariaInvoicesBody.innerHTML = filters.hasFilters
            ? '<tr><td colspan="8">No hay resultados para los filtros aplicados.</td></tr>'
            : '<tr><td colspan="8">Sin facturas solicitadas para mostrar.</td></tr>';
        selectedRequestedInvoiceRequestId = '';
        updateSecretariaInvoiceDetailButtonState();
        renderRequestedInvoicesPaginationControls(pagination);
        return;
    }

    const hasSelectedVisibleItem = pageItems.some(
        (item) => String(item.requestId || '') === selectedRequestedInvoiceRequestId
    );
    if (!hasSelectedVisibleItem) {
        selectedRequestedInvoiceRequestId = '';
    }

    pageItems.forEach((item) => {
        const row = document.createElement('tr');
        row.classList.add('secretaria-request-row');
        const isSelected = String(item.requestId || '') === selectedRequestedInvoiceRequestId;
        if (isSelected) {
            row.classList.add('is-selected');
        }

        row.innerHTML = `
            <td><input type="radio" name="secretariaInvoiceSelect" ${isSelected ? 'checked' : ''} aria-label="Seleccionar factura"></td>
            <td>${escapeHtml(formatDateTime(item.requestedAt))}</td>
            <td>${escapeHtml(item.numIngreso || '-')}</td>
            <td>${escapeHtml(item.nombreRazonSocial || '-')}</td>
            <td>${escapeHtml(item.rut || '-')}</td>
            <td>${escapeHtml(formatInvoiceAmount(item.montoFacturar))}</td>
            <td>${escapeHtml(formatInvoiceRequestStatus(item.estado))}</td>
            <td>${escapeHtml(item.destinationEmail || '-')}</td>
        `;
        row.addEventListener('click', () => {
            setSelectedRequestedInvoiceRequest(item.requestId);
        });
        const radioInput = row.querySelector('input[type="radio"]');
        if (radioInput) {
            radioInput.addEventListener('click', (event) => {
                event.stopPropagation();
                setSelectedRequestedInvoiceRequest(item.requestId);
            });
            radioInput.addEventListener('change', () => {
                setSelectedRequestedInvoiceRequest(item.requestId);
            });
        }
        ui.secretariaInvoicesBody.appendChild(row);
    });
    updateSecretariaInvoiceDetailButtonState();
    renderRequestedInvoicesPaginationControls(pagination);
}

function renderOperatorPendingInvoices() {
    if (!ui || !ui.operatorPendingInvoicesBody || !ui.operatorPendingSummary) {
        return;
    }

    const pendingCount = Array.isArray(operatorPendingInvoices) ? operatorPendingInvoices.length : 0;
    ui.operatorPendingSummary.textContent = `Facturas pendientes: ${pendingCount}`;
    ui.operatorPendingInvoicesBody.innerHTML = '';

    if (!pendingCount) {
        ui.operatorPendingInvoicesBody.innerHTML = '<tr><td colspan="6">Sin facturas pendientes para enviar.</td></tr>';
        return;
    }

    operatorPendingInvoices.forEach((item) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${escapeHtml(formatDateTime(item.createdAt))}</td>
            <td>${escapeHtml(item.numIngreso || '-')}</td>
            <td>${escapeHtml(item.nombreRazonSocial || '-')}</td>
            <td>${escapeHtml(item.rut || '-')}</td>
            <td>${escapeHtml(item.contacto || '-')}</td>
            <td>${escapeHtml(formatInvoiceAmount(item.montoFacturar))}</td>
        `;
        ui.operatorPendingInvoicesBody.appendChild(row);
    });
}

function setFacturaFieldIfEmpty(field, value, options = {}) {
    if (!field) {
        return;
    }

    const overwrite = options.overwrite === true;
    const targetValue = normalizeInvoiceText(value, Number(field.maxLength) || 255);
    if (!targetValue) {
        return;
    }

    if (!overwrite && normalizeInvoiceText(field.value, Number(field.maxLength) || 255)) {
        return;
    }

    field.value = targetValue;
}

function buildMainFormContactoSuggestion() {
    const correo = normalizeEmailValue(document.getElementById('correo')?.value);
    return normalizeInvoiceText(correo, 255);
}

function syncFacturaFromMainForm(options = {}) {
    if (!ui) {
        return;
    }

    setFacturaFieldIfEmpty(ui.facturaNombreRazon, document.getElementById('nombre')?.value, options);
    setFacturaFieldIfEmpty(ui.facturaRut, document.getElementById('rut')?.value, options);
    setFacturaFieldIfEmpty(ui.facturaComuna, document.getElementById('comuna')?.value, options);
    setFacturaFieldIfEmpty(ui.facturaCiudad, document.getElementById('region')?.value, options);
    setFacturaFieldIfEmpty(ui.facturaContacto, buildMainFormContactoSuggestion(), options);
}

function registerInvoiceAutofillBindings() {
    const sourceIds = ['nombre', 'rut', 'comuna', 'region', 'telefono', 'correo'];
    sourceIds.forEach((sourceId) => {
        const sourceNode = document.getElementById(sourceId);
        if (!sourceNode) {
            return;
        }

        sourceNode.addEventListener('change', () => {
            syncFacturaFromMainForm({ overwrite: true });
        });
        sourceNode.addEventListener('blur', () => {
            syncFacturaFromMainForm({ overwrite: true });
        });
        sourceNode.addEventListener('keydown', (event) => {
            if (event.key !== 'Enter') {
                return;
            }
            requestAnimationFrame(() => {
                syncFacturaFromMainForm({ overwrite: true });
            });
        });
    });
}

function resetFacturaForm() {
    if (!ui) {
        return;
    }

    const fields = [
        ui.facturaNombreRazon,
        ui.facturaNumero,
        ui.facturaRut,
        ui.facturaGiro,
        ui.facturaDireccion,
        ui.facturaComuna,
        ui.facturaCiudad,
        ui.facturaContacto,
        ui.facturaObservacion,
        ui.facturaMonto
    ];
    fields.forEach((field) => {
        if (field) {
            field.value = '';
        }
    });
    clearFacturaFieldErrors();
    clearCurrentInvoiceRequestStatus();
    setFormMessage(ui.facturaFormMessage, '', '');
    syncFacturaFromMainForm();
    syncFacturaNumeroVisibility({ clearOnHide: false });
}

function collectFacturaFormData() {
    const factura = collectFacturaSnapshotData();
    return {
        id: '',
        createdAt: new Date().toISOString(),
        createdBy: currentUser?.nombre || currentUser?.username || '',
        sucursal: currentUser?.sucursal || '',
        numIngreso: normalizeInvoiceText(ui.numIngreso?.value, 20),
        ...factura
    };
}

function collectFacturaSnapshotData() {
    return {
        nombreRazonSocial: normalizeInvoiceText(ui.facturaNombreRazon?.value, 255),
        numeroFactura: normalizeInvoiceText(ui.facturaNumero?.value, FACTURA_NUMBER_MAX_LENGTH),
        rut: normalizeInvoiceText(ui.facturaRut?.value, 20),
        giro: normalizeInvoiceText(ui.facturaGiro?.value, 255),
        direccion: normalizeInvoiceText(ui.facturaDireccion?.value, 255),
        comuna: normalizeInvoiceText(ui.facturaComuna?.value, 120),
        ciudad: normalizeInvoiceText(ui.facturaCiudad?.value, 120),
        contacto: normalizeInvoiceText(ui.facturaContacto?.value, 255),
        observacion: normalizeInvoiceText(ui.facturaObservacion?.value, 1000),
        montoFacturar: normalizeInvoiceAmount(ui.facturaMonto?.value)
    };
}

function applyFacturaSnapshotToForm(factura) {
    if (!ui || !factura || typeof factura !== 'object') {
        return;
    }

    if (ui.facturaNombreRazon) {
        ui.facturaNombreRazon.value = normalizeInvoiceText(factura.nombreRazonSocial, 255);
    }
    if (ui.facturaNumero) {
        ui.facturaNumero.value = normalizeInvoiceText(factura.numeroFactura, FACTURA_NUMBER_MAX_LENGTH);
    }
    if (ui.facturaRut) {
        ui.facturaRut.value = normalizeInvoiceText(factura.rut, 20);
    }
    if (ui.facturaGiro) {
        ui.facturaGiro.value = normalizeInvoiceText(factura.giro, 255);
    }
    if (ui.facturaDireccion) {
        ui.facturaDireccion.value = normalizeInvoiceText(factura.direccion, 255);
    }
    if (ui.facturaComuna) {
        ui.facturaComuna.value = normalizeInvoiceText(factura.comuna, 120);
    }
    if (ui.facturaCiudad) {
        ui.facturaCiudad.value = normalizeInvoiceText(factura.ciudad, 120);
    }
    if (ui.facturaContacto) {
        ui.facturaContacto.value = normalizeInvoiceText(factura.contacto, 255);
    }
    if (ui.facturaObservacion) {
        ui.facturaObservacion.value = normalizeInvoiceText(factura.observacion, 1000);
    }
    if (ui.facturaMonto) {
        const normalizedMonto = normalizeInvoiceAmount(factura.montoFacturar);
        ui.facturaMonto.value = normalizedMonto ? formatInvoiceAmountForInput(normalizedMonto) : '';
    }
}

function clearFacturaFieldErrors() {
    FACTURA_INLINE_ERROR_FIELD_IDS.forEach((fieldId) => {
        clearRegistroFieldError(fieldId);
    });
}

function resolveFacturaFieldIdFromErrorMessage(message) {
    const normalizedMessage = String(message || '').trim().toLowerCase();
    if (!normalizedMessage) {
        return '';
    }
    const compactMessage = normalizedMessage
        .replace(/\s*\/\s*/g, '/')
        .replace(/\s+/g, ' ');

    if (compactMessage.includes('nombre/razon social') || compactMessage.includes('nombre o razon social')) {
        return 'facturaNombreRazon';
    }
    if (compactMessage.includes('factura rut') || compactMessage.includes('rut invalido')) {
        return 'facturaRut';
    }
    if (compactMessage.includes('factura numero')) {
        return 'facturaNumero';
    }
    if (compactMessage.includes('giro')) {
        return 'facturaGiro';
    }
    if (compactMessage.includes('direccion')) {
        return 'facturaDireccion';
    }
    if (compactMessage.includes('factura comuna')) {
        return 'facturaComuna';
    }
    if (compactMessage.includes('factura ciudad')) {
        return 'facturaCiudad';
    }
    if (compactMessage.includes('factura contacto')) {
        return 'facturaContacto';
    }
    if (compactMessage.includes('observacion')) {
        return 'facturaObservacion';
    }
    if (compactMessage.includes('monto a facturar')) {
        return 'facturaMonto';
    }

    return '';
}

function applyFacturaInlineErrorMessage(message, preferredFieldId = '') {
    const normalizedMessage = String(message || '').trim();
    if (!normalizedMessage) {
        return false;
    }

    const preferred = String(preferredFieldId || '').trim();
    const resolvedFieldId = FACTURA_INLINE_ERROR_FIELD_IDS.includes(preferred)
        ? preferred
        : resolveFacturaFieldIdFromErrorMessage(normalizedMessage);
    if (!resolvedFieldId) {
        return false;
    }

    setRegistroFieldError(resolvedFieldId, normalizedMessage);
    const targetField = document.getElementById(resolvedFieldId);
    if (targetField && typeof targetField.focus === 'function') {
        targetField.focus();
    }
    return true;
}

function registerFacturaInlineValidationListeners() {
    FACTURA_INLINE_ERROR_FIELD_IDS.forEach((fieldId) => {
        const field = document.getElementById(fieldId);
        if (!field) {
            return;
        }

        const syncFieldErrorState = () => {
            const value = String(field.value || '').trim();
            if (!value) {
                if (FACTURA_REQUIRED_FIELD_IDS.includes(fieldId)) {
                    setRegistroFieldError(fieldId, 'Campo obligatorio.');
                    return;
                }
                clearRegistroFieldError(fieldId);
                return;
            }

            if (fieldId === 'facturaMonto' && !normalizeInvoiceAmount(value)) {
                setRegistroFieldError(fieldId, 'Ingresa un MONTO A FACTURAR valido (mayor a 0).');
                return;
            }

            clearRegistroFieldError(fieldId);
        };

        field.addEventListener('input', syncFieldErrorState);
        field.addEventListener('change', syncFieldErrorState);
        field.addEventListener('blur', syncFieldErrorState);
    });
}

function validateInvoiceDraft(rawData, options = {}) {
    if (!rawData) {
        return { error: 'No fue posible leer los datos de factura.', data: null, errorFieldId: '' };
    }

    const requireAmount = options.requireAmount !== false;
    const data = sanitizePendingInvoiceDraft(rawData);
    if (!data) {
        const nombreRazonSocial = normalizeInvoiceText(rawData?.nombreRazonSocial || rawData?.nombreCliente, 255);
        const rut = normalizeInvoiceText(rawData?.rut || rawData?.rutCliente, 20);
        if (!nombreRazonSocial) {
            return { error: 'Completa NOMBRE / RAZON SOCIAL para la factura.', data: null, errorFieldId: 'facturaNombreRazon' };
        }
        if (!rut) {
            return { error: 'Completa RUT para la factura.', data: null, errorFieldId: 'facturaRut' };
        }
        return { error: 'Completa nombre/razon social y RUT para la factura.', data: null, errorFieldId: 'facturaNombreRazon' };
    }

    if (!data.giro) {
        return { error: 'Debes completar el GIRO para la factura.', data: null, errorFieldId: 'facturaGiro' };
    }

    if (!data.direccion) {
        return { error: 'Debes completar la DIRECCION para la factura.', data: null, errorFieldId: 'facturaDireccion' };
    }

    if (requireAmount && !data.montoFacturar) {
        return { error: 'Ingresa un MONTO A FACTURAR valido (mayor a 0).', data: null, errorFieldId: 'facturaMonto' };
    }

    if (!data.numIngreso) {
        return { error: 'Debes cargar un registro con NRO INGRESO antes de enviar o agregar factura.', data: null, errorFieldId: '' };
    }

    if (!/^\d+-\d{4}$/.test(data.numIngreso)) {
        return { error: 'NRO INGRESO invalido. Usa formato 123-2026.', data: null, errorFieldId: '' };
    }

    return { error: '', data, errorFieldId: '' };
}

function getLastInvoiceDestinationEmail() {
    try {
        return normalizeInvoiceText(window.localStorage.getItem(OPERATOR_LAST_DEST_EMAIL_STORAGE), 255).toLowerCase();
    } catch (error) {
        return '';
    }
}

function persistLastInvoiceDestinationEmail(email) {
    try {
        window.localStorage.setItem(OPERATOR_LAST_DEST_EMAIL_STORAGE, normalizeInvoiceText(email, 255).toLowerCase());
    } catch (error) {
        // Si localStorage falla, se ignora.
    }
}

function closeInvoiceEmailPrompt(resolvedEmail = null) {
    if (ui?.invoiceEmailPromptOverlay) {
        ui.invoiceEmailPromptOverlay.classList.add('hidden');
    }
    if (ui?.invoiceEmailPromptInput) {
        ui.invoiceEmailPromptInput.value = '';
    }
    setFormMessage(ui?.invoiceEmailPromptMessage, '', '');

    if (!invoiceEmailPromptState) {
        return;
    }

    const pendingPrompt = invoiceEmailPromptState;
    invoiceEmailPromptState = null;
    pendingPrompt.resolve(resolvedEmail);
}

function cancelInvoiceEmailPrompt() {
    closeInvoiceEmailPrompt(null);
}

function confirmInvoiceEmailPrompt() {
    if (!invoiceEmailPromptState) {
        return;
    }

    const destinationEmail = normalizeInvoiceText(ui?.invoiceEmailPromptInput?.value, 255).toLowerCase();
    if (!destinationEmail) {
        setFormMessage(ui?.invoiceEmailPromptMessage, 'Debes ingresar el correo del contador.', 'error');
        if (ui?.invoiceEmailPromptInput) {
            ui.invoiceEmailPromptInput.focus();
        }
        return;
    }
    if (!isValidEmail(destinationEmail)) {
        setFormMessage(ui?.invoiceEmailPromptMessage, 'El correo del contador no es valido.', 'error');
        if (ui?.invoiceEmailPromptInput) {
            ui.invoiceEmailPromptInput.focus();
            ui.invoiceEmailPromptInput.select();
        }
        return;
    }

    persistLastInvoiceDestinationEmail(destinationEmail);
    closeInvoiceEmailPrompt(destinationEmail);
}

function openInvoiceEmailPrompt(defaultEmail = '') {
    if (!ui?.invoiceEmailPromptOverlay || !ui?.invoiceEmailPromptInput) {
        return Promise.resolve(null);
    }

    if (invoiceEmailPromptState) {
        invoiceEmailPromptState.resolve(null);
        invoiceEmailPromptState = null;
    }

    const normalizedEmail = normalizeInvoiceText(defaultEmail, 255).toLowerCase();
    ui.invoiceEmailPromptInput.value = normalizedEmail;
    setFormMessage(ui?.invoiceEmailPromptMessage, '', '');
    ui.invoiceEmailPromptOverlay.classList.remove('hidden');
    window.requestAnimationFrame(() => {
        ui.invoiceEmailPromptInput.focus();
        ui.invoiceEmailPromptInput.select();
    });

    return new Promise((resolve) => {
        invoiceEmailPromptState = { resolve };
    });
}

async function promptAccountantEmail(preferredEmail = '') {
    const normalizedPreferred = normalizeInvoiceText(preferredEmail, 255).toLowerCase();
    const suggestedEmail = normalizedPreferred || getLastInvoiceDestinationEmail();

    if (ui?.invoiceEmailPromptOverlay && ui?.invoiceEmailPromptInput) {
        return openInvoiceEmailPrompt(suggestedEmail);
    }

    const destinationInput = window.prompt('Ingresa el correo del contador:', suggestedEmail || '');
    if (destinationInput === null) {
        return null;
    }
    const destinationEmail = normalizeInvoiceText(destinationInput, 255).toLowerCase();
    if (!destinationEmail) {
        return '';
    }
    if (!isValidEmail(destinationEmail)) {
        return 'INVALID_EMAIL';
    }
    persistLastInvoiceDestinationEmail(destinationEmail);
    return destinationEmail;
}

function buildMailSubject(value) {
    return String(value || '')
        .replace(/[\r\n]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .replace(/([)\]\d])r$/i, '$1');
}

function toEmailHtmlMultiline(value) {
    return escapeHtml(String(value || '-')).replace(/\r\n|\r|\n/g, '<br>');
}

function buildInvoiceDetailHtmlRows(item) {
    return buildInvoiceDetailTableRows(item)
        .map(([label, value]) => {
            return `
                <tr>
                    <td style="width: 230px; padding: 8px 10px; border: 1px solid #d8e0ee; background: #eef3fb; font-weight: 700; color: #22365f;">
                        ${escapeHtml(label)}
                    </td>
                    <td style="padding: 8px 10px; border: 1px solid #d8e0ee; color: #10213f;">
                        ${toEmailHtmlMultiline(value)}
                    </td>
                </tr>
            `.trim();
        })
        .join('');
}

function buildInvoiceEmailCardHtml(title, summaryRowsHtml, contentHtml) {
    return `
        <!doctype html>
        <html lang="es">
            <body style="margin:0;padding:16px;background:#eef3fb;font-family:Arial,Helvetica,sans-serif;color:#10213f;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="max-width:920px;margin:0 auto;background:#ffffff;border:1px solid #d2dced;border-radius:10px;overflow:hidden;">
                    <tr>
                        <td style="padding:16px 18px;background:#d7e4f7;border-bottom:1px solid #c2d4ee;">
                            <h2 style="margin:0;font-size:20px;line-height:1.3;color:#163463;">${escapeHtml(title)}</h2>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:14px 18px;">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
                                ${summaryRowsHtml}
                            </table>
                            ${contentHtml}
                        </td>
                    </tr>
                </table>
            </body>
        </html>
    `.trim();
}

function buildSingleInvoiceEmailHtml(item) {
    const nowText = new Date().toLocaleString('es-CL', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    });
    const operatorName = currentUser?.nombre || currentUser?.username || 'Operador';
    const branch = currentUser?.sucursal || 'Sin sucursal';
    const summaryRowsHtml = `
        <tr>
            <td style="padding:0 0 6px;font-weight:700;color:#22365f;">Fecha</td>
            <td style="padding:0 0 6px;">${escapeHtml(nowText)}</td>
        </tr>
        <tr>
            <td style="padding:0 0 6px;font-weight:700;color:#22365f;">Operador</td>
            <td style="padding:0 0 6px;">${escapeHtml(operatorName)}</td>
        </tr>
        <tr>
            <td style="padding:0 0 12px;font-weight:700;color:#22365f;">Sucursal</td>
            <td style="padding:0 0 12px;">${escapeHtml(branch)}</td>
        </tr>
    `.trim();
    const detailRowsHtml = buildInvoiceDetailHtmlRows(item);
    const contentHtml = `
        <div style="margin-top:2px;padding-top:8px;">
            <h3 style="margin:0 0 10px;font-size:16px;color:#163463;">Detalle de factura</h3>
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
                ${detailRowsHtml}
            </table>
        </div>
    `.trim();

    const customTemplate = getRuntimeMailTemplateHtml('facturaSingleHtml');
    if (customTemplate) {
        return applyMailTemplateHtmlString(customTemplate, {
            TITULO: 'Solicitud de facturacion',
            RESUMEN_FILAS_HTML: summaryRowsHtml,
            DETALLE_FILAS_HTML: detailRowsHtml
        });
    }

    return buildInvoiceEmailCardHtml('Solicitud de facturacion', summaryRowsHtml, contentHtml);
}

function buildPendingInvoicesEmailHtml(items = operatorPendingInvoices) {
    const normalizedItems = Array.isArray(items) ? items.filter((item) => Boolean(item)) : [];
    const nowText = new Date().toLocaleString('es-CL', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    });
    const operatorName = currentUser?.nombre || currentUser?.username || 'Operador';
    const branch = currentUser?.sucursal || 'Sin sucursal';

    const summaryRowsHtml = `
        <tr>
            <td style="padding:0 0 6px;font-weight:700;color:#22365f;">Fecha</td>
            <td style="padding:0 0 6px;">${escapeHtml(nowText)}</td>
        </tr>
        <tr>
            <td style="padding:0 0 6px;font-weight:700;color:#22365f;">Operador</td>
            <td style="padding:0 0 6px;">${escapeHtml(operatorName)}</td>
        </tr>
        <tr>
            <td style="padding:0 0 6px;font-weight:700;color:#22365f;">Sucursal</td>
            <td style="padding:0 0 6px;">${escapeHtml(branch)}</td>
        </tr>
        <tr>
            <td style="padding:0 0 12px;font-weight:700;color:#22365f;">Total pendientes</td>
            <td style="padding:0 0 12px;">${escapeHtml(String(normalizedItems.length))}</td>
        </tr>
    `.trim();

    const summaryTableRowsHtml = normalizedItems
        .map((item, index) => {
            return `
                <tr>
                    <td style="padding:7px 8px;border:1px solid #d8e0ee;text-align:center;">${escapeHtml(String(index + 1))}</td>
                    <td style="padding:7px 8px;border:1px solid #d8e0ee;">${escapeHtml(item.numIngreso || '-')}</td>
                    <td style="padding:7px 8px;border:1px solid #d8e0ee;">${escapeHtml(item.nombreRazonSocial || '-')}</td>
                    <td style="padding:7px 8px;border:1px solid #d8e0ee;">${escapeHtml(formatInvoiceAmount(item.montoFacturar))}</td>
                    <td style="padding:7px 8px;border:1px solid #d8e0ee;">${escapeHtml(item.contacto || '-')}</td>
                </tr>
            `.trim();
        })
        .join('');

    const detailSectionsHtml = normalizedItems
        .map((item, index) => {
            return `
                <div style="margin-top:14px;">
                    <h4 style="margin:0 0 8px;font-size:14px;color:#163463;">Factura ${escapeHtml(String(index + 1))}</h4>
                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
                        ${buildInvoiceDetailHtmlRows(item)}
                    </table>
                </div>
            `.trim();
        })
        .join('');

    const contentHtml = `
        <div style="margin-top:2px;padding-top:8px;">
            <h3 style="margin:0 0 10px;font-size:16px;color:#163463;">Resumen de facturas pendientes</h3>
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
                <thead>
                    <tr>
                        <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:center;">#</th>
                        <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">NRO INGRESO</th>
                        <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">CLIENTE / RAZON SOCIAL</th>
                        <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">MONTO</th>
                        <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">CONTACTO</th>
                    </tr>
                </thead>
                <tbody>
                    ${summaryTableRowsHtml}
                </tbody>
            </table>
            ${detailSectionsHtml}
        </div>
    `.trim();

    const summaryTableHtml = `
        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
            <thead>
                <tr>
                    <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:center;">#</th>
                    <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">NRO INGRESO</th>
                    <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">CLIENTE / RAZON SOCIAL</th>
                    <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">MONTO</th>
                    <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">CONTACTO</th>
                </tr>
            </thead>
            <tbody>
                ${summaryTableRowsHtml}
            </tbody>
        </table>
    `.trim();

    const customTemplate = getRuntimeMailTemplateHtml('facturaPendingHtml');
    if (customTemplate) {
        return applyMailTemplateHtmlString(customTemplate, {
            TITULO: 'Resumen de facturas pendientes',
            RESUMEN_FILAS_HTML: summaryRowsHtml,
            TABLA_RESUMEN_HTML: summaryTableHtml,
            DETALLES_FACTURAS_HTML: detailSectionsHtml
        });
    }

    return buildInvoiceEmailCardHtml('Resumen de facturas pendientes', summaryRowsHtml, contentHtml);
}

function normalizeEmailTableValue(value, maxLength = 120) {
    const normalized = String(value || '')
        .replace(/[\r\n\t]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    const fallback = normalized || '-';
    if (!Number.isInteger(maxLength) || maxLength <= 4 || fallback.length <= maxLength) {
        return fallback;
    }

    return `${fallback.slice(0, maxLength - 3)}...`;
}

function buildReadableTextTable(headers, rows, options = {}) {
    const safeHeaders = Array.isArray(headers) ? headers.map((header) => normalizeEmailTableValue(header, 60)) : [];
    const safeRows = Array.isArray(rows)
        ? rows.map((row) => {
              const rowValues = Array.isArray(row) ? row : [];
              return safeHeaders.map((_, index) => normalizeEmailTableValue(rowValues[index], 120));
          })
        : [];

    if (safeHeaders.length === 0) {
        return '';
    }

    const maxWidths = Array.isArray(options.columnMaxWidths) ? options.columnMaxWidths : [];
    const widths = safeHeaders.map((header, index) => {
        const maxRowWidth = safeRows.reduce((maxWidth, row) => Math.max(maxWidth, String(row[index] || '-').length), 0);
        const naturalWidth = Math.max(header.length, maxRowWidth, 3);
        const maxWidth = Number(maxWidths[index] || 0);
        return maxWidth > 0 ? Math.min(naturalWidth, maxWidth) : naturalWidth;
    });
    const separator = widths.map((width) => '-'.repeat(width)).join('-+-');
    const renderRow = (columns) =>
        columns
            .map((column, index) => {
                const text = normalizeEmailTableValue(column, widths[index]);
                return text.padEnd(widths[index], ' ');
            })
            .join(' | ');

    const lines = [renderRow(safeHeaders), separator];
    safeRows.forEach((row) => lines.push(renderRow(row)));
    return lines.join('\n');
}

function buildInvoiceDetailTableRows(item) {
    return [
        ['NRO INGRESO', normalizeEmailTableValue(item.numIngreso, 20)],
        ['NOMBRE / RAZON SOCIAL', normalizeEmailTableValue(item.nombreRazonSocial, 80)],
        ['RUT', normalizeEmailTableValue(item.rut, 20)],
        ['GIRO', normalizeEmailTableValue(item.giro, 80)],
        ['DIRECCION', normalizeEmailTableValue(item.direccion, 90)],
        ['COMUNA / CIUDAD', normalizeEmailTableValue(buildInvoiceComunaRegionLabel(item.comuna, item.ciudad), 80)],
        ['CONTACTO', normalizeEmailTableValue(item.contacto, 80)],
        ['MONTO A FACTURAR', normalizeEmailTableValue(formatInvoiceAmount(item.montoFacturar), 30)],
        ['OBSERVACION', normalizeEmailTableValue(item.observacion, 120)]
    ];
}

function buildSingleInvoiceEmailBody(item) {
    const nowText = new Date().toLocaleString('es-CL', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    });

    const operatorName = currentUser?.nombre || currentUser?.username || 'Operador';
    const branch = currentUser?.sucursal || 'Sin sucursal';
    const lines = [
        'Solicitud de facturacion',
        `Fecha: ${nowText}`,
        `Operador: ${operatorName}`,
        `Sucursal: ${branch}`,
        '',
        'DETALLE DE FACTURA',
        buildReadableTextTable(['CAMPO', 'VALOR'], buildInvoiceDetailTableRows(item), {
            columnMaxWidths: [28, 110]
        })
    ];
    return lines.join('\n');
}

function buildPendingInvoicesEmailBody(items = operatorPendingInvoices) {
    const normalizedItems = Array.isArray(items) ? items.filter((item) => Boolean(item)) : [];
    const nowText = new Date().toLocaleString('es-CL', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    });
    const operatorName = currentUser?.nombre || currentUser?.username || 'Operador';
    const branch = currentUser?.sucursal || 'Sin sucursal';

    const lines = [
        'Resumen de facturas pendientes',
        `Fecha: ${nowText}`,
        `Operador: ${operatorName}`,
        `Sucursal: ${branch}`,
        `Total pendientes: ${normalizedItems.length}`,
        '',
        'RESUMEN',
        buildReadableTextTable(
            ['#', 'NRO INGRESO', 'CLIENTE / RAZON SOCIAL', 'MONTO', 'CONTACTO'],
            normalizedItems.map((item, index) => [
                String(index + 1),
                normalizeEmailTableValue(item.numIngreso, 18),
                normalizeEmailTableValue(item.nombreRazonSocial, 42),
                normalizeEmailTableValue(formatInvoiceAmount(item.montoFacturar), 18),
                normalizeEmailTableValue(item.contacto, 28)
            ]),
            {
                columnMaxWidths: [3, 18, 42, 18, 28]
            }
        ),
        '',
        'DETALLE COMPLETO PARA FACTURACION'
    ];

    normalizedItems.forEach((item, index) => {
        lines.push('');
        lines.push(`FACTURA ${index + 1}`);
        lines.push(
            buildReadableTextTable(['CAMPO', 'VALOR'], buildInvoiceDetailTableRows(item), {
                columnMaxWidths: [28, 110]
            })
        );
    });

    return lines.join('\n');
}

async function registerInvoiceRequestInDatabase(invoiceData, options = {}) {
    const status = normalizeInvoiceRequestStatus(options.estado);
    const destinationEmail = normalizeInvoiceText(options.destinationEmail, 255).toLowerCase();

    if (!status) {
        throw new Error('Estado de solicitud invalido para registrar.');
    }

    const payload = {
        ...invoiceData,
        estado: status,
        destinationEmail
    };

    return apiRequest('/facturas/solicitudes', {
        method: 'POST',
        body: JSON.stringify(payload)
    });
}

async function sendInvoiceEmailViaSmtp(options = {}) {
    const destinationEmail = normalizeInvoiceText(options.destinationEmail, 255).toLowerCase();
    const subject = buildMailSubject(options.subject);
    const textBody = String(options.textBody || '').trim();
    const htmlBody = String(options.htmlBody || '').trim();

    if (!destinationEmail) {
        throw new Error('Debes ingresar el correo del contador.');
    }
    if (!subject) {
        throw new Error('Debes indicar el asunto del correo.');
    }
    if (!textBody && !htmlBody) {
        throw new Error('Debes indicar el contenido del correo.');
    }

    return apiRequest('/facturas/correo/enviar', {
        method: 'POST',
        body: JSON.stringify({
            destinationEmail,
            subject,
            textBody,
            htmlBody
        })
    });
}

async function markPendingInvoiceRequestsAsSent(destinationEmail, numIngresos = []) {
    const normalizedNumIngresos = Array.isArray(numIngresos)
        ? Array.from(
              new Set(
                  numIngresos
                      .map((item) => normalizeInvoiceText(item, 20))
                      .filter((item) => Boolean(item) && /^\d+-\d{4}$/.test(item))
              )
          )
        : [];

    const result = await apiRequest('/facturas/solicitudes/marcar-enviadas', {
        method: 'POST',
        body: JSON.stringify({
            destinationEmail: normalizeInvoiceText(destinationEmail, 255).toLowerCase(),
            numIngresos: normalizedNumIngresos
        })
    });

    return Number(result.updatedCount || 0);
}

async function handleSendCurrentInvoiceToAccountant() {
    if (!canUseInvoiceWorkflow(currentUser)) {
        setFormMessage(ui.facturaFormMessage, 'Tu rol no tiene permisos para gestionar facturas.', 'error');
        return;
    }

    clearFacturaFieldErrors();
    const parsed = validateInvoiceDraft(collectFacturaFormData(), { requireAmount: true });
    if (parsed.error) {
        const appliedInline = applyFacturaInlineErrorMessage(parsed.error, parsed.errorFieldId || '');
        setFormMessage(ui.facturaFormMessage, appliedInline ? '' : parsed.error, appliedInline ? '' : 'error');
        return;
    }

    const destinationEmail = await promptAccountantEmail();
    if (destinationEmail === null) {
        return;
    }
    if (destinationEmail === '') {
        setFormMessage(ui.facturaFormMessage, 'Debes ingresar el correo del contador.', 'error');
        return;
    }
    if (destinationEmail === 'INVALID_EMAIL') {
        setFormMessage(ui.facturaFormMessage, 'El correo del contador no es valido.', 'error');
        return;
    }

    const targetIngreso = normalizeInvoiceText(parsed?.data?.numIngreso, 20);
    if (!targetIngreso) {
        setFormMessage(ui.facturaFormMessage, 'Debes cargar un NRO INGRESO valido antes de enviar al contador.', 'error');
        return;
    }
    let registerResult = null;
    let shouldSyncFacturaSnapshot = false;
    const sendButtons = [ui?.operatorEmitInvoiceBtn, ui?.facturaEnviarContadorBtn].filter((button) => Boolean(button));
    sendButtons.forEach((button) => {
        button.disabled = true;
    });
    setFormMessage(ui.facturaFormMessage, 'Enviando factura por correo...', 'success');
    showMailSendProgress('Enviando factura por correo...');

    try {
        registerResult = await registerInvoiceRequestInDatabase(parsed.data, {
            estado: 'PENDIENTE'
        });
        shouldSyncFacturaSnapshot = true;
        const subject = buildMailSubject(`Solicitud factura ${parsed.data.numIngreso || ''} ${parsed.data.nombreRazonSocial}`.trim());
        await sendInvoiceEmailViaSmtp({
            destinationEmail,
            subject,
            textBody: buildSingleInvoiceEmailBody(parsed.data),
            htmlBody: buildSingleInvoiceEmailHtml(parsed.data)
        });
        await markPendingInvoiceRequestsAsSent(destinationEmail, [targetIngreso]);

        setCurrentInvoiceRequestStatus('ENVIADA');
        await refreshOperatorPendingInvoices();
        if (canUseSecretaryFeatures(currentUser)) {
            await refreshRequestedInvoicesHistory();
        }
        if (syncLoadedFacturaSnapshotWithCurrentForm()) {
            updateRegistroActionButtons();
        }
        await hideMailSendProgress('Correo enviado correctamente.', 'success');

        setFormMessage(
            ui.facturaFormMessage,
            registerResult?.message || `Factura enviada correctamente a ${destinationEmail}.`,
            'success'
        );
        setFormMessage(ui.operatorInvoiceMessage, `Factura enviada al contador (${destinationEmail}).`, 'success');
    } catch (error) {
        if (shouldSyncFacturaSnapshot && syncLoadedFacturaSnapshotWithCurrentForm()) {
            updateRegistroActionButtons();
        }
        await hideMailSendProgress('No fue posible enviar el correo.', 'error', {
            minDurationMs: MAIL_SEND_PROGRESS_MIN_MS,
            finalDurationMs: 1200
        });
        const appliedInline = applyFacturaInlineErrorMessage(error.message);
        setFormMessage(ui.facturaFormMessage, appliedInline ? '' : error.message, appliedInline ? '' : 'error');
    } finally {
        applyInvoiceSendButtonsState();
    }
}

async function handleAddCurrentInvoiceToPending() {
    if (!canUseInvoiceWorkflow(currentUser)) {
        setFormMessage(ui.facturaFormMessage, 'Tu rol no tiene permisos para gestionar facturas.', 'error');
        return;
    }

    if (normalizeInvoiceRequestStatus(currentInvoiceRequestStatus) === 'ENVIADA') {
        setFormMessage(ui.facturaFormMessage, 'Esta factura ya fue enviada al contador y no puede agregarse a pendientes.', 'error');
        setFormMessage(ui.operatorInvoiceMessage, 'La factura ya fue enviada al contador.', 'error');
        applyInvoiceSendButtonsState();
        return;
    }

    clearFacturaFieldErrors();
    const parsed = validateInvoiceDraft(collectFacturaFormData(), { requireAmount: true });
    if (parsed.error) {
        const appliedInline = applyFacturaInlineErrorMessage(parsed.error, parsed.errorFieldId || '');
        setFormMessage(ui.facturaFormMessage, appliedInline ? '' : parsed.error, appliedInline ? '' : 'error');
        return;
    }

    const targetIngreso = normalizeInvoiceText(parsed?.data?.numIngreso, 20);
    if (targetIngreso && /^\d+-\d{4}$/.test(targetIngreso)) {
        await syncCurrentInvoiceRequestStatusByIngreso(targetIngreso);
        if (normalizeInvoiceRequestStatus(currentInvoiceRequestStatus) === 'ENVIADA') {
            setFormMessage(
                ui.facturaFormMessage,
                'Esta factura ya fue enviada al contador y no puede agregarse a pendientes.',
                'error'
            );
            setFormMessage(ui.operatorInvoiceMessage, 'La factura ya fue enviada al contador.', 'error');
            applyInvoiceSendButtonsState();
            return;
        }
    }

    let registerResult = null;
    try {
        registerResult = await registerInvoiceRequestInDatabase(parsed.data, {
            estado: 'PENDIENTE'
        });
        setCurrentInvoiceRequestStatus(registerResult?.estado || 'PENDIENTE');
        await refreshOperatorPendingInvoices({
            showErrors: true,
            messageElement: ui.operatorInvoiceMessage
        });
        if (canUseSecretaryFeatures(currentUser)) {
            await refreshRequestedInvoicesHistory();
        }
    } catch (error) {
        const appliedInline = applyFacturaInlineErrorMessage(error.message);
        setFormMessage(ui.facturaFormMessage, appliedInline ? '' : error.message, appliedInline ? '' : 'error');
        return;
    }

    if (syncLoadedFacturaSnapshotWithCurrentForm()) {
        updateRegistroActionButtons();
    }

    const finalStatus = normalizeInvoiceRequestStatus(registerResult?.estado) || 'PENDIENTE';
    if (finalStatus === 'ENVIADA') {
        setFormMessage(
            ui.facturaFormMessage,
            registerResult?.message || 'La factura ya estaba enviada. No se cambio a pendiente.',
            'error'
        );
        setFormMessage(ui.operatorInvoiceMessage, 'La factura ya estaba enviada al contador.', 'error');
        return;
    }

    setFormMessage(
        ui.facturaFormMessage,
        registerResult?.message || `Factura de ${parsed.data.nombreRazonSocial} agregada a la lista de pendientes.`,
        'success'
    );
    setFormMessage(
        ui.operatorInvoiceMessage,
        `Pendiente agregado: ${parsed.data.nombreRazonSocial} (${formatInvoiceAmount(parsed.data.montoFacturar)}).`,
        'success'
    );
}

async function handleSendPendingInvoicesByEmail() {
    if (!canUseInvoiceWorkflow(currentUser)) {
        setFormMessage(ui.operatorInvoiceMessage, 'Tu rol no tiene permisos para gestionar facturas.', 'error');
        return;
    }

    const refreshed = await refreshOperatorPendingInvoices({
        showErrors: true,
        messageElement: ui.operatorInvoiceMessage
    });
    if (!refreshed) {
        return;
    }
    if (!Array.isArray(operatorPendingInvoices) || operatorPendingInvoices.length === 0) {
        setFormMessage(ui.operatorInvoiceMessage, 'No hay facturas pendientes para enviar.', 'error');
        return;
    }

    const destinationEmail = await promptAccountantEmail();
    if (destinationEmail === null) {
        return;
    }
    if (destinationEmail === '') {
        setFormMessage(ui.operatorInvoiceMessage, 'Debes ingresar el correo del contador.', 'error');
        return;
    }
    if (destinationEmail === 'INVALID_EMAIL') {
        setFormMessage(ui.operatorInvoiceMessage, 'El correo del contador no es valido.', 'error');
        return;
    }

    const pendingSnapshot = Array.isArray(operatorPendingInvoices) ? [...operatorPendingInvoices] : [];
    const pendingIngresos = pendingSnapshot
        .map((item) => normalizeInvoiceText(item?.numIngreso, 20))
        .filter((item) => Boolean(item) && /^\d+-\d{4}$/.test(item));
    if (pendingIngresos.length === 0) {
        setFormMessage(ui.operatorInvoiceMessage, 'No se encontraron NRO INGRESO validos para enviar.', 'error');
        return;
    }
    let updatedCount = 0;
    showMailSendProgress('Enviando lista de facturas pendientes...');
    try {
        const subject = buildMailSubject(
            `Facturas pendientes turno ${new Date().toLocaleDateString('es-CL')} (${pendingSnapshot.length})`
        );
        await sendInvoiceEmailViaSmtp({
            destinationEmail,
            subject,
            textBody: buildPendingInvoicesEmailBody(pendingSnapshot),
            htmlBody: buildPendingInvoicesEmailHtml(pendingSnapshot)
        });
        updatedCount = await markPendingInvoiceRequestsAsSent(destinationEmail, pendingIngresos);
        await hideMailSendProgress('Correo enviado correctamente.', 'success');
    } catch (error) {
        await hideMailSendProgress('No fue posible enviar el correo.', 'error', {
            minDurationMs: MAIL_SEND_PROGRESS_MIN_MS,
            finalDurationMs: 1200
        });
        setFormMessage(ui.operatorInvoiceMessage, error.message, 'error');
        return;
    }

    await refreshOperatorPendingInvoices();
    if (canUseSecretaryFeatures(currentUser)) {
        await refreshRequestedInvoicesHistory();
    }
    await syncCurrentInvoiceRequestStatusByIngreso(ui?.numIngreso?.value);

    if (updatedCount <= 0) {
        setFormMessage(ui.operatorInvoiceMessage, 'No habia solicitudes pendientes para marcar como enviadas.', 'error');
        return;
    }

    setFormMessage(
        ui.operatorInvoiceMessage,
        `Se marcaron ${updatedCount} solicitudes pendientes como enviadas a ${destinationEmail}.`,
        'success'
    );
    setFormMessage(ui.facturaFormMessage, `Lista pendiente enviada a ${destinationEmail}.`, 'success');
}

function setRegistroMainVisibility(isVisible) {
    if (!ui || !ui.registroForm) {
        return;
    }

    if (ui.registroMainSection) {
        ui.registroMainSection.classList.toggle('hidden', !isVisible);
        return;
    }

    // Fallback para estructuras antiguas sin #registroMainSection.
    const children = Array.from(ui.registroForm.children);
    children.forEach((child) => {
        const childId = child.id || '';
        const keepVisible = childId === 'adminPanel' || child.classList.contains('session-bar');
        if (keepVisible) {
            return;
        }

        child.classList.toggle('hidden', !isVisible);
    });
}

function formatMonthlyUtmPeriodLabel(status) {
    const source = status && typeof status === 'object' ? status : {};
    const monthRaw = Number(source.mes ?? source?.utm?.mes);
    const yearRaw = Number(source.anio ?? source?.utm?.anio);
    const safeMonth = Number.isInteger(monthRaw) && monthRaw >= 1 && monthRaw <= 12 ? monthRaw : new Date().getMonth() + 1;
    const safeYear = Number.isInteger(yearRaw) && yearRaw >= 2000 ? yearRaw : new Date().getFullYear();
    const monthLabel = MONTH_NAMES_ES[safeMonth - 1] || 'mes actual';
    return `${monthLabel.toUpperCase()} ${safeYear}`;
}

function normalizeMonthlyUtmInput(value) {
    const parsed = parseInvoiceAmountNumber(value);
    if (!Number.isFinite(parsed) || parsed <= 0) {
        return '';
    }

    return parsed.toFixed(2);
}

function showMonthlyUtmCaptureCard(user, status, message = '', options = {}) {
    const manualEdit = Boolean(options?.manualEdit);
    pendingMonthlyUtmStatus = status || null;
    if (ui.appTopPanels) {
        ui.appTopPanels.classList.add('hidden');
    }
    closeInvoiceModal();
    document.body.classList.remove('app-authenticated');
    ui.authCard.classList.add('hidden');
    ui.registroForm.classList.add('hidden');
    ui.passwordChangeCard.classList.add('hidden');
    if (ui.utmCaptureCard) {
        ui.utmCaptureCard.classList.remove('hidden');
    }

    const userName = user?.nombre || user?.username || 'secretaria';
    if (ui.utmCaptureIntro) {
        if (manualEdit) {
            ui.utmCaptureIntro.textContent = `Usuario ${userName}: puedes editar y confirmar la UTM del periodo cuando lo necesites.`;
        } else {
            ui.utmCaptureIntro.textContent = `Usuario ${userName}: debes registrar y confirmar la UTM del periodo para continuar.`;
        }
    }
    if (ui.utmCapturePeriod) {
        ui.utmCapturePeriod.textContent = `Periodo: ${formatMonthlyUtmPeriodLabel(status)}`;
    }

    if (ui.utmCaptureForm) {
        ui.utmCaptureForm.reset();
    }
    const suggestedValue = normalizeMonthlyUtmInput(status?.suggestedUtm?.valor);
    const existingValue = normalizeMonthlyUtmInput(status?.utm?.valor);
    const prefilledValue = suggestedValue || existingValue || '';
    if (ui.utmValue) {
        ui.utmValue.value = prefilledValue ? formatInvoiceAmountForInput(prefilledValue) : '';
    }
    if (ui.utmConfirmValue) {
        ui.utmConfirmValue.value = manualEdit && prefilledValue ? formatInvoiceAmountForInput(prefilledValue) : '';
    }
    if (ui.utmCaptureSource) {
        if (suggestedValue) {
            const sourceLabel = String(status?.suggestedUtm?.fuente || 'mindicador.cl').trim();
            ui.utmCaptureSource.textContent = `Valor sugerido: ${formatInvoiceAmount(suggestedValue)} (${sourceLabel}). Puedes editar ambos campos.`;
        } else if (existingValue) {
            ui.utmCaptureSource.textContent = `Valor ya registrado: ${formatInvoiceAmount(existingValue)}. Puedes actualizar y confirmar.`;
        } else {
            ui.utmCaptureSource.textContent = 'No se pudo cargar sugerencia automatica. Ingresa el valor manualmente.';
        }
    }
    setFormMessage(ui.loginMessage, '', '');
    setFormMessage(ui.formMessage, '', '');
    setFormMessage(ui.passwordChangeMessage, '', '');
    setFormMessage(ui.utmCaptureMessage, message, message ? 'error' : '');

    if (ui.utmValue) {
        requestAnimationFrame(() => {
            ui.utmValue.focus();
        });
    }
}

async function handleUtmCaptureSubmit(event) {
    event.preventDefault();

    if (!currentUser || !canUseSecretaryFeatures(currentUser)) {
        showAuthCard('Debes iniciar sesion para registrar la UTM mensual.');
        return;
    }

    const value = normalizeMonthlyUtmInput(ui.utmValue?.value);
    const confirmation = normalizeMonthlyUtmInput(ui.utmConfirmValue?.value);
    if (!value || !confirmation) {
        setFormMessage(ui.utmCaptureMessage, 'Ingresa y confirma un valor UTM valido mayor a 0.', 'error');
        return;
    }
    if (value !== confirmation) {
        setFormMessage(ui.utmCaptureMessage, 'El valor UTM y su confirmacion no coinciden.', 'error');
        return;
    }

    try {
        const result = await apiRequest('/utm/mes-actual', {
            method: 'POST',
            body: JSON.stringify({
                valor: value,
                confirmacion: confirmation
            }),
            skipAuthRedirect: true,
            skipPasswordRedirect: true
        });

        pendingMonthlyUtmStatus = result?.status || pendingMonthlyUtmStatus;
        setFormMessage(ui.utmCaptureMessage, result.message || 'UTM mensual guardada correctamente.', 'success');
        await activateAuthenticatedUser(currentUser, { passwordMessage: '' });
    } catch (error) {
        setFormMessage(ui.utmCaptureMessage, error.message, 'error');
    }
}

function showPasswordChangeCard(user, message = '') {
    if (ui.appTopPanels) {
        ui.appTopPanels.classList.add('hidden');
    }
    closeInvoiceModal();
    document.body.classList.remove('app-authenticated');
    ui.authCard.classList.add('hidden');
    ui.registroForm.classList.add('hidden');
    ui.passwordChangeCard.classList.remove('hidden');
    if (ui.utmCaptureCard) {
        ui.utmCaptureCard.classList.add('hidden');
    }
    setFormMessage(ui.loginMessage, '', '');
    setFormMessage(ui.formMessage, '', '');
    setFormMessage(ui.passwordChangeMessage, message, message ? 'error' : '');
    setFormMessage(ui.utmCaptureMessage, '', '');
    if (ui.utmCaptureSource) {
        ui.utmCaptureSource.textContent = '';
    }

    if (ui.passwordChangeIntro) {
        const userName = user?.nombre || user?.username || 'usuario';
        ui.passwordChangeIntro.textContent = `Usuario ${userName}: debes reemplazar tu clave temporal para continuar.`;
    }
}

async function handlePasswordChangeSubmit(event) {
    event.preventDefault();

    const currentPassword = ui.currentPassword.value;
    const newPassword = ui.newPassword.value;
    const confirmPassword = ui.confirmPassword.value;

    if (!currentPassword || !newPassword || !confirmPassword) {
        setFormMessage(ui.passwordChangeMessage, 'Completa todos los campos de clave.', 'error');
        return;
    }

    if (newPassword.length < 6) {
        setFormMessage(ui.passwordChangeMessage, 'La nueva clave debe tener al menos 6 caracteres.', 'error');
        return;
    }

    if (newPassword !== confirmPassword) {
        setFormMessage(ui.passwordChangeMessage, 'La confirmacion no coincide con la nueva clave.', 'error');
        return;
    }

    try {
        const result = await apiRequest('/auth/change-password', {
            method: 'POST',
            body: JSON.stringify({ currentPassword, newPassword }),
            skipAuthRedirect: true,
            skipPasswordRedirect: true
        });

        ui.passwordChangeForm.reset();
        setFormMessage(ui.passwordChangeMessage, result.message || 'Clave actualizada correctamente.', 'success');
        await activateAuthenticatedUser(result.user || currentUser, { passwordMessage: '' });
    } catch (error) {
        setFormMessage(ui.passwordChangeMessage, error.message, 'error');
    }
}

function configureIngresoFieldByRole(user) {
    if (!canUseHistoricalIngresoMode(user)) {
        ui.numIngreso.readOnly = true;
        ui.numIngreso.classList.remove('admin-editable');
        ui.numIngresoHint.textContent = 'Autogenerado.';
        setAdminModeButtonsActive('current');
        setAdminYearMessage('');
        return;
    }

    void setAdminIngresoMode(adminIngresoMode, { refreshIngreso: false, clearIngreso: false });
}

function setAdminManagementView(view) {
    if (!ui || !canAccessAdminConsole(currentUser)) {
        return;
    }

    let normalizedView =
        view === 'users' ||
        view === 'branches' ||
        view === 'documents' ||
        view === 'email' ||
        view === 'mailTemplates' ||
        view === 'usage' ||
        view === 'maintenance'
            ? view
            : 'registro';
    if (!canManageAdminCatalogs(currentUser)) {
        normalizedView = 'registro';
    }
    if (normalizedView === 'email' && !isSuperUser(currentUser)) {
        normalizedView = 'registro';
    }
    if (normalizedView === 'mailTemplates' && !isSuperUser(currentUser)) {
        normalizedView = 'registro';
    }
    if (normalizedView === 'usage' && !isSuperUser(currentUser)) {
        normalizedView = 'registro';
    }
    if (normalizedView === 'maintenance' && !isSuperUser(currentUser)) {
        normalizedView = 'registro';
    }
    if (ui.adminUsersSection) {
        ui.adminUsersSection.classList.toggle('hidden', normalizedView !== 'users');
    }
    if (ui.adminBranchesSection) {
        ui.adminBranchesSection.classList.toggle('hidden', normalizedView !== 'branches');
    }
    if (ui.adminDocumentsSection) {
        ui.adminDocumentsSection.classList.toggle('hidden', normalizedView !== 'documents');
    }
    if (ui.adminEmailSection) {
        ui.adminEmailSection.classList.toggle('hidden', normalizedView !== 'email');
    }
    if (ui.adminMailTemplatesSection) {
        ui.adminMailTemplatesSection.classList.toggle('hidden', normalizedView !== 'mailTemplates');
    }
    if (ui.adminUsageSection) {
        ui.adminUsageSection.classList.toggle('hidden', normalizedView !== 'usage');
    }
    if (ui.adminMaintenanceSection) {
        ui.adminMaintenanceSection.classList.toggle('hidden', normalizedView !== 'maintenance');
    }
    if (normalizedView !== 'documents') {
        exitDocumentEditMode();
        setFormMessage(ui.adminDocumentMessage, '', '');
    }
    if (normalizedView !== 'email') {
        setFormMessage(ui.adminEmailMessage, '', '');
    }
    if (normalizedView !== 'mailTemplates') {
        setFormMessage(ui.adminMailTemplateMessage, '', '');
    }
    if (normalizedView !== 'usage') {
        setFormMessage(ui.adminUsageMessage, '', '');
    }
    if (normalizedView !== 'maintenance') {
        setFormMessage(ui.adminMaintenanceMessage, '', '');
        setFormMessage(ui.adminLoginNoticeMessage, '', '');
    }

    appViewMode = normalizedView === 'registro' ? 'registro' : `admin-${normalizedView}`;
    setRegistroMainVisibility(normalizedView === 'registro');
}

function setAdminUsersVisibility(isVisible) {
    if (!ui || !ui.adminUsersSection || !ui.adminToggleUsersBtn) {
        return;
    }

    ui.adminToggleUsersBtn.textContent = 'GESTIONAR USUARIOS';
    setAdminManagementView(isVisible ? 'users' : 'registro');
}

function setAdminBranchesVisibility(isVisible) {
    if (!ui || !ui.adminBranchesSection || !ui.adminToggleBranchesBtn) {
        return;
    }

    ui.adminToggleBranchesBtn.textContent = 'GESTIONAR SUCURSALES';
    setAdminManagementView(isVisible ? 'branches' : 'registro');
}

function setAdminDocumentsVisibility(isVisible) {
    if (!ui || !ui.adminDocumentsSection || !ui.adminToggleDocumentsBtn) {
        return;
    }

    ui.adminToggleDocumentsBtn.textContent = 'GESTIONAR DOCUMENTOS';
    setAdminManagementView(isVisible ? 'documents' : 'registro');
}

function setAdminEmailVisibility(isVisible) {
    if (!ui || !ui.adminEmailSection || !ui.adminToggleEmailBtn) {
        return;
    }

    ui.adminToggleEmailBtn.textContent = 'CONFIGURAR CORREO';
    setAdminManagementView(isVisible ? 'email' : 'registro');
}

function setAdminMailTemplatesVisibility(isVisible) {
    if (!ui || !ui.adminMailTemplatesSection || !ui.adminToggleMailTemplatesBtn) {
        return;
    }

    ui.adminToggleMailTemplatesBtn.textContent = 'PLANTILLAS CORREO';
    setAdminManagementView(isVisible ? 'mailTemplates' : 'registro');
}

function setAdminUsageVisibility(isVisible) {
    if (!ui || !ui.adminUsageSection || !ui.adminToggleUsageBtn) {
        return;
    }

    ui.adminToggleUsageBtn.textContent = 'EQUIPOS ACTIVOS';
    setAdminManagementView(isVisible ? 'usage' : 'registro');
}

function setAdminMaintenanceVisibility(isVisible) {
    if (!ui || !ui.adminMaintenanceSection || !ui.adminToggleMaintenanceBtn) {
        return;
    }

    ui.adminToggleMaintenanceBtn.textContent = 'MODO MANTENCION';
    setAdminManagementView(isVisible ? 'maintenance' : 'registro');
}

function handleAdminToggleUsers() {
    if (!canManageAdminCatalogs(currentUser)) {
        return;
    }

    setAdminUsersVisibility(true);
    void loadAdminUsers();
    void loadBranches();
}

function handleAdminToggleBranches() {
    if (!canManageAdminCatalogs(currentUser)) {
        return;
    }

    setAdminBranchesVisibility(true);
    void loadBranches();
}

function handleAdminToggleDocuments() {
    if (!canManageAdminCatalogs(currentUser)) {
        return;
    }

    setAdminDocumentsVisibility(true);
    renderAdminDocumentGroupTitleInputs(configuredDocumentGroups);
    renderAdminDocumentGroupSelect(configuredDocumentGroups);
    renderAdminDocuments(configuredDocumentOptions);
}

async function handleAdminToggleEmail() {
    if (!isSuperUser(currentUser)) {
        return;
    }

    setAdminEmailVisibility(true);
    await loadAdminEmailConfig();
}

async function handleAdminToggleMailTemplates() {
    if (!isSuperUser(currentUser)) {
        return;
    }

    setAdminMailTemplatesVisibility(true);
    await loadAdminMailTemplates();
}

async function handleAdminToggleUsage() {
    if (!isSuperUser(currentUser)) {
        return;
    }

    setAdminUsageVisibility(true);
    await loadAdminUsageStats();
}

async function handleAdminToggleMaintenance() {
    if (!isSuperUser(currentUser)) {
        return;
    }

    setAdminMaintenanceVisibility(true);
    await loadAdminMaintenanceStatus();
}

function clearAdminMaintenanceState() {
    adminMaintenanceStatusLoaded = null;
    adminMaintenanceUpdateInProgress = false;
    adminLoginRouteNoticeStatusLoaded = null;
    adminLoginRouteNoticeUpdateInProgress = false;
    if (!ui) {
        return;
    }

    if (ui.adminMaintenanceState) {
        ui.adminMaintenanceState.textContent = '-';
    }
    if (ui.adminMaintenanceUpdatedAt) {
        ui.adminMaintenanceUpdatedAt.textContent = '-';
    }
    if (ui.adminMaintenanceUpdatedBy) {
        ui.adminMaintenanceUpdatedBy.textContent = '-';
    }
    if (ui.adminLoginNoticeState) {
        ui.adminLoginNoticeState.textContent = '-';
    }
    if (ui.adminLoginNoticeUrl) {
        ui.adminLoginNoticeUrl.value = getDefaultLoginNoticeUrlByRuntime();
    }
    if (ui.adminLoginNoticeExpiresAt) {
        ui.adminLoginNoticeExpiresAt.textContent = '-';
    }
    if (ui.adminLoginNoticeUpdatedBy) {
        ui.adminLoginNoticeUpdatedBy.textContent = '-';
    }
    setAdminLoginRouteNoticeButtonsBusy(false);
    setAdminMaintenanceButtonsBusy(false);
}

function applyAdminMaintenanceStatus(status) {
    adminMaintenanceStatusLoaded = status && typeof status === 'object' ? status : null;
    if (!ui) {
        return;
    }

    const enabled = Boolean(adminMaintenanceStatusLoaded?.enabled);
    if (ui.adminMaintenanceState) {
        ui.adminMaintenanceState.textContent = enabled ? 'ACTIVO' : 'INACTIVO';
    }
    if (ui.adminMaintenanceUpdatedAt) {
        ui.adminMaintenanceUpdatedAt.textContent = adminMaintenanceStatusLoaded?.updatedAt
            ? formatDateTime(adminMaintenanceStatusLoaded.updatedAt)
            : '-';
    }
    if (ui.adminMaintenanceUpdatedBy) {
        ui.adminMaintenanceUpdatedBy.textContent = String(adminMaintenanceStatusLoaded?.updatedBy || '-').trim() || '-';
    }
    setAdminMaintenanceButtonsBusy(false);
}

function normalizeAdminLoginNoticeUrl(value) {
    const raw = String(value || '').trim();
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

function getDefaultLoginNoticeUrlByRuntime() {
    const host = String(window.location.hostname || '')
        .trim()
        .toLowerCase();
    if (host === 'localhost' || host === '127.0.0.1') {
        return 'http://localhost/geo_rural/registro/login/registro.html';
    }
    return 'https://mi-registro.cl/geo_rural/registro/login/registro.html';
}

function applyAdminLoginRouteNoticeStatus(status) {
    adminLoginRouteNoticeStatusLoaded = status && typeof status === 'object' ? status : null;
    if (!ui) {
        return;
    }

    const isEnabled = Boolean(adminLoginRouteNoticeStatusLoaded?.enabled);
    const isActive = Boolean(adminLoginRouteNoticeStatusLoaded?.active);
    const expiresAt = adminLoginRouteNoticeStatusLoaded?.expiresAt;
    const stateText = isActive ? 'ACTIVO' : isEnabled ? 'EXPIRADO' : 'INACTIVO';

    if (ui.adminLoginNoticeState) {
        ui.adminLoginNoticeState.textContent = stateText;
    }
    if (ui.adminLoginNoticeUrl) {
        const safeUrl = normalizeAdminLoginNoticeUrl(adminLoginRouteNoticeStatusLoaded?.url) || getDefaultLoginNoticeUrlByRuntime();
        ui.adminLoginNoticeUrl.value = safeUrl;
    }
    if (ui.adminLoginNoticeExpiresAt) {
        ui.adminLoginNoticeExpiresAt.textContent = expiresAt ? formatDateTime(expiresAt) : '-';
    }
    if (ui.adminLoginNoticeUpdatedBy) {
        ui.adminLoginNoticeUpdatedBy.textContent = String(adminLoginRouteNoticeStatusLoaded?.updatedBy || '-').trim() || '-';
    }
    setAdminLoginRouteNoticeButtonsBusy(false);
}

function setAdminMaintenanceButtonsBusy(isBusy) {
    if (!ui) {
        return;
    }

    const isEnabled = Boolean(adminMaintenanceStatusLoaded?.enabled);
    const disabled = Boolean(isBusy) || adminMaintenanceUpdateInProgress;

    if (ui.adminMaintenanceRefreshBtn) {
        ui.adminMaintenanceRefreshBtn.disabled = disabled;
    }
    if (ui.adminMaintenanceEnableBtn) {
        ui.adminMaintenanceEnableBtn.disabled = disabled || isEnabled;
    }
    if (ui.adminMaintenanceDisableBtn) {
        ui.adminMaintenanceDisableBtn.disabled = disabled || !isEnabled;
    }
}

function setAdminLoginRouteNoticeButtonsBusy(isBusy) {
    if (!ui) {
        return;
    }

    const isEnabled = Boolean(adminLoginRouteNoticeStatusLoaded?.enabled);
    const isActive = Boolean(adminLoginRouteNoticeStatusLoaded?.active);
    const disabled = Boolean(isBusy) || adminLoginRouteNoticeUpdateInProgress;

    if (ui.adminLoginNoticeUrl) {
        ui.adminLoginNoticeUrl.disabled = disabled;
    }
    if (ui.adminLoginNoticeRefreshBtn) {
        ui.adminLoginNoticeRefreshBtn.disabled = disabled;
    }
    if (ui.adminLoginNoticeEnableBtn) {
        ui.adminLoginNoticeEnableBtn.disabled = disabled || isActive;
    }
    if (ui.adminLoginNoticeDisableBtn) {
        ui.adminLoginNoticeDisableBtn.disabled = disabled || !isEnabled;
    }
}

async function loadAdminMaintenanceStatus(options = {}) {
    if (!isSuperUser(currentUser)) {
        return;
    }

    const silent = options.silent === true;
    setAdminMaintenanceButtonsBusy(true);
    setAdminLoginRouteNoticeButtonsBusy(true);
    if (!silent) {
        setFormMessage(ui.adminMaintenanceMessage, 'Consultando estado de mantencion...', '');
    }

    try {
        const data = await apiRequest('/system/mantenimiento');
        applyAdminMaintenanceStatus(data?.maintenance || null);
        applyAdminLoginRouteNoticeStatus(data?.loginRouteNotice || null);
        if (!silent) {
            setFormMessage(ui.adminMaintenanceMessage, 'Estado actualizado.', 'success');
        }
    } catch (error) {
        if (!silent) {
            setFormMessage(ui.adminMaintenanceMessage, error.message || 'No fue posible consultar el estado.', 'error');
        }
        setAdminMaintenanceButtonsBusy(false);
        setAdminLoginRouteNoticeButtonsBusy(false);
    }
}

async function updateAdminMaintenanceMode(enabled) {
    if (!isSuperUser(currentUser) || adminMaintenanceUpdateInProgress) {
        return;
    }

    adminMaintenanceUpdateInProgress = true;
    setAdminMaintenanceButtonsBusy(true);
    setFormMessage(
        ui.adminMaintenanceMessage,
        enabled ? 'Activando modo mantencion...' : 'Volviendo a modo normal...',
        ''
    );

    try {
        const data = await apiRequest('/system/mantenimiento', {
            method: 'PUT',
            body: JSON.stringify({ enabled: Boolean(enabled) })
        });
        applyAdminMaintenanceStatus(data?.maintenance || null);
        applyAdminLoginRouteNoticeStatus(data?.loginRouteNotice || null);
        setFormMessage(
            ui.adminMaintenanceMessage,
            data?.message || (enabled ? 'Modo mantencion activado.' : 'Modo mantencion desactivado.'),
            'success'
        );
    } catch (error) {
        setFormMessage(ui.adminMaintenanceMessage, error.message || 'No fue posible actualizar el estado.', 'error');
    } finally {
        adminMaintenanceUpdateInProgress = false;
        setAdminMaintenanceButtonsBusy(false);
    }
}

async function updateAdminLoginRouteNotice(enabled) {
    if (!isSuperUser(currentUser) || adminLoginRouteNoticeUpdateInProgress) {
        return;
    }

    const requestedUrl = normalizeAdminLoginNoticeUrl(ui?.adminLoginNoticeUrl?.value || '');
    if (enabled && !requestedUrl) {
        setFormMessage(ui.adminLoginNoticeMessage, 'Ingresa la URL del login para activar el aviso.', 'error');
        return;
    }

    adminLoginRouteNoticeUpdateInProgress = true;
    setAdminLoginRouteNoticeButtonsBusy(true);
    setFormMessage(
        ui.adminLoginNoticeMessage,
        enabled ? 'Activando aviso de nueva ruta por 15 dias...' : 'Desactivando aviso de nueva ruta...',
        ''
    );

    try {
        const data = await apiRequest('/system/login-route-notice', {
            method: 'PUT',
            body: JSON.stringify({
                enabled: Boolean(enabled),
                url: requestedUrl
            })
        });
        applyAdminMaintenanceStatus(data?.maintenance || null);
        applyAdminLoginRouteNoticeStatus(data?.loginRouteNotice || null);
        setFormMessage(
            ui.adminLoginNoticeMessage,
            data?.message || (enabled ? 'Aviso activado por 15 dias.' : 'Aviso desactivado.'),
            'success'
        );
    } catch (error) {
        setFormMessage(ui.adminLoginNoticeMessage, error.message || 'No fue posible actualizar el aviso.', 'error');
    } finally {
        adminLoginRouteNoticeUpdateInProgress = false;
        setAdminLoginRouteNoticeButtonsBusy(false);
    }
}

async function updateAdminLoginRouteNoticeUrl() {
    if (!isSuperUser(currentUser) || adminLoginRouteNoticeUpdateInProgress) {
        return;
    }

    const requestedUrl = normalizeAdminLoginNoticeUrl(ui?.adminLoginNoticeUrl?.value || '');
    if (!requestedUrl) {
        setFormMessage(ui.adminLoginNoticeMessage, 'Ingresa la URL de redireccion.', 'error');
        return;
    }

    adminLoginRouteNoticeUpdateInProgress = true;
    setAdminLoginRouteNoticeButtonsBusy(true);
    setFormMessage(ui.adminLoginNoticeMessage, 'Guardando URL de redireccion...', '');

    try {
        const data = await apiRequest('/system/login-route-notice/url', {
            method: 'PUT',
            body: JSON.stringify({
                url: requestedUrl
            })
        });
        applyAdminMaintenanceStatus(data?.maintenance || null);
        applyAdminLoginRouteNoticeStatus(data?.loginRouteNotice || null);
        setFormMessage(ui.adminLoginNoticeMessage, data?.message || 'URL de redireccion actualizada.', 'success');
    } catch (error) {
        setFormMessage(ui.adminLoginNoticeMessage, error.message || 'No fue posible actualizar la URL.', 'error');
    } finally {
        adminLoginRouteNoticeUpdateInProgress = false;
        setAdminLoginRouteNoticeButtonsBusy(false);
    }
}

async function handleAdminYearModeClick(mode) {
    if (!canUseHistoricalIngresoMode(currentUser)) {
        return;
    }

    setAdminUsersVisibility(false);
    setAdminBranchesVisibility(false);
    setAdminDocumentsVisibility(false);
    setAdminEmailVisibility(false);
    setAdminMailTemplatesVisibility(false);
    setAdminUsageVisibility(false);
    setAdminMaintenanceVisibility(false);
    await setAdminIngresoMode(mode);
}

function setAdminModeButtonsActive(mode) {
    if (!ui || !ui.adminYearCurrentBtn || !ui.adminYearPastBtn) {
        return;
    }

    const isHistorical = mode === 'historical';
    ui.adminYearCurrentBtn.classList.toggle('active', !isHistorical);
    ui.adminYearPastBtn.classList.toggle('active', isHistorical);
}

function setAdminYearMessage(message) {
    if (!ui || !ui.adminYearMessage) {
        return;
    }

    ui.adminYearMessage.textContent = message || '';
}

async function setAdminIngresoMode(mode, options = {}) {
    const normalizedMode = mode === 'historical' ? 'historical' : 'current';
    const refreshIngreso = options.refreshIngreso !== false;
    const clearIngreso = options.clearIngreso !== false;
    adminIngresoMode = normalizedMode;

    if (!canUseHistoricalIngresoMode(currentUser)) {
        return;
    }

    setAdminModeButtonsActive(normalizedMode);

    if (normalizedMode === 'historical') {
        ui.numIngreso.readOnly = false;
        ui.numIngreso.classList.add('admin-editable');
        ui.numIngresoHint.textContent = 'Modo ANOS ANTERIORES: escribe NRO INGRESO manual (ejemplo 120-2024).';
        setAdminYearMessage('Modo activo: ANOS ANTERIORES. Puedes crear o modificar registros historicos.');
        if (clearIngreso && !loadedIngresoOriginal) {
            ui.numIngreso.value = '';
        }
        return;
    }

    ui.numIngreso.readOnly = true;
    ui.numIngreso.classList.remove('admin-editable');
    ui.numIngresoHint.textContent = 'Modo ANO ACTUAL: autogenerado.';
    setAdminYearMessage('Modo activo: ANO ACTUAL. El NRO INGRESO se genera automaticamente.');
    if (refreshIngreso) {
        await loadNextIngreso(ui.numIngreso, ui.formMessage);
    }
}

function getAssignableRolesForCurrentUser() {
    if (isSuperUser(currentUser)) {
        return ['SUPER', 'ADMIN', 'SECRETARIA', 'OPERADOR', 'SUPERVISOR'];
    }

    return ['SECRETARIA', 'OPERADOR', 'SUPERVISOR'];
}

function ensureAdminRoleOptions() {
    if (!ui || !ui.adminRole) {
        return;
    }

    const expectedRoles = [
        { value: 'OPERADOR', label: 'OPERADOR' },
        { value: 'SUPERVISOR', label: 'SUPERVISOR' },
        { value: 'SECRETARIA', label: 'SECRETARIA' },
        { value: 'ADMIN', label: 'ADMIN' },
        { value: 'SUPER', label: 'SUPER' }
    ];

    const selectedRole = String(ui.adminRole.value || '').trim().toUpperCase();
    const existingOptions = new Map();
    Array.from(ui.adminRole.options).forEach((option) => {
        const normalizedValue = String(option.value || '').trim().toUpperCase();
        if (!normalizedValue || existingOptions.has(normalizedValue)) {
            return;
        }

        existingOptions.set(normalizedValue, option);
    });

    const fragment = document.createDocumentFragment();
    expectedRoles.forEach((roleItem) => {
        const option = existingOptions.get(roleItem.value) || document.createElement('option');
        option.value = roleItem.value;
        option.textContent = roleItem.label;
        option.hidden = false;
        option.disabled = false;
        fragment.appendChild(option);
    });

    ui.adminRole.replaceChildren(fragment);
    if (expectedRoles.some((roleItem) => roleItem.value === selectedRole)) {
        ui.adminRole.value = selectedRole;
    }
}

function syncAdminRoleOptionsForCurrentUser() {
    if (!ui || !ui.adminRole) {
        return;
    }

    ensureAdminRoleOptions();

    const canAssignSuper = isSuperUser(currentUser);
    const canAssignAdmin = isSuperUser(currentUser) || (editingAdminUserId && editingAdminUserRole === 'ADMIN');
    const options = Array.from(ui.adminRole.options);
    options.forEach((option) => {
        const optionValue = String(option.value || '').toUpperCase();
        if (optionValue === 'SUPER') {
            option.hidden = !canAssignSuper;
            option.disabled = !canAssignSuper;
            return;
        }
        if (optionValue === 'ADMIN') {
            option.hidden = !canAssignAdmin;
            option.disabled = !canAssignAdmin;
        }
    });

    const allowed = getAssignableRolesForCurrentUser();
    if (editingAdminUserId && editingAdminUserRole === 'ADMIN') {
        allowed.push('ADMIN');
    }
    const currentValue = String(ui.adminRole.value || '').toUpperCase();
    if (!allowed.includes(currentValue)) {
        ui.adminRole.value = 'OPERADOR';
    }
}

function clearAdminEmailConfigState() {
    adminEmailConfigLoaded = null;
    if (!ui) {
        return;
    }

    if (ui.adminEmailForm && typeof ui.adminEmailForm.reset === 'function') {
        ui.adminEmailForm.reset();
    }
    if (ui.adminEmailHost) {
        ui.adminEmailHost.value = '';
    }
    if (ui.adminEmailPort) {
        ui.adminEmailPort.value = '587';
    }
    if (ui.adminEmailSecure) {
        ui.adminEmailSecure.checked = false;
    }
    if (ui.adminEmailUser) {
        ui.adminEmailUser.value = '';
    }
    if (ui.adminEmailFromEmail) {
        ui.adminEmailFromEmail.value = '';
    }
    if (ui.adminEmailFromName) {
        ui.adminEmailFromName.value = 'Geo Rural';
    }
    if (ui.adminEmailKeepPassword) {
        ui.adminEmailKeepPassword.checked = true;
    }
    if (ui.adminEmailPassword) {
        ui.adminEmailPassword.value = '';
    }
    if (ui.adminEmailSource) {
        ui.adminEmailSource.textContent = '';
    }
    setFormMessage(ui.adminEmailMessage, '', '');
}

function applyAdminEmailConfigResult(result) {
    adminEmailConfigLoaded = result || null;
    if (!ui) {
        return;
    }

    const config = result?.config || {};
    if (ui.adminEmailHost) {
        ui.adminEmailHost.value = normalizeInvoiceText(config.smtpHost, 255);
    }
    if (ui.adminEmailPort) {
        const normalizedPort = Number(config.smtpPort);
        ui.adminEmailPort.value = Number.isInteger(normalizedPort) && normalizedPort > 0 ? String(normalizedPort) : '587';
    }
    if (ui.adminEmailSecure) {
        ui.adminEmailSecure.checked = Boolean(config.smtpSecure);
    }
    if (ui.adminEmailUser) {
        ui.adminEmailUser.value = normalizeInvoiceText(config.smtpUser, 255);
    }
    if (ui.adminEmailPassword) {
        ui.adminEmailPassword.value = '';
    }
    if (ui.adminEmailKeepPassword) {
        ui.adminEmailKeepPassword.checked = true;
    }
    if (ui.adminEmailFromEmail) {
        ui.adminEmailFromEmail.value = normalizeInvoiceText(config.smtpFromEmail, 255).toLowerCase();
    }
    if (ui.adminEmailFromName) {
        ui.adminEmailFromName.value = normalizeInvoiceText(config.smtpFromName || 'Geo Rural', 120);
    }

    const sourceMap = {
        DATABASE: 'Base de datos',
        ENV: 'Variables del servidor',
        NONE: 'Sin configuracion'
    };
    const sourceLabel = sourceMap[String(result?.source || '').toUpperCase()] || 'Sin configuracion';
    const readyLabel = result?.isReady ? 'Servicio listo para envio.' : 'Servicio incompleto: faltan datos.';
    const updatedAtText = normalizeInvoiceText(result?.updatedAt, 80);
    const parsedUpdatedAt = updatedAtText ? new Date(updatedAtText) : null;
    const updatedAtLabel =
        parsedUpdatedAt && !Number.isNaN(parsedUpdatedAt.getTime())
            ? parsedUpdatedAt.toLocaleString('es-CL')
            : updatedAtText;
    const updatedBy = normalizeInvoiceText(result?.updatedBy, 120);
    const details = [`Fuente activa: ${sourceLabel}`, readyLabel];
    if (updatedAtLabel) {
        details.push(`Actualizado: ${updatedAtLabel}`);
    }
    if (updatedBy) {
        details.push(`Por: ${updatedBy}`);
    }
    if (ui.adminEmailSource) {
        ui.adminEmailSource.textContent = details.join(' | ');
    }
}

async function loadAdminEmailConfig() {
    if (!isSuperUser(currentUser) || !ui) {
        return;
    }

    try {
        const result = await apiRequest('/admin/correo-config');
        applyAdminEmailConfigResult(result);
        setFormMessage(ui.adminEmailMessage, '', '');
    } catch (error) {
        setFormMessage(ui.adminEmailMessage, error.message, 'error');
    }
}

async function handleAdminEmailSave() {
    if (!isSuperUser(currentUser) || !ui) {
        setFormMessage(ui.adminEmailMessage, 'Solo el superusuario puede configurar el correo.', 'error');
        return;
    }

    const payload = {
        smtpHost: normalizeInvoiceText(ui.adminEmailHost?.value, 255),
        smtpPort: Number(ui.adminEmailPort?.value),
        smtpSecure: Boolean(ui.adminEmailSecure?.checked),
        smtpUser: normalizeInvoiceText(ui.adminEmailUser?.value, 255),
        smtpPassword: String(ui.adminEmailPassword?.value || '').trim(),
        keepPassword: Boolean(ui.adminEmailKeepPassword?.checked),
        smtpFromEmail: normalizeInvoiceText(ui.adminEmailFromEmail?.value, 255).toLowerCase(),
        smtpFromName: normalizeInvoiceText(ui.adminEmailFromName?.value, 120)
    };

    if (!payload.smtpHost) {
        setFormMessage(ui.adminEmailMessage, 'Ingresa el host SMTP.', 'error');
        return;
    }
    if (!Number.isInteger(payload.smtpPort) || payload.smtpPort <= 0 || payload.smtpPort > 65535) {
        setFormMessage(ui.adminEmailMessage, 'Ingresa un puerto SMTP valido (1-65535).', 'error');
        return;
    }
    if (!payload.smtpFromEmail) {
        setFormMessage(ui.adminEmailMessage, 'Ingresa el correo remitente.', 'error');
        return;
    }
    if (!isValidEmail(payload.smtpFromEmail)) {
        setFormMessage(ui.adminEmailMessage, 'El correo remitente no es valido.', 'error');
        return;
    }

    try {
        const result = await apiRequest('/admin/correo-config', {
            method: 'PUT',
            body: JSON.stringify(payload)
        });
        applyAdminEmailConfigResult(result);
        setFormMessage(ui.adminEmailMessage, result.message || 'Configuracion de correo guardada.', 'success');
    } catch (error) {
        setFormMessage(ui.adminEmailMessage, error.message, 'error');
    }
}

function normalizeMailTemplateType(value) {
    const normalized = String(value || '').trim();
    return MAIL_TEMPLATE_TYPES.includes(normalized) ? normalized : MAIL_TEMPLATE_TYPES[0];
}

function normalizeMailTemplateHtmlValue(value, maxLength = 400000) {
    const normalized = String(value || '').replace(/\u0000/g, '').trim();
    if (!normalized) {
        return '';
    }
    if (!Number.isInteger(maxLength) || maxLength <= 0) {
        return normalized;
    }
    return normalized.slice(0, maxLength);
}

function normalizeMailTemplatesObject(rawTemplates) {
    const source = rawTemplates && typeof rawTemplates === 'object' ? rawTemplates : {};
    return {
        cotizacionHtml: normalizeMailTemplateHtmlValue(source.cotizacionHtml),
        facturaSingleHtml: normalizeMailTemplateHtmlValue(source.facturaSingleHtml),
        facturaPendingHtml: normalizeMailTemplateHtmlValue(source.facturaPendingHtml)
    };
}

function escapeTemplateTokenRegex(value) {
    return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function applyMailTemplateHtmlString(template, replacements = {}) {
    let html = String(template || '');
    const entries = replacements && typeof replacements === 'object' ? Object.entries(replacements) : [];
    entries.forEach(([token, rawValue]) => {
        const pattern = new RegExp(`{{\\s*${escapeTemplateTokenRegex(token)}\\s*}}`, 'gi');
        html = html.replace(pattern, String(rawValue ?? ''));
    });
    return html;
}

function getRuntimeMailTemplateHtml(type) {
    const normalizedType = normalizeMailTemplateType(type);
    return normalizeMailTemplateHtmlValue(runtimeMailTemplates?.[normalizedType]);
}

function getMailTemplateHints(type) {
    const normalizedType = normalizeMailTemplateType(type);
    if (normalizedType === 'cotizacionHtml') {
        return [
            '{{SALUDO}}',
            '{{FECHA}}',
            '{{PERIODO_UTM}}',
            '{{UTM_VIGENTE}}',
            '{{REFERENCIA_BLOQUE_HTML}}',
            '{{TRAMO_BLOQUE_HTML}}',
            '{{TABLA_VALORES_HTML}}',
            '{{PARRAFOS_POST_TABLA_HTML}}',
            '{{ATENCION_NOMBRE}}',
            '{{ATENCION_SUCURSAL}}',
            '{{FIRMA_HTML}}',
            '{{LOGO_HTML}}'
        ];
    }
    if (normalizedType === 'facturaSingleHtml') {
        return ['{{TITULO}}', '{{RESUMEN_FILAS_HTML}}', '{{DETALLE_FILAS_HTML}}'];
    }
    return ['{{TITULO}}', '{{RESUMEN_FILAS_HTML}}', '{{TABLA_RESUMEN_HTML}}', '{{DETALLES_FACTURAS_HTML}}'];
}

function buildCotizacionPreviewReplacements() {
    const nowText = new Date().toLocaleString('es-CL');
    const sampleTable = `
        <table role="presentation" cellspacing="0" cellpadding="0" style="width:92%;max-width:680px;margin:0 auto;border-collapse:collapse;font-size:12px;color:#2d84d2;">
            <thead>
                <tr style="background:#e6f4ff;color:#1f74bf;">
                    <th style="border:1px solid #d9deea;padding:6px 8px;text-align:left;">Tramo lotes</th>
                    <th style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">Valor UTM</th>
                    <th style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">Neto CLP</th>
                    <th style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">IVA CLP</th>
                    <th style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">Total CLP</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="border:1px solid #d9deea;padding:6px 8px;">7 a 15</td>
                    <td style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">6,33</td>
                    <td style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">$ 439.321</td>
                    <td style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">$ 83.471</td>
                    <td style="border:1px solid #d9deea;padding:6px 8px;text-align:right;">$ 522.792</td>
                </tr>
            </tbody>
        </table>
    `.trim();
    const postTableParagraphs = `
        <p style="margin:12px 0 6px 0;font-weight:700;color:#1f74bf;">2).- Alcance general de los valores informados:</p>
        <p style="margin:0 0 6px 0;color:#2d84d2;">Los valores del cuadro son referenciales y corresponden al mes consultado.</p>
    `.trim();
    const signatureHtml = `
        <div style="margin-top:18px;">
            <p style="margin:0 0 16px 0;">Saludos cordiales,</p>
            <p style="margin:0 0 10px 0;">Pedro Antonio Gerardo Herrera Mendez</p>
            <p style="margin:0 0 10px 0;">Pedro Ignacio Albornoz Sateler</p>
            <p style="margin:0 0 10px 0;font-weight:700;color:#1f74bf;">GEO RURAL VERIFICACIONES LIMITADA</p>
            <p style="margin:0;">Uno Norte N° 801, oficina 306, Talca</p>
        </div>
    `.trim();
    return {
        SALUDO: 'Estimada/o Cliente de Prueba:',
        FECHA: escapeHtml(nowText),
        PERIODO_UTM: 'MARZO 2026',
        UTM_VIGENTE: '$ 69.403',
        REFERENCIA_BLOQUE_HTML: '<p style="margin:0 0 12px 0;"><strong>Referencia cliente:</strong> Carpeta SPR 7 lotes.</p>',
        TRAMO_BLOQUE_HTML:
            '<p style="margin:0 0 10px 0;"><strong>1).- Cotizacion por los servicios de verificacion:</strong></p><p style="margin:0 0 14px 0;">Tramo 7 a 15 lotes.</p>',
        TABLA_VALORES_HTML: sampleTable,
        PARRAFOS_POST_TABLA_HTML: postTableParagraphs,
        ATENCION_NOMBRE: 'Administrador Geo Rural',
        ATENCION_SUCURSAL: 'Casa Matriz',
        FIRMA_HTML: signatureHtml,
        LOGO_HTML: ''
    };
}

function buildFacturaSinglePreviewReplacements() {
    const sampleItem = {
        numIngreso: '28-2026',
        nombreRazonSocial: 'Cliente de Prueba SpA',
        rut: '76.123.456-7',
        giro: 'Servicios Profesionales',
        direccion: 'Uno Norte 801, Oficina 306',
        comuna: 'Talca',
        ciudad: 'Maule',
        contacto: 'cliente@correo.cl',
        montoFacturar: '125000',
        observacion: 'Factura para prueba de plantilla'
    };
    const summaryRowsHtml = `
        <tr><td style="padding:0 0 6px;font-weight:700;color:#22365f;">Fecha</td><td style="padding:0 0 6px;">${escapeHtml(
            new Date().toLocaleString('es-CL')
        )}</td></tr>
        <tr><td style="padding:0 0 6px;font-weight:700;color:#22365f;">Operador</td><td style="padding:0 0 6px;">Operador de Prueba</td></tr>
        <tr><td style="padding:0 0 12px;font-weight:700;color:#22365f;">Sucursal</td><td style="padding:0 0 12px;">Casa Matriz</td></tr>
    `.trim();
    return {
        TITULO: 'Solicitud de facturacion',
        RESUMEN_FILAS_HTML: summaryRowsHtml,
        DETALLE_FILAS_HTML: buildInvoiceDetailHtmlRows(sampleItem)
    };
}

function clampAdminPendingMailPreviewCount(value) {
    const numeric = Number.parseInt(value, 10);
    if (!Number.isFinite(numeric)) {
        return ADMIN_PENDING_MAIL_PREVIEW_DEFAULT_ITEMS;
    }
    return Math.min(ADMIN_PENDING_MAIL_PREVIEW_MAX_ITEMS, Math.max(ADMIN_PENDING_MAIL_PREVIEW_MIN_ITEMS, numeric));
}

function createAdminPendingMailPreviewItems(count = adminMailTemplatePendingPreviewCount) {
    const safeCount = clampAdminPendingMailPreviewCount(count);
    const items = [];
    for (let index = 0; index < safeCount; index += 1) {
        const correlativo = 28 + index;
        const baseAmount = 125000 + index * 18500;
        items.push({
            numIngreso: `${correlativo}-2026`,
            nombreRazonSocial: `Cliente de Prueba ${index + 1} SpA`,
            rut: `76.123.${String(450 + index).padStart(3, '0')}-7`,
            giro: 'Servicios Profesionales',
            direccion: `Uno Norte 801, Oficina ${306 + index}`,
            comuna: index % 2 === 0 ? 'Talca' : 'Curico',
            ciudad: 'Maule',
            contacto: `cliente${index + 1}@correo.cl`,
            montoFacturar: String(baseAmount),
            observacion: `Factura de prueba #${index + 1}`
        });
    }
    return items;
}

function setAdminMailTemplatePendingPreviewCount(value) {
    adminMailTemplatePendingPreviewCount = clampAdminPendingMailPreviewCount(value);
    updateAdminMailTemplatePendingControls(adminMailTemplateActiveType);
}

function updateAdminMailTemplatePendingControls(type = adminMailTemplateActiveType) {
    if (!ui || !ui.adminMailTemplatePendingControls) {
        return;
    }
    const normalizedType = normalizeMailTemplateType(type);
    const showControls = normalizedType === 'facturaPendingHtml';
    ui.adminMailTemplatePendingControls.classList.toggle('hidden', !showControls);
    if (!showControls) {
        return;
    }

    const safeCount = clampAdminPendingMailPreviewCount(adminMailTemplatePendingPreviewCount);
    adminMailTemplatePendingPreviewCount = safeCount;
    if (ui.adminMailTemplatePendingCountLabel) {
        ui.adminMailTemplatePendingCountLabel.textContent = `Vista previa: ${safeCount} factura${safeCount === 1 ? '' : 's'}`;
    }
    if (ui.adminMailTemplatePendingLessBtn) {
        ui.adminMailTemplatePendingLessBtn.disabled = safeCount <= ADMIN_PENDING_MAIL_PREVIEW_MIN_ITEMS;
    }
    if (ui.adminMailTemplatePendingMoreBtn) {
        ui.adminMailTemplatePendingMoreBtn.disabled = safeCount >= ADMIN_PENDING_MAIL_PREVIEW_MAX_ITEMS;
    }
}

function buildFacturaPendingPreviewReplacements() {
    const previewItems = createAdminPendingMailPreviewItems(adminMailTemplatePendingPreviewCount);
    const summaryRowsHtml = `
        <tr><td style="padding:0 0 6px;font-weight:700;color:#22365f;">Fecha</td><td style="padding:0 0 6px;">${escapeHtml(
            new Date().toLocaleString('es-CL')
        )}</td></tr>
        <tr><td style="padding:0 0 6px;font-weight:700;color:#22365f;">Operador</td><td style="padding:0 0 6px;">Operador de Prueba</td></tr>
        <tr><td style="padding:0 0 6px;font-weight:700;color:#22365f;">Sucursal</td><td style="padding:0 0 6px;">Casa Matriz</td></tr>
        <tr><td style="padding:0 0 12px;font-weight:700;color:#22365f;">Total pendientes</td><td style="padding:0 0 12px;">${escapeHtml(
            String(previewItems.length)
        )}</td></tr>
    `.trim();

    const summaryTableRowsHtml = previewItems
        .map((item, index) => {
            return `
                <tr>
                    <td style="padding:7px 8px;border:1px solid #d8e0ee;text-align:center;">${escapeHtml(String(index + 1))}</td>
                    <td style="padding:7px 8px;border:1px solid #d8e0ee;">${escapeHtml(item.numIngreso || '-')}</td>
                    <td style="padding:7px 8px;border:1px solid #d8e0ee;">${escapeHtml(item.nombreRazonSocial || '-')}</td>
                    <td style="padding:7px 8px;border:1px solid #d8e0ee;">${escapeHtml(formatInvoiceAmount(item.montoFacturar))}</td>
                    <td style="padding:7px 8px;border:1px solid #d8e0ee;">${escapeHtml(item.contacto || '-')}</td>
                </tr>
            `.trim();
        })
        .join('');

    const summaryTableHtml = `
        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
            <thead>
                <tr>
                    <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:center;">#</th>
                    <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">NRO INGRESO</th>
                    <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">CLIENTE / RAZON SOCIAL</th>
                    <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">MONTO</th>
                    <th style="padding:7px 8px;border:1px solid #d8e0ee;background:#eef3fb;text-align:left;">CONTACTO</th>
                </tr>
            </thead>
            <tbody>
                ${summaryTableRowsHtml}
            </tbody>
        </table>
    `.trim();

    const detailHtml = previewItems
        .map((item, index) => {
            return `
                <div style="margin-top:14px;">
                    <h4 style="margin:0 0 8px;font-size:14px;color:#163463;">Factura ${escapeHtml(String(index + 1))}</h4>
                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse;">
                        ${buildInvoiceDetailHtmlRows(item)}
                    </table>
                </div>
            `.trim();
        })
        .join('');

    return {
        TITULO: 'Resumen de facturas pendientes',
        RESUMEN_FILAS_HTML: summaryRowsHtml,
        TABLA_RESUMEN_HTML: summaryTableHtml,
        DETALLES_FACTURAS_HTML: detailHtml
    };
}

function buildAdminMailTemplatePreview(type, templateHtml) {
    const normalizedType = normalizeMailTemplateType(type);
    const sourceTemplate = normalizeMailTemplateHtmlValue(templateHtml);
    if (!sourceTemplate) {
        return '<html><body style="font-family:Arial,Helvetica,sans-serif;padding:12px;color:#2a3d6b;">Sin plantilla para previsualizar.</body></html>';
    }

    if (normalizedType === 'cotizacionHtml') {
        return applyMailTemplateHtmlString(sourceTemplate, buildCotizacionPreviewReplacements());
    }
    if (normalizedType === 'facturaSingleHtml') {
        return applyMailTemplateHtmlString(sourceTemplate, buildFacturaSinglePreviewReplacements());
    }
    return applyMailTemplateHtmlString(sourceTemplate, buildFacturaPendingPreviewReplacements());
}

function renderAdminMailTemplateHints(type = adminMailTemplateActiveType) {
    if (!ui || !ui.adminMailTemplateHints) {
        return;
    }
    const hints = getMailTemplateHints(type);
    ui.adminMailTemplateHints.textContent = `Tokens disponibles: ${hints.join(' | ')}`;
}

function persistCurrentAdminTemplateEditorValue() {
    if (!ui || !ui.adminMailTemplateEditor || !adminMailTemplateDrafts) {
        return;
    }
    const activeType = normalizeMailTemplateType(adminMailTemplateActiveType);
    adminMailTemplateDrafts[activeType] = normalizeMailTemplateHtmlValue(ui.adminMailTemplateEditor.value);
}

function renderAdminMailTemplateEditor(type = adminMailTemplateActiveType) {
    if (!ui || !ui.adminMailTemplateEditor || !ui.adminMailTemplateType) {
        return;
    }
    const normalizedType = normalizeMailTemplateType(type);
    adminMailTemplateActiveType = normalizedType;
    ui.adminMailTemplateType.value = normalizedType;
    ui.adminMailTemplateEditor.value = normalizeMailTemplateHtmlValue(adminMailTemplateDrafts?.[normalizedType]);
    renderAdminMailTemplateHints(normalizedType);
    updateAdminMailTemplatePendingControls(normalizedType);
}

function renderAdminMailTemplatePreview(type = adminMailTemplateActiveType) {
    if (!ui || !ui.adminMailTemplatePreviewFrame) {
        return;
    }
    const normalizedType = normalizeMailTemplateType(type);
    updateAdminMailTemplatePendingControls(normalizedType);
    const templateHtml = normalizeMailTemplateHtmlValue(adminMailTemplateDrafts?.[normalizedType]);
    ui.adminMailTemplatePreviewFrame.srcdoc = buildAdminMailTemplatePreview(normalizedType, templateHtml);
}

function applyAdminMailTemplatesResult(result) {
    const templates = normalizeMailTemplatesObject(result?.templates);
    adminMailTemplatesLoaded = result || null;
    adminMailTemplateDrafts = { ...templates };
    runtimeMailTemplates = { ...templates };

    if (!ui) {
        return;
    }

    const sourceUpdatedAt = normalizeInvoiceText(result?.updatedAt, 80);
    const parsedUpdatedAt = sourceUpdatedAt ? new Date(sourceUpdatedAt) : null;
    const updatedAtLabel =
        parsedUpdatedAt && !Number.isNaN(parsedUpdatedAt.getTime())
            ? parsedUpdatedAt.toLocaleString('es-CL')
            : sourceUpdatedAt;
    const updatedBy = normalizeInvoiceText(result?.updatedBy, 120);
    const details = [];
    if (updatedAtLabel) {
        details.push(`Actualizado: ${updatedAtLabel}`);
    }
    if (updatedBy) {
        details.push(`Por: ${updatedBy}`);
    }
    if (ui.adminMailTemplateSource) {
        ui.adminMailTemplateSource.textContent = details.join(' | ');
    }

    adminMailTemplateActiveType = normalizeMailTemplateType(ui.adminMailTemplateType?.value || adminMailTemplateActiveType);
    renderAdminMailTemplateEditor(adminMailTemplateActiveType);
    renderAdminMailTemplatePreview(adminMailTemplateActiveType);
}

async function loadRuntimeMailTemplates(options = {}) {
    try {
        const result = await apiRequest('/mail-templates');
        runtimeMailTemplates = normalizeMailTemplatesObject(result?.templates);
        return runtimeMailTemplates;
    } catch (error) {
        runtimeMailTemplates = null;
        if (options.showErrors && ui?.adminMailTemplateMessage) {
            setFormMessage(ui.adminMailTemplateMessage, error.message, 'error');
        }
        return null;
    }
}

async function loadAdminMailTemplates() {
    if (!isSuperUser(currentUser) || !ui) {
        return;
    }

    try {
        const result = await apiRequest('/admin/correo-plantillas');
        applyAdminMailTemplatesResult(result);
        setFormMessage(ui.adminMailTemplateMessage, '', '');
    } catch (error) {
        setFormMessage(ui.adminMailTemplateMessage, error.message, 'error');
    }
}

async function handleAdminMailTemplatesSave() {
    if (!isSuperUser(currentUser) || !ui) {
        setFormMessage(ui.adminMailTemplateMessage, 'Solo el superusuario puede editar plantillas de correo.', 'error');
        return;
    }

    persistCurrentAdminTemplateEditorValue();
    const payload = {
        templates: normalizeMailTemplatesObject(adminMailTemplateDrafts)
    };

    try {
        const result = await apiRequest('/admin/correo-plantillas', {
            method: 'PUT',
            body: JSON.stringify(payload)
        });
        applyAdminMailTemplatesResult(result);
        setFormMessage(ui.adminMailTemplateMessage, result.message || 'Plantillas guardadas correctamente.', 'success');
    } catch (error) {
        setFormMessage(ui.adminMailTemplateMessage, error.message, 'error');
    }
}

function clearAdminMailTemplatesState() {
    adminMailTemplatesLoaded = null;
    adminMailTemplateDrafts = null;
    adminMailTemplateActiveType = 'cotizacionHtml';
    adminMailTemplatePendingPreviewCount = ADMIN_PENDING_MAIL_PREVIEW_DEFAULT_ITEMS;
    if (!ui) {
        return;
    }
    if (ui.adminMailTemplateType) {
        ui.adminMailTemplateType.value = 'cotizacionHtml';
    }
    if (ui.adminMailTemplateEditor) {
        ui.adminMailTemplateEditor.value = '';
    }
    if (ui.adminMailTemplateHints) {
        ui.adminMailTemplateHints.textContent = '';
    }
    if (ui.adminMailTemplateSource) {
        ui.adminMailTemplateSource.textContent = '';
    }
    if (ui.adminMailTemplatePreviewFrame) {
        ui.adminMailTemplatePreviewFrame.srcdoc = '';
    }
    updateAdminMailTemplatePendingControls('cotizacionHtml');
    setFormMessage(ui.adminMailTemplateMessage, '', '');
}

function clearAdminUsageStatsState() {
    adminUsageStatsLoaded = null;
    if (!ui) {
        return;
    }

    if (ui.adminUsageSummary) {
        ui.adminUsageSummary.textContent = 'Sin datos de actividad disponibles.';
    }
    if (ui.adminUsageBody) {
        ui.adminUsageBody.innerHTML = '<tr><td colspan="5">Sin actividad para mostrar.</td></tr>';
    }
    setFormMessage(ui.adminUsageMessage, '', '');
}

function applyAdminUsageStatsResult(result) {
    adminUsageStatsLoaded = result || null;
    if (!ui) {
        return;
    }

    const windowMinutes = Math.max(1, parseInt(result?.windowMinutes, 10) || 15);
    const equiposActivos = Math.max(0, parseInt(result?.equiposActivos, 10) || 0);
    const usuariosActivos = Math.max(0, parseInt(result?.usuariosActivos, 10) || 0);
    const sesionesVigentes = Math.max(0, parseInt(result?.sesionesVigentes, 10) || 0);
    const updatedAtLabel = formatDateTime(result?.actualizadoEn);

    if (ui.adminUsageSummary) {
        const summaryParts = [
            `Equipos activos (ultimos ${windowMinutes} min): ${equiposActivos}`,
            `Usuarios activos: ${usuariosActivos}`,
            `Sesiones vigentes: ${sesionesVigentes}`
        ];
        if (updatedAtLabel) {
            summaryParts.push(`Actualizado: ${updatedAtLabel}`);
        }
        ui.adminUsageSummary.textContent = summaryParts.join(' | ');
    }

    if (!ui.adminUsageBody) {
        return;
    }

    const items = Array.isArray(result?.equipos) ? result.equipos : [];
    ui.adminUsageBody.innerHTML = '';

    if (items.length === 0) {
        ui.adminUsageBody.innerHTML = '<tr><td colspan="5">No hay equipos activos en este momento.</td></tr>';
        return;
    }

    items.forEach((item) => {
        const row = document.createElement('tr');
        const usuarioCell = document.createElement('td');
        const roleCell = document.createElement('td');
        const sucursalCell = document.createElement('td');
        const actividadCell = document.createElement('td');
        const expiracionCell = document.createElement('td');

        const userName = normalizeInvoiceText(item?.nombre || item?.username, 120);
        const username = normalizeInvoiceText(item?.username, 80);
        usuarioCell.textContent = userName && username ? `${userName} (${username})` : userName || username || 'Sin usuario';
        roleCell.textContent = normalizeUserRole(item?.role) || 'OPERADOR';
        sucursalCell.textContent = normalizeInvoiceText(item?.sucursal, 120) || 'Sin sucursal';
        actividadCell.textContent = formatDateTime(item?.lastActivityAt || item?.ultimaActividad);
        expiracionCell.textContent = formatDateTime(item?.expiresAt || item?.expiraEn);

        row.appendChild(usuarioCell);
        row.appendChild(roleCell);
        row.appendChild(sucursalCell);
        row.appendChild(actividadCell);
        row.appendChild(expiracionCell);
        ui.adminUsageBody.appendChild(row);
    });
}

async function loadAdminUsageStats() {
    if (!isSuperUser(currentUser) || !ui) {
        return;
    }

    try {
        const result = await apiRequest('/admin/sesiones-activas');
        applyAdminUsageStatsResult(result);
        setFormMessage(ui.adminUsageMessage, '', '');
    } catch (error) {
        setFormMessage(ui.adminUsageMessage, error.message, 'error');
    }
}

function resetAdminUserInputs() {
    ui.adminUsername.value = '';
    ui.adminNombre.value = '';
    ui.adminSucursal.value = '';
    ui.adminPassword.value = '';
    editingAdminUserRole = '';
    ui.adminRole.value = 'OPERADOR';
    syncAdminRoleOptionsForCurrentUser();
}

function populateAdminSucursalSelect(sucursales, selectedSucursal = '') {
    if (!ui || !ui.adminSucursal) {
        return;
    }

    const selectedValue = selectedSucursal || ui.adminSucursal.value || '';
    ui.adminSucursal.innerHTML = '<option value="">Seleccione sucursal</option>';

    const names = new Set();
    (sucursales || []).forEach((item) => {
        const branchName = String(item.nombre || '').trim();
        if (!branchName || names.has(branchName)) {
            return;
        }

        names.add(branchName);
        const option = document.createElement('option');
        option.value = branchName;
        option.textContent = branchName;
        ui.adminSucursal.appendChild(option);
    });

    if (selectedValue) {
        const existingOption = Array.from(ui.adminSucursal.options).find((option) => option.value === selectedValue);
        if (existingOption) {
            ui.adminSucursal.value = selectedValue;
        } else {
            const missingOption = document.createElement('option');
            missingOption.value = selectedValue;
            missingOption.textContent = `${selectedValue} (no disponible)`;
            ui.adminSucursal.appendChild(missingOption);
            ui.adminSucursal.value = selectedValue;
        }
    } else {
        ui.adminSucursal.value = '';
    }
}

function enterAdminEditMode(usuario) {
    editingAdminUserId = Number(usuario.id);
    editingAdminUserRole = normalizeUserRole(usuario.role || 'OPERADOR');
    ui.adminUsername.value = usuario.username || '';
    ui.adminNombre.value = usuario.nombre || '';
    populateAdminSucursalSelect(availableBranches, usuario.sucursal || '');
    ui.adminPassword.value = '';
    syncAdminRoleOptionsForCurrentUser();
    ui.adminRole.value = usuario.role || 'OPERADOR';
    ui.adminCreateUserBtn.textContent = 'GUARDAR CAMBIOS';
    if (ui.adminCancelEditBtn) {
        ui.adminCancelEditBtn.classList.remove('hidden');
    }
    setFormMessage(ui.adminMessage, `Editando usuario ${usuario.username}.`, 'success');
}

function exitAdminEditMode() {
    editingAdminUserId = null;
    editingAdminUserRole = '';
    resetAdminUserInputs();
    syncAdminRoleOptionsForCurrentUser();
    ui.adminCreateUserBtn.textContent = 'CREAR USUARIO';
    if (ui.adminCancelEditBtn) {
        ui.adminCancelEditBtn.classList.add('hidden');
    }
}

function handleAdminCancelEdit() {
    exitAdminEditMode();
    setFormMessage(ui.adminMessage, 'Edicion cancelada.', 'success');
}

function resetBranchInputs() {
    if (ui && ui.adminBranchName) {
        ui.adminBranchName.value = '';
    }
}

function enterBranchEditMode(sucursal) {
    editingBranchId = Number(sucursal.id);
    ui.adminBranchName.value = sucursal.nombre || '';
    ui.adminBranchSaveBtn.textContent = 'GUARDAR SUCURSAL';
    ui.adminBranchCancelBtn.classList.remove('hidden');
    setFormMessage(ui.adminBranchMessage, `Editando sucursal ${sucursal.nombre}.`, 'success');
}

function exitBranchEditMode() {
    editingBranchId = null;
    resetBranchInputs();
    if (ui && ui.adminBranchSaveBtn) {
        ui.adminBranchSaveBtn.textContent = 'AGREGAR SUCURSAL';
    }
    if (ui && ui.adminBranchCancelBtn) {
        ui.adminBranchCancelBtn.classList.add('hidden');
    }
}

function handleBranchCancelEdit() {
    exitBranchEditMode();
    setFormMessage(ui.adminBranchMessage, 'Edicion de sucursal cancelada.', 'success');
}

async function handleBranchSave() {
    if (!canManageAdminCatalogs(currentUser)) {
        setFormMessage(ui.adminBranchMessage, 'Solo ADMIN o SUPER pueden administrar sucursales.', 'error');
        return;
    }

    const nombre = String(ui.adminBranchName.value || '').trim();
    if (!nombre) {
        setFormMessage(ui.adminBranchMessage, 'Ingresa un nombre de sucursal.', 'error');
        return;
    }

    try {
        if (editingBranchId) {
            await apiRequest(`/admin/sucursales/${encodeURIComponent(editingBranchId)}`, {
                method: 'PUT',
                body: JSON.stringify({ nombre })
            });
            setFormMessage(ui.adminBranchMessage, `Sucursal ${nombre} actualizada correctamente.`, 'success');
        } else {
            await apiRequest('/admin/sucursales', {
                method: 'POST',
                body: JSON.stringify({ nombre })
            });
            setFormMessage(ui.adminBranchMessage, `Sucursal ${nombre} creada correctamente.`, 'success');
        }

        exitBranchEditMode();
        await loadBranches({ selectedSucursal: nombre });
        await loadAdminUsers();
    } catch (error) {
        setFormMessage(ui.adminBranchMessage, error.message, 'error');
    }
}

async function loadBranches(options = {}) {
    if (!canManageAdminCatalogs(currentUser)) {
        return;
    }

    try {
        const data = await apiRequest('/admin/sucursales');
        availableBranches = Array.isArray(data.sucursales) ? data.sucursales : [];
        renderBranches(availableBranches);
        populateAdminSucursalSelect(availableBranches, options.selectedSucursal || '');
    } catch (error) {
        setFormMessage(ui.adminBranchMessage, error.message, 'error');
    }
}

function renderBranches(sucursales) {
    if (!ui || !ui.adminBranchesBody) {
        return;
    }

    ui.adminBranchesBody.innerHTML = '';
    if (!Array.isArray(sucursales) || sucursales.length === 0) {
        ui.adminBranchesBody.innerHTML = '<tr><td colspan="2">Sin sucursales para mostrar.</td></tr>';
        return;
    }

    sucursales.forEach((sucursal) => {
        const row = document.createElement('tr');
        row.innerHTML = `<td>${escapeHtml(sucursal.nombre || '')}</td><td></td>`;

        const actionCell = row.lastElementChild;
        const editBtn = document.createElement('button');
        editBtn.type = 'button';
        editBtn.className = 'admin-edit-btn';
        editBtn.textContent = 'EDITAR';
        editBtn.addEventListener('click', () => enterBranchEditMode(sucursal));
        actionCell.appendChild(editBtn);

        const deleteBtn = document.createElement('button');
        deleteBtn.type = 'button';
        deleteBtn.className = 'admin-delete-btn';
        deleteBtn.textContent = 'ELIMINAR';
        deleteBtn.addEventListener('click', () => void handleBranchDelete(sucursal));
        actionCell.appendChild(deleteBtn);

        ui.adminBranchesBody.appendChild(row);
    });
}

async function handleBranchDelete(sucursal) {
    if (!canManageAdminCatalogs(currentUser)) {
        setFormMessage(ui.adminBranchMessage, 'Solo ADMIN o SUPER pueden eliminar sucursales.', 'error');
        return;
    }

    const confirmDelete = window.confirm(`Eliminar sucursal ${sucursal.nombre}?`);
    if (!confirmDelete) {
        return;
    }

    try {
        await apiRequest(`/admin/sucursales/${encodeURIComponent(sucursal.id)}`, {
            method: 'DELETE'
        });

        if (editingBranchId && Number(editingBranchId) === Number(sucursal.id)) {
            exitBranchEditMode();
        }
        setFormMessage(ui.adminBranchMessage, `Sucursal ${sucursal.nombre} eliminada.`, 'success');
        await loadBranches();
        await loadAdminUsers();
    } catch (error) {
        setFormMessage(ui.adminBranchMessage, error.message, 'error');
    }
}

function initializeDocumentOptions() {
    defaultDocumentGroups = getDefaultDocumentGroupsFromDom();
    defaultDocumentOptions = getDefaultDocumentOptionsFromDom(defaultDocumentGroups);
    configuredDocumentGroups = loadDocumentGroups(defaultDocumentGroups);
    configuredDocumentOptions = loadDocumentOptions(defaultDocumentOptions, configuredDocumentGroups);
    renderAdminDocumentGroupTitleInputs(configuredDocumentGroups);
    renderAdminDocumentGroupSelect(configuredDocumentGroups);
    renderDocumentOptions(configuredDocumentOptions);
    renderAdminDocuments(configuredDocumentOptions);
}

function buildDefaultDocumentGroups() {
    return DOCUMENT_GROUP_IDS.map((groupId, index) => ({
        id: groupId,
        label: DEFAULT_DOCUMENT_GROUP_TITLES[index] || `Seccion ${index + 1}`
    }));
}

function normalizeDocumentGroupId(value) {
    const normalized = String(value || '')
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9_]/g, '_')
        .replace(/_+/g, '_')
        .replace(/^_+|_+$/g, '');
    return DOCUMENT_GROUP_IDS.includes(normalized) ? normalized : '';
}

function normalizeDocumentGroupLabel(value) {
    return String(value || '')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, 80);
}

function sanitizeDocumentGroups(groups, fallbackGroups = null) {
    const fallbackMap = new Map();
    const fallbackSource = Array.isArray(fallbackGroups) && fallbackGroups.length > 0 ? fallbackGroups : buildDefaultDocumentGroups();
    fallbackSource.forEach((item, index) => {
        const defaultId = DOCUMENT_GROUP_IDS[index];
        const id = normalizeDocumentGroupId(item?.id) || defaultId;
        const label = normalizeDocumentGroupLabel(item?.label) || DEFAULT_DOCUMENT_GROUP_TITLES[index] || `Seccion ${index + 1}`;
        if (id && !fallbackMap.has(id)) {
            fallbackMap.set(id, label);
        }
    });

    const sourceMap = new Map();
    (Array.isArray(groups) ? groups : []).forEach((item) => {
        const id = normalizeDocumentGroupId(item?.id);
        const label = normalizeDocumentGroupLabel(item?.label);
        if (id && label) {
            sourceMap.set(id, label);
        }
    });

    return DOCUMENT_GROUP_IDS.map((groupId, index) => ({
        id: groupId,
        label: sourceMap.get(groupId) || fallbackMap.get(groupId) || DEFAULT_DOCUMENT_GROUP_TITLES[index] || `Seccion ${index + 1}`
    }));
}

function cloneDocumentGroups(groups) {
    return sanitizeDocumentGroups(groups).map((item) => ({ ...item }));
}

function getDefaultDocumentGroupsFromDom() {
    if (!ui || !ui.documentosContainer) {
        return buildDefaultDocumentGroups();
    }

    const groupsFromDom = Array.from(ui.documentosContainer.querySelectorAll('[data-doc-group-id]')).map((groupNode) => ({
        id: normalizeDocumentGroupId(groupNode?.dataset?.docGroupId),
        label: normalizeDocumentGroupLabel(groupNode?.dataset?.docGroupTitle || groupNode.querySelector('.docs-group-title')?.textContent)
    }));

    return sanitizeDocumentGroups(groupsFromDom, buildDefaultDocumentGroups());
}

function getDefaultDocumentOptionsFromDom(groups = defaultDocumentGroups) {
    if (!ui || !ui.documentosContainer) {
        return [];
    }

    const options = [];
    ui.documentosContainer.querySelectorAll('label.doc-item').forEach((labelNode) => {
        const inputNode = labelNode.querySelector('input[name="documentos"]');
        if (!inputNode) {
            return;
        }

        const value = normalizeDocumentOptionValue(inputNode.value);
        const label = normalizeDocumentOptionLabel(labelNode.textContent);
        if (!value || !label) {
            return;
        }

        const groupIdFromNode = normalizeDocumentGroupId(
            inputNode?.dataset?.docGroupId ||
                labelNode?.dataset?.docGroupId ||
                labelNode.closest('[data-doc-group-id]')?.dataset?.docGroupId
        );

        options.push({ value, label, groupId: groupIdFromNode });
    });

    return sanitizeDocumentOptions(options, [], groups);
}

function normalizeDocumentOptionLabel(value) {
    return String(value || '')
        .replace(/\s+/g, ' ')
        .trim();
}

function normalizeDocumentOptionValue(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9_]/g, '_')
        .replace(/_+/g, '_')
        .replace(/^_+|_+$/g, '')
        .slice(0, 80);
}

function buildDocumentOptionValue(label) {
    const normalizedLabel = String(label || '')
        .toLocaleLowerCase('es-CL')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');

    const value = normalizeDocumentOptionValue(normalizedLabel.replace(/[^a-z0-9]+/g, '_'));
    return value || 'documento';
}

function sanitizeDocumentOptions(options, fallbackOptions = [], groups = configuredDocumentGroups) {
    const safeGroups = cloneDocumentGroups(groups);
    const validGroupIds = new Set(safeGroups.map((item) => item.id));
    const fallbackGroupMap = new Map();
    (Array.isArray(fallbackOptions) ? fallbackOptions : []).forEach((item) => {
        const value = normalizeDocumentOptionValue(item?.value);
        const groupId = normalizeDocumentGroupId(item?.groupId);
        if (value && groupId && validGroupIds.has(groupId) && !fallbackGroupMap.has(value)) {
            fallbackGroupMap.set(value, groupId);
        }
    });

    const sanitized = [];
    const seenValues = new Set();
    let autoGroupIndex = 0;
    (Array.isArray(options) ? options : []).forEach((item) => {
        const value = normalizeDocumentOptionValue(item?.value);
        const label = normalizeDocumentOptionLabel(item?.label);
        if (!value || !label || seenValues.has(value)) {
            return;
        }

        let groupId = normalizeDocumentGroupId(item?.groupId);
        if (!validGroupIds.has(groupId)) {
            groupId = '';
        }
        if (!groupId) {
            const fallbackGroupId = fallbackGroupMap.get(value);
            if (fallbackGroupId && validGroupIds.has(fallbackGroupId)) {
                groupId = fallbackGroupId;
            }
        }
        if (!groupId) {
            groupId = safeGroups[autoGroupIndex % safeGroups.length]?.id || DOCUMENT_GROUP_IDS[0];
            autoGroupIndex += 1;
        }

        seenValues.add(value);
        sanitized.push({ value, label, groupId });
    });
    return sanitized;
}

function cloneDocumentOptions(options, groups = configuredDocumentGroups) {
    return sanitizeDocumentOptions(options, [], groups).map((item) => ({ ...item }));
}

function loadDocumentGroups(fallbackGroups) {
    const fallback = cloneDocumentGroups(fallbackGroups);

    try {
        const raw = window.localStorage.getItem(DOCUMENT_OPTIONS_STORAGE);
        if (!raw) {
            return fallback;
        }

        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
            return fallback;
        }
        return sanitizeDocumentGroups(parsed?.groups, fallback);
    } catch (error) {
        return fallback;
    }
}

function loadDocumentOptions(fallbackOptions, groups = configuredDocumentGroups) {
    const safeGroups = cloneDocumentGroups(groups);
    const fallback = sanitizeDocumentOptions(fallbackOptions, fallbackOptions, safeGroups);

    try {
        const raw = window.localStorage.getItem(DOCUMENT_OPTIONS_STORAGE);
        if (!raw) {
            return fallback;
        }

        const parsed = JSON.parse(raw);
        const sourceOptions = !Array.isArray(parsed) && parsed && Array.isArray(parsed.options) ? parsed.options : parsed;
        const sanitized = sanitizeDocumentOptions(sourceOptions, fallback, safeGroups);
        return sanitized.length > 0 ? sanitized : fallback;
    } catch (error) {
        return fallback;
    }
}

function persistDocumentOptions(options) {
    try {
        const payload = {
            version: DOCUMENT_OPTIONS_STORAGE_VERSION,
            groups: cloneDocumentGroups(configuredDocumentGroups),
            options: cloneDocumentOptions(options, configuredDocumentGroups)
        };
        window.localStorage.setItem(DOCUMENT_OPTIONS_STORAGE, JSON.stringify(payload));
    } catch (error) {
        // Si localStorage no esta disponible, se mantiene solo en memoria.
    }
}

function getCurrentCheckedDocumentValues() {
    if (!ui || !ui.documentosContainer) {
        return [];
    }

    return Array.from(ui.documentosContainer.querySelectorAll('input[name="documentos"]:checked'))
        .map((input) => String(input.value || '').trim())
        .filter((value) => value.length > 0);
}

function renderDocumentOptions(options, selectedValues = null) {
    if (!ui || !ui.documentosContainer) {
        return;
    }

    const safeGroups = cloneDocumentGroups(configuredDocumentGroups.length > 0 ? configuredDocumentGroups : defaultDocumentGroups);
    const safeOptions = sanitizeDocumentOptions(options, defaultDocumentOptions, safeGroups);
    const selectedSet = new Set(
        Array.isArray(selectedValues) ? selectedValues.map((value) => String(value || '').trim()) : getCurrentCheckedDocumentValues()
    );
    ui.documentosContainer.innerHTML = '';

    if (!Array.isArray(safeOptions) || safeOptions.length === 0) {
        const emptyNode = document.createElement('p');
        emptyNode.className = 'docs-empty-message';
        emptyNode.textContent = 'Sin documentos configurados.';
        ui.documentosContainer.appendChild(emptyNode);
        return;
    }

    safeGroups.forEach((group) => {
        const groupSection = document.createElement('section');
        groupSection.className = 'docs-group';
        groupSection.dataset.docGroupId = group.id;
        groupSection.dataset.docGroupTitle = group.label;

        const titleNode = document.createElement('h5');
        titleNode.className = 'docs-group-title';
        titleNode.textContent = group.label;
        groupSection.appendChild(titleNode);

        const gridNode = document.createElement('div');
        gridNode.className = 'docs-grid';

        const groupOptions = safeOptions.filter((option) => normalizeDocumentGroupId(option?.groupId) === group.id);
        if (groupOptions.length === 0) {
            const groupEmptyNode = document.createElement('p');
            groupEmptyNode.className = 'docs-group-empty';
            groupEmptyNode.textContent = 'Sin documentos en esta seccion.';
            groupSection.appendChild(groupEmptyNode);
            ui.documentosContainer.appendChild(groupSection);
            return;
        }

        groupOptions.forEach((option) => {
            const itemLabel = document.createElement('label');
            itemLabel.className = 'doc-item';
            itemLabel.dataset.docGroupId = group.id;

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.name = 'documentos';
            checkbox.value = option.value;
            checkbox.checked = selectedSet.has(option.value);
            checkbox.dataset.docGroupId = group.id;

            itemLabel.appendChild(checkbox);
            itemLabel.appendChild(document.createTextNode(` ${option.label}`));
            gridNode.appendChild(itemLabel);
        });

        groupSection.appendChild(gridNode);
        ui.documentosContainer.appendChild(groupSection);
    });
}

function renderAdminDocuments(options) {
    if (!ui || !ui.adminDocumentsBody) {
        return;
    }

    const safeOptions = sanitizeDocumentOptions(options, defaultDocumentOptions, configuredDocumentGroups);
    ui.adminDocumentsBody.innerHTML = '';
    if (!Array.isArray(safeOptions) || safeOptions.length === 0) {
        ui.adminDocumentsBody.innerHTML = '<tr><td colspan="4">Sin documentos para mostrar.</td></tr>';
        return;
    }

    safeOptions.forEach((option) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${escapeHtml(option.label)}</td>
            <td><span class="admin-doc-id">${escapeHtml(option.value)}</span></td>
            <td>${escapeHtml(getDocumentGroupLabelById(option.groupId))}</td>
            <td></td>
        `;

        const actionCell = row.lastElementChild;
        const editBtn = document.createElement('button');
        editBtn.type = 'button';
        editBtn.className = 'admin-edit-btn';
        editBtn.textContent = 'EDITAR';
        editBtn.addEventListener('click', () => enterDocumentEditMode(option));
        actionCell.appendChild(editBtn);

        const deleteBtn = document.createElement('button');
        deleteBtn.type = 'button';
        deleteBtn.className = 'admin-delete-btn';
        deleteBtn.textContent = 'ELIMINAR';
        deleteBtn.addEventListener('click', () => handleAdminDocumentDelete(option));
        actionCell.appendChild(deleteBtn);

        ui.adminDocumentsBody.appendChild(row);
    });
}

function getDocumentGroupLabelById(groupId) {
    const normalizedGroupId = normalizeDocumentGroupId(groupId);
    if (!normalizedGroupId) {
        return 'Seccion';
    }

    const match = cloneDocumentGroups(configuredDocumentGroups).find((item) => item.id === normalizedGroupId);
    return normalizeDocumentGroupLabel(match?.label || '').trim() || 'Seccion';
}

function renderAdminDocumentGroupSelect(groups, selectedGroupId = '') {
    if (!ui?.adminDocumentGroup) {
        return;
    }

    const safeGroups = cloneDocumentGroups(groups);
    const selected = normalizeDocumentGroupId(selectedGroupId);
    const selectedIsValid = safeGroups.some((group) => group.id === selected);
    const fallbackGroupId = safeGroups[0]?.id || DOCUMENT_GROUP_IDS[0];

    ui.adminDocumentGroup.innerHTML = '';
    safeGroups.forEach((group) => {
        const optionNode = document.createElement('option');
        optionNode.value = group.id;
        optionNode.textContent = group.label;
        ui.adminDocumentGroup.appendChild(optionNode);
    });
    ui.adminDocumentGroup.value = selectedIsValid ? selected : fallbackGroupId;
}

function renderAdminDocumentGroupTitleInputs(groups) {
    if (!ui?.adminDocumentGroupTitles) {
        return;
    }

    const safeGroups = cloneDocumentGroups(groups);
    ui.adminDocumentGroupTitles.innerHTML = '';

    safeGroups.forEach((group, index) => {
        const fieldWrap = document.createElement('label');
        fieldWrap.className = 'admin-doc-group-item';

        const titleNode = document.createElement('span');
        titleNode.textContent = `Seccion ${index + 1}`;
        fieldWrap.appendChild(titleNode);

        const inputNode = document.createElement('input');
        inputNode.type = 'text';
        inputNode.maxLength = 80;
        inputNode.value = group.label;
        inputNode.dataset.docGroupId = group.id;
        fieldWrap.appendChild(inputNode);

        ui.adminDocumentGroupTitles.appendChild(fieldWrap);
    });
}

function readDocumentGroupsFromInputs(fallbackGroups) {
    const fallback = cloneDocumentGroups(fallbackGroups);
    if (!ui?.adminDocumentGroupTitles) {
        return fallback;
    }

    const parsed = fallback.map((group) => {
        const inputNode = ui.adminDocumentGroupTitles.querySelector(`input[data-doc-group-id="${group.id}"]`);
        const label = normalizeDocumentGroupLabel(inputNode?.value) || group.label;
        return { id: group.id, label };
    });
    return sanitizeDocumentGroups(parsed, fallback);
}

function handleAdminDocumentGroupTitlesSave() {
    if (!canManageAdminCatalogs(currentUser)) {
        setFormMessage(ui.adminDocumentMessage, 'Solo ADMIN o SUPER pueden editar secciones de documentos.', 'error');
        return;
    }

    const updatedGroups = readDocumentGroupsFromInputs(configuredDocumentGroups.length > 0 ? configuredDocumentGroups : defaultDocumentGroups);
    configuredDocumentGroups = updatedGroups;
    configuredDocumentOptions = sanitizeDocumentOptions(configuredDocumentOptions, defaultDocumentOptions, configuredDocumentGroups);
    persistDocumentOptions(configuredDocumentOptions);
    renderAdminDocumentGroupTitleInputs(configuredDocumentGroups);
    renderAdminDocumentGroupSelect(configuredDocumentGroups, ui?.adminDocumentGroup?.value);
    renderDocumentOptions(configuredDocumentOptions);
    renderAdminDocuments(configuredDocumentOptions);
    setFormMessage(ui.adminDocumentMessage, 'Titulos de secciones actualizados correctamente.', 'success');
}

function resetDocumentInputs() {
    if (!ui || !ui.adminDocumentLabel) {
        return;
    }

    ui.adminDocumentLabel.value = '';
    renderAdminDocumentGroupSelect(configuredDocumentGroups);
}

function enterDocumentEditMode(option) {
    if (!ui || !ui.adminDocumentLabel || !ui.adminDocumentSaveBtn || !ui.adminDocumentCancelBtn) {
        return;
    }

    editingDocumentValue = option.value;
    ui.adminDocumentLabel.value = option.label;
    const selectedGroupId = normalizeDocumentGroupId(option?.groupId) || configuredDocumentGroups[0]?.id || DOCUMENT_GROUP_IDS[0];
    renderAdminDocumentGroupSelect(configuredDocumentGroups, selectedGroupId);
    ui.adminDocumentSaveBtn.textContent = 'GUARDAR DOCUMENTO';
    ui.adminDocumentCancelBtn.classList.remove('hidden');
    setFormMessage(ui.adminDocumentMessage, `Editando documento ${option.label}.`, 'success');
}

function exitDocumentEditMode() {
    editingDocumentValue = '';
    resetDocumentInputs();

    if (ui && ui.adminDocumentSaveBtn) {
        ui.adminDocumentSaveBtn.textContent = 'AGREGAR DOCUMENTO';
    }
    if (ui && ui.adminDocumentCancelBtn) {
        ui.adminDocumentCancelBtn.classList.add('hidden');
    }
}

function handleAdminDocumentCancel() {
    exitDocumentEditMode();
    setFormMessage(ui.adminDocumentMessage, 'Edicion de documento cancelada.', 'success');
}

function handleAdminDocumentSave() {
    if (!canManageAdminCatalogs(currentUser)) {
        setFormMessage(ui.adminDocumentMessage, 'Solo ADMIN o SUPER pueden gestionar documentos.', 'error');
        return;
    }

    const label = normalizeDocumentOptionLabel(ui.adminDocumentLabel?.value);
    if (!label) {
        setFormMessage(ui.adminDocumentMessage, 'Ingresa un nombre para el documento.', 'error');
        return;
    }

    if (label.length > 120) {
        setFormMessage(ui.adminDocumentMessage, 'El nombre del documento no puede exceder 120 caracteres.', 'error');
        return;
    }

    const selectedGroupId = normalizeDocumentGroupId(ui?.adminDocumentGroup?.value);
    const validGroupIds = new Set(cloneDocumentGroups(configuredDocumentGroups).map((item) => item.id));
    if (!selectedGroupId || !validGroupIds.has(selectedGroupId)) {
        setFormMessage(ui.adminDocumentMessage, 'Selecciona una seccion valida para el documento.', 'error');
        return;
    }

    const options = cloneDocumentOptions(configuredDocumentOptions, configuredDocumentGroups);
    if (editingDocumentValue) {
        const targetIndex = options.findIndex((item) => item.value === editingDocumentValue);
        if (targetIndex === -1) {
            setFormMessage(ui.adminDocumentMessage, 'No fue posible encontrar el documento en edicion.', 'error');
            exitDocumentEditMode();
            return;
        }

        options[targetIndex].label = label;
        options[targetIndex].groupId = selectedGroupId;
        configuredDocumentOptions = options;
        persistDocumentOptions(configuredDocumentOptions);
        renderDocumentOptions(configuredDocumentOptions);
        renderAdminDocuments(configuredDocumentOptions);
        exitDocumentEditMode();
        setFormMessage(ui.adminDocumentMessage, `Documento ${label} actualizado correctamente.`, 'success');
        return;
    }

    let generatedValue = buildDocumentOptionValue(label);
    if (!generatedValue) {
        setFormMessage(ui.adminDocumentMessage, 'No fue posible generar el identificador del documento.', 'error');
        return;
    }

    const existingValues = new Set(options.map((item) => item.value));
    if (existingValues.has(generatedValue)) {
        let suffix = 2;
        while (existingValues.has(`${generatedValue}_${suffix}`)) {
            suffix += 1;
        }
        generatedValue = `${generatedValue}_${suffix}`;
    }

    options.push({ value: generatedValue, label, groupId: selectedGroupId });
    configuredDocumentOptions = options;
    persistDocumentOptions(configuredDocumentOptions);
    renderDocumentOptions(configuredDocumentOptions);
    renderAdminDocuments(configuredDocumentOptions);
    exitDocumentEditMode();
    setFormMessage(ui.adminDocumentMessage, `Documento ${label} agregado correctamente.`, 'success');
}

function handleAdminDocumentDelete(option) {
    if (!canManageAdminCatalogs(currentUser)) {
        setFormMessage(ui.adminDocumentMessage, 'Solo ADMIN o SUPER pueden eliminar documentos.', 'error');
        return;
    }

    const confirmDelete = window.confirm(`Eliminar documento ${option.label}?`);
    if (!confirmDelete) {
        return;
    }

    configuredDocumentOptions = cloneDocumentOptions(configuredDocumentOptions, configuredDocumentGroups).filter(
        (item) => item.value !== option.value
    );
    persistDocumentOptions(configuredDocumentOptions);
    renderDocumentOptions(configuredDocumentOptions);
    renderAdminDocuments(configuredDocumentOptions);

    if (editingDocumentValue === option.value) {
        exitDocumentEditMode();
    }

    setFormMessage(ui.adminDocumentMessage, `Documento ${option.label} eliminado.`, 'success');
}

async function handleAdminCreateUser() {
    if (!canManageAdminCatalogs(currentUser)) {
        setFormMessage(ui.adminMessage, 'Solo ADMIN o SUPER pueden crear usuarios.', 'error');
        return;
    }

    const username = ui.adminUsername.value.trim();
    const nombre = ui.adminNombre.value.trim();
    const sucursal = ui.adminSucursal.value.trim();
    const role = String(ui.adminRole.value || '').trim().toUpperCase();
    const password = ui.adminPassword.value;
    const assignableRoles = getAssignableRolesForCurrentUser();
    if (editingAdminUserId && editingAdminUserRole === 'ADMIN') {
        assignableRoles.push('ADMIN');
    }

    if (!assignableRoles.includes(role)) {
        setFormMessage(ui.adminMessage, `Rol invalido. Usa ${assignableRoles.join(', ')}.`, 'error');
        return;
    }

    if (!username || !nombre || !sucursal) {
        setFormMessage(ui.adminMessage, 'Completa usuario, nombre y sucursal.', 'error');
        return;
    }

    if (editingAdminUserId) {
        if (password && password.length < 6) {
            setFormMessage(ui.adminMessage, 'Si vas a cambiar la contrasena, debe tener al menos 6 caracteres.', 'error');
            return;
        }

        const payload = { username, nombre, sucursal, role };
        if (password) {
            payload.password = password;
        }

        try {
            const result = await apiRequest(`/admin/usuarios/${encodeURIComponent(editingAdminUserId)}`, {
                method: 'PUT',
                body: JSON.stringify(payload)
            });

            const editedSelf = currentUser && Number(currentUser.id) === Number(editingAdminUserId);
            if (editedSelf) {
                currentUser.username = username;
                currentUser.nombre = nombre;
                currentUser.sucursal = sucursal;
                currentUser.role = role;
                if (result?.usuario && Object.prototype.hasOwnProperty.call(result.usuario, 'mustChangePassword')) {
                    currentUser.mustChangePassword = Boolean(result.usuario.mustChangePassword);
                }
                await activateAuthenticatedUser(currentUser, {
                    passwordMessage: currentUser.mustChangePassword ? 'Debes cambiar tu clave temporal para continuar.' : ''
                });
            }

            exitAdminEditMode();
            const passwordNote = password ? ' y contrasena actualizada' : '';
            setFormMessage(ui.adminMessage, `Usuario ${username} actualizado correctamente${passwordNote}.`, 'success');
            await loadAdminUsers();
        } catch (error) {
            setFormMessage(ui.adminMessage, error.message, 'error');
        }
        return;
    }

    if (password.length < 6) {
        setFormMessage(ui.adminMessage, 'La contrasena debe tener al menos 6 caracteres.', 'error');
        return;
    }

    const payload = {
        username,
        nombre,
        sucursal,
        password,
        role
    };

    try {
        await apiRequest('/admin/usuarios', {
            method: 'POST',
            body: JSON.stringify(payload)
        });

        resetAdminUserInputs();
        const roleLabel = String(payload.role || '').toUpperCase();
        const message =
            roleLabel === 'SUPER'
                ? `Usuario ${payload.username} creado con clave fija de cuenta SUPER.`
                : `Usuario ${payload.username} creado. Debe cambiar su clave temporal en el primer ingreso.`;
        setFormMessage(ui.adminMessage, message, 'success');
        await loadAdminUsers();
    } catch (error) {
        setFormMessage(ui.adminMessage, error.message, 'error');
    }
}

async function loadAdminUsers() {
    if (!canManageAdminCatalogs(currentUser)) {
        return;
    }

    try {
        const data = await apiRequest('/admin/usuarios');
        renderAdminUsers(data.usuarios || []);
    } catch (error) {
        setFormMessage(ui.adminMessage, error.message, 'error');
    }
}

function renderAdminUsers(usuarios) {
    if (!ui || !ui.adminUsersBody) {
        return;
    }

    ui.adminUsersBody.innerHTML = '';
    if (!Array.isArray(usuarios) || usuarios.length === 0) {
        ui.adminUsersBody.innerHTML = '<tr><td colspan="6">Sin usuarios para mostrar.</td></tr>';
        return;
    }

    const visibleUsers = isSuperUser(currentUser)
        ? [...usuarios]
        : [...usuarios].filter((item) => {
              const isGuestAccount =
                  Boolean(item && item.isGuest) ||
                  String(item && item.username ? item.username : '')
                      .trim()
                      .toLowerCase() === 'invitado';
              return !isGuestAccount;
          });

    const sortedUsers = visibleUsers.sort((left, right) => {
        const leftIsGuest =
            Boolean(left && left.isGuest) || String(left && left.username ? left.username : '').trim().toLowerCase() === 'invitado';
        const rightIsGuest =
            Boolean(right && right.isGuest) || String(right && right.username ? right.username : '').trim().toLowerCase() === 'invitado';
        if (leftIsGuest !== rightIsGuest) {
            return leftIsGuest ? 1 : -1;
        }

        return Number(left?.id || 0) - Number(right?.id || 0);
    });

    sortedUsers.forEach((usuario) => {
        const row = document.createElement('tr');
        const isGuestAccount =
            Boolean(usuario && usuario.isGuest) || String(usuario && usuario.username ? usuario.username : '').trim().toLowerCase() === 'invitado';
        const mustChangePasswordText = isGuestAccount ? 'N/A' : usuario.mustChangePassword ? 'SI' : 'NO';
        const roleLabel = isGuestAccount ? 'INVITADO' : usuario.role || 'OPERADOR';

        row.innerHTML = `
            <td>${escapeHtml(usuario.username || '')}</td>
            <td>${escapeHtml(usuario.nombre || '')}</td>
            <td>${escapeHtml(usuario.sucursal || '')}</td>
            <td>${escapeHtml(roleLabel)}</td>
            <td>${mustChangePasswordText}</td>
            <td></td>
        `;

        const actionCell = row.lastElementChild;
        if (currentUser) {
            if (isGuestAccount) {
                const guestLabel = document.createElement('label');
                guestLabel.className = 'doc-item';
                guestLabel.textContent = ' Invitado habilitado';

                const guestCheckbox = document.createElement('input');
                guestCheckbox.type = 'checkbox';
                guestCheckbox.checked = Boolean(usuario.activo);
                guestCheckbox.addEventListener('change', () => void handleAdminGuestStatusToggle(usuario, guestCheckbox));

                guestLabel.prepend(guestCheckbox);
                actionCell.appendChild(guestLabel);
            } else {
                const editBtn = document.createElement('button');
                editBtn.type = 'button';
                editBtn.className = 'admin-edit-btn';
                editBtn.textContent = 'EDITAR';
                editBtn.addEventListener('click', () => enterAdminEditMode(usuario));
                actionCell.appendChild(editBtn);

                if (Number(usuario.id) !== Number(currentUser.id)) {
                    const resetBtn = document.createElement('button');
                    resetBtn.type = 'button';
                    resetBtn.className = 'admin-reset-btn';
                    resetBtn.textContent = 'RESET CLAVE';
                    resetBtn.addEventListener('click', () => void handleAdminResetPassword(usuario));
                    actionCell.appendChild(resetBtn);

                    const deleteBtn = document.createElement('button');
                    deleteBtn.type = 'button';
                    deleteBtn.className = 'admin-delete-btn';
                    deleteBtn.textContent = 'ELIMINAR';
                    deleteBtn.addEventListener('click', () => void handleAdminDeleteUser(usuario));
                    actionCell.appendChild(deleteBtn);
                } else {
                    const currentTag = document.createElement('span');
                    currentTag.textContent = 'Actual';
                    actionCell.appendChild(currentTag);
                }
            }
        } else {
            actionCell.textContent = '-';
        }

        ui.adminUsersBody.appendChild(row);
    });
}

async function handleAdminGuestStatusToggle(usuario, checkboxElement) {
    if (!isSuperUser(currentUser)) {
        if (checkboxElement) {
            checkboxElement.checked = !checkboxElement.checked;
        }
        setFormMessage(ui.adminMessage, 'Solo SUPER puede cambiar el estado del invitado.', 'error');
        return;
    }

    const nextActive = Boolean(checkboxElement && checkboxElement.checked);
    const actionLabel = nextActive ? 'habilitar' : 'deshabilitar';
    const confirmToggle = window.confirm(`Deseas ${actionLabel} la cuenta invitado?`);
    if (!confirmToggle) {
        if (checkboxElement) {
            checkboxElement.checked = !nextActive;
        }
        return;
    }

    if (checkboxElement) {
        checkboxElement.disabled = true;
    }

    try {
        const result = await apiRequest(`/admin/usuarios/${encodeURIComponent(usuario.id)}/guest-status`, {
            method: 'POST',
            body: JSON.stringify({ activo: nextActive })
        });

        setFormMessage(
            ui.adminMessage,
            result.message || (nextActive ? 'Cuenta invitado habilitada.' : 'Cuenta invitado deshabilitada.'),
            'success'
        );
        await loadAdminUsers();
    } catch (error) {
        if (checkboxElement) {
            checkboxElement.checked = !nextActive;
        }
        setFormMessage(ui.adminMessage, error.message, 'error');
    } finally {
        if (checkboxElement) {
            checkboxElement.disabled = false;
        }
    }
}

async function handleAdminDeleteUser(usuario) {
    if (!canManageAdminCatalogs(currentUser)) {
        setFormMessage(ui.adminMessage, 'Solo ADMIN o SUPER pueden eliminar usuarios.', 'error');
        return;
    }

    const confirmDelete = window.confirm(`Eliminar usuario ${usuario.username}?`);
    if (!confirmDelete) {
        return;
    }

    try {
        await apiRequest(`/admin/usuarios/${encodeURIComponent(usuario.id)}`, {
            method: 'DELETE'
        });

        if (editingAdminUserId && Number(editingAdminUserId) === Number(usuario.id)) {
            exitAdminEditMode();
        }
        setFormMessage(ui.adminMessage, `Usuario ${usuario.username} eliminado.`, 'success');
        await loadAdminUsers();
    } catch (error) {
        setFormMessage(ui.adminMessage, error.message, 'error');
    }
}

async function handleAdminResetPassword(usuario) {
    if (!canManageAdminCatalogs(currentUser)) {
        setFormMessage(ui.adminMessage, 'Solo ADMIN o SUPER pueden restablecer claves.', 'error');
        return;
    }

    const targetRole = String(usuario.role || '').toUpperCase();
    const isTargetSuper = targetRole === 'SUPER';
    const promptLabel = isTargetSuper
        ? `Nueva clave fija para SUPER ${usuario.username} (minimo 6 caracteres):`
        : `Nueva clave temporal para ${usuario.username} (minimo 6 caracteres):`;
    const temporaryPassword = window.prompt(promptLabel, 'Temporal123!');

    if (temporaryPassword === null) {
        return;
    }

    const newPassword = String(temporaryPassword).trim();
    if (newPassword.length < 6) {
        setFormMessage(ui.adminMessage, 'La clave temporal debe tener al menos 6 caracteres.', 'error');
        return;
    }

    try {
        const result = await apiRequest(`/admin/usuarios/${encodeURIComponent(usuario.id)}/reset-password`, {
            method: 'POST',
            body: JSON.stringify({ newPassword })
        });

        setFormMessage(ui.adminMessage, result.message || `Clave actualizada para ${usuario.username}.`, 'success');
        await loadAdminUsers();
    } catch (error) {
        setFormMessage(ui.adminMessage, error.message, 'error');
    }
}

async function handleGuardar() {
    if (!canCreateRegistros(currentUser)) {
        if (isGuestUser(currentUser)) {
            setFormMessage(ui.formMessage, 'Modo invitado: solo lectura.', 'error');
            return;
        }
        if (isSupervisorUser(currentUser)) {
            setFormMessage(ui.formMessage, 'Tu rol solo puede buscar registros y agregar comentarios.', 'error');
            return;
        }

        setFormMessage(ui.formMessage, 'Tu rol solo puede agregar comentarios y actualizar documentacion recibida.', 'error');
        return;
    }

    if (!canEditRegistros(currentUser)) {
        setFormMessage(ui.formMessage, 'Modo invitado: solo lectura.', 'error');
        return;
    }

    if (loadedIngresoOriginal) {
        showExistingRecordSaveWarningPopup();
        setFormMessage(ui.formMessage, 'Registro ya existe. Si hiciste cambios, usa MODIFICAR.', 'error');
        updateRegistroActionButtons();
        return;
    }

    const data = collectFormData(ui.registroForm);
    if (!data.comentario) {
        data.comentario = DEFAULT_CREATION_COMMENT;
    }

    clearRegistroFieldErrors();
    const missingRequiredOnSave = highlightMissingRegistroRequiredFields(data);
    if (missingRequiredOnSave.length > 0) {
        setFormMessage(ui.formMessage, '', '');
        focusFirstMissingRegistroField(missingRequiredOnSave);
        return;
    }
    clearRegistroFieldErrors();

    const normalizedRegistroData = validateAndNormalizeRutCorreo(data);
    if (normalizedRegistroData.error || !normalizedRegistroData.data) {
        const errorFieldId = normalizedRegistroData.errorField || resolveRegistroFieldIdFromErrorMessage(normalizedRegistroData.error);
        if (errorFieldId) {
            setRegistroFieldError(errorFieldId, normalizedRegistroData.error);
            focusFirstMissingRegistroField([errorFieldId]);
            setFormMessage(ui.formMessage, '', '');
            return;
        }
        setFormMessage(ui.formMessage, normalizedRegistroData.error, 'error');
        return;
    }
    Object.assign(data, normalizedRegistroData.data);
    document.getElementById('rut').value = data.rut;
    document.getElementById('correo').value = data.correo;

    if (!validateNumLotes(data.numLotes)) {
        setRegistroFieldError('numLotes', 'NRO DE LOTES debe ser un numero entero positivo.');
        focusFirstMissingRegistroField(['numLotes']);
        setFormMessage(ui.formMessage, '', '');
        return;
    }
    clearRegistroFieldErrors();

    try {
        const result = await apiRequest('/registros', {
            method: 'POST',
            body: JSON.stringify(data)
        });

        await resetRegistroFormAfterSubmit();
        setFormMessage(ui.formMessage, `Registro ${result.numIngreso || data.numIngreso} guardado correctamente.`, 'success');
    } catch (error) {
        if (applyRegistroInlineErrorMessage(error.message)) {
            setFormMessage(ui.formMessage, '', '');
            return;
        }
        setFormMessage(ui.formMessage, error.message, 'error');
    }
}

async function handleModificar() {
    if (!canEditRegistros(currentUser)) {
        if (isGuestUser(currentUser)) {
            setFormMessage(ui.formMessage, 'Modo invitado: solo lectura.', 'error');
            return;
        }
        if (isSupervisorUser(currentUser)) {
            setFormMessage(ui.formMessage, 'Tu rol solo puede buscar registros y agregar comentarios.', 'error');
            return;
        }
        setFormMessage(ui.formMessage, 'Tu rol no tiene permisos para modificar registros.', 'error');
        return;
    }

    const isOperatorRestricted = isOperatorUser(currentUser) && !isGuestUser(currentUser);
    const data = collectFormData(ui.registroForm);
    const routeIngreso = loadedIngresoOriginal || data.numIngreso;
    if (!routeIngreso) {
        setFormMessage(ui.formMessage, 'Debes cargar un registro existente antes de modificar.', 'error');
        return;
    }

    let payloadToSend = {
        ...data,
        numIngresoNuevo: data.numIngreso
    };
    if (isOperatorRestricted) {
        clearRegistroFieldErrors();
        payloadToSend = {
            numIngresoNuevo: routeIngreso,
            documentos: Array.isArray(data.documentos) ? data.documentos : []
        };
    } else {
        clearRegistroFieldErrors();
        const missingRequiredOnModify = highlightMissingRegistroRequiredFields(data);
        if (missingRequiredOnModify.length > 0) {
            setFormMessage(ui.formMessage, '', '');
            focusFirstMissingRegistroField(missingRequiredOnModify);
            return;
        }
        clearRegistroFieldErrors();

        const normalizedRegistroData = validateAndNormalizeRutCorreo(data);
        if (normalizedRegistroData.error || !normalizedRegistroData.data) {
            const errorFieldId = normalizedRegistroData.errorField || resolveRegistroFieldIdFromErrorMessage(normalizedRegistroData.error);
            if (errorFieldId) {
                setRegistroFieldError(errorFieldId, normalizedRegistroData.error);
                focusFirstMissingRegistroField([errorFieldId]);
                setFormMessage(ui.formMessage, '', '');
                return;
            }
            setFormMessage(ui.formMessage, normalizedRegistroData.error, 'error');
            return;
        }
        Object.assign(data, normalizedRegistroData.data);
        document.getElementById('rut').value = data.rut;
        document.getElementById('correo').value = data.correo;

        if (!validateNumLotes(data.numLotes)) {
            setRegistroFieldError('numLotes', 'NRO DE LOTES debe ser un numero entero positivo.');
            focusFirstMissingRegistroField(['numLotes']);
            setFormMessage(ui.formMessage, '', '');
            return;
        }
        clearRegistroFieldErrors();

        // El comentario manual se guarda solo con el boton "AGREGAR COMENTARIO".
        // En MODIFICAR mantenemos el comentario actual del registro.
        data.comentario = loadedIngresoOriginal ? loadedComentarioOriginal : data.comentario;
        payloadToSend = {
            ...data,
            numIngresoNuevo: data.numIngreso
        };
    }

    try {
        const result = await apiRequest(`/registros/${encodeURIComponent(routeIngreso)}`, {
            method: 'PUT',
            body: JSON.stringify(payloadToSend)
        });

        if (result && result.noChanges) {
            setFormMessage(ui.formMessage, result.message || 'No se detectaron cambios para modificar.', 'success');
            return;
        }

        const finalIngreso = result.numIngreso || data.numIngreso || routeIngreso;
        loadedIngresoOriginal = finalIngreso;
        suggestedIngresoForNewRecord = '';
        if (!isOperatorRestricted) {
            loadedComentarioOriginal = formatComentarioTitleCase(data.comentario || '');
        }
        if (ui.numIngreso) {
            ui.numIngreso.value = finalIngreso;
        }

        await syncCurrentInvoiceRequestStatusByIngreso(finalIngreso);
        await loadHistoryForIngreso(finalIngreso);
        saveLoadedRegistroSnapshot();
        updateRegistroActionButtons();
        if (canUseOperatorDocumentAlerts(currentUser)) {
            void refreshOperatorDocumentAlerts({ silent: true });
        }
        setFormMessage(
            ui.formMessage,
            isOperatorRestricted
                ? `Documentacion del registro ${finalIngreso} actualizada correctamente. La vista permanece cargada hasta que pulses LIMPIAR.`
                : `Registro ${finalIngreso} modificado correctamente. La vista permanece cargada hasta que pulses LIMPIAR.`,
            'success'
        );
    } catch (error) {
        if (applyRegistroInlineErrorMessage(error.message)) {
            setFormMessage(ui.formMessage, '', '');
            return;
        }
        setFormMessage(ui.formMessage, error.message, 'error');
    }
}

async function handleGuardarSoloComentario() {
    if (!canUseQuickCommentSave(currentUser)) {
        setFormMessage(ui.formMessage, 'Modo invitado: solo lectura.', 'error');
        return;
    }

    const routeIngreso = loadedIngresoOriginal || String(ui?.numIngreso?.value || '').trim();
    if (!loadedIngresoOriginal || !routeIngreso) {
        setFormMessage(ui.formMessage, 'Debes cargar un registro existente antes de guardar comentario.', 'error');
        return;
    }

    const comentario = formatComentarioTitleCase(String(ui?.comentario?.value || '')).trim();
    if (!comentario) {
        setFormMessage(ui.formMessage, 'Ingresa un comentario para guardar.', 'error');
        updateQuickCommentButtonState();
        return;
    }

    try {
        const result = await apiRequest(`/registros/${encodeURIComponent(routeIngreso)}/comentario`, {
            method: 'POST',
            body: JSON.stringify({ comentario })
        });

        loadedComentarioOriginal = comentario;
        if (ui.comentario) {
            ui.comentario.value = '';
        }

        await loadHistoryForIngreso(result.numIngreso || routeIngreso);
        setFormMessage(ui.formMessage, result.message || 'Comentario guardado correctamente.', 'success');
        updateRegistroActionButtons();
    } catch (error) {
        setFormMessage(ui.formMessage, error.message, 'error');
    }
}

async function handleEliminarRegistro() {
    if (!canManageAdminCatalogs(currentUser)) {
        setFormMessage(ui.formMessage, 'Solo ADMIN o SUPER pueden eliminar registros.', 'error');
        return;
    }

    const routeIngreso = String(loadedIngresoOriginal || '').trim();
    if (!routeIngreso) {
        setFormMessage(ui.formMessage, 'Debes buscar un registro antes de eliminar.', 'error');
        return;
    }

    const shouldDelete = window.confirm(`Se eliminara definitivamente el registro ${routeIngreso}. Esta accion no se puede deshacer.`);
    if (!shouldDelete) {
        return;
    }

    try {
        const result = await apiRequest(`/registros/${encodeURIComponent(routeIngreso)}`, {
            method: 'DELETE'
        });

        await resetRegistroFormAfterSubmit();
        setFormMessage(ui.formMessage, result.message || `Registro ${routeIngreso} eliminado correctamente.`, 'success');
    } catch (error) {
        setFormMessage(ui.formMessage, error.message, 'error');
    }
}

async function resetRegistroFormAfterSubmit() {
    suppressResetHandler = true;
    ui.registroForm.reset();
    suppressResetHandler = false;

    populateComunas(ui.comuna, '');
    renderHistory([]);
    loadedIngresoOriginal = '';
    clearLoadedRegistroSnapshot();
    loadedRecordUnknownDocuments = [];
    loadedComentarioOriginal = '';
    resetFacturaForm();

    if (canUseHistoricalIngresoMode(currentUser) && adminIngresoMode === 'historical') {
        ui.numIngreso.value = '';
        suggestedIngresoForNewRecord = '';
    } else {
        await loadNextIngreso(ui.numIngreso, ui.formMessage);
    }

    focusNombreInput();
    updateRegistroActionButtons();
}

function focusNombreInput() {
    const nombreInput = document.getElementById('nombre');
    if (!nombreInput) {
        return;
    }

    nombreInput.focus();
    nombreInput.select();
}

async function handleBuscar() {
    if (loadedIngresoOriginal && hasPendingLoadedRegistroChanges()) {
        showPendingModifyRequiredWarningPopup();
        return;
    }

    closeSearchMatchModal();

    const numIngresoInput = document.getElementById('numIngreso').value.trim();
    const nombreInput = document.getElementById('nombre').value.trim();
    const rutInput = document.getElementById('rut').value.trim();
    const rolInput = document.getElementById('rol').value.trim();
    const hasTextCriteria = Boolean(nombreInput || rutInput || rolInput);
    const autoSuggestedIngreso =
        !loadedIngresoOriginal &&
        Boolean(suggestedIngresoForNewRecord) &&
        numIngresoInput === suggestedIngresoForNewRecord;
    const shouldUseNumIngreso = Boolean(numIngresoInput) && (!autoSuggestedIngreso || !hasTextCriteria);

    if (!shouldUseNumIngreso && !hasTextCriteria) {
        setFormMessage(ui.formMessage, 'Ingresa NRO INGRESO, Nombre/Razon social, RUT o ROL para buscar.', 'error');
        return;
    }

    const query = new URLSearchParams();
    if (shouldUseNumIngreso) {
        query.set('numIngreso', numIngresoInput);
    }
    if (nombreInput) {
        query.set('nombre', nombreInput);
    }
    if (rutInput) {
        query.set('rut', rutInput);
    }
    if (rolInput) {
        query.set('rol', rolInput);
    }

    try {
        const match = await apiRequest(`/registros/buscar?${query.toString()}`);
        if (Array.isArray(match.coincidencias) && match.coincidencias.length > 1) {
            const wasModalOpened = openSearchMatchModal(match.coincidencias, match.message);
            if (!wasModalOpened) {
                setFormMessage(ui.formMessage, 'No fue posible mostrar las coincidencias de busqueda.', 'error');
            }
            return;
        }

        await loadRegistroByIngreso(match.numIngreso || '', 'base de datos');
    } catch (error) {
        setFormMessage(ui.formMessage, error.message, 'error');
    }
}

async function loadHistoryForIngreso(numIngreso) {
    if (!numIngreso) {
        renderHistory([]);
        return;
    }

    const data = await apiRequest(`/registros/${encodeURIComponent(numIngreso)}/historial`);
    renderHistory(data.historial || []);
}

function createApiError(message, details = {}) {
    const error = new Error(String(message || 'Ocurrio un error al procesar la solicitud.'));
    if (details && typeof details === 'object') {
        Object.entries(details).forEach(([key, value]) => {
            error[key] = value;
        });
    }

    return error;
}

function isAuthFailureError(error) {
    return Number(error?.status) === 401;
}

function isWriteApiKeyAuthFailure(status, code, message) {
    if (Number(status) !== 401) {
        return false;
    }

    const normalizedCode = String(code || '').trim().toUpperCase();
    if (normalizedCode === 'WRITE_API_KEY_REQUIRED') {
        return true;
    }

    const normalizedMessage = String(message || '').toLowerCase();
    return normalizedMessage.includes('operaciones de escritura') || normalizedMessage.includes('x-api-key');
}

async function apiRequest(path, options = {}) {
    const { skipAuthRedirect = false, skipPasswordRedirect = false, ...fetchOptions } = options;

    const headers = {
        'Content-Type': 'application/json',
        ...(fetchOptions.headers || {})
    };

    const apiKey = getApiKey();
    if (apiKey) {
        headers['X-API-Key'] = apiKey;
    }

    const authToken = getAuthToken();
    if (authToken) {
        headers.Authorization = `Bearer ${authToken}`;
    }

    const config = {
        method: fetchOptions.method || 'GET',
        ...fetchOptions,
        credentials: 'include',
        headers
    };

    const apiBasesToTry = [API_BASE, ...API_BASE_FALLBACKS.filter((base) => base !== API_BASE)];
    let response;

    for (const base of apiBasesToTry) {
        try {
            response = await fetch(`${base}${path}`, config);
            API_BASE = base;
            break;
        } catch (error) {
            // Se prueba con el siguiente endpoint candidato.
        }
    }

    if (!response) {
        const triedBases = apiBasesToTry.join(', ');
        throw createApiError(`No hay conexion con el servidor API. Se intento con: ${triedBases}.`, {
            status: 0,
            code: 'NETWORK_ERROR',
            isNetworkError: true
        });
    }

    const payload = await response.json().catch(() => ({}));

    if (!response.ok) {
        const message = payload.message || 'Ocurrio un error de conexion con el servidor.';
        const responseCode = payload.code || '';
        const isWriteAuthFailure = isWriteApiKeyAuthFailure(response.status, responseCode, message);
        if (response.status === 403 && payload.code === 'PASSWORD_CHANGE_REQUIRED' && !skipPasswordRedirect) {
            if (currentUser) {
                currentUser.mustChangePassword = true;
            }
            showPasswordChangeCard(currentUser, message);
        }

        if (response.status === 401 && !skipAuthRedirect && !isWriteAuthFailure) {
            handleSessionExpired(message);
        }

        throw createApiError(message, {
            status: response.status,
            code: responseCode,
            payload,
            isNetworkError: false
        });
    }

    return payload;
}

async function loadNextIngreso(numIngresoInput, messageElement) {
    try {
        const data = await apiRequest('/next-ingreso');
        const nextIngreso = String(data.numIngreso || '').trim();
        numIngresoInput.value = nextIngreso;
        suggestedIngresoForNewRecord = nextIngreso;
        loadedIngresoOriginal = '';
        clearLoadedRegistroSnapshot();
        loadedRecordUnknownDocuments = [];
        loadedComentarioOriginal = '';
        updateRegistroActionButtons();
    } catch (error) {
        numIngresoInput.value = '';
        suggestedIngresoForNewRecord = '';
        setFormMessage(
            messageElement,
            'No fue posible obtener el siguiente NRO INGRESO. Verifica que el servidor Node este activo.',
            'error'
        );
    }
}

function formatComentarioTitleCase(value) {
    return String(value || '')
        .toLocaleLowerCase('es-CL')
        .replace(/(^|[\s]+)([a-zA-Z\u00C0-\u00FF])/g, (match, prefix, letter) => {
            return `${prefix}${letter.toLocaleUpperCase('es-CL')}`;
        });
}

function enforceComentarioTitleCase(element) {
    if (!element) {
        return;
    }

    const currentValue = String(element.value || '');
    const formattedValue = formatComentarioTitleCase(currentValue);
    if (formattedValue === currentValue) {
        return;
    }

    const cursorStart = element.selectionStart;
    const cursorEnd = element.selectionEnd;
    element.value = formattedValue;

    if (typeof cursorStart === 'number' && typeof cursorEnd === 'number') {
        element.setSelectionRange(cursorStart, cursorEnd);
    }
}

function formatNombreRazonSocialTitleCase(value) {
    return String(value || '')
        .toLocaleLowerCase('es-CL')
        .replace(/(^|[\s\-\/]+)([a-zA-Z\u00C0-\u00FF])/g, (match, prefix, letter) => {
            return `${prefix}${letter.toLocaleUpperCase('es-CL')}`;
        });
}

function enforceNombreTitleCase(element) {
    if (!element) {
        return;
    }

    const currentValue = String(element.value || '');
    const formattedValue = formatNombreRazonSocialTitleCase(currentValue);
    if (formattedValue === currentValue) {
        return;
    }

    const cursorStart = element.selectionStart;
    const cursorEnd = element.selectionEnd;
    element.value = formattedValue;

    if (typeof cursorStart === 'number' && typeof cursorEnd === 'number') {
        element.setSelectionRange(cursorStart, cursorEnd);
    }
}

function collectFormData(form) {
    const selectedDocuments = Array.from(form.querySelectorAll('input[name="documentos"]:checked'))
        .map((input) => String(input.value || '').trim())
        .filter((value) => value.length > 0);
    const retainedUnknownDocs = loadedIngresoOriginal ? loadedRecordUnknownDocuments : [];
    const documentos = [...new Set([...selectedDocuments, ...retainedUnknownDocs])];

    const comentarioIngresado = formatComentarioTitleCase(document.getElementById('comentario').value).trim();
    const comentarioFinal = comentarioIngresado || (loadedIngresoOriginal ? loadedComentarioOriginal : '');

    return {
        numIngreso: document.getElementById('numIngreso').value.trim(),
        nombre: formatNombreRazonSocialTitleCase(document.getElementById('nombre').value).trim(),
        rut: document.getElementById('rut').value.trim(),
        telefono: document.getElementById('telefono').value.trim(),
        correo: document.getElementById('correo').value.trim(),
        region: document.getElementById('region').value,
        comuna: document.getElementById('comuna').value,
        nombrePredio: document.getElementById('nombrePredio').value.trim(),
        rol: document.getElementById('rol').value.trim(),
        numLotes: document.getElementById('numLotes').value,
        estado: document.getElementById('estado').value,
        comentario: comentarioFinal,
        documentos,
        factura: collectFacturaSnapshotData()
    };
}

function fillForm(form, data, comunaSelect) {
    clearRegistroFieldErrors();
    const parsedRut = parseChileanRut(data.rut || '');
    document.getElementById('numIngreso').value = data.numIngreso || '';
    document.getElementById('nombre').value = formatNombreRazonSocialTitleCase(data.nombre || '');
    document.getElementById('rut').value = parsedRut.isValid ? parsedRut.formatted : String(data.rut || '').trim();
    document.getElementById('telefono').value = data.telefono || '';
    document.getElementById('correo').value = normalizeEmailValue(data.correo || '');
    document.getElementById('region').value = data.region || '';

    populateComunas(comunaSelect, data.region || '', data.comuna || '');

    document.getElementById('nombrePredio').value = data.nombrePredio || '';
    document.getElementById('rol').value = data.rol || '';
    document.getElementById('numLotes').value = data.numLotes ?? '';
    document.getElementById('estado').value = normalizeEstadoValue(data.estado) || 'no_facturada';
    loadedComentarioOriginal = isGuestUser(currentUser) ? '' : formatComentarioTitleCase(data.comentario || '');
    document.getElementById('comentario').value = '';

    const incomingDocs = Array.isArray(data.documentos)
        ? data.documentos.map((item) => String(item || '').trim()).filter((item) => item.length > 0)
        : [];
    const visibleValues = new Set(configuredDocumentOptions.map((item) => item.value));
    const selectedVisibleDocs = incomingDocs.filter((doc) => visibleValues.has(doc));
    loadedRecordUnknownDocuments = incomingDocs.filter((doc) => !visibleValues.has(doc));
    renderDocumentOptions(configuredDocumentOptions, selectedVisibleDocs);
    resetFacturaForm();
    applyFacturaSnapshotToForm(data.factura);
    syncFacturaNumeroVisibility({ clearOnHide: false });
}

function canEditHistoryComments(user) {
    return canManageAdminCatalogs(user);
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
        if (
            !Number.isNaN(localDate.getTime()) &&
            localDate.getFullYear() === year &&
            localDate.getMonth() + 1 === month &&
            localDate.getDate() === day &&
            localDate.getHours() === hour &&
            localDate.getMinutes() === minute &&
            localDate.getSeconds() === second
        ) {
            return localDate;
        }
    }

    const parsed = new Date(text);
    if (Number.isNaN(parsed.getTime())) {
        return null;
    }

    return parsed;
}

function toLocalDateTimeInputValue(value) {
    const parsed = parseDateValue(value);
    if (!parsed) {
        return '';
    }

    const timezoneOffsetMs = parsed.getTimezoneOffset() * 60000;
    return new Date(parsed.getTime() - timezoneOffsetMs).toISOString().slice(0, 16);
}

function closeHistoryEditModal() {
    if (ui?.historyEditModalOverlay) {
        ui.historyEditModalOverlay.classList.add('hidden');
    }
    if (ui?.historyEditForm && typeof ui.historyEditForm.reset === 'function') {
        ui.historyEditForm.reset();
    }
    setFormMessage(ui?.historyEditMessage, '', '');
    editingHistoryEntryId = null;
    editingHistoryIngreso = '';
    editingHistoryAction = '';
    if (ui?.historyEditDeleteBtn) {
        ui.historyEditDeleteBtn.classList.add('hidden');
    }
}

function openHistoryEditModal(historyItem) {
    if (!canEditHistoryComments(currentUser)) {
        setFormMessage(ui?.formMessage, 'Solo ADMIN o SUPER pueden editar comentarios del historial de progreso.', 'error');
        return;
    }

    const historyId = Number(historyItem?.id);
    const targetIngreso = String(historyItem?.numIngreso || loadedIngresoOriginal || '').trim();
    const actionLabel = String(historyItem?.accion || '').trim().toUpperCase();
    if (!Number.isInteger(historyId) || historyId <= 0 || !targetIngreso) {
        setFormMessage(ui?.formMessage, 'No fue posible abrir este comentario para edicion.', 'error');
        return;
    }
    if (!ui?.historyEditModalOverlay || !ui?.historyEditFecha || !ui?.historyEditComentario) {
        return;
    }

    const localDateTime = toLocalDateTimeInputValue(historyItem?.fecha);
    if (!localDateTime) {
        setFormMessage(ui?.formMessage, 'No fue posible interpretar la fecha de este comentario.', 'error');
        return;
    }

    editingHistoryEntryId = historyId;
    editingHistoryIngreso = targetIngreso;
    editingHistoryAction = actionLabel;
    ui.historyEditFecha.value = localDateTime;
    ui.historyEditComentario.value = formatComentarioTitleCase(String(historyItem?.comentario || '')).trim();
    if (ui.historyEditDeleteBtn) {
        const showDeleteBtn = isSuperUser(currentUser);
        ui.historyEditDeleteBtn.classList.toggle('hidden', !showDeleteBtn);
    }
    setFormMessage(ui?.historyEditMessage, '', '');
    ui.historyEditModalOverlay.classList.remove('hidden');
    window.requestAnimationFrame(() => {
        ui.historyEditComentario.focus();
        ui.historyEditComentario.select();
    });
}

async function handleHistoryEditSubmit(event) {
    event.preventDefault();

    if (!canEditHistoryComments(currentUser)) {
        setFormMessage(ui?.historyEditMessage, 'Solo ADMIN o SUPER pueden editar comentarios del historial de progreso.', 'error');
        return;
    }

    const historyId = Number(editingHistoryEntryId);
    const targetIngreso = String(editingHistoryIngreso || loadedIngresoOriginal || '').trim();
    if (!Number.isInteger(historyId) || historyId <= 0 || !targetIngreso) {
        setFormMessage(ui?.historyEditMessage, 'No se encontro el comentario seleccionado para editar.', 'error');
        return;
    }

    const fechaRaw = String(ui?.historyEditFecha?.value || '').trim();
    const comentario = formatComentarioTitleCase(String(ui?.historyEditComentario?.value || '')).trim();

    if (!fechaRaw) {
        setFormMessage(ui?.historyEditMessage, 'Debes ingresar fecha y hora del comentario.', 'error');
        return;
    }
    if (!comentario) {
        setFormMessage(ui?.historyEditMessage, 'Debes ingresar el comentario.', 'error');
        return;
    }

    try {
        await apiRequest(`/registros/${encodeURIComponent(targetIngreso)}/historial/${encodeURIComponent(historyId)}`, {
            method: 'PUT',
            body: JSON.stringify({
                fecha: fechaRaw,
                comentario
            })
        });

        try {
            await loadHistoryForIngreso(targetIngreso);
            closeHistoryEditModal();
            setFormMessage(ui?.formMessage, 'Comentario de historial de progreso actualizado correctamente.', 'success');
        } catch (refreshError) {
            closeHistoryEditModal();
            setFormMessage(
                ui?.formMessage,
                `Comentario actualizado, pero no se pudo recargar el historial de progreso: ${refreshError.message}`,
                'error'
            );
        }
    } catch (error) {
        setFormMessage(ui?.historyEditMessage, error.message, 'error');
    }
}

async function handleHistoryDeleteClick() {
    if (!isSuperUser(currentUser)) {
        setFormMessage(ui?.historyEditMessage, 'Solo SUPER puede eliminar comentarios del historial.', 'error');
        return;
    }

    const historyId = Number(editingHistoryEntryId);
    const targetIngreso = String(editingHistoryIngreso || loadedIngresoOriginal || '').trim();
    const actionLabel = String(editingHistoryAction || '').trim().toUpperCase();
    if (!Number.isInteger(historyId) || historyId <= 0 || !targetIngreso) {
        setFormMessage(ui?.historyEditMessage, 'No se encontro el comentario seleccionado para eliminar.', 'error');
        return;
    }
    const label = actionLabel || 'ENTRADA';
    const confirmDelete = window.confirm(`Se eliminara esta entrada del historial (${label}). Esta accion no se puede deshacer. Deseas continuar?`);
    if (!confirmDelete) {
        return;
    }

    try {
        await apiRequest(`/registros/${encodeURIComponent(targetIngreso)}/historial/${encodeURIComponent(historyId)}`, {
            method: 'DELETE'
        });

        try {
            await loadHistoryForIngreso(targetIngreso);
            closeHistoryEditModal();
            setFormMessage(ui?.formMessage, 'Comentario eliminado correctamente del historial.', 'success');
        } catch (refreshError) {
            closeHistoryEditModal();
            setFormMessage(
                ui?.formMessage,
                `Comentario eliminado, pero no se pudo recargar el historial: ${refreshError.message}`,
                'error'
            );
        }
    } catch (error) {
        setFormMessage(ui?.historyEditMessage, error.message, 'error');
    }
}

function renderHistory(historial) {
    if (!ui || !ui.historialBody) {
        return;
    }

    const canManageHistory = canEditHistoryComments(currentUser);
    if (ui.historialActionHeader) {
        ui.historialActionHeader.classList.toggle('hidden', !canManageHistory);
    }

    ui.historialBody.innerHTML = '';

    if (!Array.isArray(historial) || historial.length === 0) {
        const emptyCols = canManageHistory ? 4 : 3;
        ui.historialBody.innerHTML = `<tr><td colspan="${emptyCols}">Sin historial de progreso para mostrar.</td></tr>`;
        return;
    }

    const sortedHistory = [...historial].sort((left, right) => {
        const leftTime = parseDateValue(left?.fecha)?.getTime() || 0;
        const rightTime = parseDateValue(right?.fecha)?.getTime() || 0;
        if (rightTime !== leftTime) {
            return rightTime - leftTime;
        }

        const leftPriority = String(left?.accion || '').trim().toUpperCase() === 'COMENTARIO' ? 0 : 1;
        const rightPriority = String(right?.accion || '').trim().toUpperCase() === 'COMENTARIO' ? 0 : 1;
        if (leftPriority !== rightPriority) {
            return leftPriority - rightPriority;
        }

        const leftId = Number(left?.id || 0);
        const rightId = Number(right?.id || 0);
        return rightId - leftId;
    });

    sortedHistory.forEach((item) => {
        const row = document.createElement('tr');

        const fechaCell = document.createElement('td');
        fechaCell.textContent = formatDateTime(item.fecha);

        const comentarioCell = document.createElement('td');
        comentarioCell.textContent = item.comentario || '';

        const usuarioCell = document.createElement('td');
        usuarioCell.textContent = item.usuario || '';

        row.appendChild(fechaCell);
        row.appendChild(comentarioCell);
        row.appendChild(usuarioCell);
        if (canManageHistory) {
            const actionCell = document.createElement('td');
            const canEditEntry = Number.isInteger(Number(item?.id)) && Number(item?.id) > 0;
            if (canEditEntry) {
                const editBtn = document.createElement('button');
                editBtn.type = 'button';
                editBtn.className = 'admin-reset-btn';
                editBtn.textContent = 'EDITAR';
                editBtn.addEventListener('click', () => openHistoryEditModal(item));
                actionCell.appendChild(editBtn);
            } else {
                actionCell.textContent = '-';
            }
            row.appendChild(actionCell);
        }
        ui.historialBody.appendChild(row);
    });
}

function formatDateTime(value) {
    const parsed = parseDateValue(value);
    if (!parsed) {
        return String(value || '');
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

function normalizeText(value) {
    return String(value ?? '').trim();
}

function normalizeEstadoValue(value) {
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

function setRegistroFieldRequiredState(fieldId, isMissing) {
    const field = document.getElementById(fieldId);
    if (!field) {
        return;
    }

    const formGroup = field.closest('.form-group');
    field.classList.toggle('required-missing', isMissing);
    field.setAttribute('aria-invalid', isMissing ? 'true' : 'false');
    if (formGroup) {
        formGroup.classList.toggle('required-missing', isMissing);
    }
}

function getRegistroFieldErrorElement(fieldId) {
    const field = document.getElementById(fieldId);
    if (!field) {
        return null;
    }

    const formGroup = field.closest('.form-group');
    if (!formGroup) {
        return null;
    }

    let messageElement = formGroup.querySelector(`.field-error-message[data-field-error-for="${fieldId}"]`);
    if (!messageElement) {
        messageElement = document.createElement('small');
        messageElement.className = 'field-error-message hidden';
        messageElement.dataset.fieldErrorFor = fieldId;
        formGroup.appendChild(messageElement);
    }

    return messageElement;
}

function setRegistroFieldError(fieldId, message) {
    const normalizedMessage = String(message || '').trim();
    const hasError = normalizedMessage.length > 0;
    setRegistroFieldRequiredState(fieldId, hasError);

    const messageElement = getRegistroFieldErrorElement(fieldId);
    if (!messageElement) {
        return;
    }

    messageElement.textContent = normalizedMessage;
    messageElement.classList.toggle('hidden', !hasError);
}

function clearRegistroFieldError(fieldId) {
    setRegistroFieldError(fieldId, '');
}

function clearRegistroFieldErrors() {
    REGISTRO_INLINE_ERROR_FIELD_IDS.forEach((fieldId) => {
        clearRegistroFieldError(fieldId);
    });
}

function clearRegistroRequiredFieldMarks() {
    REGISTRO_REQUIRED_FIELD_IDS.forEach((fieldId) => {
        setRegistroFieldRequiredState(fieldId, false);
    });
}

function getMissingRegistroRequiredFieldIds(data) {
    if (!data) {
        return [...REGISTRO_REQUIRED_FIELD_IDS];
    }

    return REGISTRO_REQUIRED_FIELD_IDS.filter((fieldId) => !String(data[fieldId] || '').trim());
}

function highlightMissingRegistroRequiredFields(data) {
    const missingFieldIds = getMissingRegistroRequiredFieldIds(data);
    REGISTRO_REQUIRED_FIELD_IDS.forEach((fieldId) => {
        if (missingFieldIds.includes(fieldId)) {
            setRegistroFieldError(fieldId, 'Campo obligatorio.');
        } else {
            clearRegistroFieldError(fieldId);
        }
    });
    return missingFieldIds;
}

function validateRequired(data) {
    return getMissingRegistroRequiredFieldIds(data).length === 0;
}

function focusFirstMissingRegistroField(missingFieldIds) {
    if (!Array.isArray(missingFieldIds) || missingFieldIds.length === 0) {
        return;
    }

    const firstMissingField = document.getElementById(missingFieldIds[0]);
    if (firstMissingField && typeof firstMissingField.focus === 'function') {
        firstMissingField.focus();
    }
}

function registerRegistroRequiredFieldValidationListeners() {
    REGISTRO_INLINE_ERROR_FIELD_IDS.forEach((fieldId) => {
        const field = document.getElementById(fieldId);
        if (!field) {
            return;
        }

        const syncFieldErrorState = () => {
            const hasValue = String(field.value || '').trim().length > 0;
            if (hasValue) {
                clearRegistroFieldError(fieldId);
                return;
            }

            if (REGISTRO_REQUIRED_FIELD_IDS.includes(fieldId)) {
                setRegistroFieldError(fieldId, 'Campo obligatorio.');
                return;
            }

            clearRegistroFieldError(fieldId);
        };

        field.addEventListener('input', syncFieldErrorState);
        field.addEventListener('change', syncFieldErrorState);
        field.addEventListener('blur', syncFieldErrorState);
    });
}

function resolveRegistroFieldIdFromErrorMessage(message) {
    const normalizedMessage = String(message || '').trim().toLowerCase();
    if (!normalizedMessage) {
        return '';
    }

    if (normalizedMessage.includes('rut invalido')) {
        return 'rut';
    }
    if (normalizedMessage.includes('correo invalido')) {
        return 'correo';
    }
    if (normalizedMessage.includes('nro de lotes')) {
        return 'numLotes';
    }

    return '';
}

function applyRegistroInlineErrorMessage(message) {
    const fieldId = resolveRegistroFieldIdFromErrorMessage(message);
    if (!fieldId) {
        return false;
    }

    setRegistroFieldError(fieldId, String(message || '').trim());
    focusFirstMissingRegistroField([fieldId]);
    return true;
}

function validateNumLotes(value) {
    if (value === '' || value === null || typeof value === 'undefined') {
        return true;
    }

    const parsed = Number(value);
    return Number.isInteger(parsed) && parsed > 0;
}

function resolveApiBase() {
    const override = sanitizeApiBaseOverride(getApiBaseOverride());
    if (override) {
        return override;
    }

    const metaValue = normalizeApiBase(document.querySelector('meta[name=\"geo-rural-api-base\"]')?.content || '');
    if (metaValue) {
        return metaValue;
    }

    const host = String(window.location.hostname || '').trim().toLowerCase();
    const port = String(window.location.port || '').trim();
    const isLocalHost = host === 'localhost' || host === '127.0.0.1';

    if (isLocalHost && port !== '3000') {
        return `http://${host}:3000/api`;
    }

    // En internet (Nginx), la API debe resolverse por la misma origin en /api.
    return '/api';
}

function resolveApiFallbackBases(primaryBase) {
    const candidates = [];

    const appendCandidate = (value) => {
        const normalized = normalizeApiBase(value);
        if (!normalized || normalized === primaryBase || candidates.includes(normalized)) {
            return;
        }
        candidates.push(normalized);
    };

    const host = String(window.location.hostname || '').trim().toLowerCase();
    const isLocalHost = host === 'localhost' || host === '127.0.0.1';

    if (isLocalHost) {
        appendCandidate(`http://${host}:3000/api`);
        if (host === 'localhost') {
            appendCandidate('http://127.0.0.1:3000/api');
        } else if (host === '127.0.0.1') {
            appendCandidate('http://localhost:3000/api');
        }
    }

    appendCandidate('/api');
    return candidates;
}

function getApiBaseOverride() {
    try {
        return normalizeApiBase(window.localStorage.getItem(API_BASE_STORAGE) || '');
    } catch (error) {
        return '';
    }
}

function isLoopbackHost(hostname) {
    const host = String(hostname || '').trim().toLowerCase();
    return host === 'localhost' || host === '127.0.0.1' || host === '::1' || host === '[::1]';
}

function sanitizeApiBaseOverride(overrideValue) {
    const normalized = normalizeApiBase(overrideValue);
    if (!normalized) {
        return '';
    }

    const pageHost = String(window.location.hostname || '').trim().toLowerCase();
    const pageIsLoopback = isLoopbackHost(pageHost);

    try {
        const parsed = new URL(normalized, window.location.origin);
        const overrideHost = String(parsed.hostname || '').trim().toLowerCase();
        const overrideIsLoopback = isLoopbackHost(overrideHost);
        const overridePort = String(parsed.port || '').trim();
        const pagePort = String(window.location.port || '').trim();

        // Si estamos en internet (host publico), nunca usar override local 127.0.0.1/localhost.
        if (!pageIsLoopback && overrideIsLoopback) {
            try {
                window.localStorage.removeItem(API_BASE_STORAGE);
            } catch (storageError) {
                // Ignorar errores de storage.
            }
            return '';
        }

        // Si estamos en internet por 80/443, evitar override duro hacia :3000 del mismo host.
        if (!pageIsLoopback && pagePort !== '3000' && overrideHost === pageHost && overridePort === '3000') {
            try {
                window.localStorage.removeItem(API_BASE_STORAGE);
            } catch (storageError) {
                // Ignorar errores de storage.
            }
            return '';
        }
    } catch (error) {
        // Si no se puede parsear como URL, conservar valor normalizado.
    }

    return normalized;
}

function normalizeApiBase(value) {
    const cleanValue = String(value || '').trim();
    if (!cleanValue) {
        return '';
    }

    return cleanValue.replace(/\/+$/, '');
}

function getApiKey() {
    if (apiKeyMemory) {
        return apiKeyMemory;
    }

    try {
        const sessionValue = (window.sessionStorage.getItem(API_KEY_STORAGE) || '').trim();
        if (sessionValue) {
            apiKeyMemory = sessionValue;
            return apiKeyMemory;
        }

        const legacyValue = (window.localStorage.getItem(API_KEY_STORAGE) || '').trim();
        if (legacyValue) {
            window.sessionStorage.setItem(API_KEY_STORAGE, legacyValue);
            window.localStorage.removeItem(API_KEY_STORAGE);
            apiKeyMemory = legacyValue;
            return apiKeyMemory;
        }

        return '';
    } catch (error) {
        return apiKeyMemory;
    }
}

function setApiKey(value) {
    apiKeyMemory = String(value || '').trim();
    try {
        if (apiKeyMemory) {
            window.sessionStorage.setItem(API_KEY_STORAGE, apiKeyMemory);
        } else {
            window.sessionStorage.removeItem(API_KEY_STORAGE);
            window.localStorage.removeItem(API_KEY_STORAGE);
        }
    } catch (error) {
        return '';
    }
}

function getAuthToken() {
    if (authTokenMemory) {
        return authTokenMemory;
    }

    try {
        const sessionValue = (window.sessionStorage.getItem(AUTH_TOKEN_STORAGE) || '').trim();
        if (sessionValue) {
            authTokenMemory = sessionValue;
            return authTokenMemory;
        }

        const legacyValue = (window.localStorage.getItem(AUTH_TOKEN_STORAGE) || '').trim();
        if (legacyValue) {
            window.sessionStorage.setItem(AUTH_TOKEN_STORAGE, legacyValue);
            window.localStorage.removeItem(AUTH_TOKEN_STORAGE);
            authTokenMemory = legacyValue;
            return authTokenMemory;
        }
    } catch (error) {
        // Si el almacenamiento no esta disponible, se mantiene solo en memoria.
    }

    return '';
}

function setAuthToken(token) {
    authTokenMemory = String(token || '').trim();
    try {
        if (authTokenMemory) {
            window.sessionStorage.setItem(AUTH_TOKEN_STORAGE, authTokenMemory);
        } else {
            window.sessionStorage.removeItem(AUTH_TOKEN_STORAGE);
            window.localStorage.removeItem(AUTH_TOKEN_STORAGE);
        }
    } catch (error) {
        // Si sessionStorage no esta disponible, se mantiene solo en memoria.
    }
}

function clearAuthSession() {
    authTokenMemory = '';
    apiKeyMemory = '';
    try {
        window.sessionStorage.removeItem(AUTH_TOKEN_STORAGE);
        window.localStorage.removeItem(AUTH_TOKEN_STORAGE);
        window.sessionStorage.removeItem(API_KEY_STORAGE);
        window.localStorage.removeItem(API_KEY_STORAGE);
    } catch (error) {
        // Ignorar errores de almacenamiento local.
    }
}

function escapeHtml(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function setFormMessage(element, text, type) {
    if (!element) {
        return;
    }

    element.textContent = text;
    element.className = 'form-message';

    if (type) {
        element.classList.add(type);
    }
}

function populateRegiones(regionSelect, emptyLabel = 'Seleccione region') {
    if (!regionSelect) {
        return;
    }

    const previousValue = String(regionSelect.value || '').trim();
    regionSelect.innerHTML = '';
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = String(emptyLabel || 'Seleccione region');
    regionSelect.appendChild(defaultOption);
    Object.keys(REGION_COMUNAS).forEach((region) => {
        const option = document.createElement('option');
        option.value = region;
        option.textContent = region;
        regionSelect.appendChild(option);
    });

    setSelectValueWithFallback(regionSelect, previousValue);
}

function populateComunas(comunaSelect, region, selectedComuna = '', emptyLabel = 'Seleccione comuna') {
    if (!comunaSelect) {
        return;
    }

    comunaSelect.innerHTML = '';
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = String(emptyLabel || 'Seleccione comuna');
    comunaSelect.appendChild(defaultOption);

    if (!region || !REGION_COMUNAS[region]) {
        comunaSelect.disabled = true;
        return;
    }

    REGION_COMUNAS[region].forEach((comuna) => {
        const option = document.createElement('option');
        option.value = comuna;
        option.textContent = comuna;
        comunaSelect.appendChild(option);
    });

    comunaSelect.disabled = false;
    if (selectedComuna) {
        comunaSelect.value = selectedComuna;
    }
}

function setSelectValueWithFallback(selectElement, value) {
    if (!selectElement) {
        return;
    }

    const normalizedValue = String(value || '').trim();
    if (!normalizedValue) {
        selectElement.value = '';
        return;
    }

    selectElement.value = normalizedValue;
    if (selectElement.value === normalizedValue) {
        return;
    }

    const customOption = document.createElement('option');
    customOption.value = normalizedValue;
    customOption.textContent = normalizedValue;
    selectElement.appendChild(customOption);
    selectElement.value = normalizedValue;
}
