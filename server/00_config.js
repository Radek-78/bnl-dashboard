/**
 * Centrální konfigurace šablony.
 *
 * Hodnoty specifické pro konkrétní projekt (název, podtitul) se po inicializaci
 * čtou z listu _settings v databázi — zde jsou jen výchozí hodnoty a konstanty,
 * které se mezi projekty nemění.
 */
/**
 * Logo Lidl jako SVG přímo v kódu (stejné jako v Planung Dashboardu). Dřív se bralo
 * z Disku (drive.google.com/thumbnail) ze souboru sdíleného jen s vlastníkem - ostatní
 * uživatelé ho nemuseli vidět. Do stránky jde jako vložený obrázek (data URI, viz
 * CONFIG.logoUrl), takže všechna místa s <img src="logoUrl"> zůstávají beze změny.
 */
const LOGO_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 115 115"><path fill="#0050aa" d="M1 1h113v113H1z"/><path fill="#fff" d="M114 1v113H1V1h113m1-1H0v115h115V0z"/><path fill="#fff000" fill-rule="evenodd" d="M57.5 7.38a50.12 50.12 0 1 0 50.12 50.14A50.14 50.14 0 0 0 57.5 7.38z"/><path fill="#e60a14" fill-rule="evenodd" d="M54.39 58.91l-8.85-8.85-10.2 10.22v3.43l2.57-2.58 7.12 7.14-2.63 2.62 1.71 1.72 14.24-14.26v-3.42l-3.96 3.98z"/><path fill="#0050aa" fill-rule="evenodd" d="M13.08 48.2h15.76v3.4h-2.63v11l9.13-5.08v9.31H13.08v-3.42h2.64V51.6h-2.64v-3.4zM79.53 48.2v3.4h2.64v11.81h-2.64v3.42h22.28v-9.31l-9.14 5.08v-11h2.64v-3.4H79.53z"/><path fill="#e60a14" fill-rule="evenodd" d="M44.24 37.61a5.61 5.61 0 1 1-5.61 5.59 5.59 5.59 0 0 1 5.61-5.59z"/><path fill="#e60a14" fill-rule="evenodd" d="M57.5 4A53.51 53.51 0 1 0 111 57.52 53.53 53.53 0 0 0 57.5 4zm0 103.63a50.12 50.12 0 1 1 50.12-50.1 50.13 50.13 0 0 1-50.12 50.09z"/><path fill="#0050aa" fill-rule="evenodd" d="M70.75 48.2h-15v3.4h2.63v11.81h-2.66v3.42h15c11.14 0 11.28-18.63.03-18.63z"/><path fill="#fff000" fill-rule="evenodd" d="M68.64 61h-.75v-7h.63c3.29 0 3.29 7 .12 7z"/></svg>';

const CONFIG = {
  defaultAppName: 'Výchozí aplikace',
  defaultAppSubtitle: '',
  version: 'v3.1.169',
  releaseDate: '9.10.2026',
  logoUrl: 'data:image/svg+xml;base64,' + Utilities.base64Encode(LOGO_SVG),
  defaultSyncFolderUrl: 'https://drive.google.com/drive/folders/1DX1VFWt5fAztALgMdVGJvLwQEjm-t83B?lfhs=2',
  // Stejný font jako ve stylesheetu appky (--font v ui/styles.html) - viz applySheetFont_ v 10_util.js.
  sheetFont: 'Lidl Font Pro',
  theme: {
    blue: '#0050aa',
    darkBlue: '#002466',
    lightBlue: '#008cd2',
    yellow: '#fff000',
    red: '#e60a14',
    white: '#ffffff',
    black: '#000000',
  },
};

/** Klíče ve Script Properties. */
const PROPS = {
  DB_ID: 'DB_SPREADSHEET_ID',
  SETUP_AT: 'SETUP_COMPLETED_AT',
};

/** Role a jejich hierarchie. Vyšší číslo = vyšší oprávnění. */
const ROLES = {
  SUPERADMIN: 'SUPERADMIN',
  ADMIN: 'ADMIN',
  USER: 'USER',
};

const ROLE_LEVEL = {
  SUPERADMIN: 3,
  ADMIN: 2,
  USER: 1,
};

/** Systémové listy v DB spreadsheetu. */
const SHEETS = {
  USERS: '_users',
  SETTINGS: '_settings',
  AUDIT: '_audit_log',
  STORES: 'stores',
  LOGISTICS: 'logistics',
  APPS: 'apps',
  ROLE_PERMISSIONS: '_role_permissions',
};
