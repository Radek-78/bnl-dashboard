/**
 * Centrální konfigurace šablony.
 *
 * Hodnoty specifické pro konkrétní projekt (název, podtitul) se po inicializaci
 * čtou z listu _settings v databázi — zde jsou jen výchozí hodnoty a konstanty,
 * které se mezi projekty nemění.
 */
const CONFIG = {
  defaultAppName: 'Výchozí aplikace',
  defaultAppSubtitle: '',
  version: 'v3.1.92',
  releaseDate: '20.7.2026',
  logoUrl: 'https://drive.google.com/thumbnail?id=18mu_Lq1F_FqqSZcolMjLwG0aaQDPMdyD&sz=w320',
  defaultSyncFolderUrl: 'https://drive.google.com/drive/folders/1DX1VFWt5fAztALgMdVGJvLwQEjm-t83B?lfhs=2',
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
