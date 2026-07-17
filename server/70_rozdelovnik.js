/**
 * Subaplikace "Rozdělovník 20 artiklů" (slug rozdelovnik-20-artiklu).
 *
 * Vlastní DB (apps.db_spreadsheet_id, viz 25_subapp_db.js) odděleně od hlavní
 * databáze appky. Záložky Artikly/_settings mají pevné schéma (RZ_SCHEMA) přes
 * createDbRepo_. Záložky Odprodej/Teo_stavy/Vyskladňovací_listy/Přidělení_po_artiklech
 * mají schéma dynamické — sloupce odpovídají přesně hlavičce naimportovaného
 * zdrojového souboru, proto se ukládají mimo RZ_SCHEMA jako syrová mřížka
 * (rzReadGrid_/rzWriteGrid_).
 */
const RZ_SLUG = 'rozdelovnik-20-artiklu';
const RZ_ARTIKLY_ROWS = 20;

const RZ_SCHEMA = {
  '_settings': ['id', 'key', 'value', 'updated_at', 'updated_by'],
  'artikly': ['id', 'poradi', 'cislo_artiklu', 'nazev', 'k_rozdeleni', 'pocet_dni', 'metropol', 'created_at', 'created_by', 'updated_at'],
};

const RZ_IMPORT_TABLES = ['odprodej', 'teo_stavy', 'vyskladnovaci_listy', 'prideleni_po_artiklech'];

const RZ_PATTERN_SETTING_KEY = {
  odprodej: 'patternOdprodej',
  teo_stavy: 'patternTeoStavy',
  vyskladnovaci_listy: 'patternVyskladnovaciListy',
  prideleni_po_artiklech: 'patternPrideleniPoArtiklech',
};

function rzApp_() {
  const app = dbGetAll_(SHEETS.APPS).find((a) => a.slug === RZ_SLUG);
  if (!app) throw new Error('Subaplikace Rozdělovník 20 artiklů nenalezena (založte ji v sekci Aplikace).');
  if (!app.db_spreadsheet_id) throw new Error('Subaplikace ještě nemá vlastní databázi — zkuste ji znovu uložit v sekci Aplikace.');
  return app;
}

let rzRepoCache_ = null;
function rzRepo_() {
  if (rzRepoCache_) return rzRepoCache_;
  rzRepoCache_ = createDbRepo_(rzApp_().db_spreadsheet_id, RZ_SCHEMA);
  return rzRepoCache_;
}

/** Obal endpointů subaplikace: základní přístup (ROLES.USER) + kontrola allowed_apps pro tuto appku. */
function rzGuard_(fn) {
  return guard_(ROLES.USER, (user) => {
    const app = rzApp_();
    if (!isAppAllowedForUser_(user, app)) throw new Error('Nemáte přístup do této subaplikace.');
    return fn(user, app);
  });
}

function rzCanWrite_(user) {
  return user.role === 'SUPERADMIN' || user.permission === 'EDITOR';
}

/* ── Nastavení ────────────────────────────────────────────────── */

function rzSettingsAll_() {
  const repo = rzRepo_();
  repo.ensureSchema();
  const settings = {};
  repo.getAll('_settings').forEach((row) => { settings[row.key] = row.value; });
  return settings;
}

function rzSettingsSet_(key, value) {
  const repo = rzRepo_();
  const rows = repo.getAll('_settings');
  const existing = rows.find((r) => r.key === key);
  if (existing) {
    repo.update('_settings', existing.id, { value: value, updated_by: currentEmail_() });
  } else {
    repo.insert('_settings', { key: key, value: value, updated_by: currentEmail_() });
  }
}

function apiRzGetSettings() {
  return rzGuard_(() => rzSettingsAll_());
}

function apiRzSaveSettings(payload) {
  return rzGuard_((user) => {
    if ((ROLE_LEVEL[user.role] || 0) < ROLE_LEVEL[ROLES.ADMIN]) throw new Error('Nemáte oprávnění měnit nastavení.');
    const keys = ['syncFolderUrl'].concat(Object.values(RZ_PATTERN_SETTING_KEY));
    keys.forEach((key) => rzSettingsSet_(key, String((payload && payload[key]) || '').trim()));
    audit_('rz_settings_update', 'Aktualizace nastavení Rozdělovníku 20 artiklů.');
    return rzSettingsAll_();
  });
}

/** Podporuje URL i přímo vložené ID složky. */
function rzExtractFolderId_(input) {
  const raw = String(input || '').trim();
  if (!raw) return null;
  const match = raw.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(raw)) return raw;
  return null;
}

/* ── Artikly ──────────────────────────────────────────────────── */

function apiRzListArtikly() {
  return rzGuard_(() => {
    const repo = rzRepo_();
    repo.ensureSchema();
    return repo.getAll('artikly').sort((a, b) => (Number(a.poradi) || 0) - (Number(b.poradi) || 0));
  });
}

/** Uloží najednou všech 20 řádků (přepíše existující podle pořadí, chybějící založí). */
function apiRzSaveArtikly(rows) {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k zápisu.');
    if (!Array.isArray(rows)) throw new Error('Neplatná data.');
    const repo = rzRepo_();
    repo.ensureSchema();
    const existing = repo.getAll('artikly');
    const byPoradi = new Map(existing.map((r) => [Number(r.poradi), r]));

    const saved = rows.slice(0, RZ_ARTIKLY_ROWS).map((row, i) => {
      const poradi = i + 1;
      const data = {
        poradi: poradi,
        cislo_artiklu: String((row && row.cislo_artiklu) || '').trim(),
        nazev: String((row && row.nazev) || '').trim(),
        k_rozdeleni: (row && row.k_rozdeleni !== '' && row.k_rozdeleni != null) ? Number(row.k_rozdeleni) || 0 : '',
        pocet_dni: (row && row.pocet_dni !== '' && row.pocet_dni != null) ? Number(row.pocet_dni) || 0 : '',
        metropol: !!(row && row.metropol),
      };
      const ex = byPoradi.get(poradi);
      return ex ? repo.update('artikly', ex.id, data) : repo.insert('artikly', data);
    });

    audit_('rz_artikly_save', saved.filter((r) => r.cislo_artiklu).length + ' vyplněných řádků z ' + RZ_ARTIKLY_ROWS);
    return saved.sort((a, b) => (Number(a.poradi) || 0) - (Number(b.poradi) || 0));
  });
}

/**
 * Referenční přehled filiálek (Metropol se nastavuje v hlavním dashboardu — čteme napřímo hlavní DB).
 * Appka je postavená pro jedno konkrétní LC — vracíme jen filiálky spadající pod
 * výchozí LC nastavené v hlavním dashboardu (Log. centra), ne celou síť.
 */
function apiRzListStores() {
  return rzGuard_(() => {
    const defaultLcCode = String(settingsAll_().defaultLcCode || '').trim().toUpperCase();
    if (!defaultLcCode) throw new Error('V hlavním dashboardu není nastaveno výchozí logistické centrum (sekce Log. centra).');
    return dbGetAll_(SHEETS.STORES)
      .filter((s) => s.active === true && String(s.lc_code).trim().toUpperCase() === defaultLcCode)
      .map((s) => ({ code: s.code, name: s.name, metropolitni: !!s.metropolitni }))
      .sort((a, b) => String(a.code).localeCompare(String(b.code)));
  });
}

/* ── Import zdrojových souborů (dynamické schéma) ────────────────
 * Odprodej/Teo_stavy/Vyskladňovací_listy/Přidělení_po_artiklech nemají
 * pevně daný seznam sloupců v kódu — ukládají se jako syrová mřížka
 * (hlavička + řádky) přesně podle toho, co obsahuje zdrojový soubor. */

function rzImportSheet_(fileKey) {
  const ss = rzRepo_().spreadsheet();
  let sheet = ss.getSheetByName(fileKey);
  if (!sheet) sheet = ss.insertSheet(fileKey);
  return sheet;
}

function rzReadGrid_(fileKey) {
  const sheet = rzImportSheet_(fileKey);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { headers: [], rows: [] };
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  return { headers: values[0].map((h) => String(h)), rows: values.slice(1) };
}

function rzWriteGrid_(fileKey, headers, rows) {
  const sheet = rzImportSheet_(fileKey);
  sheet.clearContents();
  if (!headers || !headers.length) return;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

/** Najde ve složce nejnovější soubor (.xlsx/.csv/Google Sheets), jehož název obsahuje daný výraz. */
function rzFindFileInFolderByName_(folderId, namePattern) {
  const folder = DriveApp.getFolderById(folderId);
  const needle = String(namePattern).toLowerCase();
  let newest = null;
  let newestDate = null;
  [MimeType.MICROSOFT_EXCEL, MimeType.GOOGLE_SHEETS, MimeType.CSV].forEach((mimeType) => {
    const files = folder.getFilesByType(mimeType);
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().toLowerCase().indexOf(needle) === -1) continue;
      const date = file.getLastUpdated();
      if (!newestDate || date > newestDate) { newest = file; newestDate = date; }
    }
  });
  return newest;
}

/**
 * Přečte zdrojový soubor (.xlsx/.csv/Google Sheets) a vrátí { headers, rows }.
 * U .xlsx/Sheets použije stejný trik jako hlavní synchronizace (60_sync.js) —
 * dočasná kopie jako Google Sheet, přečtení, smazání kopie.
 */
function rzReadSourceFile_(file) {
  if (file.getMimeType() === MimeType.CSV) {
    const text = file.getBlob().getDataAsString('UTF-8');
    const data = Utilities.parseCsv(text);
    if (!data.length) return { headers: [], rows: [] };
    return { headers: data[0].map(String), rows: data.slice(1) };
  }

  let tempSheetId = null;
  try {
    const scriptFolder = scriptFolder_();
    const copyMeta = { name: '__rz_import_tmp__', mimeType: 'application/vnd.google-apps.spreadsheet' };
    if (scriptFolder) copyMeta.parents = [scriptFolder.getId()];
    const copy = Drive.Files.copy(copyMeta, file.getId());
    tempSheetId = copy.id;
    const ss = SpreadsheetApp.openById(tempSheetId);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return { headers: [], rows: [] };
    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    return { headers: values[0].map(String), rows: values.slice(1) };
  } catch (e) {
    throw new Error('Nepodařilo se načíst soubor "' + file.getName() + '": ' + e.message);
  } finally {
    if (tempSheetId) { try { Drive.Files.remove(tempSheetId); } catch (_) { /* dočasný soubor, chyba mazání není kritická */ } }
  }
}

function apiRzImport(fileKey) {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k importu.');
    if (RZ_IMPORT_TABLES.indexOf(fileKey) === -1) throw new Error('Neplatný typ importu.');

    const settings = rzSettingsAll_();
    const folderId = rzExtractFolderId_(settings.syncFolderUrl);
    if (!folderId) throw new Error('Není nastavena složka se zdrojovými soubory (Nastavení).');

    const patternKey = RZ_PATTERN_SETTING_KEY[fileKey];
    const pattern = settings[patternKey] || '';
    if (!pattern) throw new Error('Není nastaven výraz pro vyhledání tohoto souboru (Nastavení).');

    const file = rzFindFileInFolderByName_(folderId, pattern);
    if (!file) throw new Error('Ve složce nebyl nalezen žádný soubor obsahující v názvu „' + pattern + '".');

    const { headers, rows } = rzReadSourceFile_(file);
    rzWriteGrid_(fileKey, headers, rows);
    audit_('rz_import_' + fileKey, file.getName() + ' (' + rows.length + ' řádků)');
    return { fileName: file.getName(), headers: headers, rowCount: rows.length };
  });
}

function apiRzGetImportTable(fileKey) {
  return rzGuard_(() => {
    if (RZ_IMPORT_TABLES.indexOf(fileKey) === -1) throw new Error('Neplatný typ importu.');
    return rzReadGrid_(fileKey);
  });
}

function apiRzSaveImportTable(payload) {
  return rzGuard_((user) => {
    if (!rzCanWrite_(user)) throw new Error('Nemáte oprávnění k zápisu.');
    const fileKey = payload && payload.fileKey;
    if (RZ_IMPORT_TABLES.indexOf(fileKey) === -1) throw new Error('Neplatný typ importu.');
    rzWriteGrid_(fileKey, (payload && payload.headers) || [], (payload && payload.rows) || []);
    return { ok: true };
  });
}
