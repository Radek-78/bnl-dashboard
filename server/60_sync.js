/**
 * Synchronizace dat filiálek z externího souboru (.xlsx nebo Google Sheets).
 *
 * Postup: najde nejnovější tabulkový soubor v zadané Drive složce (volitelně jen ty, jejichž
 * název obsahuje nastavený výraz - viz syncFileNamePattern) → zkopíruje jako Google Sheet
 * (u .xlsx tím proběhne konverze, u už existujícího Sheets souboru jde o obyčejnou kopii)
 * → přečte list "Organizace_Detail" (filiálky), "Zavrene_Openings" (dočasná uzavření)
 * a nepovinně "Organizace" (datum otevření) → porovná s DB → provede INSERT/UPDATE/deaktivaci.
 *
 * Uzavírky ze zdroje se ukládají zvlášť (sync_closed_ranges) a každá synchronizace je
 * celé nahradí - zrušená nebo zkrácená uzavírka ve zdroji tak zmizí i v appce. Ručně
 * zadané uzavírky (temp_closed_ranges) sync nemění.
 *
 * Sloupec LC nese v souboru celý název logistického centra (např. "Brandýs nad Labem"),
 * ne zkratku — sync ho páruje na existující záznam v Log. centrech podle názvu. Filiálka,
 * jejíž LC název nejde spárovat, se v daném běhu přeskočí a nahlásí jako chyba (nic se
 * u ní nezmění).
 */

const STORES_COL_MAP = {
  'Číslo':            'code',
  'Název':            'name',
  'LC':               'lc_name',
  'Telefon prodejny': 'phone',
  'VT':               'area_manager',
  'Telefon VT':       'vt_phone',
  'RM':               'regional_manager',
  'Telefon RM':       'rm_phone',
  'Pondělí otevřeno': 'mon_open',
  'Pondělí zavřeno':  'mon_close',
  'Úterý otevřeno':   'tue_open',
  'Úterý zavřeno':    'tue_close',
  'Středa otevřeno':  'wed_open',
  'Středa zavřeno':   'wed_close',
  'Čtvrtek otevřeno': 'thu_open',
  'Čtvrtek zavřeno':  'thu_close',
  'Pátek otevřeno':   'fri_open',
  'Pátek zavřeno':    'fri_close',
  'Sobota otevřeno':  'sat_open',
  'Sobota zavřeno':   'sat_close',
  'Neděle otevřeno':  'sun_open',
  'Neděle zavřeno':   'sun_close',
  'Ulice':            'street',
  'Město':            'city',
  'PSČ':              'zip',
  'Zástupce RM':      'deputy_rm',
  'Telefon zástupce': 'deputy_phone',
};

/** List s datem oficiálního otevření filiálky - NEPOVINNÝ (stejně jako v Planung Dashboardu). */
const SYNC_OPENINGS_SHEET = 'Organizace';

/** Klíč v CacheService - ruční a automatická synchronizace nesmí běžet současně. */
const SYNC_RUNNING_KEY_ = 'sync:running';

/* ── Veřejné API ──────────────────────────────────────────────── */

function apiRunSync() {
  return guard_(ROLES.ADMIN, () => {
    const result = runSyncCore_(settingsAll_());
    audit_('sync_run',
      'Soubor: ' + result.fileName +
      ' | Filiálky: +' + result.stores.added + ' u' + result.stores.updated + ' d' + result.stores.deactivated
    );
    return result;
  });
}

/**
 * Přečte z aktuálně nakonfigurovaného zdrojového souboru všechny názvy LC ve
 * sloupci LC a vrátí ty, které zatím nemají záznam v Log. centrech. Používá
 * tlačítko "Doplnit LC ze souboru" — nic nezapisuje do DB.
 */
function apiFindMissingLcInFile() {
  return guard_(ROLES.ADMIN, () => {
    const settings = settingsAll_();
    const folderUrl = settings.syncFolderUrl || '';
    if (!folderUrl) throw new Error('Není nastavena URL složky. Vyplňte ji v sekci Synchronizace.');
    const folderId = extractFolderIdFromUrl_(folderUrl);
    if (!folderId) throw new Error('Z URL složky se nepodařilo rozpoznat ID.');

    const file = findSyncFileInFolder_(folderId, settings.syncFileNamePattern);
    if (!file) {
      throw new Error('Ve složce nebyl nalezen žádný soubor .xlsx ani Google Sheets'
        + (settings.syncFileNamePattern ? ' odpovídající výrazu "' + settings.syncFileNamePattern + '".' : '.'));
    }

    const storesSheetName = settings.syncStoresSheet || 'Organizace_Detail';
    let tempSheetId = null;
    let names = [];
    try {
      const scriptFolder = scriptFolder_();
      const copyMeta = { name: '__sync_lc_tmp__', mimeType: 'application/vnd.google-apps.spreadsheet' };
      if (scriptFolder) copyMeta.parents = [scriptFolder.getId()];
      const copy = Drive.Files.copy(copyMeta, file.getId(), { supportsAllDrives: true });
      tempSheetId = copy.id;
      const ss = SpreadsheetApp.openById(tempSheetId);
      const sheet = ss.getSheetByName(storesSheetName);
      if (!sheet) throw new Error('List "' + storesSheetName + '" nebyl v souboru nalezen.');
      const rows = parseSheetRows_(sheet, STORES_COL_MAP);
      const seen = new Set();
      rows.forEach((r) => { const name = String(r.lc_name || '').trim(); if (name) seen.add(name); });
      names = [...seen];
    } finally {
      trashTempFile_(tempSheetId);
    }

    const known = new Set(dbGetAll_(SHEETS.LOGISTICS).map((lc) => String(lc.name || '').trim().toLowerCase()));
    const missing = names.filter((n) => !known.has(n.toLowerCase())).sort((a, b) => a.localeCompare(b, 'cs'));
    return { missing: missing };
  });
}

/** Poznamená čas a výsledek automatické kontroly — zobrazuje se v Konfiguraci. */
function autoSyncNoteCheck_(outcome) {
  settingsSet_('autoSyncLastCheckAt', nowIso_());
  settingsSet_('autoSyncLastCheckResult', outcome);
}

/**
 * Cíl časovaného triggeru (viz apiSaveSyncSettings) — jednou denně zkontroluje,
 * zda se ve složce od poslední kontroly změnil soubor (jiné ID nebo novější úprava),
 * a pokud ano, spustí stejnou synchronizaci jako ruční tlačítko.
 */
function autoSyncCheck_() {
  try {
    const settings = settingsAll_();
    if (settings.autoSyncEnabled !== true && settings.autoSyncEnabled !== 'true') return;

    const folderUrl = settings.syncFolderUrl || '';
    if (!folderUrl) { autoSyncNoteCheck_('složka není nastavena'); return; }
    const folderId = extractFolderIdFromUrl_(folderUrl);
    if (!folderId) { autoSyncNoteCheck_('z URL složky nelze rozpoznat ID'); return; }

    const file = findSyncFileInFolder_(folderId, settings.syncFileNamePattern);
    if (!file) { autoSyncNoteCheck_('ve složce nebyl nalezen žádný odpovídající soubor .xlsx ani Google Sheets'); return; }

    const signature = file.getId() + ':' + file.getLastUpdated().getTime();
    if (signature === settings.syncLastFileSignature) {
      autoSyncNoteCheck_('soubor beze změny, synchronizace nebyla potřeba');
      return;
    }

    const result = runSyncCore_(settings, true);
    autoSyncNoteCheck_('soubor se změnil, synchronizace proběhla');
    audit_('sync_run_auto',
      'Soubor: ' + result.fileName +
      ' | Filiálky: +' + result.stores.added + ' u' + result.stores.updated + ' d' + result.stores.deactivated
    );
  } catch (e) {
    console.error('Automatická synchronizace selhala: ' + e);
    try { autoSyncNoteCheck_('chyba: ' + String(e && e.message ? e.message : e)); } catch (_) {}
    // Chyba nočního běhu musí být vidět v historii synchronizací, ne jen v auditu.
    try { appendSyncHistory_(settingsAll_(), null, true, String(e && e.message ? e.message : e)); } catch (_) {}
    audit_('sync_run_auto_error', String(e && e.message ? e.message : e));
    throw e; // necháme GAS poslat vlastníkovi e-mail o selhání triggeru
  }
}

/* ── Interní funkce ───────────────────────────────────────────── */

/**
 * Zapíše kompaktní záznam o proběhlé synchronizaci do _settings.syncHistory
 * (posledních 20 běhů — kdo, kdy, soubor, počty). Detail změn drží jen
 * poslední běh (lastSyncResult), historie je jen souhrn. errorMessage = běh
 * selhal (result je pak null) - zapíše se jako záznam s chybou.
 */
function appendSyncHistory_(settings, result, isAuto, errorMessage) {
  let history = [];
  try { history = settings.syncHistory ? JSON.parse(settings.syncHistory) : []; } catch (e) { history = []; }
  const s = (result && result.stores) || {};
  history.unshift({
    at: nowIso_(),
    by: currentEmail_() || 'system',
    auto: isAuto === true,
    file: (result && result.fileName) || '',
    added: s.added || 0,
    updated: s.updated || 0,
    deactivated: s.deactivated || 0,
    reactivated: s.reactivated || 0,
    unchanged: s.unchanged || 0,
    closed: s.closedNew || 0,
    reopened: s.closedEnded || 0,
    errors: (s.errors || []).length,
    failure: errorMessage || '',
  });
  if (history.length > 20) history = history.slice(0, 20);
  settingsSet_('syncHistory', JSON.stringify(history));
}

/** Jádro synchronizace sdílené ruční (apiRunSync) i automatickou (autoSyncCheck_) cestou. */
function runSyncCore_(settings, isAuto) {
  // Ruční a automatický běh nesmí běžet současně (oba přepisují celý list stores).
  // Příznak v cache místo LockService - zámek by po celou dobu syncu blokoval i běžné zápisy do DB.
  const cache = CacheService.getScriptCache();
  if (cache.get(SYNC_RUNNING_KEY_)) throw new Error('Synchronizace právě probíhá, zkuste to prosím za chvíli.');
  cache.put(SYNC_RUNNING_KEY_, '1', 900);
  try {
    return runSyncCoreUnlocked_(settings, isAuto);
  } finally {
    try { cache.remove(SYNC_RUNNING_KEY_); } catch (e) { /* příznak by stejně vypršel */ }
  }
}

function runSyncCoreUnlocked_(settings, isAuto) {
  const folderUrl = settings.syncFolderUrl || '';
  if (!folderUrl) throw new Error('Není nastavena URL složky. Vyplňte ji v sekci Synchronizace.');

  const folderId = extractFolderIdFromUrl_(folderUrl);
  if (!folderId) throw new Error('Z URL složky se nepodařilo rozpoznat ID. Použijte URL ve tvaru https://drive.google.com/drive/folders/...');

  const xlsxFile = findSyncFileInFolder_(folderId, settings.syncFileNamePattern);
  if (!xlsxFile) {
    throw new Error('Ve složce nebyl nalezen žádný soubor .xlsx ani Google Sheets'
      + (settings.syncFileNamePattern ? ' odpovídající výrazu "' + settings.syncFileNamePattern + '".' : '.'));
  }

  let ss;
  let tempSheetId = null;
  try {
    const scriptFolder = scriptFolder_();
    const copyMeta = { name: '__sync_tmp__', mimeType: 'application/vnd.google-apps.spreadsheet' };
    if (scriptFolder) copyMeta.parents = [scriptFolder.getId()];
    // supportsAllDrives: soubor může ležet ve sdíleném disku (Shared Drive) - bez
    // tohoto parametru Drive.Files.copy hlásí "File not found", i když DriveApp
    // stejný soubor bez problémů najde (stejná oprava jako v 70_rozdelovnik.js, v3.1.53).
    const copy = Drive.Files.copy(copyMeta, xlsxFile.getId(), { supportsAllDrives: true });
    tempSheetId = copy.id;
    ss = SpreadsheetApp.openById(tempSheetId);
  } catch (e) {
    throw new Error('Nepodařilo se převést soubor "' + xlsxFile.getName() + '" na Google Sheet: ' + e.message);
  }

  let result;
  try {
    result = {
      fileName: xlsxFile.getName(),
      stores: syncStores_(ss, settings),
    };
  } finally {
    trashTempFile_(tempSheetId);
  }

  settingsSet_('lastSyncAt', nowIso_());
  settingsSet_('lastSyncResult', JSON.stringify(result));
  settingsSet_('syncLastFileSignature', xlsxFile.getId() + ':' + xlsxFile.getLastUpdated().getTime());
  appendSyncHistory_(settings, result, isAuto);
  return result;
}

function syncStores_(ss, settings) {
  dbEnsureSchema_(dbSpreadsheet_());

  const mainSheetName = settings.syncStoresSheet || 'Organizace_Detail';
  const closuresSheetName = settings.syncClosuresSheet || 'Zavrene_Openings';

  const sheet1 = ss.getSheetByName(mainSheetName);
  if (!sheet1) throw new Error('List "' + mainSheetName + '" nebyl v souboru nalezen.');

  // Filiálky s číslem nad 900 se ze zdroje nikdy nenačítají ani nezakládají (testovací/vyhrazený rozsah čísel).
  const mainRows = parseSheetRows_(sheet1, STORES_COL_MAP).filter((r) => !(parseInt(r.code, 10) > 900));
  const xlsxMap = new Map(mainRows.map((r) => [r.code, r]));
  const sourceCodes = new Set(mainRows.map((r) => String(r.code)));

  // Mapa LC: název (malými, trimovaný) → zkratka. Zdroj nese jen celý název LC.
  const lcByName = {};
  dbGetAll_(SHEETS.LOGISTICS).forEach((lc) => {
    const key = String(lc.name || '').trim().toLowerCase();
    if (key) lcByName[key] = lc.abbreviation;
  });
  const resolveLc_ = (xlsxRow) => lcByName[String(xlsxRow.lc_name || '').trim().toLowerCase()] || null;

  const currentRecords = dbGetAll_(SHEETS.STORES);

  const CHANGES_LIMIT = 50;
  const stats = { added: 0, updated: 0, deactivated: 0, reactivated: 0, unchanged: 0, errors: [],
                  changes: { added: [], updated: [], deactivated: [], reactivated: [],
                             closed: [], reopened: [], notYetOpen: [], manualInactive: [] },
                  closedNew: 0, closedEnded: 0, notYetOpen: 0, manualInactive: 0,
                  closuresSheetFound: false, openingsSheetFound: false };
  const now = nowIso_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const newRecords = [];

  // Datum otevření z listu Organizace (nepovinný). Bez listu zůstávají data otevření beze změny.
  let openingsByCode = null;
  const sheetOpenings = ss.getSheetByName(SYNC_OPENINGS_SHEET);
  if (sheetOpenings) {
    const parsed = parseOpeningsRows_(sheetOpenings);
    if (parsed) {
      stats.openingsSheetFound = true;
      openingsByCode = {};
      parsed.forEach((r) => { openingsByCode[r.code] = r.date; });
    } else {
      stats.errors.push('List "' + SYNC_OPENINGS_SHEET + '" nemá sloupce "Číslo" a "Datum Otevření" — data otevření zůstala beze změny.');
    }
  }
  const openingFor_ = (code) => (openingsByCode ? (openingsByCode[String(code)] || '') : undefined);

  // Zpracování stávajících DB záznamů
  currentRecords.forEach((existing) => {
    const codeKey = String(existing.code);
    if (!xlsxMap.has(codeKey)) {
      if (existing.active === true) {
        newRecords.push(Object.assign({}, existing, { active: false, updated_at: now }));
        stats.deactivated++;
        if (stats.changes.deactivated.length < CHANGES_LIMIT)
          stats.changes.deactivated.push({ code: existing.code, name: existing.name });
      } else {
        newRecords.push(existing);
        stats.unchanged++;
      }
      return;
    }

    const xlsxRow = xlsxMap.get(codeKey);
    xlsxMap.delete(codeKey);
    const lcAbbr = resolveLc_(xlsxRow);
    if (!lcAbbr) {
      stats.errors.push('Filiálka ' + codeKey + ' (' + (xlsxRow.name || existing.name || '') + '): LC "' + (xlsxRow.lc_name || '') + '" nenalezeno v Log. centrech — filiálka nebyla v tomto běhu aktualizována.');
      newRecords.push(existing); // beze změny, žádné riziko ztráty dat kvůli nerozpoznanému LC
      stats.unchanged++;
      return;
    }

    const patch = buildStorePatch_(xlsxRow, now, existing, lcAbbr, openingFor_(codeKey));

    if (existing.manually_inactive === true) {
      // Ručně deaktivovaná filiálka — sync ji neaktivuje zpět
      newRecords.push(Object.assign({}, existing, patch, { active: false, manually_inactive: true }));
      stats.unchanged++;
    } else {
      const changedFields = storeChangedFields_(existing, patch);
      const wasInactive = existing.active !== true;

      if (changedFields.length > 0 || wasInactive) {
        newRecords.push(Object.assign({}, existing, patch));
        if (wasInactive) {
          stats.reactivated++;
          if (stats.changes.reactivated.length < CHANGES_LIMIT)
            stats.changes.reactivated.push({ code: existing.code, name: patch.name || existing.name });
        } else {
          stats.updated++;
          if (stats.changes.updated.length < CHANGES_LIMIT)
            stats.changes.updated.push({ code: existing.code, name: patch.name || existing.name, fields: changedFields });
        }
      } else {
        // Bez hlášené změny - přesto se převezme i první doplnění data otevření (viz storeChangedFields_),
        // časové značky zůstávají původní.
        newRecords.push(Object.assign({}, existing, patch, { updated_at: existing.updated_at, synced_at: existing.synced_at }));
        stats.unchanged++;
      }
    }
  });

  // Nové záznamy z xlsx (nezpracované = nebyly v DB)
  xlsxMap.forEach((xlsxRow, code) => {
    const lcAbbr = resolveLc_(xlsxRow);
    if (!lcAbbr) {
      stats.errors.push('Filiálka ' + code + ' (' + (xlsxRow.name || '') + '): LC "' + (xlsxRow.lc_name || '') + '" nenalezeno v Log. centrech — filiálka nebyla založena.');
      return;
    }
    newRecords.push(Object.assign(buildStorePatch_(xlsxRow, now, null, lcAbbr, openingFor_(code)), {
      id: uuid_(),
      created_at: now,
      created_by: currentEmail_() || 'sync',
      synced_at: now,
    }));
    stats.added++;
    if (stats.changes.added.length < CHANGES_LIMIT)
      stats.changes.added.push({ code, name: xlsxRow.name || '' });
  });

  // Dočasná uzavření z listu Zavrene_Openings - zrcadlo zdroje v sync_closed_ranges (každý běh je
  // celé nahradí). Ručně zadané uzavírky (temp_closed_ranges) se nemění; jen se z nich odstraní
  // rozsahy, které jsou přesně stejné jako ve zdroji (dřívější verze syncu je tam slučovala).
  const sheet2 = ss.getSheetByName(closuresSheetName);
  if (sheet2) {
    stats.closuresSheetFound = true;
    const closuresByCode = {};
    parseClosuresRows_(sheet2).forEach((r) => { (closuresByCode[r.code] = closuresByCode[r.code] || []).push({ from: r.from, to: r.to }); });

    newRecords.forEach((rec) => {
      const code = String(rec.code);
      if (!sourceCodes.has(code)) return; // filiálka mimo zdroj - uzavírky se nechávají, jak jsou
      const fresh = normalizeClosureRanges_(closuresByCode[code] || []);
      const old = normalizeClosureRanges_(parseClosureRanges_(rec.sync_closed_ranges));
      const freshKeys = new Set(fresh.map((r) => r.from + '|' + r.to));
      const oldKeys = new Set(old.map((r) => r.from + '|' + r.to));
      const opened = fresh.filter((r) => !oldKeys.has(r.from + '|' + r.to));
      const ended = old.filter((r) => !freshKeys.has(r.from + '|' + r.to));

      const manual = parseClosureRanges_(rec.temp_closed_ranges);
      const manualLeft = manual.filter((r) => !freshKeys.has(r.from + '|' + r.to));
      if (manualLeft.length !== manual.length) rec.temp_closed_ranges = manualLeft.length ? JSON.stringify(manualLeft) : '';

      if (opened.length || ended.length) {
        rec.sync_closed_ranges = fresh.length ? JSON.stringify(fresh) : '';
        rec.updated_at = now;
      }
      opened.forEach((r) => {
        stats.closedNew++;
        if (stats.changes.closed.length < CHANGES_LIMIT) stats.changes.closed.push({ code: rec.code, name: rec.name, from: r.from, to: r.to });
      });
      ended.forEach((r) => {
        stats.closedEnded++;
        if (stats.changes.reopened.length < CHANGES_LIMIT) stats.changes.reopened.push({ code: rec.code, name: rec.name, from: r.from, to: r.to });
      });
    });
  }

  // Stav k dnešku pro report: filiálky před otevřením a ručně deaktivované, které jsou ve zdroji
  // a nejsou zavřené (dřív se ručně vypínaly kvůli uzavření - sync je sám nezapne, jen upozorní).
  newRecords.forEach((rec) => {
    rec.temporarily_closed = isTempClosedNow_(rec);
    const opening = normalizeIsoDate_(rec.opening_date);
    if (rec.active === true && opening > today) {
      stats.notYetOpen++;
      if (stats.changes.notYetOpen.length < CHANGES_LIMIT) stats.changes.notYetOpen.push({ code: rec.code, name: rec.name, date: opening });
    }
    if (rec.manually_inactive === true && sourceCodes.has(String(rec.code)) && !rec.temporarily_closed && !(opening > today)) {
      stats.manualInactive++;
      if (stats.changes.manualInactive.length < CHANGES_LIMIT) stats.changes.manualInactive.push({ code: rec.code, name: rec.name });
    }
  });

  // Datum otevření a textové údaje jako prostý text - jinak by Sheets text, který vypadá jako datum
  // (název filiálky nebo ulice "28. října"), při zápisu převedl na datum.
  const storesSheet = dbSheet_(SHEETS.STORES);
  if (storesSheet.getMaxRows() > 1) {
    ['opening_date', 'name', 'street', 'city', 'area_manager', 'regional_manager', 'deputy_rm'].forEach((f) => {
      const col = DB_SCHEMA[SHEETS.STORES].indexOf(f) + 1;
      storesSheet.getRange(2, col, storesSheet.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  }

  dbBatchReplace_(SHEETS.STORES, newRecords);
  return stats;
}

/* ── Pomocné funkce ───────────────────────────────────────────── */

const HOUR_FIELDS_ = [
  'mon_open','mon_close','tue_open','tue_close','wed_open','wed_close',
  'thu_open','thu_close','fri_open','fri_close','sat_open','sat_close','sun_open','sun_close',
];

/**
 * Sestaví patch pro jednu filiálku. lcAbbr je už vyřešená zkratka LC (viz
 * resolveLc_ v syncStores_) — sem přichází vždy platná, jinak se řádek
 * nezpracovává vůbec. Zdroj je zrcadlo: prázdná buňka přepíše i stávající
 * hodnotu v DB. Jen když sloupec ve zdroji úplně chybí, hodnota z DB zůstane.
 * openingDate: undefined = list Organizace chybí (datum otevření se nemění).
 */
function buildStorePatch_(xlsxRow, now, existing, lcAbbr, openingDate) {
  const NON_HOUR_FIELDS = ['code', 'name', 'phone', 'area_manager', 'vt_phone', 'regional_manager', 'rm_phone',
    'street', 'city', 'zip', 'deputy_rm', 'deputy_phone'];
  const patch = { active: true, synced_at: now, updated_at: now, lc_code: lcAbbr };
  NON_HOUR_FIELDS.concat(HOUR_FIELDS_).forEach((f) => {
    const dbVal = existing ? (existing[f] || '') : '';
    patch[f] = xlsxRow[f] !== undefined ? xlsxRow[f] : dbVal;
  });
  patch.opening_date = openingDate !== undefined ? openingDate : (existing ? normalizeIsoDate_(existing.opening_date) : '');
  return patch;
}

const STORE_DIFF_FIELDS = [
  'name','lc_code','phone','area_manager','vt_phone','regional_manager','rm_phone',
  'mon_open','mon_close','tue_open','tue_close','wed_open','wed_close',
  'thu_open','thu_close','fri_open','fri_close','sat_open','sat_close','sun_open','sun_close',
  'opening_date', 'street', 'city', 'zip', 'deputy_rm', 'deputy_phone',
];

const STORE_FIELD_LABELS = {
  name: 'Název', lc_code: 'LC', phone: 'Telefon prodejny',
  area_manager: 'VT', vt_phone: 'Telefon VT', regional_manager: 'RM', rm_phone: 'Telefon RM',
  mon_open: 'Po otevřeno', mon_close: 'Po zavřeno',
  tue_open: 'Út otevřeno', tue_close: 'Út zavřeno',
  wed_open: 'St otevřeno', wed_close: 'St zavřeno',
  thu_open: 'Čt otevřeno', thu_close: 'Čt zavřeno',
  fri_open: 'Pá otevřeno', fri_close: 'Pá zavřeno',
  sat_open: 'So otevřeno', sat_close: 'So zavřeno',
  sun_open: 'Ne otevřeno', sun_close: 'Ne zavřeno',
  opening_date: 'Datum otevření',
  street: 'Ulice', city: 'Město', zip: 'PSČ', deputy_rm: 'Zástupce RM', deputy_phone: 'Telefon zástupce',
};

/** Pole přidaná později - jejich první doplnění (v DB dosud prázdné) se nehlásí jako změna. */
const STORE_SILENT_FIRST_FILL = ['opening_date', 'street', 'city', 'zip', 'deputy_rm', 'deputy_phone'];

/**
 * Hlášené změny polí. První doplnění později přidaných polí (STORE_SILENT_FIRST_FILL, dosud prázdné)
 * se nehlásí - jinak by první běh po zavedení sloupce ohlásil jako změněnou každou filiálku.
 * Hodnota se přesto uloží.
 */
function storeChangedFields_(existing, patch) {
  const result = [];
  STORE_DIFF_FIELDS.forEach((f) => {
    const oldVal = f === 'opening_date' ? normalizeIsoDate_(existing[f]) : String(existing[f] || '');
    const newVal = String(patch[f] || '');
    if (STORE_SILENT_FIRST_FILL.indexOf(f) !== -1 && !oldVal) return;
    if (oldVal !== newVal)
      result.push({ field: STORE_FIELD_LABELS[f] || f, old: oldVal, new: newVal });
  });
  return result;
}

/** Vrátí trimovaná záhlaví z prvního řádku listu. */
function readSheetHeaders_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h).trim());
}

/**
 * Přečte list tabulky a vrátí pole objektů namapovaných přes colMap.
 * Záhlaví je na řádku 1 (trimované). Prázdné řádky (bez code) jsou přeskočeny.
 */
function parseSheetRows_(sheet, colMap) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sheet.getLastColumn();
  const headers = readSheetHeaders_(sheet);

  // Index každého cílového pole
  const colIndices = {};
  Object.keys(colMap).forEach((xlsxHeader) => {
    const idx = headers.indexOf(xlsxHeader);
    if (idx !== -1) colIndices[colMap[xlsxHeader]] = idx;
  });

  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = range.getValues();
  // Zobrazený text buněk - když Sheets text s datem (název/ulice "28. října") sám převede
  // na datum, vezme se přesně to, co je ve zdroji vidět, ne prázdná hodnota.
  const display = range.getDisplayValues();
  return data
    .map((row, ri) => {
      const record = {};
      Object.keys(colIndices).forEach((dbField) => {
        const ci = colIndices[dbField];
        record[dbField] = formatCellValue_(row[ci], display[ri][ci]);
      });
      return record;
    })
    .filter((r) => r.code);  // přeskočit řádky bez kódu
}

/**
 * Přečte list "Zavrene_Openings": číslo filiálky + rozsah dočasného uzavření
 * (Od/Do). Řádek bez rozpoznaného čísla nebo obou dat se přeskočí.
 */
function parseClosuresRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sheet.getLastColumn();
  const headers = readSheetHeaders_(sheet);
  const idx = { code: headers.indexOf('Číslo'), from: headers.indexOf('Od'), to: headers.indexOf('Do') };
  if (idx.code === -1 || idx.from === -1 || idx.to === -1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data
    .map((row) => ({
      code: formatCellValue_(row[idx.code]),
      from: formatDateCellValue_(row[idx.from]),
      to: formatDateCellValue_(row[idx.to]),
    }))
    .filter((r) => r.code && r.from && r.to);
}

/**
 * Přečte list "Organizace": číslo filiálky + datum oficiálního otevření
 * ('yyyy-MM-dd'). Vrací null, když list nemá potřebné sloupce.
 */
function parseOpeningsRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const headers = readSheetHeaders_(sheet);
  const idx = { code: headers.indexOf('Číslo'), date: headers.indexOf('Datum Otevření') };
  if (idx.code === -1 || idx.date === -1) return null;
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return data
    .map((row) => ({ code: formatCellValue_(row[idx.code]), date: formatDateCellValue_(row[idx.date]) }))
    .filter((r) => r.code && r.date);
}

/** Pole rozsahů {from, to} z JSON textu uloženého v DB ([] při prázdné/neplatné hodnotě). */
function parseClosureRanges_(json) {
  if (!json) return [];
  try {
    const parsed = JSON.parse(String(json));
    return Array.isArray(parsed) ? parsed.filter((r) => r && r.from && r.to).map((r) => ({ from: r.from, to: r.to })) : [];
  } catch (e) {
    return [];
  }
}

/** Seřadí rozsahy podle data a odstraní přesné duplicity (porovnání starého a nového stavu). */
function normalizeClosureRanges_(ranges) {
  const seen = new Set();
  return ranges
    .filter((r) => { const k = r.from + '|' + r.to; if (seen.has(k)) return false; seen.add(k); return true; })
    .sort((a, b) => (a.from + a.to < b.from + b.to ? -1 : a.from + a.to > b.from + b.to ? 1 : 0));
}

/**
 * Datum v DB → 'yyyy-MM-dd'. Kdyby Sheets textové datum přece jen převedl na Date,
 * dbDeserialize_ ho vrátí jako ISO čas v UTC ("…T23:00:00.000Z" = místní půlnoc dalšího dne) -
 * proto se takový čas převádí přes časové pásmo skriptu, ne useknutím textu.
 */
function normalizeIsoDate_(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const str = val !== undefined && val !== null ? String(val).trim() : '';
  if (!str) return '';
  if (/^\d{4}-\d{2}-\d{2}T/.test(str)) {
    const d = new Date(str);
    return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return formatDateCellValue_(str);
}

/**
 * Převede hodnotu buňky na string.
 * Časové buňky (h:mm) GAS vrací jako Date s datem 30.12.1899 — formátujeme jako "H:mm".
 * Jiné datum je text, který Sheets sám převedl na datum (např. název filiálky nebo ulice
 * "28. října") - vrací se zobrazený text buňky (displayVal), nikdy prázdná hodnota.
 */
function formatCellValue_(val, displayVal) {
  if (val instanceof Date) {
    if (val.getFullYear() === 1899 && val.getMonth() === 11 && val.getDate() === 30) {
      const h = val.getHours();
      const m = val.getMinutes();
      return h + ':' + (m < 10 ? '0' + m : m);
    }
    return displayVal !== undefined && displayVal !== null ? String(displayVal).trim() : '';
  }
  const str = (val !== undefined && val !== null) ? String(val).trim() : '';
  // Normalizace časového formátu "07:00" → "7:00" (h:mm)
  const timeMatch = str.match(/^(\d{1,2}):(\d{2})$/);
  if (timeMatch) return parseInt(timeMatch[1], 10) + ':' + timeMatch[2];
  return str;
}

/** Buňka se skutečným kalendářním datem (Od/Do) → 'yyyy-MM-dd', nebo '' když nejde rozpoznat. */
function formatDateCellValue_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = (val !== undefined && val !== null) ? String(val).trim() : '';
  if (!str) return '';
  const iso = str.match(/^(\d{4}-\d{2}-\d{2})/);
  if (iso) return iso[1];
  const cz = str.match(/^(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})$/);
  if (cz) return cz[3] + '-' + cz[2].padStart(2, '0') + '-' + cz[1].padStart(2, '0');
  return '';
}

/** Extrahuje ID složky z Google Drive URL. */
function extractFolderIdFromUrl_(url) {
  const match = String(url).match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Vrátí nejnovější tabulkový soubor (.xlsx nebo Google Sheets) ve složce nebo
 * null. Je-li zadaný namePattern, bere v potaz jen soubory, jejichž název ho
 * obsahuje (bez ohledu na velikost písmen) - jinak (prázdné/nezadané) vybere
 * nejnovější soubor bez ohledu na název, jak appka fungovala dřív.
 */
function findSyncFileInFolder_(folderId, namePattern) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const needle = String(namePattern || '').trim().toLowerCase();
    let newest = null;
    let newestDate = null;
    [MimeType.MICROSOFT_EXCEL, MimeType.GOOGLE_SHEETS].forEach((mimeType) => {
      const files = folder.getFilesByType(mimeType);
      while (files.hasNext()) {
        const file = files.next();
        if (needle && file.getName().toLowerCase().indexOf(needle) === -1) continue;
        const date = file.getLastUpdated();
        if (!newestDate || date > newestDate) { newest = file; newestDate = date; }
      }
    });
    return newest;
  } catch (e) {
    throw new Error('Nepodařilo se otevřít složku: ' + e.message);
  }
}
