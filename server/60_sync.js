/**
 * Synchronizace dat filiálek a logistických center z externího souboru (.xlsx nebo Google Sheets).
 *
 * Postup: najde nejnovější tabulkový soubor v zadané Drive složce → zkopíruje jako Google Sheet
 * (u .xlsx tím proběhne konverze, u už existujícího Sheets souboru jde o obyčejnou kopii)
 * → přečte listy dle konfigurace → porovná s DB → provede INSERT/UPDATE/deaktivaci.
 */

const STORES_COL_MAP = {
  'Č.prodejny':       'code',
  'Název':            'name',
  'LC':               'lc_code',
  'Telefon prodejny': 'phone',
  'VT':               'area_manager',
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
};

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

    const file = findSyncFileInFolder_(folderId);
    if (!file) { autoSyncNoteCheck_('ve složce není žádný soubor .xlsx ani Google Sheets'); return; }

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
    audit_('sync_run_auto_error', String(e && e.message ? e.message : e));
    throw e; // necháme GAS poslat vlastníkovi e-mail o selhání triggeru
  }
}

/* ── Interní funkce ───────────────────────────────────────────── */

/**
 * Zapíše kompaktní záznam o proběhlé synchronizaci do _settings.syncHistory
 * (posledních 20 běhů — kdo, kdy, soubor, počty). Detail změn drží jen
 * poslední běh (lastSyncResult), historie je jen souhrn.
 */
function appendSyncHistory_(settings, result, isAuto) {
  let history = [];
  try { history = settings.syncHistory ? JSON.parse(settings.syncHistory) : []; } catch (e) { history = []; }
  const s = result.stores || {};
  history.unshift({
    at: nowIso_(),
    by: currentEmail_() || 'system',
    auto: isAuto === true,
    file: result.fileName || '',
    added: s.added || 0,
    updated: s.updated || 0,
    deactivated: s.deactivated || 0,
    reactivated: s.reactivated || 0,
    unchanged: s.unchanged || 0,
    errors: (s.errors || []).length,
  });
  if (history.length > 20) history = history.slice(0, 20);
  settingsSet_('syncHistory', JSON.stringify(history));
}

/** Jádro synchronizace sdílené ruční (apiRunSync) i automatickou (autoSyncCheck_) cestou. */
function runSyncCore_(settings, isAuto) {
  const folderUrl = settings.syncFolderUrl || '';
  if (!folderUrl) throw new Error('Není nastavena URL složky. Vyplňte ji v sekci Synchronizace.');

  const folderId = extractFolderIdFromUrl_(folderUrl);
  if (!folderId) throw new Error('Z URL složky se nepodařilo rozpoznat ID. Použijte URL ve tvaru https://drive.google.com/drive/folders/...');

  const xlsxFile = findSyncFileInFolder_(folderId);
  if (!xlsxFile) throw new Error('Ve složce nebyl nalezen žádný soubor .xlsx ani Google Sheets.');

  let ss;
  let tempSheetId = null;
  try {
    const scriptFolder = scriptFolder_();
    const copyMeta = { name: '__sync_tmp__', mimeType: 'application/vnd.google-apps.spreadsheet' };
    if (scriptFolder) copyMeta.parents = [scriptFolder.getId()];
    const copy = Drive.Files.copy(copyMeta, xlsxFile.getId());
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
    if (tempSheetId) {
      try { Drive.Files.remove(tempSheetId); } catch (_) {}
    }
  }

  settingsSet_('lastSyncAt', nowIso_());
  settingsSet_('lastSyncResult', JSON.stringify(result));
  settingsSet_('syncLastFileSignature', xlsxFile.getId() + ':' + xlsxFile.getLastUpdated().getTime());
  appendSyncHistory_(settings, result, isAuto);
  return result;
}

function syncStores_(ss, settings) {
  dbEnsureSchema_(dbSpreadsheet_());

  const mainSheetName = settings.syncStoresSheet || 'VTBZL_export';
  const tempPrefix    = settings.syncTempClosedPrefix || 'Dočasné zavření';

  const sheet1 = ss.getSheetByName(mainSheetName);
  if (!sheet1) throw new Error('List "' + mainSheetName + '" nebyl v souboru nalezen.');

  const sheet2 = ss.getSheets().find((s) => s.getName().startsWith(tempPrefix)) || null;

  const mainRows = parseSheetRows_(sheet1, STORES_COL_MAP);
  const tempRows = sheet2 ? parseSheetRows_(sheet2, STORES_COL_MAP) : [];

  // Mapa: code → data (hlavní list je autoritativní zdroj dat)
  const xlsxMap = new Map(mainRows.map((r) => [r.code, r]));

  // Filiálky jen na listu "Dočasné zavření" (nejsou v hlavním listu)
  tempRows.forEach((r) => { if (!xlsxMap.has(r.code)) xlsxMap.set(r.code, r); });

  const currentRecords = dbGetAll_(SHEETS.STORES);

  const CHANGES_LIMIT = 50;
  const stats = { added: 0, updated: 0, deactivated: 0, reactivated: 0, unchanged: 0, errors: [],
                  changes: { added: [], updated: [], deactivated: [], reactivated: [] } };
  const now = nowIso_();
  const newRecords = [];

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
    } else {
      const xlsxRow = xlsxMap.get(codeKey);
      const patch = buildStorePatch_(xlsxRow, now, existing);

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
          newRecords.push(existing);
          stats.unchanged++;
        }
      }
      xlsxMap.delete(codeKey);
    }
  });

  // Nové záznamy z xlsx (nezpracované = nebyly v DB)
  xlsxMap.forEach((xlsxRow, code) => {
    newRecords.push(Object.assign(buildStorePatch_(xlsxRow, now), {
      id: uuid_(),
      created_at: now,
      created_by: currentEmail_() || 'sync',
      synced_at: now,
    }));
    stats.added++;
    if (stats.changes.added.length < CHANGES_LIMIT)
      stats.changes.added.push({ code, name: xlsxRow.name || '' });
  });

  dbBatchReplace_(SHEETS.STORES, newRecords);
  return stats;
}

/* ── Pomocné funkce ───────────────────────────────────────────── */

const HOUR_FIELDS_ = [
  'mon_open','mon_close','tue_open','tue_close','wed_open','wed_close',
  'thu_open','thu_close','fri_open','fri_close','sat_open','sat_close','sun_open','sun_close',
];

function buildStorePatch_(xlsxRow, now, existing) {
  const NON_HOUR_FIELDS = ['code','name','lc_code','phone','area_manager','regional_manager','rm_phone'];
  // temporarily_closed je computed z temp_closed_ranges — sync ho nepřebírá z xlsx
  const patch = { temporarily_closed: existing ? isTempClosedNow_(existing) : false, active: true, synced_at: now, updated_at: now };
  NON_HOUR_FIELDS.forEach((f) => { patch[f] = xlsxRow[f] !== undefined ? xlsxRow[f] : ''; });
  // Otevírací doby: přepsat jen pokud xlsx má hodnotu NEBO DB ji dosud nemá
  HOUR_FIELDS_.forEach((f) => {
    const xlsxVal = xlsxRow[f] !== undefined ? xlsxRow[f] : '';
    const dbVal = existing ? (existing[f] || '') : '';
    patch[f] = xlsxVal || dbVal;
  });
  return patch;
}

const STORE_DIFF_FIELDS = [
  'name','lc_code','phone','area_manager','regional_manager','rm_phone',
  'mon_open','mon_close','tue_open','tue_close','wed_open','wed_close',
  'thu_open','thu_close','fri_open','fri_close','sat_open','sat_close','sun_open','sun_close',
];

const STORE_FIELD_LABELS = {
  name: 'Název', lc_code: 'LC', phone: 'Telefon prodejny',
  area_manager: 'VT', regional_manager: 'RM', rm_phone: 'Telefon RM',
  mon_open: 'Po otevřeno', mon_close: 'Po zavřeno',
  tue_open: 'Út otevřeno', tue_close: 'Út zavřeno',
  wed_open: 'St otevřeno', wed_close: 'St zavřeno',
  thu_open: 'Čt otevřeno', thu_close: 'Čt zavřeno',
  fri_open: 'Pá otevřeno', fri_close: 'Pá zavřeno',
  sat_open: 'So otevřeno', sat_close: 'So zavřeno',
  sun_open: 'Ne otevřeno', sun_close: 'Ne zavřeno',
};

function storeChangedFields_(existing, patch) {
  const result = [];
  STORE_DIFF_FIELDS.forEach((f) => {
    const oldVal = String(existing[f] || '');
    const newVal = String(patch[f] || '');
    if (oldVal !== newVal)
      result.push({ field: STORE_FIELD_LABELS[f] || f, old: oldVal, new: newVal });
  });
  return result;
}

/**
 * Přečte list tabulky a vrátí pole objektů namapovaných přes colMap.
 * Záhlaví je na řádku 1 (trimované). Prázdné řádky (bez code) jsou přeskočeny.
 */
function parseSheetRows_(sheet, colMap) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sheet.getLastColumn();
  const rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headers = rawHeaders.map((h) => String(h).trim());

  // Index každého cílového pole
  const colIndices = {};
  Object.keys(colMap).forEach((xlsxHeader) => {
    const idx = headers.indexOf(xlsxHeader);
    if (idx !== -1) colIndices[colMap[xlsxHeader]] = idx;
  });

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data
    .map((row) => {
      const record = {};
      Object.keys(colIndices).forEach((dbField) => {
        record[dbField] = formatCellValue_(row[colIndices[dbField]]);
      });
      return record;
    })
    .filter((r) => r.code);  // přeskočit řádky bez kódu
}

/**
 * Převede hodnotu buňky na string.
 * Časové buňky (h:mm) GAS vrací jako Date s datem 30.12.1899 — formátujeme jako "H:mm".
 */
function formatCellValue_(val) {
  if (val instanceof Date) {
    if (val.getFullYear() === 1899 && val.getMonth() === 11 && val.getDate() === 30) {
      const h = val.getHours();
      const m = val.getMinutes();
      return h + ':' + (m < 10 ? '0' + m : m);
    }
    return '';
  }
  const str = (val !== undefined && val !== null) ? String(val).trim() : '';
  // Normalizace časového formátu "07:00" → "7:00" (h:mm)
  const timeMatch = str.match(/^(\d{1,2}):(\d{2})$/);
  if (timeMatch) return parseInt(timeMatch[1], 10) + ':' + timeMatch[2];
  return str;
}

/** Extrahuje ID složky z Google Drive URL. */
function extractFolderIdFromUrl_(url) {
  const match = String(url).match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/** Vrátí nejnovější tabulkový soubor (.xlsx nebo Google Sheets) ve složce nebo null. */
function findSyncFileInFolder_(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    let newest = null;
    let newestDate = null;
    [MimeType.MICROSOFT_EXCEL, MimeType.GOOGLE_SHEETS].forEach((mimeType) => {
      const files = folder.getFilesByType(mimeType);
      while (files.hasNext()) {
        const file = files.next();
        const date = file.getLastUpdated();
        if (!newestDate || date > newestDate) { newest = file; newestDate = date; }
      }
    });
    return newest;
  } catch (e) {
    throw new Error('Nepodařilo se otevřít složku: ' + e.message);
  }
}
