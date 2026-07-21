/**
 * Pomocné funkce: jednotná odpovědní obálka, identifikátory, audit.
 *
 * Každý veřejný endpoint vrací { ok: true, data } nebo { ok: false, error },
 * frontend tak nikdy nedostane syrovou výjimku.
 */
function ok_(data) {
  return { ok: true, data: data === undefined ? null : data };
}

function fail_(message) {
  return { ok: false, error: String(message && message.message ? message.message : message) };
}

function userDisplayName_(user) {
  return [user.firstName, user.lastName].filter(Boolean).join(' ').trim() || user.email || '';
}

function uuid_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}

/**
 * Nastaví na celý (nově vytvořený) list stejný font jako appka (CONFIG.sheetFont)
 * - volá se hned po insertSheet()/vytvoření spreadsheetu, ať listy vypadají
 * konzistentně s appkou samotnou, ne v defaultním fontu Sheets. Přes celý
 * dostupný rozsah listu (ne jen vyplněné buňky), ať font sedí i na buňky,
 * které se teprve vyplní.
 */
function applySheetFont_(sheet) {
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily(CONFIG.sheetFont);
  } catch (e) {
    console.error('Nastavení fontu listu selhalo: ' + e);
  }
}

/**
 * Zapíše záznam do auditního logu. Selhání auditu nesmí shodit hlavní operaci.
 */
function audit_(action, detail) {
  try {
    dbAppendRow_(SHEETS.AUDIT, {
      timestamp: nowIso_(),
      user: currentEmail_() || 'system',
      action: action,
      detail: detail || '',
    });
  } catch (e) {
    console.error('Zápis do auditu selhal: ' + e);
  }
}
