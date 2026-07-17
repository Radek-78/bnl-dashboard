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
