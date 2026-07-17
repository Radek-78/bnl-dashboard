/**
 * Identita a oprávnění.
 *
 * Aplikace je nasazena jako "Execute as me" — uživatelé nepotřebují přístup
 * k DB spreadsheetu, o přístupu rozhoduje výhradně tabulka _users.
 */
function currentEmail_() {
  return Session.getActiveUser().getEmail() || '';
}

/** Vrátí aktivního uživatele z _users, nebo null (bez přístupu). */
function getCurrentUser_() {
  const email = currentEmail_();
  if (!email) return null;
  const user = dbGetAll_(SHEETS.USERS).find(
    (u) => String(u.email).toLowerCase() === email.toLowerCase() && u.active === true
  ) || null;

  if (user) {
    // Výchozí hodnoty pro zachování přístupu stávajících uživatelů
    if (!user.permission) user.permission = 'EDITOR';
    if (!user.location) user.location = 'HQ';
  }
  return user;
}

/** Inicializuje výchozí hodnoty rolí v tabulce _role_permissions, pokud je prázdná. */
function dbEnsureRolePermissions_() {
  dbEnsureApps_();
  const list = dbGetAll_(SHEETS.ROLE_PERMISSIONS);
  if (list.length > 0) return;

  const defaultRoles = [
    {
      role: 'SUPERADMIN',
      stores_read: true,
      stores_write: true,
      logistics_read: true,
      logistics_write: true,
      users_manage: true,
      settings_manage: true,
      allowed_apps: '*',
      updated_at: new Date().toISOString(),
      updated_by: 'system'
    },
    {
      role: 'ADMIN',
      stores_read: true,
      stores_write: true,
      logistics_read: true,
      logistics_write: true,
      users_manage: true,
      settings_manage: false,
      allowed_apps: '*',
      updated_at: new Date().toISOString(),
      updated_by: 'system'
    },
    {
      role: 'USER',
      stores_read: true,
      stores_write: true,
      logistics_read: true,
      logistics_write: false,
      users_manage: false,
      settings_manage: false,
      allowed_apps: '*',
      updated_at: new Date().toISOString(),
      updated_by: 'system'
    }
  ];

  defaultRoles.forEach((r) => dbInsert_(SHEETS.ROLE_PERMISSIONS, r));
}

/**
 * Načte oprávnění pro konkrétní roli z DB s fallbackem pro SUPERADMINA.
 */
function getRolePermissions_(roleName) {
  dbEnsureRolePermissions_();
  const nameClean = String(roleName || '').trim().toUpperCase();
  
  if (nameClean === 'SUPERADMIN') {
    return {
      role: 'SUPERADMIN',
      stores_read: true, stores_write: true,
      logistics_read: true, logistics_write: true,
      users_manage: true, settings_manage: true,
      allowed_apps: '*'
    };
  }
  
  const roles = dbGetAll_(SHEETS.ROLE_PERMISSIONS);
  const found = roles.find(r => String(r.role).trim().toUpperCase() === nameClean);
  
  return found || {
    role: nameClean,
    stores_read: false, stores_write: false,
    logistics_read: false, logistics_write: false,
    users_manage: false, settings_manage: false,
    allowed_apps: ''
  };
}

/**
 * Centrální vyhodnocení přístupových práv uživatele.
 * 
 * @param {Object} user Aktivní uživatel z DB
 * @param {string} permissionKey Název oprávnění (např. 'stores_write', 'users_manage' nebo role 'ADMIN')
 * @param {string} [resourceLcCode] LC kód pro kontrolu datového rozsahu (např. 'BNL')
 * @returns {boolean}
 */
function isAllowed_(user, permissionKey, resourceLcCode) {
  if (!user || user.active !== true) return false;

  // Superadmin má vždy plná práva ke všemu
  if (user.role === 'SUPERADMIN') return true;

  const permSet = getRolePermissions_(user.role);
  const key = String(permissionKey).trim();

  // Celoplošná oprávnění — bez vazby na konkrétní lokaci
  if (key === 'USER') return true;
  if (key === 'ADMIN' || key === 'users_manage') {
    return permSet.users_manage === true || String(permSet.users_manage) === 'true';
  }
  if (key === 'SUPERADMIN' || key === 'settings_manage') {
    return permSet.settings_manage === true || String(permSet.settings_manage) === 'true';
  }

  // Oprávnění vázaná na datový rozsah (filiálky/LC) — nejdřív základní právo z matice role
  let baseAllowed;
  if (key === 'stores_read') {
    baseAllowed = permSet.stores_read === true || String(permSet.stores_read) === 'true';
  } else if (key === 'stores_write') {
    const baseWrite = permSet.stores_write === true || String(permSet.stores_write) === 'true';
    baseAllowed = baseWrite && user.permission === 'EDITOR';
  } else if (key === 'logistics_read') {
    baseAllowed = permSet.logistics_read === true || String(permSet.logistics_read) === 'true';
  } else if (key === 'logistics_write') {
    const baseWrite = permSet.logistics_write === true || String(permSet.logistics_write) === 'true';
    baseAllowed = baseWrite && user.permission === 'EDITOR';
  } else {
    return false; // neznámý klíč oprávnění
  }
  if (!baseAllowed) return false;

  // Kontrola lokace (Row-Level Security) — jen když je zadán konkrétní záznam
  if (!resourceLcCode) return true;
  const userLoc = String(user.location || 'HQ').trim().toUpperCase();
  if (userLoc === 'HQ' || userLoc === 'CENTRÁLA' || userLoc === 'CENTRAL') {
    return true; // Uživatel z centrály vidí vše
  }
  return userLoc === String(resourceLcCode).trim().toUpperCase();
}

/**
 * Obal pro každý veřejný endpoint: ověří inicializaci, uživatele a oprávnění,
 * spustí fn(uživatel) a výsledek či chybu zabalí do jednotné obálky.
 */
function guard_(permissionKey, fn, options) {
  try {
    if (!isSetupDone_()) throw new Error('Aplikace není inicializována.');
    const user = getCurrentUser_();
    if (!user) throw new Error('Nemáte přístup do aplikace. Kontaktujte správce.');
    
    const opts = options || {};
    const checkKey = opts.permission || permissionKey;
    
    if (!isAllowed_(user, checkKey)) {
      throw new Error('Na tuto akci nemáte dostatečné oprávnění.');
    }
    return ok_(fn(user));
  } catch (e) {
    console.error(e);
    return fail_(e);
  }
}

/**
 * Ověří, zda má uživatel oprávnění k přístupu do dané subaplikace (modulu).
 * 
 * @param {Object} user Aktivní uživatel z DB
 * @param {Object} app Subaplikace z DB
 * @returns {boolean}
 */
function isAppAllowedForUser_(user, app) {
  if (!user || user.active !== true) return false;
  if (user.role === 'SUPERADMIN') return true;
  
  // 1. Zkontrolujeme specifické nastavení uživatele (pokud je sloupec vyplněný a není 'default')
  let allowed = String(user.allowed_apps || '').trim();
  if (allowed && allowed.toLowerCase() !== 'default') {
    if (allowed === '*') return true;
    const allowedList = allowed.split(',').map(s => s.trim().toLowerCase());
    const slug = String(app.slug || '').trim().toLowerCase();
    const id = String(app.id || '').trim().toLowerCase();
    return allowedList.includes(slug) || allowedList.includes(id);
  }
  
  // 2. Fallback na nastavení role
  const permSet = getRolePermissions_(user.role);
  const roleAllowed = String(permSet.allowed_apps || '').trim();
  if (roleAllowed === '*') return true;
  if (!roleAllowed) return false;
  
  const roleAllowedList = roleAllowed.split(',').map(s => s.trim().toLowerCase());
  const slug = String(app.slug || '').trim().toLowerCase();
  const id = String(app.id || '').trim().toLowerCase();
  return roleAllowedList.includes(slug) || roleAllowedList.includes(id);
}
