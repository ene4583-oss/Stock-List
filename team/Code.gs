function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Smart Stock')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function start(token) {
  const db = database_();
  const actor = principal_(db, undefined, token);
  const deleted = deletedCatalogIds_(db, 'FACTORIES');
  const factories = factoriesFor_(db, actor, deleted).map(row => ({id:row[0], name:row[1]}));
  return {name:actor.name, role:actor.role, factories,
    ...(actor.role === 'ADMIN' ? {archivedFactories:rows_(db, 'FACTORIES').filter(row => deleted.has(row[0]))
      .map(row => ({id:row[0], name:row[1]}))} : {})};
}

const SCHEMA = Object.freeze({
  FACTORIES: 'Factory_ID Factory_Name Active Created_At Updated_At',
  USERS: 'User_ID Name Email Auth_Subject Role Active Created_At Updated_At',
  USER_FACTORIES: 'User_ID Factory_ID',
  ITEMS: 'Item_ID Item_Name Category_ID Unit Pack_Size Allow_Decimal Image_File_ID Supplier Last_Price Active Created_At Updated_At',
  STOCK: 'Factory_ID Item_ID Current_Stock Initial_Daily_Usage Lead_Time_Days Review_Period_Days Safety_Days Last_Count_Date Active Updated_At',
  STOCK_LOG: 'Log_ID Request_ID Timestamp Usage_Date Factory_ID Item_ID User_ID Type Record_Mode Quantity Stock_Before Stock_After Reason Is_Consumption Reference_ID',
  DAILY_COVERAGE: 'Coverage_ID Factory_ID Item_ID Usage_Date Coverage_Status Confirmed_By Confirmed_At Notes',
  PURCHASE_ORDERS: 'Order_ID Order_Line_ID Factory_ID Item_ID Supplier Order_Date Expected_Date Ordered_Qty Received_Qty Unit_Price Status Completed_Date Created_By Updated_At',
  FORECAST: 'Factory_ID Item_ID Usage_7D Usage_28D Reliable_Days Average_Daily_Usage Average_Weekly_Usage Usage_Trend Current_Stock Incoming_Stock Days_Remaining Safety_Stock Reorder_Point Target_Stock Suggested_Order Stock_Status Data_Confidence Calculated_At',
  CATEGORIES: 'Category_ID Category_Name Active Updated_At',
  SETTINGS: 'Key Value Updated_By Updated_At',
  LABELS: 'Label_Key Language Display_Text Updated_At',
  ALERT_LOG: 'Alert_ID Alert_Key Factory_ID Item_ID User_ID Alert_Type Message Channel Status Created_At Sent_At Error_Message',
  REQUEST_LOG: 'Request_ID User_ID Operation Payload_Hash Status Reference_ID Created_At Completed_At Error_Message'
});

const DEFAULTS = Object.freeze({
  APP_NAME: 'Smart Stock', COMPANY_NAME: 'My Company', DEFAULT_LANGUAGE: 'EN',
  TIMEZONE: 'Asia/Kuala_Lumpur', FORECAST_WINDOW: 28, DEFAULT_LEAD_TIME: 3,
  DEFAULT_REVIEW_PERIOD: 7, DEFAULT_SAFETY_DAYS: 3, DAILY_ALERT_TIME: '08:00',
  TELEGRAM_ENABLED: false, UNUSUAL_USAGE_ENABLED: false, LARGE_USAGE_APPROVAL_ENABLED: false,
  BACKUP_ENABLED: true, WORKER_STOCK_VISIBILITY: false
});

function properties_() {
  return PropertiesService.getScriptProperties();
}

function database_() {
  const id = properties_().getProperty('DATABASE_ID');
  if (!id) throw new Error('Setup required.');
  return SpreadsheetApp.openById(id);
}

function width_(name) {
  return SCHEMA[name].split(' ').length;
}

function rows_(db, name) {
  const sheet = db.getSheetByName(name);
  const count = sheet.getLastRow() - 1;
  return count > 0 ? sheet.getRange(2, 1, count, width_(name)).getValues() : [];
}

function matchingRows_(db, name, column, value, limit) {
  const sheet = db.getSheetByName(name);
  const count = sheet.getLastRow() - 1;
  if (count < 1) return [];
  return sheet.getRange(2, column, count, 1).createTextFinder(value)
    .matchEntireCell(true).matchCase(true).findAll().slice(-limit).reverse()
    .map(cell => sheet.getRange(cell.getRow(), 1, 1, width_(name)).getValues()[0]);
}

function locked_(action) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('System busy. Retry the same request.');
  try {
    return action();
  } finally {
    lock.releaseLock();
  }
}

function cells_(row) {
  return {values: row.map(value => {
    if (typeof value === 'number') return {userEnteredValue: {numberValue: value}};
    if (typeof value === 'boolean') return {userEnteredValue: {boolValue: value}};
    return {userEnteredValue: {stringValue: String(value ?? '')}};
  })};
}

// A single Sheets batchUpdate applies all supplied requests atomically.
function commit_(db, requests) {
  Sheets.Spreadsheets.batchUpdate({requests}, db.getId());
}

function append_(db, name, rows) {
  return {appendCells: {
    sheetId: db.getSheetByName(name).getSheetId(), rows: rows.map(cells_), fields: 'userEnteredValue'
  }};
}

function update_(db, name, rowNumber, row) {
  return {updateCells: {
    start: {sheetId: db.getSheetByName(name).getSheetId(), rowIndex: rowNumber - 1, columnIndex: 0},
    rows: [cells_(row)], fields: 'userEnteredValue'
  }};
}

function upsert_(db, name, index, row) {
  return index < 0 ? append_(db, name, [row]) : update_(db, name, index + 2, row);
}

function now_() {
  return new Date().toISOString();
}

function today_() {
  return Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyy-MM-dd');
}

function id_() {
  return Utilities.getUuid();
}

function requiredText_(value, max = 100) {
  if (typeof value !== 'string' || !value.trim() || value.trim().length > max) {
    throw new Error('Check the required fields.');
  }
  return value.trim();
}

function enabled_(value) {
  return value === true || String(value).toUpperCase() === 'TRUE';
}

/** Run as the spreadsheet owner in the Apps Script editor. */
function setupDatabase() {
  const email = Session.getEffectiveUser().getEmail().toLowerCase();
  if (!email || email !== Session.getActiveUser().getEmail().toLowerCase()) {
    throw new Error('Run setup in the editor as the spreadsheet owner.');
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active || active.getOwner()?.getEmail().toLowerCase() !== email) {
    throw new Error('An owner-controlled, bound spreadsheet is required.');
  }
  return locked_(() => {
    const configured = properties_().getProperty('DATABASE_ID');
    if (configured && configured !== active.getId()) throw new Error('Database binding already exists.');
    // Validate every existing tab before changing any.
    for (const name of Object.keys(SCHEMA)) {
      const sheet = active.getSheetByName(name);
      if (!sheet || !sheet.getLastRow()) continue;
      const expected = SCHEMA[name].split(' ');
      const headers = sheet.getRange(1, 1, 1, Math.max(expected.length, sheet.getLastColumn())).getValues()[0];
      if (expected.some((header, index) => headers[index] !== header) ||
          headers.slice(expected.length).some(Boolean)) {
        throw new Error('Incompatible headers in ' + name + '. Back up and migrate explicitly.');
      }
    }
    const existingUsers = active.getSheetByName('USERS') ? rows_(active, 'USERS') : [];
    if (existingUsers.length && !existingUsers.some(user =>
      user[2] === email && user[4] === 'ADMIN' && enabled_(user[5]))) {
      throw new Error('The existing USERS sheet has no active administrator for this sheet owner.');
    }
    for (const name of Object.keys(SCHEMA)) {
      const sheet = active.getSheetByName(name) || active.insertSheet(name);
      if (sheet.getLastRow()) continue;
      sheet.getRange(1, 1, 1, width_(name)).setValues([SCHEMA[name].split(' ')])
        .setFontWeight('bold').setBackground('#DBEAFE');
      sheet.setFrozenRows(1);
    }
    SpreadsheetApp.flush();
    const users = rows_(active, 'USERS');
    const settings = new Set(rows_(active, 'SETTINGS').map(row => row[0]));
    const timestamp = now_();
    const requests = [];
    const missing = Object.entries(DEFAULTS).filter(([key]) => !settings.has(key));
    if (missing.length) requests.push(append_(active, 'SETTINGS', missing.map(
      ([key, value]) => [key, value, email, timestamp]
    )));
    if (!users.length) requests.push(append_(active, 'USERS', [
      [id_(), 'Administrator', email, '', 'ADMIN', true, timestamp, timestamp]
    ]));
    if (requests.length) commit_(active, requests);
    const settingsToSave = {DATABASE_ID: active.getId(), ADMIN_EMAIL: email};
    if (!properties_().getProperty('AUTH_KEY')) settingsToSave.AUTH_KEY = id_() + id_();
    properties_().setProperties(settingsToSave);
  });
}

function username_(value) {
  return typeof value === 'string' ? value.normalize('NFKC').trim().toLowerCase() : '';
}

function mac_(value) {
  const key = properties_().getProperty('AUTH_KEY');
  if (!key) throw new Error('Setup required.');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(value, key));
}

function credential_(id, password) {
  return mac_('password|' + id + '|' + password);
}

// Public deployment salt. The signing key remains in Script Properties.
function passwordSalt() {
  return mac_('smart-stock-password-salt-v2');
}

function passwordProof_(value) {
  return typeof value === 'string' && /^v2:[0-9a-f]{64}$/.test(value);
}

function savedPassword_(id, proof) {
  return 'v2:' + mac_('password-v2|' + id + '|' + proof);
}

function same_(left, right) {
  if (typeof left !== 'string' || typeof right !== 'string' || left.length !== right.length) return false;
  let diff = 0;
  for (let i = 0; i < left.length; i++) diff |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return diff === 0;
}

function owner_() {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  return email && email === properties_().getProperty('ADMIN_EMAIL') &&
    email === Session.getEffectiveUser().getEmail().trim().toLowerCase() ? email : '';
}

function sessionUser_(db, token) {
  if (typeof token !== 'string' || !/^[0-9a-f-]{36}\.[0-9a-f]{64}$/.test(token)) return null;
  const id = token.split('.')[0];
  const row = rows_(db, 'USERS').find(user => user[0] === id);
  if (!row || !['WORKER', 'SUPERVISOR', 'ADMIN'].includes(row[4]) || !enabled_(row[5])) return null;
  const saved = properties_().getProperty('SESSION_' + id);
  if (!saved) return null;
  const [expires, verifier] = saved.split('|');
  return Number(expires) > Date.now() && same_(verifier, mac_('session|' + token + '|' + row[3] + '|' + row[7])) ? row : null;
}

function principal_(db, roles, token) {
  const owner = !token && owner_();
  const row = owner ? rows_(db, 'USERS').find(user =>
    user[2] === owner && user[4] === 'ADMIN' && enabled_(user[5])) : sessionUser_(db, token);
  if (!row || (roles && !roles.includes(row[4]))) throw new Error('Access denied.');
  const factories = row[4] === 'ADMIN' ? null : new Set(
    rows_(db, 'USER_FACTORIES').filter(assignment => assignment[0] === row[0]).map(assignment => assignment[1])
  );
  return {id: row[0], name: row[1], role: row[4], factories};
}

function throttle_(key, limit, minutes) {
  const properties = properties_();
  const now = Date.now();
  let [until, count] = (properties.getProperty(key) || '0|0').split('|').map(Number);
  if (until <= now) {
    until = now + minutes * 60000;
    count = 0;
  }
  if (count >= limit) throw new Error('Login failed.');
  properties.setProperty(key, until + '|' + (count + 1));
}

function signIn(username, proof, legacyPassword) {
  const name = username_(username);
  const legacy = passwordProof_(proof) ? legacyPassword : proof;
  const hasLegacy = typeof legacy === 'string' && /^[0-9a-f]{32}$/.test(legacy);
  if (!name || (!passwordProof_(proof) && !hasLegacy)) throw new Error('Login failed.');
  return locked_(() => {
    throttle_('LOGIN_GLOBAL', 100, 5);
    const db = database_();
    const row = rows_(db, 'USERS').find(user =>
      username_(user[2]) === name && ['WORKER', 'SUPERVISOR', 'ADMIN'].includes(user[4]) && Boolean(user[3]) && enabled_(user[5]));
    if (!row) throw new Error('Login failed.');
    const id = row[0];
    throttle_('LOGIN_' + id, 10, 15);
    const stored = String(row[3]);
    const candidate = stored.startsWith('v2:')
      ? (passwordProof_(proof) ? savedPassword_(id, proof) : '')
      : (hasLegacy ? credential_(id, legacy) : '');
    if (!candidate || !same_(stored, candidate)) throw new Error('Login failed.');
    const token = id + '.' + (id_() + id_()).replace(/-/g, '');
    properties_().setProperty('SESSION_' + id,
      (Date.now() + 8 * 60 * 60000) + '|' + mac_('session|' + token + '|' + row[3] + '|' + row[7]));
    properties_().deleteProperty('LOGIN_' + id);
    return token;
  });
}

function signOut(token) {
  return locked_(() => {
    const user = sessionUser_(database_(), token);
    if (user) properties_().deleteProperty('SESSION_' + user[0]);
    return true;
  });
}

const DELETION_PREFIX = Object.freeze({FACTORIES:'FACTORY_DELETED_', ITEMS:'ITEM_DELETED_', USERS:'USER_DELETED_'});

function deletedCatalogIds_(db, kind) {
  const prefix = DELETION_PREFIX[kind];
  return new Set(rows_(db, 'SETTINGS').filter(row => String(row[0]).startsWith(prefix) && enabled_(row[1]))
    .map(row => String(row[0]).slice(prefix.length)));
}

function factoriesFor_(db, user, deleted = deletedCatalogIds_(db, 'FACTORIES')) {
  return rows_(db, 'FACTORIES').filter(row => !deleted.has(row[0]) &&
    (!user.factories || user.factories.has(row[0])));
}

function requireFactory_(db, user, id) {
  if (!factoriesFor_(db, user).some(row => row[0] === id)) throw new Error('Access denied.');
}

function saveCatalog(kind, input, token) {
  if (!['FACTORIES', 'ITEMS', 'USERS', 'STOCK'].includes(kind) || !input) throw new Error('Invalid request.');
  return locked_(() => {
    const db = database_();
    const actor = principal_(db, ['ADMIN'], token);
    const entries = rows_(db, kind);
    const timestamp = now_();
    let index = entries.findIndex(row => row[0] === input.id);
    if (kind !== 'STOCK' && input.id && index < 0) throw new Error('Record unavailable.');
    const old = entries[index];
    if (old && Object.hasOwn(DELETION_PREFIX, kind) && deletedCatalogIds_(db, kind).has(old[0])) {
      throw new Error('Record unavailable.');
    }
    const id = old ? old[0] : id_();
    let row;
    const requests = [];
    if (kind === 'FACTORIES') {
      row = [id, requiredText_(input.name), true, old?.[3] || timestamp, timestamp];
      const change = forecastSetting_(db, 'FACTORY_WORKERS_' + id, input.workerCount, actor.id, timestamp, true);
      if (change) requests.push(change);
    } else if (kind === 'ITEMS') {
      if (!Number.isFinite(input.pack) || input.pack <= 0 ||
          (!input.decimal && !Number.isInteger(input.pack))) throw new Error('Invalid pack size.');
      if (!input.decimal && rows_(db, 'STOCK').some(stock => stock[1] === id && !Number.isInteger(stock[2]))) {
        throw new Error('An existing inventory balance uses decimals.');
      }
      const name = requiredText_(input.name);
      const unit = requiredText_(input.unit, 30);
      const change = forecastSetting_(db, 'ITEM_USAGE_PER_WORKER_' + id, input.usagePerWorker, actor.id, timestamp, false);
      if (change) requests.push(change);
      if (input.itemFactories !== undefined) {
        requests.push(...itemAssignments_(db, actor, id, input.itemFactories, timestamp));
      }
      if (input.itemCode !== undefined) {
        const code = String(input.itemCode).trim() ? requiredText_(input.itemCode, 80) : '';
        const key = 'ITEM_CODE_' + id;
        const settings = rows_(db, 'SETTINGS');
        if (code && settings.some(entry => String(entry[0]).startsWith('ITEM_CODE_') && entry[0] !== key &&
            String(entry[1]).toLowerCase() === code.toLowerCase())) throw new Error('Item code already exists.');
        const index = settings.findIndex(entry => entry[0] === key);
        if ((settings[index]?.[1] ?? '') !== code) requests.push(upsert_(db, 'SETTINGS', index, [key, code, actor.id, timestamp]));
      }
      const threshold = forecastSetting_(db, 'ITEM_MINIMUM_' + id, input.minimum, actor.id, timestamp, false);
      if (threshold) requests.push(threshold);
      const image = input.photo ? storeItemImage_(input.photo) : old?.[6] || '';
      row = [id, name, old?.[2] || '', unit,
        input.pack, input.decimal === true, image, old?.[7] || '', old?.[8] ?? '',
        input.active === undefined ? (!old || enabled_(old[9])) : input.active === true, old?.[10] || timestamp, timestamp];
    } else if (kind === 'USERS') {
      const username = username_(input.username);
      if (!username) throw new Error('Enter a username.');
      if (!['WORKER', 'SUPERVISOR', 'ADMIN'].includes(input.role) || (old && !old[3] && old[4] === 'ADMIN') || id === actor.id) {
        throw new Error('Choose a role. You cannot edit your own administrator account.');
      }
      const deletedUsers = deletedCatalogIds_(db, 'USERS');
      if (entries.some(other => other[0] !== id && !deletedUsers.has(other[0]) && username_(other[2]) === username)) {
        throw new Error('Username already exists. Select that user to edit or reset their password.');
      }
      if (input.passwordProof !== undefined && !passwordProof_(input.passwordProof)) {
        throw new Error('Enter a password.');
      }
      if ((!old || (username !== username_(old[2]) && String(old[3]).startsWith('v2:'))) && !input.passwordProof) {
        throw new Error(old ? 'Enter a new password when changing the username.' : 'Enter a password.');
      }
      const verifier = input.passwordProof ? savedPassword_(id, input.passwordProof) : old[3];
      const factoryIds = [...new Set(input.factories || [])];
      const valid = new Set(factoriesFor_(db, actor).map(factory => factory[0]));
      if (!factoryIds.every(factory => valid.has(factory)) ||
          (input.active === true && input.role !== 'ADMIN' && !factoryIds.length)) {
        throw new Error('Select assigned factories.');
      }
      row = [id, username, username, verifier, input.role,
        input.active === true, old?.[6] || timestamp, timestamp];
      const assignments = rows_(db, 'USER_FACTORIES');
      const sheetId = db.getSheetByName('USER_FACTORIES').getSheetId();
      for (let offset = assignments.length - 1; offset >= 0; offset--) {
        if (assignments[offset][0] === id) {
          requests.push({deleteDimension: {range: {sheetId, dimension: 'ROWS', startIndex: offset + 1, endIndex: offset + 2}}});
        }
      }
      if (factoryIds.length) requests.push(append_(db, 'USER_FACTORIES', factoryIds.map(factory => [id, factory])));
    } else {
      requireFactory_(db, actor, input.factory);
      if (!rows_(db, 'ITEMS').some(item => item[0] === input.item && enabled_(item[9]))) {
        throw new Error('Item unavailable.');
      }
      index = entries.findIndex(stock => stock[0] === input.factory && stock[1] === input.item);
      row = entries[index] || [input.factory, input.item, 0, '', '', '', '', '', true, timestamp];
      row[8] = input.active === true;
      row[9] = timestamp;
    }
    requests.unshift(upsert_(db, kind, index, row));
    const settingsId = db.getSheetByName('SETTINGS').getSheetId();
    const insertions = requests.filter(request => request.appendCells?.sheetId === settingsId);
    if (insertions.length > 1) {
      insertions[0].appendCells.rows.push(...insertions.slice(1).flatMap(request => request.appendCells.rows));
    }
    commit_(db, requests.filter(request => !insertions.slice(1).includes(request)));
    return {id};
  });
}

function catalog(token) {
  const db = database_();
  const actor = principal_(db, ['ADMIN'], token);
  const assignments = rows_(db, 'USER_FACTORIES');
  const settings = new Map(rows_(db, 'SETTINGS').map(row => [row[0], row[1]]));
  return {
    FACTORIES: rows_(db, 'FACTORIES').filter(row => !enabled_(settings.get('FACTORY_DELETED_' + row[0]))).map(row => ({id: row[0], name: row[1],
      workerCount: settings.get('FACTORY_WORKERS_' + row[0]) ?? ''})),
    ITEMS: rows_(db, 'ITEMS').filter(row => !enabled_(settings.get('ITEM_DELETED_' + row[0]))).map(row => ({id: row[0], name: row[1], unit: row[3],
      pack: row[4], itemCode: settings.get('ITEM_CODE_' + row[0]) || '', minimum: settings.get('ITEM_MINIMUM_' + row[0]) ?? '', decimal: enabled_(row[5]),
      usagePerWorker: settings.get('ITEM_USAGE_PER_WORKER_' + row[0]) ?? '', hasPhoto: Boolean(row[6]), active: enabled_(row[9])})),
    USERS: rows_(db, 'USERS').filter(row => row[0] !== actor.id && (row[4] !== 'ADMIN' || Boolean(row[3])) &&
      !enabled_(settings.get('USER_DELETED_' + row[0]))).map(row => ({id: row[0], name: row[1], username: row[2],
      role: row[4], active: enabled_(row[5]), factories: assignments.filter(a => a[0] === row[0]).map(a => a[1])})),
    STOCK: rows_(db, 'STOCK').map(row => ({factory: row[0], item: row[1], active: enabled_(row[8]), estimatedDaily: row[3], quantity:row[2]}))
  };
}

function itemMap_(db) {
  return new Map(rows_(db, 'ITEMS').map(row => [row[0], row]));
}

function itemView_(item) {
  return {id: item[0], name: item[1], unit: item[3], ...(item[6] ? {hasPhoto:true} : {})};
}

function factoryItems(factory, token) {
  const db = database_();
  const actor = principal_(db, undefined, token);
  requireFactory_(db, actor, factory);
  const items = itemMap_(db);
  return rows_(db, 'STOCK').filter(row => row[0] === factory && enabled_(row[8]) &&
    items.has(row[1]) && enabled_(items.get(row[1])[9])).map(row => itemView_(items.get(row[1])));
}

function stockSummary(token) {
  const db = database_();
  const actor = principal_(db, ['ADMIN', 'SUPERVISOR'], token);
  const factories = new Map(factoriesFor_(db, actor).map(row =>
    [row[0], {id:row[0], name:row[1], items:0, empty:0}]));
  const activeItems = new Set(rows_(db, 'ITEMS').filter(row => enabled_(row[9])).map(row => row[0]));
  for (const row of rows_(db, 'STOCK')) {
    const factory = factories.get(row[0]);
    if (!factory || !enabled_(row[8]) || !activeItems.has(row[1])) continue;
    factory.items++;
    if (row[2] === 0) factory.empty++;
  }
  return [...factories.values()];
}

function stockOverview(factory, token) {
  const db = database_();
  const actor = principal_(db, ['ADMIN', 'SUPERVISOR'], token);
  requireFactory_(db, actor, factory);
  const items = itemMap_(db);
  return rows_(db, 'STOCK').filter(row => row[0] === factory && enabled_(row[8]) &&
    items.has(row[1]) && enabled_(items.get(row[1])[9]))
    .map(row => ({...itemView_(items.get(row[1])), quantity: row[2]}));
}

function history(factory, team, token) {
  const db = database_();
  const actor = principal_(db, team === true ? ['ADMIN', 'SUPERVISOR'] : undefined, token);
  if (team && actor.role === 'ADMIN') {
    if (!rows_(db, 'FACTORIES').some(row => row[0] === factory)) throw new Error('Access denied.');
  } else requireFactory_(db, actor, factory);
  const items = itemMap_(db);
  const selected = team === true ? matchingRows_(db, 'STOCK_LOG', 5, factory, 50) :
    matchingRows_(db, 'STOCK_LOG', 7, actor.id, 100).filter(row => row[4] === factory).slice(0, 50);
  if (!selected.length) return [];
  const names = team === true ? new Map(rows_(db, 'USERS').map(row => [row[0], row[1]])) : new Map();
  const reversed = new Set(matchingRows_(db, 'STOCK_LOG', 8, 'REVERSAL', Number.MAX_SAFE_INTEGER)
    .filter(row => row[4] === factory).map(row => row[14]));
  return selected.map(row => {
    const item = items.get(row[5]);
    const deleted = reversed.has(row[0]);
    const summary = {date: dateKey_(row[3]) || String(row[3]), name: item?.[1] || String(row[5]),
      unit: item?.[3] || '', quantity: row[9], status: deleted ? 'Deleted' : 'Saved'};
    if (team !== true) return summary;
    const removable = actor.role === 'ADMIN' && !deleted && ['USE', 'RECEIVE', 'OPENING'].includes(row[7]);
    return {...summary, type:row[7], user:names.get(row[6]) || row[6], before:row[10], after:row[11],
      ...(String(row[12]).startsWith('PHOTO:') ? {proofId:row[0]} : {}),
      ...(removable ? {deleteId:row[0]} : {})};
  });
}

function quantity_(value, decimal) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0 || value > 1e9 ||
      (decimal ? value < 0.000001 || Math.abs(value - round_(value)) > 1e-10 : !Number.isInteger(value))) {
    throw new Error('Check the quantity. Some items require whole numbers.');
  }
  return value;
}

function round_(value) {
  return Math.round(value * 1000000) / 1000000;
}

function postStock(input, token) {
  try {
    return postStock_(input, token);
  } catch (error) {
    const messages = [
      'Access denied.', 'Invalid request.', 'Check the required fields.', 'Select each item once.',
      'System busy. Retry the same request.', 'Request unavailable.', 'Item unavailable.',
      'Check the quantity. Some items require whole numbers.', 'Please contact your supervisor.',
      'This item already has stock history.', 'Stock limit reached.', 'Add a proof photo.',
      'Set a Stock In password in Settings first.', 'Enter the correct Stock In password.', 'Choose a valid item photo.'
    ];
    throw new Error(messages.includes(error.message) ? error.message : 'Could not save. Retry the same request.');
  }
}

function postStock_(input, token) {
  const type = input?.type;
  const roles = {USE: ['WORKER', 'SUPERVISOR', 'ADMIN'], RECEIVE: ['SUPERVISOR', 'ADMIN'], OPENING: ['ADMIN']};
  if (!Object.hasOwn(roles, type) || typeof input.request !== 'string' ||
      !/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(input.request) ||
      typeof input.factory !== 'string' || !Array.isArray(input.lines) ||
      input.lines.length < 1 || input.lines.length > 100) throw new Error('Invalid request.');
  const lines = input.lines.map(line => ({item: requiredText_(line?.item), quantity: line.quantity}))
    .sort((a, b) => a.item.localeCompare(b.item));
  if (new Set(lines.map(line => line.item)).size !== lines.length) throw new Error('Select each item once.');
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify([type, input.factory, lines, input.photo || '']), Utilities.Charset.UTF_8));
  return locked_(() => {
    const db = database_();
    const actor = principal_(db, roles[type], token);
    requireFactory_(db, actor, input.factory);
    const [prior] = matchingRows_(db, 'REQUEST_LOG', 1, input.request, 1);
    if (prior) {
      if (prior[1] !== actor.id || prior[2] !== type || prior[3] !== hash || prior[4] !== 'COMPLETED') {
        throw new Error('Request unavailable.');
      }
      return {saved: true};
    }
    if (type !== 'OPENING' && !input.photo) throw new Error('Add a proof photo.');
    if (type === 'RECEIVE') checkReceivePassword_(actor, input.receiveProof);
    const stocks = rows_(db, 'STOCK');
    const items = itemMap_(db);
    const timestamp = now_();
    const date = today_();
    const reference = id_();
    const logs = [];
    const updates = [];
    for (const line of lines) {
      const item = items.get(line.item);
      const matches = stocks.flatMap((row, index) => row[0] === input.factory && row[1] === line.item ? [index] : []);
      const index = matches[0];
      const stock = stocks[index];
      if (!item || !enabled_(item[9]) || matches.length !== 1 || !enabled_(stock[8])) {
        throw new Error('Item unavailable.');
      }
      const amount = quantity_(line.quantity, enabled_(item[5]));
      const before = stock[2];
      if (typeof before !== 'number' || !Number.isFinite(before) || before < 0) {
        throw new Error('Please contact your supervisor.');
      }
      if (type === 'OPENING' && (before !== 0 ||
          db.getSheetByName('STOCK_LOG').getLastRow() > 1 &&
          matchingRows_(db, 'STOCK_LOG', 6, line.item, Number.MAX_SAFE_INTEGER)
            .some(row => row[4] === input.factory))) throw new Error('This item already has stock history.');
      const after = round_(before + (type === 'USE' ? -amount : amount));
      if (after < 0) throw new Error('Please contact your supervisor.');
      if (after > 1e9) throw new Error('Stock limit reached.');
      stock[2] = after;
      stock[9] = timestamp;
      updates.push(update_(db, 'STOCK', index + 2, stock));
      logs.push([id_(), input.request, timestamp, date, input.factory, line.item, actor.id,
        type, type === 'USE' ? 'USE_NOW' : '', amount, before, after, '', type === 'USE', reference]);
    }
    if (input.photo) {
      const file = storeItemImage_(input.photo);
      for (const record of logs) record[12] = 'PHOTO:' + file;
    }
    commit_(db, [...updates, append_(db, 'STOCK_LOG', logs), append_(db, 'REQUEST_LOG', [
      [input.request, actor.id, type, hash, 'COMPLETED', reference, timestamp, timestamp, '']
    ])]);
    return {saved: true};
  });
}

function itemAssignments_(db, actor, itemId, selections, timestamp) {
  const allowed = new Set(factoriesFor_(db, actor).map(factory => factory[0]));
  if (!Array.isArray(selections) || selections.length > allowed.size) throw new Error('Check the factory assignments.');
  const assigned = new Map();
  for (const selection of selections) {
    if (!selection || !allowed.has(selection.factory) || assigned.has(selection.factory) ||
        (selection.estimatedDaily !== '' && (typeof selection.estimatedDaily !== 'number' ||
          !Number.isFinite(selection.estimatedDaily) || selection.estimatedDaily < 0 || selection.estimatedDaily > 1e9))) {
      throw new Error('Check the factory assignments and daily usage estimates.');
    }
    assigned.set(selection.factory, selection.estimatedDaily);
  }
  const updates = [];
  const newRows = [];
  rows_(db, 'STOCK').forEach((row, index) => {
    if (row[1] !== itemId || !allowed.has(row[0])) return;
    const active = assigned.has(row[0]);
    const estimate = active ? assigned.get(row[0]) : row[3];
    assigned.delete(row[0]);
    if (enabled_(row[8]) === active && row[3] === estimate) return;
    row[3] = estimate;
    row[8] = active;
    row[9] = timestamp;
    updates.push(update_(db, 'STOCK', index + 2, row));
  });
  for (const [factory, estimate] of assigned) {
    newRows.push([factory, itemId, 0, estimate, '', '', '', '', true, timestamp]);
  }
  if (newRows.length) updates.push(append_(db, 'STOCK', newRows));
  return updates;
}

function storeItemImage_(data) {
  if (typeof data !== 'string' || data.length > 180000 ||
      !/^data:image\/jpeg;base64,[a-z0-9+/]+={0,2}$/i.test(data)) throw new Error('Choose a valid item photo.');
  const bytes = Utilities.base64Decode(data.split(',')[1]);
  if (bytes.length < 4 || (bytes[0] & 255) !== 255 || (bytes[1] & 255) !== 216 ||
      (bytes[bytes.length - 2] & 255) !== 255 || (bytes[bytes.length - 1] & 255) !== 217) {
    throw new Error('Choose a valid item photo.');
  }
  const properties = properties_();
  const folderId = properties.getProperty('ITEM_PHOTO_FOLDER');
  const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.createFolder('Smart Stock item photos');
  if (!folderId) properties.setProperty('ITEM_PHOTO_FOLDER', folder.getId());
  return folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', id_() + '.jpg')).getId();
}

function itemImage(itemId, factory, token) {
  const db = database_();
  const actor = principal_(db, undefined, token);
  if (actor.role !== 'ADMIN' || factory) {
    requireFactory_(db, actor, factory);
    if (!rows_(db, 'STOCK').some(row => row[0] === factory && row[1] === itemId && enabled_(row[8]))) {
      throw new Error('Access denied.');
    }
  }
  const item = itemMap_(db).get(itemId);
  if (!item || !item[6] || (actor.role !== 'ADMIN' && !enabled_(item[9]))) return '';
  return storedPhoto_(item[6]);
}

function storedPhoto_(id) {
  const folderId = properties_().getProperty('ITEM_PHOTO_FOLDER');
  if (!folderId || !id) return '';
  try {
    const file = DriveApp.getFileById(id);
    const parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() !== folderId) continue;
      if (file.getMimeType() !== 'image/jpeg' || file.getSize() > 135000) return '';
      return 'data:image/jpeg;base64,' + Utilities.base64Encode(file.getBlob().getBytes());
    }
  } catch (error) {
    return '';
  }
  return '';
}

/** Run once as owner if Apps Script prompts for photo permissions after the update. */
function authorizeItemPhotos() {
  principal_(database_(), ['ADMIN']);
  return DriveApp.getRootFolder().getId() ? true : false;
}

function dateKey_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Kuala_Lumpur', 'yyyy-MM-dd');
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value)) ? String(value) : '';
}

function daysFrom_(day, amount) {
  const date = new Date(day + 'T00:00:00Z');
  date.setUTCDate(date.getUTCDate() + amount);
  return date.toISOString().slice(0, 10);
}

function within_(value, min, max, fallback) {
  const number = Number(value);
  return value !== '' && value != null && Number.isFinite(number) && number >= min && number <= max ? number : fallback;
}

function forecastSetting_(db, key, value, actor, timestamp, count) {
  if (value === undefined) return null;
  if (value !== '' && (typeof value !== 'number' || !Number.isFinite(value) ||
      value < 0 || value > 1000000 || (count && !Number.isInteger(value)))) {
    throw new Error(count ? 'Enter a whole number of workers, or leave it blank.' :
      'Enter a valid daily usage per worker, or leave it blank.');
  }
  const rows = rows_(db, 'SETTINGS');
  const index = rows.findIndex(row => row[0] === key);
  if ((rows[index]?.[1] ?? '') === value) return null;
  return upsert_(db, 'SETTINGS', index, [key, value, actor, timestamp]);
}

function suggestedBuy(token) {
  const db = database_();
  const actor = principal_(db, ['ADMIN'], token);
  const factories = new Map(factoriesFor_(db, actor).map(row => [row[0], row[1]]));
  const items = itemMap_(db);
  const settings = new Map(rows_(db, 'SETTINGS').map(row => [row[0], row[1]]));
  const windowDays = Math.floor(within_(settings.get('FORECAST_WINDOW'), 7, 90, 28));
  const today = today_();
  const start = daysFrom_(today, -windowDays);
  const key = (factory, item) => JSON.stringify([factory, item]);
  const usageDays = new Map();
  function observe(factory, item, date, amount) {
    if (!factories.has(factory) || !date || date < start || date >= today) return;
    const identity = key(factory, item);
    if (!usageDays.has(identity)) usageDays.set(identity, new Map());
    const days = usageDays.get(identity);
    days.set(date, (days.get(date) || 0) + amount);
  }
  const ledger = rows_(db, 'STOCK_LOG');
  const reversed = new Set(ledger.filter(row => row[7] === 'REVERSAL').map(row => row[14]));
  for (const row of ledger) {
    if (row[7] !== 'REVERSAL' && !reversed.has(row[0]) && enabled_(row[13]) && Number.isFinite(row[9])) {
      observe(row[4], row[5], dateKey_(row[3]), row[9]);
    }
  }
  for (const row of rows_(db, 'DAILY_COVERAGE')) {
    if (['COMPLETE', 'CONFIRMED', 'CONFIRMED_ZERO'].includes(row[4])) {
      observe(row[1], row[2], dateKey_(row[3]), 0);
    }
  }
  const orders = new Map();
  for (const row of rows_(db, 'PURCHASE_ORDERS')) {
    const due = dateKey_(row[6]);
    const open = Number(row[7]) - Number(row[8]);
    if (!['ORDERED', 'PARTIAL'].includes(row[10]) || !due || due < today || !Number.isFinite(open) || open <= 0) continue;
    const identity = key(row[2], row[3]);
    if (!orders.has(identity)) orders.set(identity, []);
    orders.get(identity).push({due, open});
  }
  const recommendations = [];
  let insufficient = 0;
  let missingWorkerCount = 0;
  for (const stock of rows_(db, 'STOCK')) {
    const [factory, itemId, quantity, estimated] = stock;
    const item = items.get(itemId);
    if (!factories.has(factory) || !enabled_(stock[8]) || !item || !enabled_(item[9])) continue;
    const identity = key(factory, itemId);
    const days = usageDays.get(identity) || new Map();
    const perWorker = within_(settings.get('ITEM_USAGE_PER_WORKER_' + itemId), 0, 1000000, null);
    const workerCount = within_(settings.get('FACTORY_WORKERS_' + factory), 0, 1000000, null);
    if (perWorker !== null && !Number.isInteger(workerCount)) {
      missingWorkerCount++;
      continue;
    }
    const historical = days.size >= 7;
    const daily = perWorker !== null ? workerCount * perWorker :
      historical ? [...days.values()].reduce((sum, used) => sum + Math.max(0, used), 0) / days.size :
        within_(estimated, 0.000001, 1e9, null);
    const basis = perWorker !== null ? workerCount + ' workers × ' + perWorker + '/day (estimate)' :
      historical ? days.size + ' observed days' : 'Estimated usage';
    if (daily === null || !Number.isFinite(quantity) || quantity < 0) {
      insufficient++;
      continue;
    }
    const lead = within_(stock[4], 0, 365, within_(settings.get('DEFAULT_LEAD_TIME'), 0, 365, 3));
    const review = within_(stock[5], 0, 365, within_(settings.get('DEFAULT_REVIEW_PERIOD'), 0, 365, 7));
    const safety = within_(stock[6], 0, 365, within_(settings.get('DEFAULT_SAFETY_DAYS'), 0, 365, 3));
    const horizon = daysFrom_(today, Math.floor(lead + review));
    const incoming = (orders.get(identity) || []).reduce((sum, order) => sum + (order.due <= horizon ? order.open : 0), 0);
    const pack = within_(item[4], 0.000001, 1e9, 1);
    const result = purchaseEstimate_(daily, quantity, incoming, pack, lead + review + safety);
    if (result.suggested === 0) continue;
    recommendations.push({...itemView_(item), factory, factoryName:factories.get(factory),
      quantity, suggested:result.suggested, daysLeft:result.daysLeft, basis});
  }
  recommendations.sort((a, b) => (a.daysLeft ?? Infinity) - (b.daysLeft ?? Infinity));
  return {recommendations, insufficient, missingWorkerCount};
}

function purchaseEstimate_(daily, stock, incoming, pack, days) {
  const target = round_(daily * days);
  const gap = Math.max(0, round_(target - stock - incoming));
  return {
    daily:round_(daily), target,
    suggested:round_(Math.ceil(gap / pack - 1e-9) * pack),
    daysLeft:daily > 0 ? Math.round(stock / daily * 10) / 10 : null
  };
}

function previewBuy(input, token) {
  const db = database_();
  principal_(db, ['ADMIN'], token);
  const {workers, perWorker, stock, incoming, pack} = input || {};
  if (![workers, perWorker, stock, incoming, pack].every(value =>
    typeof value === 'number' && Number.isFinite(value) && value >= 0 && value <= 1e6) ||
    !Number.isInteger(workers) || pack < 0.000001) throw new Error('Check the sample numbers.');
  const settings = new Map(rows_(db, 'SETTINGS').map(row => [row[0], row[1]]));
  const lead = within_(settings.get('DEFAULT_LEAD_TIME'), 0, 365, 3);
  const review = within_(settings.get('DEFAULT_REVIEW_PERIOD'), 0, 365, 7);
  const safety = within_(settings.get('DEFAULT_SAFETY_DAYS'), 0, 365, 3);
  return purchaseEstimate_(workers * perWorker, stock, incoming, pack, lead + review + safety);
}

const INTERFACE_DEFAULTS = Object.freeze({
  appName:'Smart Stock', factory:'Factory', signIn:'Sign in', adminSignIn:'Admin sign in', signOut:'Sign out',
  username:'Username', password:'Password', useNow:'Use now', history:'History', home:'Dashboard',
  stock:'Stock', stockMovement:'Stock In / Out', summary:'Summary', manage:'Settings', dashboard:'Dashboard', buy:'Suggested Buy',
  myHistory:'My history', teamHistory:'Team history', receive:'Receive', opening:'Opening stock',
  search:'Search items', save:'Save', confirm:'Confirm', cancel:'Cancel', saved:'Saved!',
  used:'Used', received:'Received', onHand:'On hand', noItems:'No items yet.', noHistory:'No records yet.'
});

function interfaceOverrides_(rows) {
  return Object.fromEntries(rows.filter(row => row[1] === 'EN' &&
    Object.hasOwn(INTERFACE_DEFAULTS, row[0]) && typeof row[2] === 'string' && row[2].trim())
    .map(row => [row[0], row[2].trim()]));
}

// Public login copy only. Never return other values from the database.
function interfaceText() {
  return {defaults:INTERFACE_DEFAULTS, overrides:interfaceOverrides_(rows_(database_(), 'LABELS'))};
}

function saveInterfaceText(changes, token) {
  return locked_(() => {
    const db = database_();
    principal_(db, ['ADMIN'], token);
    if (!changes || typeof changes !== 'object' || Array.isArray(changes)) throw new Error('Check the interface text.');
    const entries = Object.entries(changes);
    if (entries.some(([key, value]) => !Object.hasOwn(INTERFACE_DEFAULTS, key) ||
        typeof value !== 'string' || value.length > 80)) throw new Error('Check the interface text.');
    const rows = rows_(db, 'LABELS');
    const overrides = interfaceOverrides_(rows);
    const inserts = [];
    const requests = [];
    const timestamp = now_();
    for (const [key, raw] of entries) {
      const value = raw.trim();
      const index = rows.findIndex(row => row[0] === key && row[1] === 'EN');
      if ((overrides[key] || '') === value) continue;
      const row = [key, 'EN', value, timestamp];
      if (index < 0) inserts.push(row);
      else requests.push(update_(db, 'LABELS', index + 2, row));
      if (value) overrides[key] = value;
      else delete overrides[key];
    }
    if (inserts.length) requests.push(append_(db, 'LABELS', inserts));
    if (requests.length) commit_(db, requests);
    return {defaults:INTERFACE_DEFAULTS, overrides};
  });
}

// Deactivation keeps inventory keys, assignments and history recoverable.
function deleteCatalog(kind, id, token) {
  if (!Object.hasOwn(DELETION_PREFIX, kind) || typeof id !== 'string' || !id) throw new Error('Invalid request.');
  return locked_(() => {
    const db = database_();
    const actor = principal_(db, ['ADMIN'], token);
    const rows = rows_(db, kind);
    const index = rows.findIndex(row => row[0] === id);
    if (index < 0) throw new Error('Record unavailable.');
    const row = rows[index];
    if (kind === 'USERS' && (actor.id === id || (row[4] === 'ADMIN' && !row[3]))) {
      throw new Error('You cannot delete your own account or the Google owner account.');
    }
    const [activeColumn, updatedColumn] = {FACTORIES:[2,4], ITEMS:[9,11], USERS:[5,7]}[kind];
    const requests = [];
    const timestamp = now_();
    if (enabled_(row[activeColumn])) {
      row[activeColumn] = false;
      row[updatedColumn] = timestamp;
      requests.push(update_(db, kind, index + 2, row));
    }
    const settings = rows_(db, 'SETTINGS');
    const key = DELETION_PREFIX[kind] + id;
    const position = settings.findIndex(setting => setting[0] === key);
    if (position < 0 || !enabled_(settings[position][1])) {
      requests.push(upsert_(db, 'SETTINGS', position, [key, true, actor.id, timestamp]));
    }
    if (requests.length) commit_(db, requests);
    if (kind === 'USERS') properties_().deleteProperty('SESSION_' + id);
    return true;
  });
}

// A reversal applies the opposite movement to today's live balance. Its reference
// is the original Log_ID, which also makes repeated requests safe.
function deleteStockRecord(logId, token) {
  if (typeof logId !== 'string' || !logId) throw new Error('Invalid request.');
  return locked_(() => {
    const db = database_();
    const actor = principal_(db, ['ADMIN'], token);
    const [original] = matchingRows_(db, 'STOCK_LOG', 1, logId, 1);
    if (!original || !['USE', 'RECEIVE', 'OPENING'].includes(original[7])) {
      throw new Error('This record cannot be deleted.');
    }
    if (matchingRows_(db, 'STOCK_LOG', 15, logId, Number.MAX_SAFE_INTEGER).some(row => row[7] === 'REVERSAL')) {
      return true;
    }
    const stockRows = rows_(db, 'STOCK');
    const indexes = stockRows.flatMap((row, index) => row[0] === original[4] && row[1] === original[5] ? [index] : []);
    if (indexes.length !== 1 || !Number.isFinite(original[9]) || original[9] <= 0) {
      throw new Error('This record cannot be deleted.');
    }
    const index = indexes[0];
    const stock = stockRows[index];
    const before = stock[2];
    if (!Number.isFinite(before) || before < 0) throw new Error('This record cannot be deleted.');
    const after = round_(before + (original[7] === 'USE' ? original[9] : -original[9]));
    if (after < 0) throw new Error('Delete would make stock negative. Check later stock movements first.');
    if (after > 1e9) throw new Error('Delete would exceed the stock limit.');
    const timestamp = now_();
    const request = id_();
    const reversal = id_();
    stock[2] = after;
    stock[9] = timestamp;
    commit_(db, [
      update_(db, 'STOCK', index + 2, stock),
      append_(db, 'STOCK_LOG', [[reversal, request, timestamp, dateKey_(original[3]) || today_(),
        original[4], original[5], actor.id, 'REVERSAL', '', original[9], before, after,
        'Deleted by administrator', false, logId]]),
      append_(db, 'REQUEST_LOG', [[request, actor.id, 'REVERSAL', '', 'COMPLETED', reversal, timestamp, timestamp, '']])
    ]);
    return true;
  });
}

function checkReceivePassword_(actor, proof) {
  const properties = properties_();
  const expected = properties.getProperty('STOCK_IN_PASSWORD');
  if (!expected) throw new Error('Set a Stock In password in Settings first.');
  const key = 'STOCK_IN_TRIES_' + actor.id;
  throttle_(key, 10, 15);
  if (!passwordProof_(proof) || !same_(expected, mac_('stock-in|' + proof))) {
    throw new Error('Enter the correct Stock In password.');
  }
  properties.deleteProperty(key);
}

function companyProfile(token) {
  const db = database_();
  principal_(db, undefined, token);
  const settings = new Map(rows_(db, 'SETTINGS').map(row => [row[0], row[1]]));
  return {name:settings.get('COMPANY_NAME') || '', logo:storedPhoto_(settings.get('COMPANY_LOGO'))};
}

function saveCompany(input, token) {
  return locked_(() => {
    const db = database_();
    const actor = principal_(db, ['ADMIN'], token);
    const name = requiredText_(input?.name, 100);
    if (input.receiveProof !== undefined && !passwordProof_(input.receiveProof)) throw new Error('Enter a password.');
    const settings = rows_(db, 'SETTINGS');
    const timestamp = now_();
    const requests = [];
    const inserts = [];
    const values = [['COMPANY_NAME', name]];
    if (input.logo) values.push(['COMPANY_LOGO', storeItemImage_(input.logo)]);
    for (const [key, value] of values) {
      const index = settings.findIndex(row => row[0] === key);
      const row = [key, value, actor.id, timestamp];
      if (index < 0) inserts.push(row);
      else requests.push(update_(db, 'SETTINGS', index + 2, row));
    }
    if (inserts.length) requests.push(append_(db, 'SETTINGS', inserts));
    commit_(db, requests);
    if (input.receiveProof) properties_().setProperty('STOCK_IN_PASSWORD', mac_('stock-in|' + input.receiveProof));
    return true;
  });
}

function proofPhoto(logId, token) {
  const db = database_();
  const actor = principal_(db, ['ADMIN', 'SUPERVISOR'], token);
  const [row] = matchingRows_(db, 'STOCK_LOG', 1, String(logId), 1);
  if (!row) throw new Error('Record unavailable.');
  if (actor.role !== 'ADMIN') requireFactory_(db, actor, row[4]);
  const match = /^PHOTO:([a-z0-9_-]+)$/i.exec(String(row[12]));
  return match ? storedPhoto_(match[1]) : '';
}

function inventoryReport(factory, token) {
  const db = database_();
  const actor = principal_(db, ['ADMIN','SUPERVISOR'], token);
  const factories = new Map(factoriesFor_(db, actor).map(row => [row[0], row[1]]));
  if (factory && actor.role !== 'ADMIN' && !factories.has(factory)) throw new Error('Access denied.');
  if (factory && actor.role === 'ADMIN' && !factories.has(factory)) {
    const archived = rows_(db, 'FACTORIES').find(row => row[0] === factory);
    if (!archived) throw new Error('Access denied.');
    factories.set(archived[0], archived[1]);
  }
  const relevant = id => factories.has(id) && (!factory || id === factory);
  const items = itemMap_(db);
  const settings = new Map(rows_(db, 'SETTINGS').map(row => [row[0], row[1]]));
  const ledger = rows_(db, 'STOCK_LOG');
  const reversed = new Set(ledger.filter(row => row[7] === 'REVERSAL').map(row => row[14]));
  const today = today_();
  const start = daysFrom_(today, -28);
  const recent = daysFrom_(today, -1);
  const totals = new Map();
  for (const row of ledger) {
    if (!relevant(row[4]) || reversed.has(row[0]) || !['USE','RECEIVE','OPENING'].includes(row[7]) || !Number.isFinite(row[9])) continue;
    const id = JSON.stringify([row[4], row[5]]);
    if (!totals.has(id)) totals.set(id, {received:0, used:0, opening:0, lastType:'', lastDate:'', days:new Map()});
    const value = totals.get(id);
    value.lastType = row[7];
    value.lastDate = dateKey_(row[3]);
    if (row[7] === 'RECEIVE') value.received += row[9];
    if (row[7] === 'OPENING') value.opening += row[9];
    if (row[7] !== 'USE') continue;
    value.used += row[9];
    const day = dateKey_(row[3]);
    if (day >= start && day < today) value.days.set(day, (value.days.get(day) || 0) + row[9]);
  }
  return rows_(db, 'STOCK').flatMap(row => {
    const [factoryId, id, quantity] = row;
    const item = items.get(id);
    if (!relevant(factoryId) || !item || (!factory && (!enabled_(row[8]) || !enabled_(item[9])))) return [];
    const total = totals.get(JSON.stringify([factoryId, id])) || {received:0, used:0, opening:0, lastType:'', lastDate:'', days:new Map()};
    const observed = [...total.days.values()];
    const minimum = within_(settings.get('ITEM_MINIMUM_' + id), 0, 1e6, 0);
    const daily = observed.length ? round_(observed.reduce((sum, n) => sum + n, 0) / observed.length) : null;
    const latestRestock = total.lastType === 'RECEIVE' && total.lastDate >= recent;
    const status = quantity === 0 ? 'OUT' : quantity <= minimum ? 'LOW' : latestRestock ? 'RESTOCKED' : 'OK';
    return [{...itemView_(item), code:settings.get('ITEM_CODE_' + id) || '', factory:factoryId, factoryName:factories.get(factoryId),
      quantity, minimum, daily, received:round_(total.received), used:round_(total.used), opening:round_(total.opening),
      active:enabled_(row[8]) && enabled_(item[9]), status}];
  });
}
