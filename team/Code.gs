function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Smart Stock')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function start(token) {
  const db = database_();
  const actor = principal_(db, undefined, token);
  return {name: actor.name, role: actor.role, factories: factoriesFor_(db, actor)
    .map(row => ({id: row[0], name: row[1]}))};
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
  const name = typeof value === 'string' ? value.trim().toLowerCase() : '';
  return /^[a-z0-9][a-z0-9._-]{2,39}$/.test(name) ? name : '';
}

function mac_(value) {
  const key = properties_().getProperty('AUTH_KEY');
  if (!key) throw new Error('Setup required.');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(value, key));
}

function credential_(id, password) {
  return mac_('password|' + id + '|' + password);
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
  if (!row || !['WORKER', 'SUPERVISOR'].includes(row[4]) || !enabled_(row[5])) return null;
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

function signIn(username, password) {
  const name = username_(username);
  if (!name || typeof password !== 'string' || !/^[0-9a-f]{32}$/.test(password)) {
    throw new Error('Login failed.');
  }
  return locked_(() => {
    throttle_('LOGIN_GLOBAL', 100, 5);
    const db = database_();
    const row = rows_(db, 'USERS').find(user =>
      user[2] === name && ['WORKER', 'SUPERVISOR'].includes(user[4]) && enabled_(user[5]));
    if (!row) throw new Error('Login failed.');
    const id = row[0];
    throttle_('LOGIN_' + id, 10, 15);
    if (!same_(row[3], credential_(id, password))) throw new Error('Login failed.');
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

function factoriesFor_(db, user) {
  return rows_(db, 'FACTORIES').filter(row => enabled_(row[2]) &&
    (!user.factories || user.factories.has(row[0])));
}

function requireFactory_(db, user, id) {
  if (!factoriesFor_(db, user).some(row => row[0] === id)) throw new Error('Access denied.');
}

function saveCatalog(kind, input) {
  if (!['FACTORIES', 'ITEMS', 'USERS', 'STOCK'].includes(kind) || !input) throw new Error('Invalid request.');
  return locked_(() => {
    const db = database_();
    const actor = principal_(db, ['ADMIN']);
    const entries = rows_(db, kind);
    const timestamp = now_();
    let index = entries.findIndex(row => row[0] === input.id);
    if (kind !== 'STOCK' && input.id && index < 0) throw new Error('Record unavailable.');
    const old = entries[index];
    const id = old ? old[0] : id_();
    let row;
    let issuedPassword;
    const requests = [];
    if (kind === 'FACTORIES') {
      row = [id, requiredText_(input.name), input.active === true, old?.[3] || timestamp, timestamp];
    } else if (kind === 'ITEMS') {
      if (!Number.isFinite(input.pack) || input.pack <= 0 ||
          (!input.decimal && !Number.isInteger(input.pack))) throw new Error('Invalid pack size.');
      if (!input.decimal && rows_(db, 'STOCK').some(stock => stock[1] === id && !Number.isInteger(stock[2]))) {
        throw new Error('An existing inventory balance uses decimals.');
      }
      row = [id, requiredText_(input.name), old?.[2] || '', requiredText_(input.unit, 30),
        input.pack, input.decimal === true, old?.[6] || '', old?.[7] || '', old?.[8] ?? '',
        input.active === true, old?.[10] || timestamp, timestamp];
    } else if (kind === 'USERS') {
      const username = username_(input.username);
      if (!username) throw new Error('Use 3–40 characters for the username, starting with a letter or number. No spaces.');
      if (!['WORKER', 'SUPERVISOR'].includes(input.role) || old?.[4] === 'ADMIN') {
        throw new Error('Choose Worker or Supervisor.');
      }
      if (entries.some(other => other[0] !== id && username_(other[2]) === username)) {
        throw new Error('Username already exists. Select that user to edit or reset their password.');
      }
      if (!old || input.resetPassword === true) issuedPassword = id_().replace(/-/g, '');
      const verifier = issuedPassword ? credential_(id, issuedPassword) : old[3];
      const factoryIds = [...new Set(input.factories || [])];
      const valid = new Set(factoriesFor_(db, actor).map(factory => factory[0]));
      if (!factoryIds.every(factory => valid.has(factory)) ||
          (input.active === true && input.role !== 'ADMIN' && !factoryIds.length)) {
        throw new Error('Select assigned factories.');
      }
      row = [id, requiredText_(input.name), username, verifier, input.role,
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
    commit_(db, requests);
    return issuedPassword ? {id, password: issuedPassword} : {id};
  });
}

function catalog() {
  const db = database_();
  principal_(db, ['ADMIN']);
  const assignments = rows_(db, 'USER_FACTORIES');
  return {
    FACTORIES: rows_(db, 'FACTORIES').map(row => ({id: row[0], name: row[1], active: enabled_(row[2])})),
    ITEMS: rows_(db, 'ITEMS').map(row => ({id: row[0], name: row[1], unit: row[3],
      pack: row[4], decimal: enabled_(row[5]), active: enabled_(row[9])})),
    USERS: rows_(db, 'USERS').filter(row => row[4] !== 'ADMIN').map(row => ({id: row[0], name: row[1], username: row[2],
      role: row[4], active: enabled_(row[5]), factories: assignments.filter(a => a[0] === row[0]).map(a => a[1])})),
    STOCK: rows_(db, 'STOCK').map(row => ({factory: row[0], item: row[1], active: enabled_(row[8])}))
  };
}

function itemMap_(db) {
  return new Map(rows_(db, 'ITEMS').map(row => [row[0], row]));
}

function itemView_(item) {
  return {id: item[0], name: item[1], unit: item[3]};
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
  requireFactory_(db, actor, factory);
  const items = itemMap_(db);
  const selected = team === true ? matchingRows_(db, 'STOCK_LOG', 5, factory, 50) :
    matchingRows_(db, 'STOCK_LOG', 7, actor.id, 100).filter(row => row[4] === factory).slice(0, 50);
  return selected.map(row => {
    const item = items.get(row[5]);
    const summary = {date: String(row[3]), name: item?.[1] || String(row[5]),
      unit: item?.[3] || '', quantity: row[9], status: 'Saved'};
    return team === true ? {...summary, type: row[7], user: row[6],
      before: row[10], after: row[11], reason: row[12]} : summary;
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
      'This item already has stock history.', 'Stock limit reached.'
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
    JSON.stringify([type, input.factory, lines]), Utilities.Charset.UTF_8));
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
    commit_(db, [...updates, append_(db, 'STOCK_LOG', logs), append_(db, 'REQUEST_LOG', [
      [input.request, actor.id, type, hash, 'COMPLETED', reference, timestamp, timestamp, '']
    ])]);
    return {saved: true};
  });
}
