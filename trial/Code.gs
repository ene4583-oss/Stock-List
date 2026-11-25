const TRIAL_ITEMS = '试用物品';
const TRIAL_LOG = '试用领用记录';
const ITEM_HEADERS = ['物品名称'];
const LOG_HEADERS = ['提交编号', '登记时间', '物品名称', '领用数量'];

/** 在绑定到测试表格的 Apps Script 编辑器里运行一次。 */
function setupTrial() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const email = caller_();
  if (!spreadsheet || spreadsheet.getOwner()?.getEmail().toLowerCase() !== email) {
    throw new Error('请用表格拥有者账号，在该表格的 Apps Script 编辑器里运行 setupTrial。');
  }
  withLock_(() => {
    const tables = [[TRIAL_ITEMS, ITEM_HEADERS], [TRIAL_LOG, LOG_HEADERS]];
    for (const [name, headers] of tables) validateHeaders_(spreadsheet, name, headers);
    for (const [name, headers] of tables) {
      const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
      if (!sheet.getLastRow()) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
        sheet.setFrozenRows(1);
      }
    }
    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().setProperties({
      TRIAL_SPREADSHEET: spreadsheet.getId(),
      TRIAL_OWNER: email
    });
  });
}

function doGet() {
  trialSheet_();
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('领用登记')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function listItems() {
  return names_(trialSheet_());
}

function saveUsage(request, item, quantity) {
  if (typeof request !== 'string' || !/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(request) ||
      typeof item !== 'string' || !Number.isSafeInteger(quantity) || quantity < 1 || quantity > 1000000) {
    throw new Error('请选择物品，并输入 1 到 1,000,000 之间的整数数量。');
  }
  return withLock_(() => {
    const spreadsheet = trialSheet_();
    validateHeaders_(spreadsheet, TRIAL_LOG, LOG_HEADERS);
    const log = spreadsheet.getSheetByName(TRIAL_LOG);
    const count = log.getLastRow() - 1;
    const match = count > 0 && log.getRange(2, 1, count, 1).createTextFinder(request)
      .matchEntireCell(true).matchCase(true).findNext();
    if (match) {
      const row = log.getRange(match.getRow(), 3, 1, 2).getValues()[0];
      if (row[0] !== item || row[1] !== quantity) throw new Error('提交内容已改变，请重新登记。');
      return true;
    }
    if (!names_(spreadsheet).includes(item)) throw new Error('物品已从清单移除，请刷新网页。');
    // Leading apostrophe stores the item as literal text, even if it starts with "=".
    log.appendRow([request, new Date(), "'" + item, quantity]);
    SpreadsheetApp.flush();
    return true;
  });
}

function caller_() {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  if (!email || email !== Session.getEffectiveUser().getEmail().trim().toLowerCase()) {
    throw new Error('请用表格拥有者账号登录。');
  }
  return email;
}

function trialSheet_() {
  const email = caller_();
  const properties = PropertiesService.getScriptProperties();
  const id = properties.getProperty('TRIAL_SPREADSHEET');
  if (!id || properties.getProperty('TRIAL_OWNER') !== email) {
    throw new Error('请先在 Apps Script 编辑器运行 setupTrial。');
  }
  return SpreadsheetApp.openById(id);
}

function names_(spreadsheet) {
  validateHeaders_(spreadsheet, TRIAL_ITEMS, ITEM_HEADERS);
  const sheet = spreadsheet.getSheetByName(TRIAL_ITEMS);
  const count = sheet.getLastRow() - 1;
  if (count < 1) return [];
  return [...new Set(sheet.getRange(2, 1, count, 1).getDisplayValues().map(row => row[0].trim()).filter(Boolean))];
}

function validateHeaders_(spreadsheet, name, headers) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet || !sheet.getLastRow()) return;
  const actual = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  if (headers.some((header, index) => header !== actual[index])) {
    throw new Error('「' + name + '」已有不同的表头，请先检查。');
  }
}

function withLock_(action) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('系统忙，请再试一次。');
  try {
    return action();
  } finally {
    lock.releaseLock();
  }
}
