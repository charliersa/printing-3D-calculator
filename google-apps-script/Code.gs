/**
 * FilaCost 3D列印成本系統 — Google Sheets 資料庫 API
 *
 * 安裝:試算表 → 擴充功能 → Apps Script → 貼上本檔 → 執行 setup() →
 *       部署 → 新增部署作業 → 網頁應用程式(執行身分:我 / 存取權:所有人)
 * 詳細步驟見同資料夾 README.md
 */

// 每個資料集對應一張工作表。欄位型別:text 文字 / num 數字 / json JSON 字串
const SHEETS = {
  filaments: { name: '線材', cols: [
    ['id', 'text'], ['brand', 'text'], ['type', 'text'], ['color', 'text'], ['label', 'text'], ['hex', 'text'], ['sku', 'text'],
    ['price', 'num'], ['weight', 'num'], ['watt', 'num'], ['buyQty', 'num'],
    ['stockKg', 'num'], ['safeKg', 'num'], ['burnKg', 'num'], ['tiers', 'json'],
  ] },
  types: { name: '線材類型', cols: [['id', 'text'], ['name', 'text'], ['mode', 'text']] },
  brands: { name: '廠商', cols: [['id', 'text'], ['name', 'text']] },
  machines: { name: '機台', cols: [
    ['id', 'text'], ['label', 'text'], ['name', 'text'], ['mode', 'text'],
    ['price', 'num'], ['life', 'num'], ['watt', 'num'], ['consum', 'num'],
    ['state', 'text'], ['pct', 'num'], ['job', 'text'],
  ] },
  calcFixed: { name: '固定成本', cols: [['id', 'num'], ['name', 'text'], ['amount', 'num'], ['per', 'text']] },
  // array:true 表示前端資料是陣列列(依欄位順序),不是物件
  orderRows: { name: '訂單', array: true, cols: [
    ['訂單編號', 'text'], ['品項', 'text'], ['客戶', 'text'], ['數量', 'num'], ['金額', 'num'], ['毛利', 'num'], ['狀態', 'text'],
  ] },
  customerRows: { name: '客戶', array: true, cols: [
    ['客戶', 'text'], ['聯絡方式', 'text'], ['類型', 'text'], ['訂單數', 'num'], ['累積營收', 'num'], ['累積毛利', 'num'], ['常用材料', 'text'],
  ] },
  catalog: { name: '商品', cols: [
    ['name', 'text'], ['mat', 'text'], ['w', 'num'], ['h', 'num'], ['badge', 'text'], ['tint', 'text'], ['cat', 'text'],
  ] },
  rates: { name: '電價級距', cols: [['range', 'text'], ['summer', 'num'], ['normal', 'num']] },
};
const SETTINGS_SHEET = '設定';
const EXTRA_COL = '_extra'; // 物件資料中未定義的欄位,以 JSON 保存,避免遺失

// ---------- 入口 ----------
function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '無法解析請求內容' });
  }
  return handle_(req);
}

function handle_(req) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
    if (!token) return json_({ ok: false, error: '尚未執行 setup() 產生存取金鑰' });
    if (String(req.token || '') !== token) return json_({ ok: false, error: '存取金鑰錯誤' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    switch (req.action) {
      case 'ping':
        return json_({ ok: true, spreadsheet: ss.getName(), counts: counts_(ss) });
      case 'pull':
        return json_({ ok: true, collections: readAll_(ss), settings: readSettings_(ss), updatedAt: new Date().toISOString() });
      case 'push': {
        const lock = LockService.getScriptLock();
        lock.waitLock(20000);
        try {
          const cols = req.collections || {};
          const written = [];
          Object.keys(cols).forEach(key => {
            if (SHEETS[key] && Array.isArray(cols[key])) { writeSheet_(ss, key, cols[key]); written.push(key); }
          });
          if (req.settings && typeof req.settings === 'object') { writeSettings_(ss, req.settings); written.push('settings'); }
          return json_({ ok: true, written, updatedAt: new Date().toISOString() });
        } finally {
          lock.releaseLock();
        }
      }
      default:
        return json_({ ok: false, error: '未知的 action:' + req.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

// ---------- 首次安裝:建立工作表與存取金鑰 ----------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(key => {
    const sh = sheet_(ss, SHEETS[key].name);
    if (sh.getLastRow() === 0) {
      const header = header_(SHEETS[key]);
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  const settings = sheet_(ss, SETTINGS_SHEET);
  if (settings.getLastRow() === 0) {
    settings.getRange(1, 1, 1, 2).setValues([['key', 'value']]).setFontWeight('bold');
    settings.setFrozenRows(1);
  }
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('API_TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('API_TOKEN', token);
  }
  Logger.log('存取金鑰 API_TOKEN:' + token);
  return token;
}

// 需要更換金鑰時執行(舊金鑰立即失效)
function resetToken() {
  PropertiesService.getScriptProperties().deleteProperty('API_TOKEN');
  return setup();
}

// ---------- 讀寫資料表 ----------
function sheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function header_(def) {
  return def.cols.map(c => c[0]).concat(def.array ? [] : [EXTRA_COL]);
}

function toCell_(type, v) {
  if (v === undefined || v === null) return '';
  if (type === 'num') {
    if (v === '') return '';
    const n = Number(v);
    return isFinite(n) ? n : '';
  }
  if (type === 'json') return JSON.stringify(v);
  return String(v);
}

function fromCell_(type, v, keepBlank) {
  if (v === '' || v === null || v === undefined) return keepBlank ? (type === 'num' ? 0 : '') : undefined;
  if (type === 'num') {
    const n = Number(v);
    return isFinite(n) ? n : (keepBlank ? 0 : undefined);
  }
  if (type === 'json') {
    try { return JSON.parse(v); } catch (err) { return undefined; }
  }
  return String(v);
}

function writeSheet_(ss, key, rows) {
  const def = SHEETS[key];
  const sh = sheet_(ss, def.name);
  const header = header_(def);
  const known = {};
  def.cols.forEach(c => { known[c[0]] = true; });

  const data = rows.map(row => {
    if (def.array) return def.cols.map((c, i) => toCell_(c[1], (row || [])[i]));
    const extra = {};
    Object.keys(row || {}).forEach(k => { if (!known[k] && row[k] !== undefined) extra[k] = row[k]; });
    return def.cols.map(c => toCell_(c[1], row[c[0]])).concat([Object.keys(extra).length ? JSON.stringify(extra) : '']);
  });

  const all = [header].concat(data);
  sh.clearContents();
  // 文字與 JSON 欄設為純文字格式,避免料號 0012 之類被轉成數字
  header.forEach((h, i) => {
    const type = i < def.cols.length ? def.cols[i][1] : 'json';
    if (type !== 'num') sh.getRange(1, i + 1, all.length, 1).setNumberFormat('@');
  });
  sh.getRange(1, 1, all.length, header.length).setValues(all);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function readSheet_(ss, key) {
  const def = SHEETS[key];
  const sh = ss.getSheetByName(def.name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const idx = {};
  values[0].forEach((h, i) => { idx[String(h).trim()] = i; });

  return values.slice(1)
    .filter(r => r.some(v => v !== '' && v !== null))
    .map(r => {
      if (def.array) return def.cols.map(c => fromCell_(c[1], idx[c[0]] === undefined ? '' : r[idx[c[0]]], true));
      const o = {};
      if (idx[EXTRA_COL] !== undefined && r[idx[EXTRA_COL]] !== '') {
        try { Object.assign(o, JSON.parse(r[idx[EXTRA_COL]])); } catch (err) {}
      }
      def.cols.forEach(c => {
        if (idx[c[0]] === undefined) return;
        const v = fromCell_(c[1], r[idx[c[0]]], false);
        if (v !== undefined) o[c[0]] = v;
      });
      return o;
    });
}

function readAll_(ss) {
  const out = {};
  Object.keys(SHEETS).forEach(key => { out[key] = readSheet_(ss, key); });
  return out;
}

function counts_(ss) {
  const out = {};
  Object.keys(SHEETS).forEach(key => {
    const sh = ss.getSheetByName(SHEETS[key].name);
    out[SHEETS[key].name] = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
  });
  return out;
}

// ---------- 設定(計算機參數、預設值、報價附註) ----------
function readSettings_(ss) {
  const sh = ss.getSheetByName(SETTINGS_SHEET);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    const key = String(r[0] || '').trim();
    if (!key) return;
    try { out[key] = JSON.parse(r[1]); } catch (err) { out[key] = r[1]; }
  });
  return out;
}

function writeSettings_(ss, settings) {
  const merged = readSettings_(ss);
  Object.keys(settings).forEach(k => { if (settings[k] !== undefined) merged[k] = settings[k]; });
  merged.updatedAt = new Date().toISOString();
  const rows = [['key', 'value']].concat(Object.keys(merged).map(k => [k, JSON.stringify(merged[k])]));
  const sh = sheet_(ss, SETTINGS_SHEET);
  sh.clearContents();
  sh.getRange(1, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
