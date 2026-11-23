/**
 * 冰箱藥品管理稽核 — 後台程式（Google Apps Script）
 *
 * 使用步驟：
 * 1. 把下面三行改成你自己的員工編號、姓名、單位（這會是第一個管理員帳號）
 * 2. 在同一個專案新增 HTML 檔，命名為 Index，貼上 Index.html 的內容
 * 3. 存檔後，回到試算表重新整理，上方選單「稽核系統」→「初始化／檢查設定」
 * 4. 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *    執行身分：我；誰可以存取：所有人 → 部署。產生的網址就是 app 的網址。
 *
 * 登入方式：員工編號（必須在「員工」工作表中，且狀態為「啟用」）。
 */
const FIRST_ADMIN_ID = '00000';        // ← 改成你的員工編號
const FIRST_ADMIN_NAME = '系統管理員';  // ← 改成你的姓名
const FIRST_ADMIN_UNIT = '急診護理站';  // ← 改成你的單位

const DEFAULT_UNITS = ['急診護理站', '急診觀察室', '內科加護病房', '外科加護病房', '心臟加護病房',
  '新生兒加護病房', '內科病房', '外科病房', '婦產科病房', '兒科病房', '手術室', '恢復室', '血液透析室', '門診'];

const SHEETS = { users: '員工', units: '單位', audits: '審核紀錄', config: '設定', phrases: '常用原因' };
const DEFAULT_PHRASES = { problems: ["冰箱內存放食物或非藥品物品", "存放未經藥庫同意列管之藥品", "UD 用藥未註明病患基本資料", "同一藥盒存放兩種以上藥品", "藥盒未標示藥品名稱或數量", "未每班登錄冰箱溫度", "冰箱溫度超出 2–8℃ 未即時處理", "針劑放置於冰箱門邊", "外觀相似藥品（如 Insulin）未分層存放", "開封藥品未標示開封日期或開封後效期", "人員不熟悉高低溫度計操作及判讀", "點班本未每班點班", "護理長未每週稽核蓋章", "新進人員對規範不熟悉"],
  actions: ["單位內在職教育，宣導冰箱藥品管理規範", "於冰箱門張貼存放規範及標示範例", "備妥開封日期標籤貼紙", "交班時由組長確認溫度紀錄並簽名", "溫度異常時立即通報並依流程處理", "調整藥品擺放位置，針劑移至冰箱內層", "外觀相似藥品加貼警示標籤並分層存放", "護理長每週抽查並回饋結果", "列入新進人員訓練課程", "下次稽核持續追蹤"] };
const PHRASE_KIND = { problems: '問題分析', actions: '改善措施' };
const USER_COLS = ['員工編號', '姓名', '單位', '角色', '狀態', '建立時間', '最後登入'];
const AUDIT_COLS = ['審核單ID', '單位', '審核日期', '追蹤月份', '審核人員', '填寫人員編', '填寫人姓名', '單位主管',
  '閾值(%)', '通過率(%)', '結果', '未通過項目', '建議', '建立時間', '更新時間', '完整資料'];
const TZ = 'Asia/Taipei';

/* ================= 試算表選單與初始化 ================= */
function onOpen() {
  // 只有從試算表打開時才有選單；在 Apps Script 編輯器直接執行不會出錯
  try { SpreadsheetApp.getUi().createMenu('稽核系統').addItem('初始化／檢查設定', 'setup').addToUi(); } catch (e) { }
}

/**
 * 取得資料試算表：
 * - 從試算表「擴充功能 → Apps Script」建立的專案：用這個試算表
 * - 在 script.google.com 直接建立的專案：第一次執行 setup 時自動建立一個新的試算表，之後都用它
 */
function ss_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) { }
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('冰箱藥品管理稽核資料');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function setup() {
  const ss = ss_();
  const u = sheet_(ss, SHEETS.users, USER_COLS);
  u.getRange('A:A').setNumberFormat('@');

  const un = sheet_(ss, SHEETS.units, ['單位名稱', '單位主管']);
  if (!un.getRange(1, 2).getValue()) un.getRange(1, 2).setValue('單位主管').setFontWeight('bold').setBackground('#dcedf3');
  if (un.getLastRow() < 2) un.getRange(2, 1, DEFAULT_UNITS.length, 1).setValues(DEFAULT_UNITS.map(x => [x]));

  const a = sheet_(ss, SHEETS.audits, AUDIT_COLS);
  a.getRange('A:D').setNumberFormat('@');
  a.getRange('F:F').setNumberFormat('@');

  const ph = sheet_(ss, SHEETS.phrases, ['類別', '內容']);
  if (ph.getLastRow() < 2) {
    const rows = [];
    Object.keys(DEFAULT_PHRASES).forEach(k => DEFAULT_PHRASES[k].forEach(t => rows.push([PHRASE_KIND[k], t])));
    ph.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  const c = sheet_(ss, SHEETS.config, ['項目', '值']);
  if (c.getLastRow() < 2) c.appendRow(['預設閾值', 90]);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());

  if (!users_().some(x => x.id === FIRST_ADMIN_ID)) {
    const r = u.getLastRow() + 1;
    u.getRange(r, 1, 1, 7).setValues([[FIRST_ADMIN_ID, FIRST_ADMIN_NAME, FIRST_ADMIN_UNIT, '管理員', '啟用', new Date(), '']]);
  }

  ['工作表1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s);
  });

  const msg = '初始化完成！\n\n資料試算表：' + ss.getUrl() + '\n第一個管理員：' + FIRST_ADMIN_ID + '（' + FIRST_ADMIN_NAME + '）\n請用員工編號登入。\n\n下一步：部署 → 新增部署作業 → 網頁應用程式。';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function sheet_(ss, name, cols) {
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  if (s.getLastRow() === 0) {
    s.appendRow(cols);
    s.getRange(1, 1, 1, cols.length).setFontWeight('bold').setBackground('#dcedf3');
    s.setFrozenRows(1);
  }
  return s;
}
function sh_(name) {
  let s = ss_().getSheetByName(name);
  if (!s) {
    // 還沒初始化：第一次有人使用時自動建立所有工作表與第一個管理員
    setupSilently_();
    s = ss_().getSheetByName(name);
  }
  if (!s) throw new Error('後台初始化失敗，請管理員在試算表執行「稽核系統 → 初始化」');
  return s;
}
function setupSilently_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { setup(); } finally { lock.releaseLock(); }
}

/* ================= API 入口 ================= */
/** 打開部署網址時，直接顯示 app（Index.html）。 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('冰箱藥品管理稽核')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** app 放在 Apps Script 裡時，畫面透過 google.script.run 呼叫這個函式。 */
function api(req) {
  return handle_(req || {});
}

/** app 放在 GitHub Pages 等外部網站時，透過網址 POST 呼叫。 */
function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { }
  return ContentService.createTextOutput(JSON.stringify(handle_(req))).setMimeType(ContentService.MimeType.JSON);
}

function handle_(req) {
  try {
    const fn = ACTIONS[req.action];
    if (!fn) throw new Error('未知的操作');
    return JSON.parse(JSON.stringify({ ok: true, data: fn(req) }));
  } catch (err) {
    return { ok: false, error: err.message || String(err), code: err.code || '' };
  }
}

const ACTIONS = {
  login(r) {
    const id = String(r.id || '').trim();
    if (!id) throw new Error('請輸入員工編號');
    if (tooMany_(id)) throw new Error('錯誤次數過多，請 10 分鐘後再試');
    const u = users_().find(x => x.id === id);
    if (!u) { fail_(id); throw new Error('查無此員工編號，或尚未加入員工名單'); }
    if (u.status !== '啟用') throw new Error('此帳號已停用，請洽管理員');
    CacheService.getScriptCache().remove('f_' + id);
    sh_(SHEETS.users).getRange(u.row, 7).setValue(new Date());
    return { token: makeToken_(u, r.remember), user: pub_(u) };
  },

  me(r) { return { user: pub_(auth_(r.token)) }; },

  getConfig(r) {
    auth_(r.token);
    return { units: units_(), heads: unitHeads_(), threshold: threshold_(), phrases: phrases_(), staff: users_().filter(u => u.status === '啟用').map(u => ({ id: u.id, name: u.name, unit: u.unit })) };
  },

  saveAudit(r) {
    const u = auth_(r.token);
    const d = r.data;
    if (!d || !d.id || !Array.isArray(d.rows)) throw new Error('審核單資料不完整');
    return withLock_(() => {
      const s = sh_(SHEETS.audits);
      const row = findRow_(s, d.id);
      const now = new Date();
      let created = now, ownerId = u.id, ownerName = u.name;
      if (row) {
        const old = s.getRange(row, 1, 1, AUDIT_COLS.length).getValues()[0];
        if (String(old[5]) !== u.id && u.role !== '管理員') throw new Error('只能修改自己填寫的審核單');
        created = old[13] || now; ownerId = String(old[5]); ownerName = String(old[6]);
      }
      d.owner = ownerId; d.ownerName = ownerName;
      d.updatedAt = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm');
      const c = calc_(d);
      const vals = [[String(d.id), d.unit || '', d.auditDate || '', d.trackMonth || '', d.auditor || '', ownerId, ownerName,
        d.head || '', Number(d.threshold) || 0, c.total == null ? '' : c.total,
        c.total == null ? '' : (c.met ? '達閾值' : '未達閾值'), c.fails, d.advice || '', created, now, JSON.stringify(d)]];
      const target = row || s.getLastRow() + 1;
      s.getRange(target, 1, 1, AUDIT_COLS.length).setValues(vals);
      return { id: d.id, updatedAt: d.updatedAt, owner: ownerId, ownerName: ownerName };
    });
  },

  listAudits(r) {
    const u = auth_(r.token);
    const s = sh_(SHEETS.audits);
    const last = s.getLastRow();
    if (last < 2) return [];
    const v = s.getRange(2, 1, last - 1, AUDIT_COLS.length).getValues();
    const out = [];
    v.forEach(x => {
      if (!x[0]) return;
      if (u.role !== '管理員' && String(x[5]) !== u.id) return;
      try { out.push(JSON.parse(x[15])); } catch (e) { }
    });
    return out;
  },

  deleteAudit(r) {
    const u = auth_(r.token);
    return withLock_(() => {
      const s = sh_(SHEETS.audits);
      const row = findRow_(s, r.id);
      if (!row) throw new Error('找不到這份審核單');
      if (String(s.getRange(row, 6).getValue()) !== u.id && u.role !== '管理員') throw new Error('只能刪除自己填寫的審核單');
      s.deleteRow(row);
      return true;
    });
  },

  /* ---- 管理員 ---- */
  listUsers(r) { auth_(r.token, true); return users_().map(pub_); },

  saveUser(r) {
    const me = auth_(r.token, true);
    const x = r.user || {};
    const id = String(x.id || '').trim(), name = String(x.name || '').trim();
    if (!id || !name) throw new Error('請填寫員工編號和姓名');
    const role = x.role === '管理員' ? '管理員' : '一般';
    const status = x.status === '停用' ? '停用' : '啟用';
    return withLock_(() => {
      const u = users_().find(y => y.id === id);
      if (u) {
        if (r.isNew) throw new Error('員工編號 ' + id + ' 已經存在');
        if (u.id === me.id && (role !== '管理員' || status !== '啟用')) throw new Error('不能停用自己，或取消自己的管理員權限');
        u.name = name; u.unit = x.unit || ''; u.role = role; u.status = status;
        writeUser_(u);
        return pub_(u);
      }
      const s = sh_(SHEETS.users);
      const row = s.getLastRow() + 1;
      s.getRange(row, 1).setNumberFormat('@');
      s.getRange(row, 1, 1, 7).setValues([[id, name, x.unit || '', role, status, new Date(), '']]);
      return pub_(users_().find(y => y.id === id));
    });
  },

  saveUnits(r) {
    auth_(r.token, true);
    // units：[{ name, head }]（也接受單純的單位名稱字串）
    const seen = {}, rows = [];
    (r.units || []).forEach(x => {
      const name = String(typeof x === 'string' ? x : (x && x.name) || '').trim();
      const head = typeof x === 'string' ? '' : String((x && x.head) || '').trim();
      if (name && !seen[name]) { seen[name] = 1; rows.push([name, head]); }
    });
    if (!rows.length) throw new Error('單位清單不能是空的');
    return withLock_(() => {
      const s = sh_(SHEETS.units);
      if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).clearContent();
      s.getRange(2, 1, rows.length, 2).setValues(rows);
      return { units: units_(), heads: unitHeads_() };
    });
  },

  addPhrase(r) {
    auth_(r.token);
    const kind = PHRASE_KIND[r.kind], text = String(r.text || '').trim();
    if (!kind || !text) throw new Error('請輸入原因內容');
    return withLock_(() => {
      const list = phrases_();
      if (list[r.kind].indexOf(text) < 0) sh_(SHEETS.phrases).appendRow([kind, text]);
      return phrases_();
    });
  },

  deletePhrase(r) {
    auth_(r.token);
    const kind = PHRASE_KIND[r.kind], text = String(r.text || '').trim();
    if (!kind || !text) throw new Error('找不到這個原因');
    return withLock_(() => {
      const s = sh_(SHEETS.phrases);
      const last = s.getLastRow();
      if (last > 1) {
        const v = s.getRange(2, 1, last - 1, 2).getValues();
        for (let i = v.length - 1; i >= 0; i--) {
          if (String(v[i][0]).trim() === kind && String(v[i][1]).trim() === text) s.deleteRow(i + 2);
        }
      }
      return phrases_();
    });
  },

  savePhrases(r) {
    auth_(r.token, true);
    return withLock_(() => {
      const s = sh_(SHEETS.phrases);
      if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).clearContent();
      const rows = [];
      Object.keys(PHRASE_KIND).forEach(k => [...new Set((r.phrases && r.phrases[k] || []).map(t => String(t).trim()).filter(Boolean))].forEach(t => rows.push([PHRASE_KIND[k], t])));
      if (rows.length) s.getRange(2, 1, rows.length, 2).setValues(rows);
      return phrases_();
    });
  },

  saveConfig(r) {
    auth_(r.token, true);
    const t = Number(r.threshold);
    if (!(t >= 0 && t <= 100)) throw new Error('閾值要介於 0 到 100');
    return withLock_(() => {
      const s = sh_(SHEETS.config);
      const v = s.getDataRange().getValues();
      const i = v.findIndex(x => x[0] === '預設閾值');
      if (i > 0) s.getRange(i + 1, 2).setValue(t); else s.appendRow(['預設閾值', t]);
      return { threshold: t };
    });
  }
};

/* ================= 資料存取 ================= */
function users_() {
  const s = sh_(SHEETS.users);
  const v = s.getDataRange().getValues();
  return v.slice(1).map((r, i) => ({
    row: i + 2,
    id: String(r[0]).trim(), name: String(r[1]).trim(), unit: String(r[2]).trim(),
    role: String(r[3]).trim() === '管理員' ? '管理員' : '一般',
    status: String(r[4]).trim() === '停用' ? '停用' : '啟用',
    created: r[5] || new Date(), lastLogin: r[6]
  })).filter(u => u.id);
}
function writeUser_(u) {
  sh_(SHEETS.users).getRange(u.row, 1, 1, 7).setValues([[u.id, u.name, u.unit, u.role, u.status, u.created, u.lastLogin || '']]);
}
function pub_(u) {
  return {
    id: u.id, name: u.name, unit: u.unit, role: u.role, status: u.status,
    lastLogin: u.lastLogin ? Utilities.formatDate(new Date(u.lastLogin), TZ, 'yyyy-MM-dd HH:mm') : ''
  };
}
function units_() {
  const s = sh_(SHEETS.units);
  if (s.getLastRow() < 2) return DEFAULT_UNITS;
  return s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().map(x => String(x[0]).trim()).filter(Boolean);
}
function unitHeads_() {
  const s = sh_(SHEETS.units), out = {};
  if (s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(x => {
    const n = String(x[0]).trim(), h = String(x[1] || '').trim();
    if (n && h) out[n] = h;
  });
  return out;
}
function phrases_() {
  const s = ss_().getSheetByName(SHEETS.phrases);
  if (!s) return DEFAULT_PHRASES;
  const out = { problems: [], actions: [] };
  if (s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(x => {
    const t = String(x[1]).trim();
    if (!t) return;
    if (String(x[0]).trim() === '改善措施') out.actions.push(t); else out.problems.push(t);
  });
  return out;
}
function threshold_() {
  const v = sh_(SHEETS.config).getDataRange().getValues();
  const r = v.find(x => x[0] === '預設閾值');
  return r ? Number(r[1]) || 90 : 90;
}
function findRow_(s, id) {
  const last = s.getLastRow();
  if (last < 2) return 0;
  const ids = s.getRange(2, 1, last - 1, 1).getValues();
  const i = ids.findIndex(x => String(x[0]) === String(id));
  return i < 0 ? 0 : i + 2;
}
function calc_(d) {
  let V = 0, X = 0; const fails = [];
  (d.rows || []).forEach(r => (r.r || []).forEach((c, i) => {
    if (c === 'V') V++;
    if (c === 'X') { X++; fails.push((r.name || '未命名') + '-第' + (i + 1) + '項'); }
  }));
  const total = V + X ? Math.round(V / (V + X) * 1000) / 10 : null;
  return { total: total, met: total != null && total >= Number(d.threshold || 0), fails: fails.join('、') };
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ================= 登入憑證 ================= */
// 名字有改、帳號停用或被刪除時，舊的登入憑證會自動失效
function secret_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('SECRET');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('SECRET', k); }
  return k;
}
function sign_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, secret_())); }
function nameKey_(u) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, u.name, Utilities.Charset.UTF_8)).slice(0, 8);
}
function makeToken_(u, remember) {
  const exp = Date.now() + (remember ? 30 * 24 : 12) * 3600 * 1000;
  const p = [u.id, exp, nameKey_(u)].join('|');
  return Utilities.base64EncodeWebSafe(p, Utilities.Charset.UTF_8) + '.' + sign_(p);
}
function authErr_() { const e = new Error('登入已過期，請重新登入'); e.code = 'AUTH'; return e; }
function auth_(token, needAdmin) {
  if (!token) throw authErr_();
  const parts = String(token).split('.');
  let p;
  try { p = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString('UTF-8'); } catch (e) { throw authErr_(); }
  if (sign_(p) !== parts[1]) throw authErr_();
  const f = p.split('|');
  if (Date.now() > Number(f[1])) throw authErr_();
  const u = users_().find(x => x.id === f[0]);
  if (!u || u.status !== '啟用' || nameKey_(u) !== f[2]) throw authErr_();
  if (needAdmin && u.role !== '管理員') throw new Error('只有管理員可以使用這個功能');
  return u;
}
function tooMany_(id) { return Number(CacheService.getScriptCache().get('f_' + id) || 0) >= 5; }
function fail_(id) {
  const c = CacheService.getScriptCache();
  c.put('f_' + id, String(Number(c.get('f_' + id) || 0) + 1), 600);
}
