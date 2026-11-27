/**
 * 禾煎 HOJIAN — 線上訂單後端（Google Apps Script，綁定在 Google 試算表）
 *
 * 功能
 *  - doPost：接收網站送出的訂單 → 驗證 → 以「商品」工作表重新計價 → 寫入「訂單」「訂單明細」→ 寄通知信
 *  - doGet：健康檢查（瀏覽器打開網址可確認服務是否正常）
 *  - 試算表選單「禾煎訂單」：初始化工作表、寄測試信、今日營收摘要、每日營收報表排程
 *
 * 安裝步驟請見 gas/README.md。
 */

const TZ = 'Asia/Taipei';
const SHEETS = { orders: '訂單', items: '訂單明細', products: '商品', settings: '設定' };
const ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '用餐方式', '門市', '顧客姓名', '電話', '電子郵件', '宅配地址', '商品摘要', '件數', '商品小計', '運費', '合計', '備註'];
const ITEM_HEADERS = ['訂單編號', '建立時間', '商品ID', '品名', '單價', '數量', '小計'];
const PRODUCT_HEADERS = ['商品ID', '品名', '單價', '單位', '上架'];
const STATUSES = ['新訂單', '製作中', '可取餐 / 已出貨', '已完成', '已取消'];
const STATUS_COLORS = { '新訂單': '#fde2d2', '製作中': '#fff1c2', '可取餐 / 已出貨': '#d9ecd0', '已完成': '#e4e4e4', '已取消': '#f4cccc' };
const MODES = { dinein: '內用', takeout: '外帶', delivery: '宅配' };

// 與網站菜單相同的初始商品；之後以試算表「商品」工作表為準（網站價格請保持一致）。
const DEFAULT_PRODUCTS = [
  ['p1', '經典鮮肉鍋貼', 85, '10 入', '是'],
  ['p2', '韭香飽滿鍋貼', 90, '10 入', '是'],
  ['p3', '椒麻鮮肉鍋貼', 100, '10 入', '是'],
  ['d1', '高麗菜鮮肉水餃', 85, '10 入', '是'],
  ['d2', '鮮蝦豬肉水餃', 130, '10 入', '是'],
  ['d3', '田園蔬食水餃', 95, '10 入', '是'],
  ['s1', '暖心酸辣湯', 45, '碗', '是'],
  ['s2', '紫菜蛋花湯', 35, '碗', '是'],
  ['s3', '冷泡烏龍茶', 40, '瓶', '是'],
  ['v1', '蒜香拍黃瓜', 40, '份', '是'],
  ['v2', '香滷豆干', 40, '份', '是'],
  ['v3', '季節時蔬', 45, '份', '是'],
  ['gift', '招牌招財鍋貼禮盒', 480, '50 入 / 盒', '是'],
];

const DEFAULT_SETTINGS = [
  ['通知信箱', '', '新訂單通知寄到這裡；多個信箱用逗號分隔'],
  ['宅配運費', 120, '宅配且未達免運門檻時收取'],
  ['免運門檻', 1000, '商品小計達此金額免運'],
  ['接受訂單', '是', '填「否」可暫停網站接單'],
  ['寄送顧客確認信', '是', '填「否」則只通知店家'],
  ['每小時同一電話下單上限', 5, '防止重複或惡意下單'],
];

/* ───────────────────────── 試算表選單 ───────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('禾煎訂單')
    .addItem('① 初始化工作表', 'setup')
    .addItem('寄送測試通知信', 'sendTestEmail')
    .addSeparator()
    .addItem('今日營收摘要', 'showTodaySummary')
    .addItem('開啟每日營收報表（每晚 9 點寄信）', 'installDailyReport')
    .addItem('關閉每日營收報表', 'removeDailyReport')
    .addToUi();
}

/** 建立／整理四個工作表。可重複執行，不會清除既有資料。 */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);

  const orders = ensureSheet_(ss, SHEETS.orders, ORDER_HEADERS);
  orders.setColumnWidths(1, ORDER_HEADERS.length, 110);
  orders.setColumnWidth(10, 280);
  orders.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  orders.getRange('L:N').setNumberFormat('"NT$ "#,##0');
  const statusRange = orders.getRange('C2:C');
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
  orders.setConditionalFormatRules(STATUSES.map(s => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(s).setBackground(STATUS_COLORS[s]).setRanges([statusRange]).build()));

  const items = ensureSheet_(ss, SHEETS.items, ITEM_HEADERS);
  items.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  items.getRange('E:E').setNumberFormat('"NT$ "#,##0');
  items.getRange('G:G').setNumberFormat('"NT$ "#,##0');

  const products = ensureSheet_(ss, SHEETS.products, PRODUCT_HEADERS);
  if (products.getLastRow() < 2) products.getRange(2, 1, DEFAULT_PRODUCTS.length, PRODUCT_HEADERS.length).setValues(DEFAULT_PRODUCTS);
  products.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['是', '否'], true).build());

  const settings = ensureSheet_(ss, SHEETS.settings, ['項目', '值', '說明']);
  const existing = settings.getLastRow() > 1 ? settings.getRange(2, 1, settings.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  DEFAULT_SETTINGS.filter(r => !existing.includes(r[0])).forEach(r => settings.appendRow(r));
  settings.setColumnWidth(1, 200).setColumnWidth(2, 260).setColumnWidth(3, 320);
  const me = Session.getEffectiveUser().getEmail();
  const s = readSettings_();
  if (!s['通知信箱'] && me) setSetting_('通知信箱', me);

  const blank = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  ss.setActiveSheet(orders);
  ss.toast('工作表已就緒。下一步：部署 → 新增部署作業 → 網頁應用程式。', '禾煎訂單', 8);
}

function ensureSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#3d2c23').setFontColor('#f8f5ef');
  sh.setFrozenRows(1);
  return sh;
}

/* ───────────────────────── 網頁應用程式端點 ───────────────────────── */

function doGet() {
  const s = readSettings_();
  return json_({ ok: true, service: 'hojian-orders', accepting: s['接受訂單'] !== '否' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e && e.postData && e.postData.contents || '');
  } catch (err) {
    return json_({ ok: false, error: '訂單資料格式錯誤' });
  }
  // 隱藏欄位有值 → 多半是機器人；假裝成功但不寫入。
  if (body && body.website) return json_({ ok: true, orderId: 'HJ-0', total: 0 });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return json_({ ok: false, error: '目前訂單較多，請稍後再試一次' });
  let order;
  try {
    order = buildOrder_(body);
    writeOrder_(order);
    countOrder_(order.phone);
  } catch (err) {
    if (!err.userMessage) console.error(err);
    return json_({ ok: false, error: err.userMessage || '訂單建立失敗，請稍後再試' });
  } finally {
    lock.releaseLock();
  }
  try {
    notify_(order);
  } catch (err) {
    console.error('寄信失敗', err); // 訂單已寫入，寄信失敗不影響顧客
  }
  return json_({ ok: true, orderId: order.id, subtotal: order.subtotal, shipping: order.shipping, total: order.total, mode: MODES[order.mode] });
}

/* ───────────────────────── 訂單處理 ───────────────────────── */

function buildOrder_(body) {
  const s = readSettings_();
  if (s['接受訂單'] === '否') throw userError_('目前暫停線上接單，敬請見諒');
  if (!body || typeof body !== 'object') throw userError_('訂單資料格式錯誤');

  const mode = String(body.mode || '');
  if (!MODES[mode]) throw userError_('請選擇用餐方式');
  const customer = clean_(body.customer, 40);
  const phone = clean_(body.phone, 20);
  const email = clean_(body.email, 120);
  const store = mode === 'delivery' ? '' : clean_(body.store, 60);
  const address = mode === 'delivery' ? clean_(body.address, 150) : '';
  const note = clean_(body.note, 200);
  if (!customer) throw userError_('請填寫聯絡姓名');
  if (!/^[0-9+() -]{8,20}$/.test(phone)) throw userError_('電話格式不正確');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw userError_('電子郵件格式不正確');
  if (mode === 'delivery' && !address) throw userError_('請填寫宅配地址');
  if (mode !== 'delivery' && !store) throw userError_('請選擇門市');

  checkRateLimit_(phone, Number(s['每小時同一電話下單上限']) || 5);

  // 以「商品」工作表的價格重新計算，不採用網站傳來的金額。
  const catalog = readProducts_();
  if (!Array.isArray(body.items) || !body.items.length) throw userError_('購物車是空的');
  if (body.items.length > 50) throw userError_('品項過多');
  const merged = {};
  body.items.forEach(it => {
    const id = String(it && it.id || '');
    const qty = Math.floor(Number(it && it.qty));
    if (!catalog[id]) throw userError_('部分商品目前無法訂購，請重新整理頁面後再試');
    if (!(qty >= 1 && qty <= 99)) throw userError_('商品數量需介於 1–99');
    merged[id] = Math.min(99, (merged[id] || 0) + qty);
  });
  const lines = Object.keys(merged).map(id => ({ id, name: catalog[id].name, price: catalog[id].price, qty: merged[id], amount: catalog[id].price * merged[id] }));
  const subtotal = lines.reduce((a, l) => a + l.amount, 0);
  const fee = Number(s['宅配運費']) || 0, threshold = Number(s['免運門檻']) || 0;
  const shipping = mode === 'delivery' && subtotal < threshold ? fee : 0;

  return { mode, customer, phone, email, store, address, note, lines, subtotal, shipping, total: subtotal + shipping,
    count: lines.reduce((a, l) => a + l.qty, 0), createdAt: new Date() };
}

function writeOrder_(o) {
  const ss = SpreadsheetApp.getActive();
  const orders = ss.getSheetByName(SHEETS.orders), items = ss.getSheetByName(SHEETS.items);
  if (!orders || !items) throw new Error('找不到訂單工作表，請先在試算表執行「禾煎訂單 → ① 初始化工作表」');
  o.id = nextOrderId_(o.createdAt);
  const summary = o.lines.map(l => `${l.name} ×${l.qty}`).join('、');
  orders.appendRow([o.id, o.createdAt, '新訂單', MODES[o.mode], o.store, o.customer, o.phone, o.email, o.address, summary, o.count, o.subtotal, o.shipping, o.total, o.note].map(cell_));
  const rows = o.lines.map(l => [o.id, o.createdAt, l.id, l.name, l.price, l.qty, l.amount].map(cell_));
  items.getRange(items.getLastRow() + 1, 1, rows.length, ITEM_HEADERS.length).setValues(rows);
}

/** HJ-20261127-001：每日從 001 起算（在鎖定內呼叫，避免重號）。 */
function nextOrderId_(date) {
  const day = Utilities.formatDate(date, TZ, 'yyyyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = 'seq_' + day;
  const n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return `HJ-${day}-${String(n).padStart(3, '0')}`;
}

/** 只計算成功建立的訂單；同一電話一小時內超過上限就暫時擋下。 */
function checkRateLimit_(phone, max) {
  const n = Number(CacheService.getScriptCache().get(rateKey_(phone)) || 0);
  if (n >= max) throw userError_('短時間內下單次數過多，請稍後再試或直接來電');
}

function countOrder_(phone) {
  const cache = CacheService.getScriptCache(), key = rateKey_(phone);
  cache.put(key, String(Number(cache.get(key) || 0) + 1), 3600);
}

function rateKey_(phone) {
  return 'rl_' + phone.replace(/\D/g, '');
}

/* ───────────────────────── 通知信 ───────────────────────── */

function notify_(o) {
  const s = readSettings_();
  const owner = String(s['通知信箱'] || '').trim();
  const table = linesTable_(o);
  const where = o.mode === 'delivery' ? `宅配地址：${esc_(o.address)}` : `門市：${esc_(o.store)}`;
  if (owner) {
    MailApp.sendEmail({
      to: owner,
      subject: `【禾煎新訂單】${o.id}｜${MODES[o.mode]}｜NT$ ${o.total}`,
      htmlBody: `<h2 style="margin:0 0 8px">新訂單 ${o.id}</h2>
        <p>${MODES[o.mode]}｜${where}<br>顧客：${esc_(o.customer)}｜${esc_(o.phone)}｜${esc_(o.email)}</p>${table}
        <p><a href="${SpreadsheetApp.getActive().getUrl()}">開啟訂單試算表</a></p>`,
    });
  }
  if (o.email && s['寄送顧客確認信'] !== '否') {
    MailApp.sendEmail({
      to: o.email,
      name: '禾煎 HOJIAN',
      replyTo: owner.split(',')[0] || undefined,
      subject: `禾煎已收到你的訂單 ${o.id}`,
      htmlBody: `<p>${esc_(o.customer)} 你好，謝謝你把一餐交給禾煎。</p>
        <p>訂單編號：<b>${o.id}</b><br>${MODES[o.mode]}｜${where}</p>${table}
        <p>我們會盡快為你準備，若有問題會以電話與你聯繫。</p><p>禾煎 HOJIAN｜把日子，煎得剛剛好。</p>`,
    });
  }
}

function linesTable_(o) {
  const td = 'style="padding:6px 10px;border-bottom:1px solid #ded5c8"';
  const rows = o.lines.map(l => `<tr><td ${td}>${esc_(l.name)}</td><td ${td} align="right">${l.qty}</td><td ${td} align="right">NT$ ${l.amount}</td></tr>`).join('');
  return `<table style="border-collapse:collapse;font-size:14px">${rows}
    <tr><td ${td}>商品小計</td><td ${td}></td><td ${td} align="right">NT$ ${o.subtotal}</td></tr>
    <tr><td ${td}>運費</td><td ${td}></td><td ${td} align="right">NT$ ${o.shipping}</td></tr>
    <tr><td style="padding:6px 10px"><b>合計</b></td><td></td><td style="padding:6px 10px" align="right"><b>NT$ ${o.total}</b></td></tr></table>`;
}

function sendTestEmail() {
  const owner = String(readSettings_()['通知信箱'] || '').trim();
  const ui = SpreadsheetApp.getUi();
  if (!owner) return ui.alert('請先在「設定」工作表填寫通知信箱。');
  MailApp.sendEmail(owner, '【禾煎】測試通知信', '收到這封信，代表新訂單通知設定成功。');
  ui.alert(`測試信已寄到 ${owner}。今天還可寄送約 ${MailApp.getRemainingDailyQuota()} 封。`);
}

/* ───────────────────────── 營收摘要 ───────────────────────── */

function todaySummary_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.orders);
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd');
  const rows = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, ORDER_HEADERS.length).getValues() : [];
  const list = rows.filter(r => r[1] instanceof Date && Utilities.formatDate(r[1], TZ, 'yyyy/MM/dd') === today && r[2] !== '已取消');
  const byMode = {};
  list.forEach(r => { byMode[r[3]] = (byMode[r[3]] || 0) + 1; });
  return { today, count: list.length, revenue: list.reduce((a, r) => a + Number(r[13] || 0), 0), byMode };
}

function summaryText_(x) {
  const modes = Object.keys(x.byMode).map(k => `${k} ${x.byMode[k]} 筆`).join('、') || '—';
  return `${x.today} 營收摘要（不含已取消）\n訂單數：${x.count} 筆\n營收合計：NT$ ${x.revenue.toLocaleString()}\n用餐方式：${modes}`;
}

function showTodaySummary() {
  SpreadsheetApp.getUi().alert(summaryText_(todaySummary_()));
}

function dailyReport() {
  const owner = String(readSettings_()['通知信箱'] || '').trim();
  if (owner) MailApp.sendEmail(owner, `【禾煎】每日營收 ${Utilities.formatDate(new Date(), TZ, 'MM/dd')}`, summaryText_(todaySummary_()));
}

function installDailyReport() {
  removeDailyReport(true);
  ScriptApp.newTrigger('dailyReport').timeBased().atHour(21).everyDays(1).inTimezone(TZ).create();
  SpreadsheetApp.getUi().alert('已開啟：每天晚上 9 點寄送當日營收摘要到通知信箱。');
}

function removeDailyReport(silent) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyReport').forEach(t => ScriptApp.deleteTrigger(t));
  if (silent !== true) SpreadsheetApp.getUi().alert('已關閉每日營收報表。');
}

/* ───────────────────────── 小工具 ───────────────────────── */

function readSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.settings);
  const out = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { out[r[0]] = r[1]; });
  return out;
}

function setSetting_(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.settings);
  const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]);
  const i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(value);
}

function readProducts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.products);
  const out = {};
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, PRODUCT_HEADERS.length).getValues().forEach(r => {
      const id = String(r[0]).trim(), price = Number(r[2]);
      if (id && price > 0 && r[4] !== '否') out[id] = { name: String(r[1]), price };
    });
  }
  return out;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max);
}

/** 避免以 = + - @ 開頭的文字被試算表當成公式執行。 */
function cell_(v) {
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}

function esc_(v) {
  return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function userError_(msg) {
  const e = new Error(msg);
  e.userMessage = msg;
  return e;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
