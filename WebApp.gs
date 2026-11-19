// ===== GIAO DIỆN WEB NHẬP LIỆU (mở bằng link trên điện thoại) =====
function doGet() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  try {
    SpreadsheetApp.getActive().getName(); // chỉ người được chia sẻ Sheet mới mở được
  } catch (e) {
    return HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<div style="font:18px Times New Roman;padding:24px;max-width:520px;margin:auto">' +
      '<h2>🔒 Chưa được cấp quyền</h2><p>Tài khoản <b>' + (email || 'của bạn') + '</b> chưa được chia sẻ file kho.</p>' +
      '<p>Nhờ chủ quán mở Google Sheet → bấm <b>Chia sẻ</b> → thêm email này với quyền <b>Người chỉnh sửa</b>, rồi tải lại trang.</p></div>')
      .setTitle('Chưa được cấp quyền');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Kho Lẩu gà')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function wa_tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function wa_fmt_(d) { return (d instanceof Date) ? Utilities.formatDate(d, wa_tz_(), 'yyyy-MM-dd') : ''; }
function wa_date_(s) { var p = String(s).split('-'); return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])); }
function wa_num_(v) {
  if (v === '' || v === null || v === undefined) return '';
  var n = Number(String(v).replace(',', '.'));
  return isNaN(n) ? String(v) : n;
}
function wa_isHead_(n) { n = String(n); return n === '' || n.charAt(0) === '▶' || n.charAt(0) === '■'; }

function wa_danhMuc() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Danh mục');
  var v = sh.getRange(5, 3, 296, 6).getValues(); // C..H
  var out = [];
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim()) out.push({ n: String(v[i][0]).trim(), u: String(v[i][1]), ncc: String(v[i][5]) });
  return out;
}

// ---------- 1. PHIẾU KIỂM KHO NGÀY ----------
function wa_getPhieu() {
  wa_need_('kk');
  var sh = SpreadsheetApp.getActive().getSheetByName('Phiếu kiểm kho ngày');
  var b2 = sh.getRange('B2').getValue();
  var v = sh.getRange(5, 2, 44, 9).getDisplayValues(); // B..J
  var rows = [];
  for (var i = 0; i < v.length; i++) {
    var n = v[i][0];
    if (n === '' || n === 'MẶT HÀNG') continue;
    rows.push({ r: i + 5, n: n, head: wa_isHead_(n), u: v[i][1], d: v[i][2], e: v[i][3], tong: v[i][4], may: v[i][5], lech: v[i][6], kq: v[i][7], note: v[i][8] });
  }
  return { date: wa_fmt_(b2), nguoi: sh.getRange('E2').getDisplayValue(), rows: rows };
}

function wa_savePhieu(p) {
  wa_need_('kk');
  var lock = LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName('Phiếu kiểm kho ngày');
    if (p.date) sh.getRange('B2').setValue(wa_date_(p.date));
    if (p.nguoi !== undefined) sh.getRange('E2').setValue(p.nguoi);
    (p.rows || []).forEach(function (x) {
      if (x.r < 5 || x.r > 48) return;
      var name = sh.getRange(x.r, 2).getValue();
      if (String(name) !== String(x.n)) throw new Error('Danh sách đã thay đổi, tải lại trang rồi nhập lại: ' + x.n);
      sh.getRange(x.r, 4, 1, 2).setValues([[wa_num_(x.d), wa_num_(x.e)]]);
      sh.getRange(x.r, 10).setValue(x.note || '');
    });
  } finally { lock.releaseLock(); }
  SpreadsheetApp.flush();
  return wa_getPhieu();
}

// ---------- 2. ORDER BẾP ----------
var WA_ORDER_TABS = ['Order Rau - Thịt', 'Order Hàng khô'];
function wa_getOrder(tab) {
  wa_need_('od');
  if (WA_ORDER_TABS.indexOf(tab) < 0) throw new Error('Sai tab');
  var sh = SpreadsheetApp.getActive().getSheetByName(tab);
  var last = Math.max(sh.getLastRow(), 5);
  var v = sh.getRange(5, 3, last - 4, 6).getDisplayValues(); // C..H
  var rows = [];
  for (var i = 0; i < v.length; i++) {
    var n = v[i][0]; if (!n) continue;
    rows.push({ r: i + 5, n: n, head: n.charAt(0) === '▶', u: v[i][1], dat: v[i][2], ncc: v[i][3], nhan: v[i][4], note: v[i][5] });
  }
  return { tab: tab, date: wa_fmt_(sh.getRange('B2').getValue()), rows: rows };
}

function wa_saveOrder(p) {
  wa_need_('od');
  if (WA_ORDER_TABS.indexOf(p.tab) < 0) throw new Error('Sai tab');
  var lock = LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(p.tab);
    if (p.date) sh.getRange('B2').setValue(wa_date_(p.date));
    (p.rows || []).forEach(function (x) {
      if (String(sh.getRange(x.r, 3).getValue()) !== String(x.n)) throw new Error('Danh sách đã thay đổi, tải lại trang: ' + x.n);
      sh.getRange(x.r, 5).setValue(wa_num_(x.dat));
      sh.getRange(x.r, 7, 1, 2).setValues([[wa_num_(x.nhan), x.note || '']]);
    });
  } finally { lock.releaseLock(); }
  SpreadsheetApp.flush();
  return wa_getOrder(p.tab);
}

// ---------- 3. XUẤT KHO / NHẬP HÀNG ----------
function wa_firstEmpty_(sh, col) {
  var v = sh.getRange(5, col, sh.getMaxRows() - 4, 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]) !== '') return i + 6;
  return 5;
}

function wa_addXuat(p) {
  wa_need_('xn');
  var lock = LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName('Xuất kho');
    var r = wa_firstEmpty_(sh, 2);
    var d = wa_date_(p.date);
    (p.items || []).forEach(function (x, i) {
      sh.getRange(r + i, 1, 1, 3).setValues([[d, x.n, wa_num_(x.sl)]]);
      sh.getRange(r + i, 1).setNumberFormat('dd/mm/yyyy');
      sh.getRange(r + i, 5, 1, 3).setValues([[p.nguoi || '', p.cho || '', x.note || '']]);
    });
  } finally { lock.releaseLock(); }
  return 'Đã lưu ' + (p.items || []).length + ' dòng xuất kho.';
}

function wa_addNhap(p) {
  wa_need_('xn');
  var lock = LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName('Nhập hàng');
    var r = wa_firstEmpty_(sh, 2);
    var d = wa_date_(p.date);
    (p.items || []).forEach(function (x, i) {
      sh.getRange(r + i, 1, 1, 2).setValues([[d, x.n]]);
      sh.getRange(r + i, 1).setNumberFormat('dd/mm/yyyy');
      sh.getRange(r + i, 4, 1, 3).setValues([[wa_num_(x.gia), wa_num_(x.sl), wa_num_(x.kg)]]);
      sh.getRange(r + i, 10, 1, 3).setValues([[p.tt || '', p.nguoi || '', x.note || '']]);
    });
  } finally { lock.releaseLock(); }
  return 'Đã lưu ' + (p.items || []).length + ' dòng nhập hàng.';
}

function wa_recent(tab) {
  wa_need_('xn');
  var sh = SpreadsheetApp.getActive().getSheetByName(tab === 'nhap' ? 'Nhập hàng' : 'Xuất kho');
  var r = wa_firstEmpty_(sh, 2) - 1;
  var from = Math.max(5, r - 14);
  if (r < 5) return [];
  var w = tab === 'nhap' ? 13 : 8;
  return sh.getRange(from, 1, r - from + 1, w).getDisplayValues().reverse();
}

// ---------- 4. XEM ĐỐI CHIẾU ----------
function wa_getDoiChieu() {
  wa_need_('dc');
  var ss = SpreadsheetApp.getActive();
  var a = ss.getSheetByName('Đối chiếu NL');
  var v = a.getRange(5, 2, 38, 10).getDisplayValues(); // B..K
  var rows = [];
  for (var i = 0; i < v.length; i++) if (v[i][0]) rows.push({ n: v[i][0], head: wa_isHead_(v[i][0]), u: v[i][1], dauN: v[i][2], dauL: v[i][3], nhap: v[i][4], cuoi: v[i][5], dung: v[i][6], ban: v[i][7], lech: v[i][8], kq: v[i][9] });
  var k = ss.getSheetByName('Kiểm kê tuần');
  var kv = k.getRange(6, 2, 98, 7).getDisplayValues(); // B..H
  var krows = [];
  for (var j = 0; j < kv.length; j++) if (kv[j][0]) krows.push({ n: kv[j][0], u: kv[j][1], may: kv[j][2], thuc: kv[j][3], lech: kv[j][4], kq: kv[j][5], note: kv[j][6] });
  return { ngay: a.getRange('B2').getDisplayValue(), ngayPhieu: a.getRange('E2').getDisplayValue(), rows: rows,
           kkDen: k.getRange('C3').getDisplayValue(), kkTu: k.getRange('C4').getDisplayValue(), kkNote: k.getRange('E3').getDisplayValue(), krows: krows };
}


// ---------- PHÂN QUYỀN THEO EMAIL ----------
var WA_PERM_TAB = 'Phân quyền';
var WA_KEYS = ['kk', 'od', 'xn', 'dc'];
function wa_permSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(WA_PERM_TAB);
  if (!sh) {
    sh = ss.insertSheet(WA_PERM_TAB);
    sh.getRange(1, 1, 1, 6).setValues([['Email', 'Tên', '📋 Kiểm kho', '🛒 Order', '📦 Xuất/Nhập', '📊 Đối chiếu']]).setFontWeight('bold').setBackground('#cfe0ff');
    sh.setFrozenRows(1);
    sh.getRange('C2:F200').insertCheckboxes();
    sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 160);
  }
  return sh;
}
function wa_perm_() {
  var email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  var owner = '';
  try { owner = String(SpreadsheetApp.getActive().getOwner().getEmail()).toLowerCase(); } catch (e) {}
  if (email && email === owner) return { email: email, admin: true, kk: true, od: true, xn: true, dc: true };
  var p = { email: email, admin: false, kk: false, od: false, xn: false, dc: false };
  var sh = SpreadsheetApp.getActive().getSheetByName(WA_PERM_TAB);
  if (!sh || !email) return p;
  var v = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 6).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim().toLowerCase() === email) {
    p.name = v[i][1]; p.kk = v[i][2] === true; p.od = v[i][3] === true; p.xn = v[i][4] === true; p.dc = v[i][5] === true;
  }
  return p;
}
function wa_need_(k) {
  var p = wa_perm_();
  if (!p[k]) throw new Error('Tài khoản ' + (p.email || '') + ' chưa được cấp quyền mục này. Nhờ chủ quán tick quyền ở mục ⚙ Phân quyền.');
  return p;
}
function wa_me() { return wa_perm_(); }
function wa_getPerms() {
  if (!wa_perm_().admin) throw new Error('Chỉ chủ quán mới xem được phân quyền.');
  var sh = wa_permSheet_();
  var n = Math.max(sh.getLastRow() - 1, 0);
  var v = n ? sh.getRange(2, 1, n, 6).getValues() : [];
  return v.filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    return { email: String(r[0]).trim(), name: String(r[1]), kk: r[2] === true, od: r[3] === true, xn: r[4] === true, dc: r[5] === true };
  });
}
function wa_savePerms(rows) {
  if (!wa_perm_().admin) throw new Error('Chỉ chủ quán mới sửa được phân quyền.');
  var ss = SpreadsheetApp.getActive();
  var sh = wa_permSheet_();
  var clean = (rows || []).filter(function (r) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(r.email).trim()); });
  sh.getRange(2, 1, 199, 2).clearContent();
  sh.getRange(2, 3, 199, 4).uncheck();
  if (clean.length) sh.getRange(2, 1, clean.length, 6).setValues(clean.map(function (r) {
    return [String(r.email).trim().toLowerCase(), r.name || '', !!r.kk, !!r.od, !!r.xn, !!r.dc];
  }));
  var editors = ss.getEditors().map(function (u) { return String(u.getEmail()).toLowerCase(); });
  var added = [];
  clean.forEach(function (r) {
    var e = String(r.email).trim().toLowerCase();
    if ((r.kk || r.od || r.xn || r.dc) && editors.indexOf(e) < 0) { ss.addEditor(e); added.push(e); }
  });
  return 'Đã lưu ' + clean.length + ' người.' + (added.length ? ' Đã chia sẻ Sheet cho: ' + added.join(', ') : '');
}
