/**
 * Sales.gs — Innonext Sales (ค้นบาร์โค้ด Central + บันทึกยอดขายหน้าร้าน) v3 (29 ต.ค. 2026 — เปิดเฉพาะสาขาที่กำหนด + ตาราง CEO)
 * Repo: github.com/thipwonh-ship-it/innonext-sales (ไฟล์นี้วางใน Apps Script โปรเจกต์ "Innonext HR" เพื่อใช้ตัวตน LINE/พนักงานร่วมกัน)
 *
 * ใช้ร่วมกับ Innonext HR: ss() (ชีตพนักงาน), findByDevice, lineSession, TZ, linePushTo, lineReply, isAdminUid, execs/execByUid, linePushAdmin
 * ข้อมูลอยู่ใน Google Sheet แยก (Script Property SALES_SHEET_ID) — สร้างอัตโนมัติครั้งแรกที่เรียก setupSales()
 *   แผ่น "สินค้า":  แบรนด์ · รุ่น · สี · ไซส์(มม.) · ไซส์ Central · บาร์โค้ด Central · ชื่อ Central · บาร์โค้ดเกาหลี · SKU · ราคาเต็ม · หมายเหตุ
 *   แผ่น "ยอดขาย": วันที่ · เดือน · สาขา · พนักงาน · แบรนด์ · รุ่น · สี · ไซส์ · บาร์โค้ด Central · บาร์โค้ดเกาหลี · จำนวน · ราคาเต็ม · ราคาขาย · ส่วนลด · ยอดรวม · หมายเหตุ · ยกเลิก
 * สิทธิ์: พนักงานกลุ่ม "หน้าร้าน" (ปุ่ม 🛍 ขาย ในหน้าลงเวลา) · CEO/แอดมิน พิมพ์ "sales" / "ยอดขาย" ในแชท OA · สรุปยอดรายวันส่ง CEO 20:30
 * ⚠ ไม่มี secret ในไฟล์นี้
 */
const SALES_PROP = 'SALES_SHEET_ID';
const SL_PROD = 'สินค้า', SL_SALE = 'ยอดขาย';
const SP_BRAND = 1, SP_MODEL = 2, SP_COLOR = 3, SP_SIZE = 4, SP_CSIZE = 5, SP_CBAR = 6, SP_CNAME = 7, SP_KBAR = 8, SP_SKU = 9, SP_PRICE = 10, SP_NOTE = 11;
const SS_DATE = 1, SS_MONTH = 2, SS_LOC = 3, SS_NAME = 4, SS_BRAND = 5, SS_MODEL = 6, SS_COLOR = 7, SS_SIZE = 8, SS_CBAR = 9, SS_KBAR = 10, SS_QTY = 11, SS_FULL = 12, SS_PRICE = 13, SS_DISC = 14, SS_TOTAL = 15, SS_NOTE = 16, SS_VOID = 17;
const SALES_SUMMARY_HOUR = 20; // ส่งสรุปยอดวันนี้ให้ CEO เวลา 20:xx
const SALES_LOCS_DEFAULT = 'Mega Bangna'; // v2: สาขาที่เปิดใช้ระบบขาย (ชื่อตรงกับคอลัมน์ ร้าน/สาขา ในแผ่น พนักงาน) — เพิ่มสาขาได้ที่ Script Property SALES_LOCS เช่น "Mega Bangna, Central World" ไม่ต้องแก้โค้ด
/** v2: สาขานี้เปิดใช้ระบบขายหรือยัง */
function salesAllowedLoc(loc) {
  const v = PropertiesService.getScriptProperties().getProperty('SALES_LOCS');
  return String(v || SALES_LOCS_DEFAULT).split(',').map(x => x.trim()).filter(Boolean).indexOf(String(loc || '').trim()) >= 0;
}

function salesBook() {
  const P = PropertiesService.getScriptProperties(); let id = P.getProperty(SALES_PROP);
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const book = SpreadsheetApp.create('Innonext Sales'); P.setProperty(SALES_PROP, book.getId()); return book;
}
/** รันครั้งเดียว: สร้างชีต + แผ่น + ทริกเกอร์สรุปรายวัน */
function setupSales() {
  const book = salesBook(); const out = ['ชีต: ' + book.getUrl()];
  let p = book.getSheetByName(SL_PROD);
  if (!p) {
    p = book.insertSheet(SL_PROD, 0);
    p.getRange(1, 1, 1, SP_NOTE).setValues([['แบรนด์', 'รุ่น', 'สี', 'ไซส์ (มม.)', 'ไซส์ Central', 'บาร์โค้ด Central', 'ชื่อ Central', 'บาร์โค้ดเกาหลี', 'SKU', 'ราคาเต็ม', 'หมายเหตุ']]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
    [[1, 90], [2, 130], [3, 90], [4, 80], [5, 90], [6, 120], [7, 240], [8, 120], [9, 130], [10, 80], [11, 260]].forEach(a => p.setColumnWidth(a[0], a[1]));
    [SP_CBAR, SP_KBAR, SP_SKU, SP_CSIZE].forEach(c => p.getRange(2, c, 3000, 1).setNumberFormat('@'));
    p.getRange(2, SP_PRICE, 3000, 1).setNumberFormat('#,##0'); p.setFrozenRows(1); out.push('สร้างแผ่น ' + SL_PROD);
  }
  let s = book.getSheetByName(SL_SALE);
  if (!s) {
    s = book.insertSheet(SL_SALE);
    s.getRange(1, 1, 1, SS_VOID).setValues([['วันที่', 'เดือน', 'สาขา', 'พนักงาน', 'แบรนด์', 'รุ่น', 'สี', 'ไซส์', 'บาร์โค้ด Central', 'บาร์โค้ดเกาหลี', 'จำนวน', 'ราคาเต็ม', 'ราคาขาย', 'ส่วนลด', 'ยอดรวม', 'หมายเหตุ', 'ยกเลิก']]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
    [[1, 130], [2, 75], [3, 120], [4, 160], [5, 80], [6, 120], [7, 80], [8, 60], [9, 110], [10, 110], [11, 60], [12, 80], [13, 80], [14, 70], [15, 90], [16, 200], [17, 60]].forEach(a => s.setColumnWidth(a[0], a[1]));
    s.getRange(2, SS_DATE, 5000, 1).setNumberFormat('dd/MM/yyyy HH:mm'); [SS_CBAR, SS_KBAR, SS_NOTE].forEach(c => s.getRange(2, c, 5000, 1).setNumberFormat('@'));
    [SS_FULL, SS_PRICE, SS_DISC, SS_TOTAL].forEach(c => s.getRange(2, c, 5000, 1).setNumberFormat('#,##0.00'));
    s.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('✓').setBackground('#F4CCCC').setRanges([s.getRange('Q2:Q5000')]).build()]);
    s.setFrozenRows(1); out.push('สร้างแผ่น ' + SL_SALE);
  }
  const d = book.getSheetByName('Sheet1') || book.getSheetByName('ชีต1'); if (d && book.getSheets().length > 1) book.deleteSheet(d);
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'salesDailySummary') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('salesDailySummary').timeBased().everyDays(1).atHour(SALES_SUMMARY_HOUR).inTimezone(TZ).create(); out.push('ตั้งสรุปรายวัน ' + SALES_SUMMARY_HOUR + ':xx');
  const r = out.join(' | '); Logger.log(r); return r;
}

/* ---------- สินค้า ---------- */
function salesProducts() {
  const c = CacheService.getScriptCache(); const k = 'sales_prod'; const hit = c.get(k); if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const sh = salesBook().getSheetByName(SL_PROD); const last = sh ? sh.getLastRow() : 0; const list = [];
  if (last >= 2) sh.getRange(2, 1, last - 1, SP_NOTE).getValues().forEach(r => { if (!r[SP_MODEL - 1] && !r[SP_CBAR - 1]) return;
    list.push({ brand: String(r[SP_BRAND - 1]).trim(), model: String(r[SP_MODEL - 1]).trim(), color: String(r[SP_COLOR - 1]).trim(), size: String(r[SP_SIZE - 1]).trim(), csize: String(r[SP_CSIZE - 1]).trim(), cbar: String(r[SP_CBAR - 1]).trim(), cname: String(r[SP_CNAME - 1]).trim(), kbar: String(r[SP_KBAR - 1]).trim(), sku: String(r[SP_SKU - 1]).trim(), price: Number(r[SP_PRICE - 1]) || 0, note: String(r[SP_NOTE - 1]).trim() }); });
  try { const js = JSON.stringify(list); if (js.length < 95000) c.put(k, js, 600); } catch (e) {}
  return list;
}
function salesWho(ident) {
  const e = findByDevice(ident); if (!e) return null;
  if (String(e.group || '').trim() !== 'หน้าร้าน' && !e.mgr) return { denied: true, msg: 'ใช้ได้เฉพาะพนักงานหน้าร้าน' };
  if (!salesAllowedLoc(e.loc)) return { denied: true, msg: 'สาขานี้ยังไม่เปิดใช้ระบบขาย' };
  return e;
}
/** ค้นหา: บาร์โค้ดเกาหลี / บาร์โค้ด Central / SKU / รุ่น+สี+ไซส์ (พิมพ์คั่นด้วยช่องว่าง) — คืนไม่เกิน 30 รายการ */
function salesSearch(ident, q) {
  const e = salesWho(ident); if (!e) return { ok: false, msg: 'ไม่พบผู้ใช้' }; if (e.denied) return { ok: false, msg: e.msg };
  q = String(q || '').trim(); if (!q) return { ok: true, list: [] };
  const all = salesProducts(); const digits = q.replace(/\D/g, '');
  let list;
  if (/^\d{6,}$/.test(q)) list = all.filter(p => p.kbar === q || p.cbar === q || p.kbar.endsWith(q) || p.cbar.endsWith(q));
  else { const toks = q.toLowerCase().split(/\s+/).filter(Boolean); list = all.filter(p => { const hay = (p.model + ' ' + p.color + ' ' + p.size + ' ' + p.csize + ' ' + p.sku + ' ' + p.cname).toLowerCase(); return toks.every(t => hay.includes(t)); }); }
  if (!list.length && digits.length >= 4) list = all.filter(p => p.kbar.includes(digits) || p.cbar.includes(digits));
  return { ok: true, list: list.slice(0, 30), total: list.length, q };
}
/* ---------- บันทึกยอดขาย ---------- */
function salesInfo(ident) {
  const e = salesWho(ident); if (!e) return { ok: false, msg: 'ไม่พบผู้ใช้' }; if (e.denied) return { ok: false, msg: e.msg };
  const sh = salesBook().getSheetByName(SL_SALE); const last = sh ? sh.getLastRow() : 0; const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); const mine = []; let tQty = 0, tSum = 0;
  if (last >= 2) sh.getRange(Math.max(2, last - 300), 1, Math.min(301, last - 1), SS_VOID).getValues().forEach((r, i) => { if (!(r[0] instanceof Date)) return; if (Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') !== today || String(r[SS_LOC - 1]).trim() !== e.loc) return;
    const v = String(r[SS_VOID - 1]) === '✓'; const row = Math.max(2, last - 300) + i;
    mine.push({ row, time: Utilities.formatDate(r[0], TZ, 'HH:mm'), by: String(r[SS_NAME - 1]).split(' ')[0], item: r[SS_MODEL - 1] + ' ' + r[SS_COLOR - 1] + ' ' + r[SS_SIZE - 1], qty: Number(r[SS_QTY - 1]) || 0, price: Number(r[SS_PRICE - 1]) || 0, total: Number(r[SS_TOTAL - 1]) || 0, voided: v, mine: String(r[SS_NAME - 1]).trim() === e.name });
    if (!v) { tQty += Number(r[SS_QTY - 1]) || 0; tSum += Number(r[SS_TOTAL - 1]) || 0; } });
  return { ok: true, name: e.name, loc: e.loc, today: mine.reverse(), qty: tQty, sum: tSum };
}
/** บันทึกขาย 1 รายการ: cbar = บาร์โค้ด Central ของสินค้าที่เลือก · price = ราคาขายจริง (ว่าง = ราคาเต็ม) */
function logSale(ident, cbar, qty, price, note) {
  const e = salesWho(ident); if (!e) return { ok: false, msg: 'ไม่พบผู้ใช้' }; if (e.denied) return { ok: false, msg: e.msg };
  const p = salesProducts().find(x => x.cbar === String(cbar || '').trim()); if (!p) return { ok: false, msg: 'ไม่พบสินค้า' };
  qty = Math.max(1, Math.min(99, Math.round(Number(qty) || 1))); const full = p.price; let sell = price === '' || price == null ? full : Number(price);
  if (!(sell >= 0)) return { ok: false, msg: 'ราคาไม่ถูกต้อง' };
  const sh = salesBook().getSheetByName(SL_SALE); const now = new Date(); const row = sh.getLastRow() + 1; const disc = Math.max(0, (full - sell) * qty); const total = sell * qty;
  [SS_CBAR, SS_KBAR, SS_NOTE].forEach(c => sh.getRange(row, c).setNumberFormat('@'));
  sh.getRange(row, 1, 1, SS_VOID).setValues([[now, Utilities.formatDate(now, TZ, 'yyyy-MM'), e.loc, e.name, p.brand, p.model, p.color, p.csize || p.size, p.cbar, p.kbar, qty, full, sell, disc, total, String(note || '').slice(0, 200), '']]);
  sh.getRange(row, SS_DATE).setNumberFormat('dd/MM/yyyy HH:mm');
  return { ok: true, msg: 'บันทึกแล้ว ' + p.model + ' ' + p.color + ' ' + (p.csize || p.size) + ' ×' + qty + ' = ' + total.toLocaleString('en-US') + ' บาท' + (disc ? ' (ลด ' + disc.toLocaleString('en-US') + ')' : ''), row };
}
/** ยกเลิกรายการของตัวเองในวันนี้ (ผู้จัดการ/แอดมินยกเลิกของใครก็ได้) */
function voidSale(ident, row, adminPin) {
  const admin = adminPin && isAdminPin(adminPin); const e = admin ? null : salesWho(ident); if (!admin && (!e || e.denied)) return { ok: false, msg: 'ไม่มีสิทธิ์' };
  const sh = salesBook().getSheetByName(SL_SALE); row = Number(row); if (!(row >= 2) || row > sh.getLastRow()) return { ok: false, msg: 'ไม่พบรายการ' };
  const r = sh.getRange(row, 1, 1, SS_VOID).getValues()[0]; if (!(r[0] instanceof Date)) return { ok: false, msg: 'ไม่พบรายการ' };
  if (!admin && String(r[SS_NAME - 1]).trim() !== e.name && !e.mgr) return { ok: false, msg: 'ยกเลิกได้เฉพาะรายการของตัวเอง' };
  if (!admin && Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') !== Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')) return { ok: false, msg: 'ยกเลิกได้เฉพาะรายการวันนี้ — แจ้งแอดมิน' };
  if (String(r[SS_VOID - 1]) === '✓') return { ok: false, msg: 'ยกเลิกไปแล้ว' };
  sh.getRange(row, SS_VOID).setValue('✓'); sh.getRange(row, SS_NOTE).setValue((String(r[SS_NOTE - 1]) ? String(r[SS_NOTE - 1]) + ' | ' : '') + 'ยกเลิกโดย ' + (admin ? 'แอดมิน' : e.name.split(' ')[0]) + ' ' + Utilities.formatDate(new Date(), TZ, 'd/M HH:mm'));
  return { ok: true, msg: 'ยกเลิกแล้ว' };
}
/* ---------- สรุปยอด (CEO / แอดมิน) ---------- */
/** รวมยอดช่วงวันที่ [d0,d1) แยกสาขา + top รุ่น */
function salesAgg(d0, d1) {
  const sh = salesBook().getSheetByName(SL_SALE); const last = sh ? sh.getLastRow() : 0; const byLoc = {}, byModel = {}; let qty = 0, sum = 0, disc = 0, n = 0;
  if (last >= 2) sh.getRange(2, 1, last - 1, SS_VOID).getValues().forEach(r => { if (!(r[0] instanceof Date) || r[0] < d0 || r[0] >= d1 || String(r[SS_VOID - 1]) === '✓') return;
    const q = Number(r[SS_QTY - 1]) || 0, t = Number(r[SS_TOTAL - 1]) || 0; n++; qty += q; sum += t; disc += Number(r[SS_DISC - 1]) || 0;
    const L = String(r[SS_LOC - 1]).trim(); byLoc[L] = byLoc[L] || { qty: 0, sum: 0 }; byLoc[L].qty += q; byLoc[L].sum += t;
    const M = String(r[SS_MODEL - 1]).trim(); byModel[M] = (byModel[M] || 0) + q; });
  const top = Object.keys(byModel).sort((a, b) => byModel[b] - byModel[a]).slice(0, 5).map(m => m + ' ' + byModel[m]);
  return { n, qty, sum, disc, byLoc, top };
}
function salesText(en) {
  const now = new Date(); const t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate()); const t1 = new Date(t0.getTime() + 86400000); const m0 = new Date(now.getFullYear(), now.getMonth(), 1);
  const d = salesAgg(t0, t1), m = salesAgg(m0, t1); const f = n => Math.round(n).toLocaleString('en-US');
  const locLine = a => Object.keys(a.byLoc).map(L => '  • ' + L + ': ' + a.byLoc[L].qty + (en ? ' pcs · ' : ' ชิ้น · ') + f(a.byLoc[L].sum) + (en ? ' THB' : ' บาท')).join('\n');
  if (en) return '🛍 Sales — ' + Utilities.formatDate(now, TZ, 'EEE d MMM yyyy') + '\n\nToday: ' + d.qty + ' pcs · ' + f(d.sum) + ' THB' + (d.disc ? ' (discounts ' + f(d.disc) + ')' : '') + (d.qty ? '\n' + locLine(d) : '') +
    '\n\nMonth to date (' + Utilities.formatDate(now, TZ, 'MMM yyyy') + '): ' + m.qty + ' pcs · ' + f(m.sum) + ' THB' + (m.disc ? ' (discounts ' + f(m.disc) + ')' : '') + (m.qty ? '\n' + locLine(m) : '') + (m.top.length ? '\nTop models: ' + m.top.join(', ') : '') + '\n\nType "sales" any time.';
  return '🛍 ยอดขาย — ' + Utilities.formatDate(now, TZ, 'd/M/yyyy') + '\n\nวันนี้: ' + d.qty + ' ชิ้น · ' + f(d.sum) + ' บาท' + (d.disc ? ' (ส่วนลด ' + f(d.disc) + ')' : '') + (d.qty ? '\n' + locLine(d) : '') +
    '\n\nเดือนนี้: ' + m.qty + ' ชิ้น · ' + f(m.sum) + ' บาท' + (m.disc ? ' (ส่วนลด ' + f(m.disc) + ')' : '') + (m.qty ? '\n' + locLine(m) : '') + (m.top.length ? '\nขายดี: ' + m.top.join(', ') : '');
}
/** คำสั่งแชท: "sales" / "ยอดขาย" — CEO (อังกฤษ) / แอดมิน (ไทย) · เรียกจาก onExecCommand */
function onSalesCommand(uid, txt, replyToken) {
  if (!/^(sales|ยอดขาย)$/i.test(String(txt || '').trim())) return false;
  const ex = (typeof execByUid === 'function') && !!execByUid(uid); const ad = isAdminUid(uid); if (!ex && !ad) return false;
  let url = ''; try { url = salesCeoSync().url; } catch (e) {}
  lineReply(replyToken, salesText(ex && !ad) + (url ? (ex && !ad ? '\n\n📊 Full table (model · color · size · price): ' : '\n\n📊 ตารางเต็มสำหรับ CEO: ') + url : '')); return true;
}
/** ทริกเกอร์ 20:xx — ส่งยอดวันนี้ให้ CEO ทุกคน (เฉพาะวันที่มียอด) */
function salesDailySummary() {
  const now = new Date(); const t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate()); const d = salesAgg(t0, new Date(t0.getTime() + 86400000)); if (!d.n) return;
  let url = ''; try { url = salesCeoSync().url; } catch (e) {}
  const txt = salesText(true) + (url ? '\n\n📊 Full table (model · color · size · price): ' + url : ''); Object.keys((typeof execs === 'function') ? execs() : {}).forEach(uid => linePushTo(uid, txt));
}

/* ---------- v3: ตารางยอดขายสำหรับ CEO (Google Sheet แยกไฟล์ ไม่มีชื่อพนักงาน · แชร์ด้วยอีเมล Google แบบดูอย่างเดียว) ---------- */
const SALES_CEO_PROP = 'SALES_CEO_SHEET_ID';
const SALES_CEO_HEAD = ['Date', 'Time', 'Branch', 'Brand', 'Model', 'Color', 'Size', 'Qty', 'Full price (THB)', 'Sale price (THB)', 'Discount (THB)', 'Total (THB)', 'Note'];
function salesCeoBook() {
  const P = PropertiesService.getScriptProperties(); const id = P.getProperty(SALES_CEO_PROP);
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const b = SpreadsheetApp.create('Innonext Sales — CEO Report'); P.setProperty(SALES_CEO_PROP, b.getId());
  const sh = b.getSheets()[0]; sh.setName('Sales');
  sh.getRange(1, 1, 1, SALES_CEO_HEAD.length).setValues([SALES_CEO_HEAD]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  [[1, 95], [2, 55], [3, 120], [4, 85], [5, 130], [6, 140], [7, 55], [8, 45], [9, 105], [10, 105], [11, 100], [12, 95], [13, 200]].forEach(a => sh.setColumnWidth(a[0], a[1]));
  sh.setFrozenRows(1); return b;
}
/** สร้างตารางใหม่จากแผ่น ยอดขาย (ไม่รวมรายการที่ยกเลิก · ใหม่สุดอยู่บน) — เรียกตอนพิมพ์ sales / ยอดขาย และทริกเกอร์ 20:xx */
function salesCeoSync() {
  const b = salesCeoBook(); const sh = b.getSheetByName('Sales') || b.getSheets()[0];
  const src = salesBook().getSheetByName(SL_SALE); const last = src ? src.getLastRow() : 0; const out = [];
  if (last >= 2) src.getRange(2, 1, last - 1, SS_VOID).getValues().forEach(r => { if (!(r[0] instanceof Date) || String(r[SS_VOID - 1]) === '✓') return;
    const d = r[0]; out.push([new Date(d.getFullYear(), d.getMonth(), d.getDate()), Utilities.formatDate(d, TZ, 'HH:mm'), r[SS_LOC - 1], r[SS_BRAND - 1], r[SS_MODEL - 1], r[SS_COLOR - 1], String(r[SS_SIZE - 1]), Number(r[SS_QTY - 1]) || 0, Number(r[SS_FULL - 1]) || 0, Number(r[SS_PRICE - 1]) || 0, Number(r[SS_DISC - 1]) || 0, Number(r[SS_TOTAL - 1]) || 0, String(r[SS_NOTE - 1] || '').split(' | ยกเลิกโดย')[0], d.getTime()]); });
  out.sort((a, c) => c[13] - a[13]); const rows = out.map(r => r.slice(0, 13)); const n = rows.length;
  const f = sh.getFilter(); if (f) f.remove();
  const max = sh.getMaxRows(); if (max > 1) sh.getRange(2, 1, max - 1, SALES_CEO_HEAD.length).clearContent();
  if (n) { if (sh.getMaxRows() < n + 1) sh.insertRowsAfter(sh.getMaxRows(), n + 1 - sh.getMaxRows());
    sh.getRange(2, 7, n, 1).setNumberFormat('@'); sh.getRange(2, 1, n, SALES_CEO_HEAD.length).setValues(rows);
    sh.getRange(2, 1, n, 1).setNumberFormat('dd/MM/yyyy'); sh.getRange(2, 9, n, 4).setNumberFormat('#,##0'); }
  sh.getRange(1, 1, Math.max(n, 1) + 1, SALES_CEO_HEAD.length).createFilter();
  return { url: b.getUrl(), n };
}
/** แชร์ตาราง CEO แบบดูอย่างเดียว: salesCeoShare('a@gmail.com, b@gmail.com') — รันจาก IDE (ใส่อีเมลในฟังก์ชันชั่วคราว) */
function salesCeoShare(emails) {
  const list = String(emails || '').split(/[,\s]+/).map(x => x.trim()).filter(x => /@/.test(x)); if (!list.length) return 'ไม่มีอีเมล';
  const b = salesCeoBook(); b.addViewers(list); return 'แชร์แล้ว: ' + list.join(', ') + ' → ' + b.getUrl();
}
