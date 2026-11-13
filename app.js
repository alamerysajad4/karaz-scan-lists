'use strict';
/* ماسح القوائم — تستورد قائمة Excel، وتمسح باركود كل مادة بكاميرا التلفون، وتصدّرها.
   البيانات تنحفظ على التلفون وتتزامن بين كل التلفونات عن طريق Supabase (sync.js). */

const $ = (s, r = document) => r.querySelector(s);
function h(tag, attrs, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (k === 'style' && typeof v === 'object') Object.assign(e.style, v);
    else if (k === 'html') e.innerHTML = v;
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const k of kids.flat(9)) if (k != null && k !== false) e.append(k.nodeType ? k : document.createTextNode(String(k)));
  return e;
}
const digits = s => String(s ?? '').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
const numv = v => { if (v === '' || v == null) return null; const x = parseFloat(digits(v).replace(/[,،\s]/g, '')); return isNaN(x) ? null : x; };
const money = v => v == null || v === '' ? '—' : Math.round(v).toLocaleString('en-US');
const ceil250 = x => Math.ceil(x / 250) * 250;
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
let toastT;
function toast(msg, kind) {
  document.querySelectorAll('.toast').forEach(t => t.remove());
  const t = h('div', { class: 'toast ' + (kind || '') }, msg); document.body.append(t);
  clearTimeout(toastT); toastT = setTimeout(() => t.remove(), kind === 'err' ? 4000 : 2200);
}
function sheet(title, body, footer) {
  const ov = h('div', { class: 'ov', onclick: e => { if (e.target === ov) close(); } });
  const close = () => ov.remove();
  ov.append(h('div', { class: 'sheet' }, h('h2', {}, title), body, footer ? h('div', { class: 'row' }, footer) : null));
  document.body.append(ov); return { close };
}
let actx;
function beep(ok = true) {
  try {
    actx = actx || new (window.AudioContext || window.webkitAudioContext)();
    const o = actx.createOscillator(), g = actx.createGain();
    o.frequency.value = ok ? 1450 : 300; o.type = 'square'; g.gain.value = 0.05;
    o.connect(g); g.connect(actx.destination); o.start(); o.stop(actx.currentTime + (ok ? 0.12 : 0.3));
  } catch (e) { /* */ }
  try { navigator.vibrate && navigator.vibrate(ok ? 70 : [60, 60, 60]); } catch (e) { /* */ }
}

/* ================= التخزين ================= */
const KEY = 'scanapp_v1';
let DB = { lists: [] };
function load() { try { DB = JSON.parse(localStorage.getItem(KEY)) || { lists: [] }; } catch (e) { DB = { lists: [] }; } }
function save() {
  try { localStorage.setItem(KEY, JSON.stringify(DB)); }
  catch (e) { toast('ما انحفظ — ذاكرة المتصفح ممتلئة', 'err'); }
}
load();
if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});

/* ================= قراءة Excel (xlsx) و CSV ================= */
async function inflateRaw(bytes) {
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function unzip(buf) {
  const u8 = new Uint8Array(buf), dv = new DataView(buf);
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('ملف Excel غير صالح');
  const n = dv.getUint16(eocd + 10, true); let p = dv.getUint32(eocd + 16, true);
  const files = {}; const td = new TextDecoder();
  for (let k = 0; k < n; k++) {
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
    const name = td.decode(u8.subarray(p + 46, p + 46 + nlen));
    files[name] = { method, csize, off };
    p += 46 + nlen + xlen + clen;
  }
  return async name => {
    const f = files[name]; if (!f) return null;
    const lnl = dv.getUint16(f.off + 26, true), lxl = dv.getUint16(f.off + 28, true);
    const data = u8.subarray(f.off + 30 + lnl + lxl, f.off + 30 + lnl + lxl + f.csize);
    const out = f.method === 0 ? data : await inflateRaw(data);
    return new TextDecoder().decode(out);
  };
}
function colIdx(ref) { const m = /^([A-Z]+)/.exec(ref || ''); if (!m) return -1; let n = 0; for (const c of m[1]) n = n * 26 + c.charCodeAt(0) - 64; return n - 1; }
async function readXlsx(buf) {
  const get = await unzip(buf);
  const parse = s => new DOMParser().parseFromString(s, 'application/xml');
  const ss = [];
  const sst = await get('xl/sharedStrings.xml');
  if (sst) for (const si of parse(sst).getElementsByTagName('si')) ss.push([...si.getElementsByTagName('t')].map(t => t.textContent).join(''));
  // أول ورقة بالملف
  let sheetPath = 'xl/worksheets/sheet1.xml';
  try {
    const wb = parse(await get('xl/workbook.xml')), rels = parse(await get('xl/_rels/workbook.xml.rels'));
    const first = wb.getElementsByTagName('sheet')[0];
    const rid = first.getAttribute('r:id') || first.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    for (const r of rels.getElementsByTagName('Relationship')) if (r.getAttribute('Id') === rid) { const t = r.getAttribute('Target'); sheetPath = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, ''); }
  } catch (e) { /* sheet1 */ }
  const xml = await get(sheetPath); if (!xml) throw new Error('ما لكيت ورقة بالملف');
  const rows = [];
  for (const r of parse(xml).getElementsByTagName('row')) {
    const ri = parseInt(r.getAttribute('r'), 10) - 1, row = [];
    for (const c of r.getElementsByTagName('c')) {
      const ci = colIdx(c.getAttribute('r')); if (ci < 0) continue;
      const t = c.getAttribute('t'); let v = '';
      if (t === 's') { const x = c.getElementsByTagName('v')[0]; v = x ? ss[+x.textContent] ?? '' : ''; }
      else if (t === 'inlineStr') v = [...c.getElementsByTagName('t')].map(x => x.textContent).join('');
      else { const x = c.getElementsByTagName('v')[0]; v = x ? x.textContent : ''; }
      row[ci] = v;
    }
    rows[isNaN(ri) ? rows.length : ri] = row;
  }
  return rows.filter(Boolean).map(r => Array.from(r, v => v == null ? '' : String(v)));
}
function readCsv(text) {
  text = text.replace(/^﻿/, '');
  const first = text.split(/\r?\n/)[0];
  const d = [',', ';', '\t'].sort((a, b) => first.split(b).length - first.split(a).length)[0];
  const rows = []; let row = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += c; }
    else if (c === '"') q = true;
    else if (c === d) { row.push(cur); cur = ''; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; }
    else cur += c;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows;
}
const NAME_RX = /^(اسم المادة|الاسم|اسم الصنف|الصنف|المادة|التفاصيل|اسم|name|item|product)$/i;
const BC_RX = /^(الباركود|باركود|barcode|bar code|ean)$/i;
function toList(rows, fileName) {
  rows = rows.filter(r => r && r.some(v => String(v).trim() !== ''));
  if (!rows.length) throw new Error('الملف فارغ');
  let hi = rows.findIndex(r => r.some(v => NAME_RX.test(String(v).trim())));
  if (hi < 0) hi = 0;
  const headers = rows[hi].map((v, i) => String(v).trim() || ('عمود ' + (i + 1)));
  let nameCol = headers.findIndex(v => NAME_RX.test(v));
  if (nameCol < 0) { // أطول نصوص = عمود الاسم
    const score = headers.map((_, i) => rows.slice(hi + 1).reduce((a, r) => a + (/[^\d\s.,]/.test(r[i] || '') ? 1 : 0), 0));
    nameCol = score.indexOf(Math.max(...score));
  }
  let bcCol = headers.findIndex(v => BC_RX.test(v));
  if (bcCol < 0) { headers.push('الباركود'); bcCol = headers.length - 1; }
  const items = rows.slice(hi + 1).map((r, i) => {
    const cells = headers.map((_, j) => String(r[j] ?? '').trim());
    return { id: uid(), n: i + 1, cells, status: '' };
  }).filter(it => it.cells[nameCol]);
  if (!items.length) throw new Error('ما لكيت مواد بالملف');
  const norm = v => v.replace(/\(.*?\)|دينار/g, '').replace(/\s+/g, ' ').trim();
  const find = rx => headers.findIndex(v => rx.test(norm(v)));
  const f = { cartons: find(/^(عدد )?(الكراتين|الكارتون|الكرتون|كراتين)$/), cartonCost: find(/^سعر (شراء )?(الكارتون|الكرتون)$/),
    pieces: find(/^(قطع مفردة|عدد القطع|عدد القطعة|القطع)$/), pieceCost: find(/^سعر (شراء )?(القطع|القطعة)$/),
    pack: find(/(قطع بالكرتون|قطع بالكارتون)/), sell: find(/^سعر (ال)?بيع/) };
  for (const it of items) {
    const p = f.pack >= 0 ? numv(it.cells[f.pack]) : null; it.pack = p && p > 1 ? Math.round(p) : null;
    it.sell = f.sell >= 0 ? numv(it.cells[f.sell]) : null;
  }
  return { id: uid(), name: fileName.replace(/\.(xlsx|csv|txt)$/i, ''), created: Date.now(), headers, nameCol, bcCol, f, items };
}
$('#fileIn').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try {
    const buf = await f.arrayBuffer();
    const isXlsx = /\.xlsx$/i.test(f.name) || (new Uint8Array(buf)[0] === 0x50 && new Uint8Array(buf)[1] === 0x4b);
    const rows = isXlsx ? await readXlsx(buf) : readCsv(new TextDecoder().decode(buf));
    const list = toList(rows, f.name);
    DB.lists.unshift(list); save();
    S.view = 'list'; S.listId = list.id; S.filter = 'all'; render();
    toast('انضافت ' + list.items.length + ' مادة', 'ok');
  } catch (err) { toast('ما انقرى الملف: ' + (err.message || err), 'err'); }
});

/* ================= الحالة والرسم ================= */
const S = { view: 'home', listId: null, filter: 'all', q: '', continuous: localStorage.getItem('scan_cont') !== '0' };
const curList = () => DB.lists.find(l => l.id === S.listId);
const bcOf = (l, it) => it.cells[l.bcCol] || '';
const nameOf = (l, it) => it.cells[l.nameCol] || '';
const fv = (l, it, k) => (l.f && l.f[k] >= 0 ? numv(it.cells[l.f[k]]) : null) || 0;
function info(l, it) {   // الكمية والسعر من أعمدة القائمة
  const cartons = fv(l, it, 'cartons'), cartonCost = fv(l, it, 'cartonCost'), pieces = fv(l, it, 'pieces'), pieceCost = fv(l, it, 'pieceCost');
  const boxy = cartons > 0 || (it.pack || 0) > 1;
  const buy = cartons > 0 ? cartonCost : pieceCost;
  return { cartons, cartonCost, pieces, pieceCost, boxy, buy };
}
let MARKUP = numv(localStorage.getItem('scan_markup')) ?? 15;
let SELLMODE = localStorage.getItem('scan_sellmode') || 'empty';   // empty = التسعير بعدين بالبرنامج، guess = مقترح بنسبة ربح
function meta(l, it) {
  return l.headers.map((hd, i) => ({ hd, v: it.cells[i] })).filter((x, i) => i !== l.nameCol && i !== l.bcCol && x.v && !/^(#|ت|التسلسل|م)$/.test(x.hd) && x.v !== '0')
    .slice(0, 4).map(x => x.hd + ': ' + x.v).join(' · ');
}
function render() {
  const app = $('#app');
  if (S.view === 'list' && !curList()) S.view = 'home';
  app.replaceChildren(S.view === 'list' ? listView() : homeView());
}
const logo = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round"><path d="M4 6v12M7 6v12M10 6v12M14 6v12M17 6v12M20 6v12"/></svg>';
function homeView() {
  const head = h('header', {}, h('div', { class: 'wrap hrow' }, h('div', { class: 'logo', html: logo }),
    h('div', { class: 'grow' }, h('h1', {}, 'ماسح القوائم'), h('div', { class: 'sub' }, 'استورد القائمة ← امسح باركود كل مادة ← صدّرها'))));
  const body = h('div', { class: 'wrap' });
  if (!DB.lists.length) body.append(h('div', { class: 'card empty' }, 'ماكو قوائم بعد.', h('br'), 'اضغط «استيراد قائمة» واختار ملف Excel أو CSV.'));
  body.append(h('div', { class: 'lists' }, DB.lists.map(l => {
    const t = l.items.length, bc = l.items.filter(it => bcOf(l, it)).length;
    return h('button', { class: 'card lcard', onclick: () => { S.view = 'list'; S.listId = l.id; S.filter = 'all'; S.q = ''; render(); scrollTo(0, 0); } },
      h('div', { class: 'hrow' }, h('b', { class: 'grow' }, l.name), h('span', { class: 'sub' }, t + ' مادة')),
      h('div', { class: 'bar' }, h('i', { style: { width: (bc / (t || 1) * 100) + '%' } })),
      h('div', { class: 'sub' }, bc + ' / ' + t + ' بيها باركود'));
  })));
  const dock = h('div', { class: 'dock' }, h('div', { class: 'wrap' }, h('button', { class: 'btn primary block', onclick: () => $('#fileIn').click() }, '📄 استيراد قائمة (Excel)')));
  return h('div', {}, head, body, dock);
}
function filtered(l) {
  const q = digits(S.q).trim();
  return l.items.filter(it => {
    const b = bcOf(l, it);
    if (S.filter === 'nobc' && (b || it.status === 'miss')) return false;
    if (S.filter === 'done' && !b) return false;
    if (S.filter === 'miss' && it.status !== 'miss') return false;
    if (S.filter === 'nosell' && (it.sell || it.status === 'miss')) return false;
    return !q || (nameOf(l, it) + ' ' + b).includes(q);
  });
}
function listView() {
  const l = curList();
  const t = l.items.length, bc = l.items.filter(it => bcOf(l, it)).length, miss = l.items.filter(it => it.status === 'miss').length;
  const nobc = l.items.filter(it => !bcOf(l, it) && it.status !== 'miss').length;
  const chip = (k, txt) => h('button', { class: 'chip' + (S.filter === k ? ' on' : ''), onclick: () => { S.filter = k; render(); } }, txt);
  const search = h('input', { class: 'search', type: 'search', placeholder: 'ابحث بالاسم…', value: S.q });
  search.addEventListener('input', () => { S.q = search.value; drawItems(); });
  const itemsBox = h('div', { class: 'items' });
  const drawItems = () => {
    const its = filtered(l);
    itemsBox.replaceChildren(...(its.length ? its.map(it => itemCard(l, it)) : [h('div', { class: 'card empty' }, 'ماكو مواد')]));
  };
  drawItems();
  const head = h('header', {}, h('div', { class: 'wrap' },
    h('div', { class: 'hrow' }, h('button', { class: 'icon-btn', onclick: () => { S.view = 'home'; render(); } }, '→'),
      h('div', { class: 'grow' }, h('h1', {}, l.name), h('div', { class: 'sub' }, bc + '/' + t + ' باركود' + (miss ? ' · ' + miss + ' ما وصل' : ''))),
      h('button', { class: 'icon-btn', onclick: () => listMenu(l) }, '⋯')),
    h('div', { class: 'bar' }, h('i', { style: { width: (bc / (t || 1) * 100) + '%' } })),
    h('div', { class: 'chips' }, chip('all', 'الكل (' + t + ')'), chip('nobc', 'بدون باركود (' + nobc + ')'), chip('done', 'تم (' + bc + ')'), chip('nosell', 'بدون سعر بيع (' + l.items.filter(x => x.status !== 'miss' && !x.sell).length + ')'), chip('miss', 'ما وصل (' + miss + ')')),
    search));
  const dock = h('div', { class: 'dock' }, h('div', { class: 'wrap' },
    h('button', { class: 'btn primary', style: { flex: 2 }, disabled: !nobc, onclick: () => { const it = l.items.find(x => !bcOf(l, x) && x.status !== 'miss'); if (it) openScanner(l, it); } }, '📷 امسح التالي'),
    bc ? h('button', { class: 'btn', style: { flex: 1 }, onclick: () => showBarcode(l, l.items.find(x => bcOf(l, x) && x.status !== 'miss') || l.items.find(x => bcOf(l, x))) }, '⛶ القارئ') : null,
    h('button', { class: 'btn', style: { flex: 1 }, onclick: () => exportSheet(l) }, '⇪ تصدير')));
  return h('div', {}, head, h('div', { class: 'wrap' }, itemsBox), dock);
}
function itemCard(l, it) {
  const b = bcOf(l, it);
  const inf = info(l, it);
  const field = (label, sub, val, ph, onSave, warn, cls) => {
    const inp = h('input', { class: 'fin' + (cls ? ' ' + cls : ''), type: 'text', inputmode: 'numeric', value: val ?? '', placeholder: ph || '—' });
    inp.addEventListener('change', () => { onSave(numv(inp.value)); save(); render(); });
    inp.addEventListener('keydown', e => { if (e.key === 'Enter') inp.blur(); });
    return h('label', { class: 'frow' + (warn ? ' need' : '') }, h('span', {}, label, sub ? h('small', {}, sub) : null), inp);
  };
  const low = it.sell && inf.buy && it.sell <= inf.buy;
  const fields = it.status === 'miss' ? null : h('div', { class: 'fields' },
    field('عدد القطع بالكرتون', null, it.pack, '—', v => { it.pack = v && v > 1 ? Math.round(v) : null; }, inf.cartons > 0 && !(it.pack > 1)),
    field(inf.boxy ? 'سعر بيع الكرتون' : 'سعر بيع القطعة',
      (inf.buy ? 'الشراء ' + money(inf.buy) : '') + (low ? ' — ⚠ أقل من الشراء' : (!it.sell && inf.buy ? (SELLMODE === 'guess' ? ' — فارغ = مقترح +' + MARKUP + '%' : ' — فارغ = يتسعّر بعدين') : '')),
      it.sell, inf.buy && SELLMODE === 'guess' ? money(ceil250(inf.buy * (1 + MARKUP / 100))) : '—', v => { it.sell = v && v > 0 ? Math.round(v) : null; }, low, 'wide'));
  const g = b ? barcodeSvg(b, 56) : null;
  return h('div', { class: 'card item' + (b ? ' ok' : '') + (it.status === 'miss' ? ' miss' : '') },
    h('div', { class: 'itop' }, h('div', { class: 'idx' }, it.n), h('div', { class: 'grow' }, h('div', { class: 'iname' }, nameOf(l, it)), h('div', { class: 'imeta' }, meta(l, it)))),
    fields,
    b && it.status !== 'miss' ? h('button', { class: 'bcbox', onclick: () => showBarcode(l, it) }, g ? h('div', { html: g.svg }) : h('span', { class: 'code' }, b)) : null,
    h('div', { class: 'acts' },
      it.status === 'miss' ? h('div', { class: 'grow sub', style: { alignSelf: 'center' } }, '✗ ما وصل')
        : b ? h('div', { class: 'grow bcacts' }, h('button', { class: 'btn small', onclick: () => showBarcode(l, it) }, '⛶ للقارئ'),
            h('button', { class: 'btn small', onclick: () => openScanner(l, it) }, '📷 إعادة'), h('button', { class: 'btn small', onclick: () => editCode(l, it) }, '✎ تعديل'))
          : h('button', { class: 'scanbox', onclick: () => openScanner(l, it) }, '📷 امسح الباركود'),
      h('button', { class: 'tg missb' + (it.status === 'miss' ? ' on' : ''), title: 'ما وصل', onclick: () => { it.status = it.status === 'miss' ? '' : 'miss'; save(); render(); } }, '✗')));
}
function setCode(l, it, code) {
  code = digits(code).replace(/\s/g, '');
  const dup = code && l.items.find(x => x !== it && bcOf(l, x) === code);
  it.cells[l.bcCol] = code; if (code) it.status = ''; save();
  return dup;
}
function editCode(l, it) {
  const inp = h('input', { type: 'text', inputmode: 'numeric', value: bcOf(l, it), style: { direction: 'ltr' } });
  const s = sheet(nameOf(l, it), h('label', { class: 'fld' }, h('span', {}, 'الباركود'), inp), [
    h('button', { class: 'btn primary', onclick: () => { const d = setCode(l, it, inp.value); s.close(); render(); if (d) toast('⚠ نفس الباركود عند «' + nameOf(l, d) + '»', 'err'); } }, 'حفظ'),
    h('button', { class: 'btn', onclick: () => { s.close(); openScanner(l, it); } }, '📷 امسح')]);
}
function listMenu(l) {
  const nm = h('input', { type: 'text', value: l.name });
  const s = sheet('القائمة', h('label', { class: 'fld' }, h('span', {}, 'الاسم'), nm), [
    h('button', { class: 'btn primary', onclick: () => { l.name = nm.value.trim() || l.name; save(); s.close(); render(); } }, 'حفظ'),
    h('button', { class: 'btn', style: { color: 'var(--bad)' }, onclick: () => {
      s.close();
      const c = sheet('مسح «' + l.name + '»؟', h('div', { class: 'note bad' }, 'تنمسح القائمة وكل الباركودات اللي بيها من التلفون. صدّرها أول إذا تحتاجها.'), [
        h('button', { class: 'btn', style: { background: 'var(--bad)', color: '#fff' }, onclick: () => { DB.lists = DB.lists.filter(x => x !== l); save(); c.close(); S.view = 'home'; render(); } }, 'مسح'),
        h('button', { class: 'btn', onclick: () => c.close() }, 'لا')]);
    } }, 'مسح القائمة')]);
}

/* ================= رسم باركود حقيقي (EAN-13 / EAN-8 / Code128) ================= */
const EAN_L = ['0001101','0011001','0010011','0111101','0100011','0110001','0101111','0111011','0110111','0001011'];
const EAN_G = ['0100111','0110011','0011011','0100001','0011101','0111001','0000101','0010001','0001001','0010111'];
const EAN_R = EAN_L.map(p => p.replace(/./g, c => c === '0' ? '1' : '0'));
const EAN_PAR = ['LLLLLL','LLGLGG','LLGGLG','LLGGGL','LGLLGG','LGGLLG','LGGGLL','LGLGLG','LGLGGL','LGGLGL'];
function eanCheck(body) { let s = 0; const n = body.length; for (let i = 0; i < n; i++) s += (+body[n - 1 - i]) * (i % 2 === 0 ? 3 : 1); return (10 - s % 10) % 10; }
function eanValid(c) { return /^\d+$/.test(c) && (c.length === 13 || c.length === 8 || c.length === 12) && eanCheck(c.slice(0, -1)) === +c.slice(-1); }
const C128 = ('212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 221312 231212 112232 122132 122231 113222 123122 123221 223211 221132 ' +
  '221231 213212 223112 312131 311222 321122 321221 312212 322112 322211 212123 212321 232121 111323 131123 131321 112313 132113 132311 211313 ' +
  '231113 231311 112133 112331 132131 113123 113321 133121 313121 211331 231131 213113 213311 213131 311123 311321 331121 312113 312311 332111 ' +
  '314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 112412 122114 122411 142112 142211 241211 221114 413111 241112 134111 ' +
  '111242 121142 121241 114212 124112 124211 411212 421112 421211 212141 214121 412121 111143 111341 131141 114113 114311 411113 411311 113141 ' +
  '114131 311141 411131 211412 211214 211232 2331112').split(' ');
function c128Bits(text) {
  const vals = [];
  if (/^\d+$/.test(text) && text.length % 2 === 0 && text.length >= 4) { vals.push(105); for (let i = 0; i < text.length; i += 2) vals.push(+text.slice(i, i + 2)); }
  else { vals.push(104); for (const ch of text) { const c = ch.charCodeAt(0); if (c < 32 || c > 126) return null; vals.push(c - 32); } }
  let sum = vals[0]; for (let i = 1; i < vals.length; i++) sum += vals[i] * i;
  vals.push(sum % 103, 106);
  let bits = '';
  for (const v of vals) { const w = C128[v]; for (let i = 0; i < w.length; i++) bits += (i % 2 === 0 ? '1' : '0').repeat(+w[i]); }
  return bits;
}
function barcodeBits(code) {
  code = String(code || '').trim();
  if (/^\d{12}$/.test(code) && eanValid(code)) code = '0' + code;      // UPC-A = EAN-13 يبدأ بصفر
  if (/^\d{13}$/.test(code) && eanValid(code)) {
    const par = EAN_PAR[+code[0]]; let b = '101';
    for (let i = 1; i <= 6; i++) b += (par[i - 1] === 'L' ? EAN_L : EAN_G)[+code[i]];
    b += '01010'; for (let i = 7; i <= 12; i++) b += EAN_R[+code[i]];
    return { bits: b + '101', type: 'EAN-13' };
  }
  if (/^\d{8}$/.test(code) && eanValid(code)) {
    let b = '101'; for (let i = 0; i < 4; i++) b += EAN_L[+code[i]]; b += '01010'; for (let i = 4; i < 8; i++) b += EAN_R[+code[i]];
    return { bits: b + '101', type: 'EAN-8' };
  }
  const b = c128Bits(code);
  return b ? { bits: b, type: 'Code128' } : null;
}
function barcodeSvg(code, hgt) {
  const r = barcodeBits(code); if (!r) return null;
  const q = 10, W = r.bits.length + q * 2, H = hgt || 60;
  let rects = '', i = 0;
  while (i < r.bits.length) { if (r.bits[i] === '1') { let j = i; while (r.bits[j] === '1') j++; rects += `<rect x="${i + q}" y="0" width="${j - i}" height="${H}"/>`; i = j; } else i++; }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H + 14}" preserveAspectRatio="none" shape-rendering="crispEdges" style="display:block;width:100%;height:auto;background:#fff">` +
    `<rect width="${W}" height="${H + 14}" fill="#fff"/><g fill="#000">${rects}</g>` +
    `<text x="${W / 2}" y="${H + 11}" text-anchor="middle" font-family="ui-monospace,Menlo,Consolas,monospace" font-size="10" fill="#000" letter-spacing="1">${String(code).replace(/[<>&]/g, '')}</text></svg>`;
  return { svg, type: r.type };
}
let wakeLock = null;
async function keepAwake(on) {
  try { if (on && 'wakeLock' in navigator) wakeLock = await navigator.wakeLock.request('screen'); else if (!on && wakeLock) { await wakeLock.release(); wakeLock = null; } } catch (e) { /* */ }
}
function showBarcode(l, it) {
  const lst = l.items.filter(x => bcOf(l, x) && x.status !== 'miss');
  let k = Math.max(0, lst.indexOf(it));
  const ov = h('div', { class: 'bcview' });
  const close = () => { keepAwake(false); ov.remove(); };
  const draw = () => {
    const x = lst[k]; if (!x) { close(); return; }
    const g = barcodeSvg(bcOf(l, x), 70);
    ov.replaceChildren(
      h('div', { class: 'bcv-top' }, h('div', { class: 'grow' }, h('div', { class: 'bcv-name' }, x.n + ' · ' + nameOf(l, x)),
        h('div', { class: 'bcv-sub' }, (k + 1) + ' من ' + lst.length + (x.pack > 1 ? ' · ' + x.pack + ' قطعة بالكرتون' : ''))),
        h('button', { class: 'icon-btn', onclick: close }, '✕')),
      h('div', { class: 'bcv-code' }, g ? h('div', { html: g.svg }) : h('div', { class: 'note bad' }, 'الرقم بي حروف ما تنرسم')),
      h('div', { class: 'bcv-hint' }, 'قرّب قارئ الباركود من الشاشة (سطوع الشاشة على الآخر)'),
      h('div', { class: 'bcv-bot' },
        h('button', { class: 'btn', disabled: k <= 0, onclick: () => { k--; draw(); } }, '→ السابقة'),
        h('button', { class: 'btn primary', disabled: k >= lst.length - 1, onclick: () => { k++; draw(); } }, 'التالية ←')));
  };
  draw(); document.body.append(ov); keepAwake(true);
}

/* ================= الماسح ================= */
let scanner = null, scanBusy = false;
const FORMATS = () => (window.Html5QrcodeSupportedFormats ? [Html5QrcodeSupportedFormats.EAN_13, Html5QrcodeSupportedFormats.EAN_8, Html5QrcodeSupportedFormats.UPC_A,
  Html5QrcodeSupportedFormats.UPC_E, Html5QrcodeSupportedFormats.CODE_128, Html5QrcodeSupportedFormats.CODE_39, Html5QrcodeSupportedFormats.ITF,
  Html5QrcodeSupportedFormats.QR_CODE] : undefined);
async function openScanner(l, it) {
  if (typeof Html5Qrcode === 'undefined') { toast('مكتبة الماسح ما تحملت', 'err'); return; }
  let cur = it;
  const name = h('div', { class: 'sc-name' }), sub = h('div', { class: 'sc-sub' });
  const setCur = x => { cur = x; name.textContent = nameOf(l, x); sub.textContent = '#' + x.n + (meta(l, x) ? ' · ' + meta(l, x) : ''); };
  setCur(it);
  const cont = h('input', { type: 'checkbox' }); cont.checked = S.continuous;
  cont.addEventListener('change', () => { S.continuous = cont.checked; localStorage.setItem('scan_cont', cont.checked ? '1' : '0'); });
  const reader = h('div', { id: 'reader' });
  const wrap = h('div', { class: 'scanner' },
    h('div', { class: 'sc-top' }, name, sub),
    reader,
    h('div', { class: 'sc-bot' },
      h('label', { class: 'sw' }, cont, 'بعد كل مسح انتقل للمادة اللي بعدها وحده'),
      h('button', { class: 'btn', onclick: () => { const nx = nextNo(l, cur); if (nx) setCur(nx); else close(); } }, 'تخطي ⟵'),
      h('button', { class: 'btn', onclick: () => { close(); editCode(l, cur); } }, '⌨ اكتب'),
      h('button', { class: 'btn primary', onclick: () => close() }, 'خلص')));
  document.body.append(wrap);
  const close = async () => {
    try { if (scanner && scanner.isScanning) await scanner.stop(); } catch (e) { /* */ }
    try { scanner && scanner.clear(); } catch (e) { /* */ }
    scanner = null; wrap.remove(); render();
  };
  let last = '', lastAt = 0;
  const onCode = text => {
    const now = Date.now();
    if (scanBusy || (text === last && now - lastAt < 2500)) return;
    last = text; lastAt = now; scanBusy = true;
    const code = digits(text).replace(/\s/g, '');
    const dup = l.items.find(x => x !== cur && bcOf(l, x) === code);
    if (dup) {   // نفس الباركود مسجل لمادة ثانية: ما نحفظه ونبقى على نفس المادة
      beep(false);
      const fl = h('div', { class: 'flash', style: { background: 'rgba(180,54,26,.6)' } }, '⚠ هذا باركود «' + nameOf(l, dup) + '»'); reader.append(fl);
      setTimeout(() => { fl.remove(); scanBusy = false; }, 1800);
      return;
    }
    setCode(l, cur, code);
    beep(true);
    const fl = h('div', { class: 'flash' }, '✓ ' + code); reader.append(fl);
    setTimeout(() => {
      fl.remove(); scanBusy = false;
      if (!S.continuous) { close(); return; }
      const nx = nextNo(l, cur);
      if (nx) setCur(nx); else { toast('خلصت كل المواد ✓', 'ok'); close(); }
    }, 700);
  };
  try {
    scanner = new Html5Qrcode('reader', { formatsToSupport: FORMATS(), verbose: false, experimentalFeatures: { useBarCodeDetectorIfSupported: true } });
    await scanner.start({ facingMode: 'environment' },
      { fps: 12, qrbox: (w, hh) => ({ width: Math.floor(w * 0.85), height: Math.floor(Math.min(hh * 0.45, w * 0.5)) }), aspectRatio: undefined, disableFlip: false },
      onCode, () => {});
  } catch (e) {
    reader.append(h('div', { class: 'note bad', style: { margin: '16px' } }, 'ما انفتحت الكاميرا. اسمح للمتصفح باستعمال الكاميرا من الإعدادات، أو صوّر الباركود:',
      h('div', { class: 'row' }, h('button', { class: 'btn', onclick: () => photoScan(l, cur, onCode) }, '📸 صوّر صورة'))));
  }
}
function nextNo(l, it) {
  const i = l.items.indexOf(it);
  return l.items.slice(i + 1).find(x => !bcOf(l, x) && x.status !== 'miss') || l.items.slice(0, i).find(x => !bcOf(l, x) && x.status !== 'miss');
}
function photoScan(l, it, onCode) {
  const inp = $('#photoIn'); inp.value = '';
  inp.onchange = async () => {
    const f = inp.files[0]; if (!f) return;
    try { const tmp = new Html5Qrcode('reader', { formatsToSupport: FORMATS(), verbose: false }); const txt = await tmp.scanFile(f, false); onCode(txt); }
    catch (e) { beep(false); toast('ما انقرى من الصورة — قرّب وخلي الباركود واضح', 'err'); }
  };
  inp.click();
}

/* ================= التصدير ================= */
function csvCell(v) { v = String(v ?? ''); return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
const POS_HEAD = ['الباركود', 'اسم المادة', 'القسم', 'الوحدة', 'قطع بالكرتون', 'سعر شراء الكرتون', 'سعر بيع الكرتون', 'الكمية (كراتين)', 'حد التنبيه', 'الانتهاء'];
function posRows(l) {   // نفس ترتيب أعمدة «إدخال منتجات (جدول)» بوضع «الأسعار والكمية بالكرتون»
  const m = 1 + MARKUP / 100;
  return l.items.filter(it => it.status !== 'miss').map(it => {
    const x = info(l, it);
    const pack = it.pack > 1 ? it.pack : '';
    let cost, qty;
    if (x.cartons > 0) { cost = x.cartonCost; qty = x.cartons + (it.pack > 1 && x.pieces ? x.pieces / it.pack : 0); }
    else { cost = x.pieceCost; qty = x.pieces; }
    const sell = it.sell || (SELLMODE === 'guess' && x.buy ? ceil250(x.buy * m) : '');
    return { it, guess: !it.sell, cells: [bcOf(l, it), nameOf(l, it), '', '', pack, cost || '', sell, Math.round((qty || 0) * 1000) / 1000 || '', '', ''] };
  });
}
function exportSheet(l) {
  const withStatus = l.headers.concat(['الحالة']);
  const rows = l.items.map(it => it.cells.concat([it.status === 'miss' ? 'ما وصل' : bcOf(l, it) ? 'تم' : '']));
  const csv = '﻿' + [withStatus, ...rows].map(r => r.map(csvCell).join(',')).join('\r\n');
  const fname = l.name + '-باركود.csv';
  const file = new File([csv], fname, { type: 'text/csv' });
  const tsvAll = [l.headers, ...l.items.filter(it => it.status !== 'miss').map(it => it.cells)].map(r => r.map(v => String(v).replace(/[\t\n]/g, ' ')).join('\t')).join('\n');
  const tsvRows = l.items.filter(it => it.status !== 'miss').map(it => it.cells.map(v => String(v).replace(/[\t\n]/g, ' ')).join('\t')).join('\n');
  const done = l.items.filter(it => bcOf(l, it)).length;
  const posFile = () => {
    const csv2 = '\ufeff' + [POS_HEAD, ...posRows(l).map(r => r.cells)].map(r => r.map(csvCell).join(',')).join('\r\n');
    return new File([csv2], l.name + '-للبرنامج.csv', { type: 'text/csv' });
  };
  const mk = h('input', { type: 'text', inputmode: 'numeric', value: MARKUP, style: { maxWidth: '90px' } });
  const summary = h('div');
  const drawSum = () => {
    const pr = posRows(l), guess = pr.filter(r => r.guess).length, nopack = pr.filter(r => info(l, r.it).cartons > 0 && !(r.it.pack > 1)).length;
    summary.replaceChildren(h('div', { class: 'note ' + (guess || nopack || done < l.items.length ? 'warn' : 'ok') },
      pr.length + ' مادة للبرنامج. ' + (done < l.items.length ? (l.items.length - done) + ' بدون باركود. ' : '') + (nopack ? nopack + ' ناقصها «قطع بالكرتون». ' : '') +
      (guess ? guess + (SELLMODE === 'guess' ? ' بدون سعر بيع (ينحط مقترح +' + MARKUP + '%).' : ' بدون سعر بيع — تنحفظ بالبرنامج فارغة وتتسعّر بعدين («💲 مواد بدون سعر»).') : 'كل الأسعار مكتوبة ✓')));
  };
  const mkWrap = h('label', { class: 'fld' + (SELLMODE === 'guess' ? '' : ' hidden') }, h('span', {}, 'نسبة الربح %'), mk);
  const sm = h('select', {}, h('option', { value: 'empty' }, 'اتركه فارغ — نسعّر بعدين بالبرنامج'), h('option', { value: 'guess' }, 'حط سعر مقترح بنسبة ربح'));
  sm.value = SELLMODE;
  sm.addEventListener('change', () => { SELLMODE = sm.value; localStorage.setItem('scan_sellmode', SELLMODE); mkWrap.classList.toggle('hidden', SELLMODE !== 'guess'); drawSum(); });
  mk.addEventListener('input', () => { MARKUP = numv(mk.value) || 0; localStorage.setItem('scan_markup', MARKUP); drawSum(); });
  drawSum();
  const pf = posFile();
  const s = sheet('تصدير «' + l.name + '»', h('div', {},
    summary,
    h('label', { class: 'fld' }, h('span', {}, 'المواد اللي ما كتبت سعر بيعها'), sm), mkWrap,
    h('div', { class: 'note info' }, 'للبرنامج: «📤 أرسل ملف للبرنامج» ← دزه بالواتساب لحاسبة المحل ← نزّله ← بالبرنامج: المخزون والجرد ← إدخال منتجات (جدول) ← «📂 فتح ملف» ← حفظ الكل.'),
    h('div', { class: 'row' },
      navigator.canShare && navigator.canShare({ files: [pf] })
        ? h('button', { class: 'btn primary', onclick: async () => { try { await navigator.share({ files: [posFile()], title: pf.name }); } catch (e) { /* */ } } }, '📤 أرسل ملف للبرنامج')
        : h('button', { class: 'btn primary', onclick: () => { const a = h('a', { href: URL.createObjectURL(posFile()), download: pf.name }); document.body.append(a); a.click(); a.remove(); } }, '⬇ ملف للبرنامج'),
      h('button', { class: 'btn', onclick: async () => { const t = posRows(l).map(r => r.cells.map(v => String(v).replace(/[\t\n]/g, ' ')).join('\t')).join('\n');
        try { await navigator.clipboard.writeText(t); toast('انسخت — الصقها بأول خانة باركود بالجدول', 'ok'); } catch (e) { toast('النسخ ما اشتغل', 'err'); } } }, '📋 نسخ للبرنامج')),
    h('div', { class: 'sub', style: { marginTop: '14px' } }, 'أو بنفس أعمدة القائمة الأصلية:')), [
    navigator.canShare && navigator.canShare({ files: [file] })
      ? h('button', { class: 'btn primary', onclick: async () => { try { await navigator.share({ files: [file], title: fname }); } catch (e) { /* */ } } }, '📤 مشاركة (واتساب…)') : null,
    h('button', { class: 'btn' + (navigator.canShare ? '' : ' primary'), onclick: () => { const a = h('a', { href: URL.createObjectURL(file), download: fname }); document.body.append(a); a.click(); a.remove(); } }, '⬇ تنزيل CSV'),
    h('button', { class: 'btn', onclick: async () => { try { await navigator.clipboard.writeText(tsvRows); toast('انسخت الأسطر (بدون العناوين)', 'ok'); } catch (e) { toast('النسخ ما اشتغل', 'err'); } } }, '📋 نسخ الأسطر'),
    h('button', { class: 'btn', onclick: async () => { try { await navigator.clipboard.writeText(tsvAll); toast('انسخ الجدول مع العناوين', 'ok'); } catch (e) { toast('النسخ ما اشتغل', 'err'); } } }, '📋 نسخ مع العناوين')]);
}

/* ================= تشغيل ================= */
render();
if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
