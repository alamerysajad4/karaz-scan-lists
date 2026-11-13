/* مزامنة القوائم بين كل التلفونات عن طريق Supabase.
   كل تغيير ينحفظ على التلفون أول (يشتغل بدون نت)، وينرفع للسيرفر أول ما يكون أكو نت،
   والتغييرات من التلفونات الثانية تنزل كل ١٥ ثانية. الدخول برمز المحل (PIN). */
(() => {
  'use strict';
  const SB = { url: 'https://mvuslfmhmhdkugfrhgkk.supabase.co', key: 'sb_publishable_WzeB_WF3YYhF-UlivIdlQg_gqidk0a7' };
  const SKEY = 'scanapp_sync_v1';
  let SY;
  try { SY = JSON.parse(localStorage.getItem(SKEY)) || null; } catch (e) { SY = null; }
  SY = Object.assign({ pin: '', rev: 0, snap: {}, u: {}, dl: {}, di: {} }, SY || {});
  const keep = () => { try { localStorage.setItem(SKEY, JSON.stringify(SY)); } catch (e) { /* */ } };
  const metaOf = l => { const m = Object.assign({}, l); delete m.items; return m; };

  /* يقارن البيانات بآخر نسخة معروفة ويأشر اللي تغيّر */
  function diff() {
    const now = Date.now(), seen = new Set();
    for (const l of DB.lists) {
      const k = 'L:' + l.id, js = JSON.stringify(metaOf(l));
      seen.add(k);
      if (SY.snap[k] !== js) { SY.snap[k] = js; SY.u[k] = now; SY.dl[l.id] = 1; }
      for (const it of l.items) {
        if (!it.id) it.id = l.id + '-' + Math.random().toString(36).slice(2, 8);
        const ki = 'I:' + it.id, ji = JSON.stringify(it);
        seen.add(ki);
        if (SY.snap[ki] !== ji) { SY.snap[ki] = ji; SY.u[ki] = now; SY.di[it.id] = l.id; }
      }
    }
    for (const k of Object.keys(SY.snap)) {
      if (seen.has(k)) continue;
      if (k.startsWith('L:')) { const id = k.slice(2); SY.dl[id] = 'del'; SY.u[k] = now; }
      delete SY.snap[k];
    }
    keep();
  }

  async function rpc(fn, args) {
    const r = await fetch(SB.url + '/rest/v1/rpc/' + fn, {
      method: 'POST', headers: { apikey: SB.key, 'Content-Type': 'application/json' },
      body: JSON.stringify(args) });
    const t = await r.text();
    if (!r.ok) { const e = new Error(/bad_pin/.test(t) ? 'bad_pin' : 'http_' + r.status); e.body = t; throw e; }
    return t ? JSON.parse(t) : null;
  }

  async function push() {
    const lists = Object.entries(SY.dl).map(([id, v]) => {
      const k = 'L:' + id;
      if (v === 'del') return { id, name: '', meta: {}, deleted: true, u: SY.u[k] || Date.now() };
      const l = DB.lists.find(x => x.id === id);
      return l ? { id, name: l.name, meta: metaOf(l), deleted: false, u: SY.u[k] || Date.now() } : null;
    }).filter(Boolean);
    const items = Object.entries(SY.di).map(([id, lid]) => {
      const l = DB.lists.find(x => x.id === lid); const it = l && l.items.find(x => x.id === id);
      return it ? { id, list_id: lid, data: it, u: SY.u['I:' + id] || Date.now() } : null;
    }).filter(Boolean);
    if (!lists.length && !items.length) { SY.dl = {}; SY.di = {}; return; }
    for (let i = 0; i < Math.max(items.length, 1); i += 400) {
      await rpc('sync_push', { p_pin: SY.pin, p_lists: i === 0 ? lists : [], p_items: items.slice(i, i + 400) });
    }
    for (const l of lists) delete SY.dl[l.id];
    for (const it of items) delete SY.di[it.id];
    for (const id of Object.keys(SY.dl)) if (SY.dl[id] !== 'del' && !DB.lists.find(x => x.id === id)) delete SY.dl[id];
    for (const id of Object.keys(SY.di)) if (!DB.lists.some(l => l.items.some(x => x.id === id))) delete SY.di[id];
    keep();
  }

  async function pull() {
    const r = await rpc('sync_pull', { p_pin: SY.pin, p_since: SY.rev || 0 });
    let changed = false;
    for (const x of r.lists || []) {
      const k = 'L:' + x.id, i = DB.lists.findIndex(l => l.id === x.id);
      if (x.deleted) {
        if (i >= 0 && !SY.dl[x.id]) { DB.lists.splice(i, 1); changed = true; }
        delete SY.snap[k];
        for (const key of Object.keys(SY.snap)) if (key.startsWith('I:' + x.id)) delete SY.snap[key];
        continue;
      }
      if (SY.dl[x.id] && (SY.u[k] || 0) > x.u) continue;      // تعديلي أحدث وراح ينرفع
      const meta = Object.assign({}, x.meta, { id: x.id, name: x.name });
      if (i < 0) DB.lists.unshift(Object.assign(meta, { items: [] }));
      else { const L = DB.lists[i]; for (const key of Object.keys(L)) if (key !== 'items') delete L[key]; Object.assign(L, meta); }
      SY.snap[k] = JSON.stringify(metaOf(DB.lists.find(l => l.id === x.id))); SY.u[k] = x.u; delete SY.dl[x.id];
      changed = true;
    }
    const touched = new Set();
    for (const x of r.items || []) {
      const l = DB.lists.find(l => l.id === x.list_id); if (!l) continue;
      const k = 'I:' + x.id;
      if (SY.di[x.id] && (SY.u[k] || 0) > x.u) continue;
      const j = l.items.findIndex(it => it.id === x.id);
      if (j < 0) l.items.push(x.data);
      else { const it = l.items[j]; for (const key of Object.keys(it)) delete it[key]; Object.assign(it, x.data); }   // نفس الكائن (حتى الماسح المفتوح ما يضيع)
      SY.snap[k] = JSON.stringify(x.data); SY.u[k] = x.u; delete SY.di[x.id];
      touched.add(l); changed = true;
    }
    for (const l of touched) l.items.sort((a, b) => (a.n || 0) - (b.n || 0));
    if (r.rev) SY.rev = r.rev;
    keep();
    if (changed) {
      origSave();
      const scanning = document.querySelector('.scanner');
      const typing = document.activeElement && /INPUT|SELECT|TEXTAREA/.test(document.activeElement.tagName);
      if (!scanning && !typing && !document.querySelector('.ov')) render();
      else pendingRender = true;
    }
  }

  let busy = false, timer = null, pendingRender = false, state = 'idle';
  async function sync() {
    if (!SY.pin || busy) return;
    busy = true; setState('busy');
    try { await push(); await pull(); setState('ok'); }
    catch (e) {
      if (e.message === 'bad_pin') { SY.pin = ''; keep(); setState('off'); askPin('الرمز تغيّر أو غلط — اكتب رمز المحل من جديد'); }
      else setState('wait');
    } finally { busy = false; }
    if (pendingRender && !document.querySelector('.scanner') && !document.querySelector('.ov')) { pendingRender = false; render(); }
  }
  const soon = () => { clearTimeout(timer); timer = setTimeout(sync, 1200); };

  /* حالة المزامنة: نقطة صغيرة فوق */
  const badge = document.createElement('div');
  badge.className = 'syncb';
  document.body.append(badge);
  function setState(s) {
    state = s;
    const pend = Object.keys(SY.dl).length + Object.keys(SY.di).length;
    badge.textContent = s === 'ok' ? '☁ متزامن' : s === 'busy' ? '⟳ يحدّث…' : s === 'wait' ? '⏳ ماكو نت' + (pend ? ' — ' + pend + ' تغيير ينتظر' : '') : '🔒 بدون رمز';
    badge.dataset.s = s;
  }
  const st = document.createElement('style');
  st.textContent = '.syncb{position:fixed;top:calc(8px + env(safe-area-inset-top));left:12px;z-index:40;font-size:11.5px;font-weight:800;padding:4px 10px;border-radius:99px;background:var(--soft);color:var(--muted);box-shadow:var(--shadow)}' +
    '.syncb[data-s="ok"]{background:var(--ok-l);color:var(--ok)}.syncb[data-s="wait"]{background:var(--warn-l);color:var(--warn)}.syncb[data-s="off"]{background:var(--bad-l);color:var(--bad)}';
  document.head.append(st);

  /* حفظ = حفظ على التلفون + تأشير التغييرات + رفع */
  const origSave = save;
  save = function () { origSave(); diff(); setState(SY.pin ? (navigator.onLine ? state : 'wait') : 'off'); soon(); };

  function askPin(msg) {
    if (document.querySelector('.pinov')) return;
    const inp = h('input', { type: 'text', inputmode: 'numeric', placeholder: 'مثلاً 123456', class: 'ltr', style: { fontSize: '22px', textAlign: 'center', letterSpacing: '4px' } });
    const err = h('div');
    const ok = h('button', { class: 'btn primary block', onclick: async () => {
      const p = digits(inp.value).replace(/\D/g, '');
      if (!p) return;
      ok.disabled = true; err.replaceChildren();
      try {
        await rpc('sync_pull', { p_pin: p, p_since: 0 });
        SY.pin = p; SY.rev = 0; keep(); ov.remove(); toast('تم الربط ✓', 'ok');
        diff(); await sync();
      } catch (e) {
        err.replaceChildren(h('div', { class: 'note bad' }, e.message === 'bad_pin' ? 'الرمز غلط' : 'ماكو اتصال — تأكد من الإنترنت وجرّب'));
      } finally { ok.disabled = false; }
    } }, 'دخول');
    const ov = h('div', { class: 'ov pinov' }, h('div', { class: 'sheet' },
      h('h2', {}, '🔒 رمز المحل'),
      h('div', { class: 'note info' }, msg || 'اكتب رمز المحل حتى تطلع القوائم المشتركة، وكل شغلك يوصل للباقين.'),
      h('label', { class: 'fld' }, h('span', {}, 'رمز المحل'), inp), err, ok));
    inp.addEventListener('keydown', e => { if (e.key === 'Enter') ok.click(); });
    document.body.append(ov); setTimeout(() => inp.focus(), 80);
  }

  window.addEventListener('online', soon);
  document.addEventListener('visibilitychange', () => { if (!document.hidden) sync(); });
  setInterval(() => { if (!document.hidden) sync(); }, 15000);
  diff();
  if (SY.pin) { setState('busy'); sync(); } else { setState('off'); askPin(); }
  window.__sync = { sync, state: () => state, SY };
})();
