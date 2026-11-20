// sync.js — общая база Supabase: вход, отправка изменений, приём чужих в реальном времени.
// Подключается ПОСЛЕ script.js (в index.html) и ПОСЛЕ travel.js (в enclosement.html).
// Локальный кэш (localStorage 'turan3') остаётся: без интернета сайт работает, данные уйдут при появлении связи.
(function () {
  const SB_URL = 'https://icvivxzjqegkmyblohxm.supabase.co';
  const SB_KEY = 'sb_publishable_TNIXU5f0j7TTGeeYAjbopQ_55Rt2mq_';
  const LS = 'turan3', LS_SYNC = 'turan3_sync', LS_DIRTY = 'turan3_dirty';
  const TABLES = ['rooms', 'people', 'assign', 'book', 'enc', 'kv'];
  const KV = ['depts', 'companies', 'hist'];

  // уникальный номер записи (у разных компьютеров не совпадёт)
  let lastId = 0;
  window.newId = () => {
    let n = Date.now() * 1000 + Math.floor(Math.random() * 1000);
    if (n <= lastId) n = lastId + 1;                 // внутри одного компьютера номера никогда не повторяются
    return (lastId = n);
  };

  let sb = null, SNAP = {}, ready = false, pushing = false, again = false, pushTimer = null;
  let subscribed = false, wasSubscribed = false, uiTimer = null, pendingUI = false;
  const lsGet = (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, v); } catch (e) {} };
  const persist = () => lsSet(LS, JSON.stringify(DB));
  // JSON с отсортированными ключами: Postgres хранит jsonb в своём порядке, сравнивать надо независимо от порядка
  const stab = (v) => JSON.stringify(v, (k, x) => (x && typeof x === 'object' && !Array.isArray(x) ? Object.keys(x).sort().reduce((o, kk) => ((o[kk] = x[kk]), o), {}) : x));

  // ───────── индикатор ─────────
  function badge(state, msg) {
    let b = document.getElementById('sb-badge');
    if (!b) {
      b = document.createElement('div'); b.id = 'sb-badge';
      b.style.cssText = 'position:fixed;left:10px;bottom:10px;z-index:60;font-size:11px;background:var(--sf,#161b22);border:1px solid var(--bd,#2a3040);border-radius:8px;padding:4px 10px;color:var(--mu,#7d8590)';
      document.body.appendChild(b);
    }
    const t = { ok: '☁ синхронизировано', pending: '☁ отправка…', load: '☁ загрузка…', off: '☁ ' + (msg || 'не подключено'), err: '⚠ нет связи — данные сохранены на этом компьютере' }[state];
    b.style.color = state === 'err' ? 'var(--rd,#f75f5f)' : state === 'ok' ? 'var(--gr,#3ecf8e)' : 'var(--mu,#7d8590)';
    b.innerHTML = t + ' · <a href="#" style="color:inherit" onclick="sbLogout();return false">выйти</a>';
    if (msg && state === 'err') console.warn('sync:', msg);
  }

  // ───────── строки базы ↔ DB ─────────
  function rowsOf(t) {
    const m = new Map();
    if (t === 'rooms') Object.values(DB.rooms || {}).forEach((r) => m.set(String(r.key), stab(r)));
    else if (t === 'kv') KV.forEach((k) => { if (DB[k] !== undefined) m.set(k, stab({ v: DB[k] })); });
    else (DB[t] || []).forEach((x) => m.set(String(x.id), stab(x)));
    return m;
  }
  function dbPut(t, id, data) {
    if (t === 'rooms') (DB.rooms = DB.rooms || {})[id] = data;
    else if (t === 'kv') DB[id] = data.v;
    else {
      const a = (DB[t] = DB[t] || []), i = a.findIndex((x) => String(x.id) === id);
      if (i >= 0) a[i] = data; else a.push(data);
    }
  }
  function dbRemove(t, id) {
    if (t === 'rooms') { if (DB.rooms) delete DB.rooms[id]; }
    else if (t !== 'kv') DB[t] = (DB[t] || []).filter((x) => String(x.id) !== id);
  }

  // ───────── отправка изменений ─────────
  function sbPush() {
    lsSet(LS_DIRTY, '1');
    if (!ready) return;
    clearTimeout(pushTimer);
    pushTimer = setTimeout(doPush, 500);
    badge('pending');
  }
  async function pushTable(t) {
    const cur = rowsOf(t), prev = SNAP[t] || new Map(), up = [], del = [];
    cur.forEach((j, id) => { if (prev.get(id) !== j) up.push({ id, data: JSON.parse(j) }); });
    if (t !== 'kv') prev.forEach((j, id) => { if (!cur.has(id)) del.push(id); });
    if (!up.length && !del.length) return;
    SNAP[t] = cur;
    try {
      for (let i = 0; i < up.length; i += 400) { const { error } = await sb.from(t).upsert(up.slice(i, i + 400)); if (error) throw error; }
      for (let i = 0; i < del.length; i += 200) { const { error } = await sb.from(t).delete().in('id', del.slice(i, i + 200)); if (error) throw error; }
    } catch (e) { SNAP[t] = prev; throw e; }
  }
  async function doPush() {
    pushTimer = null;
    if (pushing) { again = true; return; }
    pushing = true;
    try {
      do { again = false; for (const t of TABLES) await pushTable(t); } while (again);
      lsSet(LS_DIRTY, '0'); badge('ok');
    } catch (e) { badge('err', e.message || String(e)); }
    finally { pushing = false; }
  }

  // ───────── приём чужих изменений ─────────
  function onRemote(t, p) {
    const id = p.eventType === 'DELETE' ? p.old && p.old.id : p.new && p.new.id;
    if (id == null) return;
    const S_ = SNAP[t] || (SNAP[t] = new Map());
    if (p.eventType === 'DELETE') { if (!S_.has(id)) return; S_.delete(id); dbRemove(t, id); }
    else { const j = stab(p.new.data); if (S_.get(id) === j) return; S_.set(id, j); dbPut(t, id, p.new.data); }
    persist(); notify();
  }
  // перерисовать экран, но не мешать, пока человек печатает в поле
  function notify() {
    clearTimeout(uiTimer);
    uiTimer = setTimeout(() => {
      const a = document.activeElement;
      if (a && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName) && a.type !== 'checkbox' && a.type !== 'file') { pendingUI = true; return; }
      pendingUI = false;
      if (window.onSbData) window.onSbData();
    }, 300);
  }
  document.addEventListener('focusout', () => { if (pendingUI) setTimeout(notify, 500); });

  // ───────── загрузка из облака ─────────
  async function fetchAll(t) {
    let all = [], from = 0;
    for (;;) {
      const { data, error } = await sb.from(t).select('id,data').order('id').range(from, from + 999);
      if (error) throw error;
      all = all.concat(data);
      if (data.length < 1000) return all;
      from += 1000;
    }
  }
  function applyCloud(raw, mode) {
    SNAP = {};
    TABLES.forEach((t) => {
      SNAP[t] = new Map(raw[t].map((r) => [r.id, stab(r.data)]));
      if (mode === 'replace') {
        if (t === 'rooms') DB.rooms = {};
        else if (t !== 'kv') DB[t] = [];
      }
      raw[t].forEach((r) => {
        if (mode === 'merge') {                                   // локальные данные главнее: добавляем только то, чего нет
          if (t === 'rooms' ? (DB.rooms || {})[r.id] : t === 'kv' ? DB[r.id] !== undefined : (DB[t] || []).some((x) => String(x.id) === r.id)) return;
        }
        dbPut(t, r.id, r.data);
      });
    });
  }
  function downloadLocalBackup() {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([lsGet(LS) || '{}'], { type: 'application/json' }));
    a.download = 'turan_local_backup_' + new Date().toISOString().slice(0, 10) + '.json';
    a.click();
  }
  async function boot(silent) {
    badge('load');
    const raw = {};
    try { for (const t of TABLES) raw[t] = await fetchAll(t); } catch (e) { badge('err', e.message || String(e)); return; }
    const cloudEmpty = TABLES.every((t) => !raw[t].length);
    const localHas = (DB.people && DB.people.length) || Object.keys(DB.rooms || {}).length;
    const synced = lsGet(LS_SYNC) === '1', dirty = lsGet(LS_DIRTY) === '1';
    let push = false;
    if (cloudEmpty) {
      if (localHas && (silent || confirm('В облаке пока пусто. Загрузить в облако данные из этого браузера?'))) { SNAP = {}; push = true; }
      else SNAP = {};
    } else if (localHas && !synced && !silent) {
      if (!confirm('В облаке уже есть данные. Они ЗАМЕНЯТ данные этого браузера (копия будет скачана). Продолжить?')) { badge('off', 'синхронизация выключена'); return; }
      downloadLocalBackup(); applyCloud(raw, 'replace');
    } else if (synced && dirty) { applyCloud(raw, 'merge'); push = true; }
    else applyCloud(raw, 'replace');
    ready = true; lsSet(LS_SYNC, '1'); persist(); notify();
    subscribe();
    if (push) sbPush(); else badge('ok');
  }
  function subscribe() {
    if (subscribed) return;
    subscribed = true;
    const ch = sb.channel('turan-sync');
    TABLES.forEach((t) => ch.on('postgres_changes', { event: '*', schema: 'public', table: t }, (p) => onRemote(t, p)));
    ch.subscribe((status) => {
      if (status === 'SUBSCRIBED') { if (wasSubscribed) boot(true); wasSubscribed = true; }   // после обрыва связи — догружаем пропущенное
      else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT' || status === 'CLOSED') badge('err', status);
    });
  }

  // ───────── вход ─────────
  function showLogin(msg) {
    let o = document.getElementById('sb-login');
    if (!o) {
      o = document.createElement('div'); o.id = 'sb-login';
      o.style.cssText = 'position:fixed;inset:0;z-index:9999;background:var(--bg,#0d1117);display:flex;align-items:center;justify-content:center';
      document.body.appendChild(o);
    }
    o.innerHTML = `<div class="card" style="width:320px;padding:22px"><div class="big" style="margin-bottom:12px">TURAN CAMP — вход</div>
      <input id="sb-e" type="email" placeholder="Почта" style="width:100%;margin-bottom:8px" autocomplete="username">
      <input id="sb-p" type="password" placeholder="Пароль" style="width:100%;margin-bottom:10px" autocomplete="current-password" onkeydown="if(event.key==='Enter')sbLogin()">
      <button class="b p" style="width:100%" onclick="sbLogin()">Войти</button>
      <div class="mu" style="margin-top:10px;color:var(--rd,#f75f5f)">${msg || ''}</div></div>`;
  }
  window.sbLogin = async function () {
    const email = document.getElementById('sb-e').value.trim(), password = document.getElementById('sb-p').value;
    const { error } = await sb.auth.signInWithPassword({ email, password });
    if (error) return showLogin('Неверная почта или пароль');
    const o = document.getElementById('sb-login'); if (o) o.remove();
    boot(false);
  };
  window.sbLogout = async function () { if (sb) await sb.auth.signOut(); location.reload(); };
  window.sbPush = sbPush;

  window.addEventListener('beforeunload', (e) => { if (pushTimer || pushing) { e.preventDefault(); e.returnValue = ''; } });
  setInterval(() => { if (ready && lsGet(LS_DIRTY) === '1' && !pushing && !pushTimer) doPush(); }, 15000);

  async function start() {
    if (!window.supabase) return badge('err', 'не загрузилась библиотека Supabase');
    sb = window.supabase.createClient(SB_URL, SB_KEY);
    const { data } = await sb.auth.getSession();
    if (data.session) boot(false); else showLogin();
    sb.auth.onAuthStateChange((ev) => { if (ev === 'SIGNED_OUT' && ready) location.reload(); });
  }
  // для тестов
  window.__sbTest = { stab, rowsOf, applyCloud, onRemote, pushTable, setClient: (c) => { sb = c; }, snap: () => SNAP, setReady: (v) => { ready = v; } };
  if (!window.__SB_NO_START) start();
})();
