// Enclosement: прибытия / отправления. Данные общие с «Проживанием» (localStorage 'turan3').
const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const KEY = 'turan3';
let DB;
function load() {
  DB = { rooms: {}, people: [], assign: [], hist: [], nid: 1, book: [], enc: [] };
  try { Object.assign(DB, JSON.parse(localStorage.getItem(KEY) || '{}')); } catch (e) {}
  DB.enc = DB.enc || []; DB.book = DB.book || [];
}
const save = () => { try { localStorage.setItem(KEY, JSON.stringify(DB)); } catch (e) { alert('⚠ Не удалось сохранить данные (память браузера переполнена). Скачайте резервную копию в «Проживание → Данные».'); } if (window.sbPush) sbPush(); };
const mut = (fn) => { load(); fn(); save(); render(); }; // всегда берём свежие данные перед изменением
const iso = (d) => new Date(d - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10);
const today = () => iso(new Date());
const shiftDay = (s, n) => { const d = new Date(s + 'T00:00:00'); d.setDate(d.getDate() + n); return iso(d); };
const toast = (m) => { const t = $('#toast'); t.textContent = m; t.style.display = 'block'; setTimeout(() => (t.style.display = 'none'), 2500); };
const log = (a, d) => { DB.hist.unshift({ t: new Date().toLocaleString('ru-RU'), a, d }); DB.hist = DB.hist.slice(0, 100); };
const full = (x) => (x.name + ' ' + x.sur).trim();
const same = (a, b) => full(a).toLowerCase() === full(b).toLowerCase();
const key = (e) => String(e.block || '').trim().toUpperCase() + '|' + String(e.room || '').trim();
const S = { date: today(), q: '', tf: today(), tt: shiftDay(today(), 14) };
load();
window.addEventListener('storage', (e) => { if (e.key === KEY) { load(); render(); } });

// ───────── ОТРИСОВКА ─────────
function roomChip(e) {
  const k = key(e), r = DB.rooms[k];
  if (!r) return e.block ? `<span class="c n">${esc(e.block)}</span>` : '';
  const n = DB.people.filter((p) => p.key === k).length;
  return n >= 2 ? '<span class="c r">комната занята</span>' : n === 1 ? '<span class="c y">1 койка свободна</span>' : '<span class="c g">свободна</span>';
}
function row(e) {
  const isIn = e.type === 'in', br = e.block === e.room ? e.block : e.block + (e.room ? '-' + e.room : '');
  return `<tr class="${e.done ? 'dn' : ''}"><td class="mu">${e.n}</td><td>${esc(e.time) || '—'}</td>
  <td><button class="b s" style="padding:1px 5px; font-size:11px;" onclick="editEnc(${e.id})" title="Редактировать">✎</button><b>${esc(full(e))}</b><div class="mu">${esc(e.emp)} · ${esc(e.comp)}${e.out ? ' · выезд ' + e.out.split('-').reverse().slice(0, 2).join('.') : ''}</div></td>
  <td>${roomCell(e)}</td><td>${e.done ? `<span class="c g">✓ ${isIn ? 'заселён' : 'выехал'}</span>` : isIn ? roomChip(e) : ''}</td>
  <td style="white-space:nowrap">${e.done ? `<button class="b s" onclick="undo(${e.id})">↩</button>` : `<button class="b s p" onclick="${isIn ? 'arrive' : 'depart'}(${e.id})">${isIn ? 'Приехал' : 'Уехал'}</button>`}
  <button class="b s d" onclick="del(${e.id})">✕</button></td></tr>`;
}
function render() {
  const q = S.q.toLowerCase().trim();
  const ev = DB.enc.filter((e) => e.date === S.date && (!q || (full(e) + ' ' + e.emp + ' ' + e.comp + ' ' + e.block + ' ' + e.room).toLowerCase().includes(q))).sort((a, b) => a.n - b.n);
  const all = DB.enc.filter((e) => e.date === S.date);
  const I = all.filter((e) => e.type === 'in'), O = all.filter((e) => e.type === 'out');
  $('#stats').innerHTML = `Прибытие: <b>${I.filter((e) => e.done).length}</b>/${I.length} · Отправление: <b>${O.filter((e) => e.done).length}</b>/${O.length}`;
  const tbl = (t, title) => {
    const L = ev.filter((e) => e.type === t);
    return `<div><h3 style="margin-top:0">${title} · ${L.length}</h3><table><tr><th>№</th><th>Время</th><th>ФИО</th><th>Блок-комн.</th><th></th><th></th></tr>${L.map(row).join('') || '<tr><td colspan="6" class="mu">Пусто</td></tr>'}</table></div>`;
  };
  $('#main').innerHTML = `
  <div class="bar"><button class="b" onclick="setDate(-1)">‹</button><input type="date" value="${S.date}" onchange="S.date=this.value;render()"><button class="b" onclick="setDate(1)">›</button>
  <button class="b" onclick="S.date=today();render()">Сегодня</button>
  <span style="margin-left:auto" class="mu">Импорт Excel:</span><input type="file" accept=".xlsx" onchange="importXlsx(this.files[0]);this.value=''"><button class="b" onclick="editCompanies()">🏢 Компании</button><button class="b" onclick="report()">📊 Отчёт</button></div>
  <div class="bar"><b>✈️ Travel-календарь:</b><span class="mu">с</span><input type="date" value="${S.tf}" onchange="S.tf=this.value"><span class="mu">по</span><input type="date" value="${S.tt}" onchange="S.tt=this.value"><input type="file" accept=".xlsx" onchange="importTravel(this.files[0]);this.value=''"></div>
  <div class="card"><div class="row"><select id="ft"><option value="in">Прибытие</option><option value="out">Отправление</option></select>
  <input id="fm" placeholder="Время (18:40 / DAY SHIFT)" size="18"><input id="fe" placeholder="Таб. № (Enter — найти)" size="16" onkeydown="if(event.key==='Enter')lookup()">
  <input id="fn" placeholder="Имя"><input id="fu" placeholder="Фамилия"><input id="fc" list="dl-comp" placeholder="Компания (поиск)" size="18"><datalist id="dl-comp">${companies().map((c) => `<option value="${esc(c)}">`).join('')}</datalist></div>
  <div class="row"><input id="fb" placeholder="Блок" size="7"><input id="fr" placeholder="Комната" size="7"><button class="b" onclick="openPicker()">🏠 Выбрать комнату</button><span class="mu">Выезд (до какого числа):</span><input id="fo" type="date">
  <button class="b p" style="margin-left:auto" onclick="addEv()">+ Добавить</button></div></div>
  <div class="cols">${tbl('in', '🚌 Прибытие')}${tbl('out', '✈️ Отправление')}</div>`;
}
const setDate = (n) => { S.date = shiftDay(S.date, n); render(); };

// ───────── ДОБАВЛЕНИЕ / ПОИСК ПО ТАБ. № ─────────
function lookup() {
  const e = $('#fe').value.trim().toLowerCase();
  if (!e) return;
  const p = DB.people.find((x) => String(x.emp).trim().toLowerCase() === e) || DB.book.find((x) => String(x.emp).trim().toLowerCase() === e);
  if (!p) return toast('Табельный № не найден в проживании');
  $('#fn').value = p.name; $('#fu').value = p.sur; $('#fc').value = p.comp || '';
  if (p.key) { const [b, r] = p.key.split('|'); $('#fb').value = b; $('#fr').value = r; }
  toast('Данные подставлены: ' + full(p));
}
function addEv() {
  const e = { id: 0, type: $('#ft').value, date: S.date, time: $('#fm').value.trim(), name: $('#fn').value.trim(), sur: $('#fu').value.trim(), emp: $('#fe').value.trim(), comp: $('#fc').value.trim(), block: $('#fb').value.trim(), room: $('#fr').value.trim(), out: $('#fo').value, done: false };
  if (!e.name && !e.sur) return toast('Укажите имя');
  mut(() => { e.id = newId(); e.n = Math.max(0, ...DB.enc.filter((x) => x.date === S.date && x.type === e.type).map((x) => x.n)) + 1; DB.enc.push(e); const cs = companies(); if (e.comp && !cs.includes(e.comp)) DB.companies = [...cs, e.comp]; });
}
const del = (id) => { if (confirm('Удалить запись?')) mut(() => { DB.enc = DB.enc.filter((e) => e.id != id); }); };


// кто сейчас живёт в проживании: по табельному №, потом по имени и фамилии
function residentFor(e) {
  let p = e.emp ? DB.people.find((x) => x.emp && String(x.emp) === String(e.emp)) : null;
  if (!p) {
    const n1 = nrm(e.name) + nrm(e.sur), n2 = nrm(e.sur) + nrm(e.name);
    p = n1 ? DB.people.find((x) => { const m = nrm(x.name) + nrm(x.sur); return m && (m === n1 || m === n2); }) : null;
  }
  return p || null;
}
function depart(id) {
  mut(() => {
    const e = DB.enc.find((x) => x.id == id), p = residentFor(e);
    if (p) {
      const [b, r] = p.key.split('|');
      e.snap = p; e.block = b; e.room = r;
      DB.people = DB.people.filter((x) => x.id !== p.id);
      log('Выселение (Enclosement)', `${full(e)} из ${p.key}`);
      toast('Выселен из ' + kd(p.key));
    } else toast('Отмечен (в проживании не найден)');
    e.done = true;
  });
}
function undo(id) {
  mut(() => {
    const e = DB.enc.find((x) => x.id == id);
    if (e.type === 'in' && e.pid) DB.people = DB.people.filter((p) => p.id !== e.pid);
    if (e.type === 'out' && e.snap && !DB.people.some((p) => p.id === e.snap.id)) DB.people.push(e.snap);
    e.done = false; delete e.pid; delete e.snap;
  });
}

// ───────── ИМПОРТ EXCEL / ОТЧЁТ ─────────
const tm = (v) => typeof v === 'number' ? String(Math.floor(Math.round(v * 1440) % 1440 / 60)).padStart(2, '0') + ':' + String(Math.round(v * 1440) % 60).padStart(2, '0') : String(v || '').trim();
function importXlsx(f) {
  if (!f) return;
  const r = new FileReader();
  r.onload = (ev) => {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'array' });
      const ws = wb.Sheets['Прибытия и Отправления'] || wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const out = [];
      // значения-ошибки Excel (#N/A) считаем пустыми
      const c = (r, i) => { const s = String(r[i] ?? '').trim(); return s.startsWith('#') ? '' : s; };
      // в имени и фамилии не бывает чисто цифр (так иногда приходит код ошибки)
      const nmc = (r, i) => { const s = c(r, i); return /^\d+$/.test(s) ? '' : s; };
      let ni = 0, no = 0;
      for (let i = 4; i < raw.length; i++) {
        const r = raw[i];
        if (nmc(r, 3) || nmc(r, 4)) out.push({ type: 'in', n: ++ni, time: tm(r[1]), name: nmc(r, 3), sur: nmc(r, 4), emp: c(r, 5), comp: c(r, 6), block: c(r, 7), room: c(r, 8) });
        if (nmc(r, 19) || nmc(r, 20)) out.push({ type: 'out', n: ++no, time: tm(r[17]), name: nmc(r, 19), sur: nmc(r, 20), emp: c(r, 21), comp: c(r, 22), block: c(r, 23), room: c(r, 24) });
      }
      mut(() => {
        if (DB.enc.some((e) => e.date === S.date)) {
          if (!confirm('На ' + S.date + ' уже есть записи. Заменить их данными из файла?')) return;
          DB.enc = DB.enc.filter((e) => e.date !== S.date);
        }
        out.forEach((e) => DB.enc.push({ ...e, id: newId(), date: S.date, out: '', done: false }));
        log('Импорт Enclosement', `${S.date}: ${ni} прибытий, ${no} отправлений`);
        toast(`Импорт: ${ni} прибытий, ${no} отправлений на ${S.date}`);
      });
    } catch (err) { alert('Ошибка импорта: ' + err.message); }
  };
  r.readAsArrayBuffer(f);
}
function report() {
  const rows = DB.enc.filter((e) => e.date === S.date).sort((a, b) => a.type.localeCompare(b.type) || a.n - b.n)
    .map((e) => ({ 'Тип': e.type === 'in' ? 'Прибытие' : 'Отправление', '№': e.n, 'Время': e.time, 'Имя': e.name, 'Фамилия': e.sur, 'Таб. №': e.emp, 'Компания': e.comp, 'Блок': e.block, 'Комната': e.room, 'Отмечен': e.done ? 'да' : 'нет' }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Enclosement');
  XLSX.writeFile(wb, 'enclosement_' + S.date + '.xlsx');
}

// ───────── ВЫБОР КОМНАТЫ ─────────
const P = { blk: '', mode: 'all' };
function freeBeds(k) {
  const used =
    DB.people.filter((p) => p.key === k).length +
    DB.enc.filter((e) => e.date === S.date && e.type === 'in' && !e.done && key(e) === k).length +
    DB.book.filter((b) => b.key === k && b.from <= S.date && b.to >= S.date).length;
  return Math.max(0, 2 - used);
}



// ───────── КОМПАНИИ ─────────
const companies = () => {
  if (DB.companies && DB.companies.length) return [...DB.companies].sort((a, b) => a.localeCompare(b, 'ru'));
  return [...new Set([...DB.people.map((p) => p.comp), ...DB.enc.map((e) => e.comp)].map((s) => String(s || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ru'));
};
function editCompanies() {
  load();
  let pk = document.getElementById('pk');
  if (!pk) { pk = document.createElement('div'); pk.id = 'pk'; pk.className = 'pk'; pk.onclick = (e) => { if (e.target === pk) closePicker(); }; document.body.appendChild(pk); }
  pk.innerHTML = `<div class="pkb"><div class="row"><b class="big">Компании (${companies().length})</b><button class="b s" style="margin-left:auto" onclick="closePicker()">✕</button></div>
  <div class="mu">Одна компания на строке. Можно вставить весь список из Excel (столбцом).</div>
  <textarea id="cl" style="width:100%;height:340px;margin-top:8px;background:var(--cd);color:var(--tx);border:1px solid var(--bd);border-radius:8px;padding:8px">${esc(companies().join('\n'))}</textarea>
  <div class="row"><button class="b p" style="margin-left:auto" onclick="saveCompanies()">Сохранить</button></div></div>`;
  pk.style.display = 'block';
}
function saveCompanies() {
  const L = [...new Set($('#cl').value.split('\n').map((s) => s.trim()).filter(Boolean))];
  mut(() => { DB.companies = L; });
  closePicker(); toast('Сохранено: ' + L.length);
}

// ───────── ЗАКРЕПЛЁННЫЕ КОМНАТЫ, ВЫБОР КОМНАТЫ ДЛЯ ЧЕЛОВЕКА, ЗАСЕЛЕНИЕ ─────────
const nrm = (s) => String(s || '').toLowerCase().replace(/[^a-zа-яәғқңөұүһі0-9]/g, '');
const kd = (k) => k.replace('|', '-');
const outTxt = (p) => (p.out ? 'выезд ' + p.out.split('-').reverse().slice(0, 2).join('.') : 'дата выезда не указана');

// закреплённая комната: сначала по табельному № (если он уже запомнен), потом по имени и фамилии
function assignedFor(e) {
  const A = DB.assign || [];
  let a = e.emp ? A.find((x) => x.emp && String(x.emp) === String(e.emp)) : null;
  if (!a) {
    const n1 = nrm(e.name) + nrm(e.sur), n2 = nrm(e.sur) + nrm(e.name);
    a = n1 ? A.find((x) => { const m = nrm(x.name) + nrm(x.sur); return m && (m === n1 || m === n2); }) : null;
  }
  return a ? { key: a.key, shift: a.shift, a } : null;
}
// ячейка «Блок-комната» в строке: выбор для каждого человека + подсказка о закреплённой
function roomCell(e) {
  if (e.type === 'out' && !e.done) {
    const p = residentFor(e);
    if (p) return `<b>${kd(p.key)}</b><div class="mu">сейчас живёт${p.third ? ' (3-й)' : ''}${p.out ? ' · ' + outTxt(p) : ''}</div>`;
  }
  const k = e.room || e.block ? key(e) : '';
  const br = e.block === e.room ? e.block : e.block + (e.room ? '-' + e.room : '');
  if (e.done) return esc(br);
  let h = k ? `<a href="#" onclick="pickFor(${e.id});return false">${esc(br)}</a>` : `<button class="b s" onclick="pickFor(${e.id})">🏠 Выбрать</button>`;
  const ar = e.type === 'in' ? assignedFor(e) : null;
  if (ar && ar.key !== k) h += `<div class="mu">🔗 ${kd(ar.key)} · ${DB.people.filter((p) => p.key === ar.key).length}/2</div>`;
  return h;
}
const pickFor = (id) => { P.for = id; openPicker(); };
function pkShow(h) {
  let pk = document.getElementById('pk');
  if (!pk) { pk = document.createElement('div'); pk.id = 'pk'; pk.className = 'pk'; pk.onclick = (e) => { if (e.target === pk) closePicker(); }; document.body.appendChild(pk); }
  pk.innerHTML = h; pk.style.display = 'block';
}
function closePicker() { P.for = null; const pk = document.getElementById('pk'); if (pk) pk.style.display = 'none'; }

function openPicker() {
  load();
  const ev = P.for ? DB.enc.find((x) => x.id == P.for) : null, asg = ev ? assignedFor(ev) : null;
  const bl = [...new Set(Object.values(DB.rooms).map((r) => r.b))].sort((a, b) => a.localeCompare(b, 0, { numeric: true }));
  if (!bl.length) return toast('Нет комнат — сначала импортируйте Excel в «Проживании»');
  if (asg && P.last !== P.for) P.blk = asg.key.split('|')[0];
  P.last = P.for;
  if (!P.blk || !bl.includes(P.blk)) P.blk = bl[0];
  const cnt = {};
  Object.values(DB.rooms).forEach((r) => {
    const f = freeBeds(r.key), o = cnt[r.b] || (cnt[r.b] = { f2: 0, f1: 0 });
    if (f === 2) o.f2++; else if (f === 1) o.f1++;
  });
  const L = Object.values(DB.rooms).filter((r) => r.b === P.blk).map((r) => ({ r, f: freeBeds(r.key) }))
    .filter((x) => (asg && x.r.key === asg.key) || (P.mode === 'full' ? x.f === 2 : P.mode === 'half' ? x.f === 1 : x.f > 0))
    .sort((a, b) => a.r.n - b.r.n);
  const fls = [...new Set(L.map((x) => x.r.f))].sort();
  let h = `<div class="pkb"><div class="row"><b class="big">Выбор комнаты${ev ? ': ' + esc(full(ev)) : ''}</b><button class="b s" style="margin-left:auto" onclick="closePicker()">✕</button></div>`;
  if (asg) {
    const n = DB.people.filter((p) => p.key === asg.key).length;
    h += `<div class="row"><button class="b p" onclick="pick('${asg.key}')">🔗 Закреплённая: ${kd(asg.key)} · занято ${n}/2</button>${n >= 2 ? '<span class="mu">комната полная — можно заселить третьим</span>' : ''}</div>`;
  }
  h += `<div class="row" style="gap:3px">${bl.map((b) => `<button class="pill ${b === P.blk ? 'on' : ''}" onclick="P.blk='${b}';openPicker()">${b} <span class="mu">${cnt[b].f2}·${cnt[b].f1}</span></button>`).join('')}</div>
  <div class="row">${[['all', 'Все свободные'], ['full', 'Полностью свободные (2 места)'], ['half', 'Наполовину (1 место)']].map(([m, t]) => `<button class="pill ${P.mode === m ? 'on' : ''}" onclick="P.mode='${m}';openPicker()">${t}</button>`).join('')}</div>
  <div class="mu">Цифры у блока: пустых комнат · комнат с одной свободной койкой. Учтены проживающие, брони и ещё не заселённые на ${S.date}.</div>`;
  h += fls.map((f) => `<h3>${f} этаж</h3><div class="pg">${L.filter((x) => x.r.f === f).map((x) => `<div class="pr ${x.f === 1 ? 'h' : ''} ${asg && x.r.key === asg.key ? 'asg' : ''}" onclick="pick('${x.r.key}')"><b>${x.r.n}</b> ${asg && x.r.key === asg.key ? '🔗' : ''} <span class="c ${x.f === 2 ? 'g' : x.f === 1 ? 'y' : 'r'}">${x.f === 2 ? '2 места' : x.f === 1 ? '1 место' : 'занята'}</span><div class="mu">${esc(x.r.dept)}</div></div>`).join('')}</div>`).join('') || '<div class="mu" style="padding:20px">Нет подходящих комнат в этом блоке</div>';
  pkShow(h + '</div>');
}
function pick(k) {
  const [b, r] = k.split('|'), id = P.for;
  closePicker();
  if (id) return mut(() => { const e = DB.enc.find((x) => x.id == id); if (e) { e.block = b; e.room = r; } });
  $('#fb').value = b; $('#fr').value = r; toast('Выбрана комната ' + b + '-' + r);
}

// «Приехал»
function arrive(id) {
  load();
  const e = DB.enc.find((x) => x.id == id), ar = assignedFor(e), k = e.room || e.block ? key(e) : '';
  if (!k) return ar ? askRoom(id, ar) : (toast('Выберите комнату'), pickFor(id));      // комнаты нет — спрашиваем / открываем выбор
  if (ar && k === ar.key && DB.people.filter((p) => p.key === k).length >= 2) return askRoom(id, ar); // закреплённая полная
  doArrive(id, k, false);
}
function askRoom(id, ar) {
  const e = DB.enc.find((x) => x.id == id), k = ar.key, res = DB.people.filter((p) => p.key === k);
  pkShow(`<div class="pkb" style="max-width:480px"><div class="row"><b class="big">Заселение: ${esc(full(e))}</b><button class="b s" style="margin-left:auto" onclick="closePicker()">✕</button></div>
  <div>Закреплена комната <b>${kd(k)}</b>${ar.shift ? ' (вахта ' + ar.shift + ')' : ''}.</div>
  ${res.length ? `<div class="card">Сейчас в комнате (${res.length}/2):${res.map((p) => `<div>${esc(full(p))} <span class="mu">· ${outTxt(p)}</span></div>`).join('')}</div>` : '<div class="mu">Комната пустая</div>'}
  <div class="row">${res.length < 2
    ? `<button class="b p" onclick="qGo(${id},'${k}',false)">✅ В закреплённую ${kd(k)}</button>`
    : `<button class="b p" onclick="qGo(${id},'${k}',true)">⚠ Занято — заселить третьим</button>`}
  <button class="b" onclick="qOther(${id})">🏠 В другую комнату</button></div>
  ${res.length >= 2 ? '<div class="mu">Третьим можно: сменщик из этой комнаты скоро выезжает.</div>' : ''}</div>`);
}
const qGo = (id, k, third) => { closePicker(); doArrive(id, k, third); };
const qOther = (id) => { closePicker(); pickFor(id); };
function doArrive(id, k, third) {
  mut(() => {
    const e = DB.enc.find((x) => x.id == id), r = DB.rooms[k];
    if (!r) { e.done = true; return toast('Отмечен (комнаты нет в проживании: ' + (e.block || '—') + ')'); }
    const n = DB.people.filter((p) => p.key === k).length, ar = assignedFor(e);
    if (n >= 2 && !third) return toast('Комната ' + kd(k) + ' занята — выберите другую');
    if (n >= 2 && !(ar && ar.key === k)) return toast('Третьим можно заселять только в закреплённую комнату');
    const p = { id: newId(), key: k, name: e.name, sur: e.sur, emp: e.emp, comp: e.comp, dept: r.dept || '', shift: ar ? ar.shift : '', in: e.date, out: e.out || '', third: n >= 2 };
    DB.people.push(p); e.pid = p.id; e.block = r.b; e.room = String(r.n); e.done = true;
    if (ar && e.emp && !ar.a.emp) ar.a.emp = e.emp;                         // запоминаем табельный № в закреплении
    DB.book = DB.book.filter((b) => !(b.key === k && same(b, e)));
    log('Заселение (Enclosement)', `${full(e)} → ${k}${p.third ? ' (третьим)' : ''}`);
    toast('Заселён в ' + kd(k) + (p.third ? ' третьим' : ''));
  });
}

// ───────── РЕДАКТИРОВАНИЕ ENCLOSEMENT ─────────
function editEnc(id) {
  const e = DB.enc.find((x) => x.id == id);
  if (!e) return;

  let h = `<div class="row"><b class="big">Редактировать запись (№${e.n})</b><button class="b s" style="margin-left:auto" onclick="closePicker()">✕</button></div>
  <div class="row"><input id="ee_name" placeholder="Имя" value="${esc(e.name)}"><input id="ee_sur" placeholder="Фамилия" value="${esc(e.sur)}"></div>
  <div class="row"><input id="ee_emp" placeholder="Таб. №" value="${esc(e.emp)}"><input id="ee_comp" placeholder="Компания" value="${esc(e.comp)}"></div>
  <div class="row"><input id="ee_time" placeholder="Время" value="${esc(e.time)}"><input id="ee_blk" placeholder="Блок" value="${esc(e.block)}"><input id="ee_rm" placeholder="Комната" value="${esc(e.room)}"></div>
  <div class="row" style="margin-top:15px"><button class="b p" style="width:100%" onclick="saveEnc(${e.id})">Сохранить</button></div>`;
  
  pkShow(`<div class="pkb" style="max-width:500px">${h}</div>`);
}

function saveEnc(id) {
  const e = DB.enc.find((x) => x.id == id);
  if (!e) return;

  const name = document.getElementById('ee_name').value.trim();
  const sur = document.getElementById('ee_sur').value.trim();
  const emp = document.getElementById('ee_emp').value.trim();
  const comp = document.getElementById('ee_comp').value.trim();
  const time = document.getElementById('ee_time').value.trim();
  const blk = document.getElementById('ee_blk').value.trim();
  const rm = document.getElementById('ee_rm').value.trim();

  mut(() => {
    const e = DB.enc.find((x) => x.id == id); // после load() объекты новые — берём свежий
    if (!e) return;
    e.name = name;
    e.sur = sur;
    e.emp = emp;
    e.comp = comp;
    e.time = time;
    e.block = blk;
    e.room = rm;
    const pp = e.pid && DB.people.find((x) => x.id === e.pid); // уже заселён — обновляем и в проживании
    if (pp) { pp.name = name; pp.sur = sur; pp.emp = emp; pp.comp = comp; }
    log('Редактирование Enclosement', `№${e.n} ${e.name} ${e.sur}`);
  });

  closePicker();
  toast('Запись обновлена');
}

// обновление экрана, когда из облака пришли чужие изменения (sync.js)
window.onSbData = () => { load(); render(); };
render();
