const $ = (s) => document.querySelector(s),
  esc = (s) =>
    String(s ?? '').replace(
      /[&<>"]/g,
      (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])
    );
let DB = { rooms: {}, people: [], assign: [], hist: [], nid: 1, book: [] };
try {
  Object.assign(DB, JSON.parse(localStorage.getItem('turan3') || '{}'));
} catch (e) {}
const save = () => {
  try {
    localStorage.setItem('turan3', JSON.stringify(DB));
  } catch (e) {
    alert('⚠ Не удалось сохранить данные (память браузера переполнена). Скачайте резервную копию: Данные → Резервная копия.');
  }
  if (window.sbPush) sbPush(); // отправка в облако (sync.js)
};
const S = {
  page: 'dashboard',
  blk: '',
  q: '',
  dept: '',
  fl: 0,
  free: false,
  open: null,
  mt: 'res',
  dashDept: '',
};
const iso = (d) =>
    new Date(d - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10),
  today = () => iso(new Date());
const addDays = (s, n) => {
  const d = new Date(s + 'T00:00:00');
  d.setDate(d.getDate() + n);
  return iso(d);
};
const diff = (s) =>
  Math.round(
    (new Date(s + 'T00:00:00') - new Date(today() + 'T00:00:00')) / 864e5
  );
const fd = (s) => (s ? s.split('-').reverse().slice(0, 2).join('.') : '—');
const log = (a, d) => {
  DB.hist.unshift({ t: new Date().toLocaleString('ru-RU'), a, d });
  DB.hist = DB.hist.slice(0, 100);
};
const toast = (m) => {
  const t = $('#toast');
  t.textContent = m;
  t.style.display = 'block';
  setTimeout(() => (t.style.display = 'none'), 2200);
};
const BL = () =>
  [...new Set(Object.values(DB.rooms).map((r) => r.b))].sort((a, b) =>
    a.localeCompare(b, 0, { numeric: true })
  );
  const depts = () => {
    if (!DB.depts) {
      const seed = [
        ...new Set(
          Object.values(DB.rooms)
            .map((r) => r.dept)
            .concat(DB.people.map((p) => p.dept))
            .filter(Boolean)
        ),
      ];
      if (!seed.length) return [];
      DB.depts = seed;
      save();
    }
    return [...DB.depts].sort((a, b) => a.localeCompare(b, 'ru'));
  };
const idx = (arr) => {
  const m = {};
  arr.forEach((p) => (m[p.key] ??= []).push(p));
  return m;
};
const nm = (p) => (p.name + ' ' + p.sur).trim();
const opts = (a, sel, all) =>
  (all ? `<option value="">${all}</option>` : '') +
  a
    .map((x) => `<option ${x == sel ? 'selected' : ''}>${esc(x)}</option>`)
    .join('');
function chip(p) {
  if (!p.out) return '<span class="c n">нет даты</span>';
  const d = diff(p.out);
  if (d < 0) return `<span class="c r">просрочен ${-d}д</span>`;
  if (d == 0) return '<span class="c r">выезд сегодня</span>';
  return `<span class="c ${d <= 2 ? 'r' : d <= 5 ? 'y' : 'g'}">${d} дн.</span>`;
}

// ───────── ИМПОРТ EXCEL ─────────
function importXlsx(f) {
  if (!f) return;
  if (DB.people.length && !confirm('Импорт заменит список проживающих (сейчас ' + DB.people.length + ') данными из Excel. Заселённые позже, которых нет в файле, пропадут.\nПеред импортом будет скачана резервная копия. Продолжить?')) return;
  if (DB.people.length) backup();
  const r = new FileReader();
  r.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array' }),
        rooms = {},
        people = [],
        assign = [];
      const old = {};
      DB.people.forEach((p) => (old[p.key + '|' + nm(p).toLowerCase()] = p));
      let nid = DB.nid;
      wb.SheetNames.filter((n) => /^[ABC]\d+$/i.test(n.trim())).forEach(
        (sn) => {
          const raw = XLSX.utils.sheet_to_json(wb.Sheets[sn], {
              header: 1,
              defval: '',
            }),
            b = sn.trim().toUpperCase();
          [0, 15, 30].forEach((o, fi) => {
            let cur = null,
              k = 0;
            for (let i = 6; i < raw.length; i++) {
              const r = raw[i],
                rm = String(r[1 + o] || '').trim(),
                c = (x) => String(r[x + o] ?? '').trim();
              if (/^\d{3}$/.test(rm)) {
                cur = { key: b + '|' + rm, b, n: +rm, f: fi + 1, dept: '' };
                rooms[cur.key] = cur;
                k = 0;
              }
              if (!cur || k++ >= 2) continue;
              if (!cur.dept && c(9)) cur.dept = c(9);
              if (c(2) || c(3)) {
                const key = cur.key,
                  o2 =
                    old[key + '|' + (c(2) + ' ' + c(3)).trim().toLowerCase()];
                people.push({
                  id: newId(),
                  key,
                  name: c(2),
                  sur: c(3),
                  emp: c(4),
                  comp: c(5),
                  dept: c(9),
                  shift: o2?.shift || '',
                  in: o2?.in || '',
                  out: o2?.out || '',
                });
              }
              [
                [10, 'A'],
                [12, 'B'],
              ].forEach(([x, s]) => {
                if (c(x) || c(x + 1))
                  assign.push({
                    id: newId(),
                    key: cur.key,
                    shift: s,
                    name: c(x),
                    sur: c(x + 1),
                  });
              });
            }
          });
        }
      );
      people.forEach((p) => {
        if (p.shift) return;
        const a = assign.find(
          (a) =>
            a.key == p.key &&
            a.name.toLowerCase() == p.name.toLowerCase() &&
            a.sur.toLowerCase() == p.sur.toLowerCase()
        );
        p.shift = a ? a.shift : '';
      });
      const oldA = {};
      (DB.assign || []).forEach((a) => { if (a.emp) oldA[a.key + '|' + nm(a).toLowerCase() + '|' + a.shift] = a.emp; });
      assign.forEach((a) => { const e = oldA[a.key + '|' + nm(a).toLowerCase() + '|' + a.shift]; if (e) a.emp = e; });
      DB.rooms = rooms;
      DB.depts = [
        ...new Set(
          [
            ...(DB.depts || []),
            ...Object.values(rooms).map((r) => r.dept),
            ...people.map((p) => p.dept),
          ].filter(Boolean)
        ),
      ];
      DB.people = people;
      DB.assign = assign;
      DB.nid = nid;
      log(
        'Импорт',
        `${Object.keys(rooms).length} комнат, ${people.length} проживающих, ${
          assign.length
        } закреплений`
      );
      save();
      S.blk = BL()[0] || '';
      toast('Импорт завершён');
      go('board');
    } catch (err) {
      alert('Ошибка импорта: ' + err.message);
    }
  };
  r.readAsArrayBuffer(f);
}

// ───────── НАВИГАЦИЯ ─────────
function go(p) {
  S.page = p;
  if (!S.blk) S.blk = BL()[0] || '';
  $('#tabs').innerHTML = BL()
    .map(
      (b) =>
        `<button class="${
          b == S.blk && S.page == 'board' ? 'on' : ''
        }" onclick="S.blk='${b}';S.q='';$('#q').value='';go('board')">${b}</button>`
    )
    .join('');
  document
    .querySelectorAll('nav button')
    .forEach((b) => b.classList.toggle('on', b.dataset.p == p));
  const n = DB.people.length,
    t = Object.keys(DB.rooms).length * 2,
    l = DB.people.filter((p) => p.out && diff(p.out) <= 2).length;
  $('#stats').innerHTML = `Занято <b>${n}</b> / ${t} · Свободно мест <b>${Math.max(0, t - n)}</b> · Выезд ≤2 дн: <b>${l}</b>`;
  if (p === 'dashboard') pDashboard();
  else
    ({
      board,
      search: pSearch,
      assign: pAssign,
      out: pOut,
      book: pBook,
      employees: pEmployees,
      depts: pDepts,
      data: pData,
    })[p]();
}

const empty =
  '<div class="mu" style="padding:40px;text-align:center">Нет данных. Откройте «💾 Данные» и импортируйте Excel (ADR).</div>';

// ───────── ДАШБОРД ─────────
function pDashboard() {
  const deptsList = depts();
  const selDept = S.dashDept.toLowerCase();

  const blocks = BL();
  const matrix = {};
  DB.people.forEach((p) => {
    if (selDept && (!p.dept || !p.dept.toLowerCase().includes(selDept))) return;
    const comp = p.comp || 'Без компании';
    if (!matrix[comp]) matrix[comp] = { total: 0 };
    const b = DB.rooms[p.key]?.b || '?';
    matrix[comp][b] = (matrix[comp][b] || 0) + 1;
    matrix[comp].total++;
  });

  let h = `<div class="bar">
   <input type="text" id="dash-dept" list="dl-dash" placeholder="Умный поиск департамента..." oninput="S.dashFocus=1;S.dashDept=this.value;clearTimeout(window.dt);window.dt=setTimeout(pDashboard,300)" value="${esc(
     S.dashDept
   )}" style="width:260px">
   <datalist id="dl-dash">${deptsList
     .map((d) => `<option value="${esc(d)}">`)
     .join('')}</datalist>
   <button class="b p" onclick="exportDashboard()">📥 Экспорт в Excel (ADR)</button>
 </div>`;

  h += `<h3>Сводка по компаниям и блокам</h3><div style="overflow-x:auto"><table class="matrix-table"><thead><tr><th class="sticky-col">Компания</th>`;
  blocks.forEach((b) => (h += `<th>${b}</th>`));
  h += `<th style="background:rgba(79,142,247,.2);color:var(--ac)">ВСЕГО</th></tr></thead><tbody>`;

  const comps = Object.keys(matrix).sort();
  if (!comps.length) {
    h += `<tr><td colspan="${
      blocks.length + 2
    }" style="text-align:center;color:var(--mu)">Нет данных</td></tr>`;
  } else {
    comps.forEach((c) => {
      h += `<tr><td class="sticky-col">${esc(c)}</td>`;
      blocks.forEach((b) => {
        const val = matrix[c][b] || 0;
        h += `<td class="${val ? '' : 'cell-zero'}">${val || '—'}</td>`;
      });
      h += `<td style="font-weight:800;color:var(--ac)">${matrix[c].total}</td></tr>`;
    });
  }
  h += `</tbody></table></div>`;

  if (S.dashDept) {
    h += `<h3>Комнаты по запросу: "${esc(
      S.dashDept
    )}"</h3><div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:8px;">`;
    const filteredRooms = Object.values(DB.rooms).filter((r) => {
      if (r.dept && r.dept.toLowerCase().includes(selDept)) return true;
      const rp = DB.people.filter((p) => p.key === r.key);
      return rp.some((p) => p.dept && p.dept.toLowerCase().includes(selDept));
    });

    if (!filteredRooms.length) h += `<div class="mu">Ничего не найдено</div>`;
    else {
      filteredRooms
        .sort((a, b) => a.b.localeCompare(b.b) || a.n - b.n)
        .forEach((r) => {
          const res = DB.people.filter((p) => p.key === r.key);
          h += `<div class="rc ${
            res.length == 0 ? '' : res.length == 1 ? 'h' : 'f'
          }" onclick="openRoom('${r.key}')">
         <div class="t"><span>${r.b}-${r.n}</span><small>${esc(
            r.dept
          )}</small></div>
         ${
           res
             .map(
               (p) =>
                 `<div class="bed"><span>${esc(nm(p))}</span>${chip(p)}</div>`
             )
             .join('') || '<div class="bed fr">Свободно</div>'
         }
       </div>`;
        });
    }
    h += `</div>`;
  } else {
    h += `<h3>Обзор блоков</h3><div class="fl">`;
    blocks.forEach((b) => {
      const roomsInB = Object.values(DB.rooms).filter((r) => r.b === b);
      const totalPeople = DB.people.filter(
        (p) => DB.rooms[p.key]?.b === b
      ).length;
      h += `<div class="rc" onclick="S.blk='${b}';go('board')">
       <div class="t"><span>Блок ${b}</span><small>${roomsInB.length} комнат</small></div>
       <div class="bed"><span>Проживает:</span><b>${totalPeople} чел.</b></div>
     </div>`;
    });
    h += `</div>`;
  }
  $('#main').innerHTML = h;
  if (S.dashFocus) {
    S.dashFocus = 0;
    const di = $('#dash-dept');
    if (di) { di.focus(); di.setSelectionRange(di.value.length, di.value.length); }
  }
}
// ───────── ЭКСПОРТ В EXCEL (формат ADR) ─────────
// Собирает книгу в том же виде, что рабочий файл ADR: сводный лист + по листу на каждый блок
// (3 этажа рядом, по 2 строки на комнату), «Список ADR», «Проверка закреплений».
// Дополнительно — лист «Проживание» с датами заезда/выезда (в ADR их нет).
// Excel-формулы остаются живыми: при правке в файле всё пересчитывается.
const ADR_ORDER = ['A1','A2','A3','B1','C1','A4','A5','A6','B2','C2','A7','A8','A9','B3','C3','A10','A11','A12','B4','C4'];
const ADR_LIST = 'Список ADR', ADR_CHECK = 'Проверка закреплений', ADR_SUM = 'Blocks Cumulative Summary';
const ADR_W = [4.5, 6, 18, 22, 30, 24, 5, 3, 22, 26, 13, 16, 13, 16, 2]; // ширины колонок одного этажа
const ADR_FILL = { blue: 'BDD6EE', blue2: 'BDD7EE', room: 'FBE4D5', mid: '9CC2E5', pos: '00FF00', wa: 'FFFF00', wb: '00FFFF', flr: 'FFFF00', hdr: 'ED7D31', white: 'FFFFFF', sum: 'D9E1F2' };
const adrCol = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const adrQ = (b) => "'" + String(b).replace(/'/g, "''") + "'";
const adrOrd = (f) => (f == 1 ? '1ST' : f == 2 ? '2ND' : f == 3 ? '3RD' : f + 'TH');

function adrPlan() {
  const I = idx(DB.people), A = idx(DB.assign);
  const blocks = BL();
  const order = [...ADR_ORDER.filter((b) => blocks.includes(b)), ...blocks.filter((b) => !ADR_ORDER.includes(b))];
  const compOf = (p) => String(p.comp || '').trim();

  // ---- мета по блокам: этажи, позиции статистики ----
  const meta = {};
  order.forEach((b) => {
    const rs = Object.values(DB.rooms).filter((r) => r.b === b).sort((x, y) => x.n - y.n);
    const fl = {};
    rs.forEach((r) => (fl[r.f] = fl[r.f] || []).push(r));
    const groups = Object.keys(fl).map(Number).sort((x, y) => x - y)
      .map((f) => ({ f, o: (f - 1) * 15, rooms: fl[f], last: 6 + 2 * fl[f].length }));
    const maxR = Math.max(1, ...groups.map((g) => g.rooms.length));
    const people = rs.flatMap((r) => I[r.key] || []);
    const comps = [...new Set(people.map(compOf).filter(Boolean))].sort((x, y) => x.localeCompare(y, 'ru')).map((c) => ({ label: c, blank: false }));
    if (people.some((p) => !compOf(p)) || !comps.length) comps.push({ label: 'Без компании', blank: true });
    const s0 = 8 + 2 * maxR, c0 = s0 + 6;
    meta[b] = { groups, maxR, s0, c0, totalRow: c0 + comps.length, comps };
  });

  // ---- «слоты» = строки всех комнат (общие для Списка ADR и Проверки закреплений) ----
  const slots = [];
  order.forEach((b) => meta[b].groups.forEach((g) => g.rooms.forEach((r, i) => {
    for (let j = 0; j < 2; j++) slots.push({ b, o: g.o, row: 7 + 2 * i + j, room: r.n });
  })));
  const NL = slots.length + 1; // последняя строка Списка ADR

  const sheets = [];
  const mk = (name) => { const sh = { name, cells: {}, merges: [], widths: {}, heights: {}, cf: [], filter: null }; sheets.push(sh); return sh; };
  const put = (sh, r, c, val, s) => { sh.cells[r + ',' + c] = { val, s: s || {} }; };
  const box = (sh, r1, c1, r2, c2, val, s) => { // объединённая область со стилем на каждой ячейке
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) put(sh, r, c, r == r1 && c == c1 ? val : null, s);
    if (r1 != r2 || c1 != c2) sh.merges.push([r1, c1, r2, c2]);
  };
  const F = (f) => ({ f });
  const redCF = (sh, ref) => sh.cf.push({ ref, text: '❌' });

  // ================= 1. СВОДНЫЙ ЛИСТ =================
  const sm = mk(ADR_SUM);
  const cols = [];
  blocks.forEach((b) => meta[b] && meta[b].groups.forEach((g) => cols.push({ b, g, label: b + '-' + g.f })));
  const nC = cols.length, tc = nC + 2; // колонка TOTAL
  sm.widths[1] = 46; for (let c = 2; c <= tc; c++) sm.widths[c] = 9;
  box(sm, 1, 1, 1, tc, 'Blocks Cumulative Summary', { b: true, sz: 14, al: 'left' });
  const hdr = (r, title) => {
    put(sm, r, 1, title, { b: true, fill: ADR_FILL.sum, bd: 'thin' });
    cols.forEach((c, i) => put(sm, r, i + 2, c.label, { b: true, al: 'center', fill: ADR_FILL.sum, bd: 'thin' }));
    put(sm, r, tc, 'TOTAL', { b: true, al: 'center', fill: ADR_FILL.sum, bd: 'thin' });
  };
  hdr(2, '');
  const statRows = [
    ['Number of FREE ROOMS', 0], ['Number of FREE PLACES', 1], ['Number of OCCUPIED PLACES', 2],
    ['Number of Rooms / Количество комнат', 3], ['Total number of PLACES / Общее количество мест', 4],
  ];
  statRows.forEach(([lab, k], ri) => {
    const r = 3 + ri;
    put(sm, r, 1, lab, { b: true, bd: 'thin' });
    cols.forEach((c, i) => put(sm, r, i + 2, F(`${adrQ(c.b)}!${adrCol(c.g.o + 5)}${meta[c.b].s0 + k}`), { al: 'center', bd: 'thin' }));
    put(sm, r, tc, F(`SUM(B${r}:${adrCol(tc - 1)}${r})`), { b: true, al: 'center', bd: 'thin' });
  });
  // компании
  const allComps = [...new Set(DB.people.map(compOf).filter(Boolean))].sort((x, y) => x.localeCompare(y, 'ru'));
  const compRows = allComps.map((c) => ({ label: c, blank: false }));
  if (DB.people.some((p) => !compOf(p))) compRows.push({ label: 'Без компании', blank: true });
  const r0 = 10;
  hdr(r0 - 1, 'COMPANIES');
  compRows.forEach((cr, ri) => {
    const r = r0 + ri;
    put(sm, r, 1, cr.label, { bd: 'thin' });
    cols.forEach((c, i) => {
      const Fc = adrCol(c.g.o + 6), Gc = adrCol(c.g.o + 7), q = adrQ(c.b);
      const rng = (L) => `${q}!$${L}$7:$${L}$${c.g.last}`;
      put(sm, r, i + 2, F(cr.blank ? `COUNTIFS(${rng(Fc)},"",${rng(Gc)},1)` : `COUNTIF(${rng(Fc)},$A${r})`), { al: 'center', bd: 'thin' });
    });
    put(sm, r, tc, F(`SUM(B${r}:${adrCol(tc - 1)}${r})`), { b: true, al: 'center', bd: 'thin' });
  });
  const rT = r0 + compRows.length;
  put(sm, rT, 1, 'TOTAL', { b: true, fill: ADR_FILL.sum, bd: 'thin' });
  for (let c = 2; c <= tc; c++) put(sm, rT, c, compRows.length ? F(`SUM(${adrCol(c)}${r0}:${adrCol(c)}${rT - 1})`) : 0, { b: true, al: 'center', fill: ADR_FILL.sum, bd: 'thin' });
  // загруженность
  const rl = rT + 3;
  hdr(rl - 1, '');
  put(sm, rl, 1, 'Загруженность (%)', { b: true, bd: 'thin' });
  for (let c = 2; c <= tc; c++) put(sm, rl, c, F(`IFERROR(${adrCol(c)}5/${adrCol(c)}7,0)`), { al: 'center', bd: 'thin', nf: '0%' });
  const letters = [...new Set(cols.map((c) => c.b[0]))].sort().reverse(); // C, B, A — как в образце
  letters.forEach((L, i) => {
    const cc = cols.map((c, k) => (c.b[0] == L ? adrCol(k + 2) : null)).filter(Boolean);
    put(sm, rl + 1 + i, 1, 'Тотал ' + L, { b: true, bd: 'thin' });
    put(sm, rl + 1 + i, 2, F(`IFERROR(SUM(${cc.map((x) => x + '5').join(',')})/SUM(${cc.map((x) => x + '7').join(',')}),0)`), { al: 'center', bd: 'thin', nf: '0%' });
  });
  put(sm, rl + 1 + letters.length, 1, 'Grand Total', { b: true, fill: ADR_FILL.sum, bd: 'thin' });
  put(sm, rl + 1 + letters.length, 2, F(`IFERROR(${adrCol(tc)}5/${adrCol(tc)}7,0)`), { b: true, al: 'center', fill: ADR_FILL.sum, bd: 'thin', nf: '0%' });

  // ================= 2. ЛИСТЫ БЛОКОВ =================
  order.forEach((b) => {
    const m = meta[b], sh = mk(b), isA = b[0] == 'A';
    const asg = {};
    m.groups.forEach((g) => g.rooms.forEach((r) => {
      const L = A[r.key] || [];
      asg[r.key] = { A: L.filter((a) => a.shift == 'A'), B: L.filter((a) => a.shift == 'B') };
    }));
    m.groups.forEach((g) => {
      const o = g.o, c = (k) => o + k, L = (k) => adrCol(o + k);
      ADR_W.forEach((w, k) => (sh.widths[o + k + 1] = w));
      // шапка
      box(sh, 4, c(4), 4, c(5), b, { b: true, sz: 14 });
      box(sh, 5, c(1), 5, c(8), `${adrOrd(g.f)} FLOOR / ${g.f} ЭТАЖ`, { b: true, fill: ADR_FILL.flr, bd: 'medium' });
      const H = (k, t, s) => put(sh, 6, c(k), t, Object.assign({ b: true, al: 'center', wrap: true, bd: 'medium' }, s));
      H(1, 'Resident№'); H(2, 'Room         №'); H(3, 'Name'); H(4, 'Surname'); H(5, 'Employee Number'); H(6, 'Company');
      H(7, 'Number for total amount control', { sz: 8 }); H(8, 'Booking Plan', { sz: 8 });
      H(10, isA ? 'POSITION' : 'Department');
      box(sh, 6, c(11), 6, c(12), isA ? 'Вахта А' : 'Закрепления', { b: true, fill: ADR_FILL.hdr, bd: 'medium' });
      box(sh, 6, c(13), 6, c(14), isA ? 'Вахта В' : 'Закрепления', { b: true, fill: ADR_FILL.hdr, bd: 'medium' });
      sh.heights[6] = 34.5;
      // комнаты
      g.rooms.forEach((r, i) => {
        const r1 = 7 + 2 * i, r2 = r1 + 1, res = (I[r.key] || []).slice(0, 2), as = asg[r.key];
        const bg = i % 2 == 0 ? ADR_FILL.blue : ADR_FILL.white;
        box(sh, r1, c(2), r2, c(2), r.n, { b: true, al: 'center', fill: ADR_FILL.room });
        box(sh, r1, c(8), r2, c(8), F(`IF(${L(7)}${r1}+${L(7)}${r2}>0,"1","0")`), { al: 'center', fill: ADR_FILL.mid });
        box(sh, r1, c(15), r2, c(15), F(`COUNTA(${L(6)}${r1}:${L(6)}${r2})`), { al: 'center', fill: ADR_FILL.mid });
        for (let j = 0; j < 2; j++) {
          const rr = r1 + j, p = res[j], a1 = as.A[j], a2 = as.B[j];
          const T = (v) => (v === undefined || v === null || v === '' ? null : v);
          put(sh, rr, c(1), F(`IF(${L(9)}${rr}="","",IF(COUNTIFS('${ADR_LIST}'!$A$2:$A$${NL},${L(3)}${rr}&"",'${ADR_LIST}'!$B$2:$B$${NL},${L(4)}${rr}&"",'${ADR_LIST}'!$C$2:$C$${NL},${L(5)}${rr}&"")>1,"❌","✅"))`), { fill: ADR_FILL.blue });
          put(sh, rr, c(3), p ? T(p.name) : null, { fill: bg });
          put(sh, rr, c(4), p ? T(p.sur) : null, { fill: bg });
          put(sh, rr, c(5), p ? T(p.emp) : null, { fill: bg });
          put(sh, rr, c(6), p ? T(p.comp) : null, { fill: bg });
          put(sh, rr, c(7), F(`COUNTA(${L(3)}${rr})`), { al: 'center', fill: ADR_FILL.blue });
          put(sh, rr, c(9), F(`IF(OR(ISBLANK(${L(3)}${rr}),ISBLANK(${L(4)}${rr})),"",CONCATENATE(${L(3)}${rr}," ",${L(4)}${rr},"",${L(5)}${rr}))`), { fill: ADR_FILL.mid, bd: 'none' });
          put(sh, rr, c(10), T((p && p.dept) || r.dept), { b: true, fill: ADR_FILL.pos });
          put(sh, rr, c(11), a1 ? T(a1.name) : null, { b: true, fill: ADR_FILL.wa });
          put(sh, rr, c(12), a1 ? T(a1.sur) : null, { b: true, fill: ADR_FILL.wa });
          put(sh, rr, c(13), a2 ? T(a2.name) : null, { b: true, fill: ADR_FILL.wb });
          put(sh, rr, c(14), a2 ? T(a2.sur) : null, { b: true, fill: ADR_FILL.wb });
        }
      });
      redCF(sh, `${L(1)}7:${L(1)}${g.last}`);
      // статистика
      const s0 = m.s0, V = L(5);
      const SR = (i, text, v) => { put(sh, s0 + i, c(4), text, { b: true, wrap: true, bd: 'thin' }); put(sh, s0 + i, c(5), v, { b: true, al: 'center', bd: 'thin' }); };
      const rng = (k) => `$${L(k)}$7:$${L(k)}$${g.last}`;
      SR(0, 'Number of FREE ROOMS/Количество свободных номеров:', F(`${g.rooms.length}-COUNTIF(${rng(8)},"1")`));
      SR(1, 'Number of FREE PLACES/Количество свободных мест:', F(`${V}${s0 + 4}-${V}${s0 + 2}`));
      SR(2, 'Number of OCCUPIED PLACES/Количество оккупированных мест:', F(`SUM(${rng(7)})`));
      SR(3, 'Number of Rooms', g.rooms.length);
      SR(4, 'Total number of PLACES/Общее количество мест:', F(`${V}${s0 + 3}*2`));
      put(sh, s0 + 5, c(4), 'OCCUPANTS BY COMPANY', { b: true, fill: ADR_FILL.sum, bd: 'thin' });
      put(sh, s0 + 5, c(5), null, { fill: ADR_FILL.sum, bd: 'thin' });
      m.comps.forEach((cm, i) => {
        const rr = m.c0 + i;
        put(sh, rr, c(4), cm.label, { bd: 'thin' });
        put(sh, rr, c(5), F(cm.blank ? `COUNTIFS(${rng(6)},"",${rng(7)},1)` : `COUNTIF(${rng(6)},${L(4)}${rr})`), { al: 'center', bd: 'thin' });
      });
      put(sh, m.totalRow, c(4), 'TOTAL', { b: true, fill: ADR_FILL.sum, bd: 'thin' });
      put(sh, m.totalRow, c(5), F(`SUM(${V}${m.c0}:${V}${m.totalRow - 1})`), { b: true, al: 'center', fill: ADR_FILL.sum, bd: 'thin' });
    });
  });

  // ================= 3. СПИСОК ADR =================
  const ls = mk(ADR_LIST);
  ['Name', 'Surname', 'Employee Number', 'Company', 'BLOCK', 'Room', 'Name Surname'].forEach((t, i) => put(ls, 1, i + 1, t, { b: true, fill: ADR_FILL.sum, bd: 'thin' }));
  [16, 20, 28, 28, 8, 8, 30].forEach((w, i) => (ls.widths[i + 1] = w));
  slots.forEach((s, i) => {
    const r = i + 2, q = adrQ(s.b);
    [3, 4, 5, 6].forEach((k, ci) => { const ref = `${q}!${adrCol(s.o + k)}${s.row}`; put(ls, r, ci + 1, F(`IF(${ref}="","",${ref})`), {}); });
    put(ls, r, 5, s.b, {}); put(ls, r, 6, s.room, {});
    put(ls, r, 7, F(`IF(OR(A${r}="",B${r}=""),"",A${r}&" "&B${r})`), {});
  });
  ls.filter = `A1:G${NL}`;

  // ================= 4. ПРОВЕРКА ЗАКРЕПЛЕНИЙ =================
  const ck = mk(ADR_CHECK);
  ['✅/ ❌', 'Name', 'Surname', 'Name Surname', 'Block', 'Room', '✅/ ❌', 'Name', 'Surname', 'Name Surname', 'Block', 'Room']
    .forEach((t, i) => put(ck, 1, i + 1, t, { b: true, fill: i < 6 ? ADR_FILL.wa : ADR_FILL.wb, bd: 'thin' }));
  [8, 16, 20, 30, 8, 8, 8, 16, 20, 30, 8, 8].forEach((w, i) => (ck.widths[i + 1] = w));
  slots.forEach((s, i) => {
    const r = i + 2, q = adrQ(s.b);
    const ref = (k) => `${q}!${adrCol(s.o + k)}${s.row}`;
    put(ck, r, 1, F(`IF(D${r}="","",IF(COUNTIF($D$2:$D$${NL},D${r})+COUNTIF($J$2:$J$${NL},D${r})>1,"❌","✅"))`), {});
    put(ck, r, 2, F(`IF(${ref(11)}="","",${ref(11)})`), {}); put(ck, r, 3, F(`IF(${ref(12)}="","",${ref(12)})`), {});
    put(ck, r, 4, F(`IF(OR(B${r}="",C${r}=""),"",CONCATENATE(B${r}," ",C${r}))`), {});
    put(ck, r, 5, s.b, {}); put(ck, r, 6, s.room, {});
    put(ck, r, 7, F(`IF(J${r}="","",IF(COUNTIF($D$2:$D$${NL},J${r})+COUNTIF($J$2:$J$${NL},J${r})>1,"❌","✅"))`), {});
    put(ck, r, 8, F(`IF(${ref(13)}="","",${ref(13)})`), {}); put(ck, r, 9, F(`IF(${ref(14)}="","",${ref(14)})`), {});
    put(ck, r, 10, F(`IF(OR(H${r}="",I${r}=""),"",CONCATENATE(H${r}," ",I${r}))`), {});
    put(ck, r, 11, s.b, {}); put(ck, r, 12, s.room, {});
  });
  redCF(ck, `A2:A${NL}`); redCF(ck, `G2:G${NL}`);

  // ================= 5. ПРОЖИВАНИЕ (даты заезда/выезда) =================
  const pr = mk('Проживание');
  ['Блок', 'Комната', 'Имя', 'Фамилия', 'Таб. №', 'Компания', 'Департамент', 'Вахта', 'Заезд', 'Выезд', 'Осталось дней']
    .forEach((t, i) => put(pr, 1, i + 1, t, { b: true, fill: ADR_FILL.sum, bd: 'thin' }));
  [8, 10, 16, 20, 14, 28, 24, 8, 12, 12, 14].forEach((w, i) => (pr.widths[i + 1] = w));
  const rank = (b) => { const i = order.indexOf(b); return i < 0 ? 999 : i; };
  const ppl = DB.people.map((p) => ({ p, r: DB.rooms[p.key] })).filter((x) => x.r)
    .sort((x, y) => rank(x.r.b) - rank(y.r.b) || x.r.n - y.r.n || (x.p.sur || '').localeCompare(y.p.sur || ''));
  ppl.forEach(({ p, r }, i) => {
    [r.b, r.n, p.name, p.sur, p.emp, p.comp, p.dept, p.shift, p.in, p.out, p.out ? diff(p.out) : '']
      .forEach((v, k) => put(pr, i + 2, k + 1, v === '' || v === undefined ? null : v, {}));
  });
  pr.filter = `A1:K${ppl.length + 1}`;

  // порядок вкладок: сводка, блоки, затем служебные
  return sheets;
}

async function adrWrite(sheets) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Turan Camp';
  wb.calcProperties = { fullCalcOnLoad: true };
  const thin = { style: 'thin', color: { argb: 'FF000000' } }, med = { style: 'medium', color: { argb: 'FF000000' } };
  sheets.forEach((sh) => {
    const ws = wb.addWorksheet(sh.name);
    Object.entries(sh.widths).forEach(([c, w]) => (ws.getColumn(+c).width = w));
    Object.entries(sh.heights).forEach(([r, h]) => (ws.getRow(+r).height = h));
    Object.entries(sh.cells).forEach(([key, { val, s }]) => {
      const [r, c] = key.split(',').map(Number), cell = ws.getCell(r, c);
      if (val !== null && val !== undefined) cell.value = typeof val === 'object' ? { formula: val.f } : val;
      cell.font = { name: 'Calibri', size: s.sz || 12, bold: !!s.b };
      if (s.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + s.fill } };
      if (s.bd !== 'none') { const bd = s.bd == 'medium' ? med : thin; if (s.bd || s.fill) cell.border = { top: bd, left: bd, bottom: bd, right: bd }; }
      cell.alignment = { horizontal: s.al || 'left', vertical: 'middle', wrapText: !!s.wrap };
      if (s.nf) cell.numFmt = s.nf;
    });
    sh.merges.forEach(([r1, c1, r2, c2]) => ws.mergeCells(r1, c1, r2, c2));
    sh.cf.forEach(({ ref, text }) => ws.addConditionalFormatting({ ref, rules: [{ type: 'cellIs', operator: 'equal', formulae: ['"' + text + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF9999' } } } }] }));
    if (sh.filter) ws.autoFilter = sh.filter;
  });
  return wb.xlsx.writeBuffer();
}

async function exportADR() {
  if (typeof ExcelJS === 'undefined') return alert('Библиотека ExcelJS не загрузилась (нужен интернет). Обновите страницу и попробуйте снова.');
  if (!Object.keys(DB.rooms).length) return toast('Нет данных для экспорта');
  try {
    toast('Готовлю Excel…');
    const buf = await adrWrite(adrPlan());
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    a.download = today().split('-').reverse().join('_') + '_ADR_TURAN.xlsx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 3000);
    toast('Экспорт готов');
  } catch (err) {
    console.error(err);
    alert('Ошибка экспорта: ' + err.message);
  }
}
function exportDashboard() { return exportADR(); }
function report() { return exportADR(); }

// ───────── БЛОКИ ─────────
function board() {
  const M = $('#main'),
    q = S.q.trim().toLowerCase();
  if (!Object.keys(DB.rooms).length) {
    M.innerHTML = empty;
    return;
  }

  // Поиск (остается без изменений)
  if (q) {
    const L = DB.people
      .filter((p) =>
        (nm(p) + ' ' + p.emp + ' ' + p.key + ' ' + p.dept + ' ' + p.comp)
          .toLowerCase()
          .includes(q)
      )
      .slice(0, 200);
    M.innerHTML =
      `<h3>Найдено: ${L.length}</h3><table><tr><th>ФИО</th><th>Таб. №</th><th>Компания</th><th>Деп.</th><th>Блок-комн.</th><th>Вахта</th><th>Осталось</th></tr>` +
      L.map(
        (p) =>
          `<tr class="k" onclick="openRoom('${p.key}')"><td>${esc(
            nm(p)
          )}</td><td>${esc(p.emp)}</td><td>${esc(p.comp)}</td><td>${esc(
            p.dept
          )}</td><td>${p.key.replace('|', '-')}</td><td>${
            p.shift ? `<span class="sh ${p.shift}">${p.shift}</span>` : ''
          }</td><td>${chip(p)}</td></tr>`
      ).join('') +
      '</table>';
    return;
  }

  const I = idx(DB.people),
    A = idx(DB.assign),
    rs = Object.values(DB.rooms)
      .filter((r) => r.b == S.blk)
      .sort((a, b) => a.n - b.n);

  // Верхняя панель фильтров (убрал кнопки этажей, так как они теперь всегда на экране)
  let h = `<div class="bar">
  <input list="dl-board" value="${esc(S.dept || '')}" placeholder="Департамент (поиск)…" onchange="S.dept=this.value.trim();board()" style="width:230px"><datalist id="dl-board">${depts().map((d) => `<option value="${esc(d)}">`).join('')}</datalist>
 <label class="mu"><input type="checkbox" ${
   S.free ? 'checked' : ''
 } onchange="S.free=this.checked;board()"> только со свободными местами</label></div>`;

  // 3 колонки для этажей
// колонки только для тех этажей, которые есть в блоке
const fls = [...new Set(rs.map((r) => r.f))].sort((a, b) => a - b);
h += `<div class="floors-row" style="--cols:${fls.length || 1}">`;

fls.forEach((f) => {
    // Фильтруем комнаты для текущего этажа
    const L = rs.filter(
      (r) =>
        r.f == f &&
        (!S.dept ||
          (r.dept || '').toLowerCase().includes(S.dept.toLowerCase()) ||
          (I[r.key] || []).some((p) =>
            (p.dept || '').toLowerCase().includes(S.dept.toLowerCase())
          )) &&
        (!S.free || (I[r.key] || []).length < 2)
    );

    h += `<div class="floor-column">`;
    h += `<h3 style="margin-top:0">${f} этаж · ${L.length} комнат</h3>`;
    h += `<div style="display:grid;grid-template-columns:1fr;gap:8px;">`;

    h += L.map((r) => {
      const res = I[r.key] || [],
        n = res.length;
      const beds = Array.from({ length: Math.max(2, res.length) }, (_, i) => i)
        .map((i) => {
          const p = res[i];
          return p
            ? `<div class="bed"><span>${esc(nm(p))}${
                p.shift ? ` <span class="sh ${p.shift}">${p.shift}</span>` : ''
              }</span>${chip(p)}</div>`
            : `<div class="bed fr">— свободно</div>`;
        })
        .join('');
      return `<div class="rc ${
        n == 0 ? '' : n == 1 ? 'h' : 'f'
      }" onclick="openRoom('${r.key}')">
       <div class="t"><span>${r.n}${(A[r.key] || []).length ? ' 🔗' : ''}${
        bk().some((b) => b.key == r.key) ? ' 📅' : ''
      }</span><small>${esc(r.dept)}</small></div>
       ${beds}
     </div>`;
    }).join('');

    h += `</div>`; // конец .fl
    if (!L.length) h += `<div class="mu" style="padding:10px">Нет комнат</div>`;
    h += `</div>`; // конец .floor-column
  });

  h += `</div>`; // конец .floors-row
  M.innerHTML = h;
}

// ───────── КАРТОЧКА КОМНАТЫ ─────────
// Модалка показывает ОДИН раздел за раз (вкладки), а не всё сразу.
// Полные списки закреплений / броней по-прежнему доступны на страницах «Закрепления» и «Брони».
const MTABS = [
  ['res', 'Жильцы'],
  ['in', 'Заселить'],
  ['book', 'Брони'],
  ['assign', 'Закрепления'],
];
function setMT(t) {
  S.mt = t;
  if (S.open) openRoom(S.open);
}
function roomTabRes(r, key, res) {
  if (!res.length)
    return `<div class="mu" style="padding:14px 0">Комната свободна.</div>
      <button class="b p" onclick="setMT('in')">Заселить →</button>`;
  return res
    .map(
      (p) => `<div class="card"><div class="row"><b>${esc(nm(p))}</b><button class="b s" style="padding:1px 6px;font-size:11px" onclick="editPerson(${p.id})" title="Редактировать сотрудника">✎</button>${
        p.shift ? `<span class="sh ${p.shift}">Вахта ${p.shift}</span>` : ''
      }<span style="margin-left:auto">${chip(p)}</span></div>
  <div class="mu">Таб. № ${esc(p.emp) || '—'} · ${esc(p.comp) || '—'} · ${esc(p.dept) || '—'} · заезд ${fd(p.in)}</div>
  <div class="row"><span class="mu">Выезд:</span><input type="date" value="${p.out}" onchange="setOut(${p.id},this.value)">
  <span class="mu">Овертайм:</span>${[1, 3, 7, 15]
    .map((n) => `<button class="b s" onclick="ext(${p.id},${n})">+${n}</button>`)
    .join('')}
  <button class="b s d" style="margin-left:auto" onclick="checkout(${p.id})">Выселить</button></div></div>`
    )
    .join('');
}
function roomTabIn(r, key, res) {
  if (res.length >= 2) return '<div class="mu" style="padding:14px 0">Мест нет — сначала выселите жильца.</div>';
  return `<div class="card">
 <div class="row"><input id="fe" placeholder="Таб. №" size="8" onkeydown="lookupEmp(event)"><span class="mu" style="font-size:10px;margin-right:auto">(Enter для поиска)</span></div>
 <div class="row"><input id="fn" placeholder="Имя"><input id="fu" placeholder="Фамилия"></div>
 <div class="row"><input id="fc" placeholder="Компания"><input id="fd" list="dl" placeholder="Департамент" value="${esc(r.dept)}"><select id="fs">${opts(['A', 'B'], '', 'Вахта?')}</select></div>
 <div class="row"><span class="mu">Выезд:</span><input id="fo" type="date" value="${addDays(today(), 15)}">${[15, 16]
    .map((n) => `<button class="b s" onclick="$('#fo').value=addDays(today(),${n})">+${n}</button>`)
    .join('')}<button class="b p" style="margin-left:auto" onclick="doIn('${key}')">Заселить</button></div></div>`;
}
function roomTabBook(r, key, res) {
  const bs = bk().filter((b) => b.key == key).sort((a, c) => a.from.localeCompare(c.from));
  return (
    (bs
      .map(
        (b) => `<div class="row"><span>${esc(nm(b))}</span>${b.shift ? `<span class="sh ${b.shift}">${b.shift}</span>` : ''}<span class="mu">${fd(b.from)} → ${fd(b.to)}</span>${fromChip(b)}
 ${res.length < 2 ? `<button class="b s p" onclick="checkInBook(${b.id})">Заселить</button>` : ''}<button class="b s d" style="margin-left:auto" onclick="delBook(${b.id})">✕</button></div>`
      )
      .join('') || '<div class="mu">Броней нет</div>') +
    `<div class="card"><div class="row"><input id="bn" placeholder="Имя"><input id="bu" placeholder="Фамилия"><input id="be" placeholder="Таб. №" size="8" onkeydown="if(event.key==='Enter')fillBook()"></div>
 <div class="row"><input id="bc" placeholder="Компания"><input id="bd" list="dl" placeholder="Департамент" value="${esc(r.dept)}"><select id="bs">${opts(['A', 'B'], '', 'Вахта?')}</select></div>
 <div class="row"><span class="mu">С:</span><input id="bf" type="date" value="${today()}"><span class="mu">По:</span><input id="bt" type="date" value="${addDays(today(), 15)}"><button class="b p" style="margin-left:auto" onclick="addBook('${key}')">Забронировать</button></div></div>`
  );
}
function roomTabAssign(r, key, res, as) {
  return (
    (as
      .map((a) => {
        const here = res.some((p) => nm(p).toLowerCase() == nm(a).toLowerCase());
        return `<div class="row"><span class="sh ${a.shift}">${a.shift}</span><span>${esc(nm(a))}${a.emp ? ' <span class="mu">№ ' + esc(a.emp) + '</span>' : ''}</span>${
          here
            ? '<span class="c g">живёт</span>'
            : res.length < 2
            ? `<button class="b s p" onclick="fillIn('${esc(a.name).replace(/'/g, "\\'")}','${esc(a.sur).replace(/'/g, "\\'")}','${a.shift}')">Заселить</button>`
            : ''
        }<button class="b s d" style="margin-left:auto" onclick="delA(${a.id})">✕</button></div>`;
      })
      .join('') || '<div class="mu">Нет закреплений за людьми</div>') +
    `<div class="row"><select id="as">${opts(['A', 'B'])}</select><input id="an" placeholder="Имя"><input id="au" placeholder="Фамилия"><button class="b s" onclick="addA('${key}')">+ Закрепить человека</button></div>`
  );
}
function openRoom(key, tab) {
  const r = DB.rooms[key];
  if (!r) return toast('Комната не найдена. Возможно, данные обновились.');
  // новая комната → всегда начинаем с вкладки «Жильцы»; обновление той же комнаты → остаёмся на своей вкладке
  if (tab) S.mt = tab;
  else if (S.open !== key || !S.mt) S.mt = 'res';
  S.open = key;
  const res = DB.people.filter((p) => p.key == key),
    as = DB.assign.filter((a) => a.key == key),
    nb = bk().filter((b) => b.key == key).length;
  const cnt = { res: `${res.length}/2`, in: '', book: nb || '', assign: as.length || '' };
  const body =
    S.mt == 'in' ? roomTabIn(r, key, res)
    : S.mt == 'book' ? roomTabBook(r, key, res)
    : S.mt == 'assign' ? roomTabAssign(r, key, res, as)
    : roomTabRes(r, key, res);
  const h = `<div class="row" style="margin-top:0"><h2>Блок ${r.b} · комната ${r.n}</h2><button class="b s" style="margin-left:auto" onclick="closeM()">✕</button></div>
 <div class="row"><span class="mu">Департамент:</span><input id="rd" list="dl" value="${esc(r.dept)}" style="flex:1;min-width:120px"><button class="b s" onclick="setDept('${key}')">Закрепить</button></div>
 <datalist id="dl">${depts().map((d) => `<option value="${esc(d)}">`).join('')}</datalist>
 <div class="mtabs">${MTABS.map(
   ([t, n]) => `<button class="${S.mt == t ? 'on' : ''}" onclick="setMT('${t}')">${n}${cnt[t] !== '' ? ` <small>${cnt[t]}</small>` : ''}</button>`
 ).join('')}</div>
 <div class="mbody">${body}</div>`;
  $('#md').innerHTML = h;
  $('#ov').classList.add('on');
}

function lookupEmp(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    const emp = e.target.value.trim();
    if (!emp) return;
    const ne = (x) => String(x || '').trim().replace(/^0+(?=\d)/, '').toLowerCase();
    const found = DB.people.find((p) => ne(p.emp) === ne(emp));
    if (found) {
      $('#fn').value = found.name || '';
      $('#fu').value = found.sur || '';
      $('#fc').value = found.comp || '';
      $('#fd').value = found.dept || '';
      $('#fs').value = found.shift || '';
      toast('Сотрудник найден!');
    } else {
      toast('Сотрудник не найден');
    }
  }
}

function closeM() {
  S.open = null;
  S.mt = 'res';
  $('#ov').classList.remove('on');
}
function refresh() {
  save();
  go(S.page);
  if (S.open) openRoom(S.open);
}
function setDept(k) {
  DB.rooms[k].dept = $('#rd').value.trim();
  log('Закрепление', `${k} → ${DB.rooms[k].dept}`);
  refresh();
}
function setOut(id, v) {
  DB.people.find((p) => p.id == id).out = v;
  refresh();
}
function ext(id, n) {
  const p = DB.people.find((p) => p.id == id);
  p.out = addDays(p.out && diff(p.out) >= 0 ? p.out : today(), n);
  log('Овертайм', `${nm(p)} +${n}д до ${p.out}`);
  refresh();
  toast(
    bk().some((b) => b.key == p.key && bOver(b, today(), p.out))
      ? '⚠ Продление пересекается с бронью!'
      : 'Продлено до ' + fd(p.out)
  );
}
function checkout(id) {
  const p = DB.people.find((p) => p.id == id);
  if (!confirm('Выселить ' + nm(p) + '?')) return;
  DB.people = DB.people.filter((x) => x.id != id);
  log('Выселение', `${nm(p)} из ${p.key}`);
  refresh();
}
function fillIn(n, u, s) {
  S.mt = 'in';
  if (S.open) openRoom(S.open);
  $('#fn').value = n;
  $('#fu').value = u;
  $('#fs').value = s;
  $('#fe').focus();
}
function doIn(key) {
  const p = {
    id: newId(),
    key,
    name: $('#fn').value.trim(),
    sur: $('#fu').value.trim(),
    emp: $('#fe').value.trim(),
    comp: $('#fc').value.trim(),
    dept: $('#fd').value.trim(),
    shift: $('#fs').value,
    in: today(),
    out: $('#fo').value,
  };
  if (!p.name && !p.sur) return toast('Укажите имя');
  if (DB.people.filter((x) => x.key == key).length >= 2) return toast('Мест нет — сначала выселите жильца');
  DB.people.push(p);
  log('Заселение', `${nm(p)} → ${key}, до ${p.out}`);
  S.mt = 'res';
  refresh();
  toast('Заселён');
}
function addA(key) {
  const a = {
    id: newId(),
    key,
    shift: $('#as').value,
    name: $('#an').value.trim(),
    sur: $('#au').value.trim(),
  };
  if (!a.name && !a.sur) return;
  DB.assign.push(a);
  log('Закрепление', `${nm(a)} → ${key} (${a.shift})`);
  refresh();
}
function delA(id) {
  DB.assign = DB.assign.filter((a) => a.id != id);
  refresh();
}

// 

// ───────── ВСЕ СОТРУДНИКИ ─────────
function pEmployees(){
  $('#main').innerHTML=`<div class="bar"><input id="eq" placeholder="ФИО, таб. №, компания…" oninput="listEmp()" style="width:280px">
  <select id="eb" onchange="listEmp()">${opts(BL(),'','Все блоки')}</select>
  <select id="ed" onchange="listEmp()">${opts(depts(),'','Все департаменты')}</select>
  <select id="es" onchange="listEmp()">${opts(['A','B'],'','Любая вахта')}</select></div><div id="res"></div>`;
  listEmp()}
function listEmp(){
  const q=$('#eq').value.toLowerCase(),b=$('#eb').value,d=$('#ed').value,s=$('#es').value;
  const L=DB.people.filter(p=>(!b||p.key.split('|')[0]==b)&&(!d||p.dept==d)&&(!s||p.shift==s)&&(!q||(nm(p)+' '+p.emp+' '+p.comp).toLowerCase().includes(q)))
    .sort((a,c)=>(a.sur||'').localeCompare(c.sur||''));
  $('#res').innerHTML=`<h3>Найдено: ${L.length}</h3><table><tr><th>ФИО</th><th>Таб. №</th><th>Компания</th><th>Департамент</th><th>Блок-комн.</th><th>Вахта</th><th>Осталось</th></tr>`+
  L.slice(0,500).map(p=>`<tr class="k" onclick="openRoom('${p.key}')"><td>${esc(nm(p))}</td><td>${esc(p.emp)}</td><td>${esc(p.comp)}</td><td>${esc(p.dept)}</td><td>${p.key.replace('|','-')}</td><td>${p.shift?`<span class="sh ${p.shift}">${p.shift}</span>`:''}</td><td>${chip(p)}</td></tr>`).join('')+'</table>'+(L.length>500?'<div class="mu">Показаны первые 500, уточните фильтр</div>':'')}

  // ───────── ДЕПАРТАМЕНТЫ ─────────
function pDepts() {
  const ds = depts(), R = {}, P = {};
  Object.values(DB.rooms).forEach((r) => { if (r.dept) R[r.dept] = (R[r.dept] || 0) + 1; });
  DB.people.forEach((p) => { if (p.dept) P[p.dept] = (P[p.dept] || 0) + 1; });
  $('#main').innerHTML =
    `<div class="bar"><input id="nd" placeholder="Новый департамент" style="width:260px" onkeydown="if(event.key==='Enter')addDept()">
    <button class="b p" onclick="addDept()">+ Добавить</button><span class="mu">Всего: <b>${ds.length}</b></span></div>
    <table><tr><th>Департамент</th><th>Комнат</th><th>Людей</th><th></th></tr>` +
    ds.map((d, i) => `<tr><td><b>${esc(d)}</b></td><td>${R[d] || 0}</td><td>${P[d] || 0}</td>
      <td><button class="b s" onclick="renameDept(${i})">✎</button> <button class="b s d" onclick="delDept(${i})">✕</button></td></tr>`).join('') + '</table>';
}
function addDept() {
  const v = $('#nd').value.trim();
  if (!v) return;
  if (depts().some((d) => d.toLowerCase() === v.toLowerCase())) return toast('Такой департамент уже есть');
  DB.depts.push(v); log('Департамент', 'добавлен: ' + v); save(); toast('Добавлен'); pDepts();
}
function delDept(i) {
  const d = depts()[i];
  const rc = Object.values(DB.rooms).filter((r) => r.dept === d).length;
  const pc = DB.people.filter((p) => p.dept === d).length;
  if (!confirm(`Удалить «${d}»?` + (rc || pc ? `\nОн указан у комнат: ${rc}, людей: ${pc}. У них департамент будет очищен.` : ''))) return;
  DB.depts = DB.depts.filter((x) => x !== d);
  Object.values(DB.rooms).forEach((r) => { if (r.dept === d) r.dept = ''; });
  DB.people.forEach((p) => { if (p.dept === d) p.dept = ''; });
  log('Департамент', 'удалён: ' + d); save(); go('depts');
}
function renameDept(i) {
  const d = depts()[i], v = (prompt('Новое название:', d) || '').trim();
  if (!v || v === d) return;
  if (depts().some((x) => x.toLowerCase() === v.toLowerCase())) return toast('Такой департамент уже есть');
  DB.depts = DB.depts.map((x) => (x === d ? v : x));
  Object.values(DB.rooms).forEach((r) => { if (r.dept === d) r.dept = v; });
  DB.people.forEach((p) => { if (p.dept === d) p.dept = v; });
  log('Департамент', `${d} → ${v}`); save(); go('depts');
}

// ───────── БРОНИ ─────────
const bk = () => DB.book || (DB.book = []);
const bOver = (b, f, t) => b.from <= t && b.to >= f;
function freeAt(key, f, t) {
  // свободных мест в комнате на период [f,t]
  const res = DB.people.filter(
    (p) => p.key == key && (!p.out || p.out >= f)
  ).length;
  const bs = bk().filter((b) => b.key == key && bOver(b, f, t)).length;
  return 2 - res - bs;
}
function fromChip(b) {
  const d = diff(b.from);
  return d < 0
    ? `<span class="c r">опоздал ${-d}д</span>`
    : d == 0
    ? '<span class="c y">заезд сегодня</span>'
    : `<span class="c g">через ${d} дн.</span>`;
}
function addBook(key) {
  const b = {
    id: newId(),
    key,
    name: $('#bn').value.trim(),
    sur: $('#bu').value.trim(),
    emp: $('#be').value.trim(),
    comp: $('#bc').value.trim(),
    dept: $('#bd').value.trim(),
    shift: $('#bs').value,
    from: $('#bf').value,
    to: $('#bt').value,
  };
  if (!b.name && !b.sur) return toast('Укажите имя');
  if (!b.from || !b.to || b.to < b.from) return toast('Проверьте даты');
  if (freeAt(key, b.from, b.to) < 1) return toast('На эти даты мест нет');
  bk().push(b);
  log('Бронь', `${nm(b)} → ${key}, ${b.from} — ${b.to}`);
  refresh();
  toast('Бронь создана');
}
function delBook(id) {
  DB.book = bk().filter((b) => b.id != id);
  refresh();
}
function fillBook(){
  const e=$('#be').value.trim().toLowerCase();if(!e)return;
  const p=DB.people.find(x=>String(x.emp).trim().toLowerCase()==e)||bk().find(x=>String(x.emp).trim().toLowerCase()==e);
  if(!p)return toast('Табельный № не найден');
  $('#bn').value=p.name;$('#bu').value=p.sur;$('#bc').value=p.comp||'';
  if(p.dept)$('#bd').value=p.dept;
  if(p.shift)$('#bs').value=p.shift;
  toast('Данные подставлены: '+nm(p));$('#bf').focus()}

function checkInBook(id) {
  const b = bk().find((x) => x.id == id);
  if (DB.people.filter((p) => p.key == b.key).length >= 2)
    return toast('Сейчас нет места — сначала выселите соседа');
  DB.people.push({
    id: newId(),
    key: b.key,
    name: b.name,
    sur: b.sur,
    emp: b.emp,
    comp: b.comp,
    dept: b.dept,
    shift: b.shift,
    in: today(),
    out: b.to,
  });
  DB.book = bk().filter((x) => x.id != id);
  log('Заселение по брони', `${nm(b)} → ${b.key}, до ${b.to}`);
  S.mt = 'res';
  refresh();
  toast('Заселён');
}
function pBook() {
  const L = [...bk()].sort((a, b) => a.from.localeCompare(b.from));
  $('#main').innerHTML =
    `<div class="mu">Все брони по дате заезда.</div><table><tr><th>ФИО</th><th>Блок-комн.</th><th>Вахта</th><th>Период</th><th>Заезд</th><th></th></tr>` +
    L.map(
      (b) => `<tr class="k" onclick="openRoom('${b.key}')"><td>${esc(
        nm(b)
      )}</td><td>${b.key.replace('|', '-')}</td><td>${
        b.shift ? `<span class="sh ${b.shift}">${b.shift}</span>` : ''
      }</td><td>${fd(b.from)} → ${fd(b.to)}</td><td>${fromChip(b)}</td>
  <td><button class="b s p" onclick="event.stopPropagation();checkInBook(${
    b.id
  })">Заселить</button></td></tr>`
    ).join('') +
    '</table>' +
    (L.length ? '' : '<div class="mu" style="padding:20px">Броней нет</div>');
}

// ───────── ЗАКРЕПЛЕНИЯ ─────────
function pAssign() {
  $(
    '#main'
  ).innerHTML = `<div class="card"><div class="row"><b>Закрепить комнату за департаментом</b><select id="gb" onchange="">${opts(
    BL()
  )}</select><input id="gr" placeholder="Комната (напр. 101)" size="10"><input id="gd" list="dl2" placeholder="Департамент"><datalist id="dl2">${depts()
    .map((d) => `<option value="${esc(d)}">`)
    .join(
      ''
    )}</datalist><button class="b p" onclick="gDept()">Закрепить</button></div>
 <div class="row"><b>Закрепить за человеком</b><select id="pb">${opts(
   BL()
 )}</select><input id="pr" placeholder="Комната" size="10"><select id="ps">${opts(
    ['A', 'B']
  )}</select><input id="pn" placeholder="Имя"><input id="pu" placeholder="Фамилия"><button class="b p" onclick="gPers()">Закрепить</button></div></div>
 <div class="bar"><input id="aq" placeholder="Фильтр: имя, комната, вахта…" oninput="listA()" style="width:280px"></div><div id="res"></div>`;
  listA();
}
function gDept() {
  const k = $('#gb').value + '|' + $('#gr').value.trim();
  if (!DB.rooms[k]) return toast('Комната не найдена');
  DB.rooms[k].dept = $('#gd').value.trim();
  log('Закрепление', `${k} → ${DB.rooms[k].dept}`);
  save();
  toast('Готово');
}
function gPers() {
  const k = $('#pb').value + '|' + $('#pr').value.trim();
  if (!DB.rooms[k]) return toast('Комната не найдена');
  const a = {
    id: newId(),
    key: k,
    shift: $('#ps').value,
    name: $('#pn').value.trim(),
    sur: $('#pu').value.trim(),
  };
  if (!a.name && !a.sur) return;
  DB.assign.push(a);
  log('Закрепление', `${nm(a)} → ${k} (${a.shift})`);
  save();
  toast('Готово');
  listA();
}
function listA() {
  const q = ($('#aq')?.value || '').toLowerCase(),
    cnt = {};
  DB.assign.forEach((a) => {
    const c = a.shift + nm(a).toLowerCase();
    cnt[c] = (cnt[c] || 0) + 1;
  });
  const I = idx(DB.people),
    L = DB.assign.filter((a) =>
      (nm(a) + ' ' + a.key + ' ' + a.shift + ' ' + (a.emp || '')).toLowerCase().includes(q)
    );
  $('#res').innerHTML =
    `<h3>Закреплений: ${L.length}${
      Object.values(cnt).some((c) => c > 1)
        ? ' · <span class="w">⚠ есть дубли</span>'
        : ''
    }</h3><table><tr><th>Вахта</th><th>ФИО</th><th>Блок-комн.</th><th>Деп. комнаты</th><th></th><th></th></tr>` +
    L.slice(0, 300)
      .map(
        (a) => `<tr class="k" onclick="openRoom('${
          a.key
        }')"><td><span class="sh ${a.shift}">${a.shift}</span></td><td>${esc(
          nm(a)
        )}</td><td>${a.key.replace('|', '-')}</td><td>${esc(
          DB.rooms[a.key]?.dept
        )}</td>
 <td>${
   (I[a.key] || []).some((p) => nm(p).toLowerCase() == nm(a).toLowerCase())
     ? '<span class="c g">живёт</span>'
     : ''
 }${
          cnt[a.shift + nm(a).toLowerCase()] > 1
            ? ' <span class="c y">⚠ дубль</span>'
            : ''
        }</td><td><button class="b s d" onclick="event.stopPropagation();delA(${
          a.id
        })">✕</button></td></tr>`
      )
      .join('') +
    '</table>';
}

// ───────── ВЫЕЗДЫ ─────────
function pOut() {
  const L = [...DB.people].sort((a, b) =>
    (a.out || '9').localeCompare(b.out || '9')
  );
  $('#main').innerHTML =
    `<div class="mu">Все проживающие по дате выезда.</div><table><tr><th>ФИО</th><th>Блок-комн.</th><th>Вахта</th><th>Выезд</th><th>Осталось</th><th>Овертайм</th></tr>` +
    L.slice(0, 500)
      .map(
        (p) => `<tr><td class="k" onclick="openRoom('${p.key}')">${esc(
          nm(p)
        )}</td><td>${p.key.replace('|', '-')}</td><td>${
          p.shift ? `<span class="sh ${p.shift}">${p.shift}</span>` : ''
        }</td><td><input type="date" value="${p.out}" onchange="setOut(${
          p.id
        },this.value)"></td><td>${chip(p)}</td>
 <td>${[1, 3, 7]
   .map((n) => `<button class="b s" onclick="ext(${p.id},${n})">+${n}</button>`)
   .join(' ')}</td></tr>`
      )
      .join('') +
    '</table>';
}

// ───────── ДАННЫЕ ─────────
function pData() {
  $('#main').innerHTML = `
  <div class="card">
    <div class="row"><b>Импорт из Excel (ADR)</b><input type="file" accept=".xlsx" onchange="importXlsx(this.files[0]);this.value=''"></div>
    <div class="mu">Читает листы A1…C4: комнаты, проживающих, «Вахта А / Вахта В», департаменты.</div>
  </div>
  
  <div class="card">
    <div class="row">
      <button class="b p" onclick="report()">📊 Отчёт в Excel (формат ADR)</button>
      <button class="b" onclick="backup()">⬇ Резервная копия (JSON)</button>
      <span class="mu">Восстановить:</span><input type="file" accept=".json" onchange="restore(this.files[0]);this.value=''">
    </div>
  </div>
  
  <!-- НОВАЯ КАРТОЧКА ДЛЯ JSON WFM -->
  <div class="card" style="margin-top: 16px; border-color: var(--gr); border-left-width: 4px;">
    <div class="row">
      <b style="color: var(--gr)">🔥 Импорт JSON (новый WFM)</b>
      <input type="file" accept=".json" onchange="importWfmJson(this.files[0]);this.value=''">
    </div>
    <div class="mu">Загрузите JSON-файл, скопированный из DevTools (Network -> Copy response).</div>
  </div>
  <!-- КОНЕЦ НОВОЙ КАРТОЧКИ -->

  <h3>История</h3>
  ${DB.hist.slice(0, 50).map((h) => `<div class="mu">${h.t} · <b>${esc(h.a)}</b> · ${esc(h.d)}</div>`).join('') || '<div class="mu">Пусто</div>'}
  `;
}
function backup() {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(
    new Blob([JSON.stringify(DB)], { type: 'application/json' })
  );
  a.download = 'turan_backup_' + today() + '.json';
  a.click();
}
function restore(f) {
  const r = new FileReader();
  r.onload = (e) => {
    try {
      DB = Object.assign(
        { rooms: {}, people: [], assign: [], hist: [], nid: 1, book: [] },
        JSON.parse(e.target.result)
      );
      save();
      go('board');
      toast('Восстановлено');
    } catch (x) {
      alert('Неверный файл');
    }
  };
  r.readAsText(f);
}

// ───────── ПОИСК МЕСТ / ОВЕРТАЙМ ─────────
function pSearch() {
  S.qOt = S.qOt || '';

  // 1. Блок поиска конкретного человека для овертайма
  let h = `<div class="card" style="margin-bottom: 20px; border-color: var(--yl);">
    <div class="row">
      <b style="color: var(--yl)">⏳ Оформить овертайм сотруднику:</b>
      <input type="text" placeholder="Найти по ФИО или таб. №" value="${esc(
        S.qOt
      )}" oninput="S.qOt=this.value; runSearch()" style="width:250px">
    </div>
    <div id="ot-person-result"></div>
  </div>`;

  // 2. Блок поиска свободных комнат (как было)
  h += `<div class="bar"><span class="mu">Нужно свободное место с даты:</span><input type="date" id="sf" value="${today()}" onchange="runSearch()">
    <select id="sd" onchange="runSearch()">${opts(
      depts(),
      '',
      'Любой департамент'
    )}</select><select id="sb" onchange="runSearch()">${opts(
    BL(),
    '',
    'Все блоки'
  )}</select>
    <label class="mu"><input type="checkbox" id="sw" onchange="runSearch()"> только пустые комнаты</label>
  </div>
  <div class="mu">Показаны комнаты, где к выбранной дате есть свободное место.</div><div id="res"></div>`;

  $('#main').innerHTML = h;
  runSearch();
}

function runSearch() {
  // Отрисовка поиска человека
  const otQ = (S.qOt || '').toLowerCase();
  if (otQ) {
    const people = DB.people.filter((p) =>
      (nm(p) + ' ' + p.emp).toLowerCase().includes(otQ)
    );
    if (people.length > 0) {
      $('#ot-person-result').innerHTML = people
        .map(
          (p) => `
        <div class="row" style="border-top:1px solid var(--bd); padding-top:5px; margin-top:5px;">
          <b>${esc(nm(p))}</b> <span class="mu">(Таб. ${esc(p.emp)})</span>
          <span>Блок ${p.key.replace('|', '-')}</span>
          <span class="mu">До: ${fd(p.out)}</span>
          ${[1, 3, 7, 14]
            .map(
              (n) =>
                `<button class="b s p" onclick="ext(${p.id},${n});runSearch()">+${n} дн.</button>`
            )
            .join(' ')}
        </div>
      `
        )
        .join('');
    } else {
      $(
        '#ot-person-result'
      ).innerHTML = `<div class="mu">Сотрудник не найден</div>`;
    }
  } else {
    $(
      '#ot-person-result'
    ).innerHTML = `<div class="mu">Введите имя или табельный номер, чтобы продлить пребывание (овертайм).</div>`;
  }

  // Отрисовка свободных комнат (как было)
  const f = $('#sf').value || today(),
    d = $('#sd').value,
    b = $('#sb').value,
    full = $('#sw').checked,
    I = idx(DB.people),
    A = idx(DB.assign),
    L = [];
  Object.values(DB.rooms).forEach((r) => {
    if (b && r.b != b) return;
    const res = I[r.key] || [],
      busy = res.filter((p) => !p.out || p.out >= f),
      fr = 2 - busy.length - bk().filter((x) => x.key === r.key && x.from <= f && x.to >= f).length;
    if (fr < (full ? 2 : 1)) return;
    L.push({
      r,
      fr,
      res,
      as: A[r.key] || [],
      m: d && r.dept == d ? 0 : d && !r.dept ? 1 : 2,
    });
  });
  L.sort(
    (x, y) =>
      x.m - y.m ||
      y.fr - x.fr ||
      x.r.b.localeCompare(y.r.b, 0, { numeric: true }) ||
      x.r.n - y.r.n
  );

  $('#res').innerHTML =
    `<h3>Свободных комнат: ${L.length}</h3><table><tr><th>Блок-комн.</th><th>Мест</th><th>Департамент</th><th>Сосед</th></tr>` +
    L.slice(0, 150)
      .map(
        (x) => `
    <tr class="k" onclick="openRoom('${x.r.key}')">
      <td><b>${x.r.b}-${x.r.n}</b></td>
      <td><span class="c ${x.fr == 2 ? 'g' : 'y'}">${x.fr}</span></td>
      <td>${x.m == 0 ? '<b style="color:var(--gr)">' : '<span>'}${
          esc(x.r.dept) || '—'
        }${x.m == 0 ? '</b>' : '</span>'}</td>
      <td>${
        x.res
          .map(
            (p) =>
              esc(nm(p)) +
              (p.out ? ` <span class="mu">(до ${fd(p.out)})</span>` : '')
          )
          .join('<br>') || '—'
      }</td>
    </tr>
  `
      )
      .join('') +
    '</table>';
}
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') closeM();
});

['nav', '#tabs'].forEach((sel) => {
  const el = $(sel);
  if (!el) return;
  el.addEventListener(
    'wheel',
    (e) => {
      if (el.scrollWidth <= el.clientWidth) return;
      e.preventDefault();
      el.scrollLeft += e.deltaY || e.deltaX;
    },
    { passive: false }
  );
});

window.addEventListener('storage', (e) => {
  if (e.key === 'turan3') { Object.assign(DB, JSON.parse(e.newValue)); go(S.page); }
});

// ───────── РЕДАКТИРОВАНИЕ СОТРУДНИКА ─────────
function editPerson(id) {
  const p = DB.people.find((x) => x.id == id);
  if (!p) return;

  let h = `<div class="row"><h2>Редактировать: ${esc(nm(p))}</h2><button class="b s" style="margin-left:auto" onclick="closeM()">✕</button></div>
  <div class="row">
    <input id="ep_name" placeholder="Имя" value="${esc(p.name)}">
    <input id="ep_sur" placeholder="Фамилия" value="${esc(p.sur)}">
  </div>
  <div class="row">
    <input id="ep_emp" placeholder="Таб. №" value="${esc(p.emp)}">
    <input id="ep_comp" placeholder="Компания" value="${esc(p.comp)}">
  </div>
  <div class="row">
    <input id="ep_dept" list="dl-edit" placeholder="Департамент" value="${esc(p.dept)}">
    <datalist id="dl-edit">${depts().map(d => `<option value="${esc(d)}">`).join('')}</datalist>
    <select id="ep_shift">${opts(['A','B'], p.shift, 'Вахта?')}</select>
  </div>
  <div class="row">
    <span class="mu">Заезд:</span><input type="date" id="ep_in" value="${p.in}">
    <span class="mu">Выезд:</span><input type="date" id="ep_out" value="${p.out}">
  </div>
  <div class="row" style="margin-top:15px">
    <button class="b p" style="width:100%; padding:10px" onclick="savePerson(${p.id})">Сохранить изменения</button>
  </div>`;
  $('#md').innerHTML = h;
  $('#ov').classList.add('on');
}

function savePerson(id) {
  const p = DB.people.find((x) => x.id == id);
  if (!p) return;

  p.name = $('#ep_name').value.trim();
  p.sur = $('#ep_sur').value.trim();
  p.emp = $('#ep_emp').value.trim();
  p.comp = $('#ep_comp').value.trim();
  p.dept = $('#ep_dept').value.trim();
  p.shift = $('#ep_shift').value;
  p.in = $('#ep_in').value;
  p.out = $('#ep_out').value;

  log('Редактирование', `${nm(p)} обновлен`);
  save();
  const k = p.key, wasOpen = !!S.open;
  closeM();
  toast('Данные сохранены');
  go(S.page);
  if (wasOpen) openRoom(k);
}
function importWfmJson(file) {
  if (!file) return;
  const reader = new FileReader();
  
  reader.onload = function(e) {
    try {
      let text = e.target.result;
      if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
      
      const json = JSON.parse(text);
      const rows = json.data || [];
      let updated = 0, notFound = 0;

      // Очистка табельного (убираем пробелы и ведущие нули)
      const normEmp = (e) => {
        if (!e) return '';
        let s = String(e).trim().replace(/\s+/g, '');
        if (/^\d+$/.test(s)) s = String(Number(s)); 
        return s.toLowerCase();
      };

      // Очистка имени: оставляем только буквы и цифры (убираем запятые, пробелы, дефисы)
      const cleanName = (s) => String(s || '').toLowerCase().replace(/[^a-zа-я0-9]/g, '');

      const parseDate = (d) => {
        if (!d) return '';
        const parts = String(d).split('/');
        return parts.length === 3 ? `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}` : '';
      };

      rows.forEach(row => {
        let person = null;
        const wfmEmp = normEmp(row.emp);

        // 1. Ищем по табельному номеру
        if (wfmEmp) {
          person = DB.people.find(p => normEmp(p.emp) === wfmEmp);
        }
        
        // 2. Если не нашли, ищем по ФИО (разбиваем по запятой)
        if (!person && row.fio) {
          const parts = row.fio.split(',');
          if (parts.length >= 2) {
            const wfmName1 = cleanName(parts[0]); // То, что до запятой
            const wfmName2 = cleanName(parts[1]); // То, что после запятой
            
            person = (wfmName1.length < 2 || wfmName2.length < 2) ? null : DB.people.find(p => {
              const dbSur = cleanName(p.sur);   // Фамилия в ADR
              const dbName = cleanName(p.name); // Имя в ADR
              
              // Проверяем оба варианта (Фамилия-Имя и Имя-Фамилия)
              return (dbSur.includes(wfmName1) && dbName.includes(wfmName2)) || 
                     (dbName.includes(wfmName1) && dbSur.includes(wfmName2));
            });
          }
        }

        if (person) {
          const newOut = parseDate(row.out);
          if (newOut && person.out !== newOut) {
            person.out = newOut;
          }

          if (row.emp) {
            // Обновляем компанию на основе табельного из WFM
            if (/\d/.test(String(row.emp))) {
              person.comp = 'KAL';
            } else if (row.comp) {
              person.comp = row.comp;
            }
            
            // Если в базе ADR табельного вообще не было — заполним его из WFM
            if (!person.emp) {
                person.emp = String(row.emp).trim();
            }
          }
          updated++;
        } else {
          notFound++;
          if (notFound <= 5) console.log('❌ WFM не найден:', row);
        }
      });

      save(); 
      toast(`WFM: Обновлено ${updated}, Не найдено ${notFound}`);
      go(S.page);
      
    } catch (err) {
      console.error("Ошибка:", err);
      alert(`Ошибка: ${err.message}`);
    }
  };
  
  reader.readAsText(file, 'UTF-8'); 
}

// Не забудь про эти две строчки в самом конце файла!
window.importWfmJson = importWfmJson;
// обновление экрана, когда из облака пришли чужие изменения (sync.js)
window.onSbData = () => {
  go(S.page);
  if (S.open && DB.rooms[S.open]) openRoom(S.open);
};
go('dashboard');
