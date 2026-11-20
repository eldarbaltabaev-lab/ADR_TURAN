// travel.js — импорт Travel-календаря (новая структура)
// ФИО: переводим в латиницу, отделяем имя от фамилии.
const TR_MAP = { а:'a',б:'b',в:'v',г:'g',д:'d',е:'e',ё:'yo',ж:'zh',з:'z',и:'i',й:'y',к:'k',л:'l',м:'m',н:'n',о:'o',п:'p',р:'r',с:'s',т:'t',у:'u',ф:'f',х:'kh',ц:'ts',ч:'ch',ш:'sh',щ:'shch',ъ:'',ы:'y',ь:'',э:'e',ю:'yu',я:'ya',ә:'a',ғ:'g',қ:'k',ң:'n',ө:'o',ұ:'u',ү:'u',һ:'h',і:'i' };
const TR_C2L = { а:'a',в:'b',е:'e',к:'k',м:'m',н:'h',о:'o',р:'p',с:'c',т:'t',х:'x',у:'y',і:'i' };
const TR_L2C = { a:'а',b:'в',e:'е',k:'к',m:'м',h:'н',o:'о',p:'р',c:'с',t:'т',x:'х',y:'у',i:'і' };
const TR_VOW = 'аеёиоуыэюяәөұүіьъ';
const TR_PART = ['оглы', 'угли', 'уғли', 'улы', 'кызы', 'қызы', 'кизи', 'гызы', 'оглу'];
const TR_CYR = /[а-яёәғқңөұүһі]/, TR_LAT = /[a-z]/;

function trFix(w) {
  w = w.toLowerCase();
  const ch = [...w];
  const lat = ch.filter((c) => TR_LAT.test(c)), cyr = ch.filter((c) => TR_CYR.test(c));
  const latOnly = lat.some((c) => !TR_L2C[c]), cyrOnly = cyr.some((c) => !TR_C2L[c]);
  const toLat = latOnly ? true : cyrOnly ? false : lat.length >= cyr.length;
  if (!lat.length || !cyr.length) return w;
  return ch.map((c) => (toLat ? TR_C2L[c] || c : TR_L2C[c] || c)).join('');
}

function trPart(p) {
  p = trFix(p);
  let o = '';
  for (let i = 0; i < p.length; i++) {
    const ch = p[i];
    if (ch === 'е') o += i === 0 || TR_VOW.includes(p[i - 1]) ? 'ye' : 'e';
    else o += TR_MAP[ch] ?? ch;
  }
  return o.charAt(0).toUpperCase() + o.slice(1);
}

const trWord = (w) => w.split('-').map(trPart).join('-');
const trJoin = (a) => a.map(trWord).join(' ');

// Парсим ФИО (в новом формате они в одной ячейке)
function parseFio(raw) {
  const s = String(raw || '').replace(/\(.*?\)/g, ' ').replace(/=/g, ' ').replace(/\s+/g, ' ').trim();
  const t = s.split(' ').filter((x) => x && !/^\d{6,}$/.test(x));
  if (!t.length) return { name: '', sur: '' };
  
  const f = t.map(trFix), cyr = f.map((w) => TR_CYR.test(w)), n = t.length;
  
  if (!cyr.some(Boolean)) {
    // Иностранцы (пишутся как Имя Фамилия)
    return { sur: trWord(t[0]), name: trJoin(t.slice(1)) };
  }
  if (n >= 4 && TR_PART.includes(f[1])) {
    // «Мамуч оглы Шариф Омарович»
    return { sur: trJoin([t[0], t[1]]), name: trWord(t[2]) };
  }
  if (n >= 4 && TR_PART.includes(f[n - 1])) {
    // «… Тулкун Уғли»
    return { sur: trWord(t[0]), name: trJoin(t.slice(1, n - 2)) };
  }
  if (n >= 3) {
    // Отбрасываем отчество (3-е слово)
    return { sur: trWord(t[0]), name: trJoin(t.slice(1, n - 1)) };
  }
  // Если просто 2 слова (Имя Фамилия или Фамилия Имя)
  return { sur: trWord(t[0]), name: trJoin(t.slice(1)) };
}

// Парсим компанию/табельный
function trCompany(c) {
  if (typeof c === 'number') return { emp: String(c), comp: 'KAL' };
  const s = String(c ?? '').trim();
  if (!s) return { emp: '', comp: '' };
  if (/^\d+$/.test(s)) return { emp: s, comp: 'KAL' }; 
  const o = trJoin(s.split(/\s+/));
  return { emp: '', comp: s === s.toUpperCase() ? o.toUpperCase() : o };
}

// Утилита для дат из Excel
const trIso = (n) => {
  if (typeof n === 'number') return new Date(Math.round((n - 25569) * 864e5)).toISOString().slice(0, 10);
  const s = String(n || '').trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);            // «2026-09-03 00:00:00» — без перевода во время UTC (иначе сдвиг на день)
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})/);   // «03.09.2026»
  return m ? m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0') : '';
};
const trHead = (v) => { const s = String(v || ''); return /дата/i.test(s) && /date/i.test(s); }; // строка-заголовок с датой

const trTime = (v) => {
  if (typeof v === 'number') { 
    const m = Math.round(v * 1440) % 1440; 
    return String(Math.floor(m / 60)).padStart(2, '0') + ':' + String(m % 60).padStart(2, '0'); 
  }
  return String(v || '').trim();
};

// Главная функция парсинга листа
function trParseSheet(ws, type, from, to) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const out = [];
  
  let currentDateOut = '';
  let currentDateIn = '';

  // Начинаем с 3-й строки (индекс 3)
  for (let i = 3; i < raw.length; i++) {
    const r = raw[i];
    let isHeaderRow = false; // Флаг: является ли эта строка заголовком с датой

    // --- ЛЕВАЯ ЧАСТЬ (ОТПРАВЛЕНИЯ) ---
    // Ищем строку с датой
    if (trHead(r[1])) {
      currentDateOut = trIso(r[2]);
      isHeaderRow = true; // Это заголовок, людей тут нет
    }
    
    // --- ПРАВАЯ ЧАСТЬ (ПРИБЫТИЯ) ---
    // Ищем строку с датой
    if (trHead(r[9])) {
      currentDateIn = trIso(r[10]);
      isHeaderRow = true;
    }

    // Если это была строка с датой — пропускаем её, чтобы не создавать фейкового пассажира
    if (isHeaderRow) continue;

    // Если в колонке B (индекс 1) есть имя (и это не машина/служебная строка)
    const nameLeft = String(r[1] || '').trim();
    if (nameLeft && !nameLeft.includes('Көлік') && !nameLeft.includes('Вахта') && !nameLeft.includes('Шығу') && !nameLeft.includes('Без пассажира')) {
      if (currentDateOut >= from && currentDateOut <= to) {
        const { name, sur } = parseFio(nameLeft);
        const { emp, comp } = trCompany(r[2]);
        out.push({ 
          type: 'out', // Отправление
          date: currentDateOut, 
          time: trTime(r[5]), 
          name, sur, emp, comp, 
          block: '', room: '' 
        });
      }
    }

    // Если в колонке J (индекс 9) есть имя
    const nameRight = String(r[9] || '').trim();
    if (nameRight && !nameRight.includes('Көлік') && !nameRight.includes('Вахта') && !nameRight.includes('Шығу') && !nameRight.includes('Без пассажира')) {
      if (currentDateIn >= from && currentDateIn <= to) {
        const { name, sur } = parseFio(nameRight);
        const { emp, comp } = trCompany(r[10]);
        out.push({ 
          type: 'in', // Прибытие
          date: currentDateIn, 
          time: trTime(r[13]), 
          name, sur, emp, comp, 
          block: '', room: '' 
        });
      }
    }
  }
  
  return out;
}

function importTravel(file) {
  const from = S.tf;
  const to = S.tt;
  if (!file) return;
  if (!from || !to || to < from) return toast('Проверьте период «с / по»');
  
  toast('Читаю файл…');
  const rd = new FileReader();
  
  rd.onload = (ev) => setTimeout(() => {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'array' });
      
      // Берем первый лист (Рассылка на SITE)
      const ws = wb.Sheets[wb.SheetNames[0]];
      
      // Парсим данные на весь диапазон
      const rows = trParseSheet(ws, 'all', from, to);
      
      let added = 0, skipped = 0;
      mut(() => {
        const k = (e) => [e.type, e.date, e.emp || '', full(e).toLowerCase()].join('|');
        const have = new Set(DB.enc.map(k)), cnt = {};
        DB.enc.forEach((e) => { const g = e.type + e.date; cnt[g] = Math.max(cnt[g] || 0, e.n || 0); });
        
        rows.forEach((e) => {
          if (have.has(k(e))) return skipped++;
          have.add(k(e)); 
          const g = e.type + e.date; 
          cnt[g] = (cnt[g] || 0) + 1;
          
          DB.enc.push({ ...e, id: newId(), n: cnt[g], out: '', done: false, src: 'travel' }); 
          added++;
        });
        
        log('Импорт Travel', `${from} — ${to}: добавлено ${added}, пропущено ${skipped}`);
      });
      
      S.date = from; 
      render();
      toast(`Travel: добавлено ${added}, уже было ${skipped}`);
    } catch (err) { 
      alert('Ошибка импорта: ' + err.message); 
      console.error(err);
    }
  }, 30);
  
  rd.readAsArrayBuffer(file);
}