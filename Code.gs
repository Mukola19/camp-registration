// ═══════════════════════════════════════════════════════════════
//  Google Apps Script — бекенд для реєстрації табору
//
//  Аркуш «Учасники» (заголовки в рядку 1):
//    ID | Дата | ПІБ | Група | ГрупаID | Вік | Телефон | Дійсний
//    (колонку «ГрупаID» скрипт додасть сам, якщо її немає)
//    «ГрупаID» — стабільний номер групи (1..5), головний звʼязок.
//    «Група» — назва для читабельності (може застаріти після перейменування).
//
//  Аркуш «Групи» (створюється автоматично) — УСЯ конфігурація груп:
//    ID | Група | Наставники | Колір | ВікВід | ВікДо
//    «ID» — стабільний номер (1..5), НЕ змінюється. «Група» — назва (редагована).
//  Назви, наставники, колір і вікові межі зберігаються ТУТ (в таблиці),
//  тож їх можна редагувати/перейменовувати без зміни коду.
//  Вікові межі також оновлює «Перерозподіл».
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME   = 'Учасники';
const GROUPS_SHEET = 'Групи';

const GROUPS_HEADERS = ['ID', 'Група', 'Наставники', 'Колір', 'ВікВід', 'ВікДо'];

// Значення за замовчуванням (для першого створення аркуша «Групи»).
// Перша колонка — стабільний ID. Далі редагується прямо в таблиці.
const DEFAULT_GROUPS = [
  [1, 'Група 1', 'Величко Аліна, Дроган Дана, Лозовик Сергій', '#E0665A', 5,  6],
  [2, 'Група 2', 'Мельник Ігор, Похільчук Яна, Мельник Юля',    '#E8944A', 7,  8],
  [3, 'Група 3', 'Чміль Віталік і Діана, Аня',                  '#3DAA6E', 9,  10],
  [4, 'Група 4', 'Войтович Коля, Гусак Юля',                    '#4A90D9', 11, 12],
  [5, 'Група 5', 'Мельник Назар і Даша',                        '#9B59B6', 13, 15],
];

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  let result;
  try {
    if      (action === 'getParticipants')   result = getParticipants();
    else if (action === 'addParticipant')    result = addParticipant(e.parameter);
    else if (action === 'updateParticipant') result = updateParticipant(e.parameter);
    else if (action === 'deleteParticipant') result = deleteParticipant(e.parameter);
    else if (action === 'getGroups')         result = { groups: readGroups(getOrCreateGroupsSheet()) };
    else if (action === 'reorganize')        result = reorganize();
    else result = { error: 'Невідома дія: ' + action };
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Не знайдено аркуш «' + SHEET_NAME + '»');
  return sh;
}

// Номер колонки за назвою заголовка (1-based), -1 якщо немає
function colIndex(sh, name) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = headers.findIndex(h => String(h).trim() === name);
  return idx === -1 ? -1 : idx + 1;
}

// ─── УЧАСНИКИ: ЧИТАННЯ ───────────────────────────────────────
function getParticipants() {
  const sh = sheet();
  if (sh.getLastRow() < 2) return { participants: [] };
  const values  = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const participants = [];
  for (let i = 1; i < values.length; i++) {
    const obj = { row: i + 1 };            // row = номер рядка в таблиці
    headers.forEach((h, j) => { obj[h] = values[i][j]; });
    participants.push(obj);
  }
  return { participants };
}

// Гарантує наявність колонки «ГрупаID» в аркуші «Учасники»
function ensureParticipantColumns(sh) {
  const headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim());
  if (headers.indexOf('ГрупаID') === -1) {
    sh.getRange(1, headers.length + 1).setValue('ГрупаID');
  }
}

// ─── УЧАСНИКИ: ДОДАВАННЯ ─────────────────────────────────────
function addParticipant(p) {
  const sh = sheet();
  ensureParticipantColumns(sh);
  const id = Date.now();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const map = {
    'ID': id,
    'Дата':    p.date    || '',
    'ПІБ':     p.pib     || '',
    'Група':   p.group   || '',
    'ГрупаID': p.groupId || '',
    'Вік':     p.age     || '',
    'Телефон': p.phone   || '',
    'Дійсний': true,
  };
  sh.appendRow(headers.map(h => (h in map ? map[h] : '')));
  return { ok: true, id: id };
}

// ─── УЧАСНИКИ: ОНОВЛЕННЯ (ПІБ / Група / ГрупаID / Вік / Телефон) ───
function updateParticipant(p) {
  const sh  = sheet();
  const row = Number(p.row);
  if (!row || row < 2) throw new Error('Невірний номер рядка');
  ensureParticipantColumns(sh);
  if (p.pib     !== undefined) setCell(sh, row, 'ПІБ',     p.pib);
  if (p.group   !== undefined) setCell(sh, row, 'Група',   p.group);
  if (p.groupId !== undefined) setCell(sh, row, 'ГрупаID', p.groupId);
  if (p.age     !== undefined) setCell(sh, row, 'Вік',     p.age);
  if (p.phone   !== undefined) setCell(sh, row, 'Телефон', p.phone);
  return { ok: true };
}

function setCell(sh, row, header, value) {
  const c = colIndex(sh, header);
  if (c !== -1) sh.getRange(row, c).setValue(value);
}

// ─── УЧАСНИКИ: «ВИДАЛЕННЯ» (позначити недійсним) ─────────────
function deleteParticipant(p) {
  const sh  = sheet();
  const row = Number(p.row);
  if (!row || row < 2) throw new Error('Невірний номер рядка');
  setCell(sh, row, 'Дійсний', false);
  return { ok: true };
}

// ═══════════════════════════════════════════════════════════════
//  ГРУПИ (динамічні вікові межі)
// ═══════════════════════════════════════════════════════════════
function getOrCreateGroupsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let gsh = ss.getSheetByName(GROUPS_SHEET);
  if (!gsh) {
    gsh = ss.insertSheet(GROUPS_SHEET);
    gsh.getRange(1, 1, 1, GROUPS_HEADERS.length).setValues([GROUPS_HEADERS]);
    gsh.getRange(2, 1, DEFAULT_GROUPS.length, GROUPS_HEADERS.length).setValues(DEFAULT_GROUPS);
    return gsh;
  }
  ensureGroupColumns(gsh);   // міграція старого аркуша (лише межі → повна схема)
  return gsh;
}

// Додає відсутні колонки та заповнює ID/наставників/колір за замовчуванням
function ensureGroupColumns(gsh) {
  let headers = gsh.getRange(1, 1, 1, Math.max(gsh.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim());
  GROUPS_HEADERS.forEach(h => {
    if (headers.indexOf(h) === -1) {
      gsh.getRange(1, headers.length + 1).setValue(h);
      headers.push(h);
    }
  });
  const vals = gsh.getDataRange().getValues();
  const hz   = vals[0].map(h => String(h).trim());
  const cId = hz.indexOf('ID'), cName = hz.indexOf('Група'), cMen = hz.indexOf('Наставники'), cCol = hz.indexOf('Колір');
  for (let i = 1; i < vals.length; i++) {
    const name = String(vals[i][cName] || '').trim();
    if (!name) continue;
    // ID за замовчуванням = порядковий номер рядка (стабільний після цього)
    if (cId !== -1 && (vals[i][cId] === '' || vals[i][cId] == null)) gsh.getRange(i + 1, cId + 1).setValue(i);
    const def = DEFAULT_GROUPS.find(d => d[1] === name);
    if (!def) continue;
    if (cMen !== -1 && !String(vals[i][cMen] || '').trim()) gsh.getRange(i + 1, cMen + 1).setValue(def[2]);
    if (cCol !== -1 && !String(vals[i][cCol] || '').trim()) gsh.getRange(i + 1, cCol + 1).setValue(def[3]);
  }
}

function readGroups(gsh) {
  const vals = gsh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0].map(h => String(h).trim());
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const o = {};
    headers.forEach((h, j) => { o[h] = vals[i][j]; });
    if (String(o['Група'] || '').trim()) out.push(o);
  }
  return out;
}

function writeGroupRanges(gsh, ids, ranges) {
  const vals    = gsh.getDataRange().getValues();
  const headers = vals[0].map(h => String(h).trim());
  const cId  = headers.indexOf('ID');
  const cMin = headers.indexOf('ВікВід');
  const cMax = headers.indexOf('ВікДо');
  for (let i = 1; i < vals.length; i++) {
    const idx = ids.map(String).indexOf(String(vals[i][cId]));
    if (idx === -1) continue;
    const r = ranges[idx];
    if (cMin !== -1) gsh.getRange(i + 1, cMin + 1).setValue(r.min == null ? '' : r.min);
    if (cMax !== -1) gsh.getRange(i + 1, cMax + 1).setValue(r.max == null ? '' : r.max);
  }
}

// ─── АЛГОРИТМ РІВНОМІРНОГО РОЗПОДІЛУ ─────────────────────────
// Ділить дітей на G суміжних вікових груп максимально рівних за кількістю.
// Діти одного віку НІКОЛИ не розриваються між групами.
//
// partitionCounts: розбиває масив кількостей (по одному на кожен унікальний
// вік) на рівно G суміжних частин. Критерій (лексикографічно):
//   1) мінімізувати НАЙБІЛЬШУ групу (щоб не було однієї роздутої);
//   2) за рівності — мінімізувати суму квадратів розмірів (загальна рівність).
// Динамічне програмування — оптимальний поділ.
function partitionCounts(counts, G) {
  const m = counts.length;
  if (m === 0) return [];
  if (m <= G) return counts.map((_, i) => i); // кожен вік — окрема група

  const prefix = [0];
  for (let i = 0; i < m; i++) prefix.push(prefix[i] + counts[i]);
  const size = (a, b) => prefix[b] - prefix[a];   // сума на [a, b)

  // краще: спершу менший максимум, потім менша сума квадратів
  const better = (a, b) => (a.mx !== b.mx ? a.mx < b.mx : a.sq < b.sq);

  // dp[k][i] = найкраща пара {mx, sq} для поділу перших i віків на k частин
  const dp  = [];
  const cut = [];
  for (let k = 0; k <= G; k++) { dp.push(new Array(m + 1).fill(null)); cut.push(new Array(m + 1).fill(-1)); }
  dp[0][0] = { mx: 0, sq: 0 };
  for (let k = 1; k <= G; k++) {
    for (let i = k; i <= m; i++) {
      for (let j = k - 1; j < i; j++) {
        if (!dp[k - 1][j]) continue;
        const s = size(j, i);
        const cand = { mx: Math.max(dp[k - 1][j].mx, s), sq: dp[k - 1][j].sq + s * s };
        if (dp[k][i] === null || better(cand, dp[k][i])) { dp[k][i] = cand; cut[k][i] = j; }
      }
    }
  }

  const bucketOf = new Array(m);
  let i = m;
  for (let k = G; k >= 1; k--) {
    const j = cut[k][i];
    for (let a = j; a < i; a++) bucketOf[a] = k - 1;
    i = j;
  }
  return bucketOf;
}

// ages — відсортований за зростанням масив віків. Повертає індекс групи (0..G-1)
// для кожного елемента.
function evenBuckets(ages, G) {
  const n = ages.length;
  const res = new Array(n).fill(0);
  if (n === 0) return res;

  const counts = [];
  for (let i = 0; i < n; i++) {
    if (i === 0 || ages[i] !== ages[i - 1]) counts.push(1);
    else counts[counts.length - 1]++;
  }

  const bucketOf = partitionCounts(counts, G);

  let di = -1;
  for (let i = 0; i < n; i++) {
    if (i === 0 || ages[i] !== ages[i - 1]) di++;
    res[i] = bucketOf[di];
  }
  return res;
}

function reorganize() {
  const gsh = getOrCreateGroupsSheet();
  const grp = readGroups(gsh);
  if (!grp.length) throw new Error('Немає груп в аркуші «' + GROUPS_SHEET + '»');
  const groupIds   = grp.map(g => g['ID']);
  const groupNames = grp.map(g => String(g['Група']).trim());
  const G = grp.length;

  // Дійсні учасники з валідним віком
  const all   = getParticipants().participants;
  const valid = all
    .filter(p => {
      const d   = p['Дійсний'];
      const isV = d === true || d === '' || d == null || String(d).toLowerCase() === 'true';
      const age = Number(p['Вік']);
      return isV && !isNaN(age) && age > 0;
    })
    .map(p => ({ row: p.row, age: Number(p['Вік']) }))
    .sort((a, b) => a.age - b.age);

  const assign = evenBuckets(valid.map(v => v.age), G);

  // Межі груп + оновлення групи (id + назви) кожного учасника
  const ranges = grp.map(() => ({ min: null, max: null }));
  const psh    = sheet();
  ensureParticipantColumns(psh);
  const grpCol = colIndex(psh, 'Група');
  const gidCol = colIndex(psh, 'ГрупаID');

  valid.forEach((v, i) => {
    const gi = assign[i];
    if (gidCol !== -1) psh.getRange(v.row, gidCol).setValue(groupIds[gi]);
    if (grpCol !== -1) psh.getRange(v.row, grpCol).setValue(groupNames[gi]);
    const r = ranges[gi];
    if (r.min == null || v.age < r.min) r.min = v.age;
    if (r.max == null || v.age > r.max) r.max = v.age;
  });

  writeGroupRanges(gsh, groupIds, ranges);
  return { ok: true, count: valid.length, groups: readGroups(gsh) };
}
