'use strict';

// 양식의 열 정의. match 는 헤더 글자로 열 위치를 찾을 때 쓴다.
const COLS = [
  { key: 'no',     label: '연번',               match: /연번/,      width: 6.6 },
  { key: 'region', label: '지역\n(시/군)',      match: /지역/,      width: 9.5 },
  { key: 'level',  label: '학교급',             match: /학교급/,    width: 8.5 },
  { key: 'school', label: '학교명',             match: /학교명/,    width: 22.5 },
  { key: 'grade',  label: '학년',               match: /학년/,      width: 6.5 },
  { key: 'name',   label: '성명',               match: /성명|이름/, width: 12 },
  { key: 'gender', label: '성별',               match: /성별/,      width: 6.5 },
  { key: 'birth',  label: '생년월일',           match: /생년월일/,  width: 14 },
  { key: 'date',   label: '일시',               match: /일시/,      width: 24 },
  { key: 'place',  label: '장소',               match: /장소/,      width: 19.6 },
  { key: 'merit',  label: '공적개요  (50-100자)', match: /공적/,    width: 69.4 },
];
const DATA_KEYS = COLS.filter(c => c.key !== 'no').map(c => c.key);
const REQUIRED = DATA_KEYS;
const LABEL = Object.fromEntries(COLS.map(c => [c.key, c.label.replace(/\n/g, '').replace(/\s*\(50-100자\)/, '')]));
const CAT_ORDER = ['초', '중', '고', '특'];
const DEC_NOTE = '졸업식 12월인 경우 노란색 음영 표시하여 제출';
const EXAMPLE_SCHOOL = /^\s*(OO|○○|00|ㅇㅇ|△△)/i;
const WEEK = ['일', '월', '화', '수', '목', '금', '토'];

// fileName -> { name, sheets: [{ cat, title, rows }], error }
const files = new Map();
let activeTab = null;

const $ = id => document.getElementById(id);

// ---------- 읽기 ----------

function pad(n) { return String(n).padStart(2, '0'); }

function formatDate(d, key) {
  // ExcelJS 의 날짜는 UTC 기준
  const y = d.getUTCFullYear(), m = d.getUTCMonth() + 1, day = d.getUTCDate();
  if (key === 'date') {
    const hh = d.getUTCHours(), mm = d.getUTCMinutes();
    const base = `${y}. ${m}. ${day}.(${WEEK[d.getUTCDay()]})`;
    return hh || mm ? `${base} ${pad(hh)}:${pad(mm)}` : base;
  }
  return `${y}.${pad(m)}.${pad(day)}`;
}

function cellText(cell, key) {
  let v = cell.value;
  if (v == null) return '';
  if (v instanceof Date) return formatDate(v, key);
  if (typeof v === 'object') {
    if (v.richText) v = v.richText.map(t => t.text).join('');
    else if ('result' in v) v = v.result instanceof Date ? formatDate(v.result, key) : v.result;
    else if (v.text != null) v = v.text;
    else if (v.error) v = '';
  }
  return v == null ? '' : String(v).trim();
}

function isFilled(cell) {
  const f = cell.fill;
  if (!f || f.type !== 'pattern' || !f.pattern || f.pattern === 'none') return false;
  const c = f.fgColor || {};
  if (c.argb) return !/^(FF)?FFFFFF$/i.test(c.argb);
  return c.theme != null && !(c.theme === 0 && !c.tint);
}

function findHeader(ws) {
  const maxScan = Math.min(ws.rowCount, 15);
  const rowTexts = r => {
    const out = [];
    ws.getRow(r).eachCell({ includeEmpty: false }, (c, col) => { out[col] = cellText(c).replace(/\s+/g, ''); });
    return out;
  };
  for (let r = 1; r <= maxScan; r++) {
    const texts = rowTexts(r);
    if (!texts.some(t => t && /성명|이름/.test(t)) || !texts.some(t => t && /학교명/.test(t))) continue;
    const sub = rowTexts(r + 1);
    const map = {};
    for (const row of [texts, sub]) row.forEach((t, col) => {
      if (!t) return;
      for (const def of COLS) if (!(def.key in map) && def.match.test(t)) { map[def.key] = col; break; }
    });
    // 못 찾은 열은 양식 기본 위치(A~K) 사용
    COLS.forEach((def, i) => { if (!(def.key in map)) map[def.key] = i + 1; });
    const hasSub = sub.some(t => t && /일시|장소/.test(t));
    return { map, start: r + (hasSub ? 2 : 1), headerRow: r };
  }
  return null;
}

function catOf(sheetName) {
  const m = sheetName.match(/\((초|중|고|특)/);
  return m ? m[1] : null;
}

// ExcelJS 는 메모·그림 부분이 표준과 조금만 달라도 읽기에 실패한다(한셀·기타 프로그램 저장 파일).
// 값을 읽는 데는 필요 없으므로 미리 걷어낸다.
async function sanitize(buf) {
  const zip = await JSZip.loadAsync(buf);
  const drop = /\/(comments|vmlDrawing|drawing|threadedComment|image)$/;
  for (const path of Object.keys(zip.files)) {
    if (/^xl\/worksheets\/_rels\/.+\.rels$/.test(path)) {
      let xml = await zip.file(path).async('string');
      xml = xml.replace(/<Relationship\b[^>]*\/>/g, rel => {
        const type = (rel.match(/Type="([^"]+)"/) || [])[1] || '';
        return drop.test(type) ? '' : rel;
      });
      zip.file(path, xml);
    } else if (/^xl\/worksheets\/[^/]+\.xml$/.test(path)) {
      let xml = await zip.file(path).async('string');
      xml = xml.replace(/<(\w+:)?(legacyDrawing|legacyDrawingHF|drawing)\b[^>]*\/>/g, '');
      zip.file(path, xml);
    } else if (/^xl\/(comments|threadedComments|drawings|persons)\//.test(path) || /^xl\/comments\d*\.xml$/.test(path)) {
      zip.remove(path);
    }
  }
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function loadWorkbook(buf) {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buf);
    return wb;
  } catch (first) {
    const retry = new ExcelJS.Workbook();
    await retry.xlsx.load(await sanitize(buf));
    return retry;
  }
}

async function readFile(file) {
  const entry = { name: file.name, sheets: [], error: null };
  try {
    if (/\.xls$/i.test(file.name)) throw new Error('예전 형식(.xls) 파일입니다. 엑셀에서 “다른 이름으로 저장 → Excel 통합 문서(*.xlsx)”로 저장 후 올려 주세요.');
    if (!/\.xlsx$/i.test(file.name)) throw new Error('엑셀 파일(.xlsx)이 아닙니다.');
    const wb = await loadWorkbook(await file.arrayBuffer());
    wb.eachSheet(ws => {
      const sheetName = ws.name.trim();
      if (/예시/.test(sheetName) || ws.state === 'hidden' || ws.state === 'veryHidden') return;
      const hdr = findHeader(ws);
      if (!hdr) return;
      const titleCell = hdr.headerRow > 1 ? cellText(ws.getRow(1).getCell(1)) : '';
      const rows = [];
      for (let r = hdr.start; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const rec = { file: file.name, sheet: sheetName, srcRow: r };
        let any = false;
        for (const key of DATA_KEYS) {
          rec[key] = cellText(row.getCell(hdr.map[key]), key);
          if (rec[key]) any = true;
        }
        if (!any) continue;
        // 표 아래 “※ …” 안내 문구 줄은 건너뜀
        if (!rec.name && DATA_KEYS.some(k => rec[k].startsWith('※'))) continue;
        rec.filled = isFilled(row.getCell(hdr.map.date));
        rows.push(rec);
      }
      entry.sheets.push({ name: sheetName, cat: catOf(sheetName), title: titleCell, rows });
    });
    if (!entry.sheets.length) throw new Error('양식 시트를 찾지 못했습니다. (학교명·성명 헤더가 있는 시트가 없음)');
  } catch (e) {
    entry.error = e.message || String(e);
  }
  return entry;
}

// ---------- 취합 ----------

function opts() {
  return {
    skipExample: $('optSkipExample').checked,
    decYellow: $('optDecYellow').checked,
    sort: $('optSort').checked,
    summary: $('optSummary').checked,
  };
}

function isDecember(text) {
  const m = text.match(/(\d{4})\s*[.\-/년]\s*(\d{1,2})/);
  return !!m && Number(m[2]) === 12;
}

// 2008.1.5 / 2008-01-05 / 20080105 / 2008년 1월 5일 → 2008.01.05
function normBirth(text) {
  const m = text.match(/^(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})\s*[.일]?$/) || text.match(/^(\d{4})(\d{2})(\d{2})$/);
  return m ? `${m[1]}.${pad(m[2])}.${pad(m[3])}` : text;
}

function meritLength(text) { return [...text.replace(/\s+/g, ' ').trim()].length; }

function validate(rec) {
  const issues = [];
  const missing = REQUIRED.filter(k => !rec[k]);
  if (missing.length) issues.push({ level: 'err', msg: `빈 칸: ${missing.map(k => LABEL[k]).join(', ')}` });
  if (rec.birth && !/^\d{4}\.\d{2}\.\d{2}$/.test(rec.birth)) issues.push({ level: 'warn', msg: `생년월일 형식 확인(예: 2014.03.05): “${rec.birth}”` });
  if (rec.gender && !/^(남|여)$/.test(rec.gender)) issues.push({ level: 'warn', msg: `성별 값 확인: “${rec.gender}”` });
  if (rec.merit) {
    const n = meritLength(rec.merit);
    if (n < 50 || n > 100) issues.push({ level: 'warn', msg: `공적개요 ${n}자 (50~100자 권장)` });
  }
  return issues;
}

function collect() {
  const o = opts();
  const sheets = new Map(); // sheetName -> { name, cat, title, rows }
  const perFile = [];
  const notices = [];

  for (const f of files.values()) {
    const info = { name: f.name, school: '', counts: {}, excluded: 0, error: f.error, issues: 0 };
    perFile.push(info);
    if (f.error) { notices.push({ level: 'err', where: f.name, msg: f.error }); continue; }
    const schools = {};
    for (const s of f.sheets) {
      if (!sheets.has(s.name)) sheets.set(s.name, { name: s.name, cat: s.cat, title: s.title, rows: [] });
      const target = sheets.get(s.name);
      if (!target.title && s.title) target.title = s.title;
      info.counts[s.name] = 0;
      for (const rec of s.rows) {
        if (o.skipExample && EXAMPLE_SCHOOL.test(rec.school)) { info.excluded++; continue; }
        const r = { ...rec, birth: normBirth(rec.birth) };
        r.issues = validate(r);
        r.yellow = rec.filled || (o.decYellow && isDecember(rec.date));
        target.rows.push(r);
        info.counts[s.name]++;
        if (rec.school) schools[rec.school] = (schools[rec.school] || 0) + 1;
      }
    }
    info.school = Object.entries(schools).sort((a, b) => b[1] - a[1]).map(e => e[0]).join(', ');
    if (info.excluded) notices.push({ level: 'warn', where: f.name, msg: `예시 행 ${info.excluded}개를 제외했습니다. (학교명이 “OO…”로 시작)` });
    if (!Object.values(info.counts).some(Boolean) && !info.excluded) notices.push({ level: 'warn', where: f.name, msg: '입력된 대상자가 없습니다.' });
  }

  // 중복(학교명+성명+생년월일) 확인
  const seen = new Map();
  for (const s of sheets.values()) for (const r of s.rows) {
    const k = [r.school, r.name, r.birth].join('|');
    if (!r.name) continue;
    if (seen.has(k)) {
      const p = seen.get(k);
      r.issues.push({ level: 'warn', msg: `중복 의심: ${p.file} [${p.sheet}] ${p.srcRow}행과 동일` });
    } else seen.set(k, r);
  }
  const byName = new Map(perFile.map(f => [f.name, f]));
  for (const s of sheets.values()) for (const r of s.rows) byName.get(r.file).issues += r.issues.length;

  const order = name => {
    const c = catOf(name);
    const i = c ? CAT_ORDER.indexOf(c) : 99;
    return (/교육감/.test(name) ? 0 : /교육장/.test(name) ? 100 : 200) + i;
  };
  const list = [...sheets.values()].sort((a, b) => order(a.name) - order(b.name) || a.name.localeCompare(b.name, 'ko'));
  if (o.sort) for (const s of list) {
    s.rows = s.rows.map((r, i) => [r, i])
      .sort((a, b) => a[0].school.localeCompare(b[0].school, 'ko') || a[1] - b[1])
      .map(x => x[0]);
  }
  for (const s of list) for (const r of s.rows) for (const is of r.issues) {
    notices.push({ level: is.level, where: `${r.file} › ${s.name} ${r.srcRow}행 (${r.school || '학교명 없음'} ${r.name || ''})`, msg: is.msg });
  }
  return { sheets: list, perFile, notices };
}

// ---------- 화면 ----------

function esc(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function render() {
  const has = files.size > 0;
  for (const id of ['filesCard', 'previewCard', 'warnCard', 'exportCard']) $(id).hidden = !has;
  if (!has) return;

  const data = collect();
  const sheetNames = data.sheets.map(s => s.name);

  // 파일 목록
  $('fileCount').textContent = files.size;
  let h = '<thead><tr><th>#</th><th>파일명</th><th>학교명</th>' +
    sheetNames.map(n => `<th>${esc(shortName(n))}</th>`).join('') + '<th>합계</th><th>상태</th><th></th></tr></thead><tbody>';
  data.perFile.forEach((f, i) => {
    const total = Object.values(f.counts).reduce((a, b) => a + b, 0);
    const status = f.error ? '<span class="tag err">읽기 실패</span>'
      : f.issues ? `<span class="tag warn">확인 ${f.issues}</span>` : '<span class="tag ok">정상</span>';
    h += `<tr class="${f.error ? 'has-err' : ''}"><td>${i + 1}</td><td class="left">${esc(f.name)}</td><td class="left">${esc(f.school || '-')}</td>` +
      sheetNames.map(n => {
        const c = f.counts[n];
        return `<td class="num ${c ? '' : 'zero'}">${c == null ? '·' : c}</td>`;
      }).join('') +
      `<td class="num"><b>${total}</b></td><td>${status}</td>` +
      `<td><button class="btn icon" data-remove="${esc(f.name)}" title="이 파일 빼기">✕</button></td></tr>`;
  });
  $('fileTable').innerHTML = h + '</tbody>';

  // 통계
  const totalRows = data.sheets.reduce((a, s) => a + s.rows.length, 0);
  const schoolSet = new Set(data.sheets.flatMap(s => s.rows.map(r => r.school)).filter(Boolean));
  $('stats').innerHTML = `<span>학교 <b>${schoolSet.size}</b>곳</span><span>대상자 <b>${totalRows}</b>명</span>` +
    data.sheets.map(s => `<span>${esc(shortName(s.name))} <b>${s.rows.length}</b></span>`).join('');

  // 탭
  if (!sheetNames.includes(activeTab)) activeTab = sheetNames[0] || null;
  $('tabs').innerHTML = data.sheets.map(s =>
    `<button role="tab" data-tab="${esc(s.name)}" aria-selected="${s.name === activeTab}">${esc(s.name)} (${s.rows.length})</button>`).join('');
  const cur = data.sheets.find(s => s.name === activeTab);
  if (!cur || !cur.rows.length) {
    $('previewTable').innerHTML = `<tbody><tr><td class="empty">이 시트에 취합된 대상자가 없습니다.</td></tr></tbody>`;
  } else {
    let p = '<thead><tr>' + COLS.map(c => `<th>${esc(LABEL[c.key])}</th>`).join('') + '<th>원본</th></tr></thead><tbody>';
    cur.rows.forEach((r, i) => {
      const lv = r.issues.some(x => x.level === 'err') ? 'has-err' : r.issues.length ? 'has-warn' : '';
      p += `<tr class="${lv}" title="${esc(r.issues.map(x => x.msg).join('\n'))}"><td>${i + 1}</td>` +
        DATA_KEYS.map(k => {
          if (k === 'merit') return `<td class="merit">${esc(r.merit)}<small>${meritLength(r.merit)}자</small></td>`;
          return `<td class="${k === 'date' && r.yellow ? 'yellow' : ''}">${esc(r[k])}</td>`;
        }).join('') +
        `<td class="src">${esc(r.file)}<br>${r.srcRow}행</td></tr>`;
    });
    $('previewTable').innerHTML = p + '</tbody>';
  }

  // 경고
  $('warnCount').textContent = data.notices.length;
  $('warnList').innerHTML = data.notices.length
    ? data.notices.map(n => `<li class="${n.level}">${esc(n.msg)} <span class="where">— ${esc(n.where)}</span></li>`).join('')
    : '<li>확인이 필요한 항목이 없습니다.</li>';

  $('download').disabled = totalRows === 0 && !data.sheets.length;
}

function shortName(n) {
  const m = n.match(/\(([^)]+)\)/);
  return m ? m[1] : n;
}

async function addFiles(list) {
  const arr = [...list];
  if (!arr.length) return;
  const results = await Promise.all(arr.map(readFile));
  for (const r of results) files.set(r.name, r);
  render();
}

// ---------- 내보내기 ----------

const THIN = { style: 'thin', color: { argb: 'FF000000' } };
const MEDIUM = { style: 'medium', color: { argb: 'FF000000' } };
const FONT = { name: '맑은 고딕', size: 11 };
const CENTER = { horizontal: 'center', vertical: 'middle', wrapText: true };

function defaultTitle(name) {
  const year = ($('outName').value.match(/(\d{4})학년도/) || [])[1] || new Date().getFullYear();
  const kinds = { 초: '초등학교', 중: '중학교', 고: '고등학교', 특: '특수학교' };
  const c = catOf(name);
  const who = (name.match(/^(교육감|교육장)/) || [])[1];
  return `${year}학년도 졸업생 표창 대상자 추천 명단(${who && c ? `${who}_${kinds[c]}` : name})`;
}

function rowHeight(r) {
  // K열 폭 69 기준 한글 약 33자/줄
  const lines = Math.max(1, Math.ceil(meritLength(r.merit) / 33), Math.ceil([...r.school].length / 10));
  return Math.max(28.9, lines * 16.5 + 6);
}

function writeSheet(wb, s) {
  const ws = wb.addWorksheet(s.name.slice(0, 31), {
    views: [{ state: 'frozen', ySplit: 3 }],
    pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      horizontalCentered: true, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 } },
  });
  ws.columns = COLS.map(c => ({ width: c.width }));
  ws.pageSetup.printTitlesRow = '2:3';

  ws.mergeCells('A1:K1');
  const t = ws.getCell('A1');
  t.value = s.title || defaultTitle(s.name);
  t.font = { ...FONT, size: 20, bold: true };
  t.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 58.5;

  // 2줄 헤더 (졸업식 → 일시/장소)
  COLS.forEach((c, i) => {
    const col = i + 1;
    if (c.key === 'date' || c.key === 'place') return;
    ws.mergeCells(2, col, 3, col);
    ws.getCell(2, col).value = c.label;
  });
  ws.mergeCells('I2:J2');
  ws.getCell('I2').value = '졸업식';
  ws.getCell('I2').note = DEC_NOTE;
  ws.getCell('I3').value = '일시';
  ws.getCell('J3').value = '장소';
  ws.getRow(2).height = 20.25;
  ws.getRow(3).height = 23.25;
  for (let r = 2; r <= 3; r++) for (let col = 1; col <= COLS.length; col++) {
    const cell = ws.getCell(r, col);
    cell.font = { ...FONT, bold: true };
    cell.alignment = CENTER;
    cell.border = { top: r === 2 ? MEDIUM : THIN, left: THIN, right: THIN, bottom: THIN };
  }

  s.rows.forEach((r, i) => {
    const rowNo = 4 + i;
    const values = [i + 1, ...DATA_KEYS.map(k => (k === 'grade' && /^\d+$/.test(r[k]) ? Number(r[k]) : r[k]))];
    const row = ws.getRow(rowNo);
    values.forEach((v, j) => {
      const cell = row.getCell(j + 1);
      cell.value = v === '' ? null : v;
      cell.font = FONT;
      cell.alignment = CENTER;
      cell.border = { top: THIN, left: THIN, right: THIN, bottom: THIN };
    });
    if (r.yellow) row.getCell(9).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    row.height = rowHeight(r);
  });

  // 대상자가 없어도 빈 칸 몇 줄은 남겨 양식 모양 유지
  for (let k = s.rows.length; k < 3; k++) {
    const row = ws.getRow(4 + k);
    for (let col = 1; col <= COLS.length; col++) {
      row.getCell(col).border = { top: THIN, left: THIN, right: THIN, bottom: THIN };
    }
    row.height = 28.9;
  }
}

function writeSummary(wb, data) {
  const ws = wb.addWorksheet('취합현황', { views: [{ state: 'frozen', ySplit: 2 }] });
  const names = data.sheets.map(s => s.name);
  const head = ['연번', '파일명', '학교명', ...names.map(shortName), '합계', '비고'];
  ws.columns = head.map((_, i) => ({ width: i === 1 ? 48 : i === 2 ? 24 : i === head.length - 1 ? 30 : 9 }));
  ws.mergeCells(1, 1, 1, head.length);
  ws.getCell('A1').value = '학교별 취합 현황';
  ws.getCell('A1').font = { ...FONT, size: 16, bold: true };
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 36;
  const box = { top: THIN, left: THIN, right: THIN, bottom: THIN };
  ws.addRow(head).eachCell(c => { c.font = { ...FONT, bold: true }; c.alignment = CENTER; c.border = box;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F3F8' } }; });
  const totals = names.map(() => 0);
  data.perFile.forEach((f, i) => {
    const counts = names.map((n, j) => { const c = f.counts[n] || 0; totals[j] += c; return c; });
    const note = f.error ? `읽기 실패: ${f.error}` : f.issues ? `확인 필요 ${f.issues}건` : '';
    ws.addRow([i + 1, f.name, f.school, ...counts, counts.reduce((a, b) => a + b, 0), note])
      .eachCell({ includeEmpty: true }, (c, col) => { c.font = FONT; c.border = box;
        c.alignment = col === 2 || col === head.length ? { vertical: 'middle', wrapText: true } : CENTER; });
  });
  ws.addRow(['합계', '', '', ...totals, totals.reduce((a, b) => a + b, 0), ''])
    .eachCell({ includeEmpty: true }, c => { c.font = { ...FONT, bold: true }; c.alignment = CENTER; c.border = box; });
}

async function download() {
  const btn = $('download');
  btn.disabled = true;
  const label = btn.textContent;
  btn.textContent = '만드는 중…';
  try {
    const data = collect();
    const wb = new ExcelJS.Workbook();
    wb.creator = '졸업생 표창 명단 취합';
    wb.created = new Date();
    for (const s of data.sheets) writeSheet(wb, s);
    if ($('optSummary').checked) writeSummary(wb, data);
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    let name = $('outName').value.trim() || '취합.xlsx';
    if (!/\.xlsx$/i.test(name)) name += '.xlsx';
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  } catch (e) {
    alert('엑셀 파일을 만드는 중 오류가 발생했습니다.\n' + (e.message || e));
  } finally {
    btn.textContent = label;
    btn.disabled = false;
  }
}

// ---------- 이벤트 ----------

const drop = $('drop');
['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => addFiles(e.dataTransfer.files));
$('fileInput').addEventListener('change', e => { addFiles(e.target.files); e.target.value = ''; });
// 드롭존 밖에 떨어뜨려도 브라우저가 파일을 열지 않도록
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => { e.preventDefault(); if (!drop.contains(e.target)) addFiles(e.dataTransfer.files); });

$('fileTable').addEventListener('click', e => {
  const b = e.target.closest('[data-remove]');
  if (b) { files.delete(b.dataset.remove); render(); }
});
$('tabs').addEventListener('click', e => {
  const b = e.target.closest('[data-tab]');
  if (b) { activeTab = b.dataset.tab; render(); }
});
$('clearAll').addEventListener('click', () => {
  if (confirm('업로드한 파일을 모두 지울까요?')) { files.clear(); render(); }
});
for (const id of ['optSkipExample', 'optDecYellow', 'optSort']) $(id).addEventListener('change', render);
$('download').addEventListener('click', download);

if (typeof ExcelJS === 'undefined' || typeof JSZip === 'undefined') {
  document.querySelector('main').insertAdjacentHTML('afterbegin',
    '<section class="card" style="border-color:#c62828;color:#c62828">엑셀 처리 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인하고 새로고침해 주세요.</section>');
}
