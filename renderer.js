// ===================== Genel durum =====================
const state = {
  mode: 'start',      // 'start' | 'word' | 'excel'
  path: null,
  dirty: false,
  sheets: [],         // excel: [{name, data}]
  activeSheet: 0,
  sel: { r: 0, c: 0 }, // excel seçili hücre
};

const $ = (sel) => document.querySelector(sel);
const screens = { start: $('#start'), word: $('#word'), excel: $('#excel') };
const editor = $('#editor');
const grid = $('#grid');
const sheetTabs = $('#sheetTabs');
const formulaInput = $('#formulaInput');

const MIN_ROWS = 40;
const MIN_COLS = 15;

function showStatus(msg, ms = 2000) {
  const el = $('#status');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(showStatus._t);
  showStatus._t = setTimeout(() => el.classList.remove('show'), ms);
}

function fileName() {
  if (!state.path) return state.mode === 'word' ? 'Yeni Belge' : state.mode === 'excel' ? 'Yeni Tablo' : '';
  return state.path.split(/[\\/]/).pop();
}

function updateTitle() {
  const name = fileName();
  const title = name ? `${state.dirty ? '• ' : ''}${name} — Yazgaç` : 'Yazgaç';
  window.api.setTitle(title);
}

function setDirty(v) {
  state.dirty = v;
  updateTitle();
}

function showScreen(mode) {
  state.mode = mode;
  Object.entries(screens).forEach(([k, el]) => el.classList.toggle('hidden', k !== mode));
  updateTitle();
}

// Kaydedilmemiş değişiklik varsa sor. true => devam et
async function confirmDiscard() {
  if (!state.dirty) return true;
  const r = await window.api.confirm(`"${fileName()}" dosyasında kaydedilmemiş değişiklikler var. Kaydedilsin mi?`);
  if (r === 2) return false;
  if (r === 0) {
    const ok = await save(false);
    return ok;
  }
  return true;
}

// ===================== Dosya işlemleri =====================
async function openFile() {
  if (!(await confirmDiscard())) return;
  const res = await window.api.openFile();
  if (!res) return;
  if (res.error) { showStatus(res.error, 4000); return; }
  state.path = res.path;
  if (res.type === 'word') {
    editor.innerHTML = res.html || '<p><br></p>';
    showScreen('word');
    editor.focus();
  } else {
    loadSheets(res.sheets);
    showScreen('excel');
  }
  setDirty(false);
  showStatus(`Açıldı: ${fileName()}`);
}

async function newWord() {
  if (!(await confirmDiscard())) return;
  state.path = null;
  editor.innerHTML = '<p><br></p>';
  showScreen('word');
  setDirty(false);
  editor.focus();
}

async function newExcel() {
  if (!(await confirmDiscard())) return;
  state.path = null;
  loadSheets([{ name: 'Sayfa1', data: [] }]);
  showScreen('excel');
  setDirty(false);
}

async function save(saveAs) {
  if (state.mode === 'start') return false;
  const payload = { type: state.mode, path: state.path, saveAs };
  if (state.mode === 'word') payload.blocks = htmlToBlocks(editor);
  else payload.sheets = trimmedSheets();
  const res = await window.api.saveFile(payload);
  if (!res) return false;
  if (res.error) { showStatus(res.error, 4000); return false; }
  state.path = res.path;
  setDirty(false);
  showStatus(`Kaydedildi: ${fileName()}`);
  return true;
}

async function goHome() {
  if (!(await confirmDiscard())) return;
  state.path = null;
  setDirty(false);
  showScreen('start');
}

// ===================== Word editörü =====================

// Düzenleyicinin HTML içeriğini basit paragraf bloklarına çevirir.
function htmlToBlocks(root) {
  const blocks = [];

  function pushBlock(el, opts) {
    const runs = [];
    collectRuns(el, { bold: false, italic: false, underline: false, strike: false }, runs);
    // Boş paragrafları da koru (satır aralığı)
    blocks.push({ runs: runs.filter((r) => r.text !== ''), ...opts });
  }

  function collectRuns(node, fmt, runs) {
    node.childNodes.forEach((ch) => {
      if (ch.nodeType === Node.TEXT_NODE) {
        const text = ch.textContent.replace(/ /g, ' ');
        if (text) runs.push({ text, ...fmt });
        return;
      }
      if (ch.nodeType !== Node.ELEMENT_NODE) return;
      const tag = ch.tagName.toLowerCase();
      if (tag === 'br') { runs.push({ text: '\n', ...fmt }); return; }
      const cs = ch.style || {};
      const f = {
        bold: fmt.bold || tag === 'b' || tag === 'strong' || /^(bold|[6-9]00)$/.test(cs.fontWeight || ''),
        italic: fmt.italic || tag === 'i' || tag === 'em' || cs.fontStyle === 'italic',
        underline: fmt.underline || tag === 'u' || (cs.textDecoration || '').includes('underline'),
        strike: fmt.strike || tag === 's' || tag === 'strike' || tag === 'del' || (cs.textDecoration || '').includes('line-through'),
      };
      collectRuns(ch, f, runs);
    });
  }

  function alignOf(el) {
    const a = (el.style && el.style.textAlign) || el.getAttribute('align') || '';
    if (a === 'center' || a === 'right' || a === 'justify') return a;
    return 'left';
  }

  function walk(node, inherited) {
    node.childNodes.forEach((ch) => {
      if (ch.nodeType === Node.TEXT_NODE) {
        if (ch.textContent.trim()) {
          // Blok dışında kalan serbest metin
          const runs = [{ text: ch.textContent, bold: false, italic: false, underline: false, strike: false }];
          blocks.push({ runs, ...inherited });
        }
        return;
      }
      if (ch.nodeType !== Node.ELEMENT_NODE) return;
      const tag = ch.tagName.toLowerCase();
      if (tag === 'h1' || tag === 'h2' || tag === 'h3') {
        pushBlock(ch, { heading: Number(tag[1]), align: alignOf(ch), ...(inherited.list ? { list: inherited.list } : {}) });
      } else if (tag === 'ul' || tag === 'ol') {
        walk(ch, { list: tag === 'ul' ? 'bullet' : 'number' });
      } else if (tag === 'li') {
        // li içinde blok var mı?
        const hasBlock = [...ch.children].some((c) => /^(p|div|h[1-6]|ul|ol)$/i.test(c.tagName));
        if (hasBlock) walk(ch, inherited);
        else pushBlock(ch, { align: alignOf(ch), list: inherited.list });
      } else if (tag === 'table') {
        ch.querySelectorAll('tr').forEach((tr) => {
          const cells = [...tr.children].map((td) => td.textContent.trim());
          blocks.push({ runs: [{ text: cells.join('\t'), bold: false, italic: false, underline: false, strike: false }] });
        });
      } else if (tag === 'p' || tag === 'div' || tag === 'blockquote' || /^h[4-6]$/.test(tag)) {
        const hasBlock = [...ch.children].some((c) => /^(p|div|h[1-6]|ul|ol|table)$/i.test(c.tagName));
        if (hasBlock) walk(ch, { ...inherited, align: alignOf(ch) });
        else pushBlock(ch, { align: alignOf(ch), ...(inherited.list ? { list: inherited.list } : {}) });
      } else if (tag === 'br') {
        blocks.push({ runs: [] });
      } else {
        // span, b, i, a vs. doğrudan kökte: tek paragraf olarak al
        pushBlock(ch, { align: inherited.align || 'left', ...(inherited.list ? { list: inherited.list } : {}) });
      }
    });
  }

  walk(root, {});

  // Satır sonları (\n) içeren run'ları ayrı paragraflara böl
  const out = [];
  blocks.forEach((b) => {
    if (!b.runs.some((r) => r.text.includes('\n'))) { out.push(b); return; }
    let cur = { ...b, runs: [] };
    b.runs.forEach((r) => {
      const parts = r.text.split('\n');
      parts.forEach((p, i) => {
        if (i > 0) { out.push(cur); cur = { ...b, runs: [] }; }
        if (p) cur.runs.push({ ...r, text: p });
      });
    });
    out.push(cur);
  });
  return out;
}

function exec(cmd) {
  editor.focus();
  document.execCommand(cmd, false, null);
  refreshToolbar();
}

function refreshToolbar() {
  document.querySelectorAll('#word [data-cmd]').forEach((btn) => {
    const cmd = btn.dataset.cmd;
    if (['undo', 'redo'].includes(cmd)) return;
    let on = false;
    try { on = document.queryCommandState(cmd); } catch { /* yoksay */ }
    btn.classList.toggle('active', on);
  });
  // Paragraf stili
  let block = 'p';
  try {
    const v = (document.queryCommandValue('formatBlock') || '').toLowerCase();
    if (/^h[1-3]$/.test(v)) block = v;
  } catch { /* yoksay */ }
  $('#blockStyle').value = block;
}

$('#blockStyle').addEventListener('change', (e) => {
  editor.focus();
  document.execCommand('formatBlock', false, e.target.value);
  refreshToolbar();
});

editor.addEventListener('input', () => setDirty(true));
editor.addEventListener('keyup', refreshToolbar);
editor.addEventListener('mouseup', refreshToolbar);
editor.addEventListener('paste', (e) => {
  // Düz metin yapıştır (dış biçimlendirme karmaşasını engelle)
  e.preventDefault();
  const text = (e.clipboardData || window.clipboardData).getData('text/plain');
  document.execCommand('insertText', false, text);
});

// ===================== Excel editörü =====================

function colName(c) {
  let s = '';
  c += 1;
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

function loadSheets(sheets) {
  state.sheets = (sheets && sheets.length ? sheets : [{ name: 'Sayfa1', data: [] }]).map((s) => ({
    name: s.name,
    data: (s.data || []).map((row) => row.map((v) => (v === null || v === undefined ? '' : String(v)))),
  }));
  state.activeSheet = 0;
  state.sel = { r: 0, c: 0 };
  renderTabs();
  renderGrid();
}

function currentData() { return state.sheets[state.activeSheet].data; }

function dims() {
  const data = currentData();
  const rows = Math.max(MIN_ROWS, data.length + 5);
  const cols = Math.max(MIN_COLS, ...data.map((r) => r.length), 0) + 3;
  return { rows, cols: Math.max(cols, MIN_COLS) };
}

function getCell(r, c) {
  const data = currentData();
  return (data[r] && data[r][c]) || '';
}

function setCell(r, c, v) {
  const data = currentData();
  while (data.length <= r) data.push([]);
  while (data[r].length <= c) data[r].push('');
  if (data[r][c] !== v) { data[r][c] = v; setDirty(true); }
}

function renderGrid() {
  const { rows, cols } = dims();
  const frag = document.createDocumentFragment();

  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  const corner = document.createElement('th');
  corner.className = 'row-head corner';
  hr.appendChild(corner);
  for (let c = 0; c < cols; c++) {
    const th = document.createElement('th');
    th.textContent = colName(c);
    th.dataset.c = c;
    hr.appendChild(th);
  }
  thead.appendChild(hr);
  frag.appendChild(thead);

  const tbody = document.createElement('tbody');
  for (let r = 0; r < rows; r++) {
    const tr = document.createElement('tr');
    const rh = document.createElement('th');
    rh.className = 'row-head';
    rh.textContent = r + 1;
    rh.dataset.r = r;
    tr.appendChild(rh);
    for (let c = 0; c < cols; c++) {
      const td = document.createElement('td');
      td.contentEditable = 'true';
      td.dataset.r = r;
      td.dataset.c = c;
      const v = getCell(r, c);
      td.textContent = v;
      if (v !== '' && !isNaN(Number(String(v).replace(',', '.'))) && !String(v).startsWith('=')) td.classList.add('num');
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  frag.appendChild(tbody);
  grid.innerHTML = '';
  grid.appendChild(frag);
  highlightSel();
}

function renderTabs() {
  sheetTabs.innerHTML = '';
  state.sheets.forEach((s, i) => {
    const b = document.createElement('button');
    b.textContent = s.name;
    b.classList.toggle('active', i === state.activeSheet);
    b.addEventListener('click', () => {
      state.activeSheet = i;
      state.sel = { r: 0, c: 0 };
      renderTabs();
      renderGrid();
    });
    sheetTabs.appendChild(b);
  });
}

function cellEl(r, c) {
  return grid.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
}

function highlightSel() {
  grid.querySelectorAll('th.hl').forEach((th) => th.classList.remove('hl'));
  const { r, c } = state.sel;
  const ch = grid.querySelector(`thead th[data-c="${c}"]`);
  const rh = grid.querySelector(`th.row-head[data-r="${r}"]`);
  if (ch) ch.classList.add('hl');
  if (rh) rh.classList.add('hl');
  const name = `${colName(c)}${r + 1}`;
  $('#cellName').textContent = name;
  $('#cellInfo').textContent = name;
  if (document.activeElement !== formulaInput) formulaInput.value = getCell(r, c);
}

function selectCell(r, c, focus = true) {
  const { rows, cols } = dims();
  r = Math.max(0, Math.min(rows - 1, r));
  c = Math.max(0, Math.min(cols - 1, c));
  state.sel = { r, c };
  highlightSel();
  const td = cellEl(r, c);
  if (td && focus) {
    td.focus();
    // İmleci sona taşı
    const range = document.createRange();
    range.selectNodeContents(td);
    range.collapse(false);
    const s = window.getSelection();
    s.removeAllRanges();
    s.addRange(range);
    td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
}

function commitCell(td) {
  const r = Number(td.dataset.r), c = Number(td.dataset.c);
  const v = td.textContent.replace(/\n/g, '');
  td.textContent = v;
  setCell(r, c, v);
  td.classList.toggle('num', v !== '' && !isNaN(Number(v.replace(',', '.'))) && !v.startsWith('='));
  if (state.sel.r === r && state.sel.c === c) formulaInput.value = v;
}

grid.addEventListener('focusin', (e) => {
  const td = e.target.closest('td');
  if (!td) return;
  state.sel = { r: Number(td.dataset.r), c: Number(td.dataset.c) };
  highlightSel();
});

grid.addEventListener('input', (e) => {
  const td = e.target.closest('td');
  if (td) commitCell(td);
});

grid.addEventListener('keydown', (e) => {
  const td = e.target.closest('td');
  if (!td) return;
  const r = Number(td.dataset.r), c = Number(td.dataset.c);
  if (e.key === 'Enter') { e.preventDefault(); commitCell(td); selectCell(e.shiftKey ? r - 1 : r + 1, c); }
  else if (e.key === 'Tab') { e.preventDefault(); commitCell(td); selectCell(r, e.shiftKey ? c - 1 : c + 1); }
  else if (e.key === 'ArrowDown') { e.preventDefault(); selectCell(r + 1, c); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); selectCell(r - 1, c); }
  else if (e.key === 'ArrowLeft' && isCaretAtStart(td)) { e.preventDefault(); selectCell(r, c - 1); }
  else if (e.key === 'ArrowRight' && isCaretAtEnd(td)) { e.preventDefault(); selectCell(r, c + 1); }
  else if (e.key === 'Escape') { td.textContent = getCell(r, c); selectCell(r, c); }
  else if (e.key === 'Delete' && window.getSelection().toString() === '' && td.textContent === '') { /* boş */ }
});

function isCaretAtStart(td) {
  const s = window.getSelection();
  if (!s.rangeCount) return true;
  const range = s.getRangeAt(0).cloneRange();
  range.selectNodeContents(td);
  range.setEnd(s.anchorNode, s.anchorOffset);
  return range.toString().length === 0;
}
function isCaretAtEnd(td) {
  const s = window.getSelection();
  if (!s.rangeCount) return true;
  const range = s.getRangeAt(0).cloneRange();
  range.selectNodeContents(td);
  range.setStart(s.focusNode, s.focusOffset);
  return range.toString().length === 0;
}

// Hücreye yapıştırmada çok satırlı veriyi tabloya dağıt
grid.addEventListener('paste', (e) => {
  const td = e.target.closest('td');
  if (!td) return;
  const text = (e.clipboardData || window.clipboardData).getData('text/plain');
  if (!text.includes('\n') && !text.includes('\t')) return; // normal yapıştır
  e.preventDefault();
  const r0 = Number(td.dataset.r), c0 = Number(td.dataset.c);
  const lines = text.replace(/\r/g, '').split('\n');
  if (lines[lines.length - 1] === '') lines.pop();
  lines.forEach((line, i) => line.split('\t').forEach((v, j) => setCell(r0 + i, c0 + j, v)));
  renderGrid();
  selectCell(r0, c0);
});

// Formül çubuğu
formulaInput.addEventListener('input', () => {
  const { r, c } = state.sel;
  setCell(r, c, formulaInput.value);
  const td = cellEl(r, c);
  if (td) td.textContent = formulaInput.value;
});
formulaInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); selectCell(state.sel.r + 1, state.sel.c); }
  if (e.key === 'Escape') { selectCell(state.sel.r, state.sel.c); }
});

function excelAction(action) {
  const data = currentData();
  const { r, c } = state.sel;
  switch (action) {
    case 'add-row':
      while (data.length <= r) data.push([]);
      data.splice(r + 1, 0, []);
      setDirty(true); renderGrid(); selectCell(r + 1, c);
      break;
    case 'add-col':
      data.forEach((row) => { while (row.length <= c) row.push(''); row.splice(c + 1, 0, ''); });
      setDirty(true); renderGrid(); selectCell(r, c + 1);
      break;
    case 'del-row':
      if (r < data.length) { data.splice(r, 1); setDirty(true); }
      renderGrid(); selectCell(r, c);
      break;
    case 'del-col':
      data.forEach((row) => { if (c < row.length) row.splice(c, 1); });
      setDirty(true); renderGrid(); selectCell(r, c);
      break;
    case 'add-sheet': {
      let n = state.sheets.length + 1;
      while (state.sheets.some((s) => s.name === `Sayfa${n}`)) n++;
      state.sheets.push({ name: `Sayfa${n}`, data: [] });
      state.activeSheet = state.sheets.length - 1;
      state.sel = { r: 0, c: 0 };
      setDirty(true); renderTabs(); renderGrid();
      break;
    }
    case 'rename-sheet': {
      const name = prompt('Sayfa adı:', state.sheets[state.activeSheet].name);
      if (name && name.trim()) { state.sheets[state.activeSheet].name = name.trim().slice(0, 31); setDirty(true); renderTabs(); }
      break;
    }
    case 'del-sheet':
      if (state.sheets.length <= 1) { showStatus('Son sayfa silinemez'); break; }
      if (!confirm(`"${state.sheets[state.activeSheet].name}" sayfası silinsin mi?`)) break;
      state.sheets.splice(state.activeSheet, 1);
      state.activeSheet = Math.max(0, state.activeSheet - 1);
      state.sel = { r: 0, c: 0 };
      setDirty(true); renderTabs(); renderGrid();
      break;
    default:
      break;
  }
}

// Kayıt için boş satır/sütunları kırp
function trimmedSheets() {
  return state.sheets.map((s) => {
    const rows = s.data.map((row) => {
      let end = row.length;
      while (end > 0 && (row[end - 1] === '' || row[end - 1] == null)) end--;
      return row.slice(0, end);
    });
    let end = rows.length;
    while (end > 0 && rows[end - 1].length === 0) end--;
    return { name: s.name, data: rows.slice(0, end) };
  });
}

// ===================== Olaylar =====================

document.addEventListener('click', (e) => {
  const btn = e.target.closest('button');
  if (!btn) return;
  if (btn.dataset.action) handleAction(btn.dataset.action);
  else if (btn.dataset.cmd) exec(btn.dataset.cmd);
  else if (btn.dataset.xl) excelAction(btn.dataset.xl);
});

// Araç çubuğu butonları editörün odağını çalmasın
document.querySelectorAll('.toolbar button').forEach((b) => b.addEventListener('mousedown', (e) => e.preventDefault()));

function handleAction(action) {
  switch (action) {
    case 'open': return openFile();
    case 'new-word': return newWord();
    case 'new-excel': return newExcel();
    case 'save': return save(false);
    case 'save-as': return save(true);
    case 'home': return goHome();
    default: return undefined;
  }
}

window.api.onMenu(handleAction);

document.addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  const k = e.key.toLowerCase();
  if (k === 's') { e.preventDefault(); save(e.shiftKey); }
  else if (k === 'o') { e.preventDefault(); openFile(); }
  else if (k === 'n') { e.preventDefault(); e.shiftKey ? newExcel() : newWord(); }
});

window.api.onRequestClose(async () => {
  if (await confirmDiscard()) window.api.closeNow();
});

showScreen('start');
