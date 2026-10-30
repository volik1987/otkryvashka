(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const el = {
    start: $('start'), view: $('view'), drop: $('drop'), file: $('file'),
    fileinfo: $('fileinfo'), badge: $('badge'), fname: $('fname'), tools: $('tools'),
    content: $('content'), scroller: $('scroller'), sheets: $('sheets'), note: $('note'),
    cellbar: $('cellbar'), addr: $('addr'), cval: $('cval'),
    q: $('q'), qcount: $('qcount'), zval: $('zval'),
    loading: $('loading'), dropmask: $('dropmask'), toast: $('toast')
  };

  const ROWS_CHUNK = 1000;
  const S = { reflow: matchMedia('(max-width: 720px)').matches, kind: null, name: '', wb: null, sheet: 0, rowsShown: 0, zoom: 1, fit: true, hits: [], hit: -1 };

  // ---------- общее ----------
  let toastTimer;
  function toast(msg, ms = 4500) {
    el.toast.textContent = msg;
    el.toast.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { el.toast.hidden = true; }, ms);
  }
  const esc = (s) => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const ext = (name) => (name.match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();

  function setPrintPage(css) {
    let st = $('pagecss');
    if (!st) { st = document.createElement('style'); st.id = 'pagecss'; document.head.appendChild(st); }
    st.textContent = '@media print{@page{' + css + '}}';
  }

  function showView(kind, name) {
    S.kind = kind; S.name = name;
    el.start.hidden = true; el.view.hidden = false;
    el.fileinfo.hidden = false; el.tools.hidden = false;
    const isSheet = kind === 'sheet';
    el.badge.textContent = { docx: 'DOCX', doc: 'DOC', sheet: (ext(name) || 'xlsx').toUpperCase() }[kind];
    el.badge.classList.toggle('x', isSheet);
    el.fname.textContent = name;
    el.fname.title = name;
    el.content.className = 'content ' + (isSheet ? 'sheet' : 'word');
    el.sheets.hidden = !isSheet;
    $('layout').hidden = isSheet;
    setLayout(S.reflow, true);
    el.cellbar.hidden = !isSheet;
    el.note.hidden = true;
    if (!isSheet) el.sheets.textContent = '';
    document.title = name + ' — Открывашка';
    setPrintPage(isSheet ? 'margin:12mm' : 'size:auto;margin:0');
    if (!history.state || !history.state.view) history.pushState({ view: 1 }, '');
  }

  function goHome() {
    S.kind = null; S.wb = null;
    el.content.textContent = '';
    el.sheets.textContent = '';
    el.view.hidden = true; el.start.hidden = false;
    el.fileinfo.hidden = true; el.tools.hidden = true;
    el.q.value = ''; el.qcount.textContent = '';
    el.file.value = '';
    document.title = 'Открывашка — просмотр Word и Excel онлайн';
  }

  // ---------- открытие ----------
  async function openFile(file) {
    if (!file) return;
    el.loading.hidden = false;
    // даём браузеру нарисовать «Открываю…» до тяжёлой работы
    await new Promise(r => requestAnimationFrame(() => setTimeout(r, 0)));
    try {
      const buf = await file.arrayBuffer();
      const u8 = new Uint8Array(buf);
      const kind = await detect(u8, file.name);
      el.q.value = ''; clearHits();
      if (kind === 'docx') await openDocx(buf, file.name);
      else if (kind === 'doc') openDoc(u8, file.name);
      else openSheet(u8, file.name);
    } catch (e) {
      console.error(e);
      toast(e && e.message ? e.message : 'Не получилось открыть файл.');
      if (S.kind === null) goHome();
    } finally {
      el.loading.hidden = true;
    }
  }

  async function detect(u8, name) {
    if (u8.length < 4) throw new Error('Файл пустой.');
    const sig = String.fromCharCode(u8[0], u8[1], u8[2], u8[3]);
    if (sig === 'PK\x03\x04') {
      const zip = await JSZip.loadAsync(u8);
      if (zip.file('word/document.xml')) return 'docx';
      if (zip.file('xl/workbook.xml') || zip.file('xl/workbook.bin') || zip.file('content.xml')) return 'sheet';
      if (zip.file('ppt/presentation.xml')) throw new Error('Это презентация PowerPoint — её Открывашка пока не показывает.');
      throw new Error('Это ZIP-архив, но не документ Word или Excel.');
    }
    if (u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0) {
      const cfb = XLSX.CFB.read(u8, { type: 'array' });
      if (WordDoc.isWord(cfb)) return 'doc';
      if (XLSX.CFB.find(cfb, 'Workbook') || XLSX.CFB.find(cfb, 'Book')) return 'sheet';
      if (XLSX.CFB.find(cfb, 'EncryptedPackage')) throw new Error('Файл защищён паролем — открыть его можно только в Office.');
      if (XLSX.CFB.find(cfb, 'PowerPoint Document')) throw new Error('Это презентация PowerPoint — её Открывашка пока не показывает.');
      throw new Error('Не удалось понять, что это за файл Office.');
    }
    if (sig === '{\\rt') throw new Error('Это файл RTF, хоть и с расширением .' + ext(name) + '. Откройте его в Word или TextEdit.');
    if (sig === '%PDF') throw new Error('Это PDF — его браузер умеет открывать сам.');
    const e = ext(name);
    if (/^(docx?|dotx?|docm)$/.test(e)) throw new Error('Файл повреждён или это не документ Word.');
    return 'sheet';   // csv, xml-таблицы, html под видом .xls — пусть разбирает SheetJS
  }

  // ---------- Word ----------
  async function openDocx(buf, name) {
    showView('docx', name);
    el.content.textContent = '';
    await docx.renderAsync(buf, el.content, null, {
      className: 'docx', inWrapper: true, breakPages: true, ignoreLastRenderedPageBreak: true,
      renderHeaders: true, renderFooters: true, renderFootnotes: true, renderEndnotes: true,
      useBase64URL: true, experimental: true
    });
    S.fit = !S.reflow; S.zoom = 1; applyZoom();
    el.scroller.scrollTo(0, 0);
  }

  function openDoc(u8, name) {
    showView('doc', name);
    WordDoc.render(u8, el.content);
    el.note.textContent = 'Старый формат .doc показан упрощённо: текст и таблицы, без оформления и картинок.';
    el.note.hidden = false;
    S.fit = !S.reflow; S.zoom = 1; applyZoom();
    el.scroller.scrollTo(0, 0);
  }

  // ---------- Excel ----------
  function decodeText(u8) {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(u8); }
    catch (e) { return new TextDecoder('windows-1251').decode(u8); }
  }

  function openSheet(u8, name) {
    const opts = { cellDates: true, cellStyles: true, cellNF: true, dense: false };
    let wb;
    const e = ext(name);
    if (e === 'csv' || e === 'txt' || e === 'tsv') wb = XLSX.read(decodeText(u8), Object.assign({ type: 'string' }, opts));
    else wb = XLSX.read(u8, Object.assign({ type: 'array' }, opts));
    if (!wb.SheetNames.length) throw new Error('В книге нет листов.');
    showView('sheet', name);
    S.wb = wb;
    const hidden = (wb.Workbook && wb.Workbook.Sheets) || [];
    el.sheets.innerHTML = wb.SheetNames.map((n, i) =>
      '<button data-i="' + i + '" class="' + (hidden[i] && hidden[i].Hidden ? 'hid' : '') + '" title="' +
      esc(n) + (hidden[i] && hidden[i].Hidden ? ' (скрытый лист)' : '') + '">' + esc(n) + '</button>').join('');
    let first = hidden.findIndex(s => s && !s.Hidden);
    if (first < 0 || !hidden.length) first = 0;
    S.zoom = 1; S.fit = false;
    showSheet(first);
  }

  function usedRange(ws) {
    if (!ws['!ref']) return null;
    const r = XLSX.utils.decode_range(ws['!ref']);
    // реальная граница по заполненным ячейкам: !ref бывает раздут форматированием
    let maxR = -1, maxC = -1;
    for (const k in ws) {
      if (k[0] === '!') continue;
      const c = ws[k];
      if (c.v === undefined || c.v === null || c.v === '') continue;
      const a = XLSX.utils.decode_cell(k);
      if (a.r > maxR) maxR = a.r;
      if (a.c > maxC) maxC = a.c;
    }
    (ws['!merges'] || []).forEach(m => { if (m.e.r > maxR && m.s.r <= maxR) maxR = m.e.r; if (m.e.c > maxC && m.s.c <= maxC) maxC = m.e.c; });
    if (maxR < 0) return null;
    return { s: { r: Math.min(r.s.r, maxR), c: Math.min(r.s.c, maxC) }, e: { r: maxR, c: maxC } };
  }

  function colWidth(ci) {
    if (!ci) return 72;
    if (ci.wpx) return Math.round(ci.wpx);
    if (ci.wch) return Math.round(ci.wch * 7 + 5);
    if (ci.width) return Math.round(ci.width * 7);
    return 72;
  }

  function showSheet(i) {
    S.sheet = i;
    [...el.sheets.children].forEach(b => b.classList.toggle('on', +b.dataset.i === i));
    const btn = el.sheets.children[i];
    if (btn) btn.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    el.addr.textContent = '—'; el.cval.textContent = 'нажмите на ячейку, чтобы увидеть значение или формулу';
    const ws = S.wb.Sheets[S.wb.SheetNames[i]];
    const rng = ws && usedRange(ws);
    el.scroller.scrollTo(0, 0);
    if (!rng) {
      el.content.innerHTML = '<div class="emptysheet">Лист пустой' + (ws && ws['!type'] === 'chart' ? ' (это диаграмма — её Открывашка не рисует)' : '') + '</div>';
      S.rowsShown = 0;
      runSearch(true);
      return;
    }
    S.rng = rng;
    const cols = ws['!cols'] || [];
    S.visCols = [];
    for (let c = rng.s.c; c <= rng.e.c; c++) if (!(cols[c] && cols[c].hidden)) S.visCols.push(c);

    // склеенные ячейки
    S.mergeStart = new Map(); S.covered = new Set();
    (ws['!merges'] || []).forEach(m => {
      let cs = 0;
      for (let c = m.s.c; c <= m.e.c; c++) if (!(cols[c] && cols[c].hidden)) cs++;
      S.mergeStart.set(m.s.r + ',' + m.s.c, { rs: m.e.r - m.s.r + 1, cs: Math.max(cs, 1) });
      for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) if (r !== m.s.r || c !== m.s.c) S.covered.add(r + ',' + c);
    });

    let head = '<colgroup><col style="width:46px">' + S.visCols.map(c => '<col style="width:' + colWidth(cols[c]) + 'px">').join('') + '</colgroup>';
    head += '<thead><tr><th class="corner"></th>' + S.visCols.map(c => '<th>' + XLSX.utils.encode_col(c) + '</th>').join('') + '</tr></thead>';
    const total = S.visCols.reduce((a, c) => a + colWidth(cols[c]), 46);
    el.content.innerHTML = '<table class="grid" style="width:' + total + 'px">' + head + '<tbody></tbody></table>';
    S.tbody = el.content.querySelector('tbody');
    S.ws = ws;
    S.rowsShown = 0;
    S.nextRow = rng.s.r;
    appendRows(ROWS_CHUNK);
    applyZoom();
    runSearch(true);
  }

  // Даты и числа — по-русски: 15.01.2026, 2 990,5
  function fmtDate(d) {
    d = new Date(Math.round(d.getTime() / 60000) * 60000);   // SheetJS даёт даты со сдвигом в секунды
    const hasTime = d.getHours() || d.getMinutes() || d.getSeconds();
    return hasTime ? d.toLocaleString('ru-RU', { dateStyle: 'short', timeStyle: 'short' }).replace(',', '') : d.toLocaleDateString('ru-RU');
  }
  function cellText(cell) {
    if (cell.v instanceof Date) return /^[\[\]hHmsS:AMP\/ ]+$/.test(cell.z || '') && cell.w ? cell.w : fmtDate(cell.v);
    if (cell.t === 'n' && typeof cell.w === 'string' && /^[^\d]*[\d,]*\d(\.\d+)?[^\d]*$/.test(cell.w)) {
      return cell.w.replace(/[,.]/g, ch => ch === ',' ? '\u202F' : ',');
    }
    if (cell.w !== undefined) return cell.w;
    return cell.v === undefined || cell.v === null ? '' : String(cell.v);
  }

  function cellHtml(ws, r, c) {
    const key = r + ',' + c;
    if (S.covered.has(key)) return '';
    const a = XLSX.utils.encode_cell({ r, c });
    const cell = ws[a];
    const m = S.mergeStart.get(key);
    let attrs = ' data-a="' + a + '"';
    if (m) { if (m.rs > 1) attrs += ' rowspan="' + m.rs + '"'; if (m.cs > 1) attrs += ' colspan="' + m.cs + '"'; }
    if (!cell) return '<td' + attrs + '></td>';
    let text = cellText(cell);
    const cls = [];
    if (cell.t === 'n' || cell.t === 'd') cls.push('n');
    else if (cell.t === 'b') { cls.push('b'); text = cell.v ? 'ИСТИНА' : 'ЛОЖЬ'; }
    else if (cell.t === 'e') cls.push('e');
    if (/\n/.test(text)) cls.push('wrap');
    if (cls.length) attrs += ' class="' + cls.join(' ') + '"';
    const s = cell.s;
    if (s && s.patternType === 'solid' && s.fgColor && s.fgColor.rgb) {
      const rgb = s.fgColor.rgb.slice(-6);
      if (/^[0-9a-f]{6}$/i.test(rgb) && rgb.toUpperCase() !== 'FFFFFF') attrs += ' style="background:#' + rgb + ';color:#111"';
    }
    if (text.length > 40) attrs += ' title="' + esc(text.slice(0, 500)) + '"';
    return '<td' + attrs + '>' + esc(text) + '</td>';
  }

  function appendRows(n) {
    const ws = S.ws, rng = S.rng, rows = ws['!rows'] || [];
    const old = S.tbody.querySelector('tr.more');
    if (old) old.remove();
    let html = '';
    let added = 0;
    while (S.nextRow <= rng.e.r && added < n) {
      const r = S.nextRow++;
      const ri = rows[r];
      if (ri && ri.hidden) continue;
      const h = ri && ri.hpx ? ' style="height:' + Math.round(ri.hpx) + 'px"' : '';
      html += '<tr' + h + '><th>' + (r + 1) + '</th>' + S.visCols.map(c => cellHtml(ws, r, c)).join('') + '</tr>';
      added++;
    }
    S.rowsShown += added;
    if (S.nextRow <= rng.e.r) {
      const left = rng.e.r - S.nextRow + 1;
      html += '<tr class="more"><td colspan="' + (S.visCols.length + 1) + '"><button type="button">Показать ещё ' +
        Math.min(left, ROWS_CHUNK).toLocaleString('ru-RU') + ' строк (осталось ' + left.toLocaleString('ru-RU') + ')</button></td></tr>';
    }
    S.tbody.insertAdjacentHTML('beforeend', html);
  }

  el.content.addEventListener('click', (e) => {
    if (S.kind !== 'sheet') return;
    const more = e.target.closest('tr.more button');
    if (more) { appendRows(ROWS_CHUNK); if (el.q.value) runSearch(true); return; }
    const td = e.target.closest('td[data-a]');
    if (!td) return;
    const prev = el.content.querySelector('td.on');
    if (prev) prev.classList.remove('on');
    td.classList.add('on');
    const a = td.dataset.a, cell = S.ws[a];
    el.addr.textContent = a;
    el.cval.textContent = !cell ? '' : cell.f ? '=' + cell.f : cell.v instanceof Date ? fmtDate(cell.v) : cell.t === 'n' ? String(cell.v).replace('.', ',') : String(cell.v === undefined ? '' : cell.v);
    el.cval.title = el.cval.textContent;
  });

  el.sheets.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-i]');
    if (b) showSheet(+b.dataset.i);
  });

  // ---------- масштаб ----------
  function fitZoom() {
    const page = el.content.querySelector('section.docx, .docpage');
    if (!page) return 1;
    const w = page.offsetWidth + 32;
    return Math.min(1, Math.max(0.3, el.scroller.clientWidth / w));
  }
  function applyZoom() {
    if (S.fit && S.kind !== 'sheet' && !S.reflow) S.zoom = fitZoom();
    el.content.style.zoom = S.zoom;
    el.zval.textContent = Math.round(S.zoom * 100) + '%';
  }
  function setLayout(reflow, quiet) {
    S.reflow = reflow;
    el.content.classList.toggle('reflow', reflow && S.kind !== 'sheet');
    [...$('layout').children].forEach(b => b.classList.toggle('on', (b.dataset.m === 'text') === reflow));
    if (quiet) return;
    S.fit = !reflow;
    if (reflow) S.zoom = 1;
    applyZoom();
    if (el.q.value) runSearch();
  }
  $('layout').addEventListener('click', (e) => {
    const b = e.target.closest('button[data-m]');
    if (b) setLayout(b.dataset.m === 'text');
  });

  function stepZoom(dir) {
    const steps = [0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];
    const next = dir > 0
      ? (steps.find(s => s > S.zoom + 0.001) || steps[steps.length - 1])
      : ([...steps].reverse().find(s => s < S.zoom - 0.001) || steps[0]);
    S.fit = false; S.zoom = next; applyZoom();
  }
  $('zin').onclick = () => stepZoom(1);
  $('zout').onclick = () => stepZoom(-1);
  el.zval.onclick = () => {
    if (S.kind === 'sheet' || S.reflow) { S.zoom = 1; S.fit = false; } else S.fit = true;
    applyZoom();
  };
  let rt;
  window.addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(() => { if (S.fit && S.kind) applyZoom(); }, 120); });

  // ---------- поиск ----------
  function clearHits() {
    const parents = new Set();
    el.content.querySelectorAll('mark.hit').forEach(m => {
      parents.add(m.parentNode);
      m.replaceWith(document.createTextNode(m.textContent));
    });
    parents.forEach(p => p.normalize());
    S.hits = []; S.hit = -1;
  }

  function markAll(q) {
    const ql = q.toLocaleLowerCase('ru');
    const walker = document.createTreeWalker(el.content, NodeFilter.SHOW_TEXT, {
      acceptNode: (n) => n.parentNode.closest('style,script,thead,tr.more') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT
    });
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    const hits = [];
    for (const node of nodes) {
      const t = node.nodeValue, tl = t.toLocaleLowerCase('ru');
      let i = tl.indexOf(ql);
      if (i < 0) continue;
      const frag = document.createDocumentFragment();
      let last = 0;
      while (i >= 0) {
        if (i > last) frag.appendChild(document.createTextNode(t.slice(last, i)));
        const m = document.createElement('mark');
        m.className = 'hit';
        m.textContent = t.slice(i, i + q.length);
        frag.appendChild(m); hits.push(m);
        last = i + q.length;
        i = tl.indexOf(ql, last);
      }
      if (last < t.length) frag.appendChild(document.createTextNode(t.slice(last)));
      node.replaceWith(frag);
    }
    return hits;
  }

  // совпадения на других листах — по данным книги, а не по таблице на экране
  function sheetCounts(q) {
    const ql = q.toLocaleLowerCase('ru');
    return S.wb.SheetNames.map(n => {
      const ws = S.wb.Sheets[n];
      let k = 0;
      for (const a in ws) {
        if (a[0] === '!') continue;
        const c = ws[a];
        const t = cellText(c);
        if (!t) continue;
        const tl = String(t).toLocaleLowerCase('ru');
        let i = tl.indexOf(ql);
        while (i >= 0) { k++; i = tl.indexOf(ql, i + ql.length); }
      }
      return k;
    });
  }

  function runSearch() {
    clearHits();
    const q = el.q.value.trim();
    if (S.kind === 'sheet' && S.wb) {
      const counts = q ? sheetCounts(q) : [];
      [...el.sheets.children].forEach((b, i) => {
        const n = S.wb.SheetNames[i];
        b.textContent = counts[i] ? n + ' · ' + counts[i] : n;
      });
    }
    if (!q) { el.qcount.textContent = ''; return; }
    S.hits = markAll(q);
    if (!S.hits.length) {
      el.qcount.textContent = S.kind === 'sheet' && S.wb && sheetCounts(q).some(Boolean) ? 'на других листах' : 'нет';
      return;
    }
    goHit(0);
  }

  function goHit(i) {
    if (!S.hits.length) return;
    if (S.hit >= 0 && S.hits[S.hit]) S.hits[S.hit].classList.remove('cur');
    S.hit = (i + S.hits.length) % S.hits.length;
    const m = S.hits[S.hit];
    m.classList.add('cur');
    m.scrollIntoView({ block: 'center', inline: 'center' });
    el.qcount.textContent = (S.hit + 1) + ' из ' + S.hits.length;
  }

  let st;
  el.q.addEventListener('input', () => { clearTimeout(st); st = setTimeout(runSearch, 250); });
  el.q.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      clearTimeout(st);
      if (!S.hits.length || el.q.value.trim() === '') { runSearch(); return; }
      goHit(S.hit + (e.shiftKey ? -1 : 1));
    } else if (e.key === 'Escape') {
      el.q.value = ''; runSearch(); el.q.blur();
    }
  });
  $('qnext').onclick = () => goHit(S.hit + 1);
  $('qprev').onclick = () => goHit(S.hit - 1);

  // ---------- ввод файлов ----------
  el.file.addEventListener('change', () => openFile(el.file.files[0]));
  $('open2').onclick = () => el.file.click();
  $('home').onclick = () => { if (S.kind) history.back(); };
  $('print').onclick = () => window.print();
  window.addEventListener('popstate', () => { if (S.kind) goHome(); });

  let dragDepth = 0;
  const hasFiles = (e) => e.dataTransfer && [...e.dataTransfer.types].includes('Files');
  window.addEventListener('dragenter', (e) => { if (!hasFiles(e)) return; e.preventDefault(); dragDepth++; el.dropmask.hidden = false; el.drop.classList.add('over'); });
  window.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; el.dropmask.hidden = true; el.drop.classList.remove('over'); } });
  window.addEventListener('drop', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault(); dragDepth = 0;
    el.dropmask.hidden = true; el.drop.classList.remove('over');
    openFile(e.dataTransfer.files[0]);
  });
  window.addEventListener('paste', (e) => {
    const f = e.clipboardData && e.clipboardData.files && e.clipboardData.files[0];
    if (f) { e.preventDefault(); openFile(f); }
  });
  window.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    if (!mod) return;
    if (e.key === 'o' || e.key === 'O' || e.key === 'щ' || e.key === 'Щ') { e.preventDefault(); el.file.click(); }
    else if (S.kind && (e.key === 'f' || e.key === 'а')) { e.preventDefault(); el.q.focus(); el.q.select(); }
    else if (S.kind && (e.key === '=' || e.key === '+')) { e.preventDefault(); stepZoom(1); }
    else if (S.kind && e.key === '-') { e.preventDefault(); stepZoom(-1); }
    else if (S.kind && e.key === '0') { e.preventDefault(); el.zval.click(); }
  });

  // открытие из системы, когда приложение установлено (file_handlers)
  if ('launchQueue' in window) {
    window.launchQueue.setConsumer(async (params) => {
      if (params.files && params.files.length) openFile(await params.files[0].getFile());
    });
  }

  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }

  window.__otk = { S, openFile };
})();
