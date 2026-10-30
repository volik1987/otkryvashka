// Чтение старого Word (.doc, Word 97–2003): достаём текст основного документа
// через таблицу кусков (piece table) и собираем абзацы и таблицы.
// Форматирование и картинки не разбираем — только содержимое.
(function () {
  'use strict';

  function u16(b, o) { return b[o] | (b[o + 1] << 8); }
  function u32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }

  function streamBytes(cfb, name) {
    const e = XLSX.CFB.find(cfb, name);
    if (!e || !e.content) return null;
    return e.content instanceof Uint8Array ? e.content : Uint8Array.from(e.content);
  }

  const cp1252 = new TextDecoder('windows-1252');
  const utf16 = new TextDecoder('utf-16le');

  // Возвращает текст основного документа со служебными символами Word.
  function extractText(cfb) {
    const wd = streamBytes(cfb, 'WordDocument');
    if (!wd) throw new Error('В файле нет потока WordDocument — это не документ Word.');
    if (u16(wd, 0) !== 0xA5EC) throw new Error('Неизвестная сигнатура документа Word.');
    const nFib = u16(wd, 2);
    const flags = u16(wd, 0x0A);
    if (flags & 0x0100) throw new Error('Документ защищён паролем — открыть его можно только в Word.');
    if (nFib < 101) throw new Error('Это документ Word 6/95 — такой старый формат не поддерживается.');

    const csw = u16(wd, 32);
    const lwPos = 34 + csw * 2 + 2;          // начало FibRgLw97
    const cslw = u16(wd, lwPos - 2);
    const ccpText = u32(wd, lwPos + 3 * 4);   // длина основного текста в символах
    const fcPos = lwPos + cslw * 4 + 2;       // начало FibRgFcLcb97
    const fcClx = u32(wd, fcPos + 33 * 8);
    const lcbClx = u32(wd, fcPos + 33 * 8 + 4);

    const tbl = streamBytes(cfb, (flags & 0x0200) ? '1Table' : '0Table');
    if (!tbl || !lcbClx) throw new Error('Не нашлась таблица текста документа.');

    let i = fcClx;
    while (tbl[i] === 0x01) i += 3 + u16(tbl, i + 1);   // пропускаем Prc
    if (tbl[i] !== 0x02) throw new Error('Повреждена таблица текста документа.');
    const lcb = u32(tbl, i + 1);
    const plc = i + 5;
    const n = (lcb - 4) / 12;

    let out = '';
    const marks = [];   // свойства абзацев для каждого \r и \x07 по порядку
    let papx = null;
    try { papx = readPapx(wd, tbl, fcPos); } catch (e) { papx = null; }
    for (let k = 0; k < n && out.length < ccpText; k++) {
      const cpStart = u32(tbl, plc + k * 4);
      const cpEnd = u32(tbl, plc + (k + 1) * 4);
      const pcd = plc + (n + 1) * 4 + k * 8;
      const fcRaw = u32(tbl, pcd + 2);
      const len = cpEnd - cpStart;
      if (len <= 0) continue;
      const compressed = !!(fcRaw & 0x40000000);
      const fc0 = compressed ? (fcRaw & 0x3FFFFFFF) / 2 : fcRaw;
      const piece = compressed ? cp1252.decode(wd.subarray(fc0, fc0 + len)) : utf16.decode(wd.subarray(fc0, fc0 + len * 2));
      if (papx) {
        for (let j = 0; j < piece.length && out.length + j < ccpText; j++) {
          const ch = piece[j];
          if (ch === '\r' || ch === '\x07') marks.push(papx(fc0 + j * (compressed ? 1 : 2)));
        }
      }
      out += piece;
    }
    return { text: out.slice(0, ccpText), marks: papx ? marks : null };
  }

  // Свойства абзацев из PAPX FKP: функция fc → { inTable, ttp, jc, istd }.
  function readPapx(wd, tbl, fcPos) {
    const fcBte = u32(wd, fcPos + 13 * 8), lcbBte = u32(wd, fcPos + 13 * 8 + 4);
    if (!lcbBte) return null;
    const nb = (lcbBte - 4) / 8;
    const runs = [];   // [fcStart, fcEnd, props]
    for (let i = 0; i < nb; i++) {
      const pn = u32(tbl, fcBte + (nb + 1) * 4 + i * 4) & 0x3FFFFF;
      const page = pn * 512;
      const crun = wd[page + 511];
      for (let r = 0; r < crun; r++) {
        const fa = u32(wd, page + r * 4), fb = u32(wd, page + (r + 1) * 4);
        const bOff = wd[page + (crun + 1) * 4 + r * 13];
        const props = { inTable: false, ttp: false, jc: 0, istd: 0 };
        if (bOff) {
          let p = page + bOff * 2;
          let cb = wd[p++];
          let size;
          if (cb === 0) { size = wd[p++] * 2; } else { size = cb * 2 - 1; }
          const end = p + size;
          props.istd = u16(wd, p); p += 2;
          while (p + 2 <= end) {
            const op = u16(wd, p); p += 2;
            const spra = op >> 13;
            let len;
            if (spra === 0 || spra === 1) len = 1;
            else if (spra === 2 || spra === 4 || spra === 5) len = 2;
            else if (spra === 3) len = 4;
            else if (spra === 7) len = 3;
            else if (op === 0xD608) { len = u16(wd, p) + 1; }
            else if (op === 0xC615) { len = wd[p] === 255 ? 0 : wd[p] + 1; if (!len) break; }
            else len = wd[p] + 1;
            if (op === 0x2416) props.inTable = wd[p] !== 0;
            else if (op === 0x6649) props.inTable = props.inTable || u32(wd, p) > 0;
            else if (op === 0x2417 || op === 0x244C) props.ttp = props.ttp || wd[p] !== 0;
            else if (op === 0x2403 || op === 0x2461) props.jc = wd[p];
            p += len;
          }
        }
        runs.push([fa, fb, props]);
      }
    }
    runs.sort((a, b) => a[0] - b[0]);
    const none = { inTable: false, ttp: false, jc: 0, istd: 0 };
    return (fc) => {
      let lo = 0, hi = runs.length - 1;
      while (lo <= hi) {
        const m = (lo + hi) >> 1;
        if (fc < runs[m][0]) hi = m - 1;
        else if (fc >= runs[m][1]) lo = m + 1;
        else return runs[m][2];
      }
      return none;
    };
  }

  // Убирает коды полей (оставляя их результат) и объекты; свойства абзацев идут следом.
  function cleanFields(src) {
    const s = src.text, marks = src.marks;
    let r = '';
    const kept = [];
    let mi = 0;
    const stack = [];   // true — внутри кода поля, false — внутри результата
    for (const ch of s) {
      const c = ch.charCodeAt(0);
      const mark = c === 0x0D || c === 0x07 ? (marks ? marks[mi++] : null) : undefined;
      if (c === 0x13) { stack.push(true); continue; }
      if (c === 0x14) { if (stack.length) stack[stack.length - 1] = false; continue; }
      if (c === 0x15) { stack.pop(); continue; }
      if (stack.some(Boolean)) continue;
      if (c === 0x01 || c === 0x08 || c === 0x02 || c === 0x05 || c === 0x1F) continue;
      if (c === 0x1E) { r += '-'; continue; }
      if (c === 0x0B) { r += '\n'; continue; }
      if (mark !== undefined) kept.push(mark);
      r += ch;
    }
    return { text: r, marks: marks ? kept : null };
  }

  // Разбивает текст на блоки: { p, jc, istd } | { pb: true } | { table: [[ячейки]] }.
  function toBlocks(src) {
    const blocks = [];
    const para = (t, pr) => {
      t.split('\x0C').forEach((part, j) => {
        if (j > 0) blocks.push({ pb: true });
        if (j === 0 || part) blocks.push({ p: part, jc: pr ? pr.jc : 0, istd: pr ? pr.istd : 0 });
      });
    };
    return src.marks ? blocksByProps(src, blocks, para) : blocksByGuess(src.text, blocks, para);
  }

  // Точный путь: знаем из PAPX, какие абзацы в таблице и где конец строки.
  function blocksByProps(src, blocks, para) {
    let cur = '', mi = 0, table = null, row = null, cellParas = [];
    const closeTable = () => {
      if (row && row.length) table.push(row);
      if (table && table.length) blocks.push({ table });
      table = null; row = null; cellParas = [];
    };
    for (const ch of src.text) {
      if (ch !== '\r' && ch !== '\x07') { cur += ch; continue; }
      const p = src.marks[mi++] || {};
      if (p.ttp) {
        if (table && row) { table.push(row); row = null; }
      } else if (p.inTable) {
        if (!table) table = [];
        if (!row) row = [];
        if (ch === '\x07') { row.push(cellParas.concat(cur).join('\n')); cellParas = []; }
        else cellParas.push(cur);
      } else {
        if (table) closeTable();
        para(cur, p);
      }
      cur = '';
    }
    if (table) closeTable();
    if (cur) para(cur);
    return blocks;
  }

  // Запасной путь без PAPX: конец ячейки — \x07, конец строки — ещё один \x07 сразу после.
  function blocksByGuess(text, blocks, para) {
    let cur = '';
    let table = null, row = null, cell = null, cols = 0;
    let lastMark = '';
    const closeTable = () => {
      if (table && table.length) blocks.push({ table });
      table = null; row = null; cell = null; cols = 0;
    };
    for (const ch of text) {
      if (ch === '\x07') {
        if (!row) {
          if (!table) table = [];
          row = []; cell = null;
        }
        const isRowEnd = cur === '' && lastMark === '\x07' && cell === null &&
          (cols === 0 ? row.length > 0 : row.length >= cols);
        if (isRowEnd) {
          if (!cols) cols = row.length;
          table.push(row);
          row = null;
        } else {
          row.push(cell === null ? cur : cell + '\n' + cur);
          cell = null;
        }
        cur = ''; lastMark = '\x07';
      } else if (ch === '\r') {
        if (row) {
          cell = cell === null ? cur : cell + '\n' + cur;
        } else {
          if (table) closeTable();
          para(cur);
        }
        cur = ''; lastMark = '\r';
      } else {
        cur += ch;
      }
    }
    if (row && row.length) { if (!table) table = []; table.push(row); }
    closeTable();
    if (cur) para(cur);
    return blocks;
  }

  function esc(s) {
    return s.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  }

  // Раскладывает блоки по «страницам» A4 (по разрывам страниц из документа).
  function render(blocks, container) {
    container.textContent = '';
    let page = document.createElement('div');
    page.className = 'docpage';
    container.appendChild(page);
    let html = '';
    const flush = () => { page.insertAdjacentHTML('beforeend', html); html = ''; };
    for (const b of blocks) {
      if (b.pb) {
        flush();
        page = document.createElement('div');
        page.className = 'docpage';
        container.appendChild(page);
      } else if (b.table) {
        const width = Math.max(...b.table.map(r => r.length));
        html += '<table>' + b.table.map(r => {
          let cells = r.map(c => '<td>' + esc(c) + '</td>').join('');
          if (r.length < width) cells += '<td colspan="' + (width - r.length) + '"></td>';
          return '<tr>' + cells + '</tr>';
        }).join('') + '</table>';
      } else {
        const t = b.p.replace(/\t/g, '    ');
        const cls = (['', 'c', 'r', 'j'][b.jc] || '') + (b.istd >= 1 && b.istd <= 9 ? ' h' : '');
        html += t.trim() ? '<p' + (cls.trim() ? ' class="' + cls.trim() + '"' : '') + '>' + esc(t) + '</p>' : '<p class="empty"></p>';
      }
    }
    flush();
  }

  window.WordDoc = {
    render(bytes, container) {
      const cfb = XLSX.CFB.read(bytes, { type: 'array' });
      const blocks = toBlocks(cleanFields(extractText(cfb)));
      render(blocks, container);
      return blocks;
    },
    isWord(cfb) { return !!XLSX.CFB.find(cfb, 'WordDocument'); },
    _test: { extractText, cleanFields, toBlocks }
  };
})();
