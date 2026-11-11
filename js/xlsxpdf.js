/* Zeichnet den Druckbereich eines Excel-Blatts originalgetreu in ein jsPDF-Dokument.
 * Liest Spaltenbreiten, Zeilenhöhen, verbundene Zellen, Rahmen, Füllungen, Schriften,
 * Ausrichtung (inkl. gedrehter Texte), Zahlenformate und Dropdown-Steuerelemente direkt aus dem XLSX.
 * Läuft im Browser (window.XlsxPdf) und in Node (module.exports).
 */
(function (root) {
  'use strict';
  const PT2MM = 25.4 / 72;

  // ---------- XML-Helfer ----------
  const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');
  const attr = (tag, name) => { const m = new RegExp('\\b' + name + '="([^"]*)"').exec(tag); return m ? m[1] : null; };
  function colNum(col) { let n = 0; for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function parseRef(ref) { const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref); return { c: colNum(m[1]), r: +m[2] }; }
  function parseRange(rg) { const [a, b] = rg.replace(/^.*!/, '').split(':'); const A = parseRef(a), B = parseRef(b || a); return { c1: A.c, r1: A.r, c2: B.c, r2: B.r }; }

  // ---------- Farben ----------
  const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
  function hexToRgb(h) { return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
  function applyTint(rgb, tint) {
    if (!tint) return rgb;
    let [r, g, b] = rgb.map((v) => v / 255);
    const max = Math.max(r, g, b), min = Math.min(r, g, b); let h = 0, s = 0; let l = (max + min) / 2;
    if (max !== min) { const d = max - min; s = l > 0.5 ? d / (2 - max - min) : d / (max + min); h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    const hue = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
    if (s === 0) r = g = b = l; else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = hue(p, q, h + 1 / 3); g = hue(p, q, h); b = hue(p, q, h - 1 / 3); }
    return [r, g, b].map((v) => Math.round(v * 255));
  }
  function resolveColor(tag, theme, fallback) {
    if (!tag) return fallback;
    const rgb = attr(tag, 'rgb'), th = attr(tag, 'theme'), ix = attr(tag, 'indexed'), tint = parseFloat(attr(tag, 'tint') || '0');
    let c = null;
    if (rgb) c = hexToRgb(rgb.slice(-6));
    else if (th != null) c = hexToRgb(theme[+th] || '000000');
    else if (ix != null) { const i = +ix; if (i === 64) c = [0, 0, 0]; else if (i === 65) c = [255, 255, 255]; else c = hexToRgb(INDEXED[i] || '000000'); }
    else if (/auto="1"/.test(tag)) c = [0, 0, 0];
    return c ? applyTint(c, tint) : fallback;
  }

  // ---------- Styles ----------
  function parseTheme(xml) {
    const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    const out = [];
    for (const n of order) {
      const m = new RegExp('<a:' + n + '>([\\s\\S]*?)</a:' + n + '>').exec(xml || '');
      const v = m ? (/lastClr="([0-9A-F]{6})"/i.exec(m[1]) || /val="([0-9A-F]{6})"/i.exec(m[1])) : null;
      out.push(v ? v[1] : (n === 'lt1' ? 'FFFFFF' : '000000'));
    }
    return out;
  }
  function parseStyles(xml, theme) {
    const numFmts = {};
    for (const m of xml.matchAll(/<numFmt [^>]*\/>/g)) numFmts[attr(m[0], 'numFmtId')] = unesc(attr(m[0], 'formatCode'));
    const section = (name) => { const m = new RegExp('<' + name + '[^>]*>([\\s\\S]*?)</' + name + '>').exec(xml); return m ? m[1] : ''; };
    const fonts = [...section('fonts').matchAll(/<font>([\s\S]*?)<\/font>|<font\/>/g)].map((m) => {
      const f = m[1] || '';
      return { sz: parseFloat((/<sz val="([\d.]+)"/.exec(f) || [0, 10])[1]), b: /<b\/>|<b val="1"/.test(f), i: /<i\/>/.test(f), u: /<u[ \/>]/.test(f),
        color: resolveColor((/<color [^>]*\/>/.exec(f) || [null])[0], theme, [0, 0, 0]), sup: /superscript/.test(f) };
    });
    const fills = [...section('fills').matchAll(/<fill>([\s\S]*?)<\/fill>/g)].map((m) => {
      const pt = attr(m[1], 'patternType'); if (pt !== 'solid') return null;
      return resolveColor((/<fgColor [^>]*\/>/.exec(m[1]) || [null])[0], theme, null);
    });
    const borders = [...section('borders').matchAll(/<border(?: [^>]*)?>([\s\S]*?)<\/border>|<border\/>/g)].map((m) => {
      const b = m[1] || ''; const o = {};
      for (const side of ['left', 'right', 'top', 'bottom']) {
        const sm = new RegExp('<' + side + '(?: style="(\\w+)")?\\s*(?:/>|>([\\s\\S]*?)</' + side + '>)').exec(b);
        if (sm && sm[1]) o[side] = { style: sm[1], color: resolveColor((/<color [^>]*\/>/.exec(sm[2] || '') || [null])[0], theme, [0, 0, 0]) };
      }
      return o;
    });
    const xfs = [...section('cellXfs').matchAll(/<xf ([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g)].map((m) => {
      const a = m[1], inner = m[2] || ''; const al = /<alignment [^>]*\/>/.exec(inner);
      return { numFmtId: attr(a, 'numFmtId') || '0', fontId: +(attr(a, 'fontId') || 0), fillId: +(attr(a, 'fillId') || 0), borderId: +(attr(a, 'borderId') || 0),
        h: al ? attr(al[0], 'horizontal') : null, v: al ? attr(al[0], 'vertical') : null, wrap: al ? attr(al[0], 'wrapText') === '1' : false,
        rot: al ? +(attr(al[0], 'textRotation') || 0) : 0, indent: al ? +(attr(al[0], 'indent') || 0) : 0 };
    });
    return { numFmts, fonts, fills, borders, xfs };
  }

  // ---------- Zahlenformate ----------
  const BUILTIN = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 14: 'DD.MM.YYYY', 20: 'h:mm', 49: '@' };
  function splitSections(f) { const out = []; let cur = '', q = false; for (let i = 0; i < f.length; i++) { const ch = f[i]; if (ch === '"') q = !q; if (ch === '\\' && !q) { cur += ch + f[++i]; continue; } if (ch === ';' && !q) { out.push(cur); cur = ''; } else cur += ch; } out.push(cur); return out; }
  function fmtNumberSection(sec, n) {
    // Literale extrahieren, Platzhalter-Block finden
    let out = '', i = 0, numDone = false; const s = sec.replace(/\[[^\]]*\]/g, '');
    const isPh = (ch) => '0#?.,%'.includes(ch);
    while (i < s.length) {
      const ch = s[i];
      if (ch === '"') { const j = s.indexOf('"', i + 1); out += s.slice(i + 1, j < 0 ? s.length : j); i = j < 0 ? s.length : j + 1; continue; }
      if (ch === '\\') { out += s[i + 1] || ''; i += 2; continue; }
      if (ch === '_') { out += ' '; i += 2; continue; }
      if (ch === '*') { i += 2; continue; }
      if (isPh(ch) && !numDone) {
        let j = i; while (j < s.length && (isPh(s[j]) || (s[j] === ' ' && isPh(s[j + 1] || '')))) j++;
        const ph = s.slice(i, j).replace(/ /g, '');
        const pct = ph.includes('%'); let v = pct ? n * 100 : n;
        const [ip, dp = ''] = ph.replace(/%/g, '').split('.');
        const dec = (dp.match(/[0#?]/g) || []).length; const thou = ip.includes(',');
        const minInt = (ip.match(/0/g) || []).length;
        let str = Math.abs(v).toFixed(dec);
        let [a, b] = str.split('.');
        if (a.length < minInt) a = a.padStart(minInt, '0');
        if (/^0+$/.test(a) && minInt === 0) a = (ip.includes('?') ? ' ' : '');
        if (thou) a = a.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
        if (b != null && /^0*$/.test(b) && !/0/.test(dp) && /\?/.test(dp)) b = b.replace(/0/g, ' ');
        out += a + (dec ? ',' + b : '') + (pct ? '%' : '');
        numDone = true; i = j; continue;
      }
      if (ch === '@') { i++; continue; }
      out += ch; i++;
    }
    return out;
  }
  function formatValue(val, type, fmtCode) {
    if (val == null || val === '') return '';
    if (type === 'e') return String(val);
    const f = fmtCode || 'General';
    if (type === 's' || type === 'str' || type === 'inlineStr' || typeof val === 'string') {
      const secs = splitSections(f);
      if (secs.length >= 4) return secs[3].replace(/"/g, '').replace(/_./g, ' ').replace(/\*./g, '').replace(/\\(.)/g, '$1').replace('@', String(val)).trim();
      return String(val);
    }
    const n = Number(val);
    if (f === 'General') {
      if (Number.isInteger(n)) return String(n);
      return String(Math.round(n * 1e9) / 1e9).replace('.', ',');
    }
    const secs = splitSections(f);
    if (/h+:mm|\[h\]/i.test(secs[0]) && !/0\.0/.test(secs[0])) {
      const tot = Math.round((n % 1 + (n < 0 ? 1 : 0)) * 1440); const hh = Math.floor(tot / 60) % 24, mm = tot % 60;
      return (/hh/.test(secs[0]) ? String(hh).padStart(2, '0') : String(hh)) + ':' + String(mm).padStart(2, '0');
    }
    if (/[dmy]{2}/i.test(secs[0]) && !/[0#]/.test(secs[0])) {
      const d = new Date(Date.UTC(1899, 11, 30) + n * 864e5);
      return `${String(d.getUTCDate()).padStart(2, '0')}.${String(d.getUTCMonth() + 1).padStart(2, '0')}.${d.getUTCFullYear()}`;
    }
    let sec = secs[0], v = n;
    if (n < 0 && secs.length > 1) { sec = secs[1]; v = -n; } else if (n === 0 && secs.length > 2) { sec = secs[2]; v = 0; } else if (n < 0) { return '-' + fmtNumberSection(sec, -n); }
    return fmtNumberSection(sec, v).replace(/\s+$/, '').replace(/^\s+/, '');
  }

  // ---------- Blatt lesen ----------
  async function loadSheet(zip, sheetPath) {
    const read = async (p) => (zip.file(p) ? zip.file(p).async('string') : null);
    const theme = parseTheme(await read('xl/theme/theme1.xml'));
    const styles = parseStyles(await read('xl/styles.xml'), theme);
    const sstXml = await read('xl/sharedStrings.xml');
    const sst = [], sstRuns = [];
    if (sstXml) for (const m of sstXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      sst.push([...m[1].matchAll(/<t(?: [^>]*)?>([\s\S]*?)<\/t>/g)].map((t) => unesc(t[1])).join(''));
      if (/<r>/.test(m[1])) {
        sstRuns.push([...m[1].matchAll(/<r>([\s\S]*?)<\/r>/g)].map((r) => {
          const rp = (/<rPr>([\s\S]*?)<\/rPr>/.exec(r[1]) || [0, ''])[1];
          const t = [...r[1].matchAll(/<t(?: [^>]*)?>([\s\S]*?)<\/t>/g)].map((x) => unesc(x[1])).join('');
          return { t, rp: rp ? { sz: parseFloat((/<sz val="([\d.]+)"/.exec(rp) || [0, 0])[1]) || null, b: /<b\/>/.test(rp), i: /<i\/>/.test(rp), u: /<u[ \/>]/.test(rp), sup: /superscript/.test(rp), sub: /subscript/.test(rp), color: resolveColor((/<color [^>]*\/>/.exec(rp) || [null])[0], theme, null) } : null };
        }));
      } else sstRuns.push(null);
    }
    const xml = await read(sheetPath);
    const fp = /<sheetFormatPr [^>]*>/.exec(xml); const defRowH = fp ? parseFloat(attr(fp[0], 'defaultRowHeight') || '12.75') : 12.75;
    const defColW = fp ? parseFloat(attr(fp[0], 'defaultColWidth') || '9.140625') : 9.140625;
    const cols = {}; // index -> {w, hidden}
    for (const m of xml.matchAll(/<col [^>]*\/>/g)) {
      const a = +attr(m[0], 'min'), b = +attr(m[0], 'max');
      for (let c = a; c <= b; c++) cols[c] = { w: parseFloat(attr(m[0], 'width')), hidden: attr(m[0], 'hidden') === '1' };
    }
    const rows = {}; const cells = {};
    for (const rm of xml.matchAll(/<row ([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const r = +attr(rm[1], 'r');
      rows[r] = { ht: attr(rm[1], 'ht') ? parseFloat(attr(rm[1], 'ht')) : null, hidden: attr(rm[1], 'hidden') === '1', s: attr(rm[1], 's'), customFormat: attr(rm[1], 'customFormat') === '1' };
      for (const cm of (rm[2] || '').matchAll(/<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const t = attr(cm[3], 't'), s = +(attr(cm[3], 's') || 0), inner = cm[4] || '';
        let v = null; const vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
        let runs = null;
        if (t === 's') { v = vm ? sst[+vm[1]] : null; runs = vm ? sstRuns[+vm[1]] : null; }
        else if (t === 'inlineStr') v = [...inner.matchAll(/<t(?: [^>]*)?>([\s\S]*?)<\/t>/g)].map((x) => unesc(x[1])).join('');
        else if (t === 'str' || t === 'e') v = vm ? unesc(vm[1]) : null;
        else if (t === 'b') v = vm ? (vm[1] === '1' ? 'WAHR' : 'FALSCH') : null;
        else v = vm ? Number(vm[1]) : null;
        cells[colNum(cm[1]) + ',' + cm[2]] = { v, t: t || 'n', s, runs };
      }
    }
    const merges = [...xml.matchAll(/<mergeCell ref="([^"]+)"\/>/g)].map((m) => parseRange(m[1]));
    // Formular-Steuerelemente (Dropdowns)
    const controls = [];
    const relsXml = await read(sheetPath.replace('worksheets/', 'worksheets/_rels/') + '.rels');
    const rels = {}; if (relsXml) for (const m of relsXml.matchAll(/<Relationship [^>]*\/>/g)) rels[attr(m[0], 'Id')] = attr(m[0], 'Target');
    for (const m of xml.matchAll(/<control [^>]*r:id="([^"]+)"[^>]*>[\s\S]*?<from><xdr:col>(\d+)<\/xdr:col><xdr:colOff>(\d+)<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row><xdr:rowOff>(\d+)<\/xdr:rowOff><\/from><to><xdr:col>(\d+)<\/xdr:col><xdr:colOff>(\d+)<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row><xdr:rowOff>(\d+)<\/xdr:rowOff><\/to>/g)) {
      const target = rels[m[1]]; if (!target) continue;
      const cp = await read('xl/' + target.replace('../', ''));
      if (!cp || attr(cp, 'objectType') !== 'Drop') continue;
      const link = attr(cp, 'fmlaLink'), range = attr(cp, 'fmlaRange');
      let text = '';
      if (link && range && !/#REF/.test(link)) {
        const L = parseRef(link); const idx = cells[L.c + ',' + L.r] ? Number(cells[L.c + ',' + L.r].v) : Number(attr(cp, 'sel') || 0);
        const R = parseRange(range); const cell = cells[R.c1 + ',' + (R.r1 + idx - 1)];
        text = cell && cell.v != null ? String(cell.v).trim() : '';
      }
      controls.push({ from: { c: +m[2] + 1, cOff: +m[3] / 12700, r: +m[4] + 1, rOff: +m[5] / 12700 }, to: { c: +m[6] + 1, cOff: +m[7] / 12700, r: +m[8] + 1, rOff: +m[9] / 12700 }, text });
    }
    // Druckbereich
    const wb = await read('xl/workbook.xml');
    const sheets = [...wb.matchAll(/<sheet [^>]*\/>/g)].map((x) => ({ name: unesc(attr(x[0], 'name')), rid: attr(x[0], 'r:id') }));
    const wbRels = await read('xl/_rels/workbook.xml.rels'); const ridToPath = {};
    for (const x of wbRels.matchAll(/<Relationship [^>]*\/>/g)) ridToPath[attr(x[0], 'Id')] = 'xl/' + attr(x[0], 'Target').replace(/^\/?xl\//, '');
    const idx = sheets.findIndex((s) => ridToPath[s.rid] === sheetPath);
    let area = null;
    for (const x of wb.matchAll(/<definedName [^>]*name="_xlnm.Print_Area"[^>]*localSheetId="(\d+)"[^>]*>([^<]+)<\/definedName>/g)) if (+x[1] === idx) area = parseRange(unesc(x[2]).split(',')[0]);
    const ps = /<pageSetup [^>]*>/.exec(xml); const pm = /<pageMargins [^>]*>/.exec(xml);
    const page = { landscape: ps ? attr(ps[0], 'orientation') === 'landscape' : false, margins: pm ? { l: +attr(pm[0], 'left'), r: +attr(pm[0], 'right'), t: +attr(pm[0], 'top'), b: +attr(pm[0], 'bottom') } : { l: 0.7, r: 0.7, t: 0.75, b: 0.75 } };
    return { styles, cols, rows, cells, merges, controls, area, defRowH, defColW, page, name: idx >= 0 ? sheets[idx].name : '' };
  }

  // ---------- Zeichnen ----------
  const colWidthPt = (w) => Math.floor(w * 7 + 5) * 0.75;
  const BORDER_W = { hair: 0.25, thin: 0.5, dotted: 0.5, dashed: 0.5, mediumDashed: 1, dashDot: 0.5, medium: 1.1, thick: 1.6, double: 1.4 };

  function renderSheet(doc, sh, opts = {}) {
    const st = sh.styles; const A = sh.area;
    // Spalten-/Zeilenpositionen in Blatt-Punkten
    const cx = {}; let x = 0;
    for (let c = A.c1; c <= A.c2 + 1; c++) { cx[c] = x; if (c > A.c2) break; const col = sh.cols[c]; const w = col ? (col.hidden ? 0 : colWidthPt(col.w)) : colWidthPt(sh.defColW); x += w; }
    const ry = {}; let y = 0;
    for (let r = A.r1; r <= A.r2 + 1; r++) { ry[r] = y; if (r > A.r2) break; const row = sh.rows[r]; const h = row ? (row.hidden ? 0 : (row.ht != null ? row.ht : sh.defRowH)) : sh.defRowH; y += h; }
    const W = x, H = y;
    const pw = doc.internal.pageSize.getWidth(), ph = doc.internal.pageSize.getHeight(); // mm
    const m = sh.page.margins; const ml = m.l * 25.4, mr = m.r * 25.4, mt = m.t * 25.4, mb = m.b * 25.4;
    const availW = pw - ml - mr - 2, availH = ph - mt - mb - (opts.footer ? 6 : 2);
    const k = Math.min(availW / (W * PT2MM), availH / (H * PT2MM)); // Skalierung Blatt-pt -> mm
    const S = PT2MM * k; const ox = ml + (availW - W * S) / 2, oy = mt;
    const X = (c) => ox + cx[c] * S, Y = (r) => oy + ry[r] * S;
    const fontScale = k; // Schriftgröße in pt * k

    // verbundene Zellen
    const mergeAt = {}; const covered = {};
    for (const mg of sh.merges) {
      if (mg.r2 < A.r1 || mg.r1 > A.r2 || mg.c2 < A.c1 || mg.c1 > A.c2) continue;
      mergeAt[mg.c1 + ',' + mg.r1] = mg;
      for (let r = mg.r1; r <= mg.r2; r++) for (let c = mg.c1; c <= mg.c2; c++) if (r !== mg.r1 || c !== mg.c1) covered[c + ',' + r] = mg;
    }
    const xfOf = (c, r) => { const cell = sh.cells[c + ',' + r]; if (cell) return st.xfs[cell.s] || st.xfs[0]; const row = sh.rows[r]; if (row && row.customFormat && row.s != null) return st.xfs[+row.s] || st.xfs[0]; return st.xfs[0]; };
    const visibleCol = (c) => !(sh.cols[c] && sh.cols[c].hidden);
    const visibleRow = (r) => !(sh.rows[r] && sh.rows[r].hidden);

    // 1) Füllungen
    for (let r = A.r1; r <= A.r2; r++) {
      if (!visibleRow(r)) continue;
      for (let c = A.c1; c <= A.c2; c++) {
        if (!visibleCol(c) || covered[c + ',' + r]) continue;
        const xf = xfOf(c, r); const fill = st.fills[xf.fillId]; if (!fill) continue;
        const mg = mergeAt[c + ',' + r]; const c2 = mg ? Math.min(mg.c2, A.c2) : c, r2 = mg ? Math.min(mg.r2, A.r2) : r;
        doc.setFillColor(fill[0], fill[1], fill[2]); doc.rect(X(c), Y(r), X(c2 + 1) - X(c), Y(r2 + 1) - Y(r), 'F');
      }
    }
    // 2) Texte
    const textOf = (c, r) => { const cell = sh.cells[c + ',' + r]; if (!cell || cell.v == null || cell.v === '') return ''; const xf = st.xfs[cell.s] || st.xfs[0]; return formatValue(cell.v, cell.t, st.numFmts[xf.numFmtId] || BUILTIN[xf.numFmtId] || 'General'); };
    for (let r = A.r1; r <= A.r2; r++) {
      if (!visibleRow(r)) continue;
      for (let c = A.c1; c <= A.c2; c++) {
        if (!visibleCol(c) || covered[c + ',' + r]) continue;
        const cell = sh.cells[c + ',' + r]; if (!cell) continue;
        let txt = textOf(c, r); if (!txt || !txt.trim()) continue;
        const xf = st.xfs[cell.s] || st.xfs[0]; const font = st.fonts[xf.fontId] || st.fonts[0];
        const mg = mergeAt[c + ',' + r]; let c2 = mg ? Math.min(mg.c2, A.c2) : c; const r2 = mg ? Math.min(mg.r2, A.r2) : r;
        const isNum = cell.t === 'n' && typeof cell.v === 'number';
        let h = xf.h || (isNum ? 'right' : 'left'); if (h === 'centerContinuous' || h === 'general') h = isNum ? 'right' : 'left'; if (h === 'fill' || h === 'justify' || h === 'distributed') h = 'left';
        const v = xf.v || 'bottom';
        let x1 = X(c), x2 = X(c2 + 1); const y1 = Y(r), y2 = Y(r2 + 1);
        const fsPt = font.sz * fontScale * (font.sup ? 0.65 : 1);
        doc.setFont('helvetica', font.b && font.i ? 'bolditalic' : font.b ? 'bold' : font.i ? 'italic' : 'normal');
        doc.setFontSize(fsPt); doc.setTextColor(font.color[0], font.color[1], font.color[2]);
        const pad = 1.6 * S; const lh = fsPt * 1.18 * PT2MM;
        if (xf.rot === 90 || xf.rot === 180 || (xf.rot > 0 && xf.rot <= 180)) {
          const ang = xf.rot <= 90 ? xf.rot : -(xf.rot - 90);
          const lines = xf.wrap ? doc.splitTextToSize(txt, (y2 - y1) - 2 * pad) : [txt];
          const tw = Math.max(...lines.map((l) => doc.getTextWidth(l)));
          const blockW = lines.length * lh; let bx = (x1 + x2) / 2 - blockW / 2 + lh * 0.78;
          if (h === 'left') bx = x1 + pad + lh * 0.78; if (h === 'right') bx = x2 - pad - blockW + lh * 0.78;
          lines.forEach((l, i) => {
            const lw = doc.getTextWidth(l); let by;
            if (v === 'top') by = y1 + pad + lw; else if (v === 'center') by = (y1 + y2) / 2 + lw / 2; else by = y2 - pad;
            if (ang < 0) by = by - lw;
            doc.text(l, bx + i * lh, by, { angle: ang });
          });
          void tw; continue;
        }
        // Überlauf in leere Nachbarzellen (nur ohne Umbruch, nicht verbunden)
        if (!xf.wrap && !mg) {
          const need = doc.getTextWidth(txt) + 2 * pad;
          if (h === 'left') { let cc = c; while (x2 - x1 < need && cc < A.c2 && !textOf(cc + 1, r) && !covered[(cc + 1) + ',' + r] && !mergeAt[(cc + 1) + ',' + r]) { cc++; x2 = X(cc + 1); } }
          else if (h === 'right') { let cc = c; while (x2 - x1 < need && cc > A.c1 && !textOf(cc - 1, r) && !covered[(cc - 1) + ',' + r] && !mergeAt[(cc - 1) + ',' + r]) { cc--; x1 = X(cc); } }
          else if (h === 'center') {
            let a = c, b = c2;
            while (x2 - x1 < need) {
              let moved = false;
              if (a > A.c1 && !textOf(a - 1, r) && !covered[(a - 1) + ',' + r] && !mergeAt[(a - 1) + ',' + r]) { a--; x1 = X(a); moved = true; }
              if (b < A.c2 && !textOf(b + 1, r) && !covered[(b + 1) + ',' + r] && !mergeAt[(b + 1) + ',' + r]) { b++; x2 = X(b + 1); moved = true; }
              if (!moved) break;
            }
          }
        }
        if (cell.runs && !xf.wrap) {
          const runs = cell.runs.map((ru) => {
            const rp = ru.rp || {}; const sz = (rp.sz || font.sz) * fontScale * (rp.sup || rp.sub ? 0.62 : 1);
            const b = ru.rp ? rp.b : font.b; const it = ru.rp ? rp.i : font.i;
            doc.setFont('helvetica', b && it ? 'bolditalic' : b ? 'bold' : it ? 'italic' : 'normal'); doc.setFontSize(sz);
            return { t: ru.t, sz, b, it, u: ru.rp ? rp.u : font.u, sup: rp.sup, sub: rp.sub, color: (ru.rp && rp.color) || font.color, w: doc.getTextWidth(ru.t) };
          });
          const totalW = runs.reduce((a, ru) => a + ru.w, 0);
          let tx = h === 'right' ? x2 - pad - totalW : h === 'center' ? (x1 + x2) / 2 - totalW / 2 : x1 + pad + xf.indent * 3 * S;
          const base = fsPt * PT2MM;
          const ty0 = v === 'top' ? y1 + pad * 0.6 + base * 0.85 : v === 'center' ? (y1 + y2) / 2 + base * 0.35 : y2 - pad * 0.5 - base * 0.18;
          doc.saveGraphicsState(); doc.rect(x1, y1, x2 - x1, y2 - y1, null); doc.clip(); doc.discardPath();
          for (const ru of runs) {
            doc.setFont('helvetica', ru.b && ru.it ? 'bolditalic' : ru.b ? 'bold' : ru.it ? 'italic' : 'normal'); doc.setFontSize(ru.sz); doc.setTextColor(ru.color[0], ru.color[1], ru.color[2]);
            const yy = ty0 - (ru.sup ? base * 0.38 : 0) + (ru.sub ? base * 0.15 : 0);
            doc.text(ru.t, tx, yy);
            if (ru.u) { doc.setDrawColor(ru.color[0], ru.color[1], ru.color[2]); doc.setLineWidth(0.15); const uo = ru.sz * PT2MM * 0.1; doc.line(tx, yy + uo, tx + ru.w, yy + uo); }
            tx += ru.w;
          }
          doc.restoreGraphicsState();
          continue;
        }
        const innerW = x2 - x1 - 2 * pad - xf.indent * 3 * S;
        let lines = xf.wrap ? doc.splitTextToSize(txt.replace(/\r/g, ''), Math.max(innerW, 1)) : txt.split(/\r?\n/);
        const blockH = lines.length * lh;
        let ty;
        if (v === 'top') ty = y1 + pad * 0.6 + fsPt * PT2MM * 0.85;
        else if (v === 'center') ty = (y1 + y2) / 2 - blockH / 2 + fsPt * PT2MM * 0.85;
        else ty = y2 - pad * 0.5 - blockH + fsPt * PT2MM * 0.85;
        doc.saveGraphicsState();
        doc.rect(x1, y1, x2 - x1, y2 - y1, null); doc.clip(); doc.discardPath();
        lines.forEach((l, i) => {
          let tx = x1 + pad + xf.indent * 3 * S, align = 'left';
          if (h === 'right') { tx = x2 - pad; align = 'right'; } else if (h === 'center') { tx = (x1 + x2) / 2; align = 'center'; }
          doc.text(l, tx, ty + i * lh, { align, baseline: 'alphabetic' });
          if (font.u) { const lw = doc.getTextWidth(l); const ux = align === 'left' ? tx : align === 'right' ? tx - lw : tx - lw / 2; doc.setDrawColor(font.color[0], font.color[1], font.color[2]); doc.setLineWidth(0.15); const uo = fsPt * PT2MM * 0.1; doc.line(ux, ty + i * lh + uo, ux + lw, ty + i * lh + uo); }
        });
        doc.restoreGraphicsState();
      }
    }
    // 3) Rahmen (je Kante, verbundene Bereiche über ihre Randzellen)
    const edge = (x1, y1, x2, y2, b) => {
      if (!b) return;
      doc.setDrawColor(b.color[0], b.color[1], b.color[2]);
      doc.setLineWidth((BORDER_W[b.style] || 0.5) * PT2MM * Math.max(0.55, Math.min(1, k * 1.6)));
      if (b.style === 'dotted' || b.style === 'hair') doc.setLineDashPattern([0.3, 0.5], 0); else if (/dash/i.test(b.style)) doc.setLineDashPattern([1.2, 0.6], 0); else doc.setLineDashPattern([], 0);
      doc.line(x1, y1, x2, y2);
    };
    const bordOf = (c, r) => st.borders[xfOf(c, r).borderId] || {};
    for (let r = A.r1; r <= A.r2; r++) {
      if (!visibleRow(r)) continue;
      for (let c = A.c1; c <= A.c2; c++) {
        if (!visibleCol(c)) continue;
        const mg = mergeAt[c + ',' + r] || covered[c + ',' + r];
        const b = bordOf(c, r);
        // Bei verbundenen Zellen nur die Außenkanten zeichnen
        const isTop = !mg || r === mg.r1, isBottom = !mg || r === Math.min(mg.r2, A.r2), isLeft = !mg || c === mg.c1, isRight = !mg || c === Math.min(mg.c2, A.c2);
        let bt = b.top, bb = b.bottom, bl = b.left, br = b.right;
        if (mg) { // Rahmen der verbundenen Zelle stammen aus der linken oberen bzw. der jeweiligen Randzelle
          const tl = bordOf(mg.c1, mg.r1);
          bt = bt || tl.top; bl = bl || tl.left; bb = bb || bordOf(mg.c1, mg.r2).bottom || tl.bottom; br = br || bordOf(mg.c2, mg.r1).right || tl.right;
        }
        if (isTop) edge(X(c), Y(r), X(c + 1), Y(r), bt);
        if (isBottom) edge(X(c), Y(r + 1), X(c + 1), Y(r + 1), bb);
        if (isLeft) edge(X(c), Y(r), X(c), Y(r + 1), bl);
        if (isRight) edge(X(c + 1), Y(r), X(c + 1), Y(r + 1), br);
      }
    }
    doc.setLineDashPattern([], 0);
    // 4) Dropdown-Steuerelemente (Länderauswahl) wie in Excel
    for (const ct of sh.controls) {
      if (ct.from.r < A.r1 || ct.to.r > A.r2 + 1 || ct.from.c < A.c1 || ct.to.c > A.c2 + 1) continue;
      const x1 = X(ct.from.c) + ct.from.cOff * S, y1 = Y(ct.from.r) + ct.from.rOff * S, x2 = X(ct.to.c) + ct.to.cOff * S, y2 = Y(ct.to.r) + ct.to.rOff * S;
      if (x2 <= x1 || y2 <= y1) continue;
      doc.setFillColor(255, 255, 255); doc.setDrawColor(150, 150, 150); doc.setLineWidth(0.12); doc.rect(x1, y1, x2 - x1, y2 - y1, 'FD');
      const bw = Math.min(y2 - y1, (x2 - x1) * 0.25); doc.setFillColor(235, 235, 235); doc.rect(x2 - bw, y1, bw, y2 - y1, 'FD');
      doc.setFillColor(80, 80, 80); const ax = x2 - bw / 2, ay = (y1 + y2) / 2; doc.triangle(ax - bw * 0.22, ay - bw * 0.1, ax + bw * 0.22, ay - bw * 0.1, ax, ay + bw * 0.16, 'F');
      if (ct.text) {
        doc.setFont('helvetica', 'normal'); doc.setTextColor(0, 0, 0);
        let fs = 8 * fontScale; doc.setFontSize(fs);
        let t = ct.text; const maxW = (x2 - x1) - bw - 1.4 * S;
        while (t.length > 1 && doc.getTextWidth(t) > maxW) t = t.slice(0, -1);
        doc.text(t, x1 + 0.8 * S, (y1 + y2) / 2 + fs * PT2MM * 0.35);
      }
    }
    return { scale: k, width: W * S, height: H * S };
  }

  async function renderWorkbookSheet(doc, zip, sheetPath, opts) {
    const sh = await loadSheet(zip, sheetPath);
    if (!sh.area) { // ohne Druckbereich: benutzten Bereich nehmen
      let c2 = 1, r2 = 1; for (const k of Object.keys(sh.cells)) { const [c, r] = k.split(',').map(Number); if (c > c2) c2 = c; if (r > r2) r2 = r; }
      sh.area = { c1: 1, r1: 1, c2, r2 };
    }
    return renderSheet(doc, sh, opts);
  }

  // Alle Blätter in Reihenfolge der Arbeitsmappe (Pfad + Name)
  async function listSheets(zip) {
    const wb = await zip.file('xl/workbook.xml').async('string');
    const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string'); const map = {};
    for (const x of rels.matchAll(/<Relationship [^>]*\/>/g)) map[attr(x[0], 'Id')] = 'xl/' + attr(x[0], 'Target').replace(/^\/?xl\//, '');
    return [...wb.matchAll(/<sheet [^>]*\/>/g)].filter((x) => attr(x[0], 'state') !== 'hidden').map((x) => ({ name: unesc(attr(x[0], 'name')), path: map[attr(x[0], 'r:id')] }));
  }

  const api = { loadSheet, renderSheet, renderWorkbookSheet, formatValue, listSheets };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.XlsxPdf = api;
})(typeof window !== 'undefined' ? window : globalThis);
