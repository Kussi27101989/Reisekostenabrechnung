/* RB003 – Befüllen der Original-Excel-Vorlage "Reisekostenabrechnung" (VYGON, Version 21)
 * Arbeitet direkt auf dem XML der Vorlage: Layout, Formeln, Dropdowns und Schutz bleiben erhalten.
 * Läuft im Browser (window.RB003) und in Node (module.exports) – JSZip wird übergeben.
 */
(function (root) {
  'use strict';

  // ---------- Kategorien (Spalten der Vorlage) ----------
  // sheet: 1 = Reisekostenabrechnung, 2 = Seite2
  const CATEGORIES = {
    verpflegung:   { nr: 2,  label: 'Verpflegung (Einzelnachweis)',      short: 'Verpflegung',      sheet: 1, txt: 'AA', amt: 'AB', vat: '19/7' },
    uebern_pausch: { nr: 3,  label: 'Übernachtung Pauschale',            short: 'Übern.-Pauschale', sheet: 1, txt: 'AC', amt: 'AD', manualOnly: true },
    uebernachtung: { nr: 4,  label: 'Übernachtung (Einzelnachweis)',     short: 'Hotel',            sheet: 1, txt: 'AE', amt: 'AF', vat: '7' },
    bewirtung:     { nr: 5,  label: 'Bewirtungskosten',                  short: 'Bewirtung',        sheet: 1, txt: 'AG', amt: 'AH', vat: '19' },
    geschenk19:    { nr: 6,  label: 'Geschenke 19 % MwSt',               short: 'Geschenk 19 %',    sheet: 1, txt: 'AI', amt: 'AJ', vat: '19' },
    geschenk7:     { nr: 7,  label: 'Geschenke 7 % MwSt (Blumen, Bücher…)', short: 'Geschenk 7 %',  sheet: 1, txt: 'AK', amt: 'AL', vat: '7' },
    neben19:       { nr: 8,  label: 'Nebenkosten 19 % (Büromat., Parken…)', short: 'Nebenk. 19 %',  sheet: 1, txt: 'AM', amt: 'AN', vat: '19' },
    neben7:        { nr: 9,  label: 'Nebenkosten 7 % (Fachliteratur…)',  short: 'Nebenk. 7 %',      sheet: 1, txt: 'AO', amt: 'AP', vat: '7' },
    neben0:        { nr: 10, label: 'Nebenkosten ohne MwSt (Parken, Porto)', short: 'Nebenk. o. MwSt', sheet: 1, txt: 'AQ', amt: 'AR', vat: '0' },
    dw_fuel:       { nr: 11, label: 'Dienstwagen – Kraftstoff',          short: 'DW Tanken',        sheet: 2, amt: 'F', liter: 'E' },
    dw_oel:        { nr: 11, label: 'Dienstwagen – Öl',                  short: 'DW Öl',            sheet: 2, amt: 'H', liter: 'G' },
    dw_garage:     { nr: 11, label: 'Dienstwagen – Garage',              short: 'DW Garage',        sheet: 2, amt: 'I' },
    dw_pflege:     { nr: 11, label: 'Dienstwagen – Wagenpflege',         short: 'DW Pflege',        sheet: 2, amt: 'J' },
    dw_reparatur:  { nr: 11, label: 'Dienstwagen – Reparatur/Inspektion', short: 'DW Reparatur',    sheet: 2, amt: 'K' },
    mw_fuel:       { nr: 12, label: 'Mietwagen – Kraftstoff',            short: 'MW Tanken',        sheet: 2, amt: 'P', liter: 'O' },
    mw_miete:      { nr: 12, label: 'Mietwagen – Miete',                 short: 'MW Miete',         sheet: 2, amt: 'Q' },
    oepnv_bis50:   { nr: 14, label: 'Öffentl. Verkehrsmittel bis 50 km', short: 'ÖPNV ≤ 50 km',     sheet: 2, amt: 'U' },
    oepnv_ueber50: { nr: 15, label: 'Öffentl. Verkehrsmittel über 50 km (Bahn, Flug)', short: 'Bahn/Flug > 50 km', sheet: 2, amt: 'V' },
  };

  const S1 = 'xl/worksheets/sheet1.xml';
  const S2 = 'xl/worksheets/sheet2.xml';

  // ---------- kleine Helfer ----------
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const r2 = (n) => Math.round((Number(n) + Number.EPSILON) * 100) / 100;
  function colNum(col) { let n = 0; for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function splitRef(ref) { const m = /^([A-Z]+)(\d+)$/.exec(ref); return { col: m[1], row: +m[2] }; }

  function parseSharedStrings(xml) {
    const out = [];
    const re = /<si>([\s\S]*?)<\/si>/g; let m;
    while ((m = re.exec(xml))) {
      const parts = []; const tr = /<t(?: [^>]*)?>([\s\S]*?)<\/t>/g; let t;
      while ((t = tr.exec(m[1]))) parts.push(unesc(t[1]));
      out.push(parts.join(''));
    }
    return out;
  }

  // Liest einen Zellwert (Zahl oder Text) aus dem Sheet-XML
  function readCell(xml, ref, sst) {
    const m = new RegExp('<c r="' + ref + '"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)').exec(xml);
    if (!m || !m[2]) return null;
    const attrs = m[1]; const inner = m[2];
    const v = /<v>([\s\S]*?)<\/v>/.exec(inner);
    if (/t="s"/.test(attrs)) return v ? sst[+v[1]] : null;
    if (/t="inlineStr"/.test(attrs)) { const t = /<t[^>]*>([\s\S]*?)<\/t>/.exec(inner); return t ? unesc(t[1]) : null; }
    if (/t="str"/.test(attrs)) return v ? unesc(v[1]) : null;
    return v ? Number(v[1]) : null;
  }

  // ---------- Ländertabelle aus der Vorlage lesen (G349:R577) ----------
  async function readTemplateInfo(JSZip, buf) {
    const zip = await JSZip.loadAsync(buf);
    const sst = parseSharedStrings(await zip.file('xl/sharedStrings.xml').async('string'));
    const xml = await zip.file(S1).async('string');
    const countries = [];
    for (let r = 349; r <= 700; r++) {
      const idx = readCell(xml, 'G' + r, sst);
      const name = readCell(xml, 'J' + r, sst);
      if (idx == null || name == null) { if (r > 360) break; else continue; }
      countries.push({
        idx: Number(idx), name: String(name).trim(),
        full: Number(readCell(xml, 'K' + r, sst)) || 0,      // > 24 Std.
        partial: Number(readCell(xml, 'L' + r, sst)) || 0,   // > 8 < 24 Std. / An- und Abreisetag
        hotel: Number(readCell(xml, 'R' + r, sst)) || 0,     // Übernachtungspauschale (Ausland)
      });
    }
    const version = readCell(xml, 'A55', sst) || '';
    const company = (readCell(xml, 'A37', sst) || '').trim();
    return { countries, version: String(version), company };
  }

  // ---------- Verpflegungspauschale exakt wie die Excel-Formeln (U–Y) ----------
  function hoursOf(day) {
    if (!day || day.startH === '' || day.startH == null) return null;
    const s = Number(day.startH) + Number(day.startM || 0) / 60;
    const e = Number(day.endH || 0) + Number(day.endM || 0) / 60;
    return e - s;
  }
  function perDiem(day, countries) {
    const J = hoursOf(day);
    if (J == null) return { hours: null, gross: null, net: null, display: '-' };
    const c = countries.find((x) => x.idx === Number(day.country || 1)) || countries[0];
    let U;
    if (day.arrival) U = c.partial;
    else if (J > 8 && J < 24) U = c.partial;
    else if (J <= 8) U = 0;
    else U = c.full;
    let X = U;
    if (day.breakfast) X -= c.full * 0.2;
    if (day.lunch) X -= c.full * 0.4;
    if (day.dinner) X -= c.full * 0.4;
    X = r2(X);
    return { hours: J, gross: U, net: X < 0 ? 0 : X, display: X < 0 ? '-' : X, raw: X, country: c };
  }

  // ---------- Monat aggregieren ----------
  // data = { month:'01', year:'2026', name, costCenter, kmRate, days:{1:{...}}, vehicle:{1:{...}}, receipts:[...] }
  function aggregate(data, countries) {
    const days = {};
    for (let d = 1; d <= 31; d++) days[d] = { pd: perDiem(data.days && data.days[d], countries), cats: {} };
    const receipts = (data.receipts || []).filter((r) => r.day >= 1 && r.day <= 31 && CATEGORIES[r.category]);
    for (const r of receipts) {
      const slot = days[r.day].cats[r.category] || (days[r.day].cats[r.category] = { amount: 0, liters: 0, refs: [], texts: [] });
      slot.amount = r2(slot.amount + Number(r.amountEur || 0));
      slot.liters = r2(slot.liters + Number(r.liters || 0));
      slot.refs.push(r.belegNr);
      if (r.text) slot.texts.push(r.text);
    }
    // manuelle Übernachtungspauschalen aus den Tagesdaten
    for (let d = 1; d <= 31; d++) {
      const dd = data.days && data.days[d];
      if (dd && Number(dd.hotelFlat) > 0) days[d].cats.uebern_pausch = { amount: r2(dd.hotelFlat), liters: 0, refs: [], texts: ['Pauschale'] };
    }
    const kmRate = Number(data.kmRate || 0.30);
    const veh = {};
    for (let d = 1; d <= 31; d++) {
      const v = (data.vehicle && data.vehicle[d]) || {};
      const kmStart = v.kmStart === '' || v.kmStart == null ? null : Number(v.kmStart);
      const kmEnd = v.kmEnd === '' || v.kmEnd == null ? null : Number(v.kmEnd);
      const priv = Number(v.kmPrivate || 0);
      const total = kmStart != null && kmEnd != null ? kmEnd - kmStart : null;
      const privKm = Number(v.privCarKm || 0);
      veh[d] = {
        kmStart, kmEnd, dwPrivate: v.kmPrivate === '' || v.kmPrivate == null ? null : priv,
        dwBusiness: total != null ? total - priv : null,
        mwBusiness: v.mwKmBusiness === '' || v.mwKmBusiness == null ? null : Number(v.mwKmBusiness),
        mwPrivate: v.mwKmPrivate === '' || v.mwKmPrivate == null ? null : Number(v.mwKmPrivate),
        privKm: privKm || null, privAmount: privKm ? r2(privKm * kmRate) : null,
        note: v.note || '',
      };
    }
    // Summen wie in der Vorlage
    const col = {}; // nr -> Summe
    let pauschSum = 0;
    for (let d = 1; d <= 31; d++) {
      if (typeof days[d].pd.display === 'number') pauschSum = r2(pauschSum + days[d].pd.display);
      for (const [k, s] of Object.entries(days[d].cats)) col[k] = r2((col[k] || 0) + s.amount);
    }
    const sheet1Cols = ['verpflegung', 'uebern_pausch', 'uebernachtung', 'bewirtung', 'geschenk19', 'geschenk7', 'neben19', 'neben7', 'neben0'];
    const sum1 = r2(pauschSum + sheet1Cols.reduce((a, k) => a + (col[k] || 0), 0));
    let privSum = 0; for (let d = 1; d <= 31; d++) privSum = r2(privSum + (veh[d].privAmount || 0));
    const sheet2Cols = ['dw_fuel', 'dw_oel', 'dw_garage', 'dw_pflege', 'dw_reparatur', 'mw_fuel', 'mw_miete', 'oepnv_bis50', 'oepnv_ueber50'];
    const sum2 = r2(privSum + sheet2Cols.reduce((a, k) => a + (col[k] || 0), 0));
    return { days, veh, col, pauschSum, privSum, sum1, sum2, total: r2(sum1 + sum2), receipts };
  }

  // ---------- XML-Zellen setzen ----------
  function cellRegex(ref) { return new RegExp('<c r="' + ref + '"((?: [^>]*?)?)(/>|>[\\s\\S]*?</c>)'); }
  function styleOf(attrs) { const m = / s="(\d+)"/.exec(attrs || ''); return m ? ' s="' + m[1] + '"' : ''; }

  function insertCell(xml, ref, cellXml) {
    const { col, row } = splitRef(ref);
    const rowRe = new RegExp('<row r="' + row + '"[^>]*?(?:/>|>([\\s\\S]*?)</row>)');
    const m = rowRe.exec(xml);
    if (!m) throw new Error('Zeile ' + row + ' fehlt in der Vorlage');
    if (m[0].endsWith('/>')) {
      const open = m[0].slice(0, -2) + '>';
      return xml.replace(m[0], open + cellXml + '</row>');
    }
    // vor der ersten Zelle mit größerer Spalte einfügen
    const inner = m[1]; const cn = colNum(col);
    const cre = /<c r="([A-Z]+)\d+"/g; let c; let pos = inner.length;
    while ((c = cre.exec(inner))) { if (colNum(c[1]) > cn) { pos = c.index; break; } }
    const newInner = inner.slice(0, pos) + cellXml + inner.slice(pos);
    return xml.replace(m[0], m[0].replace(inner, newInner));
  }

  // value: number | string | null(=leeren)
  function setValue(xml, ref, value) {
    const re = cellRegex(ref); const m = re.exec(xml);
    const s = m ? styleOf(m[1]) : '';
    let cell;
    if (value === null || value === undefined || value === '') cell = '<c r="' + ref + '"' + s + '/>';
    else if (typeof value === 'number' && isFinite(value)) cell = '<c r="' + ref + '"' + s + '><v>' + value + '</v></c>';
    else cell = '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(value) + '</t></is></c>';
    if (m) return xml.replace(m[0], cell);
    return insertCell(xml, ref, cell);
  }

  // Ergebnis einer Formelzelle vorbelegen (für Vorschau-Apps ohne Neuberechnung, z. B. iOS Quick Look)
  function setCached(xml, ref, value) {
    const re = cellRegex(ref); const m = re.exec(xml);
    if (!m || !/<f[ >]/.test(m[2])) return xml;
    let attrs = m[1].replace(/ t="[^"]*"/, '');
    let inner = m[2].slice(1, -4); // ohne '>' und '</c>'
    inner = inner.replace(/<v>[\s\S]*?<\/v>|<v\/>/, '');
    let v;
    if (typeof value === 'number' && isFinite(value)) v = '<v>' + value + '</v>';
    else { attrs += ' t="str"'; v = '<v>' + esc(value == null ? '' : value) + '</v>'; }
    return xml.replace(m[0], '<c r="' + ref + '"' + attrs + '>' + inner + v + '</c>');
  }

  function pad2(n) { return String(n).padStart(2, '0'); }
  function shortText(slot, extra) {
    const refs = slot.refs.filter(Boolean);
    let t = refs.join('+');
    if (refs.length === 1 && slot.texts[0]) t += ' ' + slot.texts[0];
    if (!refs.length && slot.texts.length) t = slot.texts.join(', ');
    if (extra) t = (t ? t + ' ' : '') + extra;
    return t.length > 40 ? t.slice(0, 39) + '…' : t;
  }

  // ---------- Hauptfunktion: befüllte Excel erzeugen ----------
  async function buildWorkbook(JSZip, templateBuf, data, countries) {
    const zip = await JSZip.loadAsync(templateBuf);
    let x1 = await zip.file(S1).async('string');
    let x2 = await zip.file(S2).async('string');
    const agg = aggregate(data, countries);

    // Kopf
    x1 = setValue(x1, 'E43', pad2(data.month));
    x1 = setValue(x1, 'G43', String(data.year));
    x1 = setValue(x1, 'E44', data.name || '');
    x1 = setValue(x1, 'E45', /^\d+$/.test(String(data.costCenter || '')) ? Number(data.costCenter) : (data.costCenter || ''));

    const ctrlSel = {}; // Zeile -> Länderindex für Dropdown

    for (let d = 1; d <= 31; d++) {
      const row = 4 + d;
      const dd = (data.days && data.days[d]) || {};
      const has = dd.startH !== '' && dd.startH != null;
      const country = Number(dd.country || 1);
      x1 = setValue(x1, 'B' + row, has ? Number(dd.startH) : null);
      x1 = setValue(x1, 'D' + row, has ? pad2(dd.startM || 0) : null);
      x1 = setValue(x1, 'E' + row, has ? Number(dd.endH) : null);
      x1 = setValue(x1, 'G' + row, has ? pad2(dd.endM || 0) : null);
      x1 = setValue(x1, 'K' + row, has && dd.arrival ? 'X' : null);
      x1 = setValue(x1, 'L' + row, dd.destination || null);
      x1 = setValue(x1, 'O' + row, has && dd.breakfast ? 'x' : null);
      x1 = setValue(x1, 'P' + row, has && dd.lunch ? 'x' : null);
      x1 = setValue(x1, 'Q' + row, has && dd.dinner ? 'x' : null);
      x1 = setValue(x1, 'R' + row, country);
      ctrlSel[row] = country;

      // Vorberechnete Formel-Ergebnisse
      const pd = agg.days[d].pd; const c = countries.find((k) => k.idx === country) || countries[0];
      const hs = has ? Number(dd.startH) + ':' + pad2(dd.startM || 0) : ':';
      const he = has ? Number(dd.endH) + ':' + pad2(dd.endM || 0) : ':';
      x1 = setCached(x1, 'H' + row, hs);
      x1 = setCached(x1, 'I' + row, he);
      x1 = setCached(x1, 'J' + row, has ? pd.hours : '-');
      x1 = setCached(x1, 'S' + row, c.full);
      x1 = setCached(x1, 'T' + row, c.partial);
      if (has) {
        let U = pd.gross; x1 = setCached(x1, 'U' + row, U);
        let V = dd.breakfast ? U - c.full * 0.2 : U; x1 = setCached(x1, 'V' + row, r2(V));
        let W = dd.lunch ? V - c.full * 0.4 : V; x1 = setCached(x1, 'W' + row, r2(W));
        let X = dd.dinner ? W - c.full * 0.4 : W; x1 = setCached(x1, 'X' + row, r2(X));
        x1 = setCached(x1, 'Y' + row, X < 0 ? '-' : r2(X));
      } else {
        for (const col of ['U', 'V', 'W', 'X', 'Y']) x1 = setCached(x1, col + row, '-');
      }

      // Belegspalten Blatt 1
      for (const [key, cat] of Object.entries(CATEGORIES)) {
        if (cat.sheet !== 1) continue;
        const slot = agg.days[d].cats[key];
        x1 = setValue(x1, cat.txt + row, slot ? shortText(slot) : null);
        x1 = setValue(x1, cat.amt + row, slot ? slot.amount : null);
      }

      // Seite 2 (Doppelzeilen: ungerade = Werte / km-Anfang, gerade = km-Ende)
      const r1 = 5 + (d - 1) * 2, rEnd = r1 + 1;
      const v = agg.veh[d];
      x2 = setValue(x2, 'B' + r1, v.kmStart);
      x2 = setValue(x2, 'B' + rEnd, v.kmEnd);
      x2 = setValue(x2, 'C' + r1, v.dwBusiness);
      x2 = setValue(x2, 'D' + r1, v.dwPrivate);
      x2 = setValue(x2, 'M' + r1, v.mwBusiness);
      x2 = setValue(x2, 'N' + r1, v.mwPrivate);
      x2 = setValue(x2, 'S' + r1, v.privKm);
      x2 = setValue(x2, 'T' + r1, v.privAmount);
      const refs2 = [];
      for (const [key, cat] of Object.entries(CATEGORIES)) {
        if (cat.sheet !== 2) continue;
        const slot = agg.days[d].cats[key];
        x2 = setValue(x2, cat.amt + r1, slot ? slot.amount : null);
        if (cat.liter) x2 = setValue(x2, cat.liter + r1, slot && slot.liters ? slot.liters : null);
        if (slot) refs2.push(...slot.refs.filter(Boolean));
      }
      const note = [refs2.join('+'), v.note].filter(Boolean).join(' ');
      x2 = setValue(x2, 'W' + r1, note || null);
      const g = (k) => (agg.days[d].cats[k] ? agg.days[d].cats[k].amount : 0);
      x2 = setCached(x2, 'L' + r1, r2(g('dw_fuel') + g('dw_oel') + g('dw_garage') + g('dw_pflege') + g('dw_reparatur')));
      x2 = setCached(x2, 'R' + r1, r2(g('mw_fuel') + g('mw_miete')));
    }

    // Summen vorbelegen
    const colSum = (k) => agg.col[k] || 0;
    x1 = setCached(x1, 'Y36', agg.pauschSum);
    const s1map = { AA36: 'verpflegung', AC36: 'uebern_pausch', AE36: 'uebernachtung', AG36: 'bewirtung', AI36: 'geschenk19', AK36: 'geschenk7', AM36: 'neben19', AO36: 'neben7', AQ36: 'neben0' };
    for (const [ref, k] of Object.entries(s1map)) x1 = setCached(x1, ref, colSum(k));
    x1 = setCached(x1, 'AG52', agg.sum1);
    x1 = setCached(x1, 'AG54', agg.sum2);
    x1 = setCached(x1, 'AG56', agg.total);
    const s2map = { F67: 'dw_fuel', H67: 'dw_oel', I67: 'dw_garage', J67: 'dw_pflege', K67: 'dw_reparatur', P67: 'mw_fuel', Q67: 'mw_miete', U67: 'oepnv_bis50', V67: 'oepnv_ueber50' };
    for (const [ref, k] of Object.entries(s2map)) x2 = setCached(x2, ref, colSum(k));
    x2 = setCached(x2, 'L67', r2(['dw_fuel', 'dw_oel', 'dw_garage', 'dw_pflege', 'dw_reparatur'].reduce((a, k) => a + colSum(k), 0)));
    x2 = setCached(x2, 'R67', r2(colSum('mw_fuel') + colSum('mw_miete')));
    x2 = setCached(x2, 'T67', agg.privSum);
    // Letzter km-Stand (B66) ist kein Formelfeld – bereits über Tag 31 gesetzt.

    zip.file(S1, x1);
    zip.file(S2, x2);

    // Länder-Dropdowns (Formularsteuerelemente) auf das gewählte Land stellen
    const ctrlFiles = Object.keys(zip.files).filter((f) => /^xl\/ctrlProps\/ctrlProp\d+\.xml$/.test(f));
    for (const f of ctrlFiles) {
      let cx = await zip.file(f).async('string');
      const m = /fmlaLink="\$R\$(\d+)"/.exec(cx);
      if (m && ctrlSel[+m[1]]) {
        cx = /sel="\d+"/.test(cx) ? cx.replace(/sel="\d+"/, 'sel="' + ctrlSel[+m[1]] + '"') : cx.replace('<formControlPr ', '<formControlPr sel="' + ctrlSel[+m[1]] + '" ');
        zip.file(f, cx);
      }
    }

    // Excel beim Öffnen komplett neu rechnen lassen
    let wb = await zip.file('xl/workbook.xml').async('string');
    wb = wb.replace(/<calcPr([^>]*?)\/>/, (all, a) => '<calcPr' + a.replace(/ fullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"/>');
    zip.file('xl/workbook.xml', wb);

    return { zip, agg };
  }

  const api = { CATEGORIES, readTemplateInfo, perDiem, aggregate, buildWorkbook, hoursOf, r2 };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.RB003 = api;
})(typeof window !== 'undefined' ? window : globalThis);
