/* Monatsexport: befüllte RB003-Excel + PDF (Abrechnung + Belegliste + Einzelnachweise) */
(function () {
  'use strict';
  const C = RB003.CATEGORIES;
  const MONTHS = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
  const eur = (n) => (n == null || n === '' ? '' : Number(n).toLocaleString('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €');
  const num = (n, d = 2) => Number(n).toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });
  const pad = (n) => String(n).padStart(2, '0');
  const fmtDate = (iso) => (iso ? iso.split('-').reverse().join('.') : '');
  const fileSafe = (s) => String(s || '').replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/Ä/g, 'Ae').replace(/Ö/g, 'Oe').replace(/Ü/g, 'Ue').replace(/ß/g, 'ss').normalize('NFKD').replace(/[^\w\-]+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');

  function base64ToBuf(b64) { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u.buffer; }

  function blobToDataURL(blob) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(blob); }); }
  function imgSize(dataUrl) { return new Promise((res) => { const i = new Image(); i.onload = () => res({ w: i.naturalWidth, h: i.naturalHeight }); i.onerror = () => res({ w: 1000, h: 1400 }); i.src = dataUrl; }); }

  function baseName(ctx) { return `RB003_Reisekostenabrechnung_${ctx.year}-${pad(ctx.month)}_${fileSafe(ctx.name) || 'Mitarbeiter'}`; }

  async function buildExcel(ctx) {
    const { zip, agg } = await RB003.buildWorkbook(window.JSZip, ctx.templateBuf, ctx.data, ctx.countries);
    const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, agg, filename: baseName(ctx) + '.xlsx' };
  }

  // ---------- PDF ----------
  function table(doc, x, y, cols, rows, opts = {}) {
    const fs = opts.fontSize || 7.5, rh = opts.rowH || 4.6, pageH = doc.internal.pageSize.getHeight();
    const head = () => {
      doc.setFillColor(0, 68, 49); doc.setTextColor(255); doc.setFont('helvetica', 'bold'); doc.setFontSize(fs);
      let cx = x; doc.rect(x, y, cols.reduce((a, c) => a + c.w, 0), rh + 0.6, 'F');
      for (const c of cols) { doc.text(c.t, c.align === 'right' ? cx + c.w - 1.2 : cx + 1.2, y + rh - 1.2, { align: c.align === 'right' ? 'right' : 'left' }); cx += c.w; }
      y += rh + 0.6; doc.setTextColor(20); doc.setFont('helvetica', 'normal');
    };
    head();
    rows.forEach((r, i) => {
      if (y + rh > pageH - 12) { doc.addPage(); y = 14; head(); }
      if (r._bold) doc.setFont('helvetica', 'bold');
      if (i % 2 === 1 && !r._bold) { doc.setFillColor(236, 244, 240); doc.rect(x, y, cols.reduce((a, c) => a + c.w, 0), rh, 'F'); }
      if (r._bold) { doc.setDrawColor(0, 68, 49); doc.line(x, y, x + cols.reduce((a, c) => a + c.w, 0), y); }
      let cx = x;
      cols.forEach((c, j) => {
        let v = r[j] == null ? '' : String(r[j]);
        const maxW = c.w - 2.4;
        while (v && doc.getTextWidth(v) > maxW) v = v.slice(0, -2) + '…';
        doc.text(v, c.align === 'right' ? cx + c.w - 1.2 : cx + 1.2, y + rh - 1.3, { align: c.align === 'right' ? 'right' : 'left' });
        cx += c.w;
      });
      doc.setFont('helvetica', 'normal');
      y += rh;
    });
    return y;
  }

  function header(doc, ctx, title, landscape) {
    const W = doc.internal.pageSize.getWidth();
    doc.setFillColor(0, 68, 49); doc.rect(0, 0, W, 9, 'F');
    doc.setFillColor(0, 174, 66); doc.rect(0, 9, W, 1.2, 'F');
    doc.setTextColor(255); doc.setFont('helvetica', 'bold'); doc.setFontSize(9);
    doc.text(title, 10, 6);
    doc.setFont('helvetica', 'normal');
    doc.text(`${ctx.name || ''} · Kst. ${ctx.costCenter || '–'} · ${MONTHS[ctx.month - 1]} ${ctx.year}`, W - 10, 6, { align: 'right' });
    doc.setTextColor(20);
    return landscape ? 15 : 15;
  }

  function footer(doc, ctx) {
    const n = doc.getNumberOfPages();
    for (let i = 1; i <= n; i++) {
      doc.setPage(i);
      const W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight();
      doc.setFontSize(6.5); doc.setTextColor(120);
      doc.text(`${ctx.company || ''} · ${ctx.version || 'RB003'} · erstellt ${new Date().toLocaleString('de-DE')}`, 10, H - 5);
      doc.text(`Seite ${i} / ${n}`, W - 10, H - 5, { align: 'right' });
    }
  }

  async function buildPdf(ctx, agg, receipts, images) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
    const W = doc.internal.pageSize.getWidth();
    let y = header(doc, ctx, `Reisekostenabrechnung ${MONTHS[ctx.month - 1]} ${ctx.year}`, true);

    // Kopfdaten + Summen
    doc.setFontSize(8.5);
    const kv = [['Name', ctx.name || '–'], ['Kostenstelle', ctx.costCenter || '–'], ['Monat/Jahr', `${pad(ctx.month)}/${ctx.year}`], ['Firma', ctx.company || '–']];
    kv.forEach(([k, v], i) => { doc.setFont('helvetica', 'bold'); doc.text(k + ':', 10, y + i * 4.6); doc.setFont('helvetica', 'normal'); doc.text(String(v), 36, y + i * 4.6); });
    const sums = [['Summe Brutto-Spalten 1–10', agg.sum1], ['Summe Brutto-Spalten 11–15', agg.sum2], ['Reisekosten insgesamt', agg.total]];
    sums.forEach(([k, v], i) => {
      doc.setFont('helvetica', i === 2 ? 'bold' : 'normal'); doc.setFontSize(i === 2 ? 10 : 8.5);
      doc.text(k, W - 70, y + i * 5); doc.text(eur(v), W - 10, y + i * 5, { align: 'right' });
    });
    doc.setFontSize(8.5);
    y += 22;

    // Reisetage
    const dayRows = [];
    for (let d = 1; d <= 31; d++) {
      const dd = (ctx.data.days || {})[d]; const pd = agg.days[d].pd;
      if (!dd || pd.hours == null) continue;
      const meals = [dd.breakfast ? 'F' : '', dd.lunch ? 'M' : '', dd.dinner ? 'A' : ''].filter(Boolean).join('/');
      dayRows.push([`${pad(d)}.${pad(ctx.month)}.`, `${dd.startH}:${pad(dd.startM || 0)}`, `${dd.endH}:${pad(dd.endM || 0)}`, num(pd.hours, 2), dd.arrival ? 'X' : '', dd.destination || '', pd.country ? pd.country.name : '', meals, typeof pd.display === 'number' ? eur(pd.display) : '–']);
    }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.text('Reisetage und Verpflegungspauschalen (Spalte 1)', 10, y); y += 2;
    if (dayRows.length) {
      dayRows.push({ 0: 'Summe', 8: eur(agg.pauschSum), _bold: true });
      y = table(doc, 10, y, [
        { t: 'Tag', w: 15 }, { t: 'Beginn', w: 14 }, { t: 'Ende', w: 14 }, { t: 'Std.', w: 13, align: 'right' }, { t: 'An/Ab', w: 12 },
        { t: 'Reiseziel (Ort, Firma, Zweck)', w: 110 }, { t: 'Land', w: 48 }, { t: 'Kürz.', w: 14 }, { t: 'Pauschale', w: 22, align: 'right' }], dayRows);
    } else { doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.text('Keine Reisetage erfasst.', 10, y + 4); y += 6; }
    y += 6;

    // Spaltensummen
    const colRows = [];
    colRows.push(['1', 'Verpflegung Pauschale', eur(agg.pauschSum)]);
    const order = ['verpflegung', 'uebern_pausch', 'uebernachtung', 'bewirtung', 'geschenk19', 'geschenk7', 'neben19', 'neben7', 'neben0'];
    for (const k of order) if (agg.col[k]) colRows.push([String(C[k].nr), C[k].label, eur(agg.col[k])]);
    const dw = ['dw_fuel', 'dw_oel', 'dw_garage', 'dw_pflege', 'dw_reparatur'].reduce((a, k) => a + (agg.col[k] || 0), 0);
    const mw = (agg.col.mw_fuel || 0) + (agg.col.mw_miete || 0);
    if (dw) colRows.push(['11', 'Dienstwagen (Kraftstoff, Öl, Garage, Pflege, Reparatur)', eur(dw)]);
    if (mw) colRows.push(['12', 'Mietwagen (Kraftstoff, Miete)', eur(mw)]);
    if (agg.privSum) colRows.push(['13', `Privatwagen km-Pauschale (${num(ctx.data.kmRate || 0.3)} €/km)`, eur(agg.privSum)]);
    if (agg.col.oepnv_bis50) colRows.push(['14', C.oepnv_bis50.label, eur(agg.col.oepnv_bis50)]);
    if (agg.col.oepnv_ueber50) colRows.push(['15', C.oepnv_ueber50.label, eur(agg.col.oepnv_ueber50)]);
    colRows.push({ 1: 'Reisekosten insgesamt', 2: eur(agg.total), _bold: true });
    if (y > 150) { doc.addPage(); y = header(doc, ctx, 'Reisekostenabrechnung – Fortsetzung', true); }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.text('Summen je Spalte der RB003', 10, y); y += 2;
    const yAfterCols = table(doc, 10, y, [{ t: 'Sp.', w: 12 }, { t: 'Kostenart', w: 100 }, { t: 'Betrag', w: 28, align: 'right' }], colRows);

    // Unterschrift
    let sy = Math.max(yAfterCols + 14, y + 20);
    if (sy > 190) { doc.addPage(); sy = 40; }
    doc.setDrawColor(120); doc.line(165, sy, 225, sy); doc.line(232, sy, 287, sy);
    doc.setFontSize(7.5); doc.setFont('helvetica', 'normal'); doc.text('Datum, Unterschrift Mitarbeiter/in', 165, sy + 4); doc.text('Genehmigt (Vorgesetzte/r)', 232, sy + 4);

    // Belegliste
    doc.addPage();
    y = header(doc, ctx, 'Belegliste (Einzelnachweise)', true);
    const lr = receipts.map((r) => [r.belegNr, fmtDate(r.date), pad(r.day) + '.', `${C[r.category].nr} ${C[r.category].short}`, [r.merchant, r.text].filter(Boolean).join(' – '),
      r.currency && r.currency !== 'EUR' ? `${num(r.amount)} ${r.currency}` : '', eur(r.amountEur), r.images.length + ' S.']);
    lr.push({ 4: 'Summe Belege', 6: eur(receipts.reduce((a, r) => a + Number(r.amountEur || 0), 0)), _bold: true });
    y = table(doc, 10, y + 1, [{ t: 'Beleg', w: 14 }, { t: 'Datum', w: 20 }, { t: 'Tag', w: 10 }, { t: 'Spalte', w: 38 }, { t: 'Händler / Beschreibung', w: 120 }, { t: 'Fremdw.', w: 24, align: 'right' }, { t: 'Betrag EUR', w: 26, align: 'right' }, { t: 'Seiten', w: 15, align: 'right' }], lr);

    // Fahrzeug (Seite 2)
    const vr = [];
    for (let d = 1; d <= 31; d++) {
      const v = agg.veh[d];
      if (v.kmStart == null && v.kmEnd == null && v.privKm == null && v.mwBusiness == null && !v.note) continue;
      vr.push([pad(d) + '.', v.kmStart ?? '', v.kmEnd ?? '', v.dwBusiness ?? '', v.dwPrivate ?? '', v.mwBusiness ?? '', v.mwPrivate ?? '', v.privKm ?? '', v.privAmount != null ? eur(v.privAmount) : '', v.note || '']);
    }
    if (vr.length) {
      y += 8; if (y > 170) { doc.addPage(); y = header(doc, ctx, 'Fahrzeug (Seite 2)', true); }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.text('Fahrzeug / Kilometer (Seite 2)', 10, y); y += 2;
      table(doc, 10, y, [{ t: 'Tag', w: 12 }, { t: 'DW km Anf.', w: 22, align: 'right' }, { t: 'DW km Ende', w: 22, align: 'right' }, { t: 'DW gesch.', w: 20, align: 'right' }, { t: 'DW privat', w: 20, align: 'right' },
        { t: 'MW gesch.', w: 20, align: 'right' }, { t: 'MW privat', w: 20, align: 'right' }, { t: 'Privat-PKW km', w: 26, align: 'right' }, { t: 'km-Pauschale', w: 25, align: 'right' }, { t: 'Notiz', w: 80 }], vr);
    }

    // Einzelnachweise: je Belegbild eine Seite (Hochformat)
    for (const r of receipts) {
      const imgs = r.images.map((id) => images[id]).filter(Boolean);
      for (let p = 0; p < imgs.length; p++) {
        doc.addPage('a4', 'portrait');
        const PW = doc.internal.pageSize.getWidth(), PH = doc.internal.pageSize.getHeight();
        header(doc, ctx, `Einzelnachweis ${r.belegNr}${imgs.length > 1 ? ` (Seite ${p + 1}/${imgs.length})` : ''}`, false);
        doc.setFontSize(8.5); doc.setFont('helvetica', 'bold');
        doc.text(`${r.belegNr} · ${fmtDate(r.date) || pad(r.day) + '.' + pad(ctx.month) + '.' + ctx.year} · Spalte ${C[r.category].nr}: ${C[r.category].label}`, 10, 15);
        doc.setFont('helvetica', 'normal');
        const line2 = [r.merchant, r.text, r.currency && r.currency !== 'EUR' ? `${num(r.amount)} ${r.currency} = ${eur(r.amountEur)}` : eur(r.amountEur)].filter(Boolean).join(' · ');
        doc.text(line2, 10, 19.5);
        let top = 23;
        if (r.category === 'bewirtung') {
          const t1 = doc.splitTextToSize(`Anlass: ${r.occasion || '–'}   Teilnehmer: ${r.participants || '–'}`, PW - 20);
          doc.text(t1, 10, 24); top = 24 + t1.length * 4;
        }
        const { dataUrl, w, h } = imgs[p];
        const maxW = PW - 20, maxH = PH - top - 12;
        const s = Math.min(maxW / w, maxH / h);
        const iw = w * s, ih = h * s;
        doc.addImage(dataUrl, 'JPEG', (PW - iw) / 2, top, iw, ih, undefined, 'FAST');
        doc.setDrawColor(200); doc.rect((PW - iw) / 2, top, iw, ih);
      }
    }
    footer(doc, ctx);
    return { blob: doc.output('blob'), filename: baseName(ctx) + '_Belege.pdf' };
  }

  async function exportMonth(ctx, receiptsSorted, loadImageBlob, onStep) {
    onStep && onStep('Excel wird befüllt …');
    const xl = await buildExcel(ctx);
    onStep && onStep('Belege werden ins PDF übernommen …');
    const images = {};
    for (const r of receiptsSorted) for (const id of r.images) {
      const b = await loadImageBlob(id); if (!b) continue;
      const dataUrl = await blobToDataURL(b); const sz = await imgSize(dataUrl);
      images[id] = { dataUrl, w: sz.w, h: sz.h };
    }
    const pdf = await buildPdf(ctx, xl.agg, receiptsSorted, images);
    return { xlsx: xl, pdf };
  }

  window.Exporter = { exportMonth, buildExcel, base64ToBuf, MONTHS, eur, num, pad, fmtDate };
})();
