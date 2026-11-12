/* Belegaufnahme + Texterkennung
 * - Fotos/Scans werden verkleinert (max. 2000 px) und als JPEG gespeichert
 * - PDF-Belege (z. B. Hotelrechnung per Mail) werden seitenweise in Bilder umgewandelt
 * - KI-Erkennung über die Claude API (direkt vom Gerät, eigener API-Key)
 * - Offline-Fallback: Tesseract (läuft komplett auf dem Gerät)
 */
(function () {
  'use strict';
  const CAT_KEYS = Object.keys(RB003.CATEGORIES).filter((k) => !RB003.CATEGORIES[k].manualOnly);

  // ---------- Bilder ----------
  function loadImage(src) {
    return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('Bild konnte nicht gelesen werden')); i.src = src; });
  }
  function canvasToBlob(c, type = 'image/jpeg', q = 0.85) { return new Promise((res) => c.toBlob(res, type, q)); }

  async function imageFileToJpeg(file, maxSide = 2000) {
    const url = URL.createObjectURL(file);
    try {
      const img = await loadImage(url);
      const scale = Math.min(1, maxSide / Math.max(img.naturalWidth, img.naturalHeight));
      const c = document.createElement('canvas');
      c.width = Math.round(img.naturalWidth * scale); c.height = Math.round(img.naturalHeight * scale);
      const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      ctx.drawImage(img, 0, 0, c.width, c.height);
      return { blob: await canvasToBlob(c), w: c.width, h: c.height };
    } finally { URL.revokeObjectURL(url); }
  }

  let pdfjsReady = null;
  function loadScript(src) {
    return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Konnte ' + src + ' nicht laden')); document.head.appendChild(s); });
  }
  function ensurePdfjs() {
    if (!pdfjsReady) pdfjsReady = loadScript('vendor/pdf.min.js').then(() => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js'; });
    return pdfjsReady;
  }
  async function pdfFileToJpegs(file, maxPages = 6) {
    await ensurePdfjs();
    const pdf = await window.pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    const out = [];
    for (let p = 1; p <= Math.min(pdf.numPages, maxPages); p++) {
      const page = await pdf.getPage(p);
      const vp0 = page.getViewport({ scale: 1 });
      const scale = Math.min(2.5, 2000 / Math.max(vp0.width, vp0.height));
      const vp = page.getViewport({ scale });
      const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
      const ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      out.push({ blob: await canvasToBlob(c), w: c.width, h: c.height });
    }
    return out;
  }

  async function fileToImages(file) {
    if (file.type === 'application/pdf' || /\.pdf$/i.test(file.name)) return pdfFileToJpegs(file);
    return [await imageFileToJpeg(file)];
  }

  function blobToBase64(blob) {
    return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.onerror = rej; r.readAsDataURL(blob); });
  }

  // ---------- KI-Erkennung (Claude API) ----------
  function buildPrompt() {
    const cats = CAT_KEYS.map((k) => `- "${k}": ${RB003.CATEGORIES[k].label}`).join('\n');
    return `Du bist ein Assistent für die deutsche Reisekostenabrechnung. Lies den Beleg auf dem Bild/den Bildern aus.
Antworte AUSSCHLIESSLICH mit einem JSON-Objekt, ohne Erklärungstext, ohne Markdown, nach diesem Schema:
{
  "date": "YYYY-MM-DD" | null,            // Belegdatum (bei Hotel: Abreise-/Rechnungsdatum)
  "merchant": string | null,              // Händler/Firma, kurz
  "description": string | null,           // max. 25 Zeichen, z. B. "Hotel 2 Nächte", "Tanken Diesel", "Parkhaus"
  "total_gross": number | null,           // Endbetrag brutto (zu zahlen), Punkt als Dezimaltrenner
  "currency": "EUR" | "CHF" | ...,        // ISO-Code
  "vat_rates": [{"rate": number, "gross": number}],  // ausgewiesene MwSt-Sätze, leer wenn keine
  "category": string,                     // genau einer der Schlüssel unten
  "liters": number | null,                // nur bei Kraftstoff/Öl
  "hotel": {"checkin": "YYYY-MM-DD" | null, "checkout": "YYYY-MM-DD" | null, "nights": number | null, "breakfast_included": boolean} | null,
  "confidence": number                    // 0..1, wie sicher bist du bei Betrag und Datum
}
Kategorien (Spalten der Reisekostenabrechnung RB003):
${cats}
Regeln: Hotel/Übernachtung -> "uebernachtung". Tankstelle -> "dw_fuel" (außer klar Mietwagen -> "mw_fuel"). Mietwagenrechnung -> "mw_miete".
Bahn, Flug, Fernbus -> "oepnv_ueber50"; Taxi, U-Bahn, Bus, Straßenbahn im Ort -> "oepnv_bis50".
Restaurant mit Bewirtung von Geschäftspartnern -> "bewirtung"; eigenes Essen -> "verpflegung".
Parken: mit ausgewiesener 19 % MwSt -> "neben19", sonst "neben0". Porto -> "neben0". Fachliteratur -> "neben7".`;
  }

  async function recognizeAI(blobs, settings) {
    if (!settings.apiKey) throw new Error('Kein API-Key hinterlegt');
    const content = [];
    for (const b of blobs.slice(0, 4)) content.push({ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: await blobToBase64(b) } });
    content.push({ type: 'text', text: buildPrompt() });
    const ctrl = new AbortController(); const to = setTimeout(() => ctrl.abort(), 60000);
    let res;
    try {
      res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST', signal: ctrl.signal,
        headers: {
          'content-type': 'application/json',
          'x-api-key': settings.apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({ model: settings.aiModel || 'claude-haiku-5-5', max_tokens: 1024, messages: [{ role: 'user', content }] }),
      });
    } finally { clearTimeout(to); }
    if (!res.ok) {
      let msg = 'HTTP ' + res.status;
      try { const j = await res.json(); if (j.error && j.error.message) msg += ': ' + j.error.message; } catch (e) { /* ignore */ }
      throw new Error('KI-Erkennung fehlgeschlagen (' + msg + ')');
    }
    const j = await res.json();
    const text = (j.content || []).filter((c) => c.type === 'text').map((c) => c.text).join('');
    const m = /\{[\s\S]*\}/.exec(text);
    if (!m) throw new Error('KI-Antwort enthielt kein JSON');
    const r = JSON.parse(m[0]);
    return normalize({ ...r, engine: 'KI' });
  }

  // ---------- Google Cloud Vision (Texterkennung in der Cloud) ----------
  async function recognizeVision(blobs, settings) {
    if (!settings.visionKey) throw new Error('Kein Google-Vision-API-Key hinterlegt');
    const requests = [];
    for (const b of blobs.slice(0, 3)) requests.push({ image: { content: await blobToBase64(b) }, features: [{ type: 'DOCUMENT_TEXT_DETECTION' }], imageContext: { languageHints: ['de', 'en', 'fr'] } });
    const ctrl = new AbortController(); const to = setTimeout(() => ctrl.abort(), 45000);
    let res;
    try {
      res = await fetch('https://vision.googleapis.com/v1/images:annotate?key=' + encodeURIComponent(settings.visionKey), {
        method: 'POST', signal: ctrl.signal, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ requests }),
      });
    } finally { clearTimeout(to); }
    let j = null; try { j = await res.json(); } catch (e) { /* ignore */ }
    if (!res.ok) throw new Error('Google Vision fehlgeschlagen (HTTP ' + res.status + (j && j.error && j.error.message ? ': ' + j.error.message : '') + ')');
    const parts = (j.responses || []).map((r) => {
      if (r.error) throw new Error('Google Vision: ' + r.error.message);
      return (r.fullTextAnnotation && r.fullTextAnnotation.text) || (r.textAnnotations && r.textAnnotations[0] && r.textAnnotations[0].description) || '';
    });
    const text = parts.join('\n');
    if (!text.trim()) throw new Error('Google Vision hat keinen Text gefunden');
    const r = parseReceiptText(text);
    r.confidence = r.total_gross != null && r.date ? 0.75 : 0.4;
    return normalize({ ...r, engine: 'Google Vision', rawText: text });
  }

  // ---------- Offline-OCR (Tesseract) ----------
  let workerP = null;
  function abs(p) { return new URL(p, location.href).href; }
  async function getWorker(onProgress) {
    if (!workerP) {
      workerP = (async () => {
        if (!window.Tesseract) await loadScript('vendor/tesseract/tesseract.min.js');
        return window.Tesseract.createWorker('deu', 1, {
          workerPath: abs('vendor/tesseract/worker.min.js'),
          corePath: abs('vendor/tesseract/'),
          langPath: abs('vendor/tesseract/lang'),
          gzip: true,
          logger: (m) => { if (onProgress && m && typeof m.progress === 'number') onProgress(m); },
        });
      })();
      workerP.catch(() => { workerP = null; });
    }
    return workerP;
  }

  async function prepForOcr(blob) {
    const url = URL.createObjectURL(blob);
    try {
      const img = await loadImage(url);
      const scale = Math.min(2, 1800 / img.naturalWidth);
      const c = document.createElement('canvas'); c.width = Math.round(img.naturalWidth * scale); c.height = Math.round(img.naturalHeight * scale);
      const ctx = c.getContext('2d'); ctx.drawImage(img, 0, 0, c.width, c.height);
      const d = ctx.getImageData(0, 0, c.width, c.height); const px = d.data;
      // Graustufen + Kontrastspreizung
      let min = 255, max = 0; const g = new Uint8ClampedArray(px.length / 4);
      for (let i = 0, j = 0; i < px.length; i += 4, j++) { const v = 0.299 * px[i] + 0.587 * px[i + 1] + 0.114 * px[i + 2]; g[j] = v; if (v < min) min = v; if (v > max) max = v; }
      const span = Math.max(40, max - min);
      for (let i = 0, j = 0; i < px.length; i += 4, j++) { const v = ((g[j] - min) / span) * 255; px[i] = px[i + 1] = px[i + 2] = v; }
      ctx.putImageData(d, 0, 0);
      return c;
    } finally { URL.revokeObjectURL(url); }
  }

  async function recognizeOffline(blobs, onProgress) {
    const worker = await getWorker(onProgress);
    let text = '';
    for (const b of blobs.slice(0, 3)) {
      const c = await prepForOcr(b);
      const r = await worker.recognize(c);
      text += '\n' + r.data.text;
    }
    return normalize({ ...parseReceiptText(text), engine: 'Offline-OCR', rawText: text });
  }

  // ---------- Text-Heuristiken für Offline-OCR ----------
  function toNum(s) {
    if (s == null) return null;
    let t = String(s).replace(/[^\d,.\-]/g, '');
    if (/,\d{2}$/.test(t)) t = t.replace(/\./g, '').replace(',', '.');
    else if (/\.\d{2}$/.test(t)) t = t.replace(/,/g, '');
    else t = t.replace(',', '.');
    const n = parseFloat(t); return isFinite(n) ? n : null;
  }
  const AMT = /(\d{1,3}(?:[.\s]\d{3})*[,.]\d{2})(?!\d)/g;

  function parseReceiptText(text) {
    const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    const lower = text.toLowerCase();
    // Datum
    let date = null;
    const dm = /(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/.exec(text);
    if (dm) { let y = +dm[3]; if (y < 100) y += 2000; const mo = +dm[2], d = +dm[1]; if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) date = `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`; }
    // Betrag: Zeilen mit Schlüsselwörtern bevorzugen, sonst größter Betrag
    const key = /(summe|gesamt|total|zu zahlen|zahlbetrag|endbetrag|betrag|bar|ec-?karte|kartenzahlung|visa|mastercard)/i;
    let best = null;
    for (const l of lines) {
      if (key.test(l) && !/(mwst|ust|netto|steuer|rückgeld|gegeben)/i.test(l)) {
        const all = [...l.matchAll(AMT)].map((m) => toNum(m[1])).filter((n) => n != null);
        if (all.length) { const v = Math.max(...all); if (!best || /summe|gesamt|total|zu zahlen/i.test(l)) best = v; }
      }
    }
    if (best == null) { const all = [...text.matchAll(AMT)].map((m) => toNum(m[1])).filter((n) => n != null && n < 10000); if (all.length) best = Math.max(...all); }
    // MwSt
    const vat = [];
    if (/19[,.]?0?0?\s?%/.test(text)) vat.push({ rate: 19, gross: null });
    if (/\b7[,.]?0?0?\s?%/.test(text)) vat.push({ rate: 7, gross: null });
    // Liter
    let liters = null; const lm = /(\d{1,3}[,.]\d{1,3})\s?(?:l\b|liter|ltr)/i.exec(text); if (lm) liters = toNum(lm[1]);
    // Kategorie
    let category = 'neben19';
    if (/(hotel|übernachtung|uebernachtung|logis|zimmer|room|nacht|night)/.test(lower)) category = 'uebernachtung';
    else if (/(diesel|super|benzin|e10|e5|kraftstoff|tankstelle|aral|shell|esso|jet\b|total ?energies|avia|agip|eni\b|omv|bft|star tank|zapfs)/.test(lower)) category = 'dw_fuel';
    else if (/(sixt|europcar|hertz|avis|enterprise|mietwagen|rental)/.test(lower)) category = 'mw_miete';
    else if (/(deutsche bahn|db fernverkehr|bahncard|ice\b|ic\b|flug|lufthansa|eurowings|ryanair|flixbus|boarding)/.test(lower)) category = 'oepnv_ueber50';
    else if (/(taxi|kvb|vrs|bvg|mvv|hvv|rmv|vrr|einzelfahrt|tageskarte|u-bahn|straßenbahn|bus\b)/.test(lower)) category = 'oepnv_bis50';
    else if (/(park|apcoa|contipark|q-park|parkhaus|parkschein)/.test(lower)) category = vat.length ? 'neben19' : 'neben0';
    else if (/(restaurant|gaststätte|bewirtung|trattoria|ristorante|bistro|café|cafe|bar\b|speisen|getränke)/.test(lower)) category = 'verpflegung';
    else if (/(porto|briefmarke|deutsche post)/.test(lower)) category = 'neben0';
    else if (/(buch|fachliteratur|zeitschrift)/.test(lower) && vat.some((v) => v.rate === 7)) category = 'neben7';
    // Händler: erste "sinnvolle" Zeile
    const merchant = (lines.find((l) => /[a-zA-ZäöüÄÖÜ]{3,}/.test(l) && !/^(rechnung|beleg|quittung|kassenbon)/i.test(l)) || '').slice(0, 40) || null;
    return { date, merchant, description: null, total_gross: best, currency: /\bchf\b/i.test(text) ? 'CHF' : 'EUR', vat_rates: vat, category, liters, hotel: null, confidence: best != null && date ? 0.5 : 0.25 };
  }

  function normalize(r) {
    const out = { ...r };
    out.total_gross = r.total_gross == null ? null : Math.round(Number(r.total_gross) * 100) / 100;
    if (!CAT_KEYS.includes(out.category)) out.category = 'neben19';
    out.currency = (r.currency || 'EUR').toUpperCase().slice(0, 3);
    if (out.date && !/^\d{4}-\d{2}-\d{2}$/.test(out.date)) out.date = null;
    return out;
  }

  // ---------- Steuerung ----------
  // mode: 'auto' | 'ai' | 'vision' | 'offline' | 'manual'
  // Automatisch: Claude-KI → Google Vision → Offline-OCR (je nachdem, welche Keys hinterlegt sind und ob Internet da ist)
  async function recognize(blobs, settings, onProgress) {
    let mode = settings.ocrMode || 'auto';
    if (mode === 'manual') mode = 'auto'; // „Text erkennen" auf Knopfdruck
    const online = navigator.onLine !== false;
    const chain = [];
    if (mode === 'ai') chain.push('ai');
    else if (mode === 'vision') chain.push('vision');
    else if (mode === 'offline') chain.push('offline');
    else { if (settings.apiKey && online) chain.push('ai'); if (settings.visionKey && online) chain.push('vision'); chain.push('offline'); }
    const warn = [];
    for (let i = 0; i < chain.length; i++) {
      try {
        const r = chain[i] === 'ai' ? await recognizeAI(blobs, settings) : chain[i] === 'vision' ? await recognizeVision(blobs, settings) : await recognizeOffline(blobs, onProgress);
        if (warn.length) r.warning = warn.join(' · ') + ' – ' + r.engine + ' verwendet';
        return r;
      } catch (e) {
        if (i === chain.length - 1) throw e;
        console.warn(e); warn.push(e.message);
      }
    }
  }

  async function warmupOffline(onProgress) { await getWorker(onProgress); await ensurePdfjs(); }

  window.OCR = { fileToImages, recognize, recognizeAI, recognizeVision, recognizeOffline, parseReceiptText, warmupOffline, blobToBase64 };
})();
