/* Reisekosten RB003 – App-Logik */
(function () {
  'use strict';
  const C = RB003.CATEGORIES;
  const E = window.Exporter;
  const { eur, num, pad, fmtDate, MONTHS } = E;
  const WD = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
  const CURRENCIES = ['EUR', 'CHF', 'GBP', 'USD', 'DKK', 'SEK', 'NOK', 'PLN', 'CZK', 'HUF', 'RON', 'TRY', 'AED', 'JPY', 'CNY'];
  const SHEET1 = ['verpflegung', 'uebernachtung', 'bewirtung', 'geschenk19', 'geschenk7', 'neben19', 'neben7', 'neben0'];
  const SHEET2 = ['dw_fuel', 'dw_oel', 'dw_garage', 'dw_pflege', 'dw_reparatur', 'mw_fuel', 'mw_miete', 'oepnv_bis50', 'oepnv_ueber50'];

  const $ = (sel, el = document) => el.querySelector(sel);
  const h = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const r2 = RB003.r2;
  const parseNum = (v) => { if (v === '' || v == null) return null; const n = parseFloat(String(v).replace(/\./g, (m, i, s) => (/,/.test(s) ? '' : m)).replace(',', '.')); return isFinite(n) ? n : null; };
  const numIn = (n, d = 2) => (n == null || n === '' ? '' : String(r2(n)).replace('.', ','));

  const S = {
    settings: {}, monthKey: '', month: { days: {}, vehicle: {} }, receipts: [], view: 'overview',
    countries: [], tplInfo: null, templateBuf: null, filter: 'all', imgUrls: new Map(), queue: Promise.resolve(), processing: 0, lastExport: null,
  };

  // ---------- Hilfen Monat ----------
  const ym = (key) => { const [y, m] = key.split('-').map(Number); return { y, m }; };
  const keyOf = (y, m) => `${y}-${pad(m)}`;
  const daysIn = (key) => { const { y, m } = ym(key); return new Date(y, m, 0).getDate(); };
  const wd = (key, d) => { const { y, m } = ym(key); return new Date(y, m - 1, d).getDay(); };
  const monthName = (key) => { const { y, m } = ym(key); return `${MONTHS[m - 1]} ${y}`; };
  const shiftMonth = (key, delta) => { const { y, m } = ym(key); const d = new Date(y, m - 1 + delta, 1); return keyOf(d.getFullYear(), d.getMonth() + 1); };
  const todayKey = () => { const d = new Date(); return keyOf(d.getFullYear(), d.getMonth() + 1); };

  function toast(msg, ms = 2600) {
    const t = $('#toast'); t.textContent = msg; t.classList.remove('hidden');
    clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.add('hidden'), ms);
  }

  // ---------- Daten laden/speichern ----------
  async function loadSettings() {
    const s = (await DB.get('settings', 'profile')) || {};
    S.settings = Object.assign({ name: '', costCenter: '', kmRate: 0.30, ocrMode: 'auto', apiKey: '', aiModel: 'claude-haiku-5-5' }, s);
  }
  const saveSettings = () => DB.put('settings', S.settings, 'profile');

  async function loadTemplate() {
    const custom = await DB.get('settings', 'template');
    S.templateBuf = custom || E.base64ToBuf(window.RB003_TEMPLATE_B64);
    S.tplInfo = await RB003.readTemplateInfo(window.JSZip, S.templateBuf);
    S.tplInfo.custom = !!custom;
    S.countries = S.tplInfo.countries;
  }

  async function loadMonth(key) {
    S.monthKey = key;
    S.month = (await DB.get('months', key)) || { days: {}, vehicle: {} };
    S.month.days = S.month.days || {}; S.month.vehicle = S.month.vehicle || {};
    S.receipts = (await DB.receiptsOfMonth(key)) || [];
    localStorage.setItem('rk_month', key);
  }
  const saveMonth = () => DB.put('months', S.month, S.monthKey);

  async function saveReceipt(r) {
    await DB.put('receipts', r);
    const i = S.receipts.findIndex((x) => x.id === r.id);
    if (r.monthKey !== S.monthKey) { if (i >= 0) S.receipts.splice(i, 1); }
    else if (i >= 0) S.receipts[i] = r; else S.receipts.push(r);
  }

  // Belegnummern: fortlaufend nach Tag/Datum – identisch in App, Excel und PDF
  function numbered() {
    const rank = (r) => (r.day && r.status !== 'error' && r.status !== 'processing' ? 0 : 1);
    const list = S.receipts.slice().sort((a, b) => rank(a) - rank(b) || (a.day || 99) - (b.day || 99) || String(a.date || '').localeCompare(String(b.date || '')) || a.createdAt - b.createdAt);
    list.forEach((r, i) => { r.belegNr = 'B' + pad(i + 1); });
    return list;
  }

  function exportData() {
    const { y, m } = ym(S.monthKey);
    const list = numbered();
    return {
      month: pad(m), year: String(y), name: S.settings.name, costCenter: S.settings.costCenter, kmRate: Number(S.settings.kmRate) || 0.30,
      days: S.month.days, vehicle: S.month.vehicle,
      receipts: list.filter((r) => r.status !== 'processing' && r.status !== 'error' && r.day).map((r) => ({
        day: r.day, category: r.category, amountEur: Number(r.amountEur) || 0, liters: r.liters, belegNr: r.belegNr, text: shortLabel(r),
      })),
    };
  }
  const shortLabel = (r) => (r.text || r.merchant || '').slice(0, 24);
  const agg = () => RB003.aggregate(exportData(), S.countries);

  async function imgUrl(id) {
    if (S.imgUrls.has(id)) return S.imgUrls.get(id);
    const b = await DB.get('images', id); if (!b) return '';
    const u = URL.createObjectURL(b); S.imgUrls.set(id, u); return u;
  }
  async function hydrateThumbs(root) {
    for (const el of root.querySelectorAll('img[data-img]')) { const u = await imgUrl(el.dataset.img); if (u) el.src = u; }
  }

  // ---------- Rendering ----------
  function setView(v) {
    S.view = v;
    document.querySelectorAll('#tabs button').forEach((b) => b.classList.toggle('active', b.dataset.view === v));
    render();
    window.scrollTo(0, 0);
  }

  function render() {
    const a = agg();
    $('#mLabel').textContent = monthName(S.monthKey);
    $('#mTotal').textContent = eur(a.total);
    $('#fab').classList.toggle('hidden', !['overview', 'receipts'].includes(S.view));
    const v = $('#view');
    const fn = { overview: viewOverview, receipts: viewReceipts, days: viewDays, vehicle: viewVehicle, settings: viewSettings }[S.view];
    v.innerHTML = fn(a);
    hydrateThumbs(v);
  }

  function isIOS() { return /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1); }
  function isStandalone() { return navigator.standalone === true || matchMedia('(display-mode: standalone)').matches; }

  function warnings(a) {
    const w = [];
    const proc = S.receipts.filter((r) => r.status === 'processing').length;
    if (proc) w.push(['info', `<span class="spin"></span> ${proc} Beleg${proc > 1 ? 'e werden' : ' wird'} erkannt …`]);
    const rev = S.receipts.filter((r) => r.status === 'review').length;
    if (rev) w.push(['warn', `${rev} Beleg${rev > 1 ? 'e' : ''} noch nicht geprüft.`, 'receipts:review']);
    const err = S.receipts.filter((r) => r.status === 'error').length;
    if (err) w.push(['err', `${err} Beleg${err > 1 ? 'e konnten' : ' konnte'} nicht erkannt werden – bitte manuell erfassen.`, 'receipts:review']);
    const noday = S.receipts.filter((r) => r.status !== 'processing' && !r.day).length;
    if (noday) w.push(['warn', `${noday} Beleg${noday > 1 ? 'e' : ''} ohne Datum – wird nicht in die Excel übernommen.`, 'receipts:review']);
    const bew = S.receipts.filter((r) => r.category === 'bewirtung' && (!r.occasion || !r.participants)).length;
    if (bew) w.push(['warn', `Bewirtung: Anlass und Teilnehmer fehlen bei ${bew} Beleg${bew > 1 ? 'en' : ''} (Pflicht laut Reisekostenordnung).`, 'receipts:all']);
    for (let d = 1; d <= 31; d++) { const v = a.veh[d]; if (v.kmStart != null && v.kmEnd != null && v.kmEnd < v.kmStart) w.push(['err', `Fahrzeug Tag ${d}: km-Stand Ende ist kleiner als Anfang.`, 'vehicle']); }
    if (!S.settings.name) w.push(['warn', 'Name und Kostenstelle fehlen (Pflichtfelder der RB003).', 'settings']);
    return w;
  }

  function viewOverview(a) {
    const nDays = Object.values(a.days).filter((d) => d.pd.hours != null).length;
    const veh = r2(a.sum2);
    let html = '';
    if (isIOS() && !isStandalone()) html += `<div class="notice info"><div><b>Als App installieren:</b> In Safari auf <b>Teilen</b> ⬆︎ tippen → <b>„Zum Home-Bildschirm"</b>. Erst dann bleiben deine Daten dauerhaft gespeichert und die App funktioniert offline.</div></div>`;
    const open = S.receipts.filter((r) => r.status !== 'ok').length;
    html += `<div class="kpis">
      <div class="kpi total"><div class="l">Reisekosten insgesamt</div><div class="v">${eur(a.total)}</div></div>
      <div class="kpi"><div class="l">Verpflegungspauschalen</div><div class="v">${eur(a.pauschSum)}</div><div class="muted small">${nDays} Reisetag${nDays === 1 ? '' : 'e'}</div></div>
      <div class="kpi"><div class="l">Belege Blatt 1</div><div class="v">${eur(r2(a.sum1 - a.pauschSum))}</div><div class="muted small">Hotel, Bewirtung, Nebenkosten …</div></div>
      <div class="kpi"><div class="l">Seite 2</div><div class="v">${eur(veh)}</div><div class="muted small">Fahrzeug, km, Bahn/ÖPNV</div></div>
      <div class="kpi"><div class="l">Belege erfasst</div><div class="v">${S.receipts.length}</div><div class="muted small">${open ? open + ' zu prüfen' : 'alle geprüft'}</div></div>
    </div>`;
    for (const [type, msg, go] of warnings(a)) html += `<div class="notice ${type}" ${go ? `data-go="${go}" style="cursor:pointer"` : ''}><div>${msg}</div></div>`;
    html += `<div class="card pad"><div class="btns">
      <button class="btn accent" data-act="camera">📷 Beleg fotografieren</button>
      <button class="btn" data-act="files">Aus Fotos / PDF</button>
    </div><div class="btns" style="margin-top:10px">
      <button class="btn" data-act="trip">Reise anlegen</button>
      <button class="btn primary" data-act="export">Monat exportieren</button>
    </div></div>`;
    // Spaltensummen
    const rows = [['1', 'Verpflegung Pauschale', a.pauschSum]];
    for (const k of ['verpflegung', 'uebern_pausch', 'uebernachtung', 'bewirtung', 'geschenk19', 'geschenk7', 'neben19', 'neben7', 'neben0']) if (a.col[k]) rows.push([C[k].nr, C[k].label, a.col[k]]);
    const dw = ['dw_fuel', 'dw_oel', 'dw_garage', 'dw_pflege', 'dw_reparatur'].reduce((s, k) => s + (a.col[k] || 0), 0);
    if (dw) rows.push([11, 'Dienstwagen', dw]);
    if ((a.col.mw_fuel || 0) + (a.col.mw_miete || 0)) rows.push([12, 'Mietwagen', (a.col.mw_fuel || 0) + (a.col.mw_miete || 0)]);
    if (a.privSum) rows.push([13, 'Privatwagen km-Pauschale', a.privSum]);
    if (a.col.oepnv_bis50) rows.push([14, 'ÖPNV bis 50 km', a.col.oepnv_bis50]);
    if (a.col.oepnv_ueber50) rows.push([15, 'Bahn/Flug über 50 km', a.col.oepnv_ueber50]);
    html += `<h3>Spalten der RB003</h3><div class="card list">${rows.map(([n, l, v]) => `<div class="item" style="cursor:default;min-height:46px"><span class="chip">${n}</span><div class="grow">${h(l)}</div><div class="amt">${eur(v)}</div></div>`).join('')}
      <div class="item" style="cursor:default;min-height:46px"><div class="grow"><b>Reisekosten insgesamt</b></div><div class="amt">${eur(a.total)}</div></div></div>`;
    const recent = numbered().slice(-4).reverse();
    if (recent.length) html += `<h3>Zuletzt erfasst</h3><div class="card list">${recent.map(receiptItem).join('')}</div>`;
    return html;
  }

  function statusChip(r) {
    if (r.status === 'processing') return '<span class="chip"><span class="spin"></span> erkennt</span>';
    if (r.status === 'error') return '<span class="chip err">Fehler</span>';
    if (r.status === 'review') return '<span class="chip warn">prüfen</span>';
    return '<span class="chip ok">✓</span>';
  }
  function receiptItem(r) {
    const cat = C[r.category];
    const title = r.merchant || r.text || (cat ? cat.label : 'Beleg');
    const sub = [r.belegNr, r.day ? `${pad(r.day)}.${pad(ym(S.monthKey).m)}.` : 'ohne Datum', cat ? `Sp. ${cat.nr} ${cat.short}` : ''].filter(Boolean).join(' · ');
    const fx = r.currency && r.currency !== 'EUR' && r.amount != null ? `<div class="muted small right">${num(r.amount)} ${h(r.currency)}</div>` : '';
    return `<div class="item" data-receipt="${r.id}"><img class="thumb" data-img="${r.images[0] || ''}" alt="">
      <div class="grow"><div class="t">${h(title)}</div><div class="s">${h(sub)}</div></div>
      <div class="right"><div class="amt">${r.amountEur != null ? eur(r.amountEur) : '–'}</div>${fx}<div style="margin-top:3px">${statusChip(r)}</div></div></div>`;
  }

  function viewReceipts() {
    const list = numbered().filter((r) => S.filter === 'all' || r.status !== 'ok');
    let html = `<div class="row between"><h2>Belege</h2><button class="btn sm" data-act="files">Fotos / PDF</button></div>
      <div class="seg" style="margin-bottom:12px"><button data-filter="all" class="${S.filter === 'all' ? 'on' : ''}">Alle (${S.receipts.length})</button><button data-filter="review" class="${S.filter === 'review' ? 'on' : ''}">Zu prüfen (${S.receipts.filter((r) => r.status !== 'ok').length})</button></div>`;
    if (!list.length) return html + `<div class="card empty"><div class="big">🧾</div><div>${S.filter === 'all' ? 'Noch keine Belege in diesem Monat.<br>Tippe auf <b>＋</b>, um einen Beleg zu fotografieren.' : 'Alles geprüft.'}</div></div>`;
    let cur = null; let open = false;
    for (const r of list) {
      const d = r.day || 0;
      if (d !== cur) {
        if (open) html += '</div>';
        cur = d; open = true;
        html += `<h3>${d ? `${WD[wd(S.monthKey, d)]}, ${pad(d)}.${pad(ym(S.monthKey).m)}.` : 'Ohne Datum'}</h3><div class="card list">`;
      }
      html += receiptItem(r);
    }
    if (open) html += '</div>';
    return html;
  }

  function dayInfo(d) {
    const dd = S.month.days[d];
    if (!dd || dd.startH === '' || dd.startH == null) return null;
    return { dd, pd: RB003.perDiem(dd, S.countries) };
  }
  const tstr = (H, M) => `${pad(H)}:${pad(M || 0)}`;

  function viewDays(a) {
    const n = daysIn(S.monthKey);
    let html = `<div class="row between"><h2>Reisetage</h2><button class="btn sm primary" data-act="trip">＋ Reise anlegen</button></div>
      <div class="muted small" style="margin:-4px 4px 12px">Verpflegungspauschalen werden wie in der RB003 berechnet (inkl. Kürzung bei gestellten Mahlzeiten: Frühstück 20 %, Mittag-/Abendessen je 40 %).</div><div class="card list">`;
    for (let d = 1; d <= n; d++) {
      const info = dayInfo(d); const w = wd(S.monthKey, d);
      const cls = info ? 'on' : (w === 0 || w === 6 ? 'we' : '');
      let t = '<span class="muted">–</span>', s = '', amt = '';
      if (info) {
        const { dd, pd } = info;
        t = h(dd.destination || 'Reisetag');
        const meals = [dd.breakfast && 'F', dd.lunch && 'M', dd.dinner && 'A'].filter(Boolean).join('/');
        s = `${tstr(dd.startH, dd.startM)}–${tstr(dd.endH, dd.endM)} · ${num(pd.hours, 1)} Std. · ${h(pd.country.name)}${dd.arrival ? ' · An/Ab' : ''}${meals ? ' · Kürzung ' + meals : ''}`;
        amt = typeof pd.display === 'number' ? eur(pd.display) : '0,00 €';
      }
      const recN = S.receipts.filter((r) => r.day === d).length;
      html += `<div class="item" data-day="${d}"><div class="daybadge ${cls}"><b>${d}</b><span>${WD[w]}</span></div>
        <div class="grow"><div class="t">${t}</div>${s ? `<div class="s">${s}</div>` : ''}</div>
        <div class="right"><div class="amt">${amt}</div>${recN ? `<div class="muted small">${recN} Beleg${recN > 1 ? 'e' : ''}</div>` : ''}</div></div>`;
    }
    return html + `</div><div class="card pad"><div class="row between"><b>Summe Pauschalen</b><b class="num">${eur(a.pauschSum)}</b></div></div>`;
  }

  function viewVehicle(a) {
    const n = daysIn(S.monthKey);
    const shown = [];
    for (let d = 1; d <= n; d++) {
      const v = S.month.vehicle[d]; const hasV = v && Object.values(v).some((x) => x !== '' && x != null);
      const hasR = S.receipts.some((r) => r.day === d && C[r.category] && C[r.category].sheet === 2);
      if (hasV || hasR || dayInfo(d)) shown.push(d);
    }
    let html = `<div class="row between"><h2>Fahrzeug</h2><button class="btn sm primary" data-act="vehicle-add">＋ Tag</button></div>
      <div class="muted small" style="margin:-4px 4px 12px">km-Stände und gefahrene km für Seite 2 der RB003. Tank-, Mietwagen- und ÖPNV-Belege kommen automatisch aus der Belegerfassung.</div>`;
    if (!shown.length) return html + `<div class="card empty"><div class="big">🚗</div>Keine Fahrten in diesem Monat.<br>Lege zuerst Reisetage an oder tippe auf <b>＋ Tag</b>.</div>`;
    html += '<div class="card list">';
    for (const d of shown) {
      const v = a.veh[d]; const parts = [];
      if (v.kmStart != null || v.kmEnd != null) parts.push(`DW ${v.kmStart != null ? v.kmStart.toLocaleString('de-DE') : '?'} → ${v.kmEnd != null ? v.kmEnd.toLocaleString('de-DE') : '?'}${v.dwBusiness != null ? ` (${v.dwBusiness} km gesch.${v.dwPrivate ? `, ${v.dwPrivate} privat` : ''})` : ''}`);
      if (v.mwBusiness != null) parts.push(`Mietwagen ${v.mwBusiness} km`);
      if (v.privKm) parts.push(`Privat-PKW ${v.privKm} km`);
      const cats = SHEET2.filter((k) => a.days[d].cats[k]).map((k) => C[k].short);
      const sum = r2(SHEET2.reduce((s, k) => s + (a.days[d].cats[k] ? a.days[d].cats[k].amount : 0), 0) + (v.privAmount || 0));
      const w = wd(S.monthKey, d);
      html += `<div class="item" data-vehicle="${d}"><div class="daybadge ${dayInfo(d) ? 'on' : ''}"><b>${d}</b><span>${WD[w]}</span></div>
        <div class="grow"><div class="t">${parts.length ? h(parts[0]) : '<span class="muted">km eintragen</span>'}</div><div class="s">${h([...parts.slice(1), ...cats].join(' · ') || (dayInfo(d) ? h(S.month.days[d].destination || '') : ''))}</div></div>
        <div class="amt">${sum ? eur(sum) : ''}</div></div>`;
    }
    return html + `</div><div class="card pad"><div class="row between"><b>Summe Seite 2 (Spalten 11–15)</b><b class="num">${eur(a.sum2)}</b></div></div>`;
  }

  function viewSettings() {
    const s = S.settings; const t = S.tplInfo;
    return `<h2>Einstellungen</h2>
    <h3>Profil (Kopf der RB003)</h3>
    <div class="card form">
      <div class="field"><label>Name</label><input type="text" id="sName" value="${h(s.name)}" placeholder="Vor- und Nachname" autocomplete="name"></div>
      <div class="field"><label>Kostenstelle</label><input type="text" id="sCost" value="${h(s.costCenter)}" inputmode="numeric" placeholder="z. B. 720120"></div>
      <div class="field"><label>km-Pauschale Privatwagen (€/km)</label><input type="text" id="sKm" value="${num(s.kmRate)}" inputmode="decimal"><div class="hint">Spalte 13 auf Seite 2. Üblich: 0,30 €/km – bitte mit eurer Reisekostenordnung abgleichen.</div></div>
    </div>
    <h3>Texterkennung</h3>
    <div class="card form">
      <div class="field"><label>Modus</label><div class="seg" id="sOcr">
        <button data-ocr="auto" class="${s.ocrMode === 'auto' ? 'on' : ''}">Automatisch</button>
        <button data-ocr="ai" class="${s.ocrMode === 'ai' ? 'on' : ''}">Nur KI</button>
        <button data-ocr="offline" class="${s.ocrMode === 'offline' ? 'on' : ''}">Offline</button>
        <button data-ocr="manual" class="${s.ocrMode === 'manual' ? 'on' : ''}">Aus</button></div>
        <div class="hint">${s.ocrMode === 'manual' ? '<b>Manuell:</b> Beleg wird nur fotografiert, du trägst Datum, Betrag und Kategorie selbst ein. Keine Texterkennung, kein API-Key, nichts verlässt das Gerät.' : 'Automatisch: KI-Erkennung, wenn Internet und API-Key vorhanden sind – sonst Offline-Texterkennung auf dem Gerät. „Aus": nur Foto, Daten von Hand eintragen.'}</div></div>
      <div class="field"><label>Claude API-Key</label><input type="password" id="sKey" value="${h(s.apiKey)}" placeholder="sk-ant-…" autocomplete="off" autocapitalize="off" spellcheck="false">
        <div class="hint">Wird nur auf diesem Gerät gespeichert. Das Belegbild wird zur Auswertung an die Anthropic-API gesendet (ca. 1–2 Cent pro Beleg). Key erstellen unter console.anthropic.com.</div></div>
      <div class="field"><label>KI-Modell</label><select id="sModel">
        ${['claude-haiku-5-5', 'claude-sonnet-5-5'].map((m) => `<option value="${m}" ${s.aiModel === m ? 'selected' : ''}>${m === 'claude-haiku-5-5' ? 'Claude Haiku – schnell & günstig' : 'Claude Sonnet – genauer bei schwierigen Belegen'}</option>`).join('')}
        ${!['claude-haiku-5-5', 'claude-sonnet-5-5'].includes(s.aiModel) ? `<option value="${h(s.aiModel)}" selected>${h(s.aiModel)}</option>` : ''}
      </select></div>
      <div class="field"><div class="btns"><button class="btn sm" data-act="test-ai">KI-Verbindung testen</button><button class="btn sm" data-act="warmup">Offline-Erkennung laden</button></div>
        <div class="hint" id="ocrStatus">Lädt die Offline-Texterkennung (ca. 15 MB) einmalig vor, damit sie auch ohne Netz funktioniert.</div></div>
    </div>
    <h3>Vorlage</h3>
    <div class="card pad">
      <div><b>${h(t.version || 'RB003')}</b> · ${h(t.company || '')}</div>
      <div class="muted small" style="margin:4px 0 10px">${t.countries.length} Länder mit Pauschalen · ${t.custom ? 'eigene Vorlage geladen' : 'integrierte Vorlage'}</div>
      <div class="btns"><button class="btn sm" data-act="tpl-load">Neue RB003-Version laden</button>${t.custom ? '<button class="btn sm ghost" data-act="tpl-reset">Integrierte Vorlage</button>' : ''}</div>
      <div class="hint muted small" style="margin-top:8px">Wenn die Firma eine neue Version der RB003 herausgibt, hier die leere Excel laden. Pauschalen und Länder werden automatisch aus der Datei übernommen.</div>
    </div>
    <h3>Datensicherung</h3>
    <div class="card pad">
      <div class="muted small" id="storeInfo" style="margin-bottom:10px">Alle Daten liegen nur auf diesem Gerät.</div>
      <div class="btns"><button class="btn sm" data-act="backup">Backup erstellen</button><button class="btn sm" data-act="restore">Backup einspielen</button></div>
      <div class="btns" style="margin-top:10px"><button class="btn sm danger" data-act="wipe">Alle Daten löschen</button></div>
    </div>
    <div class="muted small" style="text-align:center;margin:18px 0">Reisekosten RB003 · App-Version ${APP_VERSION}</div>`;
  }

  // ---------- Sheets ----------
  function openSheet(title, body, { left = 'Schließen', right = null, onRight = null, onClose = null } = {}) {
    const wrap = $('#sheetWrap');
    wrap.innerHTML = '<div class="sheet" id="sheet"></div>'; // frisches Element: keine alten Listener
    const sh = $('#sheet');
    sh.innerHTML = `<div class="sheet-head"><button class="btn ghost sm" data-sh="close">${h(left)}</button><div class="ttl">${h(title)}</div>${right ? `<button class="btn ghost sm" data-sh="right" style="font-weight:700">${h(right)}</button>` : '<span style="width:70px"></span>'}</div><div class="sheet-body">${body}</div>`;
    wrap.classList.remove('hidden'); document.body.style.overflow = 'hidden';
    sh.scrollTop = 0;
    sh.onclick = (e) => {
      const b = e.target.closest('[data-sh]'); if (!b) return;
      if (b.dataset.sh === 'close') { closeSheet(); onClose && onClose(); }
      if (b.dataset.sh === 'right' && onRight) onRight();
    };
    hydrateThumbs(sh);
    return sh;
  }
  function closeSheet() { const w = $('#sheetWrap'); w.classList.add('hidden'); document.body.style.overflow = ''; w.innerHTML = '<div class="sheet" id="sheet"></div>'; }

  // ---------- Beleg aufnehmen ----------
  async function handleFiles(files) {
    if (!files || !files.length) return;
    try { navigator.storage && navigator.storage.persist && navigator.storage.persist(); } catch (e) { /* ignore */ }
    const manual = S.settings.ocrMode === 'manual'; let lastId = null;
    for (const f of files) {
      const r = { id: uid(), monthKey: S.monthKey, createdAt: Date.now(), images: [], status: 'processing', category: 'neben19', currency: 'EUR', amount: null, amountEur: null, day: null, date: null };
      try {
        const imgs = await OCR.fileToImages(f);
        for (const im of imgs) { const id = uid(); await DB.put('images', im.blob, id); r.images.push(id); }
      } catch (e) { toast('Datei konnte nicht gelesen werden: ' + e.message, 4000); continue; }
      if (manual) {
        r.status = 'review'; r.engine = 'manuell';
        const t = new Date();
        if (keyOf(t.getFullYear(), t.getMonth() + 1) === S.monthKey) setReceiptDate(r, `${t.getFullYear()}-${pad(t.getMonth() + 1)}-${pad(t.getDate())}`);
        await saveReceipt(r); lastId = r.id;
      } else {
        await saveReceipt(r);
        enqueueRecognition(r.id);
      }
    }
    if (S.view !== 'receipts' && S.view !== 'overview') S.view = 'receipts';
    render();
    if (manual) { if (files.length === 1 && lastId) openReceipt(lastId); else toast(`${files.length} Belege gespeichert – bitte unter „Zu prüfen" ausfüllen`, 3500); }
    else toast(files.length > 1 ? `${files.length} Belege werden erkannt …` : 'Beleg wird erkannt …');
  }

  function enqueueRecognition(id) {
    S.processing++;
    S.queue = S.queue.then(() => recognizeReceipt(id)).catch((e) => console.error(e)).finally(() => { S.processing--; });
  }

  async function recognizeReceipt(id) {
    const r = await DB.get('receipts', id); if (!r) return;
    const prevMonth = r.monthKey;
    const blobs = []; for (const iid of r.images) { const b = await DB.get('images', iid); if (b) blobs.push(b); }
    try {
      const res = await OCR.recognize(blobs, S.settings, (m) => {
        const st = document.getElementById('ocrProg'); if (st && m.status) st.textContent = `${m.status} ${Math.round(m.progress * 100)} %`;
      });
      applyRecognition(r, res);
      r.status = 'review';
    } catch (e) {
      r.status = 'error'; r.error = e.message;
    }
    await saveReceipt(r);
    if (r.monthKey !== prevMonth) toast(`Beleg wurde dem Monat ${monthName(r.monthKey)} zugeordnet`, 3500);
    if ($('#sheetWrap').classList.contains('hidden')) render();
    else { const open = $('#sheet [data-editing]'); if (open && open.dataset.editing === r.id) openReceipt(r.id); else render(); }
  }

  function applyRecognition(r, res) {
    r.engine = res.engine; r.confidence = res.confidence; r.warning = res.warning || null; r.rawText = res.rawText || null;
    if (res.merchant) r.merchant = String(res.merchant).slice(0, 40);
    if (res.description) r.text = String(res.description).slice(0, 25);
    if (res.category) r.category = res.category;
    r.currency = res.currency || 'EUR';
    if (res.total_gross != null) { r.amount = res.total_gross; if (r.currency === 'EUR') r.amountEur = res.total_gross; else r.amountEur = r.rate ? r2(res.total_gross / r.rate) : null; }
    r.vat = res.vat_rates || [];
    if (res.liters) r.liters = res.liters;
    if (res.hotel) r.hotel = res.hotel;
    if (res.date) setReceiptDate(r, res.date);
  }

  function setReceiptDate(r, iso) {
    r.date = iso || null;
    if (!iso) { r.day = null; return; }
    const [y, m, d] = iso.split('-').map(Number);
    r.monthKey = keyOf(y, m); r.day = d;
  }

  // ---------- Beleg-Editor ----------
  async function openReceipt(id) {
    const r = await DB.get('receipts', id); if (!r) return;
    numbered(); const nr = (S.receipts.find((x) => x.id === id) || {}).belegNr || '';
    const cat = C[r.category];
    const opt = (keys) => keys.map((k) => `<option value="${k}" ${r.category === k ? 'selected' : ''}>${C[k].nr} – ${h(C[k].label)}</option>`).join('');
    const vatTxt = (r.vat || []).length ? 'MwSt auf dem Beleg: ' + r.vat.map((v) => `${v.rate} %${v.gross ? ` (${num(v.gross)})` : ''}`).join(', ') : '';
    const engine = r.engine ? `<span class="chip ${r.engine === 'KI' ? 'ai' : ''}">${h(r.engine === 'manuell' ? 'manuelle Eingabe' : r.engine)}${r.confidence != null ? ` · ${Math.round(r.confidence * 100)} %` : ''}</span>` : '';
    let hotelHint = '';
    if (r.category === 'uebernachtung' && r.hotel && r.hotel.breakfast_included && r.hotel.checkin && r.hotel.checkout) {
      hotelHint = `<div class="notice warn"><div>Frühstück ist in der Hotelrechnung enthalten → die Verpflegungspauschale muss an den Folgetagen um 20 % gekürzt werden.<div style="margin-top:8px"><button class="btn sm" data-act="breakfast" data-id="${r.id}">Frühstück ${fmtDate(r.hotel.checkin)} – ${fmtDate(r.hotel.checkout)} kürzen</button></div></div></div>`;
    }
    const body = `<div data-editing="${r.id}">
      <div class="preview" data-act="zoom" data-id="${r.id}">${r.images.length ? `<img data-img="${r.images[0]}" alt="Beleg">${r.images.length > 1 ? `<span class="pager">${r.images.length} Seiten</span>` : ''}` : ''}</div>
      ${r.status === 'processing' ? `<div class="notice info"><div><span class="spin"></span> Beleg wird erkannt … <span id="ocrProg" class="small"></span></div></div>` : ''}
      ${r.status === 'error' ? `<div class="notice err"><div>Erkennung fehlgeschlagen: ${h(r.error)}<br>Bitte Felder manuell ausfüllen.</div></div>` : ''}
      ${r.warning ? `<div class="notice warn"><div>${h(r.warning)}</div></div>` : ''}
      ${r.monthKey !== S.monthKey ? `<div class="notice info"><div>Dieser Beleg gehört zum Monat <b>${monthName(r.monthKey)}</b>.</div></div>` : ''}
      ${hotelHint}
      <div class="row between" style="margin:0 2px 8px"><b>${h(nr)}</b>${engine}</div>
      <div class="card form">
        <div class="field"><label>Belegdatum</label><input type="date" id="rDate" value="${h(r.date || '')}"></div>
        <div class="field"><label>Spalte / Kostenart</label><select id="rCat"><optgroup label="Blatt 1 – Reisekostenabrechnung">${opt(SHEET1)}</optgroup><optgroup label="Seite 2 – Fahrzeug & Verkehrsmittel">${opt(SHEET2)}</optgroup></select>
          ${vatTxt ? `<div class="hint">${h(vatTxt)}</div>` : ''}</div>
        <div class="field"><label>Händler / Firma</label><input type="text" id="rMerchant" value="${h(r.merchant || '')}" placeholder="z. B. Hotel Ibis Köln"></div>
        <div class="field"><label>Kurzbeschreibung (erscheint in der Excel)</label><input type="text" id="rText" maxlength="25" value="${h(r.text || '')}" placeholder="z. B. Hotel 2 Nächte"></div>
        <div class="field"><div class="cols"><div><label>Betrag brutto</label><input type="text" id="rAmount" inputmode="decimal" value="${numIn(r.amount)}" placeholder="0,00"></div>
          <div><label>Währung</label><select id="rCur">${[...new Set([...CURRENCIES, r.currency || 'EUR'])].map((c) => `<option ${c === (r.currency || 'EUR') ? 'selected' : ''}>${c}</option>`).join('')}</select></div></div></div>
        <div class="field" id="fxRow" ${!r.currency || r.currency === 'EUR' ? 'hidden' : ''}><div class="cols"><div><label>Kurs: 1 EUR =</label><input type="text" id="rRate" inputmode="decimal" value="${numIn(r.rate)}" placeholder="z. B. 0,94"></div>
          <div><label>Betrag in EUR</label><input type="text" id="rEur" inputmode="decimal" value="${numIn(r.amountEur)}" placeholder="0,00"></div></div>
          <div class="hint">Kurs eingeben oder den Euro-Betrag direkt aus der Kreditkartenabrechnung übernehmen.</div></div>
        <div class="field" id="literRow" ${cat && cat.liter ? '' : 'hidden'}><label>Liter</label><input type="text" id="rLiter" inputmode="decimal" value="${numIn(r.liters)}"></div>
        <div class="field" id="bewRow" ${r.category === 'bewirtung' ? '' : 'hidden'}><label>Anlass der Bewirtung</label><input type="text" id="rOcc" value="${h(r.occasion || '')}" placeholder="z. B. Projektbesprechung Klinikum">
          <label style="margin-top:8px">Teilnehmer (Name, Firma)</label><textarea id="rPart" rows="2" placeholder="z. B. Dr. Müller (Uniklinik), M. Schmidt (VYGON)">${h(r.participants || '')}</textarea>
          <div class="hint">Pflichtangaben laut Reisekostenordnung – werden auf der Belegseite im PDF ausgewiesen.</div></div>
      </div>
      <div class="btns"><button class="btn accent block" data-act="r-save" data-id="${r.id}">✓ Geprüft – speichern</button></div>
      <div class="btns" style="margin-top:10px"><button class="btn sm" data-act="r-addpage" data-id="${r.id}">＋ Seite fotografieren</button><button class="btn sm" data-act="r-rerun" data-id="${r.id}">${r.engine === 'manuell' ? 'Text erkennen' : 'Erneut erkennen'}</button><button class="btn sm danger" data-act="r-del" data-id="${r.id}">Löschen</button></div>
      ${r.rawText ? `<details style="margin-top:14px"><summary class="muted small">Erkannter Text</summary><pre class="small" style="white-space:pre-wrap">${h(r.rawText)}</pre></details>` : ''}
    </div>`;
    const sh = openSheet('Beleg', body, { left: 'Schließen', right: 'Sichern', onRight: () => saveReceiptForm(r.id, false) });
    const cur = $('#rCur', sh), catSel = $('#rCat', sh);
    const recalcEur = () => {
      const a = parseNum($('#rAmount', sh).value), rate = parseNum($('#rRate', sh).value);
      if ($('#rCur', sh).value === 'EUR') return;
      if (a != null && rate) $('#rEur', sh).value = numIn(a / rate);
    };
    cur.onchange = () => { $('#fxRow', sh).hidden = cur.value === 'EUR'; recalcEur(); };
    $('#rRate', sh).oninput = recalcEur; $('#rAmount', sh).oninput = recalcEur;
    catSel.onchange = () => { $('#literRow', sh).hidden = !C[catSel.value].liter; $('#bewRow', sh).hidden = catSel.value !== 'bewirtung'; };
  }

  async function saveReceiptForm(id, confirmIt) {
    const sh = $('#sheet'); const r = await DB.get('receipts', id); if (!r) return;
    setReceiptDate(r, $('#rDate', sh).value);
    r.category = $('#rCat', sh).value;
    r.merchant = $('#rMerchant', sh).value.trim();
    r.text = $('#rText', sh).value.trim();
    r.amount = parseNum($('#rAmount', sh).value);
    r.currency = $('#rCur', sh).value;
    if (r.currency === 'EUR') { r.amountEur = r.amount; r.rate = null; } else { r.rate = parseNum($('#rRate', sh).value); r.amountEur = parseNum($('#rEur', sh).value); }
    r.liters = C[r.category].liter ? parseNum($('#rLiter', sh).value) : null;
    if (r.category === 'bewirtung') { r.occasion = $('#rOcc', sh).value.trim(); r.participants = $('#rPart', sh).value.trim(); }
    if (confirmIt) {
      const miss = [];
      if (!r.date) miss.push('Datum'); if (r.amountEur == null) miss.push('Betrag in EUR');
      if (miss.length) { toast('Bitte ergänzen: ' + miss.join(', '), 3500); return; }
      r.status = 'ok';
    } else if (r.status === 'error' && r.date && r.amountEur != null) r.status = 'review';
    const moved = r.monthKey !== S.monthKey;
    await saveReceipt(r);
    closeSheet(); render();
    toast(moved ? `Gespeichert – Beleg liegt im ${monthName(r.monthKey)}` : 'Beleg gespeichert');
  }

  async function deleteReceipt(id) {
    if (!confirm('Beleg wirklich löschen?')) return;
    const r = await DB.get('receipts', id);
    if (r) for (const iid of r.images) { await DB.del('images', iid); const u = S.imgUrls.get(iid); if (u) { URL.revokeObjectURL(u); S.imgUrls.delete(iid); } }
    await DB.del('receipts', id);
    S.receipts = S.receipts.filter((x) => x.id !== id);
    closeSheet(); render(); toast('Beleg gelöscht');
  }

  async function zoomReceipt(id) {
    const r = await DB.get('receipts', id); if (!r || !r.images.length) return;
    const z = document.createElement('div'); z.className = 'zoom';
    const urls = []; for (const iid of r.images) urls.push(await imgUrl(iid));
    z.innerHTML = urls.map((u) => `<img src="${u}" alt="">`).join('<div style="height:8px"></div>') + '<button class="btn x">Schließen</button>';
    z.onclick = (e) => { if (e.target.closest('.x')) z.remove(); };
    document.body.appendChild(z);
  }

  async function markBreakfast(id) {
    const r = await DB.get('receipts', id); if (!r || !r.hotel) return;
    const start = new Date(r.hotel.checkin + 'T12:00:00'); const end = new Date(r.hotel.checkout + 'T12:00:00');
    const touched = {};
    for (let d = new Date(start.getTime() + 864e5); d <= end; d = new Date(d.getTime() + 864e5)) {
      const key = keyOf(d.getFullYear(), d.getMonth() + 1);
      const m = key === S.monthKey ? S.month : (touched[key] || (touched[key] = (await DB.get('months', key)) || { days: {}, vehicle: {} }));
      const dd = m.days[d.getDate()] || (m.days[d.getDate()] = {});
      dd.breakfast = true;
      if (key !== S.monthKey) touched[key] = m;
    }
    await saveMonth(); for (const [k, m] of Object.entries(touched)) await DB.put('months', m, k);
    toast('Frühstückskürzung gesetzt'); closeSheet(); render();
  }

  // ---------- Tag-Editor ----------
  function openDay(d) {
    const dd = Object.assign({}, S.month.days[d] || {});
    const has = dd.startH !== '' && dd.startH != null;
    const prevDd = S.month.days[d - 1];
    const countryOpts = S.countries.map((c) => `<option value="${c.idx}" ${Number(dd.country || 1) === c.idx ? 'selected' : ''}>${h(c.name)}</option>`).join('');
    const body = `
      <div class="pd-box"><div><div class="l">Verpflegungspauschale</div><div class="l" id="pdInfo"></div></div><div class="v" id="pdVal">–</div></div>
      <div class="card form">
        <div class="field"><div class="cols">
          <div><label>Beginn</label><input type="time" id="dStart" value="${has ? tstr(dd.startH, dd.startM) : ''}"></div>
          <div><label>Ende</label><input type="time" id="dEnd" value="${has && Number(dd.endH) < 24 ? tstr(dd.endH, dd.endM) : ''}" ${has && Number(dd.endH) === 24 ? 'disabled' : ''}></div></div></div>
        <div class="switch-row"><div>Ab 0:00 Uhr<div class="sub">Reise läuft vom Vortag weiter</div></div><label class="switch"><input type="checkbox" id="dFrom0" ${has && Number(dd.startH) === 0 && Number(dd.startM || 0) === 0 ? 'checked' : ''}><span></span></label></div>
        <div class="switch-row"><div>Bis 24:00 Uhr<div class="sub">Reise geht am Folgetag weiter</div></div><label class="switch"><input type="checkbox" id="dTo24" ${has && Number(dd.endH) === 24 ? 'checked' : ''}><span></span></label></div>
        <div class="switch-row"><div>An- / Abreisetag<div class="sub">einer mehrtägigen Reise mit Übernachtung</div></div><label class="switch"><input type="checkbox" id="dArr" ${dd.arrival ? 'checked' : ''}><span></span></label></div>
      </div>
      <div class="card form">
        <div class="field"><label>Reiseziel (Ort, Firma, Zweck der Reise)</label><input type="text" id="dDest" value="${h(dd.destination || '')}" placeholder="z. B. Köln, Uniklinik, Produktschulung"></div>
        <div class="field"><label>Land</label><select id="dCountry">${countryOpts}</select></div>
      </div>
      <h3>Gestellte Mahlzeiten (Kürzung)</h3>
      <div class="card form">
        <div class="switch-row"><div>Frühstück<div class="sub" id="cutB"></div></div><label class="switch"><input type="checkbox" id="dB" ${dd.breakfast ? 'checked' : ''}><span></span></label></div>
        <div class="switch-row"><div>Mittagessen<div class="sub" id="cutL"></div></div><label class="switch"><input type="checkbox" id="dL" ${dd.lunch ? 'checked' : ''}><span></span></label></div>
        <div class="switch-row"><div>Abendessen<div class="sub" id="cutD"></div></div><label class="switch"><input type="checkbox" id="dD" ${dd.dinner ? 'checked' : ''}><span></span></label></div>
      </div>
      <div class="card form">
        <div class="field"><label>Übernachtungspauschale (€) – nur ohne Hotelrechnung</label><input type="text" id="dHotel" inputmode="decimal" value="${numIn(dd.hotelFlat)}" placeholder="0,00"><div class="hint" id="hotelHint"></div></div>
      </div>
      <div class="btns">${prevDd && prevDd.startH != null && prevDd.startH !== '' ? '<button class="btn sm" data-dact="copy">Wie Vortag</button>' : ''}<button class="btn sm danger" data-dact="clear">Tag leeren</button></div>`;
    const sh = openSheet(`${WD[wd(S.monthKey, d)]}, ${pad(d)}.${pad(ym(S.monthKey).m)}.${ym(S.monthKey).y}`, body, { left: 'Abbrechen', right: 'Fertig', onRight: () => save() });
    const read = () => {
      const st = $('#dStart', sh).value, en = $('#dEnd', sh).value;
      const from0 = $('#dFrom0', sh).checked, to24 = $('#dTo24', sh).checked;
      const o = {
        destination: $('#dDest', sh).value.trim(), country: Number($('#dCountry', sh).value), arrival: $('#dArr', sh).checked,
        breakfast: $('#dB', sh).checked, lunch: $('#dL', sh).checked, dinner: $('#dD', sh).checked, hotelFlat: parseNum($('#dHotel', sh).value),
      };
      if (from0 || st) { const [H, M] = from0 ? [0, 0] : st.split(':').map(Number); o.startH = H; o.startM = M; } else { o.startH = ''; o.startM = ''; }
      if (to24) { o.endH = 24; o.endM = 0; } else if (en) { const [H, M] = en.split(':').map(Number); o.endH = H; o.endM = M; } else { o.endH = ''; o.endM = ''; }
      if (o.startH !== '' && o.endH === '') { o.endH = o.startH; o.endM = o.startM; }
      return o;
    };
    const upd = () => {
      $('#dStart', sh).disabled = $('#dFrom0', sh).checked; $('#dEnd', sh).disabled = $('#dTo24', sh).checked;
      const o = read(); const pd = RB003.perDiem(o, S.countries);
      const c = S.countries.find((x) => x.idx === o.country) || S.countries[0];
      $('#pdVal', sh).textContent = pd.hours == null ? '–' : eur(typeof pd.display === 'number' ? pd.display : 0);
      $('#pdInfo', sh).textContent = pd.hours == null ? 'Beginn und Ende eintragen' : `${num(pd.hours, 2)} Std. · ${c.name} (24 h: ${eur(c.full)} / Teil: ${eur(c.partial)})${pd.raw < 0 ? ' · Kürzung > Pauschale' : ''}`;
      $('#cutB', sh).textContent = `− ${eur(c.full * 0.2)} (20 %)`; $('#cutL', sh).textContent = `− ${eur(c.full * 0.4)} (40 %)`; $('#cutD', sh).textContent = `− ${eur(c.full * 0.4)} (40 %)`;
      $('#hotelHint', sh).textContent = c.idx === 1 ? 'Inland: steuerfreie Pauschale 20 € je Übernachtung (nur in genehmigten Fällen).' : `Pauschale ${c.name}: ${c.hotel ? eur(c.hotel) : 'k. A.'} je Übernachtung`;
    };
    sh.addEventListener('input', upd); sh.addEventListener('change', upd); upd();
    const save = async () => {
      const o = read();
      if (o.startH !== '' && Number(o.endH) * 60 + Number(o.endM) < Number(o.startH) * 60 + Number(o.startM)) { toast('Ende liegt vor dem Beginn. Mehrtägig? → „Bis 24:00" aktivieren.', 4000); return; }
      if (o.startH === '' && !o.destination && !o.hotelFlat && !o.breakfast) delete S.month.days[d]; else S.month.days[d] = o;
      await saveMonth(); closeSheet(); render();
    };
    sh.querySelectorAll('[data-dact]').forEach((b) => (b.onclick = async () => {
      if (b.dataset.dact === 'clear') { delete S.month.days[d]; await saveMonth(); closeSheet(); render(); toast('Tag geleert'); }
      if (b.dataset.dact === 'copy') { const p = S.month.days[d - 1]; S.month.days[d] = Object.assign({}, p, { hotelFlat: p.hotelFlat }); await saveMonth(); closeSheet(); openDay(d); }
    }));
  }

  // ---------- Reise anlegen (mehrtägig) ----------
  function openTrip() {
    const { y, m } = ym(S.monthKey);
    const t = new Date(); const def = t.getFullYear() === y && t.getMonth() + 1 === m ? `${y}-${pad(m)}-${pad(t.getDate())}` : `${y}-${pad(m)}-01`;
    const countryOpts = S.countries.map((c) => `<option value="${c.idx}">${h(c.name)}</option>`).join('');
    const body = `<div class="card form">
      <div class="field"><div class="cols"><div><label>Abreise am</label><input type="date" id="tFromD" value="${def}"></div><div><label>um</label><input type="time" id="tFromT" value="07:00"></div></div></div>
      <div class="field"><div class="cols"><div><label>Rückkehr am</label><input type="date" id="tToD" value="${def}"></div><div><label>um</label><input type="time" id="tToT" value="18:00"></div></div></div>
      <div class="field"><label>Reiseziel (Ort, Firma, Zweck)</label><input type="text" id="tDest" placeholder="z. B. München, Klinikum rechts der Isar, Schulung"></div>
      <div class="field"><label>Land</label><select id="tCountry">${countryOpts}</select></div>
      <div class="switch-row"><div>Frühstück im Hotel enthalten<div class="sub">Kürzung um 20 % an den Tagen nach jeder Übernachtung</div></div><label class="switch"><input type="checkbox" id="tBreak"><span></span></label></div>
    </div><div class="notice info"><div>Die App legt alle Tage automatisch an: Abreisetag (ab Uhrzeit bis 24:00, An-/Abreisetag), volle Zwischentage (0–24 Uhr) und Rückreisetag. Einzelne Tage kannst du danach noch anpassen.</div></div>
    <div id="tPreview" class="muted small" style="margin:0 4px 10px"></div>
    <button class="btn primary block" id="tSave">Reise übernehmen</button>`;
    const sh = openSheet('Reise anlegen', body, { left: 'Abbrechen' });
    const read = () => ({ fromD: $('#tFromD', sh).value, fromT: $('#tFromT', sh).value, toD: $('#tToD', sh).value, toT: $('#tToT', sh).value, dest: $('#tDest', sh).value.trim(), country: Number($('#tCountry', sh).value), br: $('#tBreak', sh).checked });
    const plan = () => {
      const o = read(); if (!o.fromD || !o.toD || !o.fromT || !o.toT) return null;
      const a = new Date(o.fromD + 'T12:00:00'), b = new Date(o.toD + 'T12:00:00');
      const n = Math.round((b - a) / 864e5); if (n < 0 || n > 62) return null;
      const [fh, fm] = o.fromT.split(':').map(Number), [th, tm] = o.toT.split(':').map(Number);
      if (n === 0 && th * 60 + tm <= fh * 60 + fm) return null;
      const days = [];
      for (let i = 0; i <= n; i++) {
        const d = new Date(a.getTime() + i * 864e5);
        const e = { destination: o.dest, country: o.country, arrival: n > 0 && (i === 0 || i === n), breakfast: o.br && i > 0, lunch: false, dinner: false };
        if (n === 0) Object.assign(e, { startH: fh, startM: fm, endH: th, endM: tm });
        else if (i === 0) Object.assign(e, { startH: fh, startM: fm, endH: 24, endM: 0 });
        else if (i === n) Object.assign(e, { startH: 0, startM: 0, endH: th, endM: tm });
        else Object.assign(e, { startH: 0, startM: 0, endH: 24, endM: 0 });
        days.push({ key: keyOf(d.getFullYear(), d.getMonth() + 1), day: d.getDate(), e });
      }
      return days;
    };
    const prev = () => {
      const p = plan(); const el = $('#tPreview', sh);
      if (!p) { el.textContent = 'Bitte gültigen Zeitraum angeben.'; return; }
      const sum = p.reduce((s, x) => s + (RB003.perDiem(x.e, S.countries).net || 0), 0);
      el.textContent = `${p.length} Tag${p.length > 1 ? 'e' : ''} · Verpflegungspauschalen gesamt ${eur(sum)}`;
    };
    sh.addEventListener('input', prev); sh.addEventListener('change', prev); prev();
    $('#tSave', sh).onclick = async () => {
      const p = plan(); if (!p) { toast('Bitte gültigen Zeitraum angeben'); return; }
      const months = {};
      for (const x of p) {
        const m = x.key === S.monthKey ? S.month : (months[x.key] || (months[x.key] = (await DB.get('months', x.key)) || { days: {}, vehicle: {} }));
        const existing = m.days[x.day] || {};
        m.days[x.day] = Object.assign({}, x.e, { hotelFlat: existing.hotelFlat || null, lunch: existing.lunch || false, dinner: existing.dinner || false });
      }
      await saveMonth(); for (const [k, m] of Object.entries(months)) await DB.put('months', m, k);
      closeSheet(); S.view = 'days'; setView('days'); toast('Reise angelegt');
    };
  }

  // ---------- Fahrzeug-Editor ----------
  function lastKmEnd(beforeDay) {
    for (let d = beforeDay - 1; d >= 1; d--) { const v = S.month.vehicle[d]; if (v && v.kmEnd !== '' && v.kmEnd != null) return v.kmEnd; }
    return S.month.lastKm != null ? S.month.lastKm : null;
  }
  function openVehicle(d) {
    const v = Object.assign({}, S.month.vehicle[d] || {});
    const last = lastKmEnd(d);
    const recs = S.receipts.filter((r) => r.day === d && C[r.category] && C[r.category].sheet === 2);
    const body = `
      <h3 style="margin-top:4px">Dienstwagen (Spalte 11)</h3>
      <div class="card form">
        <div class="field"><div class="cols"><div><label>km-Stand Anfang</label><input type="text" inputmode="numeric" id="vS" value="${v.kmStart ?? ''}" placeholder="${last != null ? last : ''}"></div>
          <div><label>km-Stand Ende</label><input type="text" inputmode="numeric" id="vE" value="${v.kmEnd ?? ''}"></div></div>
          ${last != null && (v.kmStart === '' || v.kmStart == null) ? `<div class="hint">Letzter km-Stand: ${last.toLocaleString('de-DE')} – <a href="#" id="vTake">übernehmen</a></div>` : ''}</div>
        <div class="field"><label>davon privat gefahren (km)</label><input type="text" inputmode="numeric" id="vP" value="${v.kmPrivate ?? ''}" placeholder="0"><div class="hint" id="vInfo"></div></div>
      </div>
      <h3>Mietwagen (Spalte 12)</h3>
      <div class="card form"><div class="field"><div class="cols"><div><label>km geschäftlich</label><input type="text" inputmode="numeric" id="vMB" value="${v.mwKmBusiness ?? ''}"></div><div><label>km privat</label><input type="text" inputmode="numeric" id="vMP" value="${v.mwKmPrivate ?? ''}"></div></div></div></div>
      <h3>Privatwagen (Spalte 13)</h3>
      <div class="card form"><div class="field"><label>geschäftlich gefahrene km</label><input type="text" inputmode="numeric" id="vPK" value="${v.privCarKm ?? ''}"><div class="hint" id="vPInfo"></div></div></div>
      <div class="card form"><div class="field"><label>Notiz (Spalte „TEXT" auf Seite 2)</label><input type="text" id="vN" value="${h(v.note || '')}" placeholder="z. B. Fahrt Aachen–Köln–Aachen"></div></div>
      ${recs.length ? `<h3>Belege an diesem Tag</h3><div class="card list">${numbered().filter((r) => recs.some((x) => x.id === r.id)).map(receiptItem).join('')}</div>` : ''}
      <div class="btns"><button class="btn sm danger" id="vClear">Fahrzeugdaten leeren</button></div>`;
    const sh = openSheet(`Fahrzeug · ${WD[wd(S.monthKey, d)]}, ${pad(d)}.${pad(ym(S.monthKey).m)}.`, body, { left: 'Abbrechen', right: 'Fertig', onRight: () => save() });
    const intOrEmpty = (id) => { const x = $(id, sh).value.replace(/[^\d]/g, ''); return x === '' ? '' : Number(x); };
    const read = () => ({ kmStart: intOrEmpty('#vS'), kmEnd: intOrEmpty('#vE'), kmPrivate: intOrEmpty('#vP'), mwKmBusiness: intOrEmpty('#vMB'), mwKmPrivate: intOrEmpty('#vMP'), privCarKm: intOrEmpty('#vPK'), note: $('#vN', sh).value.trim() });
    const upd = () => {
      const o = read();
      $('#vInfo', sh).textContent = o.kmStart !== '' && o.kmEnd !== '' ? `Gefahren: ${o.kmEnd - o.kmStart} km · davon geschäftlich ${o.kmEnd - o.kmStart - (o.kmPrivate || 0)} km` : '';
      $('#vPInfo', sh).textContent = o.privCarKm ? `${o.privCarKm} km × ${num(S.settings.kmRate)} € = ${eur(o.privCarKm * S.settings.kmRate)}` : `km-Pauschale ${num(S.settings.kmRate)} €/km`;
    };
    sh.addEventListener('input', upd); upd();
    const take = $('#vTake', sh); if (take) take.onclick = (e) => { e.preventDefault(); $('#vS', sh).value = last; upd(); };
    const save = async () => {
      const o = read();
      if (o.kmStart !== '' && o.kmEnd !== '' && o.kmEnd < o.kmStart) { toast('km-Stand Ende ist kleiner als Anfang'); return; }
      if (Object.values(o).every((x) => x === '' || x == null)) delete S.month.vehicle[d]; else S.month.vehicle[d] = o;
      await saveMonth(); closeSheet(); render();
    };
    $('#vClear', sh).onclick = async () => { delete S.month.vehicle[d]; await saveMonth(); closeSheet(); render(); };
  }
  function openVehicleAdd() {
    const n = daysIn(S.monthKey);
    let opts = ''; for (let d = 1; d <= n; d++) opts += `<option value="${d}">${WD[wd(S.monthKey, d)]}, ${pad(d)}.${pad(ym(S.monthKey).m)}.</option>`;
    const sh = openSheet('Tag wählen', `<div class="card form"><div class="field"><label>Tag</label><select id="vaDay">${opts}</select></div></div><button class="btn primary block" id="vaGo">Weiter</button>`);
    $('#vaGo', sh).onclick = () => { const d = Number($('#vaDay', sh).value); closeSheet(); openVehicle(d); };
  }

  // ---------- Export ----------
  function openExport() {
    const a = agg();
    const list = numbered();
    const issues = warnings(a).filter(([t]) => t !== 'info');
    const proc = list.filter((r) => r.status === 'processing').length;
    const body = `
      <div class="card pad"><div class="row between"><div><b>${monthName(S.monthKey)}</b><div class="muted small">${h(S.settings.name || 'Name fehlt')} · Kst. ${h(S.settings.costCenter || '–')}</div></div><div class="right"><div class="muted small">Gesamt</div><b class="num" style="font-size:20px">${eur(a.total)}</b></div></div></div>
      ${issues.map(([t, m]) => `<div class="notice ${t}"><div>${m}</div></div>`).join('')}
      ${proc ? '<div class="notice info"><div>Bitte warten, bis alle Belege erkannt sind.</div></div>' : ''}
      <div class="card list">
        <div class="item" style="cursor:default"><div class="grow"><div class="t">Excel (RB003, Original-Vorlage)</div><div class="s">Blatt 1 + Seite 2 befüllt, alle Formeln bleiben erhalten</div></div></div>
        <div class="item" style="cursor:default"><div class="grow"><div class="t">PDF mit Einzelnachweisen</div><div class="s">RB003 im Excel-Layout (Blatt 1 + Seite 2), Belegliste und ${list.filter((r) => r.day).length} nummerierte Belege</div></div></div>
      </div>
      <div id="exState"></div>
      <button class="btn primary block" id="exGo" ${proc ? 'disabled' : ''}>Export erstellen</button>`;
    const sh = openSheet('Monat exportieren', body);
    $('#exGo', sh).onclick = async () => {
      const btn = $('#exGo', sh); btn.disabled = true;
      const st = $('#exState', sh);
      try {
        const { y, m } = ym(S.monthKey);
        const data = exportData();
        const sorted = numbered().filter((r) => r.day && r.status !== 'processing' && r.status !== 'error');
        const ctx = { templateBuf: S.templateBuf, countries: S.countries, data, month: m, year: y, name: S.settings.name, costCenter: S.settings.costCenter, company: S.tplInfo.company, version: S.tplInfo.version };
        const res = await E.exportMonth(ctx, sorted.map((r) => Object.assign({}, r, { text: r.text, merchant: r.merchant })), (id) => DB.get('images', id), (msg) => { st.innerHTML = `<div class="notice info"><div><span class="spin"></span> ${h(msg)}</div></div>`; });
        S.lastExport = res;
        const files = [new File([res.xlsx.blob], res.xlsx.filename, { type: res.xlsx.blob.type }), new File([res.pdf.blob], res.pdf.filename, { type: 'application/pdf' })];
        const canShare = navigator.canShare && navigator.canShare({ files });
        st.innerHTML = `<div class="notice ok"><div><b>Export fertig.</b><br>${h(res.xlsx.filename)}<br>${h(res.pdf.filename)} (${Math.round(res.pdf.blob.size / 1024)} KB)</div></div>
          <div class="btns" style="margin-bottom:10px">${canShare ? '<button class="btn accent" id="exShare">Teilen (Mail, Teams, OneDrive …)</button>' : ''}</div>
          <div class="btns"><button class="btn" id="exX">Excel öffnen</button><button class="btn" id="exP">PDF öffnen</button></div>`;
        btn.classList.add('hidden');
        if (canShare) $('#exShare', sh).onclick = async () => { try { await navigator.share({ files, title: `Reisekostenabrechnung ${monthName(S.monthKey)}` }); } catch (e) { if (e.name !== 'AbortError') toast('Teilen nicht möglich: ' + e.message); } };
        $('#exX', sh).onclick = () => download(res.xlsx.blob, res.xlsx.filename);
        $('#exP', sh).onclick = () => download(res.pdf.blob, res.pdf.filename);
      } catch (e) {
        console.error(e); st.innerHTML = `<div class="notice err"><div>Export fehlgeschlagen: ${h(e.message)}</div></div>`; btn.disabled = false;
      }
    };
  }
  function download(blob, name) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 60000);
  }

  // ---------- Einstellungen: Aktionen ----------
  function bindSettings() {
    const v = $('#view');
    const bind = (id, key, conv = (x) => x) => { const el = $(id, v); if (el) el.onchange = async () => { S.settings[key] = conv(el.value); await saveSettings(); render(); }; };
    bind('#sName', 'name', (x) => x.trim()); bind('#sCost', 'costCenter', (x) => x.trim());
    bind('#sKm', 'kmRate', (x) => parseNum(x) || 0.30); bind('#sKey', 'apiKey', (x) => x.trim()); bind('#sModel', 'aiModel');
    v.querySelectorAll('[data-ocr]').forEach((b) => (b.onclick = async () => { S.settings.ocrMode = b.dataset.ocr; await saveSettings(); render(); }));
    if (navigator.storage && navigator.storage.estimate) navigator.storage.estimate().then((e) => { const el = $('#storeInfo'); if (el) el.textContent = `Alle Daten liegen nur auf diesem Gerät · belegt ca. ${(e.usage / 1048576).toFixed(1)} MB`; });
  }

  async function testAI() {
    const st = $('#ocrStatus');
    if (!S.settings.apiKey) { toast('Bitte zuerst den API-Key eintragen'); return; }
    st.innerHTML = '<span class="spin"></span> Teste Verbindung …';
    try {
      const c = document.createElement('canvas'); c.width = 400; c.height = 220; const x = c.getContext('2d');
      x.fillStyle = '#fff'; x.fillRect(0, 0, 400, 220); x.fillStyle = '#000'; x.font = '20px monospace';
      ['PARKHAUS DOM KÖLN', 'Datum: 12.03.2026', 'Parkgebühr', 'inkl. 19% MwSt 1,04', 'SUMME EUR   6,50'].forEach((t, i) => x.fillText(t, 20, 36 + i * 38));
      const blob = await new Promise((r) => c.toBlob(r, 'image/jpeg', 0.9));
      const res = await OCR.recognizeAI([blob], S.settings);
      st.innerHTML = `✅ KI-Erkennung funktioniert: ${h(res.merchant || '')} · ${res.total_gross != null ? eur(res.total_gross) : '?'} · ${fmtDate(res.date)} · ${h(C[res.category].short)}`;
    } catch (e) { st.innerHTML = `❌ ${h(e.message)}`; }
  }
  async function warmup() {
    const st = $('#ocrStatus'); st.innerHTML = '<span class="spin"></span> Lade Offline-Erkennung …';
    try { await OCR.warmupOffline((m) => { if (m.status) st.innerHTML = `<span class="spin"></span> ${h(m.status)} ${Math.round(m.progress * 100)} %`; }); st.innerHTML = '✅ Offline-Erkennung ist bereit und funktioniert auch ohne Netz.'; } catch (e) { st.innerHTML = `❌ ${h(e.message)}`; }
  }

  async function loadNewTemplate(file) {
    try {
      const buf = await file.arrayBuffer();
      const info = await RB003.readTemplateInfo(window.JSZip, buf);
      if (info.countries.length < 50) throw new Error('Das sieht nicht wie eine RB003-Vorlage aus (Ländertabelle fehlt).');
      await DB.put('settings', buf, 'template'); await loadTemplate(); render();
      toast(`Vorlage geladen: ${info.version}`);
    } catch (e) { toast('Vorlage ungültig: ' + e.message, 4500); }
  }

  async function backup() {
    toast('Backup wird erstellt …');
    const zip = new JSZip();
    const settings = Object.assign({}, S.settings, { apiKey: '' });
    const months = {}; for (const k of await DB.keys('months')) months[k] = await DB.get('months', k);
    const receipts = await DB.all('receipts');
    zip.file('daten.json', JSON.stringify({ app: 'reisekosten-rb003', version: APP_VERSION, created: new Date().toISOString(), settings, months, receipts }, null, 1));
    for (const id of await DB.keys('images')) zip.file('bilder/' + id + '.jpg', await DB.get('images', id));
    const blob = await zip.generateAsync({ type: 'blob' });
    const name = `Reisekosten_Backup_${new Date().toISOString().slice(0, 10)}.zip`;
    const f = new File([blob], name, { type: 'application/zip' });
    if (navigator.canShare && navigator.canShare({ files: [f] })) { try { await navigator.share({ files: [f], title: name }); return; } catch (e) { if (e.name === 'AbortError') return; } }
    download(blob, name);
  }
  async function restore(file) {
    if (!confirm('Backup einspielen? Vorhandene Daten auf diesem Gerät werden ersetzt.')) return;
    try {
      const zip = await JSZip.loadAsync(await file.arrayBuffer());
      const j = JSON.parse(await zip.file('daten.json').async('string'));
      if (j.app !== 'reisekosten-rb003') throw new Error('Keine Reisekosten-Sicherung');
      const key = S.settings.apiKey; const tpl = await DB.get('settings', 'template');
      await DB.clearAll();
      await DB.put('settings', Object.assign({}, j.settings, { apiKey: key }), 'profile'); if (tpl) await DB.put('settings', tpl, 'template');
      for (const [k, m] of Object.entries(j.months || {})) await DB.put('months', m, k);
      for (const r of j.receipts || []) await DB.put('receipts', r);
      for (const name of Object.keys(zip.files).filter((n) => n.startsWith('bilder/') && !zip.files[n].dir)) {
        const id = name.slice(7).replace(/\.jpg$/, ''); await DB.put('images', new Blob([await zip.file(name).async('arraybuffer')], { type: 'image/jpeg' }), id);
      }
      S.imgUrls.clear(); await loadSettings(); await loadMonth(S.monthKey); render(); toast('Backup eingespielt');
    } catch (e) { toast('Backup fehlgeschlagen: ' + e.message, 4500); }
  }
  async function wipe() {
    if (!confirm('Wirklich ALLE Belege, Reisetage und Einstellungen auf diesem Gerät löschen?')) return;
    if (!confirm('Letzte Sicherheitsabfrage: Löschen kann nicht rückgängig gemacht werden. Fortfahren?')) return;
    await DB.clearAll(); S.imgUrls.clear(); await loadSettings(); await loadTemplate(); await loadMonth(todayKey()); render(); toast('Alle Daten gelöscht');
  }

  // ---------- Events ----------
  function bindGlobal() {
    $('#mPrev').onclick = async () => { await loadMonth(shiftMonth(S.monthKey, -1)); render(); };
    $('#mNext').onclick = async () => { await loadMonth(shiftMonth(S.monthKey, 1)); render(); };
    $('#tabs').onclick = (e) => { const b = e.target.closest('button[data-view]'); if (b) setView(b.dataset.view); };
    $('#fab').onclick = () => $('#fileCamera').click();
    $('#fileCamera').onchange = (e) => { const f = [...e.target.files]; e.target.value = ''; if (S.addPageTo) { addPage(S.addPageTo, f); S.addPageTo = null; } else handleFiles(f); };
    $('#fileAny').onchange = (e) => { const f = [...e.target.files]; e.target.value = ''; handleFiles(f); };
    $('#fileTemplate').onchange = (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) loadNewTemplate(f); };
    $('#fileBackup').onchange = (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) restore(f); };
    $('#sheetWrap').addEventListener('click', (e) => { if (e.target.id === 'sheetWrap') closeSheet(); });

    document.addEventListener('click', (e) => {
      const go = e.target.closest('[data-go]');
      if (go) { const [v, f] = go.dataset.go.split(':'); if (f) S.filter = f; setView(v); return; }
      const fl = e.target.closest('[data-filter]'); if (fl) { S.filter = fl.dataset.filter; render(); return; }
      const rec = e.target.closest('[data-receipt]'); if (rec) { openReceipt(rec.dataset.receipt); return; }
      const day = e.target.closest('[data-day]'); if (day) { openDay(Number(day.dataset.day)); return; }
      const veh = e.target.closest('[data-vehicle]'); if (veh) { openVehicle(Number(veh.dataset.vehicle)); return; }
      const act = e.target.closest('[data-act]'); if (!act) return;
      const id = act.dataset.id;
      switch (act.dataset.act) {
        case 'camera': $('#fileCamera').click(); break;
        case 'files': $('#fileAny').click(); break;
        case 'trip': openTrip(); break;
        case 'export': openExport(); break;
        case 'vehicle-add': openVehicleAdd(); break;
        case 'r-save': saveReceiptForm(id, true); break;
        case 'r-del': deleteReceipt(id); break;
        case 'r-rerun': rerun(id); break;
        case 'r-addpage': S.addPageTo = id; $('#fileCamera').click(); break;
        case 'zoom': zoomReceipt(id); break;
        case 'breakfast': markBreakfast(id); break;
        case 'test-ai': testAI(); break;
        case 'warmup': warmup(); break;
        case 'tpl-load': $('#fileTemplate').click(); break;
        case 'tpl-reset': DB.del('settings', 'template').then(loadTemplate).then(() => { render(); toast('Integrierte Vorlage aktiv'); }); break;
        case 'backup': backup(); break;
        case 'restore': $('#fileBackup').click(); break;
        case 'wipe': wipe(); break;
      }
    });
    const obs = new MutationObserver(() => { if (S.view === 'settings') bindSettings(); });
    obs.observe($('#view'), { childList: true });
  }

  async function rerun(id) {
    const r = await DB.get('receipts', id); if (!r) return;
    r.status = 'processing'; await saveReceipt(r); closeSheet(); render(); enqueueRecognition(id);
  }
  async function addPage(id, files) {
    const r = await DB.get('receipts', id); if (!r) return;
    for (const f of files) for (const im of await OCR.fileToImages(f)) { const iid = uid(); await DB.put('images', im.blob, iid); r.images.push(iid); }
    await saveReceipt(r); openReceipt(id); toast('Seite hinzugefügt');
  }

  // ---------- Start ----------
  const APP_VERSION = '1.3.0';
  async function init() {
    await loadSettings();
    await loadTemplate();
    await loadMonth(localStorage.getItem('rk_month') || todayKey());
    // hängengebliebene Erkennungen (App wurde während der Erkennung geschlossen) neu starten
    for (const r of (await DB.all('receipts')).filter((x) => x.status === 'processing')) enqueueRecognition(r.id);
    bindGlobal();
    render();
    if (!S.settings.name) { S.view = 'settings'; setView('settings'); toast('Willkommen! Bitte zuerst Name und Kostenstelle eintragen.', 4000); }
    if ('serviceWorker' in navigator && (location.protocol === 'https:' || ['localhost', '127.0.0.1'].includes(location.hostname))) navigator.serviceWorker.register('sw.js').catch((e) => console.warn('SW', e));
  }
  window.addEventListener('DOMContentLoaded', () => init().catch((e) => { document.body.innerHTML = `<div style="padding:30px;font-family:sans-serif"><h2>Start fehlgeschlagen</h2><pre>${h(e.stack || e.message)}</pre></div>`; }));
  window.__RK = S; // für Tests
})();
