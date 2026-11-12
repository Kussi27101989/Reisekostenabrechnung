/* Service Worker – App funktioniert nach dem ersten Laden komplett offline */
const VERSION = 'rk-1.4.0';
const SHELL = [
  './', './index.html', './manifest.webmanifest', './css/app.css',
  './js/app.js', './js/rb003.js', './js/xlsxpdf.js', './js/db.js', './js/ocr.js', './js/export.js', './js/template.js',
  './vendor/jszip.min.js', './vendor/jspdf.umd.min.js',
  './icons/icon-180.png', './icons/icon-192.png', './icons/icon-512.png',
];
// Große Dateien (Offline-OCR, PDF-Import) werden beim ersten Gebrauch zwischengespeichert
const LAZY = /\/vendor\/(tesseract|pdf)/;

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION && k !== 'rk-lazy').map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return; // API-Aufrufe nie cachen
  if (LAZY.test(url.pathname)) {
    e.respondWith(caches.open('rk-lazy').then(async (c) => {
      const hit = await c.match(e.request); if (hit) return hit;
      const res = await fetch(e.request); if (res.ok) c.put(e.request, res.clone()); return res;
    }));
    return;
  }
  // App-Dateien: Cache zuerst, im Hintergrund aktualisieren
  e.respondWith(caches.open(VERSION).then(async (c) => {
    const hit = await c.match(e.request, { ignoreSearch: true });
    const net = fetch(e.request).then((res) => { if (res.ok) c.put(e.request, res.clone()); return res; }).catch(() => hit);
    return hit || net;
  }));
});
