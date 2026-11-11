/* Lokale Datenhaltung (IndexedDB) – alle Daten bleiben auf dem Gerät */
(function () {
  'use strict';
  const DB_NAME = 'reisekosten-rb003';
  const VERSION = 1;
  let dbp = null;

  function open() {
    if (dbp) return dbp;
    dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains('settings')) db.createObjectStore('settings');
        if (!db.objectStoreNames.contains('months')) db.createObjectStore('months');
        if (!db.objectStoreNames.contains('receipts')) {
          const s = db.createObjectStore('receipts', { keyPath: 'id' });
          s.createIndex('month', 'monthKey');
        }
        if (!db.objectStoreNames.contains('images')) db.createObjectStore('images');
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbp;
  }

  function tx(store, mode, fn) {
    return open().then((db) => new Promise((resolve, reject) => {
      const t = db.transaction(store, mode);
      const s = t.objectStore(store);
      let result;
      Promise.resolve(fn(s)).then((r) => { result = r; });
      t.oncomplete = () => resolve(result);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error || new Error('Transaktion abgebrochen'));
    }));
  }
  const reqP = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });

  const DB = {
    get: (store, key) => open().then((db) => reqP(db.transaction(store).objectStore(store).get(key))),
    put: (store, value, key) => tx(store, 'readwrite', (s) => { key === undefined ? s.put(value) : s.put(value, key); }),
    del: (store, key) => tx(store, 'readwrite', (s) => { s.delete(key); }),
    all: (store) => open().then((db) => reqP(db.transaction(store).objectStore(store).getAll())),
    keys: (store) => open().then((db) => reqP(db.transaction(store).objectStore(store).getAllKeys())),
    receiptsOfMonth: (monthKey) => open().then((db) => reqP(db.transaction('receipts').objectStore('receipts').index('month').getAll(monthKey))),
    clearAll: () => open().then((db) => Promise.all(['settings', 'months', 'receipts', 'images'].map((n) => reqP(db.transaction(n, 'readwrite').objectStore(n).clear())))),
  };
  window.DB = DB;
})();
