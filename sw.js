// Офлайн-режим «Открывашки». Файл собирается make-sw.py из sw-template.js — правьте шаблон.
const VERSION = 'f87fde6a79';
const CACHE = 'otkryvashka-' + VERSION;
const ASSETS = [
"./",
"app.js",
"apple-touch-icon.png",
"doc.js",
"favicon-32.png",
"icon-192.png",
"icon-512.png",
"icon-maskable-512.png",
"index.html",
"manifest.webmanifest",
"style.css",
"vendor/docx-preview.min.js",
"vendor/jszip.min.js",
"vendor/xlsx.full.min.js"
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS.map(u => new Request(u, {cache: 'reload'})))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('otkryvashka-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// страница — из сети, если она есть (чтобы обновления приходили), иначе из кэша; скрипты и картинки — из кэша
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put('./', cp)); return r; })
      .catch(() => caches.match('./', {ignoreSearch: true})));
    return;
  }
  e.respondWith(caches.match(req, {ignoreSearch: true}).then(hit => hit || fetch(req).then(r => {
    if (r.ok) { const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); }
    return r;
  })));
});
