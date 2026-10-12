/* Кэш приложения для работы без интернета. Кэшируются только файлы самого инструмента;
   данные людей сюда не попадают — страница их никуда не передаёт. */
var CACHE = 'sb-dostup-v3';
var FILES = [
  './', './index.html', './app.css', './app.js', './core.js', './io.js', './gosb.js',
  './pdfgen.js', './docx-render.js', './fill.js', './vendor/fonts.js',
  './manifest.webmanifest', './icon-180.png', './icon-192.png', './icon-512.png',
  './vendor/jszip.min.js', './vendor/pdf.min.mjs', './vendor/pdf.worker.min.mjs'
];
self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(FILES); }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(function (hit) {
    return hit || fetch(e.request).catch(function () { return caches.match('./index.html'); });
  }));
});
