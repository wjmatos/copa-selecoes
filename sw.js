/* Service worker do painel: abre mesmo sem internet, mas sempre tenta buscar a versão nova primeiro. */
const V = 'copa-v1';
const CORE = ['./', 'index.html', 'dados.js', 'manifest.json', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (u.origin !== location.origin) return; // planilha do Google e fontes ficam por conta do navegador
  e.respondWith(
    fetch(r)
      .then(res => {
        if (res.ok) { const cp = res.clone(); caches.open(V).then(c => c.put(r, cp)); }
        return res;
      })
      .catch(() => caches.match(r).then(m => m || caches.match('index.html')))
  );
});
