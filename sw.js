// ═══════════════════════════════════════════════════════════
//  SERVICE WORKER — Buku Kas Keluarga
//  Strategi    : Cache-First (Stale-While-Revalidate)
//  Versi       : 2.0.0
//  Dibuat      : 2025
// ═══════════════════════════════════════════════════════════

// ──────────────────────────────────────────────
//  KONFIGURASI
// ──────────────────────────────────────────────

/** Nama unik cache — ganti versi saat deployment untuk memicu update */
const CACHE_NAME = 'buku-kas-v2.1.7';

/** Daftar aset statis yang langsung di-cache saat instalasi */
const PRECACHE_ASSETS = [
  // ── Aplikasi Inti ──────────────────────────
  './',
  './index.html',
  './manifest.json',

  // ── Ikon & Gambar ──────────────────────────
  './icon-72.png',
  './icon-96.png',
  './icon-128.png',
  './icon-144.png',
  './icon-152.png',
  './icon-192.png',
  './icon-384.png',
  './icon-512.png',
  './maskable-icon-192.png',
  './maskable-icon-512.png',

  // ── CDN Eksternal (Library) ────────────────
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js',
  'https://cdn.jsdelivr.net/npm/localforage@1.10.0/dist/localforage.min.js'
];

/** Pola URL yang TIDAK PERNAH di-cache (selalu fetch dari jaringan) */
const NETWORK_ONLY_PATTERNS = [
  /\/api\//,             // API calls (jika ada)
  /\/analytics\//,       // Tracking/analytics
  /chrome-extension:\/\//, // Chrome extension requests
];

// ──────────────────────────────────────────────
//  EVENT: INSTALL
//  Dipicu saat SW pertama kali diinstal
// ──────────────────────────────────────────────

self.addEventListener('install', (event) => {
  console.log('[SW] ⬇️  Install — versi:', CACHE_NAME);

  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log(`[SW] 📦 Pre-caching ${PRECACHE_ASSETS.length} aset...`);
        // Gunakan addAll dengan promise individu agar satu gagal tidak menggagalkan semua
        return Promise.allSettled(
          PRECACHE_ASSETS.map(url =>
            cache.add(url).catch(err => {
              console.warn(`[SW] ⚠️  Gagal cache: ${url}`, err.message);
            })
          )
        );
      })
      .then(() => {
        console.log('[SW] ✅ Pre-cache selesai — skip waiting');
        // Langsung aktifkan SW tanpa menunggu tab lama ditutup
        return self.skipWaiting();
      })
  );
});

// ──────────────────────────────────────────────
//  EVENT: ACTIVATE
//  Dipicu saat SW menjadi aktif (setelah install/update)
// ──────────────────────────────────────────────

self.addEventListener('activate', (event) => {
  console.log('[SW] 🟢 Activate — versi:', CACHE_NAME);

  event.waitUntil(
    caches.keys()
      .then((cacheNames) => {
        // Hapus semua cache versi lama yang tidak lagi digunakan
        const oldCaches = cacheNames.filter(name => name !== CACHE_NAME);
        if (oldCaches.length > 0) {
          console.log(`[SW] 🧹 Membersihkan ${oldCaches.length} cache lama:`, oldCaches);
        }
        return Promise.all(
          oldCaches.map(name => {
            console.log('[SW] 🗑️  Hapus cache:', name);
            return caches.delete(name);
          })
        );
      })
      .then(() => {
        console.log('[SW] ✅ Aktivasi selesai — claim clients');
        // Ambil alih semua tab yang terbuka tanpa perlu refresh
        return self.clients.claim();
      })
  );
});

// ──────────────────────────────────────────────
//  EVENT: FETCH
//  Strategi: Cache-First dengan Stale-While-Revalidate
// ──────────────────────────────────────────────

self.addEventListener('fetch', (event) => {
  const { request } = event;

  // ── Abaikan request non-GET ──────────────────
  if (request.method !== 'GET') return;

  // ── Abaikan skema non-http/https ──────────────
  const url = new URL(request.url);
  if (!url.protocol.startsWith('http')) return;

  // ── Network-only untuk pola tertentu ──────────
  const isNetworkOnly = NETWORK_ONLY_PATTERNS.some(pattern => pattern.test(request.url));
  if (isNetworkOnly) {
    // Jangan di-cache, langsung fetch dari network
    return;
  }

  // ── Cache-First Strategy ─────────────────────
  event.respondWith(
    caches.match(request)
      .then((cachedResponse) => {
        if (cachedResponse) {
          // ✅ DITEMUKAN DI CACHE — langsung kembalikan (instant load)
          // Di belakang layar: update cache dari network (stale-while-revalidate)
          fetch(request)
            .then((networkResponse) => {
              if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
                caches.open(CACHE_NAME)
                  .then(cache => cache.put(request, networkResponse.clone()))
                  .catch(() => { /* Gagal update cache, abaikan */ });
              }
            })
            .catch(() => {
              // Network tidak tersedia — tidak masalah, kita sudah punya cache
            });

          return cachedResponse;
        }

        // ❌ TIDAK DI CACHE — fetch dari network
        return fetch(request)
          .then((networkResponse) => {
            // Jangan cache respons non-200 atau opaque (cross-origin tanpa CORS)
            if (!networkResponse || networkResponse.status !== 200 || networkResponse.type === 'opaque') {
              return networkResponse;
            }

            // Simpan ke cache untuk pemakaian berikutnya
            const responseToCache = networkResponse.clone();
            caches.open(CACHE_NAME)
              .then(cache => cache.put(request, responseToCache))
              .catch(() => { /* Gagal simpan cache, abaikan */ });

            return networkResponse;
          })
          .catch(() => {
            // 🌐 Network gagal dan tidak ada di cache
            // Untuk request navigasi (halaman HTML), kembalikan index.html sebagai fallback
            if (request.mode === 'navigate' || request.destination === 'document') {
              console.log('[SW] 🔄 Offline fallback — mengembalikan index.html');
              return caches.match('./index.html');
            }

            // Untuk aset lain (gambar, font, dll), biarkan error
            console.warn('[SW] ❌ Tidak dapat memuat:', request.url);
            return new Response('Offline — Aset tidak tersedia', {
              status: 503,
              statusText: 'Service Unavailable',
              headers: { 'Content-Type': 'text/plain; charset=utf-8' }
            });
          });
      })
  );
});

// ──────────────────────────────────────────────
//  EVENT: MESSAGE
//  Menerima pesan dari halaman utama (opsional)
// ──────────────────────────────────────────────

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    console.log('[SW] 📨 Menerima perintah SKIP_WAITING');
    self.skipWaiting();
  }

  if (event.data && event.data.type === 'CLEAR_CACHE') {
    console.log('[SW] 📨 Menerima perintah CLEAR_CACHE');
    event.waitUntil(
      caches.delete(CACHE_NAME)
        .then(() => {
          console.log('[SW] ✅ Cache dibersihkan');
          // Kirim konfirmasi ke klien
          if (event.source) {
            event.source.postMessage({ type: 'CACHE_CLEARED' });
          }
        })
    );
  }
});

// ──────────────────────────────────────────────
//  LOG AWAL
// ──────────────────────────────────────────────

console.log(`[SW] 🚀 Service Worker v${CACHE_NAME} siap`);
