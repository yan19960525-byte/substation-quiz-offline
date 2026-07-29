const CACHE_NAME = "beiti-local-v2-3";
const BASE_PATH = new URL("./", self.location.href).pathname;
const CORE_FILES = [BASE_PATH, `${BASE_PATH}manifest.webmanifest`, `${BASE_PATH}icon-192.png`, `${BASE_PATH}icon-512.png`];

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    const page = await fetch(BASE_PATH);
    const copy = page.clone();
    const html = await page.text();
    await cache.put(BASE_PATH, copy);
    const assetUrls = Array.from(html.matchAll(/(?:src|href)=["']([^"']+)["']/g))
      .map((match) => match[1])
      .filter((url) => url.startsWith("/") && !url.startsWith("//"));
    await cache.addAll(Array.from(new Set([...CORE_FILES.slice(1), ...assetUrls])));
    await self.skipWaiting();
  })());
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.filter((name) => name !== CACHE_NAME).map((name) => caches.delete(name)));
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== "GET" || url.origin !== self.location.origin) return;

  event.respondWith((async () => {
    const cached = await caches.match(request, { ignoreSearch: request.mode === "navigate" });
    if (cached) return cached;
    try {
      const response = await fetch(request);
      if (response.ok) {
        const cache = await caches.open(CACHE_NAME);
        await cache.put(request, response.clone());
      }
      return response;
    } catch {
      if (request.mode === "navigate") return (await caches.match(BASE_PATH)) || Response.error();
      return Response.error();
    }
  })());
});
