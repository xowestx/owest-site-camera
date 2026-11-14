/* O-Plus Client Portal service worker: lets phones install the portal as an app.
   Network only (the portal always shows live data). Push notifications will be added here later. */
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", () => {});
