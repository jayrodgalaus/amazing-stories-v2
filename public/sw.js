// sw.js
self.addEventListener('install', () => {
  self.skipWaiting(); // Activate immediately
});

self.addEventListener('activate', () => {
  self.clients.claim(); // Take control of the page
});
