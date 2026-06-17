// Nova service worker — receives push notifications and shows them.
self.addEventListener("install", (e) => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));

self.addEventListener("push", (event) => {
  let data = { title: "Nova", body: "You have a reminder." };
  try { if (event.data) data = event.data.json(); } catch (_) {}
  const title = data.title || "Nova";
  const options = {
    body: data.body || "",
    tag: data.tag || "nova",
    renotify: true,
    icon: "/icon-192.png",
    badge: "/icon-192.png",
    data: { url: "/" },
  };
  event.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((list) => {
      for (const c of list) { if ("focus" in c) return c.focus(); }
      if (self.clients.openWindow) return self.clients.openWindow("/");
    })
  );
});
