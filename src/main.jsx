import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import "./app.css";
import "./auth-fix.css";
import "./preview-overflow-fix.css";

// Keep API POSTs explicitly marked so Cloudflare never treats them like
// static-asset navigations. The Worker uses the pathname, so query strings
// do not change the API route.
const nativeFetch = window.fetch.bind(window);
window.fetch = (input, init = {}) => {
  try {
    const rawUrl = typeof input === "string" ? input : input?.url;
    if (rawUrl) {
      const u = new URL(rawUrl, window.location.href);
      const method = (init.method || input?.method || "GET").toUpperCase();
      if (u.pathname === "/api/setup/bootstrap" && method === "POST") {
        u.searchParams.set("_setup", "1");
        if (typeof input === "string") return nativeFetch(u.toString(), init);
        return nativeFetch(new Request(u.toString(), input), init);
      }
      if (u.pathname === "/api/auth/login" && method === "POST") {
        u.searchParams.set("_auth", "1");
        if (typeof input === "string") return nativeFetch(u.toString(), init);
        return nativeFetch(new Request(u.toString(), input), init);
      }
    }
  } catch (_) {}
  return nativeFetch(input, init);
};


ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

if ("MutationObserver" in window) installPasswordEyes();

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("/sw.js").catch(() => {}));
}
