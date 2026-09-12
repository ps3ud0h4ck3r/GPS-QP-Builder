import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import "./app.css";
import "./auth-fix.css";
import "./preview-overflow-fix.css";

// Keep the setup request explicitly marked so Cloudflare never treats the
// POST like a static-asset navigation. The Worker accepts the same endpoint
// with this query string because the pathname remains /api/setup/bootstrap.
const nativeFetch = window.fetch.bind(window);
window.fetch = (input, init = {}) => {
  try {
    const rawUrl = typeof input === "string" ? input : input?.url;
    if (rawUrl) {
      const u = new URL(rawUrl, window.location.href);
      if (u.pathname === "/api/setup/bootstrap" && (init.method || input?.method || "GET").toUpperCase() === "POST") {
        u.searchParams.set("_setup", "1");
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

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("/sw.js").catch(() => {}));
}
