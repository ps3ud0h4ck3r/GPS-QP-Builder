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

function addPasswordEye(input) {
  if (!input || input.dataset.passwordEye === "1") return;
  const parent = input.parentElement;
  if (!parent) return;
  input.dataset.passwordEye = "1";
  const wrap = document.createElement("div");
  wrap.style.position = "relative";
  wrap.style.width = "100%";
  parent.insertBefore(wrap, input);
  wrap.appendChild(input);
  input.style.paddingRight = "42px";
  const button = document.createElement("button");
  button.type = "button";
  button.setAttribute("aria-label", "Show password");
  button.title = "Show password";
  button.textContent = "👁️";
  Object.assign(button.style, {
    position: "absolute",
    right: "7px",
    top: "50%",
    transform: "translateY(-50%)",
    border: "0",
    background: "transparent",
    cursor: "pointer",
    padding: "5px 6px",
    fontSize: "17px",
    lineHeight: "1",
    color: "#475569",
    zIndex: "2"
  });
  button.addEventListener("click", () => {
    const visible = input.type === "text";
    input.type = visible ? "password" : "text";
    button.textContent = visible ? "👁️" : "🙈";
    button.setAttribute("aria-label", visible ? "Show password" : "Hide password");
    button.title = visible ? "Show password" : "Hide password";
    input.focus();
  });
  wrap.appendChild(button);
}

function installPasswordEyes() {
  document.querySelectorAll('input[type="password"]').forEach(addPasswordEye);
  const observer = new MutationObserver(() => {
    document.querySelectorAll('input[type="password"]').forEach(addPasswordEye);
  });
  observer.observe(document.body, { childList: true, subtree: true });
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

if ("MutationObserver" in window) installPasswordEyes();

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("/sw.js").catch(() => {}));
}
