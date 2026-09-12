// GPS QP Builder deployment trigger, 2026-09-12
import app from "./index.js";

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (request.method === "POST" && url.pathname === "/api/setup" && url.searchParams.get("action") === "bootstrap") {
      url.pathname = "/api/setup/bootstrap";
      url.search = "";
      request = new Request(url.toString(), request);
    }
    return app.fetch(request, env, ctx);
  }
};
