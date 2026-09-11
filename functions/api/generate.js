export async function onRequestPost({ request, env }) {
  try {
    if (!env.ANTHROPIC_API_KEY) {
      return json({ error: "AI service is not configured. Add ANTHROPIC_API_KEY to the site's environment variables." }, 500);
    }

    const body = await request.json();
    const prompt = typeof body?.prompt === "string" ? body.prompt.trim() : "";
    if (!prompt) return json({ error: "Missing prompt." }, 400);
    if (prompt.length > 12000) return json({ error: "Prompt is too long." }, 400);

    const upstream = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01"
      },
      body: JSON.stringify({
        model: env.ANTHROPIC_MODEL || "claude-sonnet-4-6",
        max_tokens: 3000,
        temperature: 0.3,
        messages: [{ role: "user", content: prompt }]
      })
    });

    const data = await upstream.json();
    if (!upstream.ok) {
      return json({ error: data?.error?.message || "Anthropic request failed." }, upstream.status);
    }

    const text = (data.content || []).map(block => block.text || "").join("");
    return json({ content: text });
  } catch (error) {
    return json({ error: error?.message || "Unexpected server error." }, 500);
  }
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" }
  });
}
