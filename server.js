const express = require("express");

const app = express();
const port = process.env.PORT || 3000;
const apiBaseUrl = process.env.OMNIROUTE_API_URL || "http://13.62.104.102:20128/v1/chat/completions";

app.use(express.json({ limit: "100kb" }));
app.use(express.static("dist"));

app.post("/api/chat", async (request, response) => {
  if (!process.env.OMNIROUTE_API_KEY) {
    return response.status(500).json({ error: "OMNIROUTE_API_KEY n'est pas configurée sur le serveur." });
  }

  try {
    const upstream = await fetch(apiBaseUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${process.env.OMNIROUTE_API_KEY}`,
      },
      body: JSON.stringify({
        model: request.body.model || "auto",
        messages: request.body.messages || [],
        stream: false,
      }),
    });
    const body = await upstream.text();
    response.status(upstream.status).type("application/json").send(body);
  } catch (error) {
    response.status(502).json({ error: "Impossible de joindre Omniroute." });
  }
});

app.get("/health", (_request, response) => response.json({ ok: true }));
app.get("/{*splat}", (_request, response) => response.sendFile("taskpane.html", { root: "dist" }));

app.listen(port, () => console.log(`Omniroute Word add-in listening on port ${port}`));
