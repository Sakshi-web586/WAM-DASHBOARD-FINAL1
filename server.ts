import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Body parser for JSON
  app.use(express.json());

  // API Proxy for Google Apps Script
  // This bypasses CORS by doing the fetch from the server
  app.all("/api/proxy", async (req, res) => {
    const targetUrl = req.query.url as string;
    if (!targetUrl) {
      return res.status(400).json({ error: "Missing url parameter" });
    }

    try {
      const method = req.method;
      console.log(`Proxying ${method} request to: ${targetUrl}`);
      const headers: Record<string, string> = {
        "User-Agent": "Mozilla/5.0 (Node.js Proxy)",
      };

      // Forward Content-Type if it exists
      if (req.headers["content-type"]) {
        headers["Content-Type"] = req.headers["content-type"];
      }

      const options: any = {
        method,
        headers,
        redirect: "follow",
      };

      if (method !== "GET" && method !== "HEAD") {
        options.body = JSON.stringify(req.body);
        headers["Content-Type"] = "application/json";
      }

      const response = await fetch(targetUrl, options);
      const data = await response.text();

      // Forward status and content type
      res.status(response.status);
      const contentType = response.headers.get("content-type");
      if (contentType) {
        res.setHeader("Content-Type", contentType);
      }
      
      res.send(data);
    } catch (error: any) {
      console.error("Proxy error:", error);
      res.status(500).json({ 
        error: "Proxy failed to reach Google Apps Script", 
        message: error.message 
      });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    // Production serving
    const distPath = path.join(__dirname, "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
