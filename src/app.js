import express from "express";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const __dirname = dirname(fileURLToPath(import.meta.url));

/**
 * Build the Express application.
 *
 * The item store is kept in-memory on purpose: this is a starter scaffold
 * whose main job is to give the Cloud Agent environment something real to
 * install, test, build, and run end-to-end. Swap it for a real datastore
 * when the SMEC application grows.
 */
export function createApp() {
  const app = express();
  app.use(express.json());

  /** @type {{ id: number, title: string, done: boolean, createdAt: string }[]} */
  const items = [];
  let nextId = 1;

  app.get("/api/health", (_req, res) => {
    res.json({ status: "ok", service: "smec", time: new Date().toISOString() });
  });

  app.get("/api/items", (_req, res) => {
    res.json({ items });
  });

  app.post("/api/items", (req, res) => {
    const title = typeof req.body?.title === "string" ? req.body.title.trim() : "";
    if (!title) {
      res.status(400).json({ error: "title is required" });
      return;
    }
    const item = { id: nextId++, title, done: false, createdAt: new Date().toISOString() };
    items.push(item);
    res.status(201).json({ item });
  });

  app.patch("/api/items/:id", (req, res) => {
    const id = Number(req.params.id);
    const item = items.find((it) => it.id === id);
    if (!item) {
      res.status(404).json({ error: "item not found" });
      return;
    }
    if (typeof req.body?.done === "boolean") item.done = req.body.done;
    if (typeof req.body?.title === "string" && req.body.title.trim()) item.title = req.body.title.trim();
    res.json({ item });
  });

  app.delete("/api/items/:id", (req, res) => {
    const id = Number(req.params.id);
    const index = items.findIndex((it) => it.id === id);
    if (index === -1) {
      res.status(404).json({ error: "item not found" });
      return;
    }
    const [removed] = items.splice(index, 1);
    res.json({ item: removed });
  });

  app.use(express.static(join(__dirname, "..", "public")));

  return app;
}
