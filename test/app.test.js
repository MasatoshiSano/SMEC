import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { createApp } from "../src/app.js";

let server;
let baseUrl;

before(async () => {
  const app = createApp();
  await new Promise((resolve) => {
    server = app.listen(0, "127.0.0.1", resolve);
  });
  const { port } = server.address();
  baseUrl = `http://127.0.0.1:${port}`;
});

after(() => {
  server?.close();
});

test("health endpoint reports ok", async () => {
  const res = await fetch(`${baseUrl}/api/health`);
  assert.equal(res.status, 200);
  const body = await res.json();
  assert.equal(body.status, "ok");
  assert.equal(body.service, "smec");
});

test("items start empty", async () => {
  const res = await fetch(`${baseUrl}/api/items`);
  const body = await res.json();
  assert.deepEqual(body.items, []);
});

test("create, update and delete an item end to end", async () => {
  const createRes = await fetch(`${baseUrl}/api/items`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title: "Ship SMEC" }),
  });
  assert.equal(createRes.status, 201);
  const { item } = await createRes.json();
  assert.equal(item.title, "Ship SMEC");
  assert.equal(item.done, false);

  const patchRes = await fetch(`${baseUrl}/api/items/${item.id}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ done: true }),
  });
  assert.equal(patchRes.status, 200);
  const patched = await patchRes.json();
  assert.equal(patched.item.done, true);

  const listRes = await fetch(`${baseUrl}/api/items`);
  const listed = await listRes.json();
  assert.equal(listed.items.length, 1);

  const deleteRes = await fetch(`${baseUrl}/api/items/${item.id}`, { method: "DELETE" });
  assert.equal(deleteRes.status, 200);

  const afterList = await (await fetch(`${baseUrl}/api/items`)).json();
  assert.equal(afterList.items.length, 0);
});

test("rejects empty titles", async () => {
  const res = await fetch(`${baseUrl}/api/items`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title: "   " }),
  });
  assert.equal(res.status, 400);
});

test("returns 404 for unknown item", async () => {
  const res = await fetch(`${baseUrl}/api/items/9999`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: "{}" });
  assert.equal(res.status, 404);
});
