# SMEC

A minimal starter web application (REST API + static frontend) used to bootstrap the
Cloud Agent development environment. Replace the in-memory item store and UI with the
real SMEC application as it grows.

## Requirements

- Node.js >= 20 (developed against Node 22)
- npm

## Getting started

```bash
npm ci        # install dependencies
npm run dev   # start the dev server with auto-reload on http://localhost:3000
```

Then open http://localhost:3000 in a browser.

## Scripts

| Command | Description |
| --- | --- |
| `npm run dev` | Start the server with `--watch` auto-reload. |
| `npm start` | Start the server without watch. |
| `npm test` | Run the automated tests (Node's built-in test runner). |
| `npm run lint` | Syntax-check the source files. |

## API

| Method | Path | Description |
| --- | --- | --- |
| `GET` | `/api/health` | Service health check. |
| `GET` | `/api/items` | List all items. |
| `POST` | `/api/items` | Create an item (`{ "title": "..." }`). |
| `PATCH` | `/api/items/:id` | Update an item (`{ "done": true }` and/or `{ "title": "..." }`). |
| `DELETE` | `/api/items/:id` | Delete an item. |

## Cloud Agent environment

`.cursor/environment.json` configures the Cloud Agent development environment:

- `install`: `npm ci` refreshes dependencies from the lockfile.
- `terminals`: a `dev-server` terminal runs `npm run dev`.
- `ports`: exposes port `3000` for the web app.
