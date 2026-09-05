const itemsEl = document.getElementById("items");
const emptyEl = document.getElementById("empty");
const formEl = document.getElementById("new-item");
const titleEl = document.getElementById("title");
const healthEl = document.getElementById("health");

async function refreshHealth() {
  try {
    const res = await fetch("/api/health");
    const data = await res.json();
    healthEl.textContent = `API: ${data.status}`;
    healthEl.className = "badge badge--ok";
  } catch {
    healthEl.textContent = "API: unreachable";
    healthEl.className = "badge badge--err";
  }
}

function render(items) {
  itemsEl.innerHTML = "";
  emptyEl.classList.toggle("hidden", items.length > 0);
  for (const item of items) {
    const li = document.createElement("li");
    li.className = `item${item.done ? " item--done" : ""}`;

    const title = document.createElement("span");
    title.className = "item__title";
    title.textContent = item.title;

    const toggle = document.createElement("button");
    toggle.className = "item__toggle";
    toggle.textContent = item.done ? "Undo" : "Done";
    toggle.onclick = () => updateItem(item.id, { done: !item.done });

    const del = document.createElement("button");
    del.className = "item__delete";
    del.textContent = "Delete";
    del.onclick = () => deleteItem(item.id);

    li.append(title, toggle, del);
    itemsEl.appendChild(li);
  }
}

async function loadItems() {
  const res = await fetch("/api/items");
  const data = await res.json();
  render(data.items);
}

async function addItem(title) {
  await fetch("/api/items", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title }),
  });
  await loadItems();
}

async function updateItem(id, patch) {
  await fetch(`/api/items/${id}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(patch),
  });
  await loadItems();
}

async function deleteItem(id) {
  await fetch(`/api/items/${id}`, { method: "DELETE" });
  await loadItems();
}

formEl.addEventListener("submit", (event) => {
  event.preventDefault();
  const title = titleEl.value.trim();
  if (!title) return;
  titleEl.value = "";
  addItem(title);
});

refreshHealth();
loadItems();
