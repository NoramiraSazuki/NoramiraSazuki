"use strict";
// Plain JS, no external libraries. All document text is inserted with textContent (never as HTML).

const TOKEN = document.querySelector('meta[name="pilot-token"]').content;
const $ = (s, el = document) => el.querySelector(s);
const state = { docs: [], sort: "name", asc: true, selected: new Map(), modes: [], methods: {},
                review: { doc: null, page: 1, count: 0, pages: [] } };

async function api(path, opts = {}) {
  const o = { ...opts, headers: { ...(opts.headers || {}) } };
  if (o.method && o.method !== "GET") { o.headers["X-Pilot-Token"] = TOKEN; o.headers["Content-Type"] = "application/json"; }
  const r = await fetch(path, o);
  const ct = r.headers.get("content-type") || "";
  const body = ct.includes("json") ? await r.json() : await r.text();
  if (!r.ok) { const e = new Error((body && body.detail) || body || r.statusText); e.status = r.status; e.body = body; throw e; }
  return body;
}
const post = (path, data) => api(path, { method: "POST", body: JSON.stringify(data || {}) });

function el(tag, attrs = {}, ...children) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") e.className = v;
    else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
    else if (v !== undefined && v !== null && v !== false) e.setAttribute(k, v === true ? "" : v);
  }
  for (const c of children.flat()) if (c !== null && c !== undefined) e.append(c instanceof Node ? c : String(c));
  return e;
}
const badge = (s) => el("span", { class: `badge b-${s || "none"}` }, s || "—");
function toast(msg, ms = 4000) { const t = $("#toast"); t.textContent = msg; t.classList.remove("hidden");
  clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.add("hidden"), ms); }
const fmtSecs = (s) => s == null ? "" : s < 90 ? `${s}s` : s < 5400 ? `${Math.round(s / 60)} min` : `${(s / 3600).toFixed(1)} h`;

// ------------------------------------------------------------------ tabs + status
document.querySelectorAll("button.tab").forEach(b => b.addEventListener("click", () => showTab(b.dataset.tab)));
function showTab(name) {
  document.querySelectorAll("button.tab").forEach(b => b.classList.toggle("active", b.dataset.tab === name));
  document.querySelectorAll("section.panel").forEach(s => s.classList.toggle("hidden", s.id !== `tab-${name}`));
  if (name === "jobs") loadJobs();
  if (name === "review") loadReviewDocs();
}

async function refreshState() {
  try {
    const s = await api("/api/state");
    state.modes = s.modes; state.methods = s.methods;
    // Simple status: the NAS connection, plus a short note only when something needs attention.
    const st = $("#status"), sc = s.scan || {};
    const notes = [];
    if (sc.running) notes.push(el("span", { class: "muted" }, sc.phase === "checking NAS"
      ? `Checking NAS… ${(sc.checked ?? 0).toLocaleString()} PDFs checked` : "Reading inventory…"));
    if (!s.worker_alive) notes.push(el("span", { class: "bad" }, "Background worker not running"));
    if (s.disk_free_gb < s.min_free_gb) notes.push(el("span", { class: "bad" }, `Low disk space: ${s.disk_free_gb} GB free`));
    st.replaceChildren(el("span", { class: "nas " + (s.source_available ? "ok" : "bad") },
      s.source_available ? "● NAS – Connected" : "● NAS – Not connected"), ...notes.flatMap(n => ["  ·  ", n]));
    st.title = [`NAS folder: ${s.source_root}`, `Worker: ${s.worker_alive ? "running" : "not responding"} (${activityText(s.worker_activity)})`,
      `Disk: ${s.disk_free_gb} GB free`, `Data folder: ${s.data_root}`].join("\n");
    state.lastScan = sc;
    state.largeJobPages = s.large_job_pages;
    if (!$("#mode").dataset.built) buildModes();
    if (state.invLoaded !== undefined && s.inventory_loaded !== state.invLoaded) loadDocs();  // catalogue (re)loaded
    state.invLoaded = s.inventory_loaded;
    // While the NAS check runs, refresh statuses every ~10 s so files turn OK/MISSING as their folder is read.
    if (sc.running && sc.checked !== state.lastChecked && Date.now() - (state.lastDocsLoad || 0) > 10000
        && !$("#tab-catalogue").classList.contains("hidden") && !document.activeElement?.closest?.("#docs, .pager")) {
      state.lastChecked = sc.checked; state.lastDocsLoad = Date.now(); loadDocs();
    }
  } catch (e) { $("#status").textContent = "App not responding: " + e.message; }
}
function activityText(a) {
  if (!a || !a.state || a.state === "idle") return "idle";
  if (a.state === "hashing") return `job ${a.job}: checking source identity`;
  return `job ${a.job}: document ${a.doc}, page ${a.page}`;
}

function buildModes() {
  const sel = $("#mode"); sel.dataset.built = "1";
  sel.replaceChildren(...state.modes.map(m => el("option", { value: m.id, disabled: !m.implemented, title: m.help },
    m.label + (m.implemented ? "" : " — not yet implemented"))));
  sel.value = (state.modes.find(m => m.implemented) || {}).id || "";
  sel.title = (state.modes.find(m => m.id === sel.value) || {}).help || "";
}

// ------------------------------------------------------------------ catalogue (folder tree)
const nameCmp = (a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
const store_ = { get(k, d) { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch { return d; } },
                 set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* private mode */ } } };
state.expanded = new Set(store_.get("pilot.expanded", []));
state.page = 1;
state.pageSize = store_.get("pilot.pageSize", 100);
state.folders = [];

async function loadDocs() {
  const [docs, rep] = await Promise.all([api("/api/documents"), api("/api/inventory/report")]);
  state.docs = docs.map(d => { const i = d.rel_path.lastIndexOf("/");
    return { ...d, folder: i >= 0 ? d.rel_path.slice(0, i) : "", name: i >= 0 ? d.rel_path.slice(i + 1) : d.rel_path }; });
  state.folders = rep.folders || [];
  renderBatches(); renderDocs(); showInvReport(rep);
}
function showInvReport(r) {
  const parts = Object.entries(r.counts).map(([k, v]) => `${k}: ${v}`).join(", ");
  const t = r.titles || {};
  const titleNote = t.error ? ` · titles: ${t.error}` : ` · titles: ${r.with_title}/${r.documents}` +
    (t.duplicate_count ? `, ${t.duplicate_count} duplicate path(s) left blank` : "");
  $("#inv-report").textContent = r.documents ? `${r.documents} PDFs · ${r.pages ?? "?"} pages · ${(r.bytes / 1e9).toFixed(1)} GB (inventory figures) · ${parts}${titleNote}` : "Inventory not loaded yet.";
  const sc = state.lastScan || {};
  if (!sc.running && sc.error) $("#inv-report").append(el("span", { class: "summary-warn" }, ` · Last NAS check: ${sc.error}`));
  if (sc.previous_interrupted && sc.running) $("#inv-report").append(el("span", { class: "muted" }, " · previous NAS check was interrupted; checking again"));
  const ul = $("#unlisted"); ul.replaceChildren(...r.unlisted.map(u => el("li", {}, u)));
  $("#unlisted-box").classList.toggle("hidden", !r.unlisted.length);
  $("#unlisted-count").textContent = r.unlisted.length;
}

function filteredDocs() {
  const q = $("#search").value.trim().toLowerCase(), f = $("#filter-status").value, b = $("#filter-batch").value;
  return state.docs.filter(d => (!b || d.rel_path.startsWith(b + "/"))
    && (!q || d.rel_path.toLowerCase().includes(q) || (d.title || "").toLowerCase().includes(q))
    && (!f || d.resolve_status === f));
}
function buildTree(docs, includeEmpty, batch = "") {
  const root = { name: "", path: "", folders: new Map(), files: [] };
  const node = (path) => {
    let n = root;
    if (!path) return n;
    let acc = "";
    for (const part of path.split("/")) {
      acc = acc ? `${acc}/${part}` : part;
      if (!n.folders.has(part)) n.folders.set(part, { name: part, path: acc, folders: new Map(), files: [] });
      n = n.folders.get(part);
    }
    return n;
  };
  if (includeEmpty) state.folders.filter(f => !batch || f === batch || f.startsWith(batch + "/")).forEach(node);
  docs.forEach(d => node(d.folder).files.push(d));
  const agg = (n) => {
    n.count = n.files.length; n.pages = 0; n.bytes = 0; n.processed = 0; n.issues = 0; n.ids = [];
    for (const d of n.files) { n.pages += d.inv_pages || 0; n.bytes += d.inv_size || 0;
      n.processed += d.processed_before ? 1 : 0; n.issues += d.resolve_status === "OK" ? 0 : 1;
      if (d.resolve_status === "OK") n.ids.push(d.id); }
    for (const c of n.folders.values()) { agg(c); n.count += c.count; n.pages += c.pages; n.bytes += c.bytes;
      n.processed += c.processed; n.issues += c.issues; n.ids.push(...c.ids); }
  };
  agg(root);
  return root;
}
function sortedChildren(n) {
  const k = state.sort, dir = state.asc ? 1 : -1;
  const folders = [...n.folders.values()].map(f => ({ kind: "folder", key: f.name, f }));
  if (k === "name") {  // Finder-like: folders and files mixed, by name
    return [...folders, ...n.files.map(d => ({ kind: "file", key: d.name, d }))]
      .sort((a, b) => nameCmp(a.key, b.key) * dir);
  }
  const files = [...n.files].sort((a, b) => {
    const x = a[k], y = b[k];
    if (x == null && y == null) return nameCmp(a.name, b.name); if (x == null) return 1; if (y == null) return -1;
    return (typeof x === "number" && typeof y === "number" ? x - y : nameCmp(String(x), String(y))) * dir;
  });
  return [...folders.sort((a, b) => nameCmp(a.key, b.key)), ...files.map(d => ({ kind: "file", key: d.name, d }))];
}
const fmtSize = (b) => !b ? "" : b >= 1e9 ? `${(b / 1e9).toFixed(1)} GB` : `${(b / 1e6).toFixed(1)} MB`;

function renderDocs(keepPage = true) {
  if (!keepPage) state.page = 1;
  const batch = $("#filter-batch").value;
  const filtering = !!($("#search").value.trim() || $("#filter-status").value);
  const root = buildTree(filteredDocs(), !filtering, batch);
  const isOpen = (path) => filtering || state.expanded.has(path) || path === batch;
  // Flatten the visible tree into rows, then show one page of them.
  const rows = [];
  const walk = (n, depth) => {
    for (const c of sortedChildren(n)) {
      if (c.kind === "folder") {
        const open = isOpen(c.f.path);
        rows.push({ kind: "folder", f: c.f, depth, open });
        if (open) walk(c.f, depth + 1);
      } else rows.push({ kind: "file", d: c.d, depth, parent: n.path });
    }
  };
  walk(root, 0);
  const size = state.pageSize, pages = Math.max(1, Math.ceil(rows.length / size));
  state.page = Math.min(Math.max(1, state.page || 1), pages);
  const from = (state.page - 1) * size, slice = rows.slice(from, from + size);
  const tb = $("#docs tbody"); tb.replaceChildren();
  if (slice.length && slice[0].depth > 0) {  // page starts inside a folder: show where we are
    const first = slice[0], parent = first.kind === "file" ? first.parent : first.f.path.split("/").slice(0, -1).join("/");
    tb.append(el("tr", { class: "context-row" }, el("td", {}), el("td", { colspan: 8, class: "small muted" },
      "📂 ", parent.split("/").join(" › "), " (continued)")));
  }
  for (const r of slice) tb.append(r.kind === "folder" ? folderRow(r.f, r.depth, r.open) : fileRow(r.d, r.depth));
  if (!rows.length) tb.append(el("tr", {}, el("td", { colspan: 9, class: "muted" },
    state.docs.length ? "No matching PDFs." : "Catalogue not loaded yet — see the status bar for the NAS check.")));
  renderPager(rows.length, pages, from, slice.length);
  updateSel();
}
function renderPager(total, pages, from, shown) {
  const go = (p) => { state.page = p; renderDocs(); $("#tab-catalogue .tablewrap").scrollTop = 0; };
  for (const box of document.querySelectorAll(".pager")) {
    const input = el("input", { type: "number", min: 1, max: pages, value: state.page, class: "narrow", title: "Go to page" });
    input.addEventListener("change", () => go(Math.min(pages, Math.max(1, Number(input.value) || 1))));
    const sizeSel = el("select", { title: "Rows per page" }, [50, 100, 200, 500].map(n =>
      el("option", { value: n, selected: n === state.pageSize }, `${n} rows per page`)));
    sizeSel.addEventListener("change", () => { state.pageSize = Number(sizeSel.value); store_.set("pilot.pageSize", state.pageSize); renderDocs(false); });
    box.replaceChildren(
      el("button", { class: "small", disabled: state.page <= 1, onclick: () => go(1) }, "« First"),
      el("button", { class: "small", disabled: state.page <= 1, onclick: () => go(state.page - 1) }, "‹ Prev"),
      el("span", { class: "small" }, " Page "), input, el("span", { class: "small" }, ` of ${pages} `),
      el("button", { class: "small", disabled: state.page >= pages, onclick: () => go(state.page + 1) }, "Next ›"),
      el("button", { class: "small", disabled: state.page >= pages, onclick: () => go(pages) }, "Last »"),
      el("span", { class: "small muted" }, total ? `  Showing rows ${(from + 1).toLocaleString()}–${(from + shown).toLocaleString()} of ${total.toLocaleString()}  ` : "  "),
      sizeSel);
  }
}
function renderBatches() {
  const sel = $("#filter-batch"), cur = sel.value;
  const counts = new Map();
  for (const f of state.folders) if (!f.includes("/")) counts.set(f, 0);
  for (const d of state.docs) { const b = d.rel_path.includes("/") ? d.rel_path.split("/")[0] : ""; if (b) counts.set(b, (counts.get(b) || 0) + 1); }
  sel.replaceChildren(el("option", { value: "" }, "All batches"),
    ...[...counts.keys()].sort(nameCmp).map(b => el("option", { value: b }, `${b} (${counts.get(b)} PDFs)`)));
  sel.value = counts.has(cur) ? cur : "";
}
function setIndent(e, depth) { e.style.paddingLeft = (6 + depth * 18) + "px"; return e; }
function folderRow(f, depth, open) {
  const sel = f.ids.filter(id => state.selected.has(id)).length;
  const cb = el("input", { type: "checkbox", disabled: !f.ids.length, title: f.ids.length ? `Select all ${f.ids.length} processable PDFs in this folder` : "No processable PDFs here" });
  cb.checked = f.ids.length > 0 && sel === f.ids.length; cb.indeterminate = sel > 0 && sel < f.ids.length;
  cb.addEventListener("change", () => { for (const id of f.ids) cb.checked ? state.selected.set(id, state.selected.get(id) || "") : state.selected.delete(id); renderDocs(); });
  const toggle = () => { open ? state.expanded.delete(f.path) : state.expanded.add(f.path); store_.set("pilot.expanded", [...state.expanded]); renderDocs(); };
  const name = setIndent(el("td", { class: "tree-name folder", onclick: toggle, title: f.path },
    el("span", { class: "caret" }, open ? "▾" : "▸"), el("span", { class: "icon" }, "📁"), f.name,
    el("span", { class: "muted small" }, f.count ? `  ${f.count} PDF${f.count === 1 ? "" : "s"}` : "  no PDFs")), depth);
  return el("tr", { class: "folder-row" + (sel ? " sel" : "") }, el("td", {}, cb), name, el("td", {}),
    el("td", { class: "num" }, f.pages || ""), el("td", { class: "num" }, fmtSize(f.bytes)),
    el("td", {}, f.issues ? el("span", { class: "flag small" }, `${f.issues} not matched`) : ""),
    el("td", {}), el("td", { class: "small muted" }, f.processed ? `${f.processed} processed` : ""), el("td", {}));
}
function fileRow(d, depth) {
  const ok = d.resolve_status === "OK";
  const cb = el("input", { type: "checkbox", disabled: !ok, checked: state.selected.has(d.id),
    title: ok ? "" : "Cannot process: " + (d.note || d.resolve_status) });
  const pages = el("input", { class: "pages", placeholder: "all", disabled: !ok, value: state.selected.get(d.id) || "" });
  const tr = el("tr", { class: state.selected.has(d.id) ? "sel" : "" });
  cb.addEventListener("change", () => { cb.checked ? state.selected.set(d.id, pages.value) : state.selected.delete(d.id); tr.classList.toggle("sel", cb.checked); updateSel(); });
  pages.addEventListener("input", () => { if (pages.value && !cb.checked) { cb.checked = true; tr.classList.add("sel"); }
    if (cb.checked) state.selected.set(d.id, pages.value); updateSel(); });
  tr.append(el("td", {}, cb),
    setIndent(el("td", { class: "tree-name", title: `${d.rel_path}${d.inv_row_id ? "  (inventory # " + d.inv_row_id + ")" : ""}${d.note ? "\n" + d.note : ""}` },
      el("span", { class: "caret" }, ""), el("span", { class: "icon" }, "📄"), d.name), depth),
    el("td", { title: d.title_source ? `from DIRt sheet "${d.title_source}"` : "" }, d.title || el("span", { class: "muted" }, "—")),
    el("td", { class: "num" }, d.inv_pages ?? ""), el("td", { class: "num" }, fmtSize(d.inv_size)),
    el("td", {}, badge(d.resolve_status), d.note ? el("div", { class: "muted small" }, d.note) : null),
    el("td", {}, d.inv_duplicate_group ? `${d.inv_duplicate_group} (${d.inv_copies ?? "?"} copies)` : ""),
    el("td", {}, d.last_status ? [badge(d.last_status), el("span", { class: "muted small" }, ` job ${d.last_job}`)] : ""),
    el("td", {}, pages));
  return tr;
}
function selectedPages() {
  const byId = new Map(state.docs.map(d => [d.id, d]));
  return [...state.selected.entries()].reduce((t, [id, spec]) => t + (spec ? countSpec(spec) : (byId.get(id)?.inv_pages || 0)), 0);
}
function updateSel() {
  const n = state.selected.size, pages = selectedPages(), big = pages > (state.largeJobPages || 1000);
  $("#sel-summary").textContent = n ? `${n} document(s), about ${pages.toLocaleString()} page(s) selected.` : "Nothing selected.";
  $("#sel-summary").classList.toggle("flag", big);
  $("#start").disabled = !n;
  const all = $("#select-all"), ids = filteredDocs().filter(d => d.resolve_status === "OK").map(d => d.id);
  const s = ids.filter(id => state.selected.has(id)).length;
  all.checked = ids.length > 0 && s === ids.length; all.indeterminate = s > 0 && s < ids.length;
}
function countSpec(spec) { let n = 0; for (const p of spec.split(",")) { const m = p.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/); if (m) n += (Number(m[2] || m[1]) - Number(m[1]) + 1); } return n; }
function estimate(pages) {  // rough: digital text ~0.05 s/page, OCR ~1-3 s/page on this Mac
  const lo = Math.max(1, Math.round(pages * 0.05)), hi = Math.round(pages * 3);
  return `${fmtSecs(lo)} (all digital) to ${fmtSecs(hi)} (all scanned)`;
}

document.querySelectorAll("#docs th[data-sort]").forEach(th => th.addEventListener("click", () => {
  state.asc = state.sort === th.dataset.sort ? !state.asc : true; state.sort = th.dataset.sort; renderDocs(false); }));
let searchTimer;
$("#search").addEventListener("input", () => { clearTimeout(searchTimer); searchTimer = setTimeout(() => renderDocs(false), 200); });
$("#filter-status").addEventListener("change", () => renderDocs(false));
$("#filter-batch").addEventListener("change", () => renderDocs(false));
$("#select-all").addEventListener("change", (e) => {
  for (const d of filteredDocs()) if (d.resolve_status === "OK") e.target.checked ? state.selected.set(d.id, state.selected.get(d.id) || "") : state.selected.delete(d.id);
  renderDocs(); });
$("#expand-all").addEventListener("click", () => { state.folders.forEach(f => state.expanded.add(f)); store_.set("pilot.expanded", [...state.expanded]); renderDocs(); });
$("#collapse-all").addEventListener("click", () => { state.expanded.clear(); store_.set("pilot.expanded", []); renderDocs(false); });
$("#clear-sel").addEventListener("click", () => { state.selected.clear(); renderDocs(); });
$("#reload-inv").addEventListener("click", async () => {
  try { const r = await post("/api/inventory/reload"); toast(r.message + " Progress is shown in the status bar."); refreshState(); }
  catch (e) { toast("Reload failed: " + e.message, 8000); } });
state.pendingLarge = null;
$("#start").addEventListener("click", async () => {
  const mode = $("#mode").value;
  if (!mode) return toast("Choose an extraction mode.");
  const items = [...state.selected.entries()].map(([doc_id, pages]) => ({ doc_id, pages: pages || "all" }));
  const key = JSON.stringify(items);
  const confirm_large = !!(state.pendingLarge && state.pendingLarge.key === key && Date.now() - state.pendingLarge.t < 15000);
  try {
    const r = await post("/api/jobs", { items, mode, confirm_large });
    state.pendingLarge = null;
    toast(`Job ${r.job_id} queued (${r.pages.toLocaleString()} pages).`);
    state.selected.clear(); renderDocs(); showTab("jobs");
  } catch (e) {
    if (e.status === 409 && e.body?.pages) {
      state.pendingLarge = { key, t: Date.now() };
      const unk = e.body.unknown_page_counts ? ` (${e.body.unknown_page_counts} document(s) with unknown page count not included)` : "";
      toast(`Large job: ${e.body.pages.toLocaleString()} pages${unk}, above the ${e.body.threshold.toLocaleString()}-page limit. ` +
            `Estimated time ${estimate(e.body.pages)}. Click "Queue extraction job" again within 15 s to confirm.`, 15000);
    } else toast("Could not queue: " + e.message, 8000);
  } });
// window.confirm blocks automation tools; a second click within 5 s confirms instead.
function confirmInline(msg) { const now = Date.now(); if (confirmInline.t && now - confirmInline.t < 5000) { confirmInline.t = 0; return true; }
  confirmInline.t = now; toast(msg + " Click again within 5 s to confirm.", 5000); return false; }

// ------------------------------------------------------------------ jobs
async function loadJobs() {
  const list = $("#jobs-list");
  const jobs = await api("/api/jobs");
  if (!jobs.length) { list.replaceChildren(el("p", { class: "muted" }, "No jobs yet. Select documents in the Catalogue.")); return; }
  list.replaceChildren(...jobs.map(jobCard));
}
function jobCard(j) {
  const pct = j.pages_total ? Math.round(100 * j.pages_done / j.pages_total) : 0;
  const bar = el("div", { class: "progress" }, el("div")); bar.firstChild.style.width = pct + "%";
  const act = (a, label, enabled) => el("button", { class: "small", disabled: !enabled, onclick: async () => {
    try { const r = await post(`/api/jobs/${j.id}/${a}`); toast(a === "export" ? `Exported to ${r.export_dir}` : `${label}: ok`, a === "export" ? 12000 : 3000); loadJobs(); }
    catch (e) { toast(`${label} failed: ${e.message}`, 8000); } } }, label);
  const s = j.status, active = s === "running" || s === "queued";
  return el("div", { class: "card" },
    el("div", { class: "row" }, el("strong", {}, `Job ${j.id}`), badge(s), el("span", { class: "muted small" }, j.mode),
      j.reason ? el("span", { class: "flag" }, j.reason) : null,
      bar, el("span", { class: "small" }, `${j.pages_done}/${j.pages_total} pages · ${j.pages_failed} failed · ${j.pages_reused} reused`),
      j.eta_seconds ? el("span", { class: "muted small" }, `≈ ${fmtSecs(j.eta_seconds)} left`) : null),
    el("div", { class: "row" },
      act("pause", "Pause", active || s === "waiting_for_source"), act("resume", "Resume", ["paused", "cancelled", "waiting_for_source"].includes(s)),
      act("cancel", "Cancel", active || ["paused", "waiting_for_source"].includes(s)),
      act("retry", "Retry failed", ["done_with_errors", "paused", "cancelled", "waiting_for_source"].includes(s)),
      act("export", "Export results", !active),
      el("button", { class: "small", disabled: active, title: "Text, metadata, Excel and JSON for every page of this job, as one .zip",
        onclick: () => download(`/api/jobs/${j.id}/download`, {}, `job${j.id}.zip`) }, "Download results (.zip)"),
      el("button", { class: "small", onclick: async () => { const t = await api(`/api/jobs/${j.id}/sanitized-report`);
        try { await navigator.clipboard.writeText(t); toast("Sanitised report copied (ids, statuses, error types only)."); }
        catch { toast(t, 15000); } } }, "Copy sanitised report")),
    el("details", {}, el("summary", { class: "small" }, `${j.items.length} document(s)`),
      el("table", {}, el("tbody", {}, j.items.map(i => el("tr", {},
        el("td", {}, i.file_name), el("td", { class: "small" }, i.page_spec), el("td", {}, badge(i.status)),
        el("td", { class: "small" }, `${i.pages_done}/${i.pages_total}`, i.pages_failed ? ` · ${i.pages_failed} failed` : ""),
        el("td", { class: "small flag" }, [i.error, i.warning].filter(Boolean).join(" · ")),
        el("td", {}, i.sha256 ? el("button", { class: "small", onclick: () => openReview(i.doc_id, firstPage(i.page_spec)) }, "Review") : null)))))));
}
const firstPage = (spec) => { const m = String(spec).match(/\d+/); return m ? Number(m[0]) : 1; };

// ------------------------------------------------------------------ review
async function loadReviewDocs() {
  const docs = await api("/api/processed-documents"); const sel = $("#rv-doc"); const cur = sel.value;
  sel.replaceChildren(el("option", { value: "" }, "Choose a processed document…"),
    ...docs.map(d => el("option", { value: d.id }, `${d.file_name}${d.title ? " — " + d.title : ""} (${d.page_count} p.)`)));
  if (cur) sel.value = cur;
  const m = $("#rv-method"); if (!m.options.length) for (const [k, v] of Object.entries(state.methods)) m.append(el("option", { value: k }, v));
}
async function openReview(docId, page) { showTab("review"); await loadReviewDocs(); $("#rv-doc").value = docId; await selectDoc(page); }
async function selectDoc(page = 1) {
  const id = Number($("#rv-doc").value); if (!id) return;
  state.review.doc = id; state.review.pages = await api(`/api/docs/${id}/processed-pages`);
  await showPage(page || state.review.pages[0] || 1);
}
$("#rv-doc").addEventListener("change", () => selectDoc(null));
$("#rv-page").addEventListener("change", () => showPage(Number($("#rv-page").value)));
$("#rv-prev").addEventListener("click", () => step(-1));
$("#rv-next").addEventListener("click", () => step(1));
function step(d) {
  const r = state.review; let p = r.page + d;
  if ($("#rv-only-processed").checked && r.pages.length) {
    const idx = r.pages.indexOf(r.page);
    p = d > 0 ? r.pages.find(x => x > r.page) ?? r.page : [...r.pages].reverse().find(x => x < r.page) ?? r.page;
    if (idx === -1 && d > 0) p = r.pages.find(x => x > r.page) ?? r.page;
  }
  if (p >= 1 && (!r.count || p <= r.count)) showPage(p);
}
async function showPage(p) {
  const r = state.review; if (!r.doc) return; r.page = p; $("#rv-page").value = p;
  let v; try { v = await api(`/api/docs/${r.doc}/pages/${p}`); } catch (e) { return toast(e.message); }
  r.count = v.page_count; r.view = v; $("#rv-of").textContent = `of ${v.page_count}`;
  const img = $("#rv-img"); $("#rv-img-msg").textContent = "Rendering page…";
  img.onload = () => $("#rv-img-msg").textContent = "";
  img.onerror = async () => { try { await api(img.src); } catch (e) { $("#rv-img-msg").textContent = "Page image unavailable: " + e.message; } };
  img.src = `/api/docs/${r.doc}/pages/${p}/image`;

  const route = $("#rv-route");
  if (!v.route) route.replaceChildren(el("div", { class: "box muted" }, "This page has not been processed."));
  else route.replaceChildren(el("div", { class: "box" },
    el("h3", {}, "Routing decision"),
    el("div", { class: "small" }, `Automatic: ${v.route.method} — ${v.route.reason}`,
      v.override ? el("strong", {}, ` · Your override: ${v.override}`) : ""),
    v.route.flags.length ? el("div", { class: "flag" }, "Flags: " + v.route.flags.join("; ")) : null,
    el("details", {}, el("summary", { class: "small" }, "Evidence"),
      el("div", { class: "kv" }, Object.entries(v.route.evidence).flatMap(([k, x]) => [el("span", { class: "muted" }, k), el("span", {}, String(x))])))));

  $("#rv-results").replaceChildren(...v.results.map(res => el("div", { class: "box" },
    el("h3", {}, `${state.methods[res.engine] || res.engine} `, res.engine === v.selected_method ? badge("selected") : null,
      res.current_settings ? null : el("span", { class: "muted small" }, " (older settings/version)")),
    el("div", { class: "small" }, `status ${res.status} · ${res.chars ?? 0} characters`,
      res.stats?.mean_conf != null ? ` · Tesseract mean confidence ${res.stats.mean_conf} (a hint, not proof)` : "",
      res.seconds != null ? ` · ${res.seconds}s` : ""),
    res.flags.length ? el("div", { class: "flag" }, "Automatic checks: " + res.flags.join(", ")) : el("div", { class: "muted small" }, "Automatic checks: none raised (not a verification)"),
    (res.warnings || []).length ? el("div", { class: "flag" }, res.warnings.join("; ")) : null,
    res.error ? el("div", { class: "flag" }, "Error: " + res.error) : null,
    res.reading_order ? el("div", { class: "muted small" }, "Reading order: " + res.reading_order) : null,
    res.text != null ? el("pre", { class: "text" }, res.text || "(no text)") : null)));

  const rv = v.review; $("#rv-review-state").textContent = rv ? `Current: ${rv.status} (revision ${v.review_history.length})` : "Current: unreviewed";
  document.querySelectorAll('input[name="rv-status"]').forEach(x => x.checked = rv ? x.value === rv.status : false);
  const sel = v.results.find(x => x.engine === v.selected_method);
  $("#rv-corrected").value = rv?.corrected_text ?? (sel?.text || "");
  $("#rv-note").value = "";
  $("#rv-history").replaceChildren(...v.review_history.map(h => el("li", {},
    `${new Date(h.created * 1000).toLocaleString()} — ${h.status}${h.note ? " — " + h.note : ""}${h.corrected_text ? " (with correction)" : ""}`)));
}
// Downloads go through fetch (with the token) and are saved via a temporary object URL.
async function download(path, payload, fallbackName) {
  try {
    toast("Preparing download…");
    const r = await fetch(path, { method: "POST", headers: { "X-Pilot-Token": TOKEN, "Content-Type": "application/json" },
                                  body: JSON.stringify(payload) });
    if (!r.ok) { const b = await r.json().catch(() => ({})); throw new Error(b.detail || r.statusText); }
    const cd = r.headers.get("content-disposition") || "";
    const m = cd.match(/filename\*=UTF-8''([^;]+)/) || cd.match(/filename="([^"]+)"/);
    const name = m ? decodeURIComponent(m[1]) : fallbackName;
    const url = URL.createObjectURL(await r.blob());
    const a = el("a", { href: url, download: name }); document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    toast(`Downloaded ${name}`);
  } catch (e) { toast("Download failed: " + e.message, 8000); }
}
for (const fmt of ["txt", "csv", "json"]) $(`#dl-${fmt}`).addEventListener("click", () => {
  const r = state.review; if (!r.doc) return toast("Choose a processed document first.");
  download(`/api/docs/${r.doc}/download`, { format: fmt, pages: $("#dl-range").value.trim() || String(r.page) }, `pages.${fmt}`);
});
$("#rv-retry").addEventListener("click", async () => {
  const r = state.review; if (!r.doc) return;
  try { const res = await post(`/api/docs/${r.doc}/pages/${r.page}/retry`, { method: $("#rv-method").value, pages: $("#rv-range").value.trim() });
        toast(`Retry queued as job ${res.job_id}.`); } catch (e) { toast("Retry failed: " + e.message, 8000); } });
$("#rv-save").addEventListener("click", async () => {
  const r = state.review; const status = document.querySelector('input[name="rv-status"]:checked')?.value;
  if (!r.doc || !status) return toast("Choose a review status.");
  const sel = r.view?.results.find(x => x.engine === r.view.selected_method);
  try { await post(`/api/docs/${r.doc}/pages/${r.page}/review`, { status, result_id: sel?.id ?? null,
          corrected_text: status === "corrected" ? $("#rv-corrected").value : null, note: $("#rv-note").value || null });
        toast("Review saved as a new revision."); showPage(r.page); } catch (e) { toast("Save failed: " + e.message, 8000); } });

$("#shutdown").addEventListener("click", async () => {
  if (!confirmInline("Shut down the pilot app?")) return;
  try { const r = await post("/api/shutdown"); toast(r.message, 10000); } catch (e) { toast(e.message); } });

// ------------------------------------------------------------------ start
refreshState().then(loadDocs).catch(e => toast(e.message));
setInterval(refreshState, 3000);
setInterval(() => { if (!$("#tab-jobs").classList.contains("hidden")) loadJobs().catch(() => {}); }, 2000);
