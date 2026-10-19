"use strict";
// Plain JS, no external libraries. All document text is inserted with textContent (never as HTML).

const TOKEN = document.querySelector('meta[name="pilot-token"]').content;
const $ = (s, el = document) => el.querySelector(s);
const state = { docs: [], sort: "file_name", asc: true, selected: new Map(), modes: [], methods: {},
                review: { doc: null, page: 1, count: 0, pages: [] } };

async function api(path, opts = {}) {
  const o = { ...opts, headers: { ...(opts.headers || {}) } };
  if (o.method && o.method !== "GET") { o.headers["X-Pilot-Token"] = TOKEN; o.headers["Content-Type"] = "application/json"; }
  const r = await fetch(path, o);
  const ct = r.headers.get("content-type") || "";
  const body = ct.includes("json") ? await r.json() : await r.text();
  if (!r.ok) throw new Error((body && body.detail) || body || r.statusText);
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
    const st = $("#status"); st.replaceChildren(
      el("span", { class: s.worker_alive ? "ok" : "bad" }, s.worker_alive ? "● worker running" : "● worker not responding"),
      " · ", activityText(s.worker_activity), " · ",
      el("span", { class: s.source_available ? "ok" : "bad" }, s.source_available ? "source folder reachable" : "source folder NOT reachable"),
      " · ", el("span", { class: s.disk_free_gb < s.min_free_gb ? "bad" : "" }, `${s.disk_free_gb} GB free`));
    if (!$("#modes").dataset.built) buildModes();
  } catch (e) { $("#status").textContent = "App not responding: " + e.message; }
}
function activityText(a) {
  if (!a || !a.state || a.state === "idle") return "idle";
  if (a.state === "hashing") return `job ${a.job}: checking source identity`;
  return `job ${a.job}: document ${a.doc}, page ${a.page}`;
}

function buildModes() {
  const fs = $("#modes"); fs.dataset.built = "1";
  state.modes.forEach((m, i) => fs.append(el("label", { class: m.implemented ? "" : "disabled", title: m.help },
    el("input", { type: "radio", name: "mode", value: m.id, disabled: !m.implemented, checked: i === 0 && m.implemented }),
    " ", m.label, m.implemented ? "" : " — not yet implemented",
    el("div", { class: "muted small" }, m.help))));
}

// ------------------------------------------------------------------ catalogue
async function loadDocs() {
  state.docs = (await api("/api/documents")).map(d => ({ ...d, folder: d.rel_path.includes("/") ? d.rel_path.slice(0, d.rel_path.lastIndexOf("/")) : "" }));
  renderDocs(); loadInvReport();
}
async function loadInvReport() {
  const r = await api("/api/inventory/report");
  const parts = Object.entries(r.counts).map(([k, v]) => `${k}: ${v}`).join(", ");
  const t = r.titles || {};
  const titleNote = t.error ? ` · titles: ${t.error}` : ` · titles: ${r.with_title}/${r.documents}` +
    (t.duplicates?.length ? `, ${t.duplicates.length} duplicate path(s) left blank` : "");
  $("#inv-report").textContent = r.documents ? `${r.documents} inventory rows · ${r.pages ?? "?"} pages · ${(r.bytes / 1e9).toFixed(2)} GB (inventory figures) · ${parts}${titleNote}` : "Inventory not loaded.";
  const ul = $("#unlisted"); ul.replaceChildren(...r.unlisted.map(u => el("li", {}, u)));
  $("#unlisted-box").classList.toggle("hidden", !r.unlisted.length);
  if (!r.source_available && r.documents) $("#inv-report").append(el("strong", { class: "flag" }, " · source folder was not available at last reload"));
}
function shownDocs() {
  const q = $("#search").value.trim().toLowerCase(), f = $("#filter-status").value;
  const rows = state.docs.filter(d => (!q || d.rel_path.toLowerCase().includes(q) || (d.title || "").toLowerCase().includes(q)) && (!f || d.resolve_status === f));
  const k = state.sort, dir = state.asc ? 1 : -1;
  return rows.sort((a, b) => {
    const x = a[k], y = b[k];
    if (x == null && y == null) return 0; if (x == null) return 1; if (y == null) return -1;
    const nx = Number(x), ny = Number(y);
    return (!isNaN(nx) && !isNaN(ny) ? nx - ny : String(x).localeCompare(String(y))) * dir;
  });
}
function renderDocs() {
  const tb = $("#docs tbody"); tb.replaceChildren();
  for (const d of shownDocs()) {
    const ok = d.resolve_status === "OK";
    const cb = el("input", { type: "checkbox", disabled: !ok, checked: state.selected.has(d.id),
      title: ok ? "" : "Cannot process: " + (d.note || d.resolve_status) });
    const pages = el("input", { class: "pages", placeholder: "all", disabled: !ok, value: state.selected.get(d.id) || "" });
    cb.addEventListener("change", () => { cb.checked ? state.selected.set(d.id, pages.value) : state.selected.delete(d.id); tr.classList.toggle("sel", cb.checked); updateSel(); });
    pages.addEventListener("input", () => { if (pages.value && !cb.checked) { cb.checked = true; tr.classList.add("sel"); }
      if (cb.checked) state.selected.set(d.id, pages.value); updateSel(); });
    const tr = el("tr", { class: state.selected.has(d.id) ? "sel" : "" },
      el("td", {}, cb), el("td", {}, d.inv_row_id ?? ""), el("td", { title: d.note || "" }, d.file_name),
      el("td", {}, d.title || el("span", { class: "muted" }, "—")),
      el("td", {}, d.folder), el("td", { class: "num" }, d.inv_pages ?? ""),
      el("td", { class: "num" }, d.inv_size ? (d.inv_size / 1e6).toFixed(1) : ""),
      el("td", { title: d.note || "" }, badge(d.resolve_status), d.note ? el("div", { class: "muted small" }, d.note) : null),
      el("td", {}, d.inv_duplicate_group ? `${d.inv_duplicate_group} (${d.inv_copies ?? "?"} copies)` : ""),
      el("td", {}, d.last_status ? [badge(d.last_status), el("span", { class: "muted small" }, ` job ${d.last_job}`)] : ""),
      el("td", {}, pages));
    tb.append(tr);
  }
  updateSel();
}
function updateSel() {
  const n = state.selected.size;
  const pages = [...state.selected.entries()].reduce((t, [id, spec]) => {
    const d = state.docs.find(x => x.id === id); return t + (spec ? countSpec(spec) : (d?.inv_pages || 0)); }, 0);
  $("#sel-summary").textContent = n ? `${n} document(s), about ${pages} page(s) selected.` : "Nothing selected.";
  $("#start").disabled = !n;
}
function countSpec(spec) { let n = 0; for (const p of spec.split(",")) { const m = p.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/); if (m) n += (Number(m[2] || m[1]) - Number(m[1]) + 1); } return n; }

document.querySelectorAll("#docs th[data-sort]").forEach(th => th.addEventListener("click", () => {
  state.asc = state.sort === th.dataset.sort ? !state.asc : true; state.sort = th.dataset.sort; renderDocs(); }));
$("#search").addEventListener("input", renderDocs);
$("#filter-status").addEventListener("change", renderDocs);
$("#select-all").addEventListener("change", (e) => {
  for (const d of shownDocs()) if (d.resolve_status === "OK") e.target.checked ? state.selected.set(d.id, state.selected.get(d.id) || "") : state.selected.delete(d.id);
  renderDocs(); });
$("#reload-inv").addEventListener("click", async () => {
  try { toast("Reloading inventory…"); await post("/api/inventory/reload"); await loadDocs(); toast("Inventory reloaded."); }
  catch (e) { toast("Reload failed: " + e.message, 8000); } });
$("#start").addEventListener("click", async () => {
  const mode = document.querySelector('input[name="mode"]:checked')?.value;
  if (!mode) return toast("Choose an extraction mode.");
  const items = [...state.selected.entries()].map(([doc_id, pages]) => ({ doc_id, pages: pages || "all" }));
  const big = items.reduce((t, i) => t + (i.pages === "all" ? (state.docs.find(d => d.id === i.doc_id)?.inv_pages || 0) : countSpec(i.pages)), 0);
  if (big > 500 && !confirmInline(`This job covers about ${big} pages and may take a long time. Queue it?`)) return;
  try { const r = await post("/api/jobs", { items, mode }); toast(`Job ${r.job_id} queued.`);
        state.selected.clear(); renderDocs(); showTab("jobs"); }
  catch (e) { toast("Could not queue: " + e.message, 8000); } });
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
