/* Category Replicator - front end (vanilla JS, no build step). */
(function () {
  "use strict";

  function apiBase() {
    try { if (typeof getWebAppBackendUrl === "function") return getWebAppBackendUrl(""); } catch (e) {}
    return "";
  }
  const API = apiBase().replace(/\/$/, "");

  async function call(path, opts) {
    opts = opts || {};
    const headers = {};
    let body = opts.body;
    if (opts.json !== undefined) { headers["Content-Type"] = "application/json"; body = JSON.stringify(opts.json); }
    const res = await fetch(API + path, { method: opts.method || "GET", headers, body });
    let data = null;
    try { data = await res.json(); } catch (e) {}
    if (!res.ok || (data && data.success === false)) {
      throw new Error((data && (data.error || data.message)) || ("HTTP " + res.status));
    }
    return data;
  }

  const $ = (s, r) => (r || document).querySelector(s);
  function h(tag, attrs, kids) {
    const n = document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v == null || v === false) return;
      if (k === "class") n.className = v;
      else if (k === "text") n.textContent = v;
      else if (k === "html") n.innerHTML = v;
      else if (k === "value") n.value = v;
      else if (k === "checked") n.checked = !!v;
      else if (k.slice(0, 2) === "on") n.addEventListener(k.slice(2), v);
      else n.setAttribute(k, v === true ? "" : v);
    });
    (kids || []).forEach(c => { if (c != null && c !== false) n.appendChild(typeof c === "string" ? document.createTextNode(c) : c); });
    return n;
  }
  const clear = n => { while (n.firstChild) n.removeChild(n.firstChild); return n; };
  function toast(msg, err) {
    const t = $("#toast");
    t.textContent = msg; t.className = "toast" + (err ? " err" : "");
    clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.add("hidden"), err ? 7000 : 3000);
  }
  const fmtPct = v => v == null ? "—" : (v * 100).toFixed(v === 1 || v === 0 ? 0 : 1) + "%";
  const when = s => s ? s.replace("T", " ").replace("Z", "").slice(0, 16) : "";
  const main = () => $("#main");
  const KIND = { copy: "Copy", const: "Constant", text: "Text", datepart: "Date part", arith: "Maths", concat: "Join text",
    lookup: "Mapping lookup", cond: "IF / bands", learned_map: "Learned map", formula: "Excel formula" };
  const STATUS = { ok: ["ok", "OK"], approx: ["warn", "Approximate"], unresolved: ["bad", "Needs a rule"], manual: ["manual", "Manual"], unchecked: ["acc", "From formula"], excel: ["acc", "Excel calculates"] };
  const OPS = ["==", "!=", ">", ">=", "<", "<=", "in", "not in", "contains", "startswith", "endswith", "blank", "not blank"];

  function dl(url) { const a = document.createElement("a"); a.href = API + url; a.click(); }

  /* ================================================================ routing */
  let ROUTE_TOKEN = 0;
  function route() {
    const hash = location.hash.replace(/^#/, "") || "/";
    const tok = ++ROUTE_TOKEN;
    const m = hash.match(/^\/s\/([^/]+)(?:\/(\w+))?/);
    if (hash === "/new") return viewBuilder();
    if (m) return viewSystem(m[1], m[2] || "rules", tok);
    return viewHome();
  }
  window.addEventListener("hashchange", route);
  // Dataiku can inject a <base href>, which turns href="#/x" into a jump to the Dataiku page.
  // Handle in-app links ourselves: never follow the href, just switch the hash route.
  document.addEventListener("click", e => {
    const a = e.target.closest && e.target.closest("a[href^='#/']");
    if (!a) return;
    e.preventDefault();
    const target = a.getAttribute("href").slice(1);
    if (location.hash.replace(/^#/, "") === target) route(); else location.hash = target;
  });

  /* ================================================================== home */
  async function viewHome() {
    const root = clear(main());
    root.appendChild(h("div", { class: "head" }, [
      h("div", {}, [h("h1", { text: "Systems" }),
        h("div", { class: "muted", text: "Each system is a category: it learned how its output is built from its input + mapping files." })]),
      h("a", { class: "btn", href: "#/new", text: "+ New system" })]));
    const box = h("div", { class: "grid g3" }, [h("div", { class: "muted" }, [h("span", { class: "spin" }), "Loading…"])]);
    root.appendChild(box);
    try {
      const r = await call("/api/systems");
      clear(box);
      if (!r.systems.length) {
        box.className = "";
        box.appendChild(h("div", { class: "card empty" }, [h("h2", { text: "No systems yet" }),
          h("p", { text: "Upload an input file, its mapping file(s) and the output you expect. The app learns every derived column." }),
          h("a", { class: "btn", href: "#/new", text: "Create the first system" })]));
      }
      r.systems.forEach(s => box.appendChild(h("a", { class: "card", href: "#/s/" + s.id, style: "color:inherit;display:block;margin:0" }, [
        h("div", { class: "row" }, [h("h3", { text: s.name, style: "margin:0" }), h("span", { class: "spacer" }),
]),
        h("div", { class: "muted small", text: s.description || "" }),
        h("div", { class: "row", style: "margin-top:10px" }, [
          h("span", { class: "pill", text: s.n_sheets + " sheet(s)" }), h("span", { class: "pill ok", text: s.n_formulas + " formulas" }), h("span", { class: "pill", text: s.n_tables + " mappings" }),
          s.split ? h("span", { class: "pill acc", text: "multi-sheet" }) : null,
          s.n_review ? h("span", { class: "pill warn", text: s.n_review + " column(s) without formula" }) : null]),
        h("div", { class: "muted small", style: "margin-top:8px", text: "Updated " + when(s.updated || s.created) })])));
    } catch (e) { clear(box); box.appendChild(h("div", { class: "note err", text: e.message })); }
  }

  /* =============================================================== builder */
  const B = { draft: null, did: null, name: "", desc: "", mode: "auto", column: "", step: 1 };

  function resetBuilder() { B.draft = null; B.did = null; B.name = ""; B.desc = ""; B.mode = "auto"; B.column = ""; B.step = 1; }
  function viewBuilder() { if (B.step === 0) resetBuilder(); renderBuilder(); }

  function allHeaders() {
    const set = [];
    ((B.draft && B.draft.files) || []).forEach(f => f.sheets.forEach(s => { if (s.role !== "ignore") s.headers.forEach(x => { if (set.indexOf(x) < 0) set.push(x); }); }));
    return set;
  }

  async function ensureDraft() {
    if (!B.did) B.did = (await call("/api/drafts", { method: "POST" })).draft_id;
    return B.did;
  }

  // Uploads run one at a time (a second drop while the first is still parsing used to race), with progress.
  let UP_QUEUE = Promise.resolve();
  let UP_BUSY = 0;
  let UP_ERR = "";
  function uploadFiles(files, hint) {
    const list = Array.prototype.slice.call(files || []);
    if (!list.length) return UP_QUEUE;
    UP_BUSY += 1; UP_ERR = "";
    const names = list.map(f => f.name).join(", ");
    renderBuilder();
    UP_QUEUE = UP_QUEUE.then(async () => {
      try {
        const fd = new FormData();
        list.forEach(f => fd.append("files", f));
        fd.append("hint", hint);
        const did = await ensureDraft();
        B.draft = (await call("/api/drafts/" + did + "/files", { method: "POST", body: fd })).draft;
      } catch (e) {
        UP_ERR = names + ": " + e.message; toast(UP_ERR, true);
      } finally { UP_BUSY -= 1; renderBuilder(); }
    });
    return UP_QUEUE;
  }

  function dropZone(hint, title, sub) {
    const inp = h("input", { type: "file", multiple: true, accept: ".xlsx,.xlsm,.xls,.csv" });
    const z = h("label", { class: "drop" }, [h("div", { text: "Drop files or click to browse" }), h("div", { class: "small", text: sub }), inp]);
    inp.addEventListener("change", () => uploadFiles(inp.files, hint));
    z.addEventListener("dragover", e => { e.preventDefault(); z.classList.add("over"); });
    z.addEventListener("dragleave", () => z.classList.remove("over"));
    z.addEventListener("drop", e => { e.preventDefault(); z.classList.remove("over"); uploadFiles(e.dataTransfer.files, hint); });
    return h("div", {}, [h("div", { class: "zone-title" }, [h("span", { class: "dot " + hint }), h("h3", { text: title, style: "margin:0" })]), z]);
  }

  function fileBoxes(role) {
    const out = [];
    ((B.draft && B.draft.files) || []).forEach(f => {
      const sheets = f.sheets;
      const rows = sheets.map(s => {
        const sel = h("select", { class: "inl" }, [["input", "Input"], ["mapping", "Mapping"], ["output", "Expected output"], ["ignore", "Ignore"]]
          .map(([v, l]) => h("option", { value: v, text: l, ...(s.role === v ? { selected: true } : {}) })));
        sel.addEventListener("change", async () => {
          const roles = {}; roles[f.fid + "|" + s.sheet] = sel.value;
          try { B.draft = (await call("/api/drafts/" + B.did + "/roles", { method: "PUT", json: { roles } })).draft; renderBuilder(); }
          catch (e) { toast(e.message, true); }
        });
        return h("div", {}, [h("div", { class: "sheetrow" }, [h("span", { class: "nm", text: s.sheet }),
          h("span", { class: "muted small", text: s.rows + " rows × " + s.cols + " cols" + (s.n_formulas ? " · " + s.n_formulas + " formulas" : "") }),
          h("span", { class: "spacer" }), sel]),
          h("div", { class: "chips" }, s.headers.slice(0, 18).map(x => h("span", { class: "chip", text: x })).concat(s.headers.length > 18 ? [h("span", { class: "chip", text: "+" + (s.headers.length - 18) })] : []))]);
      });
      const issues = (f.issues || []).map(i => h("div", { class: "note", style: "margin:6px 12px", text: i.sheet + " — " + i.issue }));
      out.push(h("div", { class: "filebox" }, [h("div", { class: "fh" }, [h("span", { text: f.name }),
        h("button", { class: "icon-btn", title: "Remove file", text: "✕", onclick: async () => {
          B.draft = (await call("/api/drafts/" + B.did + "/files/" + f.fid, { method: "DELETE" })).draft; renderBuilder(); } })]), ...issues, ...rows]));
    });
    return out;
  }

  function renderBuilder() {
    const root = clear(main());
    const steps = ["1 · Setup", "2 · Files", "3 · Learn"];
    root.appendChild(h("div", { class: "head" }, [h("div", {}, [h("h1", { text: "New system" }),
      h("div", { class: "muted", text: "Show the app one worked example; it keeps its layout, filters and formulas as a template." })])]));
    root.appendChild(h("div", { class: "steps" }, steps.map((s, i) => h("span", { class: "step " + (B.step === i + 1 ? "on" : B.step > i + 1 ? "done" : ""), text: s }))));

    if (B.step === 1) {
      const name = h("input", { type: "text", value: B.name, placeholder: "e.g. Credit RWA – Retail", oninput: e => { B.name = e.target.value; } });
      const desc = h("input", { type: "text", value: B.desc, placeholder: "optional", oninput: e => { B.desc = e.target.value; } });
      const modes = [["none", "One output sheet", "Everything goes to a single sheet."],
        ["auto", "Several outputs — detect from my example", "If your example output has multiple sheets (e.g. one per Region), the app finds the column that splits them."],
        ["column", "Several outputs — split by a column I choose", "One sheet per value of a column. You pick the column after uploading the files."],
        ["custom", "Several outputs — custom filters per sheet", "Define each sheet's filter yourself after learning (e.g. Amount ≥ 1000)."]];
      root.appendChild(h("div", { class: "card" }, [h("label", { class: "l", text: "System / category name" }), name,
        h("label", { class: "l", text: "Description" }), desc,
        h("h2", { text: "Outputs", style: "margin-top:20px" }),
        ...modes.map(([v, t, d]) => h("label", { class: "radio" + (B.mode === v ? " on" : "") }, [
          h("input", { type: "radio", name: "mode", ...(B.mode === v ? { checked: true } : {}), onchange: () => { B.mode = v; renderBuilder(); } }),
          h("b", { text: t }), h("div", { class: "muted small", text: d, style: "margin-left:22px" })])),
        h("div", { class: "row", style: "margin-top:14px" }, [h("span", { class: "spacer" }),
          h("button", { class: "btn", text: "Next: upload files →", onclick: () => { if (!B.name.trim()) return toast("Give the system a name", true); B.step = 2; renderBuilder(); } })])]));
      return;
    }

    if (B.step === 2) {
      root.appendChild(h("div", { class: "grid g3" }, [
        h("div", { class: "card" }, [dropZone("input", "Input file(s)", "The raw data the output is built from")]),
        h("div", { class: "card" }, [dropZone("mapping", "Mapping file(s)", "Lookup / rate / reference tables")]),
        h("div", { class: "card" }, [dropZone("output", "Expected output", "The finished file: several sheets allowed, including the mapping sheets its formulas use (detected automatically)")])]));
      const boxes = fileBoxes();
      if (UP_BUSY) root.appendChild(h("div", { class: "note good" }, [h("span", { class: "spin" }), "Uploading and reading " + UP_BUSY + " file group(s)… large workbooks can take a little while. You can keep adding files; they are processed in order."]));
      if (UP_ERR) root.appendChild(h("div", { class: "note err", text: "Upload failed — " + UP_ERR + " (drop the file again to retry)" }));
      root.appendChild(h("div", { class: "card" }, [h("h2", { text: "Uploaded sheets" }),
        h("div", { class: "muted small", text: "Every sheet is classified automatically — change a sheet's role if the guess is wrong. Several input sheets with the same columns are stacked; a workbook can hold input and mapping sheets together." }),
        ...(boxes.length ? boxes : [h("div", { class: "empty", text: "Nothing uploaded yet" })])]));
      const roles = { input: 0, mapping: 0, output: 0 };
      ((B.draft && B.draft.files) || []).forEach(f => f.sheets.forEach(s => { if (roles[s.role] != null) roles[s.role]++; }));
      const ready = roles.input > 0 && roles.output > 0;
      const col = B.mode === "column" ? h("div", { class: "card" }, [h("h3", { text: "Split the output by which column?" }),
        h("select", { onchange: e => { B.column = e.target.value; } }, [h("option", { value: "", text: "— choose a column —" }),
          ...allHeaders().map(x => h("option", { value: x, text: x, ...(B.column === x ? { selected: true } : {}) }))])]) : null;
      if (col) root.appendChild(col);
      root.appendChild(h("div", { class: "row" }, [h("button", { class: "btn ghost", text: "← Back", onclick: () => { B.step = 1; renderBuilder(); } }), h("span", { class: "spacer" }),
        h("span", { class: "muted small", text: roles.input + " input · " + roles.mapping + " mapping · " + roles.output + " output sheet(s)" }),
        h("button", { class: "btn", text: "Build the template →", ...(ready && !UP_BUSY ? {} : { disabled: true }), onclick: learnNow })]));
      return;
    }

    root.appendChild(h("div", { class: "card empty" }, [h("span", { class: "spin" }), "Reading your example and building the template…"]));
  }

  async function learnNow() {
    if (B.mode === "column" && !B.column) return toast("Choose the split column", true);
    B.step = 3; renderBuilder();
    try {
      const split = B.mode === "column" ? { mode: "column", column: B.column } : { mode: B.mode };
      const r = await call("/api/drafts/" + B.did + "/learn", { method: "POST", json: { name: B.name, description: B.desc, split } });
      const id = r.system.id;
      B.step = 0; resetBuilder();
      location.hash = "#/s/" + id;
    } catch (e) { B.step = 2; renderBuilder(); toast(e.message, true); }
  }

  /* =============================================================== system */
  let SYS = null;
  let LAST_RUN = null;
  let REPLAY = null;

  async function viewSystem(sid, tab, tok) {
    const root = clear(main());
    root.appendChild(h("div", { class: "muted" }, [h("span", { class: "spin" }), "Loading…"]));
    try {
      if (!SYS || SYS.id !== sid) { SYS = (await call("/api/systems/" + sid)).system; LAST_RUN = null; REPLAY = null; }
      else SYS = (await call("/api/systems/" + sid)).system;
    } catch (e) { clear(root).appendChild(h("div", { class: "note err", text: e.message })); return; }
    if (tok !== ROUTE_TOKEN) return;
    renderSystem(tab);
  }

  function renderSystem(tab) {
    const s = SYS, root = clear(main());
    const nF = s.sheets.reduce((a, d) => a + d.columns.filter(c => c.kind === "formula").length, 0);
    const rv = s.sheets.reduce((a, d) => a + d.columns.filter(c => c.kind === "empty").length, 0);
    root.appendChild(h("div", { class: "head" }, [h("div", {}, [h("h1", { text: s.name }), h("div", { class: "muted", text: s.description || "" }),
      h("div", { class: "row small", style: "margin-top:6px" }, [
        h("span", { class: "pill", text: s.sheets.length + " output sheet(s)" }), h("span", { class: "pill ok", text: nF + " formulas stored" }),
        h("span", { class: "pill", text: Object.keys(s.tables).length + " mapping / toggle sheets" }),
        rv ? h("span", { class: "pill warn", text: rv + " column(s) without a formula in the sample" }) : null])]),
      h("div", { class: "row" }, [
        h("button", { class: "btn ghost small", text: "Input template", onclick: () => dl("/api/systems/" + s.id + "/template") }),
        h("button", { class: "btn ghost small", text: "Rebuild", onclick: relearn }),
        h("button", { class: "btn ghost danger small", text: "Delete", onclick: async () => { if (!confirm("Delete system '" + s.name + "'?")) return;
          await call("/api/systems/" + s.id, { method: "DELETE" }); SYS = null; location.hash = "#/"; } })])]));
    const tabs = [["rules", "Template"], ["formulas", "Formulas"], ["run", "Run on new data"], ["filters", "Filters & outputs"], ["maps", "Mappings"], ["history", "History"]];
    root.appendChild(h("div", { class: "tabs" }, tabs.map(([k, l]) => h("a", { class: "tab" + (tab === k ? " on" : ""), href: "#/s/" + s.id + "/" + k, text: l }))));
    const body = h("div", {});
    root.appendChild(body);
    ({ rules: tabRules, formulas: tabFormulas, run: tabRun, filters: tabFilters, maps: tabMaps, history: tabHistory }[tab] || tabRules)(body);
  }

  async function relearn() {
    if (!confirm("Rebuild the template from the stored example files? Filter edits will be lost.")) return;
    const mode = SYS.filters.split.enabled ? "auto" : "auto";
    toast("Re-learning…");
    try { SYS = (await call("/api/systems/" + SYS.id + "/relearn", { method: "POST", json: { split: { mode } } })).system; renderSystem("rules"); toast("Done"); }
    catch (e) { toast(e.message, true); }
  }

  async function saveSystem(patch, msg) {
    SYS = (await call("/api/systems/" + SYS.id, { method: "PUT", json: patch })).system;
    toast(msg || "Saved");
  }

  /* ------------------------------------------------------------- template tab */
  function tabRules(body) {
    const s = SYS, sp = s.filters.split;
    body.appendChild(h("div", { class: "card" }, [h("h2", { text: "The template" }),
      h("div", { class: "muted small", text: "Each output sheet is laid out exactly like your sample. Input columns are pasted from the input rows that pass the sheet's filter; derived columns are the sample's formulas, stored as written and filled down. Nothing is calculated by the app." }),
      ...(s.warnings || []).map(w => h("div", { class: "note", text: w }))]));
    s.sheets.forEach(d => {
      const cond = sp.enabled ? ((sp.sheets || []).find(x => x.name === d.name) || {}).conditions : null;
      const where = sp.enabled && sp.mode === "column" ? "rows where " + sp.column + " = " + ((Object.entries(sp.map || {}).find(([k, v]) => v === d.name) || [d.name])[0])
        : (cond && cond.length ? "rows where " + cond.map(c => c.col + " " + c.op + " " + (Array.isArray(c.value) ? c.value.join(" | ") : c.value)).join(" and ") : "all rows");
      const col = n => { let t = ""; n += 1; while (n > 0) { const m = (n - 1) % 26; t = String.fromCharCode(65 + m) + t; n = Math.floor((n - 1) / 26); } return t; };
      body.appendChild(h("div", { class: "card" }, [
        h("div", { class: "row" }, [h("h2", { text: d.name, style: "margin:0" }), h("span", { class: "pill acc", text: where }),
          h("span", { class: "muted small", text: "header at row " + (d.origin.row + 1) + ", starts in column " + col(d.origin.col) })]),
        h("div", { class: "scroll", style: "margin-top:10px;max-height:340px" }, [h("table", { class: "t" }, [
          h("thead", {}, [h("tr", {}, ["Cell", "Column", "What goes there"].map(x => h("th", { text: x })))]),
          h("tbody", {}, d.columns.map(c => h("tr", {}, [
            h("td", { class: "mono", text: col(c.col) + (d.origin.row + 1) }), h("td", {}, [h("b", { text: c.name })]),
            c.kind === "input" ? h("td", {}, [h("span", { class: "pill acc", text: "input" }), " pasted from the input column ", h("b", { text: c.source })])
              : c.kind === "formula" ? h("td", {}, [h("span", { class: "pill ok", text: "formula" }), c.orig !== undefined ? h("span", { class: "pill manual", text: "edited", style: "margin-left:4px" }) : null, " ", h("span", { class: "desc", text: c.raw })])
              : h("td", {}, [h("span", { class: "pill warn", text: "empty" }), " no formula in the sample"])])))])]),
        (d.params || []).length ? h("div", { class: "muted small", style: "margin-top:8px", text: "Parameter cell(s) the formulas point at: " + d.params.map(p => p.cell + " = " + p.value).join(", ") }) : null]));
    });
    body.appendChild(h("div", { class: "card" }, [h("h3", { text: "Also written into every output workbook" }),
      h("div", { class: "row small" }, [...Object.keys(s.tables).map(t => h("span", { class: "pill warn", text: t + ((s.toggles || []).some(g => g.sheet === t) ? " (toggle)" : " (mapping)") })),
        h("span", { class: "pill", text: "Run summary" })])]));
  }

  /* ------------------------------------------------------------- formulas tab */
  // The same fill-down rule the server uses (row numbers without a $ move; $ refs, ranges' $ parts, sheet names and text stay).
  function shiftF(f, from, to) {
    const d = to - from; if (!d || !f) return f;
    const parts = f.split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/);
    for (let i = 0; i < parts.length; i += 2) {
      parts[i] = parts[i].replace(/(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_!])/g,
        (m, a, b, c, e) => (c === "$" || +e + d < 1) ? m : a + b + c + (+e + d));
    }
    return parts.join("");
  }
  const colLetter = n => { let t = ""; n += 1; while (n > 0) { const m = (n - 1) % 26; t = String.fromCharCode(65 + m) + t; n = Math.floor((n - 1) / 26); } return t; };

  function tabFormulas(body) {
    const s = SYS;
    const edits = {};                                   // "sheet\u0001column" -> text typed (only when different from saved)
    const K = (sh, c) => sh + "\u0001" + c;
    const rows = [];                                    // every derived (formula / empty) cell of the template
    s.sheets.forEach(d => d.columns.forEach(c => { if (c.kind !== "input") rows.push({ d, c, key: K(d.name, c.name), inp: null, prev: null, badge: null, undo: null }); }));
    const saved = r => r.c.raw || "";
    const current = r => (r.key in edits) ? edits[r.key] : saved(r);
    const rowOf = r => r.c.sample_row || (r.d.origin.row + 2);
    const norm = r => saved(r) ? shiftF(saved(r), rowOf(r), 1) : null;
    const same = {}; rows.forEach(r => { const n = norm(r); if (n) same[n] = (same[n] || 0) + 1; });

    const pendingLbl = h("span", { class: "muted small" });
    const saveBtn = h("button", { class: "btn", text: "Save formula changes", disabled: true, onclick: save });
    const discardBtn = h("button", { class: "btn ghost", text: "Discard", disabled: true, onclick: () => { Object.keys(edits).forEach(k => delete edits[k]); rows.forEach(refresh); update(); } });

    function refresh(r) {
      const t = current(r), changed = (r.key in edits);
      r.inp.value = t;
      r.inp.style.borderColor = changed ? "var(--warn)" : "";
      const nxt = t ? shiftF(t, rowOf(r), rowOf(r) + 1) : "";
      r.prev.textContent = t ? nxt : "—";
      r.badge.textContent = changed ? "unsaved" : (r.c.orig !== undefined ? "edited" : (saved(r) ? "" : "empty"));
      r.badge.className = "pill " + (changed ? "warn" : (r.c.orig !== undefined ? "manual" : (saved(r) ? "" : "bad")));
      r.badge.style.display = r.badge.textContent ? "" : "none";
    }
    function setText(r, t) {
      if (t === saved(r)) delete edits[r.key]; else edits[r.key] = t;
      refresh(r); update();
    }
    function update() {
      const n = Object.keys(edits).length;
      saveBtn.disabled = !n; discardBtn.disabled = !n;
      pendingLbl.textContent = n ? n + " unsaved change(s)" : "No unsaved changes";
    }
    async function save() {
      const payload = Object.keys(edits).map(k => { const [sheet, column] = k.split("\u0001"); return { sheet, column, raw: edits[k] }; });
      saveBtn.disabled = true;
      try {
        SYS = (await call("/api/systems/" + SYS.id, { method: "PUT", json: { formulas: payload } })).system;
        toast(payload.length + " formula(s) saved"); renderSystem("formulas");
      } catch (e) { toast(e.message, true); saveBtn.disabled = false; err.textContent = e.message; err.style.display = ""; }
    }

    const err = h("div", { class: "note err", style: "display:none" });
    const fIn = h("input", { type: "text", class: "inl", placeholder: "find, e.g. $B$3", style: "min-width:200px" });
    const rIn = h("input", { type: "text", class: "inl", placeholder: "replace with, e.g. $B$9", style: "min-width:200px" });
    const scope = h("select", { class: "inl" }, [h("option", { value: "", text: "all sheets" }), ...s.sheets.map(d => h("option", { value: d.name, text: d.name }))]);
    const hits = h("span", { class: "muted small" });
    function doReplace() {
      const f = fIn.value; if (!f) return toast("Type the text to find", true);
      let n = 0;
      rows.forEach(r => { if (scope.value && r.d.name !== scope.value) return; const t = current(r);
        if (t && t.indexOf(f) >= 0) { setText(r, t.split(f).join(rIn.value)); n++; } });
      hits.textContent = n ? n + " formula(s) changed — review, then save" : "no formula contains that text";
    }

    body.appendChild(h("div", { class: "card" }, [
      h("div", { class: "row" }, [h("h2", { text: "Formulas", style: "margin:0" }), h("span", { class: "spacer" }), pendingLbl, discardBtn, saveBtn]),
      h("div", { class: "muted small", text: "Every derived cell of the template: its sheet, address and formula. Edit the text — fix a reference, point at another mapping range, change the logic — and save. Type each formula as it would read on the FIRST data row; every other row is the same text moved down (fill-down). Clear the box to leave the column empty; fill an empty one to add a formula." }),
      h("div", { class: "row", style: "margin-top:10px" }, [h("b", { class: "small", text: "Find & replace" }), fIn, rIn, scope,
        h("button", { class: "btn ghost small", text: "Replace in formulas", onclick: doReplace }), hits]),
      err]));

    s.sheets.forEach(d => {
      const mine = rows.filter(r => r.d === d);
      if (!mine.length) { body.appendChild(h("div", { class: "card" }, [h("b", { text: d.name }), h("span", { class: "muted small", text: "  — every column is an input column; no formulas on this sheet." })])); return; }
      body.appendChild(h("div", { class: "card" }, [
        h("div", { class: "row" }, [h("h3", { text: d.name, style: "margin:0" }), h("span", { class: "muted small", text: mine.length + " derived column(s)" })]),
        h("div", { class: "scroll", style: "margin-top:8px;max-height:none" }, [h("table", { class: "t" }, [
          h("thead", {}, [h("tr", {}, ["Cell", "Column", "Formula (as on the first data row)", "Next row reads", ""].map(x => h("th", { text: x })))]),
          h("tbody", {}, mine.map(r => {
            r.inp = h("input", { type: "text", class: "mono", spellcheck: "false", placeholder: "= no formula in the sample — type one to add it", style: "width:100%;font-size:12.5px",
              oninput: e => setText(r, e.target.value) });
            r.prev = h("span", { class: "desc muted" });
            r.badge = h("span", { class: "pill" });
            r.undo = h("button", { class: "icon-btn", title: "Back to the formula from your sample", text: "↺", style: "display:" + (r.c.orig !== undefined ? "" : "none"),
              onclick: () => setText(r, r.c.orig) });
            const n = same[norm(r)];
            const tr = h("tr", {}, [h("td", { class: "mono", text: colLetter(r.c.col) + rowOf(r), style: "white-space:nowrap" }),
              h("td", {}, [h("b", { text: r.c.name }), n > 1 ? h("div", { class: "muted small", text: "same formula in " + (n - 1) + " other place(s)" }) : null]),
              h("td", { style: "min-width:340px" }, [r.inp]), h("td", {}, [r.prev]), h("td", { style: "white-space:nowrap" }, [r.badge, " ", r.undo])]);
            refresh(r);
            return tr;
          }))])])]));
    });
    update();
  }

  /* ---------------------------------------------------------------- run tab */
  function tabRun(body) {
    const s = SYS;
    const files = { in: [], map: [] }; const tog = {};
    const togBox = (s.toggles || []).length ? h("div", { class: "card", style: "margin:12px 0 0;box-shadow:none;background:var(--surface-2)" }, [
      h("h3", { text: "Toggles used by the formulas" }),
      h("div", { class: "muted small", text: "Taken from the Toggle sheet of your mapping (attach an updated mapping to change them for good) — or override for this run only:" }),
      ...(s.toggles || []).map(t => h("div", { class: "cond" }, [h("b", { text: t.name }),
        h("select", { class: "inl", onchange: e => { tog[t.name] = e.target.value; } },
          ["Yes", "No"].map(v => h("option", { value: v, text: v, ...(t.value === v ? { selected: true } : {}) }))),
        h("span", { class: "muted small", text: t.sheet ? "(" + t.sheet + "!" + t.cell + ")" : "" })]))]) : null;
    function zone(key, title, sub, multiple) {
      const list = h("div", { class: "small muted", style: "margin-top:6px" });
      const refresh = () => { list.textContent = files[key].map(f => f.name).join(", "); };
      const inp = h("input", { type: "file", ...(multiple ? { multiple: true } : {}), accept: ".xlsx,.xlsm,.xls,.csv" });
      const set = fl => { files[key] = Array.prototype.slice.call(fl); refresh(); };
      inp.addEventListener("change", () => set(inp.files));
      const z = h("label", { class: "drop" }, [h("div", { text: sub }), inp]);
      z.addEventListener("dragover", e => { e.preventDefault(); z.classList.add("over"); });
      z.addEventListener("dragleave", () => z.classList.remove("over"));
      z.addEventListener("drop", e => { e.preventDefault(); z.classList.remove("over"); set(e.dataTransfer.files); });
      return h("div", {}, [h("h3", { text: title }), z, list]);
    }
    const out = h("div", {});
    const btn = h("button", { class: "btn", text: "Run", onclick: async () => {
      if (!files.in.length) return toast("Add at least one input file", true);
      const fd = new FormData();
      files.in.forEach(f => fd.append("files", f)); files.map.forEach(f => fd.append("mappings", f));
      if (Object.keys(tog).length) fd.append("toggles", JSON.stringify(tog));
      btn.disabled = true; btn.textContent = "Running…";
      try { LAST_RUN = (await call("/api/systems/" + s.id + "/run", { method: "POST", body: fd })).run; showRun(out, LAST_RUN); }
      catch (e) { toast(e.message, true); }
      btn.disabled = false; btn.textContent = "Run";
    } });
    body.appendChild(h("div", { class: "card" }, [h("h2", { text: "Run on new data" }),
      h("div", { class: "muted small", text: "Upload new input in the same layout (" + s.input.required.slice(0, 6).join(", ") + (s.input.required.length > 6 ? "…" : "") + "). The system builds every derived column" + (s.filters.split.enabled ? " and splits the rows into the sample's sheets" : "") + "." }),
      h("div", { class: "grid g2", style: "margin-top:12px" }, [zone("in", "Input file(s)", "Drop new input here", true),
        zone("map", "Updated mappings / toggles (optional)", "Replace the stored mapping or Toggle sheets for this run", true)]),
      togBox,
      h("div", { class: "row", style: "margin-top:14px" }, [
        h("span", { class: "muted small", text: "The output holds your sample's formulas; Excel calculates them when the file is opened." }),
        h("span", { class: "spacer" }), btn])]));
    body.appendChild(out);
    if (LAST_RUN && LAST_RUN.system === s.id) showRun(out, LAST_RUN);
  }

  function showRun(box, run) {
    clear(box);
    const card = h("div", { class: "card" }, [h("div", { class: "row" }, [h("h2", { text: "Result", style: "margin:0" }), h("span", { class: "spacer" }),
      h("button", { class: "btn", text: "Download .xlsx", onclick: () => dl("/api/systems/" + run.system + "/runs/" + run.run_id + "/download") })]),
      h("div", { class: "stat", style: "margin-top:12px" }, [h("div", {}, [h("b", { text: run.input_rows }), h("span", { text: "input rows" })]),
        h("div", {}, [h("b", { text: run.rows }), h("span", { text: "output rows" })]),
        h("div", {}, [h("b", { text: Object.keys(run.sheets).length }), h("span", { text: "sheet(s)" })]),
])]);
    if ((run.toggles || []).length) card.appendChild(h("div", { class: "note good", text: "Toggles used: " + run.toggles.map(t => t.name + " = " + t.value).join(", ") }));
    (run.notes || []).forEach(n => card.appendChild(h("div", { class: "note", text: n })));
    Object.keys(run.rule_errors || {}).forEach(k => card.appendChild(h("div", { class: "note err", text: k + ": " + run.rule_errors[k] })));
    const names = Object.keys(run.preview);
    const holder = h("div", {});
    const tabs = h("div", { class: "tabs", style: "margin-top:14px" });
    function show(n) {
      Array.prototype.forEach.call(tabs.children, c => c.classList.toggle("on", c.dataset.n === n));
      const p = run.preview[n];
      clear(holder).appendChild(h("div", { class: "scroll" }, [h("table", { class: "t" }, [
        h("thead", {}, [h("tr", {}, p.headers.map(x => h("th", { text: x })))]),
        h("tbody", {}, p.data.map(r => h("tr", {}, r.map(v => h("td", { text: v == null ? "" : String(v).replace(/T00:00:00$/, "") })))))])]));
      holder.appendChild(h("div", { class: "muted small", text: "Showing " + p.data.length + " of " + p.rows + " rows" }));
    }
    names.forEach(n => tabs.appendChild(h("a", { class: "tab", "data-n": n, text: n + " (" + run.preview[n].rows + ")", onclick: () => show(n) })));
    card.appendChild(tabs); card.appendChild(holder);
    box.appendChild(card);
    if (names.length) show(names[0]);
  }

  /* ------------------------------------------------------------ filters tab */
  function tabFilters(body) {
    const F = JSON.parse(JSON.stringify(SYS.filters));
    const cols = SYS.input.columns;
    function condEditor(list, rerender) {
      const wrap = h("div", {});
      list.forEach((c, i) => {
        const col = h("select", { class: "inl", onchange: e => { c.col = e.target.value; } }, cols.map(x => h("option", { value: x, text: x, ...(x === c.col ? { selected: true } : {}) })));
        const op = h("select", { class: "inl", onchange: e => { c.op = e.target.value; rerender(); } }, OPS.map(x => h("option", { value: x, text: x, ...(x === c.op ? { selected: true } : {}) })));
        const noval = c.op === "blank" || c.op === "not blank";
        const val = h("input", { type: "text", class: "inl", value: Array.isArray(c.value) ? c.value.join(",") : (c.value == null ? "" : c.value), placeholder: "value", ...(noval ? { style: "display:none" } : {}),
          oninput: e => { c.value = e.target.value; } });
        wrap.appendChild(h("div", { class: "cond" }, [col, op, val, h("button", { class: "icon-btn", text: "✕", onclick: () => { list.splice(i, 1); rerender(); } })]));
      });
      wrap.appendChild(h("button", { class: "btn ghost small", text: "+ condition", onclick: () => { list.push({ col: cols[0], op: "==", value: "" }); rerender(); } }));
      return wrap;
    }
    function draw() {
      clear(body);
      body.appendChild(h("div", { class: "card" }, [h("h2", { text: "Input filter" }),
        h("div", { class: "muted small", text: "Rows must satisfy ALL of these to be processed (applied before any column is built)." }),
        condEditor(F.pre, draw)]));
      const sp = F.split;
      const card = h("div", { class: "card" }, [h("div", { class: "row" }, [h("h2", { text: "Split into several output sheets", style: "margin:0" }), h("span", { class: "spacer" }),
        h("label", { class: "small" }, [h("input", { type: "checkbox", ...(sp.enabled ? { checked: true } : {}), onchange: e => { sp.enabled = e.target.checked; draw(); } }), " enabled"])])]);
      if (sp.enabled) {
        const mode = h("select", { class: "inl", onchange: e => { sp.mode = e.target.value; if (sp.mode === "custom" && !(sp.sheets || []).length) sp.sheets = [{ name: "Sheet 1", match: "all", conditions: [] }]; draw(); } },
          [["column", "One sheet per value of a column"], ["custom", "Custom filters per sheet"]].map(([v, l]) => h("option", { value: v, text: l, ...(sp.mode === v ? { selected: true } : {}) })));
        card.appendChild(h("div", { class: "row", style: "margin-top:10px" }, [h("span", { text: "Mode" }), mode]));
        if (sp.mode === "column") {
          card.appendChild(h("div", { class: "row", style: "margin-top:8px" }, [h("span", { text: "Column" }),
            h("select", { class: "inl", onchange: e => { sp.column = e.target.value; sp.map = {}; draw(); } }, [h("option", { value: "", text: "— choose —" }), ...cols.map(x => h("option", { value: x, text: x, ...(sp.column === x ? { selected: true } : {}) }))]),
            h("span", { text: "New values:" }),
            h("select", { class: "inl", onchange: e => { sp.new_values = e.target.value; } }, [["create", "create a sheet automatically"], ["drop", "ignore those rows"]].map(([v, l]) => h("option", { value: v, text: l, ...(sp.new_values === v ? { selected: true } : {}) })))]));
          const keys = Object.keys(sp.map || {});
          if (keys.length) card.appendChild(h("div", { class: "scroll", style: "margin-top:10px" }, [h("table", { class: "t" }, [h("thead", {}, [h("tr", {}, [h("th", { text: "Value of " + sp.column }), h("th", { text: "Sheet name" })])]),
            h("tbody", {}, keys.map(k => h("tr", {}, [h("td", { text: k }), h("td", {}, [h("input", { type: "text", class: "inl", value: sp.map[k], oninput: e => { sp.map[k] = e.target.value; } })])])))])]));
          card.appendChild(h("div", { class: "muted small", style: "margin-top:6px", text: "Values not listed here get a sheet named after the value (or are ignored, per the setting above)." }));
        } else {
          (sp.sheets = sp.sheets || []).forEach((sh, i) => card.appendChild(h("div", { class: "card", style: "margin:10px 0 0;box-shadow:none;background:var(--surface-2)" }, [
            h("div", { class: "row" }, [h("input", { type: "text", class: "inl", value: sh.name, oninput: e => { sh.name = e.target.value; } }), h("span", { text: "rows matching" }),
              h("select", { class: "inl", onchange: e => { sh.match = e.target.value; } }, [["all", "ALL"], ["any", "ANY"]].map(([v, l]) => h("option", { value: v, text: l, ...(sh.match === v ? { selected: true } : {}) }))),
              h("span", { text: "of:" }), h("span", { class: "spacer" }), h("button", { class: "icon-btn", text: "✕", onclick: () => { sp.sheets.splice(i, 1); draw(); } })]),
            condEditor(sh.conditions = sh.conditions || [], draw)])));
          card.appendChild(h("button", { class: "btn ghost small", style: "margin-top:10px", text: "+ sheet", onclick: () => { sp.sheets.push({ name: "Sheet " + (sp.sheets.length + 1), match: "all", conditions: [] }); draw(); } }));
        }
        card.appendChild(h("div", { class: "row", style: "margin-top:12px" }, [h("span", { text: "Sheet for everything else (optional)" }),
          h("input", { type: "text", class: "inl", value: sp.rest_sheet || "", placeholder: "e.g. Other", oninput: e => { sp.rest_sheet = e.target.value; } })]));
      }
      body.appendChild(card);
      body.appendChild(h("div", { class: "row" }, [h("span", { class: "spacer" }), h("button", { class: "btn", text: "Save filters", onclick: async () => {
        F.pre = F.pre.filter(c => c.col); try { await saveSystem({ filters: F }, "Filters saved"); } catch (e) { toast(e.message, true); } } })]));
    }
    draw();
  }

  /* -------------------------------------------------------------- mappings tab */
  function tabMaps(body) {
    const names = Object.keys(SYS.tables);
    if (!names.length) { body.appendChild(h("div", { class: "card empty", text: "This system has no mapping tables — every column is built from the input itself." })); return; }
    const used = {};
    SYS.sheets.forEach(d => d.columns.forEach(c => { if (c.kind === "formula") (c.raw.match(/(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!/g) || []).forEach(x => {
      const n = x.replace(/!$/, "").replace(/^'|'$/g, ""); (used[n] = used[n] || []); if (used[n].indexOf(d.name + "." + c.name) < 0) used[n].push(d.name + "." + c.name); }); }));
    names.forEach(n => {
      const t = SYS.tables[n];
      body.appendChild(h("div", { class: "card" }, [h("div", { class: "row" }, [h("h2", { text: n, style: "margin:0" }), h("span", { class: "pill", text: t.records.length + " rows" }), h("span", { class: "spacer" }),
        h("span", { class: "muted small", text: (SYS.table_sources[n] || "") + (used[n] ? " · used by " + used[n].join(", ")
          : ((SYS.toggles || []).some(g => g.sheet === n) ? " · toggle sheet: " + SYS.toggles.filter(g => g.sheet === n).map(g => g.name + " = " + g.value).join(", ") : " · not used by any rule")) }),
        h("button", { class: "btn ghost small", text: "Download", onclick: () => dl("/api/systems/" + SYS.id + "/tables/" + encodeURIComponent(n) + "/download") })]),
        h("div", { class: "scroll", style: "margin-top:10px;max-height:260px" }, [h("table", { class: "t" }, [h("thead", {}, [h("tr", {}, t.headers.map(x => h("th", { text: x })))]),
          h("tbody", {}, t.records.slice(0, 40).map(r => h("tr", {}, t.headers.map(x => h("td", { text: r[x] == null ? "" : String(r[x]) })))))])])]));
    });
    body.appendChild(h("div", { class: "muted small", text: "To use updated mappings, attach the new mapping file on the \"Run on new data\" tab — tables are matched by sheet name." }));
  }

  /* --------------------------------------------------------------- history tab */
  async function tabHistory(body) {
    body.appendChild(h("div", { class: "muted" }, [h("span", { class: "spin" }), "Loading…"]));
    try {
      const r = await call("/api/systems/" + SYS.id + "/runs");
      clear(body);
      if (!r.runs.length) { body.appendChild(h("div", { class: "card empty", text: "No runs yet." })); return; }
      body.appendChild(h("div", { class: "card" }, [h("table", { class: "t" }, [h("thead", {}, [h("tr", {}, ["When (UTC)", "Inputs", "Rows in → out", "Sheets", ""].map(x => h("th", { text: x })))]),
        h("tbody", {}, r.runs.map(m => h("tr", {}, [h("td", { text: when(m.created) }), h("td", { class: "small", text: (m.inputs || []).map(i => i.file).join(", ") }),
          h("td", { text: m.input_rows + " → " + m.rows }), h("td", { class: "small", text: Object.keys(m.sheets).join(", ") }),
          h("td", {}, [h("button", { class: "btn ghost small", text: "Download", onclick: () => dl("/api/systems/" + SYS.id + "/runs/" + m.run_id + "/download") })])])))])]));
    } catch (e) { clear(body).appendChild(h("div", { class: "note err", text: e.message })); }
  }

  /* ==================================================================== boot */
  (async function () {
    try { const r = await call("/api/health"); $("#build-tag").textContent = "engine " + r.engine_build; } catch (e) {}
    route();
  })();
})();
