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
          s.match != null ? h("span", { class: "pill " + (s.match >= 0.9995 ? "ok" : "warn"), text: fmtPct(s.match) + " replay" }) : null]),
        h("div", { class: "muted small", text: s.description || "" }),
        h("div", { class: "row", style: "margin-top:10px" }, [
          h("span", { class: "pill", text: s.columns + " columns" }), h("span", { class: "pill", text: s.n_tables + " mappings" }),
          s.split ? h("span", { class: "pill acc", text: "multi-sheet" }) : null,
          s.n_review ? h("span", { class: "pill warn", text: s.n_review + " to review" }) : null]),
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
      h("div", { class: "muted", text: "Show the app one worked example; it learns the rules." })])]));
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
        h("button", { class: "btn", text: "Learn the rules →", ...(ready && !UP_BUSY ? {} : { disabled: true }), onclick: learnNow })]));
      return;
    }

    root.appendChild(h("div", { class: "card empty" }, [h("span", { class: "spin" }), "Learning every output column… this can take a few seconds."]));
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
    const rv = s.rules.filter(r => r.status === "unresolved" || r.status === "approx").length;
    root.appendChild(h("div", { class: "head" }, [h("div", {}, [h("h1", { text: s.name }), h("div", { class: "muted", text: s.description || "" }),
      h("div", { class: "row small", style: "margin-top:6px" }, [
        h("span", { class: "pill " + ((s.training.replay_match || 0) >= 0.9995 ? "ok" : "warn"), text: "Replays training output at " + fmtPct(s.training.replay_match) }),
        h("span", { class: "pill", text: s.output.columns.length + " columns" }), h("span", { class: "pill", text: Object.keys(s.tables).length + " mapping tables" }),
        rv ? h("span", { class: "pill warn", text: rv + " column(s) to review" }) : null])]),
      h("div", { class: "row" }, [
        h("button", { class: "btn ghost small", text: "Input template", onclick: () => dl("/api/systems/" + s.id + "/template") }),
        h("button", { class: "btn ghost small", text: "Re-learn", onclick: relearn }),
        h("button", { class: "btn ghost danger small", text: "Delete", onclick: async () => { if (!confirm("Delete system '" + s.name + "'?")) return;
          await call("/api/systems/" + s.id, { method: "DELETE" }); SYS = null; location.hash = "#/"; } })])]));
    const tabs = [["rules", "Rules"], ["run", "Run on new data"], ["filters", "Filters & outputs"], ["maps", "Mappings"], ["history", "History"]];
    root.appendChild(h("div", { class: "tabs" }, tabs.map(([k, l]) => h("a", { class: "tab" + (tab === k ? " on" : ""), href: "#/s/" + s.id + "/" + k, text: l }))));
    const body = h("div", {});
    root.appendChild(body);
    ({ rules: tabRules, run: tabRun, filters: tabFilters, maps: tabMaps, history: tabHistory }[tab] || tabRules)(body);
  }

  async function relearn() {
    if (!confirm("Re-learn all rules from the stored example files? Manual edits will be lost.")) return;
    const mode = SYS.filters.split.enabled ? "auto" : "auto";
    toast("Re-learning…");
    try { SYS = (await call("/api/systems/" + SYS.id + "/relearn", { method: "POST", json: { split: { mode } } })).system; renderSystem("rules"); toast("Done"); }
    catch (e) { toast(e.message, true); }
  }

  async function saveSystem(patch, msg) {
    SYS = (await call("/api/systems/" + SYS.id, { method: "PUT", json: patch })).system;
    toast(msg || "Saved");
  }

  /* ---------------------------------------------------------------- rules tab */
  function layoutCard() {
    const s = SYS, sc = s.output.sheet_columns || {}, rules = {};
    s.rules.forEach(r => { rules[r.name] = r; });
    const isInput = c => { const r = rules[c]; return r && r.kind === "copy" && r.col === c; };
    const sheets = Object.keys(sc);
    const split = s.filters.split;
    return h("div", { class: "card" }, [h("h2", { text: "Output sheets" }),
      h("div", { class: "muted small", text: "Each sheet is laid out exactly like your sample: the same columns in the same order, starting in the same cell. The sample's formulas are stored as written and placed in the derived columns (only the row number moves, like filling down). The mapping and toggle sheets the formulas depend on are added to the workbook." }),
      ...sheets.map(sh => {
        const cols = sc[sh], inp = cols.filter(isInput), der = cols.filter(c => !isInput(c));
        const cond = split.enabled ? ((split.sheets || []).find(x => x.name === sh) || {}).conditions : null;
        return h("div", { style: "margin-top:12px" }, [
          h("div", { class: "row" }, [h("b", { text: sh }), h("span", { class: "pill", text: inp.length + " input + " + der.length + " derived" }),
            cond && cond.length ? h("span", { class: "muted small", text: "rows where " + cond.map(c => c.col + " " + c.op + " " + (Array.isArray(c.value) ? c.value.join("|") : c.value)).join(" and ") }) : null]),
          h("div", { class: "chips", style: "margin:6px 0 0" }, [
            ...inp.map(c => h("span", { class: "chip", text: c })),
            ...der.map(c => h("span", { class: "chip", style: "background:var(--ok-soft);color:var(--ok);font-weight:600", title: (rules[c] || {}).description || "", text: c }))])]);
      }),
      h("div", { class: "row small", style: "margin-top:12px" }, [h("span", { class: "muted", text: "Also in the workbook:" }),
        ...Object.keys(s.tables).map(t => h("span", { class: "pill warn", text: t + ((s.toggles || []).some(g => g.sheet === t) ? " (toggle)" : " (mapping)") })),
        h("span", { class: "pill", text: "Run summary" })])]);
  }

  function tabRules(body) {
    const s = SYS;
    body.appendChild(layoutCard());
    body.appendChild(h("div", { class: "card" }, [
      h("div", { class: "row" }, [h("h2", { text: "How each output column is built", style: "margin:0" }), h("span", { class: "spacer" }),
        h("button", { class: "btn small", text: "Check against training output", onclick: runReplay })]),
      h("div", { class: "muted small", text: "Columns are learned in output order; later columns may use earlier ones. Edit a rule's JSON to correct it — a rule of kind \"formula\" accepts a pandas expression such as df['Amount'] * 0.05." }),
      ...(s.warnings || []).map(w => h("div", { class: "note", text: w })),
      h("div", { class: "scroll", style: "margin-top:10px;max-height:none" }, [rulesTable()])]));
    if (REPLAY) body.appendChild(replayCard());
  }

  function rulesTable() {
    const s = SYS;
    const tb = h("tbody");
    s.rules.forEach((r, i) => {
      const [cls, lab] = STATUS[r.status] || ["", r.status];
      const tr = h("tr", {}, [h("td", { text: i + 1, class: "muted" }),
        h("td", {}, [h("b", { text: r.name }), r.disabled ? h("span", { class: "pill", text: "disabled", style: "margin-left:6px" }) : null]),
        h("td", {}, [h("span", { class: "pill acc", text: KIND[r.kind] || r.kind })]),
        h("td", { class: "desc", text: r.description || "" }),
        h("td", {}, [r.match != null ? h("div", { class: "row" }, [h("div", { class: "bar " + (r.match >= .9995 ? "" : r.match >= .9 ? "warn" : "bad"), style: "width:60px" }, [h("i", { style: "width:" + (r.match * 100) + "%" })]), h("span", { class: "small", text: fmtPct(r.match) })]) : h("span", { class: "muted", text: "—" })]),
        h("td", {}, [h("span", { class: "pill " + cls, text: lab })]),
        h("td", { style: "white-space:nowrap" }, [
          h("button", { class: "icon-btn", text: "↑", title: "Move up", onclick: () => move(i, -1) }), " ",
          h("button", { class: "icon-btn", text: "↓", title: "Move down", onclick: () => move(i, 1) }), " ",
          h("button", { class: "btn ghost small", text: "Edit", onclick: () => toggleEditor(tr, r) })])]);
      tb.appendChild(tr);
    });
    return h("table", { class: "t" }, [h("thead", {}, [h("tr", {}, ["#", "Output column", "Type", "Rule", "Match", "Status", ""].map(x => h("th", { text: x })))]), tb]);
  }

  async function move(i, d) {
    const j = i + d; if (j < 0 || j >= SYS.rules.length) return;
    const rules = SYS.rules.slice(); const t = rules[i]; rules[i] = rules[j]; rules[j] = t;
    await saveSystem({ rules: rules.map(r => ({ name: r.name })) }, "Order saved"); renderSystem("rules");
  }

  const META = ["name", "match", "status", "description", "source", "notes", "error", "edited"];
  function toggleEditor(tr, r) {
    const nxt = tr.nextSibling;
    if (nxt && nxt.classList && nxt.classList.contains("ed")) { nxt.remove(); return; }
    const spec = {}; Object.keys(r).forEach(k => { if (META.indexOf(k) < 0) spec[k] = r[k]; });
    const ta = h("textarea", { spellcheck: "false" }); ta.value = JSON.stringify(spec, null, 2);
    const dis = h("input", { type: "checkbox", ...(r.disabled ? { checked: true } : {}) });
    const cell = h("td", { colspan: 7, style: "background:var(--surface-2)" }, [
      h("div", { class: "muted small", text: "Kinds: copy · const · text · datepart · arith · concat · lookup · cond · learned_map · formula. Excel formula: " + (r.excel_formula || "—") }),
      ta, h("div", { class: "row", style: "margin-top:8px" }, [h("label", { class: "small" }, [dis, " disable (output blank)"]), h("span", { class: "spacer" }),
        h("button", { class: "btn ghost small", text: "Cancel", onclick: () => cell.parentNode.remove() }),
        h("button", { class: "btn small", text: "Save rule", onclick: async () => {
          let obj; try { obj = JSON.parse(ta.value); } catch (e) { return toast("Invalid JSON: " + e.message, true); }
          if (!obj.kind) return toast("A rule needs a \"kind\"", true);
          obj.name = r.name; obj.edited = true; if (dis.checked) obj.disabled = true;
          const rules = SYS.rules.map(x => x.name === r.name ? obj : { name: x.name });
          await saveSystem({ rules }, "Rule saved — run the check to see its effect"); REPLAY = null; renderSystem("rules"); } })])]);
    tr.parentNode.insertBefore(h("tr", { class: "ed" }, [cell]), tr.nextSibling);
  }

  async function runReplay() {
    toast("Running the system on the training input…");
    try { const r = await call("/api/systems/" + SYS.id + "/replay", { method: "POST" }); REPLAY = r.report; SYS = r.system; renderSystem("rules"); }
    catch (e) { toast(e.message, true); }
  }

  function replayCard() {
    const R = REPLAY;
    const card = h("div", { class: "card" }, [h("h2", { text: "Replay check" })]);
    if (!R.aligned) { card.appendChild(h("div", { class: "note err", text: R.note })); return card; }
    card.appendChild(h("div", { class: "stat" }, [h("div", {}, [h("b", { text: fmtPct(R.match) }), h("span", { text: "cells identical" })]),
      h("div", {}, [h("b", { text: R.rows_pred + " / " + R.rows_expected }), h("span", { text: "rows produced / expected" })])]));
    Object.keys(R.rule_errors || {}).forEach(k => card.appendChild(h("div", { class: "note err", text: k + ": " + R.rule_errors[k] })));
    const bad = R.columns.filter(c => c.mismatches);
    if (!bad.length) card.appendChild(h("div", { class: "note good", text: "Every column reproduces the expected output exactly." }));
    bad.forEach(c => card.appendChild(h("details", { open: true }, [h("summary", { text: c.column + " — " + c.mismatches + " mismatches (" + fmtPct(c.match) + " match)" }),
      h("table", { class: "t" }, [h("thead", {}, [h("tr", {}, ["Row", "Expected", "Got"].map(x => h("th", { text: x })))]),
        h("tbody", {}, c.examples.map(e => h("tr", {}, [h("td", { text: e.row }), h("td", { text: String(e.expected) }), h("td", { text: String(e.got) })])))])])));
    return card;
  }

  /* ---------------------------------------------------------------- run tab */
  function tabRun(body) {
    const s = SYS;
    const files = { in: [], map: [], exp: [] }; let live = true; const tog = {};
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
      files.in.forEach(f => fd.append("files", f)); files.map.forEach(f => fd.append("mappings", f)); files.exp.forEach(f => fd.append("expected", f));
      if (live) fd.append("live_formulas", "1");
      if (Object.keys(tog).length) fd.append("toggles", JSON.stringify(tog));
      btn.disabled = true; btn.textContent = "Running…";
      try { LAST_RUN = (await call("/api/systems/" + s.id + "/run", { method: "POST", body: fd })).run; showRun(out, LAST_RUN); }
      catch (e) { toast(e.message, true); }
      btn.disabled = false; btn.textContent = "Run";
    } });
    body.appendChild(h("div", { class: "card" }, [h("h2", { text: "Run on new data" }),
      h("div", { class: "muted small", text: "Upload new input in the same layout (" + s.input.required.slice(0, 6).join(", ") + (s.input.required.length > 6 ? "…" : "") + "). The system builds every derived column" + (s.filters.split.enabled ? " and splits the result into sheets" : "") + "." }),
      h("div", { class: "grid g3", style: "margin-top:12px" }, [zone("in", "Input file(s)", "Drop new input here", true),
        zone("map", "Updated mappings (optional)", "Replace stored mapping tables for this run", true),
        zone("exp", "Expected output (optional)", "Compare the result to a known answer", false)]),
      togBox,
      h("div", { class: "row", style: "margin-top:14px" }, [
        h("label", { class: "small" }, [h("input", { type: "checkbox", checked: true, onchange: e => { live = e.target.checked; } }), " Write the formulas into the output cells (untick for plain values)"]),
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
        run.compare && run.compare.aligned ? h("div", {}, [h("b", { text: fmtPct(run.compare.match) }), h("span", { text: "match vs expected" })]) : null])]);
    if ((run.toggles || []).length) card.appendChild(h("div", { class: "note good", text: "Toggles used: " + run.toggles.map(t => t.name + " = " + t.value).join(", ") }));
    (run.notes || []).forEach(n => card.appendChild(h("div", { class: "note", text: n })));
    Object.keys(run.rule_errors || {}).forEach(k => card.appendChild(h("div", { class: "note err", text: k + ": " + run.rule_errors[k] })));
    if (run.compare) {
      if (!run.compare.aligned) card.appendChild(h("div", { class: "note err", text: run.compare.note }));
      else {
        const bad = run.compare.columns.filter(c => c.mismatches);
        if (!bad.length) card.appendChild(h("div", { class: "note good", text: "Identical to the expected output." }));
        bad.forEach(c => card.appendChild(h("details", {}, [h("summary", { text: "vs expected · " + c.column + ": " + c.mismatches + " differences" }),
          h("table", { class: "t" }, [h("thead", {}, [h("tr", {}, ["Row", "Expected", "Got"].map(x => h("th", { text: x })))]),
            h("tbody", {}, c.examples.map(e => h("tr", {}, [h("td", { text: e.row }), h("td", { text: String(e.expected) }), h("td", { text: String(e.got) })])))])])));
      }
    }
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
    const cols = SYS.input.columns.concat(SYS.output.columns.filter(c => SYS.input.columns.indexOf(c) < 0));
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
    SYS.rules.forEach(r => { const t = r.kind === "lookup" ? [r.table] : (r.kind === "arith" ? JSON.stringify(r.expr).match(/"table":"([^"]+)"/g) : null);
      (t || []).forEach(x => { const n = x.replace(/"table":"|"/g, ""); (used[n] = used[n] || []).push(r.name); }); });
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
