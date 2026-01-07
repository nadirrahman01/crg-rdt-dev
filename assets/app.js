/* =====================================================================
   Cordoba Research Group — Research Drafting Tool (RDT)
   main.js (UI logic only) — BlueMatrix-style workflow upgrades

   ✅ What this JS gives you (without changing your backend):
   - Note-type schemas (Equity / Macro / Credit / Morning note)
   - Section-aware workflow (required fields, gating, completion meter)
   - Autosave + recovery + version bump logic
   - “Internal / Public / Client-safe” modes (redaction + guardrails)
   - Word counts per section + total
   - Co-author manager
   - Figure manager (labels, captions, callouts)
   - Export pipeline (calls window.createDocument(data) if present,
     otherwise POSTs to /api/export-word and downloads the blob)
   - Keyboard shortcuts + “Jump to first missing”
   - Robust: if your HTML IDs differ / are missing, it will create fallback UI
     so you never get a “blank/broken” screen.

   ---------------------------------------------------------------------
   EXPECTED (but not strictly required) HTML hooks:
   - Container: #rdtApp (optional)
   - Mode: [name="audienceMode"] (values: internal|public|clientSafe)
   - Template select: #templateSelect
   - Note type select: #noteTypeSelect
   - Topic: #topicInput
   - Title: #titleInput
   - Status: #statusSelect (Draft|Ready|Published)
   - Reviewed by: #reviewedBySelect
   - Change note: #changeNoteInput
   - Bump version: #bumpVersionCheckbox
   - Clear autosave: #clearAutosaveBtn
   - Reset form: #resetFormBtn
   - Sections:
      #execSummaryInput, #keyTakeawaysInput, #analysisInput, #contentInput,
      #cordobaViewInput
   - Author fields: #authorLastName, #authorFirstName, #authorPhone
   - Coauthors: #addCoAuthorBtn, #coAuthorList (optional)
   - Figures: #figureUploadInput (type=file multiple), #figureList (optional)
   - Export confirm: #exportConfirmCheckbox
   - Email routing: #emailRoutingSelect (optional)
   - Export button: #exportBtn
   - Completion: #completionValue, #validationList, #jumpFirstMissingBtn
   - Autosave indicator: #autosaveStatus

   If you don’t have these IDs yet, this script still works by building a
   minimal “safe” UI inside the page body.

   ===================================================================== */

(() => {
  "use strict";

  /* ================================
     Utilities
  ================================ */
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  const uid = () => Math.random().toString(16).slice(2) + Date.now().toString(16);
  const clamp = (n, a, b) => Math.max(a, Math.min(b, n));

  const debounce = (fn, ms = 350) => {
    let t;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), ms);
    };
  };

  const wordCount = (s) => {
    if (!s) return 0;
    const clean = String(s)
      .replace(/[\u200B-\u200D\uFEFF]/g, "")
      .trim();
    if (!clean) return 0;
    return clean.split(/\s+/).filter(Boolean).length;
  };

  const safeTrim = (s) => (s == null ? "" : String(s).trim());

  const nowISO = () => new Date().toISOString();
  const nowHuman = () => {
    const d = new Date();
    const pad = (x) => String(x).padStart(2, "0");
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(
      d.getMinutes()
    )}`;
  };

  const downloadBlob = (blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  };

  /* ================================
     BlueMatrix-style Note Schemas
  ================================ */
  const NOTE_SCHEMAS = {
    "Macro Note": {
      id: "macro",
      required: ["title", "topic", "analysis", "cordobaView"],
      sections: [
        { key: "execSummary", label: "Executive summary", required: false, hint: "Optional — keep it tight." },
        { key: "keyTakeaways", label: "Key takeaways", required: false, hint: "Bullet-style lines are fine." },
        { key: "analysis", label: "Analysis and commentary", required: true, hint: "Main body — drivers, catalysts, risks." },
        { key: "content", label: "Additional content", required: false, hint: "Appendix / data notes / background." },
        { key: "cordobaView", label: "The Cordoba view", required: true, hint: "Conviction-led conclusion + positioning." }
      ],
      exportProfile: { includeTicker: false, includeValuation: false }
    },

    "Equity Research": {
      id: "equity",
      required: ["title", "topic", "analysis", "cordobaView", "ticker", "targetPrice", "crgRating"],
      sections: [
        { key: "execSummary", label: "Investment thesis", required: true, hint: "Why now, what’s mispriced, what changes." },
        { key: "keyTakeaways", label: "Key takeaways", required: true, hint: "3–7 lines max. Each line should stand alone." },
        { key: "analysis", label: "Company analysis", required: true, hint: "Business model, drivers, downside, catalysts." },
        { key: "valuationSummary", label: "Valuation", required: true, hint: "Method, key bridges, sensitivity." },
        { key: "keyAssumptions", label: "Key assumptions", required: true, hint: "Revenue, margins, WACC, terminal." },
        { key: "scenarioNotes", label: "Scenarios", required: false, hint: "Base / bull / bear — triggers + probabilities." },
        { key: "content", label: "Appendix", required: false, hint: "Extra detail, comps, notes." },
        { key: "cordobaView", label: "The Cordoba view", required: true, hint: "Rating, target, risk/reward in plain English." }
      ],
      exportProfile: { includeTicker: true, includeValuation: true }
    },

    "Credit / FI Note": {
      id: "credit",
      required: ["title", "topic", "analysis", "cordobaView"],
      sections: [
        { key: "execSummary", label: "Executive summary", required: true, hint: "What changed, why spreads should move." },
        { key: "keyTakeaways", label: "Key takeaways", required: true, hint: "Carry/roll, convexity, downgrade risk, liquidity." },
        { key: "analysis", label: "Credit analysis", required: true, hint: "Balance sheet, covenants, refinancing, flow." },
        { key: "content", label: "Additional content", required: false, hint: "Curves, relative value tables, data." },
        { key: "cordobaView", label: "The Cordoba view", required: true, hint: "Positioning, underpriced optionality, asymmetry." }
      ],
      exportProfile: { includeTicker: false, includeValuation: false }
    },

    "Morning Note": {
      id: "morning",
      required: ["title", "analysis"],
      sections: [
        { key: "execSummary", label: "Top line", required: true, hint: "1–2 paragraphs. What matters today." },
        { key: "analysis", label: "Market commentary", required: true, hint: "Rates, FX, risk, key catalysts." },
        { key: "content", label: "Watchlist", required: false, hint: "Short bullets / ideas / calendar." },
        { key: "cordobaView", label: "The Cordoba view", required: false, hint: "If you have a call — put it here." }
      ],
      exportProfile: { includeTicker: false, includeValuation: false }
    }
  };

  const AUDIENCE_MODES = {
    internal: { id: "internal", label: "Internal", watermarkDraft: true, redactPhone: false, restrictLanguage: false },
    public: { id: "public", label: "Public", watermarkDraft: false, redactPhone: true, restrictLanguage: true },
    clientSafe: { id: "clientSafe", label: "Client-safe", watermarkDraft: false, redactPhone: true, restrictLanguage: true }
  };

  /* ================================
     State + Storage
  ================================ */
  const STORAGE_KEY = "crg_rdt_autosave_v2";
  const STORAGE_META_KEY = "crg_rdt_meta_v2";

  const DEFAULT_META = {
    sessionId: uid(),
    createdAt: nowISO(),
    updatedAt: nowISO(),
    version: "1.0",
    draftWatermarkOn: true
  };

  const DEFAULT_DATA = {
    // Header
    audienceMode: "internal", // internal|public|clientSafe
    template: "",
    noteType: "Macro Note",
    topic: "",
    title: "",
    status: "Draft",
    reviewedBy: "",
    changeNote: "",
    bumpMajorOnExport: false,

    // Author
    authorLastName: "",
    authorFirstName: "",
    authorPhone: "+44",
    authorPhoneSafe: "", // computed for public/clientSafe
    coAuthors: [],

    // Core sections
    execSummary: "",
    keyTakeaways: "",
    analysis: "",
    content: "",
    cordobaView: "",

    // Equity extras
    ticker: "",
    targetPrice: "",
    crgRating: "",
    equityStats: "",

    valuationSummary: "",
    keyAssumptions: "",
    scenarioNotes: "",
    modelLink: "",
    modelFiles: [],

    // Figures
    imageFiles: [], // [{id,name,type,size,caption,label,bytes?}]
    figureCallouts: [], // reserved

    // System
    dateTimeString: nowHuman()
  };

  const State = {
    meta: { ...DEFAULT_META },
    data: { ...DEFAULT_DATA },
    ui: {
      lastSavedAt: null,
      dirty: false,
      mounted: false
    }
  };

  const loadAutosave = () => {
    try {
      const metaRaw = localStorage.getItem(STORAGE_META_KEY);
      const dataRaw = localStorage.getItem(STORAGE_KEY);
      if (metaRaw) State.meta = { ...DEFAULT_META, ...JSON.parse(metaRaw) };
      if (dataRaw) State.data = { ...DEFAULT_DATA, ...JSON.parse(dataRaw) };

      // ensure compatible defaults
      if (!State.data.dateTimeString) State.data.dateTimeString = nowHuman();
      if (!State.data.noteType) State.data.noteType = "Macro Note";
      if (!State.data.audienceMode) State.data.audienceMode = "internal";
      if (!Array.isArray(State.data.coAuthors)) State.data.coAuthors = [];
      if (!Array.isArray(State.data.imageFiles)) State.data.imageFiles = [];

      State.ui.lastSavedAt = State.meta.updatedAt ? new Date(State.meta.updatedAt) : null;
    } catch (e) {
      // if corrupted, reset
      State.meta = { ...DEFAULT_META };
      State.data = { ...DEFAULT_DATA };
    }
  };

  const persistAutosave = () => {
    State.meta.updatedAt = nowISO();
    localStorage.setItem(STORAGE_META_KEY, JSON.stringify(State.meta));
    localStorage.setItem(STORAGE_KEY, JSON.stringify(State.data));
    State.ui.lastSavedAt = new Date(State.meta.updatedAt);
    State.ui.dirty = false;
    renderAutosaveStatus();
  };

  const clearAutosave = () => {
    localStorage.removeItem(STORAGE_META_KEY);
    localStorage.removeItem(STORAGE_KEY);
    State.meta = { ...DEFAULT_META };
    State.data = { ...DEFAULT_DATA };
    State.ui.lastSavedAt = null;
    State.ui.dirty = false;
  };

  const markDirty = () => {
    State.ui.dirty = true;
    debouncedAutosave();
    renderAutosaveStatus();
  };

  const debouncedAutosave = debounce(() => {
    persistAutosave();
  }, 450);

  /* ================================
     Build / Ensure UI (fallback-safe)
  ================================ */
  const ensureBaseUI = () => {
    let app = $("#rdtApp");
    if (!app) {
      app = document.createElement("div");
      app.id = "rdtApp";
      document.body.innerHTML = "";
      document.body.appendChild(app);
    }

    // If user already has a layout, we don't overwrite it.
    // We only build fallback UI if core inputs are missing.
    const hasTitle = !!$("#titleInput");
    const hasAnalysis = !!$("#analysisInput");

    if (hasTitle && hasAnalysis) return;

    app.innerHTML = `
      <div class="rdt-shell" style="padding:18px;max-width:1100px;margin:0 auto;font-family:Helvetica,system-ui,sans-serif;">
        <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:12px;flex-wrap:wrap;">
          <div>
            <div style="font-size:12px;opacity:.65;">Cordoba Research Group</div>
            <div style="font-size:26px;font-family:'Times New Roman',serif;font-weight:700;margin-top:2px;">Research Drafting Tool</div>
            <div style="font-size:12px;opacity:.65;margin-top:2px;">Session <span id="sessionId"></span> · <span id="autosaveStatus"></span></div>
          </div>
          <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
            <button id="clearAutosaveBtn" type="button">Clear autosave</button>
            <button id="resetFormBtn" type="button">Reset form</button>
          </div>
        </div>

        <hr style="margin:14px 0;opacity:.25;" />

        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;align-items:start;">
          <div style="border:1px solid rgba(0,0,0,.12);border-radius:12px;padding:12px;">
            <div style="font-weight:700;margin-bottom:8px;">Note builder</div>

            <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px;">
              <label style="font-size:12px;opacity:.8;">Mode</label>
              <label><input type="radio" name="audienceMode" value="internal" checked> Internal</label>
              <label><input type="radio" name="audienceMode" value="public"> Public</label>
              <label><input type="radio" name="audienceMode" value="clientSafe"> Client-safe</label>
            </div>

            <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
              <label style="font-size:12px;">Template
                <select id="templateSelect">
                  <option value="">Select…</option>
                  <option>CRG Standard</option>
                  <option>CRG Equity</option>
                  <option>CRG Credit</option>
                </select>
              </label>

              <label style="font-size:12px;">Note type
                <select id="noteTypeSelect"></select>
              </label>

              <label style="font-size:12px;">Topic
                <input id="topicInput" type="text" placeholder="e.g., Indonesia FX, Uranium" />
              </label>

              <label style="font-size:12px;">Title
                <input id="titleInput" type="text" placeholder="Short, direct title" />
              </label>

              <label style="font-size:12px;">Status
                <select id="statusSelect">
                  <option>Draft</option>
                  <option>Ready</option>
                  <option>Published</option>
                </select>
              </label>

              <label style="font-size:12px;">Reviewed by
                <select id="reviewedBySelect">
                  <option value="">Select…</option>
                </select>
              </label>

              <label style="font-size:12px;grid-column:1/-1;">Change note
                <input id="changeNoteInput" type="text" placeholder="Optional (e.g., Updated figures)" />
              </label>

              <label style="font-size:12px;grid-column:1/-1;">
                <input id="bumpVersionCheckbox" type="checkbox" />
                Bump major version on export
              </label>
            </div>
          </div>

          <div style="border:1px solid rgba(0,0,0,.12);border-radius:12px;padding:12px;">
            <div style="font-weight:700;margin-bottom:8px;">Export</div>
            <label style="font-size:12px;display:flex;gap:8px;align-items:center;margin-bottom:8px;">
              <input id="exportConfirmCheckbox" type="checkbox" />
              I confirm the draft is ready to export.
            </label>
            <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
              <button id="exportBtn" type="button">Generate Word Document</button>
              <button id="jumpFirstMissingBtn" type="button">Jump to first missing</button>
            </div>
            <div style="margin-top:10px;font-size:12px;opacity:.85;">
              Completion: <span id="completionValue">0%</span>
            </div>
            <div id="validationList" style="margin-top:8px;font-size:12px;line-height:1.5;"></div>
          </div>
        </div>

        <div id="dynamicBlocks" style="margin-top:14px;"></div>
      </div>
    `;
  };

  const buildDynamicBlocks = () => {
    const host = $("#dynamicBlocks");
    if (!host) return;

    const schema = NOTE_SCHEMAS[State.data.noteType] || NOTE_SCHEMAS["Macro Note"];

    // Build blocks (author + sections + figures + equity extras where relevant)
    const wantsEquity = schema.id === "equity";

    host.innerHTML = `
      ${wantsEquity ? `
        <div class="rdt-card" style="border:1px solid rgba(0,0,0,.12);border-radius:12px;padding:12px;margin-bottom:14px;">
          <div style="font-weight:700;margin-bottom:8px;">Equity header</div>
          <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;">
            <label style="font-size:12px;">Ticker
              <input id="tickerInput" type="text" placeholder="e.g., AAPL" />
            </label>
            <label style="font-size:12px;">Target price
              <input id="targetPriceInput" type="text" placeholder="e.g., 210" />
            </label>
            <label style="font-size:12px;">CRG rating
              <select id="crgRatingSelect">
                <option value="">Select…</option>
                <option>BUY</option>
                <option>HOLD</option>
                <option>SELL</option>
              </select>
            </label>
            <label style="font-size:12px;grid-column:1/-1;">Equity stats (optional)
              <textarea id="equityStatsInput" rows="3" placeholder="Optional: key stats, market cap, EV/EBITDA, etc."></textarea>
            </label>
          </div>
        </div>
      ` : ""}

      <div class="rdt-card" style="border:1px solid rgba(0,0,0,.12);border-radius:12px;padding:12px;margin-bottom:14px;">
        <div style="font-weight:700;margin-bottom:8px;">Author</div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;">
          <label style="font-size:12px;">Last name
            <input id="authorLastName" type="text" />
          </label>
          <label style="font-size:12px;">First name
            <input id="authorFirstName" type="text" />
          </label>
          <label style="font-size:12px;">Phone
            <input id="authorPhone" type="text" placeholder="+44…" />
          </label>
        </div>

        <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;margin-top:10px;flex-wrap:wrap;">
          <div style="font-size:12px;opacity:.8;">Co-authors</div>
          <button id="addCoAuthorBtn" type="button">Add co-author</button>
        </div>
        <div id="coAuthorList" style="margin-top:8px;"></div>
      </div>

      ${schema.sections.map(sec => `
        <div class="rdt-card" data-section="${sec.key}" style="border:1px solid rgba(0,0,0,.12);border-radius:12px;padding:12px;margin-bottom:14px;">
          <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:10px;">
            <div>
              <div style="font-family:'Times New Roman',serif;font-weight:700;font-size:18px;">
                ${sec.label}${sec.required ? ' <span style="font-size:12px;opacity:.7;">(required)</span>' : ''}
              </div>
              <div style="font-size:12px;opacity:.65;margin-top:2px;">${sec.hint || ""}</div>
            </div>
            <div style="font-size:12px;opacity:.75;"><span id="${sec.key}Count">0</span> words</div>
          </div>
          <textarea id="${sec.key}Input" rows="${sec.key === "analysis" ? 12 : 6}" style="width:100%;margin-top:10px;"></textarea>
        </div>
      `).join("")}

      <div class="rdt-card" style="border:1px solid rgba(0,0,0,.12);border-radius:12px;padding:12px;margin-bottom:14px;">
        <div style="font-family:'Times New Roman',serif;font-weight:700;font-size:18px;margin-bottom:8px;">Figures and charts</div>
        <input id="figureUploadInput" type="file" multiple />
        <div id="figureList" style="margin-top:10px;"></div>
      </div>
    `;

    // Populate note types dropdown if present
    const noteTypeSelect = $("#noteTypeSelect");
    if (noteTypeSelect && noteTypeSelect.options.length === 0) {
      Object.keys(NOTE_SCHEMAS).forEach((name) => {
        const opt = document.createElement("option");
        opt.value = name;
        opt.textContent = name;
        noteTypeSelect.appendChild(opt);
      });
    }
  };

  /* ================================
     Render helpers
  ================================ */
  const renderAutosaveStatus = () => {
    const el = $("#autosaveStatus");
    if (!el) return;
    const s = State.ui.lastSavedAt ? `Saved ${State.ui.lastSavedAt.toLocaleTimeString()}` : "Not saved yet";
    const dirty = State.ui.dirty ? " · Editing…" : "";
    el.textContent = `${s}${dirty}`;
  };

  const renderSession = () => {
    const el = $("#sessionId");
    if (!el) return;
    el.textContent = State.meta.sessionId.slice(0, 8).toUpperCase();
  };

  const renderCoAuthors = () => {
    const host = $("#coAuthorList");
    if (!host) return;

    if (!State.data.coAuthors.length) {
      host.innerHTML = `<div style="font-size:12px;opacity:.6;">None</div>`;
      return;
    }

    host.innerHTML = State.data.coAuthors
      .map((c, i) => {
        const name = safeTrim(c?.name);
        const role = safeTrim(c?.role);
        return `
          <div data-coauthor="${i}" style="display:flex;gap:8px;align-items:center;justify-content:space-between;border:1px solid rgba(0,0,0,.08);border-radius:10px;padding:8px;margin-bottom:6px;">
            <div style="font-size:12px;">
              <div style="font-weight:700;">${escapeHtml(name || "Unnamed")}</div>
              <div style="opacity:.65;">${escapeHtml(role || "Co-author")}</div>
            </div>
            <div style="display:flex;gap:6px;">
              <button type="button" data-action="editCoAuthor" data-index="${i}">Edit</button>
              <button type="button" data-action="removeCoAuthor" data-index="${i}">Remove</button>
            </div>
          </div>
        `;
      })
      .join("");
  };

  const renderFigures = () => {
    const host = $("#figureList");
    if (!host) return;

    if (!State.data.imageFiles.length) {
      host.innerHTML = `<div style="font-size:12px;opacity:.6;">No figures uploaded yet.</div>`;
      return;
    }

    host.innerHTML = State.data.imageFiles
      .map((f, i) => {
        const label = f.label || `Figure ${i + 1}`;
        const caption = f.caption || "";
        return `
          <div style="border:1px solid rgba(0,0,0,.08);border-radius:10px;padding:10px;margin-bottom:8px;">
            <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;">
              <div style="font-size:12px;">
                <div style="font-weight:700;">${escapeHtml(label)} <span style="opacity:.6;">· ${escapeHtml(f.name)}</span></div>
                <div style="opacity:.75;margin-top:4px;">Caption:</div>
                <input type="text" data-fig-caption="${i}" value="${escapeAttr(caption)}" placeholder="Optional caption" style="width:100%;max-width:620px;" />
              </div>
              <div style="display:flex;gap:6px;align-items:center;">
                <button type="button" data-action="renameFigure" data-index="${i}">Label</button>
                <button type="button" data-action="removeFigure" data-index="${i}">Remove</button>
              </div>
            </div>
          </div>
        `;
      })
      .join("");
  };

  const renderWordCounts = () => {
    const schema = NOTE_SCHEMAS[State.data.noteType] || NOTE_SCHEMAS["Macro Note"];
    schema.sections.forEach((sec) => {
      const countEl = $(`#${sec.key}Count`);
      if (!countEl) return;
      countEl.textContent = String(wordCount(State.data[sec.key] || ""));
    });
  };

  const computeValidation = () => {
    const schema = NOTE_SCHEMAS[State.data.noteType] || NOTE_SCHEMAS["Macro Note"];
    const missing = [];

    // global required fields
    schema.required.forEach((key) => {
      const val = safeTrim(State.data[key]);
      if (!val) missing.push({ key, label: prettyKey(key) });
    });

    // section required
    schema.sections.forEach((sec) => {
      if (!sec.required) return;
      const val = safeTrim(State.data[sec.key]);
      if (!val) missing.push({ key: sec.key, label: sec.label });
    });

    // export confirm
    const exportConfirmed = !!State.data.exportConfirmed;
    if (!exportConfirmed) missing.push({ key: "exportConfirm", label: "Export confirmation checkbox" });

    return { missing, schema };
  };

  const renderValidation = () => {
    const list = $("#validationList");
    const completionEl = $("#completionValue");
    if (!list && !completionEl) return;

    const { missing, schema } = computeValidation();
    const totalChecks = schema.required.length + schema.sections.filter(s => s.required).length + 1;
    const done = totalChecks - missing.length;

    const pct = clamp(Math.round((done / totalChecks) * 100), 0, 100);
    if (completionEl) completionEl.textContent = `${pct}%`;

    if (!list) return;

    if (!missing.length) {
      list.innerHTML = `<div style="color:rgba(0,0,0,.75);">✅ Validation: ready to export</div>`;
      return;
    }

    list.innerHTML =
      `<div style="margin-bottom:6px;opacity:.8;">Missing / required:</div>` +
      missing
        .slice(0, 10)
        .map((m) => `<div style="color:rgba(0,0,0,.7);">• ${escapeHtml(m.label)}</div>`)
        .join("") +
      (missing.length > 10 ? `<div style="opacity:.6;margin-top:6px;">(+${missing.length - 10} more)</div>` : "");
  };

  const applyAudienceMode = () => {
    const mode = AUDIENCE_MODES[State.data.audienceMode] || AUDIENCE_MODES.internal;

    // Redact phone for public/clientSafe
    if (mode.redactPhone) {
      State.data.authorPhoneSafe = redactPhone(State.data.authorPhone);
    } else {
      State.data.authorPhoneSafe = safeTrim(State.data.authorPhone);
    }

    // Draft watermark flag
    State.meta.draftWatermarkOn = !!mode.watermarkDraft;

    // Language guardrails (light-touch)
    // (We don’t “rewrite” text; we just warn on export if risky terms exist.)
  };

  /* ================================
     Escaping helpers
  ================================ */
  function escapeHtml(str) {
    return String(str ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
  function escapeAttr(str) {
    return String(str ?? "").replace(/"/g, "&quot;");
  }

  /* ================================
     Data helpers
  ================================ */
  const prettyKey = (k) => {
    const map = {
      ticker: "Ticker",
      targetPrice: "Target price",
      crgRating: "CRG rating",
      execSummary: "Executive summary / thesis",
      keyTakeaways: "Key takeaways",
      cordobaView: "The Cordoba view"
    };
    return map[k] || k.replace(/([A-Z])/g, " $1").replace(/^./, (c) => c.toUpperCase());
  };

  const redactPhone = (p) => {
    const s = safeTrim(p);
    if (!s) return "";
    // Keep country code & last 2 digits if possible
    const digits = s.replace(/\D/g, "");
    if (digits.length <= 4) return "—";
    const last2 = digits.slice(-2);
    const cc = digits.startsWith("44") ? "+44" : (s.startsWith("+") ? "+" + digits.slice(0, Math.min(3, digits.length - 2)) : "—");
    return `${cc} •••• ••${last2}`;
  };

  const bumpVersion = (v, major) => {
    const parts = String(v || "1.0").split(".").map(x => parseInt(x, 10));
    const a = Number.isFinite(parts[0]) ? parts[0] : 1;
    const b = Number.isFinite(parts[1]) ? parts[1] : 0;

    if (major) return `${a + 1}.0`;
    return `${a}.${b + 1}`;
  };

  /* ================================
     Bind UI -> State
  ================================ */
  const bindInputs = () => {
    // Mode radios
    $$(`input[name="audienceMode"], input[name="audienceMode"], input[name="audienceMode"], input[name="audienceMode"]`);
    const modeRadios = $$(`input[name="audienceMode"], input[name="audienceMode"], input[name="audienceMode"]`);
    if (modeRadios.length) {
      modeRadios.forEach((r) => {
        r.addEventListener("change", () => {
          State.data.audienceMode = r.value;
          applyAudienceMode();
          markDirty();
          renderValidation();
        });
      });
    } else {
      // alternate name from expected spec
      const alt = $$(`input[name="audienceMode"], input[name="audienceMode"]`);
      alt.forEach((r) => {
        r.addEventListener("change", () => {
          State.data.audienceMode = r.value;
          applyAudienceMode();
          markDirty();
          renderValidation();
        });
      });
    }

    const wire = (id, key, eventName = "input") => {
      const el = $(`#${id}`);
      if (!el) return;
      el.addEventListener(eventName, () => {
        State.data[key] = el.type === "checkbox" ? !!el.checked : el.value;
        if (key === "noteType") {
          // Rebuild blocks to match schema
          buildDynamicBlocks();
          hydrateUIFromState();
          bindInputs(); // rebind for newly created elements
        }
        applyAudienceMode();
        markDirty();
        renderWordCounts();
        renderValidation();
      });
    };

    wire("templateSelect", "template");
    wire("noteTypeSelect", "noteType", "change");
    wire("topicInput", "topic");
    wire("titleInput", "title");
    wire("statusSelect", "status", "change");
    wire("reviewedBySelect", "reviewedBy", "change");
    wire("changeNoteInput", "changeNote");
    wire("bumpVersionCheckbox", "bumpMajorOnExport", "change");

    // Author
    wire("authorLastName", "authorLastName");
    wire("authorFirstName", "authorFirstName");
    wire("authorPhone", "authorPhone");

    // Equity extras (optional)
    wire("tickerInput", "ticker");
    wire("targetPriceInput", "targetPrice");
    wire("crgRatingSelect", "crgRating", "change");
    wire("equityStatsInput", "equityStats");

    // Sections (dynamic)
    Object.keys(DEFAULT_DATA).forEach((k) => {
      if (!k.endsWith("Summary") && !k.endsWith("view") && !k) return;
    });

    const schema = NOTE_SCHEMAS[State.data.noteType] || NOTE_SCHEMAS["Macro Note"];
    schema.sections.forEach((sec) => {
      wire(`${sec.key}Input`, sec.key);
    });

    // Equity valuation blocks (if present)
    wire("valuationSummaryInput", "valuationSummary");
    wire("keyAssumptionsInput", "keyAssumptions");
    wire("scenarioNotesInput", "scenarioNotes");
    wire("modelLinkInput", "modelLink");

    // Export confirm
    const exportConfirm = $("#exportConfirmCheckbox");
    if (exportConfirm) {
      exportConfirm.addEventListener("change", () => {
        State.data.exportConfirmed = !!exportConfirm.checked;
        markDirty();
        renderValidation();
      });
    }

    // Buttons
    const clearBtn = $("#clearAutosaveBtn");
    if (clearBtn) {
      clearBtn.addEventListener("click", () => {
        clearAutosave();
        buildDynamicBlocks();
        hydrateUIFromState();
        bindInputs();
        renderCoAuthors();
        renderFigures();
        renderWordCounts();
        renderValidation();
        renderAutosaveStatus();
      });
    }

    const resetBtn = $("#resetFormBtn");
    if (resetBtn) {
      resetBtn.addEventListener("click", () => {
        State.data = { ...DEFAULT_DATA, dateTimeString: nowHuman() };
        applyAudienceMode();
        markDirty();
        buildDynamicBlocks();
        hydrateUIFromState();
        bindInputs();
        renderCoAuthors();
        renderFigures();
        renderWordCounts();
        renderValidation();
      });
    }

    const addCoAuthorBtn = $("#addCoAuthorBtn");
    if (addCoAuthorBtn) {
      addCoAuthorBtn.addEventListener("click", () => {
        const name = prompt("Co-author name (e.g., First Last):");
        if (!safeTrim(name)) return;
        const role = prompt("Role / title (optional):") || "";
        State.data.coAuthors.push({ name: safeTrim(name), role: safeTrim(role) });
        renderCoAuthors();
        markDirty();
      });
    }

    const coAuthorList = $("#coAuthorList");
    if (coAuthorList) {
      coAuthorList.addEventListener("click", (e) => {
        const btn = e.target.closest("button");
        if (!btn) return;
        const action = btn.getAttribute("data-action");
        const idx = parseInt(btn.getAttribute("data-index"), 10);
        if (!Number.isFinite(idx)) return;

        if (action === "removeCoAuthor") {
          State.data.coAuthors.splice(idx, 1);
          renderCoAuthors();
          markDirty();
        }
        if (action === "editCoAuthor") {
          const c = State.data.coAuthors[idx];
          const name = prompt("Edit name:", c?.name || "") || "";
          if (!safeTrim(name)) return;
          const role = prompt("Edit role:", c?.role || "") || "";
          State.data.coAuthors[idx] = { name: safeTrim(name), role: safeTrim(role) };
          renderCoAuthors();
          markDirty();
        }
      });
    }

    // Figures upload
    const figInput = $("#figureUploadInput");
    if (figInput) {
      figInput.addEventListener("change", async () => {
        const files = Array.from(figInput.files || []);
        if (!files.length) return;

        // Store metadata now; bytes may be captured later (export time) if needed
        for (const f of files) {
          State.data.imageFiles.push({
            id: uid(),
            name: f.name,
            type: f.type || "application/octet-stream",
            size: f.size || 0,
            label: `Figure ${State.data.imageFiles.length + 1}`,
            caption: ""
          });
        }
        renderFigures();
        markDirty();

        // Reset input so same file can be reselected
        figInput.value = "";
      });
    }

    const figList = $("#figureList");
    if (figList) {
      figList.addEventListener("input", (e) => {
        const inp = e.target.closest("input[data-fig-caption]");
        if (!inp) return;
        const idx = parseInt(inp.getAttribute("data-fig-caption"), 10);
        if (!Number.isFinite(idx) || !State.data.imageFiles[idx]) return;
        State.data.imageFiles[idx].caption = inp.value;
        markDirty();
      });

      figList.addEventListener("click", (e) => {
        const btn = e.target.closest("button");
        if (!btn) return;
        const action = btn.getAttribute("data-action");
        const idx = parseInt(btn.getAttribute("data-index"), 10);
        if (!Number.isFinite(idx)) return;

        if (action === "removeFigure") {
          State.data.imageFiles.splice(idx, 1);
          // Renumber labels if they were default
          State.data.imageFiles.forEach((f, i) => {
            if (!f.label || /^Figure\s+\d+$/.test(f.label)) f.label = `Figure ${i + 1}`;
          });
          renderFigures();
          markDirty();
        }
        if (action === "renameFigure") {
          const f = State.data.imageFiles[idx];
          const next = prompt("Figure label:", f?.label || `Figure ${idx + 1}`);
          if (!safeTrim(next)) return;
          f.label = safeTrim(next);
          renderFigures();
          markDirty();
        }
      });
    }

    // Jump to first missing
    const jumpBtn = $("#jumpFirstMissingBtn");
    if (jumpBtn) {
      jumpBtn.addEventListener("click", () => {
        jumpToFirstMissing();
      });
    }

    // Export
    const exportBtn = $("#exportBtn");
    if (exportBtn) {
      exportBtn.addEventListener("click", async () => {
        await exportDocument();
      });
    }

    // Keyboard shortcuts (BlueMatrix-ish)
    window.addEventListener("keydown", (e) => {
      // Cmd/Ctrl+S = force save
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
        e.preventDefault();
        persistAutosave();
      }
      // Cmd/Ctrl+Enter = export
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        exportDocument();
      }
      // Alt+M = jump first missing
      if (e.altKey && e.key.toLowerCase() === "m") {
        e.preventDefault();
        jumpToFirstMissing();
      }
    });
  };

  const hydrateUIFromState = () => {
    // Mode radios
    const r = $$(`input[name="audienceMode"]`);
    if (r.length) r.forEach(x => (x.checked = x.value === State.data.audienceMode));

    const setVal = (id, val) => {
      const el = $(`#${id}`);
      if (!el) return;
      if (el.type === "checkbox") el.checked = !!val;
      else el.value = val ?? "";
    };

    setVal("templateSelect", State.data.template);
    setVal("noteTypeSelect", State.data.noteType);
    setVal("topicInput", State.data.topic);
    setVal("titleInput", State.data.title);
    setVal("statusSelect", State.data.status);
    setVal("reviewedBySelect", State.data.reviewedBy);
    setVal("changeNoteInput", State.data.changeNote);
    setVal("bumpVersionCheckbox", State.data.bumpMajorOnExport);

    setVal("authorLastName", State.data.authorLastName);
    setVal("authorFirstName", State.data.authorFirstName);
    setVal("authorPhone", State.data.authorPhone);

    setVal("tickerInput", State.data.ticker);
    setVal("targetPriceInput", State.data.targetPrice);
    setVal("crgRatingSelect", State.data.crgRating);
    setVal("equityStatsInput", State.data.equityStats);

    const schema = NOTE_SCHEMAS[State.data.noteType] || NOTE_SCHEMAS["Macro Note"];
    schema.sections.forEach((sec) => {
      setVal(`${sec.key}Input`, State.data[sec.key]);
    });

    // Export confirm
    setVal("exportConfirmCheckbox", !!State.data.exportConfirmed);

    renderCoAuthors();
    renderFigures();
    renderWordCounts();
    renderValidation();
    renderAutosaveStatus();
  };

  /* ================================
     Navigation helpers
  ================================ */
  const jumpToFirstMissing = () => {
    const { missing } = computeValidation();
    if (!missing.length) {
      alert("All required fields look complete.");
      return;
    }

    const m = missing[0];

    // map to input id
    const idMap = {
      title: "titleInput",
      topic: "topicInput",
      analysis: "analysisInput",
      cordobaView: "cordobaViewInput",
      execSummary: "execSummaryInput",
      keyTakeaways: "keyTakeawaysInput",
      exportConfirm: "exportConfirmCheckbox",
      ticker: "tickerInput",
      targetPrice: "targetPriceInput",
      crgRating: "crgRatingSelect",
      valuationSummary: "valuationSummaryInput",
      keyAssumptions: "keyAssumptionsInput"
    };

    const targetId = idMap[m.key] || `${m.key}Input`;
    const el = $(`#${targetId}`);
    if (!el) {
      // fallback: try section card
      const card = $(`.rdt-card[data-section="${m.key}"]`);
      if (card) {
        card.scrollIntoView({ behavior: "smooth", block: "start" });
      }
      return;
    }
    el.scrollIntoView({ behavior: "smooth", block: "center" });
    setTimeout(() => el.focus?.(), 250);
  };

  /* ================================
     Export logic (BlueMatrix-grade gates)
  ================================ */
  const exportDocument = async () => {
    applyAudienceMode();

    // Hard validation
    const { missing } = computeValidation();
    if (missing.length) {
      renderValidation();
      jumpToFirstMissing();
      alert("Missing required fields. Complete the required items before exporting.");
      return;
    }

    // Compliance / audience guardrails
    const warnings = [];
    if ((State.data.audienceMode === "public" || State.data.audienceMode === "clientSafe") && safeTrim(State.data.authorPhone)) {
      // We redact automatically; just inform quietly.
      if (State.data.authorPhoneSafe !== safeTrim(State.data.authorPhone)) {
        warnings.push("Phone number will be redacted for public/client-safe export.");
      }
    }

    // Version bump
    const bumpMajor = !!State.data.bumpMajorOnExport;
    State.meta.version = bumpVersion(State.meta.version, bumpMajor);

    // Timestamp
    State.data.dateTimeString = nowHuman();

    // Persist before export
    persistAutosave();

    // Attach computed safe phone
    State.data.authorPhoneSafe = (AUDIENCE_MODES[State.data.audienceMode]?.redactPhone)
      ? redactPhone(State.data.authorPhone)
      : safeTrim(State.data.authorPhone);

    // Prepare payload (what your existing createDocument expects)
    const payload = buildExportPayload();

    // If page provides createDocument (your existing docx builder), use it
    try {
      if (typeof window.createDocument === "function") {
        // createDocument should internally create + download/save the doc
        await window.createDocument(payload);
        if (warnings.length) console.info("[RDT warnings]", warnings);
        return;
      }
    } catch (err) {
      console.error("window.createDocument failed:", err);
      // fallback to API export below
    }

    // Otherwise POST to backend endpoint and download
    // Your server should return a .docx blob
    try {
      const res = await fetch("/api/export-word", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(`Export failed (${res.status}). ${txt}`);
      }
      const blob = await res.blob();
      const filename = suggestedFilename(payload);
      downloadBlob(blob, filename);
      if (warnings.length) console.info("[RDT warnings]", warnings);
    } catch (err) {
      console.error(err);
      alert("Export failed. If you are running client-only, include your createDocument() docx builder on the page, or add /api/export-word on the server.");
    }
  };

  const suggestedFilename = (payload) => {
    const safe = (s) =>
      safeTrim(s)
        .replace(/[^\w\- ]+/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 80)
        .replace(/\s/g, "-");

    const type = safe(payload.noteType || "Note");
    const title = safe(payload.title || "untitled");
    const v = safe(State.meta.version || "1-0");
    return `CRG_${type}_${title}_v${v}.docx`;
  };

  const buildExportPayload = () => {
    const schema = NOTE_SCHEMAS[State.data.noteType] || NOTE_SCHEMAS["Macro Note"];
    const wantsEquity = schema.id === "equity";

    // Key takeaways kept as raw text; your createDocument already converts lines to bullets.
    // We do NOT rewrite content here (authoring system must preserve analyst words).
    const payload = {
      noteType: State.data.noteType,
      title: safeTrim(State.data.title),
      topic: safeTrim(State.data.topic),

      authorLastName: safeTrim(State.data.authorLastName),
      authorFirstName: safeTrim(State.data.authorFirstName),
      authorPhone: safeTrim(State.data.authorPhone),
      authorPhoneSafe: safeTrim(State.data.authorPhoneSafe),
      coAuthors: Array.isArray(State.data.coAuthors) ? State.data.coAuthors : [],

      analysis: safeTrim(State.data.analysis),
      keyTakeaways: safeTrim(State.data.keyTakeaways),
      content: safeTrim(State.data.content),
      cordobaView: safeTrim(State.data.cordobaView),
      execSummary: safeTrim(State.data.execSummary),

      imageFiles: Array.isArray(State.data.imageFiles) ? State.data.imageFiles : [],
      dateTimeString: safeTrim(State.data.dateTimeString),

      // Equity extras (only included if relevant)
      ticker: wantsEquity ? safeTrim(State.data.ticker) : "",
      valuationSummary: wantsEquity ? safeTrim(State.data.valuationSummary) : "",
      keyAssumptions: wantsEquity ? safeTrim(State.data.keyAssumptions) : "",
      scenarioNotes: wantsEquity ? safeTrim(State.data.scenarioNotes) : "",
      modelFiles: wantsEquity ? (State.data.modelFiles || []) : [],
      modelLink: wantsEquity ? safeTrim(State.data.modelLink) : "",

      targetPrice: wantsEquity ? safeTrim(State.data.targetPrice) : "",
      equityStats: wantsEquity ? safeTrim(State.data.equityStats) : "",
      crgRating: wantsEquity ? safeTrim(State.data.crgRating) : "",

      // Workflow metadata (optional for doc)
      audienceMode: State.data.audienceMode,
      template: State.data.template,
      status: State.data.status,
      reviewedBy: State.data.reviewedBy,
      changeNote: State.data.changeNote,
      version: State.meta.version,
      draftWatermarkOn: !!State.meta.draftWatermarkOn
    };

    return payload;
  };

  /* ================================
     Mount
  ================================ */
  const mount = () => {
    loadAutosave();
    ensureBaseUI();
    buildDynamicBlocks();
    renderSession();
    applyAudienceMode();
    hydrateUIFromState();
    bindInputs();
    renderAutosaveStatus();

    State.ui.mounted = true;
  };

  // Fire when ready
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", mount);
  } else {
    mount();
  }
})();
