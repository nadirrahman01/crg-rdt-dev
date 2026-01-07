/* =========================================================
   Cordoba Research Group — Research Drafting Tool (RDT)
   Bank-grade Equity Research UI + Word Export (Tear Sheet)
   File: app.js
   Dependencies (load in HTML):
     1) docx (https://www.npmjs.com/package/docx) via CDN build that exposes window.docx
        Example:
        <script src="https://unpkg.com/docx@8.5.0/build/index.js"></script>

   Notes:
   - This script renders the entire UI into #rdtRoot (or <body> if missing).
   - Autosave is localStorage-based, session-scoped.
   - Export generates a Word .docx with an Equity tear-sheet layout and section formatting.
   ========================================================= */

/* ================================
   App Config
================================ */
const RDT = {
  appName: "Cordoba Research Group — Research Drafting Tool",
  version: "2.0",
  brand: {
    ink: "#0B0E14",
    muted: "rgba(11,14,20,.62)",
    gold: "#845F0F",
    soft: "#FFF7F0",
    border: "rgba(11,14,20,.10)",
    border2: "rgba(11,14,20,.14)",
    bg: "#FFFCF9",
    card: "#FFFFFF",
  },
  autosave: {
    enabled: true,
    intervalMs: 2500,
    keyPrefix: "CRG_RDT_V2",
  },
  required: {
    common: ["audience", "template", "noteType", "topic", "title", "status"],
    equity: ["ticker", "companyName", "rating", "investmentThesis", "keyTakeaways", "companyAnalysis", "valuation", "cordobaView"],
  },
  templates: [
    { value: "equity", label: "CRG Equity (Tear Sheet)" },
    { value: "macro", label: "CRG Macro Note" },
    { value: "credit", label: "CRG Credit / FI" },
    { value: "thematic", label: "CRG Thematic" },
    { value: "event", label: "CRG Event / Reaction" },
  ],
  audiences: [
    { value: "internal", label: "Internal" },
    { value: "public", label: "Public" },
    { value: "client_safe", label: "Client-safe" },
  ],
  ratings: [
    { value: "BUY", label: "BUY" },
    { value: "HOLD", label: "HOLD" },
    { value: "SELL", label: "SELL" },
    { value: "WATCH", label: "WATCH" },
  ],
  noteTypes: [
    { value: "equity", label: "Equity Research" },
    { value: "macro", label: "Macro" },
    { value: "fixed_income", label: "Fixed Income" },
    { value: "commodities", label: "Commodities" },
    { value: "geopolitics", label: "Geopolitics" },
  ],
};

/* ================================
   Utilities
================================ */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

function uid(len = 10) {
  const s = Math.random().toString(16).slice(2) + Date.now().toString(16);
  return s.slice(0, len).toUpperCase();
}

function nowStamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const yyyy = d.getFullYear();
  const mm = pad(d.getMonth() + 1);
  const dd = pad(d.getDate());
  const hh = pad(d.getHours());
  const mi = pad(d.getMinutes());
  const ss = pad(d.getSeconds());
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
}

function safeTrim(v) {
  return (v ?? "").toString().trim();
}

function wordCount(text) {
  const t = safeTrim(text);
  if (!t) return 0;
  return t.split(/\s+/).filter(Boolean).length;
}

function debounce(fn, ms = 250) {
  let t = null;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

function toTitleCase(s) {
  return (s || "")
    .split(" ")
    .filter(Boolean)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ");
}

function sanitizeFileName(name) {
  return (name || "CRG_Note")
    .replace(/[^\w\-]+/g, "_")
    .replace(/_+/g, "_")
    .slice(0, 90);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 500);
}

/* ================================
   State
================================ */
const AppState = {
  sessionId: uid(8),
  lastSavedAt: null,
  dirty: false,
  autosaveTimer: null,

  // Data model (single source of truth)
  data: {
    // Common
    audience: "internal",
    template: "equity",
    noteType: "equity",
    status: "DRAFT",
    reviewedBy: "",
    changeNote: "",
    bumpMajorOnExport: false,

    topic: "",
    title: "",
    tags: "",

    // Author
    authorFirstName: "",
    authorLastName: "",
    authorRole: "Research Analyst",
    authorEmail: "",
    authorPhoneCountry: "+44",
    authorPhone: "",
    coAuthors: [], // {id, name, role, email}

    // Equity tear sheet
    companyName: "",
    ticker: "",
    exchange: "",
    country: "",
    sector: "",
    industry: "",
    currency: "USD",

    rating: "HOLD",
    targetPrice: "",
    currentPrice: "",
    upsidePct: "",
    marketCap: "",
    sharesOut: "",
    floatPct: "",
    avgVol3m: "",
    beta: "",
    peFwd: "",
    evEbitda: "",
    dividendYield: "",

    // Core sections (equity)
    investmentThesis: "",
    keyTakeaways: "",
    companyAnalysis: "",
    valuation: "",
    keyAssumptions: "",
    scenarios: "",
    catalysts: "",
    risks: "",
    appendix: "",
    cordobaView: "",

    // Figures
    imageFiles: [], // {id, name, type, size, dataUrl}
    figureCaptions: {}, // {fileId: "caption"}

    // Export controls
    confirmReady: false,
    emailRouting: false,

    // Meta
    draftWatermark: true,
  },
};

/* ================================
   Storage (Autosave)
================================ */
function storageKey() {
  return `${RDT.autosave.keyPrefix}:${AppState.sessionId}`;
}

function saveToLocalStorage() {
  try {
    const payload = {
      sessionId: AppState.sessionId,
      savedAt: nowStamp(),
      version: RDT.version,
      data: AppState.data,
    };
    localStorage.setItem(storageKey(), JSON.stringify(payload));
    AppState.lastSavedAt = payload.savedAt;
    AppState.dirty = false;
    renderTopBarMeta();
  } catch (e) {
    console.warn("Autosave failed:", e);
  }
}

function loadFromLocalStorage(existingSessionId = null) {
  try {
    const key = existingSessionId ? `${RDT.autosave.keyPrefix}:${existingSessionId}` : storageKey();
    const raw = localStorage.getItem(key);
    if (!raw) return false;
    const parsed = JSON.parse(raw);
    if (!parsed?.data) return false;

    AppState.sessionId = parsed.sessionId || AppState.sessionId;
    AppState.data = mergeDefaults(AppState.data, parsed.data);
    AppState.lastSavedAt = parsed.savedAt || null;
    AppState.dirty = false;
    return true;
  } catch (e) {
    console.warn("Load autosave failed:", e);
    return false;
  }
}

function clearAutosave() {
  try {
    localStorage.removeItem(storageKey());
    AppState.lastSavedAt = null;
    AppState.dirty = false;
    renderTopBarMeta();
  } catch (e) {
    console.warn("Clear autosave failed:", e);
  }
}

function mergeDefaults(defaults, incoming) {
  // Deep-ish merge for plain objects/arrays used here
  const out = Array.isArray(defaults) ? [...defaults] : { ...defaults };
  for (const k of Object.keys(incoming || {})) {
    const dv = defaults?.[k];
    const iv = incoming?.[k];
    if (dv && typeof dv === "object" && !Array.isArray(dv) && iv && typeof iv === "object" && !Array.isArray(iv)) {
      out[k] = mergeDefaults(dv, iv);
    } else {
      out[k] = iv;
    }
  }
  return out;
}

function startAutosaveLoop() {
  if (!RDT.autosave.enabled) return;
  if (AppState.autosaveTimer) clearInterval(AppState.autosaveTimer);
  AppState.autosaveTimer = setInterval(() => {
    if (AppState.dirty) saveToLocalStorage();
  }, RDT.autosave.intervalMs);
}

/* ================================
   Validation
================================ */
function computeRequiredFields() {
  const base = [...RDT.required.common];
  const isEquity = AppState.data.template === "equity" || AppState.data.noteType === "equity";
  if (isEquity) return base.concat(RDT.required.equity);
  return base; // expand for other templates later
}

function validate() {
  const required = computeRequiredFields();
  const missing = [];
  for (const key of required) {
    const val = AppState.data[key];
    const ok =
      typeof val === "boolean" ? val === true : Array.isArray(val) ? val.length > 0 : safeTrim(val) !== "";
    if (!ok) missing.push(key);
  }

  // Special: confirmReady gates export only if status is FINAL or user ticks confirm
  const exportGateOk = AppState.data.confirmReady === true;

  return {
    required,
    missing,
    ok: missing.length === 0,
    exportGateOk,
  };
}

function fieldLabel(key) {
  const map = {
    audience: "Audience",
    template: "Template",
    noteType: "Note type",
    topic: "Topic",
    title: "Title",
    status: "Status",

    ticker: "Ticker",
    companyName: "Company name",
    rating: "Rating",
    investmentThesis: "Investment thesis",
    keyTakeaways: "Key takeaways",
    companyAnalysis: "Company analysis",
    valuation: "Valuation",
    cordobaView: "The Cordoba view",
  };
  return map[key] || toTitleCase(key.replace(/([A-Z])/g, " $1"));
}

/* ================================
   UI Rendering (Bank-grade layout)
================================ */
function mount() {
  const root = document.getElementById("rdtRoot") || document.body;
  root.innerHTML = "";

  // Shell
  const shell = document.createElement("div");
  shell.className = "rdt-shell";

  // Top bar (fixed-like)
  const top = document.createElement("div");
  top.className = "rdt-topbar";
  top.innerHTML = `
    <div class="rdt-topbar-left">
      <div class="rdt-brand">
        <div class="rdt-brand-title">Research Drafting Tool</div>
        <div class="rdt-brand-sub">
          <span>Cordoba Research Group</span>
          <span class="rdt-dot">•</span>
          <span>Session <strong id="rdtSession"></strong></span>
          <span class="rdt-dot">•</span>
          <span>Saved <strong id="rdtSavedAt">—</strong></span>
        </div>
      </div>
    </div>
    <div class="rdt-topbar-right">
      <button class="rdt-btn rdt-btn-ghost" id="btnLoadSession" title="Load a previous session from localStorage">Load</button>
      <button class="rdt-btn rdt-btn-ghost" id="btnClearAutosave" title="Clear autosave for this session">Clear autosave</button>
      <button class="rdt-btn rdt-btn-ghost" id="btnReset" title="Reset the form">Reset</button>
    </div>
  `;
  shell.appendChild(top);

  // Main split: Left editor + Right action/validation panel
  const main = document.createElement("div");
  main.className = "rdt-main";

  const left = document.createElement("div");
  left.className = "rdt-left";

  const right = document.createElement("div");
  right.className = "rdt-right";

  main.appendChild(left);
  main.appendChild(right);
  shell.appendChild(main);

  // Build left (editor)
  left.appendChild(renderHeaderGrid());
  left.appendChild(renderEquityHeader()); // shown/hidden based on template
  left.appendChild(renderAuthorCard());
  left.appendChild(renderSectionCard("Investment thesis", "investmentThesis", {
    required: true,
    helper: "Why now, what’s mispriced, what changes. Keep it tight.",
    rows: 6,
  }));
  left.appendChild(renderSectionCard("Key takeaways", "keyTakeaways", {
    required: true,
    helper: "3–7 lines max. Each line stands alone (PM-ready).",
    rows: 5,
    placeholder: "- Takeaway 1\n- Takeaway 2\n- Takeaway 3",
  }));
  left.appendChild(renderSectionCard("Company analysis", "companyAnalysis", {
    required: true,
    helper: "Business model, drivers, unit economics, downside, catalysts.",
    rows: 10,
  }));
  left.appendChild(renderSectionCard("Valuation", "valuation", {
    required: true,
    helper: "Method, key bridges, sensitivity, what the market is assuming.",
    rows: 8,
  }));
  left.appendChild(renderSectionCard("Key assumptions", "keyAssumptions", {
    required: false,
    helper: "Revenue, margins, WACC, terminal, FX, commodities, etc.",
    rows: 6,
  }));
  left.appendChild(renderSectionCard("Scenarios", "scenarios", {
    required: false,
    helper: "Base / bull / bear — triggers and probabilities.",
    rows: 6,
  }));

  // Bank-style “Risks & Catalysts” row
  const rcRow = document.createElement("div");
  rcRow.className = "rdt-row-2";
  rcRow.appendChild(renderSectionCard("Catalysts", "catalysts", {
    required: false,
    helper: "What makes the market re-rate? Timeline matters.",
    rows: 6,
    compact: true,
  }));
  rcRow.appendChild(renderSectionCard("Risks", "risks", {
    required: false,
    helper: "What breaks the thesis? Be explicit.",
    rows: 6,
    compact: true,
  }));
  left.appendChild(rcRow);

  left.appendChild(renderSectionCard("Appendix", "appendix", {
    required: false,
    helper: "Extra detail, comps, data notes, disclosures (if needed).",
    rows: 7,
  }));

  left.appendChild(renderSectionCard("The Cordoba view", "cordobaView", {
    required: true,
    helper: "Plain English. Positioning, risk, and what we’d do next.",
    rows: 7,
  }));

  left.appendChild(renderFiguresCard());

  // Build right (export + validation + quick actions)
  right.appendChild(renderExportCard());
  right.appendChild(renderValidationCard());
  right.appendChild(renderShortcutsCard());

  root.appendChild(shell);

  // Bind + initial render updates
  bindEvents();
  refreshAllComputed();
  renderTopBarMeta();
  startAutosaveLoop();
}

function renderHeaderGrid() {
  const card = document.createElement("div");
  card.className = "rdt-card";

  card.innerHTML = `
    <div class="rdt-card-head">
      <div class="rdt-card-title">Note builder</div>
      <div class="rdt-card-sub">Structured like a bank research workstation. Export matches the on-screen sections.</div>
    </div>

    <div class="rdt-grid-3">
      <div class="rdt-field">
        <label>Mode</label>
        <div class="rdt-seg" role="tablist" aria-label="Audience mode">
          ${RDT.audiences
            .map(
              (a) => `
              <button type="button" class="rdt-seg-btn" data-seg="audience" data-value="${a.value}">
                ${a.label}
              </button>`
            )
            .join("")}
        </div>
      </div>

      <div class="rdt-field">
        <label>Template</label>
        <select id="template">
          ${RDT.templates.map((t) => `<option value="${t.value}">${t.label}</option>`).join("")}
        </select>
      </div>

      <div class="rdt-field">
        <label>Note type</label>
        <select id="noteType">
          ${RDT.noteTypes.map((n) => `<option value="${n.value}">${n.label}</option>`).join("")}
        </select>
      </div>
    </div>

    <div class="rdt-grid-2" style="margin-top:12px;">
      <div class="rdt-field">
        <label>Topic <span class="rdt-req">*</span></label>
        <input id="topic" type="text" placeholder="e.g., Vietnam consumer demand, Indonesia FX, uranium supply..." />
      </div>

      <div class="rdt-field">
        <label>Title <span class="rdt-req">*</span></label>
        <input id="title" type="text" placeholder="Investor-ready title (tight, specific)" />
      </div>
    </div>

    <div class="rdt-grid-3" style="margin-top:12px;">
      <div class="rdt-field">
        <label>Status <span class="rdt-req">*</span></label>
        <select id="status">
          <option value="DRAFT">Draft</option>
          <option value="READY">Ready</option>
          <option value="FINAL">Final</option>
        </select>
      </div>

      <div class="rdt-field">
        <label>Reviewed by</label>
        <input id="reviewedBy" type="text" placeholder="Optional (Lead / Editor / IC)" />
      </div>

      <div class="rdt-field">
        <label>Change note</label>
        <input id="changeNote" type="text" placeholder="Optional (e.g., Updated data / Revised view)" />
      </div>
    </div>

    <div class="rdt-inline" style="margin-top:12px;">
      <label class="rdt-check"><input id="bumpMajorOnExport" type="checkbox" /> Bump major version on export</label>
      <label class="rdt-check"><input id="draftWatermark" type="checkbox" /> Draft watermark ON</label>
    </div>
  `;

  return card;
}

function renderEquityHeader() {
  const card = document.createElement("div");
  card.className = "rdt-card";
  card.id = "equityHeaderCard";

  card.innerHTML = `
    <div class="rdt-card-head">
      <div class="rdt-card-title">Equity tear sheet header</div>
      <div class="rdt-card-sub">This block drives the Word “tear sheet” layout (top of the note).</div>
    </div>

    <div class="rdt-grid-3">
      <div class="rdt-field">
        <label>Company name <span class="rdt-req">*</span></label>
        <input id="companyName" type="text" placeholder="e.g., Masan Group" />
      </div>

      <div class="rdt-field">
        <label>Ticker <span class="rdt-req">*</span></label>
        <input id="ticker" type="text" placeholder="e.g., AAPL" />
      </div>

      <div class="rdt-field">
        <label>Exchange</label>
        <input id="exchange" type="text" placeholder="e.g., NASDAQ / HOSE" />
      </div>
    </div>

    <div class="rdt-grid-4" style="margin-top:12px;">
      <div class="rdt-field">
        <label>Rating <span class="rdt-req">*</span></label>
        <select id="rating">
          ${RDT.ratings.map((r) => `<option value="${r.value}">${r.label}</option>`).join("")}
        </select>
      </div>

      <div class="rdt-field">
        <label>Target price</label>
        <input id="targetPrice" type="text" inputmode="decimal" placeholder="e.g., 210" />
      </div>

      <div class="rdt-field">
        <label>Current price</label>
        <input id="currentPrice" type="text" inputmode="decimal" placeholder="e.g., 192" />
      </div>

      <div class="rdt-field">
        <label>Upside (%)</label>
        <input id="upsidePct" type="text" inputmode="decimal" placeholder="Auto / manual" />
      </div>
    </div>

    <div class="rdt-grid-4" style="margin-top:12px;">
      <div class="rdt-field">
        <label>Sector</label>
        <input id="sector" type="text" placeholder="e.g., Consumer / Tech" />
      </div>

      <div class="rdt-field">
        <label>Industry</label>
        <input id="industry" type="text" placeholder="e.g., Retail / Semis" />
      </div>

      <div class="rdt-field">
        <label>Country</label>
        <input id="country" type="text" placeholder="e.g., Vietnam" />
      </div>

      <div class="rdt-field">
        <label>Currency</label>
        <input id="currency" type="text" placeholder="USD" />
      </div>
    </div>

    <div class="rdt-divider"></div>

    <div class="rdt-card-title" style="font-size:13px; opacity:.85; margin-bottom:8px;">Key stats (optional)</div>

    <div class="rdt-grid-4">
      <div class="rdt-field"><label>Market cap</label><input id="marketCap" type="text" placeholder="e.g., 3.2T" /></div>
      <div class="rdt-field"><label>3M avg vol</label><input id="avgVol3m" type="text" placeholder="e.g., 62M" /></div>
      <div class="rdt-field"><label>Fwd P/E</label><input id="peFwd" type="text" placeholder="e.g., 27x" /></div>
      <div class="rdt-field"><label>EV/EBITDA</label><input id="evEbitda" type="text" placeholder="e.g., 18x" /></div>
    </div>

    <div class="rdt-grid-4" style="margin-top:12px;">
      <div class="rdt-field"><label>Dividend yield</label><input id="dividendYield" type="text" placeholder="e.g., 0.5%" /></div>
      <div class="rdt-field"><label>Beta</label><input id="beta" type="text" placeholder="e.g., 1.2" /></div>
      <div class="rdt-field"><label>Shares out</label><input id="sharesOut" type="text" placeholder="e.g., 15.5B" /></div>
      <div class="rdt-field"><label>Float (%)</label><input id="floatPct" type="text" placeholder="e.g., 98%" /></div>
    </div>
  `;

  return card;
}

function renderAuthorCard() {
  const card = document.createElement("div");
  card.className = "rdt-card";

  card.innerHTML = `
    <div class="rdt-card-head">
      <div class="rdt-card-title">Author</div>
      <div class="rdt-card-sub">Shown in the Word header. Client-safe mode strips phone/email by default.</div>
    </div>

    <div class="rdt-grid-4">
      <div class="rdt-field">
        <label>Last name</label>
        <input id="authorLastName" type="text" placeholder="Rahman" />
      </div>

      <div class="rdt-field">
        <label>First name</label>
        <input id="authorFirstName" type="text" placeholder="Nadir" />
      </div>

      <div class="rdt-field">
        <label>Role</label>
        <input id="authorRole" type="text" placeholder="Research Analyst" />
      </div>

      <div class="rdt-field">
        <label>Phone</label>
        <div class="rdt-split">
          <select id="authorPhoneCountry" aria-label="Country code">
            <option value="+44">+44</option>
            <option value="+1">+1</option>
            <option value="+971">+971</option>
            <option value="+65">+65</option>
          </select>
          <input id="authorPhone" type="tel" placeholder="National number" />
        </div>
        <div class="rdt-micro">Client-safe: <strong id="authorPhoneSafe">—</strong></div>
      </div>
    </div>

    <div class="rdt-grid-2" style="margin-top:12px;">
      <div class="rdt-field">
        <label>Email</label>
        <input id="authorEmail" type="email" placeholder="Optional" />
      </div>
      <div class="rdt-field">
        <label>Tags</label>
        <input id="tags" type="text" placeholder="Optional (comma-separated): Asia, Consumer, FX..." />
      </div>
    </div>

    <div class="rdt-divider"></div>

    <div class="rdt-inline" style="justify-content:space-between;">
      <div class="rdt-card-title" style="font-size:13px; opacity:.85;">Co-authors</div>
      <button type="button" class="rdt-btn rdt-btn-ghost" id="btnAddCoAuthor">Add co-author</button>
    </div>

    <div id="coAuthorList" class="rdt-coauthors" style="margin-top:10px;"></div>
  `;

  return card;
}

function renderSectionCard(title, key, opts = {}) {
  const required = !!opts.required;
  const helper = opts.helper || "";
  const rows = opts.rows || 7;
  const placeholder = opts.placeholder || "";
  const compact = !!opts.compact;

  const card = document.createElement("div");
  card.className = "rdt-card";
  card.dataset.sectionKey = key;

  card.innerHTML = `
    <div class="rdt-card-head ${compact ? "rdt-card-head-compact" : ""}">
      <div class="rdt-card-title">
        ${title} ${required ? `<span class="rdt-req">*</span>` : ""}
      </div>
      <div class="rdt-wc"><span id="wc_${key}">0</span> words</div>
    </div>
    ${helper ? `<div class="rdt-helper">${helper}</div>` : ""}
    <textarea id="${key}" rows="${rows}" placeholder="${placeholder}"></textarea>
  `;

  return card;
}

function renderFiguresCard() {
  const card = document.createElement("div");
  card.className = "rdt-card";

  card.innerHTML = `
    <div class="rdt-card-head">
      <div class="rdt-card-title">Figures and charts</div>
      <div class="rdt-card-sub">Upload at the end. Reference in text as (Figure 1), (Figure 2), etc.</div>
    </div>

    <div class="rdt-grid-2">
      <div class="rdt-field">
        <label>Upload images</label>
        <input id="imageFiles" type="file" accept="image/*" multiple />
        <div class="rdt-helper">Keep file names clean. Add captions so the Word export is publication-ready.</div>
      </div>

      <div>
        <div class="rdt-card-title" style="font-size:13px; opacity:.85; margin-bottom:8px;">Figure list</div>
        <div id="figureList" class="rdt-figures"></div>
      </div>
    </div>
  `;

  return card;
}

function renderExportCard() {
  const card = document.createElement("div");
  card.className = "rdt-card rdt-sticky";

  card.innerHTML = `
    <div class="rdt-card-head">
      <div class="rdt-card-title">Export</div>
      <div class="rdt-card-sub">Export mirrors bank research formatting: tear sheet + structured sections.</div>
    </div>

    <label class="rdt-check">
      <input id="confirmReady" type="checkbox" />
      I confirm the draft is ready to export.
    </label>

    <div class="rdt-actions">
      <button type="button" class="rdt-btn rdt-btn-primary" id="btnExport">Generate Word Document</button>
      <button type="button" class="rdt-btn rdt-btn-ghost" id="btnJumpMissing">Jump to first missing</button>
    </div>

    <div class="rdt-kpis">
      <div class="rdt-kpi">
        <div class="rdt-kpi-label">Completion</div>
        <div class="rdt-kpi-value" id="kpiCompletion">0%</div>
      </div>
      <div class="rdt-kpi">
        <div class="rdt-kpi-label">Validation</div>
        <div class="rdt-kpi-value" id="kpiValidation">—</div>
      </div>
    </div>
  `;

  return card;
}

function renderValidationCard() {
  const card = document.createElement("div");
  card.className = "rdt-card";

  card.innerHTML = `
    <div class="rdt-card-head">
      <div class="rdt-card-title">Checks</div>
      <div class="rdt-card-sub">Required fields and quick navigation.</div>
    </div>

    <div id="validationList" class="rdt-validation">—</div>
  `;

  return card;
}

function renderShortcutsCard() {
  const card = document.createElement("div");
  card.className = "rdt-card";

  card.innerHTML = `
    <div class="rdt-card-head">
      <div class="rdt-card-title">Shortcuts</div>
      <div class="rdt-card-sub">Built for analyst flow.</div>
    </div>

    <div class="rdt-shortcuts">
      <div><span class="rdt-chip">Ctrl</span> + <span class="rdt-chip">S</span> Save</div>
      <div><span class="rdt-chip">Ctrl</span> + <span class="rdt-chip">E</span> Export</div>
      <div><span class="rdt-chip">Ctrl</span> + <span class="rdt-chip">K</span> Jump to first missing</div>
    </div>
  `;

  return card;
}

/* ================================
   Bind events + model syncing
================================ */
function bindEvents() {
  // Top bar actions
  $("#btnClearAutosave")?.addEventListener("click", () => {
    clearAutosave();
    toast("Autosave cleared.");
  });

  $("#btnReset")?.addEventListener("click", () => {
    if (!confirm("Reset the form? This will clear current inputs (autosave remains unless cleared).")) return;
    AppState.data = mergeDefaults(AppState.data, {
      // Keep audience/template defaults
      audience: "internal",
      template: "equity",
      noteType: "equity",
      status: "DRAFT",
      reviewedBy: "",
      changeNote: "",
      bumpMajorOnExport: false,
      topic: "",
      title: "",
      tags: "",

      authorFirstName: "",
      authorLastName: "",
      authorRole: "Research Analyst",
      authorEmail: "",
      authorPhoneCountry: "+44",
      authorPhone: "",
      coAuthors: [],

      companyName: "",
      ticker: "",
      exchange: "",
      country: "",
      sector: "",
      industry: "",
      currency: "USD",
      rating: "HOLD",
      targetPrice: "",
      currentPrice: "",
      upsidePct: "",
      marketCap: "",
      sharesOut: "",
      floatPct: "",
      avgVol3m: "",
      beta: "",
      peFwd: "",
      evEbitda: "",
      dividendYield: "",

      investmentThesis: "",
      keyTakeaways: "",
      companyAnalysis: "",
      valuation: "",
      keyAssumptions: "",
      scenarios: "",
      catalysts: "",
      risks: "",
      appendix: "",
      cordobaView: "",

      imageFiles: [],
      figureCaptions: {},

      confirmReady: false,
      emailRouting: false,
      draftWatermark: true,
    });
    AppState.dirty = true;
    mount(); // rerender
    toast("Form reset.");
  });

  $("#btnLoadSession")?.addEventListener("click", () => {
    const sid = prompt("Paste a session ID to load (e.g., D034C7F1). Leave blank to cancel.");
    if (!sid) return;
    const ok = loadFromLocalStorage(sid.trim());
    if (!ok) return alert("No autosave found for that session ID.");
    mount();
    toast("Session loaded.");
  });

  // Segment buttons
  $$(".rdt-seg-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const seg = btn.dataset.seg;
      const val = btn.dataset.value;
      if (seg === "audience") setField("audience", val);
      refreshAllComputed();
      updateSegUI();
    });
  });

  // Selects/inputs/textareas
  const bind = (id, key = id, transform = null) => {
    const el = document.getElementById(id);
    if (!el) return;
    const handler = debounce(() => {
      let v = el.type === "checkbox" ? el.checked : el.value;
      if (transform) v = transform(v);
      setField(key, v);
      refreshAllComputed();
    }, 50);
    el.addEventListener("input", handler);
    el.addEventListener("change", handler);
  };

  bind("template");
  bind("noteType");
  bind("topic");
  bind("title");
  bind("status");
  bind("reviewedBy");
  bind("changeNote");
  bind("bumpMajorOnExport", "bumpMajorOnExport");
  bind("draftWatermark", "draftWatermark");

  bind("companyName");
  bind("ticker", "ticker", (v) => safeTrim(v).toUpperCase());
  bind("exchange");
  bind("rating");
  bind("targetPrice");
  bind("currentPrice");
  bind("upsidePct");
  bind("sector");
  bind("industry");
  bind("country");
  bind("currency", "currency", (v) => safeTrim(v).toUpperCase());
  bind("marketCap");
  bind("avgVol3m");
  bind("peFwd");
  bind("evEbitda");
  bind("dividendYield");
  bind("beta");
  bind("sharesOut");
  bind("floatPct");

  bind("authorLastName");
  bind("authorFirstName");
  bind("authorRole");
  bind("authorEmail");
  bind("authorPhoneCountry");
  bind("authorPhone");
  bind("tags");

  bind("investmentThesis");
  bind("keyTakeaways");
  bind("companyAnalysis");
  bind("valuation");
  bind("keyAssumptions");
  bind("scenarios");
  bind("catalysts");
  bind("risks");
  bind("appendix");
  bind("cordobaView");

  bind("confirmReady", "confirmReady");

  // Co-authors
  $("#btnAddCoAuthor")?.addEventListener("click", () => {
    AppState.data.coAuthors.push({
      id: uid(8),
      name: "",
      role: "Research Analyst",
      email: "",
    });
    AppState.dirty = true;
    renderCoAuthors();
  });

  // Image uploads
  $("#imageFiles")?.addEventListener("change", async (e) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;

    for (const f of files) {
      const dataUrl = await fileToDataURL(f);
      AppState.data.imageFiles.push({
        id: uid(10),
        name: f.name,
        type: f.type,
        size: f.size,
        dataUrl,
      });
    }
    AppState.dirty = true;
    renderFigures();
    refreshAllComputed();
    toast(`${files.length} figure(s) added.`);
    // reset input so same file can be re-added if needed
    e.target.value = "";
  });

  // Export
  $("#btnExport")?.addEventListener("click", async () => {
    try {
      const v = validate();
      if (!v.ok) {
        toast("Missing required fields. Jumping to first missing…");
        jumpToFirstMissing();
        return;
      }
      if (!AppState.data.confirmReady) {
        toast("Please confirm the draft is ready to export.");
        $("#confirmReady")?.focus();
        return;
      }
      await exportWord();
    } catch (err) {
      console.error(err);
      alert("Export failed. Check console for details.");
    }
  });

  $("#btnJumpMissing")?.addEventListener("click", jumpToFirstMissing);

  // Keyboard shortcuts
  document.addEventListener("keydown", (e) => {
    const isMac = navigator.platform.toUpperCase().includes("MAC");
    const ctrl = isMac ? e.metaKey : e.ctrlKey;
    if (!ctrl) return;

    const k = e.key.toLowerCase();
    if (k === "s") {
      e.preventDefault();
      saveToLocalStorage();
      toast("Saved.");
    }
    if (k === "e") {
      e.preventDefault();
      $("#btnExport")?.click();
    }
    if (k === "k") {
      e.preventDefault();
      jumpToFirstMissing();
    }
  });

  // Initial data -> UI sync
  hydrateUIFromState();
  updateSegUI();
  renderCoAuthors();
  renderFigures();
  toggleEquityVisibility();
}

function setField(key, value) {
  AppState.data[key] = value;
  AppState.dirty = true;

  // Auto upside calc if prices given (unless user manually overrides upsidePct)
  if ((key === "targetPrice" || key === "currentPrice") && safeTrim(AppState.data.upsidePct) === "") {
    const tp = parseFloat(String(AppState.data.targetPrice).replace(/[^0-9.]/g, ""));
    const cp = parseFloat(String(AppState.data.currentPrice).replace(/[^0-9.]/g, ""));
    if (isFinite(tp) && isFinite(cp) && cp !== 0) {
      const up = ((tp / cp) - 1) * 100;
      AppState.data.upsidePct = `${up.toFixed(1)}%`;
      const el = $("#upsidePct");
      if (el) el.value = AppState.data.upsidePct;
    }
  }

  // Phone safe display
  if (key === "authorPhone" || key === "authorPhoneCountry" || key === "audience") {
    renderPhoneSafe();
  }

  // Template switches
  if (key === "template" || key === "noteType") {
    toggleEquityVisibility();
  }
}

function hydrateUIFromState() {
  // Set basic inputs
  const setVal = (id, val) => {
    const el = document.getElementById(id);
    if (!el) return;
    if (el.type === "checkbox") el.checked = !!val;
    else el.value = val ?? "";
  };

  // Common
  setVal("template", AppState.data.template);
  setVal("noteType", AppState.data.noteType);
  setVal("topic", AppState.data.topic);
  setVal("title", AppState.data.title);
  setVal("status", AppState.data.status);
  setVal("reviewedBy", AppState.data.reviewedBy);
  setVal("changeNote", AppState.data.changeNote);
  setVal("bumpMajorOnExport", AppState.data.bumpMajorOnExport);
  setVal("draftWatermark", AppState.data.draftWatermark);

  // Equity
  setVal("companyName", AppState.data.companyName);
  setVal("ticker", AppState.data.ticker);
  setVal("exchange", AppState.data.exchange);
  setVal("rating", AppState.data.rating);
  setVal("targetPrice", AppState.data.targetPrice);
  setVal("currentPrice", AppState.data.currentPrice);
  setVal("upsidePct", AppState.data.upsidePct);
  setVal("sector", AppState.data.sector);
  setVal("industry", AppState.data.industry);
  setVal("country", AppState.data.country);
  setVal("currency", AppState.data.currency);
  setVal("marketCap", AppState.data.marketCap);
  setVal("avgVol3m", AppState.data.avgVol3m);
  setVal("peFwd", AppState.data.peFwd);
  setVal("evEbitda", AppState.data.evEbitda);
  setVal("dividendYield", AppState.data.dividendYield);
  setVal("beta", AppState.data.beta);
  setVal("sharesOut", AppState.data.sharesOut);
  setVal("floatPct", AppState.data.floatPct);

  // Author
  setVal("authorLastName", AppState.data.authorLastName);
  setVal("authorFirstName", AppState.data.authorFirstName);
  setVal("authorRole", AppState.data.authorRole);
  setVal("authorEmail", AppState.data.authorEmail);
  setVal("authorPhoneCountry", AppState.data.authorPhoneCountry);
  setVal("authorPhone", AppState.data.authorPhone);
  setVal("tags", AppState.data.tags);

  // Sections
  setVal("investmentThesis", AppState.data.investmentThesis);
  setVal("keyTakeaways", AppState.data.keyTakeaways);
  setVal("companyAnalysis", AppState.data.companyAnalysis);
  setVal("valuation", AppState.data.valuation);
  setVal("keyAssumptions", AppState.data.keyAssumptions);
  setVal("scenarios", AppState.data.scenarios);
  setVal("catalysts", AppState.data.catalysts);
  setVal("risks", AppState.data.risks);
  setVal("appendix", AppState.data.appendix);
  setVal("cordobaView", AppState.data.cordobaView);

  // Export
  setVal("confirmReady", AppState.data.confirmReady);

  renderPhoneSafe();
}

function updateSegUI() {
  const current = AppState.data.audience;
  $$(".rdt-seg-btn[data-seg='audience']").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.value === current);
  });
}

function toggleEquityVisibility() {
  const isEquity = AppState.data.template === "equity" || AppState.data.noteType === "equity";
  const card = $("#equityHeaderCard");
  if (card) card.style.display = isEquity ? "" : "none";
}

/* ================================
   Co-authors UI
================================ */
function renderCoAuthors() {
  const wrap = $("#coAuthorList");
  if (!wrap) return;
  wrap.innerHTML = "";

  const items = AppState.data.coAuthors || [];
  if (!items.length) {
    wrap.innerHTML = `<div class="rdt-empty">No co-authors added.</div>`;
    return;
  }

  for (const ca of items) {
    const row = document.createElement("div");
    row.className = "rdt-coauthor";

    row.innerHTML = `
      <div class="rdt-coauthor-main">
        <div class="rdt-coauthor-name">
          <input class="rdt-coauthor-input" data-ca-id="${ca.id}" data-ca-field="name" type="text" placeholder="Name" value="${escapeHtml(ca.name)}" />
        </div>
        <div class="rdt-coauthor-meta">
          <input class="rdt-coauthor-input" data-ca-id="${ca.id}" data-ca-field="role" type="text" placeholder="Role" value="${escapeHtml(ca.role || "")}" />
          <input class="rdt-coauthor-input" data-ca-id="${ca.id}" data-ca-field="email" type="email" placeholder="Email (optional)" value="${escapeHtml(ca.email || "")}" />
        </div>
      </div>
      <div class="rdt-coauthor-actions">
        <button class="rdt-btn rdt-btn-ghost" data-ca-remove="${ca.id}" type="button">Remove</button>
      </div>
    `;

    wrap.appendChild(row);
  }

  // Bind inputs
  $$(".rdt-coauthor-input").forEach((inp) => {
    inp.addEventListener(
      "input",
      debounce(() => {
        const id = inp.dataset.caId;
        const field = inp.dataset.caField;
        const item = AppState.data.coAuthors.find((x) => x.id === id);
        if (!item) return;
        item[field] = inp.value;
        AppState.dirty = true;
      }, 50)
    );
  });

  // Bind remove
  $$("[data-ca-remove]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const id = btn.dataset.caRemove;
      AppState.data.coAuthors = AppState.data.coAuthors.filter((x) => x.id !== id);
      AppState.dirty = true;
      renderCoAuthors();
    });
  });
}

/* ================================
   Figures UI
================================ */
function renderFigures() {
  const list = $("#figureList");
  if (!list) return;
  list.innerHTML = "";

  const figs = AppState.data.imageFiles || [];
  if (!figs.length) {
    list.innerHTML = `<div class="rdt-empty">No figures uploaded yet.</div>`;
    return;
  }

  figs.forEach((f, idx) => {
    const n = idx + 1;
    const cap = AppState.data.figureCaptions?.[f.id] || "";
    const row = document.createElement("div");
    row.className = "rdt-figure";

    row.innerHTML = `
      <div class="rdt-figure-left">
        <div class="rdt-figure-tag">Figure ${n}</div>
        <div class="rdt-figure-name">${escapeHtml(f.name)}</div>
        <div class="rdt-figure-meta">${formatBytes(f.size)} • ${escapeHtml(f.type || "image")}</div>
        <input class="rdt-figure-caption" data-fig-id="${f.id}" type="text" placeholder="Caption (appears under the figure in Word)" value="${escapeHtml(cap)}" />
      </div>
      <div class="rdt-figure-actions">
        <button class="rdt-btn rdt-btn-ghost" type="button" data-fig-remove="${f.id}">Remove</button>
      </div>
    `;

    list.appendChild(row);
  });

  $$(".rdt-figure-caption").forEach((inp) => {
    inp.addEventListener(
      "input",
      debounce(() => {
        const id = inp.dataset.figId;
        AppState.data.figureCaptions[id] = inp.value;
        AppState.dirty = true;
      }, 80)
    );
  });

  $$("[data-fig-remove]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const id = btn.dataset.figRemove;
      AppState.data.imageFiles = AppState.data.imageFiles.filter((x) => x.id !== id);
      delete AppState.data.figureCaptions[id];
      AppState.dirty = true;
      renderFigures();
      refreshAllComputed();
    });
  });
}

function formatBytes(bytes) {
  const b = Number(bytes || 0);
  if (b < 1024) return `${b} B`;
  const kb = b / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KB`;
  const mb = kb / 1024;
  return `${mb.toFixed(1)} MB`;
}

function fileToDataURL(file) {
  return new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = () => res(fr.result);
    fr.onerror = rej;
    fr.readAsDataURL(file);
  });
}

function renderPhoneSafe() {
  const el = $("#authorPhoneSafe");
  if (!el) return;

  // Client-safe => strip phone entirely
  if (AppState.data.audience === "client_safe") {
    el.textContent = "—";
    return;
  }

  const country = safeTrim(AppState.data.authorPhoneCountry) || "";
  const phone = safeTrim(AppState.data.authorPhone) || "";
  if (!phone) {
    el.textContent = "—";
    return;
  }

  // Mask middle digits
  const digits = phone.replace(/[^\d]/g, "");
  if (digits.length <= 4) {
    el.textContent = `${country}${digits}`;
    return;
  }
  const head = digits.slice(0, 2);
  const tail = digits.slice(-2);
  el.textContent = `${country}${head}${"•".repeat(Math.max(4, digits.length - 4))}${tail}`;
}

/* ================================
   Computed UI (words, validation, KPIs)
================================ */
function refreshAllComputed() {
  // Word counts
  const keys = [
    "investmentThesis",
    "keyTakeaways",
    "companyAnalysis",
    "valuation",
    "keyAssumptions",
    "scenarios",
    "catalysts",
    "risks",
    "appendix",
    "cordobaView",
  ];
  keys.forEach((k) => {
    const el = $(`#wc_${k}`);
    if (el) el.textContent = wordCount(AppState.data[k]);
  });

  // Validation + KPIs
  const v = validate();
  const totalReq = v.required.length;
  const miss = v.missing.length;
  const done = totalReq - miss;
  const pct = totalReq ? Math.round((done / totalReq) * 100) : 0;

  const kpiCompletion = $("#kpiCompletion");
  if (kpiCompletion) kpiCompletion.textContent = `${pct}%`;

  const kpiValidation = $("#kpiValidation");
  if (kpiValidation) kpiValidation.textContent = v.ok ? "ready to export" : `${miss} missing`;

  // Render list
  const list = $("#validationList");
  if (list) {
    if (v.ok) {
      list.innerHTML = `<div class="rdt-ok">✓ Validation: ready to export</div>`;
    } else {
      list.innerHTML = `
        <div class="rdt-warn">Missing required fields:</div>
        <ul class="rdt-miss">
          ${v.missing
            .slice(0, 12)
            .map((m) => `<li><a href="#" data-jump="${m}">${escapeHtml(fieldLabel(m))}</a></li>`)
            .join("")}
          ${v.missing.length > 12 ? `<li class="rdt-muted">+ ${v.missing.length - 12} more…</li>` : ""}
        </ul>
      `;

      $$("[data-jump]", list).forEach((a) => {
        a.addEventListener("click", (e) => {
          e.preventDefault();
          focusField(a.dataset.jump);
        });
      });
    }
  }

  // Save (throttled by autosave loop)
  renderTopBarMeta();
}

function renderTopBarMeta() {
  const sid = $("#rdtSession");
  const saved = $("#rdtSavedAt");
  if (sid) sid.textContent = AppState.sessionId;
  if (saved) saved.textContent = AppState.lastSavedAt || (AppState.dirty ? "—" : "—");
}

/* ================================
   Jump / Focus helpers
================================ */
function focusField(key) {
  const id = key;
  const el = document.getElementById(id);
  if (el) {
    el.scrollIntoView({ behavior: "smooth", block: "center" });
    el.focus();
    return true;
  }

  // Map to alternate IDs
  const aliases = {
    rating: "rating",
    companyName: "companyName",
    investmentThesis: "investmentThesis",
    keyTakeaways: "keyTakeaways",
    companyAnalysis: "companyAnalysis",
    valuation: "valuation",
    cordobaView: "cordobaView",
    audience: null,
  };

  const aid = aliases[key];
  if (aid && document.getElementById(aid)) {
    document.getElementById(aid).scrollIntoView({ behavior: "smooth", block: "center" });
    document.getElementById(aid).focus();
    return true;
  }

  // Audience segment: scroll to top card
  if (key === "audience") {
    const topCard = $$(".rdt-card")[0];
    if (topCard) topCard.scrollIntoView({ behavior: "smooth", block: "start" });
    return true;
  }

  return false;
}

function jumpToFirstMissing() {
  const v = validate();
  if (!v.missing.length) {
    toast("All required fields complete.");
    return;
  }
  const first = v.missing[0];
  focusField(first);
}

/* ================================
   Toast (minimal)
================================ */
function toast(msg) {
  let t = document.getElementById("rdtToast");
  if (!t) {
    t = document.createElement("div");
    t.id = "rdtToast";
    t.className = "rdt-toast";
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove("show"), 1600);
}

/* ================================
   HTML escape
================================ */
function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* =========================================================
   Word Export — Equity Research Tear Sheet
   Uses docx (window.docx)
========================================================= */
async function exportWord() {
  if (!window.docx) {
    alert("docx library not found. Please include docx in your HTML before app.js.");
    return;
  }

  const d = getExportData();
  const doc = await buildEquityDocx(d);

  const blob = await window.docx.Packer.toBlob(doc);

  // File naming: CRG_Equity_TICKER_Title_YYYYMMDD
  const date = new Date();
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const base = `CRG_${d.templateLabel}_${d.ticker || "TICKER"}_${d.title || "Note"}_${y}${m}${day}`;
  downloadBlob(blob, `${sanitizeFileName(base)}.docx`);

  // Save after export
  saveToLocalStorage();
  toast("Word document generated.");
}

function getExportData() {
  // Client-safe stripping
  const isClientSafe = AppState.data.audience === "client_safe";

  const author = `${safeTrim(AppState.data.authorFirstName)} ${safeTrim(AppState.data.authorLastName)}`.trim();
  const coAuthors = (AppState.data.coAuthors || [])
    .map((x) => safeTrim(x.name))
    .filter(Boolean);

  const phone = isClientSafe ? "" : `${safeTrim(AppState.data.authorPhoneCountry)} ${safeTrim(AppState.data.authorPhone)}`.trim();
  const email = isClientSafe ? "" : safeTrim(AppState.data.authorEmail);

  return {
    ...AppState.data,
    templateLabel: "Equity",
    author,
    coAuthors,
    phone,
    email,
    dateTimeString: nowStamp(),
    draft: AppState.data.status !== "FINAL" && AppState.data.draftWatermark === true,
  };
}

/* ================================
   docx helpers
================================ */
function linesToParagraphs(text, spacingAfter = 120) {
  const { Paragraph } = window.docx;
  const t = safeTrim(text);
  if (!t) return [new Paragraph({ text: "" })];

  const lines = t.split("\n");
  const out = [];
  for (const line of lines) {
    const raw = line.replace(/\r/g, "");
    if (!raw.trim()) {
      out.push(new Paragraph({ text: "", spacing: { after: spacingAfter } }));
      continue;
    }
    // Bullet detection: -, *, •
    const m = raw.match(/^\s*([-*•])\s+(.*)$/);
    if (m) {
      out.push(
        new Paragraph({
          text: m[2].trim(),
          bullet: { level: 0 },
          spacing: { after: spacingAfter },
        })
      );
    } else {
      out.push(
        new Paragraph({
          text: raw.trim(),
          spacing: { after: spacingAfter },
        })
      );
    }
  }
  return out;
}

function heading(text, level = 2) {
  const { Paragraph, HeadingLevel } = window.docx;
  const map = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
  };
  return new Paragraph({
    text,
    heading: map[level] || HeadingLevel.HEADING_2,
    spacing: { before: 180, after: 120 },
  });
}

function smallMuted(text) {
  const { Paragraph, TextRun } = window.docx;
  return new Paragraph({
    children: [
      new TextRun({
        text,
        size: 20,
        color: "666666",
      }),
    ],
    spacing: { after: 120 },
  });
}

async function imageToDocxImage(dataUrl, maxWidthTwips = 6.8 * 1440) {
  // dataUrl => Uint8Array
  const base64 = dataUrl.split(",")[1] || "";
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);

  // We can’t read actual image dimensions reliably without extra work.
  // Use a sensible default width; docx scales image by width in pixels.
  // docx uses 'transformation' in pixels. We'll approximate:
  // 680px wide is usually safe for A4 with margins.
  const { ImageRun } = window.docx;
  return new ImageRun({
    data: bytes,
    transformation: {
      width: 680,
      height: 380,
    },
  });
}

/* ================================
   Equity Doc Builder (Tear Sheet)
================================ */
async function buildEquityDocx(d) {
  const {
    Document,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    WidthType,
    AlignmentType,
    PageNumber,
    Footer,
    Header,
    BorderStyle,
  } = window.docx;

  const borderNone = {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };

  const gold = "845F0F";

  // ===== Header: Brand + Title + Meta
  const headerPara = new Paragraph({
    children: [
      new TextRun({ text: "Cordoba Research Group", bold: true, size: 22, color: gold }),
      new TextRun({ text: "  ", size: 22 }),
      new TextRun({ text: "Equity Research", size: 22, color: "333333" }),
    ],
    spacing: { after: 120 },
  });

  const metaLine = new Paragraph({
    children: [
      new TextRun({ text: d.dateTimeString, size: 20, color: "666666" }),
      new TextRun({ text: "   |   ", size: 20, color: "BBBBBB" }),
      new TextRun({ text: d.audience === "client_safe" ? "Client-safe" : d.audience === "public" ? "Public" : "Internal", size: 20, color: "666666" }),
      new TextRun({ text: "   |   ", size: 20, color: "BBBBBB" }),
      new TextRun({ text: d.status, size: 20, color: d.status === "FINAL" ? "2E7D32" : "B26A00" }),
    ],
    spacing: { after: 180 },
  });

  const titlePara = new Paragraph({
    children: [
      new TextRun({ text: d.title || "—", bold: true, size: 34, color: "111111" }),
    ],
    spacing: { after: 120 },
  });

  const topicPara = new Paragraph({
    children: [
      new TextRun({ text: "Topic: ", bold: true, size: 22, color: "333333" }),
      new TextRun({ text: d.topic || "—", size: 22, color: "333333" }),
    ],
    spacing: { after: 180 },
  });

  // ===== Tear Sheet Table
  const leftCol = [
    ["Company", d.companyName || "—"],
    ["Ticker", d.ticker ? `${d.ticker}${d.exchange ? ` (${d.exchange})` : ""}` : "—"],
    ["Country / Sector", [d.country, d.sector].filter(Boolean).join(" / ") || "—"],
    ["Industry", d.industry || "—"],
    ["Currency", d.currency || "—"],
  ];

  const rightCol = [
    ["Rating", d.rating || "—"],
    ["Target / Current", [d.targetPrice, d.currentPrice].filter(Boolean).join(" / ") || "—"],
    ["Upside", d.upsidePct || "—"],
    ["Market cap", d.marketCap || "—"],
    ["Key multiples", [d.peFwd ? `P/E: ${d.peFwd}` : "", d.evEbitda ? `EV/EBITDA: ${d.evEbitda}` : ""].filter(Boolean).join("  |  ") || "—"],
  ];

  function tearRow(label, value) {
    return new TableRow({
      children: [
        new TableCell({
          width: { size: 28, type: WidthType.PERCENTAGE },
          borders: borderNone,
          children: [
            new Paragraph({
              children: [new TextRun({ text: label, bold: true, size: 20, color: "333333" })],
              spacing: { after: 60 },
            }),
          ],
        }),
        new TableCell({
          width: { size: 72, type: WidthType.PERCENTAGE },
          borders: borderNone,
          children: [
            new Paragraph({
              children: [new TextRun({ text: value || "—", size: 20, color: "333333" })],
              spacing: { after: 60 },
            }),
          ],
        }),
      ],
    });
  }

  const tearSheet = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: borderNone,
            children: [
              new Paragraph({ children: [new TextRun({ text: "Tear Sheet", bold: true, size: 22, color: gold })], spacing: { after: 120 } }),
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: leftCol.map(([l, v]) => tearRow(l, v)),
              }),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: borderNone,
            children: [
              new Paragraph({ children: [new TextRun({ text: "Recommendation", bold: true, size: 22, color: gold })], spacing: { after: 120 } }),
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: rightCol.map(([l, v]) => tearRow(l, v)),
              }),
            ],
          }),
        ],
      }),
    ],
  });

  // ===== Analyst line
  const analystBits = [];
  if (d.author) analystBits.push(d.author);
  if (safeTrim(d.authorRole)) analystBits.push(safeTrim(d.authorRole));
  const analystLine = new Paragraph({
    children: [
      new TextRun({ text: "Analyst: ", bold: true, size: 20, color: "333333" }),
      new TextRun({ text: analystBits.join(" — ") || "—", size: 20, color: "333333" }),
      ...(d.coAuthors?.length
        ? [
            new TextRun({ text: "   |   ", size: 20, color: "BBBBBB" }),
            new TextRun({ text: "Co-authors: ", bold: true, size: 20, color: "333333" }),
            new TextRun({ text: d.coAuthors.join(", "), size: 20, color: "333333" }),
          ]
        : []),
    ],
    spacing: { after: 180 },
  });

  // ===== Body sections
  const body = [];

  // Draft watermark note (simple, non-invasive)
  if (d.draft) {
    body.push(
      new Paragraph({
        children: [
          new TextRun({ text: "DRAFT — Internal working document", bold: true, color: "B26A00", size: 22 }),
        ],
        spacing: { after: 180 },
      })
    );
  }

  body.push(heading("Investment thesis", 2));
  body.push(...linesToParagraphs(d.investmentThesis, 140));

  body.push(heading("Key takeaways", 2));
  body.push(...linesToParagraphs(d.keyTakeaways, 120));

  body.push(heading("Company analysis", 2));
  body.push(...linesToParagraphs(d.companyAnalysis, 140));

  body.push(heading("Valuation", 2));
  body.push(...linesToParagraphs(d.valuation, 140));

  if (safeTrim(d.keyAssumptions)) {
    body.push(heading("Key assumptions", 3));
    body.push(...linesToParagraphs(d.keyAssumptions, 140));
  }

  if (safeTrim(d.scenarios)) {
    body.push(heading("Scenarios", 3));
    body.push(...linesToParagraphs(d.scenarios, 140));
  }

  // Risks & Catalysts (bank-style)
  if (safeTrim(d.catalysts) || safeTrim(d.risks)) {
    body.push(heading("Catalysts and risks", 2));
    if (safeTrim(d.catalysts)) {
      body.push(heading("Catalysts", 3));
      body.push(...linesToParagraphs(d.catalysts, 120));
    }
    if (safeTrim(d.risks)) {
      body.push(heading("Risks", 3));
      body.push(...linesToParagraphs(d.risks, 120));
    }
  }

  if (safeTrim(d.appendix)) {
    body.push(heading("Appendix", 2));
    body.push(...linesToParagraphs(d.appendix, 140));
  }

  body.push(heading("The Cordoba view", 2));
  body.push(...linesToParagraphs(d.cordobaView, 140));

  // Figures
  if ((d.imageFiles || []).length) {
    body.push(heading("Figures", 2));
    for (let i = 0; i < d.imageFiles.length; i++) {
      const fig = d.imageFiles[i];
      const n = i + 1;
      const cap = (d.figureCaptions && d.figureCaptions[fig.id]) ? d.figureCaptions[fig.id] : fig.name;

      // Insert image
      const imgRun = await imageToDocxImage(fig.dataUrl);
      body.push(
        new Paragraph({
          children: [imgRun],
          spacing: { after: 120 },
          alignment: AlignmentType.CENTER,
        })
      );
      body.push(
        new Paragraph({
          children: [
            new TextRun({ text: `Figure ${n}: `, bold: true, size: 20, color: "333333" }),
            new TextRun({ text: cap || "", size: 20, color: "333333" }),
          ],
          spacing: { after: 240 },
          alignment: AlignmentType.CENTER,
        })
      );
    }
  }

  // ===== Footer with page numbers
  const footer = new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: "Cordoba Research Group", size: 18, color: "777777" }),
          new TextRun({ text: "   |   ", size: 18, color: "BBBBBB" }),
          new TextRun({ text: d.audience === "client_safe" ? "Client-safe" : d.audience === "public" ? "Public" : "Internal", size: 18, color: "777777" }),
          new TextRun({ text: "   |   ", size: 18, color: "BBBBBB" }),
          new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], size: 18, color: "777777" }),
        ],
        alignment: AlignmentType.CENTER,
      }),
    ],
  });

  // ===== Document
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 720, right: 720, bottom: 720, left: 720 }, // 0.5 inch
          },
        },
        headers: {
          default: new Header({
            children: [headerPara],
          }),
        },
        footers: { default: footer },
        children: [
          metaLine,
          titlePara,
          topicPara,
          tearSheet,
          analystLine,
          ...body,
          smallMuted("Disclosures: For internal use unless explicitly marked client-safe/public. Not investment advice."),
        ],
      },
    ],
  });

  return doc;
}

/* =========================================================
   Boot
========================================================= */
(function boot() {
  // Try load current session autosave if present; otherwise start fresh
  loadFromLocalStorage();

  // Mount UI
  mount();

  // Immediately save session header so ID exists in storage (optional)
  if (!AppState.lastSavedAt) {
    saveToLocalStorage();
  }
})();
