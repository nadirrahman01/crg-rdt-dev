/* =========================================================
   Cordoba Research Group — Research Documentation Tool (RDT)
   app.js (Institutional v2.2 — Workspace layout edition)
   ---------------------------------------------------------
   Includes:
   ✅ Existing export logic
   ✅ Autosave
   ✅ Co-author management
   ✅ Equity tear sheet support
   ✅ Top action button wiring
   ✅ Live workspace preview sync
   ========================================================= */

(() => {
  "use strict";

  console.log("RDT workspace app.js loaded (v2.2)");

  // ------------------------------
  // Brand (Cordoba)
  // ------------------------------
  const BRAND = {
    name: "Cordoba Research Group",
    short: "CRG",
    version: "RDT v2.2.0",
    tagline: "Values that bind",
    colors: {
      gold: "9A690F",
      goldDark: "845F0F",
      cream: "FFF7F0",
      ink: "0B0E14",
      muted: "6B7280",
      lightMuted: "9CA3AF",
      border: "E5E7EB",
      rail: "F3F4F6",
      callout: "F6F1E8",
      red: "B91C1C"
    },
    fonts: {
      heading: "Times New Roman",
      body: "Helvetica"
    },
    disclaimers: {
      internal:
        "Internal use only. Outputs are draft research documentation generated from user inputs and third-party market data. Verify all figures, tickers, and assumptions before circulation.",
      publicInfo: "Public Information"
    }
  };

  // ------------------------------
  // Utilities
  // ------------------------------
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const safeTrim = (v) => (v ?? "").toString().trim();
  const digitsOnly = (v) => (v || "").toString().replace(/\D/g, "");

  function naIfBlank(v) {
    const s = safeTrim(v);
    return s ? s : "N/A";
  }

  function safeNum(v) {
    const n = Number(String(v ?? "").replace(/,/g, ""));
    return Number.isFinite(n) ? n : null;
  }

  function pct(x) {
    if (!Number.isFinite(x)) return "—";
    return `${(x * 100).toFixed(1)}%`;
  }

  function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  function showMsg(kind, text) {
    const el = document.getElementById("message");
    if (!el) return;
    el.className = `message ${kind || ""}`.trim();
    el.textContent = text || "";
    el.style.display = text ? "block" : "none";
  }

  function formatDateTime(date) {
    const months = [
      "January","February","March","April","May","June",
      "July","August","September","October","November","December"
    ];
    const month = months[date.getMonth()];
    const day = date.getDate();
    const year = date.getFullYear();
    let hours = date.getHours();
    const minutes = date.getMinutes().toString().padStart(2, "0");
    const ampm = hours >= 12 ? "PM" : "AM";
    hours = hours % 12 || 12;
    return `${day} ${month} ${year} ${hours}:${minutes} ${ampm}`;
  }

  function formatDateShortISO(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  function buildMailto(to, cc, subject, body) {
    const crlfBody = (body || "").replace(/\n/g, "\r\n");
    const parts = [];
    if (cc) parts.push(`cc=${encodeURIComponent(cc)}`);
    parts.push(`subject=${encodeURIComponent(subject || "")}`);
    parts.push(`body=${encodeURIComponent(crlfBody)}`);
    return `mailto:${encodeURIComponent(to)}?${parts.join("&")}`;
  }

  function ccForNoteType(noteTypeRaw) {
    const t = (noteTypeRaw || "").toLowerCase();
    if (t.includes("equity")) return "tommaso@cordobarg.com";
    if (t.includes("macro") || t.includes("market")) return "tim@cordobarg.com";
    if (t.includes("fixed")) return "tim@cordobarg.com";
    if (t.includes("commodity")) return "uhayd@cordobarg.com";
    return "";
  }

  function formatNationalLoose(rawDigits) {
    const d = digitsOnly(rawDigits);
    if (!d) return "";
    const p1 = d.slice(0, 4);
    const p2 = d.slice(4, 7);
    const p3 = d.slice(7, 10);
    const rest = d.slice(10);
    return [p1, p2, p3, rest].filter(Boolean).join(" ");
  }

  function buildInternationalHyphen(ccDigits, nationalDigits) {
    const cc = digitsOnly(ccDigits);
    const nn = digitsOnly(nationalDigits);
    if (!cc && !nn) return "";
    if (cc && !nn) return `${cc}-`;
    if (!cc && nn) return nn;
    return `${cc}-${nn}`;
  }

  // ------------------------------
  // Draft persistence
  // ------------------------------
  const DRAFT_KEY = "crg_rdt_draft_v22";

  const DRAFT_FIELDS = [
    "noteType","title","topic",
    "authorLastName","authorFirstName","authorPhone",
    "authorPhoneCountry","authorPhoneNational",
    "analysis","keyTakeaways","content","cordobaView",
    "ticker","crgRating","targetPrice",
    "modelLink","valuationSummary","keyAssumptions","scenarioNotes"
  ];

  function snapshotDraft() {
    const draft = {};
    DRAFT_FIELDS.forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      draft[id] = el.value ?? "";
    });

    const coAuthors = $$(".coauthor-entry").map(entry => ({
      lastName: safeTrim($(".coauthor-lastname", entry)?.value),
      firstName: safeTrim($(".coauthor-firstname", entry)?.value),
      phone: safeTrim($(".coauthor-phone", entry)?.value),
      cc: safeTrim($(".coauthor-country", entry)?.value),
      local: safeTrim($(".coauthor-phone-local", entry)?.value)
    })).filter(x => x.lastName || x.firstName || x.phone || x.local);

    draft.__coAuthors = coAuthors;
    draft.__chartRange = $("#chartRange")?.value || "";
    draft.__equityStats = equityStats || null;
    draft.__savedAt = new Date().toISOString();
    return draft;
  }

  function setDraftStatus(text) {
    const el = document.getElementById("draftStatus");
    if (el) el.textContent = text || "";
  }

  function saveDraftNow() {
    try {
      localStorage.setItem(DRAFT_KEY, JSON.stringify(snapshotDraft()));
      setDraftStatus("Saved");
    } catch (_) {}
  }

  let draftSaveTimer = null;
  function scheduleDraftSave() {
    setDraftStatus("Saving…");
    clearTimeout(draftSaveTimer);
    draftSaveTimer = setTimeout(() => saveDraftNow(), 350);
  }

  function loadDraft() {
    try {
      const raw = localStorage.getItem(DRAFT_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch (_) {
      return null;
    }
  }

  function clearDraft() {
    try { localStorage.removeItem(DRAFT_KEY); } catch(_) {}
  }

  // ------------------------------
  // Phone wiring (primary + coauthors)
  // ------------------------------
  const authorPhoneCountryEl = document.getElementById("authorPhoneCountry");
  const authorPhoneNationalEl = document.getElementById("authorPhoneNational");
  const authorPhoneHiddenEl = document.getElementById("authorPhone");

  function syncPrimaryPhone() {
    if (!authorPhoneHiddenEl) return;
    const cc = authorPhoneCountryEl ? authorPhoneCountryEl.value : "";
    const nn = digitsOnly(authorPhoneNationalEl ? authorPhoneNationalEl.value : "");
    authorPhoneHiddenEl.value = buildInternationalHyphen(cc, nn);
  }

  function formatPrimaryVisible() {
    if (!authorPhoneNationalEl) return;
    const caret = authorPhoneNationalEl.selectionStart || 0;
    const beforeLen = authorPhoneNationalEl.value.length;

    authorPhoneNationalEl.value = formatNationalLoose(authorPhoneNationalEl.value);

    const afterLen = authorPhoneNationalEl.value.length;
    const delta = afterLen - beforeLen;
    const next = Math.max(0, caret + delta);
    authorPhoneNationalEl.setSelectionRange(next, next);

    syncPrimaryPhone();
  }

  if (authorPhoneNationalEl) {
    authorPhoneNationalEl.addEventListener("input", () => { formatPrimaryVisible(); scheduleDraftSave(); });
    authorPhoneNationalEl.addEventListener("blur", () => { syncPrimaryPhone(); scheduleDraftSave(); });
  }
  if (authorPhoneCountryEl) {
    authorPhoneCountryEl.addEventListener("change", () => { syncPrimaryPhone(); scheduleDraftSave(); });
  }

  // ------------------------------
  // Co-author management
  // ------------------------------
  let coAuthorCount = 0;
  const addCoAuthorBtn = document.getElementById("addCoAuthor");
  const coAuthorsList = document.getElementById("coAuthorsList");

  const countryOptionsHtml = `
    <option value="44" selected>🇬🇧 +44</option>
    <option value="1">🇺🇸 +1</option>
    <option value="353">🇮🇪 +353</option>
    <option value="33">🇫🇷 +33</option>
    <option value="49">🇩🇪 +49</option>
    <option value="31">🇳🇱 +31</option>
    <option value="34">🇪🇸 +34</option>
    <option value="39">🇮🇹 +39</option>
    <option value="971">🇦🇪 +971</option>
    <option value="966">🇸🇦 +966</option>
    <option value="92">🇵🇰 +92</option>
    <option value="880">🇧🇩 +880</option>
    <option value="91">🇮🇳 +91</option>
    <option value="234">🇳🇬 +234</option>
    <option value="254">🇰🇪 +254</option>
    <option value="27">🇿🇦 +27</option>
    <option value="995">🇬🇪 +995</option>
    <option value="">Other</option>
  `;

  function wireCoauthorPhone(coAuthorDiv) {
    const ccEl = $(".coauthor-country", coAuthorDiv);
    const nationalEl = $(".coauthor-phone-local", coAuthorDiv);
    const hiddenEl = $(".coauthor-phone", coAuthorDiv);
    if (!hiddenEl) return;

    function syncHidden() {
      const cc = ccEl ? ccEl.value : "";
      const nn = digitsOnly(nationalEl ? nationalEl.value : "");
      hiddenEl.value = buildInternationalHyphen(cc, nn);
    }

    function formatVisible() {
      if (!nationalEl) return;
      const caret = nationalEl.selectionStart || 0;
      const beforeLen = nationalEl.value.length;

      nationalEl.value = formatNationalLoose(nationalEl.value);

      const afterLen = nationalEl.value.length;
      const delta = afterLen - beforeLen;
      const next = Math.max(0, caret + delta);
      nationalEl.setSelectionRange(next, next);

      syncHidden();
    }

    if (nationalEl) {
      nationalEl.addEventListener("input", () => { formatVisible(); scheduleDraftSave(); updateCompletionMeter(); });
      nationalEl.addEventListener("blur", () => { syncHidden(); scheduleDraftSave(); });
    }
    if (ccEl) {
      ccEl.addEventListener("change", () => { syncHidden(); scheduleDraftSave(); });
    }

    syncHidden();
  }

  function createCoauthorNode() {
    coAuthorCount += 1;
    const div = document.createElement("div");
    div.className = "coauthor-entry";
    div.id = `coauthor-${coAuthorCount}`;

    div.innerHTML = `
      <input type="text" placeholder="Last Name" class="coauthor-lastname" autocomplete="family-name">
      <input type="text" placeholder="First Name" class="coauthor-firstname" autocomplete="given-name">
      <div class="phone-row phone-row--compact">
        <select class="phone-country coauthor-country" aria-label="Country code">
          ${countryOptionsHtml}
        </select>
        <input type="text" placeholder="Phone number" class="phone-number coauthor-phone-local" inputmode="numeric" autocomplete="tel-national">
      </div>
      <input type="text" class="coauthor-phone" style="display:none;">
      <button type="button" class="remove-coauthor" data-remove-id="${coAuthorCount}">Remove</button>
    `;

    const hidden = $(".coauthor-phone", div);
    if (hidden) hidden.required = false;

    wireCoauthorPhone(div);

    ["input","change","keyup"].forEach(evt => {
      div.addEventListener(evt, () => scheduleDraftSave(), { passive: true });
    });

    return div;
  }

  if (addCoAuthorBtn && coAuthorsList) {
    addCoAuthorBtn.addEventListener("click", () => {
      coAuthorsList.appendChild(createCoauthorNode());
      updateCompletionMeter();
      scheduleDraftSave();
      syncWorkspacePreview();
    });

    document.addEventListener("click", (e) => {
      const btn = e.target.closest(".remove-coauthor");
      if (!btn) return;
      const id = btn.getAttribute("data-remove-id");
      const node = document.getElementById(`coauthor-${id}`);
      if (node) node.remove();
      updateCompletionMeter();
      scheduleDraftSave();
      syncWorkspacePreview();
    });
  }

  // ------------------------------
  // Equity section toggle
  // ------------------------------
  const noteTypeEl = document.getElementById("noteType");
  const equitySectionEl = document.getElementById("equitySection");

  function toggleEquitySection() {
    if (!noteTypeEl || !equitySectionEl) return;
    equitySectionEl.style.display = (noteTypeEl.value === "Equity Research") ? "block" : "none";
  }

  function isEquityMode() {
    return !!(noteTypeEl && noteTypeEl.value === "Equity Research");
  }

  if (noteTypeEl && equitySectionEl) {
    noteTypeEl.addEventListener("change", () => {
      toggleEquitySection();
      updateCompletionMeter();
      scheduleDraftSave();
      syncWorkspacePreview();
    });
    toggleEquitySection();
  }

  // ------------------------------
  // Completion meter
  // ------------------------------
  const completionTextEl = document.getElementById("completionText");
  const completionBarEl = document.getElementById("completionBar");

  function isFilled(el) {
    if (!el) return false;
    if (el.type === "file") return el.files && el.files.length > 0;
    return safeTrim(el.value).length > 0;
  }

  const baseCoreIds = [
    "noteType","title","topic",
    "authorLastName","authorFirstName",
    "keyTakeaways","analysis"
  ];
  const equityCoreIds = ["crgRating"];

  function updateCompletionMeter() {
    const ids = isEquityMode() ? baseCoreIds.concat(equityCoreIds) : baseCoreIds;
    let done = 0;
    ids.forEach(id => {
      const el = document.getElementById(id);
      if (isFilled(el)) done += 1;
    });

    const total = ids.length;
    const pctDone = total ? Math.round((done / total) * 100) : 0;

    if (completionTextEl) completionTextEl.textContent = `${done} / ${total} publish-core`;
    if (completionBarEl) completionBarEl.style.width = `${pctDone}%`;

    const bar = completionBarEl?.parentElement;
    if (bar) bar.setAttribute("aria-valuenow", String(pctDone));
  }

  ["input","change","keyup"].forEach(evt => {
    document.addEventListener(evt, (e) => {
      const t = e.target;
      if (!t) return;
      if (t.closest && t.closest("#researchForm")) {
        updateCompletionMeter();
        scheduleDraftSave();
        syncWorkspacePreview();
      }
    }, { passive: true });
  });

  // ------------------------------
  // Attachment summary
  // ------------------------------
  const modelFilesEl = document.getElementById("modelFiles");
  const attachSummaryHeadEl = document.getElementById("attachmentSummaryHead");
  const attachSummaryListEl = document.getElementById("attachmentSummaryList");

  function escapeHtml(s) {
    return (s || "").toString()
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function updateAttachmentSummary() {
    if (!modelFilesEl || !attachSummaryHeadEl || !attachSummaryListEl) return;

    const files = Array.from(modelFilesEl.files || []);
    if (!files.length) {
      attachSummaryHeadEl.textContent = "No files selected";
      attachSummaryListEl.style.display = "none";
      attachSummaryListEl.innerHTML = "";
      return;
    }

    attachSummaryHeadEl.textContent = `${files.length} file${files.length === 1 ? "" : "s"} selected`;
    attachSummaryListEl.style.display = "block";
    attachSummaryListEl.innerHTML = files.map(f => `<div class="attachment-file">${escapeHtml(f.name)}</div>`).join("");
  }

  if (modelFilesEl) {
    modelFilesEl.addEventListener("change", () => {
      updateAttachmentSummary();
      updateCompletionMeter();
      scheduleDraftSave();
      syncWorkspacePreview();
    });
  }

  // ------------------------------
  // Reset
  // ------------------------------
  const resetBtn = document.getElementById("resetFormBtn");
  const formEl = document.getElementById("researchForm");

  function clearChartUI() {
    setText("currentPrice", "—");
    setText("realisedVol", "—");
    setText("rangeReturn", "—");
    setText("upsideToTarget", "—");

    const chartStatus = document.getElementById("chartStatus");
    if (chartStatus) chartStatus.textContent = "";

    if (priceChart) {
      try { priceChart.destroy(); } catch (_) {}
      priceChart = null;
    }

    priceChartImageBytes = null;
    equityStats = { currentPrice: null, realisedVolAnn: null, rangeReturn: null, startPrice: null };
  }

  if (resetBtn && formEl) {
    resetBtn.addEventListener("click", () => {
      const ok = confirm("Reset the form? This clears all fields and removes any saved draft.");
      if (!ok) return;

      formEl.reset();
      if (coAuthorsList) coAuthorsList.innerHTML = "";
      if (modelFilesEl) modelFilesEl.value = "";
      updateAttachmentSummary();
      clearChartUI();
      syncPrimaryPhone();
      toggleEquitySection();
      updateCompletionMeter();
      showMsg("", "");
      clearDraft();
      setDraftStatus("");
      syncWorkspacePreview();
    });
  }

  // ------------------------------
  // Email to CRG
  // ------------------------------
  const emailToCrgBtn = document.getElementById("emailToCrgBtn");

  function buildCrgEmailPayload() {
    const noteType = safeTrim($("#noteType")?.value || "Research Note");
    const title = safeTrim($("#title")?.value || "");
    const topic = safeTrim($("#topic")?.value || "");

    const authorFirstName = safeTrim($("#authorFirstName")?.value || "");
    const authorLastName = safeTrim($("#authorLastName")?.value || "");
    const authorLine = [authorFirstName, authorLastName].filter(Boolean).join(" ").trim();

    const ticker = safeTrim($("#ticker")?.value || "");
    const crgRating = safeTrim($("#crgRating")?.value || "");
    const targetPrice = safeTrim($("#targetPrice")?.value || "");

    const now = new Date();
    const subject = [noteType, formatDateShortISO(now), title ? `— ${title}` : ""].filter(Boolean).join(" ");

    const paragraphs = [];
    paragraphs.push("Hi CRG Research,");
    paragraphs.push("Please find my most recent note attached.");

    const metaLines = [
      `Note type: ${noteType || "N/A"}`,
      title ? `Title: ${title}` : null,
      topic ? `Topic: ${topic}` : null,
      ticker ? `Ticker (Stooq): ${ticker}` : null,
      crgRating ? `CRG Rating: ${crgRating}` : null,
      targetPrice ? `Target Price: ${targetPrice}` : null,
      `Generated: ${formatDateTime(now)}`
    ].filter(Boolean);

    paragraphs.push(metaLines.join("\n"));
    paragraphs.push("Best,");
    paragraphs.push(authorLine || "");

    return { subject, body: paragraphs.join("\n\n"), cc: ccForNoteType(noteType) };
  }

  if (emailToCrgBtn) {
    emailToCrgBtn.addEventListener("click", () => {
      const { subject, body, cc } = buildCrgEmailPayload();
      const to = "research@cordobarg.com";
      window.location.href = buildMailto(to, cc, subject, body);
    });
  }

  // ------------------------------
  // Top button wiring
  // ------------------------------
  const emailToCrgBtnTop = document.getElementById("emailToCrgBtnTop");
  if (emailToCrgBtnTop && emailToCrgBtn) {
    emailToCrgBtnTop.addEventListener("click", () => emailToCrgBtn.click());
  }

  const resetFormBtnTop = document.getElementById("resetFormBtnTop");
  if (resetFormBtnTop && resetBtn) {
    resetFormBtnTop.addEventListener("click", () => resetBtn.click());
  }

  // ------------------------------
  // Price chart + stats
  // ------------------------------
  let priceChart = null;
  let priceChartImageBytes = null;

  let equityStats = {
    currentPrice: null,
    realisedVolAnn: null,
    rangeReturn: null,
    startPrice: null
  };

  const chartStatusEl = document.getElementById("chartStatus");
  const fetchChartBtn = document.getElementById("fetchPriceChart");
  const chartRangeEl = document.getElementById("chartRange");
  const priceChartCanvas = document.getElementById("priceChart");
  const targetPriceEl = document.getElementById("targetPrice");

  function stooqSymbolFromTicker(ticker) {
    const t = safeTrim(ticker);
    if (!t) return null;
    if (t.includes(".")) return t.toLowerCase();
    return `${t.toLowerCase()}.us`;
  }

  function computeStartDate(range) {
    const now = new Date();
    const d = new Date(now);
    if (range === "6mo") d.setMonth(d.getMonth() - 6);
    else if (range === "1y") d.setFullYear(d.getFullYear() - 1);
    else if (range === "2y") d.setFullYear(d.getFullYear() - 2);
    else if (range === "5y") d.setFullYear(d.getFullYear() - 5);
    else d.setFullYear(d.getFullYear() - 1);
    return d;
  }

  function extractStooqCSV(text) {
    const lines = (text || "").split("\n").map(l => l.trim()).filter(Boolean);
    const headerIdx = lines.findIndex(l => l.toLowerCase().startsWith("date,open,high,low,close,volume"));
    if (headerIdx === -1) return null;
    return lines.slice(headerIdx).join("\n");
  }

  async function fetchStooqDaily(symbol) {
    const stooqUrl = `http://stooq.com/q/d/l/?s=${encodeURIComponent(symbol)}&i=d`;
    const proxyUrl = `https://r.jina.ai/${stooqUrl}`;

    const res = await fetch(proxyUrl, { cache: "no-store" });
    if (!res.ok) throw new Error("Could not fetch price data (proxy blocked or down).");

    const rawText = await res.text();
    const csvText = extractStooqCSV(rawText) || rawText;

    const lines = csvText.trim().split("\n");
    if (lines.length < 20) throw new Error("Not enough data returned. Check ticker.");

    const rows = lines.slice(1).map(line => line.split(","));
    const out = rows
      .map(r => ({ date: r[0], close: Number(r[4]) }))
      .filter(x => x.date && Number.isFinite(x.close));

    if (!out.length) throw new Error("No usable price data.");
    return out;
  }

  function renderChart({ labels, values, title }) {
    if (!priceChartCanvas || typeof Chart === "undefined") return;

    if (priceChart) {
      try { priceChart.destroy(); } catch (_) {}
      priceChart = null;
    }

    priceChart = new Chart(priceChartCanvas, {
      type: "line",
      data: {
        labels,
        datasets: [{
          label: title,
          data: values,
          pointRadius: 0,
          borderWidth: 2,
          tension: 0.12
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: false,
        plugins: {
          legend: { display: false },
          tooltip: { intersect: false, mode: "index" }
        },
        scales: {
          x: {
            ticks: { maxTicksLimit: 6, color: "#374151" },
            grid: { color: "rgba(148,163,184,.22)" }
          },
          y: {
            ticks: { maxTicksLimit: 6, color: "#374151" },
            grid: { color: "rgba(148,163,184,.22)" }
          }
        }
      }
    });
  }

  function canvasToPngBytes(canvas) {
    const dataUrl = canvas.toDataURL("image/png");
    const b64 = dataUrl.split(",")[1];
    return Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  }

  function computeDailyReturns(closes) {
    const rets = [];
    for (let i = 1; i < closes.length; i++) {
      const prev = closes[i - 1];
      const cur = closes[i];
      if (prev > 0 && Number.isFinite(prev) && Number.isFinite(cur)) {
        rets.push((cur / prev) - 1);
      }
    }
    return rets;
  }

  function stddev(arr) {
    if (!arr.length) return null;
    const mean = arr.reduce((a, b) => a + b, 0) / arr.length;
    const v = arr.reduce((a, b) => a + (b - mean) ** 2, 0) / (arr.length - 1 || 1);
    return Math.sqrt(v);
  }

  function computeUpsideToTarget(currentPrice, targetPrice) {
    if (!Number.isFinite(currentPrice) || !Number.isFinite(targetPrice) || currentPrice <= 0) return null;
    return (targetPrice / currentPrice) - 1;
  }

  function paintEquityStats() {
    const s = equityStats || {};
    setText("currentPrice", Number.isFinite(s.currentPrice) ? s.currentPrice.toFixed(2) : "—");
    setText("rangeReturn", Number.isFinite(s.rangeReturn) ? pct(s.rangeReturn) : "—");
    setText("realisedVol", Number.isFinite(s.realisedVolAnn) ? pct(s.realisedVolAnn) : "—");

    const tp = safeNum(targetPriceEl?.value);
    const up = computeUpsideToTarget(s.currentPrice, tp);
    setText("upsideToTarget", up === null ? "—" : pct(up));
  }

  if (targetPriceEl) {
    targetPriceEl.addEventListener("input", () => {
      paintEquityStats();
      updateCompletionMeter();
      scheduleDraftSave();
      syncWorkspacePreview();
    });
  }

  async function buildPriceChart() {
    try {
      const tickerVal = safeTrim($("#ticker")?.value || "");
      if (!tickerVal) throw new Error("Enter a ticker first.");

      const range = chartRangeEl ? chartRangeEl.value : "6mo";
      const symbol = stooqSymbolFromTicker(tickerVal);
      if (!symbol) throw new Error("Invalid ticker.");

      if (chartStatusEl) chartStatusEl.textContent = "Fetching price data…";

      const data = await fetchStooqDaily(symbol);

      const start = computeStartDate(range);
      const filtered = data.filter(x => new Date(x.date) >= start);
      if (filtered.length < 10) throw new Error("Not enough data for selected range.");

      const labels = filtered.map(x => x.date);
      const values = filtered.map(x => x.close);

      renderChart({
        labels,
        values,
        title: `${tickerVal.toUpperCase()} Close`
      });

      await new Promise(r => setTimeout(r, 150));
      priceChartImageBytes = canvasToPngBytes(priceChartCanvas);

      const closes = values;
      const currentPrice = closes[closes.length - 1];
      const startPrice = closes[0];
      const rangeReturn = (startPrice && currentPrice) ? (currentPrice / startPrice) - 1 : null;

      const dailyRets = computeDailyReturns(closes);
      const volDaily = stddev(dailyRets);
      const realisedVolAnn = (volDaily !== null) ? volDaily * Math.sqrt(252) : null;

      equityStats.currentPrice = currentPrice;
      equityStats.startPrice = startPrice;
      equityStats.rangeReturn = rangeReturn;
      equityStats.realisedVolAnn = realisedVolAnn;

      paintEquityStats();

      if (chartStatusEl) chartStatusEl.textContent = `✓ Chart ready (${range.toUpperCase()})`;
      scheduleDraftSave();
      syncWorkspacePreview();
    } catch (e) {
      priceChartImageBytes = null;
      equityStats = { currentPrice: null, realisedVolAnn: null, rangeReturn: null, startPrice: null };
      paintEquityStats();
      if (chartStatusEl) chartStatusEl.textContent = `✗ ${e.message}`;
    } finally {
      updateCompletionMeter();
    }
  }

  if (fetchChartBtn) fetchChartBtn.addEventListener("click", buildPriceChart);

  // =========================================================
  // WORD EXPORT (docx)
  // =========================================================

  function ensureLibs() {
    if (typeof docx === "undefined") throw new Error("docx library not loaded. Refresh the page.");
    if (typeof saveAs === "undefined") throw new Error("FileSaver library not loaded. Refresh the page.");
  }

  function validatePublishCore() {
    const missing = [];
    const core = isEquityMode() ? baseCoreIds.concat(equityCoreIds) : baseCoreIds;
    core.forEach(id => {
      const el = document.getElementById(id);
      if (el && !isFilled(el)) missing.push(id);
    });
    return missing;
  }

  function run(text, opts = {}) {
    return new docx.TextRun({
      text: text ?? "",
      font: opts.font ?? BRAND.fonts.body,
      size: opts.size ?? 20,
      bold: !!opts.bold,
      italics: !!opts.italics,
      color: opts.color ?? BRAND.colors.ink
    });
  }

  function para(children, opts = {}) {
    return new docx.Paragraph({
      children: Array.isArray(children) ? children : [children],
      spacing: opts.spacing ?? { after: 140 },
      alignment: opts.align ?? undefined,
      border: opts.border ?? undefined,
      tabStops: opts.tabStops ?? undefined
    });
  }

  function heading(text, size = 34, after = 160) {
    return para(run(text, { font: BRAND.fonts.heading, size, bold: true }), { spacing: { after } });
  }

  function subheading(text, size = 24, after = 120) {
    return para(run(text, { font: BRAND.fonts.heading, size, bold: true }), { spacing: { after } });
  }

  function bodyLine(text, after = 140) {
    return para(run(text, { font: BRAND.fonts.body, size: 20 }), { spacing: { after } });
  }

  function linesToParagraphs(text, spacingAfter = 140) {
    const lines = (text || "").split("\n");
    return lines.map(line => {
      if (line.trim() === "") return para(run(""), { spacing: { after: spacingAfter } });
      return bodyLine(line, spacingAfter);
    });
  }

  function bulletLines(text, spacingAfter = 90) {
    const lines = (text || "").split("\n");
    const bullets = [];
    lines.forEach(line => {
      const t = line.replace(/^[-*•]\s*/, "").trim();
      if (!t) return;
      bullets.push(new docx.Paragraph({
        text: t,
        bullet: { level: 0 },
        spacing: { after: spacingAfter }
      }));
    });
    return bullets.length ? bullets : [bodyLine("—", spacingAfter)];
  }

  function noBorderTable(rows, widthPct = 100) {
    return new docx.Table({
      width: { size: widthPct, type: docx.WidthType.PERCENTAGE },
      borders: {
        top: { style: docx.BorderStyle.NONE },
        bottom: { style: docx.BorderStyle.NONE },
        left: { style: docx.BorderStyle.NONE },
        right: { style: docx.BorderStyle.NONE },
        insideHorizontal: { style: docx.BorderStyle.NONE },
        insideVertical: { style: docx.BorderStyle.NONE }
      },
      rows
    });
  }

  function cell(children, opts = {}) {
    return new docx.TableCell({
      width: opts.width ? { size: opts.width, type: docx.WidthType.PERCENTAGE } : undefined,
      shading: opts.shading ? { fill: opts.shading } : undefined,
      margins: opts.margins ?? { top: 180, bottom: 180, left: 180, right: 180 },
      borders: opts.borders ?? {
        top: { style: docx.BorderStyle.NONE },
        bottom: { style: docx.BorderStyle.NONE },
        left: { style: docx.BorderStyle.NONE },
        right: { style: docx.BorderStyle.NONE }
      },
      verticalAlign: opts.vAlign ?? docx.VerticalAlign.TOP,
      children: Array.isArray(children) ? children : [children]
    });
  }

  function thinRule(after = 160) {
    return new docx.Paragraph({
      border: { bottom: { color: BRAND.colors.border, space: 1, style: docx.BorderStyle.SINGLE, size: 2 } },
      spacing: { after }
    });
  }

  function pageBreak() {
    return para(new docx.PageBreak(), { spacing: { after: 0 } });
  }

  function ratingToDisplay(r) {
    const x = safeTrim(r);
    if (!x) return "—";
    if (x.toLowerCase() === "hold") return "Neutral";
    return x;
  }

  function splitIntoParagraphBlocks(text) {
    const raw = (text || "").replace(/\r/g, "");
    const blocks = raw.split(/\n\s*\n/g).map(b => b.trim()).filter(Boolean);
    return blocks;
  }

  async function addImages(files) {
    const out = [];
    const list = Array.from(files || []);
    for (let i = 0; i < list.length; i++) {
      const file = list[i];
      try {
        const buf = await file.arrayBuffer();
        const cap = file.name.replace(/\.[^/.]+$/, "");
        out.push(
          new docx.Paragraph({
            children: [new docx.ImageRun({ data: buf, transformation: { width: 520, height: 340 } })],
            alignment: docx.AlignmentType.CENTER,
            spacing: { before: 140, after: 70 }
          }),
          new docx.Paragraph({
            children: [run(`Figure ${i + 1}: ${cap}`, { size: 18, italics: true, color: BRAND.colors.muted })],
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 220 }
          })
        );
      } catch (e) {
        console.error("Image error:", e);
      }
    }
    return out;
  }

  function headerTable(noteType, isoDate) {
    const leftText = `${BRAND.short} | ${noteType} | ${isoDate}`;
    return noBorderTable([
      new docx.TableRow({
        children: [
          cell(para(run(leftText, { size: 16, color: BRAND.colors.lightMuted }), { spacing: { after: 0 } }), { width: 70, margins: { top: 80, bottom: 80, left: 0, right: 0 } }),
          cell(para(run(BRAND.short, { size: 16, color: BRAND.colors.lightMuted }), { spacing: { after: 0 }, align: docx.AlignmentType.RIGHT }), { width: 30, margins: { top: 80, bottom: 80, left: 0, right: 0 } })
        ]
      })
    ]);
  }

  function footerLine() {
    return new docx.Paragraph({
      border: { top: { color: BRAND.colors.border, space: 1, style: docx.BorderStyle.SINGLE, size: 2 } },
      spacing: { before: 70, after: 0 }
    });
  }

  function footerTable() {
    return new docx.Paragraph({
      children: [
        run(BRAND.disclaimers.internal, { size: 14, italics: true, color: BRAND.colors.muted }),
        new docx.TextRun({ text: "\t" }),
        new docx.TextRun({
          children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
          size: 14,
          italics: true,
          color: BRAND.colors.muted,
          font: BRAND.fonts.body
        })
      ],
      tabStops: [{ type: docx.TabStopType.RIGHT, position: 9000 }],
      spacing: { after: 0 }
    });
  }

  function equityMasthead(dateLabel) {
    return noBorderTable([
      new docx.TableRow({
        children: [
          cell([
            para(run(BRAND.name.toUpperCase(), { font: BRAND.fonts.heading, size: 28, bold: true }), { spacing: { after: 30 } }),
            para(run(BRAND.tagline, { size: 16, color: BRAND.colors.muted }), { spacing: { after: 0 } })
          ], { width: 68, margins: { top: 0, bottom: 0, left: 0, right: 0 } }),
          cell([
            para(run("INSTITUTIONAL EQUITY RESEARCH", { size: 16, bold: true, color: BRAND.colors.red }), { spacing: { after: 30 }, align: docx.AlignmentType.RIGHT }),
            para(run(dateLabel, { size: 16, color: BRAND.colors.ink }), { spacing: { after: 20 }, align: docx.AlignmentType.RIGHT }),
            para(run("Initiating coverage", { size: 16, bold: true, color: BRAND.colors.red }), { spacing: { after: 0 }, align: docx.AlignmentType.RIGHT })
          ], { width: 32, margins: { top: 0, bottom: 0, left: 0, right: 0 } })
        ]
      })
    ]);
  }

  function equityKeyDataSidebar({ ticker, rating, currentPrice, targetPrice, rangeReturn, volAnn, miniChartBytes }) {
    const rows = [];

    rows.push(para(run("Key data", { size: 18, bold: true, color: BRAND.colors.muted }), { spacing: { after: 120 } }));

    const kv = (k, v) => para([
      run(`${k}: `, { size: 18, bold: true }),
      run(v, { size: 18 })
    ], { spacing: { after: 90 } });

    rows.push(kv("Ticker", naIfBlank(ticker)));
    rows.push(kv("Recommendation", ratingToDisplay(rating)));
    rows.push(kv("Current price", currentPrice));
    rows.push(kv("Target price", targetPrice));

    rows.push(para(run(" ", { size: 10 }), { spacing: { after: 60 } }));
    rows.push(para(run("Price performance (range)", { size: 18, bold: true, color: BRAND.colors.muted }), { spacing: { after: 120 } }));
    rows.push(kv("Return", rangeReturn));
    rows.push(kv("Vol (ann.)", volAnn));

    rows.push(para(run(" ", { size: 10 }), { spacing: { after: 60 } }));
    rows.push(para(run("Chart", { size: 18, bold: true, color: BRAND.colors.muted }), { spacing: { after: 90 } }));

    if (miniChartBytes) {
      rows.push(
        new docx.Paragraph({
          children: [
            new docx.ImageRun({
              data: miniChartBytes,
              transformation: { width: 170, height: 120 }
            })
          ],
          spacing: { after: 0 }
        })
      );
    } else {
      rows.push(para(run("—", { size: 18, color: BRAND.colors.muted }), { spacing: { after: 0 } }));
    }

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      borders: {
        top: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border },
        bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border },
        left: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border },
        right: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border },
        insideHorizontal: { style: docx.BorderStyle.NONE },
        insideVertical: { style: docx.BorderStyle.NONE }
      },
      rows: [
        new docx.TableRow({
          children: [
            cell(rows, {
              shading: BRAND.colors.rail,
              margins: { top: 220, bottom: 220, left: 220, right: 220 }
            })
          ]
        })
      ]
    });
  }

  function equityMainHeader({ ticker, rating, currentPrice, targetPrice, title }) {
    const tkr = safeTrim(ticker) || "—";
    const rec = ratingToDisplay(rating);

    const priceLine = [
      run("Current Price: ", { size: 18, bold: true, color: BRAND.colors.muted }),
      run(`${currentPrice}`, { size: 18, color: BRAND.colors.muted }),
      run("    ", { size: 18 }),
      run("Target price: ", { size: 18, bold: true, color: BRAND.colors.muted }),
      run(`${targetPrice}`, { size: 18, color: BRAND.colors.muted })
    ];

    return [
      para([
        run(tkr.toUpperCase(), { font: BRAND.fonts.heading, size: 44, bold: true }),
        run("  ", { size: 10 }),
        run(rec, { font: BRAND.fonts.body, size: 22, bold: true, color: BRAND.colors.ink })
      ], { spacing: { after: 90 } }),
      para(priceLine, { spacing: { after: 160 } }),
      para(run(safeTrim(title) || "—", { font: BRAND.fonts.heading, size: 30, bold: true }), { spacing: { after: 120 } })
    ];
  }

  function tearSheetStatsRow({ currentPrice, targetPrice, upside, volAnn }) {
    const one = (label, value) => cell([
      para(run(label, { size: 16, bold: true, color: BRAND.colors.muted }), { spacing: { after: 40 } }),
      para(run(value, { size: 22, bold: true }), { spacing: { after: 0 } })
    ], {
      shading: "F3F4F6",
      borders: {
        top: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border },
        bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border },
        left: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border },
        right: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border }
      },
      margins: { top: 180, bottom: 180, left: 180, right: 180 },
      width: 25
    });

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        new docx.TableRow({
          children: [
            one("Current price", currentPrice),
            one("Target price", targetPrice),
            one("Upside", upside),
            one("Vol (ann.)", volAnn)
          ]
        })
      ]
    });
  }

  async function createInstitutionalDocument(payload) {
    const {
      noteType, title, topic,
      authorLastName, authorFirstName, authorPhoneSafe,
      coAuthors,
      analysis, keyTakeaways, content, cordobaView,
      imageFiles, dateTimeString,
      ticker, valuationSummary, keyAssumptions, scenarioNotes, modelFiles, modelLink,
      priceChartImageBytes,
      targetPrice, equityStats,
      crgRating
    } = payload;

    const nt = noteType || "Research Note";
    const now = new Date();
    const iso = formatDateShortISO(now);

    const cp = Number.isFinite(equityStats?.currentPrice) ? equityStats.currentPrice : null;
    const tp = safeNum(targetPrice);
    const upside = (cp !== null && tp !== null && cp > 0) ? ((tp / cp) - 1) : null;

    const displayCurrent = cp === null ? "—" : cp.toFixed(2);
    const displayTarget = tp === null ? (safeTrim(targetPrice) || "—") : tp.toFixed(2);
    const displayReturn = Number.isFinite(equityStats?.rangeReturn) ? pct(equityStats.rangeReturn) : "—";
    const displayVol = Number.isFinite(equityStats?.realisedVolAnn) ? pct(equityStats.realisedVolAnn) : "—";
    const displayUpside = upside === null ? "—" : pct(upside);

    const blocks = splitIntoParagraphBlocks(analysis);
    const p1Blocks = blocks.slice(0, 3);
    const restBlocks = blocks.slice(3);

    const p1Paras = p1Blocks.length
      ? p1Blocks.flatMap(b => linesToParagraphs(b, 140)).concat([bodyLine(" ", 60)])
      : [bodyLine("—")];

    const restParas = restBlocks.length
      ? restBlocks.flatMap(b => linesToParagraphs(b, 140))
      : [];

    const figures = await addImages(imageFiles);

    const attachedModelNames = (modelFiles && modelFiles.length)
      ? Array.from(modelFiles).map(f => f.name)
      : [];

    let children = [];

    if (nt === "Equity Research") {
      children.push(
        equityMasthead(dateTimeString.split(" ").slice(0, 3).join(" ")),
        thinRule(220)
      );

      const sidebar = equityKeyDataSidebar({
        ticker: safeTrim(ticker),
        rating: safeTrim(crgRating),
        currentPrice: displayCurrent,
        targetPrice: displayTarget,
        rangeReturn: displayReturn,
        volAnn: displayVol,
        miniChartBytes: priceChartImageBytes
      });

      const mainHeader = equityMainHeader({
        ticker: safeTrim(ticker),
        rating: safeTrim(crgRating),
        currentPrice: displayCurrent,
        targetPrice: displayTarget,
        title: safeTrim(title)
      });

      const page1Layout = new docx.Table({
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
        borders: {
          top: { style: docx.BorderStyle.NONE },
          bottom: { style: docx.BorderStyle.NONE },
          left: { style: docx.BorderStyle.NONE },
          right: { style: docx.BorderStyle.NONE },
          insideHorizontal: { style: docx.BorderStyle.NONE },
          insideVertical: { style: docx.BorderStyle.SINGLE, color: BRAND.colors.border, size: 2 }
        },
        rows: [
          new docx.TableRow({
            children: [
              cell([sidebar], { width: 24, margins: { top: 0, bottom: 0, left: 0, right: 220 } }),
              cell([...mainHeader, ...p1Paras], { width: 76, margins: { top: 0, bottom: 0, left: 220, right: 0 } })
            ]
          })
        ]
      });

      children.push(
        page1Layout,
        para(run(BRAND.disclaimers.internal, { size: 14, color: BRAND.colors.muted }), { spacing: { before: 220, after: 0 } }),
        pageBreak()
      );

      children.push(
        heading("Tear sheet", 34, 140),
        thinRule(160),
        tearSheetStatsRow({
          currentPrice: displayCurrent,
          targetPrice: displayTarget,
          upside: displayUpside,
          volAnn: displayVol
        }),
        para(run(" ", { size: 10 }), { spacing: { after: 120 } })
      );

      if (priceChartImageBytes) {
        children.push(
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: priceChartImageBytes,
                transformation: { width: 620, height: 210 }
              })
            ],
            spacing: { after: 160 },
            alignment: docx.AlignmentType.CENTER
          })
        );
      } else {
        children.push(bodyLine("Chart not attached (fetch chart before export).", 160));
      }

      if (safeTrim(valuationSummary)) {
        children.push(subheading("Valuation", 28, 110), ...linesToParagraphs(valuationSummary, 140));
      }
      if (safeTrim(keyAssumptions)) {
        children.push(subheading("Key assumptions", 28, 110), ...bulletLines(keyAssumptions, 90));
      }

      children.push(pageBreak());
      children.push(heading("Investment thesis", 34, 140), thinRule(180));

      if (restParas.length) children.push(...restParas);
      if (safeTrim(content)) children.push(subheading("Additional detail", 26, 110), ...linesToParagraphs(content, 140));
      if (safeTrim(scenarioNotes)) children.push(subheading("Scenario / sensitivity notes", 26, 110), ...linesToParagraphs(scenarioNotes, 140));
      if (safeTrim(cordobaView)) children.push(subheading("The Cordoba view", 26, 110), ...linesToParagraphs(cordobaView, 140));

      if (safeTrim(modelLink) || attachedModelNames.length) {
        children.push(subheading("Model and attachments", 26, 110));
        if (safeTrim(modelLink)) children.push(bodyLine(`Model link: ${safeTrim(modelLink)}`, 120));
        if (attachedModelNames.length) {
          attachedModelNames.forEach(name => {
            children.push(new docx.Paragraph({ text: name, bullet: { level: 0 }, spacing: { after: 70 } }));
          });
        } else {
          children.push(bodyLine("No files uploaded.", 120));
        }
      }

      if (figures.length) children.push(subheading("Figures and charts", 26, 110), ...figures);
      if (safeTrim(keyTakeaways)) children.push(subheading("Key takeaways", 26, 110), ...bulletLines(keyTakeaways, 90));
    } else {
      children.push(
        heading(BRAND.name.toUpperCase(), 30, 40),
        bodyLine(BRAND.tagline, 180),
        thinRule(180),
        heading(safeTrim(title) || "—", 34, 100),
        bodyLine(safeTrim(topic) || "—", 160),
        subheading("Key takeaways", 26, 110),
        ...bulletLines(keyTakeaways, 90),
        subheading("Analysis", 26, 110),
        ...linesToParagraphs(analysis, 140)
      );

      if (safeTrim(content)) children.push(subheading("Additional detail", 26, 110), ...linesToParagraphs(content, 140));
      if (safeTrim(cordobaView)) children.push(subheading("The Cordoba view", 26, 110), ...linesToParagraphs(cordobaView, 140));
      if (figures.length) children.push(subheading("Figures and charts", 26, 110), ...figures);
    }

    return new docx.Document({
      styles: {
        default: {
          document: {
            run: { font: BRAND.fonts.body, size: 20, color: BRAND.colors.ink },
            paragraph: { spacing: { after: 140 } }
          }
        }
      },
      sections: [{
        properties: {
          page: {
            margin: { top: 720, right: 720, bottom: 720, left: 720 },
            pageSize: { orientation: docx.PageOrientation.PORTRAIT }
          }
        },
        headers: {
          default: new docx.Header({
            children: [
              headerTable(nt, iso),
              thinRule(0)
            ]
          })
        },
        footers: {
          default: new docx.Footer({
            children: [
              footerLine(),
              footerTable()
            ]
          })
        },
        children
      }]
    });
  }

  // ------------------------------
  // Submit / Export
  // ------------------------------
  const form = document.getElementById("researchForm");
  if (form) form.noValidate = true;

  if (form) {
    form.addEventListener("submit", async (e) => {
      e.preventDefault();

      const button = form.querySelector('button[type="submit"]');
      if (!button) return;

      showMsg("", "");

      const missing = validatePublishCore();
      if (missing.length) {
        showMsg("error", `✗ Missing publish-core fields: ${missing.join(", ")}`);
        const first = document.getElementById(missing[0]);
        if (first && typeof first.focus === "function") first.focus();
        return;
      }

      button.disabled = true;
      button.classList.add("loading");
      button.textContent = "Generating…";

      try {
        ensureLibs();

        const noteType = safeTrim($("#noteType")?.value);
        const title = safeTrim($("#title")?.value);
        const topic = safeTrim($("#topic")?.value);

        const authorLastName = safeTrim($("#authorLastName")?.value);
        const authorFirstName = safeTrim($("#authorFirstName")?.value);
        const authorPhoneSafe = naIfBlank(safeTrim($("#authorPhone")?.value));

        const analysis = $("#analysis")?.value || "";
        const keyTakeaways = $("#keyTakeaways")?.value || "";
        const content = $("#content")?.value || "";
        const cordobaView = $("#cordobaView")?.value || "";

        const imageFiles = $("#imageUpload")?.files || [];

        const ticker = $("#ticker")?.value || "";
        const valuationSummary = $("#valuationSummary")?.value || "";
        const keyAssumptions = $("#keyAssumptions")?.value || "";
        const scenarioNotes = $("#scenarioNotes")?.value || "";
        const modelFiles = $("#modelFiles")?.files || null;
        const modelLink = $("#modelLink")?.value || "";

        const targetPrice = $("#targetPrice")?.value || "";
        const crgRating = $("#crgRating")?.value || "";

        const now = new Date();
        const dateTimeString = formatDateTime(now);

        const coAuthors = [];
        $$(".coauthor-entry").forEach(entry => {
          const lastName = safeTrim($(".coauthor-lastname", entry)?.value);
          const firstName = safeTrim($(".coauthor-firstname", entry)?.value);
          const phone = safeTrim($(".coauthor-phone", entry)?.value);
          if (lastName || firstName) {
            coAuthors.push({ lastName: lastName || "", firstName: firstName || "", phone: naIfBlank(phone) });
          }
        });

        const doc = await createInstitutionalDocument({
          noteType, title, topic,
          authorLastName, authorFirstName, authorPhoneSafe,
          coAuthors,
          analysis, keyTakeaways, content, cordobaView,
          imageFiles, dateTimeString,
          ticker, valuationSummary, keyAssumptions, scenarioNotes, modelFiles, modelLink,
          priceChartImageBytes,
          targetPrice,
          equityStats,
          crgRating
        });

        const blob = await docx.Packer.toBlob(doc);

        const fileName =
          `${(title || "research_note").replace(/[^a-z0-9]/gi, "_").toLowerCase()}_${(noteType || "note").replace(/\s+/g, "_").toLowerCase()}_${formatDateShortISO(now)}.docx`;

        saveAs(blob, fileName);

        showMsg("success", `✓ Document "${fileName}" generated successfully.`);
        saveDraftNow();
        syncWorkspacePreview();
      } catch (error) {
        console.error(error);
        showMsg("error", `✗ Error: ${error.message}`);
      } finally {
        button.disabled = false;
        button.classList.remove("loading");
        button.textContent = "Generate Word Document";
      }
    });
  }

  // ------------------------------
  // Draft restore
  // ------------------------------
  function applyDraft(draft) {
    if (!draft) return;
    DRAFT_FIELDS.forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      if (typeof draft[id] === "string") el.value = draft[id];
    });

    if (Array.isArray(draft.__coAuthors) && draft.__coAuthors.length) {
      if (coAuthorsList) {
        coAuthorsList.innerHTML = "";
        draft.__coAuthors.forEach(ca => {
          const node = createCoauthorNode();
          $(".coauthor-lastname", node).value = ca.lastName || "";
          $(".coauthor-firstname", node).value = ca.firstName || "";
          $(".coauthor-country", node).value = ca.cc || "44";
          $(".coauthor-phone-local", node).value = ca.local ? formatNationalLoose(ca.local) : "";
          wireCoauthorPhone(node);
          const hidden = $(".coauthor-phone", node);
          if (hidden) hidden.value = ca.phone || hidden.value || "";
          coAuthorsList.appendChild(node);
        });
      }
    }

    if (draft.__chartRange && $("#chartRange")) $("#chartRange").value = draft.__chartRange;

    if (draft.__equityStats) {
      equityStats = draft.__equityStats;
      paintEquityStats();
    }

    syncPrimaryPhone();
  }

  // ------------------------------
  // Workspace preview sync
  // ------------------------------
  function syncWorkspacePreview() {
    const titleEl = document.getElementById("title");
    const topicEl = document.getElementById("topic");
    const noteTypeElLocal = document.getElementById("noteType");
    const firstEl = document.getElementById("authorFirstName");
    const lastEl = document.getElementById("authorLastName");
    const takeawaysEl = document.getElementById("keyTakeaways");
    const analysisEl = document.getElementById("analysis");
    const cordobaViewEl = document.getElementById("cordobaView");

    const liveTitle = document.getElementById("liveTitle");
    const previewNoteType = document.getElementById("previewNoteType");
    const previewTopic = document.getElementById("previewTopic");
    const previewAuthor = document.getElementById("previewAuthor");
    const previewTakeaways = document.getElementById("previewTakeaways");
    const previewAnalysis = document.getElementById("previewAnalysis");
    const previewCordobaView = document.getElementById("previewCordobaView");
    const draftStatusMirror = document.getElementById("draftStatusMirror");
    const draftStatusLive = document.getElementById("draftStatus");

    const safe = (v, fallback) => {
      const s = (v || "").trim();
      return s || fallback;
    };

    if (liveTitle) liveTitle.textContent = safe(titleEl?.value, "Document preview");
    if (previewNoteType) previewNoteType.textContent = safe(noteTypeElLocal?.value, "—");
    if (previewTopic) previewTopic.textContent = safe(topicEl?.value, "—");

    if (previewAuthor) {
      const fullName = [firstEl?.value, lastEl?.value].map(v => (v || "").trim()).filter(Boolean).join(" ");
      previewAuthor.textContent = fullName || "—";
    }

    if (previewTakeaways) {
      previewTakeaways.textContent = safe(takeawaysEl?.value, "Your key takeaways will appear here as you draft.");
    }

    if (previewAnalysis) {
      previewAnalysis.textContent = safe(analysisEl?.value, "Your analysis preview will appear here.");
    }

    if (previewCordobaView) {
      previewCordobaView.textContent = safe(cordobaViewEl?.value, "Your Cordoba view will appear here.");
    }

    if (draftStatusMirror && draftStatusLive) {
      draftStatusMirror.textContent = draftStatusLive.textContent || "—";
    }
  }

  ["input","change","keyup"].forEach(evt => {
    document.addEventListener(evt, syncWorkspacePreview, { passive: true });
  });

  let previewSyncInterval = null;

  // ------------------------------
  // Init
  // ------------------------------
  function init() {
    syncPrimaryPhone();
    updateAttachmentSummary();
    updateCompletionMeter();

    const draft = loadDraft();
    if (draft) {
      applyDraft(draft);
      setDraftStatus("Restored");
      updateAttachmentSummary();
      updateCompletionMeter();
    } else {
      setDraftStatus("");
    }

    syncWorkspacePreview();

    if (!previewSyncInterval) {
      previewSyncInterval = setInterval(syncWorkspacePreview, 500);
    }
  }

  window.addEventListener("DOMContentLoaded", init);
  window.addEventListener("beforeunload", () => {
    if (previewSyncInterval) clearInterval(previewSyncInterval);
  });
})();
