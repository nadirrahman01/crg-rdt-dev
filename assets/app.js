/* =========================================================
   Cordoba Research Group — Research Documentation Tool (RDT)
   assets/app.js (BlueMatrix-grade v2.1 — Equity-first Export)
   ---------------------------------------------------------
   What this version guarantees:
   1) Backwards compatible with your current HTML IDs/classes.
   2) Autosave + restore (drafts), robust validation, completion meter.
   3) Equity Research Word export formatted like the classic sell-side
      “Initiating Coverage” cover shown in your screenshot:
        - Institutional header + date + “Initiating Coverage”
        - Left sidebar “Key data” block + mini chart
        - Main area: Ticker, Recommendation, Current Price, Target Price
        - Big headline + Investment thesis paragraphs
        - Page 2+: Valuation, Drivers, Risks, Catalysts, Appendix/Figures
   4) Chart fetch embeds chart image + computed stats into the Word tear sheet.
   ========================================================= */

(() => {
  "use strict";

  console.log("RDT app.js loaded (Equity-first v2.1)");

  // ------------------------------
  // Brand (Córdoba)
  // ------------------------------
  const BRAND = {
    name: "Cordoba Research Group",
    short: "CRG",
    version: "RDT v2.1.0",
    colors: {
      gold: "9A690F",
      goldDark: "845F0F",
      cream: "FFF7F0",
      ink: "0B0E14",
      muted: "6B7280",
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
      publicInfo: "Cordoba Research Group Public Information"
    }
  };

  // ------------------------------
  // DOM utilities
  // ------------------------------
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const clamp = (n, a, b) => Math.min(Math.max(n, a), b);

  function safeTrim(v) {
    return (v ?? "").toString().trim();
  }

  function digitsOnly(v) {
    return (v || "").toString().replace(/\D/g, "");
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

  function escapeHtml(s) {
    return (s || "").toString()
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
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

  // Mailto: CRLF body
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

  // ------------------------------
  // Draft persistence (autosave)
  // ------------------------------
  const DRAFT_KEY = "crg_rdt_draft_v21";

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
    } catch (_) { /* ignore */ }
  }

  let draftSaveTimer = null;
  function scheduleDraftSave() {
    setDraftStatus("Saving…");
    clearTimeout(draftSaveTimer);
    draftSaveTimer = setTimeout(() => saveDraftNow(), 320);
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

  function applyDraft(draft) {
    if (!draft) return;

    DRAFT_FIELDS.forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      if (typeof draft[id] === "string") el.value = draft[id];
    });

    // Coauthors
    if (Array.isArray(draft.__coAuthors) && draft.__coAuthors.length) {
      const list = document.getElementById("coAuthorsList");
      if (list) {
        list.innerHTML = "";
        draft.__coAuthors.forEach(ca => {
          const node = createCoauthorNode();
          $(".coauthor-lastname", node).value = ca.lastName || "";
          $(".coauthor-firstname", node).value = ca.firstName || "";
          $(".coauthor-country", node).value = ca.cc || "44";
          $(".coauthor-phone-local", node).value = ca.local ? formatNationalLoose(ca.local) : "";
          wireCoauthorPhone(node);
          const hidden = $(".coauthor-phone", node);
          if (hidden) hidden.value = ca.phone || hidden.value || "";
          list.appendChild(node);
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
  // Phone wiring (primary)
  // ------------------------------
  const authorPhoneCountryEl = document.getElementById("authorPhoneCountry");
  const authorPhoneNationalEl = document.getElementById("authorPhoneNational");
  const authorPhoneHiddenEl = document.getElementById("authorPhone");

  function syncPrimaryPhone() {
    if (!authorPhoneHiddenEl) return;
    const cc = authorPhoneCountryEl ? authorPhoneCountryEl.value : "";
    const nationalDigits = digitsOnly(authorPhoneNationalEl ? authorPhoneNationalEl.value : "");
    authorPhoneHiddenEl.value = buildInternationalHyphen(cc, nationalDigits);
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

    const coAuthorDiv = document.createElement("div");
    coAuthorDiv.className = "coauthor-entry";
    coAuthorDiv.id = `coauthor-${coAuthorCount}`;

    coAuthorDiv.innerHTML = `
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

    const phoneHidden = $(".coauthor-phone", coAuthorDiv);
    if (phoneHidden) phoneHidden.required = false;

    wireCoauthorPhone(coAuthorDiv);

    ["input","change","keyup"].forEach(evt => {
      coAuthorDiv.addEventListener(evt, () => scheduleDraftSave(), { passive: true });
    });

    return coAuthorDiv;
  }

  if (addCoAuthorBtn && coAuthorsList) {
    addCoAuthorBtn.addEventListener("click", () => {
      coAuthorsList.appendChild(createCoauthorNode());
      updateCompletionMeter();
      scheduleDraftSave();
    });

    document.addEventListener("click", (e) => {
      const btn = e.target.closest(".remove-coauthor");
      if (!btn) return;
      const id = btn.getAttribute("data-remove-id");
      const coAuthorDiv = document.getElementById(`coauthor-${id}`);
      if (coAuthorDiv) coAuthorDiv.remove();
      updateCompletionMeter();
      scheduleDraftSave();
    });
  }

  // ------------------------------
  // Equity section toggle
  // ------------------------------
  const noteTypeEl = document.getElementById("noteType");
  const equitySectionEl = document.getElementById("equitySection");

  function isEquityMode() {
    return !!(noteTypeEl && noteTypeEl.value === "Equity Research" && equitySectionEl && equitySectionEl.style.display !== "none");
  }

  function toggleEquitySection() {
    if (!noteTypeEl || !equitySectionEl) return;
    equitySectionEl.style.display = (noteTypeEl.value === "Equity Research") ? "block" : "none";
  }

  if (noteTypeEl && equitySectionEl) {
    noteTypeEl.addEventListener("change", () => {
      toggleEquitySection();
      updateCompletionMeter();
      scheduleDraftSave();
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

  // Equity requires a rating (like a bank note: must have a call)
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
      }
    }, { passive: true });
  });

  // ------------------------------
  // Attachment summary
  // ------------------------------
  const modelFilesEl = document.getElementById("modelFiles");
  const attachSummaryHeadEl = document.getElementById("attachmentSummaryHead");
  const attachSummaryListEl = document.getElementById("attachmentSummaryList");

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
  // Price chart + stats (Stooq via r.jina.ai)
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
    // Prefer HTTPS to avoid mixed content.
    const stooqUrl = `https://stooq.com/q/d/l/?s=${encodeURIComponent(symbol)}&i=d`;
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
          tension: 0.18
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { intersect: false, mode: "index" }
        },
        scales: {
          x: { ticks: { maxTicksLimit: 6 } },
          y: { ticks: { maxTicksLimit: 6 } }
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

      renderChart({ labels, values, title: `${tickerVal.toUpperCase()} Close` });

      // allow chart to render before capture
      await new Promise(r => setTimeout(r, 180));
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

  // ============================================================
  // Word export (sell-side layout)
  // ============================================================

  function ensureLibs() {
    if (typeof docx === "undefined") throw new Error("docx library not loaded. Refresh the page.");
    if (typeof saveAs === "undefined") throw new Error("FileSaver library not loaded. Refresh the page.");
  }

  // --- docx helpers (consistent typography) ---
  function P(text, opts = {}) {
    return new docx.Paragraph({
      children: [
        new docx.TextRun({
          text: text ?? "",
          font: opts.font ?? BRAND.fonts.body,
          size: opts.size ?? 20, // half-points (20 => 10pt)
          bold: !!opts.bold,
          italics: !!opts.italics,
          color: opts.color ?? BRAND.colors.ink
        })
      ],
      spacing: opts.spacing ?? { after: 140 },
      alignment: opts.align
    });
  }

  function HR(spacingAfter = 160) {
    return new docx.Paragraph({
      border: { bottom: { color: BRAND.colors.border, space: 1, style: docx.BorderStyle.SINGLE, size: 4 } },
      spacing: { after: spacingAfter }
    });
  }

  function linesToParas(text, style = {}) {
    const lines = (text || "").split("\n");
    return lines.map(line => {
      if (!line.trim()) return new docx.Paragraph({ text: "", spacing: { after: style.after ?? 140 } });
      return new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: line,
            font: style.font ?? BRAND.fonts.body,
            size: style.size ?? 22,
            color: style.color ?? BRAND.colors.ink,
            bold: !!style.bold
          })
        ],
        spacing: { after: style.after ?? 140 }
      });
    });
  }

  function bulletLines(text, spacingAfter = 90, size = 20) {
    const lines = (text || "").split("\n");
    const bullets = [];
    lines.forEach(line => {
      const t = line.replace(/^[-*•]\s*/, "").trim();
      if (!t) return;
      bullets.push(new docx.Paragraph({
        children: [new docx.TextRun({ text: t, font: BRAND.fonts.body, size, color: BRAND.colors.ink })],
        bullet: { level: 0 },
        spacing: { after: spacingAfter }
      }));
    });
    return bullets.length ? bullets : [new docx.Paragraph({ text: "—", spacing: { after: spacingAfter } })];
  }

  function shadedBox(children, fillHex, pad = 180, borderHex = BRAND.colors.border) {
    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      borders: {
        top: { style: docx.BorderStyle.SINGLE, size: 2, color: borderHex },
        bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: borderHex },
        left: { style: docx.BorderStyle.SINGLE, size: 2, color: borderHex },
        right: { style: docx.BorderStyle.SINGLE, size: 2, color: borderHex },
        insideHorizontal: { style: docx.BorderStyle.NONE },
        insideVertical: { style: docx.BorderStyle.NONE }
      },
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              shading: { fill: fillHex },
              margins: { top: pad, bottom: pad, left: pad, right: pad },
              children
            })
          ]
        })
      ]
    });
  }

  function metaHeaderLine(left, right) {
    return new docx.Paragraph({
      spacing: { after: 80 },
      tabStops: [{ type: docx.TabStopType.RIGHT, position: 9000 }],
      children: [
        new docx.TextRun({ text: left, size: 16, font: BRAND.fonts.body, color: BRAND.colors.muted }),
        new docx.TextRun({ text: "\t" + (right || ""), size: 16, font: BRAND.fonts.body, color: BRAND.colors.muted })
      ]
    });
  }

  function pageBreak() {
    return new docx.Paragraph({ children: [new docx.PageBreak()] });
  }

  function recommendationFromCRGRating(crgRating) {
    const r = safeTrim(crgRating);
    // Match your screenshot language more closely
    if (r === "Buy") return "Accumulate";
    if (r === "Hold") return "Neutral";
    if (r === "Sell") return "Reduce";
    return r || "—";
  }

  function firstNonEmptyLine(text) {
    const lines = (text || "").split("\n").map(x => x.trim()).filter(Boolean);
    return lines[0] || "";
  }

  function firstParagraphBlock(text, maxChars = 650) {
    const t = safeTrim(text);
    if (!t) return "—";
    const compact = t.replace(/\n+/g, " ").replace(/\s+/g, " ").trim();
    return compact.length > maxChars ? compact.slice(0, maxChars - 1) + "…" : compact;
  }

  async function addImagesToAppendix(imageFiles) {
    const list = Array.from(imageFiles || []);
    if (!list.length) return [];

    const out = [];
    for (let i = 0; i < list.length; i++) {
      const file = list[i];
      try {
        const arrayBuffer = await file.arrayBuffer();
        const fileNameWithoutExt = file.name.replace(/\.[^/.]+$/, "");
        out.push(
          P(`Figure ${i + 1}: ${fileNameWithoutExt}`, { italics: true, size: 18, color: BRAND.colors.muted, spacing: { after: 90 } , font: BRAND.fonts.body }),
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: arrayBuffer,
                transformation: { width: 620, height: 360 }
              })
            ],
            spacing: { after: 220 },
            alignment: docx.AlignmentType.CENTER
          })
        );
      } catch (e) {
        out.push(P(`Figure ${i + 1}: (could not embed ${file.name})`, { size: 18, color: BRAND.colors.muted }));
      }
    }
    return out;
  }

  // ----------------------------------------------------------
  // Equity report cover page — “Initiating Coverage” format
  // ----------------------------------------------------------
  function equityCoverPage(payload) {
    const {
      title,
      ticker,
      crgRating,
      targetPrice,
      equityStats
    } = payload;

    const now = new Date();
    const dateStr = `${now.getDate()} ${["January","February","March","April","May","June","July","August","September","October","November","December"][now.getMonth()]} ${now.getFullYear()}`;

    const rec = recommendationFromCRGRating(crgRating);
    const cp = Number.isFinite(equityStats?.currentPrice) ? equityStats.currentPrice : null;
    const tp = safeNum(targetPrice);

    const headline = safeTrim(title) || firstNonEmptyLine(payload.analysis) || "—";

    // Top “institutional” header (like the screenshot)
    const topRow = new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      borders: {
        top: { style: docx.BorderStyle.NONE },
        bottom: { style: docx.BorderStyle.NONE },
        left: { style: docx.BorderStyle.NONE },
        right: { style: docx.BorderStyle.NONE },
        insideHorizontal: { style: docx.BorderStyle.NONE },
        insideVertical: { style: docx.BorderStyle.NONE }
      },
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              width: { size: 55, type: docx.WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 60, left: 0, right: 0 },
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: BRAND.name.toUpperCase(), font: BRAND.fonts.heading, size: 26, bold: true, color: BRAND.colors.ink })
                  ],
                  spacing: { after: 20 }
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: "Values that bind", font: BRAND.fonts.body, size: 16, color: BRAND.colors.muted })
                  ],
                  spacing: { after: 0 }
                })
              ]
            }),
            new docx.TableCell({
              width: { size: 45, type: docx.WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 60, left: 0, right: 0 },
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: "INSTITUTIONAL EQUITY RESEARCH", font: BRAND.fonts.body, size: 16, bold: true, color: BRAND.colors.red })
                  ],
                  alignment: docx.AlignmentType.RIGHT,
                  spacing: { after: 20 }
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: dateStr, font: BRAND.fonts.body, size: 16, color: BRAND.colors.muted })
                  ],
                  alignment: docx.AlignmentType.RIGHT,
                  spacing: { after: 20 }
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: "Initiating coverage", font: BRAND.fonts.body, size: 16, bold: true, color: BRAND.colors.red })
                  ],
                  alignment: docx.AlignmentType.RIGHT,
                  spacing: { after: 0 }
                })
              ]
            })
          ]
        })
      ]
    });

    // Sidebar “Key data”
    const keyDataLines = [];

    keyDataLines.push(P("Key data", { bold: true, size: 22, font: BRAND.fonts.body, spacing: { after: 110 } }));
    keyDataLines.push(P(`Ticker: ${safeTrim(ticker) || "—"}`, { size: 18, color: BRAND.colors.ink, spacing: { after: 70 } }));
    keyDataLines.push(P(`Recommendation: ${rec || "—"}`, { size: 18, color: BRAND.colors.ink, spacing: { after: 70 } }));
    keyDataLines.push(P(`Current price: ${cp === null ? "—" : cp.toFixed(2)}`, { size: 18, color: BRAND.colors.ink, spacing: { after: 70 } }));
    keyDataLines.push(P(`Target price: ${tp === null ? "—" : tp.toFixed(2)}`, { size: 18, color: BRAND.colors.ink, spacing: { after: 70 } }));

    // Light “stats” block under key data (mirrors the “data rail” feel)
    const rr = Number.isFinite(equityStats?.rangeReturn) ? equityStats.rangeReturn : null;
    const vol = Number.isFinite(equityStats?.realisedVolAnn) ? equityStats.realisedVolAnn : null;

    keyDataLines.push(new docx.Paragraph({ spacing: { after: 90 } }));
    keyDataLines.push(P("Price performance (range)", { bold: true, size: 18, color: BRAND.colors.muted, spacing: { after: 70 } }));
    keyDataLines.push(P(`Return: ${rr === null ? "—" : pct(rr)}`, { size: 18, spacing: { after: 60 } }));
    keyDataLines.push(P(`Vol (ann.): ${vol === null ? "—" : pct(vol)}`, { size: 18, spacing: { after: 100 } }));

    // Embed a mini chart if available
    if (priceChartImageBytes) {
      keyDataLines.push(P("Chart", { bold: true, size: 18, color: BRAND.colors.muted, spacing: { after: 60 } }));
      keyDataLines.push(
        new docx.Paragraph({
          children: [
            new docx.ImageRun({ data: priceChartImageBytes, transformation: { width: 250, height: 120 } })
          ],
          spacing: { after: 0 },
          alignment: docx.AlignmentType.CENTER
        })
      );
    }

    const sidebar = shadedBox(keyDataLines, BRAND.colors.rail, 160);

    // Main cover body (ticker header line + headline + thesis block)
    const mainTop = [];

    mainTop.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: (safeTrim(ticker) || "—").toUpperCase(), font: BRAND.fonts.heading, size: 40, bold: true, color: BRAND.colors.ink }),
          new docx.TextRun({ text: "   ", size: 20 }),
          new docx.TextRun({ text: rec || "—", font: BRAND.fonts.body, size: 22, bold: true, color: BRAND.colors.ink })
        ],
        spacing: { after: 90 }
      })
    );

    mainTop.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: "Current Price: ", font: BRAND.fonts.body, size: 18, bold: true, color: BRAND.colors.muted }),
          new docx.TextRun({ text: cp === null ? "—" : `Rs${cp.toFixed(0)}`, font: BRAND.fonts.body, size: 18, color: BRAND.colors.ink }),
          new docx.TextRun({ text: "    ", size: 18 }),
          new docx.TextRun({ text: "Target price: ", font: BRAND.fonts.body, size: 18, bold: true, color: BRAND.colors.muted }),
          new docx.TextRun({ text: tp === null ? "—" : `Rs${tp.toFixed(0)}`, font: BRAND.fonts.body, size: 18, color: BRAND.colors.ink })
        ],
        spacing: { after: 120 }
      })
    );

    mainTop.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: headline, font: BRAND.fonts.heading, size: 32, bold: true, color: BRAND.colors.ink })
        ],
        spacing: { after: 140 }
      })
    );

    // Thesis text = first chunk of analysis (cover page should read like a short initiation thesis)
    const thesis = firstParagraphBlock(payload.analysis, 900);
    mainTop.push(...linesToParas(thesis, { font: BRAND.fonts.body, size: 22, after: 140 }));

    // Cover layout: left sidebar + main body (like screenshot)
    const coverGrid = new docx.Table({
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
            new docx.TableCell({
              width: { size: 28, type: docx.WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 0, right: 240 },
              children: [sidebar],
              verticalAlign: docx.VerticalAlign.TOP
            }),
            new docx.TableCell({
              width: { size: 72, type: docx.WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 0, right: 0 },
              children: mainTop,
              verticalAlign: docx.VerticalAlign.TOP
            })
          ]
        })
      ]
    });

    return [topRow, HR(120), coverGrid];
  }

  // ----------------------------------------------------------
  // Equity pages (post-cover)
  // ----------------------------------------------------------
  function equityBodyPages(payload) {
    const now = new Date();
    const cp = Number.isFinite(payload.equityStats?.currentPrice) ? payload.equityStats.currentPrice : null;
    const tp = safeNum(payload.targetPrice);
    const upside = (cp !== null && tp !== null && cp > 0) ? ((tp / cp) - 1) : null;

    const out = [];

    // Page 2: Investment thesis + in-brief rail (BlueMatrix-ish rhythm)
    out.push(pageBreak());
    out.push(metaHeaderLine(`${BRAND.short} | Equity Research | ${formatDateShortISO(now)}`, "Public Information"));
    out.push(HR(120));

    const inBrief = shadedBox([
      P("In brief", { bold: true, size: 20, spacing: { after: 90 } }),
      ...bulletLines(payload.keyTakeaways || "", 80, 20)
    ], BRAND.colors.callout, 160);

    const left = [
      P("Investment thesis", { bold: true, font: BRAND.fonts.heading, size: 30, spacing: { after: 140 } }),
      ...linesToParas(payload.analysis || "—", { font: BRAND.fonts.body, size: 22, after: 140 })
    ];

    const thesisGrid = new docx.Table({
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
            new docx.TableCell({
              width: { size: 68, type: docx.WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 0, right: 240 },
              children: left,
              verticalAlign: docx.VerticalAlign.TOP
            }),
            new docx.TableCell({
              width: { size: 32, type: docx.WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 0, right: 0 },
              children: [inBrief],
              verticalAlign: docx.VerticalAlign.TOP
            })
          ]
        })
      ]
    });

    out.push(thesisGrid);

    // Page 3: Tear sheet (clean)
    out.push(pageBreak());
    out.push(metaHeaderLine(`${BRAND.short} | Equity Research | Tear sheet`, "Public Information"));
    out.push(HR(120));
    out.push(P("Tear sheet", { bold: true, font: BRAND.fonts.heading, size: 30, spacing: { after: 140 } }));

    // Small stats row
    const statsCells = [
      { k: "Current price", v: cp === null ? "—" : cp.toFixed(2) },
      { k: "Target price", v: tp === null ? "—" : tp.toFixed(2) },
      { k: "Upside", v: upside === null ? "—" : pct(upside) },
      { k: "Vol (ann.)", v: Number.isFinite(payload.equityStats?.realisedVolAnn) ? pct(payload.equityStats.realisedVolAnn) : "—" }
    ];

    const statsRow = new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        new docx.TableRow({
          children: statsCells.map(c => new docx.TableCell({
            width: { size: 25, type: docx.WidthType.PERCENTAGE },
            margins: { top: 160, bottom: 160, left: 160, right: 160 },
            shading: { fill: BRAND.colors.rail },
            borders: {
              top: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border },
              bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border },
              left: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border },
              right: { style: docx.BorderStyle.SINGLE, size: 2, color: BRAND.colors.border }
            },
            children: [
              P(c.k, { size: 16, color: BRAND.colors.muted, bold: true, spacing: { after: 40 } }),
              P(c.v, { size: 24, bold: true, spacing: { after: 0 } })
            ]
          }))
        })
      ]
    });

    out.push(statsRow);

    out.push(new docx.Paragraph({ spacing: { after: 120 } }));

    if (priceChartImageBytes) {
      out.push(
        new docx.Paragraph({
          children: [new docx.ImageRun({ data: priceChartImageBytes, transformation: { width: 620, height: 260 } })],
          alignment: docx.AlignmentType.CENTER,
          spacing: { after: 140 }
        })
      );
    } else {
      out.push(P("Chart not attached (fetch chart before export).", { size: 18, color: BRAND.colors.muted }));
    }

    // Valuation inputs
    if (safeTrim(payload.valuationSummary)) {
      out.push(P("Valuation", { bold: true, font: BRAND.fonts.heading, size: 28, spacing: { before: 160, after: 120 } }));
      out.push(...linesToParas(payload.valuationSummary, { font: BRAND.fonts.body, size: 22, after: 140 }));
    }

    if (safeTrim(payload.keyAssumptions)) {
      out.push(P("Key assumptions", { bold: true, font: BRAND.fonts.heading, size: 26, spacing: { before: 140, after: 90 } }));
      out.push(...bulletLines(payload.keyAssumptions, 80, 20));
    }

    if (safeTrim(payload.scenarioNotes)) {
      out.push(P("Scenario / sensitivities", { bold: true, font: BRAND.fonts.heading, size: 26, spacing: { before: 140, after: 90 } }));
      out.push(...linesToParas(payload.scenarioNotes, { font: BRAND.fonts.body, size: 22, after: 140 }));
    }

    // Page 4+: Additional detail + Córdoba view
    const extra = safeTrim(payload.content);
    const cv = safeTrim(payload.cordobaView);

    if (extra || cv) {
      out.push(pageBreak());
      out.push(metaHeaderLine(`${BRAND.short} | Equity Research | Detail`, "Public Information"));
      out.push(HR(120));

      if (extra) {
        out.push(P("Company / industry detail", { bold: true, font: BRAND.fonts.heading, size: 28, spacing: { after: 120 } }));
        out.push(...linesToParas(extra, { font: BRAND.fonts.body, size: 22, after: 140 }));
      }

      if (cv) {
        out.push(P("The Cordoba View", { bold: true, font: BRAND.fonts.heading, size: 28, spacing: { before: 160, after: 120 } }));
        out.push(...linesToParas(cv, { font: BRAND.fonts.body, size: 22, after: 140 }));
      }
    }

    return out;
  }

  // ----------------------------------------------------------
  // Non-equity: clean institutional template (kept strong)
  // ----------------------------------------------------------
  function nonEquityDocument(payload) {
    const now = new Date();

    const top = shadedBox([
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: BRAND.short, font: BRAND.fonts.body, size: 20, bold: true, color: "FFFFFF" }),
          new docx.TextRun({ text: "  ", size: 20 }),
          new docx.TextRun({ text: (payload.noteType || "Research Note").toUpperCase(), font: BRAND.fonts.body, size: 18, bold: true, color: "FFFFFF" })
        ],
        spacing: { after: 40 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: safeTrim(payload.title) || "—", font: BRAND.fonts.heading, size: 34, bold: true, color: "FFFFFF" })
        ],
        spacing: { after: 30 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: safeTrim(payload.topic) || "—", font: BRAND.fonts.body, size: 18, color: "FFFFFF" })
        ],
        spacing: { after: 0 }
      })
    ], BRAND.colors.goldDark, 220, BRAND.colors.goldDark);

    const authorLine = `${safeTrim(payload.authorFirstName)} ${safeTrim(payload.authorLastName)}`.trim() || "—";

    const meta = shadedBox([
      P("Author", { bold: true, size: 18, color: BRAND.colors.muted, spacing: { after: 60 } }),
      P(authorLine, { bold: true, size: 22, spacing: { after: 60 } }),
      P(`Phone: ${naIfBlank(payload.authorPhone)}`, { size: 18, color: BRAND.colors.muted, spacing: { after: 70 } }),
      P(`Generated: ${formatDateTime(now)}`, { size: 16, color: BRAND.colors.muted, spacing: { after: 0 } })
    ], BRAND.colors.rail, 160);

    const page1 = new docx.Table({
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
            new docx.TableCell({
              width: { size: 70, type: docx.WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 0, right: 240 },
              children: [
                P("Key takeaways", { bold: true, font: BRAND.fonts.heading, size: 28, spacing: { after: 90 } }),
                ...bulletLines(payload.keyTakeaways || "", 90, 20),
                new docx.Paragraph({ spacing: { after: 140 } }),
                P("Analysis", { bold: true, font: BRAND.fonts.heading, size: 28, spacing: { after: 120 } }),
                ...linesToParas(payload.analysis || "—", { font: BRAND.fonts.body, size: 22, after: 140 })
              ],
              verticalAlign: docx.VerticalAlign.TOP
            }),
            new docx.TableCell({
              width: { size: 30, type: docx.WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 0, right: 0 },
              children: [meta],
              verticalAlign: docx.VerticalAlign.TOP
            })
          ]
        })
      ]
    });

    const rest = [];
    const extra = safeTrim(payload.content);
    const cv = safeTrim(payload.cordobaView);

    if (extra || cv) {
      rest.push(pageBreak());
      rest.push(metaHeaderLine(`${BRAND.short} | ${(payload.noteType || "Note")} | ${formatDateShortISO(now)}`, "Public Information"));
      rest.push(HR(120));
      if (extra) {
        rest.push(P("Additional detail", { bold: true, font: BRAND.fonts.heading, size: 28, spacing: { after: 120 } }));
        rest.push(...linesToParas(extra, { font: BRAND.fonts.body, size: 22, after: 140 }));
      }
      if (cv) {
        rest.push(P("The Cordoba View", { bold: true, font: BRAND.fonts.heading, size: 28, spacing: { before: 160, after: 120 } }));
        rest.push(...linesToParas(cv, { font: BRAND.fonts.body, size: 22, after: 140 }));
      }
    }

    return [top, new docx.Paragraph({ spacing: { after: 120 } }), page1, ...rest];
  }

  // ----------------------------------------------------------
  // Main builder — routes equity vs non-equity
  // ----------------------------------------------------------
  async function createInstitutionalDocument(payload) {
    const now = new Date();

    const sectionsChildren = [];

    if (payload.noteType === "Equity Research") {
      // Cover page (initiating coverage format)
      sectionsChildren.push(...equityCoverPage(payload));

      // Disclaimers under cover (like sell-side fine print)
      sectionsChildren.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: BRAND.disclaimers.internal,
              font: BRAND.fonts.body,
              size: 14,
              color: BRAND.colors.muted
            })
          ],
          spacing: { before: 160, after: 0 }
        })
      );

      // Body pages
      sectionsChildren.push(...equityBodyPages(payload));
    } else {
      sectionsChildren.push(...nonEquityDocument(payload));
    }

    // Figures appendix (always last if present)
    const figs = await addImagesToAppendix(payload.imageFiles);
    if (figs.length) {
      sectionsChildren.push(pageBreak());
      sectionsChildren.push(metaHeaderLine(`${BRAND.short} | Appendix | Figures`, "Public Information"));
      sectionsChildren.push(HR(120));
      sectionsChildren.push(P("Figures and charts", { bold: true, font: BRAND.fonts.heading, size: 28, spacing: { after: 120 } }));
      sectionsChildren.push(...figs);
    }

    // Construct document with consistent headers/footers
    const doc = new docx.Document({
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
              metaHeaderLine(`${BRAND.short} | ${payload.noteType || "Research"} | ${formatDateShortISO(now)}`, BRAND.short)
            ]
          })
        },
        footers: {
          default: new docx.Footer({
            children: [
              new docx.Paragraph({
                border: { top: { color: BRAND.colors.border, space: 1, style: docx.BorderStyle.SINGLE, size: 2 } },
                spacing: { after: 0 }
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({ text: BRAND.disclaimers.publicInfo, size: 14, italics: true, color: BRAND.colors.muted }),
                  new docx.TextRun({ text: "\t" }),
                  new docx.TextRun({
                    children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
                    size: 14,
                    italics: true,
                    color: BRAND.colors.muted
                  })
                ],
                tabStops: [{ type: docx.TabStopType.RIGHT, position: 9000 }],
                spacing: { before: 70, after: 0 }
              })
            ]
          })
        },
        children: sectionsChildren
      }]
    });

    return doc;
  }

  // ------------------------------
  // Validation
  // ------------------------------
  function validatePublishCore() {
    const missing = [];
    const core = isEquityMode() ? baseCoreIds.concat(equityCoreIds) : baseCoreIds;
    core.forEach(id => {
      const el = document.getElementById(id);
      if (el && !isFilled(el)) missing.push(id);
    });
    return missing;
  }

  // ------------------------------
  // Submit (Generate Word)
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
        const authorPhone = safeTrim($("#authorPhone")?.value);

        const analysis = $("#analysis")?.value || "";
        const keyTakeaways = $("#keyTakeaways")?.value || "";
        const content = $("#content")?.value || "";
        const cordobaView = $("#cordobaView")?.value || "";

        const imageFiles = $("#imageUpload")?.files || [];

        const ticker = $("#ticker") ? $("#ticker").value : "";
        const valuationSummary = $("#valuationSummary") ? $("#valuationSummary").value : "";
        const keyAssumptions = $("#keyAssumptions") ? $("#keyAssumptions").value : "";
        const scenarioNotes = $("#scenarioNotes") ? $("#scenarioNotes").value : "";
        const modelFiles = $("#modelFiles") ? $("#modelFiles").files : null;
        const modelLink = $("#modelLink") ? $("#modelLink").value : "";

        const targetPrice = $("#targetPrice") ? $("#targetPrice").value : "";
        const crgRating = $("#crgRating") ? $("#crgRating").value : "";

        const coAuthors = [];
        $$(".coauthor-entry").forEach(entry => {
          const lastName = safeTrim($(".coauthor-lastname", entry)?.value);
          const firstName = safeTrim($(".coauthor-firstname", entry)?.value);
          const phone = safeTrim($(".coauthor-phone", entry)?.value);
          if (lastName || firstName) {
            coAuthors.push({ lastName, firstName, phone: naIfBlank(phone) });
          }
        });

        const doc = await createInstitutionalDocument({
          noteType, title, topic,
          authorLastName, authorFirstName, authorPhone,
          coAuthors,
          analysis, keyTakeaways, content, cordobaView,
          imageFiles,
          ticker, valuationSummary, keyAssumptions, scenarioNotes, modelFiles, modelLink,
          priceChartImageBytes,
          targetPrice,
          equityStats,
          crgRating
        });

        const now = new Date();
        const fileName =
          `${(title || "crg_note").replace(/[^a-z0-9]/gi, "_").toLowerCase()}_${(noteType || "note").replace(/\s+/g, "_").toLowerCase()}_${formatDateShortISO(now)}.docx`;

        const blob = await docx.Packer.toBlob(doc);
        saveAs(blob, fileName);

        showMsg("success", `✓ Document "${fileName}" generated successfully.`);
        saveDraftNow();
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
  // Init / restore
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
  }

  window.addEventListener("DOMContentLoaded", init);
})();
