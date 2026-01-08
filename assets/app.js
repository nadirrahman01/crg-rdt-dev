/* =========================================================
   Cordoba Research Group — Research Documentation Tool (RDT)
   app.js (Institutional v2)
   ---------------------------------------------------------
   Goals:
   - “BlueMatrix-grade” authoring UX: autosave drafts, robust validation,
     institutional actions (export/email/reset), attachment summaries,
     equity tear-sheet workflow (chart + stats), section-aware contents.
   - Word export: Peel Hunt-style *layout language* (banner + sidebar/rail,
     strong typographic hierarchy, callout rail), but Córdoba-branded.
   - Backwards compatible with your current HTML IDs/classes where possible.
   ========================================================= */

(() => {
  "use strict";

  console.log("RDT institutional app.js loaded");

  // ------------------------------
  // Config (Córdoba brand)
  // ------------------------------
  const BRAND = {
    name: "Cordoba Research Group",
    short: "CRG",
    version: "RDT v2.0.0",
    colors: {
      gold: "9A690F",
      goldDark: "845F0F",
      cream: "FFF7F0",
      ink: "0B0E14",
      muted: "6B7280",
      border: "E5E7EB",
      rail: "F3F4F6",
      callout: "F6F1E8",
      teal: "0EA5A6" // used sparingly for “highlight strip” effect (optional)
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
  // Utilities
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

  function pct(x) {
    if (!Number.isFinite(x)) return "—";
    return `${(x * 100).toFixed(1)}%`;
  }

  function safeNum(v) {
    const n = Number(String(v).replace(/,/g, ""));
    return Number.isFinite(n) ? n : null;
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

  // Mailto: avoid URLSearchParams (+ issues), force CRLF
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
  const DRAFT_KEY = "crg_rdt_draft_v2";

  const DRAFT_FIELDS = [
    "noteType","title","topic",
    "authorLastName","authorFirstName","authorPhone",
    "authorPhoneCountry","authorPhoneNational",
    "analysis","keyTakeaways","content","cordobaView",
    "ticker","crgRating","targetPrice",
    "modelLink","valuationSummary","keyAssumptions","scenarioNotes"
    // files excluded by browser security
  ];

  function snapshotDraft() {
    const draft = {};
    DRAFT_FIELDS.forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      draft[id] = el.value ?? "";
    });

    // Coauthors (dynamic)
    const coAuthors = $$(".coauthor-entry").map(entry => ({
      lastName: safeTrim($(".coauthor-lastname", entry)?.value),
      firstName: safeTrim($(".coauthor-firstname", entry)?.value),
      phone: safeTrim($(".coauthor-phone", entry)?.value),
      cc: safeTrim($(".coauthor-country", entry)?.value),
      local: safeTrim($(".coauthor-phone-local", entry)?.value)
    })).filter(x => x.lastName || x.firstName || x.phone || x.local);

    draft.__coAuthors = coAuthors;

    // Chart range + last fetched ticker
    const chartRange = $("#chartRange")?.value || "";
    draft.__chartRange = chartRange;

    // Stats (so UI restores)
    draft.__equityStats = equityStats || null;

    // Last updated
    draft.__savedAt = new Date().toISOString();

    return draft;
  }

  function saveDraftNow() {
    try {
      const draft = snapshotDraft();
      localStorage.setItem(DRAFT_KEY, JSON.stringify(draft));
      setDraftStatus("Saved");
    } catch (_) {
      // ignore
    }
  }

  let draftSaveTimer = null;
  function scheduleDraftSave() {
    setDraftStatus("Saving…");
    clearTimeout(draftSaveTimer);
    draftSaveTimer = setTimeout(() => {
      saveDraftNow();
    }, 350);
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

  function applyDraft(draft) {
    if (!draft) return;
    DRAFT_FIELDS.forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      if (typeof draft[id] === "string") el.value = draft[id];
    });

    // Rebuild coauthors
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
          understandingWireCoauthorPhone(node); // sync hidden
          // if explicit phone saved, keep it (but prefer rebuilt)
          const hidden = $(".coauthor-phone", node);
          if (hidden) hidden.value = ca.phone || hidden.value || "";
          list.appendChild(node);
        });
      }
    }

    // Chart range restore
    if (draft.__chartRange && $("#chartRange")) $("#chartRange").value = draft.__chartRange;

    // Stats restore
    if (draft.__equityStats) {
      equityStats = draft.__equityStats;
      paintEquityStats();
    }

    // Sync phone (primary)
    syncPrimaryPhone();
  }

  function clearDraft() {
    try { localStorage.removeItem(DRAFT_KEY); } catch(_) {}
  }

  // Optional “Saved” indicator if HTML includes it
  function setDraftStatus(text) {
    const el = document.getElementById("draftStatus");
    if (el) el.textContent = text || "";
  }

  // ------------------------------
  // Phone wiring (primary + coauthors)
  // ------------------------------
  const authorPhoneCountryEl = document.getElementById("authorPhoneCountry");
  const authorPhoneNationalEl = document.getElementById("authorPhoneNational");
  const authorPhoneHiddenEl = document.getElementById("authorPhone"); // source of truth

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
  // Co-author management (institutional)
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

  function understandingWireCoauthorPhone(coAuthorDiv) {
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

    // make optional (even if future HTML marks required)
    const phoneHidden = $(".coauthor-phone", coAuthorDiv);
    if (phoneHidden) phoneHidden.required = false;

    // wire
    understandingWireCoauthorPhone(coAuthorDiv);

    // autosave events
    ["input","change","keyup"].forEach(evt => {
      coAuthorDiv.addEventListener(evt, () => scheduleDraftSave(), { passive: true });
    });

    return coAuthorDiv;
  }

  if (addCoAuthorBtn && coAuthorsList) {
    addCoAuthorBtn.addEventListener("click", () => {
      const node = createCoauthorNode();
      coAuthorsList.appendChild(node);
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
  // Completion meter (institutional core)
  // ------------------------------
  const completionTextEl = document.getElementById("completionText");
  const completionBarEl = document.getElementById("completionBar");

  function isFilled(el) {
    if (!el) return false;
    if (el.type === "file") return el.files && el.files.length > 0;
    return safeTrim(el.value).length > 0;
  }

  // “Core” should reflect publish reality:
  // - For non-equity: noteType, title, topic, author name, key takeaways, analysis
  // - Córdoba View is optional in your HTML, but institutional notes usually require a house view;
  //   keep it as core if present in UX.
  const baseCoreIds = [
    "noteType","title","topic",
    "authorLastName","authorFirstName",
    "keyTakeaways","analysis"
  ];

  const equityCoreIds = ["crgRating"]; // keep strict minimal; ticker/target/model optional in practice

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
  // Attachment summary (models)
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

  function escapeHtml(s) {
    return (s || "").toString()
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  if (modelFilesEl) {
    modelFilesEl.addEventListener("change", () => {
      updateAttachmentSummary();
      updateCompletionMeter();
      scheduleDraftSave();
    });
  }

  // ------------------------------
  // Reset (with confirm)
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
  // Email to CRG (prefilled)
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
    const subject = [
      noteType, formatDateShortISO(now), title ? `— ${title}` : ""
    ].filter(Boolean).join(" ");

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

  function updateUpsideDisplay() {
    paintEquityStats();
  }

  if (targetPriceEl) {
    targetPriceEl.addEventListener("input", () => {
      updateUpsideDisplay();
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

      renderChart({
        labels,
        values,
        title: `${tickerVal.toUpperCase()} Close`
      });

      // allow chart to render before capture
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

  // ------------------------------
  // Word export (Peel Hunt-style, Córdoba-branded)
  // ------------------------------
  function linesToParagraphs(text, spacingAfter = 140, style = {}) {
    const lines = (text || "").split("\n");
    return lines.map(line => {
      if (line.trim() === "") {
        return new docx.Paragraph({ text: "", spacing: { after: spacingAfter } });
      }
      return new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: line,
            ...style
          })
        ],
        spacing: { after: spacingAfter }
      });
    });
  }

  function bulletLines(text, spacingAfter = 100) {
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
    return bullets.length ? bullets : [new docx.Paragraph({ text: "—", spacing: { after: spacingAfter } })];
  }

  function coAuthorLine(coAuthor) {
    const ln = safeTrim(coAuthor.lastName).toUpperCase();
    const fn = safeTrim(coAuthor.firstName).toUpperCase();
    const ph = naIfBlank(coAuthor.phone);
    return `${ln}, ${fn} (${ph})`;
  }

  function hyperlinkParagraph(label, url) {
    const safeUrl = safeTrim(url);
    if (!safeUrl) return null;

    return new docx.Paragraph({
      children: [
        new docx.TextRun({ text: label, bold: true }),
        new docx.TextRun({ text: " " }),
        new docx.ExternalHyperlink({
          children: [new docx.TextRun({ text: safeUrl, style: "Hyperlink" })],
          link: safeUrl
        })
      ],
      spacing: { after: 120 }
    });
  }

  async function addImages(files) {
    const imageParagraphs = [];
    const list = Array.from(files || []);
    for (let i = 0; i < list.length; i++) {
      const file = list[i];
      try {
        const arrayBuffer = await file.arrayBuffer();
        const fileNameWithoutExt = file.name.replace(/\.[^/.]+$/, "");

        imageParagraphs.push(
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: arrayBuffer,
                transformation: { width: 520, height: 360 }
              })
            ],
            spacing: { before: 180, after: 90 },
            alignment: docx.AlignmentType.CENTER
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: `Figure ${i + 1}: ${fileNameWithoutExt}`,
                italics: true,
                size: 18,
                font: BRAND.fonts.body
              })
            ],
            spacing: { after: 240 },
            alignment: docx.AlignmentType.CENTER
          })
        );
      } catch (error) {
        console.error(`Error processing image ${file.name}:`, error);
      }
    }
    return imageParagraphs;
  }

  function hr(spacingAfter = 200) {
    return new docx.Paragraph({
      border: { bottom: { color: BRAND.colors.ink, space: 1, style: docx.BorderStyle.SINGLE, size: 4 } },
      spacing: { after: spacingAfter }
    });
  }

  function heading(text, size = 28, extra = {}) {
    return new docx.Paragraph({
      children: [
        new docx.TextRun({
          text,
          bold: true,
          size,
          font: BRAND.fonts.heading,
          color: BRAND.colors.ink,
          ...extra
        })
      ],
      spacing: { after: 140 }
    });
  }

  function subheading(text, size = 20, extra = {}) {
    return new docx.Paragraph({
      children: [
        new docx.TextRun({
          text,
          bold: true,
          size,
          font: BRAND.fonts.body,
          color: BRAND.colors.ink,
          ...extra
        })
      ],
      spacing: { after: 110 }
    });
  }

  function smallLabel(text, color = BRAND.colors.muted) {
    return new docx.Paragraph({
      children: [
        new docx.TextRun({
          text,
          size: 18,
          font: BRAND.fonts.body,
          color
        })
      ],
      spacing: { after: 80 }
    });
  }

  function shadedBox(children, hexFill, padding = 220) {
    return new docx.Table({
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
              shading: { fill: hexFill },
              margins: { top: padding, bottom: padding, left: padding, right: padding },
              children
            })
          ]
        })
      ]
    });
  }

  function bannerBlock(noteType, dateStr) {
    // Peel Hunt uses a full-width band with big category + sub-brand.
    // We recreate that language with Córdoba gold.
    const left = [
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: noteType.toUpperCase(),
            font: BRAND.fonts.heading,
            size: 48,
            bold: true,
            color: "FFFFFF"
          })
        ],
        spacing: { after: 60 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: noteType.includes("Macro") ? "MACRO INSIGHTS" :
                  noteType.includes("Commodity") ? "COMMODITY INSIGHTS" :
                  noteType.includes("Fixed") ? "FIXED INCOME" :
                  noteType.includes("Equity") ? "EQUITY RESEARCH" :
                  "RESEARCH NOTE",
            font: BRAND.fonts.body,
            size: 22,
            bold: true,
            color: "FFFFFF"
          })
        ],
        spacing: { after: 40 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: dateStr,
            font: BRAND.fonts.body,
            size: 18,
            color: "FFFFFF"
          })
        ]
      })
    ];

    const right = [
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: BRAND.short,
            font: BRAND.fonts.body,
            size: 22,
            bold: true,
            color: "FFFFFF"
          })
        ],
        alignment: docx.AlignmentType.RIGHT,
        spacing: { after: 30 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "MARKETING COMMUNICATION",
            font: BRAND.fonts.body,
            size: 14,
            color: "FFFFFF"
          })
        ],
        alignment: docx.AlignmentType.RIGHT
      })
    ];

    return new docx.Table({
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
              shading: { fill: BRAND.colors.goldDark },
              width: { size: 72, type: docx.WidthType.PERCENTAGE },
              margins: { top: 220, bottom: 220, left: 320, right: 180 },
              children: left,
              verticalAlign: docx.VerticalAlign.CENTER
            }),
            new docx.TableCell({
              shading: { fill: BRAND.colors.goldDark },
              width: { size: 28, type: docx.WidthType.PERCENTAGE },
              margins: { top: 220, bottom: 220, left: 180, right: 320 },
              children: right,
              verticalAlign: docx.VerticalAlign.CENTER
            })
          ]
        })
      ]
    });
  }

  function authorCard(author, coAuthors) {
    const lines = [];

    lines.push(new docx.Paragraph({
      children: [
        new docx.TextRun({ text: `${safeTrim(author.firstName)} ${safeTrim(author.lastName)}`.trim() || "—", bold: true, size: 20 })
      ],
      spacing: { after: 60 }
    }));

    if (author.phoneWrapped) {
      lines.push(new docx.Paragraph({
        children: [new docx.TextRun({ text: author.phoneWrapped, size: 18, color: BRAND.colors.muted })],
        spacing: { after: 80 }
      }));
    }

    if (coAuthors && coAuthors.length) {
      lines.push(new docx.Paragraph({
        children: [new docx.TextRun({ text: "Co-authors", bold: true, size: 18 })],
        spacing: { after: 70 }
      }));
      coAuthors.forEach(ca => {
        lines.push(new docx.Paragraph({
          children: [new docx.TextRun({ text: coAuthorLine(ca), size: 16, color: BRAND.colors.muted })],
          spacing: { after: 55 }
        }));
      });
    }

    return shadedBox(lines, BRAND.colors.rail, 180);
  }

  function contentsBox(noteType) {
    // Peel Hunt right sidebar “Contents” with page refs.
    // We create a smart contents list based on noteType; analysts can still write freely.
    const defaults = {
      "Macro Research": [
        "Overview: Policymakers’ high-wire act",
        "Forecast summary",
        "Assumptions and risks",
        "UK: Headwinds restart growth",
        "US: Demand, fiscal and AI",
        "Eurozone: Multi-speed expansion",
        "China: Growth targets and disinflation",
        "Japan: Policy normalisation watch",
        "Summary of projections"
      ],
      "Fixed Income Research": [
        "Rates: Macro backdrop",
        "Curve and term premium",
        "Credit: Spreads and positioning",
        "Risks and catalysts",
        "Strategy: What we’d do"
      ],
      "Commodity Insights": [
        "Setup: Why this market matters",
        "Supply / demand balance",
        "Cost curve and marginal pricing",
        "Policy and geopolitics",
        "Trade idea / positioning"
      ],
      "Equity Research": [
        "Investment thesis",
        "Tear sheet / valuation",
        "Key drivers",
        "Risks",
        "Catalysts",
        "Appendix / model notes"
      ],
      "General Note": [
        "Executive summary",
        "Analysis",
        "Risks",
        "What we’d watch next"
      ]
    };

    const items = defaults[noteType] || defaults["General Note"];

    const children = [
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "Contents", bold: true, size: 18 })],
        spacing: { after: 90 }
      })
    ];

    items.slice(0, 10).forEach((t, idx) => {
      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({ text: t, size: 16, color: BRAND.colors.ink }),
            new docx.TextRun({ text: "\t" }),
            new docx.TextRun({ text: String(2 + Math.floor(idx / 2)), size: 16, color: BRAND.colors.muted })
          ],
          tabStops: [{ type: docx.TabStopType.RIGHT, position: 8600 }],
          spacing: { after: 55 }
        })
      );
    });

    return shadedBox(children, BRAND.colors.rail, 180);
  }

  function highlightStrip(titleText, bodyText) {
    // Peel Hunt teal strip box on page 1.
    return shadedBox([
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: titleText,
            bold: true,
            size: 18,
            color: "FFFFFF",
            font: BRAND.fonts.body
          })
        ],
        spacing: { after: 70 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: bodyText,
            size: 18,
            color: "FFFFFF",
            font: BRAND.fonts.body
          })
        ]
      })
    ], BRAND.colors.teal, 180);
  }

  function calloutRailFromTakeaways(keyTakeaways) {
    const lines = (keyTakeaways || "").split("\n").map(l => l.replace(/^[-*•]\s*/, "").trim()).filter(Boolean);
    const items = lines.slice(0, 8);

    const children = [];
    children.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "In brief", bold: true, size: 18 })],
        spacing: { after: 90 }
      })
    );

    if (items.length) {
      items.forEach(t => {
        children.push(new docx.Paragraph({
          text: t,
          bullet: { level: 0 },
          spacing: { after: 80 }
        }));
      });
    } else {
      children.push(new docx.Paragraph({ text: "—", spacing: { after: 80 } }));
    }

    return shadedBox(children, BRAND.colors.callout, 180);
  }

  async function createInstitutionalDocument(payload) {
    const {
      noteType, title, topic,
      authorLastName, authorFirstName, authorPhone, authorPhoneSafe,
      coAuthors,
      analysis, keyTakeaways, content, cordobaView,
      imageFiles, dateTimeString,
      ticker, valuationSummary, keyAssumptions, scenarioNotes, modelFiles, modelLink,
      priceChartImageBytes,
      targetPrice, equityStats,
      crgRating
    } = payload;

    // Defensive
    const nt = noteType || "Research Note";

    // Phone formatting (single bracket pair)
    const authorPhonePrintable = authorPhoneSafe ? authorPhoneSafe : naIfBlank(authorPhone);
    const authorPhoneWrapped = authorPhonePrintable ? `(${authorPhonePrintable})` : "(N/A)";

    // Cover (page 1)
    const banner = bannerBlock(nt, dateTimeString);

    const mainTitle = safeTrim(title);
    const mainTopic = safeTrim(topic);

    const mainIntro = [
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: "TOPIC: ", bold: true, size: 18, font: BRAND.fonts.body, color: BRAND.colors.muted }),
          new docx.TextRun({ text: mainTopic || "—", size: 18, font: BRAND.fonts.body, color: BRAND.colors.ink })
        ],
        spacing: { after: 100 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: mainTitle || "—", bold: true, size: 30, font: BRAND.fonts.heading, color: BRAND.colors.ink })
        ],
        spacing: { after: 120 }
      })
    ];

    const side = [
      authorCard(
        {
          firstName: safeTrim(authorFirstName),
          lastName: safeTrim(authorLastName),
          phoneWrapped: authorPhoneWrapped
        },
        coAuthors || []
      ),
      new docx.Paragraph({ spacing: { after: 120 } }),
      contentsBox(nt)
    ];

    // Page1: two-column layout (main + sidebar)
    const page1Table = new docx.Table({
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
              margins: { top: 120, bottom: 120, left: 80, right: 240 },
              children: [
                ...mainIntro,
                // Peel Hunt: a coloured strip highlight — we derive from first takeaways line(s)
                highlightStrip(
                  `${nt.includes("Macro") ? "Global outlook" : "Note highlight"}: ${safeTrim(keyTakeaways.split("\n")[0] || "Key message")}`,
                  safeTrim(keyTakeaways.split("\n").slice(1, 4).join(" ").slice(0, 220)) || "—"
                ),
                new docx.Paragraph({ spacing: { after: 140 } }),
                ...linesToParagraphs(analysis, 140)
              ],
              verticalAlign: docx.VerticalAlign.TOP
            }),
            new docx.TableCell({
              width: { size: 30, type: docx.WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 240, right: 80 },
              children: side,
              verticalAlign: docx.VerticalAlign.TOP
            })
          ]
        })
      ]
    });

    // Page2+: “Overview” with callout rail (Peel Hunt language)
    const pageBreak = new docx.Paragraph({ children: [new docx.PageBreak()] });

    const overviewTitle = nt.includes("Macro") ? "Overview: Policymakers’ high-wire act"
                         : nt.includes("Equity") ? "Investment view"
                         : "Overview";

    const overviewRail = calloutRailFromTakeaways(keyTakeaways);

    const page2Table = new docx.Table({
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
              width: { size: 68, type: docx.WidthType.PERCENTAGE },
              margins: { top: 80, bottom: 80, left: 80, right: 220 },
              children: [
                smallLabel(`${nt.toUpperCase()}  |  ${formatDateShortISO(new Date())}`, BRAND.colors.muted),
                heading(overviewTitle, 34),
                ...linesToParagraphs(analysis, 140)
              ],
              verticalAlign: docx.VerticalAlign.TOP
            }),
            new docx.TableCell({
              width: { size: 32, type: docx.WidthType.PERCENTAGE },
              margins: { top: 80, bottom: 80, left: 220, right: 80 },
              children: [overviewRail],
              verticalAlign: docx.VerticalAlign.TOP
            })
          ]
        })
      ]
    });

    // Equity tear sheet (sell-side structure) — inserted after page2 header if equity
    const equityBlocks = [];
    if (nt === "Equity Research") {
      equityBlocks.push(
        subheading("Tear Sheet", 24),
        hr(140)
      );

      if (safeTrim(ticker)) {
        equityBlocks.push(new docx.Paragraph({
          children: [
            new docx.TextRun({ text: "Ticker / Company: ", bold: true }),
            new docx.TextRun({ text: safeTrim(ticker) })
          ],
          spacing: { after: 90 }
        }));
      }

      if (safeTrim(crgRating)) {
        equityBlocks.push(new docx.Paragraph({
          children: [
            new docx.TextRun({ text: "CRG Rating: ", bold: true }),
            new docx.TextRun({ text: safeTrim(crgRating) })
          ],
          spacing: { after: 90 }
        }));
      }

      const modelLinkPara = hyperlinkParagraph("Model link:", modelLink);
      if (modelLinkPara) equityBlocks.push(modelLinkPara);

      if (priceChartImageBytes) {
        equityBlocks.push(
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "Price Chart", bold: true, size: 22, font: BRAND.fonts.body })],
            spacing: { before: 80, after: 90 }
          }),
          new docx.Paragraph({
            children: [
              new docx.ImageRun({ data: priceChartImageBytes, transformation: { width: 600, height: 260 } })
            ],
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 160 }
          })
        );
      }

      if (equityStats && Number.isFinite(equityStats.currentPrice)) {
        const tp = safeNum(targetPrice);
        const upside = computeUpsideToTarget(equityStats.currentPrice, tp);

        // stats table (clean sell-side feel)
        equityBlocks.push(
          new docx.Table({
            width: { size: 100, type: docx.WidthType.PERCENTAGE },
            borders: {
              top: { style: docx.BorderStyle.SINGLE, color: BRAND.colors.border, size: 2 },
              bottom: { style: docx.BorderStyle.SINGLE, color: BRAND.colors.border, size: 2 },
              left: { style: docx.BorderStyle.SINGLE, color: BRAND.colors.border, size: 2 },
              right: { style: docx.BorderStyle.SINGLE, color: BRAND.colors.border, size: 2 },
              insideHorizontal: { style: docx.BorderStyle.SINGLE, color: BRAND.colors.border, size: 2 },
              insideVertical: { style: docx.BorderStyle.SINGLE, color: BRAND.colors.border, size: 2 }
            },
            rows: [
              new docx.TableRow({
                children: [
                  cellKV("Current price", equityStats.currentPrice.toFixed(2)),
                  cellKV("Volatility (ann.)", Number.isFinite(equityStats.realisedVolAnn) ? pct(equityStats.realisedVolAnn) : "—"),
                  cellKV("Return (range)", Number.isFinite(equityStats.rangeReturn) ? pct(equityStats.rangeReturn) : "—"),
                  cellKV("+/- to target", upside === null ? "—" : pct(upside))
                ]
              })
            ]
          }),
          new docx.Paragraph({ spacing: { after: 140 } })
        );

        if (tp !== null) {
          equityBlocks.push(new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Target price: ", bold: true }),
              new docx.TextRun({ text: tp.toFixed(2) })
            ],
            spacing: { after: 120 }
          }));
        }
      }

      const attachedModelNames = (modelFiles && modelFiles.length) ? Array.from(modelFiles).map(f => f.name) : [];
      equityBlocks.push(subheading("Attachments", 20));
      if (attachedModelNames.length) {
        attachedModelNames.forEach(name => {
          equityBlocks.push(new docx.Paragraph({ text: name, bullet: { level: 0 }, spacing: { after: 70 } }));
        });
      } else {
        equityBlocks.push(new docx.Paragraph({ text: "None uploaded", spacing: { after: 110 } }));
      }

      if (safeTrim(valuationSummary)) {
        equityBlocks.push(subheading("Valuation Summary", 20));
        equityBlocks.push(...linesToParagraphs(valuationSummary, 120));
      }

      if (safeTrim(keyAssumptions)) {
        equityBlocks.push(subheading("Key Assumptions", 20));
        equityBlocks.push(...bulletLines(keyAssumptions, 70));
      }

      if (safeTrim(scenarioNotes)) {
        equityBlocks.push(subheading("Scenario / Sensitivity Notes", 20));
        equityBlocks.push(...linesToParagraphs(scenarioNotes, 120));
      }
    }

    function cellKV(k, v) {
      return new docx.TableCell({
        shading: { fill: "FFFFFF" },
        margins: { top: 140, bottom: 140, left: 160, right: 160 },
        children: [
          new docx.Paragraph({
            children: [new docx.TextRun({ text: k, size: 16, color: BRAND.colors.muted, font: BRAND.fonts.body })],
            spacing: { after: 50 }
          }),
          new docx.Paragraph({
            children: [new docx.TextRun({ text: v, bold: true, size: 20, color: BRAND.colors.ink, font: BRAND.fonts.body })]
          })
        ],
        verticalAlign: docx.VerticalAlign.TOP
      });
    }

    // Remaining core sections (sell-side rhythm)
    const takeawaysHeading = subheading("Key Takeaways", 22);
    const takeawaysBullets = bulletLines(keyTakeaways, 85);

    const bodyHeading = subheading("Analysis and Commentary", 22);
    const analysisParas = linesToParagraphs(analysis, 140);

    const contentParas = safeTrim(content) ? [subheading("Additional Detail", 20), ...linesToParagraphs(content, 140)] : [];

    const cordobaViewParas = safeTrim(cordobaView)
      ? [subheading("The Cordoba View", 22), ...linesToParagraphs(cordobaView, 140)]
      : [];

    const imageParagraphs = await addImages(imageFiles);

    const figuresBlock = imageParagraphs.length
      ? [subheading("Figures and Charts", 22), ...imageParagraphs]
      : [];

    // Build document children
    const children = [
      banner,
      page1Table,
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: BRAND.disclaimers.internal,
            size: 14,
            color: BRAND.colors.muted,
            font: BRAND.fonts.body
          })
        ],
        spacing: { before: 160, after: 0 }
      }),
      pageBreak,
      page2Table,
      new docx.Paragraph({ spacing: { after: 160 } }),
      ...equityBlocks,
      takeawaysHeading,
      ...takeawaysBullets,
      new docx.Paragraph({ spacing: { after: 180 } }),
      bodyHeading,
      ...analysisParas,
      ...contentParas,
      ...cordobaViewParas,
      ...figuresBlock
    ];

    // A4 portrait, institutional header/footer
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
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    text: `${BRAND.short} | ${nt} | Published ${formatDateShortISO(new Date())}`,
                    size: 16,
                    font: BRAND.fonts.body,
                    color: BRAND.colors.muted
                  }),
                  new docx.TextRun({ text: "\t" }),
                  new docx.TextRun({
                    text: BRAND.short,
                    size: 16,
                    font: BRAND.fonts.body,
                    color: BRAND.colors.muted
                  })
                ],
                tabStops: [{ type: docx.TabStopType.RIGHT, position: 9000 }],
                spacing: { after: 80 },
                border: { bottom: { color: BRAND.colors.border, space: 1, style: docx.BorderStyle.SINGLE, size: 2 } }
              })
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
        children
      }]
    });

    return doc;
  }

  // ------------------------------
  // Main submit (export)
  // ------------------------------
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
        // gentle focus
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

        // Hidden field is source-of-truth
        const authorPhone = safeTrim($("#authorPhone")?.value);
        const authorPhoneSafe = naIfBlank(authorPhone);

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

        const now = new Date();
        const dateTimeString = formatDateTime(now);

        const coAuthors = [];
        $$(".coauthor-entry").forEach(entry => {
          const lastName = safeTrim($(".coauthor-lastname", entry)?.value);
          const firstName = safeTrim($(".coauthor-firstname", entry)?.value);
          const phone = safeTrim($(".coauthor-phone", entry)?.value); // hidden combined
          if (lastName || firstName) {
            coAuthors.push({
              lastName: lastName || "",
              firstName: firstName || "",
              phone: naIfBlank(phone)
            });
          }
        });

        const doc = await createInstitutionalDocument({
          noteType, title, topic,
          authorLastName, authorFirstName, authorPhone,
          authorPhoneSafe,
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
          `${title.replace(/[^a-z0-9]/gi, "_").toLowerCase()}_${noteType.replace(/\s+/g, "_").toLowerCase()}_${formatDateShortISO(now)}.docx`;

        saveAs(blob, fileName);

        showMsg("success", `✓ Document "${fileName}" generated successfully.`);
        saveDraftNow(); // keep draft aligned with what was exported
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

  // DOM ready
  window.addEventListener("DOMContentLoaded", init);
})();
