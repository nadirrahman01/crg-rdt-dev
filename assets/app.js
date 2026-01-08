/* assets/app.js
   Cordoba Research Group — Research Documentation Tool (RDT)
   Institutional-grade UI/workflow + Peel Hunt–style Word export scaffold (Cordoba-branded)

   Notes:
   - This script is backward-compatible with your current HTML IDs/classes where possible.
   - It also supports additional “institutional” fields if/when you add them in HTML (it will gracefully ignore missing inputs).
   - Word export is now A4 portrait with a Peel Hunt–inspired first-page layout (banner + right-hand sidebar + contents/callouts style).
*/

console.log("app.js loaded successfully");

window.addEventListener("DOMContentLoaded", () => {
  "use strict";

  // ============================================================
  // DOM helpers
  // ============================================================
  const $ = (id) => document.getElementById(id);
  const q = (sel, root = document) => root.querySelector(sel);
  const qa = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  function on(el, evt, fn, opts) {
    if (!el) return;
    el.addEventListener(evt, fn, opts);
  }

  function safeTrim(v) {
    return (v ?? "").toString().trim();
  }

  function digitsOnly(v) {
    return (v || "").toString().replace(/\D/g, "");
  }

  function naIfBlank(v) {
    const s = safeTrim(v);
    return s ? s : "N/A";
  }

  // ============================================================
  // Date/time formatting (Peel Hunt-esque)
  // ============================================================
  function formatDateLong(date) {
    const months = [
      "January","February","March","April","May","June",
      "July","August","September","October","November","December"
    ];
    const month = months[date.getMonth()];
    const day = date.getDate();
    const year = date.getFullYear();
    return `${day} ${month} ${year}`;
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

    return `${month} ${day}, ${year} ${hours}:${minutes} ${ampm}`;
  }

  function formatDateShortISO(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  // ============================================================
  // Phone formatting (primary + coauthors)
  // ============================================================
  // Visible formatting: groups into readable blocks
  function formatNationalLoose(rawDigits) {
    const d = digitsOnly(rawDigits);
    if (!d) return "";

    const p1 = d.slice(0, 4);
    const p2 = d.slice(4, 7);
    const p3 = d.slice(7, 10);
    const rest = d.slice(10);

    return [p1, p2, p3, rest].filter(Boolean).join(" ");
  }

  // Hidden combined format used for export
  function buildInternationalHyphen(ccDigits, nationalDigits) {
    const cc = digitsOnly(ccDigits);
    const nn = digitsOnly(nationalDigits);
    if (!cc && !nn) return "";
    if (cc && !nn) return `${cc}-`;
    if (!cc && nn) return nn;
    return `${cc}-${nn}`;
  }

  // Primary author phone wiring
  const authorPhoneCountryEl = $("authorPhoneCountry");
  const authorPhoneNationalEl = $("authorPhoneNational");
  const authorPhoneHiddenEl = $("authorPhone"); // source of truth for export

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

  on(authorPhoneNationalEl, "input", formatPrimaryVisible);
  on(authorPhoneNationalEl, "blur", syncPrimaryPhone);
  on(authorPhoneCountryEl, "change", syncPrimaryPhone);
  syncPrimaryPhone();

  // ============================================================
  // Co-author management (institutional: phone optional, clean formatting)
  // ============================================================
  let coAuthorCount = 0;
  const addCoAuthorBtn = $("addCoAuthor");
  const coAuthorsList = $("coAuthorsList");

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
    const ccEl = q(".coauthor-country", coAuthorDiv);
    const nationalEl = q(".coauthor-phone-local", coAuthorDiv);
    const hiddenEl = q(".coauthor-phone", coAuthorDiv);
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

    on(nationalEl, "input", formatVisible);
    on(nationalEl, "blur", syncHidden);
    on(ccEl, "change", syncHidden);
    syncHidden();
  }

  if (addCoAuthorBtn && coAuthorsList) {
    on(addCoAuthorBtn, "click", () => {
      coAuthorCount++;

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
          <input type="text" placeholder="Phone (optional)" class="phone-number coauthor-phone-local" inputmode="numeric" autocomplete="tel-national">
        </div>

        <!-- hidden combined phone -->
        <input type="text" class="coauthor-phone" style="display:none;">
        <button type="button" class="remove-coauthor" data-remove-id="${coAuthorCount}">Remove</button>
      `;

      coAuthorsList.appendChild(coAuthorDiv);

      const phoneHidden = q(".coauthor-phone", coAuthorDiv);
      if (phoneHidden) phoneHidden.required = false;

      wireCoauthorPhone(coAuthorDiv);
      updateCompletionMeter();
      saveDraftDebounced();
    });

    on(document, "click", (e) => {
      const btn = e.target && e.target.closest ? e.target.closest(".remove-coauthor") : null;
      if (!btn) return;
      const id = btn.getAttribute("data-remove-id");
      const div = $(`coauthor-${id}`);
      if (div) div.remove();
      updateCompletionMeter();
      saveDraftDebounced();
    });
  }

  // ============================================================
  // Note type toggles + “institutional” section routing
  // (keeps your current equity toggle; adds hooks for future Macro/FI/Commodity)
  // ============================================================
  const noteTypeEl = $("noteType");
  const equitySectionEl = $("equitySection");
  const macroSectionEl = $("macroSection"); // optional future
  const fiSectionEl = $("fixedIncomeSection"); // optional future
  const commoditySectionEl = $("commoditySection"); // optional future

  function setDisplay(el, show) {
    if (!el) return;
    el.style.display = show ? "block" : "none";
  }

  function toggleSections() {
    const type = safeTrim(noteTypeEl?.value);

    setDisplay(equitySectionEl, type === "Equity Research");
    setDisplay(macroSectionEl, type === "Macro Research");
    setDisplay(fiSectionEl, type === "Fixed Income Research");
    setDisplay(commoditySectionEl, type === "Commodity Insights");

    // completion recalculation
    setTimeout(updateCompletionMeter, 0);
    saveDraftDebounced();
  }

  on(noteTypeEl, "change", toggleSections);
  toggleSections();

  // ============================================================
  // Completion meter (institutional core fields by note type)
  // ============================================================
  const completionTextEl = $("completionText");
  const completionBarEl = $("completionBar");

  function isFilled(el) {
    if (!el) return false;
    if (el.type === "file") return el.files && el.files.length > 0;
    const v = safeTrim(el.value);
    return v.length > 0;
  }

  // Baseline “publishable” minimum
  // (We keep your 8, but make it truly scalable by type)
  const CORE = {
    base: ["noteType", "title", "topic", "authorLastName", "authorFirstName", "keyTakeaways", "analysis", "cordobaView"],
    equity: ["crgRating"], // targetPrice/modelFiles remain optional; ticker optional
    macro: [],            // add later in HTML if you want e.g. ["macroRegion","macroCall"]
    fixedIncome: [],      // add later
    commodity: []         // add later
  };

  function coreIdsForType(type) {
    const ids = CORE.base.slice();
    if (type === "Equity Research") ids.push(...CORE.equity);
    if (type === "Macro Research") ids.push(...CORE.macro);
    if (type === "Fixed Income Research") ids.push(...CORE.fixedIncome);
    if (type === "Commodity Insights") ids.push(...CORE.commodity);
    return ids;
  }

  function updateCompletionMeter() {
    const type = safeTrim(noteTypeEl?.value);
    const ids = coreIdsForType(type);

    let done = 0;
    ids.forEach((id) => {
      const el = $(id);
      if (isFilled(el)) done++;
    });

    const total = ids.length;
    const pct = total ? Math.round((done / total) * 100) : 0;

    if (completionTextEl) completionTextEl.textContent = `${done} / ${total} core fields`;
    if (completionBarEl) completionBarEl.style.width = `${pct}%`;

    const bar = completionBarEl?.parentElement;
    if (bar) bar.setAttribute("aria-valuenow", String(pct));
  }

  // ============================================================
  // Attachment summary strip (model files)
  // ============================================================
  const modelFilesEl = $("modelFiles");
  const attachSummaryHeadEl = $("attachmentSummaryHead");
  const attachSummaryListEl = $("attachmentSummaryList");

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
    attachSummaryListEl.innerHTML = files.map(f => `<div class="attachment-file">${f.name}</div>`).join("");
  }

  on(modelFilesEl, "change", () => {
    updateAttachmentSummary();
    updateCompletionMeter();
    saveDraftDebounced();
  });

  // ============================================================
  // Draft autosave (localStorage) — institutional must-have
  // ============================================================
  const DRAFT_KEY = "crg_rdt_draft_v2";
  const draftStatusEl = $("draftStatus"); // optional future badge in HTML

  function setDraftStatus(text) {
    if (!draftStatusEl) return;
    draftStatusEl.textContent = text;
  }

  function serializeDraft() {
    // Collect all known fields; ignore missing gracefully
    const fileNames = (el) => el?.files ? Array.from(el.files).map(f => f.name) : [];

    // coauthors
    const coAuthors = qa(".coauthor-entry").map(entry => ({
      lastName: safeTrim(q(".coauthor-lastname", entry)?.value),
      firstName: safeTrim(q(".coauthor-firstname", entry)?.value),
      phoneCountry: safeTrim(q(".coauthor-country", entry)?.value),
      phoneLocal: safeTrim(q(".coauthor-phone-local", entry)?.value),
      phoneHidden: safeTrim(q(".coauthor-phone", entry)?.value)
    })).filter(x => x.lastName || x.firstName || x.phoneHidden || x.phoneLocal);

    return {
      v: 2,
      ts: Date.now(),
      fields: {
        noteType: safeTrim($("noteType")?.value),
        title: safeTrim($("title")?.value),
        topic: safeTrim($("topic")?.value),
        authorLastName: safeTrim($("authorLastName")?.value),
        authorFirstName: safeTrim($("authorFirstName")?.value),

        authorPhoneCountry: safeTrim($("authorPhoneCountry")?.value),
        authorPhoneNational: safeTrim($("authorPhoneNational")?.value),
        authorPhone: safeTrim($("authorPhone")?.value),

        keyTakeaways: safeTrim($("keyTakeaways")?.value),
        analysis: safeTrim($("analysis")?.value),
        content: safeTrim($("content")?.value),
        cordobaView: safeTrim($("cordobaView")?.value),

        // equity
        ticker: safeTrim($("ticker")?.value),
        crgRating: safeTrim($("crgRating")?.value),
        targetPrice: safeTrim($("targetPrice")?.value),
        chartRange: safeTrim($("chartRange")?.value),
        modelLink: safeTrim($("modelLink")?.value),
        valuationSummary: safeTrim($("valuationSummary")?.value),
        keyAssumptions: safeTrim($("keyAssumptions")?.value),
        scenarioNotes: safeTrim($("scenarioNotes")?.value),

        // optional future institutional fields (safe no-ops if missing)
        authorEmail: safeTrim($("authorEmail")?.value),
        sector: safeTrim($("sector")?.value),
        region: safeTrim($("region")?.value),
        complianceTag: safeTrim($("complianceTag")?.value),
        distribution: safeTrim($("distribution")?.value)
      },
      files: {
        modelFiles: fileNames($("modelFiles")),
        imageUpload: fileNames($("imageUpload"))
      },
      coAuthors,
      equityStats
    };
  }

  function applyDraft(d) {
    if (!d || !d.fields) return;

    const F = d.fields;

    const setVal = (id, v) => {
      const el = $(id);
      if (!el) return;
      el.value = v ?? "";
    };

    setVal("noteType", F.noteType);
    setVal("title", F.title);
    setVal("topic", F.topic);
    setVal("authorLastName", F.authorLastName);
    setVal("authorFirstName", F.authorFirstName);

    setVal("authorPhoneCountry", F.authorPhoneCountry);
    setVal("authorPhoneNational", F.authorPhoneNational);
    setVal("authorPhone", F.authorPhone);

    setVal("keyTakeaways", F.keyTakeaways);
    setVal("analysis", F.analysis);
    setVal("content", F.content);
    setVal("cordobaView", F.cordobaView);

    setVal("ticker", F.ticker);
    setVal("crgRating", F.crgRating);
    setVal("targetPrice", F.targetPrice);
    setVal("chartRange", F.chartRange);
    setVal("modelLink", F.modelLink);
    setVal("valuationSummary", F.valuationSummary);
    setVal("keyAssumptions", F.keyAssumptions);
    setVal("scenarioNotes", F.scenarioNotes);

    // optional future
    setVal("authorEmail", F.authorEmail);
    setVal("sector", F.sector);
    setVal("region", F.region);
    setVal("complianceTag", F.complianceTag);
    setVal("distribution", F.distribution);

    // restore coauthors UI
    if (coAuthorsList) coAuthorsList.innerHTML = "";
    coAuthorCount = 0;

    (d.coAuthors || []).forEach(ca => {
      coAuthorCount++;
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
          <input type="text" placeholder="Phone (optional)" class="phone-number coauthor-phone-local" inputmode="numeric" autocomplete="tel-national">
        </div>

        <input type="text" class="coauthor-phone" style="display:none;">
        <button type="button" class="remove-coauthor" data-remove-id="${coAuthorCount}">Remove</button>
      `;
      coAuthorsList.appendChild(div);

      q(".coauthor-lastname", div).value = ca.lastName || "";
      q(".coauthor-firstname", div).value = ca.firstName || "";

      const cc = q(".coauthor-country", div);
      const local = q(".coauthor-phone-local", div);
      const hidden = q(".coauthor-phone", div);

      if (cc) cc.value = ca.phoneCountry ?? "44";
      if (local) local.value = ca.phoneLocal || "";
      if (hidden) hidden.value = ca.phoneHidden || "";

      wireCoauthorPhone(div);
    });

    // restore stats (chart must be refetched to embed image)
    equityStats = d.equityStats || equityStats;

    // rerun formatting + sections
    formatPrimaryVisible();
    syncPrimaryPhone();
    toggleSections();
    updateAttachmentSummary();
    updateCompletionMeter();
    updateUpsideDisplay();
  }

  function saveDraft() {
    try {
      setDraftStatus("Saving…");
      const payload = serializeDraft();
      localStorage.setItem(DRAFT_KEY, JSON.stringify(payload));
      setDraftStatus("Saved");
    } catch (e) {
      console.warn("Draft save failed:", e);
      setDraftStatus("Draft not saved");
    }
  }

  let saveT = null;
  function saveDraftDebounced() {
    if (saveT) clearTimeout(saveT);
    saveT = setTimeout(saveDraft, 350);
  }

  function loadDraft() {
    try {
      const raw = localStorage.getItem(DRAFT_KEY);
      if (!raw) return;
      const d = JSON.parse(raw);
      if (!d || !d.fields) return;
      applyDraft(d);
      setDraftStatus("Draft restored");
      setTimeout(() => setDraftStatus("Saved"), 800);
    } catch (e) {
      console.warn("Draft load failed:", e);
    }
  }

  // Save draft on changes across the form
  ["input", "change", "keyup"].forEach(evt => {
    on(document, evt, (e) => {
      const t = e.target;
      if (!t || !t.closest) return;
      if (!t.closest("#researchForm")) return;
      updateCompletionMeter();
      saveDraftDebounced();
    }, { passive: true });
  });

  // Restore draft on load
  loadDraft();

  // ============================================================
  // Email to CRG (prefilled mailto) — same behaviour, cleaner encoding
  // ============================================================
  const emailToCrgBtn = $("emailToCrgBtn");

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
    if (t.includes("commodity")) return "uhayd@cordobarg.com";
    return "";
  }

  function buildCrgEmailPayload() {
    const noteType = safeTrim($("noteType")?.value || "Research Note");
    const title = safeTrim($("title")?.value);
    const topic = safeTrim($("topic")?.value);

    const authorFirstName = safeTrim($("authorFirstName")?.value);
    const authorLastName = safeTrim($("authorLastName")?.value);

    const ticker = safeTrim($("ticker")?.value);
    const crgRating = safeTrim($("crgRating")?.value);
    const targetPrice = safeTrim($("targetPrice")?.value);

    const now = new Date();
    const dateShort = formatDateShortISO(now);
    const dateLong = formatDateTime(now);

    const subjectParts = [noteType, dateShort, title ? `— ${title}` : ""].filter(Boolean);
    const subject = subjectParts.join(" ");

    const authorLine = [authorFirstName, authorLastName].filter(Boolean).join(" ").trim();

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
      `Generated: ${dateLong}`
    ].filter(Boolean);

    paragraphs.push(metaLines.join("\n"));
    paragraphs.push("Best,");
    paragraphs.push(authorLine || "");

    const body = paragraphs.join("\n\n");
    const cc = ccForNoteType(noteType);

    return { subject, body, cc };
  }

  on(emailToCrgBtn, "click", () => {
    const { subject, body, cc } = buildCrgEmailPayload();
    const to = "research@cordobarg.com";
    const mailto = buildMailto(to, cc, subject, body);
    window.location.href = mailto;
  });

  // ============================================================
  // Images to Word (user uploads)
  // ============================================================
  async function addImages(files) {
    const imageParagraphs = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      try {
        const arrayBuffer = await file.arrayBuffer();
        const fileNameWithoutExt = file.name.replace(/\.[^/.]+$/, "");

        imageParagraphs.push(
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: arrayBuffer,
                // institutional: slightly wider, keeps within margins
                transformation: { width: 560, height: 350 }
              })
            ],
            spacing: { before: 220, after: 110 },
            alignment: docx.AlignmentType.CENTER
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: `Figure ${i + 1}: ${fileNameWithoutExt}`,
                italics: true,
                size: 18,
                font: "Times New Roman",
                color: "4B4B4B"
              })
            ],
            spacing: { after: 260 },
            alignment: docx.AlignmentType.CENTER
          })
        );
      } catch (error) {
        console.error(`Error processing image ${file.name}:`, error);
      }
    }
    return imageParagraphs;
  }

  function linesToParagraphs(text, spacingAfter = 140) {
    const lines = (text || "").split("\n");
    return lines.map((line) => {
      if (safeTrim(line) === "") {
        return new docx.Paragraph({ text: "", spacing: { after: spacingAfter } });
      }
      return new docx.Paragraph({
        children: [new docx.TextRun({ text: line, font: "Times New Roman" })],
        spacing: { after: spacingAfter }
      });
    });
  }

  function bulletsFromLines(text, spacingAfter = 90) {
    const lines = (text || "").split("\n");
    const out = [];
    lines.forEach(line => {
      const s = safeTrim(line);
      if (!s) {
        out.push(new docx.Paragraph({ text: "", spacing: { after: spacingAfter } }));
        return;
      }
      const clean = s.replace(/^[-*•]\s*/, "").trim();
      out.push(
        new docx.Paragraph({
          text: clean,
          bullet: { level: 0 },
          spacing: { after: spacingAfter }
        })
      );
    });
    return out;
  }

  function hyperlinkParagraph(label, url) {
    const safeUrl = safeTrim(url);
    if (!safeUrl) return null;

    return new docx.Paragraph({
      children: [
        new docx.TextRun({ text: label, bold: true, font: "Times New Roman" }),
        new docx.TextRun({ text: " " }),
        new docx.ExternalHyperlink({
          children: [new docx.TextRun({ text: safeUrl, style: "Hyperlink" })],
          link: safeUrl
        })
      ],
      spacing: { after: 120 }
    });
  }

  // ============================================================
  // Price chart (Stooq -> Chart.js -> Word image)
  // ============================================================
  let priceChart = null;
  let priceChartImageBytes = null;

  let equityStats = {
    currentPrice: null,
    realisedVolAnn: null,
    rangeReturn: null
  };

  const chartStatus = $("chartStatus");
  const fetchChartBtn = $("fetchPriceChart");
  const chartRangeEl = $("chartRange");
  const priceChartCanvas = $("priceChart");
  const targetPriceEl = $("targetPrice");

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
    if (lines.length < 5) throw new Error("Not enough data returned. Check ticker.");

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
      priceChart.destroy();
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

  // stats helpers
  function pct(x) { return `${(x * 100).toFixed(1)}%`; }

  function safeNum(v) {
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
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

  function setText(id, text) {
    const el = $(id);
    if (el) el.textContent = text;
  }

  function computeUpsideToTarget(currentPrice, targetPrice) {
    if (!currentPrice || !targetPrice) return null;
    return (targetPrice / currentPrice) - 1;
  }

  function updateUpsideDisplay() {
    const current = equityStats.currentPrice;
    const target = safeNum(targetPriceEl?.value);
    const up = computeUpsideToTarget(current, target);
    setText("upsideToTarget", up === null ? "—" : pct(up));
  }

  on(targetPriceEl, "input", () => {
    updateUpsideDisplay();
    updateCompletionMeter();
    saveDraftDebounced();
  });

  async function buildPriceChart() {
    try {
      const tickerVal = safeTrim($("ticker")?.value);
      if (!tickerVal) throw new Error("Enter a ticker first.");

      const range = chartRangeEl ? chartRangeEl.value : "6mo";
      const symbol = stooqSymbolFromTicker(tickerVal);
      if (!symbol) throw new Error("Invalid ticker.");

      if (chartStatus) chartStatus.textContent = "Fetching price data…";

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

      await new Promise(r => setTimeout(r, 160));
      priceChartImageBytes = canvasToPngBytes(priceChartCanvas);

      const closes = values;
      const currentPrice = closes[closes.length - 1];
      const startPrice = closes[0];

      const rangeReturn = (startPrice && currentPrice) ? (currentPrice / startPrice) - 1 : null;

      const dailyRets = computeDailyReturns(closes);
      const volDaily = stddev(dailyRets);
      const realisedVolAnn = (volDaily !== null) ? volDaily * Math.sqrt(252) : null;

      equityStats.currentPrice = currentPrice;
      equityStats.rangeReturn = rangeReturn;
      equityStats.realisedVolAnn = realisedVolAnn;

      setText("currentPrice", currentPrice ? currentPrice.toFixed(2) : "—");
      setText("rangeReturn", rangeReturn === null ? "—" : pct(rangeReturn));
      setText("realisedVol", realisedVolAnn === null ? "—" : pct(realisedVolAnn));

      updateUpsideDisplay();

      if (chartStatus) chartStatus.textContent = `✓ Chart ready (${range.toUpperCase()})`;
      saveDraftDebounced();
    } catch (e) {
      priceChartImageBytes = null;
      equityStats = { currentPrice: null, realisedVolAnn: null, rangeReturn: null };

      setText("currentPrice", "—");
      setText("rangeReturn", "—");
      setText("realisedVol", "—");
      setText("upsideToTarget", "—");

      if (chartStatus) chartStatus.textContent = `✗ ${e.message}`;
    } finally {
      updateCompletionMeter();
    }
  }

  on(fetchChartBtn, "click", buildPriceChart);

  // ============================================================
  // Reset form button (keeps your behaviour + clears draft)
  // ============================================================
  const resetBtn = $("resetFormBtn");
  const formEl = $("researchForm");

  function clearChartUI() {
    setText("currentPrice", "—");
    setText("realisedVol", "—");
    setText("rangeReturn", "—");
    setText("upsideToTarget", "—");
    if (chartStatus) chartStatus.textContent = "";

    if (priceChart) {
      try { priceChart.destroy(); } catch (_) {}
      priceChart = null;
    }

    priceChartImageBytes = null;
    equityStats = { currentPrice: null, realisedVolAnn: null, rangeReturn: null };
  }

  on(resetBtn, "click", () => {
    if (!formEl) return;
    const ok = confirm("Reset the form? This will clear all fields on this page.");
    if (!ok) return;

    formEl.reset();

    if (coAuthorsList) coAuthorsList.innerHTML = "";
    coAuthorCount = 0;

    if (modelFilesEl) modelFilesEl.value = "";
    updateAttachmentSummary();

    const messageDiv = $("message");
    if (messageDiv) {
      messageDiv.className = "message";
      messageDiv.textContent = "";
      messageDiv.style.display = "none";
    }

    clearChartUI();
    syncPrimaryPhone();
    toggleSections();
    updateCompletionMeter();

    try { localStorage.removeItem(DRAFT_KEY); } catch (_) {}
    setDraftStatus("Draft cleared");
  });

  // ============================================================
  // WORD EXPORT — Peel Hunt–style (Cordoba branded)
  // A4 portrait, institutional banner, right sidebar, and templated sections.
  // ============================================================

  // Cordoba theme palette (Word-safe hex; no #)
  const CRG_COLORS = {
    gold: "9A690F",
    goldDark: "845F0F",
    cream: "FFF7F0",
    ink: "0B0E14",
    muted: "6B7280",
    line: "D1D5DB",
    lightPanel: "F7F4EF"
  };

  function coAuthorLine(coAuthor) {
    const ln = safeTrim(coAuthor.lastName).toUpperCase();
    const fn = safeTrim(coAuthor.firstName).toUpperCase();
    const ph = naIfBlank(coAuthor.phone);
    return `${ln}, ${fn} (${ph})`;
  }

  // Peel Hunt-inspired: a structured “Contents” list based on what sections exist
  function buildContentsList(noteType, hasEquity, hasImages, hasCordobaView, hasContent) {
    const items = [];

    // Always
    items.push("Overview");
    items.push("Key takeaways");
    items.push("Analysis");

    if (hasContent) items.push("Additional content");
    if (hasEquity) items.push("Equity: valuation & model");
    if (hasEquity) items.push("Equity: market data & chart");
    if (hasCordobaView) items.push("The Cordoba view");
    if (hasImages) items.push("Figures & charts");

    // Type-specific (future)
    if (noteType === "Macro Research") items.push("Macro: scenarios & risks");
    if (noteType === "Fixed Income Research") items.push("Fixed income: security details");
    if (noteType === "Commodity Insights") items.push("Commodity: drivers & positioning");

    // Make unique while preserving order
    return Array.from(new Set(items));
  }

  // Shared text runs
  function TR(text, opts = {}) {
    return new docx.TextRun({
      text,
      font: opts.font || "Times New Roman",
      size: opts.size ?? 22, // 11pt
      bold: !!opts.bold,
      italics: !!opts.italics,
      color: opts.color || CRG_COLORS.ink
    });
  }

  function H(text, size = 28) { // 14pt
    return new docx.Paragraph({
      children: [TR(text, { bold: true, size })],
      spacing: { before: 140, after: 120 }
    });
  }

  function small(text, bold = false) {
    return new docx.Paragraph({
      children: [TR(text, { size: 18, bold, color: CRG_COLORS.muted })],
      spacing: { after: 60 }
    });
  }

  function rule(after = 220) {
    return new docx.Paragraph({
      border: { bottom: { color: CRG_COLORS.line, space: 1, style: docx.BorderStyle.SINGLE, size: 8 } },
      spacing: { after }
    });
  }

  function shadedBoxParagraph(text, shade = CRG_COLORS.lightPanel) {
    return new docx.Paragraph({
      children: [TR(text, { size: 20, color: CRG_COLORS.ink })],
      shading: { type: docx.ShadingType.CLEAR, color: "auto", fill: shade },
      spacing: { before: 80, after: 120 },
      indent: { left: 120, right: 120 },
    });
  }

  // Banner table: full width, gold strip with white text (Peel Hunt vibe, Cordoba palette)
  function buildBanner(noteType, dateLong) {
    const left = new docx.Paragraph({
      children: [
        TR(noteType.toUpperCase(), { bold: true, size: 28, color: "FFFFFF" })
      ],
      spacing: { after: 0 }
    });

    const right = new docx.Paragraph({
      children: [
        TR(dateLong, { size: 20, color: "FFFFFF" })
      ],
      alignment: docx.AlignmentType.RIGHT,
      spacing: { after: 0 }
    });

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              children: [left],
              shading: { type: docx.ShadingType.CLEAR, color: "auto", fill: CRG_COLORS.goldDark },
              borders: { top: { style: docx.BorderStyle.NONE }, bottom: { style: docx.BorderStyle.NONE }, left: { style: docx.BorderStyle.NONE }, right: { style: docx.BorderStyle.NONE } },
              margins: { top: 140, bottom: 140, left: 220, right: 220 }
            }),
            new docx.TableCell({
              children: [right],
              shading: { type: docx.ShadingType.CLEAR, color: "auto", fill: CRG_COLORS.goldDark },
              borders: { top: { style: docx.BorderStyle.NONE }, bottom: { style: docx.BorderStyle.NONE }, left: { style: docx.BorderStyle.NONE }, right: { style: docx.BorderStyle.NONE } },
              margins: { top: 140, bottom: 140, left: 220, right: 220 }
            })
          ]
        })
      ]
    });
  }

  // Peel Hunt–style “first page”: two-column table:
  // Left: title + overview box + first bullets
  // Right: analyst box + contents list
  function buildFirstPageLayout(payload) {
    const {
      noteType, title, topic,
      authorLastName, authorFirstName, authorPhonePrintable,
      authorEmail,
      coAuthors,
      dateLong,
      keyTakeaways,
      analysis,
      hasEquity,
      hasImages,
      hasCordobaView,
      hasContent
    } = payload;

    const contents = buildContentsList(noteType, hasEquity, hasImages, hasCordobaView, hasContent);

    // LEFT column
    const topicLine = new docx.Paragraph({
      children: [
        TR("TOPIC: ", { bold: true, size: 18, color: CRG_COLORS.muted }),
        TR(topic || "—", { size: 18, color: CRG_COLORS.muted })
      ],
      spacing: { after: 60 }
    });

    const titlePara = new docx.Paragraph({
      children: [TR(title || "Untitled research note", { bold: true, size: 44, color: CRG_COLORS.ink })],
      spacing: { after: 160 }
    });

    // Overview: take first ~2 paragraphs of analysis if present, else from cordobaView/content
    const overviewText = (() => {
      const a = safeTrim(analysis);
      if (a) {
        const parts = a.split("\n").map(s => safeTrim(s)).filter(Boolean);
        return parts.slice(0, 4).join(" ");
      }
      return "—";
    })();

    const overviewHead = new docx.Paragraph({
      children: [TR("Overview", { bold: true, size: 22, color: CRG_COLORS.ink })],
      spacing: { after: 80 }
    });

    const overviewBox = shadedBoxParagraph(overviewText, CRG_COLORS.lightPanel);

    const ktHead = new docx.Paragraph({
      children: [TR("Key takeaways", { bold: true, size: 22 })],
      spacing: { before: 120, after: 80 }
    });

    const ktBullets = bulletsFromLines(keyTakeaways || "", 70).slice(0, 6);

    const leftChildren = [
      topicLine,
      titlePara,
      overviewHead,
      overviewBox,
      ktHead,
      ...ktBullets
    ];

    // RIGHT column (analyst + contents) — Peel Hunt vibe
    const analystName = `${safeTrim(authorFirstName)} ${safeTrim(authorLastName)}`.trim();
    const analystPhone = authorPhonePrintable ? `+${authorPhonePrintable.replace("-", " ")}` : "N/A";
    const analystEmailLine = safeTrim(authorEmail);

    const analystHead = new docx.Paragraph({
      children: [TR("Author", { bold: true, size: 20, color: CRG_COLORS.ink })],
      spacing: { after: 70 }
    });

    const analystBox = new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              children: [
                analystHead,
                new docx.Paragraph({ children: [TR(analystName || "—", { bold: true, size: 22 })], spacing: { after: 40 } }),
                new docx.Paragraph({ children: [TR(analystPhone, { size: 18, color: CRG_COLORS.muted })], spacing: { after: 40 } }),
                ...(analystEmailLine
                  ? [new docx.Paragraph({ children: [TR(analystEmailLine, { size: 18, color: CRG_COLORS.muted })], spacing: { after: 60 } })]
                  : []),
                ...(coAuthors && coAuthors.length
                  ? [
                      new docx.Paragraph({ children: [TR("Co-authors", { bold: true, size: 18, color: CRG_COLORS.ink })], spacing: { before: 80, after: 50 } }),
                      ...coAuthors.map(ca => new docx.Paragraph({
                        children: [TR(coAuthorLine(ca), { size: 18, color: CRG_COLORS.muted })],
                        spacing: { after: 40 }
                      }))
                    ]
                  : [])
              ],
              shading: { type: docx.ShadingType.CLEAR, color: "auto", fill: "FFFFFF" },
              borders: {
                top: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 6 },
                bottom: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 6 },
                left: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 6 },
                right: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 6 }
              },
              margins: { top: 160, bottom: 160, left: 180, right: 180 }
            })
          ]
        })
      ]
    });

    const contentsHead = new docx.Paragraph({
      children: [TR("Contents", { bold: true, size: 20 })],
      spacing: { before: 180, after: 80 }
    });

    const contentsParas = contents.map(item =>
      new docx.Paragraph({
        children: [TR(item, { size: 18, color: CRG_COLORS.muted })],
        spacing: { after: 50 }
      })
    );

    const rightChildren = [
      small("Cordoba Research Group", true),
      small(`Published: ${dateLong}`),
      analystBox,
      contentsHead,
      ...contentsParas
    ];

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              width: { size: 66, type: docx.WidthType.PERCENTAGE },
              children: leftChildren,
              borders: { top: { style: docx.BorderStyle.NONE }, bottom: { style: docx.BorderStyle.NONE }, left: { style: docx.BorderStyle.NONE }, right: { style: docx.BorderStyle.NONE } },
              margins: { top: 180, bottom: 0, left: 0, right: 240 },
              verticalAlign: docx.VerticalAlign.TOP
            }),
            new docx.TableCell({
              width: { size: 34, type: docx.WidthType.PERCENTAGE },
              children: rightChildren,
              borders: { top: { style: docx.BorderStyle.NONE }, bottom: { style: docx.BorderStyle.NONE }, left: { style: docx.BorderStyle.NONE }, right: { style: docx.BorderStyle.NONE } },
              margins: { top: 180, bottom: 0, left: 240, right: 0 },
              verticalAlign: docx.VerticalAlign.TOP
            })
          ]
        })
      ]
    });
  }

  // Equity “tear sheet” section (sell-side-ish, scalable)
  function buildEquityTearSheet(payload) {
    const {
      ticker,
      crgRating,
      targetPrice,
      equityStats,
      modelLink,
      valuationSummary
    } = payload;

    const tpNum = safeNum(targetPrice);
    const upside = (equityStats?.currentPrice && tpNum) ? computeUpsideToTarget(equityStats.currentPrice, tpNum) : null;

    const rows = [
      ["Ticker", safeTrim(ticker) || "—"],
      ["CRG Rating", safeTrim(crgRating) || "—"],
      ["Current Price", equityStats?.currentPrice ? equityStats.currentPrice.toFixed(2) : "—"],
      ["Target Price", tpNum ? tpNum.toFixed(2) : "—"],
      ["Upside / Downside", upside == null ? "—" : pct(upside)],
      ["Volatility (ann.)", equityStats?.realisedVolAnn == null ? "—" : pct(equityStats.realisedVolAnn)],
      ["Return (range)", equityStats?.rangeReturn == null ? "—" : pct(equityStats.rangeReturn)]
    ];

    const tableRows = rows.map(([k, v]) =>
      new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [new docx.Paragraph({ children: [TR(k, { bold: true, size: 20, color: CRG_COLORS.ink })], spacing: { after: 0 } })],
            shading: { type: docx.ShadingType.CLEAR, color: "auto", fill: CRG_COLORS.lightPanel },
            borders: {
              top: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 4 },
              bottom: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 4 },
              left: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 4 },
              right: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 4 }
            },
            margins: { top: 80, bottom: 80, left: 120, right: 120 }
          }),
          new docx.TableCell({
            children: [new docx.Paragraph({ children: [TR(v, { size: 20, color: CRG_COLORS.ink })], spacing: { after: 0 } })],
            borders: {
              top: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 4 },
              bottom: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 4 },
              left: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 4 },
              right: { style: docx.BorderStyle.SINGLE, color: CRG_COLORS.line, size: 4 }
            },
            margins: { top: 80, bottom: 80, left: 120, right: 120 }
          })
        ]
      })
    );

    const tearSheetTable = new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: tableRows
    });

    const out = [
      H("Equity tear sheet", 30),
      tearSheetTable,
      new docx.Paragraph({ spacing: { after: 120 } })
    ];

    const linkPara = hyperlinkParagraph("Model link:", modelLink);
    if (linkPara) out.push(linkPara);

    if (safeTrim(valuationSummary)) {
      out.push(
        H("Valuation summary", 26),
        ...linesToParagraphs(valuationSummary, 120)
      );
    }

    return out;
  }

  async function createDocument(data) {
    const {
      noteType, title, topic,
      authorLastName, authorFirstName, authorPhone,
      authorPhoneSafe,
      authorEmail,
      coAuthors,
      analysis, keyTakeaways, content, cordobaView,
      imageFiles, dateTimeString,

      ticker, valuationSummary, keyAssumptions, scenarioNotes, modelFiles, modelLink,
      priceChartImageBytes,

      targetPrice,
      equityStats,

      crgRating
    } = data;

    // Derived flags
    const hasEquity = noteType === "Equity Research";
    const hasImages = imageFiles && imageFiles.length > 0;
    const hasCordobaView = safeTrim(cordobaView).length > 0;
    const hasContent = safeTrim(content).length > 0;

    // First page metadata
    const now = new Date();
    const dateLong = formatDateLong(now);

    const authorPhonePrintable = authorPhoneSafe ? authorPhoneSafe : naIfBlank(authorPhone);

    // Build sections
    const imageParagraphs = await addImages(imageFiles);

    const ktBulletsAll = bulletsFromLines(keyTakeaways || "", 90);
    const analysisParagraphs = linesToParagraphs(analysis, 140);
    const contentParagraphs = linesToParagraphs(content, 140);
    const cordobaViewParagraphs = linesToParagraphs(cordobaView, 140);

    // Equity extras
    const attachedModelNames = (modelFiles && modelFiles.length) ? Array.from(modelFiles).map(f => f.name) : [];
    const equityAssumptionsBullets = (safeTrim(keyAssumptions))
      ? bulletsFromLines(keyAssumptions, 70).filter(p => true)
      : [];

    // Compose Word children (Peel Hunt-like)
    const docChildren = [];

    // Banner + first-page two-column layout
    docChildren.push(
      buildBanner(noteType || "Research", dateLong),
      buildFirstPageLayout({
        noteType: noteType || "Research",
        title,
        topic,
        authorLastName,
        authorFirstName,
        authorPhonePrintable,
        authorEmail,
        coAuthors,
        dateLong,
        keyTakeaways,
        analysis,
        hasEquity,
        hasImages,
        hasCordobaView,
        hasContent
      }),
      new docx.Paragraph({ children: [new docx.PageBreak()] })
    );

    // Main body (institutional: clear headings, consistent spacing)
    docChildren.push(
      H("Key takeaways", 30),
      ...ktBulletsAll,
      rule(260),
      H("Analysis", 30),
      ...analysisParagraphs
    );

    if (hasContent) {
      docChildren.push(
        rule(220),
        H("Additional content", 28),
        ...contentParagraphs
      );
    }

    if (hasEquity) {
      docChildren.push(
        rule(220),
        ...buildEquityTearSheet({
          ticker,
          crgRating,
          targetPrice,
          equityStats,
          modelLink,
          valuationSummary
        })
      );

      // Price chart
      if (priceChartImageBytes) {
        docChildren.push(
          H("Price chart", 28),
          new docx.Paragraph({
            children: [new docx.ImageRun({ data: priceChartImageBytes, transformation: { width: 560, height: 260 } })],
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 220 }
          })
        );
      }

      // Key assumptions (bullets)
      if (equityAssumptionsBullets.length) {
        docChildren.push(
          H("Key assumptions", 26),
          ...equityAssumptionsBullets
        );
      }

      // Scenario notes
      if (safeTrim(scenarioNotes)) {
        docChildren.push(
          H("Scenario / sensitivity notes", 26),
          ...linesToParagraphs(scenarioNotes, 120)
        );
      }

      // Model attachments list
      docChildren.push(
        H("Model files", 26)
      );

      if (attachedModelNames.length) {
        attachedModelNames.forEach(name => {
          docChildren.push(new docx.Paragraph({ text: name, bullet: { level: 0 }, spacing: { after: 70 } }));
        });
      } else {
        docChildren.push(new docx.Paragraph({ children: [TR("None uploaded", { size: 20, color: CRG_COLORS.muted })], spacing: { after: 120 } }));
      }
    }

    // Cordoba View
    if (hasCordobaView) {
      docChildren.push(
        rule(220),
        H("The Cordoba view", 30),
        ...cordobaViewParagraphs
      );
    }

    // Figures
    if (imageParagraphs.length > 0) {
      docChildren.push(
        rule(220),
        H("Figures & charts", 30),
        ...imageParagraphs
      );
    }

    // Build doc with Peel Hunt-like header/footer conventions
    const doc = new docx.Document({
      styles: {
        default: {
          document: {
            run: { font: "Times New Roman", size: 22, color: CRG_COLORS.ink },
            paragraph: { spacing: { after: 140 } }
          }
        }
      },
      sections: [{
        properties: {
          page: {
            // A4 portrait
            margin: { top: 720, right: 720, bottom: 720, left: 720 },
            pageSize: { orientation: docx.PageOrientation.PORTRAIT }
          }
        },
        headers: {
          default: new docx.Header({
            children: [
              new docx.Paragraph({
                children: [
                  TR("Cordoba Research Group", { size: 16, color: CRG_COLORS.muted, bold: true }),
                  TR("  |  ", { size: 16, color: CRG_COLORS.muted }),
                  TR(noteType || "Research", { size: 16, color: CRG_COLORS.muted }),
                  TR("  |  ", { size: 16, color: CRG_COLORS.muted }),
                  TR(`Published ${dateTimeString}`, { size: 16, color: CRG_COLORS.muted })
                ],
                alignment: docx.AlignmentType.RIGHT,
                spacing: { after: 80 },
                border: { bottom: { color: CRG_COLORS.line, space: 1, style: docx.BorderStyle.SINGLE, size: 6 } }
              })
            ]
          })
        },
        footers: {
          default: new docx.Footer({
            children: [
              new docx.Paragraph({
                border: { top: { color: CRG_COLORS.line, space: 1, style: docx.BorderStyle.SINGLE, size: 6 } },
                spacing: { after: 0 }
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({ text: "\t" }),
                  TR("Internal marketing communication — not investment advice. Verify all figures before circulation.", { size: 16, italics: true, color: CRG_COLORS.muted }),
                  new docx.TextRun({ text: "\t" }),
                  new docx.TextRun({
                    children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
                    size: 16,
                    italics: true,
                    font: "Times New Roman",
                    color: CRG_COLORS.muted
                  })
                ],
                spacing: { before: 0, after: 0 },
                tabStops: [
                  { type: docx.TabStopType.CENTER, position: 5000 },
                  { type: docx.TabStopType.RIGHT, position: 10000 }
                ]
              })
            ]
          })
        },
        children: docChildren
      }]
    });

    return doc;
  }

  // ============================================================
  // Main form submission
  // ============================================================
  const form = $("researchForm");
  if (form) form.noValidate = true;

  on(form, "submit", async (e) => {
    e.preventDefault();

    const button = q('button[type="submit"]', form);
    const messageDiv = $("message");

    if (button) {
      button.disabled = true;
      button.classList.add("loading");
      button.textContent = "Generating Document…";
    }

    if (messageDiv) {
      messageDiv.className = "message";
      messageDiv.textContent = "";
      messageDiv.style.display = "none";
    }

    try {
      if (typeof docx === "undefined") throw new Error("docx library not loaded. Please refresh the page.");
      if (typeof saveAs === "undefined") throw new Error("FileSaver library not loaded. Please refresh the page.");

      const noteType = safeTrim($("noteType")?.value);
      const title = safeTrim($("title")?.value);
      const topic = safeTrim($("topic")?.value);
      const authorLastName = safeTrim($("authorLastName")?.value);
      const authorFirstName = safeTrim($("authorFirstName")?.value);

      // Hidden phone is source of truth
      const authorPhone = safeTrim($("authorPhone")?.value);
      const authorPhoneSafe = naIfBlank(authorPhone);

      const authorEmail = safeTrim($("authorEmail")?.value); // optional future

      const analysis = safeTrim($("analysis")?.value);
      const keyTakeaways = safeTrim($("keyTakeaways")?.value);
      const content = safeTrim($("content")?.value);
      const cordobaView = safeTrim($("cordobaView")?.value);
      const imageFiles = $("imageUpload")?.files || [];

      const ticker = safeTrim($("ticker")?.value);
      const valuationSummary = safeTrim($("valuationSummary")?.value);
      const keyAssumptions = safeTrim($("keyAssumptions")?.value);
      const scenarioNotes = safeTrim($("scenarioNotes")?.value);
      const modelFiles = $("modelFiles")?.files || null;
      const modelLink = safeTrim($("modelLink")?.value);

      const targetPrice = safeTrim($("targetPrice")?.value);
      const crgRating = safeTrim($("crgRating")?.value);

      const now = new Date();
      const dateTimeString = formatDateTime(now);

      const coAuthors = [];
      qa(".coauthor-entry").forEach(entry => {
        const lastName = safeTrim(q(".coauthor-lastname", entry)?.value);
        const firstName = safeTrim(q(".coauthor-firstname", entry)?.value);
        const phone = safeTrim(q(".coauthor-phone", entry)?.value); // hidden combined
        if (lastName || firstName) coAuthors.push({ lastName, firstName, phone: naIfBlank(phone) });
      });

      const doc = await createDocument({
        noteType, title, topic,
        authorLastName, authorFirstName, authorPhone,
        authorPhoneSafe,
        authorEmail,
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
        `${(title || "research_note").replace(/[^a-z0-9]/gi, "_").toLowerCase()}_${(noteType || "note").replace(/\s+/g, "_").toLowerCase()}.docx`;

      saveAs(blob, fileName);

      if (messageDiv) {
        messageDiv.className = "message success";
        messageDiv.textContent = `✓ Document "${fileName}" generated successfully.`;
        messageDiv.style.display = "block";
      }

      // persist latest draft status
      saveDraftDebounced();
    } catch (error) {
      console.error("Error generating document:", error);
      if (messageDiv) {
        messageDiv.className = "message error";
        messageDiv.textContent = `✗ Error: ${error.message}`;
        messageDiv.style.display = "block";
      }
    } finally {
      if (button) {
        button.disabled = false;
        button.classList.remove("loading");
        button.textContent = "Generate Word Document";
      }
    }
  });

  // Initial paint
  updateAttachmentSummary();
  updateCompletionMeter();
  updateUpsideDisplay();
});
