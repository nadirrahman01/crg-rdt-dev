(() => {
  "use strict";

  const BRAND = {
    name: "Cordoba Research Group",
    short: "CRG",
    version: "RDT v4.0",
    tagline: "Values that bind",
    previewLogoPath: "./assets/cordoba-wordmark.svg",
    logoPath: "assets/cordoba-logo.png",
    colors: {
      gold: "B38633",
      navy: "102B46",
      navyMid: "234A72",
      ink: "162436",
      muted: "607086",
      border: "D7DEE8",
      cream: "F6F3ED",
      white: "FFFFFF",
      green: "245E48",
      red: "9E3737"
    },
    fonts: {
      heading: "Times New Roman",
      body: "Arial"
    },
    disclaimers: {
      internal:
        "Internal use only. Outputs are draft research documentation generated from user inputs and third-party market data. Verify all figures, tickers, assumptions, ratings and model outputs before circulation."
    }
  };

  const DRAFT_KEY = "crg_rdt_institutional_v40";
  const DRAFT_FIELDS = [
    "noteType",
    "noteStage",
    "distributionClass",
    "noteHorizon",
    "title",
    "topic",
    "thesisHeadline",
    "authorLastName",
    "authorFirstName",
    "authorPhone",
    "authorPhoneCountry",
    "authorPhoneNational",
    "keyTakeaways",
    "analysis",
    "variantPerception",
    "catalysts",
    "keyRisks",
    "content",
    "cordobaView",
    "ticker",
    "crgRating",
    "targetPrice",
    "valuationSummary",
    "keyAssumptions",
    "scenarioNotes",
    "projectedFinancialsTitle",
    "projectedFinancialsTable",
    "modelLink",
    "customDisclaimer"
  ];

  const state = {
    priceChart: null,
    priceChartImageBytes: null,
    priceChartDataUrl: "",
    equityStats: {
      currentPrice: null,
      realisedVolAnn: null,
      rangeReturn: null,
      startPrice: null
    }
  };

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const safeTrim = (value) => (value ?? "").toString().trim();
  const digitsOnly = (value) => (value || "").toString().replace(/\D/g, "");
  const safeNum = (value) => {
    const n = Number(String(value ?? "").replace(/,/g, ""));
    return Number.isFinite(n) ? n : null;
  };
  const pct = (value) => (Number.isFinite(value) ? `${(value * 100).toFixed(1)}%` : "—");
  const naIfBlank = (value) => (safeTrim(value) ? safeTrim(value) : "N/A");

  function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  function escapeHtml(value) {
    return (value || "")
      .toString()
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
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
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    let hours = date.getHours();
    const minutes = String(date.getMinutes()).padStart(2, "0");
    const ampm = hours >= 12 ? "PM" : "AM";
    hours = hours % 12 || 12;
    return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()} ${hours}:${minutes} ${ampm}`;
  }

  function formatDateLong(date) {
    const months = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
  }

  function formatDateShortISO(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
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
    const value = (noteTypeRaw || "").toLowerCase();
    if (value.includes("equity")) return "tommaso@cordobarg.com";
    if (value.includes("macro") || value.includes("market")) return "tim@cordobarg.com";
    if (value.includes("fixed")) return "tim@cordobarg.com";
    if (value.includes("commodity")) return "uhayd@cordobarg.com";
    return "";
  }

  function formatNationalLoose(rawDigits) {
    const digits = digitsOnly(rawDigits);
    if (!digits) return "";
    const parts = [digits.slice(0, 4), digits.slice(4, 7), digits.slice(7, 10), digits.slice(10)];
    return parts.filter(Boolean).join(" ");
  }

  function buildInternationalHyphen(countryCode, nationalDigits) {
    const cc = digitsOnly(countryCode);
    const nn = digitsOnly(nationalDigits);
    if (!cc && !nn) return "";
    if (cc && !nn) return `${cc}-`;
    if (!cc && nn) return nn;
    return `${cc}-${nn}`;
  }

  function toParagraphs(text) {
    const raw = safeTrim(text);
    if (!raw) return [];
    return raw.split(/\n\s*\n/g).map((block) => block.trim()).filter(Boolean);
  }

  function toBulletItems(text) {
    const raw = safeTrim(text);
    if (!raw) return [];
    return raw
      .split("\n")
      .map((line) => line.replace(/^[-*•]\s*/, "").trim())
      .filter(Boolean);
  }

  function toSimpleRows(text) {
    return safeTrim(text)
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);
  }

  function previewParagraphsHtml(text, fallback = "") {
    const blocks = toParagraphs(text);
    if (!blocks.length) {
      return `<div class="doc-paragraph doc-empty">${escapeHtml(fallback)}</div>`;
    }
    return blocks.map((block) => `<div class="doc-paragraph">${escapeHtml(block)}</div>`).join("");
  }

  function previewBulletsHtml(text, fallback = "") {
    const items = toBulletItems(text);
    if (!items.length) {
      return `<div class="doc-paragraph doc-empty">${escapeHtml(fallback)}</div>`;
    }
    return `<ul class="doc-bullets">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
  }

  function previewFileListHtml(items, fallback = "") {
    if (!items.length) {
      return `<div class="doc-paragraph doc-empty">${escapeHtml(fallback)}</div>`;
    }
    return `<div class="doc-file-list">${items.map((item) => `<div class="doc-file-item">${escapeHtml(item)}</div>`).join("")}</div>`;
  }

  function detectTableDelimiter(line) {
    if (line.includes("\t")) return "\t";
    if (line.includes("|")) return "|";
    return ",";
  }

  function splitTableRow(line, delimiter) {
    return line.split(delimiter).map((cell) => cell.trim());
  }

  function parseStructuredTable(text) {
    const rows = toSimpleRows(text);
    if (!rows.length) return [];
    const delimiter = detectTableDelimiter(rows[0]);
    const parsed = rows
      .map((row) => splitTableRow(row, delimiter))
      .filter((row) => row.some((cell) => safeTrim(cell)));
    const maxCols = Math.max(0, ...parsed.map((row) => row.length));
    return parsed.map((row) => {
      const next = row.slice();
      while (next.length < maxCols) next.push("");
      return next;
    });
  }

  function previewTableHtml(text, fallback = "") {
    const rows = parseStructuredTable(text);
    if (!rows.length) {
      return `<div class="doc-paragraph doc-empty">${escapeHtml(fallback)}</div>`;
    }

    const [header, ...body] = rows;
    return `
      <div class="doc-table-wrap">
        <table class="doc-table">
          <thead>
            <tr>${header.map((cell) => `<th>${escapeHtml(cell)}</th>`).join("")}</tr>
          </thead>
          <tbody>
            ${body.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function brandLogoHtml() {
    return `
      <div class="doc-brand-lockup">
        <img
          src="${escapeHtml(BRAND.previewLogoPath)}"
          alt="${escapeHtml(BRAND.name)}"
          onerror="this.style.display='none'; this.nextElementSibling.style.display='block';"
        />
        <div class="doc-brand-fallback">${escapeHtml(BRAND.name)}</div>
      </div>
    `;
  }

  function ratingToDisplay(value) {
    const rating = safeTrim(value);
    if (!rating) return "—";
    if (rating.toLowerCase() === "hold") return "Neutral";
    return rating;
  }

  function noteTypeCode(noteType) {
    const mapping = {
      "Macro Research": "MR",
      "Fixed Income Research": "FI",
      "Commodity Insights": "CI",
      "Equity Research": "ER",
      "General Note": "GN"
    };
    return mapping[noteType] || "RN";
  }

  function stageCode(noteStage) {
    const mapping = {
      "Initiation": "INIT",
      "Update": "UPD",
      "Flash Note": "FLASH",
      "Results Review": "RESULTS",
      "Preview": "PREV",
      "Thematic": "THEM",
      "Sector Note": "SECTOR"
    };
    return mapping[noteStage] || "DRAFT";
  }

  function tokenFromText(text) {
    const clean = safeTrim(text)
      .replace(/[^A-Za-z0-9\s]/g, " ")
      .split(/\s+/)
      .filter(Boolean);
    if (!clean.length) return "NOTE";
    if (clean.length === 1) return clean[0].slice(0, 8).toUpperCase();
    return clean.slice(0, 3).map((part) => part[0]).join("").toUpperCase();
  }

  function buildNoteReference(data) {
    const suffix = safeTrim(data.ticker) ? safeTrim(data.ticker).replace(/[^A-Za-z0-9]/g, "").toUpperCase() : tokenFromText(data.title);
    return `CRG-${noteTypeCode(data.noteType)}-${stageCode(data.noteStage)}-${formatDateShortISO(new Date()).replaceAll("-", "")}-${suffix || "NOTE"}`;
  }

  async function fetchAssetBytes(path) {
    try {
      const res = await fetch(path, { cache: "no-store" });
      if (!res.ok) return null;
      const buf = await res.arrayBuffer();
      return new Uint8Array(buf);
    } catch {
      return null;
    }
  }

  const authorPhoneCountryEl = document.getElementById("authorPhoneCountry");
  const authorPhoneNationalEl = document.getElementById("authorPhoneNational");
  const authorPhoneHiddenEl = document.getElementById("authorPhone");

  function syncPrimaryPhone() {
    if (!authorPhoneHiddenEl) return;
    const cc = authorPhoneCountryEl ? authorPhoneCountryEl.value : "";
    const nn = authorPhoneNationalEl ? authorPhoneNationalEl.value : "";
    authorPhoneHiddenEl.value = buildInternationalHyphen(cc, nn);
  }

  function formatPrimaryVisible() {
    if (!authorPhoneNationalEl) return;
    const caret = authorPhoneNationalEl.selectionStart || 0;
    const beforeLength = authorPhoneNationalEl.value.length;
    authorPhoneNationalEl.value = formatNationalLoose(authorPhoneNationalEl.value);
    const afterLength = authorPhoneNationalEl.value.length;
    const next = Math.max(0, caret + (afterLength - beforeLength));
    authorPhoneNationalEl.setSelectionRange(next, next);
    syncPrimaryPhone();
  }

  const countryOptionsHtml = `
    <option value="44" selected>+44</option>
    <option value="1">+1</option>
    <option value="353">+353</option>
    <option value="33">+33</option>
    <option value="49">+49</option>
    <option value="31">+31</option>
    <option value="34">+34</option>
    <option value="39">+39</option>
    <option value="971">+971</option>
    <option value="966">+966</option>
    <option value="92">+92</option>
    <option value="91">+91</option>
    <option value="27">+27</option>
    <option value="">Other</option>
  `;

  let coAuthorCount = 0;
  const addCoAuthorBtn = document.getElementById("addCoAuthor");
  const coAuthorsList = document.getElementById("coAuthorsList");

  function getCoauthorData() {
    return $$(".coauthor-entry").map((entry) => ({
      lastName: safeTrim($(".coauthor-lastname", entry)?.value),
      firstName: safeTrim($(".coauthor-firstname", entry)?.value),
      cc: safeTrim($(".coauthor-country", entry)?.value),
      local: safeTrim($(".coauthor-phone-local", entry)?.value),
      phone: safeTrim($(".coauthor-phone", entry)?.value)
    })).filter((author) => author.lastName || author.firstName || author.phone || author.local);
  }

  function buildAnalystRoster() {
    const roster = [];
    const firstName = safeTrim($("#authorFirstName")?.value);
    const lastName = safeTrim($("#authorLastName")?.value);
    const primaryName = [firstName, lastName].filter(Boolean).join(" ").trim();
    const primaryPhone = safeTrim($("#authorPhone")?.value);
    if (primaryName) {
      roster.push({
        name: primaryName,
        phone: primaryPhone
      });
    }

    getCoauthorData().forEach((author) => {
      const name = [author.firstName, author.lastName].filter(Boolean).join(" ").trim();
      if (!name) return;
      roster.push({
        name,
        phone: author.phone || buildInternationalHyphen(author.cc, author.local)
      });
    });

    return roster;
  }

  function wireCoauthorPhone(node) {
    const ccEl = $(".coauthor-country", node);
    const nationalEl = $(".coauthor-phone-local", node);
    const hiddenEl = $(".coauthor-phone", node);

    function syncHidden() {
      if (!hiddenEl) return;
      hiddenEl.value = buildInternationalHyphen(ccEl?.value, nationalEl?.value);
    }

    function formatVisible() {
      if (!nationalEl) return;
      const caret = nationalEl.selectionStart || 0;
      const beforeLength = nationalEl.value.length;
      nationalEl.value = formatNationalLoose(nationalEl.value);
      const afterLength = nationalEl.value.length;
      const next = Math.max(0, caret + (afterLength - beforeLength));
      nationalEl.setSelectionRange(next, next);
      syncHidden();
    }

    if (nationalEl) {
      nationalEl.addEventListener("input", () => {
        formatVisible();
        scheduleDraftSave();
        refreshUI();
      });
      nationalEl.addEventListener("blur", () => {
        syncHidden();
        scheduleDraftSave();
      });
    }

    if (ccEl) {
      ccEl.addEventListener("change", () => {
        syncHidden();
        scheduleDraftSave();
        refreshUI();
      });
    }

    syncHidden();
  }

  function createCoauthorNode() {
    coAuthorCount += 1;

    const node = document.createElement("div");
    node.className = "coauthor-entry";
    node.id = `coauthor-${coAuthorCount}`;
    node.innerHTML = `
      <div>
        <label>Last Name</label>
        <input type="text" class="coauthor-lastname" placeholder="Surname" autocomplete="family-name" />
      </div>
      <div>
        <label>First Name</label>
        <input type="text" class="coauthor-firstname" placeholder="Given name" autocomplete="given-name" />
      </div>
      <div class="phone-row--compact">
        <div>
          <label>Code</label>
          <select class="phone-country coauthor-country" aria-label="Country code">
            ${countryOptionsHtml}
          </select>
        </div>
        <div>
          <label>Phone</label>
          <input type="text" class="phone-number coauthor-phone-local" placeholder="Phone number" inputmode="numeric" />
        </div>
      </div>
      <input type="text" class="coauthor-phone" hidden />
      <button type="button" class="remove-coauthor" data-remove-id="${coAuthorCount}">Remove</button>
    `;

    wireCoauthorPhone(node);

    ["input", "change"].forEach((eventName) => {
      node.addEventListener(eventName, () => {
        scheduleDraftSave();
        refreshUI();
      }, { passive: true });
    });

    return node;
  }

  const modelFilesEl = document.getElementById("modelFiles");
  const attachmentSummaryHeadEl = document.getElementById("attachmentSummaryHead");
  const attachmentSummaryListEl = document.getElementById("attachmentSummaryList");

  function updateAttachmentSummary() {
    if (!modelFilesEl || !attachmentSummaryHeadEl || !attachmentSummaryListEl) return;
    const files = Array.from(modelFilesEl.files || []);
    if (!files.length) {
      attachmentSummaryHeadEl.textContent = "No files selected";
      attachmentSummaryListEl.innerHTML = "";
      attachmentSummaryListEl.style.display = "none";
      return;
    }
    attachmentSummaryHeadEl.textContent = `${files.length} file${files.length === 1 ? "" : "s"} selected`;
    attachmentSummaryListEl.innerHTML = files.map((file) => `<div class="attachment-file">${escapeHtml(file.name)}</div>`).join("");
    attachmentSummaryListEl.style.display = "block";
  }

  function isEquityMode() {
    return safeTrim($("#noteType")?.value) === "Equity Research";
  }

  function toggleEquityPanels() {
    $$(".equity-only").forEach((section) => {
      section.style.display = isEquityMode() ? "block" : "none";
    });
  }

  function isFilled(el) {
    if (!el) return false;
    if (el.type === "file") return !!(el.files && el.files.length);
    return safeTrim(el.value).length > 0;
  }

  const completionTextEl = document.getElementById("completionText");
  const completionBarEl = document.getElementById("completionBar");
  const readinessStatusEl = document.getElementById("readinessStatus");
  const qualityChecklistEl = document.getElementById("qualityChecklist");
  const qualityGateSummaryEl = document.getElementById("qualityGateSummary");

  const baseCoreIds = [
    "noteType",
    "noteStage",
    "distributionClass",
    "title",
    "topic",
    "thesisHeadline",
    "authorLastName",
    "authorFirstName",
    "keyTakeaways",
    "analysis",
    "catalysts",
    "keyRisks"
  ];

  const equityCoreIds = ["ticker", "crgRating", "targetPrice", "valuationSummary"];

  function updateCompletionMeter() {
    const ids = isEquityMode() ? baseCoreIds.concat(equityCoreIds) : baseCoreIds;
    const done = ids.reduce((count, id) => count + (isFilled(document.getElementById(id)) ? 1 : 0), 0);
    const total = ids.length;
    const completion = total ? Math.round((done / total) * 100) : 0;

    if (completionTextEl) completionTextEl.textContent = `${done} / ${total} publish-core`;
    if (completionBarEl) completionBarEl.style.width = `${completion}%`;

    return { done, total, completion };
  }

  function buildQualityGates() {
    const noteType = $("#noteType");
    const noteStage = $("#noteStage");
    const distributionClass = $("#distributionClass");
    const noteHorizon = $("#noteHorizon");
    const title = $("#title");
    const topic = $("#topic");
    const thesisHeadline = $("#thesisHeadline");
    const authorLastName = $("#authorLastName");
    const authorFirstName = $("#authorFirstName");
    const keyTakeaways = $("#keyTakeaways");
    const analysis = $("#analysis");
    const catalysts = $("#catalysts");
    const keyRisks = $("#keyRisks");
    const cordobaView = $("#cordobaView");
    const variantPerception = $("#variantPerception");
    const content = $("#content");
    const ticker = $("#ticker");
    const crgRating = $("#crgRating");
    const targetPrice = $("#targetPrice");
    const valuationSummary = $("#valuationSummary");
    const keyAssumptions = $("#keyAssumptions");
    const projectedFinancialsTable = $("#projectedFinancialsTable");
    const imageUpload = $("#imageUpload");

    const gates = [
      {
        label: "Document control complete",
        pass: [noteType, noteStage, distributionClass, title, topic].every(isFilled)
      },
      {
        label: "Horizon and routing locked",
        pass: [distributionClass, noteHorizon].every(isFilled)
      },
      {
        label: "Lead analyst identified",
        pass: [authorLastName, authorFirstName].every(isFilled)
      },
      {
        label: "House call is explicit",
        pass: [thesisHeadline, keyTakeaways].every(isFilled)
      },
      {
        label: "Core thesis drafted",
        pass: isFilled(analysis)
      },
      {
        label: "Catalysts and risks covered",
        pass: [catalysts, keyRisks].every(isFilled)
      },
      {
        label: "Cordoba differentiation captured",
        pass: isFilled(cordobaView) || isFilled(variantPerception)
      }
    ];

    if (isEquityMode()) {
      gates.push(
        {
          label: "Recommendation and target set",
          pass: [ticker, crgRating, targetPrice].every(isFilled)
        },
        {
          label: "Valuation support present",
          pass: isFilled(valuationSummary) || isFilled(keyAssumptions) || isFilled(projectedFinancialsTable)
        }
      );
    } else {
      gates.push({
        label: "Support material or figures added",
        pass: isFilled(content) || isFilled(imageUpload)
      });
    }

    return gates;
  }

  function renderQualityGates() {
    const gates = buildQualityGates();
    const passed = gates.filter((gate) => gate.pass).length;
    const total = gates.length;
    const ratio = total ? passed / total : 0;
    const completion = updateCompletionMeter().completion;

    if (qualityChecklistEl) {
      qualityChecklistEl.innerHTML = gates
        .map((gate) => `<li class="${gate.pass ? "is-pass" : "is-fail"}">${escapeHtml(gate.label)}</li>`)
        .join("");
    }

    if (qualityGateSummaryEl) {
      qualityGateSummaryEl.textContent = `${passed} / ${total} gates closed`;
    }

    let label = "Draft";
    if (ratio >= 0.85 && completion >= 85) label = "Publish Candidate";
    else if (ratio >= 0.7 && completion >= 70) label = "Desk Review";
    else if (ratio >= 0.5 && completion >= 55) label = "Analyst Ready";

    if (readinessStatusEl) readinessStatusEl.textContent = label;
    return { passed, total, label };
  }

  function setDraftStatus(text) {
    const el = document.getElementById("draftStatus");
    if (el) el.textContent = text || "";
    const mirror = document.getElementById("draftStatusMirror");
    if (mirror) mirror.textContent = text || "—";
  }

  function snapshotDraft() {
    const draft = {};
    DRAFT_FIELDS.forEach((id) => {
      const el = document.getElementById(id);
      if (el) draft[id] = el.value ?? "";
    });
    draft.__coAuthors = getCoauthorData();
    draft.__chartRange = $("#chartRange")?.value || "";
    draft.__equityStats = state.equityStats || null;
    draft.__savedAt = new Date().toISOString();
    return draft;
  }

  function saveDraftNow() {
    try {
      localStorage.setItem(DRAFT_KEY, JSON.stringify(snapshotDraft()));
      setDraftStatus("Saved");
    } catch {
      setDraftStatus("Local save failed");
    }
  }

  let draftSaveTimer = null;
  function scheduleDraftSave() {
    setDraftStatus("Saving…");
    clearTimeout(draftSaveTimer);
    draftSaveTimer = setTimeout(() => {
      saveDraftNow();
      syncWorkspacePreview();
    }, 300);
  }

  function loadDraft() {
    try {
      const raw = localStorage.getItem(DRAFT_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  }

  function clearDraft() {
    try {
      localStorage.removeItem(DRAFT_KEY);
    } catch {
      // noop
    }
  }

  function applyDraft(draft) {
    if (!draft) return;

    DRAFT_FIELDS.forEach((id) => {
      const el = document.getElementById(id);
      if (el && typeof draft[id] === "string") el.value = draft[id];
    });

    if (Array.isArray(draft.__coAuthors) && coAuthorsList) {
      coAuthorsList.innerHTML = "";
      draft.__coAuthors.forEach((author) => {
        const node = createCoauthorNode();
        $(".coauthor-lastname", node).value = author.lastName || "";
        $(".coauthor-firstname", node).value = author.firstName || "";
        $(".coauthor-country", node).value = author.cc || "44";
        $(".coauthor-phone-local", node).value = author.local ? formatNationalLoose(author.local) : "";
        $(".coauthor-phone", node).value = author.phone || "";
        wireCoauthorPhone(node);
        coAuthorsList.appendChild(node);
      });
    }

    if (draft.__chartRange && $("#chartRange")) {
      $("#chartRange").value = draft.__chartRange;
    }

    if (draft.__equityStats) {
      state.equityStats = draft.__equityStats;
    }

    syncPrimaryPhone();
    paintEquityStats();
  }

  function buildPayloadFromForm() {
    const noteType = safeTrim($("#noteType")?.value || "Research Note");
    const noteStage = safeTrim($("#noteStage")?.value || "Draft");
    const distributionClass = safeTrim($("#distributionClass")?.value || "Unassigned");
    const noteHorizon = safeTrim($("#noteHorizon")?.value || "—");
    const title = safeTrim($("#title")?.value || "Untitled research note");
    const topic = safeTrim($("#topic")?.value || "—");
    const thesisHeadline = safeTrim($("#thesisHeadline")?.value || "");
    const authorFirstName = safeTrim($("#authorFirstName")?.value || "");
    const authorLastName = safeTrim($("#authorLastName")?.value || "");
    const authorPhoneSafe = naIfBlank(safeTrim($("#authorPhone")?.value || ""));
    const keyTakeaways = $("#keyTakeaways")?.value || "";
    const analysis = $("#analysis")?.value || "";
    const variantPerception = $("#variantPerception")?.value || "";
    const catalysts = $("#catalysts")?.value || "";
    const keyRisks = $("#keyRisks")?.value || "";
    const content = $("#content")?.value || "";
    const cordobaView = $("#cordobaView")?.value || "";
    const ticker = safeTrim($("#ticker")?.value || "");
    const crgRating = safeTrim($("#crgRating")?.value || "");
    const targetPrice = safeTrim($("#targetPrice")?.value || "");
    const valuationSummary = $("#valuationSummary")?.value || "";
    const keyAssumptions = $("#keyAssumptions")?.value || "";
    const scenarioNotes = $("#scenarioNotes")?.value || "";
    const projectedFinancialsTitle = safeTrim($("#projectedFinancialsTitle")?.value || "Projected Financials");
    const projectedFinancialsTable = $("#projectedFinancialsTable")?.value || "";
    const modelLink = safeTrim($("#modelLink")?.value || "");
    const customDisclaimer = $("#customDisclaimer")?.value || "";
    const imageFiles = $("#imageUpload")?.files || [];
    const modelFiles = $("#modelFiles")?.files || [];
    const noteReference = buildNoteReference({ noteType, noteStage, title, ticker });
    const deskCc = ccForNoteType(noteType);
    const analysts = buildAnalystRoster();

    return {
      noteType,
      noteStage,
      distributionClass,
      noteHorizon,
      title,
      topic,
      thesisHeadline,
      authorFirstName,
      authorLastName,
      authorPhoneSafe,
      keyTakeaways,
      analysis,
      variantPerception,
      catalysts,
      keyRisks,
      content,
      cordobaView,
      ticker,
      crgRating,
      targetPrice,
      valuationSummary,
      keyAssumptions,
      scenarioNotes,
      projectedFinancialsTitle,
      projectedFinancialsTable,
      modelLink,
      customDisclaimer,
      imageFiles,
      imageFileNames: Array.from(imageFiles).map((file) => file.name),
      modelFiles,
      modelFileNames: Array.from(modelFiles).map((file) => file.name),
      analysts,
      noteReference,
      deskCc,
      currentPriceText: $("#currentPrice")?.textContent?.trim() || "—",
      rangeReturnText: $("#rangeReturn")?.textContent?.trim() || "—",
      realisedVolText: $("#realisedVol")?.textContent?.trim() || "—",
      upsideText: $("#upsideToTarget")?.textContent?.trim() || "—",
      chartDataUrl: state.priceChartDataUrl,
      dateLabel: formatDateLong(new Date()),
      dateTimeString: formatDateTime(new Date())
    };
  }

  function previewChartHtml(data) {
    if (!data.chartDataUrl) {
      return `<div class="doc-preview-chart"><div class="doc-preview-chart-empty">Fetch a chart to include price action in the preview and export.</div></div>`;
    }

    return `
      <div class="doc-preview-chart">
        <img src="${data.chartDataUrl}" alt="${escapeHtml(data.ticker || "Price chart")}" />
      </div>
    `;
  }

  function previewSideCardHtml(data) {
    return `
      <aside class="doc-side-card">
        <div class="doc-side-block">
          <span class="doc-side-label">Analysts</span>
          <span class="doc-side-value">
            ${data.analysts.length ? data.analysts.map((analyst) => `${escapeHtml(analyst.name)}${analyst.phone ? ` · ${escapeHtml(analyst.phone)}` : ""}`).join("<br>") : "—"}
          </span>
        </div>
        <div class="doc-side-block">
          <span class="doc-side-label">Distribution</span>
          <span class="doc-side-value">${escapeHtml(data.distributionClass)}</span>
        </div>
        <div class="doc-side-block">
          <span class="doc-side-label">Note Reference</span>
          <span class="doc-side-value">${escapeHtml(data.noteReference)}</span>
        </div>
        <div class="doc-side-block">
          <span class="doc-side-label">Horizon</span>
          <span class="doc-side-value">${escapeHtml(data.noteHorizon)}</span>
        </div>
        <div class="doc-side-block">
          <span class="doc-side-label">Desk CC</span>
          <span class="doc-side-value">${escapeHtml(data.deskCc || "Desk only")}</span>
        </div>
        ${data.noteType === "Equity Research" ? `
          <div class="doc-side-block">
            <span class="doc-side-label">Recommendation</span>
            <span class="doc-side-value">${escapeHtml(ratingToDisplay(data.crgRating))}</span>
          </div>
          <div class="doc-side-block">
            <span class="doc-side-label">Current / Target</span>
            <span class="doc-side-value">${escapeHtml(data.currentPriceText)} / ${escapeHtml(data.targetPrice || "—")}</span>
          </div>
          <div class="doc-side-block">
            <span class="doc-side-label">Range Return / Upside</span>
            <span class="doc-side-value">${escapeHtml(data.rangeReturnText)} / ${escapeHtml(data.upsideText)}</span>
          </div>
          <div class="doc-side-block">
            <span class="doc-side-label">Price context</span>
            <div class="doc-side-value">${previewChartHtml(data)}</div>
          </div>
        ` : ""}
      </aside>
    `;
  }

  function previewHeroHtml(data) {
    return `
      <div class="doc-brand-bar">
        ${brandLogoHtml()}
        <div class="doc-date-block">
          <strong>${escapeHtml(data.distributionClass)}</strong>
          <span>${escapeHtml(data.dateLabel)}</span>
        </div>
      </div>

      <div class="doc-hero">
        <div class="doc-kicker">${escapeHtml(`${data.noteType} · ${data.noteStage}`)}</div>
        <div class="doc-hero-title">${escapeHtml(data.title)}</div>
        <p class="doc-hero-subtitle">${escapeHtml(data.thesisHeadline || data.topic)}</p>
        <div class="doc-pill-row">
          <span class="doc-pill">${escapeHtml(data.topic)}</span>
          <span class="doc-pill">${escapeHtml(data.noteHorizon)}</span>
          <span class="doc-pill">${escapeHtml(data.noteReference)}</span>
        </div>
      </div>
    `;
  }

  function previewEquityHtml(data) {
    const ratingDisplay = ratingToDisplay(data.crgRating);
    const titleForTable = safeTrim(data.projectedFinancialsTitle) || "Projected Financials";

    return `
      ${previewHeroHtml(data)}

      <div class="doc-summary-strip">
        <div class="doc-summary-card">
          <span>Ticker</span>
          <strong>${escapeHtml(data.ticker || "—")}</strong>
        </div>
        <div class="doc-summary-card">
          <span>Recommendation</span>
          <strong>${escapeHtml(ratingDisplay)}</strong>
        </div>
        <div class="doc-summary-card">
          <span>Current / Target</span>
          <strong>${escapeHtml(data.currentPriceText)} / ${escapeHtml(data.targetPrice || "—")}</strong>
        </div>
        <div class="doc-summary-card">
          <span>Upside</span>
          <strong>${escapeHtml(data.upsideText)}</strong>
        </div>
      </div>

      <div class="doc-body-grid">
        ${previewSideCardHtml(data)}

        <div>
          <section class="doc-section">
            <h3 class="doc-section-title">Key Takeaways</h3>
            ${previewBulletsHtml(data.keyTakeaways, "Key takeaways will appear here.")}
          </section>

          <section class="doc-section">
            <h3 class="doc-section-title">Investment Thesis</h3>
            <p class="doc-section-subtitle">Lead with the key call, then support it with evidence and why the market is mis-framing the setup.</p>
            ${previewParagraphsHtml(data.analysis, "Your investment thesis will appear here.")}
          </section>

          <div class="doc-callout-grid">
            <article class="doc-callout-card">
              <h3>Variant Perception</h3>
              ${previewParagraphsHtml(data.variantPerception, "What Cordoba sees differently will appear here.")}
            </article>
            <article class="doc-callout-card">
              <h3>Catalysts</h3>
              ${previewBulletsHtml(data.catalysts, "Catalysts will appear here.")}
            </article>
            <article class="doc-callout-card">
              <h3>Key Risks</h3>
              ${previewBulletsHtml(data.keyRisks, "Risks will appear here.")}
            </article>
            <article class="doc-callout-card">
              <h3>Valuation</h3>
              ${previewParagraphsHtml(data.valuationSummary, "Valuation commentary will appear here.")}
            </article>
          </div>

          <section class="doc-section">
            <h3 class="doc-section-title">Key Assumptions</h3>
            ${previewBulletsHtml(data.keyAssumptions, "Key assumptions will appear here.")}
          </section>

          <section class="doc-section">
            <h3 class="doc-section-title">${escapeHtml(titleForTable)}</h3>
            ${previewTableHtml(data.projectedFinancialsTable, "Projected financials will appear here. Use commas, tabs or pipes.")}
          </section>

          <section class="doc-section">
            <h3 class="doc-section-title">Scenario / Sensitivity Notes</h3>
            ${previewParagraphsHtml(data.scenarioNotes, "Scenario notes will appear here.")}
          </section>

          <section class="doc-section">
            <h3 class="doc-section-title">Additional Detail</h3>
            ${previewParagraphsHtml(data.content, "Additional supporting detail will appear here.")}
          </section>

          <section class="doc-section">
            <h3 class="doc-section-title">The Cordoba View</h3>
            ${previewParagraphsHtml(data.cordobaView, "The Cordoba house view will appear here.")}
          </section>

          <section class="doc-section">
            <h3 class="doc-section-title">Model & Supporting Files</h3>
            ${previewFileListHtml(data.modelFileNames, "Model file names will appear here once uploaded.")}
          </section>
        </div>
      </div>
    `;
  }

  function previewGeneralHtml(data) {
    return `
      ${previewHeroHtml(data)}

      <div class="doc-summary-strip">
        <div class="doc-summary-card">
          <span>Publication Type</span>
          <strong>${escapeHtml(data.noteStage)}</strong>
        </div>
        <div class="doc-summary-card">
          <span>Distribution</span>
          <strong>${escapeHtml(data.distributionClass)}</strong>
        </div>
        <div class="doc-summary-card">
          <span>Horizon</span>
          <strong>${escapeHtml(data.noteHorizon)}</strong>
        </div>
        <div class="doc-summary-card">
          <span>Lead Analyst</span>
          <strong>${escapeHtml(data.analysts[0]?.name || "—")}</strong>
        </div>
      </div>

      <div class="doc-body-grid">
        ${previewSideCardHtml(data)}

        <div>
          <section class="doc-section">
            <h3 class="doc-section-title">Key Takeaways</h3>
            ${previewBulletsHtml(data.keyTakeaways, "Key takeaways will appear here.")}
          </section>

          <section class="doc-section">
            <h3 class="doc-section-title">Core Analysis</h3>
            ${previewParagraphsHtml(data.analysis, "Your analysis will appear here.")}
          </section>

          <div class="doc-callout-grid">
            <article class="doc-callout-card">
              <h3>Variant Perception</h3>
              ${previewParagraphsHtml(data.variantPerception, "Variant perception will appear here.")}
            </article>
            <article class="doc-callout-card">
              <h3>Catalysts</h3>
              ${previewBulletsHtml(data.catalysts, "Catalysts will appear here.")}
            </article>
            <article class="doc-callout-card">
              <h3>Key Risks</h3>
              ${previewBulletsHtml(data.keyRisks, "Risks will appear here.")}
            </article>
            <article class="doc-callout-card">
              <h3>The Cordoba View</h3>
              ${previewParagraphsHtml(data.cordobaView, "The Cordoba house view will appear here.")}
            </article>
          </div>

          <section class="doc-section">
            <h3 class="doc-section-title">Additional Detail</h3>
            ${previewParagraphsHtml(data.content, "Additional supporting detail will appear here.")}
          </section>

          <section class="doc-section">
            <h3 class="doc-section-title">Figures & Attachments</h3>
            ${previewFileListHtml(data.imageFileNames, "Uploaded figure names will appear here.")}
          </section>
        </div>
      </div>
    `;
  }

  function syncWorkspacePreview() {
    const data = buildPayloadFromForm();
    const previewBody = document.getElementById("docPreviewBody");
    if (!previewBody) return;

    setText("docHeaderLeft", `CRG | ${data.noteType} | ${formatDateShortISO(new Date())}`);
    setText("docHeaderRight", data.distributionClass);
    setText("previewMode", data.noteType);
    setText("previewPillType", data.noteType);
    setText("previewPillStage", data.noteStage);
    setText("distributionMirror", data.distributionClass);
    setText("deskRouting", data.deskCc || "research@cordobarg.com");
    setText("noteReference", data.noteReference);

    const disclaimer = [BRAND.disclaimers.internal, safeTrim(data.customDisclaimer)].filter(Boolean).join(" ");
    setText("previewDisclaimer", disclaimer);

    previewBody.innerHTML = data.noteType === "Equity Research"
      ? previewEquityHtml(data)
      : previewGeneralHtml(data);
  }

  function refreshUI() {
    toggleEquityPanels();
    updateAttachmentSummary();
    updateCompletionMeter();
    renderQualityGates();
    syncWorkspacePreview();
  }

  function clearChartUI() {
    setText("currentPrice", "—");
    setText("rangeReturn", "—");
    setText("realisedVol", "—");
    setText("upsideToTarget", "—");
    setText("chartStatus", "");

    if (state.priceChart) {
      try {
        state.priceChart.destroy();
      } catch {
        // noop
      }
      state.priceChart = null;
    }

    state.priceChartImageBytes = null;
    state.priceChartDataUrl = "";
    state.equityStats = {
      currentPrice: null,
      realisedVolAnn: null,
      rangeReturn: null,
      startPrice: null
    };
  }

  function stooqSymbolFromTicker(ticker) {
    const value = safeTrim(ticker);
    if (!value) return null;
    if (value.includes(".")) return value.toLowerCase();
    return `${value.toLowerCase()}.us`;
  }

  function computeStartDate(range) {
    const date = new Date();
    if (range === "6mo") date.setMonth(date.getMonth() - 6);
    else if (range === "1y") date.setFullYear(date.getFullYear() - 1);
    else if (range === "2y") date.setFullYear(date.getFullYear() - 2);
    else if (range === "5y") date.setFullYear(date.getFullYear() - 5);
    else date.setFullYear(date.getFullYear() - 1);
    return date;
  }

  function extractStooqCSV(text) {
    const lines = (text || "").split("\n").map((line) => line.trim()).filter(Boolean);
    const headerIndex = lines.findIndex((line) => line.toLowerCase().startsWith("date,open,high,low,close,volume"));
    if (headerIndex === -1) return null;
    return lines.slice(headerIndex).join("\n");
  }

  async function fetchStooqDaily(symbol) {
    const stooqUrl = `http://stooq.com/q/d/l/?s=${encodeURIComponent(symbol)}&i=d`;
    const proxyUrl = `https://r.jina.ai/${stooqUrl}`;
    const res = await fetch(proxyUrl, { cache: "no-store" });
    if (!res.ok) throw new Error("Could not fetch price data.");

    const rawText = await res.text();
    const csvText = extractStooqCSV(rawText) || rawText;
    const lines = csvText.trim().split("\n");
    if (lines.length < 10) throw new Error("Not enough data returned for that ticker.");

    const rows = lines.slice(1).map((line) => line.split(","));
    const series = rows
      .map((row) => ({ date: row[0], close: Number(row[4]) }))
      .filter((point) => point.date && Number.isFinite(point.close));

    if (!series.length) throw new Error("No usable price data returned.");
    return series;
  }

  function computeDailyReturns(closes) {
    const output = [];
    for (let i = 1; i < closes.length; i += 1) {
      const prev = closes[i - 1];
      const current = closes[i];
      if (prev > 0 && Number.isFinite(prev) && Number.isFinite(current)) {
        output.push((current / prev) - 1);
      }
    }
    return output;
  }

  function stddev(values) {
    if (!values.length) return null;
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const variance = values.reduce((sum, value) => sum + ((value - mean) ** 2), 0) / (values.length - 1 || 1);
    return Math.sqrt(variance);
  }

  function computeUpsideToTarget(currentPrice, targetPrice) {
    if (!Number.isFinite(currentPrice) || !Number.isFinite(targetPrice) || currentPrice <= 0) return null;
    return (targetPrice / currentPrice) - 1;
  }

  function renderChart({ labels, values, title }) {
    const canvas = document.getElementById("priceChart");
    if (!canvas || typeof Chart === "undefined") return;

    if (state.priceChart) {
      try {
        state.priceChart.destroy();
      } catch {
        // noop
      }
    }

    state.priceChart = new Chart(canvas, {
      type: "line",
      data: {
        labels,
        datasets: [{
          label: title,
          data: values,
          pointRadius: 0,
          borderWidth: 2.25,
          borderColor: "#123455",
          tension: 0.14,
          fill: true,
          backgroundColor: "rgba(179, 134, 51, 0.12)"
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
            ticks: { maxTicksLimit: 6, color: "#607086" },
            grid: { color: "rgba(16, 43, 70, 0.06)" }
          },
          y: {
            ticks: { maxTicksLimit: 6, color: "#607086" },
            grid: { color: "rgba(16, 43, 70, 0.06)" }
          }
        }
      }
    });
  }

  function canvasToPngBytes(canvas) {
    const dataUrl = canvas.toDataURL("image/png");
    const base64 = dataUrl.split(",")[1];
    return Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
  }

  function paintEquityStats() {
    const stats = state.equityStats || {};
    setText("currentPrice", Number.isFinite(stats.currentPrice) ? stats.currentPrice.toFixed(2) : "—");
    setText("rangeReturn", Number.isFinite(stats.rangeReturn) ? pct(stats.rangeReturn) : "—");
    setText("realisedVol", Number.isFinite(stats.realisedVolAnn) ? pct(stats.realisedVolAnn) : "—");
    const targetPrice = safeNum($("#targetPrice")?.value);
    const upside = computeUpsideToTarget(stats.currentPrice, targetPrice);
    setText("upsideToTarget", upside === null ? "—" : pct(upside));
  }

  async function buildPriceChart() {
    try {
      const tickerValue = safeTrim($("#ticker")?.value || "");
      if (!tickerValue) throw new Error("Enter a ticker first.");

      const symbol = stooqSymbolFromTicker(tickerValue);
      if (!symbol) throw new Error("Invalid ticker.");

      const range = $("#chartRange")?.value || "1y";
      setText("chartStatus", "Fetching price data…");

      const data = await fetchStooqDaily(symbol);
      const startDate = computeStartDate(range);
      const filtered = data.filter((point) => new Date(point.date) >= startDate);
      if (filtered.length < 10) throw new Error("Not enough data for the selected range.");

      const labels = filtered.map((point) => point.date);
      const values = filtered.map((point) => point.close);

      renderChart({
        labels,
        values,
        title: `${tickerValue.toUpperCase()} Close`
      });

      const canvas = document.getElementById("priceChart");
      await new Promise((resolve) => setTimeout(resolve, 120));
      state.priceChartImageBytes = canvas ? canvasToPngBytes(canvas) : null;
      state.priceChartDataUrl = canvas ? canvas.toDataURL("image/png") : "";

      const currentPrice = values[values.length - 1];
      const startPrice = values[0];
      const rangeReturn = currentPrice && startPrice ? (currentPrice / startPrice) - 1 : null;
      const realisedVolAnn = (() => {
        const dailyVol = stddev(computeDailyReturns(values));
        return dailyVol === null ? null : dailyVol * Math.sqrt(252);
      })();

      state.equityStats = {
        currentPrice,
        realisedVolAnn,
        rangeReturn,
        startPrice
      };

      paintEquityStats();
      setText("chartStatus", `Chart ready (${range.toUpperCase()})`);
      scheduleDraftSave();
      syncWorkspacePreview();
    } catch (error) {
      clearChartUI();
      paintEquityStats();
      setText("chartStatus", `Chart unavailable: ${error.message}`);
    } finally {
      updateCompletionMeter();
      renderQualityGates();
    }
  }

  function bindSectionJumpButtons() {
    $$(".jump-btn").forEach((button) => {
      button.addEventListener("click", () => {
        const targetId = button.getAttribute("data-scroll-target");
        const target = targetId ? document.getElementById(targetId) : null;
        if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    });

    const scrollPreviewBtn = document.getElementById("scrollPreviewBtn");
    if (scrollPreviewBtn) {
      scrollPreviewBtn.addEventListener("click", () => {
        document.getElementById("previewPanel")?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    }
  }

  function buildCrgEmailPayload() {
    const data = buildPayloadFromForm();
    const subject = `${data.noteReference} — ${data.title}`;
    const paragraphs = [
      "Hi CRG Research,",
      "Please find the latest research note attached.",
      [
        `Reference: ${data.noteReference}`,
        `Type: ${data.noteType}`,
        `Publication type: ${data.noteStage}`,
        `Distribution: ${data.distributionClass}`,
        `Horizon: ${data.noteHorizon}`,
        data.ticker ? `Ticker: ${data.ticker}` : null,
        data.crgRating ? `Rating: ${data.crgRating}` : null,
        data.targetPrice ? `Target price: ${data.targetPrice}` : null,
        `Generated: ${data.dateTimeString}`
      ].filter(Boolean).join("\n"),
      "Best,",
      data.analysts[0]?.name || ""
    ];

    return {
      subject,
      body: paragraphs.join("\n\n"),
      cc: data.deskCc
    };
  }

  function ensureLibs() {
    if (typeof docx === "undefined") throw new Error("docx library not loaded. Refresh the page.");
    if (typeof saveAs === "undefined") throw new Error("FileSaver library not loaded. Refresh the page.");
  }

  function validatePublishCore() {
    const ids = isEquityMode() ? baseCoreIds.concat(equityCoreIds) : baseCoreIds;
    return ids.filter((id) => {
      const el = document.getElementById(id);
      return el && !isFilled(el);
    });
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
      tabStops: opts.tabStops ?? undefined,
      border: opts.border ?? undefined
    });
  }

  function heading(text, size = 34, color = BRAND.colors.navy) {
    return para(run(text, {
      font: BRAND.fonts.heading,
      size,
      bold: true,
      color
    }), { spacing: { after: 140 } });
  }

  function subheading(text, size = 26, color = BRAND.colors.navy) {
    return para(run(text, {
      font: BRAND.fonts.heading,
      size,
      bold: true,
      color
    }), { spacing: { before: 60, after: 100 } });
  }

  function bodyLine(text, after = 140) {
    return para(run(text, { size: 20 }), { spacing: { after } });
  }

  function bodyParagraphs(text) {
    const blocks = toParagraphs(text);
    if (!blocks.length) return [bodyLine("—")];
    return blocks.map((block) => bodyLine(block));
  }

  function bulletParagraphs(text) {
    const items = toBulletItems(text);
    if (!items.length) return [bodyLine("—")];
    return items.map((item) => new docx.Paragraph({
      text: item,
      bullet: { level: 0 },
      spacing: { after: 90 }
    }));
  }

  function pageBreak() {
    return para(new docx.PageBreak(), { spacing: { after: 0 } });
  }

  function tableCell(children, opts = {}) {
    return new docx.TableCell({
      width: opts.width ? { size: opts.width, type: docx.WidthType.PERCENTAGE } : undefined,
      shading: opts.shading ? { fill: opts.shading } : undefined,
      margins: opts.margins ?? { top: 130, bottom: 130, left: 130, right: 130 },
      borders: opts.borders ?? {
        top: { style: docx.BorderStyle.SINGLE, size: 1, color: BRAND.colors.border },
        bottom: { style: docx.BorderStyle.SINGLE, size: 1, color: BRAND.colors.border },
        left: { style: docx.BorderStyle.SINGLE, size: 1, color: BRAND.colors.border },
        right: { style: docx.BorderStyle.SINGLE, size: 1, color: BRAND.colors.border }
      },
      children: Array.isArray(children) ? children : [children]
    });
  }

  function buildSummaryMatrix(items, columns = 4) {
    const padded = items.slice();
    while (padded.length % columns !== 0) {
      padded.push({ label: "", value: "" });
    }

    const rows = [];
    for (let i = 0; i < padded.length; i += columns) {
      rows.push(new docx.TableRow({
        children: padded.slice(i, i + columns).map((item) => tableCell([
          para(run(item.label || "", { size: 16, bold: true, color: BRAND.colors.muted }), { spacing: { after: 20 } }),
          para(run(item.value || "—", { size: 22, bold: true, color: BRAND.colors.navy }), { spacing: { after: 0 } })
        ], {
          width: Math.floor(100 / columns),
          shading: BRAND.colors.cream
        }))
      }));
    }

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows
    });
  }

  function buildProjectedFinancialsTableDoc(text) {
    const rows = parseStructuredTable(text);
    if (!rows.length) return null;

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: rows.map((row, index) => new docx.TableRow({
        children: row.map((value) => tableCell(
          para(run(value, {
            size: 18,
            bold: index === 0,
            color: index === 0 ? BRAND.colors.navy : BRAND.colors.ink
          }), { spacing: { after: 0 } }),
          {
            width: Math.floor(100 / row.length),
            shading: index === 0 ? BRAND.colors.cream : undefined
          }
        ))
      }))
    });
  }

  function buildCalloutTable(items) {
    const columns = items.length || 1;
    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        new docx.TableRow({
          children: items.map((item) => tableCell([
            para(run(item.title, {
              font: BRAND.fonts.heading,
              size: 24,
              bold: true,
              color: BRAND.colors.navy
            }), { spacing: { after: 70 } }),
            ...(item.bullets ? bulletParagraphs(item.body) : bodyParagraphs(item.body))
          ], {
            width: Math.floor(100 / columns)
          }))
        })
      ]
    });
  }

  async function addImages(files) {
    const output = [];
    const list = Array.from(files || []);
    for (let index = 0; index < list.length; index += 1) {
      const file = list[index];
      try {
        const buffer = await file.arrayBuffer();
        const caption = file.name.replace(/\.[^/.]+$/, "");
        output.push(
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: buffer,
                transformation: { width: 520, height: 320 }
              })
            ],
            alignment: docx.AlignmentType.CENTER,
            spacing: { before: 120, after: 60 }
          }),
          para(run(`Figure ${index + 1}: ${caption}`, {
            size: 18,
            italics: true,
            color: BRAND.colors.muted
          }), { align: docx.AlignmentType.CENTER, spacing: { after: 180 } })
        );
      } catch {
        // skip broken image
      }
    }
    return output;
  }

  function headerTable(noteType, reference, distributionClass) {
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
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: [para(run(`${BRAND.short} | ${noteType} | ${reference}`, { size: 16, color: BRAND.colors.muted }), { spacing: { after: 0 } })]
            }),
            new docx.TableCell({
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: [para(run(distributionClass, { size: 16, color: BRAND.colors.muted, bold: true }), {
                spacing: { after: 0 },
                align: docx.AlignmentType.RIGHT
              })]
            })
          ]
        })
      ]
    });
  }

  function footerParagraph(customDisclaimer) {
    const text = [BRAND.disclaimers.internal, safeTrim(customDisclaimer)].filter(Boolean).join(" ");
    return new docx.Paragraph({
      children: [
        run(text, { size: 14, italics: true, color: BRAND.colors.muted }),
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

  function buildCoverBand(logoBytes, dateLabel, distributionClass) {
    const logoBlock = logoBytes
      ? new docx.Paragraph({
        children: [
          new docx.ImageRun({
            data: logoBytes,
            transformation: { width: 220, height: 88 }
          })
        ],
        spacing: { after: 0 }
      })
      : para(run(BRAND.name, {
        font: BRAND.fonts.heading,
        size: 34,
        bold: true,
        color: BRAND.colors.gold
      }), { spacing: { after: 0 } });

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
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: [logoBlock]
            }),
            new docx.TableCell({
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: [
                para(run(distributionClass, { size: 18, bold: true, color: BRAND.colors.navy }), {
                  align: docx.AlignmentType.RIGHT,
                  spacing: { after: 20 }
                }),
                para(run(dateLabel, { size: 18, color: BRAND.colors.muted }), {
                  align: docx.AlignmentType.RIGHT,
                  spacing: { after: 0 }
                })
              ]
            })
          ]
        })
      ]
    });
  }

  function buildHeroBand(noteType, noteStage, title, thesisHeadline, reference) {
    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        new docx.TableRow({
          children: [
            tableCell([
              para(run(`${noteType} | ${noteStage}`, { size: 22, bold: true, color: BRAND.colors.white }), { spacing: { after: 100 } }),
              para(run(title, {
                font: BRAND.fonts.heading,
                size: 48,
                color: BRAND.colors.white,
                bold: true
              }), { spacing: { after: 110 } }),
              para(run(thesisHeadline || reference, { size: 22, color: BRAND.colors.white }), { spacing: { after: 0 } })
            ], {
              shading: BRAND.colors.navy,
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              margins: { top: 260, bottom: 260, left: 260, right: 260 }
            })
          ]
        })
      ]
    });
  }

  async function createInstitutionalDocument(payload) {
    const logoBytes = await fetchAssetBytes(BRAND.logoPath);
    const noteReference = payload.noteReference;
    const longDate = formatDateLong(new Date());
    const displayCurrent = payload.currentPriceText || "—";
    const displayTarget = payload.targetPrice || "—";
    const displayUpside = payload.upsideText || "—";
    const displayReturn = payload.rangeReturnText || "—";
    const displayVol = payload.realisedVolText || "—";
    const figures = await addImages(payload.imageFiles);
    const projectedFinancialsTable = buildProjectedFinancialsTableDoc(payload.projectedFinancialsTable);
    const analystLines = payload.analysts.length
      ? payload.analysts.map((analyst) => `${analyst.name}${analyst.phone ? ` | ${analyst.phone}` : ""}`)
      : ["—"];

    const summaryItems = [
      { label: "Publication Type", value: payload.noteStage || "—" },
      { label: "Distribution", value: payload.distributionClass || "—" },
      { label: "Horizon", value: payload.noteHorizon || "—" },
      { label: "Reference", value: noteReference }
    ];

    if (payload.noteType === "Equity Research") {
      summaryItems.push(
        { label: "Ticker", value: payload.ticker || "—" },
        { label: "Recommendation", value: ratingToDisplay(payload.crgRating) },
        { label: "Current / Target", value: `${displayCurrent} / ${displayTarget}` },
        { label: "Upside", value: displayUpside }
      );
    }

    const children = [
      buildCoverBand(logoBytes, longDate, payload.distributionClass),
      buildHeroBand(payload.noteType, payload.noteStage, payload.title, payload.thesisHeadline, noteReference),
      bodyLine(" ", 60),
      buildSummaryMatrix(summaryItems, 4),
      bodyLine(" ", 40),
      subheading("Lead Analysts"),
      ...analystLines.map((line) => bodyLine(line, 90)),
      subheading("Key Takeaways"),
      ...bulletParagraphs(payload.keyTakeaways)
    ];

    if (payload.noteType === "Equity Research") {
      children.push(
        bodyLine(" ", 50),
        buildSummaryMatrix([
          { label: "Range Return", value: displayReturn },
          { label: "Volatility (ann.)", value: displayVol },
          { label: "Current Price", value: displayCurrent },
          { label: "Target Price", value: displayTarget }
        ], 4)
      );
    }

    children.push(
      pageBreak(),
      heading("Investment Thesis"),
      ...bodyParagraphs(payload.analysis),
      buildCalloutTable([
        { title: "Variant Perception", body: payload.variantPerception || "—", bullets: false },
        { title: "Catalysts", body: payload.catalysts || "—", bullets: true },
        { title: "Key Risks", body: payload.keyRisks || "—", bullets: true }
      ])
    );

    if (payload.noteType === "Equity Research") {
      children.push(
        subheading("Valuation"),
        ...bodyParagraphs(payload.valuationSummary),
        subheading("Key Assumptions"),
        ...bulletParagraphs(payload.keyAssumptions)
      );

      if (projectedFinancialsTable) {
        children.push(
          subheading(payload.projectedFinancialsTitle || "Projected Financials"),
          projectedFinancialsTable
        );
      }

      if (payload.priceChartImageBytes) {
        children.push(
          subheading("Price Chart"),
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: payload.priceChartImageBytes,
                transformation: { width: 560, height: 220 }
              })
            ],
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 160 }
          })
        );
      }

      children.push(
        subheading("Scenario / Sensitivity Notes"),
        ...bodyParagraphs(payload.scenarioNotes)
      );
    }

    children.push(
      subheading("Additional Detail"),
      ...bodyParagraphs(payload.content),
      subheading("The Cordoba View"),
      ...bodyParagraphs(payload.cordobaView)
    );

    if (payload.modelLink || payload.modelFileNames.length) {
      children.push(subheading("Model & Supporting Files"));
      if (payload.modelLink) children.push(bodyLine(`Model link: ${payload.modelLink}`));
      if (payload.modelFileNames.length) {
        payload.modelFileNames.forEach((fileName) => {
          children.push(new docx.Paragraph({
            text: fileName,
            bullet: { level: 0 },
            spacing: { after: 70 }
          }));
        });
      }
    }

    if (figures.length) {
      children.push(subheading("Figures & Charts"), ...figures);
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
            children: [headerTable(payload.noteType, noteReference, payload.distributionClass)]
          })
        },
        footers: {
          default: new docx.Footer({
            children: [footerParagraph(payload.customDisclaimer)]
          })
        },
        children
      }]
    });
  }

  function bindTopButtons() {
    const emailToCrgBtn = document.getElementById("emailToCrgBtn");
    const emailToCrgBtnTop = document.getElementById("emailToCrgBtnTop");
    const resetBtn = document.getElementById("resetFormBtn");
    const resetBtnTop = document.getElementById("resetFormBtnTop");

    function emailDesk() {
      const { subject, body, cc } = buildCrgEmailPayload();
      window.location.href = buildMailto("research@cordobarg.com", cc, subject, body);
    }

    function resetForm() {
      const ok = window.confirm("Reset the form? This clears all fields and removes the saved draft.");
      if (!ok) return;

      const form = document.getElementById("researchForm");
      if (form) form.reset();
      if (coAuthorsList) coAuthorsList.innerHTML = "";
      if (modelFilesEl) modelFilesEl.value = "";
      if ($("#imageUpload")) $("#imageUpload").value = "";
      clearChartUI();
      clearDraft();
      setDraftStatus("—");
      syncPrimaryPhone();
      paintEquityStats();
      showMsg("", "");
      refreshUI();
    }

    if (emailToCrgBtn) emailToCrgBtn.addEventListener("click", emailDesk);
    if (emailToCrgBtnTop) emailToCrgBtnTop.addEventListener("click", emailDesk);
    if (resetBtn) resetBtn.addEventListener("click", resetForm);
    if (resetBtnTop) resetBtnTop.addEventListener("click", resetForm);
  }

  function bindFormEvents() {
    const form = document.getElementById("researchForm");
    if (!form) return;

    ["input", "change"].forEach((eventName) => {
      form.addEventListener(eventName, () => {
        syncPrimaryPhone();
        scheduleDraftSave();
        refreshUI();
      }, { passive: true });
    });

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      showMsg("", "");

      const missing = validatePublishCore();
      if (missing.length) {
        showMsg("error", `Missing publish-core fields: ${missing.join(", ")}`);
        const first = document.getElementById(missing[0]);
        if (first) first.focus();
        return;
      }

      const submitButton = form.querySelector('button[type="submit"]');
      if (submitButton) {
        submitButton.disabled = true;
        submitButton.classList.add("loading");
        submitButton.textContent = "Generating…";
      }

      try {
        ensureLibs();
        const payload = buildPayloadFromForm();
        payload.priceChartImageBytes = state.priceChartImageBytes;
        const doc = await createInstitutionalDocument(payload);
        const blob = await docx.Packer.toBlob(doc);
        const fileName = `${payload.title.replace(/[^a-z0-9]/gi, "_").toLowerCase()}_${noteTypeCode(payload.noteType).toLowerCase()}_${formatDateShortISO(new Date())}.docx`;
        saveAs(blob, fileName);
        saveDraftNow();
        showMsg("success", `Document "${fileName}" generated successfully.`);
      } catch (error) {
        showMsg("error", `Error: ${error.message}`);
      } finally {
        if (submitButton) {
          submitButton.disabled = false;
          submitButton.classList.remove("loading");
          submitButton.textContent = "Generate Word Document";
        }
      }
    });
  }

  function init() {
    bindSectionJumpButtons();
    bindTopButtons();
    bindFormEvents();

    if (authorPhoneNationalEl) {
      authorPhoneNationalEl.addEventListener("input", () => {
        formatPrimaryVisible();
        scheduleDraftSave();
        refreshUI();
      });
    }

    if (authorPhoneCountryEl) {
      authorPhoneCountryEl.addEventListener("change", () => {
        syncPrimaryPhone();
        scheduleDraftSave();
        refreshUI();
      });
    }

    if (addCoAuthorBtn && coAuthorsList) {
      addCoAuthorBtn.addEventListener("click", () => {
        coAuthorsList.appendChild(createCoauthorNode());
        scheduleDraftSave();
        refreshUI();
      });
    }

    document.addEventListener("click", (event) => {
      const button = event.target.closest(".remove-coauthor");
      if (!button) return;
      const id = button.getAttribute("data-remove-id");
      const node = id ? document.getElementById(`coauthor-${id}`) : null;
      if (node) node.remove();
      scheduleDraftSave();
      refreshUI();
    });

    const fetchChartBtn = document.getElementById("fetchPriceChart");
    if (fetchChartBtn) {
      fetchChartBtn.addEventListener("click", buildPriceChart);
    }

    const targetPriceEl = document.getElementById("targetPrice");
    if (targetPriceEl) {
      targetPriceEl.addEventListener("input", () => {
        paintEquityStats();
        refreshUI();
      });
    }

    if (modelFilesEl) {
      modelFilesEl.addEventListener("change", () => {
        updateAttachmentSummary();
        refreshUI();
      });
    }

    const draft = loadDraft();
    if (draft) {
      applyDraft(draft);
      setDraftStatus("Restored");
    } else {
      setDraftStatus("—");
    }

    syncPrimaryPhone();
    paintEquityStats();
    refreshUI();
  }

  window.addEventListener("DOMContentLoaded", init);
})();
