// assets/app.js
console.log("app.js loaded successfully");

window.addEventListener("DOMContentLoaded", () => {
  // ============================================================
  // CRG RDT — BlueMatrix-style workflow + versioning + templates
  // ============================================================

  // ----------------
  // Utilities
  // ----------------
  const $ = (id) => document.getElementById(id);

  function digitsOnly(v){ return (v||"").toString().replace(/\D/g,""); }
  function safeTrim(v){ return (v ?? "").toString().trim(); }
  function naIfBlank(v){ const s = safeTrim(v); return s ? s : "N/A"; }

  function formatDateTime(date){
    const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
    const month = months[date.getMonth()];
    const day = date.getDate();
    const year = date.getFullYear();
    let hours = date.getHours();
    const minutes = date.getMinutes().toString().padStart(2,"0");
    const ampm = hours >= 12 ? "PM" : "AM";
    hours = hours % 12 || 12;
    return `${month} ${day}, ${year} ${hours}:${minutes} ${ampm}`;
  }
  function formatDateShort(date){
    const y = date.getFullYear();
    const m = String(date.getMonth()+1).padStart(2,"0");
    const d = String(date.getDate()).padStart(2,"0");
    return `${y}-${m}-${d}`;
  }

  // ----------------
  // Mailto builder (avoid '+' spaces)
  // ----------------
  function buildMailto(to, cc, subject, body){
    const crlfBody = (body || "").replace(/\n/g, "\r\n");
    const parts = [];
    if (cc) parts.push(`cc=${encodeURIComponent(cc)}`);
    parts.push(`subject=${encodeURIComponent(subject || "")}`);
    parts.push(`body=${encodeURIComponent(crlfBody)}`);
    return `mailto:${encodeURIComponent(to)}?${parts.join("&")}`;
  }

  // ----------------
  // Phone formatting (primary + coauthors)
  // ----------------
  function formatNationalLoose(rawDigits){
    const d = digitsOnly(rawDigits);
    if (!d) return "";
    const p1 = d.slice(0,4);
    const p2 = d.slice(4,7);
    const p3 = d.slice(7,10);
    const rest = d.slice(10);
    return [p1,p2,p3,rest].filter(Boolean).join(" ");
  }
  function buildInternationalHyphen(ccDigits, nationalDigits){
    const cc = digitsOnly(ccDigits);
    const nn = digitsOnly(nationalDigits);
    if (!cc && !nn) return "";
    if (cc && !nn) return `${cc}-`;
    if (!cc && nn) return nn;
    return `${cc}-${nn}`;
  }

  const authorPhoneCountryEl  = $("authorPhoneCountry");
  const authorPhoneNationalEl = $("authorPhoneNational");
  const authorPhoneHiddenEl   = $("authorPhone"); // hidden source-of-truth

  function syncPrimaryPhone(){
    if (!authorPhoneHiddenEl) return;
    const cc = authorPhoneCountryEl ? authorPhoneCountryEl.value : "";
    const nn = digitsOnly(authorPhoneNationalEl ? authorPhoneNationalEl.value : "");
    authorPhoneHiddenEl.value = buildInternationalHyphen(cc, nn);
  }
  function formatPrimaryVisible(){
    if (!authorPhoneNationalEl) return;
    const caret = authorPhoneNationalEl.selectionStart || 0;
    const beforeLen = authorPhoneNationalEl.value.length;

    authorPhoneNationalEl.value = formatNationalLoose(authorPhoneNationalEl.value);

    const afterLen = authorPhoneNationalEl.value.length;
    const delta = afterLen - beforeLen;
    const next = Math.max(0, caret + delta);
    try { authorPhoneNationalEl.setSelectionRange(next,next); } catch(_){}

    syncPrimaryPhone();
  }

  if (authorPhoneNationalEl){
    authorPhoneNationalEl.addEventListener("input", formatPrimaryVisible);
    authorPhoneNationalEl.addEventListener("blur", syncPrimaryPhone);
  }
  if (authorPhoneCountryEl){
    authorPhoneCountryEl.addEventListener("change", syncPrimaryPhone);
  }
  syncPrimaryPhone();

  // ----------------
  // Co-author management
  // ----------------
  let coAuthorCount = 0;
  const addCoAuthorBtn = $("addCoAuthor");
  const coAuthorsList  = $("coAuthorsList");

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
    <option value="880">+880</option>
    <option value="91">+91</option>
    <option value="234">+234</option>
    <option value="254">+254</option>
    <option value="27">+27</option>
    <option value="995">+995</option>
    <option value="">Other</option>
  `;

  function wireCoauthorPhone(coAuthorDiv){
    const ccEl      = coAuthorDiv.querySelector(".coauthor-country");
    const nationalEl= coAuthorDiv.querySelector(".coauthor-phone-local");
    const hiddenEl  = coAuthorDiv.querySelector(".coauthor-phone");
    if (!hiddenEl) return;

    function syncHidden(){
      const cc = ccEl ? ccEl.value : "";
      const nn = digitsOnly(nationalEl ? nationalEl.value : "");
      hiddenEl.value = buildInternationalHyphen(cc, nn);
    }
    function formatVisible(){
      if (!nationalEl) return;
      const caret = nationalEl.selectionStart || 0;
      const beforeLen = nationalEl.value.length;

      nationalEl.value = formatNationalLoose(nationalEl.value);

      const afterLen = nationalEl.value.length;
      const delta = afterLen - beforeLen;
      const next = Math.max(0, caret + delta);
      try { nationalEl.setSelectionRange(next,next); } catch(_){}

      syncHidden();
    }

    if (nationalEl){
      nationalEl.addEventListener("input", formatVisible);
      nationalEl.addEventListener("blur", syncHidden);
    }
    if (ccEl) ccEl.addEventListener("change", syncHidden);

    syncHidden();
  }

  if (addCoAuthorBtn){
    addCoAuthorBtn.addEventListener("click", () => {
      coAuthorCount++;

      const coAuthorDiv = document.createElement("div");
      coAuthorDiv.className = "coauthor-row";
      coAuthorDiv.id = `coauthor-${coAuthorCount}`;

      coAuthorDiv.innerHTML = `
        <div class="coauthor-grid">
          <div>
            <label>Last Name</label>
            <input type="text" class="coauthor-lastname" placeholder="e.g., Rahman">
          </div>
          <div>
            <label>First Name</label>
            <input type="text" class="coauthor-firstname" placeholder="e.g., Nadir">
          </div>
          <div>
            <label>Phone (Optional)</label>
            <div class="phone-row phone-row--compact">
              <select class="phone-country coauthor-country" aria-label="Country code">${countryOptionsHtml}</select>
              <input type="text" class="phone-number coauthor-phone-local" inputmode="numeric" placeholder="e.g., 7323 324 120">
            </div>
            <input type="text" class="coauthor-phone" style="display:none;">
          </div>
          <div style="display:flex; align-items:flex-end;">
            <button type="button" class="btn btn-danger remove-coauthor" data-remove-id="${coAuthorCount}">Remove</button>
          </div>
        </div>
      `;

      coAuthorsList.appendChild(coAuthorDiv);
      wireCoauthorPhone(coAuthorDiv);
      updateCompletionMeter();
      refreshValidationSummary();
    });

    document.addEventListener("click", (e) => {
      const btn = e.target.closest(".remove-coauthor");
      if (!btn) return;
      const id = btn.getAttribute("data-remove-id");
      const row = document.getElementById(`coauthor-${id}`);
      if (row) row.remove();
      updateCompletionMeter();
      refreshValidationSummary();
    });
  }

  function coAuthorLine(coAuthor){
    const ln = (coAuthor.lastName || "").toUpperCase();
    const fn = (coAuthor.firstName || "").toUpperCase();
    const ph = naIfBlank(coAuthor.phone);
    return `${ln}, ${fn} (${ph})`;
  }

  // ----------------
  // Equity section toggle
  // ----------------
  const noteTypeEl = $("noteType");
  const equitySectionEl = $("equitySection");
  const equityRailLink = $("equityRailLink"); // optional in left rail

  function toggleEquitySection(){
    if (!noteTypeEl || !equitySectionEl) return;
    const isEquity = noteTypeEl.value === "Equity Research";
    equitySectionEl.style.display = isEquity ? "block" : "none";
    if (equityRailLink) equityRailLink.style.display = isEquity ? "list-item" : "none";
  }
  if (noteTypeEl && equitySectionEl){
    noteTypeEl.addEventListener("change", () => {
      toggleEquitySection();
      setTimeout(() => {
        updateCompletionMeter();
        refreshValidationSummary();
      }, 0);
    });
    toggleEquitySection();
  }

  // ============================================================
  // (NEW) BlueMatrix-style workflow: Status + Reviewer + Presets
  // ============================================================
  const statusEl     = $("workflowStatus");     // Draft / Reviewed / Cleared
  const reviewedByEl = $("reviewedBy");         // dropdown
  const presetEl     = $("distributionPreset"); // Internal only / Public pack / Client-safe

  // Routing + redaction rules by preset
  const PRESET = {
    INTERNAL: "Internal only",
    PUBLIC:   "Public pack",
    CLIENT:   "Client-safe"
  };

  function getPreset(){
    return safeTrim(presetEl?.value) || PRESET.INTERNAL;
  }
  function getStatus(){
    return safeTrim(statusEl?.value) || "Draft";
  }

  // ============================================================
  // (NEW) Versioning + Change Note (BlueMatrix “knows what changed”)
  // ============================================================
  const versionTextEl = $("versionText"); // UI display like v1.0
  const changeNoteEl  = $("changeNote");  // optional text area

  function noteKey(){
    // stable key by note type + title (trimmed, lower)
    const t = safeTrim($("title")?.value).toLowerCase();
    const nt = safeTrim(noteTypeEl?.value).toLowerCase();
    // if empty title, still keep a key so version doesn't crash
    return `crg_rdt_version__${nt}__${t || "untitled"}`;
  }

  function readVersionState(){
    const raw = localStorage.getItem(noteKey());
    if (!raw) return { major: 1, minor: 0, lastStatus: "Draft" };
    try {
      const obj = JSON.parse(raw);
      const major = Number.isFinite(+obj.major) ? +obj.major : 1;
      const minor = Number.isFinite(+obj.minor) ? +obj.minor : 0;
      return { major, minor, lastStatus: obj.lastStatus || "Draft" };
    } catch(_) {
      return { major: 1, minor: 0, lastStatus: "Draft" };
    }
  }

  function writeVersionState(state){
    localStorage.setItem(noteKey(), JSON.stringify(state));
  }

  function versionString(state){
    return `v${state.major}.${state.minor}`;
  }

  function paintVersion(){
    const st = readVersionState();
    if (versionTextEl) versionTextEl.textContent = versionString(st);
  }

  // Rule set:
  // - Every export increments minor by 1 by default.
  // - If status transitions into Cleared (from non-cleared), bump major and reset minor=0.
  // - If title or noteType changes, the noteKey changes, so version resets to v1.0 automatically.
  function bumpVersionOnExport(){
    const cur = readVersionState();
    const status = getStatus();

    let major = cur.major;
    let minor = cur.minor;
    const lastStatus = cur.lastStatus || "Draft";

    if (status === "Cleared" && lastStatus !== "Cleared"){
      major = major + 1;
      minor = 0;
    } else {
      minor = minor + 1;
    }

    const next = { major, minor, lastStatus: status };
    writeVersionState(next);
    paintVersion();
    return versionString(next);
  }

  if ($("title")) $("title").addEventListener("input", paintVersion);
  if (noteTypeEl) noteTypeEl.addEventListener("change", paintVersion);
  if (statusEl) statusEl.addEventListener("change", paintVersion);
  paintVersion();

  // ============================================================
  // (NEW) Templates (Saved templates)
  // ============================================================
  const templateEl = $("noteTemplate"); // Macro update / Equity initiation / Event note
  const thesisEl   = $("thesis");       // thesis field
  const execToggleEl = $("autoExecSummary"); // checkbox toggle
  const execSummaryEl= $("execSummary");     // editable textarea

  function setVal(id, value){
    const el = $(id);
    if (!el) return;
    el.value = value;
  }

  const TEMPLATES = {
    "Macro update": {
      noteType: "Macro Research",
      topicPH: "e.g., UK inflation dynamics, EM FX, global liquidity...",
      titlePH: "Aim for a clear investor-led statement (not a blog headline).",
      thesis: "In one line: what’s changed, what’s priced, what’s mispriced.",
      keyTakeaways:
        "- What changed and why it matters\n- What markets are pricing\n- What we think is mispriced\n- Key risks / what would change our view",
      analysis:
        "Write thesis → evidence → transmission → risks → positioning.\n\nInclude key data points, comparables, and cross-asset implications.",
      cordobaView:
        "Our stance, conditions for change, and what we would position for."
    },
    "Equity initiation": {
      noteType: "Equity Research",
      topicPH: "e.g., Company / sector / key driver (pricing power, margins, capex cycle...)",
      titlePH: "Initiation: concise view + angle (e.g., 'Underpriced optionality in ...').",
      thesis: "We initiate with a clear view driven by X; market is mispricing Y.",
      keyTakeaways:
        "- Rating + target + time horizon\n- Core driver(s) of upside/downside\n- What the market is missing\n- Key risks / disconfirmers",
      analysis:
        "Business quality → unit economics → catalysts → valuation → risks.\n\nBe explicit on what needs to happen for the view to work.",
      cordobaView:
        "Positioning, catalysts, and what would change the view."
    },
    "Event note": {
      noteType: "General Note",
      topicPH: "e.g., Earnings, CPI print, policy decision, regulatory change...",
      titlePH: "Event-driven: what happened + so-what.",
      thesis: "What happened and what it changes (and what it doesn’t).",
      keyTakeaways:
        "- Event summary in one line\n- Immediate market reaction\n- Second-order implications\n- Risks / next catalysts",
      analysis:
        "What happened → why → who is affected → pricing response → second-order effects.\n\nClose with a clear so-what.",
      cordobaView:
        "Our stance, next steps, and positioning bias."
    }
  };

  function applyTemplate(name){
    const t = TEMPLATES[name];
    if (!t) return;

    // note type
    if (noteTypeEl && t.noteType){
      noteTypeEl.value = t.noteType;
      toggleEquitySection();
    }

    // placeholders (keep your existing placeholders if you prefer)
    const topic = $("topic");
    const title = $("title");
    if (topic && t.topicPH) topic.placeholder = t.topicPH;
    if (title && t.titlePH) title.placeholder = t.titlePH;

    // prefill (only if empty, to avoid clobbering work)
    if (thesisEl && !safeTrim(thesisEl.value)) thesisEl.value = t.thesis || "";
    if ($("keyTakeaways") && !safeTrim($("keyTakeaways").value)) $("keyTakeaways").value = t.keyTakeaways || "";
    if ($("analysis") && !safeTrim($("analysis").value)) $("analysis").value = t.analysis || "";
    if ($("cordobaView") && !safeTrim($("cordobaView").value)) $("cordobaView").value = t.cordobaView || "";

    updateCompletionMeter();
    refreshValidationSummary();
  }

  if (templateEl){
    templateEl.addEventListener("change", () => {
      applyTemplate(templateEl.value);
    });
  }

  // ============================================================
  // Completion meter + validation summary (updated core logic)
  // ============================================================
  const completionTextEl = $("completionText");
  const completionBarEl  = $("completionBar");
  const validationSummaryEl = $("validationSummary"); // optional

  function isFilled(el){
    if (!el) return false;
    if (el.type === "file") return el.files && el.files.length > 0;
    const v = (el.value ?? "").toString().trim();
    return v.length > 0;
  }

  // Required core fields (institutional baseline)
  const baseCoreIds = [
    "noteType",
    "topic",
    "title",
    "authorLastName",
    "authorFirstName",
    "thesis",
    "keyTakeaways",
    "analysis",
    "cordobaView"
  ];

  // Equity module required fields *when noteType is Equity Research*
  // (keeps this strict + practical)
  const equityCoreIds = [
    "crgRating",
    "targetPrice",
    "modelFiles"
  ];

  // Workflow requirements: reviewer for Reviewed/Cleared, plus compliance ack always.
  function workflowRequirementsMet(){
    const status = getStatus();
    const ack = $("complianceAck");
    const ackOk = !!ack?.checked;

    if (!ackOk) return false;
    if ((status === "Reviewed" || status === "Cleared") && !isFilled(reviewedByEl)) return false;

    // Client-safe preset must be Cleared
    const preset = getPreset();
    if (preset === PRESET.CLIENT && status !== "Cleared") return false;

    return true;
  }

  function updateCompletionMeter(){
    const isEquity = (noteTypeEl && noteTypeEl.value === "Equity Research" && equitySectionEl && equitySectionEl.style.display !== "none");
    const ids = isEquity ? baseCoreIds.concat(equityCoreIds) : baseCoreIds;

    let done = 0;
    ids.forEach((id) => {
      const el = $(id);
      if (isFilled(el)) done++;
    });

    // compliance ack + routing adds “institutional” completion
    const ack = $("complianceAck");
    const reviewerNeeded = (getStatus() === "Reviewed" || getStatus() === "Cleared");
    const ackDone = !!ack?.checked;
    const reviewerDone = reviewerNeeded ? isFilled(reviewedByEl) : true;

    const extraTotal = 1 + (reviewerNeeded ? 1 : 0);
    const extraDone  = (ackDone ? 1 : 0) + (reviewerDone ? 1 : 0);

    const total = ids.length + extraTotal;
    const doneAll = done + extraDone;

    const pct = total ? Math.round((doneAll / total) * 100) : 0;

    if (completionTextEl) completionTextEl.textContent = `${doneAll} / ${total} required`;
    if (completionBarEl) completionBarEl.style.width = `${pct}%`;
    const bar = completionBarEl?.parentElement;
    if (bar) bar.setAttribute("aria-valuenow", String(pct));
  }

  function listMissing(){
    const isEquity = (noteTypeEl && noteTypeEl.value === "Equity Research" && equitySectionEl && equitySectionEl.style.display !== "none");
    const ids = isEquity ? baseCoreIds.concat(equityCoreIds) : baseCoreIds;

    const missing = [];
    ids.forEach((id) => {
      const el = $(id);
      if (!isFilled(el)) missing.push(id);
    });

    const ack = $("complianceAck");
    if (!ack?.checked) missing.push("complianceAck");

    const status = getStatus();
    if ((status === "Reviewed" || status === "Cleared") && !isFilled(reviewedByEl)){
      missing.push("reviewedBy");
    }

    const preset = getPreset();
    if (preset === PRESET.CLIENT && status !== "Cleared"){
      // not a field, but show as actionable
      missing.push("workflowStatus");
    }

    return missing;
  }

  function humanLabel(id){
    const map = {
      noteType: "Type of note",
      topic: "Topic",
      title: "Title",
      authorLastName: "Primary author — last name",
      authorFirstName: "Primary author — first name",
      thesis: "Thesis",
      keyTakeaways: "Key takeaways",
      analysis: "Analysis and commentary",
      cordobaView: "The Cordoba view",
      crgRating: "CRG rating",
      targetPrice: "Target price",
      modelFiles: "Model files",
      complianceAck: "Compliance acknowledgement",
      reviewedBy: "Reviewed by",
      workflowStatus: "Status (must be Cleared for Client-safe)"
    };
    return map[id] || id;
  }

  function refreshValidationSummary(){
    if (!validationSummaryEl) return;
    const missing = listMissing();
    if (!missing.length){
      validationSummaryEl.textContent = "Ready: all required fields and workflow checks complete.";
      return;
    }
    validationSummaryEl.textContent =
      "Missing: " + missing.map(humanLabel).join(" • ");
  }

  // Jump to first missing button
  const jumpMissingBtn = $("jumpMissingBtn");
  if (jumpMissingBtn){
    jumpMissingBtn.addEventListener("click", () => {
      const missing = listMissing();
      if (!missing.length) return;

      const first = missing[0];
      const el = $(first);
      if (el && typeof el.scrollIntoView === "function"){
        el.scrollIntoView({ behavior: "smooth", block: "center" });
        try { el.focus(); } catch(_){}
      }
    });
  }

  // global listeners
  ["input","change","keyup"].forEach(evt => {
    document.addEventListener(evt, (e) => {
      const t = e.target;
      if (!t) return;
      if (t.closest && t.closest("#researchForm")){
        updateCompletionMeter();
        refreshValidationSummary();
      }
    }, { passive: true });
  });

  // Workflow listeners
  if (statusEl) statusEl.addEventListener("change", () => { updateCompletionMeter(); refreshValidationSummary(); paintVersion(); });
  if (reviewedByEl) reviewedByEl.addEventListener("change", () => { updateCompletionMeter(); refreshValidationSummary(); });
  if (presetEl) presetEl.addEventListener("change", () => { updateCompletionMeter(); refreshValidationSummary(); });

  // ============================================================
  // (NEW) Auto Exec Summary (editable)
  // Pull: first takeaway bullet + thesis + target/rating (if equity)
  // ============================================================
  function firstTakeawayLine(){
    const raw = safeTrim($("keyTakeaways")?.value);
    if (!raw) return "";
    const lines = raw.split("\n").map(l => l.replace(/^[-*•]\s*/, "").trim()).filter(Boolean);
    return lines[0] || "";
  }

  function buildExecSummaryAuto(){
    const thesis = safeTrim(thesisEl?.value);
    const firstBullet = firstTakeawayLine();
    const noteType = safeTrim(noteTypeEl?.value);

    const parts = [];
    if (firstBullet) parts.push(firstBullet);
    if (thesis) parts.push(thesis);

    if (noteType === "Equity Research"){
      const rating = safeTrim($("crgRating")?.value);
      const tp = safeTrim($("targetPrice")?.value);
      const ticker = safeTrim($("ticker")?.value);

      const line = [
        ticker ? `Ticker: ${ticker}` : null,
        rating ? `Rating: ${rating}` : null,
        tp ? `Target: ${tp}` : null
      ].filter(Boolean).join(" • ");

      if (line) parts.push(line);
    }

    // keep it “one-page”: short paragraphs
    return parts.filter(Boolean).join("\n\n");
  }

  function syncExecSummary(){
    if (!execToggleEl || !execSummaryEl) return;
    if (execToggleEl.checked){
      execSummaryEl.value = buildExecSummaryAuto();
    }
  }

  if (execToggleEl){
    execToggleEl.addEventListener("change", () => {
      // if toggled on, populate immediately
      if (execToggleEl.checked) syncExecSummary();
      updateCompletionMeter();
      refreshValidationSummary();
    });
  }

  // Rebuild auto-summary when key fields change (only if toggle on)
  ["thesis","keyTakeaways","crgRating","targetPrice","ticker","noteType"].forEach((id) => {
    const el = $(id);
    if (!el) return;
    el.addEventListener("input", () => {
      if (execToggleEl?.checked) syncExecSummary();
    });
    el.addEventListener("change", () => {
      if (execToggleEl?.checked) syncExecSummary();
    });
  });

  // ============================================================
  // (NEW) Inline data citation + chart annotation
  // ============================================================
  const dataSourceEl = $("priceDataSource");     // dropdown
  const chartNoteEl  = $("chartAnnotation");     // textarea

  function selectedDataSource(){
    const v = safeTrim(dataSourceEl?.value);
    return v || "Stooq (via proxy)";
  }

  // ============================================================
  // (NEW) Scenario table generator (Bear/Base/Bull)
  // ============================================================
  const scBearPriceEl = $("scenarioBearPrice");
  const scBearKeyEl   = $("scenarioBearKey");
  const scBasePriceEl = $("scenarioBasePrice");
  const scBaseKeyEl   = $("scenarioBaseKey");
  const scBullPriceEl = $("scenarioBullPrice");
  const scBullKeyEl   = $("scenarioBullKey");

  function scenarioRows(){
    const rows = [
      { name: "Bear", price: safeTrim(scBearPriceEl?.value), key: safeTrim(scBearKeyEl?.value) },
      { name: "Base", price: safeTrim(scBasePriceEl?.value), key: safeTrim(scBaseKeyEl?.value) },
      { name: "Bull", price: safeTrim(scBullPriceEl?.value), key: safeTrim(scBullKeyEl?.value) }
    ];
    // include rows if at least one field provided
    const any = rows.some(r => r.price || r.key);
    return any ? rows : [];
  }

  // ============================================================
  // Attachment summary (modelFiles)
  // ============================================================
  const modelFilesEl2 = $("modelFiles");
  const attachSummaryHeadEl = $("attachmentSummaryHead");
  const attachSummaryListEl = $("attachmentSummaryList");

  function updateAttachmentSummary(){
    if (!modelFilesEl2 || !attachSummaryHeadEl || !attachSummaryListEl) return;

    const files = Array.from(modelFilesEl2.files || []);
    if (!files.length){
      attachSummaryHeadEl.textContent = "No files selected";
      attachSummaryListEl.style.display = "none";
      attachSummaryListEl.innerHTML = "";
      return;
    }
    attachSummaryHeadEl.textContent = `${files.length} file${files.length===1?"":"s"} selected`;
    attachSummaryListEl.style.display = "block";
    attachSummaryListEl.innerHTML = files.map(f => `<div class="attach-file">${f.name}</div>`).join("");
  }
  if (modelFilesEl2){
    modelFilesEl2.addEventListener("change", () => {
      updateAttachmentSummary();
      updateCompletionMeter();
      refreshValidationSummary();
    });
  }

  // ============================================================
  // Reset + clear autosave
  // ============================================================
  const resetBtn = $("resetFormBtn");
  const clearAutosaveBtn = $("clearAutosaveBtn");
  const formEl = $("researchForm");

  function clearChartUI(){
    const setText = (id, text) => { const el = $(id); if (el) el.textContent = text; };
    setText("currentPrice","—");
    setText("realisedVol","—");
    setText("rangeReturn","—");
    setText("upsideToTarget","—");
    const chartStatus = $("chartStatus");
    if (chartStatus) chartStatus.textContent = "";
    if (typeof priceChart !== "undefined" && priceChart){
      try { priceChart.destroy(); } catch(_){}
      priceChart = null;
    }
    if (typeof priceChartImageBytes !== "undefined") priceChartImageBytes = null;
    if (typeof equityStats !== "undefined"){
      equityStats = { currentPrice:null, realisedVolAnn:null, rangeReturn:null };
    }
  }

  function wipeAutosave(){
    // wipe all autosave keys (prefix)
    Object.keys(localStorage).forEach(k => {
      if (k.startsWith("crg_rdt_autosave__")) localStorage.removeItem(k);
    });
  }

  if (clearAutosaveBtn){
    clearAutosaveBtn.addEventListener("click", () => {
      const ok = confirm("Clear autosave? This removes locally saved drafts from this browser.");
      if (!ok) return;
      wipeAutosave();
      alert("Autosave cleared.");
    });
  }

  if (resetBtn && formEl){
    resetBtn.addEventListener("click", () => {
      const ok = confirm("Reset the form? This clears all fields on this page.");
      if (!ok) return;

      formEl.reset();
      if (coAuthorsList) coAuthorsList.innerHTML = "";
      if (modelFilesEl2) modelFilesEl2.value = "";
      updateAttachmentSummary();

      const messageDiv = $("message");
      if (messageDiv){
        messageDiv.className = "message";
        messageDiv.textContent = "";
        messageDiv.style.display = "none";
      }

      clearChartUI();
      syncPrimaryPhone();
      toggleEquitySection();

      // do NOT wipe version history; that’s per noteKey (title/type)
      // but repaint the displayed version
      paintVersion();

      setTimeout(() => {
        updateCompletionMeter();
        refreshValidationSummary();
      }, 0);
    });
  }

  // ============================================================
  // Autosave (simple, local-only)
  // ============================================================
  const AUTOSAVE_PREFIX = "crg_rdt_autosave__";
  const AUTOSAVE_FIELDS = [
    "noteTemplate","noteType","distributionPreset","workflowStatus","reviewedBy",
    "topic","title",
    "authorLastName","authorFirstName","authorPhone",
    "thesis","autoExecSummary","execSummary",
    "keyTakeaways","analysis","content","cordobaView",
    "ticker","crgRating","targetPrice","modelLink",
    "valuationSummary","keyAssumptions","scenarioNotes",
    "priceDataSource","chartAnnotation",
    "scenarioBearPrice","scenarioBearKey","scenarioBasePrice","scenarioBaseKey","scenarioBullPrice","scenarioBullKey",
    "changeNote"
  ];

  function autosaveKey(){ return `${AUTOSAVE_PREFIX}draft`; }

  function autosaveNow(){
    const obj = {};
    AUTOSAVE_FIELDS.forEach(id => {
      const el = $(id);
      if (!el) return;
      if (el.type === "checkbox") obj[id] = !!el.checked;
      else obj[id] = el.value;
    });
    // coauthors too
    const co = [];
    document.querySelectorAll(".coauthor-row").forEach(entry => {
      const ln = safeTrim(entry.querySelector(".coauthor-lastname")?.value);
      const fn = safeTrim(entry.querySelector(".coauthor-firstname")?.value);
      const ph = safeTrim(entry.querySelector(".coauthor-phone")?.value);
      if (ln || fn || ph) co.push({ ln, fn, ph });
    });
    obj.__coauthors = co;

    localStorage.setItem(autosaveKey(), JSON.stringify(obj));
  }

  function restoreAutosave(){
    const raw = localStorage.getItem(autosaveKey());
    if (!raw) return;
    try {
      const obj = JSON.parse(raw);

      AUTOSAVE_FIELDS.forEach(id => {
        const el = $(id);
        if (!el) return;
        if (el.type === "checkbox") el.checked = !!obj[id];
        else if (obj[id] != null) el.value = obj[id];
      });

      // restore coauthors
      const co = Array.isArray(obj.__coauthors) ? obj.__coauthors : [];
      if (coAuthorsList) coAuthorsList.innerHTML = "";
      co.forEach(item => {
        coAuthorCount++;
        const div = document.createElement("div");
        div.className = "coauthor-row";
        div.id = `coauthor-${coAuthorCount}`;
        div.innerHTML = `
          <div class="coauthor-grid">
            <div>
              <label>Last Name</label>
              <input type="text" class="coauthor-lastname" placeholder="e.g., Rahman" value="${(item.ln||"").replace(/"/g,"&quot;")}">
            </div>
            <div>
              <label>First Name</label>
              <input type="text" class="coauthor-firstname" placeholder="e.g., Nadir" value="${(item.fn||"").replace(/"/g,"&quot;")}">
            </div>
            <div>
              <label>Phone (Optional)</label>
              <div class="phone-row phone-row--compact">
                <select class="phone-country coauthor-country" aria-label="Country code">${countryOptionsHtml}</select>
                <input type="text" class="phone-number coauthor-phone-local" inputmode="numeric" placeholder="e.g., 7323 324 120">
              </div>
              <input type="text" class="coauthor-phone" style="display:none;">
            </div>
            <div style="display:flex; align-items:flex-end;">
              <button type="button" class="btn btn-danger remove-coauthor" data-remove-id="${coAuthorCount}">Remove</button>
            </div>
          </div>
        `;
        coAuthorsList.appendChild(div);

        // set phone hidden directly if saved; set visible formatted if possible
        const hidden = div.querySelector(".coauthor-phone");
        const local  = div.querySelector(".coauthor-phone-local");
        if (hidden) hidden.value = item.ph || "";
        if (local) local.value = formatNationalLoose((item.ph||"").split("-")[1] || "");
        wireCoauthorPhone(div);
      });

      // sync phone hidden
      syncPrimaryPhone();

      // toggle sections
      toggleEquitySection();

      // exec summary sync if toggle on
      if (execToggleEl?.checked) syncExecSummary();

      updateAttachmentSummary();
      paintVersion();
      updateCompletionMeter();
      refreshValidationSummary();
    } catch(_) {}
  }

  // autosave throttle
  let autosaveT = null;
  function autosaveSoon(){
    clearTimeout(autosaveT);
    autosaveT = setTimeout(autosaveNow, 350);
  }
  document.addEventListener("input", (e) => {
    if (e.target && e.target.closest && e.target.closest("#researchForm")) autosaveSoon();
  });
  document.addEventListener("change", (e) => {
    if (e.target && e.target.closest && e.target.closest("#researchForm")) autosaveSoon();
  });

  restoreAutosave();

  // ============================================================
  // Price chart (Stooq -> proxy -> Chart.js -> Word image)
  // + stats + source + annotation
  // ============================================================
  let priceChart = null;
  let priceChartImageBytes = null;

  let equityStats = { currentPrice: null, realisedVolAnn: null, rangeReturn: null };

  const chartStatus = $("chartStatus");
  const fetchChartBtn = $("fetchPriceChart");
  const chartRangeEl = $("chartRange");
  const priceChartCanvas = $("priceChart");
  const targetPriceEl = $("targetPrice");

  function stooqSymbolFromTicker(ticker){
    const t = safeTrim(ticker);
    if (!t) return null;
    if (t.includes(".")) return t.toLowerCase();
    return `${t.toLowerCase()}.us`;
  }

  function computeStartDate(range){
    const now = new Date();
    const d = new Date(now);
    if (range === "6mo") d.setMonth(d.getMonth() - 6);
    else if (range === "1y") d.setFullYear(d.getFullYear() - 1);
    else if (range === "2y") d.setFullYear(d.getFullYear() - 2);
    else if (range === "5y") d.setFullYear(d.getFullYear() - 5);
    else d.setFullYear(d.getFullYear() - 1);
    return d;
  }

  function extractStooqCSV(text){
    const lines = (text||"").split("\n").map(l => l.trim()).filter(Boolean);
    const headerIdx = lines.findIndex(l => l.toLowerCase().startsWith("date,open,high,low,close,volume"));
    if (headerIdx === -1) return null;
    return lines.slice(headerIdx).join("\n");
  }

  async function fetchStooqDaily(symbol){
    const stooqUrl = `http://stooq.com/q/d/l/?s=${encodeURIComponent(symbol)}&i=d`;
    const proxyUrl = `https://r.jina.ai/${stooqUrl}`;
    const res = await fetch(proxyUrl, { cache: "no-store" });
    if (!res.ok) throw new Error("Could not fetch price data (proxy blocked or down).");

    const rawText = await res.text();
    const csvText = extractStooqCSV(rawText) || rawText;

    const lines = csvText.trim().split("\n");
    if (lines.length < 5) throw new Error("Not enough data returned. Check ticker.");

    const rows = lines.slice(1).map(line => line.split(","));
    const out = rows.map(r => ({ date: r[0], close: Number(r[4]) }))
      .filter(x => x.date && Number.isFinite(x.close));
    if (!out.length) throw new Error("No usable price data.");
    return out;
  }

  function renderChart({ labels, values, title }){
    if (!priceChartCanvas || typeof Chart === "undefined") return;

    if (priceChart){
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
        plugins: { legend: { display:false }, tooltip: { intersect:false, mode:"index" } },
        scales: {
          x: { ticks: { maxTicksLimit: 6 } },
          y: { ticks: { maxTicksLimit: 6 } }
        }
      }
    });
  }

  function canvasToPngBytes(canvas){
    const dataUrl = canvas.toDataURL("image/png");
    const b64 = dataUrl.split(",")[1];
    return Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  }

  // stats helpers
  function pct(x){ return `${(x*100).toFixed(1)}%`; }
  function safeNum(v){ const n = Number(v); return Number.isFinite(n) ? n : null; }

  function computeDailyReturns(closes){
    const rets = [];
    for (let i=1;i<closes.length;i++){
      const prev = closes[i-1];
      const cur  = closes[i];
      if (prev > 0 && Number.isFinite(prev) && Number.isFinite(cur)){
        rets.push((cur/prev)-1);
      }
    }
    return rets;
  }
  function stddev(arr){
    if (!arr.length) return null;
    const mean = arr.reduce((a,b)=>a+b,0)/arr.length;
    const v = arr.reduce((a,b)=>a+(b-mean)**2,0)/(arr.length-1 || 1);
    return Math.sqrt(v);
  }
  function setText(id, text){ const el = $(id); if (el) el.textContent = text; }

  function computeUpsideToTarget(currentPrice, targetPrice){
    if (!currentPrice || !targetPrice) return null;
    return (targetPrice/currentPrice) - 1;
  }

  function updateUpsideDisplay(){
    const current = equityStats.currentPrice;
    const target = safeNum(targetPriceEl?.value);
    const up = computeUpsideToTarget(current, target);
    setText("upsideToTarget", up === null ? "—" : pct(up));
  }

  if (targetPriceEl){
    targetPriceEl.addEventListener("input", () => {
      updateUpsideDisplay();
      updateCompletionMeter();
      refreshValidationSummary();
    });
  }

  async function buildPriceChart(){
    try{
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

      renderChart({ labels, values, title: `${tickerVal.toUpperCase()} Close` });

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
      equityStats.rangeReturn = rangeReturn;
      equityStats.realisedVolAnn = realisedVolAnn;

      setText("currentPrice", currentPrice ? currentPrice.toFixed(2) : "—");
      setText("rangeReturn", rangeReturn === null ? "—" : pct(rangeReturn));
      setText("realisedVol", realisedVolAnn === null ? "—" : pct(realisedVolAnn));
      updateUpsideDisplay();

      if (chartStatus) chartStatus.textContent = `✓ Chart ready (${range.toUpperCase()})`;
    } catch(e){
      priceChartImageBytes = null;
      equityStats = { currentPrice:null, realisedVolAnn:null, rangeReturn:null };
      setText("currentPrice","—");
      setText("rangeReturn","—");
      setText("realisedVol","—");
      setText("upsideToTarget","—");
      if (chartStatus) chartStatus.textContent = `✗ ${e.message}`;
    } finally {
      updateCompletionMeter();
      refreshValidationSummary();
    }
  }

  if (fetchChartBtn) fetchChartBtn.addEventListener("click", buildPriceChart);

  // ============================================================
  // Images -> Word
  // ============================================================
  async function addImages(files){
    const imageParagraphs = [];
    for (let i=0;i<files.length;i++){
      const file = files[i];
      try{
        const arrayBuffer = await file.arrayBuffer();
        const fileNameWithoutExt = file.name.replace(/\.[^/.]+$/, "");
        imageParagraphs.push(
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: arrayBuffer,
                transformation: { width: 600, height: 420 }
              })
            ],
            spacing: { before: 180, after: 60 },
            alignment: docx.AlignmentType.CENTER
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: `Figure ${i+1}: ${fileNameWithoutExt}`,
                italics: true,
                size: 18,
                font: "Arial"
              })
            ],
            spacing: { after: 260 },
            alignment: docx.AlignmentType.CENTER
          })
        );
      } catch(err){
        console.error(`Error processing image ${file.name}:`, err);
      }
    }
    return imageParagraphs;
  }

  function linesToParagraphs(text, spacingAfter = 120){
    const lines = (text || "").split("\n");
    return lines.map((line) => {
      if (line.trim() === "") return new docx.Paragraph({ text: "", spacing: { after: spacingAfter } });
      return new docx.Paragraph({ text: line, spacing: { after: spacingAfter } });
    });
  }

  function bulletLines(text){
    const lines = (text || "").split("\n").map(l => l.trim());
    return lines
      .map(line => line.replace(/^[-*•]\s*/, "").trim())
      .filter(Boolean);
  }

  function hyperlinkParagraph(label, url){
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

  // ============================================================
  // (NEW) Compliance watermarking system
  // - Draft: watermark banner + header label
  // - Reviewed: no watermark
  // - Cleared: no watermark + “Cleared” in header
  // Presets can force watermarking:
  // - Internal only: watermark if Draft
  // - Public pack: watermark if Draft
  // - Client-safe: never watermark, but requires Cleared
  // ============================================================
  function shouldWatermark({ status, preset }){
    if (preset === PRESET.CLIENT) return false; // requires Cleared
    if (status === "Draft") return true;
    return false;
  }

  function watermarkBlock(text){
    // sell-side style: prominent banner at top of first page
    return new docx.Paragraph({
      children: [
        new docx.TextRun({
          text,
          bold: true,
          size: 22,
          font: "Arial",
          color: "8A1F1F"
        })
      ],
      spacing: { after: 220 },
      border: {
        top:    { color: "8A1F1F", style: docx.BorderStyle.SINGLE, size: 6, space: 4 },
        bottom: { color: "8A1F1F", style: docx.BorderStyle.SINGLE, size: 6, space: 4 }
      }
    });
  }

  // ============================================================
  // Create Word Document (BlueMatrix-style)
  // - Version included in header + email subject
  // - Status + reviewer in header
  // - Exec summary at top (optional auto + editable)
  // - Data source + chart annotation under chart
  // - Scenario table
  // - Presets: internal/public/client control redactions + routing
  // ============================================================
  async function createDocument(data){
    const {
      noteType, title, topic,
      authorLastName, authorFirstName, authorPhoneSafe,
      coAuthors,
      thesis,
      execSummary,
      analysis, keyTakeaways, content, cordobaView,
      imageFiles, dateTimeString,

      ticker, valuationSummary, keyAssumptions, scenarioNotes, modelFiles, modelLink,
      priceChartImageBytes,
      targetPrice,
      equityStats,
      crgRating,

      workflowStatus,
      reviewedBy,
      preset,
      version,
      changeNote,
      priceDataSource,
      chartAnnotation,
      scenarios
    } = data;

    // ---------- Build blocks ----------
    const takeawayBullets = bulletLines(keyTakeaways).map(line =>
      new docx.Paragraph({ text: line, bullet: { level: 0 }, spacing: { after: 80 } })
    );

    const thesisParas   = linesToParagraphs(thesis, 90);
    const execParas     = linesToParagraphs(execSummary, 90);
    const analysisParas = linesToParagraphs(analysis, 120);
    const contentParas  = linesToParagraphs(content, 120);
    const cordobaParas  = linesToParagraphs(cordobaView, 120);
    const imageParagraphs = await addImages(imageFiles);

    const authorLine = `${authorLastName.toUpperCase()}, ${authorFirstName.toUpperCase()} (${authorPhoneSafe})`;

    // Top info table (sell-side feel)
    const leftTop = [
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: (topic || "").toUpperCase(), size: 18, font: "Arial", color: "555555" })
        ],
        spacing: { after: 90 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: title || "", bold: true, size: 30, font: "Times New Roman", color: "111111" })
        ],
        spacing: { after: 120 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: `${noteType} • ${preset}`, size: 18, font: "Arial", color: "555555" })
        ],
        spacing: { after: 60 }
      })
    ];

    const rightTop = [
      new docx.Paragraph({
        children: [new docx.TextRun({ text: authorLine, bold: true, size: 20, font: "Arial" })],
        alignment: docx.AlignmentType.RIGHT,
        spacing: { after: 70 }
      }),
      ...(coAuthors.length
        ? coAuthors.map(ca => new docx.Paragraph({
            children: [new docx.TextRun({ text: coAuthorLine(ca), size: 18, bold: true, font: "Arial" })],
            alignment: docx.AlignmentType.RIGHT,
            spacing: { after: 50 }
          }))
        : [new docx.Paragraph({ text: "", spacing: { after: 40 } })]
      ),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: `Version: ${version}`, bold: true, size: 18, font: "Arial" })
        ],
        alignment: docx.AlignmentType.RIGHT,
        spacing: { after: 50 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: `Status: ${workflowStatus}`, bold: true, size: 18, font: "Arial" })
        ],
        alignment: docx.AlignmentType.RIGHT,
        spacing: { after: 50 }
      }),
      ...(reviewedBy && reviewedBy !== "N/A"
        ? [new docx.Paragraph({
            children: [new docx.TextRun({ text: `Reviewed by: ${reviewedBy}`, size: 18, font: "Arial" })],
            alignment: docx.AlignmentType.RIGHT,
            spacing: { after: 30 }
          })]
        : []
      )
    ];

    const infoTable = new docx.Table({
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
            new docx.TableCell({ children: leftTop,  width: { size: 65, type: docx.WidthType.PERCENTAGE }, verticalAlign: docx.VerticalAlign.TOP }),
            new docx.TableCell({ children: rightTop, width: { size: 35, type: docx.WidthType.PERCENTAGE }, verticalAlign: docx.VerticalAlign.TOP })
          ]
        })
      ]
    });

    const divider = new docx.Paragraph({
      border: { bottom: { color: "111111", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } },
      spacing: { after: 240 }
    });

    // Exec Summary “one-page” block
    const execSummaryBlock = [
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "EXECUTIVE SUMMARY", bold: true, size: 22, font: "Arial" })],
        spacing: { after: 90 }
      }),
      ...execParas,
      new docx.Paragraph({ spacing: { after: 180 } })
    ];

    // Change note block (optional, but BlueMatrix-style)
    const changeNoteBlock = (changeNote && changeNote !== "N/A")
      ? [
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "CHANGE NOTE", bold: true, size: 20, font: "Arial" })],
            spacing: { before: 40, after: 60 }
          }),
          ...linesToParagraphs(changeNote, 90),
          new docx.Paragraph({ spacing: { after: 160 } })
        ]
      : [];

    const docChildren = [];

    // Watermark banner if Draft + not client-safe
    const watermarkOn = shouldWatermark({ status: workflowStatus, preset });
    if (watermarkOn){
      docChildren.push(watermarkBlock("DRAFT — INTERNAL — NOT FOR DISTRIBUTION"));
    }

    docChildren.push(infoTable, divider);

    // Thesis + Exec Summary at top
    docChildren.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "THESIS", bold: true, size: 22, font: "Arial" })],
        spacing: { after: 90 }
      }),
      ...thesisParas,
      new docx.Paragraph({ spacing: { after: 160 } })
    );

    // Exec summary (always included if present)
    if (safeTrim(execSummary)){
      docChildren.push(...execSummaryBlock);
    }

    // include change note only for Internal preset (sell-side behaviour)
    if (preset === PRESET.INTERNAL){
      docChildren.push(...changeNoteBlock);
    }

    // Equity module (only if Equity Research)
    if (noteType === "Equity Research"){
      const attachedModelNames = (modelFiles && modelFiles.length) ? Array.from(modelFiles).map(f => f.name) : [];

      docChildren.push(
        new docx.Paragraph({
          children: [new docx.TextRun({ text: "EQUITY MODULE", bold: true, size: 22, font: "Arial" })],
          spacing: { before: 80, after: 120 }
        })
      );

      if (safeTrim(ticker)){
        docChildren.push(new docx.Paragraph({
          children: [new docx.TextRun({ text: "Ticker / Company: ", bold: true }), new docx.TextRun({ text: ticker.trim() })],
          spacing: { after: 90 }
        }));
      }
      if (safeTrim(crgRating)){
        docChildren.push(new docx.Paragraph({
          children: [new docx.TextRun({ text: "CRG Rating: ", bold: true }), new docx.TextRun({ text: crgRating.trim() })],
          spacing: { after: 90 }
        }));
      }

      const modelLinkPara = hyperlinkParagraph("Model link:", modelLink);
      if (modelLinkPara) docChildren.push(modelLinkPara);

      // Chart
      if (priceChartImageBytes){
        docChildren.push(
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "PRICE CHART", bold: true, size: 20, font: "Arial" })],
            spacing: { before: 120, after: 80 }
          }),
          new docx.Paragraph({
            children: [new docx.ImageRun({ data: priceChartImageBytes, transformation: { width: 650, height: 300 } })],
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 120 }
          }),
          // Source line (inline citation tracking)
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Source: ", bold: true, size: 18, font: "Arial" }),
              new docx.TextRun({ text: `${priceDataSource}; CRG estimates`, size: 18, font: "Arial" })
            ],
            spacing: { after: 90 }
          })
        );

        // Annotation under chart (optional)
        if (safeTrim(chartAnnotation)){
          docChildren.push(
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: "Note: ", bold: true, size: 18, font: "Arial" }),
                new docx.TextRun({ text: safeTrim(chartAnnotation), size: 18, font: "Arial" })
              ],
              spacing: { after: 140 }
            })
          );
        }
      }

      // Market stats
      if (equityStats && equityStats.currentPrice){
        const tpNum = safeNum(targetPrice);
        const upside = computeUpsideToTarget(equityStats.currentPrice, tpNum);

        docChildren.push(
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "MARKET STATS", bold: true, size: 20, font: "Arial" })],
            spacing: { before: 80, after: 80 }
          }),
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "Current price: ", bold: true }), new docx.TextRun({ text: equityStats.currentPrice.toFixed(2) })],
            spacing: { after: 70 }
          }),
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "Volatility (ann.): ", bold: true }), new docx.TextRun({ text: equityStats.realisedVolAnn == null ? "—" : pct(equityStats.realisedVolAnn) })],
            spacing: { after: 70 }
          }),
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "Return (range): ", bold: true }), new docx.TextRun({ text: equityStats.rangeReturn == null ? "—" : pct(equityStats.rangeReturn) })],
            spacing: { after: 70 }
          })
        );

        if (tpNum){
          docChildren.push(
            new docx.Paragraph({
              children: [new docx.TextRun({ text: "Target price: ", bold: true }), new docx.TextRun({ text: tpNum.toFixed(2) })],
              spacing: { after: 70 }
            }),
            new docx.Paragraph({
              children: [new docx.TextRun({ text: "+/- to target: ", bold: true }), new docx.TextRun({ text: upside == null ? "—" : pct(upside) })],
              spacing: { after: 120 }
            })
          );
        }
      }

      // Scenario table generator (Bear/Base/Bull)
      if (Array.isArray(scenarios) && scenarios.length){
        docChildren.push(
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "SCENARIOS", bold: true, size: 20, font: "Arial" })],
            spacing: { before: 80, after: 80 }
          })
        );

        const table = new docx.Table({
          width: { size: 100, type: docx.WidthType.PERCENTAGE },
          rows: [
            new docx.TableRow({
              children: [
                new docx.TableCell({ children: [new docx.Paragraph({ children: [new docx.TextRun({ text: "Case", bold: true, font: "Arial", size: 18 })] })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ children: [new docx.TextRun({ text: "Price", bold: true, font: "Arial", size: 18 })] })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ children: [new docx.TextRun({ text: "Key assumption", bold: true, font: "Arial", size: 18 })] })] })
              ]
            }),
            ...scenarios.map(r => new docx.TableRow({
              children: [
                new docx.TableCell({ children: [new docx.Paragraph({ text: r.name, spacing: { after: 0 } })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: r.price || "—", spacing: { after: 0 } })] }),
                new docx.TableCell({ children: [new docx.Paragraph({ text: r.key || "—", spacing: { after: 0 } })] })
              ]
            }))
          ]
        });

        docChildren.push(table, new docx.Paragraph({ spacing: { after: 140 } }));
      }

      // Attachments list
      docChildren.push(
        new docx.Paragraph({
          children: [new docx.TextRun({ text: "ATTACHMENTS", bold: true, size: 20, font: "Arial" })],
          spacing: { before: 80, after: 80 }
        })
      );

      if (attachedModelNames.length){
        attachedModelNames.forEach(name => {
          docChildren.push(new docx.Paragraph({ text: name, bullet: { level: 0 }, spacing: { after: 70 } }));
        });
      } else {
        docChildren.push(new docx.Paragraph({ text: "None uploaded", spacing: { after: 100 } }));
      }

      // Optional valuation / assumptions / scenario notes
      if (safeTrim(valuationSummary)){
        docChildren.push(
          new docx.Paragraph({ children: [new docx.TextRun({ text: "VALUATION SUMMARY", bold: true, size: 20, font: "Arial" })], spacing: { before: 120, after: 70 } }),
          ...linesToParagraphs(valuationSummary, 90)
        );
      }
      if (safeTrim(keyAssumptions)){
        docChildren.push(
          new docx.Paragraph({ children: [new docx.TextRun({ text: "KEY ASSUMPTIONS", bold: true, size: 20, font: "Arial" })], spacing: { before: 120, after: 70 } })
        );
        bulletLines(keyAssumptions).forEach(line => {
          docChildren.push(new docx.Paragraph({ text: line, bullet: { level: 0 }, spacing: { after: 70 } }));
        });
      }
      if (safeTrim(scenarioNotes)){
        docChildren.push(
          new docx.Paragraph({ children: [new docx.TextRun({ text: "SENSITIVITIES / NOTES", bold: true, size: 20, font: "Arial" })], spacing: { before: 120, after: 70 } }),
          ...linesToParagraphs(scenarioNotes, 90)
        );
      }

      docChildren.push(new docx.Paragraph({ spacing: { after: 220 } }));
    }

    // Core research
    docChildren.push(
      new docx.Paragraph({ children: [new docx.TextRun({ text: "KEY TAKEAWAYS", bold: true, size: 22, font: "Arial" })], spacing: { after: 120 } }),
      ...takeawayBullets,
      new docx.Paragraph({ spacing: { after: 180 } }),
      new docx.Paragraph({ children: [new docx.TextRun({ text: "ANALYSIS AND COMMENTARY", bold: true, size: 22, font: "Arial" })], spacing: { after: 120 } }),
      ...analysisParas
    );

    // Content + Cordoba view redactions by preset
    // - Public pack: include additional content, but remove internal-only metadata (already removed)
    // - Client-safe: exclude “Additional Content” if it looks like internal appendices (you can change this rule)
    if (safeTrim(content) && preset !== PRESET.CLIENT){
      docChildren.push(
        new docx.Paragraph({ spacing: { after: 160 } }),
        new docx.Paragraph({ children: [new docx.TextRun({ text: "ADDITIONAL CONTENT", bold: true, size: 22, font: "Arial" })], spacing: { after: 120 } }),
        ...contentParas
      );
    }

    if (safeTrim(cordobaView)){
      docChildren.push(
        new docx.Paragraph({ spacing: { after: 160 } }),
        new docx.Paragraph({ children: [new docx.TextRun({ text: "THE CORDOBA VIEW", bold: true, size: 22, font: "Arial" })], spacing: { after: 120 } }),
        ...cordobaParas
      );
    }

    if (imageParagraphs.length){
      docChildren.push(
        new docx.Paragraph({ spacing: { after: 200 } }),
        new docx.Paragraph({ children: [new docx.TextRun({ text: "FIGURES AND CHARTS", bold: true, size: 22, font: "Arial" })], spacing: { after: 120 } }),
        ...imageParagraphs
      );
    }

    // ---------- Header / Footer ----------
    const headerLine = `Cordoba Research Group | ${noteType} | ${formatDateShort(new Date())} | ${version} | ${preset} | ${workflowStatus}`;

    // Put watermark text in header too (Draft)
    const headerChildren = [
      new docx.Paragraph({
        children: [new docx.TextRun({ text: headerLine, size: 16, font: "Arial" })],
        alignment: docx.AlignmentType.RIGHT,
        spacing: { after: 70 },
        border: { bottom: { color: "111111", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } }
      })
    ];

    if (watermarkOn){
      headerChildren.unshift(
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: "DRAFT — INTERNAL — NOT FOR DISTRIBUTION",
              bold: true,
              size: 16,
              font: "Arial",
              color: "8A1F1F"
            })
          ],
          alignment: docx.AlignmentType.LEFT,
          spacing: { after: 40 }
        })
      );
    }

    const footerChildren = [
      new docx.Paragraph({
        border: { top: { color: "111111", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } },
        spacing: { after: 0 }
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: "\t" }),
          new docx.TextRun({
            text: "Internal research documentation — verify figures, tickers, and assumptions before circulation.",
            size: 16,
            font: "Arial",
            italics: true
          }),
          new docx.TextRun({ text: "\t" }),
          new docx.TextRun({
            children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
            size: 16,
            font: "Arial",
            italics: true
          })
        ],
        spacing: { before: 0, after: 0 },
        tabStops: [
          { type: docx.TabStopType.CENTER, position: 5000 },
          { type: docx.TabStopType.RIGHT, position: 10000 }
        ]
      })
    ];

    const doc = new docx.Document({
      styles: {
        default: {
          document: {
            run: { font: "Arial", size: 20, color: "111111" },
            paragraph: { spacing: { after: 120 } }
          }
        }
      },
      sections: [{
        properties: {
          page: {
            margin: { top: 720, right: 720, bottom: 720, left: 720 },
            // keep landscape as you had; if you want sell-side portrait, say and I’ll switch it.
            pageSize: { orientation: docx.PageOrientation.LANDSCAPE, width: 15840, height: 12240 }
          }
        },
        headers: { default: new docx.Header({ children: headerChildren }) },
        footers: { default: new docx.Footer({ children: footerChildren }) },
        children: docChildren
      }]
    });

    return doc;
  }

  // ============================================================
  // Email routing (presets + note type)
  // - subject includes version + change note marker
  // ============================================================
  function ccForNoteType(noteTypeRaw){
    const t = (noteTypeRaw || "").toLowerCase();
    if (t.includes("equity")) return "tommaso@cordobarg.com";
    if (t.includes("macro") || t.includes("market")) return "tim@cordobarg.com";
    if (t.includes("commodity")) return "uhayd@cordobarg.com";
    return "";
  }

  function toForPreset(preset){
    if (preset === PRESET.INTERNAL) return "research@cordobarg.com";
    if (preset === PRESET.PUBLIC)   return "research@cordobarg.com"; // adjust later if you create a public pack mailbox
    if (preset === PRESET.CLIENT)   return "research@cordobarg.com"; // you can change to a distribution list
    return "research@cordobarg.com";
  }

  function buildCrgEmailPayload({ noteType, title, topic, ticker, crgRating, targetPrice, version, changeNote, preset, status }){
    const now = new Date();
    const dateShort = formatDateShort(now);
    const dateLong  = formatDateTime(now);

    const subjectParts = [
      `${noteType || "Research Note"}`,
      `${version}`,
      `${preset}`,
      `${status}`,
      dateShort,
      title ? `— ${title}` : ""
    ].filter(Boolean);

    // BlueMatrix-style change note marker (short)
    if (safeTrim(changeNote)) subjectParts.push("— Chg");

    const subject = subjectParts.join(" ");

    const authorFirstName = safeTrim($("authorFirstName")?.value);
    const authorLastName  = safeTrim($("authorLastName")?.value);
    const authorLine = [authorFirstName, authorLastName].filter(Boolean).join(" ").trim();

    const paragraphs = [];
    paragraphs.push("Hi CRG Research,");
    paragraphs.push("Please find my most recent note attached.");

    const metaLines = [
      `Preset: ${preset}`,
      `Status: ${status}`,
      `Version: ${version}`,
      `Note type: ${noteType || "N/A"}`,
      title ? `Title: ${title}` : null,
      topic ? `Topic: ${topic}` : null,
      ticker ? `Ticker (Stooq): ${ticker}` : null,
      crgRating ? `CRG Rating: ${crgRating}` : null,
      targetPrice ? `Target Price: ${targetPrice}` : null,
      safeTrim(changeNote) ? `Change note: ${safeTrim(changeNote)}` : null,
      `Generated: ${dateLong}`
    ].filter(Boolean);

    paragraphs.push(metaLines.join("\n"));
    paragraphs.push("Best,");
    paragraphs.push(authorLine || "");

    const body = paragraphs.join("\n\n");
    const cc = ccForNoteType(noteType);

    return { subject, body, cc };
  }

  const emailToCrgBtn = $("emailToCrgBtn");
  if (emailToCrgBtn){
    emailToCrgBtn.addEventListener("click", () => {
      const noteType = safeTrim(noteTypeEl?.value) || "Research Note";
      const title    = safeTrim($("title")?.value);
      const topic    = safeTrim($("topic")?.value);
      const ticker   = safeTrim($("ticker")?.value);
      const rating   = safeTrim($("crgRating")?.value);
      const tp       = safeTrim($("targetPrice")?.value);
      const preset   = getPreset();
      const status   = getStatus();
      const version  = versionTextEl?.textContent || "v1.0";
      const changeNote = safeTrim(changeNoteEl?.value);

      const { subject, body, cc } = buildCrgEmailPayload({
        noteType, title, topic, ticker, crgRating: rating, targetPrice: tp,
        version, changeNote, preset, status
      });

      const to = toForPreset(preset);
      const mailto = buildMailto(to, cc, subject, body);
      window.location.href = mailto;
    });
  }

  // ============================================================
  // Main form submission (export)
  // - gates by workflow + preset rules
  // - bumps version
  // ============================================================
  const form = $("researchForm");
  if (form) form.noValidate = true;

  form.addEventListener("submit", async function(e){
    e.preventDefault();

    const button = form.querySelector('button[type="submit"]');
    const messageDiv = $("message");

    button.disabled = true;
    button.classList.add("loading");
    button.textContent = "Generating Document…";
    messageDiv.className = "message";
    messageDiv.textContent = "";

    try{
      if (typeof docx === "undefined") throw new Error("docx library not loaded. Please refresh the page.");
      if (typeof saveAs === "undefined") throw new Error("FileSaver library not loaded. Please refresh the page.");

      // Workflow gating
      if (!workflowRequirementsMet()){
        throw new Error("Export blocked: complete required fields, acknowledgement, and workflow checks (status/preset/reviewer).");
      }

      // bump version at the moment of export (BlueMatrix behaviour)
      const version = bumpVersionOnExport();

      // Gather data
      const noteType = safeTrim(noteTypeEl?.value);
      const title = safeTrim($("title")?.value);
      const topic = safeTrim($("topic")?.value);

      const authorLastName = safeTrim($("authorLastName")?.value);
      const authorFirstName = safeTrim($("authorFirstName")?.value);

      const authorPhone = safeTrim($("authorPhone")?.value);
      const authorPhoneSafe = naIfBlank(authorPhone);

      const thesis = safeTrim(thesisEl?.value);
      const execSummary = safeTrim(execSummaryEl?.value);

      const analysis = safeTrim($("analysis")?.value);
      const keyTakeaways = safeTrim($("keyTakeaways")?.value);
      const content = safeTrim($("content")?.value);
      const cordobaView = safeTrim($("cordobaView")?.value);
      const imageFiles = $("imageUpload")?.files;

      const ticker = safeTrim($("ticker")?.value);
      const valuationSummary = safeTrim($("valuationSummary")?.value);
      const keyAssumptions = safeTrim($("keyAssumptions")?.value);
      const scenarioNotes = safeTrim($("scenarioNotes")?.value);
      const modelFiles = $("modelFiles") ? $("modelFiles").files : null;
      const modelLink = safeTrim($("modelLink")?.value);

      const targetPrice = safeTrim($("targetPrice")?.value);
      const crgRating = safeTrim($("crgRating")?.value);

      const workflowStatus = getStatus();
      const reviewedBy = naIfBlank(safeTrim(reviewedByEl?.value));
      const preset = getPreset();

      const changeNote = naIfBlank(safeTrim(changeNoteEl?.value));
      const priceDataSource = selectedDataSource();
      const chartAnnotation = safeTrim(chartNoteEl?.value);

      const scenarios = scenarioRows();

      const now = new Date();
      const dateTimeString = formatDateTime(now);

      // Coauthors
      const coAuthors = [];
      document.querySelectorAll(".coauthor-row").forEach(entry => {
        const lastName = safeTrim(entry.querySelector(".coauthor-lastname")?.value);
        const firstName = safeTrim(entry.querySelector(".coauthor-firstname")?.value);
        const phone = safeTrim(entry.querySelector(".coauthor-phone")?.value);
        if (lastName || firstName || phone){
          coAuthors.push({ lastName, firstName, phone: naIfBlank(phone) });
        }
      });

      // Sync auto exec summary right before export if toggle is on
      if (execToggleEl?.checked){
        syncExecSummary();
      }

      const doc = await createDocument({
        noteType, title, topic,
        authorLastName, authorFirstName, authorPhoneSafe,
        coAuthors,
        thesis,
        execSummary: safeTrim(execSummaryEl?.value),
        analysis, keyTakeaways, content, cordobaView,
        imageFiles, dateTimeString,

        ticker, valuationSummary, keyAssumptions, scenarioNotes, modelFiles, modelLink,
        priceChartImageBytes,
        targetPrice,
        equityStats,
        crgRating,

        workflowStatus,
        reviewedBy,
        preset,
        version,
        changeNote: (changeNote === "N/A" ? "" : changeNote),
        priceDataSource,
        chartAnnotation,
        scenarios
      });

      const blob = await docx.Packer.toBlob(doc);

      const fileName = [
        (title || "research_note").replace(/[^a-z0-9]/gi,"_").toLowerCase(),
        noteType.replace(/\s+/g,"_").toLowerCase(),
        version.replace(/\./g,"_"),
        preset.replace(/\s+/g,"_").toLowerCase(),
        workflowStatus.toLowerCase()
      ].filter(Boolean).join("_") + ".docx";

      saveAs(blob, fileName);

      messageDiv.className = "message success";
      messageDiv.textContent = `✓ Document "${fileName}" generated successfully!`;

      // autosave after export too
      autosaveNow();

    } catch(err){
      console.error("Error generating document:", err);
      messageDiv.className = "message error";
      messageDiv.textContent = `✗ Error: ${err.message}`;
    } finally {
      button.disabled = false;
      button.classList.remove("loading");
      button.textContent = "Generate Word Document";
    }
  });

  // Initial paint
  updateAttachmentSummary();
  updateCompletionMeter();
  refreshValidationSummary();
  paintVersion();
});
