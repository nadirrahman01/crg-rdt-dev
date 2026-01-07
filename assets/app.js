/* ============================================================
   CRG RDT — app.js (BlueMatrix-style upgrades)
   - Versioning (v1.0 → v1.1 → v2.0), change note
   - Workflow: Status (Draft/Reviewed/Cleared) + Reviewed by
   - Distribution presets (Internal / Public pack / Client-safe)
   - Compliance watermarking (Draft + presets)
   - Auto-generated executive summary (editable)
   - Data source tracking + chart annotation
   - Scenario table generator (Bear/Base/Bull)
   - Saved templates (Macro update / Equity initiation / Event note)
   - Word header + email subject include version/status/change note
   ============================================================ */

console.log("app.js loaded successfully");

window.addEventListener("DOMContentLoaded", () => {
  // ============================================================
  // Utilities
  // ============================================================
  const $ = (id) => document.getElementById(id);

  function digitsOnly(v){ return (v || "").toString().replace(/\D/g, ""); }
  function clamp(n, a, b){ return Math.max(a, Math.min(b, n)); }
  function setText(id, text){ const el = $(id); if (el) el.textContent = text; }

  function wordCount(text){
    const t = (text || "").toString().trim();
    if (!t) return 0;
    return t.split(/\s+/).filter(Boolean).length;
  }

  function naIfBlank(v){
    const s = (v ?? "").toString().trim();
    return s ? s : "N/A";
  }

  function esc(s){
    return (s ?? "").toString().replace(/[&<>"']/g, (c) => ({
      "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
    }[c]));
  }

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

  function nowTime(){
    const d = new Date();
    const hh = String(d.getHours()).padStart(2,"0");
    const mm = String(d.getMinutes()).padStart(2,"0");
    const ss = String(d.getSeconds()).padStart(2,"0");
    return `${hh}:${mm}:${ss}`;
  }

  function formatDateTime(date){
    const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
    const month = months[date.getMonth()];
    const day = date.getDate();
    const year = date.getFullYear();

    let hours = date.getHours();
    const minutes = String(date.getMinutes()).padStart(2, "0");
    const ampm = hours >= 12 ? "PM" : "AM";
    hours = hours % 12 || 12;
    return `${month} ${day}, ${year} ${hours}:${minutes} ${ampm}`;
  }

  function formatDateShort(date){
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  // ============================================================
  // Session timer
  // ============================================================
  const sessionTimerEl = $("sessionTimer");
  const sessionStart = Date.now();
  setInterval(() => {
    if (!sessionTimerEl) return;
    const secs = Math.floor((Date.now() - sessionStart) / 1000);
    const mm = String(Math.floor(secs / 60)).padStart(2, "0");
    const ss = String(secs % 60).padStart(2, "0");
    sessionTimerEl.textContent = `${mm}:${ss}`;
  }, 1000);

  // ============================================================
  // Configuration (routing + reviewers)
  // ============================================================
  const REVIEWERS = [
    "Tim (Macro)",
    "Tommaso (Equity)",
    "Uhayd (Commodities)",
    "Research Lead",
    "Compliance"
  ];

  // Email routing by preset (adjust as needed)
  const ROUTES = {
    internal: { to: "research@cordobarg.com", cc: "" },
    public:   { to: "publishing@cordobarg.com", cc: "" },
    client:   { to: "clients@cordobarg.com", cc: "" }
  };

  // ============================================================
  // Versioning (BlueMatrix-style)
  // - Stored per note key (noteType+title+topic)
  // - Auto-increment minor on each successful export by default
  // - Optional "Bump major" -> major+1, minor=0
  // ============================================================
  const VERSION_KEY = "crg_rdt_versions_v1"; // { [noteKey]: {major, minor} }

  function getNoteKey(){
    const nt = ($("noteType")?.value || "").trim();
    const t  = ($("title")?.value || "").trim();
    const tp = ($("topic")?.value || "").trim();
    // stable enough, avoids exploding keys
    return `${nt}||${t}||${tp}`.toLowerCase();
  }

  function readVersions(){
    try{
      const raw = localStorage.getItem(VERSION_KEY);
      if (!raw) return {};
      return JSON.parse(raw) || {};
    }catch(_){ return {}; }
  }

  function writeVersions(map){
    try{ localStorage.setItem(VERSION_KEY, JSON.stringify(map || {})); }catch(_){}
  }

  function ensureVersionForKey(noteKey){
    const m = readVersions();
    if (!m[noteKey]) m[noteKey] = { major: 1, minor: 0 };
    writeVersions(m);
    return m[noteKey];
  }

  function versionString(v){ return `v${v.major}.${v.minor}`; }

  function getCurrentVersion(){
    const key = getNoteKey();
    const v = ensureVersionForKey(key);
    return { ...v };
  }

  function bumpVersion({ bumpMajor=false } = {}){
    const key = getNoteKey();
    const m = readVersions();
    const v = m[key] || { major: 1, minor: 0 };

    if (bumpMajor){
      v.major = (v.major || 1) + 1;
      v.minor = 0;
    } else {
      v.minor = (v.minor || 0) + 1;
      // Keep minor in sane bounds
      v.minor = clamp(v.minor, 0, 99);
    }

    m[key] = v;
    writeVersions(m);
    return { ...v };
  }

  function setVersionMajorMinor(major, minor){
    const key = getNoteKey();
    const m = readVersions();
    m[key] = { major: clamp(Number(major)||1, 1, 99), minor: clamp(Number(minor)||0, 0, 99) };
    writeVersions(m);
    return { ...m[key] };
  }

  function refreshVersionUI(){
    const v = getCurrentVersion();
    setText("versionDisplay", versionString(v));
  }

  // ============================================================
  // Workflow state: Status + Reviewed by + Distribution preset
  // ============================================================
  const statusEl = $("status");
  const reviewedByEl = $("reviewedBy");
  const distPresetEl = $("distPreset"); // hidden input
  const distLabelEl = $("distPresetLabel");

  function initReviewersDropdown(){
    if (!reviewedByEl) return;
    reviewedByEl.innerHTML = `<option value="">Select…</option>` + REVIEWERS.map(r => `<option value="${esc(r)}">${esc(r)}</option>`).join("");
  }

  function getStatus(){
    return (statusEl?.value || "Draft").trim();
  }

  function getPreset(){
    return (distPresetEl?.value || "internal").trim(); // internal | public | client
  }

  function setPreset(p){
    if (!distPresetEl) return;
    distPresetEl.value = p;

    const nice = p === "internal" ? "Internal only" : (p === "public" ? "Public pack" : "Client-safe");
    if (distLabelEl) distLabelEl.textContent = nice;

    // Preset implies workflow defaults (safe, not forced)
    // - Client-safe should not be Draft by default
    if (p === "client"){
      if (statusEl && statusEl.value === "Draft") statusEl.value = "Reviewed";
    }
    updateWatermarkBadge();
    updateValidationSummary();
    autosaveSoon();
  }

  // Buttons
  $("presetInternal")?.addEventListener("click", () => setPreset("internal"));
  $("presetPublic")?.addEventListener("click", () => setPreset("public"));
  $("presetClient")?.addEventListener("click", () => setPreset("client"));

  function isWatermarked(){
    // Watermark rules:
    // - Draft always watermarked
    // - Internal preset can still watermark if Draft (primary)
    // - Public/Client should not be Draft (but if they are, watermark)
    return getStatus() === "Draft";
  }

  function updateWatermarkBadge(){
    const el = $("watermarkBadge");
    if (!el) return;
    el.textContent = isWatermarked() ? "WATERMARK: ON (Draft)" : "WATERMARK: OFF";
  }

  // ============================================================
  // Templates (saved templates)
  // ============================================================
  const templateEl = $("template");
  function applyTemplate(name){
    const nt = $("noteType");
    const keyTakeaways = $("keyTakeaways");
    const analysis = $("analysis");
    const cordobaView = $("cordobaView");
    const content = $("content");

    if (!name) return;

    if (name === "macro"){
      if (nt) nt.value = "Macro Research";
      if (keyTakeaways && !keyTakeaways.value.trim()){
        keyTakeaways.value = [
          "- Thesis in one line.",
          "- What’s priced in vs what’s mispriced.",
          "- One key data point + source.",
          "- Risk / disconfirming condition."
        ].join("\n");
      }
      if (analysis && !analysis.value.trim()){
        analysis.value = [
          "Thesis:",
          "",
          "Evidence:",
          "-",
          "",
          "What’s priced in / mispriced:",
          "-",
          "",
          "Risks / disconfirmers:",
          "-",
          "",
          "Positioning / implications:",
          "-"
        ].join("\n");
      }
      if (cordobaView && !cordobaView.value.trim()){
        cordobaView.value = [
          "We would position…",
          "",
          "What changes our view:",
          "-",
          "",
          "Key watchpoints:",
          "-"
        ].join("\n");
      }
      if (content && !content.value.trim()){
        content.value = "Appendix (optional): series definitions, back-of-envelope, data notes.";
      }
    }

    if (name === "equity"){
      if (nt) nt.value = "Equity Research";
      if (keyTakeaways && !keyTakeaways.value.trim()){
        keyTakeaways.value = [
          "- One-line investment thesis.",
          "- Upside/downside and key driver.",
          "- What the market is missing.",
          "- 1–2 risks that would break the view."
        ].join("\n");
      }
      if (analysis && !analysis.value.trim()){
        analysis.value = [
          "Thesis:",
          "",
          "Why now:",
          "-",
          "",
          "Valuation:",
          "- Base case summary",
          "",
          "Catalysts:",
          "-",
          "",
          "Risks / disconfirmers:",
          "-",
          "",
          "Positioning:",
          "-"
        ].join("\n");
      }
      if (cordobaView && !cordobaView.value.trim()){
        cordobaView.value = [
          "Our stance:",
          "",
          "Conditions for change:",
          "-",
          "",
          "Implementation (how to express it):",
          "-"
        ].join("\n");
      }
    }

    if (name === "event"){
      if (nt) nt.value = "General Note";
      if (keyTakeaways && !keyTakeaways.value.trim()){
        keyTakeaways.value = [
          "- What happened (1 line).",
          "- Why it matters (1 line).",
          "- What to watch next (1 line).",
          "- Risk / uncertainty."
        ].join("\n");
      }
      if (analysis && !analysis.value.trim()){
        analysis.value = [
          "Event summary:",
          "",
          "Market reaction:",
          "-",
          "",
          "Interpretation:",
          "-",
          "",
          "Second-order effects:",
          "-",
          "",
          "Risks / disconfirmers:",
          "-"
        ].join("\n");
      }
      if (cordobaView && !cordobaView.value.trim()){
        cordobaView.value = [
          "Cordoba view:",
          "",
          "Base case:",
          "-",
          "",
          "If wrong, what would we see:",
          "-"
        ].join("\n");
      }
    }

    toggleEquitySection();
    refreshWordCounts();
    updateCompletionMeter();
    updateValidationSummary();
    refreshVersionUI();
    autosaveSoon();
  }

  templateEl?.addEventListener("change", () => applyTemplate(templateEl.value));

  // ============================================================
  // Executive summary (auto-generate toggle)
  // ============================================================
  const autoExecEl = $("autoExecSummary");
  const execSummaryEl = $("execSummary");

  function firstTakeawayLine(){
    const raw = ($("keyTakeaways")?.value || "").split("\n")
      .map(s => s.trim()).filter(Boolean);
    if (!raw.length) return "";
    const clean = raw[0].replace(/^[-*•]\s*/, "").trim();
    return clean;
  }

  function thesisLine(){
    // Use first non-empty line of analysis, but if it says "Thesis:" then take next line
    const lines = ($("analysis")?.value || "").split("\n").map(s => s.trim());
    const idx = lines.findIndex(l => l.length);
    if (idx === -1) return "";
    if (lines[idx].toLowerCase().startsWith("thesis")){
      const next = lines.slice(idx + 1).find(l => l.length);
      return next || "";
    }
    return lines[idx] || "";
  }

  function buildAutoExecSummary(){
    const noteType = ($("noteType")?.value || "Research Note").trim();
    const title = ($("title")?.value || "").trim();
    const topic = ($("topic")?.value || "").trim();
    const rating = ($("crgRating")?.value || "").trim();
    const target = ($("targetPrice")?.value || "").trim();

    const tl = firstTakeawayLine();
    const th = thesisLine();

    const parts = [];
    parts.push(`Executive Summary`);
    parts.push(``);
    if (title) parts.push(`Title: ${title}`);
    if (topic) parts.push(`Topic: ${topic}`);
    parts.push(`Type: ${noteType}`);
    if (rating) parts.push(`Rating: ${rating}`);
    if (target) parts.push(`Target: ${target}`);
    parts.push(``);
    if (tl) parts.push(`Key takeaway: ${tl}`);
    if (th) parts.push(`Thesis: ${th}`);
    parts.push(``);
    parts.push(`What changes the view:`);
    parts.push(`-`);

    return parts.join("\n");
  }

  function refreshExecSummary(){
    if (!autoExecEl || !execSummaryEl) return;
    if (autoExecEl.checked){
      // Only overwrite if empty OR last generated marker
      const cur = (execSummaryEl.value || "").trim();
      if (!cur || execSummaryEl.dataset.autogen === "1"){
        execSummaryEl.value = buildAutoExecSummary();
        execSummaryEl.dataset.autogen = "1";
      }
    } else {
      execSummaryEl.dataset.autogen = "0";
    }
    refreshWordCounts();
  }

  autoExecEl?.addEventListener("change", () => {
    refreshExecSummary();
    autosaveSoon();
  });

  // If user edits, treat as manual
  execSummaryEl?.addEventListener("input", () => {
    execSummaryEl.dataset.autogen = "0";
    autosaveSoon();
  });

  // ============================================================
  // Word counters (analyst ergonomics)
  // ============================================================
  const wcMap = [
    { field: "execSummary", out: "execWords" },
    { field: "keyTakeaways", out: "takeawaysWords" },
    { field: "analysis", out: "analysisWords" },
    { field: "content", out: "contentWords" },
    { field: "cordobaView", out: "viewWords" }
  ];

  function refreshWordCounts(){
    wcMap.forEach(({ field, out }) => {
      const el = $(field);
      if (!el) return;
      const words = wordCount(el.value || "");
      setText(out, `${words} words`);
    });
  }

  ["input", "keyup", "change"].forEach(evt => {
    document.addEventListener(evt, (e) => {
      if (!e.target) return;
      if (e.target.closest && e.target.closest("#researchForm")) {
        refreshWordCounts();
        // keep exec summary fresh when autogen enabled
        if (["keyTakeaways","analysis","noteType","title","topic","crgRating","targetPrice"].includes(e.target.id || "")){
          refreshExecSummary();
        }
      }
    }, { passive: true });
  });

  refreshWordCounts();

  // ============================================================
  // Phone wiring (primary + coauthors)
  // ============================================================
  const authorPhoneCountryEl = $("authorPhoneCountry");
  const authorPhoneNationalEl = $("authorPhoneNational");
  const authorPhoneHiddenEl = $("authorPhone");

  function syncPrimaryPhone(){
    if (!authorPhoneHiddenEl) return;
    const cc = authorPhoneCountryEl ? authorPhoneCountryEl.value : "";
    const nationalDigits = digitsOnly(authorPhoneNationalEl ? authorPhoneNationalEl.value : "");
    authorPhoneHiddenEl.value = buildInternationalHyphen(cc, nationalDigits);
  }

  function formatPrimaryVisible(){
    if (!authorPhoneNationalEl) return;
    const caret = authorPhoneNationalEl.selectionStart || 0;
    const beforeLen = authorPhoneNationalEl.value.length;

    authorPhoneNationalEl.value = formatNationalLoose(authorPhoneNationalEl.value);

    const afterLen = authorPhoneNationalEl.value.length;
    const delta = afterLen - beforeLen;
    const next = Math.max(0, caret + delta);
    try { authorPhoneNationalEl.setSelectionRange(next, next); } catch(_){}
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

  // Co-authors
  let coAuthorCount = 0;
  const addCoAuthorBtn = $("addCoAuthor");
  const coAuthorsList = $("coAuthorsList");

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

  function wireCoauthorPhone(container){
    const ccEl = container.querySelector(".coauthor-country");
    const nationalEl = container.querySelector(".coauthor-phone-local");
    const hiddenEl = container.querySelector(".coauthor-phone");
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
      try { nationalEl.setSelectionRange(next, next); } catch(_){}
      syncHidden();
    }

    nationalEl?.addEventListener("input", formatVisible);
    nationalEl?.addEventListener("blur", syncHidden);
    ccEl?.addEventListener("change", syncHidden);
    syncHidden();
  }

  addCoAuthorBtn?.addEventListener("click", () => {
    coAuthorCount++;
    const row = document.createElement("div");
    row.className = "coauthor-row";
    row.id = `coauthor-${coAuthorCount}`;

    row.innerHTML = `
      <div class="coauthor-grid">
        <input type="text" class="coauthor-lastname" placeholder="Last name" required>
        <input type="text" class="coauthor-firstname" placeholder="First name" required>

        <div class="phone-row phone-row--compact">
          <select class="phone-country coauthor-country" aria-label="Country code">${countryOptionsHtml}</select>
          <input type="text" class="phone-number coauthor-phone-local" placeholder="Phone (optional)" inputmode="numeric">
        </div>

        <input type="text" class="coauthor-phone" style="display:none;">
        <button type="button" class="btn btn-danger remove-coauthor" data-remove-id="${coAuthorCount}">Remove</button>
      </div>
    `;

    coAuthorsList?.appendChild(row);
    wireCoauthorPhone(row);

    updateCompletionMeter();
    autosaveSoon();
  });

  document.addEventListener("click", (e) => {
    const btn = e.target.closest(".remove-coauthor");
    if (!btn) return;
    const id = btn.getAttribute("data-remove-id");
    const div = document.getElementById(`coauthor-${id}`);
    div?.remove();
    updateCompletionMeter();
    autosaveSoon();
  });

  // ============================================================
  // Equity section toggle
  // ============================================================
  const noteTypeEl = $("noteType");
  const equitySectionEl = $("sec-equity");
  const equityRailLink = $("equityRailLink");
  const crgRatingEl = $("crgRating");

  function toggleEquitySection(){
    if (!noteTypeEl || !equitySectionEl) return;
    const isEquity = noteTypeEl.value === "Equity Research";
    equitySectionEl.style.display = isEquity ? "block" : "none";
    if (equityRailLink) equityRailLink.style.display = isEquity ? "block" : "none";
    if (crgRatingEl) crgRatingEl.required = isEquity;
  }

  noteTypeEl?.addEventListener("change", () => {
    toggleEquitySection();
    refreshExecSummary();
    updateCompletionMeter();
    updateValidationSummary();
    refreshVersionUI();
    autosaveSoon();
  });
  toggleEquitySection();

  // ============================================================
  // Completion + validation summary (left rail)
  // ============================================================
  const completionTextEl = $("completionText");
  const completionBarEl = $("completionBar");
  const completionPctEl = $("completionPct");
  const validationSummaryEl = $("validationSummary");

  function isFilled(el){
    if (!el) return false;
    if (el.type === "file") return el.files && el.files.length > 0;
    if (el.type === "checkbox") return !!el.checked;
    const v = (el.value ?? "").toString().trim();
    return v.length > 0;
  }

  // Core fields (institutional minimum)
  const baseCoreIds = [
    "noteType",
    "topic",
    "title",
    "authorLastName",
    "authorFirstName",
    "keyTakeaways",
    "analysis",
    "cordobaView"
  ];

  // Equity adds required rating
  const equityCoreIds = ["crgRating"];

  function requiredIds(){
    const isEquity = (noteTypeEl?.value === "Equity Research" && equitySectionEl?.style.display !== "none");
    return isEquity ? baseCoreIds.concat(equityCoreIds) : baseCoreIds;
  }

  function listMissing(){
    const missing = [];
    requiredIds().forEach((id) => {
      const el = $(id);
      if (!isFilled(el)) missing.push(id);
    });

    // Workflow rules:
    // - Reviewed/Cleared requires reviewedBy
    const st = getStatus();
    if ((st === "Reviewed" || st === "Cleared") && !isFilled(reviewedByEl)) missing.push("reviewedBy");

    return missing;
  }

  function fieldLabelFor(id){
    const label = document.querySelector(`label[for="${id}"]`);
    if (label) return label.textContent.replace(/\s+\(Optional\)$/i, "").trim();
    const map = {
      reviewedBy: "Reviewed by",
      status: "Status"
    };
    return map[id] || id;
  }

  function updateCompletionMeter(){
    const ids = requiredIds().slice();
    // include reviewer in completion count when status requires it
    const st = getStatus();
    if (st === "Reviewed" || st === "Cleared") ids.push("reviewedBy");

    let done = 0;
    ids.forEach((id) => { if (isFilled($(id))) done++; });

    const total = ids.length;
    const pct = total ? Math.round((done / total) * 100) : 0;

    if (completionTextEl) completionTextEl.textContent = `${done} / ${total} core fields`;
    if (completionPctEl) completionPctEl.textContent = `${pct}%`;
    if (completionBarEl) completionBarEl.style.width = `${pct}%`;
  }

  function updateValidationSummary(){
    const missing = listMissing();
    if (!validationSummaryEl) return;
    if (!missing.length){
      validationSummaryEl.textContent = "All required fields complete.";
      return;
    }
    const nice = missing.slice(0, 7).map(fieldLabelFor);
    const rest = missing.length > 7 ? ` +${missing.length - 7} more` : "";
    validationSummaryEl.textContent = `Missing: ${nice.join(" • ")}${rest}`;
  }

  // update on any form change
  ["input", "change", "keyup"].forEach(evt => {
    document.addEventListener(evt, (e) => {
      if (!e.target) return;
      if (e.target.closest && e.target.closest("#researchForm")){
        updateCompletionMeter();
        updateValidationSummary();
        updateWatermarkBadge();
        refreshVersionUI();
        autosaveSoon();
      }
    }, { passive: true });
  });

  // Jump to first missing
  $("jumpFirstMissing")?.addEventListener("click", () => {
    const missing = listMissing();
    if (!missing.length) return;
    const firstId = missing[0];
    const el = $(firstId);
    if (el){
      el.scrollIntoView({ behavior: "smooth", block: "center" });
      setTimeout(() => { try { el.focus(); } catch(_){} }, 250);
    }
  });

  // ============================================================
  // Autosave (localStorage)
  // ============================================================
  const AUTOSAVE_KEY = "crg_rdt_autosave_v2";
  const autosaveStatusEl = $("autosaveStatus");
  $("clearAutosave")?.addEventListener("click", () => {
    const ok = confirm("Clear autosave? This will remove any saved draft from this browser.");
    if (!ok) return;
    try{ localStorage.removeItem(AUTOSAVE_KEY); } catch(_){}
    if (autosaveStatusEl) autosaveStatusEl.textContent = "cleared";
  });

  function serializeCoAuthors(){
    const out = [];
    document.querySelectorAll(".coauthor-row").forEach(row => {
      const ln = row.querySelector(".coauthor-lastname")?.value || "";
      const fn = row.querySelector(".coauthor-firstname")?.value || "";
      const cc = row.querySelector(".coauthor-country")?.value || "44";
      const local = row.querySelector(".coauthor-phone-local")?.value || "";
      const hidden = row.querySelector(".coauthor-phone")?.value || buildInternationalHyphen(cc, digitsOnly(local));
      if ((ln || "").trim() || (fn || "").trim() || digitsOnly(local)){
        out.push({ lastName: ln, firstName: fn, cc, phoneLocal: local, phoneHidden: hidden });
      }
    });
    return out;
  }

  function serializeScenario(){
    return {
      bear: { price: ($("scBearPrice")?.value || "").trim(), assumption: ($("scBearAssump")?.value || "").trim() },
      base: { price: ($("scBasePrice")?.value || "").trim(), assumption: ($("scBaseAssump")?.value || "").trim() },
      bull: { price: ($("scBullPrice")?.value || "").trim(), assumption: ($("scBullAssump")?.value || "").trim() }
    };
  }

  function saveAutosave(){
    try{
      const payload = {
        savedAt: Date.now(),

        template: $("template")?.value || "",
        noteType: noteTypeEl?.value || "",
        title: $("title")?.value || "",
        topic: $("topic")?.value || "",

        status: $("status")?.value || "Draft",
        reviewedBy: $("reviewedBy")?.value || "",
        distPreset: getPreset(),

        bumpMajor: $("bumpMajor")?.checked || false,
        changeNote: $("changeNote")?.value || "",

        autoExecSummary: $("autoExecSummary")?.checked || false,
        execSummary: $("execSummary")?.value || "",

        authorLastName: $("authorLastName")?.value || "",
        authorFirstName: $("authorFirstName")?.value || "",
        authorPhoneCountry: authorPhoneCountryEl?.value || "44",
        authorPhoneNational: authorPhoneNationalEl?.value || "",
        authorPhoneHidden: authorPhoneHiddenEl?.value || "",

        keyTakeaways: $("keyTakeaways")?.value || "",
        analysis: $("analysis")?.value || "",
        content: $("content")?.value || "",
        cordobaView: $("cordobaView")?.value || "",

        // equity
        ticker: $("ticker")?.value || "",
        crgRating: $("crgRating")?.value || "",
        targetPrice: $("targetPrice")?.value || "",
        valuationSummary: $("valuationSummary")?.value || "",
        keyAssumptions: $("keyAssumptions")?.value || "",
        scenarioNotes: $("scenarioNotes")?.value || "",
        modelLink: $("modelLink")?.value || "",

        // data source + annotation
        chartDataSource: $("chartDataSource")?.value || "",
        chartDataSourceNote: $("chartDataSourceNote")?.value || "",
        chartAnnotation: $("chartAnnotation")?.value || "",

        scenarioTable: serializeScenario(),

        // coauthors
        coAuthors: serializeCoAuthors()
      };

      localStorage.setItem(AUTOSAVE_KEY, JSON.stringify(payload));
      if (autosaveStatusEl) autosaveStatusEl.textContent = `saved ${nowTime()}`;
    } catch(e){
      if (autosaveStatusEl) autosaveStatusEl.textContent = "autosave failed";
    }
  }

  let autosaveTimer = null;
  function autosaveSoon(){
    if (autosaveTimer) clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(saveAutosave, 650);
  }

  function restoreAutosave(){
    let raw = null;
    try{ raw = localStorage.getItem(AUTOSAVE_KEY); } catch(_){}
    if (!raw) return;

    let data = null;
    try{ data = JSON.parse(raw); } catch(_){ return; }
    if (!data) return;

    if ($("template")) $("template").value = data.template || "";

    if (noteTypeEl) noteTypeEl.value = data.noteType || "";
    if ($("title")) $("title").value = data.title || "";
    if ($("topic")) $("topic").value = data.topic || "";

    if ($("status")) $("status").value = data.status || "Draft";
    if ($("reviewedBy")) $("reviewedBy").value = data.reviewedBy || "";
    if ($("distPreset")) $("distPreset").value = data.distPreset || "internal";
    if ($("distPresetLabel")){
      $("distPresetLabel").textContent =
        (data.distPreset === "public") ? "Public pack" :
        (data.distPreset === "client") ? "Client-safe" :
        "Internal only";
    }

    if ($("bumpMajor")) $("bumpMajor").checked = !!data.bumpMajor;
    if ($("changeNote")) $("changeNote").value = data.changeNote || "";

    if ($("autoExecSummary")) $("autoExecSummary").checked = !!data.autoExecSummary;
    if ($("execSummary")) $("execSummary").value = data.execSummary || "";

    if ($("authorLastName")) $("authorLastName").value = data.authorLastName || "";
    if ($("authorFirstName")) $("authorFirstName").value = data.authorFirstName || "";

    if (authorPhoneCountryEl) authorPhoneCountryEl.value = data.authorPhoneCountry || "44";
    if (authorPhoneNationalEl) authorPhoneNationalEl.value = data.authorPhoneNational || "";
    syncPrimaryPhone();

    if ($("keyTakeaways")) $("keyTakeaways").value = data.keyTakeaways || "";
    if ($("analysis")) $("analysis").value = data.analysis || "";
    if ($("content")) $("content").value = data.content || "";
    if ($("cordobaView")) $("cordobaView").value = data.cordobaView || "";

    // equity
    if ($("ticker")) $("ticker").value = data.ticker || "";
    if ($("crgRating")) $("crgRating").value = data.crgRating || "";
    if ($("targetPrice")) $("targetPrice").value = data.targetPrice || "";
    if ($("valuationSummary")) $("valuationSummary").value = data.valuationSummary || "";
    if ($("keyAssumptions")) $("keyAssumptions").value = data.keyAssumptions || "";
    if ($("scenarioNotes")) $("scenarioNotes").value = data.scenarioNotes || "";
    if ($("modelLink")) $("modelLink").value = data.modelLink || "";

    // data source + annotation
    if ($("chartDataSource")) $("chartDataSource").value = data.chartDataSource || "";
    if ($("chartDataSourceNote")) $("chartDataSourceNote").value = data.chartDataSourceNote || "";
    if ($("chartAnnotation")) $("chartAnnotation").value = data.chartAnnotation || "";

    if (data.scenarioTable){
      if ($("scBearPrice")) $("scBearPrice").value = data.scenarioTable.bear?.price || "";
      if ($("scBearAssump")) $("scBearAssump").value = data.scenarioTable.bear?.assumption || "";
      if ($("scBasePrice")) $("scBasePrice").value = data.scenarioTable.base?.price || "";
      if ($("scBaseAssump")) $("scBaseAssump").value = data.scenarioTable.base?.assumption || "";
      if ($("scBullPrice")) $("scBullPrice").value = data.scenarioTable.bull?.price || "";
      if ($("scBullAssump")) $("scBullAssump").value = data.scenarioTable.bull?.assumption || "";
    }

    // coauthors rebuild
    if (coAuthorsList) coAuthorsList.innerHTML = "";
    coAuthorCount = 0;
    (data.coAuthors || []).forEach(ca => {
      coAuthorCount++;
      const row = document.createElement("div");
      row.className = "coauthor-row";
      row.id = `coauthor-${coAuthorCount}`;
      row.innerHTML = `
        <div class="coauthor-grid">
          <input type="text" class="coauthor-lastname" placeholder="Last name" required value="${esc(ca.lastName||"")}">
          <input type="text" class="coauthor-firstname" placeholder="First name" required value="${esc(ca.firstName||"")}">
          <div class="phone-row phone-row--compact">
            <select class="phone-country coauthor-country" aria-label="Country code">${countryOptionsHtml}</select>
            <input type="text" class="phone-number coauthor-phone-local" placeholder="Phone (optional)" inputmode="numeric" value="${esc(ca.phoneLocal||"")}">
          </div>
          <input type="text" class="coauthor-phone" style="display:none;" value="${esc(ca.phoneHidden||"")}">
          <button type="button" class="btn btn-danger remove-coauthor" data-remove-id="${coAuthorCount}">Remove</button>
        </div>
      `;
      coAuthorsList.appendChild(row);
      const ccEl = row.querySelector(".coauthor-country");
      if (ccEl) ccEl.value = ca.cc ?? "44";
      wireCoauthorPhone(row);
    });

    toggleEquitySection();
    refreshWordCounts();
    updateCompletionMeter();
    updateValidationSummary();
    refreshExecSummary();
    refreshVersionUI();
    updateWatermarkBadge();

    if (autosaveStatusEl){
      const ts = data.savedAt ? new Date(data.savedAt) : null;
      autosaveStatusEl.textContent = ts ? `restored ${String(ts.getHours()).padStart(2,"0")}:${String(ts.getMinutes()).padStart(2,"0")}` : "restored";
    }
  }

  // ============================================================
  // Attachment summary (model files)
  // ============================================================
  const modelFilesEl = $("modelFiles");
  const attachSummaryHeadEl = $("attachmentSummaryHead");
  const attachSummaryListEl = $("attachmentSummaryList");

  function updateAttachmentSummary(){
    if (!modelFilesEl || !attachSummaryHeadEl || !attachSummaryListEl) return;
    const files = Array.from(modelFilesEl.files || []);
    if (!files.length){
      attachSummaryHeadEl.textContent = "No files selected";
      attachSummaryListEl.style.display = "none";
      attachSummaryListEl.innerHTML = "";
      return;
    }
    attachSummaryHeadEl.textContent = `${files.length} file${files.length === 1 ? "" : "s"} selected`;
    attachSummaryListEl.style.display = "block";
    attachSummaryListEl.innerHTML = files.map(f => `<div class="attach-file">${esc(f.name)}</div>`).join("");
  }

  modelFilesEl?.addEventListener("change", () => {
    updateAttachmentSummary();
    updateCompletionMeter();
    updateValidationSummary();
    autosaveSoon();
  });
  updateAttachmentSummary();

  // ============================================================
  // Reset (page state)
  // ============================================================
  const resetBtn = $("resetFormBtn");
  const formEl = $("researchForm");

  let priceChart = null;
  let priceChartImageBytes = null;
  let equityStats = { currentPrice: null, realisedVolAnn: null, rangeReturn: null };

  function clearChartUI(){
    setText("currentPrice", "—");
    setText("realisedVol", "—");
    setText("rangeReturn", "—");
    setText("upsideToTarget", "—");
    setText("chartStatus", "");
    if (priceChart){
      try { priceChart.destroy(); } catch(_){}
      priceChart = null;
    }
    priceChartImageBytes = null;
    equityStats = { currentPrice: null, realisedVolAnn: null, rangeReturn: null };
  }

  resetBtn?.addEventListener("click", () => {
    const ok = confirm("Reset the form? This will clear all fields on this page.");
    if (!ok) return;

    formEl?.reset();
    if (coAuthorsList) coAuthorsList.innerHTML = "";
    coAuthorCount = 0;

    if (modelFilesEl) modelFilesEl.value = "";
    updateAttachmentSummary();
    clearChartUI();

    // Defaults
    if ($("status")) $("status").value = "Draft";
    if ($("distPreset")) $("distPreset").value = "internal";
    if ($("distPresetLabel")) $("distPresetLabel").textContent = "Internal only";
    if ($("reviewedBy")) $("reviewedBy").value = "";
    if ($("bumpMajor")) $("bumpMajor").checked = false;
    if ($("changeNote")) $("changeNote").value = "";
    if ($("autoExecSummary")) $("autoExecSummary").checked = false;
    if ($("execSummary")) $("execSummary").value = "";

    syncPrimaryPhone();
    toggleEquitySection();
    refreshWordCounts();

    const messageDiv = $("message");
    if (messageDiv){
      messageDiv.className = "message";
      messageDiv.textContent = "";
    }

    updateCompletionMeter();
    updateValidationSummary();
    refreshVersionUI();
    updateWatermarkBadge();

    autosaveSoon();
  });

  // ============================================================
  // Email (subject includes version + status + change note)
  // ============================================================
  function buildMailto(to, cc, subject, body){
    const crlfBody = (body || "").replace(/\n/g, "\r\n");
    const parts = [];
    if (cc) parts.push(`cc=${encodeURIComponent(cc)}`);
    parts.push(`subject=${encodeURIComponent(subject || "")}`);
    parts.push(`body=${encodeURIComponent(crlfBody)}`);
    return `mailto:${encodeURIComponent(to)}?${parts.join("&")}`;
  }

  function ccForNoteType(noteTypeRaw){
    const t = (noteTypeRaw || "").toLowerCase();
    if (t.includes("equity")) return "tommaso@cordobarg.com";
    if (t.includes("macro") || t.includes("market")) return "tim@cordobarg.com";
    if (t.includes("commodity")) return "uhayd@cordobarg.com";
    return "";
  }

  function buildCrgEmailPayload({ versionStr }){
    const noteType = (noteTypeEl?.value || "Research Note").trim();
    const title = ($("title")?.value || "").trim();
    const topic = ($("topic")?.value || "").trim();
    const st = getStatus();
    const changeNote = ($("changeNote")?.value || "").trim();
    const preset = getPreset();

    const authorFirstName = ($("authorFirstName")?.value || "").trim();
    const authorLastName = ($("authorLastName")?.value || "").trim();

    const ticker = ($("ticker")?.value || "").trim();
    const crgRating = ($("crgRating")?.value || "").trim();
    const targetPrice = ($("targetPrice")?.value || "").trim();

    const now = new Date();
    const dateShort = formatDateShort(now);
    const dateLong = formatDateTime(now);

    const subjectParts = [
      "CRG",
      noteType || "Research Note",
      versionStr,
      `[${st}]`,
      dateShort,
      title ? `— ${title}` : ""
    ].filter(Boolean);

    // include change note in subject line if present (BlueMatrix feel)
    if (changeNote) subjectParts.push(`— ${changeNote}`);

    const subject = subjectParts.join(" ");

    const authorLine = [authorFirstName, authorLastName].filter(Boolean).join(" ").trim();

    const paragraphs = [];
    paragraphs.push("Hi CRG Research,");
    paragraphs.push("Please find the note attached.");

    const metaLines = [
      `Note type: ${noteType || "N/A"}`,
      `Version: ${versionStr}`,
      `Status: ${st}`,
      `Distribution preset: ${preset}`,
      changeNote ? `Change note: ${changeNote}` : null,
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

    // Route by preset, with note-type cc as “secondary”
    const route = ROUTES[preset] || ROUTES.internal;
    const cc2 = ccForNoteType(noteType);
    const cc = [route.cc, cc2].filter(Boolean).join(",");

    return { subject, body: paragraphs.join("\n\n"), to: route.to, cc };
  }

  $("emailToCrgBtn")?.addEventListener("click", () => {
    const v = getCurrentVersion();
    const payload = buildCrgEmailPayload({ versionStr: versionString(v) });
    window.location.href = buildMailto(payload.to, payload.cc, payload.subject, payload.body);
  });

  // ============================================================
  // Price chart (Stooq -> Chart.js -> Word image)
  // (Stooq has no CORS. Use r.jina.ai proxy.)
  // ============================================================
  const chartStatus = $("chartStatus");
  const fetchChartBtn = $("fetchPriceChart");
  const chartRangeEl = $("chartRange");
  const priceChartCanvas = $("priceChart");
  const targetPriceEl = $("targetPrice");

  function stooqSymbolFromTicker(ticker){
    const t = (ticker || "").trim();
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
    const lines = (text || "").split("\n").map(l => l.trim()).filter(Boolean);
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

  function canvasToPngBytes(canvas){
    const dataUrl = canvas.toDataURL("image/png");
    const b64 = dataUrl.split(",")[1];
    return Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  }

  function pct(x){ return `${(x * 100).toFixed(1)}%`; }
  function safeNum(v){
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  }

  function computeDailyReturns(closes){
    const rets = [];
    for (let i = 1; i < closes.length; i++){
      const prev = closes[i - 1];
      const cur = closes[i];
      if (prev > 0 && Number.isFinite(prev) && Number.isFinite(cur)){
        rets.push((cur / prev) - 1);
      }
    }
    return rets;
  }

  function stddev(arr){
    if (!arr.length) return null;
    const mean = arr.reduce((a,b) => a + b, 0) / arr.length;
    const v = arr.reduce((a,b) => a + (b - mean) ** 2, 0) / (arr.length - 1 || 1);
    return Math.sqrt(v);
  }

  function computeUpsideToTarget(currentPrice, targetPrice){
    if (!currentPrice || !targetPrice) return null;
    return (targetPrice / currentPrice) - 1;
  }

  function updateUpsideDisplay(){
    const current = equityStats.currentPrice;
    const target = safeNum(targetPriceEl?.value);
    const up = computeUpsideToTarget(current, target);
    setText("upsideToTarget", up === null ? "—" : pct(up));
  }

  targetPriceEl?.addEventListener("input", () => {
    updateUpsideDisplay();
    autosaveSoon();
  });

  async function buildPriceChart(){
    try{
      const tickerVal = ($("ticker")?.value || "").trim();
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
      equityStats = { currentPrice: null, realisedVolAnn: null, rangeReturn: null };
      setText("currentPrice", "—");
      setText("rangeReturn", "—");
      setText("realisedVol", "—");
      setText("upsideToTarget", "—");
      if (chartStatus) chartStatus.textContent = `✗ ${e.message}`;
    } finally {
      updateCompletionMeter();
      updateValidationSummary();
    }
  }

  fetchChartBtn?.addEventListener("click", buildPriceChart);

  // ============================================================
  // Data source tracking + chart annotation
  // ============================================================
  const chartDataSourceEl = $("chartDataSource");
  const chartDataSourceNoteEl = $("chartDataSourceNote");
  const chartAnnotationEl = $("chartAnnotation");

  // ============================================================
  // Word: helpers
  // ============================================================
  async function addImages(files){
    const imageParagraphs = [];
    for (let i = 0; i < files.length; i++){
      const file = files[i];
      try{
        const arrayBuffer = await file.arrayBuffer();
        const fileNameWithoutExt = file.name.replace(/\.[^/.]+$/, "");

        imageParagraphs.push(
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: arrayBuffer,
                transformation: { width: 580, height: 420 }
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
                font: "Times New Roman"
              })
            ],
            spacing: { after: 240 },
            alignment: docx.AlignmentType.CENTER
          })
        );
      } catch (error){
        console.error(`Error processing image ${file.name}:`, error);
      }
    }
    return imageParagraphs;
  }

  function linesToParagraphs(text, spacingAfter = 120){
    const lines = (text || "").split("\n");
    return lines.map((line) => {
      if (line.trim() === "") return new docx.Paragraph({ text: "", spacing: { after: spacingAfter } });
      return new docx.Paragraph({
        children: [new docx.TextRun({ text: line, font: "Times New Roman", size: 22 })],
        spacing: { after: spacingAfter }
      });
    });
  }

  function bulletLines(text, spacingAfter=90){
    const lines = (text || "").split("\n").map(s => s.trim());
    const out = [];
    lines.forEach(line => {
      if (!line) return;
      const clean = line.replace(/^[-*•]\s*/, "").trim();
      if (!clean) return;
      out.push(new docx.Paragraph({
        children: [new docx.TextRun({ text: clean, font: "Times New Roman", size: 22 })],
        bullet: { level: 0 },
        spacing: { after: spacingAfter }
      }));
    });
    if (!out.length) out.push(new docx.Paragraph({ text: "—", spacing: { after: 120 } }));
    return out;
  }

  function sectionHeading(text){
    return new docx.Paragraph({
      children: [new docx.TextRun({ text, bold: true, font: "Times New Roman", size: 26 })],
      spacing: { before: 200, after: 140 }
    });
  }

  function smallLabel(text){
    return new docx.TextRun({ text, bold: true, font: "Times New Roman", size: 20 });
  }

  function smallValue(text){
    return new docx.TextRun({ text, font: "Times New Roman", size: 20 });
  }

  function buildScenarioTable(scenario){
    const rows = [];
    const mkRow = (name, obj) => new docx.TableRow({
      children: [
        new docx.TableCell({ children: [new docx.Paragraph({ children: [smallValue(name)] })] }),
        new docx.TableCell({ children: [new docx.Paragraph({ children: [smallValue(naIfBlank(obj.price))] })] }),
        new docx.TableCell({ children: [new docx.Paragraph({ children: [smallValue(naIfBlank(obj.assumption))] })] })
      ]
    });

    rows.push(new docx.TableRow({
      children: [
        new docx.TableCell({ children: [new docx.Paragraph({ children: [smallLabel("Scenario")] })] }),
        new docx.TableCell({ children: [new docx.Paragraph({ children: [smallLabel("Price")] })] }),
        new docx.TableCell({ children: [new docx.Paragraph({ children: [smallLabel("Key assumption")] })] })
      ]
    }));

    rows.push(mkRow("Bear", scenario.bear || {}));
    rows.push(mkRow("Base", scenario.base || {}));
    rows.push(mkRow("Bull", scenario.bull || {}));

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows
    });
  }

  // ============================================================
  // Create Word Document (BlueMatrix-style header + watermarking)
  // ============================================================
  async function createDocument(data){
    const {
      noteType, title, topic,
      status, reviewedBy, distPreset,
      versionStr, changeNote,
      execSummary,
      authorLastName, authorFirstName, authorPhoneSafe,
      coAuthors,
      analysis, keyTakeaways, content, cordobaView,
      imageFiles, dateTimeString,

      ticker, valuationSummary, keyAssumptions, scenarioNotes, modelFiles, modelLink,
      priceChartImageBytes,

      targetPrice,
      equityStats,
      crgRating,

      chartDataSource, chartDataSourceNote, chartAnnotation,
      scenarioTable
    } = data;

    const authorLine = `${authorLastName.toUpperCase()}, ${authorFirstName.toUpperCase()} (${authorPhoneSafe})`;

    // Coauthors block (right aligned)
    const coAuthorParas = (coAuthors && coAuthors.length)
      ? coAuthors.map(ca => new docx.Paragraph({
          children: [new docx.TextRun({ text: `${(ca.lastName||"").toUpperCase()}, ${(ca.firstName||"").toUpperCase()} (${naIfBlank(ca.phone)})`, bold: true, font:"Times New Roman", size: 22 })],
          alignment: docx.AlignmentType.RIGHT,
          spacing: { after: 60 }
        }))
      : [new docx.Paragraph({ text: "", spacing: { after: 40 } })];

    const metaTable = new docx.Table({
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
              width: { size: 66, type: docx.WidthType.PERCENTAGE },
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: (title || "").trim(), bold: true, font: "Times New Roman", size: 34 })
                  ],
                  spacing: { after: 80 }
                }),
                new docx.Paragraph({
                  children: [
                    smallLabel("TOPIC: "),
                    smallValue((topic || "").trim() || "—")
                  ],
                  spacing: { after: 40 }
                }),
                new docx.Paragraph({
                  children: [
                    smallLabel("TYPE: "),
                    smallValue(noteType || "—"),
                    new docx.TextRun({ text: "   " }),
                    smallLabel("VERSION: "),
                    smallValue(versionStr),
                    new docx.TextRun({ text: "   " }),
                    smallLabel("STATUS: "),
                    smallValue(status)
                  ],
                  spacing: { after: 40 }
                }),
                new docx.Paragraph({
                  children: [
                    smallLabel("DISTRIBUTION: "),
                    smallValue(distPreset)
                  ],
                  spacing: { after: 20 }
                }),
                changeNote
                  ? new docx.Paragraph({
                      children: [smallLabel("CHANGE NOTE: "), smallValue(changeNote)],
                      spacing: { after: 20 }
                    })
                  : new docx.Paragraph({ text: "", spacing: { after: 20 } })
              ]
            }),
            new docx.TableCell({
              width: { size: 34, type: docx.WidthType.PERCENTAGE },
              children: [
                new docx.Paragraph({
                  children: [new docx.TextRun({ text: authorLine, bold: true, font:"Times New Roman", size: 24 })],
                  alignment: docx.AlignmentType.RIGHT,
                  spacing: { after: 60 }
                }),
                ...coAuthorParas,
                reviewedBy
                  ? new docx.Paragraph({
                      children: [smallLabel("Reviewed by: "), smallValue(reviewedBy)],
                      alignment: docx.AlignmentType.RIGHT,
                      spacing: { after: 40 }
                    })
                  : new docx.Paragraph({ text: "", spacing: { after: 40 } })
              ]
            })
          ]
        })
      ]
    });

    const divider = new docx.Paragraph({
      border: { bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } },
      spacing: { after: 220 }
    });

    // Watermark line (header-level, repeated)
    const watermarkOn = (status === "Draft");
    const watermarkLine = watermarkOn
      ? new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: "DRAFT — INTERNAL — NOT FOR DISTRIBUTION",
              bold: true,
              font: "Times New Roman",
              size: 18,
              color: "7A7A7A"
            })
          ],
          alignment: docx.AlignmentType.CENTER,
          spacing: { after: 80 }
        })
      : null;

    const headerLine = new docx.Paragraph({
      children: [
        new docx.TextRun({
          text: `Cordoba Research Group | ${noteType} | ${versionStr} | ${status} | ${formatDateShort(new Date())}`,
          font: "Times New Roman",
          size: 16
        })
      ],
      alignment: docx.AlignmentType.RIGHT,
      spacing: { after: 80 },
      border: { bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } }
    });

    // Body sections
    const body = [];

    body.push(metaTable, divider);

    // Executive summary (one-page feel)
    if ((execSummary || "").trim()){
      body.push(sectionHeading("Executive Summary"));
      body.push(...linesToParagraphs(execSummary, 110));
      body.push(new docx.Paragraph({ spacing: { after: 120 } }));
      body.push(new docx.Paragraph({
        border: { bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } },
        spacing: { after: 180 }
      }));
    }

    // Equity module
    if (noteType === "Equity Research"){
      body.push(sectionHeading("Equity Addendum"));

      if ((ticker || "").trim()){
        body.push(new docx.Paragraph({
          children: [smallLabel("Ticker / Company: "), smallValue(ticker.trim())],
          spacing: { after: 70 }
        }));
      }

      if ((crgRating || "").trim()){
        body.push(new docx.Paragraph({
          children: [smallLabel("CRG Rating: "), smallValue(crgRating.trim())],
          spacing: { after: 70 }
        }));
      }

      if ((targetPrice || "").trim()){
        body.push(new docx.Paragraph({
          children: [smallLabel("Target price: "), smallValue(targetPrice.trim())],
          spacing: { after: 70 }
        }));
      }

      if ((modelLink || "").trim()){
        body.push(new docx.Paragraph({
          children: [
            smallLabel("Model link: "),
            new docx.ExternalHyperlink({
              children: [new docx.TextRun({ text: modelLink.trim(), style: "Hyperlink" })],
              link: modelLink.trim()
            })
          ],
          spacing: { after: 90 }
        }));
      }

      // Price chart + source + annotation
      if (priceChartImageBytes){
        body.push(new docx.Paragraph({
          children: [new docx.TextRun({ text: "Price chart", bold: true, font:"Times New Roman", size: 24 })],
          spacing: { before: 60, after: 80 }
        }));
        body.push(new docx.Paragraph({
          children: [new docx.ImageRun({ data: priceChartImageBytes, transformation: { width: 610, height: 260 } })],
          alignment: docx.AlignmentType.CENTER,
          spacing: { after: 100 }
        }));

        const sourceLine = [
          "Source: ",
          (chartDataSource || "N/A"),
          chartDataSourceNote ? `; ${chartDataSourceNote}` : ""
        ].join("");

        body.push(new docx.Paragraph({
          children: [new docx.TextRun({ text: sourceLine, italics: true, font:"Times New Roman", size: 18 })],
          alignment: docx.AlignmentType.CENTER,
          spacing: { after: 90 }
        }));

        if ((chartAnnotation || "").trim()){
          body.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: `Note: ${chartAnnotation.trim()}`, italics: true, font:"Times New Roman", size: 18 })],
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 140 }
          }));
        }
      }

      // Market stats
      if (equityStats && equityStats.currentPrice){
        const tp = (targetPrice || "").trim();
        const tpNum = safeNum(tp);
        const upside = computeUpsideToTarget(equityStats.currentPrice, tpNum);

        body.push(new docx.Paragraph({
          children: [new docx.TextRun({ text: "Market stats", bold: true, font:"Times New Roman", size: 24 })],
          spacing: { before: 80, after: 80 }
        }));

        const statLines = [
          { k: "Current price", v: equityStats.currentPrice.toFixed(2) },
          { k: "Volatility (ann.)", v: equityStats.realisedVolAnn == null ? "—" : pct(equityStats.realisedVolAnn) },
          { k: "Return (range)", v: equityStats.rangeReturn == null ? "—" : pct(equityStats.rangeReturn) },
          { k: "Upside to target", v: upside == null ? "—" : pct(upside) }
        ];

        statLines.forEach(s => {
          body.push(new docx.Paragraph({
            children: [smallLabel(`${s.k}: `), smallValue(s.v)],
            spacing: { after: 55 }
          }));
        });

        body.push(new docx.Paragraph({ spacing: { after: 90 } }));
      }

      // Scenario table generator
      if (scenarioTable){
        body.push(new docx.Paragraph({
          children: [new docx.TextRun({ text: "Scenario table", bold: true, font:"Times New Roman", size: 24 })],
          spacing: { before: 80, after: 90 }
        }));
        body.push(buildScenarioTable(scenarioTable));
        body.push(new docx.Paragraph({ spacing: { after: 120 } }));
      }

      if ((valuationSummary || "").trim()){
        body.push(sectionHeading("Valuation Summary"));
        body.push(...linesToParagraphs(valuationSummary, 110));
      }

      if ((keyAssumptions || "").trim()){
        body.push(sectionHeading("Key Assumptions"));
        body.push(...bulletLines(keyAssumptions, 80));
      }

      if ((scenarioNotes || "").trim()){
        body.push(sectionHeading("Scenario / Sensitivity Notes"));
        body.push(...linesToParagraphs(scenarioNotes, 110));
      }

      // Model files list (names only)
      const attachedModelNames = (modelFiles && modelFiles.length) ? Array.from(modelFiles).map(f => f.name) : [];
      body.push(sectionHeading("Model Attachments"));
      if (attachedModelNames.length){
        attachedModelNames.forEach(name => {
          body.push(new docx.Paragraph({
            children: [new docx.TextRun({ text: name, font:"Times New Roman", size: 22 })],
            bullet: { level: 0 },
            spacing: { after: 70 }
          }));
        });
      } else {
        body.push(new docx.Paragraph({ children: [smallValue("None uploaded")], spacing: { after: 120 } }));
      }

      body.push(new docx.Paragraph({
        border: { bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } },
        spacing: { before: 180, after: 160 }
      }));
    }

    // Core body
    body.push(sectionHeading("Key Takeaways"));
    body.push(...bulletLines(keyTakeaways, 80));

    body.push(sectionHeading("Analysis and Commentary"));
    body.push(...linesToParagraphs(analysis, 120));

    if ((content || "").trim()){
      body.push(sectionHeading("Additional Content"));
      body.push(...linesToParagraphs(content, 120));
    }

    body.push(sectionHeading("The Cordoba View"));
    body.push(...linesToParagraphs(cordobaView, 120));

    // Figures
    const imageParagraphs = await addImages(imageFiles);
    if (imageParagraphs.length){
      body.push(sectionHeading("Figures and Charts"));
      body.push(...imageParagraphs);
    }

    // Footer
    const footer = new docx.Footer({
      children: [
        new docx.Paragraph({
          border: { top: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } },
          spacing: { after: 0 }
        }),
        new docx.Paragraph({
          children: [
            new docx.TextRun({ text: "\t" }),
            new docx.TextRun({
              text: watermarkOn ? "INTERNAL — DRAFT" : "INTERNAL",
              size: 16,
              font: "Times New Roman",
              italics: true
            }),
            new docx.TextRun({ text: "\t" }),
            new docx.TextRun({
              children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
              size: 16,
              font: "Times New Roman",
              italics: true
            })
          ],
          tabStops: [
            { type: docx.TabStopType.CENTER, position: 4680 },
            { type: docx.TabStopType.RIGHT, position: 9360 }
          ]
        })
      ]
    });

    // Header (with “watermark line” when Draft)
    const headerChildren = [];
    if (watermarkLine) headerChildren.push(watermarkLine);
    headerChildren.push(headerLine);

    const doc = new docx.Document({
      styles: {
        default: {
          document: {
            run: { font: "Times New Roman", size: 22, color: "000000" },
            paragraph: { spacing: { after: 120 } }
          }
        }
      },
      sections: [{
        properties: {
          page: {
            // Sell-side standard: portrait, readable margins
            margin: { top: 720, right: 720, bottom: 720, left: 720 }
          }
        },
        headers: { default: new docx.Header({ children: headerChildren }) },
        footers: { default: footer },
        children: body
      }]
    });

    return doc;
  }

  // ============================================================
  // Submit: validation + attestation lock + export + version bump
  // ============================================================
  const generateBtn = $("generateBtn");
  const attestEl = $("attestDraft");
  const messageDiv = $("message");

  function showMessage(type, text){
    if (!messageDiv) return;
    messageDiv.className = `message ${type || ""}`.trim();
    messageDiv.textContent = text || "";
  }

  function firstMissingElement(){
    const missing = listMissing();
    if (!missing.length) return null;
    return $(missing[0]);
  }

  // Cmd/Ctrl+Enter to generate
  document.addEventListener("keydown", (e) => {
    const isMac = navigator.platform.toUpperCase().includes("MAC");
    const mod = isMac ? e.metaKey : e.ctrlKey;
    if (mod && e.key === "Enter"){
      e.preventDefault();
      formEl?.requestSubmit();
    }
  });

  if (formEl) formEl.noValidate = true;

  formEl?.addEventListener("submit", async (e) => {
    e.preventDefault();

    // lock until attestation is checked
    if (attestEl && !attestEl.checked){
      showMessage("error", "✗ Export is locked. Please tick the attestation checkbox in Review & export.");
      $("sec-review")?.scrollIntoView({ behavior: "smooth", block: "center" });
      attestEl.focus();
      return;
    }

    // required validation
    const missing = listMissing();
    if (missing.length){
      showMessage("error", `✗ Please complete required fields. ${missing.length} missing.`);
      const el = firstMissingElement();
      if (el){
        el.scrollIntoView({ behavior: "smooth", block: "center" });
        setTimeout(() => { try { el.focus(); } catch(_){} }, 250);
      }
      return;
    }

    // Reviewed/Cleared should not proceed without reviewedBy (already in missing, but keep explicit)
    const st = getStatus();
    if ((st === "Reviewed" || st === "Cleared") && !($("reviewedBy")?.value || "").trim()){
      showMessage("error", "✗ Please select ‘Reviewed by’ before exporting Reviewed/Cleared notes.");
      $("reviewedBy")?.focus();
      return;
    }

    const button = generateBtn || formEl.querySelector('button[type="submit"]');
    if (button){
      button.disabled = true;
      button.classList.add("loading");
      button.textContent = "Generating…";
    }
    showMessage("", "");

    try{
      if (typeof docx === "undefined") throw new Error("docx library not loaded. Please refresh.");
      if (typeof saveAs === "undefined") throw new Error("FileSaver library not loaded. Please refresh.");

      // Version for this export (current)
      const currentV = getCurrentVersion();
      const versionStr = versionString(currentV);

      const noteType = noteTypeEl?.value || "";
      const title = $("title")?.value || "";
      const topic = $("topic")?.value || "";

      const status = getStatus();
      const reviewedBy = ($("reviewedBy")?.value || "").trim();
      const distPreset = getPreset();
      const changeNote = ($("changeNote")?.value || "").trim();

      const bumpMajorFlag = !!$("bumpMajor")?.checked;

      const authorLastName = $("authorLastName")?.value || "";
      const authorFirstName = $("authorFirstName")?.value || "";

      const authorPhoneSafe = naIfBlank($("authorPhone")?.value || "");
      const execSummary = ($("execSummary")?.value || "");

      const analysis = $("analysis")?.value || "";
      const keyTakeaways = $("keyTakeaways")?.value || "";
      const content = $("content")?.value || "";
      const cordobaView = $("cordobaView")?.value || "";
      const imageFiles = $("imageUpload")?.files || [];

      const ticker = $("ticker") ? $("ticker").value : "";
      const valuationSummary = $("valuationSummary") ? $("valuationSummary").value : "";
      const keyAssumptions = $("keyAssumptions") ? $("keyAssumptions").value : "";
      const scenarioNotes = $("scenarioNotes") ? $("scenarioNotes").value : "";
      const modelFiles = $("modelFiles") ? $("modelFiles").files : null;
      const modelLink = $("modelLink") ? $("modelLink").value : "";

      const targetPrice = $("targetPrice") ? $("targetPrice").value : "";
      const crgRating = $("crgRating") ? $("crgRating").value : "";

      const chartDataSource = ($("chartDataSource")?.value || "").trim() || "N/A";
      const chartDataSourceNote = ($("chartDataSourceNote")?.value || "").trim();
      const chartAnnotation = ($("chartAnnotation")?.value || "").trim();

      const scenarioTable = serializeScenario();

      const now = new Date();
      const dateTimeString = formatDateTime(now);

      const coAuthors = [];
      document.querySelectorAll(".coauthor-row").forEach(row => {
        const lastName = row.querySelector(".coauthor-lastname")?.value || "";
        const firstName = row.querySelector(".coauthor-firstname")?.value || "";
        const phone = row.querySelector(".coauthor-phone")?.value || "";
        if ((lastName || "").trim() && (firstName || "").trim()){
          coAuthors.push({ lastName, firstName, phone: naIfBlank(phone) });
        }
      });

      const doc = await createDocument({
        noteType, title, topic,
        status, reviewedBy, distPreset,
        versionStr, changeNote,
        execSummary,
        authorLastName, authorFirstName, authorPhoneSafe,
        coAuthors,
        analysis, keyTakeaways, content, cordobaView,
        imageFiles, dateTimeString,
        ticker, valuationSummary, keyAssumptions, scenarioNotes, modelFiles, modelLink,
        priceChartImageBytes,
        targetPrice,
        equityStats,
        crgRating,
        chartDataSource, chartDataSourceNote, chartAnnotation,
        scenarioTable
      });

      const blob = await docx.Packer.toBlob(doc);

      const safeTitle = (title || "research_note").replace(/[^a-z0-9]/gi, "_").toLowerCase();
      const safeType = (noteType || "note").replace(/\s+/g, "_").toLowerCase();
      const fileName = `${safeTitle}_${safeType}_${versionStr}_${status.toLowerCase()}.docx`;

      saveAs(blob, fileName);

      showMessage("success", `✓ Document "${fileName}" generated successfully.`);
      saveAutosave();

      // After successful export, bump version for next export (BlueMatrix always knows what changed)
      bumpVersion({ bumpMajor: bumpMajorFlag });
      // reset bump major toggle after applying
      if ($("bumpMajor")) $("bumpMajor").checked = false;
      refreshVersionUI();

    } catch (error){
      console.error("Error generating document:", error);
      showMessage("error", `✗ Error: ${error.message}`);
    } finally {
      if (button){
        button.disabled = false;
        button.classList.remove("loading");
        button.textContent = "Generate Word Document";
      }
    }
  });

  // ============================================================
  // Status change behaviour
  // ============================================================
  statusEl?.addEventListener("change", () => {
    updateWatermarkBadge();
    updateCompletionMeter();
    updateValidationSummary();
    autosaveSoon();
  });

  // ============================================================
  // Initialise: reviewers + restore autosave + version badge
  // ============================================================
  initReviewersDropdown();
  restoreAutosave();
  refreshVersionUI();
  updateWatermarkBadge();
  refreshExecSummary();

  // ============================================================
  // Preselect default preset if none
  // ============================================================
  if (distPresetEl && !distPresetEl.value) setPreset("internal");

  // ============================================================
  // Initial left rail state
  // ============================================================
  updateCompletionMeter();
  updateValidationSummary();
});
