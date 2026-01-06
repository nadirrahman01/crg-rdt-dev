/* assets/app.js
   CRG — Research Documentation Tool (Institutional / BlueMatrix-style workflow)

   Adds:
   - Versioning: auto v1.0 -> v1.1 increments (major/minor), change note; included in Word header + email subject
   - Workflow: Status (Draft/Reviewed/Cleared), Reviewed by, export gating
   - Watermarking system: Draft + Internal-only preset watermarked; Client-safe requires Cleared
   - Executive Summary: auto-generate toggle + editable textarea
   - Data source tracking (chart/stats) + chart annotation printed in Word
   - Distribution presets: Internal only / Public pack / Client-safe controlling watermark + export sections + email routing tags
   - Scenario table generator (Bear/Base/Bull) rendered in Word
   - Saved templates (Macro update / Equity initiation / Event note) that prefill and adjust required fields
*/

console.log("app.js loaded successfully");

window.addEventListener("DOMContentLoaded", () => {

  // ================================
  // Utilities
  // ================================
  const $ = (id) => document.getElementById(id);

  function digitsOnly(v){ return (v || "").toString().replace(/\D/g, ""); }

  function formatNationalLoose(rawDigits){
    const d = digitsOnly(rawDigits);
    if (!d) return "";
    const p1 = d.slice(0, 4);
    const p2 = d.slice(4, 7);
    const p3 = d.slice(7, 10);
    const rest = d.slice(10);
    return [p1, p2, p3, rest].filter(Boolean).join(" ");
  }

  function buildInternationalHyphen(ccDigits, nationalDigits){
    const cc = digitsOnly(ccDigits);
    const nn = digitsOnly(nationalDigits);
    if (!cc && !nn) return "";
    if (cc && !nn) return `${cc}-`;
    if (!cc && nn) return nn;
    return `${cc}-${nn}`;
  }

  function naIfBlank(v){
    const s = (v ?? "").toString().trim();
    return s ? s : "N/A";
  }

  function escapeText(s){
    return (s || "").toString()
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  // ================================
  // Date/time + session timer
  // ================================
  function formatDateTime(date){
    const months = [
      "January","February","March","April","May","June",
      "July","August","September","October","November","December"
    ];
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

  // ================================
  // Autosave (localStorage)
  // ================================
  const STORAGE_KEY = "crg_rdt_institutional_v1_draft";

  function getFormSnapshot(){
    const coAuthors = [];
    document.querySelectorAll(".coauthor-entry").forEach(entry => {
      const lastName = entry.querySelector(".coauthor-lastname")?.value || "";
      const firstName = entry.querySelector(".coauthor-firstname")?.value || "";
      const phone = entry.querySelector(".coauthor-phone")?.value || "";
      const cc = entry.querySelector(".coauthor-country")?.value || "";
      const local = entry.querySelector(".coauthor-phone-local")?.value || "";
      if (lastName || firstName || phone || local) {
        coAuthors.push({ lastName, firstName, phone, cc, local });
      }
    });

    const scenario = {
      bearPrice: $("bearPrice")?.value || "",
      bearKey: $("bearKey")?.value || "",
      basePrice: $("basePrice")?.value || "",
      baseKey: $("baseKey")?.value || "",
      bullPrice: $("bullPrice")?.value || "",
      bullKey: $("bullKey")?.value || ""
    };

    return {
      template: $("noteTemplate")?.value || "",
      noteType: $("noteType")?.value || "",
      status: $("workflowStatus")?.value || "Draft",
      reviewedBy: $("reviewedBy")?.value || "",
      reviewedByOther: $("reviewedByOther")?.value || "",
      distributionPreset: $("distributionPreset")?.value || "internal",
      versionMajor: $("versionMajor")?.value || "1",
      versionMinor: $("versionMinor")?.value || "0",
      changeNote: $("changeNote")?.value || "",
      majorBumpNext: $("majorBumpNext")?.checked || false,

      title: $("title")?.value || "",
      topic: $("topic")?.value || "",
      thesis: $("thesis")?.value || "",

      authorLastName: $("authorLastName")?.value || "",
      authorFirstName: $("authorFirstName")?.value || "",
      authorPhoneCountry: $("authorPhoneCountry")?.value || "",
      authorPhoneNational: $("authorPhoneNational")?.value || "",
      authorPhone: $("authorPhone")?.value || "",

      execAuto: $("execAuto")?.checked || false,
      executiveSummary: $("executiveSummary")?.value || "",

      ticker: $("ticker")?.value || "",
      crgRating: $("crgRating")?.value || "",
      targetPrice: $("targetPrice")?.value || "",
      chartRange: $("chartRange")?.value || "6mo",
      chartDataSource: $("chartDataSource")?.value || "Stooq",
      chartDataSourceOther: $("chartDataSourceOther")?.value || "",
      chartAnnotation: $("chartAnnotation")?.value || "",

      valuationSummary: $("valuationSummary")?.value || "",
      keyAssumptions: $("keyAssumptions")?.value || "",
      scenarioNotes: $("scenarioNotes")?.value || "",
      modelLink: $("modelLink")?.value || "",

      keyTakeaways: $("keyTakeaways")?.value || "",
      analysis: $("analysis")?.value || "",
      content: $("content")?.value || "",
      cordobaView: $("cordobaView")?.value || "",

      confirmDraft: $("confirmDraft")?.checked || false,

      coAuthors,
      scenario
    };
  }

  function applyFormSnapshot(s){
    if (!s) return;

    if ($("noteTemplate")) $("noteTemplate").value = s.template ?? "";
    if ($("noteType")) $("noteType").value = s.noteType ?? "";
    if ($("workflowStatus")) $("workflowStatus").value = s.status ?? "Draft";
    if ($("reviewedBy")) $("reviewedBy").value = s.reviewedBy ?? "";
    if ($("reviewedByOther")) $("reviewedByOther").value = s.reviewedByOther ?? "";
    if ($("distributionPreset")) $("distributionPreset").value = s.distributionPreset ?? "internal";

    if ($("versionMajor")) $("versionMajor").value = s.versionMajor ?? "1";
    if ($("versionMinor")) $("versionMinor").value = s.versionMinor ?? "0";
    if ($("changeNote")) $("changeNote").value = s.changeNote ?? "";
    if ($("majorBumpNext")) $("majorBumpNext").checked = !!s.majorBumpNext;

    if ($("title")) $("title").value = s.title ?? "";
    if ($("topic")) $("topic").value = s.topic ?? "";
    if ($("thesis")) $("thesis").value = s.thesis ?? "";

    if ($("authorLastName")) $("authorLastName").value = s.authorLastName ?? "";
    if ($("authorFirstName")) $("authorFirstName").value = s.authorFirstName ?? "";
    if ($("authorPhoneCountry")) $("authorPhoneCountry").value = s.authorPhoneCountry ?? "44";
    if ($("authorPhoneNational")) $("authorPhoneNational").value = s.authorPhoneNational ?? "";
    if ($("authorPhone")) $("authorPhone").value = s.authorPhone ?? "";

    if ($("execAuto")) $("execAuto").checked = !!s.execAuto;
    if ($("executiveSummary")) $("executiveSummary").value = s.executiveSummary ?? "";

    if ($("ticker")) $("ticker").value = s.ticker ?? "";
    if ($("crgRating")) $("crgRating").value = s.crgRating ?? "";
    if ($("targetPrice")) $("targetPrice").value = s.targetPrice ?? "";
    if ($("chartRange")) $("chartRange").value = s.chartRange ?? "6mo";
    if ($("chartDataSource")) $("chartDataSource").value = s.chartDataSource ?? "Stooq";
    if ($("chartDataSourceOther")) $("chartDataSourceOther").value = s.chartDataSourceOther ?? "";
    if ($("chartAnnotation")) $("chartAnnotation").value = s.chartAnnotation ?? "";

    if ($("valuationSummary")) $("valuationSummary").value = s.valuationSummary ?? "";
    if ($("keyAssumptions")) $("keyAssumptions").value = s.keyAssumptions ?? "";
    if ($("scenarioNotes")) $("scenarioNotes").value = s.scenarioNotes ?? "";
    if ($("modelLink")) $("modelLink").value = s.modelLink ?? "";

    if ($("keyTakeaways")) $("keyTakeaways").value = s.keyTakeaways ?? "";
    if ($("analysis")) $("analysis").value = s.analysis ?? "";
    if ($("content")) $("content").value = s.content ?? "";
    if ($("cordobaView")) $("cordobaView").value = s.cordobaView ?? "";

    if ($("confirmDraft")) $("confirmDraft").checked = !!s.confirmDraft;

    // Scenario
    if ($("bearPrice")) $("bearPrice").value = s.scenario?.bearPrice ?? "";
    if ($("bearKey")) $("bearKey").value = s.scenario?.bearKey ?? "";
    if ($("basePrice")) $("basePrice").value = s.scenario?.basePrice ?? "";
    if ($("baseKey")) $("baseKey").value = s.scenario?.baseKey ?? "";
    if ($("bullPrice")) $("bullPrice").value = s.scenario?.bullPrice ?? "";
    if ($("bullKey")) $("bullKey").value = s.scenario?.bullKey ?? "";

    // Coauthors (rebuild list)
    if (Array.isArray(s.coAuthors) && $("coAuthorsList")) {
      $("coAuthorsList").innerHTML = "";
      s.coAuthors.forEach((c) => addCoAuthorRow(c));
    }
  }

  let autosaveTimer = null;
  function autosaveNow(){
    try{
      const snap = getFormSnapshot();
      localStorage.setItem(STORAGE_KEY, JSON.stringify(snap));
      setAutosaveStatus("Saved");
    } catch(e){
      console.warn("Autosave failed:", e);
      setAutosaveStatus("Autosave error");
    }
  }

  function scheduleAutosave(){
    setAutosaveStatus("Typing…");
    if (autosaveTimer) clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(() => autosaveNow(), 450);
  }

  function clearAutosave(){
    localStorage.removeItem(STORAGE_KEY);
    setAutosaveStatus("Cleared");
  }

  function setAutosaveStatus(text){
    const el = $("autosaveStatus");
    if (el) el.textContent = text;
  }

  // ================================
  // Versioning
  // ================================
  function getVersion(){
    const major = parseInt($("versionMajor")?.value || "1", 10);
    const minor = parseInt($("versionMinor")?.value || "0", 10);
    return { major: Number.isFinite(major) ? major : 1, minor: Number.isFinite(minor) ? minor : 0 };
  }

  function formatVersion(v){
    // minor is integer "tenths": 0 => .0, 1 => .1, ... 9 => .9
    const m = Math.max(0, v.major);
    const mi = Math.max(0, v.minor);
    return `v${m}.${mi}`;
  }

  function setVersion(v){
    if ($("versionMajor")) $("versionMajor").value = String(v.major);
    if ($("versionMinor")) $("versionMinor").value = String(v.minor);
    paintVersionPill();
    scheduleAutosave();
  }

  function bumpVersion({ majorBump }){
    const v = getVersion();
    if (majorBump){
      setVersion({ major: v.major + 1, minor: 0 });
    } else {
      const nextMinor = v.minor + 1;
      if (nextMinor >= 10){
        setVersion({ major: v.major + 1, minor: 0 });
      } else {
        setVersion({ major: v.major, minor: nextMinor });
      }
    }
  }

  function paintVersionPill(){
    const v = getVersion();
    const el = $("versionPill");
    if (el) el.textContent = formatVersion(v);
  }

  // ================================
  // Workflow: status + reviewer + preset rules
  // ================================
  function getStatus(){ return ($("workflowStatus")?.value || "Draft").trim(); }

  function getReviewedBy(){
    const base = ($("reviewedBy")?.value || "").trim();
    if (base === "__other__") return ($("reviewedByOther")?.value || "").trim();
    return base;
  }

  function getPreset(){ return ($("distributionPreset")?.value || "internal").trim(); }

  function presetLabel(p){
    if (p === "internal") return "Internal only";
    if (p === "public") return "Public pack";
    if (p === "client") return "Client-safe";
    return p;
  }

  function isEquity(){ return ($("noteType")?.value || "") === "Equity Research"; }

  function toggleEquitySection(){
    const show = isEquity();
    const sec = $("equitySection");
    if (sec) sec.style.display = show ? "block" : "none";
    const railLink = $("equityRailLink");
    if (railLink) railLink.style.display = show ? "list-item" : "none";
  }

  function toggleReviewedByOther(){
    const sel = $("reviewedBy");
    const other = $("reviewedByOther");
    if (!sel || !other) return;
    const on = sel.value === "__other__";
    other.style.display = on ? "block" : "none";
  }

  function computeWatermarkMode(){
    // Compliance watermarking system:
    // - Draft => watermarked always
    // - Internal preset => watermarked always (even if reviewed)
    // - Public / Client-safe => no watermark, but client-safe requires Cleared status
    const status = getStatus();
    const preset = getPreset();

    if (status === "Draft") return "DRAFT";
    if (preset === "internal") return "INTERNAL";
    return "NONE";
  }

  function canExport(){
    // Gating:
    // - Must confirm draft checkbox
    // - If status Reviewed/Cleared => reviewedBy required
    // - If preset client-safe => status must be Cleared
    if (!$("confirmDraft")?.checked) return { ok:false, reason:"Tick the confirmation before exporting." };

    const status = getStatus();
    const reviewer = getReviewedBy();
    if ((status === "Reviewed" || status === "Cleared") && !reviewer) {
      return { ok:false, reason:"Status is Reviewed/Cleared — please select ‘Reviewed by’." };
    }

    const preset = getPreset();
    if (preset === "client" && status !== "Cleared"){
      return { ok:false, reason:"Client-safe export requires Status = Cleared." };
    }

    return { ok:true, reason:"" };
  }

  // ================================
  // Templates (saved templates)
  // ================================
  const TEMPLATES = {
    "": { name:"—", noteType:"", defaults:{} },
    "macro_update": {
      name:"Macro update",
      noteType:"Macro Research",
      defaults:{
        title:"Macro Update — ",
        thesis:"",
        keyTakeaways:"- \n- \n- ",
        analysis:"",
        cordobaView:""
      }
    },
    "equity_initiation": {
      name:"Equity initiation",
      noteType:"Equity Research",
      defaults:{
        title:"Initiation — ",
        thesis:"",
        crgRating:"",
        valuationSummary:"",
        keyAssumptions:"",
        scenarioNotes:""
      }
    },
    "event_note": {
      name:"Event note",
      noteType:"General Note",
      defaults:{
        title:"Event Note — ",
        thesis:"",
        keyTakeaways:"- What happened\n- Why it matters\n- What we do next",
        analysis:"",
        cordobaView:""
      }
    }
  };

  function applyTemplate(key){
    const t = TEMPLATES[key] || TEMPLATES[""];
    if (t.noteType && $("noteType")) $("noteType").value = t.noteType;

    // apply defaults only if field is blank (avoid overwriting analyst work)
    Object.entries(t.defaults || {}).forEach(([k,v]) => {
      const el = $(k);
      if (!el) return;
      const cur = (el.value || "").trim();
      if (!cur) el.value = v;
    });

    toggleEquitySection();
    updateCompletionMeter();
    maybeGenerateExecutiveSummary();
    scheduleAutosave();
  }

  // ================================
  // Completion meter + validation panel
  // ================================
  function isFilled(el){
    if (!el) return false;
    if (el.type === "file") return el.files && el.files.length > 0;
    if (el.type === "checkbox") return !!el.checked;
    const v = (el.value ?? "").toString().trim();
    return v.length > 0;
  }

  function requiredCoreIds(){
    // Core fields for institutional output (adjusted)
    const base = [
      "noteType",
      "workflowStatus",
      "title",
      "topic",
      "authorLastName",
      "authorFirstName",
      "thesis",
      "keyTakeaways",
      "analysis"
    ];

    // ReviewedBy required when status reviewed/cleared
    const status = getStatus();
    if (status === "Reviewed" || status === "Cleared") base.push("reviewedBy");

    // Equity core extras
    if (isEquity()) {
      base.push("crgRating");
      base.push("targetPrice");
    }

    // Compliance tick is mandatory (export gate)
    base.push("confirmDraft");

    return base;
  }

  function firstMissingId(){
    const ids = requiredCoreIds();
    for (const id of ids){
      const el = $(id);
      if (id === "reviewedBy") {
        if ((getStatus() === "Reviewed" || getStatus() === "Cleared") && !getReviewedBy()) return "reviewedBy";
        continue;
      }
      if (!isFilled(el)) return id;
    }
    return null;
  }

  function updateCompletionMeter(){
    const ids = requiredCoreIds();
    let done = 0;

    ids.forEach((id) => {
      if (id === "reviewedBy") {
        if ((getStatus() === "Reviewed" || getStatus() === "Cleared") && getReviewedBy()) done++;
        return;
      }
      const el = $(id);
      if (isFilled(el)) done++;
    });

    const total = ids.length;
    const pct = total ? Math.round((done / total) * 100) : 0;

    if ($("completionText")) $("completionText").textContent = `${done} / ${total} core fields`;
    if ($("completionBar")) $("completionBar").style.width = `${pct}%`;

    const railPct = $("railPct");
    if (railPct) railPct.textContent = `${pct}%`;

    const railBar = $("railBarFill");
    if (railBar) railBar.style.width = `${pct}%`;

    // Validation summary
    renderValidationSummary();

    // Export button enable/disable
    const exportBtn = $("generateDocBtn");
    const gate = canExport();
    if (exportBtn){
      exportBtn.disabled = !gate.ok;
      exportBtn.title = gate.ok ? "" : gate.reason;
    }
  }

  function renderValidationSummary(){
    const el = $("validationSummary");
    if (!el) return;

    const missing = [];
    const missId = firstMissingId();
    if (missId) missing.push(missId);

    const gate = canExport();
    const items = [];

    // show key items
    items.push({ label:"Status", ok: !!getStatus(), note:getStatus() });
    items.push({ label:"Version", ok: true, note: formatVersion(getVersion()) });
    items.push({ label:"Preset", ok: true, note: presetLabel(getPreset()) });

    if (getStatus() === "Reviewed" || getStatus() === "Cleared"){
      items.push({ label:"Reviewed by", ok: !!getReviewedBy(), note: getReviewedBy() || "Missing" });
    }

    if (getPreset() === "client"){
      items.push({ label:"Client-safe requires Cleared", ok: getStatus() === "Cleared", note: getStatus() });
    }

    items.push({ label:"Confirmation ticked", ok: !!$("confirmDraft")?.checked, note: $("confirmDraft")?.checked ? "Yes" : "No" });

    el.innerHTML = items.map(i => {
      const mark = i.ok ? "✓" : "✗";
      const cls = i.ok ? "val-ok" : "val-bad";
      return `<div class="val-row ${cls}"><span class="val-mark">${mark}</span><span class="val-label">${escapeText(i.label)}:</span> <span class="val-note">${escapeText(i.note || "")}</span></div>`;
    }).join("");

    const warn = $("exportWarning");
    if (warn) warn.textContent = gate.ok ? "" : gate.reason;
  }

  function markInvalid(id, on){
    const el = $(id);
    if (!el) return;
    if (on) el.classList.add("is-invalid");
    else el.classList.remove("is-invalid");
  }

  function jumpToFirstMissing(){
    const miss = firstMissingId();
    if (!miss) return;
    markInvalid(miss, true);

    const el = $(miss);
    if (!el) return;

    // Ensure section visible
    if (miss === "crgRating" || miss === "targetPrice") toggleEquitySection();

    el.scrollIntoView({ behavior: "smooth", block: "center" });
    try { el.focus({ preventScroll: true }); } catch(_) {}
    setTimeout(() => markInvalid(miss, false), 1600);
  }

  // ================================
  // Word counts
  // ================================
  function countWords(text){
    const t = (text || "").trim();
    if (!t) return 0;
    return t.split(/\s+/).filter(Boolean).length;
  }

  function paintWordCounts(){
    const map = [
      ["thesis","thesisWC"],
      ["keyTakeaways","takeawaysWC"],
      ["analysis","analysisWC"],
      ["executiveSummary","execWC"]
    ];
    map.forEach(([id,out]) => {
      const el = $(id);
      const outEl = $(out);
      if (!el || !outEl) return;
      outEl.textContent = `${countWords(el.value)} words`;
    });
  }

  // ================================
  // Phone syncing
  // ================================
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
    try{ authorPhoneNationalEl.setSelectionRange(next, next); } catch(_) {}
    syncPrimaryPhone();
  }

  if (authorPhoneNationalEl){
    authorPhoneNationalEl.addEventListener("input", () => { formatPrimaryVisible(); scheduleAutosave(); updateCompletionMeter(); });
    authorPhoneNationalEl.addEventListener("blur", () => { syncPrimaryPhone(); scheduleAutosave(); });
  }
  if (authorPhoneCountryEl){
    authorPhoneCountryEl.addEventListener("change", () => { syncPrimaryPhone(); scheduleAutosave(); });
  }

  // ================================
  // Coauthor management + phone wiring
  // ================================
  let coAuthorCount = 0;

  const countryOptionsHtml = `
    <option value="44">+44</option>
    <option value="1">+1</option>
    <option value="353">+353</option>
    <option value="33">+33</option>
    <option value="49">+49</option>
    <option value="31">+31</option>
    <option value="971">+971</option>
    <option value="966">+966</option>
    <option value="92">+92</option>
    <option value="">Other</option>
  `;

  function wireCoauthorPhone(coAuthorDiv){
    const ccEl = coAuthorDiv.querySelector(".coauthor-country");
    const nationalEl = coAuthorDiv.querySelector(".coauthor-phone-local");
    const hiddenEl = coAuthorDiv.querySelector(".coauthor-phone");
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
      try{ nationalEl.setSelectionRange(next, next); } catch(_) {}
      syncHidden();
    }

    if (nationalEl){
      nationalEl.addEventListener("input", () => { formatVisible(); scheduleAutosave(); });
      nationalEl.addEventListener("blur", () => { syncHidden(); scheduleAutosave(); });
    }
    if (ccEl) ccEl.addEventListener("change", () => { syncHidden(); scheduleAutosave(); });

    syncHidden();
  }

  function addCoAuthorRow(preset){
    coAuthorCount++;
    const wrap = document.createElement("div");
    wrap.className = "coauthor-entry coauthor-row";
    wrap.id = `coauthor-${coAuthorCount}`;
    wrap.innerHTML = `
      <div class="coauthor-grid">
        <div>
          <label>Last Name</label>
          <input type="text" class="coauthor-lastname" placeholder="e.g., Rahman" value="${escapeText(preset?.lastName || "")}">
        </div>
        <div>
          <label>First Name</label>
          <input type="text" class="coauthor-firstname" placeholder="e.g., Nadir" value="${escapeText(preset?.firstName || "")}">
        </div>
        <div>
          <label>Phone (optional)</label>
          <div class="phone-row phone-row--compact">
            <select class="coauthor-country">
              ${countryOptionsHtml}
            </select>
            <input type="text" class="coauthor-phone-local" placeholder="e.g., 7323 324 120" value="${escapeText(preset?.local || "")}">
          </div>
          <input type="text" class="coauthor-phone" style="display:none;" value="${escapeText(preset?.phone || "")}">
        </div>
        <div style="display:flex; align-items:flex-end;">
          <button type="button" class="btn btn-danger remove-coauthor" data-remove-id="${coAuthorCount}">Remove</button>
        </div>
      </div>
    `;

    $("coAuthorsList")?.appendChild(wrap);

    // set cc default from preset
    const ccEl = wrap.querySelector(".coauthor-country");
    if (ccEl) ccEl.value = preset?.cc ?? "44";

    wireCoauthorPhone(wrap);

    wrap.querySelectorAll("input,select").forEach(el => {
      el.addEventListener("input", () => { scheduleAutosave(); updateCompletionMeter(); paintWordCounts(); });
      el.addEventListener("change", () => { scheduleAutosave(); updateCompletionMeter(); paintWordCounts(); });
    });

    updateCompletionMeter();
  }

  $("addCoAuthor")?.addEventListener("click", () => addCoAuthorRow());

  document.addEventListener("click", (e) => {
    const btn = e.target.closest(".remove-coauthor");
    if (!btn) return;
    const id = btn.getAttribute("data-remove-id");
    const row = document.getElementById(`coauthor-${id}`);
    if (row) row.remove();
    scheduleAutosave();
    updateCompletionMeter();
  });

  // ================================
  // Executive summary (auto-generate)
  // ================================
  function firstTakeawayLine(){
    const t = ($("keyTakeaways")?.value || "").split("\n")
      .map(x => x.replace(/^[-*•]\s*/, "").trim())
      .filter(Boolean);
    return t.length ? t[0] : "";
  }

  function buildExecSummary(){
    const thesis = ($("thesis")?.value || "").trim();
    const firstBullet = firstTakeawayLine();
    const rating = ($("crgRating")?.value || "").trim();
    const tp = ($("targetPrice")?.value || "").trim();
    const ticker = ($("ticker")?.value || "").trim();

    const lines = [];
    if (thesis) lines.push(thesis);
    if (firstBullet) lines.push(firstBullet);

    if (isEquity()){
      const bits = [];
      if (ticker) bits.push(ticker.toUpperCase());
      if (rating) bits.push(rating);
      if (tp) bits.push(`Target: ${tp}`);
      if (bits.length) lines.push(bits.join(" · "));
    }

    return lines.join("\n");
  }

  function maybeGenerateExecutiveSummary(){
    if (!$("execAuto")?.checked) return;
    const out = $("executiveSummary");
    if (!out) return;
    out.value = buildExecSummary();
    paintWordCounts();
  }

  $("execAuto")?.addEventListener("change", () => {
    maybeGenerateExecutiveSummary();
    scheduleAutosave();
    updateCompletionMeter();
  });

  $("regenExecBtn")?.addEventListener("click", () => {
    $("executiveSummary").value = buildExecSummary();
    paintWordCounts();
    scheduleAutosave();
  });

  // ================================
  // Distribution preset buttons
  // ================================
  function setPreset(p){
    if ($("distributionPreset")) $("distributionPreset").value = p;
    // preset effects: watermark + gating + small UI note
    paintPresetPill();
    updateCompletionMeter();
    scheduleAutosave();
  }

  function paintPresetPill(){
    const el = $("presetPill");
    if (!el) return;
    el.textContent = presetLabel(getPreset());
  }

  $("presetInternalBtn")?.addEventListener("click", () => setPreset("internal"));
  $("presetPublicBtn")?.addEventListener("click", () => setPreset("public"));
  $("presetClientBtn")?.addEventListener("click", () => setPreset("client"));

  // ================================
  // Equity: chart + stats + sources + annotations
  // ================================
  let priceChart = null;
  let priceChartImageBytes = null;

  let equityStats = { currentPrice:null, realisedVolAnn:null, rangeReturn:null };

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
    // Stooq has no CORS; use r.jina.ai proxy
    const stooqUrl = `http://stooq.com/q/d/l/?s=${encodeURIComponent(symbol)}&i=d`;
    const proxyUrl = `https://r.jina.ai/${stooqUrl}`;

    const res = await fetch(proxyUrl, { cache:"no-store" });
    if (!res.ok) throw new Error("Could not fetch price data (proxy blocked or down).");

    const rawText = await res.text();
    const csvText = extractStooqCSV(rawText) || rawText;

    const lines = csvText.trim().split("\n");
    if (lines.length < 5) throw new Error("Not enough data returned. Check ticker.");

    const rows = lines.slice(1).map(line => line.split(","));
    const out = rows.map(r => ({ date:r[0], close:Number(r[4]) }))
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
          legend: { display:false },
          tooltip: { intersect:false, mode:"index" }
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
    for (let i=1;i<closes.length;i++){
      const prev = closes[i-1];
      const cur = closes[i];
      if (prev > 0 && Number.isFinite(prev) && Number.isFinite(cur)) rets.push((cur/prev)-1);
    }
    return rets;
  }

  function stddev(arr){
    if (!arr.length) return null;
    const mean = arr.reduce((a,b)=>a+b,0)/arr.length;
    const v = arr.reduce((a,b)=>a+(b-mean)**2,0)/(arr.length-1||1);
    return Math.sqrt(v);
  }

  function setText(id, text){
    const el = $(id);
    if (el) el.textContent = text;
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
    scheduleAutosave();
    updateCompletionMeter();
  });

  $("chartDataSource")?.addEventListener("change", () => {
    const other = $("chartDataSourceOther");
    if (!other) return;
    other.style.display = ($("chartDataSource").value === "__other__") ? "block" : "none";
    scheduleAutosave();
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
      equityStats = { currentPrice:null, realisedVolAnn:null, rangeReturn:null };
      setText("currentPrice","—");
      setText("rangeReturn","—");
      setText("realisedVol","—");
      setText("upsideToTarget","—");
      if (chartStatus) chartStatus.textContent = `✗ ${e.message}`;
    } finally {
      scheduleAutosave();
      updateCompletionMeter();
      maybeGenerateExecutiveSummary();
    }
  }

  fetchChartBtn?.addEventListener("click", buildPriceChart);

  // ================================
  // Attachment summary (model files)
  // ================================
  function updateAttachmentSummary(){
    const modelFilesEl = $("modelFiles");
    const head = $("attachmentSummaryHead");
    const list = $("attachmentSummaryList");
    if (!modelFilesEl || !head || !list) return;

    const files = Array.from(modelFilesEl.files || []);
    if (!files.length){
      head.textContent = "No model files selected";
      list.style.display = "none";
      list.innerHTML = "";
      return;
    }

    head.textContent = `${files.length} file${files.length===1?"":"s"} selected`;
    list.style.display = "block";
    list.innerHTML = files.map(f => `<div class="attach-file">${escapeText(f.name)}</div>`).join("");
  }

  $("modelFiles")?.addEventListener("change", () => {
    updateAttachmentSummary();
    scheduleAutosave();
    updateCompletionMeter();
  });

  // ================================
  // Images to Word
  // ================================
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
                transformation: { width: 560, height: 360 }
              })
            ],
            spacing: { before: 180, after: 80 },
            alignment: docx.AlignmentType.CENTER
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: `Figure ${i+1}: ${fileNameWithoutExt}`,
                italics: true,
                size: 18,
                font: "Calibri"
              })
            ],
            spacing: { after: 220 },
            alignment: docx.AlignmentType.CENTER
          })
        );
      } catch(err){
        console.error(`Error processing image ${file.name}:`, err);
      }
    }
    return imageParagraphs;
  }

  function linesToParagraphs(text, spacingAfter=140){
    const lines = (text || "").split("\n");
    return lines.map(line => {
      if (line.trim() === "") return new docx.Paragraph({ text:"", spacing:{ after: spacingAfter } });
      return new docx.Paragraph({ text: line, spacing:{ after: spacingAfter } });
    });
  }

  function bulletsFromLines(text){
    const lines = (text || "").split("\n");
    return lines.map(line => {
      if (line.trim() === "") return new docx.Paragraph({ text:"", spacing:{ after: 80 } });
      const clean = line.replace(/^[-*•]\s*/, "").trim();
      return new docx.Paragraph({ text: clean, bullet:{ level:0 }, spacing:{ after: 80 } });
    });
  }

  function hyperlinkParagraph(label, url){
    const safeUrl = (url || "").trim();
    if (!safeUrl) return null;

    return new docx.Paragraph({
      children: [
        new docx.TextRun({ text: label, bold:true }),
        new docx.TextRun({ text: " " }),
        new docx.ExternalHyperlink({
          children: [new docx.TextRun({ text: safeUrl, style:"Hyperlink" })],
          link: safeUrl
        })
      ],
      spacing: { after: 120 }
    });
  }

  // ================================
  // Scenario table (Bear/Base/Bull)
  // ================================
  function buildScenarioRows(){
    const rows = [];

    const bearPrice = ($("bearPrice")?.value || "").trim();
    const bearKey = ($("bearKey")?.value || "").trim();
    const basePrice = ($("basePrice")?.value || "").trim();
    const baseKey = ($("baseKey")?.value || "").trim();
    const bullPrice = ($("bullPrice")?.value || "").trim();
    const bullKey = ($("bullKey")?.value || "").trim();

    if (bearPrice || bearKey) rows.push({ name:"Bear", price:bearPrice, key:bearKey });
    if (basePrice || baseKey) rows.push({ name:"Base", price:basePrice, key:baseKey });
    if (bullPrice || bullKey) rows.push({ name:"Bull", price:bullPrice, key:bullKey });

    return rows;
  }

  function scenarioTableDoc(rows){
    if (!rows || !rows.length) return null;

    const headerRow = new docx.TableRow({
      children: [
        new docx.TableCell({ children:[ new docx.Paragraph({ children:[ new docx.TextRun({ text:"Scenario", bold:true }) ] }) ] }),
        new docx.TableCell({ children:[ new docx.Paragraph({ children:[ new docx.TextRun({ text:"Price", bold:true }) ] }) ] }),
        new docx.TableCell({ children:[ new docx.Paragraph({ children:[ new docx.TextRun({ text:"Key assumption / driver", bold:true }) ] }) ] })
      ]
    });

    const bodyRows = rows.map(r => new docx.TableRow({
      children: [
        new docx.TableCell({ children:[ new docx.Paragraph(r.name) ] }),
        new docx.TableCell({ children:[ new docx.Paragraph(r.price || "—") ] }),
        new docx.TableCell({ children:[ new docx.Paragraph(r.key || "—") ] })
      ]
    }));

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [headerRow, ...bodyRows]
    });
  }

  // ================================
  // Data source + chart annotation helpers
  // ================================
  function getChartSource(){
    const v = ($("chartDataSource")?.value || "Stooq").trim();
    if (v === "__other__") return ($("chartDataSourceOther")?.value || "").trim() || "Other";
    return v;
  }

  // ================================
  // Word document — BlueMatrix-like layout
  // ================================
  async function createDocument(data){
    const {
      template,
      noteType, status, reviewedBy, distributionPreset,
      versionString, changeNote,
      title, topic, thesis,
      authorLastName, authorFirstName, authorPhonePrintable,
      coAuthors,
      execSummary,
      analysis, keyTakeaways, content, cordobaView,
      ticker, crgRating, targetPrice,
      equityStats,
      priceChartImageBytes,
      chartSource,
      chartAnnotation,
      scenarioRows,
      valuationSummary, keyAssumptions, scenarioNotes,
      modelFiles, modelLink,
      imageFiles,
      dateTimeString
    } = data;

    const watermarkMode = computeWatermarkMode();

    const isClientSafe = distributionPreset === "client";
    const isPublicPack = distributionPreset === "public";
    const isInternal = distributionPreset === "internal";

    // export inclusion rules by preset
    const includeCordobaView = isInternal;                 // hide in public/client
    const includePhones = isInternal;                      // hide in public/client
    const includeModelLink = isInternal;                   // hide in public/client
    const includeAttachedModelList = isInternal;           // hide in public/client
    const includeChangeNote = isInternal;                  // change note is internal audit trail
    const includeExecSummary = true;                       // all presets keep this (it’s the front page)

    const titleRun = new docx.TextRun({ text: title || "", bold:true, size: 32, font: "Times New Roman" });
    const topicRun = new docx.TextRun({ text: `Topic: ${topic || "—"}`, size: 20, font: "Calibri", color: "444444" });

    const authorsLine = (() => {
      const main = `${(authorLastName||"").toUpperCase()}, ${(authorFirstName||"").toUpperCase()}`;
      const phone = includePhones ? ` (${authorPhonePrintable})` : "";
      return `${main}${phone}`;
    })();

    const coLines = (coAuthors || []).map(c => {
      const ln = (c.lastName || "").toUpperCase();
      const fn = (c.firstName || "").toUpperCase();
      const ph = includePhones ? ` (${naIfBlank(c.phone)})` : "";
      return `${ln}, ${fn}${ph}`;
    });

    const headerLeft = "Cordoba Research Group";
    const headerRight = `${noteType} · ${versionString} · ${status} · ${formatDateShort(new Date())}`;

    const distTag = presetLabel(distributionPreset).toUpperCase();
    const watermarkHeaderLine =
      (watermarkMode === "DRAFT")
        ? "DRAFT — INTERNAL — NOT FOR DISTRIBUTION"
        : (watermarkMode === "INTERNAL" ? "INTERNAL USE ONLY" : "");

    // header: (optional large watermark line) + normal line
    const headerChildren = [];

    if (watermarkMode === "DRAFT" || watermarkMode === "INTERNAL"){
      // Simulated watermark in header (centered, light)
      headerChildren.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: (watermarkMode === "DRAFT") ? "DRAFT" : "INTERNAL",
              size: 72,
              color: "D9D9D9",
              bold: true,
              font: "Calibri"
            })
          ],
          alignment: docx.AlignmentType.CENTER,
          spacing: { after: 0 }
        })
      );
    }

    headerChildren.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: headerLeft, bold:true, size: 18, font:"Calibri" }),
          new docx.TextRun({ text: "\t" }),
          new docx.TextRun({ text: headerRight, size: 18, font:"Calibri" })
        ],
        tabStops: [
          { type: docx.TabStopType.RIGHT, position: 9360 }
        ],
        spacing: { after: 60 },
        border: { bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } }
      })
    );

    if (watermarkHeaderLine){
      headerChildren.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({ text: watermarkHeaderLine, bold:true, size: 18, font:"Calibri", color:"555555" }),
            new docx.TextRun({ text: `  |  ${distTag}`, size: 18, font:"Calibri", color:"555555" })
          ],
          alignment: docx.AlignmentType.LEFT,
          spacing: { after: 80 }
        })
      );
    } else {
      headerChildren.push(
        new docx.Paragraph({
          children: [ new docx.TextRun({ text: distTag, size: 18, font:"Calibri", color:"555555" }) ],
          spacing: { after: 80 }
        })
      );
    }

    const footer = new docx.Footer({
      children: [
        new docx.Paragraph({
          border: { top: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } },
          spacing: { after: 0 }
        }),
        new docx.Paragraph({
          children: [
            new docx.TextRun({ text: "Confidential — for intended recipients only. Verify all figures and sources.", size: 16, font:"Calibri", italics:true }),
            new docx.TextRun({ text: "\t" }),
            new docx.TextRun({ children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES], size: 16, font:"Calibri", italics:true })
          ],
          tabStops: [
            { type: docx.TabStopType.RIGHT, position: 9360 }
          ],
          spacing: { before: 80, after: 0 }
        })
      ]
    });

    const children = [];

    // Cover / top block (BlueMatrix-ish)
    children.push(
      new docx.Paragraph({ children:[titleRun], spacing:{ after: 80 } }),
      new docx.Paragraph({ children:[topicRun], spacing:{ after: 120 } }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: "Author: ", bold:true, font:"Calibri" }),
          new docx.TextRun({ text: authorsLine, font:"Calibri" })
        ],
        spacing: { after: 60 }
      })
    );

    if (coLines.length){
      children.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({ text: "Co-authors: ", bold:true, font:"Calibri" }),
            new docx.TextRun({ text: coLines.join(" | "), font:"Calibri" })
          ],
          spacing: { after: 80 }
        })
      );
    } else {
      children.push(new docx.Paragraph({ spacing:{ after: 60 } }));
    }

    // Internal audit metadata block
    const metaLines = [];
    metaLines.push(`Template: ${template || "—"}`);
    metaLines.push(`Version: ${versionString}`);
    metaLines.push(`Status: ${status}`);
    if (reviewedBy) metaLines.push(`Reviewed by: ${reviewedBy}`);
    metaLines.push(`Generated: ${dateTimeString}`);
    if (includeChangeNote && (changeNote || "").trim()) metaLines.push(`Change note: ${changeNote.trim()}`);

    children.push(
      new docx.Paragraph({
        children: [ new docx.TextRun({ text: metaLines.join("  |  "), size: 18, font:"Calibri", color:"555555" }) ],
        spacing: { after: 160 }
      })
    );

    // Executive summary
    if (includeExecSummary){
      children.push(
        new docx.Paragraph({
          children: [ new docx.TextRun({ text: "Executive Summary", bold:true, size: 24, font:"Calibri" }) ],
          spacing: { after: 120 }
        })
      );
      const execText = (execSummary || "").trim() || buildExecSummary();
      children.push(...linesToParagraphs(execText, 120));
      children.push(new docx.Paragraph({ spacing: { after: 160 } }));
    }

    // Equity module
    if (noteType === "Equity Research"){
      children.push(
        new docx.Paragraph({
          children: [ new docx.TextRun({ text: "Equity Snapshot", bold:true, size: 24, font:"Calibri" }) ],
          spacing: { after: 120 }
        })
      );

      const snapBits = [];
      if ((ticker || "").trim()) snapBits.push(`Ticker: ${ticker.trim()}`);
      if ((crgRating || "").trim()) snapBits.push(`CRG Rating: ${crgRating.trim()}`);
      if ((targetPrice || "").trim()) snapBits.push(`Target: ${targetPrice.trim()}`);
      if (snapBits.length){
        children.push(
          new docx.Paragraph({
            children: [ new docx.TextRun({ text: snapBits.join("  |  "), font:"Calibri" }) ],
            spacing: { after: 120 }
          })
        );
      }

      // Price chart
      if (priceChartImageBytes){
        children.push(
          new docx.Paragraph({
            children: [ new docx.TextRun({ text: "Price chart", bold:true, size: 22, font:"Calibri" }) ],
            spacing: { after: 80 }
          }),
          new docx.Paragraph({
            children: [ new docx.ImageRun({ data: priceChartImageBytes, transformation: { width: 640, height: 260 } }) ],
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 80 }
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: `Source: ${chartSource || "N/A"}; CRG estimates`, italics:true, size: 18, font:"Calibri", color:"555555" })
            ],
            spacing: { after: 80 }
          })
        );

        if ((chartAnnotation || "").trim()){
          children.push(
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: `Note: ${(chartAnnotation || "").trim()}`, size: 18, font:"Calibri" })
              ],
              spacing: { after: 120 }
            })
          );
        }
      }

      // Market stats
      if (equityStats && equityStats.currentPrice){
        const tpNum = safeNum((targetPrice || "").trim());
        const upside = computeUpsideToTarget(equityStats.currentPrice, tpNum);

        children.push(
          new docx.Paragraph({
            children: [ new docx.TextRun({ text: "Market stats", bold:true, size: 22, font:"Calibri" }) ],
            spacing: { before: 80, after: 80 }
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Current price: ", bold:true, font:"Calibri" }),
              new docx.TextRun({ text: equityStats.currentPrice.toFixed(2), font:"Calibri" })
            ],
            spacing: { after: 60 }
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Volatility (ann.): ", bold:true, font:"Calibri" }),
              new docx.TextRun({ text: equityStats.realisedVolAnn == null ? "—" : pct(equityStats.realisedVolAnn), font:"Calibri" })
            ],
            spacing: { after: 60 }
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Return (range): ", bold:true, font:"Calibri" }),
              new docx.TextRun({ text: equityStats.rangeReturn == null ? "—" : pct(equityStats.rangeReturn), font:"Calibri" })
            ],
            spacing: { after: 60 }
          })
        );

        if (tpNum){
          children.push(
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: "+/- to target: ", bold:true, font:"Calibri" }),
                new docx.TextRun({ text: upside == null ? "—" : pct(upside), font:"Calibri" })
              ],
              spacing: { after: 100 }
            })
          );
        } else {
          children.push(new docx.Paragraph({ spacing: { after: 80 } }));
        }
      }

      // Scenario table
      const scenTable = scenarioTableDoc(scenarioRows);
      if (scenTable){
        children.push(
          new docx.Paragraph({
            children: [ new docx.TextRun({ text: "Scenario table", bold:true, size: 22, font:"Calibri" }) ],
            spacing: { before: 60, after: 80 }
          }),
          scenTable,
          new docx.Paragraph({ spacing: { after: 140 } })
        );
      }

      // Model link + attachments (internal only)
      if (includeModelLink){
        const linkPara = hyperlinkParagraph("Model link:", modelLink);
        if (linkPara) children.push(linkPara);
      }

      if (includeAttachedModelList){
        const attachedNames = (modelFiles && modelFiles.length) ? Array.from(modelFiles).map(f => f.name) : [];
        children.push(
          new docx.Paragraph({
            children: [ new docx.TextRun({ text: "Attached model files", bold:true, size: 22, font:"Calibri" }) ],
            spacing: { before: 80, after: 80 }
          })
        );
        if (attachedNames.length){
          attachedNames.forEach(name => children.push(new docx.Paragraph({ text: name, bullet:{ level:0 }, spacing:{ after: 60 } })));
        } else {
          children.push(new docx.Paragraph({ text: "None uploaded", spacing: { after: 100 } }));
        }
      }

      // Valuation / assumptions / sensitivity (all presets can include; internal depth still optional)
      if ((valuationSummary || "").trim()){
        children.push(
          new docx.Paragraph({
            children: [ new docx.TextRun({ text: "Valuation summary", bold:true, size: 22, font:"Calibri" }) ],
            spacing: { before: 120, after: 80 }
          }),
          ...linesToParagraphs(valuationSummary, 120)
        );
      }

      if ((keyAssumptions || "").trim()){
        children.push(
          new docx.Paragraph({
            children: [ new docx.TextRun({ text: "Key assumptions", bold:true, size: 22, font:"Calibri" }) ],
            spacing: { before: 120, after: 80 }
          })
        );
        // bullets from assumptions
        keyAssumptions.split("\n").forEach(line => {
          if (!line.trim()) return;
          children.push(
            new docx.Paragraph({
              text: line.replace(/^[-*•]\s*/, "").trim(),
              bullet:{ level:0 },
              spacing:{ after: 60 }
            })
          );
        });
      }

      if ((scenarioNotes || "").trim()){
        children.push(
          new docx.Paragraph({
            children: [ new docx.TextRun({ text: "Scenario / sensitivity notes", bold:true, size: 22, font:"Calibri" }) ],
            spacing: { before: 120, after: 80 }
          }),
          ...linesToParagraphs(scenarioNotes, 120)
        );
      }

      children.push(new docx.Paragraph({ spacing: { after: 160 } }));
    }

    // Thesis
    if ((thesis || "").trim()){
      children.push(
        new docx.Paragraph({
          children: [ new docx.TextRun({ text: "Thesis", bold:true, size: 24, font:"Calibri" }) ],
          spacing: { after: 120 }
        }),
        ...linesToParagraphs(thesis, 120),
        new docx.Paragraph({ spacing: { after: 160 } })
      );
    }

    // Key takeaways
    children.push(
      new docx.Paragraph({
        children: [ new docx.TextRun({ text: "Key takeaways", bold:true, size: 24, font:"Calibri" }) ],
        spacing: { after: 120 }
      }),
      ...bulletsFromLines(keyTakeaways),
      new docx.Paragraph({ spacing: { after: 160 } })
    );

    // Analysis
    children.push(
      new docx.Paragraph({
        children: [ new docx.TextRun({ text: "Analysis", bold:true, size: 24, font:"Calibri" }) ],
        spacing: { after: 120 }
      }),
      ...linesToParagraphs(analysis, 140)
    );

    // Additional content
    if ((content || "").trim()){
      children.push(
        new docx.Paragraph({ spacing: { after: 140 } }),
        new docx.Paragraph({
          children: [ new docx.TextRun({ text: "Additional content", bold:true, size: 22, font:"Calibri" }) ],
          spacing: { after: 100 }
        }),
        ...linesToParagraphs(content, 140)
      );
    }

    // Cordoba view (internal only)
    if (includeCordobaView && (cordobaView || "").trim()){
      children.push(
        new docx.Paragraph({ spacing: { after: 140 } }),
        new docx.Paragraph({
          children: [ new docx.TextRun({ text: "The Cordoba view", bold:true, size: 22, font:"Calibri" }) ],
          spacing: { after: 100 }
        }),
        ...linesToParagraphs(cordobaView, 140)
      );
    }

    // Figures
    const imageParagraphs = await addImages(imageFiles);
    if (imageParagraphs.length){
      children.push(
        new docx.Paragraph({ spacing: { before: 240, after: 120 } }),
        new docx.Paragraph({
          children: [ new docx.TextRun({ text: "Figures & charts", bold:true, size: 24, font:"Calibri" }) ],
          spacing: { after: 120 }
        }),
        ...imageParagraphs
      );
    }

    const doc = new docx.Document({
      styles: {
        default: {
          document: {
            run: { font: "Calibri", size: 22, color: "000000" },
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
        headers: { default: new docx.Header({ children: headerChildren }) },
        footers: { default: footer },
        children
      }]
    });

    return doc;
  }

  // ================================
  // Email payload (subject includes version + change note tag)
  // ================================
  function ccForNoteType(noteTypeRaw){
    const t = (noteTypeRaw || "").toLowerCase();
    if (t.includes("equity")) return "tommaso@cordobarg.com";
    if (t.includes("macro") || t.includes("market")) return "tim@cordobarg.com";
    if (t.includes("commodity")) return "uhayd@cordobarg.com";
    return "";
  }

  function presetSubjectTag(preset){
    if (preset === "internal") return "[INTERNAL]";
    if (preset === "public") return "[PUBLIC]";
    if (preset === "client") return "[CLIENT]";
    return "";
  }

  // IMPORTANT FIX:
  // - Avoid URLSearchParams (it turns spaces into "+")
  // - Use encodeURIComponent directly
  // - Force CRLF
  function buildMailto(to, cc, subject, body){
    const crlfBody = (body || "").replace(/\n/g, "\r\n");
    const parts = [];
    if (cc) parts.push(`cc=${encodeURIComponent(cc)}`);
    parts.push(`subject=${encodeURIComponent(subject || "")}`);
    parts.push(`body=${encodeURIComponent(crlfBody)}`);
    return `mailto:${encodeURIComponent(to)}?${parts.join("&")}`;
  }

  function buildCrgEmailPayload(){
    const noteType = ($("noteType")?.value || "Research Note").trim();
    const title = ($("title")?.value || "").trim();
    const topic = ($("topic")?.value || "").trim();
    const thesis = ($("thesis")?.value || "").trim();
    const status = getStatus();
    const reviewedBy = getReviewedBy();
    const preset = getPreset();
    const v = formatVersion(getVersion());
    const changeNote = ($("changeNote")?.value || "").trim();

    const ticker = ($("ticker")?.value || "").trim();
    const crgRating = ($("crgRating")?.value || "").trim();
    const targetPrice = ($("targetPrice")?.value || "").trim();

    const now = new Date();
    const dateShort = formatDateShort(now);
    const dateLong = formatDateTime(now);

    const subject = [
      presetSubjectTag(preset),
      noteType,
      v,
      dateShort,
      title ? `— ${title}` : ""
 ""
    ].filter(Boolean).join(" ");

    const authorLine = [($("authorFirstName")?.value || "").trim(), ($("authorLastName")?.value || "").trim()]
      .filter(Boolean).join(" ").trim();

    const lines = [];
    lines.push("Hi CRG Research,");
    lines.push("Please find my most recent note attached.");

    const meta = [
      `Note type: ${noteType || "N/A"}`,
      title ? `Title: ${title}` : null,
      topic ? `Topic: ${topic}` : null,
      thesis ? `Thesis: ${thesis}` : null,
      `Version: ${v}`,
      `Status: ${status}`,
      reviewedBy ? `Reviewed by: ${reviewedBy}` : null,
      changeNote ? `Change note: ${changeNote}` : null,
      ticker ? `Ticker (Stooq): ${ticker}` : null,
      crgRating ? `CRG Rating: ${crgRating}` : null,
      targetPrice ? `Target Price: ${targetPrice}` : null,
      `Generated: ${dateLong}`
    ].filter(Boolean);

    lines.push(meta.join("\n"));
    lines.push("Best,");
    lines.push(authorLine || "");

    // routing: internal preset uses desk CC; public/client remove CC by default
    const cc = (preset === "internal") ? ccForNoteType(noteType) : "";

    return { subject, body: lines.join("\n\n"), cc };
  }

  $("emailToCrgBtn")?.addEventListener("click", () => {
    const { subject, body, cc } = buildCrgEmailPayload();
    const to = "research@cordobarg.com";
    window.location.href = buildMailto(to, cc, subject, body);
  });

  // ================================
  // Reset
  // ================================
  function clearChartUI(){
    setText("currentPrice","—");
    setText("realisedVol","—");
    setText("rangeReturn","—");
    setText("upsideToTarget","—");
    if (chartStatus) chartStatus.textContent = "";
    if (priceChart){
      try{ priceChart.destroy(); } catch(_) {}
      priceChart = null;
    }
    priceChartImageBytes = null;
    equityStats = { currentPrice:null, realisedVolAnn:null, rangeReturn:null };
  }

  $("resetFormBtn")?.addEventListener("click", () => {
    const ok = confirm("Reset the form? This will clear all fields on this page.");
    if (!ok) return;

    const form = $("researchForm");
    form?.reset();

    // Clear dynamic coauthors
    if ($("coAuthorsList")) $("coAuthorsList").innerHTML = "";
    coAuthorCount = 0;

    // Clear file inputs
    if ($("modelFiles")) $("modelFiles").value = "";
    if ($("imageUpload")) $("imageUpload").value = "";

    updateAttachmentSummary();
    clearChartUI();

    // reset version to v1.0
    setVersion({ major: 1, minor: 0 });

    // preset default
    setPreset("internal");

    // status default
    if ($("workflowStatus")) $("workflowStatus").value = "Draft";
    if ($("reviewedBy")) $("reviewedBy").value = "";
    if ($("reviewedByOther")) $("reviewedByOther").value = "";
    toggleReviewedByOther();

    // resync phone
    syncPrimaryPhone();

    // clear autosave too (system reset)
    clearAutosave();

    toggleEquitySection();
    paintWordCounts();
    updateCompletionMeter();

    const msg = $("message");
    if (msg){
      msg.className = "message";
      msg.textContent = "";
      msg.style.display = "none";
    }
  });

  $("clearAutosaveBtn")?.addEventListener("click", () => {
    const ok = confirm("Clear autosaved draft? This cannot be undone.");
    if (!ok) return;
    clearAutosave();
  });

  $("jumpMissingBtn")?.addEventListener("click", () => jumpToFirstMissing());

  // ================================
  // Main form submission (export)
  // ================================
  $("researchForm")?.addEventListener("submit", async (e) => {
    e.preventDefault();

    const gate = canExport();
    if (!gate.ok){
      alert(gate.reason);
      updateCompletionMeter();
      return;
    }

    const btn = $("generateDocBtn");
    const messageDiv = $("message");

    if (btn){
      btn.disabled = true;
      btn.classList.add("loading");
      btn.textContent = "Generating…";
    }
    if (messageDiv){
      messageDiv.className = "message";
      messageDiv.textContent = "";
      messageDiv.style.display = "block";
    }

    try{
      if (typeof docx === "undefined") throw new Error("docx library not loaded. Refresh the page.");
      if (typeof saveAs === "undefined") throw new Error("FileSaver library not loaded. Refresh the page.");

      const now = new Date();
      const dateTimeString = formatDateTime(now);

      // Version string (current)
      const vStr = formatVersion(getVersion());

      // Workflow
      const status = getStatus();
      const reviewedBy = getReviewedBy();
      const preset = getPreset();

      // Core fields
      const template = $("noteTemplate")?.value || "";
      const noteType = $("noteType")?.value || "";
      const title = $("title")?.value || "";
      const topic = $("topic")?.value || "";
      const thesis = $("thesis")?.value || "";

      const authorLastName = $("authorLastName")?.value || "";
      const authorFirstName = $("authorFirstName")?.value || "";
      const authorPhone = $("authorPhone")?.value || "";
      const authorPhonePrintable = naIfBlank(authorPhone);

      const changeNote = $("changeNote")?.value || "";

      const execSummary = $("executiveSummary")?.value || "";

      const analysis = $("analysis")?.value || "";
      const keyTakeaways = $("keyTakeaways")?.value || "";
      const content = $("content")?.value || "";
      const cordobaView = $("cordobaView")?.value || "";

      const imageFiles = $("imageUpload")?.files || [];

      // Equity fields
      const ticker = $("ticker")?.value || "";
      const crgRating = $("crgRating")?.value || "";
      const targetPrice = $("targetPrice")?.value || "";
      const valuationSummary = $("valuationSummary")?.value || "";
      const keyAssumptions = $("keyAssumptions")?.value || "";
      const scenarioNotes = $("scenarioNotes")?.value || "";
      const modelFiles = $("modelFiles")?.files || null;
      const modelLink = $("modelLink")?.value || "";
      const chartSource = getChartSource();
      const chartAnnotation = $("chartAnnotation")?.value || "";

      // Coauthors
      const coAuthors = [];
      document.querySelectorAll(".coauthor-entry").forEach(entry => {
        const lastName = entry.querySelector(".coauthor-lastname")?.value || "";
        const firstName = entry.querySelector(".coauthor-firstname")?.value || "";
        const phone = entry.querySelector(".coauthor-phone")?.value || "";
        if (lastName && firstName) coAuthors.push({ lastName, firstName, phone: naIfBlank(phone) });
      });

      // Scenario table rows
      const scenarioRows = buildScenarioRows();

      // Build doc
      const doc = await createDocument({
        template,
        noteType, status, reviewedBy, distributionPreset: preset,
        versionString: vStr,
        changeNote,
        title, topic, thesis,
        authorLastName, authorFirstName, authorPhonePrintable,
        coAuthors,
        execSummary,
        analysis, keyTakeaways, content, cordobaView,
        ticker, crgRating, targetPrice,
        equityStats,
        priceChartImageBytes,
        chartSource,
        chartAnnotation,
        scenarioRows,
        valuationSummary, keyAssumptions, scenarioNotes,
        modelFiles, modelLink,
        imageFiles,
        dateTimeString
      });

      const blob = await docx.Packer.toBlob(doc);

      // filename includes version + status + preset
      const safeTitle = (title || "note").replace(/[^a-z0-9]/gi, "_").toLowerCase();
      const safeType = (noteType || "research").replace(/\s+/g, "_").toLowerCase();
      const fileName = `${safeTitle}_${safeType}_${vStr}_${status.toLowerCase()}_${preset}.docx`;

      saveAs(blob, fileName);

      if (messageDiv){
        messageDiv.className = "message success";
        messageDiv.textContent = `✓ Exported: ${fileName}`;
      }

      // Auto increment version AFTER a successful export
      const majorBump = !!$("majorBumpNext")?.checked;
      bumpVersion({ majorBump });
      if ($("majorBumpNext")) $("majorBumpNext").checked = false;

      // Keep change note as-is (audit trail) — analyst can update manually
      autosaveNow();

    } catch(err){
      console.error(err);
      if (messageDiv){
        messageDiv.className = "message error";
        messageDiv.textContent = `✗ Error: ${err.message}`;
      }
    } finally {
      if (btn){
        btn.disabled = false;
        btn.classList.remove("loading");
        btn.textContent = "Generate Word Document";
      }
      updateCompletionMeter();
    }
  });

  // ================================
  // Wire global listeners (autosave + UI refresh)
  // ================================
  const listenIds = [
    "noteTemplate","noteType","workflowStatus","reviewedBy","reviewedByOther","distributionPreset",
    "changeNote","title","topic","thesis","executiveSummary",
    "ticker","crgRating","targetPrice","chartRange","chartDataSource","chartDataSourceOther","chartAnnotation",
    "valuationSummary","keyAssumptions","scenarioNotes","modelLink",
    "bearPrice","bearKey","basePrice","baseKey","bullPrice","bullKey",
    "keyTakeaways","analysis","content","cordobaView",
    "confirmDraft","majorBumpNext"
  ];

  listenIds.forEach(id => {
    const el = $(id);
    if (!el) return;
    el.addEventListener("input", () => {
      scheduleAutosave();
      updateCompletionMeter();
      paintWordCounts();
      maybeGenerateExecutiveSummary();
    });
    el.addEventListener("change", () => {
      scheduleAutosave();
      updateCompletionMeter();
      paintWordCounts();
      maybeGenerateExecutiveSummary();
    });
  });

  $("workflowStatus")?.addEventListener("change", () => {
    updateCompletionMeter();
    scheduleAutosave();
  });

  $("reviewedBy")?.addEventListener("change", () => {
    toggleReviewedByOther();
    updateCompletionMeter();
    scheduleAutosave();
  });

  $("noteType")?.addEventListener("change", () => {
    toggleEquitySection();
    updateCompletionMeter();
    maybeGenerateExecutiveSummary();
  });

  $("noteTemplate")?.addEventListener("change", () => {
    applyTemplate($("noteTemplate").value);
  });

  // Version buttons
  $("bumpMinorBtn")?.addEventListener("click", () => bumpVersion({ majorBump:false }));
  $("bumpMajorBtn")?.addEventListener("click", () => bumpVersion({ majorBump:true }));

  // ================================
  // Boot
  // ================================
  // Restore draft if present
  try{
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw){
      const snap = JSON.parse(raw);
      applyFormSnapshot(snap);
      setAutosaveStatus("Restored draft");
    } else {
      setAutosaveStatus("No draft");
    }
  } catch(e){
    console.warn("Failed to restore draft:", e);
    setAutosaveStatus("Draft restore error");
  }

  // initialise defaults if empty
  if (!$("distributionPreset")?.value) setPreset("internal");
  paintPresetPill();
  paintVersionPill();

  toggleReviewedByOther();
  toggleEquitySection();
  updateAttachmentSummary();
  syncPrimaryPhone();
  paintWordCounts();
  updateCompletionMeter();

  // session timer
  const sessionStart = Date.now();
  setInterval(() => {
    const el = $("sessionTime");
    if (!el) return;
    const s = Math.floor((Date.now() - sessionStart) / 1000);
    const mm = String(Math.floor(s / 60)).padStart(2, "0");
    const ss = String(s % 60).padStart(2, "0");
    el.textContent = `${mm}:${ss}`;
  }, 1000);

});
