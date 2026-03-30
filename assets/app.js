console.log("CRG Research Production Console loaded");

document.addEventListener("DOMContentLoaded", () => {
  const STORAGE_KEY = "crg-rdt-draft-v2";
  const COUNTRY_CODES = [
    { value: "44", label: "+44 UK" },
    { value: "1", label: "+1 US" },
    { value: "353", label: "+353 IE" },
    { value: "33", label: "+33 FR" },
    { value: "49", label: "+49 DE" },
    { value: "31", label: "+31 NL" },
    { value: "34", label: "+34 ES" },
    { value: "39", label: "+39 IT" },
    { value: "971", label: "+971 AE" },
    { value: "966", label: "+966 SA" },
    { value: "92", label: "+92 PK" },
    { value: "880", label: "+880 BD" },
    { value: "91", label: "+91 IN" },
    { value: "234", label: "+234 NG" },
    { value: "254", label: "+254 KE" },
    { value: "27", label: "+27 ZA" },
    { value: "995", label: "+995 GE" },
    { value: "", label: "Other" }
  ];

  const FIELD_LABELS = {
    noteType: "Note type",
    distribution: "Distribution",
    publicationDate: "Publication date",
    deck: "Deck / standfirst",
    title: "Research title",
    topic: "Topic / coverage angle",
    authorLastName: "Primary author last name",
    authorFirstName: "Primary author first name",
    ticker: "Ticker / company",
    crgRating: "CRG rating",
    targetPrice: "Target price",
    keyTakeaways: "Executive summary / key takeaways",
    analysis: "Analysis and commentary",
    cordobaView: "The Cordoba View"
  };

  const FIELD_SECTION = {
    noteType: "Brief",
    distribution: "Brief",
    publicationDate: "Brief",
    deck: "Brief",
    title: "Brief",
    topic: "Brief",
    authorLastName: "Authors",
    authorFirstName: "Authors",
    ticker: "Equity",
    crgRating: "Equity",
    targetPrice: "Equity",
    keyTakeaways: "Research",
    analysis: "Research",
    cordobaView: "Research"
  };

  const BASE_REQUIRED_IDS = [
    "noteType",
    "distribution",
    "publicationDate",
    "title",
    "deck",
    "topic",
    "authorLastName",
    "authorFirstName",
    "keyTakeaways",
    "analysis",
    "cordobaView"
  ];

  const EQUITY_REQUIRED_IDS = ["ticker", "crgRating", "targetPrice"];

  const SECTION_REQUIREMENTS = {
    note: ["noteType", "distribution", "publicationDate", "title", "deck", "topic"],
    authors: ["authorLastName", "authorFirstName"],
    equity: EQUITY_REQUIRED_IDS,
    body: ["keyTakeaways", "analysis", "cordobaView"]
  };

  const form = document.getElementById("researchForm");
  if (!form) return;

  const dom = {
    form,
    noteType: document.getElementById("noteType"),
    distribution: document.getElementById("distribution"),
    publicationDate: document.getElementById("publicationDate"),
    deskLine: document.getElementById("deskLine"),
    title: document.getElementById("title"),
    deck: document.getElementById("deck"),
    topic: document.getElementById("topic"),
    authorLastName: document.getElementById("authorLastName"),
    authorFirstName: document.getElementById("authorFirstName"),
    authorPhoneCountry: document.getElementById("authorPhoneCountry"),
    authorPhoneNational: document.getElementById("authorPhoneNational"),
    authorPhone: document.getElementById("authorPhone"),
    coAuthorsList: document.getElementById("coAuthorsList"),
    addCoAuthor: document.getElementById("addCoAuthor"),
    equitySection: document.getElementById("section-equity"),
    ticker: document.getElementById("ticker"),
    crgRating: document.getElementById("crgRating"),
    targetPrice: document.getElementById("targetPrice"),
    chartRange: document.getElementById("chartRange"),
    fetchPriceChart: document.getElementById("fetchPriceChart"),
    chartStatus: document.getElementById("chartStatus"),
    priceChartCanvas: document.getElementById("priceChart"),
    currentPrice: document.getElementById("currentPrice"),
    realisedVol: document.getElementById("realisedVol"),
    rangeReturn: document.getElementById("rangeReturn"),
    upsideToTarget: document.getElementById("upsideToTarget"),
    modelFiles: document.getElementById("modelFiles"),
    modelLink: document.getElementById("modelLink"),
    valuationSummary: document.getElementById("valuationSummary"),
    keyAssumptions: document.getElementById("keyAssumptions"),
    scenarioNotes: document.getElementById("scenarioNotes"),
    keyTakeaways: document.getElementById("keyTakeaways"),
    analysis: document.getElementById("analysis"),
    content: document.getElementById("content"),
    cordobaView: document.getElementById("cordobaView"),
    imageUpload: document.getElementById("imageUpload"),
    modelSummaryHead: document.getElementById("modelSummaryHead"),
    modelSummaryList: document.getElementById("modelSummaryList"),
    imageSummaryHead: document.getElementById("imageSummaryHead"),
    imageSummaryList: document.getElementById("imageSummaryList"),
    completionBar: document.getElementById("completionBar"),
    completionText: document.getElementById("completionText"),
    readinessPercent: document.getElementById("readinessPercent"),
    readinessCaption: document.getElementById("readinessCaption"),
    noteStateChip: document.getElementById("noteStateChip"),
    missingFields: document.getElementById("missingFields"),
    draftStatus: document.getElementById("draftStatus"),
    previewFileName: document.getElementById("previewFileName"),
    renderTimestamp: document.getElementById("renderTimestamp"),
    summaryType: document.getElementById("summaryType"),
    summaryTopic: document.getElementById("summaryTopic"),
    summaryAuthor: document.getElementById("summaryAuthor"),
    summaryCoAuthors: document.getElementById("summaryCoAuthors"),
    summaryOutput: document.getElementById("summaryOutput"),
    summaryOutputDetail: document.getElementById("summaryOutputDetail"),
    previewTitle: document.getElementById("previewTitle"),
    previewAuthor: document.getElementById("previewAuthor"),
    previewCoverage: document.getElementById("previewCoverage"),
    previewSupport: document.getElementById("previewSupport"),
    generateDocBtn: document.getElementById("generateDocBtn"),
    emailToCrgBtn: document.getElementById("emailToCrgBtn"),
    resetFormBtn: document.getElementById("resetFormBtn"),
    message: document.getElementById("message"),
    navNote: document.getElementById("nav-note"),
    navAuthors: document.getElementById("nav-authors"),
    navEquity: document.getElementById("nav-equity"),
    navBody: document.getElementById("nav-body"),
    navExhibits: document.getElementById("nav-exhibits"),
    navOutput: document.getElementById("nav-output")
  };

  const state = {
    coAuthorCount: 0,
    priceChart: null,
    priceChartImageBytes: null,
    equityStats: {
      currentPrice: null,
      realisedVolAnn: null,
      rangeReturn: null
    },
    saveTimer: null,
    lastSavedAt: null
  };

  const draftFieldIds = [
    "noteType",
    "distribution",
    "publicationDate",
    "deskLine",
    "title",
    "deck",
    "topic",
    "authorLastName",
    "authorFirstName",
    "authorPhoneCountry",
    "authorPhoneNational",
    "authorPhone",
    "ticker",
    "crgRating",
    "targetPrice",
    "modelLink",
    "valuationSummary",
    "keyAssumptions",
    "scenarioNotes",
    "keyTakeaways",
    "analysis",
    "content",
    "cordobaView",
    "chartRange"
  ];

  init();

  function init() {
    wireSectionNavigation();
    wirePrimaryPhone();
    wireFormEvents();
    restoreDraft();
    ensurePublicationDate();
    ensureDeskLineDefault();
    syncPrimaryPhone();
    toggleEquitySection();
    updateFileSummary(dom.modelFiles, dom.modelSummaryHead, dom.modelSummaryList, "No supporting files attached.");
    updateFileSummary(dom.imageUpload, dom.imageSummaryHead, dom.imageSummaryList, "No figures attached.");
    updateAllUI();
    checkLibraries();
  }

  function wireSectionNavigation() {
    const navButtons = Array.from(document.querySelectorAll(".section-nav-link"));
    navButtons.forEach((button) => {
      button.addEventListener("click", () => {
        const targetId = button.getAttribute("data-target");
        const target = document.getElementById(targetId);
        if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    });

    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (!entry.isIntersecting) return;
          navButtons.forEach((button) => {
            button.classList.toggle("is-active", button.getAttribute("data-target") === entry.target.id);
          });
        });
      },
      {
        rootMargin: "-30% 0px -55% 0px",
        threshold: 0
      }
    );

    document.querySelectorAll(".section-card").forEach((section) => observer.observe(section));
  }

  function wirePrimaryPhone() {
    if (!dom.authorPhoneNational || !dom.authorPhoneCountry || !dom.authorPhone) return;

    dom.authorPhoneNational.addEventListener("input", () => {
      formatPhoneInput(dom.authorPhoneNational);
      syncPrimaryPhone();
      updateAllUI();
      queueDraftSave();
    });

    dom.authorPhoneCountry.addEventListener("change", () => {
      syncPrimaryPhone();
      updateAllUI();
      queueDraftSave();
    });
  }

  function wireFormEvents() {
    dom.noteType.addEventListener("change", () => {
      ensureDeskLineDefault();
      toggleEquitySection();
      resetChartState({ keepStatusText: false });
      updateAllUI();
      queueDraftSave();
    });

    dom.deskLine.addEventListener("input", () => {
      dom.deskLine.dataset.autofill = "false";
      updateAllUI();
      queueDraftSave();
    });

    dom.targetPrice.addEventListener("input", () => {
      updateUpsideDisplay();
      updateAllUI();
      queueDraftSave();
    });

    dom.modelFiles.addEventListener("change", () => {
      updateFileSummary(dom.modelFiles, dom.modelSummaryHead, dom.modelSummaryList, "No supporting files attached.");
      updateAllUI();
    });

    dom.imageUpload.addEventListener("change", () => {
      updateFileSummary(dom.imageUpload, dom.imageSummaryHead, dom.imageSummaryList, "No figures attached.");
      updateAllUI();
    });

    dom.addCoAuthor.addEventListener("click", () => {
      addCoAuthorCard();
      updateAllUI();
      queueDraftSave();
    });

    dom.coAuthorsList.addEventListener("click", (event) => {
      const removeButton = event.target.closest(".remove-coauthor");
      if (!removeButton) return;
      const card = removeButton.closest(".coauthor-card");
      if (card) {
        card.remove();
        updateAllUI();
        queueDraftSave();
      }
    });

    dom.fetchPriceChart.addEventListener("click", buildPriceChart);
    dom.emailToCrgBtn.addEventListener("click", draftEmailToResearch);
    dom.resetFormBtn.addEventListener("click", resetDraft);

    form.addEventListener("input", (event) => {
      if (event.target instanceof HTMLInputElement && event.target.type === "file") return;
      if (event.target.id !== "authorPhoneNational") updateAllUI();
      queueDraftSave();
    });

    form.addEventListener("change", (event) => {
      if (event.target instanceof HTMLInputElement && event.target.type === "file") return;
      updateAllUI();
      queueDraftSave();
    });

    form.addEventListener("submit", handleSubmit);
  }

  function digitsOnly(value) {
    return String(value || "").replace(/\D/g, "");
  }

  function todayIsoDate() {
    const now = new Date();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${now.getFullYear()}-${month}-${day}`;
  }

  function defaultDeskLine(noteType) {
    const map = {
      "General Note": "Strategy - Global",
      "Equity Research": "Equity Research - Global",
      "Macro Research": "Macro Strategy - Global",
      "Fixed Income Research": "Fixed Income - Global",
      "Commodity Insights": "Commodities - Global"
    };

    return noteType ? (map[noteType] || "Research - Global") : "";
  }

  function ensurePublicationDate() {
    if (!dom.publicationDate.value) dom.publicationDate.value = todayIsoDate();
  }

  function ensureDeskLineDefault(force = false) {
    const current = dom.deskLine.value.trim();
    const currentAuto = dom.deskLine.dataset.autofill !== "false";
    if (force || !current || currentAuto) {
      dom.deskLine.value = defaultDeskLine(dom.noteType.value);
      dom.deskLine.dataset.autofill = "true";
    }
  }

  function formatNationalLoose(rawValue) {
    const digits = digitsOnly(rawValue);
    if (!digits) return "";

    const parts = [digits.slice(0, 4), digits.slice(4, 7), digits.slice(7, 10), digits.slice(10)];
    return parts.filter(Boolean).join(" ");
  }

  function buildInternationalHyphen(countryCode, nationalNumber) {
    const cc = digitsOnly(countryCode);
    const nn = digitsOnly(nationalNumber);

    if (!cc && !nn) return "";
    if (!cc) return nn;
    if (!nn) return `${cc}-`;
    return `${cc}-${nn}`;
  }

  function formatPhoneInput(input) {
    const caretStart = input.selectionStart || input.value.length;
    const beforeLength = input.value.length;
    input.value = formatNationalLoose(input.value);
    const afterLength = input.value.length;
    const offset = afterLength - beforeLength;
    const nextPosition = Math.max(0, caretStart + offset);
    input.setSelectionRange(nextPosition, nextPosition);
  }

  function syncPrimaryPhone() {
    dom.authorPhone.value = buildInternationalHyphen(dom.authorPhoneCountry.value, dom.authorPhoneNational.value);
  }

  function createCountryOptionsHtml(selectedValue) {
    return COUNTRY_CODES.map((option) => {
      const selected = option.value === (selectedValue ?? "44") ? " selected" : "";
      return `<option value="${option.value}"${selected}>${option.label}</option>`;
    }).join("");
  }

  function addCoAuthorCard(seed = {}) {
    state.coAuthorCount += 1;

    const card = document.createElement("div");
    card.className = "coauthor-card";
    card.dataset.id = String(state.coAuthorCount);
    card.innerHTML = `
      <div class="field">
        <label>Last Name</label>
        <input type="text" class="coauthor-lastname" placeholder="Surname" value="${escapeAttribute(seed.lastName || "")}">
      </div>
      <div class="field">
        <label>First Name</label>
        <input type="text" class="coauthor-firstname" placeholder="Given name" value="${escapeAttribute(seed.firstName || "")}">
      </div>
      <div class="field">
        <label>Phone</label>
        <div class="coauthor-phone-row">
          <select class="coauthor-country" aria-label="Co-author country code">
            ${createCountryOptionsHtml(seed.countryCode || "44")}
          </select>
          <input type="text" class="coauthor-phone-local" inputmode="numeric" placeholder="National number" value="${escapeAttribute(seed.phoneLocal || "")}">
        </div>
        <input type="hidden" class="coauthor-phone" value="${escapeAttribute(seed.phone || "")}">
      </div>
      <div class="field">
        <label>&nbsp;</label>
        <button type="button" class="btn btn-ghost remove-coauthor">Remove</button>
      </div>
    `;

    const localInput = card.querySelector(".coauthor-phone-local");
    const countrySelect = card.querySelector(".coauthor-country");
    const hiddenInput = card.querySelector(".coauthor-phone");

    const sync = () => {
      hiddenInput.value = buildInternationalHyphen(countrySelect.value, localInput.value);
    };

    localInput.addEventListener("input", () => {
      formatPhoneInput(localInput);
      sync();
      updateAllUI();
      queueDraftSave();
    });

    countrySelect.addEventListener("change", () => {
      sync();
      updateAllUI();
      queueDraftSave();
    });

    card.querySelectorAll("input[type='text']").forEach((input) => {
      input.addEventListener("input", () => {
        updateAllUI();
        queueDraftSave();
      });
    });

    sync();
    dom.coAuthorsList.appendChild(card);
  }

  function getCoAuthors() {
    return Array.from(dom.coAuthorsList.querySelectorAll(".coauthor-card"))
      .map((card) => {
        const lastName = card.querySelector(".coauthor-lastname").value.trim();
        const firstName = card.querySelector(".coauthor-firstname").value.trim();
        const countryCode = card.querySelector(".coauthor-country").value;
        const phoneLocal = card.querySelector(".coauthor-phone-local").value.trim();
        const phone = buildInternationalHyphen(countryCode, phoneLocal);

        return { lastName, firstName, phone, countryCode, phoneLocal };
      })
      .filter((coAuthor) => coAuthor.lastName || coAuthor.firstName || coAuthor.phone);
  }

  function isEquitySelected() {
    return dom.noteType.value === "Equity Research";
  }

  function toggleEquitySection() {
    const showEquity = isEquitySelected();
    dom.equitySection.hidden = !showEquity;
    dom.navEquity.textContent = showEquity ? buildSectionCompletion("equity") : "Optional";
  }

  function getRequiredIds() {
    return isEquitySelected() ? BASE_REQUIRED_IDS.concat(EQUITY_REQUIRED_IDS) : BASE_REQUIRED_IDS;
  }

  function isFilled(element) {
    if (!element) return false;
    if (element instanceof HTMLInputElement && element.type === "file") {
      return Array.from(element.files || []).length > 0;
    }
    return String(element.value || "").trim().length > 0;
  }

  function validateForm(showErrors = false) {
    const missing = [];
    const requiredIds = getRequiredIds();

    requiredIds.forEach((id) => {
      const element = document.getElementById(id);
      let valid = isFilled(element);

      if (valid && id === "targetPrice") {
        valid = parseNumber(element.value) != null;
      }

      if (!valid) {
        missing.push({
          id,
          label: id === "targetPrice" && isFilled(element) ? "Target price must be numeric" : (FIELD_LABELS[id] || id),
          section: FIELD_SECTION[id] || "General"
        });
      }

      if (showErrors && element) element.classList.toggle("is-invalid", !valid);
      if (!showErrors && element && valid) element.classList.remove("is-invalid");
    });

    return {
      valid: missing.length === 0,
      missing,
      total: requiredIds.length,
      complete: requiredIds.length - missing.length,
      percent: requiredIds.length ? Math.round(((requiredIds.length - missing.length) / requiredIds.length) * 100) : 0
    };
  }

  function buildSectionCompletion(sectionKey) {
    const ids = SECTION_REQUIREMENTS[sectionKey] || [];
    if (sectionKey === "equity" && !isEquitySelected()) return "Optional";
    const complete = ids.filter((id) => {
      const element = document.getElementById(id);
      if (!isFilled(element)) return false;
      if (id === "targetPrice") return parseNumber(element.value) != null;
      return true;
    }).length;
    return `${complete}/${ids.length}`;
  }

  function updateSectionPills() {
    dom.navNote.textContent = buildSectionCompletion("note");
    dom.navAuthors.textContent = buildSectionCompletion("authors");
    dom.navEquity.textContent = buildSectionCompletion("equity");
    dom.navBody.textContent = buildSectionCompletion("body");

    const supportCount = Array.from(dom.modelFiles.files || []).length + Array.from(dom.imageUpload.files || []).length;
    dom.navExhibits.textContent = supportCount ? `${supportCount} files` : "Optional";
    dom.navOutput.textContent = validateForm(false).valid ? "Ready" : "Draft";
  }

  function updateCompletion() {
    const validation = validateForm(false);
    dom.completionBar.style.width = `${validation.percent}%`;
    dom.completionText.textContent = `${validation.complete} / ${validation.total} required fields complete`;
    dom.readinessPercent.textContent = `${validation.percent}%`;
    dom.noteStateChip.textContent = validation.valid ? "Ready" : validation.percent >= 55 ? "In Build" : "Draft";
    dom.noteStateChip.style.background = validation.valid ? "rgba(37, 115, 75, 0.12)" : validation.percent >= 55 ? "rgba(132, 95, 15, 0.10)" : "rgba(109, 118, 130, 0.12)";
    dom.noteStateChip.style.color = validation.valid ? "#25734b" : validation.percent >= 55 ? "#845F0F" : "#5d6570";
    dom.noteStateChip.style.borderColor = validation.valid ? "rgba(37, 115, 75, 0.2)" : validation.percent >= 55 ? "rgba(132, 95, 15, 0.22)" : "rgba(109, 118, 130, 0.16)";

    const progressTrack = dom.completionBar.parentElement;
    if (progressTrack) progressTrack.setAttribute("aria-valuenow", String(validation.percent));

    if (validation.valid) {
      dom.readinessCaption.textContent = "The brief, narrative, and author metadata are complete. The Word note can now render cleanly in the publication template.";
    } else if (validation.percent >= 55) {
      dom.readinessCaption.textContent = "The draft is structurally taking shape. Complete the remaining mandatory fields before export.";
    } else {
      dom.readinessCaption.textContent = "Complete the brief, byline, and core body sections before generating the Word note.";
    }

    renderMissingFields(validation.missing);
  }

  function renderMissingFields(missing) {
    dom.missingFields.innerHTML = "";

    if (!missing.length) {
      const item = document.createElement("li");
      item.textContent = "No blocking fields. The note is structurally ready.";
      dom.missingFields.appendChild(item);
      return;
    }

    missing.slice(0, 6).forEach((entry) => {
      const item = document.createElement("li");
      item.textContent = entry.label;
      const meta = document.createElement("span");
      meta.textContent = entry.section;
      item.appendChild(meta);
      dom.missingFields.appendChild(item);
    });
  }

  function updateSummaryCards() {
    const noteType = dom.noteType.value.trim();
    const deskLine = getDeskLine();
    const deck = dom.deck.value.trim();
    const authorLine = buildPrimaryAuthorLine();
    const coAuthors = getCoAuthors();
    const validation = validateForm(false);

    dom.summaryType.textContent = noteType || "Select a note type";
    dom.summaryTopic.textContent = deck || `${deskLine || "Set the desk line"}${dom.publicationDate.value ? ` | ${formatInputDateLabel(dom.publicationDate.value)}` : ""}`;
    dom.summaryAuthor.textContent = authorLine || "Assign primary author";
    dom.summaryCoAuthors.textContent = coAuthors.length
      ? `${coAuthors.length} co-author${coAuthors.length > 1 ? "s" : ""} added.`
      : "No co-authors added.";
    dom.summaryOutput.textContent = validation.valid ? "Publication package ready" : "Draft in progress";

    const attachmentBits = [];
    const modelCount = Array.from(dom.modelFiles.files || []).length;
    const imageCount = Array.from(dom.imageUpload.files || []).length;
    if (isEquitySelected()) attachmentBits.push(state.priceChartImageBytes ? "price chart ready" : "price chart pending");
    if (modelCount) attachmentBits.push(`${modelCount} model file${modelCount > 1 ? "s" : ""}`);
    if (imageCount) attachmentBits.push(`${imageCount} exhibit${imageCount > 1 ? "s" : ""}`);

    dom.summaryOutputDetail.textContent = attachmentBits.length
      ? `Support pack includes ${attachmentBits.join(", ")}.`
      : "Word export and email payload update automatically as the note develops.";
  }

  function updatePreview() {
    const data = collectFormData();
    dom.previewFileName.textContent = buildDocumentFileName(data);
    dom.renderTimestamp.textContent = `Last refreshed ${formatDateTime(new Date())}`;
    dom.previewTitle.textContent = data.title || "No title yet";
    dom.previewAuthor.textContent = buildPrimaryAuthorLine() || "Primary author pending";

    if (isEquitySelected()) {
      const coverageBits = [
        data.deskLine || "Desk line pending",
        data.publicationDate ? formatInputDateLabel(data.publicationDate) : "Date pending",
        data.ticker || "Ticker pending",
        data.crgRating || "Rating pending"
      ];
      dom.previewCoverage.textContent = coverageBits.join(" | ");
    } else {
      const coverageBits = [
        data.deskLine || data.noteType || "Desk line not set",
        data.publicationDate ? formatInputDateLabel(data.publicationDate) : "Date pending",
        data.topic || "Topic pending"
      ];
      dom.previewCoverage.textContent = coverageBits.join(" | ");
    }

    const supportBits = [];
    const modelCount = Array.from(dom.modelFiles.files || []).length;
    const imageCount = Array.from(dom.imageUpload.files || []).length;
    if (state.priceChartImageBytes) supportBits.push("chart");
    if (modelCount) supportBits.push(`${modelCount} model file${modelCount > 1 ? "s" : ""}`);
    if (imageCount) supportBits.push(`${imageCount} image${imageCount > 1 ? "s" : ""}`);
    dom.previewSupport.textContent = supportBits.length ? supportBits.join(" | ") : "No attachments yet";
  }

  function buildPrimaryAuthorLine() {
    const lastName = dom.authorLastName.value.trim();
    const firstName = dom.authorFirstName.value.trim();
    const combined = [firstName, lastName].filter(Boolean).join(" ").trim();
    return combined;
  }

  function getDeskLine() {
    return dom.deskLine.value.trim() || defaultDeskLine(dom.noteType.value);
  }

  function updateAllUI() {
    updateCompletion();
    updateSectionPills();
    updateSummaryCards();
    updatePreview();
    updateUpsideDisplay();
  }

  function queueDraftSave() {
    window.clearTimeout(state.saveTimer);
    state.saveTimer = window.setTimeout(saveDraft, 320);
  }

  function serializeDraft() {
    const values = {};
    draftFieldIds.forEach((id) => {
      const element = document.getElementById(id);
      if (element) values[id] = element.value;
    });

    return {
      values,
      coAuthors: getCoAuthors(),
      savedAt: new Date().toISOString()
    };
  }

  function saveDraft() {
    try {
      const payload = serializeDraft();
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
      state.lastSavedAt = payload.savedAt;
      dom.draftStatus.textContent = `Draft saved locally at ${formatClock(new Date(payload.savedAt))}.`;
    } catch (error) {
      dom.draftStatus.textContent = "Local autosave could not write to storage in this browser session.";
      console.error("Autosave failed:", error);
    }
  }

  function restoreDraft() {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) return;

      const payload = JSON.parse(raw);
      Object.entries(payload.values || {}).forEach(([id, value]) => {
        const element = document.getElementById(id);
        if (!element || (element instanceof HTMLInputElement && element.type === "file")) return;
        element.value = value;
      });

      dom.coAuthorsList.innerHTML = "";
      state.coAuthorCount = 0;
      (payload.coAuthors || []).forEach((coAuthor) => addCoAuthorCard(coAuthor));
      state.lastSavedAt = payload.savedAt || null;
      dom.deskLine.dataset.autofill = dom.deskLine.value.trim() ? "false" : "true";

      if (state.lastSavedAt) {
        dom.draftStatus.textContent = `Draft restored from ${formatDateTime(new Date(state.lastSavedAt))}.`;
      }
    } catch (error) {
      console.error("Draft restore failed:", error);
      dom.draftStatus.textContent = "A saved draft existed but could not be restored cleanly.";
    }
  }

  function resetDraft() {
    const confirmed = window.confirm("Reset the entire draft? Text fields, metadata, and autosaved content will be cleared.");
    if (!confirmed) return;

    form.reset();
    dom.coAuthorsList.innerHTML = "";
    state.coAuthorCount = 0;
    state.lastSavedAt = null;
    syncPrimaryPhone();
    resetChartState({ keepStatusText: false });
    window.localStorage.removeItem(STORAGE_KEY);
    dom.draftStatus.textContent = "Autosave idle. Drafts are stored locally in this browser.";
    ensurePublicationDate();
    ensureDeskLineDefault(true);
    updateFileSummary(dom.modelFiles, dom.modelSummaryHead, dom.modelSummaryList, "No supporting files attached.");
    updateFileSummary(dom.imageUpload, dom.imageSummaryHead, dom.imageSummaryList, "No figures attached.");
    toggleEquitySection();
    clearMessage();
    updateAllUI();
  }

  function updateFileSummary(input, head, list, emptyText) {
    const files = Array.from(input.files || []);
    head.textContent = files.length
      ? `${files.length} file${files.length > 1 ? "s" : ""} attached.`
      : emptyText;

    list.innerHTML = "";
    if (!files.length) {
      list.hidden = true;
      return;
    }

    files.forEach((file) => {
      const item = document.createElement("li");
      item.textContent = file.name;
      list.appendChild(item);
    });

    list.hidden = false;
  }

  function setMetric(element, value) {
    element.textContent = value == null || value === "" ? "-" : value;
  }

  function resetChartState(options = {}) {
    if (state.priceChart) {
      try {
        state.priceChart.destroy();
      } catch (error) {
        console.warn("Unable to destroy existing chart instance:", error);
      }
    }

    state.priceChart = null;
    state.priceChartImageBytes = null;
    state.equityStats = {
      currentPrice: null,
      realisedVolAnn: null,
      rangeReturn: null
    };

    setMetric(dom.currentPrice, "-");
    setMetric(dom.realisedVol, "-");
    setMetric(dom.rangeReturn, "-");
    setMetric(dom.upsideToTarget, "-");

    if (!options.keepStatusText) dom.chartStatus.textContent = isEquitySelected() ? "No market chart fetched yet." : "";
  }

  function checkLibraries() {
    const issues = [];
    if (typeof window.docx === "undefined") issues.push("The docx export library failed to load.");
    if (typeof window.saveAs === "undefined") issues.push("The file save library failed to load.");
    if (typeof window.Chart === "undefined") issues.push("The charting library failed to load.");

    if (issues.length) setMessage("error", issues.join(" "));
  }

  async function buildPriceChart() {
    if (!isEquitySelected()) return;

    if (typeof window.Chart === "undefined") {
      setMessage("error", "Chart.js is unavailable, so the market chart cannot be rendered in this session.");
      return;
    }

    const ticker = dom.ticker.value.trim();
    if (!ticker) {
      dom.ticker.classList.add("is-invalid");
      dom.chartStatus.textContent = "Enter a ticker before fetching market data.";
      dom.ticker.focus();
      return;
    }

    dom.fetchPriceChart.disabled = true;
    dom.fetchPriceChart.classList.add("loading");
    dom.chartStatus.textContent = "Fetching price history and computing market statistics...";
    clearMessage();

    try {
      const symbol = stooqSymbolFromTicker(ticker);
      const series = await fetchStooqDaily(symbol);
      const startDate = computeStartDate(dom.chartRange.value || "6mo");
      const filtered = series.filter((item) => new Date(item.date) >= startDate);

      if (filtered.length < 10) {
        throw new Error("Not enough price history returned for the selected range.");
      }

      const labels = filtered.map((item) => item.date);
      const values = filtered.map((item) => item.close);
      renderChart(labels, values, `${ticker.toUpperCase()} close`);
      await waitForChartPaint();

      state.priceChartImageBytes = canvasToPngBytes(dom.priceChartCanvas);

      const currentPrice = values[values.length - 1];
      const rangeReturn = values[0] ? (currentPrice / values[0]) - 1 : null;
      const returns = computeDailyReturns(values);
      const dailyVol = standardDeviation(returns);
      const realisedVolAnn = dailyVol == null ? null : dailyVol * Math.sqrt(252);

      state.equityStats = {
        currentPrice,
        rangeReturn,
        realisedVolAnn
      };

      setMetric(dom.currentPrice, currentPrice != null ? currentPrice.toFixed(2) : "-");
      setMetric(dom.rangeReturn, rangeReturn != null ? formatPercent(rangeReturn) : "-");
      setMetric(dom.realisedVol, realisedVolAnn != null ? formatPercent(realisedVolAnn) : "-");
      updateUpsideDisplay();
      dom.chartStatus.textContent = `Market chart ready for ${ticker.toUpperCase()} (${(dom.chartRange.value || "6mo").toUpperCase()}).`;
      updateAllUI();
    } catch (error) {
      resetChartState({ keepStatusText: true });
      dom.chartStatus.textContent = error.message;
      setMessage("error", `Unable to build the market chart: ${error.message}`);
    } finally {
      dom.fetchPriceChart.disabled = false;
      dom.fetchPriceChart.classList.remove("loading");
    }
  }

  function stooqSymbolFromTicker(rawTicker) {
    const ticker = rawTicker.trim().toLowerCase();
    if (!ticker) return "";
    return ticker.includes(".") ? ticker : `${ticker}.us`;
  }

  async function fetchStooqDaily(symbol) {
    const urls = [
      `https://r.jina.ai/http://stooq.com/q/d/l/?s=${encodeURIComponent(symbol)}&i=d`,
      `https://r.jina.ai/http://stooq.pl/q/d/l/?s=${encodeURIComponent(symbol)}&i=d`
    ];

    let lastError = new Error("Market data request failed.");

    for (const url of urls) {
      try {
        const response = await fetchWithTimeout(url, { cache: "no-store" }, 12000);
        if (!response.ok) throw new Error(`Market data request returned ${response.status}.`);

        const rawText = await response.text();
        const csv = extractStooqCsv(rawText);
        if (!csv) throw new Error("The market data feed returned an unexpected response.");

        const rows = csv
          .trim()
          .split("\n")
          .slice(1)
          .map((line) => line.split(","))
          .map((parts) => ({ date: parts[0], close: Number(parts[4]) }))
          .filter((row) => row.date && Number.isFinite(row.close));

        if (rows.length < 5) throw new Error("The market data feed did not return enough observations.");
        return rows;
      } catch (error) {
        lastError = error;
      }
    }

    throw lastError;
  }

  function extractStooqCsv(text) {
    const lines = String(text || "")
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);

    const headerIndex = lines.findIndex((line) => line.toLowerCase().startsWith("date,open,high,low,close,volume"));
    if (headerIndex === -1) return "";
    return lines.slice(headerIndex).join("\n");
  }

  function fetchWithTimeout(resource, options, timeoutMs) {
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), timeoutMs);

    return fetch(resource, {
      ...options,
      signal: controller.signal
    }).finally(() => window.clearTimeout(timeout));
  }

  function computeStartDate(range) {
    const date = new Date();
    if (range === "6mo") date.setMonth(date.getMonth() - 6);
    else if (range === "1y") date.setFullYear(date.getFullYear() - 1);
    else if (range === "2y") date.setFullYear(date.getFullYear() - 2);
    else if (range === "5y") date.setFullYear(date.getFullYear() - 5);
    return date;
  }

  function renderChart(labels, values, label) {
    if (state.priceChart) state.priceChart.destroy();

    state.priceChart = new window.Chart(dom.priceChartCanvas, {
      type: "line",
      data: {
        labels,
        datasets: [
          {
            label,
            data: values,
            borderColor: "#9b7a36",
            backgroundColor: "rgba(155, 122, 54, 0.14)",
            pointRadius: 0,
            borderWidth: 2.2,
            tension: 0.2,
            fill: true
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
          mode: "index",
          intersect: false
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            displayColors: false,
            backgroundColor: "#101b2f",
            titleFont: { family: "Aptos", weight: "600" },
            bodyFont: { family: "Aptos" }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: {
              maxTicksLimit: 7,
              color: "#4e617d"
            }
          },
          y: {
            grid: { color: "rgba(15, 23, 42, 0.08)" },
            ticks: {
              maxTicksLimit: 6,
              color: "#4e617d"
            }
          }
        }
      }
    });
  }

  function waitForChartPaint() {
    return new Promise((resolve) => {
      requestAnimationFrame(() => {
        requestAnimationFrame(resolve);
      });
    });
  }

  function canvasToPngBytes(canvas) {
    const dataUrl = canvas.toDataURL("image/png");
    const base64 = dataUrl.split(",")[1];
    return Uint8Array.from(window.atob(base64), (char) => char.charCodeAt(0));
  }

  function computeDailyReturns(closes) {
    const returns = [];
    for (let index = 1; index < closes.length; index += 1) {
      const previous = closes[index - 1];
      const current = closes[index];
      if (previous > 0 && Number.isFinite(previous) && Number.isFinite(current)) {
        returns.push((current / previous) - 1);
      }
    }
    return returns;
  }

  function standardDeviation(values) {
    if (!values.length) return null;
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const variance = values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(values.length - 1, 1);
    return Math.sqrt(variance);
  }

  function parseNumber(value) {
    const cleaned = String(value || "").replace(/[^0-9.-]/g, "");
    const number = Number(cleaned);
    return Number.isFinite(number) ? number : null;
  }

  function computeUpsideToTarget(currentPrice, targetPrice) {
    if (!Number.isFinite(currentPrice) || !Number.isFinite(targetPrice) || currentPrice === 0) return null;
    return (targetPrice / currentPrice) - 1;
  }

  function updateUpsideDisplay() {
    const targetPrice = parseNumber(dom.targetPrice.value);
    const currentPrice = state.equityStats.currentPrice;
    const upside = computeUpsideToTarget(currentPrice, targetPrice);
    setMetric(dom.upsideToTarget, upside == null ? "-" : formatPercent(upside));
  }

  function formatPercent(value) {
    return `${(value * 100).toFixed(1)}%`;
  }

  function clearMessage() {
    dom.message.className = "message";
    dom.message.textContent = "";
  }

  function setMessage(kind, text) {
    dom.message.className = `message is-visible ${kind === "success" ? "is-success" : "is-error"}`;
    dom.message.textContent = text;
  }

  function draftEmailToResearch() {
    const data = collectFormData();
    const payload = buildCrgEmailPayload(data);
    const mailto = buildMailto("research@cordobarg.com", payload.cc, payload.subject, payload.body);
    window.location.href = mailto;
  }

  function buildMailto(to, cc, subject, body) {
    const parts = [];
    if (cc) parts.push(`cc=${encodeURIComponent(cc)}`);
    parts.push(`subject=${encodeURIComponent(subject)}`);
    parts.push(`body=${encodeURIComponent(body.replace(/\n/g, "\r\n"))}`);
    return `mailto:${encodeURIComponent(to)}?${parts.join("&")}`;
  }

  function buildCrgEmailPayload(data) {
    const cc = ccForNoteType(data.noteType);
    const support = [];
    if (isEquitySelected()) support.push(state.priceChartImageBytes ? "price chart included in export" : "price chart not yet pulled");
    if (data.modelFiles.length) support.push(`${data.modelFiles.length} model file${data.modelFiles.length > 1 ? "s" : ""}`);
    if (data.imageFiles.length) support.push(`${data.imageFiles.length} figure${data.imageFiles.length > 1 ? "s" : ""}`);
    const publicationDate = parseInputDate(data.publicationDate) || data.generatedAt || new Date();

    const subject = [data.noteType || "Research Note", formatDateShort(publicationDate), data.title ? `- ${data.title}` : ""]
      .filter(Boolean)
      .join(" ");

    const lines = [
      "Hi CRG Research,",
      "",
      "Please find the latest research note attached.",
      "",
      `Note type: ${data.noteType || "N/A"}`,
      `Distribution: ${data.distribution || "N/A"}`,
      `Publication date: ${formatDocDate(publicationDate)}`,
      `Desk line: ${data.deskLine || "N/A"}`,
      `Title: ${data.title || "N/A"}`,
      `Deck: ${data.deck || "N/A"}`,
      `Topic: ${data.topic || "N/A"}`,
      data.ticker ? `Ticker / company: ${data.ticker}` : null,
      data.crgRating ? `CRG rating: ${data.crgRating}` : null,
      data.targetPrice ? `Target price: ${data.targetPrice}` : null,
      support.length ? `Support pack: ${support.join(", ")}` : "Support pack: no additional attachments referenced yet",
      `Generated: ${formatDateTime(new Date())}`,
      "",
      "Best,",
      buildPrimaryAuthorLine() || "Analyst"
    ].filter(Boolean);

    return {
      cc,
      subject,
      body: lines.join("\n")
    };
  }

  function ccForNoteType(noteType) {
    const normalized = String(noteType || "").toLowerCase();
    if (normalized.includes("equity")) return "tommaso@cordobarg.com";
    if (normalized.includes("macro") || normalized.includes("market")) return "tim@cordobarg.com";
    if (normalized.includes("commodity")) return "uhayd@cordobarg.com";
    return "";
  }

  async function handleSubmit(event) {
    event.preventDefault();
    clearMessage();

    const validation = validateForm(true);
    if (!validation.valid) {
      const firstMissing = document.getElementById(validation.missing[0].id);
      if (firstMissing) {
        firstMissing.scrollIntoView({ behavior: "smooth", block: "center" });
        firstMissing.focus();
      }
      setMessage("error", `The note is not ready to export. Complete the remaining core fields: ${validation.missing.map((entry) => entry.label).join(", ")}.`);
      return;
    }

    if (typeof window.docx === "undefined" || typeof window.saveAs === "undefined") {
      setMessage("error", "The Word export libraries are unavailable in this browser session. Refresh the page and try again.");
      return;
    }

    dom.generateDocBtn.disabled = true;
    dom.generateDocBtn.classList.add("loading");
    dom.generateDocBtn.textContent = "Generating Document";

    try {
      syncPrimaryPhone();
      const data = collectFormData();
      const documentFileName = buildDocumentFileName(data);
      const doc = await createDocument(data);
      const blob = await window.docx.Packer.toBlob(doc);
      window.saveAs(blob, documentFileName);
      saveDraft();
      setMessage("success", `Document generated successfully as ${documentFileName}.`);
    } catch (error) {
      console.error("Document generation failed:", error);
      setMessage("error", `Document generation failed: ${error.message}`);
    } finally {
      dom.generateDocBtn.disabled = false;
      dom.generateDocBtn.classList.remove("loading");
      dom.generateDocBtn.textContent = "Generate Word Document";
    }
  }

  function collectFormData() {
    syncPrimaryPhone();

    return {
      noteType: dom.noteType.value.trim(),
      distribution: dom.distribution.value.trim(),
      publicationDate: dom.publicationDate.value.trim(),
      deskLine: getDeskLine(),
      title: dom.title.value.trim(),
      deck: dom.deck.value.trim(),
      topic: dom.topic.value.trim(),
      authorLastName: dom.authorLastName.value.trim(),
      authorFirstName: dom.authorFirstName.value.trim(),
      authorPhone: dom.authorPhone.value.trim(),
      coAuthors: getCoAuthors(),
      ticker: dom.ticker.value.trim(),
      crgRating: dom.crgRating.value.trim(),
      targetPrice: dom.targetPrice.value.trim(),
      valuationSummary: dom.valuationSummary.value.trim(),
      keyAssumptions: dom.keyAssumptions.value.trim(),
      scenarioNotes: dom.scenarioNotes.value.trim(),
      modelLink: dom.modelLink.value.trim(),
      keyTakeaways: dom.keyTakeaways.value.trim(),
      analysis: dom.analysis.value.trim(),
      content: dom.content.value.trim(),
      cordobaView: dom.cordobaView.value.trim(),
      imageFiles: Array.from(dom.imageUpload.files || []),
      modelFiles: Array.from(dom.modelFiles.files || []),
      priceChartImageBytes: state.priceChartImageBytes,
      equityStats: { ...state.equityStats },
      generatedAt: new Date()
    };
  }

  function buildDocumentFileName(data) {
    const titleSlug = slugify(data.title || "research-note");
    const typeSlug = slugify(data.noteType || "note");
    const dateSlug = formatDateShort(parseInputDate(data.publicationDate) || data.generatedAt || new Date());
    return `${dateSlug}_${titleSlug}_${typeSlug}.docx`;
  }

  function slugify(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "") || "research_note";
  }

  function formatDateShort(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}${month}${day}`;
  }

  function formatClock(date) {
    return new Intl.DateTimeFormat("en-GB", {
      hour: "numeric",
      minute: "2-digit"
    }).format(date);
  }

  function formatDateTime(date) {
    return new Intl.DateTimeFormat("en-GB", {
      day: "numeric",
      month: "long",
      year: "numeric",
      hour: "numeric",
      minute: "2-digit"
    }).format(date);
  }

  function formatDocDate(date) {
    return new Intl.DateTimeFormat("en-GB", {
      day: "numeric",
      month: "long",
      year: "numeric"
    }).format(date);
  }

  function parseInputDate(value) {
    if (!value) return null;
    const [year, month, day] = String(value).split("-").map(Number);
    if (!year || !month || !day) return null;
    return new Date(year, month - 1, day);
  }

  function formatInputDateLabel(value) {
    const date = parseInputDate(value);
    return date ? formatDocDate(date) : "Date pending";
  }

  function escapeAttribute(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function lineItems(text) {
    return String(text || "")
      .split("\n")
      .map((line) => line.replace(/^[-*•]\s*/, "").trim())
      .filter(Boolean);
  }

  function paragraphBlocks(text) {
    return String(text || "")
      .split(/\n{2,}/)
      .map((block) => block.trim())
      .filter(Boolean);
  }

  async function createDocument(data) {
    const docxLib = window.docx;
    const colors = {
      red: "845F0F",
      redDark: "6F500D",
      redDeep: "5D430B",
      black: "111111",
      ink: "1B1F24",
      muted: "5D636D",
      line: "D8DCE2",
      soft: "F5F6F8",
      white: "FFFFFF"
    };

    const publicationDate = parseInputDate(data.publicationDate) || data.generatedAt || new Date();
    const bannerBytes = await buildNomuraBannerImageBytes({
      noteType: data.noteType,
      deskLine: data.deskLine,
      publicationDate
    });

    const analystNames = [
      buildDocAuthorLine(data.authorLastName, data.authorFirstName, data.authorPhone),
      ...data.coAuthors
        .filter((coAuthor) => coAuthor.lastName || coAuthor.firstName)
        .map((coAuthor) => buildDocAuthorLine(coAuthor.lastName, coAuthor.firstName, coAuthor.phone))
    ].filter(Boolean);

    const documentChildren = [
      new docxLib.Paragraph({
        children: [
          new docxLib.ImageRun({
            data: bannerBytes,
            transformation: { width: 660, height: 112 }
          })
        ],
        spacing: { after: 110 }
      }),
      buildNomuraTitlePanel(docxLib, colors, data, analystNames, publicationDate),
      new docxLib.Paragraph({ spacing: { after: 70 } }),
      buildNomuraSubhead(docxLib, colors, "Key Takeaways"),
      ...buildNomuraBulletParagraphs(docxLib, lineItems(data.keyTakeaways))
    ];

    if (data.noteType === "Equity Research") {
      const marketRows = [];
      const targetPrice = parseNumber(data.targetPrice);
      const upside = computeUpsideToTarget(data.equityStats.currentPrice, targetPrice);

      marketRows.push(["Ticker / company", data.ticker || "N/A"]);
      marketRows.push(["CRG rating", data.crgRating || "N/A"]);
      marketRows.push(["Target price", data.targetPrice || "N/A"]);
      if (data.equityStats.currentPrice != null) marketRows.push(["Current price", data.equityStats.currentPrice.toFixed(2)]);
      if (data.equityStats.rangeReturn != null) marketRows.push(["Range return", formatPercent(data.equityStats.rangeReturn)]);
      if (data.equityStats.realisedVolAnn != null) marketRows.push(["Volatility (ann.)", formatPercent(data.equityStats.realisedVolAnn)]);
      if (upside != null) marketRows.push(["Upside / downside", formatPercent(upside)]);

      if (marketRows.length) {
        documentChildren.push(
          buildNomuraSubhead(docxLib, colors, "Market Chart"),
          buildNomuraMetricTable(docxLib, colors, marketRows)
        );
      }

      if (data.valuationSummary) {
        documentChildren.push(
          buildNomuraSubhead(docxLib, colors, "Valuation Summary"),
          ...buildNomuraBodyParagraphs(docxLib, data.valuationSummary)
        );
      }

      if (data.keyAssumptions) {
        documentChildren.push(
          buildNomuraSubhead(docxLib, colors, "Key Assumptions"),
          ...buildNomuraBulletParagraphs(docxLib, lineItems(data.keyAssumptions))
        );
      }

      if (data.scenarioNotes) {
        documentChildren.push(
          buildNomuraSubhead(docxLib, colors, "Scenario And Sensitivity Notes"),
          ...buildNomuraBodyParagraphs(docxLib, data.scenarioNotes)
        );
      }
    }

    documentChildren.push(
      buildNomuraSubhead(docxLib, colors, "Analysis And Commentary"),
      ...buildNomuraBodyParagraphs(docxLib, data.analysis)
    );

    if (data.content) {
      documentChildren.push(
        buildNomuraSubhead(docxLib, colors, "Additional Detail"),
        ...buildNomuraBodyParagraphs(docxLib, data.content)
      );
    }

    documentChildren.push(
      buildNomuraSubhead(docxLib, colors, "The Cordoba View"),
      ...buildNomuraBodyParagraphs(docxLib, data.cordobaView)
    );

    const exhibitParagraphs = await buildNomuraExhibitParagraphs(docxLib, colors, data);
    if (exhibitParagraphs.length) {
      documentChildren.push(buildNomuraSubhead(docxLib, colors, "Figures / Screenshots"), ...exhibitParagraphs);
    }

    const supportParagraphs = buildSupportingMaterialParagraphs(docxLib, colors, data);
    if (supportParagraphs.length) {
      documentChildren.push(buildNomuraSubhead(docxLib, colors, "Model Files"), ...supportParagraphs);
    }

    return new docxLib.Document({
      styles: {
        default: {
          document: {
            run: {
              font: "Arial",
              size: 18,
              color: colors.ink
            },
            paragraph: {
              spacing: {
                line: 220,
                after: 40
              }
            }
          }
        }
      },
      sections: [
        {
          properties: {
            page: {
              margin: { top: 520, right: 660, bottom: 660, left: 660 },
              pageSize: { width: 11906, height: 16838 }
            }
          },
          footers: {
            default: new docxLib.Footer({
              children: [
                new docxLib.Paragraph({
                  border: {
                    top: {
                      color: colors.line,
                      style: docxLib.BorderStyle.SINGLE,
                      size: 4
                    }
                  },
                  spacing: { after: 25 }
                }),
                new docxLib.Paragraph({
                  tabStops: [
                    { type: docxLib.TabStopType.CENTER, position: 4700 },
                    { type: docxLib.TabStopType.RIGHT, position: 9200 }
                  ],
                  children: [
                    new docxLib.TextRun({
                      text: `Published: ${formatDocDate(publicationDate)}`,
                      size: 12,
                      color: colors.muted,
                      font: "Arial"
                    }),
                    new docxLib.TextRun({ text: "\t" }),
                    new docxLib.TextRun({
                      text: `Audit: ${data.distribution} | ${nomuraDisplayType(data.noteType)} | Generated ${formatProductionTimestamp(data.generatedAt)}`,
                      size: 12,
                      color: colors.muted,
                      font: "Arial"
                    }),
                    new docxLib.TextRun({ text: "\t" }),
                    new docxLib.TextRun({
                      children: ["Page ", docxLib.PageNumber.CURRENT, " of ", docxLib.PageNumber.TOTAL_PAGES],
                      size: 12,
                      color: colors.muted,
                      font: "Arial"
                    })
                  ]
                })
              ]
            })
          },
          children: documentChildren
        }
      ]
    });
  }

  function nomuraDisplayType(noteType) {
    const map = {
      "General Note": "Strategy Trade",
      "Equity Research": "Equity Research",
      "Macro Research": "Macro View",
      "Fixed Income Research": "Fixed Income View",
      "Commodity Insights": "Commodities View"
    };

    return map[noteType] || "Research Note";
  }

  function buildDocAuthorLine(lastName, firstName, phone) {
    const name = [firstName, lastName].filter(Boolean).join(" ").trim();
    const phoneDisplay = formatPhoneDisplay(phone);
    if (name && phoneDisplay) return `${name}, ${phoneDisplay}`;
    return name || phoneDisplay;
  }

  function formatPhoneDisplay(phone) {
    const raw = String(phone || "").trim();
    if (!raw) return "";

    if (raw.startsWith("+")) {
      return `+${digitsOnly(raw)}`;
    }

    if (raw.includes("-")) {
      const [countryCode, nationalNumber] = raw.split("-");
      const cc = digitsOnly(countryCode);
      const nn = digitsOnly(nationalNumber);
      if (cc || nn) return `+${cc}${nn}`;
    }

    const digits = digitsOnly(raw);
    return digits ? `+${digits}` : raw;
  }

  async function buildNomuraBannerImageBytes(meta) {
    const canvas = document.createElement("canvas");
    const width = 1600;
    const height = 260;
    const ctx = canvas.getContext("2d");
    canvas.width = width;
    canvas.height = height;

    const publicationDate = meta.publicationDate || new Date();
    const deskLine = meta.deskLine || "Research - Global";

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    ctx.fillStyle = "#845F0F";
    ctx.fillRect(0, 0, width, 126);

    drawBannerShape(ctx, "#ffffff", 1145, 126, 104, 0, 180, 126);
    drawBannerShape(ctx, "rgba(255,255,255,0.16)", 1238, 126, 168, 0, 255, 126);
    drawBannerShape(ctx, "rgba(84,60,10,0.22)", 1332, 126, 188, 0, 265, 126);
    drawBannerShape(ctx, "rgba(84,60,10,0.14)", 1434, 126, 188, 0, 265, 126);

    ctx.fillStyle = "#ffffff";
    ctx.font = "700 42px Arial";
    ctx.textBaseline = "alphabetic";
    ctx.fillText("CORDOBA RESEARCH GROUP", 54, 74);

    ctx.fillStyle = "#111111";
    ctx.font = "700 38px Arial";
    ctx.fillText(nomuraDisplayType(meta.noteType), 54, 178);

    ctx.fillStyle = "#845F0F";
    ctx.textAlign = "right";
    ctx.font = "600 24px Arial";
    ctx.fillText("Global Markets Research", width - 46, 156);
    ctx.font = "700 30px Arial";
    ctx.fillText(formatDocDate(publicationDate), width - 46, 193);

    const stripGradient = ctx.createLinearGradient(0, 0, width, 0);
    stripGradient.addColorStop(0, "#6F500D");
    stripGradient.addColorStop(0.58, "#9C7420");
    stripGradient.addColorStop(1, "#CBA65B");
    ctx.fillStyle = stripGradient;
    ctx.fillRect(0, 214, width, 22);

    ctx.textAlign = "left";
    ctx.fillStyle = "#ffffff";
    ctx.font = "600 14px Arial";
    ctx.fillText(deskLine, 56, 229);

    return canvasToPngBytes(canvas);
  }

  function drawBannerShape(ctx, color, startX, baseY, width, topInset, skew, height) {
    ctx.save();
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(startX, baseY);
    ctx.lineTo(startX + width, baseY);
    ctx.lineTo(startX + width - skew, topInset);
    ctx.lineTo(startX - skew, topInset);
    ctx.closePath();
    ctx.fill();
    ctx.restore();
  }

  function buildNomuraTitlePanel(docxLib, colors, data, analystNames, publicationDate) {
    const analystDesk = data.deskLine || getDeskLine();
    const analystNameLines = analystNames.length ? analystNames : ["Research Analyst"];

    return new docxLib.Table({
      width: { size: 100, type: docxLib.WidthType.PERCENTAGE },
      borders: {
        top: { style: docxLib.BorderStyle.NONE },
        bottom: { style: docxLib.BorderStyle.NONE },
        left: { style: docxLib.BorderStyle.NONE },
        right: { style: docxLib.BorderStyle.NONE },
        insideHorizontal: { style: docxLib.BorderStyle.NONE },
        insideVertical: { style: docxLib.BorderStyle.NONE }
      },
      rows: [
        new docxLib.TableRow({
          children: [
            new docxLib.TableCell({
              width: { size: 72, type: docxLib.WidthType.PERCENTAGE },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              margins: { top: 0, bottom: 0, left: 0, right: 160 },
              children: [
                new docxLib.Paragraph({
                  children: [
                    new docxLib.TextRun({
                      text: data.title || "Untitled Research Note",
                      bold: true,
                      size: 30,
                      color: colors.black,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 95 }
                }),
                new docxLib.Paragraph({
                  children: [
                    new docxLib.TextRun({
                      text: data.deck || "No deck supplied",
                      bold: true,
                      size: 20,
                      color: colors.ink,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 65 }
                }),
                new docxLib.Paragraph({
                  children: [
                    new docxLib.TextRun({
                      text: `${data.topic || "Topic pending"} | ${formatDocDate(publicationDate)}`,
                      size: 17,
                      color: colors.muted,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 20 }
                })
              ]
            }),
            new docxLib.TableCell({
              width: { size: 28, type: docxLib.WidthType.PERCENTAGE },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              margins: { top: 0, bottom: 0, left: 0, right: 0 },
              children: [
                new docxLib.Paragraph({
                  children: [
                    new docxLib.TextRun({
                      text: "Research Analysts",
                      bold: true,
                      size: 15,
                      color: colors.black,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 45 }
                }),
                new docxLib.Paragraph({
                  border: {
                    bottom: {
                      color: colors.red,
                      style: docxLib.BorderStyle.SINGLE,
                      size: 6
                    }
                  },
                  children: [
                    new docxLib.TextRun({
                      text: analystDesk,
                      color: colors.red,
                      underline: {},
                      size: 15,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 60 }
                }),
                buildAnalystNameTable(docxLib, colors, analystNameLines)
              ]
            })
          ]
        })
      ]
    });
  }

  function buildAnalystNameTable(docxLib, colors, lines) {
    return new docxLib.Table({
      width: { size: 100, type: docxLib.WidthType.PERCENTAGE },
      borders: {
        top: { style: docxLib.BorderStyle.NONE },
        bottom: { style: docxLib.BorderStyle.NONE },
        left: { style: docxLib.BorderStyle.NONE },
        right: { style: docxLib.BorderStyle.NONE },
        insideHorizontal: { style: docxLib.BorderStyle.NONE },
        insideVertical: { style: docxLib.BorderStyle.NONE }
      },
      rows: [
        new docxLib.TableRow({
          children: [
            new docxLib.TableCell({
              shading: { fill: colors.black },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              margins: { top: 80, bottom: 80, left: 110, right: 110 },
              children: [
                ...lines.map((line) =>
                  new docxLib.Paragraph({
                    children: [
                      new docxLib.TextRun({
                        text: line,
                        color: colors.white,
                        size: 14,
                        font: "Arial"
                      })
                    ],
                    spacing: { after: 10 }
                  })
                )
              ]
            })
          ]
        })
      ]
    });
  }

  function buildNomuraSubhead(docxLib, colors, label) {
    return new docxLib.Paragraph({
      children: [
        new docxLib.TextRun({
          text: label,
          bold: true,
          color: colors.red,
          size: 18,
          font: "Arial"
        })
      ],
      spacing: { before: 135, after: 34 }
    });
  }

  function buildNomuraBulletParagraphs(docxLib, items) {
    const cleanItems = items.length ? items : ["No key takeaways supplied."];

    return cleanItems.map((item) =>
      new docxLib.Paragraph({
        indent: { left: 180, hanging: 120 },
        children: [
          new docxLib.TextRun({
            text: "• ",
            bold: true,
            font: "Arial",
            size: 18
          }),
          new docxLib.TextRun({
            text: item,
            font: "Arial",
            size: 18,
            color: "1B1F24"
          })
        ],
        spacing: { after: 42 }
      })
    );
  }

  function buildNomuraBodyParagraphs(docxLib, text) {
    const blocks = paragraphBlocks(text);
    if (!blocks.length) {
      return [new docxLib.Paragraph({ text: "No content supplied.", spacing: { after: 44 } })];
    }

    return blocks.map((block) =>
      new docxLib.Paragraph({
        children: [
          new docxLib.TextRun({
            text: block.replace(/\s*\n\s*/g, " "),
            font: "Arial",
            size: 18,
            color: "1B1F24"
          })
        ],
        spacing: { after: 58 }
      })
    );
  }

  function buildNomuraMetricTable(docxLib, colors, rows) {
    return new docxLib.Table({
      width: { size: 100, type: docxLib.WidthType.PERCENTAGE },
      borders: {
        top: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        bottom: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        left: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        right: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        insideHorizontal: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        insideVertical: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 }
      },
      rows: rows.map((row, index) =>
        new docxLib.TableRow({
          children: row.map((cell, cellIndex) =>
            new docxLib.TableCell({
              shading: { fill: cellIndex === 0 ? colors.soft : colors.white },
              margins: { top: 65, bottom: 65, left: 90, right: 90 },
              children: [
                new docxLib.Paragraph({
                  children: [
                    new docxLib.TextRun({
                      text: cell,
                      bold: cellIndex === 0,
                      color: cellIndex === 0 ? colors.black : colors.ink,
                      size: 17,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 10 }
                })
              ]
            })
          )
        })
      )
    });
  }

  async function buildNomuraExhibitParagraphs(docxLib, colors, data) {
    const output = [];

    if (data.priceChartImageBytes) {
      output.push(
        new docxLib.Paragraph({
          children: [
            new docxLib.ImageRun({
              data: data.priceChartImageBytes,
              transformation: { width: 560, height: 235 }
            })
          ],
          alignment: docxLib.AlignmentType.CENTER,
          spacing: { before: 60, after: 45 }
        }),
        new docxLib.Paragraph({
          alignment: docxLib.AlignmentType.CENTER,
          children: [
            new docxLib.TextRun({
              text: `Figure 1. ${(data.ticker || "Security").toUpperCase()} price performance`,
              italics: true,
              color: colors.muted,
              size: 15,
              font: "Arial"
            })
          ],
          spacing: { after: 90 }
        })
      );
    }

    const startIndex = output.length ? 2 : 1;
    const imageParagraphs = await buildImageParagraphs(docxLib, data.imageFiles, colors, startIndex);
    output.push(...imageParagraphs);

    return output;
  }

  function buildSupportingMaterialParagraphs(docxLib, colors, data) {
    const paragraphs = [];

    if (data.modelLink) {
      paragraphs.push(
        new docxLib.Paragraph({
          children: [
            new docxLib.TextRun({
              text: "Model link: ",
              bold: true,
              font: "Arial",
              size: 18
            }),
            new docxLib.ExternalHyperlink({
              children: [new docxLib.TextRun({ text: data.modelLink, style: "Hyperlink" })],
              link: data.modelLink
            })
          ],
          spacing: { after: 45 }
        })
      );
    }

    data.modelFiles.forEach((file) => {
      paragraphs.push(
        new docxLib.Paragraph({
          indent: { left: 180, hanging: 120 },
          children: [
            new docxLib.TextRun({
              text: "• ",
              bold: true,
              font: "Arial",
              size: 18
            }),
            new docxLib.TextRun({
              text: file.name,
              font: "Arial",
              size: 18,
              color: colors.ink
            })
          ],
          spacing: { after: 40 }
        })
      );
    });

    return paragraphs;
  }

  async function buildImageParagraphs(docxLib, files, colors, startIndex = 1) {
    const output = [];

    for (let index = 0; index < files.length; index += 1) {
      const file = files[index];
      const buffer = await file.arrayBuffer();
      const size = await getImageFit(file, 560, 320);
      const caption = file.name.replace(/\.[^.]+$/, "");

      output.push(
        new docxLib.Paragraph({
          children: [
            new docxLib.ImageRun({
              data: buffer,
              transformation: size
            })
          ],
          alignment: docxLib.AlignmentType.CENTER,
          spacing: { before: 45, after: 45 }
        }),
        new docxLib.Paragraph({
          alignment: docxLib.AlignmentType.CENTER,
          children: [
            new docxLib.TextRun({
              text: `Figure ${startIndex + index}. ${caption}`,
              italics: true,
              color: colors.muted,
              size: 15,
              font: "Arial"
            })
          ],
          spacing: { after: 85 }
        })
      );
    }

    return output;
  }

  function formatProductionTimestamp(date) {
    const d = date instanceof Date ? date : new Date(date);
    const year = d.getUTCFullYear();
    const month = String(d.getUTCMonth() + 1).padStart(2, "0");
    const day = String(d.getUTCDate()).padStart(2, "0");
    const hours = String(d.getUTCHours()).padStart(2, "0");
    const minutes = String(d.getUTCMinutes()).padStart(2, "0");
    return `${year}-${month}-${day} ${hours}:${minutes} UTC`;
  }

  function getImageFit(file, maxWidth, maxHeight) {
    return new Promise((resolve) => {
      const image = new Image();
      const objectUrl = URL.createObjectURL(file);

      image.onload = () => {
        const width = image.naturalWidth || maxWidth;
        const height = image.naturalHeight || maxHeight;
        const scale = Math.min(maxWidth / width, maxHeight / height, 1);

        URL.revokeObjectURL(objectUrl);
        resolve({
          width: Math.round(width * scale),
          height: Math.round(height * scale)
        });
      };

      image.onerror = () => {
        URL.revokeObjectURL(objectUrl);
        resolve({ width: maxWidth, height: Math.round(maxWidth * 0.62) });
      };

      image.src = objectUrl;
    });
  }
});
