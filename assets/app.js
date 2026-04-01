console.log("CRG Research Production Console loaded");

document.addEventListener("DOMContentLoaded", () => {
  const STORAGE_KEY = "crg-rdt-draft-v2";
  const NOTE_SEQUENCE_STORAGE_KEY = "crg-rdt-note-seq-v1";
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

  const WORLD_COUNTRY_CODES = [
    "AF","AX","AL","DZ","AS","AD","AO","AI","AQ","AG","AR","AM","AW","AU","AT","AZ","BS","BH","BD","BB","BY","BE","BZ","BJ","BM","BT","BO","BQ","BA","BW","BV","BR","IO","BN","BG","BF","BI","CV","KH","CM","CA","KY","CF","TD","CL","CN","CX","CC","CO","KM","CG","CD","CK","CR","CI","HR","CU","CW","CY","CZ","DK","DJ","DM","DO","EC","EG","SV","GQ","ER","EE","SZ","ET","FK","FO","FJ","FI","FR","GF","PF","TF","GA","GM","GE","DE","GH","GI","GR","GL","GD","GP","GU","GT","GG","GN","GW","GY","HT","HM","VA","HN","HK","HU","IS","IN","ID","IR","IQ","IE","IM","IT","JM","JP","JE","JO","KZ","KE","KI","KP","KR","KW","KG","LA","LV","LB","LS","LR","LY","LI","LT","LU","MO","MG","MW","MY","MV","ML","MT","MH","MQ","MR","MU","YT","MX","FM","MD","MC","MN","ME","MS","MA","MZ","MM","NA","NR","NP","NL","NC","NZ","NI","NE","NG","NU","NF","MK","MP","NO","OM","PK","PW","PS","PA","PG","PY","PE","PH","PN","PL","PT","PR","QA","RE","RO","RU","RW","BL","SH","KN","LC","MF","PM","VC","WS","SM","ST","SA","SN","RS","SC","SL","SG","SX","SK","SI","SB","SO","ZA","GS","SS","ES","LK","SD","SR","SJ","SE","CH","SY","TW","TJ","TZ","TH","TL","TG","TK","TO","TT","TN","TR","TM","TC","TV","UG","UA","AE","GB","US","UM","UY","UZ","VU","VE","VN","VG","VI","WF","EH","YE","ZM","ZW"
  ];

  const WORLD_COUNTRIES = buildWorldCountryOptions();
  const EQUITY_INDUSTRIES = [
    "Aerospace & Defense",
    "Air Freight & Logistics",
    "Airlines",
    "Auto Components",
    "Automobiles",
    "Banks",
    "Beverages",
    "Biotechnology",
    "Broadline Retail",
    "Building Products",
    "Capital Markets",
    "Chemicals",
    "Commercial Services & Supplies",
    "Communications Equipment",
    "Construction & Engineering",
    "Construction Materials",
    "Consumer Finance",
    "Consumer Staples Distribution & Retail",
    "Containers & Packaging",
    "Diversified Consumer Services",
    "Diversified Financial Services",
    "Diversified Telecommunication Services",
    "Distributors",
    "Electric Utilities",
    "Electrical Equipment",
    "Electronic Equipment, Instruments & Components",
    "Energy Equipment & Services",
    "Entertainment",
    "Equity Real Estate Investment Trusts (REITs)",
    "Food Products",
    "Gas Utilities",
    "Health Care Equipment & Supplies",
    "Health Care Providers & Services",
    "Health Care Technology",
    "Hotels, Restaurants & Leisure",
    "Household Durables",
    "Household Products",
    "Independent Power and Renewable Electricity Producers",
    "Industrial Conglomerates",
    "Insurance",
    "Interactive Media & Services",
    "IT Services",
    "Leisure Products",
    "Life Sciences Tools & Services",
    "Machinery",
    "Marine Transportation",
    "Media",
    "Metals & Mining",
    "Multi-Utilities",
    "Oil, Gas & Consumable Fuels",
    "Paper & Forest Products",
    "Personal Care Products",
    "Pharmaceuticals",
    "Professional Services",
    "Real Estate Management & Development",
    "Renewable Energy",
    "Road & Rail",
    "Semiconductors & Semiconductor Equipment",
    "Software",
    "Technology Hardware, Storage & Peripherals",
    "Textiles, Apparel & Luxury Goods",
    "Thrifts & Mortgage Finance",
    "Tobacco",
    "Trading Companies & Distributors",
    "Transportation Infrastructure",
    "Water Utilities",
    "Wireless Telecommunication Services"
  ];

  const ISSUER_COUNTRY_MAP = {
    US: "923213",
    GB: "814275",
    DE: "742318",
    FR: "742319",
    IT: "735482",
    ES: "731456",
    JP: "552410",
    CN: "683120",
    IN: "744205",
    SA: "615240",
    AE: "615241",
    BR: "691330",
    ZA: "604120",
    TR: "631450",
    MX: "612340",
    ID: "643210"
  };

  const FIELD_LABELS = {
    noteType: "Note type",
    distribution: "Distribution",
    publicationDate: "Publication date",
    deck: "Deck / standfirst",
    title: "Research title",
    topic: "Topic / coverage angle",
    coverageCountry: "Coverage country",
    issuerId: "Issuer number",
    authorLastName: "Primary author last name",
    authorFirstName: "Primary author first name",
    ticker: "Ticker / company",
    equitySectorLine: "Sector / industry",
    crgRating: "CRG rating",
    targetPrice: "Target price",
    marketCapUsd: "Market cap",
    businessDescription: "Business description",
    fiveYearRationale: "5 Year rationale & price target",
    esgSummary: "ESG summary",
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
    coverageCountry: "Brief",
    issuerId: "Brief",
    authorLastName: "Authors",
    authorFirstName: "Authors",
    ticker: "Equity",
    equitySectorLine: "Equity",
    crgRating: "Equity",
    targetPrice: "Equity",
    marketCapUsd: "Equity",
    businessDescription: "Equity",
    fiveYearRationale: "Research",
    esgSummary: "Research",
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

  const EQUITY_REQUIRED_IDS = ["ticker", "equitySectorLine", "crgRating", "targetPrice", "marketCapUsd", "businessDescription"];

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
    appBody: document.getElementById("appBody"),
    workflowRail: document.getElementById("workflowRail"),
    toggleRailBtn: document.getElementById("toggleRailBtn"),
    toggleRailIcon: document.getElementById("toggleRailIcon"),
    toggleRailText: document.getElementById("toggleRailText"),
    headerDateTime: document.getElementById("headerDateTime"),
    noteType: document.getElementById("noteType"),
    distribution: document.getElementById("distribution"),
    publicationDate: document.getElementById("publicationDate"),
    deskLine: document.getElementById("deskLine"),
    title: document.getElementById("title"),
    deck: document.getElementById("deck"),
    topic: document.getElementById("topic"),
    macroFiPanel: document.getElementById("macroFiPanel"),
    coverageCountryPanel: document.getElementById("coverageCountryPanel"),
    coverageCountryToggle: document.getElementById("coverageCountryToggle"),
    coverageCountrySearch: document.getElementById("coverageCountrySearch"),
    coverageCountryDisplay: document.getElementById("coverageCountryDisplay"),
    coverageCountryOptions: document.getElementById("coverageCountryOptions"),
    coverageCountry: document.getElementById("coverageCountry"),
    industryPanel: document.getElementById("industryPanel"),
    industryToggle: document.getElementById("industryToggle"),
    industrySearch: document.getElementById("industrySearch"),
    industryOptions: document.getElementById("industryOptions"),
    issuerId: document.getElementById("issuerId"),
    macroFiHeading: document.getElementById("macroFiHeading"),
    agencyRating: document.getElementById("agencyRating"),
    shortTermRating: document.getElementById("shortTermRating"),
    longTermRating: document.getElementById("longTermRating"),
    ratingsProfileRows: document.getElementById("ratingsProfileRows"),
    addRatingProfileRowBtn: document.getElementById("addRatingProfileRowBtn"),
    ratingProfileRowTemplate: document.getElementById("ratingProfileRowTemplate"),
    authorLastName: document.getElementById("authorLastName"),
    authorFirstName: document.getElementById("authorFirstName"),
    authorPhoneCountry: document.getElementById("authorPhoneCountry"),
    authorPhoneNational: document.getElementById("authorPhoneNational"),
    authorPhone: document.getElementById("authorPhone"),
    coAuthorsList: document.getElementById("coAuthorsList"),
    addCoAuthor: document.getElementById("addCoAuthor"),
    equitySection: document.getElementById("section-equity"),
    ticker: document.getElementById("ticker"),
    equityCompanyName: document.getElementById("equityCompanyName"),
    equitySecurityDisplay: document.getElementById("equitySecurityDisplay"),
    equitySectorLine: document.getElementById("equitySectorLine"),
    crgRating: document.getElementById("crgRating"),
    targetPrice: document.getElementById("targetPrice"),
    priceCurrency: document.getElementById("priceCurrency"),
    benchmarkName: document.getElementById("benchmarkName"),
    benchmarkTicker: document.getElementById("benchmarkTicker"),
    marketCapUsd: document.getElementById("marketCapUsd"),
    adtUsd: document.getElementById("adtUsd"),
    chartRange: document.getElementById("chartRange"),
    fetchPriceChart: document.getElementById("fetchPriceChart"),
    chartStatus: document.getElementById("chartStatus"),
    priceChartCanvas: document.getElementById("priceChart"),
    currentPrice: document.getElementById("currentPrice"),
    priceDate: document.getElementById("priceDate"),
    rangeReturn: document.getElementById("rangeReturn"),
    upsideToTarget: document.getElementById("upsideToTarget"),
    modelFiles: document.getElementById("modelFiles"),
    modelLink: document.getElementById("modelLink"),
    businessDescription: document.getElementById("businessDescription"),
    valuationSummary: document.getElementById("valuationSummary"),
    financialTableTitle: document.getElementById("financialTableTitle"),
    financialSourceNote: document.getElementById("financialSourceNote"),
    financialTableEditor: document.getElementById("financialTableEditor"),
    financialTableHead: document.getElementById("financialTableHead"),
    financialTableBody: document.getElementById("financialTableBody"),
    financialTableInput: document.getElementById("financialTableInput"),
    financialDownloadTemplateBtn: document.getElementById("financialDownloadTemplateBtn"),
    financialUploadTriggerBtn: document.getElementById("financialUploadTriggerBtn"),
    financialTemplateUpload: document.getElementById("financialTemplateUpload"),
    financialAddColumnBtn: document.getElementById("financialAddColumnBtn"),
    financialAddRowBtn: document.getElementById("financialAddRowBtn"),
    financialResetBtn: document.getElementById("financialResetBtn"),
    bodySections: document.getElementById("bodySections"),
    addCustomSectionBtn: document.getElementById("addCustomSectionBtn"),
    customBodySectionTemplate: document.getElementById("customBodySectionTemplate"),
    headingKeyTakeaways: document.getElementById("headingKeyTakeaways"),
    headingAnalysis: document.getElementById("headingAnalysis"),
    headingContent: document.getElementById("headingContent"),
    headingFiveYearRationale: document.getElementById("headingFiveYearRationale"),
    headingEsgSummary: document.getElementById("headingEsgSummary"),
    headingCordobaView: document.getElementById("headingCordobaView"),
    keyTakeaways: document.getElementById("keyTakeaways"),
    analysis: document.getElementById("analysis"),
    content: document.getElementById("content"),
    fiveYearRationale: document.getElementById("fiveYearRationale"),
    esgSummary: document.getElementById("esgSummary"),
    cordobaView: document.getElementById("cordobaView"),
    imageUpload: document.getElementById("imageUpload"),
    modelSummaryHead: document.getElementById("modelSummaryHead"),
    modelSummaryList: document.getElementById("modelSummaryList"),
    imageSummaryHead: document.getElementById("imageSummaryHead"),
    imageSummaryList: document.getElementById("imageSummaryList"),
    figurePlacementPanel: document.getElementById("figurePlacementPanel"),
    figurePlacementList: document.getElementById("figurePlacementList"),
    completionBar: document.getElementById("completionBar"),
    completionText: document.getElementById("completionText"),
    readinessPercent: document.getElementById("readinessPercent"),
    readinessCaption: document.getElementById("readinessCaption"),
    noteStateChip: document.getElementById("noteStateChip"),
    missingFields: document.getElementById("missingFields"),
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
    previewDocBtn: document.getElementById("previewDocBtn"),
    previewModal: document.getElementById("previewModal"),
    previewModalBackdrop: document.getElementById("previewModalBackdrop"),
    previewModalBody: document.getElementById("previewModalBody"),
    closePreviewBtn: document.getElementById("closePreviewBtn"),
    previewGenerateBtn: document.getElementById("previewGenerateBtn"),
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
      rangeReturn: null,
      priceDate: null,
      providerCurrency: "",
      benchmarkLabel: "",
      chartMode: "price",
      symbol: "",
      benchmarkSymbol: "",
      range: ""
    },
    financialTable: null,
    railCollapsed: false,
    saveTimer: null,
    lastSavedAt: null,
    customSectionCount: 0,
    brandLogoImagePromise: null,
    figurePlacements: {},
    figureFiles: [],
    figureDetails: {},
    previewObjectUrls: []
  };

  const draftFieldIds = [
    "noteType",
    "distribution",
    "publicationDate",
    "deskLine",
    "title",
    "deck",
    "topic",
    "coverageCountry",
    "issuerId",
    "macroFiHeading",
    "agencyRating",
    "shortTermRating",
    "longTermRating",
    "authorLastName",
    "authorFirstName",
    "authorPhoneCountry",
    "authorPhoneNational",
    "authorPhone",
    "ticker",
    "equityCompanyName",
    "equitySecurityDisplay",
    "equitySectorLine",
    "crgRating",
    "targetPrice",
    "priceCurrency",
    "benchmarkName",
    "benchmarkTicker",
    "marketCapUsd",
    "adtUsd",
    "modelLink",
    "businessDescription",
    "valuationSummary",
    "financialTableTitle",
    "financialSourceNote",
    "financialTableInput",
    "headingKeyTakeaways",
    "headingAnalysis",
    "headingContent",
    "headingFiveYearRationale",
    "headingEsgSummary",
    "headingCordobaView",
    "keyTakeaways",
    "analysis",
    "content",
    "fiveYearRationale",
    "esgSummary",
    "cordobaView",
    "chartRange"
  ];

  init();

  function init() {
    startHeaderClock();
    restoreRailState();
    populateCoverageCountryOptions();
    populateIndustryOptions();
    wireSectionNavigation();
    wirePrimaryPhone();
    initializeBodySectionEditor();
    wireFormEvents();
    restoreDraft();
    ensurePublicationDate();
    initializeFinancialTableEditor();
    ensureDeskLineDefault();
    syncIssuerId();
    updateCoverageCountryDisplayFromCode();
    syncPrimaryPhone();
    toggleNoteTypeSections();
    updateFileSummary(dom.modelFiles, dom.modelSummaryHead, dom.modelSummaryList, "No supporting files attached.");
    setManagedFigureFiles([]);
    syncFigurePlacementControls();
    updateAllUI();
    checkLibraries();
  }

  function startHeaderClock() {
    updateHeaderDateTime();
    window.setInterval(updateHeaderDateTime, 30000);
  }

  function updateHeaderDateTime() {
    if (!dom.headerDateTime) return;
    dom.headerDateTime.textContent = formatLocalTimestamp(new Date());
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
      toggleNoteTypeSections();
      resetChartState({ keepStatusText: false });
      syncFigurePlacementControls();
      updateAllUI();
      queueDraftSave();
    });

    dom.publicationDate.addEventListener("change", () => {
      syncIssuerId();
      updateAllUI();
      queueDraftSave();
    });

    dom.coverageCountryDisplay?.addEventListener("click", openCoverageCountryPanel);
    dom.coverageCountryToggle?.addEventListener("click", () => {
      if (dom.coverageCountryPanel?.hidden) openCoverageCountryPanel();
      else closeCoverageCountryPanel();
    });
    dom.coverageCountrySearch?.addEventListener("input", () => {
      renderCoverageCountryOptions(dom.coverageCountrySearch.value);
    });
    dom.coverageCountryOptions?.addEventListener("click", (event) => {
      const option = event.target.closest("[data-country-code]");
      if (!option) return;
      selectCoverageCountry(option.getAttribute("data-country-code"));
    });
    dom.equitySectorLine?.addEventListener("click", openIndustryPanel);
    dom.industryToggle?.addEventListener("click", () => {
      if (dom.industryPanel?.hidden) openIndustryPanel();
      else closeIndustryPanel();
    });
    dom.industrySearch?.addEventListener("input", () => {
      renderIndustryOptions(dom.industrySearch.value);
    });
    dom.industryOptions?.addEventListener("click", (event) => {
      const option = event.target.closest("[data-industry-value]");
      if (!option) return;
      selectIndustry(option.getAttribute("data-industry-value"));
    });
    document.addEventListener("click", (event) => {
      if (dom.coverageCountryPanel && !dom.coverageCountryPanel.hidden) {
        const picker = document.getElementById("coverageCountryPicker");
        if (picker && !picker.contains(event.target)) closeCoverageCountryPanel();
      }
      if (dom.industryPanel && !dom.industryPanel.hidden) {
        const picker = document.getElementById("industryPicker");
        if (picker && !picker.contains(event.target)) closeIndustryPanel();
      }
    });

    dom.macroFiHeading?.addEventListener("input", syncFigurePlacementControls);

    dom.deskLine.addEventListener("input", () => {
      dom.deskLine.dataset.autofill = "false";
      updateAllUI();
      queueDraftSave();
    });

    dom.equityCompanyName.addEventListener("input", () => {
      dom.equityCompanyName.dataset.autofill = "false";
    });

    dom.priceCurrency.addEventListener("input", () => {
      dom.priceCurrency.dataset.autofill = "false";
    });

    dom.addCustomSectionBtn?.addEventListener("click", () => {
      addCustomBodySection();
      syncFigurePlacementControls();
      updateAllUI();
      queueDraftSave();
    });

    dom.bodySections?.addEventListener("click", handleBodySectionActions);
    dom.bodySections?.addEventListener("input", handleBodySectionInput);

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
      handleFigureUploadChange();
      syncFigurePlacementControls();
      updateAllUI();
    });

    dom.addRatingProfileRowBtn?.addEventListener("click", () => {
      addRatingProfileRow();
      updateAllUI();
      queueDraftSave();
    });
    dom.ratingsProfileRows?.addEventListener("click", handleRatingProfileActions);

    dom.figurePlacementList?.addEventListener("change", handleFigurePlacementChange);
    dom.figurePlacementList?.addEventListener("input", handleFigurePlacementChange);
    dom.figurePlacementList?.addEventListener("click", handleFigurePlacementActions);

    dom.financialAddColumnBtn.addEventListener("click", addFinancialPeriodColumn);
    dom.financialAddRowBtn.addEventListener("click", addFinancialLineItem);
    dom.financialResetBtn.addEventListener("click", resetFinancialTableGrid);
    dom.financialDownloadTemplateBtn.addEventListener("click", downloadFinancialTemplate);
    dom.financialUploadTriggerBtn.addEventListener("click", () => {
      if (!ensureXlsxAvailable("import the completed Excel template")) return;
      dom.financialTemplateUpload.click();
    });
    dom.financialTemplateUpload.addEventListener("change", handleFinancialTemplateUpload);

    dom.financialTableEditor.addEventListener("input", handleFinancialGridInput);
    dom.financialTableEditor.addEventListener("focusout", handleFinancialGridFocusOut);
    dom.financialTableEditor.addEventListener("click", handleFinancialGridClick);
    dom.toggleRailBtn.addEventListener("click", toggleRail);

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
    dom.previewDocBtn?.addEventListener("click", openPreviewModal);
    dom.closePreviewBtn?.addEventListener("click", closePreviewModal);
    dom.previewModalBackdrop?.addEventListener("click", closePreviewModal);
    dom.previewGenerateBtn?.addEventListener("click", () => {
      closePreviewModal();
      form.requestSubmit(dom.generateDocBtn);
    });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && !dom.previewModal?.hidden) closePreviewModal();
    });
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

  function restoreRailState() {
    applyRailState(false);
  }

  function applyRailState(collapsed) {
    state.railCollapsed = Boolean(collapsed);
    dom.appBody.classList.toggle("is-rail-collapsed", state.railCollapsed);
    dom.toggleRailBtn.classList.toggle("is-collapsed", state.railCollapsed);
    dom.toggleRailIcon.textContent = state.railCollapsed ? ">" : "<";
    dom.toggleRailText.textContent = state.railCollapsed ? "Show Workflow Panel" : "Hide Workflow Panel";
    dom.toggleRailBtn.setAttribute("aria-expanded", String(!state.railCollapsed));
    dom.toggleRailBtn.setAttribute("aria-label", state.railCollapsed ? "Show workflow side panel" : "Hide workflow side panel");
  }

  function toggleRail() {
    applyRailState(!state.railCollapsed);
  }

  function todayIsoDate() {
    const now = new Date();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${now.getFullYear()}-${month}-${day}`;
  }

  function strategyLabelForNoteType(noteType) {
    const map = {
      "General Note": "Global Strategy",
      "Equity Research": "Global Equity Strategy",
      "Macro Research": "Global Macro Strategy",
      "Fixed Income Research": "Global Fixed Income Strategy",
      "Commodity Insights": "Global Commodity Strategy",
      "Commodity Research": "Global Commodity Strategy",
      "Short Note / Market Alert": "Global Market Alert",
      Commodity: "Global Commodity Strategy"
    };

    return noteType ? (map[noteType] || "Research - Global") : "";
  }

  function defaultDeskLine(noteType) {
    return strategyLabelForNoteType(noteType);
  }

  function buildWorldCountryOptions() {
    const displayNames = typeof Intl.DisplayNames === "function"
      ? new Intl.DisplayNames(["en"], { type: "region" })
      : null;

    return WORLD_COUNTRY_CODES
      .map((code) => ({
        code,
        name: code === "PS" ? "Palestine" : (displayNames?.of(code) || code)
      }))
      .sort((left, right) => left.name.localeCompare(right.name, "en", { sensitivity: "base" }));
  }

  function getCoverageCountryLabel(countryCode) {
    const normalized = String(countryCode || "").trim().toUpperCase();
    return WORLD_COUNTRIES.find((entry) => entry.code === normalized)?.name || normalized || "N/A";
  }

  function populateCoverageCountryOptions() {
    if (!dom.coverageCountryOptions) return;
    renderCoverageCountryOptions("");
  }

  function updateCoverageCountryDisplayFromCode() {
    if (!dom.coverageCountryDisplay) return;
    dom.coverageCountryDisplay.value = getCoverageCountryLabel(dom.coverageCountry.value).replace(/^N\/A$/, "");
  }

  function renderCoverageCountryOptions(filterText = "") {
    if (!dom.coverageCountryOptions) return;
    const normalizedFilter = normalizeComparableText(filterText);
    const matches = WORLD_COUNTRIES.filter((country) => {
      if (!normalizedFilter) return true;
      return normalizeComparableText(country.name).includes(normalizedFilter) || normalizeComparableText(country.code).includes(normalizedFilter);
    });

    dom.coverageCountryOptions.innerHTML = "";
    if (!matches.length) {
      const empty = document.createElement("div");
      empty.className = "country-picker-empty";
      empty.textContent = "No matching countries";
      dom.coverageCountryOptions.appendChild(empty);
      return;
    }

    matches.forEach((country) => {
      const option = document.createElement("button");
      option.type = "button";
      option.className = "country-picker-option";
      option.textContent = country.name;
      option.setAttribute("data-country-code", country.code);
      if (country.code === String(dom.coverageCountry.value || "").trim().toUpperCase()) {
        option.classList.add("is-selected");
      }
      dom.coverageCountryOptions.appendChild(option);
    });
  }

  function openCoverageCountryPanel() {
    if (!dom.coverageCountryPanel) return;
    dom.coverageCountryPanel.hidden = false;
    if (dom.coverageCountrySearch) {
      dom.coverageCountrySearch.value = "";
      renderCoverageCountryOptions("");
      window.setTimeout(() => dom.coverageCountrySearch.focus(), 0);
    }
  }

  function closeCoverageCountryPanel() {
    if (!dom.coverageCountryPanel) return;
    dom.coverageCountryPanel.hidden = true;
  }

  function selectCoverageCountry(countryCode) {
    const normalized = String(countryCode || "").trim().toUpperCase();
    dom.coverageCountry.value = normalized;
    updateCoverageCountryDisplayFromCode();
    syncIssuerId();
    renderCoverageCountryOptions(dom.coverageCountrySearch?.value || "");
    closeCoverageCountryPanel();
    updateAllUI();
    queueDraftSave();
  }

  function populateIndustryOptions() {
    if (!dom.industryOptions) return;
    renderIndustryOptions("");
  }

  function renderIndustryOptions(filterText = "") {
    if (!dom.industryOptions) return;
    const normalizedFilter = normalizeComparableText(filterText);
    const matches = EQUITY_INDUSTRIES.filter((industry) => {
      if (!normalizedFilter) return true;
      return normalizeComparableText(industry).includes(normalizedFilter);
    });

    dom.industryOptions.innerHTML = "";
    if (!matches.length) {
      const empty = document.createElement("div");
      empty.className = "country-picker-empty";
      empty.textContent = "No matching industries";
      dom.industryOptions.appendChild(empty);
      return;
    }

    matches.forEach((industry) => {
      const option = document.createElement("button");
      option.type = "button";
      option.className = "country-picker-option";
      option.textContent = industry;
      option.setAttribute("data-industry-value", industry);
      if (industry === String(dom.equitySectorLine?.value || "").trim()) option.classList.add("is-selected");
      dom.industryOptions.appendChild(option);
    });
  }

  function openIndustryPanel() {
    if (!dom.industryPanel) return;
    dom.industryPanel.hidden = false;
    if (dom.industrySearch) {
      dom.industrySearch.value = "";
      renderIndustryOptions("");
      window.setTimeout(() => dom.industrySearch.focus(), 0);
    }
  }

  function closeIndustryPanel() {
    if (!dom.industryPanel) return;
    dom.industryPanel.hidden = true;
  }

  function selectIndustry(industry) {
    const normalized = String(industry || "").trim();
    if (!dom.equitySectorLine) return;
    dom.equitySectorLine.value = normalized;
    dom.equitySectorLine.dataset.autofill = "false";
    renderIndustryOptions(dom.industrySearch?.value || "");
    closeIndustryPanel();
    updateAllUI();
    queueDraftSave();
  }

  function isMacroFiNoteType(noteType) {
    return noteType === "Macro Research" || noteType === "Fixed Income Research";
  }

  function isEquitySelected() {
    return dom.noteType.value === "Equity Research";
  }

  function isMacroFiSelected() {
    return isMacroFiNoteType(dom.noteType.value);
  }

  function isShortNoteSelected() {
    return dom.noteType.value === "Short Note / Market Alert";
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

    if (!nn) return "";
    if (!cc) return nn;
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

  function syncIssuerId() {
    if (!dom.issuerId) return;
    const selectedCountry = String(dom.coverageCountry?.value || "").trim().toUpperCase();
    dom.issuerId.value = selectedCountry ? (ISSUER_COUNTRY_MAP[selectedCountry] || buildIssuerIdFromCountry(selectedCountry)) : "";
  }

  function buildIssuerIdFromCountry(countryCode) {
    const normalized = String(countryCode || "").trim().toUpperCase();
    if (!normalized) return "";
    const total = normalized.split("").reduce((sum, char, index) => sum + (char.charCodeAt(0) * (index + 3)), 0);
    return String(((total * 7919) % 900000) + 100000);
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

  function initializeBodySectionEditor() {
    syncBodySectionVisibility();
  }

  function addCustomBodySection(seed = {}, shouldAppend = true) {
    if (!dom.customBodySectionTemplate || !dom.bodySections) return null;

    state.customSectionCount += 1;
    const fragment = dom.customBodySectionTemplate.content.cloneNode(true);
    const card = fragment.querySelector(".body-section-card");
    const key = seed.key || `custom-${state.customSectionCount}`;
    card.dataset.sectionKey = key;

    const headingInput = fragment.querySelector("[data-custom-field='heading']");
    const contentInput = fragment.querySelector("[data-custom-field='content']");
    if (headingInput) headingInput.value = seed.label || "Custom Section";
    if (contentInput) contentInput.value = seed.content || "";

    if (shouldAppend) {
      dom.bodySections.appendChild(fragment);
      return dom.bodySections.querySelector(`[data-section-key='${key}']`);
    }

    return card;
  }

  function handleBodySectionActions(event) {
    const actionButton = event.target.closest("[data-body-action]");
    if (!actionButton || !dom.bodySections) return;

    const card = actionButton.closest(".body-section-card");
    if (!card) return;

    const action = actionButton.getAttribute("data-body-action");
    if (action === "move-up") {
      const previous = card.previousElementSibling;
      if (previous) dom.bodySections.insertBefore(card, previous);
    } else if (action === "move-down") {
      const next = card.nextElementSibling;
      if (next) dom.bodySections.insertBefore(next, card);
    } else if (action === "remove" && card.dataset.sectionRole === "custom") {
      card.remove();
    } else {
      return;
    }

    syncFigurePlacementControls();
    updateAllUI();
    queueDraftSave();
  }

  function handleBodySectionInput(event) {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)) return;

    if (
      target.closest(".body-section-heading") ||
      target.getAttribute("data-custom-field") === "heading"
    ) {
      syncFigurePlacementControls();
    }
  }

  function getBodySectionContent(card) {
    const sectionKey = card.dataset.sectionKey;
    if (card.dataset.sectionRole === "custom") {
      const contentInput = card.querySelector("[data-custom-field='content']");
      return String(contentInput?.value || "").trim();
    }

    const element = document.getElementById(sectionKey);
    return String(element?.value || "").trim();
  }

  function getBodySectionHeading(card) {
    if (card.dataset.sectionRole === "custom") {
      return String(card.querySelector("[data-custom-field='heading']")?.value || "").trim() || "Custom Section";
    }

    const headingInput = card.querySelector(".body-section-heading input");
    return String(headingInput?.value || "").trim() || defaultHeadingForSection(card.dataset.sectionKey);
  }

  function defaultHeadingForSection(sectionKey) {
    const defaults = {
      keyTakeaways: "Key Takeaways",
      analysis: "Analysis And Commentary",
      content: "Additional Detail",
      fiveYearRationale: "5 Year Rationale & Price Target",
      esgSummary: "ESG Summary",
      cordobaView: "The Cordoba View"
    };

    return defaults[sectionKey] || "Section";
  }

  function isEquityOnlySectionKey(sectionKey) {
    return sectionKey === "fiveYearRationale" || sectionKey === "esgSummary";
  }

  function resolveBodySectionContent(data, sectionKey, fallbackContent = "") {
    const map = {
      keyTakeaways: data.keyTakeaways,
      analysis: data.analysis,
      content: data.content,
      fiveYearRationale: data.fiveYearRationale,
      esgSummary: data.esgSummary,
      cordobaView: data.cordobaView
    };

    return String(map[sectionKey] ?? fallbackContent ?? "").trim();
  }

  function normalizeBodySectionLayoutForExport(data) {
    const source = Array.isArray(data.bodySectionLayout) && data.bodySectionLayout.length
      ? data.bodySectionLayout
      : [
        { key: "keyTakeaways", label: defaultHeadingForSection("keyTakeaways"), role: "core" },
        { key: "analysis", label: defaultHeadingForSection("analysis"), role: "core" },
        { key: "content", label: defaultHeadingForSection("content"), role: "core" },
        ...(data.noteType === "Equity Research"
          ? [
            { key: "fiveYearRationale", label: defaultHeadingForSection("fiveYearRationale"), role: "core" },
            { key: "esgSummary", label: defaultHeadingForSection("esgSummary"), role: "core" }
          ]
          : []),
        { key: "cordobaView", label: defaultHeadingForSection("cordobaView"), role: "core" }
      ];

    return source
      .map((entry) => {
        const key = String(entry?.key || "").trim();
        if (!key) return null;

        const role = entry?.role === "custom" || key.startsWith("custom-") ? "custom" : "core";
        const label = String(entry?.label || defaultHeadingForSection(key)).trim() || defaultHeadingForSection(key);
        const content = role === "custom"
          ? String(entry?.content || "").trim()
          : resolveBodySectionContent(data, key, entry?.content);
        const hidden = Boolean(entry?.hidden) || (isEquityOnlySectionKey(key) && data.noteType !== "Equity Research");

        return { key, label, content, role, hidden };
      })
      .filter(Boolean);
  }

  function getNarrativeSectionPartitions(data) {
    const layout = normalizeBodySectionLayoutForExport(data)
      .filter((entry) => !entry.hidden)
      .filter((entry) => entry.role === "custom" || entry.content || entry.key === "keyTakeaways" || entry.key === "analysis");

    const keyTakeaways = layout.find((entry) => entry.key === "keyTakeaways") || {
      key: "keyTakeaways",
      label: defaultHeadingForSection("keyTakeaways"),
      content: data.keyTakeaways
    };

    const cordobaView = layout.find((entry) => entry.key === "cordobaView" && entry.content);
    const middle = layout.filter((entry) => entry.key !== "keyTakeaways" && entry.key !== "cordobaView");

    return { keyTakeaways, middle, cordobaView };
  }

  function buildNarrativeSectionBlocks(docxLib, colors, sections) {
    const blocks = [];

    sections.forEach((section) => {
      const label = String(section?.label || defaultHeadingForSection(section?.key)).trim();
      const content = String(section?.content || "").trim();
      if (!label || !content) return;

      blocks.push(
        buildNomuraSubhead(docxLib, colors, label),
        ...buildNomuraBodyParagraphs(docxLib, content)
      );
    });

    return blocks;
  }

  function collectBodySectionLayout() {
    if (!dom.bodySections) return [];

    return Array.from(dom.bodySections.querySelectorAll(".body-section-card"))
      .map((card) => ({
        key: card.dataset.sectionKey || "",
        label: getBodySectionHeading(card),
        content: getBodySectionContent(card),
        role: card.dataset.sectionRole || "core",
        hidden: Boolean(card.hidden)
      }));
  }

  function restoreBodySectionLayout(layout = []) {
    if (!dom.bodySections) return;

    Array.from(dom.bodySections.querySelectorAll(".body-section-card.is-custom")).forEach((card) => card.remove());

    const coreCards = new Map(
      Array.from(dom.bodySections.querySelectorAll(".body-section-card"))
        .filter((card) => card.dataset.sectionRole !== "custom")
        .map((card) => [card.dataset.sectionKey, card])
    );

    const orderedCards = [];
    layout.forEach((entry) => {
      if (!entry?.key) return;

      if (entry.role === "custom" || entry.key.startsWith("custom-")) {
        const customCard = addCustomBodySection(entry, false);
        if (customCard) orderedCards.push(customCard);
        return;
      }

      const card = coreCards.get(entry.key);
      if (!card) return;
      orderedCards.push(card);
      coreCards.delete(entry.key);
    });

    orderedCards.push(...Array.from(coreCards.values()));
    orderedCards.forEach((card) => dom.bodySections.appendChild(card));
    syncBodySectionVisibility();
    syncFigurePlacementControls();
  }

  function addRatingProfileRow(seed = {}) {
    if (!dom.ratingProfileRowTemplate || !dom.ratingsProfileRows) return null;
    const fragment = dom.ratingProfileRowTemplate.content.cloneNode(true);
    const row = fragment.querySelector(".ratings-profile-row");
    row.querySelector("[data-rating-field='agency']").value = seed.agency || "";
    row.querySelector("[data-rating-field='short']").value = seed.shortTerm || "";
    row.querySelector("[data-rating-field='long']").value = seed.longTerm || "";
    dom.ratingsProfileRows.appendChild(fragment);
    return dom.ratingsProfileRows.lastElementChild;
  }

  function handleRatingProfileActions(event) {
    const actionButton = event.target.closest("[data-rating-action]");
    if (!actionButton) return;

    if (actionButton.getAttribute("data-rating-action") === "remove") {
      actionButton.closest(".ratings-profile-row")?.remove();
      updateAllUI();
      queueDraftSave();
    }
  }

  function collectRatingsProfile() {
    const rows = [];
    const firstRow = {
      agency: String(dom.agencyRating?.value || "").trim(),
      shortTerm: String(dom.shortTermRating?.value || "").trim(),
      longTerm: String(dom.longTermRating?.value || "").trim()
    };
    if (firstRow.agency || firstRow.shortTerm || firstRow.longTerm) rows.push(firstRow);

    rows.push(
      ...Array.from(dom.ratingsProfileRows?.querySelectorAll(".ratings-profile-row") || []).map((row) => ({
        agency: String(row.querySelector("[data-rating-field='agency']")?.value || "").trim(),
        shortTerm: String(row.querySelector("[data-rating-field='short']")?.value || "").trim(),
        longTerm: String(row.querySelector("[data-rating-field='long']")?.value || "").trim()
      }))
    );

    return rows.filter((row) => row.agency || row.shortTerm || row.longTerm);
  }

  function restoreRatingsProfile(rows = []) {
    if (!dom.ratingsProfileRows) return;
    dom.ratingsProfileRows.innerHTML = "";

    const [primaryRow, ...extraRows] = Array.isArray(rows) ? rows : [];
    dom.agencyRating.value = primaryRow?.agency || "";
    dom.shortTermRating.value = primaryRow?.shortTerm || "";
    dom.longTermRating.value = primaryRow?.longTerm || "";
    extraRows.forEach((row) => addRatingProfileRow(row));
  }

  function initializeFinancialTableEditor(forceReset = false) {
    const publicationDate = parseInputDate(dom.publicationDate.value) || new Date();
    const rawValue = forceReset ? "" : dom.financialTableInput.value.trim();
    state.financialTable = parseFinancialTableInput(rawValue, publicationDate);
    formatEntireFinancialTableDisplay();
    syncFinancialTableStorage();
    renderFinancialTableEditor();
  }

  function syncFinancialTableStorage() {
    const publicationDate = parseInputDate(dom.publicationDate.value) || new Date();
    state.financialTable = normalizeFinancialMatrix(state.financialTable, publicationDate);
    dom.financialTableInput.value = JSON.stringify(state.financialTable);
  }

  function renderFinancialTableEditor() {
    if (!dom.financialTableHead || !dom.financialTableBody) return;

    const matrix = normalizeFinancialMatrix(state.financialTable, parseInputDate(dom.publicationDate.value) || new Date());
    state.financialTable = matrix;

    dom.financialTableHead.innerHTML = `
      <tr>
        <th class="financial-grid-metric-col">
          <span class="financial-grid-label">Line Item</span>
        </th>
        ${matrix.headers.map((header, index) => `
          <th class="financial-grid-period-col">
            <div class="financial-header-cell">
              <span class="financial-grid-label">Period ${index + 1}</span>
              <input
                type="text"
                value="${escapeAttribute(header)}"
                data-kind="header"
                data-col-index="${index}"
                placeholder="FY26F"
              >
              <div class="financial-header-actions">
                <button type="button" class="btn btn-ghost btn-xs" data-action="move-col-left" data-col-index="${index}" ${index === 0 ? "disabled" : ""}>&lt;</button>
                <button type="button" class="btn btn-ghost btn-xs" data-action="move-col-right" data-col-index="${index}" ${index === matrix.headers.length - 1 ? "disabled" : ""}>&gt;</button>
                <button type="button" class="btn btn-ghost btn-xs" data-action="delete-col" data-col-index="${index}" ${matrix.headers.length === 1 ? "disabled" : ""}>Del</button>
              </div>
            </div>
          </th>
        `).join("")}
        <th class="financial-grid-actions-col">
          <span class="financial-grid-label">Row Order</span>
        </th>
      </tr>
    `;

    dom.financialTableBody.innerHTML = matrix.rows.map((row, rowIndex) => `
      <tr>
        <td class="financial-grid-metric-col">
          <div class="financial-row-cell">
            <span class="financial-grid-label">Metric ${rowIndex + 1}</span>
            <input
              type="text"
              class="financial-row-label"
              value="${escapeAttribute(row.label)}"
              data-kind="row-label"
              data-row-index="${rowIndex}"
              placeholder="Revenue (mn)"
            >
          </div>
        </td>
        ${row.values.map((value, colIndex) => `
          <td class="financial-grid-period-col">
            <input
              type="text"
              class="financial-cell-input is-${classifyFinancialMetric(row.label)}"
              value="${escapeAttribute(value)}"
              data-kind="value"
              data-row-index="${rowIndex}"
              data-col-index="${colIndex}"
              placeholder="-"
            >
          </td>
        `).join("")}
        <td class="financial-grid-actions-col">
          <div class="financial-row-actions">
            <button type="button" class="btn btn-ghost btn-xs" data-action="move-row-up" data-row-index="${rowIndex}" ${rowIndex === 0 ? "disabled" : ""}>Up</button>
            <button type="button" class="btn btn-ghost btn-xs" data-action="move-row-down" data-row-index="${rowIndex}" ${rowIndex === matrix.rows.length - 1 ? "disabled" : ""}>Dn</button>
            <button type="button" class="btn btn-ghost btn-xs" data-action="delete-row" data-row-index="${rowIndex}" ${matrix.rows.length === 1 ? "disabled" : ""}>Del</button>
          </div>
        </td>
      </tr>
    `).join("");

    dom.financialAddColumnBtn.disabled = matrix.headers.length >= 6;
  }

  function handleFinancialGridInput(event) {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;

    const kind = target.getAttribute("data-kind");
    if (!kind || !state.financialTable) return;

    const rowIndex = Number(target.getAttribute("data-row-index"));
    const colIndex = Number(target.getAttribute("data-col-index"));

    if (kind === "header" && Number.isInteger(colIndex)) {
      state.financialTable.headers[colIndex] = target.value.trim() || `Period ${colIndex + 1}`;
    } else if (kind === "row-label" && Number.isInteger(rowIndex)) {
      state.financialTable.rows[rowIndex].label = target.value;
    } else if (kind === "value" && Number.isInteger(rowIndex) && Number.isInteger(colIndex)) {
      state.financialTable.rows[rowIndex].values[colIndex] = target.value;
    }

    syncFinancialTableStorage();
  }

  function handleFinancialGridFocusOut(event) {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;

    const kind = target.getAttribute("data-kind");
    const rowIndex = Number(target.getAttribute("data-row-index"));
    const colIndex = Number(target.getAttribute("data-col-index"));

    if (kind === "value" && Number.isInteger(rowIndex) && Number.isInteger(colIndex)) {
      const rowLabel = state.financialTable.rows[rowIndex]?.label || "";
      const formattedValue = formatFinancialValueForDisplay(target.value, rowLabel);
      state.financialTable.rows[rowIndex].values[colIndex] = formattedValue;
      target.value = formattedValue;
      syncFinancialTableStorage();
      return;
    }

    if (kind === "row-label" && Number.isInteger(rowIndex)) {
      formatFinancialRowValues(rowIndex);
      syncFinancialTableStorage();
      renderFinancialTableEditor();
      updateAllUI();
      queueDraftSave();
    }
  }

  function handleFinancialGridClick(event) {
    const actionButton = event.target.closest("[data-action]");
    if (!actionButton || !state.financialTable) return;

    const action = actionButton.getAttribute("data-action");
    const rowIndex = Number(actionButton.getAttribute("data-row-index"));
    const colIndex = Number(actionButton.getAttribute("data-col-index"));

    if (action === "move-col-left" && colIndex > 0) {
      moveFinancialColumn(colIndex, colIndex - 1);
    } else if (action === "move-col-right" && colIndex < state.financialTable.headers.length - 1) {
      moveFinancialColumn(colIndex, colIndex + 1);
    } else if (action === "delete-col") {
      removeFinancialColumn(colIndex);
    } else if (action === "move-row-up" && rowIndex > 0) {
      moveFinancialRow(rowIndex, rowIndex - 1);
    } else if (action === "move-row-down" && rowIndex < state.financialTable.rows.length - 1) {
      moveFinancialRow(rowIndex, rowIndex + 1);
    } else if (action === "delete-row") {
      removeFinancialRow(rowIndex);
    } else {
      return;
    }

    syncFinancialTableStorage();
    renderFinancialTableEditor();
    updateAllUI();
    queueDraftSave();
  }

  function addFinancialPeriodColumn() {
    if (!state.financialTable) initializeFinancialTableEditor();
    if (state.financialTable.headers.length >= 6) {
      setMessage("error", "Financial forecast export supports up to six period columns in the current layout.");
      return;
    }

    const nextHeader = buildNextFinancialHeaderLabel(state.financialTable.headers, parseInputDate(dom.publicationDate.value) || new Date());
    state.financialTable.headers.push(nextHeader);
    state.financialTable.rows.forEach((row) => row.values.push(""));
    syncFinancialTableStorage();
    renderFinancialTableEditor();
    updateAllUI();
    queueDraftSave();
  }

  function addFinancialLineItem() {
    if (!state.financialTable) initializeFinancialTableEditor();
    state.financialTable.rows.push({
      label: "",
      values: Array.from({ length: state.financialTable.headers.length }, () => "")
    });
    syncFinancialTableStorage();
    renderFinancialTableEditor();
    updateAllUI();
    queueDraftSave();
  }

  function resetFinancialTableGrid() {
    const shouldReset = !state.financialTable || !hasMeaningfulFinancialData(state.financialTable)
      ? true
      : window.confirm("Reset the financial forecast grid to the default institutional template?");

    if (!shouldReset) return;

    initializeFinancialTableEditor(true);
    updateAllUI();
    queueDraftSave();
  }

  async function downloadFinancialTemplate() {
    if (!ensureXlsxAvailable("download the Excel template")) return;

    const publicationDate = parseInputDate(dom.publicationDate.value) || new Date();
    const matrix = pruneEmptyFinancialRows(normalizeFinancialMatrix(state.financialTable, publicationDate));
    const workbook = buildFinancialTemplateWorkbook(matrix, {
      title: dom.financialTableTitle.value.trim() || "Year-end 31 Dec",
      sourceNote: dom.financialSourceNote.value.trim() || "Source: Company data, Cordoba Research Group estimates"
    });
    const tickerSlug = slugify(dom.ticker.value.trim() || "financial_forecast");
    window.XLSX.writeFile(workbook, `${tickerSlug}_forecast_template.xlsx`);
    setMessage("success", "Excel forecast template downloaded. Fill in the workbook and upload it back to replace the grid.");
  }

  async function handleFinancialTemplateUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!ensureXlsxAvailable("import the completed Excel template")) return;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, { type: "array", raw: false });
      state.financialTable = readFinancialMatrixFromWorkbook(workbook, parseInputDate(dom.publicationDate.value) || new Date());
      formatEntireFinancialTableDisplay();
      syncFinancialTableStorage();
      renderFinancialTableEditor();
      updateAllUI();
      queueDraftSave();
      setMessage("success", `Financial forecast grid updated from ${file.name}.`);
    } catch (error) {
      console.error("Financial template import failed:", error);
      setMessage("error", "The uploaded Excel template could not be parsed. Use the downloaded CRG template structure and try again.");
    } finally {
      event.target.value = "";
    }
  }

  function ensureXlsxAvailable(actionLabel) {
    if (typeof window.XLSX !== "undefined") return true;
    setMessage("error", `The Excel workflow library is unavailable, so the tool cannot ${actionLabel} in this session.`);
    return false;
  }

  function buildFinancialTemplateWorkbook(matrix, meta) {
    const workbook = window.XLSX.utils.book_new();
    const forecastSheet = window.XLSX.utils.aoa_to_sheet(financialMatrixToAoA(matrix));
    forecastSheet["!cols"] = [
      { wch: 30 },
      ...matrix.headers.map(() => ({ wch: 14 }))
    ];

    const instructionsSheet = window.XLSX.utils.aoa_to_sheet([
      ["Cordoba Research Group Financial Forecast Template"],
      [""],
      ["1. Update the Forecast Input sheet only."],
      ["2. Keep the first row for period headers and the first column for line items."],
      ["3. Use one row per metric and one column per period."],
      ["4. Percent rows can be entered as 12.5 or 12.5%; the tool will normalize formatting."],
      ["5. Upload the completed workbook back into the research production tool."],
      [""],
      [`Caption: ${meta.title}`],
      [`Source note: ${meta.sourceNote}`]
    ]);
    instructionsSheet["!cols"] = [{ wch: 88 }];

    window.XLSX.utils.book_append_sheet(workbook, forecastSheet, "Forecast Input");
    window.XLSX.utils.book_append_sheet(workbook, instructionsSheet, "Instructions");
    return workbook;
  }

  function financialMatrixToAoA(matrix) {
    return [
      ["Metric", ...(matrix?.headers || [])],
      ...(matrix?.rows || []).map((row) => [
        row.label || "",
        ...(row.values || []).map((value) => formatFinancialValueForDisplay(value, row.label))
      ])
    ];
  }

  function readFinancialMatrixFromWorkbook(workbook, publicationDate) {
    const targetSheetName = workbook.SheetNames.find((name) => /forecast/i.test(name)) || workbook.SheetNames[0];
    const sheet = workbook.Sheets[targetSheetName];
    const rows = window.XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
      raw: false
    });

    return createFinancialMatrixFromRows(rows, publicationDate);
  }

  function createFinancialMatrixFromRows(rows, publicationDate) {
    const cleanedRows = (rows || [])
      .map((row) => Array.isArray(row) ? row.map((cell) => String(cell ?? "").trim()) : [])
      .filter((row) => row.some((cell) => cell));

    if (!cleanedRows.length) return normalizeFinancialMatrix(null, publicationDate);

    const headerRow = cleanedRows[0];
    const headers = headerRow.slice(1, 7).filter(Boolean);
    const bodyRows = cleanedRows
      .slice(1)
      .map((cells) => ({
        label: cells[0] || "",
        values: normalizeFinancialValues(cells.slice(1), headers.length || buildDefaultFinancialHeaders(publicationDate).length)
      }));

    return normalizeFinancialMatrix({ headers, rows: bodyRows }, publicationDate);
  }

  function moveFinancialColumn(fromIndex, toIndex) {
    const headers = state.financialTable.headers;
    const [header] = headers.splice(fromIndex, 1);
    headers.splice(toIndex, 0, header);

    state.financialTable.rows.forEach((row) => {
      const [value] = row.values.splice(fromIndex, 1);
      row.values.splice(toIndex, 0, value);
    });
  }

  function removeFinancialColumn(index) {
    if (state.financialTable.headers.length === 1) return;
    state.financialTable.headers.splice(index, 1);
    state.financialTable.rows.forEach((row) => {
      row.values.splice(index, 1);
    });
  }

  function moveFinancialRow(fromIndex, toIndex) {
    const [row] = state.financialTable.rows.splice(fromIndex, 1);
    state.financialTable.rows.splice(toIndex, 0, row);
  }

  function removeFinancialRow(index) {
    if (state.financialTable.rows.length === 1) {
      state.financialTable.rows = [{
        label: "",
        values: Array.from({ length: state.financialTable.headers.length }, () => "")
      }];
      return;
    }
    state.financialTable.rows.splice(index, 1);
  }

  function buildNextFinancialHeaderLabel(headers, publicationDate) {
    const years = extractFinancialHeaderYears(headers);
    const nextYear = years.length ? Math.max(...years) + 1 : publicationDate.getFullYear() + 1;
    return `FY${String(nextYear).slice(-2)}F`;
  }

  function formatEntireFinancialTableDisplay() {
    if (!state.financialTable) return;
    state.financialTable.rows.forEach((row, rowIndex) => {
      formatFinancialRowValues(rowIndex);
    });
  }

  function formatFinancialRowValues(rowIndex) {
    const row = state.financialTable?.rows?.[rowIndex];
    if (!row) return;
    row.values = row.values.map((value) => formatFinancialValueForDisplay(value, row.label));
  }

  function classifyFinancialMetric(label) {
    const normalized = normalizeComparableText(label);
    if (!normalized) return "numeric";

    if (
      normalized.includes("%") ||
      /\bgrowth\b|\bmargin\b|\broe\b|\broic\b|\breturn\b|\byield\b|\bupside\b|\bdownside\b|\btax rate\b/.test(normalized)
    ) {
      return "percent";
    }

    if (/\beps\b/.test(normalized)) return "eps";

    if (
      normalized.includes("(x)") ||
      /\bpe\b|\bp e\b|ev ebitda|price book|price to book|multiple|ratio\b/.test(normalized)
    ) {
      return "multiple";
    }

    if (
      /\brevenue\b|\bprofit\b|\bebit\b|\bebitda\b|\bcash\b|\bdebt\b|\bcapex\b|\bfcf\b|\bdividend\b|\bbook value\b|\bassets\b|\bliabilities\b|\bnet income\b/.test(normalized)
    ) {
      return "monetary";
    }

    return "numeric";
  }

  function formatFinancialValueForDisplay(value, rowLabel) {
    const text = String(value ?? "").trim();
    if (!text || text === "-") return "";
    if (/^(n\/a|nm|n\.m\.|na)$/i.test(text)) return text.toUpperCase();
    if (/^[=]/.test(text)) return text;

    const numericText = text.replace(/,/g, "").replace(/%/g, "").trim();
    if (!/^-?\d*\.?\d+$/.test(numericText)) return text;

    const number = Number(numericText);
    if (!Number.isFinite(number)) return text;

    const metricType = classifyFinancialMetric(rowLabel);
    const decimalPart = numericText.includes(".") ? numericText.split(".")[1] : "";
    const decimals = decimalPart.length;

    if (metricType === "percent") {
      return `${formatNumericForDisplay(number, decimals)}%`;
    }

    if (metricType === "monetary") {
      return formatNumericForDisplay(number, decimals);
    }

    if (metricType === "eps" || metricType === "multiple" || metricType === "numeric") {
      return formatNumericForDisplay(number, decimals);
    }

    return text;
  }

  function formatNumericForDisplay(number, decimals = 0) {
    const safeDecimals = Math.min(Math.max(decimals, 0), 4);
    return new Intl.NumberFormat("en-US", {
      minimumFractionDigits: safeDecimals,
      maximumFractionDigits: safeDecimals
    }).format(number);
  }

  function toggleEquitySection() {
    const showEquity = isEquitySelected();
    dom.equitySection.hidden = !showEquity;
    dom.navEquity.textContent = showEquity ? buildSectionCompletion("equity") : "Optional";
  }

  function toggleMacroFiPanel() {
    if (!dom.macroFiPanel) return;
    dom.macroFiPanel.hidden = !isMacroFiSelected();
  }

  function syncBodySectionVisibility() {
    if (!dom.bodySections) return;
    const showEquityOnly = isEquitySelected();
    dom.bodySections.querySelectorAll(".is-equity-only").forEach((section) => {
      section.hidden = !showEquityOnly;
    });
  }

  function toggleNoteTypeSections() {
    toggleEquitySection();
    toggleMacroFiPanel();
    syncIssuerId();
    syncBodySectionVisibility();
  }

  function getRequiredIds() {
    const requiredIds = [...BASE_REQUIRED_IDS];
    if (dom.noteType.value === "Macro Research") requiredIds.push("coverageCountry", "issuerId");
    if (isEquitySelected()) return requiredIds.concat(EQUITY_REQUIRED_IDS);
    return requiredIds;
  }

  function getSectionRequirementIds(sectionKey) {
    if (sectionKey === "note") {
      const ids = [...SECTION_REQUIREMENTS.note];
      if (dom.noteType.value === "Macro Research") ids.push("coverageCountry", "issuerId");
      return ids;
    }

    if (sectionKey === "equity") {
      return isEquitySelected() ? EQUITY_REQUIRED_IDS : [];
    }

    return SECTION_REQUIREMENTS[sectionKey] || [];
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

      if (valid && (id === "targetPrice" || id === "marketCapUsd")) {
        valid = parseNumber(element.value) != null;
      }

      if (!valid) {
        missing.push({
          id,
          label:
            (id === "targetPrice" || id === "marketCapUsd") && isFilled(element)
              ? `${FIELD_LABELS[id] || id} must be numeric`
              : (FIELD_LABELS[id] || id),
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

  function normalizeComparableText(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function countMeaningfulWords(text) {
    return normalizeComparableText(text)
      .split(/\s+/)
      .filter(Boolean).length;
  }

  function containsAnyKeyword(text, keywords) {
    const source = normalizeComparableText(text);
    return keywords.some((keyword) => source.includes(normalizeComparableText(keyword)));
  }

  function buildFinding(severity, title, detail, section, options = {}) {
    return {
      severity,
      title,
      detail,
      section,
      blocking: severity === "critical",
      focusId: options.focusId || "",
      sortOrder: severity === "critical" ? 0 : severity === "warning" ? 1 : 2
    };
  }

  function extractFinancialHeaderYears(headers) {
    return (headers || [])
      .map((header) => {
        const text = String(header || "").trim();
        const fullYearMatch = text.match(/(20\d{2})/);
        if (fullYearMatch) return Number(fullYearMatch[1]);

        const shortYearMatch = text.match(/(\d{2})(?=[A-Z]?$)/i);
        if (!shortYearMatch) return null;
        const shortYear = Number(shortYearMatch[1]);
        return shortYear >= 70 ? 1900 + shortYear : 2000 + shortYear;
      })
      .filter((year) => Number.isFinite(year));
  }

  function isFinancialTableStale(matrix, publicationDate) {
    const headers = matrix?.headers || [];
    const normalizedHeaders = headers.map((header) => String(header || "").toUpperCase());
    const years = extractFinancialHeaderYears(headers);
    const publicationYear = publicationDate.getFullYear();
    const hasForecastColumn = normalizedHeaders.some((header) => header.includes("F"));
    const maxYear = years.length ? Math.max(...years) : null;

    if (!headers.length) return true;
    if (!hasForecastColumn) return true;
    if (!Number.isFinite(maxYear)) return true;
    return maxYear < publicationYear + 1;
  }

  function isLikelyGenericCordobaView(text) {
    const normalized = normalizeComparableText(text);
    if (!normalized) return true;
    if (countMeaningfulWords(text) < 12) return true;

    const genericPhrases = [
      "we remain constructive",
      "continue to monitor",
      "stay selective",
      "house view remains constructive",
      "watch the data",
      "further updates to follow",
      "positive on the story",
      "balanced view"
    ];

    return genericPhrases.some((phrase) => normalized.includes(phrase));
  }

  function buildPrePublishReview(data, validation = validateForm(false)) {
    const findings = [];
    const noteType = String(data.noteType || "");
    const missingIds = new Set(validation.missing.map((entry) => entry.id));
    const combinedResearchText = [
      ...normalizeBodySectionLayoutForExport(data).map((entry) => entry.content),
      data.businessDescription,
      data.valuationSummary
    ]
      .filter(Boolean)
      .join("\n");
    const publicationDate = parseInputDate(data.publicationDate) || data.generatedAt || new Date();

    validation.missing.forEach((entry) => {
      findings.push(
        buildFinding(
          "critical",
          entry.label,
          `Complete this ${entry.section.toLowerCase()} field before export.`,
          entry.section,
          { focusId: entry.id }
        )
      );
    });

    if (noteType === "Macro Research" || noteType === "Fixed Income Research") {
      const ratingsRows = Array.isArray(data.ratingsProfile) ? data.ratingsProfile : [];
      if (!ratingsRows.length) {
        findings.push(
          buildFinding(
            noteType === "Fixed Income Research" ? "warning" : "suggestion",
            "Ratings table not populated",
            "Add the agency, short-term, and long-term ratings profile so the note header carries the structured sovereign / issuer ratings block.",
            "Brief",
            { focusId: "agencyRating" }
          )
        );
      }

      const partialRatingsRow = ratingsRows.find((row) => (row.agency || row.shortTerm || row.longTerm) && (!row.agency || (!row.shortTerm && !row.longTerm)));
      if (partialRatingsRow) {
        findings.push(
          buildFinding(
            "warning",
            "Ratings row is incomplete",
            "Each ratings row should include the agency and at least one of the short-term or long-term ratings.",
            "Brief",
            { focusId: "agencyRating" }
          )
        );
      }
    }

    if (noteType === "Equity Research") {
      const targetPrice = parseNumber(data.targetPrice);
      const valuationWordCount = countMeaningfulWords(data.valuationSummary);
      const rationaleWordCount = countMeaningfulWords(data.fiveYearRationale);
      const businessWordCount = countMeaningfulWords(data.businessDescription);
      const esgWordCount = countMeaningfulWords(data.esgSummary);
      const financialMatrix = parseFinancialTableInput(data.financialTableInput, publicationDate);
      if (targetPrice != null && rationaleWordCount < 18) {
        findings.push(
          buildFinding(
            "critical",
            "Target price rationale missing",
            "A target price is set, but the 5 Year Rationale & Price Target section does not yet explain the basis for that target.",
            "Equity",
            { focusId: "fiveYearRationale" }
          )
        );
      }

      if (!missingIds.has("businessDescription") && businessWordCount < 18) {
        findings.push(
          buildFinding(
            "warning",
            "Business description is still light",
            "Add a clearer two- or three-sentence description of what the business does and where the core exposure sits.",
            "Equity",
            { focusId: "businessDescription" }
          )
        );
      }

      if (valuationWordCount < 18) {
        findings.push(
          buildFinding(
            "warning",
            "Valuation summary needs more support",
            "Add more detail on fair value, dislocation, and the core valuation driver so the note reads more like a sell-side product.",
            "Equity",
            { focusId: "valuationSummary" }
          )
        );
      }

      if (esgWordCount < 12) {
        findings.push(
          buildFinding(
            "suggestion",
            "ESG summary is limited",
            "Add material governance, sustainability, or environmental risk points so the equity note captures the main ESG considerations.",
            "Equity",
            { focusId: "esgSummary" }
          )
        );
      }

      if (data.crgRating && !missingIds.has("analysis") && !containsAnyKeyword(combinedResearchText, [
        "catalyst",
        "trigger",
        "driver",
        "earnings",
        "margin",
        "order",
        "pricing",
        "volume",
        "rerating",
        "re-rating",
        "valuation"
      ])) {
        findings.push(
          buildFinding(
            "warning",
            "Rating set with limited catalyst support",
            "The note carries a formal rating, but the draft does not clearly surface catalysts or investment-case support.",
            "Equity",
            { focusId: "analysis" }
          )
        );
      }

      if (!data.benchmarkTicker.trim()) {
        findings.push(
          buildFinding(
            "warning",
            "Benchmark missing",
            "Add a benchmark ticker so the tear sheet can show relative performance rather than standalone price action.",
            "Equity",
            { focusId: "benchmarkTicker" }
          )
        );
      }

      if (!hasMeaningfulFinancialData(financialMatrix)) {
        findings.push(
          buildFinding(
            "warning",
            "Financial table not populated",
            "Add live forecast values into the financial grid so the equity note renders with actual financial estimates rather than placeholders.",
            "Equity",
            { focusId: "financialTableEditor" }
          )
        );
      } else if (isFinancialTableStale(financialMatrix, publicationDate)) {
        findings.push(
          buildFinding(
            "warning",
            "Financial table may be stale",
            "The financial headers look historical-only or do not extend far enough beyond the publication year.",
            "Equity",
            { focusId: "financialTableEditor" }
          )
        );
      }
    }

    if (normalizeComparableText(data.deck) && normalizeComparableText(data.deck) === normalizeComparableText(data.title)) {
      findings.push(
        buildFinding(
          "suggestion",
          "Deck duplicates the title",
          "Use the deck to add differentiated framing, context, or timing instead of repeating the research title verbatim.",
          "Brief",
          { focusId: "deck" }
        )
      );
    }

    if (!missingIds.has("keyTakeaways") && (lineItems(data.keyTakeaways).length < 2 || countMeaningfulWords(data.keyTakeaways) < 14)) {
      findings.push(
        buildFinding(
          "warning",
          "Key Takeaways are thin",
          "Add clearer bullet points so the note opens with a sharper summary of the thesis, catalysts, and call.",
          "Research",
          { focusId: "keyTakeaways" }
        )
      );
    }

    const minimumAnalysisWords = noteType === "Short Note / Market Alert" ? 65 : 120;

    if (!missingIds.has("analysis") && countMeaningfulWords(data.analysis) < minimumAnalysisWords) {
      findings.push(
        buildFinding(
          "warning",
          "Analysis needs more development",
          "The analysis section is still too light for a publication-ready research note.",
          "Research",
          { focusId: "analysis" }
        )
      );
    }

    if (!missingIds.has("cordobaView") && isLikelyGenericCordobaView(data.cordobaView)) {
      findings.push(
        buildFinding(
          "warning",
          "Cordoba View is too generic",
          "Tighten the house view so it states the action, horizon, and what would change conviction.",
          "Research",
          { focusId: "cordobaView" }
        )
      );
    }

    if (noteType === "Macro Research" && !missingIds.has("analysis") && !containsAnyKeyword(combinedResearchText, [
      "inflation",
      "rates",
      "growth",
      "policy",
      "payroll",
      "cpi",
      "pmi",
      "scenario"
    ])) {
      findings.push(
        buildFinding(
          "suggestion",
          "Macro framing could be sharper",
          "The draft would read more institutionally with an explicit macro driver set: growth, inflation, rates, policy, or scenario markers.",
          "Research",
          { focusId: "analysis" }
        )
      );
    }

    if (noteType === "Fixed Income Research" && !missingIds.has("analysis") && !containsAnyKeyword(combinedResearchText, [
      "spread",
      "yield",
      "carry",
      "curve",
      "duration",
      "refinancing",
      "maturity",
      "coupon"
    ])) {
      findings.push(
        buildFinding(
          "suggestion",
          "Fixed income lens not yet explicit",
          "Consider surfacing spread, yield, carry, curve, or refinancing language so the FI call reads as desk-specific rather than general commentary.",
          "Research",
          { focusId: "analysis" }
        )
      );
    }

    if ((noteType === "Commodity Insights" || noteType === "Commodity Research") && !missingIds.has("analysis") && !containsAnyKeyword(combinedResearchText, [
      "supply",
      "demand",
      "inventory",
      "balance",
      "curve",
      "production",
      "opec",
      "stock"
    ])) {
      findings.push(
        buildFinding(
          "suggestion",
          "Commodity balance framing is light",
          "Highlight supply, demand, inventory, or curve dynamics so the commodity view feels more complete.",
          "Research",
          { focusId: "analysis" }
        )
      );
    }

    findings.sort((left, right) => left.sortOrder - right.sortOrder);

    const criticalCount = findings.filter((finding) => finding.severity === "critical").length;
    const warningCount = findings.filter((finding) => finding.severity === "warning").length;
    const suggestionCount = findings.filter((finding) => finding.severity === "suggestion").length;

    return {
      findings,
      blockingCount: criticalCount,
      warningCount,
      suggestionCount,
      counts: {
        critical: criticalCount,
        warning: warningCount,
        suggestion: suggestionCount
      }
    };
  }

  function buildSectionCompletion(sectionKey) {
    const ids = getSectionRequirementIds(sectionKey);
    if (sectionKey === "equity" && !isEquitySelected()) return "Optional";
    const complete = ids.filter((id) => {
      const element = document.getElementById(id);
      if (!isFilled(element)) return false;
      if (id === "targetPrice" || id === "marketCapUsd") return parseNumber(element.value) != null;
      return true;
    }).length;
    return `${complete}/${ids.length}`;
  }

  function updateSectionPills(validation, review) {
    dom.navNote.textContent = buildSectionCompletion("note");
    dom.navAuthors.textContent = buildSectionCompletion("authors");
    dom.navEquity.textContent = buildSectionCompletion("equity");
    dom.navBody.textContent = buildSectionCompletion("body");

    const supportCount = Array.from(dom.modelFiles.files || []).length + (state.figureFiles || []).length;
    dom.navExhibits.textContent = supportCount ? `${supportCount} files` : "Optional";
    if (review.blockingCount > 0) {
      dom.navOutput.textContent = "Blocked";
    } else if (review.warningCount > 0) {
      dom.navOutput.textContent = "Check";
    } else if (validation.valid) {
      dom.navOutput.textContent = "Ready";
    } else {
      dom.navOutput.textContent = "Draft";
    }
  }

  function updateCompletion(validation, review) {
    dom.completionBar.style.width = `${validation.percent}%`;
    dom.completionText.textContent = `${validation.complete} / ${validation.total} required fields complete`;
    dom.readinessPercent.textContent = `${validation.percent}%`;
    if (review.blockingCount > 0) {
      dom.noteStateChip.textContent = "Review";
      dom.noteStateChip.style.background = "rgba(132, 95, 15, 0.10)";
      dom.noteStateChip.style.color = "#845F0F";
      dom.noteStateChip.style.borderColor = "rgba(132, 95, 15, 0.22)";
    } else if (review.warningCount > 0) {
      dom.noteStateChip.textContent = "Check";
      dom.noteStateChip.style.background = "rgba(180, 139, 51, 0.12)";
      dom.noteStateChip.style.color = "#8a6521";
      dom.noteStateChip.style.borderColor = "rgba(180, 139, 51, 0.22)";
    } else if (validation.valid) {
      dom.noteStateChip.textContent = "Ready";
      dom.noteStateChip.style.background = "rgba(37, 115, 75, 0.12)";
      dom.noteStateChip.style.color = "#25734b";
      dom.noteStateChip.style.borderColor = "rgba(37, 115, 75, 0.2)";
    } else if (validation.percent >= 55) {
      dom.noteStateChip.textContent = "In Build";
      dom.noteStateChip.style.background = "rgba(132, 95, 15, 0.10)";
      dom.noteStateChip.style.color = "#845F0F";
      dom.noteStateChip.style.borderColor = "rgba(132, 95, 15, 0.22)";
    } else {
      dom.noteStateChip.textContent = "Draft";
      dom.noteStateChip.style.background = "rgba(109, 118, 130, 0.12)";
      dom.noteStateChip.style.color = "#5d6570";
      dom.noteStateChip.style.borderColor = "rgba(109, 118, 130, 0.16)";
    }

    const progressTrack = dom.completionBar.parentElement;
    if (progressTrack) progressTrack.setAttribute("aria-valuenow", String(validation.percent));

    if (review.blockingCount > 0) {
      dom.readinessCaption.textContent = `Pre-publish review has ${review.blockingCount} critical issue${review.blockingCount === 1 ? "" : "s"} to resolve before export.`;
    } else if (review.warningCount > 0) {
      dom.readinessCaption.textContent = `The note is structurally complete, but the readiness pass still shows ${review.warningCount} warning${review.warningCount === 1 ? "" : "s"} to tighten before publication.`;
    } else if (validation.valid) {
      dom.readinessCaption.textContent = "The brief, narrative, and author metadata are complete. The Word note can now render cleanly in the publication template.";
    } else if (validation.percent >= 55) {
      dom.readinessCaption.textContent = "The draft is structurally taking shape. Complete the remaining mandatory fields before export.";
    } else {
      dom.readinessCaption.textContent = "Complete the brief, byline, and core body sections before generating the Word note.";
    }

    renderMissingFields(review);
  }

  function renderMissingFields(review) {
    dom.missingFields.innerHTML = "";

    if (!review.findings.length) {
      const item = document.createElement("li");
      item.className = "finding-item finding-clear";
      item.textContent = "No critical issues or review warnings. The note is structurally ready for publication export.";
      dom.missingFields.appendChild(item);
      return;
    }

    review.findings.slice(0, 6).forEach((entry) => {
      const item = document.createElement("li");
      item.className = `finding-item finding-${entry.severity}`;

      const row = document.createElement("div");
      row.className = "finding-row";

      const copy = document.createElement("div");
      const title = document.createElement("strong");
      title.className = "finding-title";
      title.textContent = entry.title;
      copy.appendChild(title);

      if (entry.detail) {
        const detail = document.createElement("span");
        detail.className = "finding-detail";
        detail.textContent = entry.detail;
        copy.appendChild(detail);
      }

      row.appendChild(copy);

      const badge = document.createElement("span");
      badge.className = "finding-badge";
      badge.textContent = entry.severity;
      row.appendChild(badge);

      item.appendChild(row);

      const meta = document.createElement("span");
      meta.textContent = entry.section;
      item.appendChild(meta);
      dom.missingFields.appendChild(item);
    });
  }

  function updateSummaryCards(validation, review, data) {
    const noteType = data.noteType;
    const deskLine = data.deskLine;
    const deck = data.deck;
    const authorLine = buildPrimaryAuthorLine();
    const coAuthors = data.coAuthors;

    dom.summaryType.textContent = strategyLabelForNoteType(noteType) || "Select a note type";
    dom.summaryTopic.textContent = deck || `${deskLine || "Set the desk line"}${dom.publicationDate.value ? ` | ${formatInputDateLabel(dom.publicationDate.value)}` : ""}`;
    dom.summaryAuthor.textContent = authorLine || "Assign primary author";
    dom.summaryCoAuthors.textContent = coAuthors.length
      ? `${coAuthors.length} co-author${coAuthors.length > 1 ? "s" : ""} added.`
      : "No co-authors added.";
    if (review.blockingCount > 0) {
      dom.summaryOutput.textContent = "Critical issues to resolve";
    } else if (review.warningCount > 0) {
      dom.summaryOutput.textContent = "Ready with warnings";
    } else if (validation.valid) {
      dom.summaryOutput.textContent = "Publication package ready";
    } else {
      dom.summaryOutput.textContent = "Draft in progress";
    }

    const attachmentBits = [];
    const modelCount = Array.from(dom.modelFiles.files || []).length;
    const imageCount = (state.figureFiles || []).length;
    if (isEquitySelected()) attachmentBits.push(state.priceChartImageBytes ? "price chart ready" : "price chart pending");
    if (modelCount) attachmentBits.push(`${modelCount} model file${modelCount > 1 ? "s" : ""}`);
    if (imageCount) attachmentBits.push(`${imageCount} exhibit${imageCount > 1 ? "s" : ""}`);

    const reviewSummary = `Pre-publish review: ${review.counts.critical} critical, ${review.counts.warning} warning${review.counts.warning === 1 ? "" : "s"}, ${review.counts.suggestion} suggestion${review.counts.suggestion === 1 ? "" : "s"}.`;

    dom.summaryOutputDetail.textContent = attachmentBits.length
      ? `${reviewSummary} Support pack includes ${attachmentBits.join(", ")}.`
      : reviewSummary;
  }

  function updatePreview(data) {
    dom.previewTitle.textContent = data.title || "No title yet";
    dom.previewAuthor.textContent = buildPrimaryAuthorLine() || "Primary author pending";

    if (isEquitySelected()) {
      const coverageBits = [
        data.equityCompanyName || data.ticker || "Security pending",
        data.deskLine || "Desk line pending",
        data.publicationDate ? formatInputDateLabel(data.publicationDate) : "Date pending",
        data.crgRating || "Rating pending"
      ];
      dom.previewCoverage.textContent = coverageBits.join(" | ");
    } else if (isMacroFiSelected()) {
      const coverageBits = [
        data.deskLine || data.noteType || "Desk line not set",
        data.coverageCountry ? getCoverageCountryLabel(data.coverageCountry) : (data.noteType === "Fixed Income Research" ? "Country optional" : "Country pending"),
        data.issuerId || (data.noteType === "Fixed Income Research" ? "Issuer optional" : "Issuer pending"),
        data.publicationDate ? formatInputDateLabel(data.publicationDate) : "Date pending"
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
    const imageCount = (state.figureFiles || []).length;
    if (state.priceChartImageBytes) supportBits.push("chart");
    if (modelCount) supportBits.push(`${modelCount} model file${modelCount > 1 ? "s" : ""}`);
    if (imageCount) supportBits.push(`${imageCount} image${imageCount > 1 ? "s" : ""}`);
    dom.previewSupport.textContent = supportBits.length ? supportBits.join(" | ") : "No attachments yet";
  }

  function buildPrimaryAuthorLine() {
    const lastName = dom.authorLastName.value.trim();
    const firstName = dom.authorFirstName.value.trim();
    const combined = [firstName, lastName].filter(Boolean).join(" ").trim();
    if (!combined) return "";
    return `${combined}, ${formatPhoneDisplay(dom.authorPhone.value)}`;
  }

  function getDeskLine() {
    return dom.deskLine.value.trim() || defaultDeskLine(dom.noteType.value);
  }

  function updateAllUI() {
    const validation = validateForm(false);
    const data = collectFormData();
    const review = buildPrePublishReview(data, validation);
    updateCompletion(validation, review);
    updateSectionPills(validation, review);
    updateSummaryCards(validation, review, data);
    updatePreview(data);
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
      bodySectionLayout: collectBodySectionLayout(),
      ratingsProfile: collectRatingsProfile(),
      savedAt: new Date().toISOString()
    };
  }

  function saveDraft() {
    try {
      const payload = serializeDraft();
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
      state.lastSavedAt = payload.savedAt;
    } catch (error) {
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
      restoreBodySectionLayout(payload.bodySectionLayout || []);
      restoreRatingsProfile(payload.ratingsProfile || []);
      state.lastSavedAt = payload.savedAt || null;
      dom.deskLine.dataset.autofill = "true";
      dom.equityCompanyName.dataset.autofill = dom.equityCompanyName.value.trim() ? "false" : "true";
      dom.priceCurrency.dataset.autofill = dom.priceCurrency.value.trim() ? "false" : "true";
      syncIssuerId();
      updateCoverageCountryDisplayFromCode();
      renderIndustryOptions("");
      syncBodySectionVisibility();
      ensureDeskLineDefault(true);
    } catch (error) {
      console.error("Draft restore failed:", error);
    }
  }

  function resetDraft() {
    const confirmed = window.confirm("Reset the entire draft? Text fields, metadata, and autosaved content will be cleared.");
    if (!confirmed) return;

    form.reset();
    dom.coAuthorsList.innerHTML = "";
    state.coAuthorCount = 0;
    state.lastSavedAt = null;
    state.customSectionCount = 0;
    state.figurePlacements = {};
    state.figureDetails = {};
    state.figureFiles = [];
    syncPrimaryPhone();
    restoreRatingsProfile([]);
    dom.equityCompanyName.dataset.autofill = "true";
    dom.priceCurrency.dataset.autofill = "true";
    resetChartState({ keepStatusText: false });
    window.localStorage.removeItem(STORAGE_KEY);
    ensurePublicationDate();
    initializeFinancialTableEditor(true);
    ensureDeskLineDefault(true);
    updateCoverageCountryDisplayFromCode();
    renderIndustryOptions("");
    restoreBodySectionLayout([]);
    updateFileSummary(dom.modelFiles, dom.modelSummaryHead, dom.modelSummaryList, "No supporting files attached.");
    updateFigureSummary();
    syncFigurePlacementControls();
    toggleNoteTypeSections();
    closePreviewModal();
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

  function updateFigureSummary() {
    if (!dom.imageSummaryHead || !dom.imageSummaryList) return;
    const files = state.figureFiles || [];
    dom.imageSummaryHead.textContent = files.length
      ? `${files.length} figure${files.length > 1 ? "s" : ""} attached.`
      : "No figures attached.";

    dom.imageSummaryList.innerHTML = "";
    if (!files.length) {
      dom.imageSummaryList.hidden = true;
      return;
    }

    files.forEach((file, index) => {
      const item = document.createElement("li");
      const meta = getFigureDetailForFile(file, index);
      const label = `${meta.labelType} ${meta.labelNumber}`;
      item.textContent = `${label} - ${meta.caption || file.name.replace(/\.[^.]+$/, "")}`;
      dom.imageSummaryList.appendChild(item);
    });

    dom.imageSummaryList.hidden = false;
  }

  function figureFileKey(file) {
    return `${file.name}__${file.size}__${file.lastModified}`;
  }

  function defaultFigureLabelType(file) {
    return /chart/i.test(file.name) ? "Chart" : "Figure";
  }

  function defaultFigureCaption(file) {
    return file.name.replace(/\.[^.]+$/, "");
  }

  function sanitizeFigureNumber(value, fallback) {
    const parsed = Number.parseInt(String(value || "").trim(), 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
  }

  function getFigureDetailForFile(file, index = 0) {
    const key = figureFileKey(file);
    const existing = state.figureDetails[key] || {};
    return {
      placement: state.figurePlacements[key] || existing.placement || "end",
      labelType: existing.labelType || defaultFigureLabelType(file),
      labelNumber: sanitizeFigureNumber(existing.labelNumber, index + 1),
      caption: String(existing.caption || "").trim() || defaultFigureCaption(file)
    };
  }

  function setManagedFigureFiles(files) {
    state.figureFiles = files;
    const nextDetails = {};

    state.figureFiles.forEach((file, index) => {
      const key = figureFileKey(file);
      nextDetails[key] = getFigureDetailForFile(file, index);
    });

    state.figureDetails = nextDetails;
    state.figurePlacements = Object.fromEntries(
      Object.entries(nextDetails).map(([key, detail]) => [key, detail.placement || "end"])
    );
    updateFigureSummary();
  }

  function handleFigureUploadChange() {
    const incoming = Array.from(dom.imageUpload.files || []);
    if (!incoming.length) return;

    const nextFiles = [...state.figureFiles];
    const seen = new Set(nextFiles.map((file) => figureFileKey(file)));
    incoming.forEach((file) => {
      const key = figureFileKey(file);
      if (!seen.has(key)) {
        nextFiles.push(file);
        seen.add(key);
      }
    });

    setManagedFigureFiles(nextFiles);
    dom.imageUpload.value = "";
  }

  function removeFigureByKey(figureKey) {
    setManagedFigureFiles((state.figureFiles || []).filter((file) => figureFileKey(file) !== figureKey));
    syncFigurePlacementControls();
    updateAllUI();
    queueDraftSave();
  }

  function getFigurePlacementOptions(noteType = dom.noteType.value) {
    const options = [];
    const layout = collectBodySectionLayout().filter((entry) => !entry.hidden);
    const keyTakeaways = layout.find((entry) => entry.key === "keyTakeaways");
    const cordobaView = layout.find((entry) => entry.key === "cordobaView");
    const middle = layout.filter((entry) => entry.key !== "keyTakeaways" && entry.key !== "cordobaView");

    if (noteType === "Equity Research") {
      options.push(
        { value: "after-businessDescription", label: "After Business Description" },
        { value: "after-valuationSummary", label: "After Valuation Summary" }
      );

      middle.forEach((entry) => {
        options.push({ value: `after-${entry.key}`, label: `After ${entry.label}` });
      });
      if (cordobaView) options.push({ value: "after-cordobaView", label: `After ${cordobaView.label}` });

      options.push({ value: "end", label: "End of Note" });
      return options;
    }

    if (isMacroFiNoteType(noteType)) {
      const macroHeading = String(dom.macroFiHeading?.value || "").trim() || "Ratings Profile";
      options.push({ value: "after-macroFiProfile", label: `After ${macroHeading}` });
    }

    if (keyTakeaways) options.push({ value: "after-keyTakeaways", label: `After ${keyTakeaways.label}` });
    middle.forEach((entry) => {
      options.push({ value: `after-${entry.key}`, label: `After ${entry.label}` });
    });
    if (cordobaView) options.push({ value: "after-cordobaView", label: `After ${cordobaView.label}` });

    options.push({ value: "end", label: "End of Note" });
    return options;
  }

  function syncFigurePlacementControls() {
    if (!dom.figurePlacementPanel || !dom.figurePlacementList) return;

    const files = state.figureFiles || [];
    if (!files.length) {
      dom.figurePlacementPanel.hidden = true;
      dom.figurePlacementList.innerHTML = "";
      state.figurePlacements = {};
      state.figureDetails = {};
      return;
    }

    const options = getFigurePlacementOptions(dom.noteType.value);
    const allowedValues = new Set(options.map((option) => option.value));
    const nextPlacements = {};

    dom.figurePlacementList.innerHTML = "";
    files.forEach((file, index) => {
      const key = figureFileKey(file);
      const detail = getFigureDetailForFile(file, index);
      const currentValue = allowedValues.has(detail.placement) ? detail.placement : "end";
      nextPlacements[key] = currentValue;

      const row = document.createElement("div");
      row.className = "figure-placement-row";
      row.setAttribute("data-figure-row", key);

      const nameWrap = document.createElement("div");
      const name = document.createElement("div");
      name.className = "figure-placement-name";
      name.textContent = `${detail.labelType} ${detail.labelNumber}: ${detail.caption}`;
      nameWrap.appendChild(name);

      const origin = document.createElement("div");
      origin.className = "figure-placement-origin";
      origin.textContent = file.name;
      nameWrap.appendChild(origin);
      row.appendChild(nameWrap);

      const meta = document.createElement("div");
      meta.className = "figure-placement-meta";

      const metaGrid = document.createElement("div");
      metaGrid.className = "figure-placement-meta-grid";

      const kindSelect = document.createElement("select");
      kindSelect.setAttribute("data-figure-key", key);
      kindSelect.setAttribute("data-figure-field", "labelType");
      ["Figure", "Chart", "Exhibit"].forEach((labelType) => {
        const optionEl = document.createElement("option");
        optionEl.value = labelType;
        optionEl.textContent = labelType;
        if (labelType === detail.labelType) optionEl.selected = true;
        kindSelect.appendChild(optionEl);
      });
      metaGrid.appendChild(kindSelect);

      const numberInput = document.createElement("input");
      numberInput.type = "text";
      numberInput.inputMode = "numeric";
      numberInput.value = String(detail.labelNumber);
      numberInput.setAttribute("data-figure-key", key);
      numberInput.setAttribute("data-figure-field", "labelNumber");
      numberInput.setAttribute("aria-label", "Figure number");
      metaGrid.appendChild(numberInput);

      const select = document.createElement("select");
      select.setAttribute("data-figure-key", key);
      select.setAttribute("data-figure-field", "placement");
      options.forEach((option) => {
        const optionEl = document.createElement("option");
        optionEl.value = option.value;
        optionEl.textContent = option.label;
        if (option.value === currentValue) optionEl.selected = true;
        select.appendChild(optionEl);
      });
      metaGrid.appendChild(select);
      meta.appendChild(metaGrid);

      const captionInput = document.createElement("input");
      captionInput.type = "text";
      captionInput.value = detail.caption;
      captionInput.placeholder = "Figure caption";
      captionInput.setAttribute("data-figure-key", key);
      captionInput.setAttribute("data-figure-field", "caption");
      meta.appendChild(captionInput);

      const actionRow = document.createElement("div");
      actionRow.className = "figure-placement-actions";
      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "btn btn-ghost btn-xs";
      removeBtn.setAttribute("data-figure-action", "remove");
      removeBtn.setAttribute("data-figure-key", key);
      removeBtn.textContent = "Remove";
      actionRow.appendChild(removeBtn);
      meta.appendChild(actionRow);

      row.appendChild(meta);
      dom.figurePlacementList.appendChild(row);
    });

    state.figurePlacements = nextPlacements;
    dom.figurePlacementPanel.hidden = false;
  }

  function handleFigurePlacementChange(event) {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement || target instanceof HTMLInputElement)) return;
    const figureKey = target.getAttribute("data-figure-key");
    if (!figureKey) return;
    const field = target.getAttribute("data-figure-field");
    if (!field) return;

    const current = state.figureDetails[figureKey] || {};
    if (field === "labelNumber") {
      current.labelNumber = sanitizeFigureNumber(target.value, current.labelNumber || 1);
      target.value = String(current.labelNumber);
    } else {
      current[field] = target.value;
    }
    if (field === "placement") state.figurePlacements[figureKey] = current.placement || "end";
    state.figureDetails[figureKey] = current;
    const nameEl = target.closest(".figure-placement-row")?.querySelector(".figure-placement-name");
    if (nameEl) {
      const labelType = current.labelType || "Figure";
      const labelNumber = sanitizeFigureNumber(current.labelNumber, 1);
      const caption = String(current.caption || "").trim() || "Untitled figure";
      nameEl.textContent = `${labelType} ${labelNumber}: ${caption}`;
    }
    updateFigureSummary();
    updateAllUI();
  }

  function handleFigurePlacementActions(event) {
    const actionButton = event.target.closest("[data-figure-action='remove']");
    if (!actionButton) return;
    removeFigureByKey(actionButton.getAttribute("data-figure-key"));
  }

  function buildCurrentFigurePlacements(files) {
    return files.reduce((acc, file) => {
      acc[figureFileKey(file)] = getFigureDetailForFile(file).placement || state.figurePlacements[figureFileKey(file)] || "end";
      return acc;
    }, {});
  }

  function buildCurrentFigureDetails(files) {
    return files.reduce((acc, file, index) => {
      acc[figureFileKey(file)] = getFigureDetailForFile(file, index);
      return acc;
    }, {});
  }

  function resolveFigurePlacementForFile(file, data, availablePlacements) {
    const requested = data.figurePlacements?.[figureFileKey(file)] || "end";
    return availablePlacements.has(requested) ? requested : "end";
  }

  function getFigureFilesForPlacement(data, placement, availablePlacements) {
    return (data.imageFiles || []).filter((file) => resolveFigurePlacementForFile(file, data, availablePlacements) === placement);
  }

  async function appendPlacedFigures(children, docxLib, colors, files, figureCounterRef, options = {}) {
    if (!files.length) return;
    const startIndex = figureCounterRef.value;
    const imageParagraphs = await buildImageParagraphs(docxLib, files, colors, startIndex, options.figureDetails || {});
    figureCounterRef.value += files.length;

    if (options.withHeading) {
      children.push(buildNomuraSubhead(docxLib, colors, options.heading || "Figures / Screenshots"));
    }

    children.push(...imageParagraphs);
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
      rangeReturn: null,
      priceDate: null,
      providerCurrency: "",
      benchmarkLabel: "",
      chartMode: "price",
      symbol: "",
      benchmarkSymbol: "",
      range: ""
    };

    setMetric(dom.currentPrice, "-");
    setMetric(dom.priceDate, "-");
    setMetric(dom.rangeReturn, "-");
    setMetric(dom.upsideToTarget, "-");

    if (!options.keepStatusText) dom.chartStatus.textContent = isEquitySelected() ? "No tear sheet market data fetched yet." : "";
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
      setMessage("error", "Chart.js is unavailable, so the tear sheet chart cannot be rendered in this session.");
      return;
    }

    dom.fetchPriceChart.disabled = true;
    dom.fetchPriceChart.classList.add("loading");
    clearMessage();

    try {
      await loadEquityMarketSnapshot();
    } catch (error) {
      resetChartState({ keepStatusText: true });
      dom.chartStatus.textContent = error.message;
      setMessage("error", `Unable to build the tear sheet: ${error.message}`);
    } finally {
      dom.fetchPriceChart.disabled = false;
      dom.fetchPriceChart.classList.remove("loading");
    }
  }

  async function loadEquityMarketSnapshot() {
    if (typeof window.Chart === "undefined") {
      throw new Error("Chart.js is unavailable, so the equity tear sheet cannot be rendered in this session.");
    }

    const ticker = dom.ticker.value.trim();
    if (!ticker) {
      dom.ticker.classList.add("is-invalid");
      dom.chartStatus.textContent = "Enter a ticker before refreshing the tear sheet.";
      dom.ticker.focus();
      throw new Error("Enter a ticker before refreshing the tear sheet.");
    }
    dom.ticker.classList.remove("is-invalid");

    const range = dom.chartRange.value || "6mo";
    const symbol = marketDataSymbolFromTicker(ticker);
    const benchmarkInput = dom.benchmarkTicker.value.trim();
    const benchmarkSymbol = benchmarkInput ? marketDataSymbolFromTicker(benchmarkInput) : "";
    const benchmarkLabel = dom.benchmarkName.value.trim() || benchmarkSymbol || benchmarkInput.toUpperCase();

    dom.chartStatus.textContent = benchmarkSymbol
      ? "Fetching security and benchmark history for the tear sheet..."
      : "Fetching security history for the tear sheet...";

    const securityResult = await fetchMarketHistory(symbol, range);
    const filteredSecurity = securityResult.rows;
    if (filteredSecurity.length < 10) {
      throw new Error("Not enough price history returned for the selected range.");
    }

    const resolvedProfile = await resolveMarketSecurityProfile(symbol, securityResult.meta);
    applyResolvedSecurityProfile(resolvedProfile);

    let filteredBenchmark = null;
    let benchmarkNotice = "";

    if (benchmarkSymbol) {
      try {
        const benchmarkResult = await fetchMarketHistory(benchmarkSymbol, range);
        const candidateSeries = benchmarkResult.rows;
        if (candidateSeries.length >= 10) {
          filteredBenchmark = candidateSeries;
        } else {
          benchmarkNotice = "Benchmark history was too short, so the chart shows security performance only.";
        }
      } catch (error) {
        benchmarkNotice = "Benchmark history was unavailable, so the chart shows security performance only.";
      }
    }

    const chartConfig = buildEquityChartConfig({
      securitySeries: filteredSecurity,
      benchmarkSeries: filteredBenchmark,
      securityLabel: ticker.toUpperCase(),
      benchmarkLabel
    });

    renderChart(chartConfig);
    await waitForChartPaint();

    state.priceChartImageBytes = await buildExportChartImageBytes(chartConfig);

    const closeValues = filteredSecurity.map((item) => item.close);
    const currentPrice = closeValues[closeValues.length - 1];
    const rangeReturn = closeValues[0] ? (currentPrice / closeValues[0]) - 1 : null;
    const returns = computeDailyReturns(closeValues);
    const dailyVol = standardDeviation(returns);
    const realisedVolAnn = dailyVol == null ? null : dailyVol * Math.sqrt(252);
    const priceDate = filteredSecurity[filteredSecurity.length - 1]?.date || null;

    state.equityStats = {
      currentPrice,
      rangeReturn,
      realisedVolAnn,
      priceDate,
      providerCurrency: normalizeMarketCurrency(resolvedProfile.currency || securityResult.meta.currency),
      benchmarkLabel: chartConfig.mode === "relative" ? benchmarkLabel : "",
      chartMode: chartConfig.mode,
      symbol,
      benchmarkSymbol,
      range
    };

    setMetric(dom.currentPrice, currentPrice != null ? formatPriceDisplay(currentPrice, resolvePriceCurrency({
      priceCurrency: dom.priceCurrency.value.trim(),
      equityStats: state.equityStats
    })) : "-");
    setMetric(dom.priceDate, priceDate ? formatShortDisplayDate(priceDate) : "-");
    setMetric(dom.rangeReturn, rangeReturn != null ? formatPercent(rangeReturn) : "-");
    updateUpsideDisplay();

    const rangeLabel = range.toUpperCase();
    if (chartConfig.mode === "relative" && benchmarkLabel) {
      dom.chartStatus.textContent = `Tear sheet ready for ${ticker.toUpperCase()} versus ${benchmarkLabel} (${rangeLabel}).`;
    } else if (benchmarkNotice) {
      dom.chartStatus.textContent = `${benchmarkNotice} Tear sheet updated for ${ticker.toUpperCase()} (${rangeLabel}).`;
    } else {
      dom.chartStatus.textContent = `Tear sheet updated for ${ticker.toUpperCase()} (${rangeLabel}).`;
    }

    updateAllUI();
  }

  function marketDataSymbolFromTicker(rawTicker) {
    const cleaned = String(rawTicker || "")
      .trim()
      .split(/\s+/)[0]
      .replace(/^[A-Z]+:/i, "")
      .replace(/,+$/, "");

    if (!cleaned) return "";
    if (cleaned.startsWith("^")) return cleaned.toUpperCase();

    const dotIndex = cleaned.lastIndexOf(".");
    if (dotIndex === -1) return cleaned.toUpperCase();

    const base = cleaned.slice(0, dotIndex).toUpperCase();
    const suffix = cleaned.slice(dotIndex + 1).toUpperCase();
    const suffixMap = {
      US: "",
      UK: ".L",
      LN: ".L",
      L: ".L",
      TW: ".TW",
      TT: ".TW",
      JP: ".T",
      T: ".T",
      HK: ".HK",
      AU: ".AX",
      AX: ".AX",
      CA: ".TO",
      TO: ".TO",
      FR: ".PA",
      PA: ".PA",
      DE: ".DE",
      GR: ".DE",
      SW: ".SW",
      CH: ".SW",
      NL: ".AS",
      AS: ".AS",
      ES: ".MC",
      MC: ".MC",
      IT: ".MI",
      MI: ".MI",
      SE: ".ST",
      ST: ".ST",
      DK: ".CO",
      CO: ".CO",
      FI: ".HE",
      HE: ".HE",
      BR: ".BR",
      BE: ".BR",
      IR: ".IR",
      IE: ".IR",
      WA: ".WA",
      PL: ".WA"
    };

    return `${base}${suffixMap[suffix] ?? `.${suffix}`}`;
  }

  async function fetchMarketHistory(symbol, range) {
    const normalizedRange = ["6mo", "1y", "2y", "5y"].includes(range) ? range : "6mo";
    const requestPaths = [
      `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?interval=1d&range=${encodeURIComponent(normalizedRange)}&includeAdjustedClose=true`,
      `https://query2.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?interval=1d&range=${encodeURIComponent(normalizedRange)}&includeAdjustedClose=true`,
      `https://r.jina.ai/http://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?interval=1d&range=${encodeURIComponent(normalizedRange)}&includeAdjustedClose=true`,
      `https://r.jina.ai/http://query2.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(symbol)}?interval=1d&range=${encodeURIComponent(normalizedRange)}&includeAdjustedClose=true`
    ];

    let lastError = new Error(`Market data request failed for ${symbol.toUpperCase()}.`);

    for (const url of requestPaths) {
      try {
        const response = await fetchWithTimeout(url, {
          cache: "no-store",
          headers: {
            Accept: "application/json, text/plain;q=0.9, */*;q=0.8"
          }
        }, 14000);

        if (!response.ok) throw new Error(`Market data request returned ${response.status}.`);

        const rawText = await response.text();
        const result = parseYahooChartResponse(rawText, symbol);
        if (result.rows.length < 5) throw new Error("The market data feed did not return enough observations.");
        return result;
      } catch (error) {
        lastError = error;
      }
    }

    throw lastError;
  }

  async function resolveMarketSecurityProfile(symbol, chartMeta = {}) {
    const baseProfile = buildMarketSecurityProfile(chartMeta, symbol);
    const needsQuoteLookup = !baseProfile.companyName || !baseProfile.currency || !baseProfile.exchangeName;

    if (!needsQuoteLookup) return baseProfile;

    try {
      const quoteProfile = await fetchMarketQuoteProfile(symbol);
      return {
        companyName: quoteProfile.companyName || baseProfile.companyName,
        exchangeName: quoteProfile.exchangeName || baseProfile.exchangeName,
        currency: quoteProfile.currency || baseProfile.currency
      };
    } catch (error) {
      console.warn(`Unable to resolve quote profile for ${symbol}:`, error);
      return baseProfile;
    }
  }

  function buildMarketSecurityProfile(meta = {}, fallbackSymbol = "") {
    return {
      companyName: String(meta.longName || meta.shortName || meta.displayName || "").trim(),
      exchangeName: String(meta.fullExchangeName || meta.exchangeName || meta.exchange || "").trim(),
      currency: normalizeMarketCurrency(meta.currency || ""),
      symbol: String(meta.symbol || fallbackSymbol || "").trim()
    };
  }

  async function fetchMarketQuoteProfile(symbol) {
    const requestPaths = [
      `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${encodeURIComponent(symbol)}`,
      `https://query2.finance.yahoo.com/v7/finance/quote?symbols=${encodeURIComponent(symbol)}`,
      `https://r.jina.ai/http://query1.finance.yahoo.com/v7/finance/quote?symbols=${encodeURIComponent(symbol)}`,
      `https://r.jina.ai/http://query2.finance.yahoo.com/v7/finance/quote?symbols=${encodeURIComponent(symbol)}`
    ];

    let lastError = new Error(`Quote profile request failed for ${symbol.toUpperCase()}.`);

    for (const url of requestPaths) {
      try {
        const response = await fetchWithTimeout(url, {
          cache: "no-store",
          headers: {
            Accept: "application/json, text/plain;q=0.9, */*;q=0.8"
          }
        }, 12000);

        if (!response.ok) throw new Error(`Quote profile request returned ${response.status}.`);

        const rawText = await response.text();
        return parseYahooQuoteResponse(rawText, symbol);
      } catch (error) {
        lastError = error;
      }
    }

    throw lastError;
  }

  function parseYahooQuoteResponse(rawText, symbol) {
    const payloadText = extractJsonObject(rawText);
    if (!payloadText) {
      throw new Error("The quote profile feed returned an unexpected response.");
    }

    let payload = null;
    try {
      payload = JSON.parse(payloadText);
    } catch (error) {
      throw new Error("The quote profile feed returned malformed JSON.");
    }

    const quote = payload?.quoteResponse?.result?.[0];
    if (!quote) {
      throw new Error(`No quote profile was found for ${symbol.toUpperCase()}.`);
    }

    return {
      companyName: String(quote.longName || quote.shortName || quote.displayName || "").trim(),
      exchangeName: String(quote.fullExchangeName || quote.exchange || quote.exchangeName || "").trim(),
      currency: normalizeMarketCurrency(quote.currency || "")
    };
  }

  function parseYahooChartResponse(rawText, symbol) {
    const payloadText = extractJsonObject(rawText);
    if (!payloadText) {
      throw new Error("The market data feed returned an unexpected response.");
    }

    let payload = null;
    try {
      payload = JSON.parse(payloadText);
    } catch (error) {
      throw new Error("The market data feed returned malformed JSON.");
    }

    const chart = payload?.chart;
    if (!chart) {
      throw new Error("The market data feed returned an unexpected schema.");
    }

    if (chart.error?.description) {
      throw new Error(chart.error.description);
    }

    const result = Array.isArray(chart.result) ? chart.result[0] : null;
    if (!result) {
      throw new Error(`No market data was found for ${symbol.toUpperCase()}.`);
    }

    const timestamps = Array.isArray(result.timestamp) ? result.timestamp : [];
    const quoteClose = result?.indicators?.quote?.[0]?.close;
    const adjustedClose = result?.indicators?.adjclose?.[0]?.adjclose;
    const closeValues = Array.isArray(adjustedClose) && adjustedClose.length ? adjustedClose : (Array.isArray(quoteClose) ? quoteClose : []);

    if (!timestamps.length || !closeValues.length) {
      throw new Error("The market data feed returned no usable price history.");
    }

    const rows = timestamps
      .map((timestamp, index) => ({
        date: formatUnixDate(timestamp),
        close: Number(closeValues[index])
      }))
      .filter((row) => /^\d{4}-\d{2}-\d{2}$/.test(row.date) && Number.isFinite(row.close) && row.close > 0);

    if (!rows.length) {
      throw new Error("The market data feed returned no usable observations.");
    }

    return {
      rows,
      meta: result.meta || {}
    };
  }

  function extractJsonObject(rawText) {
    const source = String(rawText || "").trim();
    if (!source) return "";

    const start = source.indexOf("{");
    if (start === -1) return "";

    let depth = 0;
    let inString = false;
    let escaped = false;

    for (let index = start; index < source.length; index += 1) {
      const char = source[index];

      if (escaped) {
        escaped = false;
        continue;
      }

      if (char === "\\") {
        escaped = true;
        continue;
      }

      if (char === "\"") {
        inString = !inString;
        continue;
      }

      if (inString) continue;

      if (char === "{") depth += 1;
      if (char === "}") depth -= 1;

      if (depth === 0) {
        return source.slice(start, index + 1);
      }
    }

    return source.slice(start);
  }

  function formatUnixDate(timestamp) {
    const date = new Date(Number(timestamp) * 1000);
    if (Number.isNaN(date.getTime())) return "";

    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    const day = String(date.getUTCDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function normalizeMarketCurrency(value) {
    const normalized = String(value || "").trim();
    if (!normalized) return "";

    const currencyMap = {
      USD: "USD",
      EUR: "EUR",
      GBP: "GBP",
      GBX: "GBp",
      GBPX: "GBp",
      GBp: "GBp",
      JPY: "JPY",
      TWD: "TWD",
      HKD: "HKD",
      CAD: "CAD",
      AUD: "AUD",
      CHF: "CHF"
    };

    return currencyMap[normalized] || normalized;
  }

  function canAutoFillField(input) {
    if (!input) return false;
    const currentValue = input.value.trim();
    if (!currentValue) return true;
    return input.dataset.autofill === "true";
  }

  function setAutoFilledField(input, value) {
    const resolvedValue = String(value || "").trim();
    if (!input || !resolvedValue || !canAutoFillField(input)) return false;

    const changed = input.value !== resolvedValue;
    input.value = resolvedValue;
    input.dataset.autofill = "true";
    return changed;
  }

  function applyResolvedSecurityProfile(profile = {}) {
    let changed = false;

    if (profile.companyName) {
      changed = setAutoFilledField(dom.equityCompanyName, profile.companyName) || changed;
    }

    if (profile.currency) {
      changed = setAutoFilledField(dom.priceCurrency, profile.currency) || changed;
    }

    if (changed) {
      updateAllUI();
      queueDraftSave();
    }
  }

  function resolvePriceCurrency(data) {
    return String(data.priceCurrency || data.equityStats?.providerCurrency || "").trim();
  }

  function buildEquityChartConfig({ securitySeries, benchmarkSeries, securityLabel, benchmarkLabel }) {
    if (benchmarkSeries && benchmarkSeries.length) {
      const aligned = alignSeriesByDate(securitySeries, benchmarkSeries);
      if (aligned.length >= 10) {
        const securityBase = aligned[0].security;
        const benchmarkBase = aligned[0].benchmark;

        return {
          mode: "relative",
          labels: aligned.map((item) => item.date),
          datasets: [
            {
              label: securityLabel || "Security",
              data: aligned.map((item) => (item.security / securityBase) * 100),
              borderColor: "#9b4a2f",
              backgroundColor: "rgba(155, 74, 47, 0.08)",
              fill: false
            },
            {
              label: benchmarkLabel || "Benchmark",
              data: aligned.map((item) => (item.benchmark / benchmarkBase) * 100),
              borderColor: "#6a7482",
              backgroundColor: "rgba(106, 116, 130, 0.08)",
              fill: false
            }
          ]
        };
      }
    }

    return {
      mode: "price",
      labels: securitySeries.map((item) => item.date),
      datasets: [
        {
          label: securityLabel || "Security",
          data: securitySeries.map((item) => item.close),
          borderColor: "#9b7a36",
          backgroundColor: "rgba(155, 122, 54, 0.14)",
          fill: true
        }
      ]
    };
  }

  function alignSeriesByDate(securitySeries, benchmarkSeries) {
    const benchmarkMap = new Map(benchmarkSeries.map((item) => [item.date, item.close]));
    return securitySeries
      .filter((item) => benchmarkMap.has(item.date))
      .map((item) => ({
        date: item.date,
        security: item.close,
        benchmark: benchmarkMap.get(item.date)
      }))
      .filter((item) => Number.isFinite(item.security) && Number.isFinite(item.benchmark));
  }

  function fetchWithTimeout(resource, options, timeoutMs) {
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), timeoutMs);

    return fetch(resource, {
      ...options,
      signal: controller.signal
    }).finally(() => window.clearTimeout(timeout));
  }

  function renderChart(config) {
    if (state.priceChart) state.priceChart.destroy();

    const datasets = buildChartDatasets(config, "ui");

    state.priceChart = new window.Chart(dom.priceChartCanvas, {
      type: "line",
      data: {
        labels: config.labels,
        datasets
      },
      options: buildChartOptions(config, datasets.length, "ui")
    });
  }

  function buildChartDatasets(config, variant = "ui") {
    const isExport = variant === "export";

    return (config.datasets || []).map((dataset) => ({
      label: dataset.label,
      data: dataset.data,
      borderColor: dataset.borderColor,
      backgroundColor: dataset.backgroundColor,
      pointRadius: 0,
      borderWidth: isExport ? 4 : 2.2,
      tension: 0.2,
      fill: dataset.fill === true
    }));
  }

  function buildChartOptions(config, datasetCount, variant = "ui") {
    const isExport = variant === "export";
    const tickFontSize = isExport ? 17 : 11;
    const legendFontSize = isExport ? 19 : 11;

    return {
      responsive: !isExport,
      maintainAspectRatio: false,
      animation: isExport ? false : undefined,
      interaction: {
        mode: "index",
        intersect: false
      },
      layout: {
        padding: isExport ? { top: 8, right: 12, bottom: 8, left: 8 } : 0
      },
      plugins: {
        legend: {
          display: datasetCount > 1,
          position: "top",
          labels: {
            boxWidth: isExport ? 18 : 10,
            boxHeight: isExport ? 18 : 10,
            padding: isExport ? 18 : 12,
            color: "#4e617d",
            font: { family: "Aptos", size: legendFontSize, weight: "600" }
          }
        },
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
            color: "#4e617d",
            maxRotation: 0,
            minRotation: 0,
            padding: isExport ? 10 : 4,
            font: { family: "Aptos", size: tickFontSize }
          }
        },
        y: {
          grid: { color: "rgba(15, 23, 42, 0.08)" },
          ticks: {
            maxTicksLimit: 6,
            color: "#4e617d",
            padding: isExport ? 10 : 4,
            font: { family: "Aptos", size: tickFontSize },
            callback(value) {
              if (config.mode === "relative") return Number(value).toFixed(0);
              return value;
            }
          }
        }
      }
    };
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

  async function buildExportChartImageBytes(config) {
    const exportCanvas = document.createElement("canvas");
    exportCanvas.width = 1560;
    exportCanvas.height = 990;

    const exportDatasets = buildChartDatasets(config, "export");
    const exportChart = new window.Chart(exportCanvas, {
      type: "line",
      data: {
        labels: config.labels,
        datasets: exportDatasets
      },
      options: buildChartOptions(config, exportDatasets.length, "export")
    });

    try {
      exportChart.update("none");
      await waitForChartPaint();
      paintCanvasBackground(exportCanvas, "#FFFFFF");
      return canvasToPngBytes(exportCanvas);
    } finally {
      exportChart.destroy();
    }
  }

  function paintCanvasBackground(canvas, color) {
    const context = canvas.getContext("2d");
    if (!context) return;

    context.save();
    context.globalCompositeOperation = "destination-over";
    context.fillStyle = color;
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.restore();
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

  function formatSignedPercent(value) {
    const pct = formatPercent(value);
    return value > 0 ? `+${pct}` : pct;
  }

  function formatPriceDisplay(value, currency) {
    if (!Number.isFinite(value)) return "N/A";
    const prefix = String(currency || "").trim();
    const number = value.toFixed(2);
    return prefix ? `${prefix} ${number}` : number;
  }

  function formatShortDisplayDate(value) {
    const date = value instanceof Date ? value : new Date(`${value}T00:00:00`);
    if (Number.isNaN(date.getTime())) return "N/A";
    return new Intl.DateTimeFormat("en-GB", {
      day: "numeric",
      month: "short",
      year: "numeric"
    }).format(date);
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
      data.equityCompanyName ? `Company: ${data.equityCompanyName}` : null,
      data.ticker ? `Ticker / company: ${data.ticker}` : null,
      data.benchmarkName ? `Benchmark: ${data.benchmarkName}` : null,
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
    closePreviewModal();
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
      if (isEquitySelected()) {
        const activeSymbol = marketDataSymbolFromTicker(dom.ticker.value.trim());
        const activeBenchmarkSymbol = dom.benchmarkTicker.value.trim() ? marketDataSymbolFromTicker(dom.benchmarkTicker.value.trim()) : "";
        const activeRange = dom.chartRange.value || "6mo";
        const needsFreshTearSheet =
          !state.priceChartImageBytes ||
          state.equityStats.currentPrice == null ||
          state.equityStats.symbol !== activeSymbol ||
          state.equityStats.benchmarkSymbol !== activeBenchmarkSymbol ||
          state.equityStats.range !== activeRange;

        if (needsFreshTearSheet) {
          await loadEquityMarketSnapshot();
        }
      }
      const data = collectFormData();
      data.noteId = buildNoteId(data);
      const review = buildPrePublishReview(data, validateForm(false));

      if (review.blockingCount > 0) {
        const firstCritical = review.findings.find((finding) => finding.severity === "critical");
        const focusTarget = firstCritical?.focusId ? document.getElementById(firstCritical.focusId) : null;
        if (focusTarget) {
          focusTarget.scrollIntoView({ behavior: "smooth", block: "center" });
          focusTarget.focus();
        }
        setMessage(
          "error",
          `Pre-publish review found ${review.blockingCount} critical issue${review.blockingCount === 1 ? "" : "s"}. Resolve ${review.findings.filter((finding) => finding.severity === "critical").slice(0, 3).map((finding) => finding.title).join(", ")} before export.`
        );
        return;
      }

      const documentFileName = buildDocumentFileName(data);
      const doc = await createDocument(data);
      const blob = await window.docx.Packer.toBlob(doc);
      window.saveAs(blob, documentFileName);
      saveDraft();
      setMessage(
        "success",
        review.warningCount || review.suggestionCount
          ? `Document generated successfully as ${documentFileName}. Note ID ${data.noteId}. Pre-publish review still shows ${review.warningCount} warning${review.warningCount === 1 ? "" : "s"} and ${review.suggestionCount} suggestion${review.suggestionCount === 1 ? "" : "s"}.`
          : `Document generated successfully as ${documentFileName}. Note ID ${data.noteId}.`
      );
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
    const imageFiles = [...(state.figureFiles || [])];

    return {
      noteType: dom.noteType.value.trim(),
      distribution: dom.distribution.value.trim(),
      publicationDate: dom.publicationDate.value.trim(),
      deskLine: getDeskLine(),
      title: dom.title.value.trim(),
      deck: dom.deck.value.trim(),
      topic: dom.topic.value.trim(),
      coverageCountry: dom.coverageCountry.value.trim(),
      issuerId: dom.issuerId.value.trim(),
      macroFiHeading: dom.macroFiHeading.value.trim(),
      agencyRating: dom.agencyRating.value.trim(),
      shortTermRating: dom.shortTermRating.value.trim(),
      longTermRating: dom.longTermRating.value.trim(),
      ratingsProfile: collectRatingsProfile(),
      authorLastName: dom.authorLastName.value.trim(),
      authorFirstName: dom.authorFirstName.value.trim(),
      authorPhone: dom.authorPhone.value.trim(),
      coAuthors: getCoAuthors(),
      ticker: dom.ticker.value.trim(),
      equityCompanyName: dom.equityCompanyName.value.trim(),
      equitySecurityDisplay: dom.equitySecurityDisplay.value.trim(),
      equitySectorLine: dom.equitySectorLine.value.trim(),
      crgRating: dom.crgRating.value.trim(),
      targetPrice: dom.targetPrice.value.trim(),
      priceCurrency: dom.priceCurrency.value.trim(),
      benchmarkName: dom.benchmarkName.value.trim(),
      benchmarkTicker: dom.benchmarkTicker.value.trim(),
      marketCapUsd: dom.marketCapUsd.value.trim(),
      adtUsd: dom.adtUsd.value.trim(),
      businessDescription: dom.businessDescription.value.trim(),
      valuationSummary: dom.valuationSummary.value.trim(),
      financialTableTitle: dom.financialTableTitle.value.trim(),
      financialSourceNote: dom.financialSourceNote.value.trim(),
      financialTableInput: dom.financialTableInput.value.trim(),
      modelLink: dom.modelLink.value.trim(),
      bodySectionLayout: collectBodySectionLayout(),
      keyTakeaways: dom.keyTakeaways.value.trim(),
      analysis: dom.analysis.value.trim(),
      content: dom.content.value.trim(),
      fiveYearRationale: dom.fiveYearRationale.value.trim(),
      esgSummary: dom.esgSummary.value.trim(),
      cordobaView: dom.cordobaView.value.trim(),
      imageFiles,
      figurePlacements: buildCurrentFigurePlacements(imageFiles),
      figureDetails: buildCurrentFigureDetails(imageFiles),
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

  function noteTypeCode(noteType) {
    const map = {
      "General Note": "GN",
      "Equity Research": "EQ",
      "Macro Research": "MA",
      "Fixed Income Research": "FI",
      "Commodity Insights": "CO",
      "Commodity Research": "CO",
      "Short Note / Market Alert": "AL"
    };

    return map[noteType] || "RS";
  }

  function buildNoteId(data) {
    const generatedAt = data.generatedAt || new Date();
    const dateStamp = formatCompactDateStamp(generatedAt);
    const sequence = getNextNoteSequence(noteTypeCode(data.noteType), dateStamp);
    return `CRG-${noteTypeCode(data.noteType)}-${dateStamp}-${sequence}`;
  }

  function formatCompactDateStamp(date) {
    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const year = String(date.getFullYear()).slice(-2);
    return `${day}${month}${year}`;
  }

  function getNextNoteSequence(typeCode, dateStamp) {
    try {
      const raw = window.localStorage.getItem(NOTE_SEQUENCE_STORAGE_KEY);
      const store = raw ? JSON.parse(raw) : {};
      const key = `${typeCode}-${dateStamp}`;
      const nextValue = Number(store[key] || 0) + 1;
      store[key] = nextValue;
      window.localStorage.setItem(NOTE_SEQUENCE_STORAGE_KEY, JSON.stringify(store));
      return String(nextValue).padStart(2, "0");
    } catch (error) {
      console.warn("Unable to persist note sequence:", error);
      const fallback = generatedSequenceFromClock();
      return String(fallback).padStart(2, "0");
    }
  }

  function generatedSequenceFromClock() {
    const now = new Date();
    return ((now.getHours() * 60) + now.getMinutes()) % 99 || 1;
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

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
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

  function cleanupPreviewAssets() {
    state.previewObjectUrls.forEach((url) => URL.revokeObjectURL(url));
    state.previewObjectUrls = [];
  }

  function peekNextNoteSequence(typeCode, dateStamp) {
    try {
      const raw = window.localStorage.getItem(NOTE_SEQUENCE_STORAGE_KEY);
      const store = raw ? JSON.parse(raw) : {};
      const key = `${typeCode}-${dateStamp}`;
      const nextValue = Number(store[key] || 0) + 1;
      return String(nextValue).padStart(2, "0");
    } catch (error) {
      console.warn("Unable to inspect note sequence:", error);
      return String(generatedSequenceFromClock()).padStart(2, "0");
    }
  }

  function buildPreviewNoteId(data) {
    const generatedAt = data.generatedAt || new Date();
    const dateStamp = formatCompactDateStamp(generatedAt);
    return `CRG-${noteTypeCode(data.noteType)}-${dateStamp}-${peekNextNoteSequence(noteTypeCode(data.noteType), dateStamp)}`;
  }

  function buildPreviewAnalystPanelHtml(deskLine, contacts) {
    const people = contacts.length
      ? contacts
      : [{ name: "Research Analyst", email: "research@cordobarg.com", phone: "N/A" }];

    return `
      <section class="preview-analyst-panel">
        <h3>Research Analysts</h3>
        <div class="preview-analyst-desk">${escapeHtml(deskLine || "Global Research Strategy")}</div>
        ${people.map((person) => `
          <div class="preview-analyst-person">
            <strong>${escapeHtml(person.name)}</strong>
            <span>${escapeHtml(person.email)}</span>
            <span>${escapeHtml(person.phone)}</span>
          </div>
        `).join("")}
      </section>
    `;
  }

  function buildPreviewBannerHtml(data, publicationDate) {
    return `
      <header class="preview-export-banner">
        <div class="preview-export-logo">
          <img src="assets/cordoba-logo" alt="Cordoba Research Group logo">
        </div>
        <div class="preview-export-mark" aria-hidden="true">
          <span></span><span></span><span></span><span></span>
        </div>
        <div class="preview-export-brand">
          <strong>Global Markets Research</strong>
          <span>${escapeHtml(formatDocDate(publicationDate))}</span>
        </div>
      </header>
    `;
  }

  function buildPreviewDisplayLineHtml(data) {
    return `
      <div class="preview-export-display">${escapeHtml(nomuraDisplayType(data.noteType))}</div>
      <div class="preview-export-rule">${escapeHtml(data.deskLine || strategyLabelForNoteType(data.noteType) || "Research Note")}</div>
    `;
  }

  function buildPreviewNarrativeHtml(label, content) {
    if (!content) return "";
    return `
      <section class="preview-export-section">
        <h3>${escapeHtml(label)}</h3>
        ${paragraphBlocks(content).map((block) => `<p>${escapeHtml(block).replace(/\n/g, "<br>")}</p>`).join("")}
      </section>
    `;
  }

  function buildPreviewKeyTakeawaysBoxHtml(label, content) {
    const items = lineItems(content);
    return `
      <section class="preview-key-box">
        <h3>${escapeHtml(label)}</h3>
        <ul>
          ${(items.length ? items : ["No key takeaways supplied."]).map((item) => `<li>${escapeHtml(item)}</li>`).join("")}
        </ul>
      </section>
    `;
  }

  function getPreviewChartImageSrc(data) {
    if (data?.priceChartImageBytes?.length) {
      try {
        const blob = new Blob([data.priceChartImageBytes], { type: "image/png" });
        const objectUrl = URL.createObjectURL(blob);
        state.previewObjectUrls.push(objectUrl);
        return objectUrl;
      } catch (error) {
        console.warn("Unable to create preview chart URL from export bytes:", error);
      }
    }

    try {
      if (dom.priceChartCanvas && typeof dom.priceChartCanvas.toDataURL === "function") {
        const dataUrl = dom.priceChartCanvas.toDataURL("image/png");
        return dataUrl && dataUrl !== "data:," ? dataUrl : "";
      }
    } catch (error) {
      console.warn("Unable to create preview chart image:", error);
    }
    return "";
  }

  function buildPreviewFigureMarkup(files, data, availablePlacements, placement) {
    const placedFiles = getFigureFilesForPlacement(data, placement, availablePlacements);
    return placedFiles.map((file, index) => {
      const objectUrl = URL.createObjectURL(file);
      state.previewObjectUrls.push(objectUrl);
      const detail = data.figureDetails?.[figureFileKey(file)] || getFigureDetailForFile(file, index);
      const caption = `${detail.labelType || "Figure"} ${sanitizeFigureNumber(detail.labelNumber, index + 1)}. ${detail.caption || defaultFigureCaption(file)}`;
      return `
        <figure class="preview-figure">
          <img src="${objectUrl}" alt="${escapeAttribute(caption)}">
          <figcaption>${escapeHtml(caption)}</figcaption>
        </figure>
      `;
    }).join("");
  }

  function buildPreviewMacroFiProfileHtml(data) {
    const heading = data.macroFiHeading || (data.noteType === "Fixed Income Research" ? "Ratings Overview" : "Sovereign Ratings");
    const ratingsRows = (data.ratingsProfile || []).filter((row) => row.agency || row.shortTerm || row.longTerm);
    const showMeta = data.noteType === "Macro Research" || data.coverageCountry || data.issuerId;

    return `
      <section class="preview-export-section">
        <h3>${escapeHtml(heading)}</h3>
        ${showMeta ? `
          <div class="preview-macro-meta">
            <div><strong>Coverage Country</strong><span>${escapeHtml(data.coverageCountry ? getCoverageCountryLabel(data.coverageCountry) : "N/A")}</span></div>
            <div><strong>Issuer Number</strong><span>${escapeHtml(data.issuerId || "N/A")}</span></div>
          </div>
        ` : ""}
        <table class="preview-ratings-table">
          <thead>
            <tr><th>Agency</th><th>Short-Term Rating</th><th>Long-Term Rating</th></tr>
          </thead>
          <tbody>
            ${(ratingsRows.length ? ratingsRows : [{ agency: "N/A", shortTerm: "N/A", longTerm: "N/A" }]).map((row) => `
              <tr>
                <td>${escapeHtml(row.agency || "N/A")}</td>
                <td>${escapeHtml(row.shortTerm || "N/A")}</td>
                <td>${escapeHtml(row.longTerm || "N/A")}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </section>
    `;
  }

  function buildPreviewEquityTearSheetHtml(data, publicationDate) {
    const targetPrice = parseNumber(data.targetPrice);
    const marketPrice = Number.isFinite(data.equityStats.currentPrice) ? data.equityStats.currentPrice : null;
    const upside = computeUpsideToTarget(marketPrice, targetPrice);
    const priceDate = data.equityStats.priceDate ? formatDocDate(new Date(`${data.equityStats.priceDate}T00:00:00`)) : formatDocDate(publicationDate);
    const priceCurrency = resolvePriceCurrency(data);
    const chartSrc = getPreviewChartImageSrc(data);
    const rows = [
      { label: "Rating", value: data.crgRating || "N/A" },
      { label: "Target price", value: formatPriceDisplay(targetPrice, priceCurrency) },
      { label: "Closing price", sublabel: priceDate, value: formatPriceDisplay(marketPrice, priceCurrency) },
      { label: "Implied upside", value: upside == null ? "N/A" : formatSignedPercent(upside) },
      { label: "Market Cap (USD mn)", value: formatPlainMetricValue(data.marketCapUsd) },
      { label: "ADT (USD mn)", value: formatPlainMetricValue(data.adtUsd) }
    ];

    return `
      <section class="preview-tearsheet">
        <table class="preview-tearsheet-table">
          <tbody>
            ${rows.map((row) => `
              <tr>
                <td>
                  <strong>${escapeHtml(row.label)}</strong>
                  ${row.sublabel ? `<span>${escapeHtml(row.sublabel)}</span>` : ""}
                </td>
                <td>${escapeHtml(row.value)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div class="preview-chart-block">
          <h4>Relative performance chart</h4>
          <div class="preview-chart-frame">
            ${chartSrc
              ? `<img src="${chartSrc}" alt="Relative performance chart">`
              : `<div class="preview-chart-empty">No tear sheet chart loaded.</div>`}
          </div>
          <p>Source: Yahoo Finance market data${data.equityStats.benchmarkLabel ? `, benchmarked against ${escapeHtml(data.equityStats.benchmarkLabel)}` : ""}</p>
        </div>
      </section>
    `;
  }

  function buildPreviewFinancialTableHtml(data, publicationDate) {
    const matrix = pruneEmptyFinancialRows(parseFinancialTableInput(data.financialTableInput, publicationDate));
    if (!matrix.headers.length) return "";

    const title = data.financialTableTitle || "Year-end 31 Dec";
    const sourceNote = data.financialSourceNote || "Source: Company data, Cordoba Research Group estimates";
    const priceCurrency = resolvePriceCurrency(data);

    return `
      <section class="preview-export-section">
        <table class="preview-financial-table">
          <thead>
            <tr>
              <th>
                <div class="preview-financial-title">${escapeHtml(title)}</div>
                ${priceCurrency ? `<div class="preview-financial-currency">Currency (${escapeHtml(priceCurrency)})</div>` : ""}
              </th>
              ${matrix.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${matrix.rows.map((row) => `
              <tr>
                <td>${escapeHtml(row.label || "-")}</td>
                ${(row.values || []).map((value) => `<td>${escapeHtml(formatFinancialCellForExport(value, row.label))}</td>`).join("")}
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div class="preview-financial-source">${escapeHtml(sourceNote)}</div>
      </section>
    `;
  }

  function buildPreviewSupportHtml(data) {
    if (!data.modelLink && !data.modelFiles.length) return "";

    return `
      <section class="preview-export-section">
        <h3>Model Files</h3>
        ${data.modelLink ? `<p>${escapeHtml(data.modelLink)}</p>` : ""}
        ${data.modelFiles.length ? `<ul>${data.modelFiles.map((file) => `<li>${escapeHtml(file.name)}</li>`).join("")}</ul>` : ""}
      </section>
    `;
  }

  function getCrgRatingDefinitionIntro() {
    return "These are the CRG rating definitions. They reflect our opinion, are not statements of fact, and do not constitute regulated investment advice.";
  }

  function getCrgRatingDefinitionRows() {
    return [
      ["Outperform (O)", "Expected 15%+ upside over 12 months, supported by clear fundamental progress and identifiable catalysts."],
      ["Sector Perform (SP)", "Expected return between -10% and +15% over 12 months, or a broadly balanced risk-reward profile versus the sector."],
      ["Underperform (U)", "Expected return worse than -10% over 12 months, or a meaningfully inferior risk-reward profile versus the sector."],
      ["Restricted (R)", "Applied to non-ethical businesses or activities. CRG may still publish research on such names for market context, sector context, or thematic relevance, but does not assign a recommendation rating."],
      ["Not Rated (NR)", "Coverage is incomplete, conviction is still forming, or available evidence is not yet sufficient for a formal rating call."]
    ];
  }

  function buildPreviewCrgRatingDefinitionsHtml() {
    const rows = getCrgRatingDefinitionRows();

    return `
      <section class="preview-export-section preview-rating-definitions">
        <div class="preview-rating-definitions-head">
          <h3>CRG Rating Definitions</h3>
        </div>
        <p>${escapeHtml(getCrgRatingDefinitionIntro())}</p>
        <table class="preview-ratings-definitions-table">
          <tbody>
            ${rows.map((row) => `
              <tr>
                <th>${escapeHtml(row[0])}</th>
                <td>${escapeHtml(row[1])}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </section>
    `;
  }

  function buildPreviewFooterHtml(noteId, data, pageNumber = 1, totalPages = 1) {
    return `
      <footer class="preview-export-footer">
        <span>Note ID: ${escapeHtml(noteId)}</span>
        <span>Published: ${escapeHtml(formatProductionTimestamp(data.generatedAt))}</span>
        <span>Page ${pageNumber} of ${totalPages}</span>
      </footer>
    `;
  }

  function buildPreviewEquityCrgRatingDefinitionsPageHtml(data, pageNumber, totalPages) {
    const introParagraphs = [
      "CRG equity research ratings are relative opinion classifications intended to frame expected share-price performance over a 12-month horizon against the relevant sector opportunity set and the valuation case presented in the note. Target prices and valuation ranges reflect analytical judgment based on assumptions, scenario weighting, and market inputs that may change without notice.",
      "The definitions below are explanatory only. They are not statements of fact, do not constitute regulated investment advice, and should be read together with the valuation discussion, catalysts, and risk factors set out in the note."
    ];

    return `
      <article class="preview-export-page preview-reference-page">
        <section class="preview-reference-head">
          <h1>CRG Equity Research Rating Definitions</h1>
          ${introParagraphs.map((paragraph) => `<p>${escapeHtml(paragraph)}</p>`).join("")}
        </section>
        <section class="preview-reference-section">
          <h2>STOCKS</h2>
          ${getCrgRatingDefinitionRows().map((row) => `
            <p><strong>${escapeHtml(row[0])}:</strong> ${escapeHtml(row[1])}</p>
          `).join("")}
          <p><strong>Benchmark interpretation:</strong> Unless stated otherwise in the note, rating language should be read in the context of 12-month total return potential relative to sector alternatives, the market-implied base case, and the benchmark context identified in the valuation discussion.</p>
        </section>
        ${buildPreviewFooterHtml(data.noteId, data, pageNumber, totalPages)}
      </article>
    `;
  }

  function buildComplianceNoticeModel(data) {
    const noteReference = [data.title || "Research note", data.noteId ? `Note ID ${data.noteId}` : ""].filter(Boolean).join(" | ");
    const analystNames = collectComplianceAnalystNames(data);
    const analystLabel = analystNames.length ? formatHumanList(analystNames) : "the named analyst(s)";
    const analystVerb = analystNames.length === 1 ? "is" : "are";
    const analystPronoun = analystNames.length === 1 ? "that individual" : "those individuals";
    const publicationTimestamp = formatProductionTimestamp(data.generatedAt);
    const sections = [
      {
        heading: "Status of Publisher",
        subheading: "Regulatory perimeter and institutional status",
        paragraphs: [
          `${noteReference} has been prepared and circulated by Cordoba Research Group solely in connection with its educational, editorial, and training activities. Cordoba Research Group is not authorised, approved, supervised, or otherwise regulated by the Financial Conduct Authority and is not a company registered with the Financial Conduct Authority for the carrying on of investment business, research distribution, brokerage, advisory, arranging, dealing, discretionary management, or any other regulated activity. Cordoba Research Group is not a registered investment firm, broker, investment adviser, wealth manager, regulated research provider, or financial intermediary, and no statement in this note should be read as suggesting otherwise.`,
          `Cordoba Research Group is an educational platform and research group intended to provide students and developing analysts with opportunities to practise institutional-style writing, analytical discipline, editorial judgment, and structured market commentary. The note has therefore been produced within an educational context and must be understood as such. The fact that the note adopts institutional conventions of structure, tone, layout, valuation language, or market terminology does not alter its educational character and does not convert it into regulated investment research, financial advice, or any communication approved by an authorised person for investment purposes.`
        ]
      },
      {
        heading: "Nature of the Material",
        subheading: "Educational and informational publication only",
        paragraphs: [
          `This material is provided strictly for educational and informational purposes. It is intended to illustrate how a research note may be framed, structured, edited, and presented, including the handling of market data, valuation concepts, scenario analysis, risk articulation, and analytical narrative. It is not intended to provide personalised or general investment advice; it does not constitute legal, tax, accounting, regulatory, or financial advice; and it must not be treated as a substitute for advice from a properly authorised or qualified professional adviser.`,
          `Nothing in this note constitutes, or should be construed as constituting, investment advice, a personal recommendation, a regulated investment recommendation, independent research, or investment research prepared in accordance with legal or regulatory requirements designed to promote the independence of investment research. The note is not prepared subject to the rules, controls, separations, monitoring arrangements, research independence safeguards, or dealing restrictions that may apply to research produced by an authorised investment firm or broker-dealer.`
        ]
      },
      {
        heading: "No Offer, No Solicitation, No Financial Promotion",
        subheading: "No invitation to transact and no approved marketing communication",
        paragraphs: [
          `This note does not constitute an offer, invitation, solicitation, inducement, marketing communication, approval communication, placing document, offering memorandum, prospectus, term sheet, dealing instruction, or call to action in relation to any issuer, security, derivative, fund, commodity, rate product, credit instrument, currency, or investment strategy. It must not be treated as an offer to buy or sell, or as a solicitation of an offer to buy or sell, any financial instrument or to participate in any transaction, placement, underwriting, financing, syndication, hedging arrangement, or capital-markets activity.`,
          `This note is not intended to be, and must not be relied upon as, a financial promotion for the purposes of the Financial Services and Markets Act 2000 or any corresponding law or regulation in another jurisdiction. It has not been approved by an authorised person for communication to persons who may require that level of approval, and it must not be circulated in a manner that would cause it to be characterised as an externally approved marketing, distribution, or promotional document. Any ratings, target prices, catalysts, valuation outputs, trade illustrations, or scenario outcomes mentioned in the note are included solely within an educational framework and do not amount to a recommendation or inducement.`
        ]
      },
      {
        heading: "No Advice, Recommendation, Suitability Assessment, or Fiduciary Relationship",
        subheading: "Recipient responsibility for judgment and decision-making",
        paragraphs: [
          `Cordoba Research Group does not provide investment advice and does not make investment recommendations. No statement, chart, scenario, rating, target price, valuation range, market comment, or analytical conclusion in this note should be relied upon as a basis for any investment decision, portfolio construction decision, asset-allocation decision, hedging decision, lending decision, treasury action, suitability assessment, mandate decision, or execution instruction. This material does not take into account the investment objectives, financial situation, regulatory status, legal constraints, tax position, risk tolerance, or portfolio composition of any particular person.`,
          `No fiduciary duty, advisory duty, suitability duty, execution duty, monitoring duty, or continuing duty is owed by Cordoba Research Group or by the named analyst authors to any reader, recipient, institution, or third party by reason of the preparation, circulation, availability, or use of this note. Recipients remain solely responsible for exercising their own independent judgment. Any person considering any action in relation to a market, issuer, or instrument referred to in the note should obtain advice from properly authorised legal, tax, accounting, regulatory, and investment professionals before acting or refraining from acting.`
        ]
      },
      {
        heading: "Sources, Verification, Accuracy, Completeness, and Timeliness",
        subheading: "Public and third-party information; no guarantee of independent verification",
        paragraphs: [
          `Information used in the preparation of this note may have been derived from public filings, company announcements, exchange data, price feeds, benchmark providers, market vendors, published commentary, consensus estimates, analyst presentations, official statistics, third-party databases, and other sources believed, but not guaranteed, to be reliable. Such information may be incomplete, abbreviated, reformatted, rounded, selectively presented, translated, delayed, stale, superseded, or otherwise affected by the limitations of the original source material or by the way in which it has been incorporated into the note.`,
          `Cordoba Research Group does not represent, warrant, or guarantee that any information, statement, chart, table, model output, estimate, assumption, or summary contained in this note is accurate, complete, fair, balanced, timely, reliable, suitable, or fit for any particular purpose, nor that it has been independently verified in every respect. There may be errors, omissions, transcription issues, formatting discrepancies, stale data, or missing contextual information. The absence of any express qualification within the body of the note should not be taken to imply that a statement has been independently verified or remains current at the time of reading.`
        ]
      },
      {
        heading: "Valuations, Forecasts, Forward-Looking Statements, and Market Variables",
        subheading: "Assumption-driven analysis and uncertainty of outcomes",
        paragraphs: [
          `Any forecast, estimate, projection, expected return, target price, scenario, sensitivity, probability, valuation output, implied upside, downside estimate, or other forward-looking statement in this note is inherently uncertain and dependent upon assumptions, methodologies, judgments, simplifications, model choices, and market inputs that may prove to be incorrect, incomplete, or unstable. Different assumptions, data selections, discount rates, benchmark choices, macro views, or analytical frameworks may produce materially different results. Forward-looking statements are not statements of fact and should not be interpreted as assurances or guarantees of future performance or future market levels.`,
          `Actual outcomes may differ materially from any forward-looking view expressed in the note as a result of changes in issuer fundamentals, policy, regulation, rates, currencies, liquidity, funding conditions, market structure, competitive dynamics, geopolitical events, management decisions, commodity prices, spreads, implied volatilities, credit conditions, or broader macroeconomic developments. Prices, yields, spreads, valuations, exchange rates, and market-implied metrics may move adversely and rapidly. Past performance is not indicative of future results. Capital is at risk. Investments may fall as well as rise in value, income may vary, and investors may receive back less than originally invested. Instruments involving leverage, options, swaps, futures, or other derivatives may expose participants to amplified gains and losses and may entail losses in excess of initial capital committed or collateral posted.`
        ]
      },
      {
        heading: "Views, Opinions, Changes Without Notice, and No Duty to Update",
        subheading: "Current only as of publication and subject to revision",
        paragraphs: [
          `Any view, opinion, interpretation, analytical judgment, or conclusion contained in this note is current only as of the publication timestamp stated below or the relevant drafting date reflected in the note and may change at any time without notice. Markets evolve, information changes, company circumstances develop, and analytical judgments may be revised as new information becomes available. Cordoba Research Group and the named analyst authors reserve the right to alter, withdraw, replace, or discontinue any view expressed in the note without any duty to notify readers, recipients, or third parties.`,
          `Neither Cordoba Research Group nor the named analyst authors undertake any obligation to revise, supplement, correct, amend, republish, or update the note, whether as a result of subsequent events, new information, changes in assumptions, data revisions, market developments, or otherwise. Continued storage, forwarding, availability, download, or possession of the note should not be taken to imply that the content remains current, accurate, or complete after the publication date.`
        ]
      },
      {
        heading: "Authorship, Editorial Responsibility, and Internal Divergence of Views",
        subheading: "Named authorship and limits of attribution",
        paragraphs: [
          `${analystLabel} ${analystVerb} responsible for the authorship and substantive content of this note, and the views, observations, interpretations, estimates, and analytical judgments expressed in the note are attributable to ${analystPronoun} in an editorial sense. That attribution reflects authorship responsibility only. It does not imply that the named analyst authors are regulated analysts, approved persons, authorised advisers, or persons producing research under a regulated investment-research regime. It also does not create any duty on the part of the named analyst authors toward any reader or recipient.`,
          `Views expressed in this note may differ from views expressed elsewhere by Cordoba Research Group participants, contributors, reviewers, students, mentors, editors, or associated parties. The existence of named authors should not be read as implying institutional unanimity, house view consistency, supervisory approval of an investment thesis, or endorsement by any external institution. Cordoba Research Group may host, circulate, or archive materials containing different, inconsistent, or subsequently revised viewpoints relating to the same issuer, market, instrument, or macro theme without any obligation to reconcile those views across documents.`
        ]
      },
      {
        heading: "Distribution, Circulation, Audience, and Use Restrictions",
        subheading: "Limitations on onward circulation and jurisdictional use",
        paragraphs: [
          `This note is not directed at retail investors and is not prepared for general public distribution as regulated investment material. It should not be distributed, transmitted, published, excerpted, syndicated, marketed, or relied upon in any jurisdiction or circumstance in which such use would be unlawful or would require registration, authorisation, approval, licensing, filing, or review not obtained by Cordoba Research Group. Each recipient is responsible for satisfying itself that receipt, possession, use, and onward circulation of the note is lawful in the relevant jurisdiction and context.`,
          `Recipients must not circulate this note externally as client research, approved sales material, issuer marketing, a due-diligence substitute, or a regulated communication without first obtaining any editorial, legal, compliance, and supervisory approvals that may be necessary. The note is intended for limited educational use and should be treated accordingly. No recipient may imply that the note has been approved, endorsed, or issued by an authorised investment firm, broker, bank, or regulator merely because it adopts the structure or tone of institutional-style research.`
        ]
      },
      {
        heading: "Intellectual Property and Restrictions on Reproduction",
        subheading: "No reproduction, republication, redistribution, or commercial exploitation without permission",
        paragraphs: [
          `Unless otherwise indicated, the text, structure, compilation, editorial arrangement, and original analytical content of this note are proprietary to Cordoba Research Group and/or the named analyst authors. The note is made available for the limited educational purpose for which it was prepared and may not be copied, reproduced, republished, extracted, adapted, translated, distributed, uploaded, transmitted, licensed, sold, commercialised, or otherwise exploited, in whole or in part, without prior written permission from Cordoba Research Group, except to the extent permitted by applicable law for strictly limited non-commercial internal reference purposes.`,
          `Any unauthorised onward circulation, republication, or commercial use may mischaracterise the status of the note and may expose downstream recipients to material that has not been prepared for their use or regulatory context. Quotations, extracts, screenshots, or republication of tables, charts, or commentary from the note should not be made in a manner that creates the impression that Cordoba Research Group has issued approved investment advice, regulated research, or promotional material.`
        ]
      },
      {
        heading: "Limitation of Liability and Exclusion of Responsibility",
        subheading: "No liability for use, reliance, or consequences of circulation",
        paragraphs: [
          `To the fullest extent permitted by law, Cordoba Research Group disclaims all liability, responsibility, duty of care, and accountability arising directly or indirectly from the preparation, use, reference to, circulation of, or reliance upon this note or any part of it. This exclusion applies whether any claim sounds in contract, tort, negligence, misrepresentation, statutory duty, fiduciary duty, restitution, or otherwise, and extends to direct, indirect, incidental, consequential, punitive, special, exemplary, or other forms of loss or damage, including without limitation trading losses, opportunity costs, lost profits, lost revenues, reputational harm, loss of data, business interruption, financing costs, regulatory consequences, or losses associated with adverse market movements.`,
          `Without limitation, no responsibility is accepted for loss arising from inaccuracies, omissions, stale information, incomplete context, data vendor issues, benchmark errors, model limitations, assumption failures, formatting discrepancies, charting issues, software faults, or subsequent revisions to public information. No representation or warranty, express or implied, is made in respect of merchantability, title, non-infringement, quality, suitability, fitness for a particular purpose, completeness, timeliness, fairness, or reliability. If any part of this notice is held unenforceable, the remaining provisions shall continue to apply to the fullest extent permitted by law.`
        ]
      },
      {
        heading: "Recipient Responsibility and Professional Advice",
        subheading: "Independent assessment remains essential",
        paragraphs: [
          `Each recipient must rely on its own independent assessment and judgment when evaluating any market, issuer, instrument, or transaction referred to in this note. No person should assume that any security, instrument, structure, issuer, benchmark, valuation conclusion, or trade concept mentioned in the note is suitable for their objectives, risk profile, capital base, liquidity needs, constitutional restrictions, mandate terms, fiduciary duties, accounting treatment, tax position, or regulatory obligations. The responsibility for testing assumptions, verifying facts, understanding legal documentation, and assessing suitability remains entirely with the recipient and its advisers.`,
          `Where relevant, recipients should obtain advice from properly authorised legal, tax, accounting, regulatory, and investment professionals before taking or refraining from taking any action in connection with any matter discussed in the note. Any use of the note within an institution, classroom, committee, or investment process should be accompanied by the recipient’s own diligence, challenge, and review procedures. The educational availability of this material does not reduce the need for professional advice or independent verification.`
        ]
      },
      {
        heading: "Record of Publication",
        subheading: "Relationship of this notice to the note as a whole",
        paragraphs: [
          `Published ${publicationTimestamp}. This Regulatory Disclosure and Important Notice forms an integral part of the publication record for ${noteReference} and must be read together with the substantive analysis, charts, tables, appendices, and supporting material that precede it. By accessing, retaining, forwarding, or using this note, the recipient acknowledges the educational, non-advisory, non-promotional, and non-regulated basis on which it has been prepared and circulated.`,
          `If there is any inconsistency between the style, form, structure, or presentation of the note and the regulatory status described in this notice, this notice shall prevail. The institutional appearance of the publication is a matter of educational format and editorial discipline only and should not be interpreted as conferring any regulated status, authorisation, endorsement, or legal effect beyond that expressly stated in this notice.`
        ]
      }
    ];

    return {
      noteReference,
      introduction: `This notice sets out the regulatory status, limitations, authorship attribution, distribution restrictions, and use conditions applicable to ${noteReference}. It should be read in full and together with the substantive body of the note.`,
      sections
    };
  }

  function buildPreviewCompliancePageHtml(data, pageNumber, totalPages) {
    const model = buildComplianceNoticeModel(data);

    return `
      <article class="preview-export-page preview-reference-page preview-legal-page">
        <section class="preview-reference-head">
          <h1>Regulatory Disclosure and Important Notice</h1>
          <p>${escapeHtml(model.introduction)}</p>
        </section>
        ${model.sections.map((section) => `
          <section class="preview-reference-section">
            <h2>${escapeHtml(section.heading)}</h2>
            <h3>${escapeHtml(section.subheading)}</h3>
            ${section.paragraphs.map((paragraph) => `<p>${escapeHtml(paragraph)}</p>`).join("")}
          </section>
        `).join("")}
        ${buildPreviewFooterHtml(data.noteId, data, pageNumber, totalPages)}
      </article>
    `;
  }

  function buildPreviewMainPageHtml(data, publicationDate, analystContacts, availablePlacements, sections, pageNumber, totalPages) {
    const middleSections = sections.filter((entry) => entry.key !== "keyTakeaways" && entry.key !== "cordobaView");
    const keyTakeaways = sections.find((entry) => entry.key === "keyTakeaways");
    const cordobaView = sections.find((entry) => entry.key === "cordobaView");
    const previewParts = [`<article class="preview-export-page">`, buildPreviewBannerHtml(data, publicationDate)];

    if (data.noteType === "Equity Research") {
      previewParts.push(
        `<div class="preview-equity-security-line"><strong>${escapeHtml(data.equityCompanyName || data.title || data.ticker || "Company pending")}</strong>${data.equitySecurityDisplay ? ` <span>${escapeHtml(data.equitySecurityDisplay)}</span>` : ""}</div>`,
        `<div class="preview-equity-sector-strip">EQUITY: ${escapeHtml((data.equitySectorLine || data.topic || "Equity Research").toUpperCase())}</div>`,
        `<div class="preview-equity-front-grid">`,
        `<div class="preview-equity-main">`,
        `<h1 class="preview-export-title">${escapeHtml(data.title || "Untitled Equity Research Note")}</h1>`,
        `<p class="preview-export-deck">${escapeHtml(data.deck || "No deck supplied")}</p>`,
        `<p class="preview-export-topic">${escapeHtml(`${data.topic || data.deskLine || "Equity Research"} | ${formatDocDate(publicationDate)}`)}</p>`,
        buildPreviewKeyTakeawaysBoxHtml(keyTakeaways?.label || "Key Takeaways", keyTakeaways?.content || data.keyTakeaways),
        `</div>`,
        `<aside class="preview-equity-sidebar">`,
        buildPreviewEquityTearSheetHtml(data, publicationDate),
        buildPreviewAnalystPanelHtml(data.deskLine || "Global Equity Strategy", analystContacts),
        `</aside>`,
        `</div>`,
        buildPreviewFinancialTableHtml(data, publicationDate)
      );

      if (data.businessDescription) {
        previewParts.push(buildPreviewNarrativeHtml("Business Description", data.businessDescription));
        previewParts.push(buildPreviewFigureMarkup(data.imageFiles, data, availablePlacements, "after-businessDescription"));
      }
      if (data.valuationSummary) {
        previewParts.push(buildPreviewNarrativeHtml("Valuation Summary", data.valuationSummary));
        previewParts.push(buildPreviewFigureMarkup(data.imageFiles, data, availablePlacements, "after-valuationSummary"));
      }
      middleSections.forEach((section) => {
        previewParts.push(buildPreviewNarrativeHtml(section.label, section.content));
        previewParts.push(buildPreviewFigureMarkup(data.imageFiles, data, availablePlacements, `after-${section.key}`));
      });
      const endFigures = buildPreviewFigureMarkup(data.imageFiles, data, availablePlacements, "end");
      if (endFigures) previewParts.push(`<section class="preview-export-section"><h3>Figures / Screenshots</h3>${endFigures}</section>`);
      previewParts.push(buildPreviewSupportHtml(data));
      if (cordobaView?.content) {
        previewParts.push(buildPreviewNarrativeHtml(cordobaView.label, cordobaView.content));
        previewParts.push(buildPreviewFigureMarkup(data.imageFiles, data, availablePlacements, "after-cordobaView"));
      }
    } else {
      previewParts.push(
        buildPreviewDisplayLineHtml(data),
        `<div class="preview-generic-front-grid">`,
        `<div class="preview-generic-main">`,
        `<h1 class="preview-export-title">${escapeHtml(data.title || "Untitled Research Note")}</h1>`,
        `<p class="preview-export-deck">${escapeHtml(data.deck || "No deck supplied")}</p>`,
        `<p class="preview-export-topic">${escapeHtml(`${data.topic || "Topic pending"} | ${formatDocDate(publicationDate)}`)}</p>`,
        `${data.noteType === "Macro Research" || (isMacroFiNoteType(data.noteType) && (data.coverageCountry || data.issuerId))
          ? `<p class="preview-export-meta-line">${escapeHtml(`Coverage Country: ${data.coverageCountry ? getCoverageCountryLabel(data.coverageCountry) : "N/A"} | Issuer Number: ${data.issuerId || "N/A"}`)}</p>`
          : ""}`,
        `</div>`,
        `<aside class="preview-generic-side">${buildPreviewAnalystPanelHtml(data.deskLine || strategyLabelForNoteType(data.noteType) || "Global Research Strategy", analystContacts)}</aside>`,
        `</div>`
      );

      if (isMacroFiNoteType(data.noteType)) {
        previewParts.push(buildPreviewMacroFiProfileHtml(data));
        previewParts.push(buildPreviewFigureMarkup(data.imageFiles, data, availablePlacements, "after-macroFiProfile"));
      }

      previewParts.push(buildPreviewKeyTakeawaysBoxHtml(keyTakeaways?.label || "Key Takeaways", keyTakeaways?.content || data.keyTakeaways));
      previewParts.push(buildPreviewFigureMarkup(data.imageFiles, data, availablePlacements, "after-keyTakeaways"));

      middleSections.forEach((section) => {
        previewParts.push(buildPreviewNarrativeHtml(section.label, section.content));
        previewParts.push(buildPreviewFigureMarkup(data.imageFiles, data, availablePlacements, `after-${section.key}`));
      });

      if (cordobaView?.content) {
        previewParts.push(buildPreviewNarrativeHtml(cordobaView.label, cordobaView.content));
        previewParts.push(buildPreviewFigureMarkup(data.imageFiles, data, availablePlacements, "after-cordobaView"));
      }

      const endFigures = buildPreviewFigureMarkup(data.imageFiles, data, availablePlacements, "end");
      if (endFigures) previewParts.push(`<section class="preview-export-section"><h3>Figures / Screenshots</h3>${endFigures}</section>`);
      previewParts.push(buildPreviewSupportHtml(data));
    }

    previewParts.push(buildPreviewFooterHtml(data.noteId, data, pageNumber, totalPages), `</article>`);
    return previewParts.join("");
  }

  function buildPreviewDocumentSrcdoc(data, pagesHtml) {
    const baseHref = escapeAttribute(document.baseURI || window.location.href || "");
    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <base href="${baseHref}">
  <link rel="stylesheet" href="assets/styles.css">
  <style>
    html, body {
      margin: 0;
      padding: 0;
      background: #edf0f5;
      min-height: 100%;
    }
    body {
      font-family: "Aptos", "Segoe UI", "Helvetica Neue", Arial, sans-serif;
      color: #11151c;
    }
    .preview-document-shell {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 28px;
      padding: 28px 34px 34px;
    }
    @page {
      size: A4;
      margin: 0;
    }
  </style>
</head>
<body>
  <div class="preview-document-shell">${pagesHtml}</div>
</body>
</html>`;
  }

  function openPreviewModal() {
    if (!dom.previewModal || !dom.previewModalBody) return;

    cleanupPreviewAssets();
    const data = collectFormData();
    data.noteId = buildPreviewNoteId(data);
    const publicationDate = parseInputDate(data.publicationDate) || data.generatedAt || new Date();
    const analystContacts = collectAnalystContacts(data);
    const availablePlacements = new Set(getFigurePlacementOptions(data.noteType).map((option) => option.value));
    const sections = normalizeBodySectionLayoutForExport(data).filter((entry) => !entry.hidden && entry.content);
    const previewMarkup = buildPreviewMainPageHtml(
      data,
      publicationDate,
      analystContacts,
      availablePlacements,
      sections,
      1,
      1
    );

    dom.previewModalBody.innerHTML = `<div class="preview-page-stage">${previewMarkup}</div>`;
    dom.previewModal.hidden = false;
  }

  function closePreviewModal() {
    if (!dom.previewModal || !dom.previewModalBody) return;
    dom.previewModal.hidden = true;
    dom.previewModalBody.innerHTML = "";
    cleanupPreviewAssets();
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
      takeawayFill: "FFF9F4",
      takeawayBorder: "EADCCD",
      white: "FFFFFF"
    };

    const publicationDate = parseInputDate(data.publicationDate) || data.generatedAt || new Date();
    const bannerBytes = await buildNomuraBannerImageBytes({
      noteType: data.noteType,
      deskLine: data.deskLine,
      publicationDate
    });
    const analystContacts = collectAnalystContacts(data);
    const { keyTakeaways, middle, cordobaView } = getNarrativeSectionPartitions(data);
    const figureCounterRef = { value: 1 };
    const availablePlacements = new Set(["end", "after-keyTakeaways"]);
    if (isMacroFiNoteType(data.noteType)) availablePlacements.add("after-macroFiProfile");
    middle.forEach((section) => availablePlacements.add(`after-${section.key}`));
    if (cordobaView?.content) availablePlacements.add("after-cordobaView");

    if (data.noteType === "Equity Research") {
      return createEquityResearchDocument(docxLib, colors, data, publicationDate, bannerBytes, analystContacts, {
        keyTakeaways,
        middle,
        cordobaView
      }, figureCounterRef);
    }

    const documentChildren = [
      new docxLib.Paragraph({
        children: [
          new docxLib.ImageRun({
            data: bannerBytes,
            transformation: { width: 705, height: 122 }
          })
        ],
        spacing: { after: 110 }
      }),
      buildNomuraTitlePanel(docxLib, colors, data, analystContacts, publicationDate)
    ];

    if (isMacroFiNoteType(data.noteType)) {
      documentChildren.push(
        new docxLib.Paragraph({ spacing: { after: 34 } }),
        ...buildMacroFiProfileTable(docxLib, colors, data),
        new docxLib.Paragraph({ spacing: { after: 54 } })
      );
      await appendPlacedFigures(
        documentChildren,
        docxLib,
        colors,
        getFigureFilesForPlacement(data, "after-macroFiProfile", availablePlacements),
        figureCounterRef,
        { figureDetails: data.figureDetails }
      );
    } else {
      documentChildren.push(new docxLib.Paragraph({ spacing: { after: 70 } }));
    }

    documentChildren.push(
      buildKeyTakeawaysBox(docxLib, colors, lineItems(keyTakeaways.content), keyTakeaways.label)
    );
    await appendPlacedFigures(
      documentChildren,
      docxLib,
      colors,
      getFigureFilesForPlacement(data, "after-keyTakeaways", availablePlacements),
      figureCounterRef,
      { figureDetails: data.figureDetails }
    );

    for (const section of middle) {
      documentChildren.push(...buildNarrativeSectionBlocks(docxLib, colors, [section]));
      await appendPlacedFigures(
        documentChildren,
        docxLib,
        colors,
        getFigureFilesForPlacement(data, `after-${section.key}`, availablePlacements),
        figureCounterRef,
        { figureDetails: data.figureDetails }
      );
    }

    if (cordobaView?.content) {
      documentChildren.push(...buildNarrativeSectionBlocks(docxLib, colors, [cordobaView]));
      await appendPlacedFigures(
        documentChildren,
        docxLib,
        colors,
        getFigureFilesForPlacement(data, "after-cordobaView", availablePlacements),
        figureCounterRef,
        { figureDetails: data.figureDetails }
      );
    }

    const endFigures = getFigureFilesForPlacement(data, "end", availablePlacements);
    if (endFigures.length) {
      await appendPlacedFigures(documentChildren, docxLib, colors, endFigures, figureCounterRef, {
        withHeading: true,
        heading: "Figures / Screenshots",
        figureDetails: data.figureDetails
      });
    }

    const supportParagraphs = buildSupportingMaterialParagraphs(docxLib, colors, data);
    if (supportParagraphs.length) {
      documentChildren.push(buildNomuraSubhead(docxLib, colors, "Model Files"), ...supportParagraphs);
    }

    return buildResearchDocumentShell(
      docxLib,
      colors,
      publicationDate,
      data.generatedAt,
      data.noteId,
      documentChildren,
      buildComplianceDeclarationSection(docxLib, colors, data)
    );
  }

  function buildResearchDocumentShell(docxLib, colors, publicationDate, generatedAt, noteId, children, finalPrintChildren = []) {
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
              margin: { top: 520, right: 666, bottom: 660, left: 666, footer: 180 },
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
                  spacing: { after: 40 }
                }),
                buildFooterMetaTable(docxLib, colors, publicationDate, generatedAt, noteId)
              ]
            })
          },
          children: [...children, ...finalPrintChildren]
        }
      ]
    });
  }

  async function createEquityResearchDocument(docxLib, colors, data, publicationDate, bannerBytes, analystContacts, narrativeSections, figureCounterRef) {
    const { keyTakeaways, middle, cordobaView } = narrativeSections;
    const availablePlacements = new Set(["end"]);
    if (data.businessDescription) availablePlacements.add("after-businessDescription");
    if (data.valuationSummary) availablePlacements.add("after-valuationSummary");
    middle.forEach((section) => availablePlacements.add(`after-${section.key}`));
    if (cordobaView?.content) availablePlacements.add("after-cordobaView");
    const children = [
      new docxLib.Paragraph({
        children: [
          new docxLib.ImageRun({
            data: bannerBytes,
            transformation: { width: 705, height: 122 }
          })
        ],
        spacing: { after: 45 }
      }),
      buildEquitySecurityLine(docxLib, colors, data),
      buildEquitySectorStrip(docxLib, colors, data),
      buildEquityFrontPageTable(docxLib, colors, data, publicationDate, analystContacts, keyTakeaways),
      ...buildEquityFinancialTable(docxLib, colors, data, publicationDate)
    ];

    if (data.businessDescription) {
      children.push(
        buildNomuraSubhead(docxLib, colors, "Business Description"),
        ...buildNomuraBodyParagraphs(docxLib, data.businessDescription)
      );
      await appendPlacedFigures(
        children,
        docxLib,
        colors,
        getFigureFilesForPlacement(data, "after-businessDescription", availablePlacements),
        figureCounterRef,
        { figureDetails: data.figureDetails }
      );
    }

    if (data.valuationSummary) {
      children.push(
        buildNomuraSubhead(docxLib, colors, "Valuation Summary"),
        ...buildNomuraBodyParagraphs(docxLib, data.valuationSummary)
      );
      await appendPlacedFigures(
        children,
        docxLib,
        colors,
        getFigureFilesForPlacement(data, "after-valuationSummary", availablePlacements),
        figureCounterRef,
        { figureDetails: data.figureDetails }
      );
    }

    for (const section of middle) {
      children.push(...buildNarrativeSectionBlocks(docxLib, colors, [section]));
      await appendPlacedFigures(
        children,
        docxLib,
        colors,
        getFigureFilesForPlacement(data, `after-${section.key}`, availablePlacements),
        figureCounterRef,
        { figureDetails: data.figureDetails }
      );
    }

    const endFigures = getFigureFilesForPlacement(data, "end", availablePlacements);
    if (endFigures.length) {
      await appendPlacedFigures(children, docxLib, colors, endFigures, figureCounterRef, {
        withHeading: true,
        heading: "Figures / Screenshots",
        figureDetails: data.figureDetails
      });
    }

    const supportParagraphs = buildSupportingMaterialParagraphs(docxLib, colors, data);
    if (supportParagraphs.length) {
      children.push(buildNomuraSubhead(docxLib, colors, "Model Files"), ...supportParagraphs);
    }

    if (cordobaView?.content) {
      children.push(...buildNarrativeSectionBlocks(docxLib, colors, [cordobaView]));
      await appendPlacedFigures(
        children,
        docxLib,
        colors,
        getFigureFilesForPlacement(data, "after-cordobaView", availablePlacements),
        figureCounterRef,
        { figureDetails: data.figureDetails }
      );
    }

    return buildResearchDocumentShell(
      docxLib,
      colors,
      publicationDate,
      data.generatedAt,
      data.noteId,
      children,
      [
        ...buildCrgRatingMethodologySection(docxLib, colors),
        ...buildComplianceDeclarationSection(docxLib, colors, data, { pageBreakBefore: false })
      ]
    );
  }

  function buildEquitySecurityLine(docxLib, colors, data) {
    const companyName = data.equityCompanyName || data.title || data.ticker || "Company pending";
    const securityDisplay = data.equitySecurityDisplay || (data.ticker ? data.ticker.toUpperCase() : "");

    return new docxLib.Paragraph({
      children: [
        new docxLib.TextRun({
          text: companyName,
          bold: true,
          size: 20,
          color: colors.black,
          font: "Arial"
        }),
        new docxLib.TextRun({
          text: securityDisplay ? ` ${securityDisplay}` : "",
          size: 14,
          color: colors.muted,
          font: "Arial"
        })
      ],
      spacing: { after: 24 }
    });
  }

  function buildEquitySectorStrip(docxLib, colors, data) {
    const sectorLine = (data.equitySectorLine || data.topic || "Equity Research").toUpperCase();

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
              shading: { fill: colors.redDeep },
              margins: { top: 25, bottom: 25, left: 90, right: 90 },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              children: [
                new docxLib.Paragraph({
                  children: [
                    new docxLib.TextRun({
                      text: `EQUITY: ${sectorLine}`,
                      size: 11,
                      color: colors.white,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 0 }
                })
              ]
            })
          ]
        })
      ]
    });
  }

  function buildEquityFrontPageTable(docxLib, colors, data, publicationDate, analystContacts, keyTakeawaysSection) {
    const leftColumnChildren = [
      new docxLib.Paragraph({
        children: [
          new docxLib.TextRun({
            text: data.title || "Untitled Equity Research Note",
            bold: true,
            size: 33,
            color: colors.black,
            font: "Arial"
          })
        ],
        spacing: { before: 80, after: 54 }
      }),
      new docxLib.Paragraph({
        children: [
          new docxLib.TextRun({
            text: data.deck || "No deck supplied",
            bold: true,
            size: 21,
            color: "656D78",
            font: "Arial"
          })
        ],
        spacing: { after: 42 }
      }),
      new docxLib.Paragraph({
        children: [
          new docxLib.TextRun({
            text: data.topic || data.deskLine || "Equity Research",
            size: 14,
            color: colors.muted,
            font: "Arial"
          })
        ],
        spacing: { after: 46 }
      }),
      buildKeyTakeawaysBox(
        docxLib,
        colors,
        lineItems(keyTakeawaysSection?.content || data.keyTakeaways),
        keyTakeawaysSection?.label || "Key Takeaways"
      )
    ];

    const rightColumnChildren = [
      buildEquityTearSheetTable(docxLib, colors, data, publicationDate),
      ...buildEquitySidebarChart(docxLib, colors, data),
      buildEquityAnalystSidebar(docxLib, colors, data, analystContacts)
    ];

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
              width: { size: 68, type: docxLib.WidthType.PERCENTAGE },
              verticalAlign: docxLib.VerticalAlign.TOP,
              margins: { top: 0, bottom: 0, left: 0, right: 150 },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              children: leftColumnChildren
            }),
            new docxLib.TableCell({
              width: { size: 32, type: docxLib.WidthType.PERCENTAGE },
              verticalAlign: docxLib.VerticalAlign.TOP,
              margins: { top: 0, bottom: 0, left: 0, right: 0 },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              children: rightColumnChildren
            })
          ]
        })
      ]
    });
  }

  function buildEquityTearSheetTable(docxLib, colors, data, publicationDate) {
    const targetPrice = parseNumber(data.targetPrice);
    const marketPrice = Number.isFinite(data.equityStats.currentPrice) ? data.equityStats.currentPrice : null;
    const upside = computeUpsideToTarget(marketPrice, targetPrice);
    const priceDate = data.equityStats.priceDate ? formatDocDate(new Date(`${data.equityStats.priceDate}T00:00:00`)) : formatDocDate(publicationDate);
    const priceCurrency = resolvePriceCurrency(data);
    const tearSheetRows = [
      { label: "Rating", value: data.crgRating || "N/A" },
      { label: "Target price", value: formatPriceDisplay(targetPrice, priceCurrency) },
      { label: "Closing price", sublabel: priceDate, value: formatPriceDisplay(marketPrice, priceCurrency) },
      { label: "Implied upside", value: upside == null ? "N/A" : formatSignedPercent(upside) },
      { label: "Market Cap (USD mn)", value: formatPlainMetricValue(data.marketCapUsd) },
      { label: "ADT (USD mn)", value: formatPlainMetricValue(data.adtUsd) }
    ];

    return new docxLib.Table({
      width: { size: 100, type: docxLib.WidthType.PERCENTAGE },
      borders: {
        top: { style: docxLib.BorderStyle.NONE },
        bottom: { style: docxLib.BorderStyle.NONE },
        left: { style: docxLib.BorderStyle.NONE },
        right: { style: docxLib.BorderStyle.NONE },
        insideHorizontal: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        insideVertical: { style: docxLib.BorderStyle.NONE }
      },
      rows: tearSheetRows.map((row) =>
        new docxLib.TableRow({
          children: [
            new docxLib.TableCell({
              width: { size: 56, type: docxLib.WidthType.PERCENTAGE },
              margins: { top: 55, bottom: 55, left: 0, right: 40 },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              children: [
                new docxLib.Paragraph({
                  children: [
                    new docxLib.TextRun({
                      text: row.label,
                      size: 12,
                      color: colors.muted,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: row.sublabel ? 10 : 0 }
                }),
                ...(row.sublabel
                  ? [new docxLib.Paragraph({
                    children: [
                      new docxLib.TextRun({
                        text: row.sublabel,
                        size: 11,
                        color: colors.muted,
                        font: "Arial"
                      })
                    ],
                    spacing: { after: 0 }
                  })]
                  : [])
              ]
            }),
            new docxLib.TableCell({
              width: { size: 44, type: docxLib.WidthType.PERCENTAGE },
              margins: { top: 55, bottom: 55, left: 40, right: 0 },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              children: [
                new docxLib.Paragraph({
                  alignment: docxLib.AlignmentType.RIGHT,
                  children: [
                    new docxLib.TextRun({
                      text: row.value,
                      bold: true,
                      size: row.label === "Market Cap (USD mn)" || row.label === "ADT (USD mn)" ? 14 : 16,
                      color: colors.black,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 0 }
                })
              ]
            })
          ]
        })
      )
    });
  }

  function buildEquitySidebarChart(docxLib, colors, data) {
    const children = [
      new docxLib.Paragraph({
        children: [
          new docxLib.TextRun({
            text: "Relative performance chart",
            bold: true,
            size: 13,
            color: colors.black,
            font: "Arial"
          })
        ],
        spacing: { before: 95, after: 30 }
      })
    ];

    if (data.priceChartImageBytes) {
      children.push(
        new docxLib.Paragraph({
          children: [
            new docxLib.ImageRun({
              data: data.priceChartImageBytes,
              transformation: { width: 208, height: 132 }
            })
          ],
          spacing: { after: 20 }
        })
      );
    } else {
      children.push(
        new docxLib.Paragraph({
          border: {
            top: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
            bottom: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
            left: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
            right: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 }
          },
          shading: { fill: colors.soft },
          children: [
            new docxLib.TextRun({
              text: "No tear sheet chart loaded.",
              size: 12,
              color: colors.muted,
              font: "Arial"
            })
          ],
          spacing: { after: 20 },
          indent: { left: 40, right: 40 }
        })
      );
    }

    children.push(
      new docxLib.Paragraph({
        children: [
          new docxLib.TextRun({
            text: `Source: Yahoo Finance market data${data.equityStats.benchmarkLabel ? `, benchmarked against ${data.equityStats.benchmarkLabel}` : ""}`,
            size: 11,
            color: colors.muted,
            font: "Arial"
          })
        ],
        spacing: { after: 80 }
      })
    );

    return children;
  }

  function buildEquityAnalystSidebar(docxLib, colors, data, analystContacts) {
    const analystDesk = data.deskLine || "Global Equity Strategy";

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
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              margins: { top: 0, bottom: 0, left: 0, right: 0 },
              children: [
                ...buildAnalystContactPanel(docxLib, colors, analystDesk, analystContacts)
              ]
            })
          ]
        })
      ]
    });
  }

  function buildEquityFinancialTable(docxLib, colors, data, publicationDate) {
    const matrix = pruneEmptyFinancialRows(parseFinancialTableInput(data.financialTableInput, publicationDate));
    const title = data.financialTableTitle || "Year-end 31 Dec";
    const sourceNote = data.financialSourceNote || "Source: Company data, Cordoba Research Group estimates";
    const priceCurrency = resolvePriceCurrency(data);
    const columnCount = matrix.headers.length;
    const valueColumnWidth = columnCount ? Math.max(11, Math.floor(66 / columnCount)) : 16;
    const metricColumnWidth = 100 - (valueColumnWidth * columnCount);

    const headerRow = new docxLib.TableRow({
      children: [
        new docxLib.TableCell({
          width: { size: metricColumnWidth, type: docxLib.WidthType.PERCENTAGE },
          shading: { fill: colors.soft },
          margins: { top: 55, bottom: 55, left: 70, right: 70 },
          children: [
            new docxLib.Paragraph({
              children: [
                new docxLib.TextRun({
                  text: title,
                  bold: true,
                  size: 12,
                  color: colors.black,
                    font: "Arial"
                  })
                ],
              spacing: { after: priceCurrency ? 8 : 0 }
            }),
            ...(priceCurrency
              ? [new docxLib.Paragraph({
                children: [
                  new docxLib.TextRun({
                    text: `Currency (${priceCurrency})`,
                    size: 10,
                    color: colors.muted,
                    font: "Arial"
                  })
                ],
                spacing: { after: 0 }
              })]
              : [])
          ]
        }),
        ...matrix.headers.map((header) =>
          new docxLib.TableCell({
            width: { size: valueColumnWidth, type: docxLib.WidthType.PERCENTAGE },
            shading: { fill: colors.soft },
            margins: { top: 55, bottom: 55, left: 50, right: 50 },
            children: [
              new docxLib.Paragraph({
                alignment: docxLib.AlignmentType.CENTER,
                children: [
                  new docxLib.TextRun({
                    text: header,
                    bold: true,
                    size: 12,
                    color: colors.black,
                    font: "Arial"
                  })
                ],
                spacing: { after: 0 }
              })
            ]
          })
        )
      ]
    });

    const dataRows = matrix.rows.map((row) =>
      new docxLib.TableRow({
        children: [
          new docxLib.TableCell({
            width: { size: metricColumnWidth, type: docxLib.WidthType.PERCENTAGE },
            margins: { top: 45, bottom: 45, left: 70, right: 70 },
            children: [
              new docxLib.Paragraph({
                children: [
                  new docxLib.TextRun({
                    text: row.label,
                    bold: true,
                    size: 11,
                    color: colors.black,
                    font: "Arial"
                  })
                ],
                spacing: { after: 0 }
              })
            ]
          }),
          ...row.values.map((value) =>
            new docxLib.TableCell({
              width: { size: valueColumnWidth, type: docxLib.WidthType.PERCENTAGE },
              margins: { top: 45, bottom: 45, left: 50, right: 50 },
              children: [
                new docxLib.Paragraph({
                  alignment: docxLib.AlignmentType.CENTER,
                  children: [
                    new docxLib.TextRun({
                      text: formatFinancialCellForExport(value, row.label),
                      size: 11,
                      color: colors.black,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 0 }
                })
              ]
            })
          )
        ]
      })
    );

    return [
      new docxLib.Paragraph({ spacing: { before: 65, after: 10 } }),
      new docxLib.Table({
        width: { size: 100, type: docxLib.WidthType.PERCENTAGE },
        borders: {
          top: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
          bottom: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
          left: { style: docxLib.BorderStyle.NONE },
          right: { style: docxLib.BorderStyle.NONE },
          insideHorizontal: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
          insideVertical: { style: docxLib.BorderStyle.NONE }
        },
        rows: [headerRow, ...dataRows]
      }),
      new docxLib.Paragraph({
        children: [
          new docxLib.TextRun({
            text: sourceNote,
            italics: true,
            size: 11,
            color: colors.muted,
            font: "Arial"
          })
        ],
        spacing: { before: 18, after: 65 }
      })
    ];
  }

  function pruneEmptyFinancialRows(matrix) {
    const rows = (matrix?.rows || []).filter((row) =>
      String(row?.label || "").trim() ||
      (row?.values || []).some((value) => String(value || "").trim())
    );

    return {
      headers: matrix?.headers || [],
      rows: rows.length ? rows : buildDefaultFinancialRows((matrix?.headers || []).length || 4)
    };
  }

  function parseFinancialTableInput(text, publicationDate) {
    const raw = String(text || "").trim();
    if (!raw) return normalizeFinancialMatrix(null, publicationDate);

    if (raw.startsWith("{")) {
      try {
        const parsed = JSON.parse(raw);
        return normalizeFinancialMatrix(parsed, publicationDate);
      } catch (error) {
        console.warn("Unable to parse structured financial grid data, falling back to delimited parsing:", error);
      }
    }

    return parseDelimitedFinancialTable(raw, publicationDate);
  }

  function parseDelimitedFinancialTable(text, publicationDate) {
    const lines = String(text || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);

    if (!lines.length) return normalizeFinancialMatrix(null, publicationDate);

    const delimiter = detectTableDelimiter(lines[0]);
    const parsedLines = lines.map((line) => splitDelimitedLine(line, delimiter).map((cell) => cell.trim()));
    const headerCells = parsedLines[0];
    const hasExplicitHeader = headerCells.length > 1;
    const headers = hasExplicitHeader
      ? headerCells.slice(1, 7).filter(Boolean)
      : buildDefaultFinancialHeaders(publicationDate);

    const rows = parsedLines
      .slice(hasExplicitHeader ? 1 : 0)
      .map((cells) => ({
        label: cells[0] || "",
        values: normalizeFinancialValues(cells.slice(1), headers.length)
      }))
      .filter((row) => row.label || row.values.some((value) => String(value || "").trim()));

    return normalizeFinancialMatrix({ headers, rows }, publicationDate);
  }

  function normalizeFinancialMatrix(matrix, publicationDate) {
    const fallbackHeaders = buildDefaultFinancialHeaders(publicationDate);
    const inputHeaders = Array.isArray(matrix?.headers) && matrix.headers.length ? matrix.headers.slice(0, 6) : fallbackHeaders;
    const normalizedHeaders = inputHeaders.map((header, index) => String(header || "").trim() || fallbackHeaders[index] || `Period ${index + 1}`);
    const rawRows = Array.isArray(matrix?.rows) && matrix.rows.length ? matrix.rows : buildDefaultFinancialRows(normalizedHeaders.length);
    const rows = rawRows
      .map((row) => ({
        label: String(row?.label || ""),
        values: normalizeFinancialValues(Array.isArray(row?.values) ? row.values : [], normalizedHeaders.length)
      }));

    return {
      headers: normalizedHeaders,
      rows: rows.length ? rows : buildDefaultFinancialRows(normalizedHeaders.length)
    };
  }

  function buildDefaultFinancialHeaders(publicationDate) {
    const year = publicationDate.getFullYear();
    return [
      `FY${String(year - 1).slice(-2)}A`,
      `FY${String(year).slice(-2)}F`,
      `FY${String(year + 1).slice(-2)}F`,
      `FY${String(year + 2).slice(-2)}F`
    ];
  }

  function buildDefaultFinancialRows(columnCount) {
    const metrics = [
      "Revenue (mn)",
      "Reported net profit (mn)",
      "Normalised net profit (mn)",
      "FD norm. EPS",
      "FD norm. EPS growth (%)",
      "FD normalised PE (x)",
      "EV/EBITDA (x)",
      "Price/book (x)",
      "ROE (%)",
      "Net debt/equity (%)"
    ];

    return metrics.map((label) => ({
      label,
      values: Array.from({ length: columnCount }, () => "")
    }));
  }

  function detectTableDelimiter(line) {
    const candidates = ["\t", ",", "|", ";"];
    let best = "\t";
    let bestCount = -1;

    candidates.forEach((candidate) => {
      const count = line.split(candidate).length - 1;
      if (count > bestCount) {
        best = candidate;
        bestCount = count;
      }
    });

    return best;
  }

  function splitDelimitedLine(line, delimiter) {
    if (delimiter === "\t") return line.split("\t");

    const cells = [];
    let current = "";
    let inQuotes = false;

    for (let index = 0; index < line.length; index += 1) {
      const char = line[index];
      if (char === "\"") {
        inQuotes = !inQuotes;
        continue;
      }
      if (char === delimiter && !inQuotes) {
        cells.push(current);
        current = "";
        continue;
      }
      current += char;
    }

    cells.push(current);
    return cells;
  }

  function normalizeFinancialValues(values, targetLength) {
    const normalized = values
      .slice(0, targetLength)
      .map((value) => normalizeFinancialCellValue(value));
    while (normalized.length < targetLength) normalized.push("");
    return normalized;
  }

  function normalizeFinancialCellValue(value) {
    const text = String(value ?? "").trim();
    return text === "-" ? "" : text;
  }

  function hasMeaningfulFinancialData(matrix) {
    return (matrix?.rows || []).some((row) =>
      (row.values || []).some((value) => {
        const text = String(value || "").trim();
        return text && text !== "-";
      })
    );
  }

  function formatFinancialCellForExport(value, rowLabel) {
    const formatted = formatFinancialValueForDisplay(value, rowLabel);
    return formatted || "-";
  }

  function formatPlainMetricValue(value) {
    const numeric = parseNumber(value);
    if (numeric == null) return value ? value : "N/A";
    return new Intl.NumberFormat("en-US", {
      minimumFractionDigits: numeric % 1 === 0 ? 0 : 1,
      maximumFractionDigits: 1
    }).format(numeric);
  }

  function buildCordobaEmail(firstName, lastName) {
    const first = String(firstName || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "");
    const fallback = [firstName, lastName]
      .filter(Boolean)
      .join("")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "");
    const localPart = first || fallback || "research";
    return `${localPart}@cordobarg.com`;
  }

  function collectAnalystContacts(data) {
    const entries = [
      {
        firstName: data.authorFirstName,
        lastName: data.authorLastName,
        phone: data.authorPhone
      },
      ...(data.coAuthors || []).map((coAuthor) => ({
        firstName: coAuthor.firstName,
        lastName: coAuthor.lastName,
        phone: coAuthor.phone
      }))
    ];

    return entries
      .map((entry) => {
        const name = [entry.firstName, entry.lastName].filter(Boolean).join(" ").trim();
        if (!name) return null;
        return {
          name,
          email: buildCordobaEmail(entry.firstName, entry.lastName),
          phone: formatPhoneDisplay(entry.phone)
        };
      })
      .filter(Boolean);
  }

  function buildAnalystContactPanel(docxLib, colors, deskLine, contacts) {
    const people = contacts.length
      ? contacts
      : [{ name: "Research Analyst", email: "research@cordobarg.com", phone: "N/A" }];

    return [
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
        spacing: { after: 28 }
      }),
      new docxLib.Paragraph({
        border: {
          bottom: {
            color: colors.red,
            style: docxLib.BorderStyle.SINGLE,
            size: 5
          }
        },
        children: [
          new docxLib.TextRun({
            text: deskLine || "Global Research Strategy",
            color: colors.red,
            size: 14,
            font: "Arial"
          })
        ],
        spacing: { after: 42 }
      }),
      ...people.flatMap((person, index) => ([
        new docxLib.Paragraph({
          children: [
            new docxLib.TextRun({
              text: person.name,
              bold: true,
              size: 14,
              color: colors.black,
              font: "Arial"
            })
          ],
          spacing: { after: 8 }
        }),
        new docxLib.Paragraph({
          children: [
            new docxLib.TextRun({
              text: person.email,
              size: 12,
              color: colors.ink,
              font: "Arial"
            })
          ],
          spacing: { after: 8 }
        }),
        new docxLib.Paragraph({
          children: [
            new docxLib.TextRun({
              text: person.phone || "N/A",
              size: 12,
              color: colors.ink,
              font: "Arial"
            })
          ],
          spacing: { after: index === people.length - 1 ? 0 : 28 }
        })
      ]))
    ];
  }

  function buildMacroFiProfileTable(docxLib, colors, data) {
    const heading = data.macroFiHeading || (data.noteType === "Fixed Income Research" ? "Ratings Overview" : "Sovereign Ratings");
    const countryLabel = getCoverageCountryLabel(data.coverageCountry);
    const issuerNumber = data.issuerId || "N/A";
    const showMetadataTable = data.noteType === "Macro Research" || Boolean(data.coverageCountry || data.issuerId);
    const ratings = (Array.isArray(data.ratingsProfile) && data.ratingsProfile.length
      ? data.ratingsProfile
      : [{
        agency: data.agencyRating,
        shortTerm: data.shortTermRating,
        longTerm: data.longTermRating
      }])
      .filter((row) => row.agency || row.shortTerm || row.longTerm);

    const metadataTable = new docxLib.Table({
      width: { size: 100, type: docxLib.WidthType.PERCENTAGE },
      borders: {
        top: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        bottom: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        left: { style: docxLib.BorderStyle.NONE },
        right: { style: docxLib.BorderStyle.NONE },
        insideHorizontal: { style: docxLib.BorderStyle.NONE },
        insideVertical: { style: docxLib.BorderStyle.NONE }
      },
      rows: [
        new docxLib.TableRow({
          children: [
            ...[
              { label: "Coverage Country", value: countryLabel },
              { label: "Issuer Number", value: issuerNumber }
            ].map((cell) =>
              new docxLib.TableCell({
                width: { size: 50, type: docxLib.WidthType.PERCENTAGE },
                borders: {
                  top: { style: docxLib.BorderStyle.NONE },
                  bottom: { style: docxLib.BorderStyle.NONE },
                  left: { style: docxLib.BorderStyle.NONE },
                  right: { style: docxLib.BorderStyle.NONE }
                },
                margins: { top: 55, bottom: 55, left: 65, right: 65 },
                children: [
                  new docxLib.Paragraph({
                    children: [
                      new docxLib.TextRun({
                        text: cell.label,
                        size: 11,
                        color: colors.muted,
                        font: "Arial"
                      })
                    ],
                    spacing: { after: 10 }
                  }),
                  new docxLib.Paragraph({
                    children: [
                      new docxLib.TextRun({
                        text: cell.value,
                        bold: true,
                        size: 13,
                        color: colors.black,
                        font: "Arial"
                      })
                    ],
                    spacing: { after: 0 }
                  })
                ]
              })
            )
          ]
        })
      ]
    });

    const ratingsTable = new docxLib.Table({
      width: { size: 100, type: docxLib.WidthType.PERCENTAGE },
      borders: {
        top: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        bottom: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        left: { style: docxLib.BorderStyle.NONE },
        right: { style: docxLib.BorderStyle.NONE },
        insideHorizontal: { color: colors.line, style: docxLib.BorderStyle.SINGLE, size: 4 },
        insideVertical: { style: docxLib.BorderStyle.NONE }
      },
      rows: [
        new docxLib.TableRow({
          children: [
            "Agency",
            "Short-Term Rating",
            "Long-Term Rating"
          ].map((cell) =>
            new docxLib.TableCell({
              shading: { fill: colors.soft },
              margins: { top: 55, bottom: 55, left: 55, right: 55 },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              children: [
                new docxLib.Paragraph({
                  alignment: docxLib.AlignmentType.CENTER,
                  children: [
                    new docxLib.TextRun({
                      text: cell,
                      bold: true,
                      size: 11,
                      color: colors.black,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 0 }
                })
              ]
            })
          )
        }),
        ...((ratings.length ? ratings : [{ agency: "N/A", shortTerm: "N/A", longTerm: "N/A" }]).map((cell) =>
          new docxLib.TableRow({
            children: [
              cell.agency || "N/A",
              cell.shortTerm || "N/A",
              cell.longTerm || "N/A"
            ].map((value) =>
            new docxLib.TableCell({
              margins: { top: 55, bottom: 55, left: 55, right: 55 },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              children: [
                new docxLib.Paragraph({
                  alignment: docxLib.AlignmentType.CENTER,
                  children: [
                    new docxLib.TextRun({
                      text: value,
                      size: 13,
                      color: colors.ink,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 0 }
                })
              ]
            })
            )
          })
        ))
      ]
    });

    return [
      new docxLib.Paragraph({
        children: [
          new docxLib.TextRun({
            text: heading,
            bold: true,
            size: 18,
            color: colors.red,
            font: "Arial"
          })
        ],
        spacing: { before: 60, after: 26 }
      }),
      ...(showMetadataTable
        ? [
          metadataTable,
          new docxLib.Paragraph({ spacing: { after: 24 } })
        ]
        : []),
      ratingsTable
    ];
  }

  async function buildEquityOnlyImageParagraphs(docxLib, colors, files, figureDetails = {}) {
    return buildImageParagraphs(docxLib, files, colors, 1, figureDetails);
  }

  function nomuraDisplayType(noteType) {
    return strategyLabelForNoteType(noteType) || "Research Note";
  }

  function buildFooterMetaTable(docxLib, colors, publicationDate, generatedAt, noteId) {
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
              width: { size: 35, type: docxLib.WidthType.PERCENTAGE },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              margins: { top: 0, bottom: 0, left: 0, right: 0 },
              children: [
                new docxLib.Paragraph({
                  alignment: docxLib.AlignmentType.LEFT,
                  children: [
                    new docxLib.TextRun({
                      text: `Note ID: ${noteId || "Pending"}`,
                      size: 11,
                      color: colors.muted,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 0 }
                })
              ]
            }),
            new docxLib.TableCell({
              width: { size: 41, type: docxLib.WidthType.PERCENTAGE },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              margins: { top: 0, bottom: 0, left: 0, right: 0 },
              children: [
                new docxLib.Paragraph({
                  alignment: docxLib.AlignmentType.CENTER,
                  children: [
                    new docxLib.TextRun({
                      text: `Published: ${formatProductionTimestamp(generatedAt)}`,
                      size: 11,
                      color: colors.muted,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 0 }
                })
              ]
            }),
            new docxLib.TableCell({
              width: { size: 24, type: docxLib.WidthType.PERCENTAGE },
              borders: {
                top: { style: docxLib.BorderStyle.NONE },
                bottom: { style: docxLib.BorderStyle.NONE },
                left: { style: docxLib.BorderStyle.NONE },
                right: { style: docxLib.BorderStyle.NONE }
              },
              margins: { top: 0, bottom: 0, left: 0, right: 0 },
              children: [
                new docxLib.Paragraph({
                  alignment: docxLib.AlignmentType.RIGHT,
                  children: [
                    new docxLib.TextRun({
                      children: ["Page ", docxLib.PageNumber.CURRENT, " of ", docxLib.PageNumber.TOTAL_PAGES],
                      size: 11,
                      color: colors.muted,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 0 }
                })
              ]
            })
          ]
        })
      ]
    });
  }

  function buildComplianceDeclarationSection(docxLib, colors, data, options = {}) {
    const justification = docxLib.AlignmentType.JUSTIFIED || docxLib.AlignmentType.BOTH || docxLib.AlignmentType.LEFT;
    const model = buildComplianceNoticeModel(data);
    const legalBodySize = 14;
    const legalHeadingSize = 12;
    const legalSubheadingSize = 11;

    return [
      new docxLib.Paragraph({
        pageBreakBefore: options.pageBreakBefore !== false,
        border: {
          top: {
            color: colors.line,
            style: docxLib.BorderStyle.SINGLE,
            size: 4
          }
        },
        children: [
          new docxLib.TextRun({
            text: "Regulatory Disclosure and Important Notice",
            bold: true,
            size: 18,
            color: colors.black,
            font: "Arial"
          })
        ],
        spacing: { before: 0, after: 34 }
      }),
      new docxLib.Paragraph({
        alignment: justification,
        children: [
          new docxLib.TextRun({
            text: model.introduction,
            size: legalBodySize,
            color: colors.muted,
            font: "Arial"
          })
        ],
        spacing: { after: 108, line: 228 }
      }),
      ...model.sections.flatMap((section) => ([
        new docxLib.Paragraph({
          children: [
            new docxLib.TextRun({
              text: section.heading,
              bold: true,
              size: legalHeadingSize,
              color: colors.black,
              font: "Arial"
            })
          ],
          spacing: { before: 124, after: 18 }
        }),
        new docxLib.Paragraph({
          children: [
            new docxLib.TextRun({
              text: section.subheading,
              italics: true,
              size: legalSubheadingSize,
              color: colors.muted,
              font: "Arial"
            })
          ],
          spacing: { after: 34 }
        }),
        ...section.paragraphs.map((paragraph, index) =>
          new docxLib.Paragraph({
            alignment: justification,
            children: [
              new docxLib.TextRun({
                text: paragraph,
                size: legalBodySize,
                color: colors.muted,
                font: "Arial"
              })
            ],
            spacing: { after: index === section.paragraphs.length - 1 ? 82 : 52, line: 228 },
            indent: { left: 0, right: 0 }
          })
        )
      ]))
    ];
  }

  function collectComplianceAnalystNames(data) {
    return [
      [data.authorFirstName, data.authorLastName].filter(Boolean).join(" ").trim(),
      ...(data.coAuthors || []).map((coAuthor) => [coAuthor.firstName, coAuthor.lastName].filter(Boolean).join(" ").trim())
    ].filter(Boolean);
  }

  function formatHumanList(items) {
    const cleanItems = items.map((item) => String(item || "").trim()).filter(Boolean);
    if (!cleanItems.length) return "";
    if (cleanItems.length === 1) return cleanItems[0];
    if (cleanItems.length === 2) return `${cleanItems[0]} and ${cleanItems[1]}`;
    return `${cleanItems.slice(0, -1).join(", ")}, and ${cleanItems[cleanItems.length - 1]}`;
  }

  function buildDocAuthorLine(lastName, firstName, phone) {
    const name = [firstName, lastName].filter(Boolean).join(" ").trim();
    const phoneDisplay = formatPhoneDisplay(phone);
    if (name && phoneDisplay) return `${name}, ${phoneDisplay}`;
    return name || phoneDisplay;
  }

  function formatPhoneDisplay(phone) {
    const raw = String(phone || "").trim();
    if (!raw) return "N/A";

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
    return digits ? `+${digits}` : "N/A";
  }

  function loadBrandLogoImage() {
    if (state.brandLogoImagePromise) return state.brandLogoImagePromise;

    state.brandLogoImagePromise = new Promise((resolve) => {
      const existingLogo = document.querySelector(".brand-logo");
      if (existingLogo instanceof HTMLImageElement && existingLogo.complete && existingLogo.naturalWidth > 0) {
        resolve(existingLogo);
        return;
      }

      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => resolve(null);
      image.src = "assets/cordoba-logo";
    });

    return state.brandLogoImagePromise;
  }

  async function buildNomuraBannerImageBytes(meta) {
    const canvas = document.createElement("canvas");
    const width = 1600;
    const height = 260;
    const ctx = canvas.getContext("2d");
    canvas.width = width;
    canvas.height = height;
    const logoImage = await loadBrandLogoImage();

    const publicationDate = meta.publicationDate || new Date();
    const deskLine = meta.deskLine || defaultDeskLine(meta.noteType) || "Research - Global";

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    const headerGradient = ctx.createLinearGradient(0, 0, width, 0);
    headerGradient.addColorStop(0, "#fbf7ef");
    headerGradient.addColorStop(0.46, "#ffffff");
    headerGradient.addColorStop(1, "#f7efdc");
    ctx.fillStyle = headerGradient;
    ctx.fillRect(0, 0, width, 140);

    drawBannerShape(ctx, "rgba(132,95,15,0.14)", 1118, 140, 172, 0, 182, 140);
    drawBannerShape(ctx, "rgba(132,95,15,0.20)", 1220, 140, 214, 0, 246, 140);
    drawBannerShape(ctx, "rgba(93,67,11,0.14)", 1330, 140, 220, 0, 258, 140);
    drawBannerShape(ctx, "rgba(203,166,91,0.28)", 1448, 140, 208, 0, 246, 140);

    ctx.fillStyle = "#845F0F";
    ctx.fillRect(0, 216, width, 24);

    ctx.textBaseline = "alphabetic";
    if (logoImage) {
      const maxWidth = 560;
      const maxHeight = 92;
      const scale = Math.min(maxWidth / logoImage.naturalWidth, maxHeight / logoImage.naturalHeight);
      const renderWidth = Math.round(logoImage.naturalWidth * scale);
      const renderHeight = Math.round(logoImage.naturalHeight * scale);
      ctx.drawImage(logoImage, 44, 24, renderWidth, renderHeight);
    } else {
      ctx.fillStyle = "#111111";
      ctx.font = "700 40px Arial";
      ctx.fillText("CORDOBA RESEARCH GROUP", 50, 82);
    }

    ctx.fillStyle = "#111111";
    ctx.font = "700 40px Arial";
    ctx.fillText(nomuraDisplayType(meta.noteType), 54, 188);

    ctx.fillStyle = "#845F0F";
    ctx.textAlign = "right";
    ctx.font = "600 24px Arial";
    ctx.fillText("Global Markets Research", width - 54, 164);
    ctx.font = "700 32px Arial";
    ctx.fillText(formatDocDate(publicationDate), width - 54, 200);

    ctx.textAlign = "left";
    ctx.fillStyle = "#ffffff";
    ctx.font = "600 14px Arial";
    ctx.fillText(deskLine, 56, 232);

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

  function buildNomuraTitlePanel(docxLib, colors, data, analystContacts, publicationDate) {
    const analystDesk = data.deskLine || getDeskLine();
    const showCoverageMeta = data.noteType === "Macro Research" || (isMacroFiNoteType(data.noteType) && (data.coverageCountry || data.issuerId));
    const coverageMeta = showCoverageMeta
      ? `Coverage Country: ${getCoverageCountryLabel(data.coverageCountry)} | Issuer Number: ${data.issuerId || "N/A"}`
      : "";

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
                  spacing: { after: coverageMeta ? 16 : 20 }
                }),
                ...(coverageMeta
                  ? [new docxLib.Paragraph({
                    children: [
                      new docxLib.TextRun({
                        text: coverageMeta,
                        size: 12,
                        color: colors.red,
                        font: "Arial"
                      })
                    ],
                    spacing: { after: 12 }
                  })]
                  : [])
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
                ...buildAnalystContactPanel(docxLib, colors, analystDesk, analystContacts)
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

  function buildKeyTakeawaysBox(docxLib, colors, items, heading = "Key Takeaways") {
    const cleanItems = items.length ? items : ["No key takeaways supplied."];

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
              shading: { fill: colors.takeawayFill },
              margins: { top: 108, bottom: 108, left: 135, right: 135 },
              borders: {
                top: { color: colors.takeawayBorder, style: docxLib.BorderStyle.SINGLE, size: 5 },
                bottom: { color: colors.takeawayBorder, style: docxLib.BorderStyle.SINGLE, size: 5 },
                left: { color: colors.takeawayBorder, style: docxLib.BorderStyle.SINGLE, size: 5 },
                right: { color: colors.takeawayBorder, style: docxLib.BorderStyle.SINGLE, size: 5 }
              },
              children: [
                new docxLib.Paragraph({
                  children: [
                    new docxLib.TextRun({
                      text: heading,
                      bold: true,
                      color: colors.red,
                      size: 18,
                      font: "Arial"
                    })
                  ],
                  spacing: { after: 36 }
                }),
                ...buildKeyTakeawayBoxParagraphs(docxLib, colors, cleanItems)
              ]
            })
          ]
        })
      ]
    });
  }

  function buildKeyTakeawayBoxParagraphs(docxLib, colors, items) {
    return items.map((item, index) =>
      new docxLib.Paragraph({
        indent: { left: 120, hanging: 84 },
        children: [
          new docxLib.TextRun({
            text: "• ",
            bold: true,
            font: "Arial",
            size: 18,
            color: colors.redDark
          }),
          new docxLib.TextRun({
            text: item,
            font: "Arial",
            size: 17,
            color: colors.ink
          })
        ],
        spacing: { after: index === items.length - 1 ? 0 : 28 }
      })
    );
  }

  function buildNomuraBulletParagraphs(docxLib, items, spacingAfter = 42) {
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
        spacing: { after: spacingAfter }
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
    const imageParagraphs = await buildImageParagraphs(docxLib, data.imageFiles, colors, startIndex, data.figureDetails || {});
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

  function buildCrgRatingMethodologySection(docxLib, colors) {
    const introParagraphs = [
      "CRG equity research ratings are relative opinion classifications intended to frame expected share-price performance over a 12-month horizon against the relevant sector opportunity set and the valuation case presented in the note. Target prices and valuation ranges reflect analytical judgment based on assumptions, scenario weighting, and market inputs that may change without notice.",
      "The definitions below are explanatory only. They are not statements of fact, do not constitute regulated investment advice, and should be read together with the valuation discussion, catalysts, and risk factors set out in the note."
    ];
    const justification = docxLib.AlignmentType.JUSTIFIED || docxLib.AlignmentType.BOTH || docxLib.AlignmentType.LEFT;

    return [
      new docxLib.Paragraph({
        pageBreakBefore: true,
        border: {
          top: {
            color: colors.line,
            style: docxLib.BorderStyle.SINGLE,
            size: 4
          }
        },
        children: [
          new docxLib.TextRun({
            text: "CRG Equity Research Rating Definitions",
            bold: true,
            font: "Arial",
            size: 18,
            color: colors.black
          })
        ],
        spacing: { before: 82, after: 30 }
      }),
      ...introParagraphs.map((paragraph, index) =>
        new docxLib.Paragraph({
          children: [
            new docxLib.TextRun({
              text: paragraph,
              font: "Arial",
              size: 14,
              color: colors.muted
            })
          ],
          spacing: { after: index === introParagraphs.length - 1 ? 46 : 34, line: 228 },
          alignment: justification
        })
      ),
      new docxLib.Paragraph({
        children: [
          new docxLib.TextRun({
            text: "STOCKS",
            bold: true,
            font: "Arial",
            size: 12,
            color: colors.black
          })
        ],
        spacing: { before: 18, after: 26 }
      }),
      ...getCrgRatingDefinitionRows().map((row) =>
        new docxLib.Paragraph({
          children: [
            new docxLib.TextRun({
              text: `${row[0]}: `,
              bold: true,
              font: "Arial",
              size: 14,
              color: colors.black
            }),
            new docxLib.TextRun({
              text: row[1],
              font: "Arial",
              size: 14,
              color: colors.muted
            })
          ],
          spacing: { after: 24, line: 228 },
          alignment: justification
        })
      ),
      new docxLib.Paragraph({
        children: [
          new docxLib.TextRun({
            text: "Benchmark interpretation: ",
            bold: true,
            font: "Arial",
            size: 14,
            color: colors.black
          }),
          new docxLib.TextRun({
            text: "Unless stated otherwise in the note, rating language should be read in the context of 12-month total return potential relative to sector alternatives, the market-implied base case, and the benchmark context identified in the valuation discussion.",
            font: "Arial",
            size: 14,
            color: colors.muted
          })
        ],
        spacing: { before: 18, after: 60, line: 228 },
        alignment: justification
      })
    ];
  }

  async function buildImageParagraphs(docxLib, files, colors, startIndex = 1, figureDetails = {}) {
    const output = [];

    for (let index = 0; index < files.length; index += 1) {
      const file = files[index];
      const buffer = await file.arrayBuffer();
      const size = await getImageFit(file, 560, 320);
      const figureKey = figureFileKey(file);
      const detail = figureDetails[figureKey] || {};
      const labelType = detail.labelType || defaultFigureLabelType(file);
      const labelNumber = sanitizeFigureNumber(detail.labelNumber, startIndex + index);
      const caption = detail.caption || defaultFigureCaption(file);

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
              text: `${labelType} ${labelNumber}. ${caption}`,
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
    return formatPublishedTimestampForDocument(d);
  }

  function formatLocalTimestamp(date) {
    const d = date instanceof Date ? date : new Date(date);
    const parts = new Intl.DateTimeFormat("en-GB", {
      day: "numeric",
      month: "long",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
      timeZoneName: "short"
    }).formatToParts(d);

    const lookup = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    const zone = lookup.timeZoneName || Intl.DateTimeFormat().resolvedOptions().timeZone || "";
    return `${lookup.day} ${lookup.month} ${lookup.year}, ${lookup.hour}:${lookup.minute} ${zone}`.replace(/\s+/g, " ").trim();
  }

  function formatPublishedTimestampForDocument(date) {
    const d = date instanceof Date ? date : new Date(date);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    const hours = String(d.getHours()).padStart(2, "0");
    const minutes = String(d.getMinutes()).padStart(2, "0");
    const seconds = String(d.getSeconds()).padStart(2, "0");
    return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}`;
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
