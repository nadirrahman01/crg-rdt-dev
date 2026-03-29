(() => {
  "use strict";

  const DRAFT_KEY = "crg_rdt_v5_draft";
  const HOUSE_DISCLOSURE =
    "For professional and institutional audiences only. Not for retail distribution. All outputs are draft research publications and require analyst and compliance review before circulation.";

  const WORDMARK_SVG = `
    <svg xmlns="http://www.w3.org/2000/svg" width="920" height="210" viewBox="0 0 920 210" role="img" aria-labelledby="title desc">
      <title id="title">Cordoba Research Group</title>
      <desc id="desc">Cordoba Research Group wordmark.</desc>
      <rect width="920" height="210" fill="none"/>
      <g fill="#112B46">
        <text x="0" y="82" font-family="Georgia, 'Times New Roman', serif" font-size="62" font-weight="700" letter-spacing="1.2">Cordoba</text>
        <text x="2" y="142" font-family="Arial, Helvetica, sans-serif" font-size="31" font-weight="700" letter-spacing="7">RESEARCH GROUP</text>
      </g>
      <g fill="#B38633">
        <rect x="610" y="38" width="250" height="6" rx="3"/>
        <rect x="610" y="100" width="250" height="6" rx="3"/>
        <circle cx="886" cy="103" r="10"/>
      </g>
    </svg>
  `.trim();

  const WORDMARK_DATA_URL = `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(WORDMARK_SVG)}`;

  const DRAFT_FIELDS = [
    "noteType",
    "publicationType",
    "distributionClass",
    "publicationDate",
    "title",
    "subtitle",
    "topic",
    "houseView",
    "authorName",
    "authorRole",
    "contactLine",
    "executiveSummary",
    "keyTakeaways",
    "cordobaView",
    "analysis",
    "catalysts",
    "keyRisks",
    "supportingAnalysis",
    "ticker",
    "rating",
    "currentPriceInput",
    "targetPrice",
    "valuationSummary",
    "scenarioNotes",
    "forecastTitle",
    "forecastTable",
    "modelLink",
    "customDisclaimer"
  ];

  const state = {
    chart: {
      dataUrl: "",
      svg: "",
      stats: {
        currentPrice: null,
        rangeReturn: null,
        realisedVol: null
      }
    },
    figureAssets: [],
    draftTimer: null
  };

  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => Array.from(root.querySelectorAll(selector));

  function safeTrim(value) {
    return (value ?? "").toString().trim();
  }

  function safeNum(value) {
    const parsed = Number(String(value ?? "").replace(/,/g, ""));
    return Number.isFinite(parsed) ? parsed : null;
  }

  function pct(value) {
    return Number.isFinite(value) ? `${(value * 100).toFixed(1)}%` : "—";
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

  function formatDateISO(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function formatLongDate(dateLike) {
    const date = dateLike ? new Date(dateLike) : new Date();
    if (Number.isNaN(date.getTime())) return "Undated";
    const months = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
  }

  function formatTimeStamp() {
    const now = new Date();
    const hours = String(now.getHours()).padStart(2, "0");
    const minutes = String(now.getMinutes()).padStart(2, "0");
    return `${hours}:${minutes}`;
  }

  function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  }

  function showMessage(kind, text) {
    const el = document.getElementById("message");
    if (!el) return;
    if (!text) {
      el.textContent = "";
      el.className = "message is-hidden";
      return;
    }
    el.textContent = text;
    el.className = `message ${kind}`;
  }

  function toParagraphs(text) {
    return safeTrim(text)
      .split(/\n\s*\n/g)
      .map((block) => block.trim())
      .filter(Boolean);
  }

  function toBullets(text) {
    return safeTrim(text)
      .split("\n")
      .map((line) => line.replace(/^[-*•]\s*/, "").trim())
      .filter(Boolean);
  }

  function detectDelimiter(line) {
    if (line.includes("\t")) return "\t";
    if (line.includes("|")) return "|";
    return ",";
  }

  function parseTable(text) {
    const rows = safeTrim(text)
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);

    if (!rows.length) return [];

    const delimiter = detectDelimiter(rows[0]);
    const parsed = rows.map((row) => row.split(delimiter).map((cell) => cell.trim()));
    const width = Math.max(...parsed.map((row) => row.length));

    return parsed.map((row) => {
      const next = row.slice();
      while (next.length < width) next.push("");
      return next;
    });
  }

  function tableHtml(rows) {
    if (!rows.length) {
      return `<div class="doc-body">No table provided.</div>`;
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

  function paragraphHtml(text, fallback) {
    const blocks = toParagraphs(text);
    if (!blocks.length) {
      return `<div class="doc-body">${escapeHtml(fallback || "—")}</div>`;
    }
    return blocks.map((block) => `<div class="doc-body">${escapeHtml(block)}</div>`).join("");
  }

  function bulletsHtml(text, fallback) {
    const bullets = toBullets(text);
    if (!bullets.length) {
      return `<div class="doc-body">${escapeHtml(fallback || "—")}</div>`;
    }
    return `<ul class="doc-bullet-list">${bullets.map((bullet) => `<li>${escapeHtml(bullet)}</li>`).join("")}</ul>`;
  }

  function buildReference(data) {
    const typeCodeMap = {
      "Macro Research": "MR",
      "Equity Research": "ER",
      "Fixed Income Research": "FI",
      "Commodity Insights": "CI",
      "Strategy Note": "SN"
    };

    const publicationCodeMap = {
      "In Focus": "IF",
      "Flash Note": "FL",
      "Update": "UP",
      "Initiation": "IN",
      "Results Review": "RR",
      "Thematic": "TH"
    };

    const typeCode = typeCodeMap[data.noteType] || "RN";
    const publicationCode = publicationCodeMap[data.publicationType] || "DR";
    const dateCode = safeTrim(data.publicationDate || formatDateISO(new Date())).replaceAll("-", "");
    const suffixBase = safeTrim(data.ticker) || safeTrim(data.title).replace(/[^A-Za-z0-9]+/g, " ").split(/\s+/).slice(0, 3).join("");
    const suffix = (suffixBase || "NOTE").replace(/[^A-Za-z0-9]/g, "").toUpperCase();
    return `CRG-${typeCode}-${publicationCode}-${dateCode}-${suffix}`;
  }

  function ratingDisplay(value) {
    const rating = safeTrim(value);
    if (!rating) return "—";
    return rating.toLowerCase() === "hold" ? "Neutral" : rating;
  }

  function formatNumberDisplay(raw, decimals = 2) {
    const num = safeNum(raw);
    if (num === null) return safeTrim(raw) || "—";
    return num.toFixed(decimals);
  }

  function effectiveCurrentPrice() {
    const manual = safeNum($("#currentPriceInput")?.value);
    if (manual !== null) return manual;
    return state.chart.stats.currentPrice;
  }

  function buildNoteMap(data) {
    const sections = [
      "Executive Summary",
      "Cordoba View",
      "Investment Thesis",
      "Catalysts and Risks",
      "Supporting Analysis"
    ];

    if (data.isEquity) {
      sections.push("Valuation and Scenarios");
    }

    if (data.figureAssets.length) {
      sections.push("Figures Appendix");
    }

    return sections;
  }

  function buildModel() {
    const noteType = safeTrim($("#noteType")?.value);
    const publicationType = safeTrim($("#publicationType")?.value);
    const distributionClass = safeTrim($("#distributionClass")?.value);
    const publicationDate = safeTrim($("#publicationDate")?.value) || formatDateISO(new Date());
    const title = safeTrim($("#title")?.value);
    const subtitle = safeTrim($("#subtitle")?.value);
    const topic = safeTrim($("#topic")?.value);
    const houseView = safeTrim($("#houseView")?.value);
    const authorName = safeTrim($("#authorName")?.value);
    const authorRole = safeTrim($("#authorRole")?.value);
    const contactLine = safeTrim($("#contactLine")?.value);
    const coAuthors = $$(".coauthor-input").map((input) => safeTrim(input.value)).filter(Boolean);
    const executiveSummary = $("#executiveSummary")?.value || "";
    const keyTakeaways = $("#keyTakeaways")?.value || "";
    const cordobaView = $("#cordobaView")?.value || "";
    const analysis = $("#analysis")?.value || "";
    const catalysts = $("#catalysts")?.value || "";
    const keyRisks = $("#keyRisks")?.value || "";
    const supportingAnalysis = $("#supportingAnalysis")?.value || "";
    const ticker = safeTrim($("#ticker")?.value);
    const rating = safeTrim($("#rating")?.value);
    const currentPriceRaw = safeTrim($("#currentPriceInput")?.value);
    const targetPriceRaw = safeTrim($("#targetPrice")?.value);
    const currentPrice = effectiveCurrentPrice();
    const targetPrice = safeNum(targetPriceRaw);
    const valuationSummary = $("#valuationSummary")?.value || "";
    const scenarioNotes = $("#scenarioNotes")?.value || "";
    const forecastTitle = safeTrim($("#forecastTitle")?.value) || "Forecast Snapshot";
    const forecastTableRows = parseTable($("#forecastTable")?.value || "");
    const modelLink = safeTrim($("#modelLink")?.value);
    const customDisclaimer = $("#customDisclaimer")?.value || "";
    const figureAssets = state.figureAssets.slice();
    const isEquity = noteType === "Equity Research";
    const noteReference = buildReference({
      noteType,
      publicationType,
      publicationDate,
      title,
      ticker
    });
    const upside = currentPrice !== null && targetPrice !== null && currentPrice > 0
      ? (targetPrice / currentPrice) - 1
      : null;

    return {
      noteType,
      publicationType,
      distributionClass,
      publicationDate,
      publicationDateLabel: formatLongDate(publicationDate),
      title,
      subtitle,
      topic,
      houseView,
      authorName,
      authorRole,
      contactLine,
      coAuthors,
      analystLine: [authorName, authorRole].filter(Boolean).join(" · "),
      executiveSummary,
      keyTakeaways,
      cordobaView,
      analysis,
      catalysts,
      keyRisks,
      supportingAnalysis,
      ticker,
      rating,
      ratingDisplay: ratingDisplay(rating),
      currentPriceRaw,
      currentPrice,
      currentPriceDisplay: currentPrice !== null ? currentPrice.toFixed(2) : (currentPriceRaw || "—"),
      targetPriceRaw,
      targetPrice,
      targetPriceDisplay: formatNumberDisplay(targetPriceRaw),
      upside,
      upsideDisplay: pct(upside),
      rangeReturnDisplay: pct(state.chart.stats.rangeReturn),
      volatilityDisplay: pct(state.chart.stats.realisedVol),
      chartDataUrl: state.chart.dataUrl,
      valuationSummary,
      scenarioNotes,
      forecastTitle,
      forecastTableRows,
      modelLink,
      customDisclaimer,
      figureAssets,
      isEquity,
      noteReference
    };
  }

  function buildQualityGates(model) {
    const gates = [
      {
        label: "Document control complete",
        pass: [
          model.noteType,
          model.publicationType,
          model.distributionClass,
          model.publicationDate,
          model.title,
          model.subtitle,
          model.topic,
          model.houseView
        ].every(Boolean)
      },
      {
        label: "Front-page summary drafted",
        pass: safeTrim(model.executiveSummary) && toBullets(model.keyTakeaways).length > 0
      },
      {
        label: "Analyst line defined",
        pass: !!model.authorName
      },
      {
        label: "House view and thesis written",
        pass: safeTrim(model.cordobaView) && safeTrim(model.analysis)
      },
      {
        label: "Catalysts and risks framed",
        pass: toBullets(model.catalysts).length > 0 && toBullets(model.keyRisks).length > 0
      }
    ];

    if (model.isEquity) {
      gates.push({
        label: "Equity pricing and valuation support present",
        pass:
          !!model.ticker &&
          !!model.rating &&
          safeTrim(model.currentPriceDisplay) !== "—" &&
          !!(safeTrim(model.valuationSummary) || model.forecastTableRows.length > 0 || safeTrim(model.scenarioNotes))
      });
    } else {
      gates.push({
        label: "Supporting analysis present",
        pass: !!safeTrim(model.supportingAnalysis) || model.figureAssets.length > 0
      });
    }

    return gates;
  }

  function renderQuality(model) {
    const checklist = document.getElementById("qualityChecklist");
    const summary = document.getElementById("qualitySummary");
    const gates = buildQualityGates(model);
    const passed = gates.filter((gate) => gate.pass).length;

    if (summary) {
      summary.textContent = `${passed} / ${gates.length} gates closed`;
    }

    if (checklist) {
      checklist.innerHTML = gates
        .map((gate) => `<li class="${gate.pass ? "is-pass" : "is-fail"}">${escapeHtml(gate.label)}</li>`)
        .join("");
    }

    let label = "Draft";
    if (passed === gates.length) label = "Ready to Export";
    else if (passed >= Math.ceil(gates.length * 0.7)) label = "Editor Review";
    else if (passed >= Math.ceil(gates.length * 0.5)) label = "In Progress";

    setText("readinessStatus", label);
  }

  function renderMetricStrip(model) {
    setText("currentPriceDisplay", model.currentPriceDisplay || "—");
    setText("upsideDisplay", model.upsideDisplay);
    setText("rangeReturnDisplay", model.rangeReturnDisplay);
    setText("volatilityDisplay", model.volatilityDisplay);
  }

  function coverPageHtml(model) {
    return `
      <article class="preview-page">
        <div class="doc-cover">
          <div class="doc-top-legal">${escapeHtml(model.distributionClass || "Institutional Only")} · ${escapeHtml(HOUSE_DISCLOSURE)}</div>
          <div class="doc-cover-body">
            <div class="doc-sideband">
              <div class="doc-sideband-rule"></div>
              <div class="doc-sideband-label">${escapeHtml(model.publicationType || "Research Note")}</div>
            </div>
            <div class="doc-cover-main">
              <div class="doc-brand-row">
                <img class="doc-wordmark" src="${WORDMARK_DATA_URL}" alt="Cordoba Research Group" />
                <div class="doc-cover-date">${escapeHtml(model.publicationDateLabel)}</div>
              </div>

              <div class="doc-cover-kicker">${escapeHtml(model.noteType || "Research Note")} · ${escapeHtml(model.publicationType || "Draft")}</div>
              <div class="doc-cover-title">${escapeHtml(model.title || "Untitled Research Note")}</div>
              <div class="doc-cover-subtitle">${escapeHtml(model.subtitle || model.houseView || "Add a subtitle to frame the note.")}</div>

              <div class="doc-cover-meta">
                <div class="doc-meta-block">
                  <span>Coverage</span>
                  <strong>${escapeHtml(model.topic || "—")}</strong>
                </div>
                <div class="doc-meta-block">
                  <span>Lead Analyst</span>
                  <strong>${escapeHtml(model.analystLine || "—")}</strong>
                </div>
                <div class="doc-meta-block">
                  <span>Reference</span>
                  <strong>${escapeHtml(model.noteReference)}</strong>
                </div>
              </div>
            </div>
          </div>
        </div>
      </article>
    `;
  }

  function summarySideCards(model) {
    const noteMap = buildNoteMap(model);

    const equityCard = model.isEquity
      ? `
        <div class="doc-sidecard">
          <div class="doc-mini-label">Equity Snapshot</div>
          <strong>${escapeHtml(model.ticker || "—")} · ${escapeHtml(model.ratingDisplay)}</strong>
          <p>Current ${escapeHtml(model.currentPriceDisplay)} · Target ${escapeHtml(model.targetPriceDisplay)} · ${escapeHtml(model.upsideDisplay)} upside/downside.</p>
        </div>
        <div class="doc-sidecard">
          <div class="doc-mini-label">Price Chart</div>
          ${model.chartDataUrl
            ? `<div class="doc-chart-frame"><img src="${model.chartDataUrl}" alt="${escapeHtml(model.ticker || "Price chart")}" /></div>`
            : `<div class="doc-chart-frame"><div class="doc-chart-empty">No price chart loaded.</div></div>`}
        </div>
      `
      : `
        <div class="doc-sidecard">
          <div class="doc-mini-label">Publication Focus</div>
          <strong>${escapeHtml(model.topic || "—")}</strong>
          <p>${escapeHtml(model.houseView || "Add a house view to sharpen the front page.")}</p>
        </div>
      `;

    return `
      <div class="doc-sidecard">
        <div class="doc-mini-label">Cordoba View</div>
        <strong>${escapeHtml(model.houseView || "—")}</strong>
        <p>${escapeHtml(model.contactLine || "Add a contact line if the note needs one.")}</p>
      </div>
      <div class="doc-sidecard">
        <div class="doc-mini-label">Document Map</div>
        <ul class="doc-note-map">${noteMap.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
      </div>
      ${equityCard}
    `;
  }

  function summaryPageHtml(model) {
    return `
      <article class="preview-page">
        <div class="doc-inner-page">
          <div class="doc-page-head">
            <span>${escapeHtml(model.publicationType || "Research Note")}</span>
            <span>${escapeHtml(model.noteReference)}</span>
          </div>

          <div class="doc-summary-grid">
            <div>
              <section class="doc-section-block">
                <div class="doc-section-number">Executive Summary</div>
                <h3 class="doc-section-title">What matters on page one.</h3>
                ${paragraphHtml(model.executiveSummary, "Add an executive summary to define the front page.")}
              </section>

              <section class="doc-section-block">
                <div class="doc-section-number">Key Takeaways</div>
                <h3 class="doc-section-title">Key takeaways</h3>
                ${bulletsHtml(model.keyTakeaways, "Add investor-facing takeaways.")}
              </section>
            </div>

            <div>
              ${summarySideCards(model)}
            </div>
          </div>

          <div class="doc-footer">
            <div class="doc-footer-text">${escapeHtml(HOUSE_DISCLOSURE)}</div>
          </div>
        </div>
      </article>
    `;
  }

  function numberedSection(index, title, bodyHtml) {
    return `
      <section class="doc-numbered-section">
        <div class="doc-numbered-index">${String(index).padStart(2, "0")}</div>
        <div>
          <h3 class="doc-section-title">${escapeHtml(title)}</h3>
          ${bodyHtml}
        </div>
      </section>
    `;
  }

  function analysisPageHtml(model) {
    const sections = [
      numberedSection(1, "Cordoba View", paragraphHtml(model.cordobaView, "Add the Cordoba view.")),
      numberedSection(2, "Investment Thesis", paragraphHtml(model.analysis, "Add the main thesis.")),
      numberedSection(3, "Catalysts / Why Now", bulletsHtml(model.catalysts, "Add catalysts and timing triggers.")),
      numberedSection(4, "Key Risks", bulletsHtml(model.keyRisks, "Add risks and invalidation points.")),
      numberedSection(5, "Supporting Analysis", paragraphHtml(model.supportingAnalysis, "Add supporting analysis, evidence, or framework detail."))
    ];

    if (model.isEquity) {
      const valuationBody = `
        ${paragraphHtml(model.valuationSummary, "Add valuation support.") }
        <div class="doc-chart-note">${escapeHtml(`Current ${model.currentPriceDisplay} · Target ${model.targetPriceDisplay} · ${model.upsideDisplay}`)}</div>
        ${model.chartDataUrl
          ? `<div class="doc-chart-frame"><img src="${model.chartDataUrl}" alt="${escapeHtml(model.ticker || "Price chart")}" /></div>`
          : `<div class="doc-chart-frame"><div class="doc-chart-empty">Fetch a chart or enter price data to complete the equity section.</div></div>`}
        ${paragraphHtml(model.scenarioNotes, "Add bull/base/bear or sensitivity notes.") }
        ${model.forecastTableRows.length ? tableHtml(model.forecastTableRows) : ""}
      `;
      sections.push(numberedSection(6, model.forecastTitle || "Valuation and Scenarios", valuationBody));
    }

    return `
      <article class="preview-page">
        <div class="doc-inner-page">
          <div class="doc-page-head">
            <span>${escapeHtml(model.noteType || "Research Note")}</span>
            <span>${escapeHtml(model.noteReference)}</span>
          </div>
          ${sections.join("")}
          <div class="doc-footer">
            <div class="doc-footer-text">${escapeHtml(HOUSE_DISCLOSURE)}</div>
          </div>
        </div>
      </article>
    `;
  }

  function appendixPageHtml(model) {
    const fileListHtml = model.figureAssets.length
      ? `
        <div class="doc-appendix-grid">
          ${model.figureAssets.map((figure, index) => `
            <div class="doc-figure-card">
              <img src="${figure.dataUrl}" alt="${escapeHtml(figure.name)}" />
              <div class="doc-figure-caption">Figure ${index + 1}: ${escapeHtml(figure.name)}</div>
            </div>
          `).join("")}
        </div>
      `
      : `
        <div class="doc-file-list">
          <div class="doc-file-item">No figures uploaded.</div>
        </div>
      `;

    const modelLinkHtml = model.modelLink
      ? `<div class="doc-file-item">Model link: ${escapeHtml(model.modelLink)}</div>`
      : "";

    return `
      <article class="preview-page">
        <div class="doc-inner-page">
          <div class="doc-page-head">
            <span>Appendix</span>
            <span>${escapeHtml(model.noteReference)}</span>
          </div>

          <section class="doc-section-block">
            <div class="doc-section-number">Appendix</div>
            <h3 class="doc-section-title">Figures and disclosures</h3>
            ${fileListHtml}
          </section>

          ${(model.modelLink || safeTrim(model.customDisclaimer))
            ? `
              <section class="doc-section-block">
                <div class="doc-section-number">Disclosures</div>
                <h3 class="doc-section-title">Additional note disclosures</h3>
                <div class="doc-file-list">
                  ${modelLinkHtml}
                  <div class="doc-file-item">${escapeHtml(safeTrim(model.customDisclaimer) || "No additional disclosure supplied.")}</div>
                </div>
              </section>
            `
            : ""}

          <div class="doc-footer">
            <div class="doc-footer-text">${escapeHtml([HOUSE_DISCLOSURE, safeTrim(model.customDisclaimer)].filter(Boolean).join(" "))}</div>
          </div>
        </div>
      </article>
    `;
  }

  function renderPreview() {
    const model = buildModel();
    renderMetricStrip(model);
    renderQuality(model);

    setText("noteReferenceTop", model.noteReference);
    setText("previewReference", model.noteReference);
    setText("previewDistribution", model.distributionClass || "—");
    setText("previewPublicationType", [model.noteType, model.publicationType].filter(Boolean).join(" · ") || "Draft");

    const preview = document.getElementById("previewPages");
    if (!preview) return;

    const pages = [
      coverPageHtml(model),
      summaryPageHtml(model),
      analysisPageHtml(model)
    ];

    if (model.figureAssets.length || model.modelLink || safeTrim(model.customDisclaimer)) {
      pages.push(appendixPageHtml(model));
    }

    preview.innerHTML = pages.join("");
  }

  function setDraftStatus(text) {
    setText("draftStatus", text);
  }

  function saveDraftSnapshot() {
    const snapshot = {};
    DRAFT_FIELDS.forEach((id) => {
      const el = document.getElementById(id);
      if (el) snapshot[id] = el.value ?? "";
    });
    snapshot.__coAuthors = $$(".coauthor-input").map((input) => input.value ?? "");
    snapshot.__chart = state.chart;
    try {
      localStorage.setItem(DRAFT_KEY, JSON.stringify(snapshot));
      setDraftStatus(`Saved ${formatTimeStamp()}`);
    } catch {
      setDraftStatus("Local save failed");
    }
  }

  function scheduleDraftSave() {
    setDraftStatus("Saving…");
    clearTimeout(state.draftTimer);
    state.draftTimer = setTimeout(saveDraftSnapshot, 280);
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

  function syncEquitySection() {
    const isEquity = safeTrim($("#noteType")?.value) === "Equity Research";
    const section = document.getElementById("sectionEquity");
    if (section) {
      section.classList.toggle("is-hidden", !isEquity);
    }
  }

  function createCoAuthorRow(value = "") {
    const row = document.createElement("div");
    row.className = "coauthor-row";
    row.innerHTML = `
      <input type="text" class="coauthor-input" placeholder="Co-author name" value="${escapeHtml(value)}" />
      <button type="button" class="ghost-btn remove-coauthor">Remove</button>
    `;
    return row;
  }

  function attachCoAuthorRow(row) {
    const remove = $(".remove-coauthor", row);
    const input = $(".coauthor-input", row);

    if (remove) {
      remove.addEventListener("click", () => {
        row.remove();
        scheduleDraftSave();
        renderPreview();
      });
    }

    if (input) {
      ["input", "change"].forEach((eventName) => {
        input.addEventListener(eventName, () => {
          scheduleDraftSave();
          renderPreview();
        });
      });
    }
  }

  function restoreDraft(snapshot) {
    if (!snapshot) return;

    DRAFT_FIELDS.forEach((id) => {
      const el = document.getElementById(id);
      if (el && typeof snapshot[id] === "string") {
        el.value = snapshot[id];
      }
    });

    if (Array.isArray(snapshot.__coAuthors)) {
      const list = document.getElementById("coAuthorsList");
      if (list) {
        list.innerHTML = "";
        snapshot.__coAuthors.forEach((name) => {
          const row = createCoAuthorRow(name);
          attachCoAuthorRow(row);
          list.appendChild(row);
        });
      }
    }

    if (snapshot.__chart && typeof snapshot.__chart === "object") {
      state.chart = snapshot.__chart;
      renderChartPreview();
      setText("chartStatus", state.chart.dataUrl ? "Draft chart restored." : "No chart loaded.");
    }

    setDraftStatus("Draft restored");
  }

  function renderFigureList() {
    const figureList = document.getElementById("figureList");
    if (!figureList) return;

    if (!state.figureAssets.length) {
      figureList.innerHTML = "";
      return;
    }

    figureList.innerHTML = state.figureAssets
      .map((figure, index) => `<div class="file-item">Figure ${index + 1}: ${escapeHtml(figure.name)}</div>`)
      .join("");
  }

  async function readFilesAsDataUrls(fileList) {
    const files = Array.from(fileList || []);
    const output = [];

    for (const file of files) {
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
      output.push({ name: file.name, dataUrl });
    }

    return output;
  }

  function computeDailyReturns(closes) {
    const returns = [];
    for (let i = 1; i < closes.length; i += 1) {
      const prev = closes[i - 1];
      const current = closes[i];
      if (prev > 0 && Number.isFinite(prev) && Number.isFinite(current)) {
        returns.push((current / prev) - 1);
      }
    }
    return returns;
  }

  function stddev(values) {
    if (!values.length) return null;
    const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
    const variance = values.reduce((sum, value) => sum + ((value - mean) ** 2), 0) / (values.length - 1 || 1);
    return Math.sqrt(variance);
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

    const raw = await res.text();
    const csv = extractStooqCSV(raw) || raw;
    const lines = csv.trim().split("\n");
    if (lines.length < 10) throw new Error("Not enough price history returned.");

    return lines.slice(1)
      .map((line) => line.split(","))
      .map((row) => ({ date: row[0], close: Number(row[4]) }))
      .filter((point) => point.date && Number.isFinite(point.close));
  }

  function buildChartDataUrl(points, title) {
    const width = 700;
    const height = 260;
    const padLeft = 28;
    const padRight = 18;
    const padTop = 18;
    const padBottom = 30;
    const values = points.map((point) => point.close);
    const min = Math.min(...values);
    const max = Math.max(...values);
    const range = max - min || 1;
    const innerWidth = width - padLeft - padRight;
    const innerHeight = height - padTop - padBottom;

    const polyline = points.map((point, index) => {
      const x = padLeft + (index / Math.max(points.length - 1, 1)) * innerWidth;
      const y = padTop + ((max - point.close) / range) * innerHeight;
      return `${x.toFixed(2)},${y.toFixed(2)}`;
    }).join(" ");

    const area = `${padLeft},${height - padBottom} ${polyline} ${width - padRight},${height - padBottom}`;

    const guides = [0.25, 0.5, 0.75].map((ratio) => {
      const y = padTop + innerHeight * ratio;
      return `<line x1="${padLeft}" y1="${y}" x2="${width - padRight}" y2="${y}" stroke="#e2e8ef" stroke-width="1" />`;
    }).join("");

    const svg = `
      <svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(title)}">
        <rect width="${width}" height="${height}" rx="18" fill="#ffffff" />
        ${guides}
        <path d="M ${polyline}" fill="none" stroke="#173652" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" />
        <polygon points="${area}" fill="rgba(177, 135, 58, 0.12)" />
        <text x="${padLeft}" y="${padTop - 2}" fill="#6b7b8d" font-family="Arial, Helvetica, sans-serif" font-size="12" font-weight="700">${escapeHtml(title)}</text>
        <text x="${padLeft}" y="${height - 8}" fill="#8a97a6" font-family="Arial, Helvetica, sans-serif" font-size="11">${escapeHtml(points[0].date)}</text>
        <text x="${width - padRight}" y="${height - 8}" fill="#8a97a6" font-family="Arial, Helvetica, sans-serif" font-size="11" text-anchor="end">${escapeHtml(points[points.length - 1].date)}</text>
      </svg>
    `.trim();

    return {
      svg,
      dataUrl: `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(svg)}`
    };
  }

  function renderChartPreview() {
    const chartPreview = document.getElementById("chartPreview");
    if (!chartPreview) return;

    if (!state.chart.dataUrl) {
      chartPreview.className = "chart-preview chart-preview--empty";
      chartPreview.textContent = "Price chart preview will appear here.";
      return;
    }

    chartPreview.className = "chart-preview";
    chartPreview.innerHTML = `<img src="${state.chart.dataUrl}" alt="Price chart preview" />`;
  }

  async function buildPriceChart() {
    try {
      const ticker = safeTrim($("#ticker")?.value);
      if (!ticker) throw new Error("Enter a ticker first.");

      const symbol = stooqSymbolFromTicker(ticker);
      if (!symbol) throw new Error("Invalid ticker.");

      setText("chartStatus", "Fetching price history…");

      const range = $("#chartRange")?.value || "1y";
      const series = await fetchStooqDaily(symbol);
      const startDate = computeStartDate(range);
      const filtered = series.filter((point) => new Date(point.date) >= startDate);
      if (filtered.length < 10) throw new Error("Not enough data for the selected chart range.");

      const chart = buildChartDataUrl(filtered, `${ticker.toUpperCase()} Price`);
      const closes = filtered.map((point) => point.close);
      const currentPrice = closes[closes.length - 1];
      const rangeReturn = currentPrice && closes[0] ? (currentPrice / closes[0]) - 1 : null;
      const realisedVol = (() => {
        const dailyVol = stddev(computeDailyReturns(closes));
        return dailyVol === null ? null : dailyVol * Math.sqrt(252);
      })();

      state.chart = {
        dataUrl: chart.dataUrl,
        svg: chart.svg,
        stats: {
          currentPrice,
          rangeReturn,
          realisedVol
        }
      };

      const currentPriceInput = document.getElementById("currentPriceInput");
      if (currentPriceInput) {
        currentPriceInput.value = currentPrice.toFixed(2);
      }

      renderChartPreview();
      setText("chartStatus", `Chart ready (${range.toUpperCase()})`);
      scheduleDraftSave();
      renderPreview();
    } catch (error) {
      state.chart = {
        dataUrl: "",
        svg: "",
        stats: {
          currentPrice: null,
          rangeReturn: null,
          realisedVol: null
        }
      };
      renderChartPreview();
      setText("chartStatus", `Chart unavailable: ${error.message}`);
      renderPreview();
    }
  }

  function validateCore(model) {
    const missing = [];
    const coreIds = [
      ["noteType", model.noteType],
      ["publicationType", model.publicationType],
      ["distributionClass", model.distributionClass],
      ["publicationDate", model.publicationDate],
      ["title", model.title],
      ["subtitle", model.subtitle],
      ["topic", model.topic],
      ["houseView", model.houseView],
      ["authorName", model.authorName],
      ["executiveSummary", safeTrim(model.executiveSummary)],
      ["keyTakeaways", toBullets(model.keyTakeaways).length],
      ["cordobaView", safeTrim(model.cordobaView)],
      ["analysis", safeTrim(model.analysis)],
      ["catalysts", toBullets(model.catalysts).length],
      ["keyRisks", toBullets(model.keyRisks).length]
    ];

    coreIds.forEach(([id, value]) => {
      if (!value) missing.push(id);
    });

    if (model.isEquity) {
      if (!model.ticker) missing.push("ticker");
      if (!model.rating) missing.push("rating");
      if (!safeTrim(model.currentPriceDisplay) || model.currentPriceDisplay === "—") missing.push("currentPriceInput");
      if (!(safeTrim(model.valuationSummary) || model.forecastTableRows.length || safeTrim(model.scenarioNotes))) {
        missing.push("valuationSummary");
      }
    }

    return missing;
  }

  const EXPORT_STYLES = `
    body {
      margin: 0;
      color: #152638;
      font-family: Arial, Helvetica, sans-serif;
      background: #ffffff;
    }
    .preview-page {
      width: 100%;
      min-height: 980px;
      page-break-after: always;
      border: 1px solid #d7dee5;
      box-sizing: border-box;
      background: #ffffff;
    }
    .preview-page:last-child {
      page-break-after: auto;
    }
    .doc-cover {
      min-height: 980px;
      display: flex;
      flex-direction: column;
    }
    .doc-top-legal {
      padding: 14px 34px;
      background: #10283f;
      color: #ffffff;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.08em;
      text-transform: uppercase;
    }
    .doc-cover-body {
      display: table;
      width: 100%;
      min-height: 940px;
    }
    .doc-sideband,
    .doc-cover-main {
      display: table-cell;
      vertical-align: top;
    }
    .doc-sideband {
      width: 110px;
      padding: 32px 18px;
      background: #f5f2eb;
      border-right: 1px solid #e3e6ea;
    }
    .doc-sideband-rule {
      width: 36px;
      height: 2px;
      margin-bottom: 18px;
      background: #b1873a;
    }
    .doc-sideband-label {
      color: #254b72;
      font-size: 11px;
      font-weight: 800;
      letter-spacing: 0.18em;
      text-transform: uppercase;
    }
    .doc-cover-main {
      padding: 44px 50px 56px;
    }
    .doc-brand-row {
      display: table;
      width: 100%;
    }
    .doc-brand-row > div,
    .doc-brand-row > img {
      display: table-cell;
      vertical-align: top;
    }
    .doc-wordmark {
      width: 250px;
    }
    .doc-cover-date {
      color: #10283f;
      font-size: 13px;
      font-weight: 700;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      text-align: right;
    }
    .doc-cover-kicker,
    .doc-page-head {
      color: #254b72;
      font-size: 11px;
      font-weight: 800;
      letter-spacing: 0.16em;
      text-transform: uppercase;
    }
    .doc-cover-kicker {
      margin-top: 64px;
    }
    .doc-cover-title,
    .doc-section-title {
      color: #10283f;
      font-family: Georgia, "Times New Roman", serif;
    }
    .doc-cover-title {
      max-width: 560px;
      margin: 16px 0;
      font-size: 52px;
      line-height: 1;
    }
    .doc-cover-subtitle,
    .doc-body,
    .doc-bullet-list li,
    .doc-note-map li,
    .doc-file-item,
    .doc-chart-note,
    .doc-figure-caption {
      color: #263647;
      font-size: 13px;
      line-height: 1.7;
    }
    .doc-cover-subtitle {
      max-width: 560px;
      font-size: 18px;
    }
    .doc-cover-meta {
      margin-top: 40px;
      padding-top: 24px;
      border-top: 1px solid #e2e7ec;
    }
    .doc-meta-block {
      display: inline-block;
      width: 32%;
      vertical-align: top;
    }
    .doc-meta-block span,
    .doc-mini-label,
    .doc-footer-text {
      display: block;
      color: #8592a3;
      font-size: 11px;
      font-weight: 800;
      letter-spacing: 0.12em;
      text-transform: uppercase;
    }
    .doc-meta-block strong,
    .doc-sidecard strong {
      display: block;
      margin-top: 8px;
      color: #10283f;
      font-size: 16px;
      line-height: 1.4;
    }
    .doc-inner-page {
      min-height: 980px;
      padding: 34px 38px;
      box-sizing: border-box;
    }
    .doc-page-head {
      padding-bottom: 14px;
      border-bottom: 1px solid #dde3ea;
    }
    .doc-page-head span:last-child {
      float: right;
    }
    .doc-summary-grid {
      display: table;
      width: 100%;
      margin-top: 24px;
    }
    .doc-summary-grid > div {
      display: table-cell;
      vertical-align: top;
    }
    .doc-summary-grid > div:first-child {
      width: 68%;
      padding-right: 24px;
    }
    .doc-summary-grid > div:last-child {
      width: 32%;
    }
    .doc-section-number {
      color: #b1873a;
      font-size: 11px;
      font-weight: 800;
      letter-spacing: 0.16em;
      text-transform: uppercase;
    }
    .doc-section-title {
      margin: 8px 0 12px;
      font-size: 26px;
      line-height: 1.08;
    }
    .doc-sidecard {
      margin-bottom: 14px;
      padding: 16px;
      border: 1px solid #dde3ea;
      border-radius: 14px;
      background: #f7f4ee;
    }
    .doc-sidecard p {
      margin: 10px 0 0;
      color: #627387;
      font-size: 12px;
      line-height: 1.6;
    }
    .doc-bullet-list,
    .doc-note-map {
      margin: 0;
      padding-left: 18px;
    }
    .doc-bullet-list li,
    .doc-note-map li {
      margin-bottom: 8px;
    }
    .doc-numbered-section {
      display: table;
      width: 100%;
      padding: 18px 0;
      border-top: 1px solid #e2e7ec;
    }
    .doc-numbered-index,
    .doc-numbered-section > div:last-child {
      display: table-cell;
      vertical-align: top;
    }
    .doc-numbered-index {
      width: 42px;
      color: #b1873a;
      font-size: 22px;
      font-family: Georgia, "Times New Roman", serif;
    }
    .doc-chart-frame {
      margin-top: 12px;
      padding: 12px;
      border: 1px solid #dde3ea;
      border-radius: 12px;
      background: #ffffff;
    }
    .doc-chart-frame img {
      width: 100%;
    }
    .doc-chart-empty {
      min-height: 160px;
      color: #8592a3;
      text-align: center;
      line-height: 160px;
    }
    .doc-table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 12px;
    }
    .doc-table th,
    .doc-table td {
      border: 1px solid #dde3ea;
      padding: 9px 10px;
      text-align: left;
    }
    .doc-table th {
      background: #f5f2eb;
      color: #10283f;
      font-size: 11px;
      font-weight: 800;
      letter-spacing: 0.12em;
      text-transform: uppercase;
    }
    .doc-appendix-grid {
      margin-top: 16px;
    }
    .doc-figure-card {
      margin-bottom: 16px;
      border: 1px solid #dde3ea;
      border-radius: 12px;
      overflow: hidden;
      background: #ffffff;
    }
    .doc-figure-card img {
      width: 100%;
      display: block;
    }
    .doc-figure-caption {
      padding: 12px 14px;
    }
    .doc-file-list {
      margin-top: 12px;
    }
    .doc-file-item {
      padding: 10px 12px;
      border: 1px solid #dde3ea;
      border-radius: 12px;
      background: #f8fafb;
      margin-bottom: 8px;
    }
    .doc-footer {
      margin-top: 30px;
      padding-top: 18px;
      border-top: 1px solid #dde3ea;
    }
    .doc-footer-text {
      letter-spacing: 0.08em;
      line-height: 1.6;
    }
  `;

  function buildExportHtml(model) {
    const pages = [
      coverPageHtml(model),
      summaryPageHtml(model),
      analysisPageHtml(model)
    ];

    if (model.figureAssets.length || model.modelLink || safeTrim(model.customDisclaimer)) {
      pages.push(appendixPageHtml(model));
    }

    const pagesHtml = pages.join("");

    return `
      <!DOCTYPE html>
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8" />
        <title>${escapeHtml(model.title || "research_note")}</title>
        <!--[if gte mso 9]>
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>90</w:Zoom>
          </w:WordDocument>
        </xml>
        <![endif]-->
        <style>${EXPORT_STYLES}</style>
      </head>
      <body>${pagesHtml}</body>
      </html>
    `;
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function exportWordDocument() {
    const model = buildModel();
    const missing = validateCore(model);
    if (missing.length) {
      showMessage("error", `Complete the required fields before export: ${missing.join(", ")}`);
      const first = document.getElementById(missing[0]);
      if (first && typeof first.focus === "function") first.focus();
      return;
    }

    const button = document.getElementById("exportDocBtn");
    const topButton = document.getElementById("exportDocBtnTop");
    [button, topButton].forEach((node) => node && node.classList.add("loading"));

    try {
      const html = buildExportHtml(model);
      const blob = new Blob(["\ufeff", html], { type: "application/msword" });
      const safeName = (model.title || "research_note").replace(/[^a-z0-9]+/gi, "_").replace(/^_+|_+$/g, "").toLowerCase();
      downloadBlob(blob, `${safeName || "research_note"}_${model.noteReference.toLowerCase()}.doc`);
      saveDraftSnapshot();
      showMessage("success", "Word file exported successfully.");
    } catch (error) {
      showMessage("error", `Export failed: ${error.message}`);
    } finally {
      [button, topButton].forEach((node) => node && node.classList.remove("loading"));
    }
  }

  function bindStaticEvents() {
    $$(".rail-link").forEach((button) => {
      button.addEventListener("click", () => {
        const targetId = button.getAttribute("data-scroll-target");
        if (!targetId) return;
        const target = document.getElementById(targetId);
        if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    });

    const addCoAuthorBtn = document.getElementById("addCoAuthor");
    const coAuthorsList = document.getElementById("coAuthorsList");
    if (addCoAuthorBtn && coAuthorsList) {
      addCoAuthorBtn.addEventListener("click", () => {
        const row = createCoAuthorRow();
        attachCoAuthorRow(row);
        coAuthorsList.appendChild(row);
        scheduleDraftSave();
        renderPreview();
      });
    }

    const figureUpload = document.getElementById("figureUpload");
    if (figureUpload) {
      figureUpload.addEventListener("change", async () => {
        state.figureAssets = await readFilesAsDataUrls(figureUpload.files);
        renderFigureList();
        scheduleDraftSave();
        renderPreview();
      });
    }

    const fetchPriceChart = document.getElementById("fetchPriceChart");
    if (fetchPriceChart) {
      fetchPriceChart.addEventListener("click", buildPriceChart);
    }

    const saveDraftBtn = document.getElementById("saveDraftBtn");
    if (saveDraftBtn) {
      saveDraftBtn.addEventListener("click", () => {
        saveDraftSnapshot();
        showMessage("success", "Draft saved locally.");
      });
    }

    const refreshPreviewBtn = document.getElementById("previewRefreshBtn");
    if (refreshPreviewBtn) {
      refreshPreviewBtn.addEventListener("click", () => {
        renderPreview();
        showMessage("", "");
      });
    }

    const resetButtons = [
      document.getElementById("resetFormBtn"),
      document.getElementById("resetFormBtnBottom")
    ].filter(Boolean);

    resetButtons.forEach((button) => {
      button.addEventListener("click", () => {
        const confirmed = window.confirm("Reset the form and clear the saved draft?");
        if (!confirmed) return;

        const form = document.getElementById("researchForm");
        if (form) form.reset();
        clearDraft();
        state.chart = {
          dataUrl: "",
          svg: "",
          stats: {
            currentPrice: null,
            rangeReturn: null,
            realisedVol: null
          }
        };
        state.figureAssets = [];
        const list = document.getElementById("coAuthorsList");
        if (list) list.innerHTML = "";
        renderFigureList();
        renderChartPreview();
        setText("chartStatus", "No chart loaded.");
        setDraftStatus("Not saved");
        setDefaultDate();
        syncEquitySection();
        showMessage("", "");
        renderPreview();
      });
    });

    const topExportButton = document.getElementById("exportDocBtnTop");
    if (topExportButton) {
      topExportButton.addEventListener("click", (event) => {
        event.preventDefault();
        exportWordDocument();
      });
    }
  }

  function setDefaultDate() {
    const publicationDate = document.getElementById("publicationDate");
    if (publicationDate && !publicationDate.value) {
      publicationDate.value = formatDateISO(new Date());
    }
  }

  function bindFormSync() {
    const form = document.getElementById("researchForm");
    if (!form) return;

    ["input", "change"].forEach((eventName) => {
      form.addEventListener(eventName, () => {
        syncEquitySection();
        scheduleDraftSave();
        renderPreview();
      });
    });

    form.addEventListener("submit", (event) => {
      event.preventDefault();
      exportWordDocument();
    });
  }

  function init() {
    setDefaultDate();
    bindStaticEvents();
    bindFormSync();
    renderChartPreview();

    const draft = loadDraft();
    if (draft) restoreDraft(draft);

    syncEquitySection();
    renderFigureList();
    renderPreview();
  }

  window.addEventListener("DOMContentLoaded", init);
})();
