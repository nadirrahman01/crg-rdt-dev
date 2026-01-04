console.log("app.js loaded successfully");

window.addEventListener("DOMContentLoaded", () => {
  // ================================
  // NEW: Phone formatting helpers (shared)
  // ================================
  function digitsOnly(v) {
    return (v || "").toString().replace(/\D/g, "");
  }

  // Visible formatting only (keeps UX tidy); hidden value stays cc-NUMBERDIGITS
  // Example visible: 7398 344 190 (groups of 4/3/3-ish)
  function formatNationalLoose(rawDigits) {
    const d = digitsOnly(rawDigits);
    if (!d) return "";

    // simple, readable grouping: 4-3-3 then remainder
    const p1 = d.slice(0, 4);
    const p2 = d.slice(4, 7);
    const p3 = d.slice(7, 10);
    const rest = d.slice(10);

    return [p1, p2, p3, rest].filter(Boolean).join(" ");
  }

  // Hidden combined format used across Word export + email meta, etc.
  // Example: "44-7398344190"
  function buildInternationalHyphen(ccDigits, nationalDigits) {
    const cc = digitsOnly(ccDigits);
    const nn = digitsOnly(nationalDigits);
    if (!cc && !nn) return "";
    if (cc && !nn) return `${cc}-`;
    if (!cc && nn) return nn;
    return `${cc}-${nn}`;
  }

  // Wire primary author phone (country select + national input -> hidden #authorPhone)
  const authorPhoneCountryEl = document.getElementById("authorPhoneCountry");
  const authorPhoneNationalEl = document.getElementById("authorPhoneNational");
  const authorPhoneHiddenEl = document.getElementById("authorPhone"); // kept for existing logic

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

    // best-effort caret stability
    const afterLen = authorPhoneNationalEl.value.length;
    const delta = afterLen - beforeLen;
    const next = Math.max(0, caret + delta);
    authorPhoneNationalEl.setSelectionRange(next, next);

    syncPrimaryPhone();
  }

  if (authorPhoneNationalEl) {
    authorPhoneNationalEl.addEventListener("input", formatPrimaryVisible);
    authorPhoneNationalEl.addEventListener("blur", syncPrimaryPhone);
  }
  if (authorPhoneCountryEl) {
    authorPhoneCountryEl.addEventListener("change", syncPrimaryPhone);
  }
  // initialise
  syncPrimaryPhone();

  // ================================
  // Co-Author Management
  // ================================
  let coAuthorCount = 0;

  const addCoAuthorBtn = document.getElementById("addCoAuthor");
  const coAuthorsList = document.getElementById("coAuthorsList");

  // NEW: Create a country code select HTML once (reuse)
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
    const ccEl = coAuthorDiv.querySelector(".coauthor-country");
    const nationalEl = coAuthorDiv.querySelector(".coauthor-phone-local");
    const hiddenEl = coAuthorDiv.querySelector(".coauthor-phone"); // existing expectation

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
      nationalEl.addEventListener("input", formatVisible);
      nationalEl.addEventListener("blur", syncHidden);
    }
    if (ccEl) ccEl.addEventListener("change", syncHidden);

    // init
    syncHidden();
  }

  if (addCoAuthorBtn) {
    addCoAuthorBtn.addEventListener("click", function () {
      coAuthorCount++;

      const coAuthorDiv = document.createElement("div");
      coAuthorDiv.className = "coauthor-entry";
      coAuthorDiv.id = `coauthor-${coAuthorCount}`;

      // IMPORTANT: keep .coauthor-phone (hidden) so the rest of your code stays unchanged
      coAuthorDiv.innerHTML = `
        <input type="text" placeholder="Last Name" class="coauthor-lastname" required>
        <input type="text" placeholder="First Name" class="coauthor-firstname" required>

        <div class="phone-row phone-row--compact">
          <select class="phone-country coauthor-country" aria-label="Country code">
            ${countryOptionsHtml}
          </select>
          <input type="text" placeholder="Phone number" class="phone-number coauthor-phone-local" inputmode="numeric">
        </div>

        <input type="text" class="coauthor-phone" style="display:none;" required>
        <button type="button" class="remove-coauthor" data-remove-id="${coAuthorCount}">Remove</button>
      `;

      coAuthorsList.appendChild(coAuthorDiv);

      // NEW: make phone optional even though HTML has required on hidden
      const phoneHidden = coAuthorDiv.querySelector(".coauthor-phone");
      if (phoneHidden) phoneHidden.required = false;

      // NEW: wire formatter
      wireCoauthorPhone(coAuthorDiv);

      // NEW: update completion meter (coauthors not core, but good to refresh UX)
      updateCompletionMeter();
    });

    document.addEventListener("click", (e) => {
      const btn = e.target.closest(".remove-coauthor");
      if (!btn) return;
      const id = btn.getAttribute("data-remove-id");
      const coAuthorDiv = document.getElementById(`coauthor-${id}`);
      if (coAuthorDiv) coAuthorDiv.remove();

      // NEW
      updateCompletionMeter();
    });
  }

  // ================================
  // Equity Research section toggle
  // ================================
  const noteTypeEl = document.getElementById("noteType");
  const equitySectionEl = document.getElementById("equitySection");

  function toggleEquitySection() {
    if (!noteTypeEl || !equitySectionEl) return;
    const isEquity = noteTypeEl.value === "Equity Research";
    equitySectionEl.style.display = isEquity ? "block" : "none";
  }

  if (noteTypeEl && equitySectionEl) {
    noteTypeEl.addEventListener("change", () => {
      toggleEquitySection();
      // NEW
      setTimeout(updateCompletionMeter, 0);
    });
    toggleEquitySection();
  } else {
    console.warn("Equity toggle not wired. Missing #noteType or #equitySection in index.html");
  }

  // ================================
  // NEW: Completion meter (8 core fields; 12 when Equity selected)
  // ================================
  const completionTextEl = document.getElementById("completionText");
  const completionBarEl = document.getElementById("completionBar");

  function isFilled(el) {
    if (!el) return false;
    if (el.type === "file") return el.files && el.files.length > 0;
    const v = (el.value ?? "").toString().trim();
    return v.length > 0;
  }

  // Base core (8): general notes
  const baseCoreIds = [
    "noteType",
    "title",
    "topic",
    "authorLastName",
    "authorFirstName",
    "keyTakeaways",
    "analysis",
    "cordobaView"
  ];

  // Equity adds 4 => total 12
  const equityCoreIds = [
    "ticker",
    "crgRating",
    "targetPrice",
    "modelFiles"
  ];

  function updateCompletionMeter() {
    const isEquity = (noteTypeEl && noteTypeEl.value === "Equity Research" && equitySectionEl && equitySectionEl.style.display !== "none");
    const ids = isEquity ? baseCoreIds.concat(equityCoreIds) : baseCoreIds;

    let done = 0;
    ids.forEach((id) => {
      const el = document.getElementById(id);
      if (isFilled(el)) done++;
    });

    const total = ids.length;
    const pct = total ? Math.round((done / total) * 100) : 0;

    if (completionTextEl) completionTextEl.textContent = `${done} / ${total} core fields`;
    if (completionBarEl) completionBarEl.style.width = `${pct}%`;

    const bar = completionBarEl?.parentElement;
    if (bar) bar.setAttribute("aria-valuenow", String(pct));
  }

  // Update meter on changes across the form
  ["input", "change", "keyup"].forEach(evt => {
    document.addEventListener(evt, (e) => {
      if (!e.target) return;
      if (e.target.closest && e.target.closest("#researchForm")) updateCompletionMeter();
    }, { passive: true });
  });

  // ================================
  // NEW: Attachment summary strip (modelFiles)
  // ================================
  const modelFilesEl2 = document.getElementById("modelFiles");
  const attachSummaryHeadEl = document.getElementById("attachmentSummaryHead");
  const attachSummaryListEl = document.getElementById("attachmentSummaryList");

  function updateAttachmentSummary() {
    if (!modelFilesEl2 || !attachSummaryHeadEl || !attachSummaryListEl) return;

    const files = Array.from(modelFilesEl2.files || []);
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

  if (modelFilesEl2) {
    modelFilesEl2.addEventListener("change", () => {
      updateAttachmentSummary();
      updateCompletionMeter();
    });
  }

  // ================================
  // NEW: Reset form button with confirm
  // ================================
  const resetBtn = document.getElementById("resetFormBtn");
  const formEl = document.getElementById("researchForm");

  function clearChartUI() {
    // clear chart stats UI safely
    const setText = (id, text) => {
      const el = document.getElementById(id);
      if (el) el.textContent = text;
    };
    setText("currentPrice", "—");
    setText("realisedVol", "—");
    setText("rangeReturn", "—");
    setText("upsideToTarget", "—");

    const chartStatus = document.getElementById("chartStatus");
    if (chartStatus) chartStatus.textContent = "";

    // if chart exists, destroy
    if (typeof priceChart !== "undefined" && priceChart) {
      try { priceChart.destroy(); } catch (_) {}
      priceChart = null;
    }

    // reset globals if present
    if (typeof priceChartImageBytes !== "undefined") priceChartImageBytes = null;
    if (typeof equityStats !== "undefined") {
      equityStats = { currentPrice: null, realisedVolAnn: null, rangeReturn: null };
    }
  }

  if (resetBtn && formEl) {
    resetBtn.addEventListener("click", () => {
      const ok = confirm("Reset the form? This will clear all fields on this page.");
      if (!ok) return;

      formEl.reset();

      // Clear co-authors added dynamically
      if (coAuthorsList) coAuthorsList.innerHTML = "";

      // Clear file input + attachment summary
      if (modelFilesEl2) modelFilesEl2.value = "";
      updateAttachmentSummary();

      // Clear message box
      const messageDiv = document.getElementById("message");
      if (messageDiv) {
        messageDiv.className = "message";
        messageDiv.textContent = "";
        messageDiv.style.display = "none";
      }

      // Clear chart UI
      clearChartUI();

      // Re-sync phone hidden values
      syncPrimaryPhone();

      // Hide/show equity section appropriately
      toggleEquitySection();

      // Refresh meter
      setTimeout(updateCompletionMeter, 0);
    });
  }

  // Initial paint
  updateAttachmentSummary();
  updateCompletionMeter();

  // ================================
  // Date/time formatting
  // ================================
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

  // ================================
  // NEW: Email to CRG (prefilled mailto)
  // Note: browsers cannot auto-attach files for security reasons.
  // User will attach the downloaded Word doc manually.
  // ================================
  const emailToCrgBtn = document.getElementById("emailToCrgBtn");

  function formatDateShort(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  // IMPORTANT FIX:
  // - Avoid URLSearchParams (it turns spaces into "+")
  // - Use encodeURIComponent directly
  // - Force CRLF for clean line breaks in email clients
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
    const noteType = (document.getElementById("noteType")?.value || "Research Note").trim();
    const title = (document.getElementById("title")?.value || "").trim();
    const topic = (document.getElementById("topic")?.value || "").trim();

    const authorFirstName = (document.getElementById("authorFirstName")?.value || "").trim();
    const authorLastName = (document.getElementById("authorLastName")?.value || "").trim();

    const ticker = (document.getElementById("ticker")?.value || "").trim();
    const crgRating = (document.getElementById("crgRating")?.value || "").trim();
    const targetPrice = (document.getElementById("targetPrice")?.value || "").trim();

    const now = new Date();
    const dateShort = formatDateShort(now);
    const dateLong = formatDateTime(now);

    const subjectParts = [
      noteType || "Research Note",
      dateShort,
      title ? `— ${title}` : ""
    ].filter(Boolean);

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

  if (emailToCrgBtn) {
    emailToCrgBtn.addEventListener("click", () => {
      const { subject, body, cc } = buildCrgEmailPayload();

      const to = "research@cordobarg.com";
      const mailto = buildMailto(to, cc, subject, body);

      window.location.href = mailto;
    });
  }

  // ================================
  // NEW: optional field helpers
  // ================================
  function naIfBlank(v) {
    const s = (v ?? "").toString().trim();
    // FIX: return "N/A" (no brackets) so coAuthorLine produces "(N/A)" not "((N/A))"
    return s ? s : "N/A";
  }

  function coAuthorLine(coAuthor) {
    const ln = (coAuthor.lastName || "").toUpperCase();
    const fn = (coAuthor.firstName || "").toUpperCase();
    const ph = naIfBlank(coAuthor.phone);
    return `${ln}, ${fn} (${ph})`;
  }

  // ================================
  // Add images to Word (user uploads)
  // ================================
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
                transformation: { width: 600, height: 450 }
              })
            ],
            spacing: { before: 200, after: 100 },
            alignment: docx.AlignmentType.CENTER
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: `Figure ${i + 1}: ${fileNameWithoutExt}`,
                italics: true,
                size: 18,
                font: "Book Antiqua"
              })
            ],
            spacing: { after: 300 },
            alignment: docx.AlignmentType.CENTER
          })
        );
      } catch (error) {
        console.error(`Error processing image ${file.name}:`, error);
      }
    }
    return imageParagraphs;
  }

  function linesToParagraphs(text, spacingAfter = 150) {
    const lines = (text || "").split("\n");
    return lines.map((line) => {
      if (line.trim() === "") {
        return new docx.Paragraph({ text: "", spacing: { after: spacingAfter } });
      }
      return new docx.Paragraph({ text: line, spacing: { after: spacingAfter } });
    });
  }

  function hyperlinkParagraph(label, url) {
    const safeUrl = (url || "").trim();
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

  // ================================
  // Price chart (Stooq -> Chart.js -> Word image)
  // FIX: Stooq has no CORS. Use r.jina.ai proxy.
  // + NEW: compute stats (current price, vol, range return, upside to target)
  // ================================
  let priceChart = null;
  let priceChartImageBytes = null;

  let equityStats = {
    currentPrice: null,
    realisedVolAnn: null,
    rangeReturn: null
  };

  const chartStatus = document.getElementById("chartStatus");
  const fetchChartBtn = document.getElementById("fetchPriceChart");
  const chartRangeEl = document.getElementById("chartRange");
  const priceChartCanvas = document.getElementById("priceChart");
  const targetPriceEl = document.getElementById("targetPrice");

  function stooqSymbolFromTicker(ticker) {
    const t = (ticker || "").trim();
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

  // ================================
  // stats helpers
  // ================================
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
    const el = document.getElementById(id);
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

  if (targetPriceEl) {
    targetPriceEl.addEventListener("input", () => {
      updateUpsideDisplay();
      updateCompletionMeter(); // NEW: targetPrice is a core field for Equity
    });
  }

  async function buildPriceChart() {
    try {
      const tickerVal = (document.getElementById("ticker")?.value || "").trim();
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
    } catch (e) {
      priceChartImageBytes = null;

      equityStats = { currentPrice: null, realisedVolAnn: null, rangeReturn: null };
      setText("currentPrice", "—");
      setText("rangeReturn", "—");
      setText("realisedVol", "—");
      setText("upsideToTarget", "—");

      if (chartStatus) chartStatus.textContent = `✗ ${e.message}`;
    } finally {
      // NEW
      updateCompletionMeter();
    }
  }

  if (fetchChartBtn) fetchChartBtn.addEventListener("click", buildPriceChart);

  // ================================
  // Create Word Document
  // ================================
  async function createDocument(data) {
    const {
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
    } = data;

    const takeawayLines = (keyTakeaways || "").split("\n");
    const takeawayBullets = takeawayLines.map(line => {
      if (line.trim() === "") return new docx.Paragraph({ text: "", spacing: { after: 100 } });
      const cleanLine = line.replace(/^[-*•]\s*/, "").trim();
      return new docx.Paragraph({ text: cleanLine, bullet: { level: 0 }, spacing: { after: 100 } });
    });

    const analysisParagraphs = linesToParagraphs(analysis, 150);
    const contentParagraphs = linesToParagraphs(content, 150);
    const cordobaViewParagraphs = linesToParagraphs(cordobaView, 150);

    const imageParagraphs = await addImages(imageFiles);

    // NEW: ensure author phone prints as "(N/A)" (single bracket pair)
    const authorPhonePrintable = authorPhoneSafe ? authorPhoneSafe : naIfBlank(authorPhone);
    const authorPhoneWrapped = authorPhonePrintable ? `(${authorPhonePrintable})` : "(N/A)";

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
            new docx.TableCell({
              // ================================
              // ✅ UPDATE: Topic (10pt, not bold) + Title (14pt, bold)
              // Topic is placed where title used to be; title sits directly below it.
              // ================================
              children: [
                // TOPIC line (10pt, not bold)
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: "TOPIC: ", bold: false, size: 20, font: "Book Antiqua" }),
                    new docx.TextRun({ text: (topic || ""), bold: false, size: 20, font: "Book Antiqua" })
                  ],
                  spacing: { after: 120 }
                }),
                // TITLE line (14pt, bold)
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: (title || ""),
                      bold: true,
                      size: 28, // 14pt
                      font: "Book Antiqua"
                    })
                  ],
                  spacing: { after: 100 }
                })
              ],
              width: { size: 60, type: docx.WidthType.PERCENTAGE },
              verticalAlign: docx.VerticalAlign.TOP
            }),
            new docx.TableCell({
              children: [
                new docx.Paragraph({
                  children: [new docx.TextRun({
                    text: `${authorLastName.toUpperCase()}, ${authorFirstName.toUpperCase()} ${authorPhoneWrapped}`,
                    bold: true,
                    size: 28,
                    font: "Book Antiqua"
                  })],
                  alignment: docx.AlignmentType.RIGHT,
                  spacing: { after: 100 }
                })
              ],
              width: { size: 40, type: docx.WidthType.PERCENTAGE },
              verticalAlign: docx.VerticalAlign.TOP
            })
          ]
        }),
        new docx.TableRow({
          children: [
            new docx.TableCell({
              // keep structure; blank so topic/title aren't duplicated
              children: [new docx.Paragraph({ text: "", spacing: { after: 200 } })],
              width: { size: 60, type: docx.WidthType.PERCENTAGE },
              verticalAlign: docx.VerticalAlign.TOP
            }),
            new docx.TableCell({
              children: coAuthors.length > 0
                ? coAuthors.map(coAuthor =>
                  new docx.Paragraph({
                    children: [new docx.TextRun({
                      text: coAuthorLine(coAuthor),
                      bold: true,
                      size: 28,
                      font: "Book Antiqua"
                    })],
                    alignment: docx.AlignmentType.RIGHT,
                    spacing: { after: 100 }
                  })
                )
                : [new docx.Paragraph({ text: "" })],
              width: { size: 40, type: docx.WidthType.PERCENTAGE },
              verticalAlign: docx.VerticalAlign.TOP
            })
          ]
        })
      ]
    });

    const documentChildren = [
      infoTable,
      new docx.Paragraph({
        border: { bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } },
        spacing: { after: 300 }
      })
    ];

    if (noteType === "Equity Research") {
      const attachedModelNames = (modelFiles && modelFiles.length) ? Array.from(modelFiles).map(f => f.name) : [];

      if ((ticker || "").trim()) {
        documentChildren.push(
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Ticker / Company: ", bold: true }),
              new docx.TextRun({ text: ticker.trim() })
            ],
            spacing: { after: 120 }
          })
        );
      }

      if ((crgRating || "").trim()) {
        documentChildren.push(
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "CRG Rating: ", bold: true }),
              new docx.TextRun({ text: crgRating.trim() })
            ],
            spacing: { after: 120 }
          })
        );
      }

      const modelLinkPara = hyperlinkParagraph("Model link:", modelLink);
      if (modelLinkPara) documentChildren.push(modelLinkPara);

      if (priceChartImageBytes) {
        documentChildren.push(
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "Price Chart", bold: true, size: 24, font: "Book Antiqua" })],
            spacing: { before: 120, after: 120 }
          }),
          new docx.Paragraph({
            children: [new docx.ImageRun({ data: priceChartImageBytes, transformation: { width: 650, height: 300 } })],
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 200 }
          })
        );
      }

      if (equityStats && equityStats.currentPrice) {
        const tp = (targetPrice || "").trim();
        const tpNum = safeNum(tp);
        const upside = computeUpsideToTarget(equityStats.currentPrice, tpNum);

        documentChildren.push(
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "Market Stats", bold: true, size: 24, font: "Book Antiqua" })],
            spacing: { before: 80, after: 100 }
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Current price: ", bold: true }),
              new docx.TextRun({ text: equityStats.currentPrice.toFixed(2) })
            ],
            spacing: { after: 80 }
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Volatility (ann.): ", bold: true }),
              new docx.TextRun({ text: equityStats.realisedVolAnn == null ? "—" : pct(equityStats.realisedVolAnn) })
            ],
            spacing: { after: 80 }
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Return (range): ", bold: true }),
              new docx.TextRun({ text: equityStats.rangeReturn == null ? "—" : pct(equityStats.rangeReturn) })
            ],
            spacing: { after: 80 }
          })
        );

        if (tpNum) {
          documentChildren.push(
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: "Target price: ", bold: true }),
                new docx.TextRun({ text: tpNum.toFixed(2) })
              ],
              spacing: { after: 80 }
            }),
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: "+/- to target: ", bold: true }),
                new docx.TextRun({ text: upside == null ? "—" : pct(upside) })
              ],
              spacing: { after: 120 }
            })
          );
        } else {
          documentChildren.push(new docx.Paragraph({ spacing: { after: 80 } }));
        }
      }

      documentChildren.push(
        new docx.Paragraph({
          children: [new docx.TextRun({ text: "Attached model files:", bold: true, size: 24, font: "Book Antiqua" })],
          spacing: { after: 120 }
        })
      );

      if (attachedModelNames.length) {
        attachedModelNames.forEach(name => {
          documentChildren.push(
            new docx.Paragraph({ text: name, bullet: { level: 0 }, spacing: { after: 80 } })
          );
        });
      } else {
        documentChildren.push(new docx.Paragraph({ text: "None uploaded", spacing: { after: 120 } }));
      }

      if ((valuationSummary || "").trim()) {
        documentChildren.push(
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "Valuation Summary", bold: true, size: 24, font: "Book Antiqua" })],
            spacing: { before: 120, after: 100 }
          }),
          ...linesToParagraphs(valuationSummary, 120)
        );
      }

      if ((keyAssumptions || "").trim()) {
        documentChildren.push(
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "Key Assumptions", bold: true, size: 24, font: "Book Antiqua" })],
            spacing: { before: 120, after: 100 }
          })
        );

        keyAssumptions.split("\n").forEach(line => {
          if (!line.trim()) return;
          documentChildren.push(
            new docx.Paragraph({
              text: line.replace(/^[-*•]\s*/, "").trim(),
              bullet: { level: 0 },
              spacing: { after: 80 }
            })
          );
        });
      }

      if ((scenarioNotes || "").trim()) {
        documentChildren.push(
          new docx.Paragraph({
            children: [new docx.TextRun({ text: "Scenario / Sensitivity Notes", bold: true, size: 24, font: "Book Antiqua" })],
            spacing: { before: 120, after: 100 }
          }),
          ...linesToParagraphs(scenarioNotes, 120)
        );
      }

      documentChildren.push(new docx.Paragraph({ spacing: { after: 250 } }));
    }

    documentChildren.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "Key Takeaways", bold: true, size: 24, font: "Book Antiqua" })],
        spacing: { after: 200 }
      }),
      ...takeawayBullets,
      new docx.Paragraph({ spacing: { after: 300 } }),
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "Analysis and Commentary", bold: true, size: 24, font: "Book Antiqua" })],
        spacing: { after: 200 }
      }),
      ...analysisParagraphs,
      ...contentParagraphs
    );

    if ((cordobaView || "").trim()) {
      documentChildren.push(
        new docx.Paragraph({ spacing: { after: 300 } }),
        new docx.Paragraph({
          children: [new docx.TextRun({ text: "The Cordoba View", bold: true, size: 24, font: "Book Antiqua" })],
          spacing: { after: 200 }
        }),
        ...cordobaViewParagraphs
      );
    }

    if (imageParagraphs.length > 0) {
      documentChildren.push(
        new docx.Paragraph({
          children: [new docx.TextRun({ text: "Figures and Charts", bold: true, size: 24, font: "Book Antiqua" })],
          spacing: { before: 400, after: 200 }
        }),
        ...imageParagraphs
      );
    }

    const doc = new docx.Document({
      styles: {
        default: {
          document: {
            run: { font: "Book Antiqua", size: 20, color: "000000" },
            paragraph: { spacing: { after: 150 } }
          }
        }
      },
      sections: [{
        properties: {
          page: {
            margin: { top: 720, right: 720, bottom: 720, left: 720 },
            pageSize: { orientation: docx.PageOrientation.LANDSCAPE, width: 15840, height: 12240 }
          }
        },
        headers: {
          default: new docx.Header({
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    text: `Cordoba Research Group | ${noteType} | Published on ${dateTimeString}`,
                    size: 16,
                    font: "Book Antiqua"
                  })
                ],
                alignment: docx.AlignmentType.RIGHT,
                spacing: { after: 100 },
                border: { bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } }
              })
            ]
          })
        },
        footers: {
          default: new docx.Footer({
            children: [
              new docx.Paragraph({
                border: { top: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 } },
                spacing: { after: 0 }
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({ text: "\t" }),
                  new docx.TextRun({
                    text: "Cordoba Research Group Internal Information",
                    size: 16,
                    font: "Book Antiqua",
                    italics: true
                  }),
                  new docx.TextRun({ text: "\t" }),
                  new docx.TextRun({
                    children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
                    size: 16,
                    font: "Book Antiqua",
                    italics: true
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
        children: documentChildren
      }]
    });

    return doc;
  }

  // ================================
  // Main Form Submission
  // ================================
  const form = document.getElementById("researchForm");
  if (form) form.noValidate = true;

  form.addEventListener("submit", async function (e) {
    e.preventDefault();

    const button = form.querySelector('button[type="submit"]');
    const messageDiv = document.getElementById("message");

    button.disabled = true;
    button.classList.add("loading");
    button.textContent = "Generating Document...";
    messageDiv.className = "message";
    messageDiv.textContent = "";

    try {
      if (typeof docx === "undefined") throw new Error("docx library not loaded. Please refresh the page.");
      if (typeof saveAs === "undefined") throw new Error("FileSaver library not loaded. Please refresh the page.");

      const noteType = document.getElementById("noteType").value;
      const title = document.getElementById("title").value;
      const topic = document.getElementById("topic").value;
      const authorLastName = document.getElementById("authorLastName").value;
      const authorFirstName = document.getElementById("authorFirstName").value;

      // IMPORTANT: hidden field stays source-of-truth
      const authorPhone = document.getElementById("authorPhone").value;
      const authorPhoneSafe = naIfBlank(authorPhone);

      const analysis = document.getElementById("analysis").value;
      const keyTakeaways = document.getElementById("keyTakeaways").value;
      const content = document.getElementById("content").value;
      const cordobaView = document.getElementById("cordobaView").value;
      const imageFiles = document.getElementById("imageUpload").files;

      const ticker = document.getElementById("ticker") ? document.getElementById("ticker").value : "";
      const valuationSummary = document.getElementById("valuationSummary") ? document.getElementById("valuationSummary").value : "";
      const keyAssumptions = document.getElementById("keyAssumptions") ? document.getElementById("keyAssumptions").value : "";
      const scenarioNotes = document.getElementById("scenarioNotes") ? document.getElementById("scenarioNotes").value : "";
      const modelFiles = document.getElementById("modelFiles") ? document.getElementById("modelFiles").files : null;
      const modelLink = document.getElementById("modelLink") ? document.getElementById("modelLink").value : "";

      const targetPrice = document.getElementById("targetPrice") ? document.getElementById("targetPrice").value : "";
      const crgRating = document.getElementById("crgRating") ? document.getElementById("crgRating").value : "";

      const now = new Date();
      const dateTimeString = formatDateTime(now);

      const coAuthors = [];
      document.querySelectorAll(".coauthor-entry").forEach(entry => {
        const lastName = entry.querySelector(".coauthor-lastname").value;
        const firstName = entry.querySelector(".coauthor-firstname").value;
        const phone = entry.querySelector(".coauthor-phone").value; // hidden combined
        if (lastName && firstName) coAuthors.push({ lastName, firstName, phone: naIfBlank(phone) });
      });

      const doc = await createDocument({
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
        `${title.replace(/[^a-z0-9]/gi, "_").toLowerCase()}_${noteType.replace(/\s+/g, "_").toLowerCase()}.docx`;

      saveAs(blob, fileName);

      messageDiv.className = "message success";
      messageDiv.textContent = `✓ Document "${fileName}" generated successfully!`;
    } catch (error) {
      console.error("Error generating document:", error);
      messageDiv.className = "message error";
      messageDiv.textContent = `✗ Error: ${error.message}`;
    } finally {
      button.disabled = false;
      button.classList.remove("loading");
      button.textContent = "Generate Word Document";
    }
  });
});
