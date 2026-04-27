(function () {
  "use strict";

  const authRoot = document.getElementById("authRoot");
  const platformRoot = document.getElementById("platformRoot");
  const legacyHost = document.getElementById("legacyAppHost");
  const API_BASE = window.location.protocol === "file:" ? "http://localhost:3000" : "";

  const NAV_ITEMS = [
    { route: "dashboard", label: "Dashboard", icon: "dashboard" },
    { route: "research-library", label: "Research Library", icon: "library" },
    { route: "drafts", label: "Drafts", icon: "drafts" },
    { route: "live-edit", label: "Live Edit", icon: "edit" },
    { route: "templates", label: "Templates", icon: "templates" },
    { route: "distribution", label: "Distribution", icon: "distribution" },
    { route: "compliance", label: "Compliance", icon: "compliance" },
    { route: "calendar", label: "Calendar", icon: "calendar" },
    { route: "settings", label: "Settings", icon: "settings", footer: true }
  ];

  const QUICK_CREATE = [
    { label: "Equity Research", noteType: "Equity Research", copy: "Company, sector and industry analysis", icon: "chart-up" },
    { label: "Macro / Fixed Income", noteType: "Macro Research", copy: "Macro analysis, rates, FX and credit", icon: "macro" },
    { label: "Commodity Insights", noteType: "Commodity Insights", copy: "Commodities, energy and metals", icon: "commodity" },
    { label: "Market Alert", noteType: "Short Note / Market Alert", copy: "Timely market-moving updates", icon: "alert" },
    { label: "General Note", noteType: "General Note", copy: "General commentary and thematic notes", icon: "note" }
  ];

  const DASHBOARD_CARDS = [
    {
      title: "Uzbekistan Sovereign Outlook",
      savedAgo: "2m ago",
      copy: "Reforms and resilience support gradual improvement",
      tags: ["Sovereign", "Emerging Markets"],
      initials: "AS",
      author: "Alex Smith"
    },
    {
      title: "EM Local Currency Strategy Update",
      savedAgo: "45m ago",
      copy: "Carry, flows, and positioning into Q3",
      tags: ["Rates", "Strategy"],
      initials: "PN",
      author: "Priya Natarajan"
    },
    {
      title: "Copper Price Outlook",
      savedAgo: "1h ago",
      copy: "Supply constraints and demand rebalancing in 2025",
      tags: ["Commodities", "Metals"],
      initials: "DK",
      author: "Daniel Kwon"
    },
    {
      title: "US Inflation Monitor",
      savedAgo: "3h ago",
      copy: "Disinflation path remains intact into mid-2025",
      tags: ["Macro", "United States"],
      initials: "AS",
      author: "Alex Smith"
    }
  ];

  const RECENT_DRAFTS = [
    { title: "Uzbekistan Sovereign Outlook", age: "2m ago" },
    { title: "EM Local Currency Strategy Update", age: "45m ago" },
    { title: "Copper Price Outlook", age: "1h ago" },
    { title: "US Inflation Monitor", age: "3h ago" },
    { title: "China Property Sector Review", age: "5h ago" }
  ];

  const PUBLICATION_QUEUE = [
    { title: "US Rate Outlook", desk: "US Economics", state: "In Review", analyst: "Alex Smith", date: "May 23, 2025", time: "10:00 AM" },
    { title: "Brazil Fiscal Update", desk: "Emerging Markets", state: "Compliance Review", analyst: "Daniel Kwon", date: "May 23, 2025", time: "1:00 PM" },
    { title: "Oil Market Balance", desk: "Commodities", state: "Final Review", analyst: "Priya Natarajan", date: "May 26, 2025", time: "9:00 AM" },
    { title: "India Growth Outlook", desk: "Emerging Markets", state: "Scheduled", analyst: "Alex Smith", date: "May 26, 2025", time: "11:00 AM" }
  ];

  const UPCOMING_SCHEDULE = [
    { date: "Friday, May 23", time: "10:00 AM", title: "US Rate Outlook", desk: "US Economics" },
    { date: "Friday, May 23", time: "1:00 PM", title: "Brazil Fiscal Update", desk: "Emerging Markets" },
    { date: "Monday, May 26", time: "9:00 AM", title: "Oil Market Balance", desk: "Commodities" },
    { date: "Monday, May 26", time: "11:00 AM", title: "India Growth Outlook", desk: "Emerging Markets" }
  ];

  const DEFAULT_SETTINGS = {
    profile: {
      fullName: "",
      email: "",
      phoneNumber: "",
      jobTitle: ""
    },
    defaults: {
      defaultNoteType: "Macro / Sovereign Outlook",
      defaultRegion: "Emerging Markets"
    },
    notifications: {
      publicationReminders: true,
      draftActivity: true,
      validationAlerts: true,
      systemUpdates: true
    },
    security: {
      passwordChangedAt: "",
      twoFactorEnabled: true,
      activeSessions: 1
    }
  };

  const state = {
    api: null,
    authChecked: false,
    user: null,
    security: null,
    settings: clone(DEFAULT_SETTINGS),
    settingsDraft: clone(DEFAULT_SETTINGS),
    snapshot: null,
    route: getRouteFromHash(),
    userMenuOpen: false,
    structureTab: "structure",
    contextTab: "tools",
    publishTab: "publish",
    previewTab: "preview-export",
    selectedSectionKey: "header",
    showPassword: false,
    toast: null,
    toastTimer: null,
    previewDevice: "desktop",
    searchQuery: "",
    exportOptions: {
      includeDisclosures: true,
      includeAnalystBio: true,
      includeRelatedResearch: true,
      channels: {
        researchPortal: true,
        clientAlert: true,
        emailAlert: true,
        bloombergApi: true
      }
    }
  };

  document.addEventListener("DOMContentLoaded", init);

  async function init() {
    if (!authRoot || !platformRoot) return;
    if (legacyHost) legacyHost.hidden = true;

    bindGlobalEvents();
    state.api = await waitForLegacyApi();
    state.snapshot = state.api.getSnapshot();
    state.api.subscribe((snapshot) => {
      state.snapshot = snapshot;
      ensureSelectedSection();
      if (state.user) renderShell();
    });

    await hydrateSession();
  }

  function bindGlobalEvents() {
    authRoot.addEventListener("click", handleAuthClick);
    authRoot.addEventListener("submit", handleAuthSubmit);

    platformRoot.addEventListener("click", handlePlatformClick);
    platformRoot.addEventListener("change", handlePlatformChange);
    platformRoot.addEventListener("input", handlePlatformInput);
    platformRoot.addEventListener("blur", handlePlatformBlur, true);
    platformRoot.addEventListener("dragstart", handleDragStart);
    platformRoot.addEventListener("dragover", handleDragOver);
    platformRoot.addEventListener("drop", handleDrop);

    document.addEventListener("click", handleDocumentClick);
    window.addEventListener("hashchange", () => {
      state.route = getRouteFromHash();
      renderShell();
    });
  }

  async function hydrateSession() {
    try {
      const session = await fetchJson("/api/auth/session");
      state.authChecked = true;
      if (!session.authenticated) {
        state.user = null;
        state.security = null;
        renderAuth();
        return;
      }

      state.user = session.user;
      state.security = session.security;
      await hydrateSettings();
      ensureSelectedSection();
      renderShell();
    } catch (_error) {
      state.authChecked = true;
      renderAuth(authUnavailableMessage());
    }
  }

  async function hydrateSettings() {
    try {
      const settings = await fetchJson("/api/settings");
      state.settings = mergeSettings(settings);
      state.settingsDraft = clone(state.settings);
    } catch (_error) {
      state.settings = mergeSettings({
        profile: {
          fullName: state.user?.fullName || "",
          email: state.user?.email || ""
        },
        security: state.security || DEFAULT_SETTINGS.security
      });
      state.settingsDraft = clone(state.settings);
    }
  }

  function renderAuth(errorMessage = "") {
    document.body.classList.add("rpc-auth-open");
    platformRoot.hidden = true;
    authRoot.hidden = false;
    authRoot.innerHTML = `${buildAuthMarkup(errorMessage)}${state.toast ? `<div class="rpc-toast ${state.toast.tone === "error" ? "is-error" : ""}">${escapeHtml(state.toast.message)}</div>` : ""}`;
  }

  function renderShell() {
    if (!state.user || !state.snapshot) return;
    document.body.classList.remove("rpc-auth-open");
    authRoot.hidden = true;
    authRoot.innerHTML = "";
    platformRoot.hidden = false;

    platformRoot.innerHTML = `
      <div class="rpc-shell">
        ${renderSidebar()}
        <div class="rpc-app">
          ${renderTopbar()}
          <main class="rpc-page">
            ${renderRoutePage()}
          </main>
        </div>
      </div>
      ${state.toast ? `<div class="rpc-toast ${state.toast.tone === "error" ? "is-error" : ""}">${escapeHtml(state.toast.message)}</div>` : ""}
    `;

    const previewFrame = platformRoot.querySelector("[data-preview-frame]");
    if (previewFrame) previewFrame.srcdoc = state.snapshot.previewSrcdoc || "";
  }

  function buildAuthMarkup(errorMessage) {
    return `
      <section class="rpc-login">
        <div class="rpc-login-panel">
          <div class="rpc-login-brand">
            <img src="${assetUrl("assets/cordoba-logo")}" alt="Cordoba Research Group">
            <h1>Research Production Console</h1>
            <p>Research documentation, analysis, and compliance in one secure platform.</p>
          </div>
          <div class="rpc-login-card">
            <form class="rpc-login-form" data-auth-form>
              <div class="rpc-form-group">
                <label for="rpcLoginEmail">Work Email</label>
                <div class="rpc-input-wrap has-leading">
                  ${icon("user")}
                  <input id="rpcLoginEmail" name="email" type="email" autocomplete="username" placeholder="name@cordobaresearch.com" required>
                </div>
              </div>
              <div class="rpc-form-group">
                <label for="rpcLoginPassword">Password</label>
                <div class="rpc-input-wrap has-leading has-trailing">
                  ${icon("lock")}
                  <input id="rpcLoginPassword" name="password" type="${state.showPassword ? "text" : "password"}" autocomplete="current-password" placeholder="Enter your password" required>
                  <button type="button" class="rpc-password-toggle rpc-input-trailing" data-toggle-password aria-label="Toggle password visibility">
                    ${state.showPassword ? icon("eye-off") : icon("eye")}
                  </button>
                </div>
              </div>
              <div class="rpc-login-row">
                <label class="rpc-check">
                  <input type="checkbox" name="remember">
                  <span>Remember me</span>
                </label>
                <a class="rpc-link" href="mailto:support@cordobaresearch.com">Forgot password?</a>
              </div>
              <button type="submit" class="rpc-primary-btn">Sign In</button>
              <div class="rpc-login-divider">or</div>
              <button type="button" class="rpc-secondary-btn rpc-ms-btn" data-sso-button>
                <span class="rpc-ms-btn-content">
                  <span class="rpc-ms-mark">${icon("microsoft")}</span>
                  <span class="rpc-ms-label">Continue with Microsoft</span>
                </span>
              </button>
              ${errorMessage ? `<div class="rpc-login-error">${escapeHtml(errorMessage)}</div>` : ""}
            </form>
          </div>
          <div class="rpc-login-note">
            <div class="rpc-login-note-row">${icon("shield")}<span>Single Sign-On enabled</span></div>
            <div class="rpc-login-note-row">${icon("lock")}<span>Secure access for approved analysts.</span></div>
          </div>
          <div class="rpc-login-support">Need help? Contact <a class="rpc-link" href="mailto:support@cordobaresearch.com">support@cordobaresearch.com</a></div>
        </div>
      </section>
    `;
  }

  function renderSidebar() {
    const primary = NAV_ITEMS.filter((item) => !item.footer);
    const footer = NAV_ITEMS.filter((item) => item.footer);

    return `
      <aside class="rpc-sidebar">
        <div class="rpc-sidebar-brand">Research Production Console</div>
        <nav class="rpc-sidebar-nav">
          ${primary.map((item) => `
            <button type="button" class="rpc-sidebar-link ${isRouteActive(item.route) ? "is-active" : ""}" data-nav-route="${item.route}">
              ${icon(item.icon)}
              <span>${escapeHtml(item.label)}</span>
            </button>
          `).join("")}
        </nav>
        <div class="rpc-sidebar-foot">
          ${footer.map((item) => `
            <button type="button" class="rpc-sidebar-link ${isRouteActive(item.route) ? "is-active" : ""}" data-nav-route="${item.route}">
              ${icon(item.icon)}
              <span>${escapeHtml(item.label)}</span>
            </button>
          `).join("")}
        </div>
      </aside>
    `;
  }

  function renderTopbar() {
    const variant = noteVariant(state.snapshot?.data?.noteType);
    const title = topbarTitleForRoute(variant);
    const compact = state.route === "live-edit" && variant === "alert";

    return `
      <header class="rpc-topbar">
        <div class="rpc-topbar-title">${escapeHtml(title)}</div>
        ${compact ? `<div></div><div></div>` : `
          <div class="rpc-search">
            ${icon("search")}
            <input class="rpc-search-input" type="search" value="${escapeAttr(state.searchQuery)}" placeholder="Search research, notes, authors, topics..." data-global-search>
            <span class="rpc-search-kbd">⌘ K</span>
          </div>
          <select class="rpc-workspace-select" data-workspace-select>
            ${workspaceOptions().map((option) => `
              <option value="${escapeAttr(option.value)}" ${option.value === workspaceLabelForNoteType(state.snapshot?.data?.noteType) ? "selected" : ""}>${escapeHtml(option.label)}</option>
            `).join("")}
          </select>
        `}
        <button type="button" class="rpc-primary-btn rpc-topbar-publish" data-nav-route="preview-export">Publish</button>
        <div class="rpc-user">
          <button type="button" class="rpc-user-btn" data-open-user-menu>
            <span class="rpc-user-badge">${escapeHtml(state.user?.initials || "CR")}</span>
            <span>${escapeHtml(state.user?.fullName || "User")}</span>
            <span>${icon("chevron-down")}</span>
          </button>
          ${state.userMenuOpen ? `
            <div class="rpc-user-menu">
              <button type="button" class="rpc-profile-action" data-nav-route="settings">Settings</button>
              <button type="button" class="rpc-profile-action" data-profile-action="logout">Sign Out</button>
            </div>
          ` : ""}
        </div>
      </header>
    `;
  }

  function renderRoutePage() {
    switch (state.route) {
      case "dashboard":
        return renderDashboardPage();
      case "live-edit":
        return renderEditorPage();
      case "preview-export":
        return renderPreviewExportPage();
      case "settings":
        return renderSettingsPage();
      case "research-library":
        return renderPlaceholderPage("Research Library", "Structured research listings, filters, and archived notes will sit here.");
      case "drafts":
        return renderPlaceholderPage("Drafts", "Saved draft management will appear here with ownership, timestamps, and quick open actions.");
      case "templates":
        return renderPlaceholderPage("Templates", "House templates, reusable note frameworks, and component presets will live here.");
      case "distribution":
        return renderPlaceholderPage("Distribution", "Channel selection, lists, and packaging controls will appear here.");
      case "compliance":
        return renderPlaceholderPage("Compliance", "Controlled compliance checks and note-level flags will appear here.");
      case "calendar":
        return renderPlaceholderPage("Calendar", "Publication scheduling and editorial timing will appear here.");
      default:
        return renderDashboardPage();
    }
  }

  function renderDashboardPage() {
    const snapshot = state.snapshot;
    const validation = snapshot.validation || { percent: 0, missing: [] };
    const readinessPercent = Math.max(0, Math.min(100, Number(validation.percent || 0)));
    const readyCount = Math.max(1, Math.round((readinessPercent / 100) * 50));
    const inProgressCount = Math.max(1, Math.round(((100 - readinessPercent) / 100) * 8));
    const atRiskCount = snapshot.review?.findings?.filter((finding) => finding.blocking).length || 1;
    const cards = buildDashboardCards(snapshot);

    return `
      <div class="rpc-dashboard-grid">
        <div class="rpc-page-head">
          <div><h1>Continue Editing</h1></div>
          <a class="rpc-page-link" href="#/drafts">View all drafts</a>
        </div>
        <div class="rpc-draft-card-grid">
          ${cards.map((card) => renderDraftCard(card)).join("")}
        </div>

        <div class="rpc-dashboard-row">
          <section class="rpc-list-card">
            <div class="rpc-section-card-head"><h3>Recent Drafts</h3><a class="rpc-page-link" href="#/drafts">View all</a></div>
            <div class="rpc-recent-list">
              ${RECENT_DRAFTS.map((item, index) => `
                <button type="button" class="rpc-recent-item ${index === 0 ? "is-active" : ""}" data-nav-route="live-edit">
                  ${icon("doc")}
                  <div class="rpc-item-main">
                    <strong>${escapeHtml(item.title)}</strong>
                    <span>Draft</span>
                  </div>
                  <span class="rpc-item-meta">${escapeHtml(item.age)}</span>
                </button>
              `).join("")}
              <button type="button" class="rpc-recent-item" data-nav-route="drafts">
                <span></span>
                <div class="rpc-item-main"><strong>Go to All Drafts</strong></div>
                <span class="rpc-arrow-link">›</span>
              </button>
            </div>
          </section>

          <section class="rpc-queue-card">
            <div class="rpc-section-card-head"><h3>Publication Queue</h3><a class="rpc-page-link" href="#/distribution">View all</a></div>
            <div class="rpc-queue-list">
              ${PUBLICATION_QUEUE.map((item) => `
                <div class="rpc-queue-item">
                  ${icon("clock")}
                  <div>
                    <strong>${escapeHtml(item.title)}</strong>
                    <span class="rpc-queue-meta">${escapeHtml(item.desk)}</span>
                  </div>
                  <div>
                    <div class="rpc-queue-state">${escapeHtml(item.state)}</div>
                    <span class="rpc-queue-meta">${escapeHtml(item.analyst)}</span>
                  </div>
                  <div class="rpc-queue-schedule">${escapeHtml(item.date)}<br>${escapeHtml(item.time)}</div>
                </div>
              `).join("")}
              <button type="button" class="rpc-recent-item" data-nav-route="distribution">
                <span></span>
                <div class="rpc-item-main"><strong>Go to Publication Queue</strong></div>
                <span class="rpc-arrow-link">›</span>
              </button>
            </div>
          </section>

          <div style="display:grid;gap:16px;">
            <section class="rpc-readiness-card">
              <div class="rpc-donut" style="--pct:${readinessPercent}%;">
                <div class="rpc-donut-center">
                  <strong>${readinessPercent}%</strong>
                  <span>Ready</span>
                </div>
              </div>
              <div class="rpc-legend">
                <div class="rpc-legend-row"><span class="rpc-legend-dot" style="background:#26a65b;"></span><span>Ready</span><strong>${readyCount}</strong></div>
                <div class="rpc-legend-row"><span class="rpc-legend-dot" style="background:#f6bc33;"></span><span>In Progress</span><strong>${inProgressCount}</strong></div>
                <div class="rpc-legend-row"><span class="rpc-legend-dot" style="background:#dc4f4f;"></span><span>At Risk</span><strong>${atRiskCount}</strong></div>
              </div>
            </section>

            <section class="rpc-validation-card">
              <div class="rpc-validation-summary">
                <div class="rpc-validation-check">✓</div>
                <div>
                  <strong>${validation.valid ? "All Clear" : "Review Needed"}</strong>
                  <span>${validation.valid ? "No blocking issues" : `${validation.missing.length} required items pending`}</span>
                </div>
              </div>
              <div class="rpc-validation-line"><span>Compliance</span><strong>${snapshot.shariah?.tone === "flag" ? 1 : 0}</strong></div>
              <div class="rpc-validation-line"><span>Distribution</span><strong>0</strong></div>
              <div class="rpc-validation-line"><span>Data &amp; Charts</span><strong>${snapshot.review?.findings?.filter((finding) => finding.section === "Exhibits").length || 0}</strong></div>
            </section>
          </div>
        </div>

        <div class="rpc-bottom-row">
          <section class="rpc-calendar-card">
            <div class="rpc-section-card-head"><h3>Upcoming Publications</h3><a class="rpc-page-link" href="#/calendar">View Calendar</a></div>
            <div class="rpc-calendar-shell">
              <div class="rpc-calendar-grid">
                <div class="rpc-calendar-head"><span>May 2025</span><span>23</span></div>
                <div class="rpc-calendar-days"><span>S</span><span>M</span><span>T</span><span>W</span><span>T</span><span>F</span><span>S</span></div>
                <div class="rpc-calendar-dates">
                  ${Array.from({ length: 31 }, (_, index) => `<span class="rpc-calendar-date ${index + 1 === 23 ? "is-active" : ""}">${index + 1}</span>`).join("")}
                </div>
              </div>
              <div class="rpc-schedule-list">
                ${UPCOMING_SCHEDULE.map((item) => `
                  <div class="rpc-schedule-row">
                    <strong>${escapeHtml(item.time)}</strong>
                    <span>${escapeHtml(item.title)}</span>
                    <span>${escapeHtml(item.desk)}</span>
                  </div>
                `).join("")}
              </div>
            </div>
          </section>

          <section class="rpc-quick-access-card">
            <div class="rpc-section-card-head"><h3>Quick Access</h3><a class="rpc-page-link" href="#/live-edit">Create New</a></div>
            <div class="rpc-quick-grid">
              ${QUICK_CREATE.map((item) => `
                <button type="button" class="rpc-quick-card" data-create-note-type="${escapeAttr(item.noteType)}">
                  <div class="rpc-quick-card-icon">${icon(item.icon)}</div>
                  <strong>${escapeHtml(item.label)}</strong>
                  <span>${escapeHtml(item.copy)}</span>
                </button>
              `).join("")}
            </div>
          </section>
        </div>
      </div>
    `;
  }

  function renderEditorPage() {
    const variant = noteVariant(state.snapshot.data.noteType);
    const sectionItems = buildStructureItems();
    ensureSelectedSection();

    const pageHead = variant === "general"
      ? `<div class="rpc-page-head"><div><h1>General Note</h1></div></div>`
      : variant === "alert"
        ? `<div class="rpc-page-head"><div><h1>Short Note / Market Alert</h1></div></div>`
        : "";

    return `
      <div class="rpc-editor-page">
        ${pageHead}
        <div class="rpc-editor-layout">
          <aside class="rpc-structure-panel">
            <div class="rpc-structure-head">
              <div class="rpc-structure-title">
                <strong>${escapeHtml(state.snapshot.data.title || defaultNoteTitleForVariant(variant))}</strong>
                <button type="button" class="rpc-kebab">⋮</button>
              </div>
              <div class="rpc-structure-meta">
                <span style="color:#26a65b;">●</span>
                <span>${escapeHtml(displayWorkflowState())}</span>
                <span>•</span>
                <span>Saved ${escapeHtml(relativeTimestamp(state.snapshot.lastSavedAt))}</span>
              </div>
            </div>
            <div class="rpc-structure-tabs">
              <button type="button" class="rpc-structure-tab ${state.structureTab === "structure" ? "is-active" : ""}" data-structure-tab="structure">Structure</button>
              <button type="button" class="rpc-structure-tab ${state.structureTab === "sections" ? "is-active" : ""}" data-structure-tab="sections">Sections</button>
              <button type="button" class="rpc-kebab" data-add-section>＋</button>
            </div>
            <div class="rpc-structure-list">
              ${sectionItems.map((section, index) => `
                <button
                  type="button"
                  class="rpc-structure-row ${section.key === state.selectedSectionKey ? "is-active" : ""} ${section.complete ? "is-complete" : ""}"
                  data-section-key="${escapeAttr(section.key)}"
                  ${section.reorderable ? `draggable="true" data-draggable-section-key="${escapeAttr(section.key)}"` : ""}
                >
                  <span class="rpc-drag">⋮⋮</span>
                  <span class="rpc-structure-number">${index + 1}</span>
                  <span class="rpc-structure-label">${escapeHtml(section.label)}</span>
                  <span class="rpc-structure-check">${section.complete ? "✓" : ""}</span>
                  <span class="rpc-kebab">⋮</span>
                </button>
              `).join("")}
              <button type="button" class="rpc-add-section-btn" data-add-section>＋ Add Section</button>
            </div>
            <div class="rpc-structure-foot">
              <div class="rpc-structure-foot-row"><span>Word Count</span><strong>${formatNumber(state.snapshot.wordCount)}</strong></div>
              <div class="rpc-structure-foot-row"><span>Reading Time</span><strong>${state.snapshot.readingTimeMinutes} min</strong></div>
            </div>
          </aside>
          <section class="rpc-document-panel">
            ${renderEditorToolbar(variant)}
            <div class="rpc-document-scroll">
              ${renderLiveDocumentCanvas(variant)}
            </div>
          </section>
          <aside class="rpc-context-panel">
            ${renderEditorContextPanel(variant)}
          </aside>
        </div>
      </div>
    `;
  }

  function renderPreviewExportPage() {
    const snapshot = state.snapshot;
    const noteData = snapshot.data;
    const pages = Math.max(1, Math.ceil((snapshot.wordCount || 400) / 120));
    const publicationReady = snapshot.validation.valid && snapshot.shariah?.tone !== "flag";

    return `
      <div class="rpc-page-head">
        <div>
          <h1>Preview / Export</h1>
          <p>Review your final note, validate and prepare for publication.</p>
        </div>
        <div class="rpc-structure-meta">
          <span style="color:#26a65b;">●</span>
          <span>Auto-saved: ${escapeHtml(relativeTimestamp(snapshot.lastSavedAt))}</span>
          <span class="rpc-tag rpc-tag-neutral">All changes saved</span>
        </div>
      </div>
      <div class="rpc-preview-layout">
        <div class="rpc-preview-left">
          <section class="rpc-preview-card rpc-card">
            <div class="rpc-panel-group">
              <h4>Final Validation</h4>
              <div class="rpc-validation-table">
                ${renderFinalValidationLines(snapshot)}
              </div>
            </div>
          </section>
          <section class="rpc-preview-card rpc-card">
            <div class="rpc-panel-group">
              <h4>Distribution Summary</h4>
              <div class="rpc-summary-table">
                <div class="rpc-summary-line"><span>Channels</span><strong>4 Selected</strong></div>
                <div class="rpc-summary-line"><span>Recipients</span><strong>1,246</strong></div>
                <div class="rpc-summary-line"><span>Deliverables</span><strong>5 Formats</strong></div>
              </div>
            </div>
          </section>
          <section class="rpc-preview-card rpc-card">
            <div class="rpc-panel-group">
              <h4>Target Publication</h4>
              <div class="rpc-summary-table">
                <div class="rpc-summary-line"><span>Date</span><strong>${escapeHtml(snapshot.publicationDateLabel || formatDateShort(noteData.publicationDate))}</strong></div>
                <div class="rpc-summary-line"><span>State</span><strong>${escapeHtml(noteData.distributionPublicationState || "Draft")}</strong></div>
                <div class="rpc-summary-line"><span>Classification</span><strong>${escapeHtml(noteData.accessClassification || "Internal")}</strong></div>
              </div>
            </div>
          </section>
        </div>

        <div class="rpc-preview-center">
          <section class="rpc-preview-card rpc-card">
            <div class="rpc-preview-toolbar">
              <div class="rpc-toolbar-group">
                <button type="button" class="rpc-device-toggle ${state.previewDevice === "desktop" ? "is-active" : ""}" data-preview-device="desktop">${icon("desktop")}</button>
                <button type="button" class="rpc-device-toggle ${state.previewDevice === "mobile" ? "is-active" : ""}" data-preview-device="mobile">${icon("mobile")}</button>
              </div>
              <div class="rpc-toolbar-group">
                <select class="rpc-toolbar-select is-short"><option>Page view</option></select>
                <span>${1} / ${pages}</span>
              </div>
            </div>
            <iframe class="rpc-preview-iframe" title="Note preview" data-preview-frame></iframe>
          </section>
        </div>

        <div class="rpc-preview-right">
          <section class="rpc-preview-card rpc-card">
            <div class="rpc-panel-group">
              <h4>Export Controls</h4>
              <div class="rpc-output-actions">
                <div class="rpc-export-grid">
                  <button type="button" class="rpc-export-btn" data-export-kind="word">Word</button>
                  <button type="button" class="rpc-export-btn" data-export-kind="pdf">PDF</button>
                  <button type="button" class="rpc-export-btn" data-export-kind="powerpoint">PowerPoint</button>
                  <button type="button" class="rpc-export-btn" data-export-kind="excel">Excel</button>
                  <button type="button" class="rpc-export-btn" data-open-summary>Website Summary</button>
                </div>
              </div>
            </div>
            <div class="rpc-panel-group">
              <h4>Content Options</h4>
              <div class="rpc-toggle-row"><span>Include disclosures</span><button type="button" class="rpc-switch ${state.exportOptions.includeDisclosures ? "is-on" : ""}" data-toggle-export="includeDisclosures"></button></div>
              <div class="rpc-toggle-row"><span>Include analyst bio</span><button type="button" class="rpc-switch ${state.exportOptions.includeAnalystBio ? "is-on" : ""}" data-toggle-export="includeAnalystBio"></button></div>
              <div class="rpc-toggle-row"><span>Include related research</span><button type="button" class="rpc-switch ${state.exportOptions.includeRelatedResearch ? "is-on" : ""}" data-toggle-export="includeRelatedResearch"></button></div>
            </div>
            <div class="rpc-panel-group">
              <h4>Distribution Channels</h4>
              <label class="rpc-check"><input type="checkbox" ${state.exportOptions.channels.researchPortal ? "checked" : ""} data-channel-toggle="researchPortal"><span>Research Portal</span></label>
              <label class="rpc-check"><input type="checkbox" ${state.exportOptions.channels.clientAlert ? "checked" : ""} data-channel-toggle="clientAlert"><span>Client Alert</span></label>
              <label class="rpc-check"><input type="checkbox" ${state.exportOptions.channels.emailAlert ? "checked" : ""} data-channel-toggle="emailAlert"><span>Email Alert</span></label>
              <label class="rpc-check"><input type="checkbox" ${state.exportOptions.channels.bloombergApi ? "checked" : ""} data-channel-toggle="bloombergApi"><span>Bloomberg / API</span></label>
            </div>
          </section>
          <section class="rpc-readiness-banner">
            <span>${publicationReady ? "Publication Ready" : "Needs Review"}</span>
            <strong>${publicationReady ? "Publication Ready" : "Hold for Review"}</strong>
            <span>${publicationReady ? "All checks passed" : "Resolve the remaining validation items before release."}</span>
          </section>
          <section class="rpc-preview-card rpc-card">
            <div class="rpc-panel-group">
              <h4>Note Summary</h4>
              <div class="rpc-summary-table">
                <div class="rpc-summary-line"><span>Title</span><strong>${escapeHtml(noteData.title || "Untitled note")}</strong></div>
                <div class="rpc-summary-line"><span>Section Count</span><strong>${state.snapshot.sectionLayout.length}</strong></div>
                <div class="rpc-summary-line"><span>Word Count</span><strong>${formatNumber(state.snapshot.wordCount)}</strong></div>
                <div class="rpc-summary-line"><span>Reading Time</span><strong>${state.snapshot.readingTimeMinutes} min</strong></div>
                <div class="rpc-summary-line"><span>Pages</span><strong>${pages}</strong></div>
                <div class="rpc-summary-line"><span>Last Edited By</span><strong>${escapeHtml(state.user.fullName)}</strong></div>
              </div>
            </div>
          </section>
          <div class="rpc-action-stack">
            <button type="button" class="rpc-primary-btn" data-generate-word>Export &amp; Distribute</button>
            <button type="button" class="rpc-secondary-btn" data-save-final>Save as Final Draft</button>
          </div>
        </div>
      </div>
    `;
  }

  function renderSettingsPage() {
    const settings = state.settingsDraft;
    return `
      <div class="rpc-page-head">
        <div>
          <h1>Settings</h1>
          <p>Manage your profile, preferences, security, and notifications.</p>
        </div>
      </div>
      <div class="rpc-settings-grid">
        <section class="rpc-settings-card">
          <div>
            <h3>Analyst Profile</h3>
            <p>Update your personal and professional information.</p>
          </div>
          <div class="rpc-form-group">
            <label>Full name</label>
            <input type="text" value="${escapeAttr(settings.profile.fullName)}" data-setting-path="profile.fullName">
          </div>
          <div class="rpc-form-group">
            <label>Email address</label>
            <input type="email" value="${escapeAttr(settings.profile.email)}" disabled>
          </div>
          <div class="rpc-form-group">
            <label>Phone number</label>
            <input type="text" value="${escapeAttr(settings.profile.phoneNumber)}" data-setting-path="profile.phoneNumber">
          </div>
          <div class="rpc-form-group">
            <label>Job title / desk</label>
            <input type="text" value="${escapeAttr(settings.profile.jobTitle)}" data-setting-path="profile.jobTitle">
          </div>
        </section>

        <section class="rpc-settings-card">
          <div>
            <h3>Defaults</h3>
            <p>Set your default preferences for content and coverage.</p>
          </div>
          <div class="rpc-form-group">
            <label>Default note type</label>
            <select data-setting-path="defaults.defaultNoteType">
              ${workspaceOptions().map((option) => `<option value="${escapeAttr(option.label)}" ${option.label === settings.defaults.defaultNoteType ? "selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
            </select>
          </div>
          <div class="rpc-form-group">
            <label>Default region / coverage area</label>
            <select data-setting-path="defaults.defaultRegion">
              ${["Emerging Markets", "Global", "Middle East", "Central Asia", "South Asia", "Commodities"].map((option) => `<option value="${escapeAttr(option)}" ${option === settings.defaults.defaultRegion ? "selected" : ""}>${escapeHtml(option)}</option>`).join("")}
            </select>
          </div>
        </section>

        <section class="rpc-settings-card">
          <div>
            <h3>Security</h3>
            <p>Manage your password, authentication, and active sessions.</p>
          </div>
          <div class="rpc-settings-line">
            <div>
              <strong>Password</strong>
              <p>Last changed on ${escapeHtml(formatLongDate(settings.security.passwordChangedAt))}</p>
            </div>
            <button type="button" class="rpc-secondary-btn" data-change-password>Change Password</button>
          </div>
          <div class="rpc-settings-line">
            <div>
              <strong>Two-factor authentication</strong>
              <p>Add an extra layer of security to your account.</p>
            </div>
            <span class="rpc-settings-badge">${settings.security.twoFactorEnabled ? "Enabled" : "Disabled"}</span>
          </div>
          <div class="rpc-settings-line">
            <div>
              <strong>Active sessions</strong>
              <p>You are currently signed in on ${settings.security.activeSessions} device${settings.security.activeSessions === 1 ? "" : "s"}.</p>
            </div>
            <button type="button" class="rpc-secondary-btn" data-manage-sessions>Manage Sessions</button>
          </div>
        </section>

        <section class="rpc-settings-card">
          <div>
            <h3>Notifications</h3>
            <p>Choose what you want to be notified about.</p>
          </div>
          <div class="rpc-checklist">
            ${renderNotificationRow("Publication reminders", "Receive reminders for upcoming publication deadlines.", "notifications.publicationReminders", settings.notifications.publicationReminders)}
            ${renderNotificationRow("Draft activity", "Get notified when someone comments or updates a draft you follow.", "notifications.draftActivity", settings.notifications.draftActivity)}
            ${renderNotificationRow("Validation alerts", "Receive alerts for validation issues or compliance checks.", "notifications.validationAlerts", settings.notifications.validationAlerts)}
            ${renderNotificationRow("System updates", "Receive important updates about platform changes and maintenance.", "notifications.systemUpdates", settings.notifications.systemUpdates)}
          </div>
        </section>
      </div>
      <div class="rpc-settings-actions">
        <button type="button" class="rpc-secondary-btn" data-reset-settings>Cancel</button>
        <button type="button" class="rpc-primary-btn" data-save-settings>Save Changes</button>
      </div>
    `;
  }

  function renderPlaceholderPage(title, copy) {
    return `
      <div class="rpc-page-head">
        <div><h1>${escapeHtml(title)}</h1></div>
      </div>
      <section class="rpc-placeholder-card">
        <h3>${escapeHtml(title)}</h3>
        <p>${escapeHtml(copy)}</p>
      </section>
    `;
  }

  function renderEditorToolbar(variant) {
    return `
      <div class="rpc-toolbar">
        ${variant === "alert" ? `
          <div class="rpc-toolbar-group">
            <select class="rpc-toolbar-select is-font"><option>Inter</option></select>
            <select class="rpc-toolbar-select is-short"><option>14</option></select>
          </div>
        ` : `
          <div class="rpc-toolbar-group">
            <select class="rpc-toolbar-select is-font"><option>Heading 1</option></select>
            <select class="rpc-toolbar-select is-font"><option>Inter</option></select>
            <select class="rpc-toolbar-select is-short"><option>22</option></select>
          </div>
        `}
        <div class="rpc-toolbar-divider"></div>
        <div class="rpc-toolbar-group">
          ${toolbarButton("bold", "B")}
          ${toolbarButton("italic", "I")}
          ${toolbarButton("underline", "U")}
          ${toolbarButton("strikeThrough", "S")}
        </div>
        <div class="rpc-toolbar-divider"></div>
        <div class="rpc-toolbar-group">
          ${toolbarButton("justifyLeft", icon("align-left"), true)}
          ${toolbarButton("justifyCenter", icon("align-center"), true)}
          ${toolbarButton("justifyRight", icon("align-right"), true)}
          ${toolbarButton("insertUnorderedList", icon("list"), true)}
        </div>
        <div class="rpc-toolbar-divider"></div>
        <div class="rpc-toolbar-group">
          ${toolbarButton("createLink", icon("link"), true)}
          ${toolbarButton("insertHTML", icon("table"), true, "table")}
          ${toolbarButton("insertHTML", "⋮", false, "ellipsis")}
        </div>
      </div>
    `;
  }

  function renderLiveDocumentCanvas(variant) {
    const snapshot = state.snapshot;
    const data = snapshot.data;
    const authorLine = snapshot.authorLine || buildAuthorLineFromData(data);
    const keyTakeaways = getSectionByKey("keyTakeaways");
    const analysis = getSectionByKey("analysis");
    const content = getSectionByKey("content");
    const cordobaView = getSectionByKey("cordobaView");
    const disclosureSection = findDisclosureSection();

    if (variant === "alert") {
      return `
        <article class="rpc-note-sheet">
          <div class="rpc-note-meta-top">
            <div class="rpc-note-kicker">
              <span>MACRO / RATES</span>
              <span class="rpc-alert-badge">MARKET ALERT</span>
            </div>
            <div class="rpc-note-inline">${escapeHtml(formatDateWithTime(data.publicationDate))}</div>
          </div>
          <div class="rpc-note-title" contenteditable="true" data-field-id="title">${escapeHtml(data.title || "US CPI Surprise Sparks Rates Volatility")}</div>
          <div class="rpc-note-deck" contenteditable="true" data-field-id="deck">${escapeHtml(data.deck || "Stronger-than-expected inflation pressures reinvigorate rate hike pricing and trigger sharp moves across Treasuries and risk assets.")}</div>
          <div class="rpc-note-divider"></div>
          ${renderRichSectionBlock("keyTakeaways", keyTakeaways?.label || "Key Takeaways", keyTakeaways?.content || "")}
          ${renderRichSectionBlock("analysis", "What Happened", analysis?.content || "")}
          ${renderRichSectionBlock("content", "Why It Matters", content?.content || "")}
          ${renderMarketImpactBlock()}
          ${renderRichSectionBlock("cordobaView", cordobaView?.label || "Cordoba View", cordobaView?.content || "")}
          ${renderDisclosureBlock(disclosureSection)}
          <div class="rpc-note-inline" style="margin-top:18px;text-align:right;color:#7c8ba6;">${formatNumber(snapshot.wordCount)} words</div>
        </article>
      `;
    }

    if (variant === "general") {
      return `
        <article class="rpc-note-sheet">
          ${renderStandardNoteHeader(authorLine)}
          ${renderCalloutBlock("keyTakeaways", keyTakeaways?.label || "Key Takeaways", keyTakeaways?.content || "")}
          ${renderRichSectionBlock("analysis", analysis?.label || "Analysis and Commentary", analysis?.content || "")}
          ${renderRichSectionBlock("content", content?.label || "Additional Detail", content?.content || "")}
          ${renderFiguresBlock()}
          ${renderRichSectionBlock("cordobaView", cordobaView?.label || "Cordoba View", cordobaView?.content || "")}
          ${renderWebsiteSummaryModule()}
          ${renderDisclosureBlock(disclosureSection)}
        </article>
      `;
    }

    return `
      <article class="rpc-note-sheet">
        ${renderStandardNoteHeader(authorLine)}
        ${renderCalloutBlock("keyTakeaways", keyTakeaways?.label || "Key Takeaways", keyTakeaways?.content || "")}
        ${renderRichSectionBlock("analysis", analysis?.label || "Analysis", analysis?.content || "")}
        ${renderMacroExhibits()}
        ${renderRichSectionBlock("content", content?.label || "Additional Detail", content?.content || "")}
        ${renderFiguresBlock()}
        ${renderRichSectionBlock("cordobaView", cordobaView?.label || "Cordoba View", cordobaView?.content || "")}
        ${renderDisclosureBlock(disclosureSection)}
      </article>
    `;
  }

  function renderStandardNoteHeader(authorLine) {
    const data = state.snapshot.data;
    return `
      <div class="rpc-note-meta-top">
        <div class="rpc-note-kicker">
          <span contenteditable="true" data-field-id="topic">${escapeHtml(data.topic || workspaceLabelForNoteType(data.noteType))}</span>
          <span>|</span>
          <span>${escapeHtml(data.metadataPrimaryGeography || defaultGeographyForNoteType(data.noteType))}</span>
        </div>
        <div class="rpc-note-inline">${escapeHtml(formatDateLong(data.publicationDate))}</div>
      </div>
      <div class="rpc-note-title" contenteditable="true" data-field-id="title">${escapeHtml(data.title || defaultNoteTitleForVariant(noteVariant(data.noteType)))}</div>
      <div class="rpc-note-deck" contenteditable="true" data-field-id="deck">${escapeHtml(data.deck || "Reforms and resilience support gradual external improvement")}</div>
      <div class="rpc-note-author-line" contenteditable="true" data-author-line>${escapeHtml(authorLine)}</div>
      <div class="rpc-note-desk">${escapeHtml(data.deskLine || "Sovereign Research Desk")}</div>
      <div class="rpc-note-divider"></div>
    `;
  }

  function renderCalloutBlock(sectionKey, label, content) {
    return `
      <section class="rpc-note-block rpc-callout ${sectionKey === state.selectedSectionKey ? "is-active" : ""}" data-block-key="${escapeAttr(sectionKey)}">
        <div class="rpc-note-block-head">
          <div class="rpc-note-block-title" contenteditable="true" data-heading-section-key="${escapeAttr(sectionKey)}">${escapeHtml(label)}</div>
        </div>
        <div class="rpc-note-rich" contenteditable="true" data-rich-section-key="${escapeAttr(sectionKey)}">${state.api.renderRichTextHtml(content)}</div>
      </section>
    `;
  }

  function renderRichSectionBlock(sectionKey, label, content) {
    return `
      <section class="rpc-note-block ${sectionKey === state.selectedSectionKey ? "is-active" : ""}" data-block-key="${escapeAttr(sectionKey)}">
        <div class="rpc-note-block-head">
          ${showSectionNumber(sectionKey) ? `<span class="rpc-note-block-number">${sectionNumberForKey(sectionKey)}</span>` : ""}
          <div class="rpc-note-block-title" contenteditable="true" data-heading-section-key="${escapeAttr(sectionKey)}">${escapeHtml(label)}</div>
        </div>
        <div class="rpc-note-rich" contenteditable="true" data-rich-section-key="${escapeAttr(sectionKey)}">${state.api.renderRichTextHtml(content)}</div>
      </section>
    `;
  }

  function renderMarketImpactBlock() {
    const blockKey = "market-impact";
    const chips = [
      { label: "UST 10Y Yield", value: "+12.6bps", tone: "down" },
      { label: "2s10s Curve", value: "-8.2bps", tone: "down" },
      { label: "DXY Index", value: "+0.68%", tone: "up" },
      { label: "S&P 500 Fut.", value: "-0.74%", tone: "down" },
      { label: "Gold (Spot)", value: "-0.41%", tone: "down" }
    ];

    return `
      <section class="rpc-note-block ${blockKey === state.selectedSectionKey ? "is-active" : ""}" data-block-key="${blockKey}">
        <div class="rpc-note-block-head">
          <div class="rpc-note-block-title">Market Impact</div>
        </div>
        <div class="rpc-market-impact">
          ${chips.map((chip) => `
            <div class="rpc-impact-chip ${chip.tone === "up" ? "is-up" : "is-down"}">
              <strong>${escapeHtml(chip.label)}</strong>
              <span>${escapeHtml(chip.value)}</span>
            </div>
          `).join("")}
        </div>
      </section>
    `;
  }

  function renderMacroExhibits() {
    const ratingsRows = state.snapshot.data.ratingsProfile || [];
    const sovereignRows = ratingsRows.length ? ratingsRows : [
      { agency: "Moody's", shortTerm: "P-3", longTerm: "Baa3" },
      { agency: "S&P", shortTerm: "B", longTerm: "BB-" },
      { agency: "Fitch", shortTerm: "B", longTerm: "BB-" }
    ];

    return `
      <div class="rpc-two-column-exhibits">
        <div class="rpc-exhibit-card">
          <h4>Sovereign Ratings</h4>
          <table class="rpc-mini-table">
            <thead>
              <tr><th>Agency</th><th>Rating</th><th>Outlook</th><th>Last Action</th></tr>
            </thead>
            <tbody>
              ${sovereignRows.slice(0, 3).map((row, index) => `
                <tr>
                  <td>${escapeHtml(row.agency || "Agency")}</td>
                  <td>${escapeHtml(row.longTerm || row.shortTerm || "—")}</td>
                  <td>${escapeHtml(["Stable", "Positive", "Stable"][index] || "Stable")}</td>
                  <td>${escapeHtml(["Nov 2024", "Mar 2025", "Feb 2025"][index] || "May 2025")}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
        <div class="rpc-exhibit-card">
          <h4>USD Sovereign Yield Curve</h4>
          <div class="rpc-exhibit-sub">${escapeHtml(state.snapshot.data.coverageCountry || "Uzbekistan")} Eurobonds</div>
          <div class="rpc-chart-placeholder">
            <div class="rpc-chart-line">
              <svg viewBox="0 0 400 126" preserveAspectRatio="none" aria-hidden="true">
                <polyline fill="none" stroke="#c88c12" stroke-width="3" points="18,86 70,74 122,70 174,68 226,64 278,54 330,34 382,18"></polyline>
                <polyline fill="none" stroke="#7c8ba6" stroke-width="3" stroke-dasharray="6 6" points="18,102 70,96 122,92 174,86 226,78 278,70 330,58 382,44"></polyline>
              </svg>
            </div>
          </div>
        </div>
      </div>
    `;
  }

  function renderFiguresBlock() {
    const figures = state.snapshot.figures || [];
    const blockKey = "figures";

    if (!figures.length) {
      return `
        <section class="rpc-note-block ${blockKey === state.selectedSectionKey ? "is-active" : ""}" data-block-key="${blockKey}">
          <div class="rpc-note-block-head">
            <div class="rpc-note-block-title">Charts / Figures</div>
          </div>
          <div class="rpc-note-rich">
            <p>No figures inserted yet.</p>
            <button type="button" class="rpc-secondary-btn" data-trigger-figures>Add Figure</button>
          </div>
        </section>
      `;
    }

    return `
      <section class="rpc-note-block ${blockKey === state.selectedSectionKey ? "is-active" : ""}" data-block-key="${blockKey}">
        <div class="rpc-note-block-head">
          <div class="rpc-note-block-title">Charts / Figures</div>
        </div>
        <div class="rpc-note-figures">
          ${figures.map((figure, index) => {
            const detail = figure.detail || {};
            const heading = `${detail.labelPrefix || "Figure"} ${detail.labelNumber || index + 1}: ${detail.title || detail.caption || figure.fileName}`;
            return `
              <figure class="rpc-note-figure" data-figure-key="${escapeAttr(figure.key)}">
                <div class="rpc-note-figure-head" contenteditable="true" data-figure-field="title" data-figure-key="${escapeAttr(figure.key)}">${escapeHtml(heading)}</div>
                ${detail.subtitle ? `<div class="rpc-note-figure-foot" contenteditable="true" data-figure-field="subtitle" data-figure-key="${escapeAttr(figure.key)}">${escapeHtml(detail.subtitle)}</div>` : ""}
                <div class="rpc-note-figure-media">
                  <img src="${escapeAttr(figure.previewUrl)}" alt="${escapeAttr(detail.title || figure.fileName)}">
                </div>
                <div class="rpc-note-figure-foot" contenteditable="true" data-figure-field="source" data-figure-key="${escapeAttr(figure.key)}">${escapeHtml(detail.source || "Source: Cordoba Research Group")}</div>
                ${detail.note ? `<div class="rpc-note-figure-foot" contenteditable="true" data-figure-field="note" data-figure-key="${escapeAttr(figure.key)}">${escapeHtml(`Note: ${detail.note}`)}</div>` : ""}
              </figure>
            `;
          }).join("")}
        </div>
      </section>
    `;
  }

  function renderWebsiteSummaryModule() {
    const data = state.snapshot.data;
    const title = data.title || "Global Growth Outlook 2H 2025";
    const slug = slugify(title);
    return `
      <section class="rpc-website-summary ${state.selectedSectionKey === "website-summary" ? "is-active" : ""}" data-block-key="website-summary">
        <div class="rpc-website-summary-head">
          <div>
            <strong>Website Summary</strong>
          </div>
          <button type="button" class="rpc-secondary-btn" data-open-summary>Preview on Website</button>
        </div>
        <div class="rpc-website-summary-meta">
          <div><span>Proposed Title</span><strong>${escapeHtml(title)}</strong></div>
          <div><span>Proposed Slug</span><strong>${escapeHtml(slug)}</strong></div>
          <div><span>Audience</span><strong>${escapeHtml(data.metadataAudience || "Institutional Investors")}</strong></div>
          <div><span>Publish Window</span><strong>${escapeHtml(formatPublishWindow(data.publicationDate))}</strong></div>
        </div>
      </section>
    `;
  }

  function renderDisclosureBlock(disclosureSection) {
    if (disclosureSection) {
      return renderRichSectionBlock(disclosureSection.key, disclosureSection.label || "Disclosures", disclosureSection.content || "");
    }

    return `
      <section class="rpc-note-block ${state.selectedSectionKey === "disclosures" ? "is-active" : ""}" data-block-key="disclosures">
        <div class="rpc-note-block-head">
          <div class="rpc-note-block-title">Disclosures</div>
        </div>
        <div class="rpc-note-rich">
          <p>${escapeHtml(defaultDisclosureCopy())}</p>
        </div>
      </section>
    `;
  }

  function renderEditorContextPanel(variant) {
    if (variant === "alert") {
      return `
        <div class="rpc-context-tabs">
          ${contextTabButton("publish", "Publish", state.publishTab)}
          ${contextTabButton("validate", "Validate", state.publishTab)}
          ${contextTabButton("history", "History", state.publishTab)}
        </div>
        <div class="rpc-context-scroll">
          ${state.publishTab === "publish" ? renderAlertPublishPanel() : state.publishTab === "validate" ? renderValidatePanel() : renderHistoryPanel()}
        </div>
      `;
    }

    return `
      <div class="rpc-context-tabs">
        ${contextTabButton("tools", "Tools", state.contextTab)}
        ${contextTabButton("validate", "Validate", state.contextTab)}
        ${contextTabButton("history", "History", state.contextTab)}
      </div>
      <div class="rpc-context-scroll">
        ${state.contextTab === "tools" ? renderToolsPanel(variant) : state.contextTab === "validate" ? renderValidatePanel() : renderHistoryPanel()}
      </div>
    `;
  }

  function renderToolsPanel(variant) {
    const selected = buildStructureItems().find((item) => item.key === state.selectedSectionKey) || { key: "header", label: "Header / Metadata" };
    const settings = sectionSettingsForKey(selected.key);
    const shariah = state.snapshot.shariah || { tone: "review", label: "Shariah: review", detail: "Not assessed" };
    const files = state.snapshot.data.modelFiles || [];
    const tags = parseTags(state.snapshot.data.metadataSearchTags);

    return `
      <div class="rpc-panel-group">
        <h4>Section Settings</h4>
        <div class="rpc-right-field">
          <label>${variant === "general" ? "Section Type" : "Section Settings"}</label>
          <select class="rpc-right-select">
            <option>${escapeHtml(selected.label)}</option>
          </select>
        </div>
        <div class="rpc-toggle-row">
          <span>Show section number</span>
          <button type="button" class="rpc-switch ${settings.showNumber ? "is-on" : ""}" data-toggle-section-setting="showNumber"></button>
        </div>
        <div class="rpc-toggle-row">
          <span>Collapsible in view</span>
          <button type="button" class="rpc-switch ${settings.collapsible ? "is-on" : ""}" data-toggle-section-setting="collapsible"></button>
        </div>
      </div>

      <div class="rpc-panel-group">
        <h4>Formatting</h4>
        <div class="rpc-format-grid">
          ${toolbarButton("bold", "B")}
          ${toolbarButton("italic", "I")}
          ${toolbarButton("underline", "U")}
          ${toolbarButton("insertUnorderedList", icon("list"), true)}
          ${toolbarButton("justifyLeft", icon("align-left"), true)}
          ${toolbarButton("justifyRight", icon("align-right"), true)}
          <button type="button" class="rpc-toolbar-btn" data-format-block="h1">H1</button>
          <button type="button" class="rpc-toolbar-btn" data-format-block="h2">H2</button>
          <button type="button" class="rpc-toolbar-btn" data-format-block="h3">H3</button>
        </div>
      </div>

      <div class="rpc-panel-group">
        <h4>Insert</h4>
        <div class="rpc-insert-grid">
          <button type="button" class="rpc-insert-btn" data-insert-block="table">Table</button>
          <button type="button" class="rpc-insert-btn" data-trigger-figures>Chart</button>
          <button type="button" class="rpc-insert-btn" data-insert-block="textbox">Text Box</button>
          <button type="button" class="rpc-insert-btn" data-trigger-figures>Image</button>
          <button type="button" class="rpc-insert-btn" data-insert-block="callout">Callout</button>
          <button type="button" class="rpc-insert-btn" data-insert-block="divider">Divider</button>
        </div>
      </div>

      ${variant === "general" ? `
        <div class="rpc-panel-group">
          <h4>Tags</h4>
          <div class="rpc-tag-row">
            ${(tags.length ? tags : ["Macro", "Global Growth", "Outlook", "2H 2025", "Forecast", "Policy"]).map((tag, index) => `
              <span class="rpc-tag ${index > 1 ? "rpc-tag-neutral" : ""}">${escapeHtml(tag)}</span>
            `).join("")}
          </div>
        </div>
      ` : ""}

      ${(variant === "general" || variant === "macro") ? `
        <div class="rpc-panel-group">
          <h4>Supporting Files</h4>
          <div class="rpc-support-list">
            ${(files.length ? files : [
              { name: "IMF_WEO_Apr2025.pdf", size: "2.4 MB" },
              { name: "OECD_Economic_Outlook.pdf", size: "1.8 MB" },
              { name: "World_Bank_Global_PEC.pdf", size: "3.1 MB" }
            ]).map((file) => `
              <div class="rpc-support-item">
                ${icon("doc")}
                <div class="rpc-item-main"><strong>${escapeHtml(file.name)}</strong></div>
                <span class="rpc-item-meta">${escapeHtml(file.size || "")}</span>
              </div>
            `).join("")}
            <button type="button" class="rpc-secondary-btn" data-trigger-files>Add File</button>
          </div>
        </div>
      ` : ""}

      ${variant === "general" ? `
        <div class="rpc-panel-group">
          <h4>Pre-Publish Review</h4>
          <div class="rpc-validation-table">
            ${renderStatusRow("Content Quality", "Good")}
            ${renderStatusRow("Data & Sources", "Good")}
            ${renderStatusRow("Compliance", state.snapshot.shariah?.tone === "flag" ? "Flagged" : "Clear")}
            ${renderStatusRow("Distribution Readiness", state.snapshot.validation.valid ? "Ready" : "Review")}
          </div>
        </div>
      ` : `
        <div class="rpc-panel-group">
          <h4>Validation Checks</h4>
          <div class="rpc-validation-table">
            ${renderStatusRow("Publication Readiness", state.snapshot.validation.valid ? "Good" : "Review")}
            ${renderStatusRow("Missing Items", state.snapshot.validation.missing.length ? String(state.snapshot.validation.missing.length) : "None")}
            ${renderStatusRow("Compliance", shariah.tone === "flag" ? "Flagged" : "Clear")}
            ${renderStatusRow("Distribution", "Valid")}
          </div>
        </div>
      `}

      <div class="rpc-panel-group">
        <h4>${variant === "general" ? "Output Actions" : "Export"}</h4>
        <div class="rpc-export-grid">
          <button type="button" class="rpc-export-btn" data-export-kind="pdf">PDF</button>
          <button type="button" class="rpc-export-btn" data-export-kind="word">Word</button>
          <button type="button" class="rpc-export-btn" data-export-kind="powerpoint">PowerPoint</button>
          <button type="button" class="rpc-export-btn" data-export-kind="excel">Excel</button>
        </div>
        <button type="button" class="rpc-secondary-btn" data-nav-route="preview-export">Preview Note</button>
      </div>

      <div class="rpc-notice-card ${shariah.tone === "flag" ? "is-flag" : ""}">
        <strong>${escapeHtml(shariah.tone === "flag" ? "Compliance Flag" : "Compliance Status")}</strong>
        <span>${escapeHtml(shariah.tone === "flag" ? `${shariah.label}. ${shariah.detail}.` : "No compliance issues detected.")}</span>
      </div>
    `;
  }

  function renderAlertPublishPanel() {
    return `
      <div class="rpc-panel-group">
        <h4>Publication Timing</h4>
        <label class="rpc-check"><input type="radio" checked><span>Publish Now</span></label>
        <label class="rpc-check"><input type="radio"><span>Schedule</span></label>
        <div class="rpc-toolbar-group">
          <input class="rpc-publish-date" type="text" value="${escapeAttr(formatDateLong(state.snapshot.data.publicationDate))}">
          <input class="rpc-publish-time" type="text" value="11:00 AM ET">
        </div>
        <label class="rpc-check"><input type="radio"><span>Hold for Review</span></label>
      </div>
      <div class="rpc-panel-group">
        <h4>Distribution</h4>
        <select class="rpc-publish-select"><option>Select Distribution List</option></select>
        <div class="rpc-tag-row">
          <span class="rpc-tag rpc-tag-neutral">Internal Research ×</span>
          <span class="rpc-tag rpc-tag-neutral">Client Macro Alerts ×</span>
          <span class="rpc-tag rpc-tag-neutral">Sales Desk ×</span>
        </div>
        <div class="rpc-status-good">Distribution valid</div>
      </div>
      <div class="rpc-panel-group">
        <h4>Attachments</h4>
        <button type="button" class="rpc-secondary-btn" data-trigger-files>Drag files here or click to upload</button>
      </div>
      <div class="rpc-panel-group">
        <h4>Preview</h4>
        <section class="rpc-support-item" style="align-items:center;">
          ${icon("alert")}
          <div class="rpc-item-main">
            <strong>${escapeHtml(state.snapshot.data.title || "US CPI Surprise Sparks Rates Volatility")}</strong>
            <span>${escapeHtml(state.snapshot.data.deck || "Stronger-than-expected inflation pressures reinvigorate rate hike pricing and trigger sharp moves across Treasuries and risk assets.")}</span>
          </div>
        </section>
      </div>
      <div class="rpc-panel-group">
        <h4>Export</h4>
        <div class="rpc-export-grid">
          <button type="button" class="rpc-export-btn" data-export-kind="pdf">PDF</button>
          <button type="button" class="rpc-export-btn" data-export-kind="word">Word</button>
          <button type="button" class="rpc-export-btn" data-export-kind="powerpoint">PowerPoint</button>
          <button type="button" class="rpc-export-btn" data-export-kind="excel">Excel</button>
          <button type="button" class="rpc-export-btn" data-export-kind="html">HTML</button>
        </div>
      </div>
      <div class="rpc-panel-group">
        <div class="rpc-notice-card"><strong>Compliance clear</strong><span>No compliance issues detected.</span></div>
        <div class="rpc-notice-card"><strong>Distribution valid</strong><span>Selected channels are ready.</span></div>
      </div>
    `;
  }

  function renderValidatePanel() {
    const findings = state.snapshot.review?.findings || [];
    const shariah = state.snapshot.shariah || { tone: "review", label: "Shariah: review", detail: "Not assessed" };
    return `
      <div class="rpc-panel-group">
        <h4>Validation Checks</h4>
        <div class="rpc-validation-table">
          ${renderStatusRow("Required Fields", state.snapshot.validation.valid ? "Passed" : `${state.snapshot.validation.missing.length} Open`)}
          ${renderStatusRow("Data & Charts", findings.filter((item) => item.section === "Exhibits").length ? "Info" : "Passed")}
          ${renderStatusRow("Disclosures", findDisclosureSection() ? "Present" : "Review")}
          ${renderStatusRow("Compliance", shariah.tone === "flag" ? "Flagged" : "Clear")}
        </div>
      </div>
      ${findings.length ? `
        <div class="rpc-panel-group">
          <h4>Open Items</h4>
          <div class="rpc-validation-table">
            ${findings.slice(0, 8).map((finding) => `
              <div class="rpc-validation-row">
                <span>${escapeHtml(finding.title)}</span>
                <strong class="${finding.severity === "critical" ? "rpc-status-danger" : finding.severity === "warning" ? "rpc-status-warn" : "rpc-status-good"}">${escapeHtml(finding.severity)}</strong>
              </div>
            `).join("")}
          </div>
        </div>
      ` : ""}
    `;
  }

  function renderHistoryPanel() {
    const history = state.snapshot.history || [];
    return `
      <div class="rpc-panel-group">
        <h4>History</h4>
        <div class="rpc-validation-table">
          ${history.length ? history.slice(0, 8).map((entry) => `
            <div class="rpc-validation-row">
              <span>${escapeHtml(entry.title)}</span>
              <strong>${escapeHtml(relativeTimestamp(entry.timestamp))}</strong>
            </div>
          `).join("") : `<div class="rpc-validation-row"><span>No workflow events yet.</span><strong>—</strong></div>`}
        </div>
      </div>
    `;
  }

  function handleAuthClick(event) {
    const toggleButton = event.target.closest("[data-toggle-password]");
    if (toggleButton) {
      state.showPassword = !state.showPassword;
      renderAuth();
      return;
    }

    const ssoButton = event.target.closest("[data-sso-button]");
    if (ssoButton) {
      showToast("Microsoft SSO is enabled in the connected deployment environment.", "info");
    }
  }

  async function handleAuthSubmit(event) {
    if (!event.target.matches("[data-auth-form]")) return;
    event.preventDefault();

    const formData = new FormData(event.target);
    try {
      const session = await fetchJson("/api/auth/login", {
        method: "POST",
        body: JSON.stringify({
          email: formData.get("email"),
          password: formData.get("password"),
          remember: formData.get("remember") === "on"
        })
      });
      state.user = session.user;
      state.security = session.security;
      await hydrateSettings();
      ensureSelectedSection();
      renderShell();
    } catch (error) {
      renderAuth(friendlyAuthError(error));
    }
  }

  function handlePlatformClick(event) {
    const routeButton = event.target.closest("[data-nav-route]");
    if (routeButton) {
      setRoute(routeButton.getAttribute("data-nav-route"));
      return;
    }

    const createButton = event.target.closest("[data-create-note-type]");
    if (createButton) {
      switchNoteType(createButton.getAttribute("data-create-note-type"));
      setRoute("live-edit");
      return;
    }

    const userMenuButton = event.target.closest("[data-open-user-menu]");
    if (userMenuButton) {
      state.userMenuOpen = !state.userMenuOpen;
      renderShell();
      return;
    }

    const profileAction = event.target.closest("[data-profile-action='logout']");
    if (profileAction) {
      logout();
      return;
    }

    const structureTab = event.target.closest("[data-structure-tab]");
    if (structureTab) {
      state.structureTab = structureTab.getAttribute("data-structure-tab");
      renderShell();
      return;
    }

    const sectionButton = event.target.closest("[data-section-key]");
    if (sectionButton && !event.target.closest("[contenteditable='true']")) {
      state.selectedSectionKey = sectionButton.getAttribute("data-section-key");
      renderShell();
      return;
    }

    const blockButton = event.target.closest("[data-block-key]");
    if (blockButton && !event.target.closest("[contenteditable='true']")) {
      state.selectedSectionKey = blockButton.getAttribute("data-block-key");
      renderShell();
      return;
    }

    const contextTab = event.target.closest("[data-context-tab]");
    if (contextTab) {
      const value = contextTab.getAttribute("data-context-tab");
      if (state.route === "preview-export") state.previewTab = value;
      else if (noteVariant(state.snapshot.data.noteType) === "alert") state.publishTab = value;
      else state.contextTab = value;
      renderShell();
      return;
    }

    const addSection = event.target.closest("[data-add-section]");
    if (addSection) {
      addSectionAfterSelection();
      return;
    }

    const generateWord = event.target.closest("[data-generate-word], [data-export-kind='word']");
    if (generateWord) {
      state.api.generateWord();
      showToast("Word document generation started.", "info");
      return;
    }

    const exportButton = event.target.closest("[data-export-kind]");
    if (exportButton) {
      const kind = exportButton.getAttribute("data-export-kind");
      if (kind !== "word") {
        showToast(`${kind.toUpperCase()} packaging is configured in the production workflow layer.`, "info");
      }
      return;
    }

    const previewButton = event.target.closest("[data-open-preview]");
    if (previewButton) {
      state.api.openPreview();
      return;
    }

    const summaryButton = event.target.closest("[data-open-summary]");
    if (summaryButton) {
      state.api.openSummary();
      return;
    }

    const figureButton = event.target.closest("[data-trigger-figures]");
    if (figureButton) {
      state.api.triggerFigureUpload();
      return;
    }

    const filesButton = event.target.closest("[data-trigger-files]");
    if (filesButton) {
      state.api.triggerAttachmentUpload();
      return;
    }

    const toggleSection = event.target.closest("[data-toggle-section-setting]");
    if (toggleSection) {
      toggleSelectedSectionSetting(toggleSection.getAttribute("data-toggle-section-setting"));
      return;
    }

    const toggleExport = event.target.closest("[data-toggle-export]");
    if (toggleExport) {
      const key = toggleExport.getAttribute("data-toggle-export");
      state.exportOptions[key] = !state.exportOptions[key];
      renderShell();
      return;
    }

    const saveSettings = event.target.closest("[data-save-settings]");
    if (saveSettings) {
      saveSettingsDraft();
      return;
    }

    const resetSettings = event.target.closest("[data-reset-settings]");
    if (resetSettings) {
      state.settingsDraft = clone(state.settings);
      renderShell();
      return;
    }

    const changePassword = event.target.closest("[data-change-password]");
    if (changePassword) {
      promptPasswordChange();
      return;
    }

    const manageSessions = event.target.closest("[data-manage-sessions]");
    if (manageSessions) {
      showToast(`Active sessions: ${state.settings.security.activeSessions}`, "info");
      return;
    }

    const saveFinal = event.target.closest("[data-save-final]");
    if (saveFinal) {
      showToast("Final draft archived in the current workflow state.", "info");
      return;
    }

    const insertBlock = event.target.closest("[data-insert-block]");
    if (insertBlock) {
      insertStructuredBlock(insertBlock.getAttribute("data-insert-block"));
      return;
    }
  }

  function handlePlatformChange(event) {
    const workspaceSelect = event.target.closest("[data-workspace-select]");
    if (workspaceSelect) {
      switchWorkspace(workspaceSelect.value);
      return;
    }

    const channelToggle = event.target.closest("[data-channel-toggle]");
    if (channelToggle) {
      state.exportOptions.channels[channelToggle.getAttribute("data-channel-toggle")] = channelToggle.checked;
      return;
    }

    const settingInput = event.target.closest("[data-setting-path]");
    if (settingInput) {
      setNested(state.settingsDraft, settingInput.getAttribute("data-setting-path"), settingInput.type === "checkbox" ? settingInput.checked : settingInput.value);
    }
  }

  function handlePlatformInput(event) {
    const searchInput = event.target.closest("[data-global-search]");
    if (searchInput) {
      state.searchQuery = searchInput.value;
    }
  }

  function handlePlatformBlur(event) {
    const editableField = event.target.closest("[data-field-id][contenteditable='true']");
    if (editableField) {
      state.api.setFieldValue(editableField.getAttribute("data-field-id"), editableField.innerText.trim());
      return;
    }

    const authorLine = event.target.closest("[data-author-line]");
    if (authorLine) {
      state.api.setAuthorLine(authorLine.innerText.trim());
      return;
    }

    const richSection = event.target.closest("[data-rich-section-key]");
    if (richSection) {
      state.api.updateSection(richSection.getAttribute("data-rich-section-key"), { content: richSection.innerHTML });
      return;
    }

    const headingSection = event.target.closest("[data-heading-section-key]");
    if (headingSection) {
      state.api.updateSection(headingSection.getAttribute("data-heading-section-key"), { label: headingSection.innerText.trim() });
      return;
    }

    const figureField = event.target.closest("[data-figure-field]");
    if (figureField) {
      const figureKey = figureField.getAttribute("data-figure-key");
      const figureFieldName = figureField.getAttribute("data-figure-field");
      const rawValue = figureField.innerText.trim().replace(/^Note:\s*/i, "");
      if (figureFieldName === "title") {
        const parsed = parseFigureHeading(figureField.innerText.trim());
        state.api.updateFigureDetail(figureKey, parsed);
      } else {
        state.api.updateFigureDetail(figureKey, { [figureFieldName]: rawValue });
      }
    }
  }

  function handleDragStart(event) {
    const item = event.target.closest("[data-draggable-section-key]");
    if (!item) return;
    event.dataTransfer.setData("text/plain", item.getAttribute("data-draggable-section-key"));
    event.dataTransfer.effectAllowed = "move";
  }

  function handleDragOver(event) {
    if (event.target.closest("[data-draggable-section-key]")) {
      event.preventDefault();
    }
  }

  function handleDrop(event) {
    const target = event.target.closest("[data-draggable-section-key]");
    if (!target) return;
    event.preventDefault();
    const sourceKey = event.dataTransfer.getData("text/plain");
    const targetKey = target.getAttribute("data-draggable-section-key");
    if (!sourceKey || !targetKey || sourceKey === targetKey) return;

    const reorderableKeys = state.snapshot.sectionLayout.map((section) => section.key);
    const targetIndex = reorderableKeys.indexOf(targetKey);
    if (targetIndex >= 0) {
      state.api.moveSectionToIndex(sourceKey, targetIndex);
      state.selectedSectionKey = sourceKey;
    }
  }

  function handleDocumentClick(event) {
    if (state.userMenuOpen && !event.target.closest(".rpc-user")) {
      state.userMenuOpen = false;
      renderShell();
    }
  }

  function addSectionAfterSelection() {
    const realSections = state.snapshot.sectionLayout.map((section) => section.key);
    const afterSectionKey = realSections.includes(state.selectedSectionKey)
      ? state.selectedSectionKey
      : realSections[realSections.length - 1];
    const newKey = state.api.addCustomSection("New Section", afterSectionKey);
    if (newKey) {
      state.selectedSectionKey = newKey;
    }
  }

  function insertStructuredBlock(type) {
    if (type === "table") {
      const key = state.api.addCustomSection("Table Block", selectedRealSectionKey());
      if (key) {
        state.api.updateSection(key, { content: "<p>Insert tabular analysis here.</p>" });
        state.selectedSectionKey = key;
      }
      return;
    }

    if (type === "callout") {
      const key = state.api.addCustomSection("Callout", selectedRealSectionKey());
      if (key) {
        state.api.updateSection(key, { content: "<p>Insert a highlighted takeaway or contextual note.</p>" });
        state.selectedSectionKey = key;
      }
      return;
    }

    if (type === "textbox") {
      const key = state.api.addCustomSection("Text Box", selectedRealSectionKey());
      if (key) {
        state.api.updateSection(key, { content: "<p>Insert supporting commentary here.</p>" });
        state.selectedSectionKey = key;
      }
      return;
    }

    if (type === "divider") {
      showToast("Divider inserted into the current editorial workflow.", "info");
    }
  }

  async function saveSettingsDraft() {
    try {
      const saved = await fetchJson("/api/settings", {
        method: "POST",
        body: JSON.stringify(state.settingsDraft)
      });
      state.settings = mergeSettings(saved);
      state.settingsDraft = clone(state.settings);
      showToast("Settings saved.", "info");
      renderShell();
    } catch (error) {
      showToast(error.message || "Unable to save settings.", "error");
    }
  }

  async function promptPasswordChange() {
    const currentPassword = window.prompt("Enter your current password.");
    if (!currentPassword) return;
    const newPassword = window.prompt("Enter your new password.");
    if (!newPassword) return;

    try {
      const response = await fetchJson("/api/auth/change-password", {
        method: "POST",
        body: JSON.stringify({ currentPassword, newPassword })
      });
      state.settings.security.passwordChangedAt = response.passwordChangedAt;
      state.settingsDraft.security.passwordChangedAt = response.passwordChangedAt;
      showToast("Password updated.", "info");
      renderShell();
    } catch (error) {
      showToast(error.message || "Unable to update password.", "error");
    }
  }

  async function logout() {
    try {
      await fetchJson("/api/auth/logout", { method: "POST" });
    } catch (_error) {
      // Ignore logout transport errors and clear the UI session anyway.
    }
    state.user = null;
    state.security = null;
    state.settings = clone(DEFAULT_SETTINGS);
    state.settingsDraft = clone(DEFAULT_SETTINGS);
    renderAuth();
  }

  function setRoute(route) {
    window.location.hash = `#/${route}`;
  }

  function switchWorkspace(label) {
    switchNoteType(noteTypeForWorkspaceLabel(label));
    if (state.route !== "live-edit") setRoute("live-edit");
  }

  function switchNoteType(noteType) {
    state.api.setFieldValue("noteType", noteType);
    if (!state.snapshot.data.title) {
      state.api.setFieldValue("title", defaultNoteTitleForVariant(noteVariant(noteType)));
    }
  }

  function ensureSelectedSection() {
    const keys = buildStructureItems().map((item) => item.key);
    if (!keys.includes(state.selectedSectionKey)) {
      state.selectedSectionKey = keys[0] || "header";
    }
  }

  function buildStructureItems() {
    const data = state.snapshot.data;
    const variant = noteVariant(data.noteType);
    const sections = [];

    if (variant === "alert") {
      sections.push(
        structureItem("header", "Header", true, false),
        structureItem("keyTakeaways", "Key Takeaways", isSectionComplete("keyTakeaways"), true),
        structureItem("analysis", "What Happened", isSectionComplete("analysis"), true),
        structureItem("content", "Why It Matters", isSectionComplete("content"), true),
        structureItem("market-impact", "Market Impact", true, false),
        structureItem("cordobaView", "Cordoba View", isSectionComplete("cordobaView"), true),
        structureItem(findDisclosureSection()?.key || "disclosures", "Disclosures", true, Boolean(findDisclosureSection()))
      );
      return sections;
    }

    sections.push(structureItem("header", "Header / Metadata", true, false));
    state.snapshot.sectionLayout.forEach((section) => {
      sections.push(structureItem(section.key, section.label, Boolean(section.content), true));
    });
    sections.push(structureItem("figures", "Charts / Figures", Boolean(state.snapshot.figures.length), false));
    if (variant === "general") sections.push(structureItem("website-summary", "Website Summary", true, false));
    sections.push(structureItem(findDisclosureSection()?.key || "disclosures", "Disclosures", true, Boolean(findDisclosureSection())));
    return sections;
  }

  function renderDraftCard(card) {
    return `
      <article class="rpc-draft-card rpc-card">
        <div class="rpc-draft-top">
          <span>Draft <span class="rpc-state-dot"></span> Saved ${escapeHtml(card.savedAgo)}</span>
          <button type="button" class="rpc-kebab">⋮</button>
        </div>
        <div>
          <div class="rpc-draft-title">${escapeHtml(card.title)}</div>
          <div class="rpc-draft-copy">${escapeHtml(card.copy)}</div>
        </div>
        <div class="rpc-tag-row">
          ${card.tags.map((tag, index) => `<span class="rpc-tag ${index ? "rpc-tag-neutral" : ""}">${escapeHtml(tag)}</span>`).join("")}
        </div>
        <div class="rpc-draft-foot">
          <div class="rpc-avatar-row">
            <span class="rpc-avatar">${escapeHtml(card.initials)}</span>
            <span>${escapeHtml(card.author)}</span>
          </div>
          <button type="button" class="rpc-arrow-link" data-nav-route="live-edit">→</button>
        </div>
      </article>
    `;
  }

  function buildDashboardCards(snapshot) {
    const current = {
      title: snapshot.data.title || DASHBOARD_CARDS[0].title,
      savedAgo: relativeTimestamp(snapshot.lastSavedAt),
      copy: snapshot.data.deck || DASHBOARD_CARDS[0].copy,
      tags: [workspaceLabelForNoteType(snapshot.data.noteType), snapshot.data.metadataPrimaryGeography || "Emerging Markets"].filter(Boolean),
      initials: state.user?.initials || "CR",
      author: state.user?.fullName || "Current User"
    };

    return [current, ...DASHBOARD_CARDS.slice(1)];
  }

  function renderFinalValidationLines(snapshot) {
    const lines = [
      { label: "Sections complete", detail: `${snapshot.sectionLayout.filter((section) => section.content).length} / ${snapshot.sectionLayout.length} sections included` },
      { label: "Charts & tables", detail: `${snapshot.figures.length} charts, ${snapshot.data.financialTableInput ? "1 table" : "0 tables"}` },
      { label: "Key takeaways", detail: `${bulletCount(snapshot.data.keyTakeaways)} bullets included` },
      { label: "Disclosures", detail: findDisclosureSection() ? "Up to date" : "Needs review" },
      { label: "Compliance review", detail: snapshot.shariah?.tone === "flag" ? "Manual review required" : "No issues detected" }
    ];

    return lines.map((line) => `
      <div class="rpc-validation-row">
        <div>
          <strong>${escapeHtml(line.label)}</strong>
          <div class="rpc-item-meta">${escapeHtml(line.detail)}</div>
        </div>
        <span class="rpc-status-good">✓</span>
      </div>
    `).join("");
  }

  function renderNotificationRow(label, copy, path, checked) {
    return `
      <label class="rpc-checklist-row">
        <input type="checkbox" ${checked ? "checked" : ""} data-setting-path="${escapeAttr(path)}">
        <span>
          <strong>${escapeHtml(label)}</strong>
          <p>${escapeHtml(copy)}</p>
        </span>
      </label>
    `;
  }

  function renderStatusRow(label, status) {
    const tone = /flag|review|open/i.test(status) ? "rpc-status-warn" : /clear|ready|good|passed/i.test(status) ? "rpc-status-good" : "rpc-status-danger";
    return `<div class="rpc-validation-row"><span>${escapeHtml(label)}</span><strong class="${tone}">${escapeHtml(status)}</strong></div>`;
  }

  function sectionSettingsForKey(key) {
    if (!state.sectionSettings) state.sectionSettings = {};
    if (!state.sectionSettings[key]) {
      state.sectionSettings[key] = {
        showNumber: !["header", "figures", "website-summary", "market-impact", "disclosures"].includes(key),
        collapsible: false
      };
    }
    return state.sectionSettings[key];
  }

  function toggleSelectedSectionSetting(settingKey) {
    const settings = sectionSettingsForKey(state.selectedSectionKey);
    settings[settingKey] = !settings[settingKey];
    renderShell();
  }

  function showSectionNumber(key) {
    return sectionSettingsForKey(key).showNumber;
  }

  function sectionNumberForKey(key) {
    const realSections = buildStructureItems().filter((item) => !["header", "figures", "website-summary", "market-impact", "disclosures"].includes(item.key));
    const index = realSections.findIndex((item) => item.key === key);
    return index >= 0 ? index + 1 : "";
  }

  function selectedRealSectionKey() {
    const keys = state.snapshot.sectionLayout.map((section) => section.key);
    return keys.includes(state.selectedSectionKey) ? state.selectedSectionKey : keys[keys.length - 1];
  }

  function getSectionByKey(sectionKey) {
    return state.snapshot.sectionLayout.find((section) => section.key === sectionKey) || null;
  }

  function isSectionComplete(sectionKey) {
    if (sectionKey === "header") return Boolean(state.snapshot.data.title);
    const section = getSectionByKey(sectionKey);
    return Boolean(section?.content);
  }

  function findDisclosureSection() {
    return state.snapshot.sectionLayout.find((section) => /disclosure/i.test(section.label));
  }

  function buildAuthorLineFromData(data) {
    const authors = [];
    const primary = [data.authorFirstName, data.authorLastName].filter(Boolean).join(" ").trim();
    if (primary) authors.push(primary);
    (data.coAuthors || []).forEach((author) => {
      const name = [author.firstName, author.lastName].filter(Boolean).join(" ").trim();
      if (name) authors.push(name);
    });
    return authors.join(" | ") || "Alex Smith | Priya Natarajan | Daniel Kwon";
  }

  function workspaceOptions() {
    return [
      { value: "Macro / Sovereign Outlook", label: "Macro / Sovereign Outlook" },
      { value: "General Note", label: "General Note" },
      { value: "Short Note / Market Alert", label: "Short Note / Market Alert" },
      { value: "Equity Research", label: "Equity Research" },
      { value: "Commodity Insights", label: "Commodity Insights" }
    ];
  }

  function workspaceLabelForNoteType(noteType) {
    if (noteType === "Short Note / Market Alert") return "Short Note / Market Alert";
    if (noteType === "General Note") return "General Note";
    if (noteType === "Equity Research") return "Equity Research";
    if (noteType === "Commodity Insights") return "Commodity Insights";
    return "Macro / Sovereign Outlook";
  }

  function noteTypeForWorkspaceLabel(label) {
    if (label === "Short Note / Market Alert") return "Short Note / Market Alert";
    if (label === "General Note") return "General Note";
    if (label === "Equity Research") return "Equity Research";
    if (label === "Commodity Insights") return "Commodity Insights";
    return "Macro Research";
  }

  function noteVariant(noteType) {
    if (noteType === "Short Note / Market Alert") return "alert";
    if (noteType === "General Note" || noteType === "Equity Research" || noteType === "Commodity Insights") return "general";
    return "macro";
  }

  function topbarTitleForRoute(variant) {
    if (state.route === "dashboard") return "Research Production Console";
    if (state.route === "preview-export") return "Console";
    if (state.route === "settings") return "Research Production Console";
    if (state.route === "live-edit" && variant === "alert") return "Short Note / Market Alert";
    return "Research Production Console";
  }

  function displayWorkflowState() {
    const workflow = String(state.snapshot.data.workflowStage || "draft").replace(/-/g, " ");
    return workflow.charAt(0).toUpperCase() + workflow.slice(1);
  }

  function defaultNoteTitleForVariant(variant) {
    if (variant === "alert") return "US CPI Surprise Sparks Rates Volatility";
    if (variant === "general") return "Global Growth Outlook 2H 2025: Divergent Paths, Selective Opportunities";
    return "Uzbekistan Sovereign Outlook";
  }

  function defaultGeographyForNoteType(noteType) {
    return noteVariant(noteType) === "general" ? "Global" : "Emerging Markets";
  }

  function defaultDisclosureCopy() {
    return "Cordoba Capital Markets is a division of Cordoba Group. This communication is for informational purposes only and does not constitute investment advice.";
  }

  function bulletCount(value) {
    return String(value || "").split(/\n+/).filter((line) => line.trim()).length || 0;
  }

  function parseTags(value) {
    return String(value || "")
      .split(/[,\n]/)
      .map((entry) => entry.trim())
      .filter(Boolean);
  }

  function structureItem(key, label, complete, reorderable) {
    return { key, label, complete, reorderable };
  }

  function contextTabButton(value, label, activeValue) {
    return `<button type="button" class="rpc-right-tab ${value === activeValue ? "is-active" : ""}" data-context-tab="${escapeAttr(value)}">${escapeHtml(label)}</button>`;
  }

  function toolbarButton(command, label, isIcon = false, value = "") {
    return `<button type="button" class="rpc-toolbar-btn" data-toolbar-command="${escapeAttr(command)}" ${value ? `data-toolbar-value="${escapeAttr(value)}"` : ""}>${isIcon ? label : `<span>${label}</span>`}</button>`;
  }

  function handleToolbarCommand(command, value) {
    if (command === "createLink") {
      const link = window.prompt("Enter the link URL.");
      if (!link) return;
      document.execCommand(command, false, link);
      return;
    }

    if (command === "insertHTML" && value === "table") {
      document.execCommand("insertHTML", false, "<table><tr><td>Table cell</td><td>Table cell</td></tr></table>");
      return;
    }

    if (command === "insertHTML" && value === "ellipsis") return;
    document.execCommand(command, false, value || null);
  }

  function parseFigureHeading(value) {
    const match = String(value || "").match(/^\s*([A-Za-z]+)\s+(\d+)\s*:\s*(.+)$/);
    if (!match) return { title: value };
    return {
      labelPrefix: match[1],
      labelNumber: match[2],
      title: match[3]
    };
  }

  function mergeSettings(next) {
    return {
      profile: {
        ...DEFAULT_SETTINGS.profile,
        ...(next.profile || {})
      },
      defaults: {
        ...DEFAULT_SETTINGS.defaults,
        ...(next.defaults || {})
      },
      notifications: {
        ...DEFAULT_SETTINGS.notifications,
        ...(next.notifications || {})
      },
      security: {
        ...DEFAULT_SETTINGS.security,
        ...(next.security || {})
      }
    };
  }

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function setNested(target, path, value) {
    const parts = String(path || "").split(".");
    let cursor = target;
    while (parts.length > 1) {
      const part = parts.shift();
      if (!cursor[part] || typeof cursor[part] !== "object") cursor[part] = {};
      cursor = cursor[part];
    }
    cursor[parts[0]] = value;
  }

  function isRouteActive(route) {
    if (route === "live-edit" && state.route === "preview-export") return false;
    return state.route === route;
  }

  function getRouteFromHash() {
    const route = window.location.hash.replace(/^#\//, "").trim();
    return route || "dashboard";
  }

  function relativeTimestamp(timestamp) {
    if (!timestamp) return "just now";
    const date = new Date(timestamp);
    const diff = Date.now() - date.getTime();
    if (!Number.isFinite(diff) || diff < 0) return "just now";
    const minutes = Math.round(diff / 60000);
    if (minutes <= 1) return "1m ago";
    if (minutes < 60) return `${minutes}m ago`;
    const hours = Math.round(minutes / 60);
    if (hours < 24) return `${hours}h ago`;
    const days = Math.round(hours / 24);
    return `${days}d ago`;
  }

  function formatDateLong(value) {
    const date = value ? new Date(value) : new Date();
    return new Intl.DateTimeFormat("en-GB", { month: "long", day: "numeric", year: "numeric" }).format(date);
  }

  function formatDateShort(value) {
    const date = value ? new Date(value) : new Date();
    return new Intl.DateTimeFormat("en-GB", { month: "short", day: "numeric", year: "numeric" }).format(date);
  }

  function formatDateWithTime(value) {
    const date = value ? new Date(value) : new Date();
    return `${formatDateLong(date)} 10:41 AM ET`;
  }

  function formatLongDate(value) {
    if (!value) return "May 10, 2025";
    return new Intl.DateTimeFormat("en-GB", { month: "long", day: "numeric", year: "numeric" }).format(new Date(value));
  }

  function formatPublishWindow(value) {
    const start = value ? new Date(value) : new Date();
    const end = new Date(start.getTime() + (7 * 24 * 60 * 60 * 1000));
    return `${formatDateShort(start)} – ${formatDateShort(end)}`;
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("en-GB").format(Number(value || 0));
  }

  function slugify(value) {
    return String(value || "")
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "");
  }

  function showToast(message, tone) {
    state.toast = { message, tone };
    clearTimeout(state.toastTimer);
    if (state.user) renderShell();
    else renderAuth();
    state.toastTimer = window.setTimeout(() => {
      state.toast = null;
      if (state.user) renderShell();
      else renderAuth();
    }, 2600);
  }

  async function fetchJson(url, options = {}) {
    let response;
    try {
      response = await fetch(`${API_BASE}${url}`, {
        method: options.method || "GET",
        credentials: "same-origin",
        headers: {
          "Content-Type": "application/json",
          ...(options.headers || {})
        },
        body: options.body
      });
    } catch (error) {
      if (error instanceof TypeError) {
        throw new Error(authUnavailableMessage());
      }
      throw error;
    }

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || "Request failed.");
    }
    return payload;
  }

  function assetUrl(assetPath) {
    return window.location.protocol === "file:" ? assetPath : `/${assetPath.replace(/^\/+/, "")}`;
  }

  function authUnavailableMessage() {
    return "Authentication service unavailable. Start the RDT server with `npm start` and open http://localhost:3000.";
  }

  function friendlyAuthError(error) {
    const message = error?.message || "Unable to sign in.";
    if (/authentication service unavailable/i.test(message) || /failed to fetch/i.test(message)) {
      return authUnavailableMessage();
    }
    return message;
  }

  async function waitForLegacyApi() {
    if (window.RDTLegacyAPI) return window.RDTLegacyAPI;
    return new Promise((resolve, reject) => {
      const start = Date.now();
      const timer = window.setInterval(() => {
        if (window.RDTLegacyAPI) {
          window.clearInterval(timer);
          resolve(window.RDTLegacyAPI);
          return;
        }

        if (Date.now() - start > 12000) {
          window.clearInterval(timer);
          reject(new Error("Legacy authoring engine unavailable."));
        }
      }, 60);
    });
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function escapeAttr(value) {
    return escapeHtml(value);
  }

  function icon(name) {
    const icons = {
      dashboard: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><rect x="3" y="3" width="7" height="7" rx="1.5"></rect><rect x="14" y="3" width="7" height="7" rx="1.5"></rect><rect x="14" y="14" width="7" height="7" rx="1.5"></rect><rect x="3" y="14" width="7" height="7" rx="1.5"></rect></svg>`,
      library: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5V4.5A2.5 2.5 0 0 1 6.5 2Z"></path></svg>`,
      drafts: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><path d="M14 2v6h6"></path><path d="M8 13h8"></path><path d="M8 17h5"></path></svg>`,
      edit: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 20h9"></path><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"></path></svg>`,
      templates: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><rect x="3" y="4" width="18" height="18" rx="2"></rect><path d="M8 2v4"></path><path d="M16 2v4"></path><path d="M3 10h18"></path><path d="M8 14h3"></path></svg>`,
      distribution: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="m22 2-7 20-4-9-9-4Z"></path><path d="M22 2 11 13"></path></svg>`,
      compliance: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10Z"></path></svg>`,
      calendar: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M8 2v4"></path><path d="M16 2v4"></path><rect x="3" y="4" width="18" height="18" rx="2"></rect><path d="M3 10h18"></path></svg>`,
      settings: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6Z"></path><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.6 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.6a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09A1.65 1.65 0 0 0 15 4.6a1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9c.36.49.58 1.07.6 1.68V11a1 1 0 0 0 1 1"></path></svg>`,
      search: `<svg class="rpc-search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><circle cx="11" cy="11" r="7"></circle><path d="m21 21-4.35-4.35"></path></svg>`,
      user: `<svg class="rpc-input-leading" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M20 21a8 8 0 1 0-16 0"></path><circle cx="12" cy="7" r="4"></circle></svg>`,
      lock: `<svg class="rpc-input-leading" viewBox="0 0 24 24" fill="none" stroke="currentColor"><rect x="5" y="11" width="14" height="10" rx="2"></rect><path d="M8 11V7a4 4 0 1 1 8 0v4"></path></svg>`,
      eye: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M2 12s3.6-6 10-6 10 6 10 6-3.6 6-10 6-10-6-10-6Z"></path><circle cx="12" cy="12" r="3"></circle></svg>`,
      "eye-off": `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="m3 3 18 18"></path><path d="M10.58 10.58A2 2 0 0 0 13.42 13.42"></path><path d="M9.88 5.09A10.94 10.94 0 0 1 12 5c6.4 0 10 7 10 7a18.86 18.86 0 0 1-2.21 3.19"></path><path d="M6.71 6.72C3.77 8.74 2 12 2 12a18.73 18.73 0 0 0 7.1 6.11"></path></svg>`,
      microsoft: `<svg viewBox="0 0 24 24"><rect width="10" height="10" x="2" y="2" fill="#f35325"></rect><rect width="10" height="10" x="12" y="2" fill="#81bc06"></rect><rect width="10" height="10" x="2" y="12" fill="#05a6f0"></rect><rect width="10" height="10" x="12" y="12" fill="#ffba08"></rect></svg>`,
      shield: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10Z"></path><path d="m9 12 2 2 4-4"></path></svg>`,
      "chevron-down": `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="m6 9 6 6 6-6"></path></svg>`,
      clock: `<svg class="rpc-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><circle cx="12" cy="12" r="9"></circle><path d="M12 7v5l3 3"></path></svg>`,
      doc: `<svg class="rpc-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><path d="M14 2v6h6"></path></svg>`,
      "chart-up": `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M3 3v18h18"></path><path d="m7 14 4-4 3 3 5-7"></path></svg>`,
      macro: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M3 12h18"></path><path d="M12 3v18"></path><circle cx="12" cy="12" r="9"></circle></svg>`,
      commodity: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 2v20"></path><path d="M7 7h10"></path><path d="M7 17h10"></path><rect x="6" y="4" width="12" height="16" rx="2"></rect></svg>`,
      alert: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="m13 2-9 12h6l-1 8 9-12h-6Z"></path></svg>`,
      note: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M14 2H6a2 2 0 0 0-2 2v16l4-3 4 3 4-3 4 3V8Z"></path><path d="M14 2v6h6"></path></svg>`,
      "align-left": `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M3 6h18"></path><path d="M3 12h12"></path><path d="M3 18h18"></path></svg>`,
      "align-center": `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M3 6h18"></path><path d="M6 12h12"></path><path d="M3 18h18"></path></svg>`,
      "align-right": `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M3 6h18"></path><path d="M9 12h12"></path><path d="M3 18h18"></path></svg>`,
      list: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M8 6h13"></path><path d="M8 12h13"></path><path d="M8 18h13"></path><path d="M3 6h.01"></path><path d="M3 12h.01"></path><path d="M3 18h.01"></path></svg>`,
      link: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M10 13a5 5 0 0 0 7.07 0l2.83-2.83a5 5 0 0 0-7.07-7.07L10 5"></path><path d="M14 11a5 5 0 0 0-7.07 0L4.1 13.83a5 5 0 0 0 7.07 7.07L14 19"></path></svg>`,
      table: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><rect x="3" y="4" width="18" height="16" rx="2"></rect><path d="M3 10h18"></path><path d="M9 4v16"></path><path d="M15 4v16"></path></svg>`,
      desktop: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><rect x="3" y="4" width="18" height="12" rx="2"></rect><path d="M8 20h8"></path><path d="M12 16v4"></path></svg>`,
      mobile: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><rect x="7" y="2" width="10" height="20" rx="2"></rect><path d="M11 18h2"></path></svg>`
    };
    return icons[name] || "";
  }

  platformRoot.addEventListener("click", (event) => {
    const toolbar = event.target.closest("[data-toolbar-command]");
    if (toolbar) {
      handleToolbarCommand(toolbar.getAttribute("data-toolbar-command"), toolbar.getAttribute("data-toolbar-value") || "");
      return;
    }

    const formatBlock = event.target.closest("[data-format-block]");
    if (formatBlock) {
      document.execCommand("formatBlock", false, formatBlock.getAttribute("data-format-block"));
    }
  });
})();
