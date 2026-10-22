/**
 * Email Sticky Notes for Outlook - Taskpane (v2.1)
 * Plain ES5, no build step. Depends on notes-core.js (window.StickyNotesCore).
 */
(function () {
  "use strict";

  var Core = window.StickyNotesCore;
  var THEME_KEY = "outlook_sticky_notes_theme";
  var DEMO_SEED_FLAG_KEY = "outlook_sticky_notes_demo_seeded_v2";

  var PALETTE = {
    yellow: { label: "Classic Yellow", swatch: "#FFE255", lBg: "#FFF9C4", lHd: "#FFF176", lBd: "#FBC02D", lTx: "#3E2723", dBg: "#3D3514", dHd: "#4F4418", dBd: "#A68B22", dTx: "#FFF59D" },
    amber: { label: "Warm Amber", swatch: "#FFB74D", lBg: "#FFE0B2", lHd: "#FFCC80", lBd: "#FB8C00", lTx: "#3E2723", dBg: "#3E2914", dHd: "#523519", dBd: "#B26A1B", dTx: "#FFE0B2" },
    green: { label: "Mint Green", swatch: "#81C784", lBg: "#DCEDC8", lHd: "#C5E1A5", lBd: "#7CB342", lTx: "#1B5E20", dBg: "#183622", dHd: "#20472D", dBd: "#3E8E53", dTx: "#C8E6C9" },
    blue: { label: "Sky Blue", swatch: "#64B5F6", lBg: "#E1F5FE", lHd: "#B3E5FC", lBd: "#039BE5", lTx: "#0D47A1", dBg: "#142B3E", dHd: "#1B3B54", dBd: "#2E75B6", dTx: "#BBDEFB" },
    pink: { label: "Rose Pink", swatch: "#F06292", lBg: "#FCE4EC", lHd: "#F8BBD0", lBd: "#D81B60", lTx: "#4A148C", dBg: "#3B1828", dHd: "#4F2036", dBd: "#A83868", dTx: "#F8BBD0" },
    purple: { label: "Lavender", swatch: "#BA68C8", lBg: "#F3E5F5", lHd: "#E1BEE7", lBd: "#8E24AA", lTx: "#311B92", dBg: "#2B193D", dHd: "#3A2252", dBd: "#7B46A8", dTx: "#E1BEE7" },
  };

  var PRIORITY = {
    low: { label: "Low", rank: 1, color: "#64748b" },
    medium: { label: "Medium", rank: 2, color: "#0284c7" },
    high: { label: "High", rank: 3, color: "#d97706" },
    urgent: { label: "Urgent", rank: 4, color: "#dc2626" },
  };

  var PRESET_TAGS = ["follow-up", "waiting", "action-required", "decision", "client", "billing", "important", "meeting-notes"];

  var DEMO_EMAILS = [
    { itemId: "demo-msg-001-q4-contract", conversationId: "conv-q4-enterprise-renewal", subject: "Q4 Enterprise License Renewal & Pricing Proposal", senderName: "Sarah Jenkins (VP Procurement)", senderEmail: "s.jenkins@contoso-enterprise.com", userDisplayName: "Outlook User" },
    { itemId: "demo-msg-002-security-audit", conversationId: "conv-soc2-security-review", subject: "Action Required: SOC2 Type II Architecture Sign-off", senderName: "Marcus Vance (Security Engineering)", senderEmail: "mvance@cloudsec-ops.io", userDisplayName: "Outlook User" },
    { itemId: "demo-msg-003-board-deck", conversationId: "conv-board-deck-prep", subject: "Draft Board Deck Slides & Financial Appendix", senderName: "Elena Rostova (CFO)", senderEmail: "elena@northwind-ventures.com", userDisplayName: "Outlook User" },
  ];

  var state = {
    isInOutlook: false,
    noItem: false,
    loading: true,
    demoIndex: 0,
    ctx: DEMO_EMAILS[0],
    emailNotes: [],
    allNotes: [],
    isDark: false,
    viewScope: "current_email",
    searchQuery: "",
    selectedColor: "all",
    selectedPriority: "all",
    selectedTag: "all",
    sortBy: "updated_desc",
    onlyPinned: false,
    editor: null,
    settingsOpen: false,
    toast: null,
    toastUndo: false,
    lastDeleted: null,
  };

  try {
    state.isDark = localStorage.getItem(THEME_KEY) === "dark";
  } catch (e) {}

  var toastTimer = null;
  var loadSeq = 0;

  // ---------------------------------------------------------------------------
  // Utilities
  // ---------------------------------------------------------------------------
  function esc(str) {
    return String(str === undefined || str === null ? "" : str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function formatTimestamp(iso) {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return "";
    var mins = Math.floor((Date.now() - d.getTime()) / 60000);
    if (mins < 1) return "Just now";
    if (mins < 60) return mins + "m ago";
    var hrs = Math.floor(mins / 60);
    if (hrs < 24) return hrs + "h ago";
    return d.toLocaleDateString(undefined, { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
  }

  function currentItem() {
    return state.isInOutlook && Office.context && Office.context.mailbox ? Office.context.mailbox.item : null;
  }

  function sortPinnedFirst(list) {
    return list.slice().sort(function (a, b) {
      if (Boolean(a.isPinned) !== Boolean(b.isPinned)) return a.isPinned ? -1 : 1;
      return String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
    });
  }

  function showToast(msg, withUndo) {
    state.toast = msg;
    state.toastUndo = Boolean(withUndo);
    render();
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(function () {
      state.toast = null;
      state.toastUndo = false;
      render();
    }, withUndo ? 6000 : 3000);
  }

  // ---------------------------------------------------------------------------
  // Email context
  // ---------------------------------------------------------------------------
  function readContext() {
    if (!state.isInOutlook) return DEMO_EMAILS[state.demoIndex];
    var mailbox = Office.context.mailbox;
    var item = mailbox.item;
    if (!item) return null;
    var up = mailbox.userProfile || {};
    var from = item.from || item.sender || {};
    var key = Core.getItemKey(item);
    return {
      itemId: key,
      conversationId: item.conversationId || key,
      subject: Core.subjectOf(item),
      senderName: from.displayName || from.emailAddress || "",
      senderEmail: from.emailAddress || "",
      userDisplayName: up.displayName || "Outlook User",
    };
  }

  function ensureDemoSeed(cb) {
    var seeded = false;
    try {
      seeded = localStorage.getItem(DEMO_SEED_FLAG_KEY) === "true";
    } catch (e) {}
    if (state.isInOutlook || seeded) {
      cb();
      return;
    }
    var e1 = DEMO_EMAILS[0];
    var e2 = DEMO_EMAILS[1];
    var now = Date.now();
    var iso = function (minsAgo) {
      return new Date(now - minsAgo * 60000).toISOString();
    };
    var n1 = [
      { id: "seed-note-101", itemId: e1.itemId, conversationId: e1.conversationId, scope: "message", emailSubject: e1.subject, emailSender: e1.senderName + " <" + e1.senderEmail + ">", title: "Negotiate 12% Volume Discount Tier", content: "Sarah confirmed flexibility on the 3-year renewal if we sign before quarter-end. Verify seat count with IT Ops before replying.", color: "yellow", priority: "high", tags: ["follow-up", "billing", "client"], checklist: [{ id: "chk-1", text: "Confirm active M365 seat count with IT", completed: true }, { id: "chk-2", text: "Request revised redline PDF from Legal", completed: false }], isPinned: true, author: "Outlook User", createdAt: iso(35), updatedAt: iso(12) },
      { id: "seed-note-102", itemId: e1.itemId, conversationId: e1.conversationId, scope: "message", emailSubject: e1.subject, emailSender: e1.senderName + " <" + e1.senderEmail + ">", title: "PO Approval Reference #PO-88412", content: "Finance pre-approved budget code FIN-2026-Q4.", color: "green", priority: "medium", tags: ["decision", "billing"], checklist: [], isPinned: false, author: "Outlook User", createdAt: iso(90), updatedAt: iso(90) },
    ];
    var n2 = [
      { id: "seed-note-201", itemId: e2.itemId, conversationId: e2.conversationId, scope: "message", emailSubject: e2.subject, emailSender: e2.senderName + " <" + e2.senderEmail + ">", title: "Verify Conditional Access Evidence", content: "Auditor requested a screenshot of the MFA enforcement policy.", color: "pink", priority: "urgent", tags: ["action-required", "important"], checklist: [{ id: "chk-201", text: "Export Conditional Access policy JSON", completed: false }], isPinned: true, author: "Outlook User", createdAt: iso(180), updatedAt: iso(180) },
    ];
    Core.saveItemNotes(null, e1.itemId, n1, function () {
      Core.saveItemNotes(null, e2.itemId, n2, function () {
        Core.writeIndex(n1.concat(n2), function () {
          try {
            localStorage.setItem(DEMO_SEED_FLAG_KEY, "true");
          } catch (e) {}
          cb();
        });
      });
    });
  }

  /** Loads notes for whichever email is currently selected. */
  function loadCurrent() {
    var seq = ++loadSeq;
    var ctx = readContext();
    if (!ctx) {
      state.noItem = true;
      state.loading = false;
      state.editor = null;
      state.emailNotes = [];
      state.allNotes = Core.readIndex();
      render();
      return;
    }
    state.noItem = false;
    state.ctx = ctx;
    state.loading = true;
    render();

    ensureDemoSeed(function () {
      var item = currentItem();
      Core.loadItemNotes(item, ctx.itemId, function (notes) {
        if (seq !== loadSeq) return; // user already switched to another email
        state.emailNotes = sortPinnedFirst(notes || []);
        state.allNotes = Core.readIndex();
        state.loading = false;
        render();
        if (item) Core.syncMarkers(item, ctx.itemId, state.emailNotes, null, function () {});
      });
    });
  }

  /** Saves the full note list for the current email and refreshes the UI. */
  function persistCurrent(nextNotes, toastMsg, withUndo) {
    var item = currentItem();
    var key = state.ctx.itemId;
    Core.saveItemNotes(item, key, nextNotes, function (ok) {
      if (!ok) {
        showToast("Couldn't save to your mailbox. Please try again.");
        return;
      }
      Core.replaceItemInIndex(key, nextNotes, function () {
        if (key !== state.ctx.itemId) return; // switched emails mid-save
        state.emailNotes = sortPinnedFirst(nextNotes);
        state.allNotes = Core.readIndex();
        if (item) Core.syncMarkers(item, key, nextNotes, null, function () {});
        if (toastMsg) showToast(toastMsg, withUndo);
        else render();
      });
    });
  }

  function findCurrentNote(id) {
    return state.emailNotes.find(function (n) {
      return n.id === id;
    });
  }

  // ---------------------------------------------------------------------------
  // Filtering
  // ---------------------------------------------------------------------------
  function sourceNotes() {
    return state.viewScope === "current_email" ? state.emailNotes : state.allNotes;
  }

  function filteredNotes() {
    var q = state.searchQuery.trim().toLowerCase();
    return sourceNotes()
      .filter(function (n) {
        if (state.onlyPinned && !n.isPinned) return false;
        if (state.selectedColor !== "all" && n.color !== state.selectedColor) return false;
        if (state.selectedPriority !== "all" && n.priority !== state.selectedPriority) return false;
        if (state.selectedTag !== "all" && (n.tags || []).indexOf(state.selectedTag) === -1) return false;
        if (!q) return true;
        var hay = [n.title, n.content, n.emailSubject, n.emailSender, (n.tags || []).join(" "), (n.checklist || []).map(function (c) { return c.text; }).join(" ")].join(" ").toLowerCase();
        return hay.indexOf(q) !== -1;
      })
      .sort(function (a, b) {
        if (Boolean(a.isPinned) !== Boolean(b.isPinned)) return a.isPinned ? -1 : 1;
        if (state.sortBy === "priority_desc") return (PRIORITY[b.priority] || PRIORITY.medium).rank - (PRIORITY[a.priority] || PRIORITY.medium).rank;
        if (state.sortBy === "created_desc") return String(b.createdAt).localeCompare(String(a.createdAt));
        if (state.sortBy === "created_asc") return String(a.createdAt).localeCompare(String(b.createdAt));
        return String(b.updatedAt).localeCompare(String(a.updatedAt));
      });
  }

  function hasActiveFilters() {
    return Boolean(state.searchQuery) || state.selectedColor !== "all" || state.selectedPriority !== "all" || state.selectedTag !== "all" || state.onlyPinned;
  }

  // ---------------------------------------------------------------------------
  // Editor
  // ---------------------------------------------------------------------------
  function openEditor(note) {
    state.editor = note
      ? { id: note.id, title: note.title || "", content: note.content || "", color: note.color || "yellow", priority: note.priority || "medium", isPinned: Boolean(note.isPinned), tags: (note.tags || []).slice(), checklist: (note.checklist || []).map(function (c) { return { id: c.id, text: c.text, completed: Boolean(c.completed) }; }), newChk: "", newTag: "" }
      : { id: null, title: "", content: "", color: "yellow", priority: "medium", isPinned: false, tags: [], checklist: [], newChk: "", newTag: "" };
    render();
    var t = document.getElementById("ed-title");
    if (t) t.focus();
  }

  function syncEditorInputs() {
    var ed = state.editor;
    if (!ed) return;
    var get = function (id) {
      var el = document.getElementById(id);
      return el ? el.value : null;
    };
    var v;
    if ((v = get("ed-title")) !== null) ed.title = v;
    if ((v = get("ed-content")) !== null) ed.content = v;
    if ((v = get("ed-priority")) !== null) ed.priority = v;
    if ((v = get("ed-new-chk")) !== null) ed.newChk = v;
    if ((v = get("ed-new-tag")) !== null) ed.newTag = v;
  }

  function saveEditor() {
    syncEditorInputs();
    var ed = state.editor;
    if (!ed) return;
    if (ed.newChk.trim()) addChecklistItem();
    if (!ed.title.trim() && !ed.content.trim() && ed.checklist.length === 0) {
      showToast("Add a title, text, or checklist item first.");
      return;
    }
    var nowIso = new Date().toISOString();
    var fields = { title: ed.title.trim(), content: ed.content.trim(), color: ed.color, priority: ed.priority, tags: ed.tags.slice(), checklist: ed.checklist.slice(), isPinned: ed.isPinned, updatedAt: nowIso };
    var next;
    if (ed.id) {
      next = state.emailNotes.map(function (n) {
        return n.id === ed.id ? Object.assign({}, n, fields) : n;
      });
    } else {
      var created = Object.assign(
        { id: Core.newId("note"), itemId: state.ctx.itemId, conversationId: state.ctx.conversationId, scope: "message", emailSubject: state.ctx.subject, emailSender: state.ctx.senderName + (state.ctx.senderEmail ? " <" + state.ctx.senderEmail + ">" : ""), author: state.ctx.userDisplayName, createdAt: nowIso, syncStatus: "synced_office" },
        fields
      );
      next = [created].concat(state.emailNotes);
    }
    var wasEdit = Boolean(ed.id);
    state.editor = null;
    persistCurrent(next, wasEdit ? "Sticky note updated" : "Sticky note saved to this email");
  }

  function addChecklistItem() {
    var ed = state.editor;
    var text = (ed.newChk || "").trim();
    if (!text) return;
    ed.checklist.push({ id: Core.newId("chk"), text: text, completed: false });
    ed.newChk = "";
  }

  function addCustomTag() {
    var ed = state.editor;
    var tag = (ed.newTag || "").trim().toLowerCase().replace(/^#/, "").replace(/\s+/g, "-");
    if (tag && ed.tags.indexOf(tag) === -1) ed.tags.push(tag);
    ed.newTag = "";
  }

  // ---------------------------------------------------------------------------
  // Banner controls
  // ---------------------------------------------------------------------------
  function setBannerForThisEmail(hidden) {
    var item = currentItem();
    var key = state.ctx.itemId;
    Core.setBannerDismissed(key, hidden, function () {
      Core.syncBanner(item, key, state.emailNotes, null, function () {
        showToast(hidden ? "Banner hidden for this email" : "Banner shown for this email");
      });
    });
  }

  function setBannerGlobally(enabled) {
    var item = currentItem();
    Core.setBannerEnabled(enabled, function () {
      if (state.ctx) {
        Core.syncBanner(item, state.ctx.itemId, state.emailNotes, null, function () {
          showToast(enabled ? "Banners above emails turned on" : "Banners above emails turned off");
        });
      } else {
        render();
      }
    });
  }

  // ---------------------------------------------------------------------------
  // Export / import
  // ---------------------------------------------------------------------------
  function exportJson() {
    var payload = JSON.stringify({ application: "Email Sticky Notes for Outlook", version: "2.1", exportedAt: new Date().toISOString(), notes: Core.readIndex() }, null, 2);
    try {
      var blob = new Blob([payload], { type: "application/json" });
      if (window.navigator && window.navigator.msSaveBlob) {
        window.navigator.msSaveBlob(blob, "outlook-sticky-notes-backup.json");
      } else {
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = "outlook-sticky-notes-backup.json";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(function () {
          URL.revokeObjectURL(url);
        }, 1000);
      }
      showToast("Backup downloaded");
    } catch (e) {
      copyText(payload, "Download blocked - backup copied to clipboard instead");
    }
  }

  function copyText(text, okMsg) {
    var done = function () {
      showToast(okMsg);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done, function () {
        legacyCopy(text) ? done() : showToast("Copy failed");
      });
    } else {
      legacyCopy(text) ? done() : showToast("Copy failed");
    }
  }

  function legacyCopy(text) {
    try {
      var ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.select();
      var ok = document.execCommand("copy");
      document.body.removeChild(ta);
      return ok;
    } catch (e) {
      return false;
    }
  }

  function importJson(file) {
    var reader = new FileReader();
    reader.onload = function () {
      var parsed;
      try {
        parsed = JSON.parse(String(reader.result));
      } catch (e) {
        showToast("That file isn't a valid backup.");
        return;
      }
      var incoming = Array.isArray(parsed) ? parsed : parsed && Array.isArray(parsed.notes) ? parsed.notes : [];
      incoming = incoming.filter(function (n) {
        return n && n.id && n.itemId;
      });
      if (incoming.length === 0) {
        showToast("No notes found in that file.");
        return;
      }
      var byId = {};
      Core.readIndex().concat(incoming).forEach(function (n) {
        if (!byId[n.id] || String(n.updatedAt) >= String(byId[n.id].updatedAt)) byId[n.id] = n;
      });
      var merged = Object.keys(byId).map(function (k) {
        return byId[k];
      });
      Core.writeIndex(merged, function () {
        var forThis = merged.filter(function (n) {
          return n.itemId === state.ctx.itemId;
        });
        state.settingsOpen = false;
        if (forThis.length > 0) {
          persistCurrent(forThis, "Imported " + incoming.length + " note(s)");
        } else {
          state.allNotes = Core.readIndex();
          showToast("Imported " + incoming.length + " note(s)");
        }
      });
    };
    reader.readAsText(file);
  }

  // ---------------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------------
  function opt(value, label, selected) {
    return '<option value="' + value + '"' + (selected === value ? " selected" : "") + ">" + label + "</option>";
  }

  function renderHeader(h) {
    h.push('<header class="app-header"><div class="header-top-row"><div class="brand-group">');
    h.push('<img src="./assets/icon-64.png" alt="" class="brand-icon-img" width="32" height="32" />');
    h.push('<div><h1 class="brand-title">Email Sticky Notes</h1><div class="storage-pill-row">');
    h.push(state.isInOutlook ? '<span class="status-pill status-office">Saved in your mailbox</span>' : '<span class="status-pill status-preview">Browser preview</span>');
    h.push("</div></div></div>");
    h.push('<div class="header-actions">');
    if (!state.noItem) {
      h.push('<button type="button" class="btn-primary-compact" data-action="new-note" id="btn-new-note">+ New Note</button>');
      h.push('<button type="button" class="btn-secondary-sm" data-action="quick-note" id="btn-quick-note" title="Stamp an instant timestamped follow-up sticky note">⚡ Quick</button>');
    }
    h.push('<button type="button" class="icon-btn" data-action="refresh" id="btn-refresh" title="Refresh notes">🔄</button>');
    h.push('<button type="button" class="icon-btn" data-action="toggle-theme" id="btn-toggle-theme" title="Toggle dark mode">' + (state.isDark ? "☀️" : "🌙") + "</button>");
    h.push('<button type="button" class="icon-btn" data-action="open-settings" id="btn-open-settings" title="Settings">⚙️</button>');
    h.push("</div></div>");
    h.push('<div class="scope-tabs">');
    h.push('<button type="button" class="scope-tab ' + (state.viewScope === "current_email" ? "active" : "") + '" data-action="scope" data-value="current_email"><span>This Email</span><span class="tab-badge">' + state.emailNotes.length + "</span></button>");
    h.push('<button type="button" class="scope-tab ' + (state.viewScope === "all_emails" ? "active" : "") + '" data-action="scope" data-value="all_emails"><span>All Annotated Emails</span><span class="tab-badge">' + state.allNotes.length + "</span></button>");
    h.push("</div></header>");
  }

  function renderEmailContext(h) {
    h.push('<section class="email-context-banner">');
    if (!state.isInOutlook) {
      h.push('<div class="demo-switcher-bar"><span class="demo-label">Sample emails:</span><div class="demo-pills">');
      DEMO_EMAILS.forEach(function (em, idx) {
        var count = state.allNotes.filter(function (n) {
          return n.itemId === em.itemId;
        }).length;
        h.push('<button type="button" class="demo-pill ' + (idx === state.demoIndex ? "active" : "") + '" data-action="demo" data-value="' + idx + '">Email #' + (idx + 1) + (count ? " (📝 " + count + ")" : "") + "</button>");
      });
      h.push("</div></div>");
    }
    var c = state.ctx;
    var hasNotes = state.emailNotes.length > 0;
    var hasUrgent = state.emailNotes.some(function (n) {
      return n.priority === "urgent";
    });
    h.push('<div class="email-meta-row"><div class="email-meta-text"><div class="email-subject-row">');
    h.push('<div class="email-subject" title="' + esc(c.subject) + '">' + esc(c.subject) + "</div>");
    if (hasNotes) {
      h.push('<div class="email-category-markers"><span class="outlook-category-pill cat-yellow" title="Category applied to this email in your Inbox">' + Core.NOTE_CATEGORY + "</span>");
      if (hasUrgent) h.push('<span class="outlook-category-pill cat-red">' + Core.URGENT_CATEGORY + "</span>");
      h.push("</div>");
    }
    h.push("</div>");
    if (c.senderName || c.senderEmail) {
      h.push('<div class="email-sender">' + esc(c.senderName) + (c.senderEmail ? " (" + esc(c.senderEmail) + ")" : "") + "</div>");
    }
    h.push("</div></div>");

    // Banner control row: lets the user remove/hide the notification above the email at any time
    h.push('<div class="banner-control-row">');
    if (!Core.isBannerEnabled()) {
      h.push('<span>🔕 Top-of-email notification is turned off.</span><div class="banner-actions-group"><button type="button" class="link-btn" data-action="banner-clear-now" id="btn-banner-clear-now" title="Force-clear any visible notification above this email">Clear Email Banner</button><button type="button" class="link-btn" data-action="banner-global-on" id="btn-banner-global-on">Turn On</button></div>');
    } else if (Core.isBannerDismissed(c.itemId)) {
      h.push('<span>🔕 Notification above this email is removed.</span><div class="banner-actions-group"><button type="button" class="link-btn" data-action="banner-show" id="btn-banner-show">Restore on this email</button><button type="button" class="link-btn" data-action="banner-global-off" id="btn-banner-global-off">Turn Off Always</button></div>');
    } else {
      h.push('<span>🔔 Notification above email:</span><div class="banner-actions-group"><button type="button" class="link-btn danger-link" data-action="banner-hide" id="btn-banner-hide">✕ Remove Above Email</button><button type="button" class="link-btn" data-action="banner-global-off" id="btn-banner-global-off">Turn Off Always</button></div>');
    }
    h.push("</div>");
    h.push("</section>");
  }

  function renderFilters(h) {
    h.push('<div class="search-filter-container">');
    h.push('<div class="search-input-wrapper"><input type="text" id="input-search" class="search-input" style="padding-left:10px;" placeholder="Search notes, tags, checklists, subjects..." value="' + esc(state.searchQuery) + '" /></div>');
    h.push('<div class="filter-controls-row"><div class="color-filter-dots">');
    h.push('<button type="button" class="color-dot-all ' + (state.selectedColor === "all" ? "selected" : "") + '" data-action="filter-color" data-value="all">All</button>');
    Object.keys(PALETTE).forEach(function (k) {
      h.push('<button type="button" class="color-dot ' + (state.selectedColor === k ? "selected" : "") + '" style="background-color:' + PALETTE[k].swatch + '" data-action="filter-color" data-value="' + k + '" title="' + PALETTE[k].label + '"></button>');
    });
    h.push('</div><div class="filter-select-group">');
    h.push('<button type="button" class="pin-filter-btn ' + (state.onlyPinned ? "active" : "") + '" data-action="toggle-only-pinned" id="btn-only-pinned" title="Only pinned notes">📌</button>');
    h.push('<select id="select-priority-filter" class="compact-select" title="Priority">' + opt("all", "Any Priority", state.selectedPriority) + opt("urgent", "Urgent", state.selectedPriority) + opt("high", "High", state.selectedPriority) + opt("medium", "Medium", state.selectedPriority) + opt("low", "Low", state.selectedPriority) + "</select>");
    h.push('<select id="select-sort" class="compact-select" title="Sort">' + opt("updated_desc", "Recently Updated", state.sortBy) + opt("created_desc", "Newest", state.sortBy) + opt("created_asc", "Oldest", state.sortBy) + opt("priority_desc", "Priority", state.sortBy) + "</select>");
    h.push("</div></div>");

    var tagMap = {};
    sourceNotes().forEach(function (n) {
      (n.tags || []).forEach(function (t) {
        tagMap[t] = true;
      });
    });
    var tags = Object.keys(tagMap).sort();
    if (tags.length > 0) {
      h.push('<div class="tag-filter-strip"><button type="button" class="tag-chip-filter ' + (state.selectedTag === "all" ? "active" : "") + '" data-action="filter-tag" data-value="all">All Tags</button>');
      tags.forEach(function (t) {
        h.push('<button type="button" class="tag-chip-filter ' + (state.selectedTag === t ? "active" : "") + '" data-action="filter-tag" data-value="' + esc(t) + '">#' + esc(t) + "</button>");
      });
      if (hasActiveFilters()) h.push('<button type="button" class="reset-filters-link" data-action="reset-filters">Reset</button>');
      h.push("</div>");
    }
    h.push("</div>");
  }

  function renderEditor(h) {
    var ed = state.editor;
    h.push('<div class="note-editor-card" id="note-editor">');
    h.push('<div class="editor-header"><span class="editor-heading">' + (ed.id ? "Edit Sticky Note" : "New Sticky Note") + '</span><div class="editor-header-controls">');
    h.push('<button type="button" class="icon-toggle-btn ' + (ed.isPinned ? "active" : "") + '" data-action="ed-toggle-pin" id="ed-pin" title="Pin to top">📌</button>');
    h.push('<button type="button" class="icon-toggle-btn" data-action="ed-cancel" title="Cancel">✕</button></div></div>');
    h.push('<input type="text" id="ed-title" class="editor-title-input" placeholder="Title (e.g. Follow up with Legal)" value="' + esc(ed.title) + '" />');
    h.push('<textarea id="ed-content" class="editor-body-textarea" rows="4" placeholder="Write your note... (Ctrl+Enter to save)">' + esc(ed.content) + "</textarea>");

    h.push('<div class="editor-section"><div class="editor-section-label">☑️ Checklist</div>');
    if (ed.checklist.length > 0) {
      h.push('<ul class="editor-checklist">');
      ed.checklist.forEach(function (c) {
        h.push('<li class="editor-checklist-item"><input type="checkbox" data-action="ed-chk-toggle" data-value="' + esc(c.id) + '"' + (c.completed ? " checked" : "") + " /><span>" + esc(c.text) + '</span><button type="button" class="remove-chk-btn" data-action="ed-chk-remove" data-value="' + esc(c.id) + '" title="Remove">✕</button></li>');
      });
      h.push("</ul>");
    }
    h.push('<div class="inline-add-row"><input type="text" id="ed-new-chk" class="compact-text-input" placeholder="Add a checklist item and press Enter" value="' + esc(ed.newChk) + '" /><button type="button" class="btn-secondary-sm" data-action="ed-chk-add" id="ed-chk-add">+ Add</button></div></div>');

    h.push('<div class="editor-options-grid"><div class="editor-field"><label>Color</label><div class="editor-color-swatches">');
    Object.keys(PALETTE).forEach(function (k) {
      h.push('<button type="button" class="swatch-btn ' + (ed.color === k ? "selected" : "") + '" style="background-color:' + PALETTE[k].swatch + '" data-action="ed-color" data-value="' + k + '" title="' + PALETTE[k].label + '"></button>');
    });
    h.push("</div></div>");
    h.push('<div class="editor-field"><label>Priority</label><select id="ed-priority" class="compact-select">' + opt("low", "Low", ed.priority) + opt("medium", "Medium", ed.priority) + opt("high", "High", ed.priority) + opt("urgent", "Urgent", ed.priority) + "</select></div></div>");

    h.push('<div class="editor-section"><div class="editor-section-label">🏷️ Tags</div><div class="preset-tags-wrap">');
    var allTags = PRESET_TAGS.concat(ed.tags.filter(function (t) { return PRESET_TAGS.indexOf(t) === -1; }));
    allTags.forEach(function (t) {
      h.push('<button type="button" class="tag-pill-selectable ' + (ed.tags.indexOf(t) !== -1 ? "selected" : "") + '" data-action="ed-tag" data-value="' + esc(t) + '">#' + esc(t) + "</button>");
    });
    h.push('</div><div class="inline-add-row"><input type="text" id="ed-new-tag" class="compact-text-input" placeholder="Custom tag" value="' + esc(ed.newTag) + '" /><button type="button" class="btn-secondary-sm" data-action="ed-tag-add">Add Tag</button></div></div>');

    h.push('<div class="editor-footer"><button type="button" class="btn-ghost" data-action="ed-cancel">Cancel</button><button type="button" class="btn-primary" data-action="ed-save" id="btn-save-editor">' + (ed.id ? "Update Note" : "Save Sticky Note") + "</button></div></div>");
  }

  function renderCard(h, note) {
    var p = PALETTE[note.color] || PALETTE.yellow;
    var pr = PRIORITY[note.priority] || PRIORITY.medium;
    var dark = state.isDark;
    var editable = Boolean(findCurrentNote(note.id));
    var id = esc(note.id);

    h.push('<article class="sticky-note-card" data-note-id="' + id + '" style="background-color:' + (dark ? p.dBg : p.lBg) + ";border-color:" + (dark ? p.dBd : p.lBd) + ";color:" + (dark ? p.dTx : p.lTx) + '">');
    h.push('<div class="sticky-note-top-strip" style="background-color:' + (dark ? p.dHd : p.lHd) + '"><div class="sticky-badges-left">');
    if (note.isPinned) h.push('<span class="pinned-indicator">📌 Pinned</span>');
    h.push('<span class="priority-badge" style="background-color:' + pr.color + '">' + pr.label + "</span></div>");
    h.push('<div class="sticky-card-actions">');
    h.push('<button type="button" class="card-icon-btn" data-action="copy-note" data-value="' + id + '" title="Copy text">📋</button>');
    if (editable) {
      h.push('<button type="button" class="card-icon-btn" data-action="pin-note" data-value="' + id + '" title="' + (note.isPinned ? "Unpin" : "Pin") + '">📌</button>');
      h.push('<button type="button" class="card-icon-btn" data-action="edit-note" data-value="' + id + '" title="Edit">✏️</button>');
      h.push('<button type="button" class="card-icon-btn delete-btn" data-action="delete-note" data-value="' + id + '" title="Delete">🗑️</button>');
    }
    h.push("</div></div>");

    h.push('<div class="sticky-note-body">');
    if (state.viewScope === "all_emails") {
      h.push('<div class="cross-email-ref">✉️ <span class="cross-email-subject" title="' + esc(note.emailSender) + '">' + esc(note.emailSubject) + "</span></div>");
    }
    if (note.title) h.push('<h3 class="sticky-note-title">' + esc(note.title) + "</h3>");
    if (note.content) h.push('<p class="sticky-note-content">' + esc(note.content) + "</p>");
    if (note.checklist && note.checklist.length > 0) {
      h.push('<div class="sticky-checklist-box">');
      note.checklist.forEach(function (c) {
        h.push('<label class="sticky-checklist-row ' + (c.completed ? "completed" : "") + '"><input type="checkbox" data-action="chk-toggle" data-note="' + id + '" data-value="' + esc(c.id) + '"' + (c.completed ? " checked" : "") + (editable ? "" : " disabled") + " /><span>" + esc(c.text) + "</span></label>");
      });
      h.push("</div>");
    }
    if (note.tags && note.tags.length > 0) {
      h.push('<div class="sticky-tags-row">');
      note.tags.forEach(function (t) {
        h.push('<button type="button" class="sticky-tag-chip" data-action="filter-tag" data-value="' + esc(t) + '">#' + esc(t) + "</button>");
      });
      h.push("</div>");
    }
    if (!editable) h.push('<div class="readonly-hint">Open this email in Outlook to edit or delete this note.</div>');
    h.push('<div class="sticky-note-footer"><span class="timestamp-meta" title="Created ' + esc(new Date(note.createdAt).toLocaleString()) + '">🕒 ' + esc(formatTimestamp(note.updatedAt)) + "</span>");
    h.push('<span class="author-meta">' + esc(note.author) + "</span></div></div></article>");
  }

  function renderSettings(h) {
    var bannerOn = Core.isBannerEnabled();
    h.push('<div class="modal-backdrop" data-action="close-settings-backdrop"><div class="settings-modal" id="settings-modal">');
    h.push('<div class="modal-header"><h2>Settings</h2><button type="button" class="icon-btn" data-action="close-settings" title="Close">✕</button></div><div class="modal-body">');

    h.push('<div class="settings-group"><label class="settings-group-title">Banner above emails</label>');
    h.push('<label class="radio-card ' + (bannerOn ? "selected" : "") + '"><input type="checkbox" id="cfg-banner-enabled"' + (bannerOn ? " checked" : "") + ' /><div><strong>Show a notification banner above emails that have notes</strong><p>Turn this off to stop the "📝 Sticky Note" message from appearing at the top of emails. The yellow category tag in your Inbox still marks which emails have notes. You can also hide the banner for a single email from the main screen.</p></div></label></div>');

    h.push('<div class="settings-group"><label class="settings-group-title">Where notes are stored</label><p class="settings-hint">Notes are saved privately inside your own mailbox (on each email, plus a small index that roams with your account). Nothing is sent to any other server.</p></div>');

    h.push('<div class="settings-group"><label class="settings-group-title">Backup</label><div class="backup-actions-row">');
    h.push('<button type="button" class="btn-secondary-sm" data-action="export-json" id="btn-export-json">⬇️ Export JSON</button>');
    h.push('<button type="button" class="btn-secondary-sm" data-action="copy-json">📋 Copy JSON</button>');
    h.push('<label class="btn-secondary-sm file-upload-label">⬆️ Import JSON<input type="file" id="cfg-import-file" accept=".json,application/json" hidden /></label>');
    h.push("</div></div>");

    h.push('<div class="modal-footer"><button type="button" class="btn-primary" data-action="close-settings">Done</button></div></div></div></div>');
  }

  function render() {
    document.documentElement.setAttribute("data-theme", state.isDark ? "dark" : "light");
    var root = document.getElementById("root");
    if (!root) return;

    var h = [];
    h.push('<div class="addin-shell">');
    renderHeader(h);

    if (state.noItem) {
      h.push('<main class="notes-scroll-area"><div class="empty-state-box" id="no-item-state"><h3>Select an email</h3><p>Pick a message in your Inbox to see or add its sticky notes.</p></div></main>');
    } else {
      renderEmailContext(h);
      renderFilters(h);
      h.push('<main class="notes-scroll-area">');
      if (state.editor) renderEditor(h);
      var list = filteredNotes();
      if (state.loading) {
        h.push('<div class="loading-skeleton-stack"><div class="skeleton-card"></div></div>');
      } else if (list.length === 0) {
        h.push('<div class="empty-state-box" id="empty-state">');
        if (hasActiveFilters()) {
          h.push('<h3>No matching notes</h3><p>Try a different search or filter.</p><button type="button" class="btn-secondary-sm" data-action="reset-filters">Clear filters</button>');
        } else if (state.viewScope === "all_emails") {
          h.push("<h3>No notes yet</h3><p>Notes you add to any email will be listed here.</p>");
        } else {
          h.push('<h3>No sticky notes on this email yet</h3><p>Attach a private, timestamped note, checklist, or reminder to this email.</p><button type="button" class="btn-primary" data-action="new-note">+ Create Sticky Note</button>');
        }
        h.push("</div>");
      } else {
        h.push('<div class="sticky-notes-list">');
        list.forEach(function (n) {
          renderCard(h, n);
        });
        h.push("</div>");
      }
      h.push("</main>");
    }

    if (state.toast) {
      h.push('<div class="toast-bar" id="toast"><span>' + esc(state.toast) + "</span>");
      if (state.toastUndo && state.lastDeleted) h.push('<button type="button" class="undo-toast-btn" data-action="undo-delete" id="btn-undo">↩ Undo</button>');
      h.push("</div>");
    }
    if (state.settingsOpen) renderSettings(h);
    h.push("</div>");

    root.innerHTML = h.join("");
  }

  // ---------------------------------------------------------------------------
  // Events (delegated once on #root)
  // ---------------------------------------------------------------------------
  function onClick(e) {
    var el = e.target;
    while (el && el !== document && !(el.getAttribute && el.getAttribute("data-action"))) el = el.parentNode;
    if (!el || el === document) return;
    var action = el.getAttribute("data-action");
    var value = el.getAttribute("data-value");

    // Checkbox actions are handled on "change".
    if (action === "chk-toggle" || action === "ed-chk-toggle") return;
    if (action === "close-settings-backdrop" && e.target !== el) return;

    if (state.editor) syncEditorInputs();

    switch (action) {
      case "new-note":
        openEditor(null);
        break;
      case "quick-note":
        var qNote = Core.createQuickNote(
          {
            conversationId: state.ctx.conversationId,
            subject: state.ctx.subject,
            from: {
              displayName: state.ctx.senderName,
              emailAddress: state.ctx.senderEmail,
            },
          },
          state.ctx.itemId
        );
        persistCurrent([qNote].concat(state.emailNotes), "⚡ Quick Note stamped on this email");
        break;
      case "refresh":
        loadCurrent();
        break;
      case "toggle-theme":
        state.isDark = !state.isDark;
        try {
          localStorage.setItem(THEME_KEY, state.isDark ? "dark" : "light");
        } catch (err) {}
        render();
        break;
      case "open-settings":
        state.settingsOpen = true;
        render();
        break;
      case "close-settings":
      case "close-settings-backdrop":
        state.settingsOpen = false;
        render();
        break;
      case "scope":
        state.viewScope = value;
        render();
        break;
      case "demo":
        state.demoIndex = parseInt(value, 10) || 0;
        state.editor = null;
        loadCurrent();
        break;
      case "filter-color":
        state.selectedColor = state.selectedColor === value && value !== "all" ? "all" : value;
        render();
        break;
      case "filter-tag":
        state.selectedTag = state.selectedTag === value && value !== "all" ? "all" : value;
        render();
        break;
      case "toggle-only-pinned":
        state.onlyPinned = !state.onlyPinned;
        render();
        break;
      case "reset-filters":
        state.searchQuery = "";
        state.selectedColor = "all";
        state.selectedPriority = "all";
        state.selectedTag = "all";
        state.onlyPinned = false;
        render();
        break;
      case "banner-hide":
        setBannerForThisEmail(true);
        break;
      case "banner-show":
        setBannerForThisEmail(false);
        break;
      case "banner-clear-now":
        Core.removeBanner(currentItem(), function () {
          showToast("Removed notification banner above email");
        });
        break;
      case "banner-global-on":
        setBannerGlobally(true);
        break;
      case "banner-global-off":
        setBannerGlobally(false);
        break;
      case "ed-cancel":
        state.editor = null;
        render();
        break;
      case "ed-save":
        saveEditor();
        break;
      case "ed-toggle-pin":
        state.editor.isPinned = !state.editor.isPinned;
        render();
        break;
      case "ed-color":
        state.editor.color = value;
        render();
        break;
      case "ed-tag":
        var i = state.editor.tags.indexOf(value);
        if (i === -1) state.editor.tags.push(value);
        else state.editor.tags.splice(i, 1);
        render();
        break;
      case "ed-tag-add":
        addCustomTag();
        render();
        break;
      case "ed-chk-add":
        addChecklistItem();
        render();
        var inp = document.getElementById("ed-new-chk");
        if (inp) inp.focus();
        break;
      case "ed-chk-remove":
        state.editor.checklist = state.editor.checklist.filter(function (c) {
          return c.id !== value;
        });
        render();
        break;
      case "pin-note":
        var pn = findCurrentNote(value);
        if (!pn) return;
        persistCurrent(
          state.emailNotes.map(function (n) {
            return n.id === value ? Object.assign({}, n, { isPinned: !n.isPinned, updatedAt: new Date().toISOString() }) : n;
          }),
          pn.isPinned ? "Note unpinned" : "Note pinned"
        );
        break;
      case "edit-note":
        var en = findCurrentNote(value);
        if (en) {
          openEditor(en);
          var edEl = document.getElementById("note-editor");
          if (edEl && edEl.scrollIntoView) edEl.scrollIntoView();
        }
        break;
      case "delete-note":
        var dn = findCurrentNote(value);
        if (!dn) return;
        state.lastDeleted = { note: dn, itemId: state.ctx.itemId };
        if (state.editor && state.editor.id === value) state.editor = null;
        persistCurrent(
          state.emailNotes.filter(function (n) {
            return n.id !== value;
          }),
          "Sticky note deleted",
          true
        );
        break;
      case "undo-delete":
        var ld = state.lastDeleted;
        state.lastDeleted = null;
        if (ld && ld.itemId === state.ctx.itemId) {
          persistCurrent([ld.note].concat(state.emailNotes), "Note restored");
        }
        break;
      case "copy-note":
        var cn = sourceNotes().find(function (n) {
          return n.id === value;
        });
        if (cn) {
          var text = [cn.title, cn.content].concat((cn.checklist || []).map(function (c) { return (c.completed ? "[x] " : "[ ] ") + c.text; })).filter(Boolean).join("\n");
          copyText(text, "Note copied");
        }
        break;
      case "export-json":
        exportJson();
        break;
      case "copy-json":
        copyText(JSON.stringify({ notes: Core.readIndex() }, null, 2), "Backup JSON copied");
        break;
    }
  }

  function onChange(e) {
    var el = e.target;
    var action = el.getAttribute && el.getAttribute("data-action");
    if (action === "chk-toggle") {
      var noteId = el.getAttribute("data-note");
      var chkId = el.getAttribute("data-value");
      if (!findCurrentNote(noteId)) return;
      persistCurrent(
        state.emailNotes.map(function (n) {
          if (n.id !== noteId) return n;
          return Object.assign({}, n, {
            checklist: (n.checklist || []).map(function (c) {
              return c.id === chkId ? Object.assign({}, c, { completed: !c.completed }) : c;
            }),
            updatedAt: new Date().toISOString(),
          });
        }),
        null
      );
      return;
    }
    if (action === "ed-chk-toggle" && state.editor) {
      syncEditorInputs();
      var cid = el.getAttribute("data-value");
      state.editor.checklist.forEach(function (c) {
        if (c.id === cid) c.completed = !c.completed;
      });
      render();
      return;
    }
    if (el.id === "select-priority-filter") {
      state.selectedPriority = el.value;
      render();
    } else if (el.id === "select-sort") {
      state.sortBy = el.value;
      render();
    } else if (el.id === "cfg-banner-enabled") {
      setBannerGlobally(el.checked);
    } else if (el.id === "cfg-import-file" && el.files && el.files[0]) {
      importJson(el.files[0]);
    }
  }

  function onInput(e) {
    var el = e.target;
    if (el.id !== "input-search") return;
    state.searchQuery = el.value;
    var pos = el.selectionStart;
    render();
    var again = document.getElementById("input-search");
    if (again) {
      again.focus();
      try {
        again.setSelectionRange(pos, pos);
      } catch (err) {}
    }
  }

  function onKeyDown(e) {
    var el = e.target;
    if (state.editor) {
      if (e.key === "Escape") {
        state.editor = null;
        render();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        saveEditor();
        return;
      }
      if (e.key === "Enter" && el.id === "ed-new-chk") {
        e.preventDefault();
        syncEditorInputs();
        addChecklistItem();
        render();
        var inp = document.getElementById("ed-new-chk");
        if (inp) inp.focus();
        return;
      }
      if (e.key === "Enter" && el.id === "ed-new-tag") {
        e.preventDefault();
        syncEditorInputs();
        addCustomTag();
        render();
        return;
      }
    } else if (e.key === "Escape" && state.settingsOpen) {
      state.settingsOpen = false;
      render();
    }
  }

  function bindEvents() {
    var root = document.getElementById("root");
    root.addEventListener("click", onClick);
    root.addEventListener("change", onChange);
    root.addEventListener("input", onInput);
    root.addEventListener("keydown", onKeyDown);

    // Pick up notes added by the ribbon "Quick Note" button while the pane is open.
    var lastFocusLoad = 0;
    window.addEventListener("focus", function () {
      if (!state.isInOutlook || state.editor || Date.now() - lastFocusLoad < 2000) return;
      lastFocusLoad = Date.now();
      loadCurrent();
    });
    if (Core.onCrossFrameUpdate) {
      Core.onCrossFrameUpdate(function () {
        if (!state.editor) loadCurrent();
      });
    }
  }

  // ---------------------------------------------------------------------------
  // Boot
  // ---------------------------------------------------------------------------
  function start(isOutlook) {
    state.isInOutlook = isOutlook;
    bindEvents();
    if (isOutlook && Office.context.mailbox.addHandlerAsync) {
      try {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
          state.editor = null;
          state.lastDeleted = null;
          loadCurrent();
        });
      } catch (e) {}
    }
    loadCurrent();
  }

  window.__stickyNotesApp = { state: state, reload: loadCurrent };

  var started = false;
  function startOnce(isOutlook) {
    if (started) return;
    started = true;
    start(isOutlook);
  }

  if (window.Office && typeof Office.onReady === "function") {
    // Outside Office, onReady still resolves (with host = null), so no fallback timer is needed.
    Office.onReady(function (info) {
      startOnce(Boolean(info && info.host === Office.HostType.Outlook && Office.context && Office.context.mailbox));
    });
  } else {
    startOnce(false);
  }
})();
