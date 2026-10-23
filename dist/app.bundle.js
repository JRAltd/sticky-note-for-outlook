/**
 * Email Sticky Notes for Outlook - Taskpane (v2.2)
 * Clean Fluent 2 UI with crisp SVG icons, spacious action buttons, and high readability.
 * Plain ES5, no build step. Depends on notes-core.js (window.StickyNotesCore).
 */
(function () {
  "use strict";

  var Core = window.StickyNotesCore;
  var THEME_KEY = "outlook_sticky_notes_theme";
  var DEMO_SEED_FLAG_KEY = "outlook_sticky_notes_demo_seeded_v2";

  // Crisp inline Fluent-style SVG icons (16x16 viewBox)
  var ICONS = {
    plus: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>',
    bolt: '<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>',
    refresh: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>',
    sun: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>',
    moon: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>',
    settings: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>',
    search: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>',
    sliders: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><line x1="4" y1="21" x2="4" y2="14"/><line x1="4" y1="10" x2="4" y2="3"/><line x1="12" y1="21" x2="12" y2="12"/><line x1="12" y1="8" x2="12" y2="3"/><line x1="20" y1="21" x2="20" y2="16"/><line x1="20" y1="12" x2="20" y2="3"/><line x1="1" y1="14" x2="7" y2="14"/><line x1="9" y1="8" x2="15" y2="8"/><line x1="17" y1="16" x2="23" y2="16"/></svg>',
    pin: '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="17" x2="12" y2="22"/><path d="M5 17h14v-1.76a2 2 0 0 0-1.11-1.79l-1.78-.9A2 2 0 0 1 15 10.76V6h1a2 2 0 0 0 0-4H8a2 2 0 0 0 0 4h1v4.76a2 2 0 0 1-1.11 1.79l-1.78.9A2 2 0 0 0 5 15.24Z"/></svg>',
    copy: '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>',
    edit: '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>',
    trash: '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>',
    close: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>',
    bell: '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>',
    bellOff: '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M13.73 21a2 2 0 0 1-3.46 0"/><path d="M18.63 13A17.89 17.89 0 0 1 18 8"/><path d="M6.26 6.26A5.86 5.86 0 0 0 6 8c0 7-3 9-3 9h14"/><path d="M18 8a6 6 0 0 0-9.33-5"/><line x1="1" y1="1" x2="23" y2="23"/></svg>',
    download: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>',
    upload: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>',
    mail: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/></svg>',
    clock: '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>',
  };

  var PALETTE = {
    yellow: { label: "Classic Yellow", swatch: "#FFE255", lBg: "#FFFBEB", lHd: "#FEF08A", lBd: "#FACC15", lTx: "#3F2E05", dBg: "#362F12", dHd: "#473D16", dBd: "#A1841E", dTx: "#FEF9C3" },
    amber: { label: "Warm Amber", swatch: "#FFB74D", lBg: "#FFF7ED", lHd: "#FED7AA", lBd: "#FB923C", lTx: "#431407", dBg: "#382412", dHd: "#4A2F17", dBd: "#B4631E", dTx: "#FFEDD5" },
    green: { label: "Mint Green", swatch: "#81C784", lBg: "#F0FDF4", lHd: "#BBF7D0", lBd: "#4ADE80", lTx: "#064E3B", dBg: "#143121", dHd: "#1B422C", dBd: "#348F58", dTx: "#DCFCE7" },
    blue: { label: "Sky Blue", swatch: "#64B5F6", lBg: "#F0F9FF", lHd: "#BAE6FD", lBd: "#38BDF8", lTx: "#0C4A6E", dBg: "#12293B", dHd: "#18374F", dBd: "#2D78A8", dTx: "#E0F2FE" },
    pink: { label: "Rose Pink", swatch: "#F06292", lBg: "#FDF2F8", lHd: "#FBCFE8", lBd: "#F472B6", lTx: "#500724", dBg: "#371627", dHd: "#491D34", dBd: "#A63A6C", dTx: "#FCE7F3" },
    purple: { label: "Lavender", swatch: "#BA68C8", lBg: "#FAF5FF", lHd: "#E9D5FF", lBd: "#C084FC", lTx: "#3B0764", dBg: "#28173A", dHd: "#361F4E", dBd: "#7E48AD", dTx: "#F3E8FF" },
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
    filtersExpanded: false,
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
  // Email context & storage
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
        if (seq !== loadSeq) return;
        state.emailNotes = sortPinnedFirst(notes || []);
        state.allNotes = Core.readIndex();
        state.loading = false;
        render();
        if (item) Core.syncMarkers(item, ctx.itemId, state.emailNotes, null, function () {});
      });
    });
  }

  function persistCurrent(nextNotes, toastMsg, withUndo) {
    var item = currentItem();
    var key = state.ctx.itemId;
    Core.saveItemNotes(item, key, nextNotes, function (ok) {
      if (!ok) {
        showToast("Couldn't save to your mailbox. Please try again.");
        return;
      }
      Core.replaceItemInIndex(key, nextNotes, function () {
        if (key !== state.ctx.itemId) return;
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

  function countActiveFilterPills() {
    var c = 0;
    if (state.selectedColor !== "all") c++;
    if (state.selectedPriority !== "all") c++;
    if (state.selectedTag !== "all") c++;
    if (state.onlyPinned) c++;
    return c;
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
        showToast(hidden ? "Banner removed above this email" : "Banner restored above this email");
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
    var payload = JSON.stringify({ application: "Email Sticky Notes for Outlook", version: "2.2", exportedAt: new Date().toISOString(), notes: Core.readIndex() }, null, 2);
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
    h.push('<header class="app-header">');

    // Top row: Brand + clean utility icon buttons
    h.push('<div class="header-top-row">');
    h.push('<div class="brand-group">');
    h.push('<img src="./assets/icon-64.png" alt="" class="brand-icon-img" width="34" height="34" />');
    h.push('<div><h1 class="brand-title">Email Sticky Notes</h1><div class="storage-pill-row">');
    h.push(state.isInOutlook ? '<span class="status-pill status-office"><span class="status-dot"></span>Saved in Mailbox</span>' : '<span class="status-pill status-preview"><span class="status-dot"></span>Interactive Preview</span>');
    h.push("</div></div></div>");

    h.push('<div class="header-utility-btns">');
    h.push('<button type="button" class="icon-btn" data-action="refresh" id="btn-refresh" title="Refresh notes">' + ICONS.refresh + "</button>");
    h.push('<button type="button" class="icon-btn" data-action="toggle-theme" id="btn-toggle-theme" title="' + (state.isDark ? "Switch to Light Mode" : "Switch to Dark Mode") + '">' + (state.isDark ? ICONS.sun : ICONS.moon) + "</button>");
    h.push('<button type="button" class="icon-btn" data-action="open-settings" id="btn-open-settings" title="Settings & Backup">' + ICONS.settings + "</button>");
    h.push("</div></div>");

    // Primary Action Row: Two spacious, easy-to-read buttons
    if (!state.noItem) {
      h.push('<div class="header-cta-row">');
      h.push('<button type="button" class="btn-cta-new" data-action="new-note" id="btn-new-note">' + ICONS.plus + "<span>New Sticky Note</span></button>");
      h.push('<button type="button" class="btn-cta-quick" data-action="quick-note" id="btn-quick-note" title="Instantly stamp a follow-up sticky note with 1 click">' + ICONS.bolt + "<span>Quick Note</span></button>");
      h.push("</div>");
    }

    // Segmented Scope Pill Tabs
    h.push('<div class="scope-segmented-bar" role="tablist">');
    h.push('<button type="button" class="scope-tab ' + (state.viewScope === "current_email" ? "active" : "") + '" data-action="scope" data-value="current_email"><span>This Email</span><span class="tab-badge">' + state.emailNotes.length + "</span></button>");
    h.push('<button type="button" class="scope-tab ' + (state.viewScope === "all_emails" ? "active" : "") + '" data-action="scope" data-value="all_emails"><span>All Emails</span><span class="tab-badge">' + state.allNotes.length + "</span></button>");
    h.push("</div></header>");
  }

  function renderEmailContext(h) {
    h.push('<section class="email-context-banner">');
    if (!state.isInOutlook) {
      h.push('<div class="demo-switcher-bar"><span class="demo-label">Sample Email:</span><div class="demo-pills">');
      DEMO_EMAILS.forEach(function (em, idx) {
        var count = state.allNotes.filter(function (n) {
          return n.itemId === em.itemId;
        }).length;
        h.push('<button type="button" class="demo-pill ' + (idx === state.demoIndex ? "active" : "") + '" data-action="demo" data-value="' + idx + '">Email #' + (idx + 1) + (count ? " (" + count + ")" : "") + "</button>");
      });
      h.push("</div></div>");
    }

    var c = state.ctx;
    var hasNotes = state.emailNotes.length > 0;
    var hasUrgent = state.emailNotes.some(function (n) {
      return n.priority === "urgent";
    });
    h.push('<div class="email-meta-row">');
    h.push('<div class="email-meta-icon">' + ICONS.mail + "</div>");
    h.push('<div class="email-meta-text"><div class="email-subject-row">');
    h.push('<div class="email-subject" title="' + esc(c.subject) + '">' + esc(c.subject) + "</div>");
    if (hasNotes) {
      h.push('<div class="email-category-markers"><span class="outlook-category-pill cat-yellow" title="Category badge shown in your Outlook Inbox">' + Core.NOTE_CATEGORY + "</span>");
      if (hasUrgent) h.push('<span class="outlook-category-pill cat-red">' + Core.URGENT_CATEGORY + "</span>");
      h.push("</div>");
    }
    h.push("</div>");
    if (c.senderName || c.senderEmail) {
      h.push('<div class="email-sender">' + esc(c.senderName) + (c.senderEmail ? " &lt;" + esc(c.senderEmail) + "&gt;" : "") + "</div>");
    }
    h.push("</div></div>");

    // Clean, pill-button Banner Control Bar
    h.push('<div class="banner-control-row">');
    if (!Core.isBannerEnabled()) {
      h.push('<div class="banner-status-label">' + ICONS.bellOff + "<span>Email banner is off</span></div>");
      h.push('<div class="banner-actions-group">');
      h.push('<button type="button" class="banner-pill-btn" data-action="banner-clear-now" id="btn-banner-clear-now">Clear Now</button>');
      h.push('<button type="button" class="banner-pill-btn primary" data-action="banner-global-on" id="btn-banner-global-on">Turn On</button>');
      h.push("</div>");
    } else if (Core.isBannerDismissed(c.itemId)) {
      h.push('<div class="banner-status-label">' + ICONS.bellOff + "<span>Banner removed on this email</span></div>");
      h.push('<div class="banner-actions-group">');
      h.push('<button type="button" class="banner-pill-btn primary" data-action="banner-show" id="btn-banner-show">Restore Banner</button>');
      h.push('<button type="button" class="banner-pill-btn" data-action="banner-global-off" id="btn-banner-global-off">Turn Off Always</button>');
      h.push("</div>");
    } else {
      h.push('<div class="banner-status-label">' + ICONS.bell + "<span>Notification above email</span></div>");
      h.push('<div class="banner-actions-group">');
      h.push('<button type="button" class="banner-pill-btn danger" data-action="banner-hide" id="btn-banner-hide">' + ICONS.close + "<span>Remove Banner</span></button>");
      h.push('<button type="button" class="banner-pill-btn" data-action="banner-global-off" id="btn-banner-global-off">Turn Off Always</button>');
      h.push("</div>");
    }
    h.push("</div>");
    h.push("</section>");
  }

  function renderFilters(h) {
    var activeCount = countActiveFilterPills();
    var showDrawer = state.filtersExpanded || activeCount > 0;

    h.push('<div class="search-filter-container">');
    h.push('<div class="search-bar-row">');
    h.push('<div class="search-input-wrapper">');
    h.push('<span class="search-icon">' + ICONS.search + "</span>");
    h.push('<input type="text" id="input-search" class="search-input" placeholder="Search notes, tags, checklists, or subjects..." value="' + esc(state.searchQuery) + '" />');
    if (state.searchQuery) {
      h.push('<button type="button" class="clear-search-btn" data-action="clear-search" title="Clear search">' + ICONS.close + "</button>");
    }
    h.push("</div>");
    h.push('<button type="button" class="filter-drawer-toggle ' + (showDrawer ? "active" : "") + '" data-action="toggle-filter-drawer" id="btn-toggle-filters" title="Show or hide color, priority, and tag filters">' + ICONS.sliders + "<span>Filters" + (activeCount > 0 ? " (" + activeCount + ")" : "") + "</span></button>");
    h.push("</div>");

    if (showDrawer) {
      h.push('<div class="filter-drawer-body">');
      h.push('<div class="filter-controls-row"><div class="color-filter-dots">');
      h.push('<button type="button" class="color-dot-all ' + (state.selectedColor === "all" ? "selected" : "") + '" data-action="filter-color" data-value="all">All Colors</button>');
      Object.keys(PALETTE).forEach(function (k) {
        h.push('<button type="button" class="color-dot ' + (state.selectedColor === k ? "selected" : "") + '" style="background-color:' + PALETTE[k].swatch + '" data-action="filter-color" data-value="' + k + '" title="' + PALETTE[k].label + '"></button>');
      });
      h.push('</div><div class="filter-select-group">');
      h.push('<button type="button" class="pin-filter-btn ' + (state.onlyPinned ? "active" : "") + '" data-action="toggle-only-pinned" id="btn-only-pinned" title="Show only pinned notes">' + ICONS.pin + "<span>Pinned</span></button>");
      h.push('<select id="select-priority-filter" class="compact-select" title="Priority">' + opt("all", "Any Priority", state.selectedPriority) + opt("urgent", "Urgent", state.selectedPriority) + opt("high", "High", state.selectedPriority) + opt("medium", "Medium", state.selectedPriority) + opt("low", "Low", state.selectedPriority) + "</select>");
      h.push('<select id="select-sort" class="compact-select" title="Sort">' + opt("updated_desc", "Latest Updated", state.sortBy) + opt("created_desc", "Newest Created", state.sortBy) + opt("created_asc", "Oldest Created", state.sortBy) + opt("priority_desc", "Highest Priority", state.sortBy) + "</select>");
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
        if (hasActiveFilters()) h.push('<button type="button" class="reset-filters-link" data-action="reset-filters">Reset All</button>');
        h.push("</div>");
      }
      h.push("</div>");
    }
    h.push("</div>");
  }

  function renderEditor(h) {
    var ed = state.editor;
    h.push('<div class="note-editor-card" id="note-editor">');
    h.push('<div class="editor-header"><span class="editor-heading">' + (ed.id ? "Edit Sticky Note" : "Create Sticky Note") + '</span><div class="editor-header-controls">');
    h.push('<button type="button" class="pill-action-btn ' + (ed.isPinned ? "active" : "") + '" data-action="ed-toggle-pin" id="ed-pin" title="Pin to top of email">' + ICONS.pin + "<span>" + (ed.isPinned ? "Pinned" : "Pin") + "</span></button>");
    h.push('<button type="button" class="icon-toggle-btn" data-action="ed-cancel" title="Close editor">' + ICONS.close + "</button></div></div>");

    h.push('<input type="text" id="ed-title" class="editor-title-input" placeholder="Note title (e.g. Follow up with Legal)..." value="' + esc(ed.title) + '" />');
    h.push('<textarea id="ed-content" class="editor-body-textarea" rows="4" placeholder="Write your sticky note details here... (Ctrl+Enter to save)">' + esc(ed.content) + "</textarea>");

    h.push('<div class="editor-options-grid"><div class="editor-field"><label>Sticky Color</label><div class="editor-color-swatches">');
    Object.keys(PALETTE).forEach(function (k) {
      h.push('<button type="button" class="swatch-btn ' + (ed.color === k ? "selected" : "") + '" style="background-color:' + PALETTE[k].swatch + '" data-action="ed-color" data-value="' + k + '" title="' + PALETTE[k].label + '"></button>');
    });
    h.push("</div></div>");
    h.push('<div class="editor-field"><label>Priority Level</label><select id="ed-priority" class="compact-select">' + opt("low", "Low Priority", ed.priority) + opt("medium", "Medium Priority", ed.priority) + opt("high", "High Priority", ed.priority) + opt("urgent", "Urgent Priority", ed.priority) + "</select></div></div>");

    h.push('<div class="editor-section"><div class="editor-section-label">Action Checklist</div>');
    if (ed.checklist.length > 0) {
      h.push('<ul class="editor-checklist">');
      ed.checklist.forEach(function (c) {
        h.push('<li class="editor-checklist-item"><input type="checkbox" data-action="ed-chk-toggle" data-value="' + esc(c.id) + '"' + (c.completed ? " checked" : "") + " /><span>" + esc(c.text) + '</span><button type="button" class="remove-chk-btn" data-action="ed-chk-remove" data-value="' + esc(c.id) + '" title="Remove item">' + ICONS.close + "</button></li>");
      });
      h.push("</ul>");
    }
    h.push('<div class="inline-add-row"><input type="text" id="ed-new-chk" class="compact-text-input" placeholder="Add a checklist item and press Enter..." value="' + esc(ed.newChk) + '" /><button type="button" class="btn-secondary-sm" data-action="ed-chk-add" id="ed-chk-add">+ Add Item</button></div></div>');

    h.push('<div class="editor-section"><div class="editor-section-label">Tags</div><div class="preset-tags-wrap">');
    var allTags = PRESET_TAGS.concat(ed.tags.filter(function (t) { return PRESET_TAGS.indexOf(t) === -1; }));
    allTags.forEach(function (t) {
      h.push('<button type="button" class="tag-pill-selectable ' + (ed.tags.indexOf(t) !== -1 ? "selected" : "") + '" data-action="ed-tag" data-value="' + esc(t) + '">#' + esc(t) + "</button>");
    });
    h.push('</div><div class="inline-add-row"><input type="text" id="ed-new-tag" class="compact-text-input" placeholder="Add custom tag..." value="' + esc(ed.newTag) + '" /><button type="button" class="btn-secondary-sm" data-action="ed-tag-add">+ Add Tag</button></div></div>');

    h.push('<div class="editor-footer"><button type="button" class="btn-ghost" data-action="ed-cancel">Cancel</button><button type="button" class="btn-primary" data-action="ed-save" id="btn-save-editor">' + (ed.id ? "Save Changes" : "Save Sticky Note") + "</button></div></div>");
  }

  function renderCard(h, note) {
    var p = PALETTE[note.color] || PALETTE.yellow;
    var pr = PRIORITY[note.priority] || PRIORITY.medium;
    var dark = state.isDark;
    var editable = Boolean(findCurrentNote(note.id));
    var id = esc(note.id);

    h.push('<article class="sticky-note-card" data-note-id="' + id + '" style="background-color:' + (dark ? p.dBg : p.lBg) + ";border-color:" + (dark ? p.dBd : p.lBd) + ";color:" + (dark ? p.dTx : p.lTx) + '">');
    h.push('<div class="sticky-note-top-strip" style="background-color:' + (dark ? p.dHd : p.lHd) + '"><div class="sticky-badges-left">');
    if (note.isPinned) h.push('<span class="pinned-indicator">' + ICONS.pin + " Pinned</span>");
    h.push('<span class="priority-badge" style="background-color:' + pr.color + '">' + pr.label + "</span></div>");

    h.push('<div class="sticky-card-actions">');
    h.push('<button type="button" class="card-action-pill" data-action="copy-note" data-value="' + id + '" title="Copy note text">' + ICONS.copy + "<span>Copy</span></button>");
    if (editable) {
      h.push('<button type="button" class="card-action-pill ' + (note.isPinned ? "active" : "") + '" data-action="pin-note" data-value="' + id + '" title="' + (note.isPinned ? "Unpin note" : "Pin to top") + '">' + ICONS.pin + "</button>");
      h.push('<button type="button" class="card-action-pill" data-action="edit-note" data-value="' + id + '" title="Edit note">' + ICONS.edit + "<span>Edit</span></button>");
      h.push('<button type="button" class="card-action-pill delete-btn" data-action="delete-note" data-value="' + id + '" title="Delete note">' + ICONS.trash + "</button>");
    }
    h.push("</div></div>");

    h.push('<div class="sticky-note-body">');
    if (state.viewScope === "all_emails") {
      h.push('<div class="cross-email-ref">' + ICONS.mail + ' <span class="cross-email-subject" title="' + esc(note.emailSender) + '">' + esc(note.emailSubject) + "</span></div>");
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
    h.push('<div class="sticky-note-footer"><span class="timestamp-meta" title="Created ' + esc(new Date(note.createdAt).toLocaleString()) + '">' + ICONS.clock + " <span>" + esc(formatTimestamp(note.updatedAt)) + "</span></span>");
    h.push('<span class="author-meta">' + esc(note.author) + "</span></div></div></article>");
  }

  function renderSettings(h) {
    var bannerOn = Core.isBannerEnabled();
    h.push('<div class="modal-backdrop" data-action="close-settings-backdrop"><div class="settings-modal" id="settings-modal">');
    h.push('<div class="modal-header"><h2>Settings &amp; Backup</h2><button type="button" class="icon-btn" data-action="close-settings" title="Close">' + ICONS.close + '</button></div><div class="modal-body">');

    h.push('<div class="settings-group"><label class="settings-group-title">Notification Banner Above Emails</label>');
    h.push('<label class="radio-card ' + (bannerOn ? "selected" : "") + '"><input type="checkbox" id="cfg-banner-enabled"' + (bannerOn ? " checked" : "") + ' /><div><strong>Show summary banner at the top of annotated emails</strong><p>Uncheck this to stop the "Sticky Note" notification bar from appearing above your emails. The yellow category badge in your Inbox will still mark emails that have notes.</p></div></label></div>');

    h.push('<div class="settings-group"><label class="settings-group-title">Backup &amp; Restore (' + state.allNotes.length + ' notes)</label><div class="backup-actions-row">');
    h.push('<button type="button" class="btn-secondary-sm" data-action="export-json" id="btn-export-json">' + ICONS.download + "<span>Export JSON</span></button>");
    h.push('<button type="button" class="btn-secondary-sm" data-action="copy-json">' + ICONS.copy + "<span>Copy Backup</span></button>");
    h.push('<label class="btn-secondary-sm file-upload-label">' + ICONS.upload + '<span>Import JSON</span><input type="file" id="cfg-import-file" accept=".json,application/json" hidden /></label>');
    h.push("</div></div>");

    h.push('<div class="settings-group"><label class="settings-group-title">Privacy &amp; Storage</label><p class="settings-hint">Your sticky notes are stored privately inside your own Microsoft 365 / Exchange mailbox on each email and never leave your account.</p></div>');

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
      h.push('<main class="notes-scroll-area"><div class="empty-state-box" id="no-item-state"><h3>Select an email</h3><p>Pick any message in your Inbox to view or attach sticky notes.</p></div></main>');
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
          h.push('<h3>No matching sticky notes</h3><p>Try clearing your search or active filter pills.</p><button type="button" class="btn-secondary-sm" data-action="reset-filters">Clear All Filters</button>');
        } else if (state.viewScope === "all_emails") {
          h.push("<h3>No annotated emails yet</h3><p>Sticky notes you attach to any email will appear here.</p>");
        } else {
          h.push('<h3>No sticky notes on this email</h3><p>Attach a private, color-coded note, checklist, or follow-up reminder to this message.</p><button type="button" class="btn-primary" data-action="new-note">' + ICONS.plus + "<span>Create Sticky Note</span></button>");
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
      if (state.toastUndo && state.lastDeleted) h.push('<button type="button" class="undo-toast-btn" data-action="undo-delete" id="btn-undo">Undo</button>');
      h.push("</div>");
    }
    if (state.settingsOpen) renderSettings(h);
    h.push("</div>");

    root.innerHTML = h.join("");
  }

  // ---------------------------------------------------------------------------
  // Events
  // ---------------------------------------------------------------------------
  function onClick(e) {
    var el = e.target;
    while (el && el !== document && !(el.getAttribute && el.getAttribute("data-action"))) el = el.parentNode;
    if (!el || el === document) return;
    var action = el.getAttribute("data-action");
    var value = el.getAttribute("data-value");

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
        persistCurrent([qNote].concat(state.emailNotes), "Quick Note stamped on this email");
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
      case "clear-search":
        state.searchQuery = "";
        render();
        break;
      case "toggle-filter-drawer":
        state.filtersExpanded = !state.filtersExpanded;
        render();
        break;
      case "filter-color":
        state.selectedColor = state.selectedColor === value && value !== "all" ? "all" : value;
        render();
        break;
      case "filter-tag":
        state.selectedTag = state.selectedTag === value && value !== "all" ? "all" : value;
        state.filtersExpanded = true;
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
          copyText(text, "Note copied to clipboard");
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
    var fallbackTimer = setTimeout(function () {
      startOnce(false);
    }, 800);
    Office.onReady(function (info) {
      clearTimeout(fallbackTimer);
      startOnce(Boolean(info && info.host === Office.HostType.Outlook && Office.context && Office.context.mailbox));
    });
  } else {
    startOnce(false);
  }
})();
