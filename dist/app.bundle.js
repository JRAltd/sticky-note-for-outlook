/**
 * Email Sticky Notes for Outlook - Taskpane & Dashboard Bundle (v2.3)
 * Clean Fluent 2 UI with instant optimistic updates, zero focus-loss rendering,
 * full cross-email editing, reminders, thread-scope notes, and 1-click templates.
 * Plain ES5, no build step required. Depends on notes-core.js (window.StickyNotesCore).
 */
(function () {
  "use strict";

  var Core = window.StickyNotesCore;
  var THEME_KEY = "outlook_sticky_notes_theme";
  var DEMO_SEED_FLAG_KEY = "outlook_sticky_notes_demo_seeded_v2";

  // Crisp inline Fluent-style SVG icons
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
    thread: '<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>',
    sparkles: '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="m12 3-1.9 5.8a2 2 0 0 1-1.3 1.3L3 12l5.8 1.9a2 2 0 0 1 1.3 1.3L12 21l1.9-5.8a2 2 0 0 1 1.3-1.3L21 12l-5.8-1.9a2 2 0 0 1-1.3-1.3Z"/></svg>',
    fileText: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>',
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

  var PRESET_TAGS = [
    "follow-up",
    "waiting",
    "action-required",
    "decision",
    "client",
    "billing",
    "important",
    "meeting-notes",
  ];

  var DEMO_EMAILS = [
    {
      itemId: "demo-msg-001-q4-contract",
      conversationId: "conv-q4-enterprise-renewal",
      subject: "Q4 Enterprise License Renewal & Pricing Proposal",
      senderName: "Sarah Jenkins (VP Procurement)",
      senderEmail: "s.jenkins@contoso-enterprise.com",
      userDisplayName: "Outlook User",
    },
    {
      itemId: "demo-msg-002-security-audit",
      conversationId: "conv-soc2-security-review",
      subject: "Action Required: SOC2 Type II Architecture Sign-off",
      senderName: "Marcus Vance (Security Engineering)",
      senderEmail: "mvance@cloudsec-ops.io",
      userDisplayName: "Outlook User",
    },
    {
      itemId: "demo-msg-003-board-deck",
      conversationId: "conv-board-deck-prep",
      subject: "Draft Board Deck Slides & Financial Appendix",
      senderName: "Elena Rostova (CFO)",
      senderEmail: "elena@northwind-ventures.com",
      userDisplayName: "Outlook User",
    },
  ];

  var QUICK_TEMPLATES = {
    followup: {
      title: "Follow-up Required",
      content: "Follow up with sender regarding open action items.",
      color: "yellow",
      priority: "high",
      scope: "message",
      tags: ["follow-up", "action-required"],
      checklist: [{ text: "Review attachments & reply", completed: false }],
      isPinned: true,
    },
    approval: {
      title: "Waiting on Stakeholder Approval",
      content: "Holding reply until internal sign-off is confirmed.",
      color: "amber",
      priority: "medium",
      scope: "conversation",
      tags: ["waiting"],
      checklist: [],
      isPinned: false,
    },
    meeting: {
      title: "Key Call / Meeting Takeaways",
      content: "Summary of key decisions discussed regarding this thread:",
      color: "blue",
      priority: "medium",
      scope: "conversation",
      tags: ["meeting-notes", "decision"],
      checklist: [],
      isPinned: false,
    },
  };

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

  if (
    window.location.search.indexOf("view=all") !== -1 ||
    window.location.pathname.indexOf("dashboard") !== -1
  ) {
    state.viewScope = "all_emails";
  }

  var toastTimer = null;
  var loadSeq = 0;
  var lastPointerDownAt = 0;

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
    return d.toLocaleDateString(undefined, {
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function currentItem() {
    return state.isInOutlook && Office.context && Office.context.mailbox
      ? Office.context.mailbox.item
      : null;
  }

  function sortPinnedFirst(list) {
    return (list || []).slice().sort(function (a, b) {
      if (Boolean(a.isPinned) !== Boolean(b.isPinned)) {
        return a.isPinned ? -1 : 1;
      }
      return String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
    });
  }

  /**
   * Updates ONLY the toast DOM container without re-rendering #root,
   * so active clicks, editor inputs, and focus are NEVER interrupted!
   */
  function updateToastDom() {
    var toastRoot = document.getElementById("toast-root");
    if (!toastRoot) return;
    if (!state.toast) {
      toastRoot.innerHTML = "";
      return;
    }
    var html =
      '<div class="toast-bar" id="toast"><span>' + esc(state.toast) + "</span>";
    if (state.toastUndo && state.lastDeleted) {
      html +=
        '<button type="button" class="undo-toast-btn" data-action="undo-delete" id="btn-undo">Undo</button>';
    }
    html += "</div>";
    toastRoot.innerHTML = html;
  }

  function showToast(msg, withUndo) {
    state.toast = msg;
    state.toastUndo = Boolean(withUndo);
    updateToastDom();
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(
      function () {
        state.toast = null;
        state.toastUndo = false;
        updateToastDom();
      },
      withUndo ? 6000 : 3200
    );
  }

  // ---------------------------------------------------------------------------
  // Email context & storage
  // ---------------------------------------------------------------------------
  function resolveContextAsync(cb) {
    if (!state.isInOutlook) {
      cb(DEMO_EMAILS[state.demoIndex]);
      return;
    }
    var mailbox = Office.context && Office.context.mailbox;
    var item = mailbox ? mailbox.item : null;
    if (!item) {
      cb(null);
      return;
    }
    var up = mailbox.userProfile || {};
    var isCompose =
      item.subject && typeof item.subject.getAsync === "function";

    var finishWithIdAndSubject = function (resolvedItemId, resolvedSubject) {
      var from = (!isCompose && (item.from || item.sender)) || {};
      var senderName =
        from.displayName || from.emailAddress || up.displayName || "Sender";
      var senderEmail = from.emailAddress || up.emailAddress || "";
      var key =
        resolvedItemId || item.conversationId || "outlook-draft-item";
      cb({
        itemId: key,
        conversationId: item.conversationId || key,
        subject: resolvedSubject || "(No Subject)",
        senderName: senderName,
        senderEmail: senderEmail,
        userDisplayName: up.displayName || "Outlook User",
      });
    };

    var getSubject = function (resolvedItemId) {
      if (isCompose) {
        try {
          item.subject.getAsync(function (res) {
            var subj =
              res &&
              res.status === Office.AsyncResultStatus.Succeeded &&
              res.value
                ? res.value
                : "(Draft Message)";
            finishWithIdAndSubject(resolvedItemId, subj);
          });
        } catch (e) {
          finishWithIdAndSubject(resolvedItemId, "(Draft Message)");
        }
      } else {
        finishWithIdAndSubject(resolvedItemId, Core.subjectOf(item));
      }
    };

    if (!item.itemId && isCompose && typeof item.saveAsync === "function") {
      try {
        item.saveAsync(function (res) {
          var savedId =
            res && res.status === Office.AsyncResultStatus.Succeeded
              ? res.value
              : item.conversationId || "compose-draft-item";
          getSubject(savedId);
        });
      } catch (e) {
        getSubject(item.conversationId || "compose-draft-item");
      }
    } else {
      getSubject(Core.getItemKey(item));
    }
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
      {
        id: "seed-note-101",
        itemId: e1.itemId,
        conversationId: e1.conversationId,
        scope: "message",
        emailSubject: e1.subject,
        emailSender: e1.senderName + " <" + e1.senderEmail + ">",
        title: "Negotiate 12% Volume Discount Tier",
        content:
          "Sarah confirmed flexibility on the 3-year renewal if we sign before quarter-end. Verify seat count with IT Ops before replying.",
        color: "yellow",
        priority: "high",
        tags: ["follow-up", "billing", "client"],
        checklist: [
          {
            id: "chk-1",
            text: "Confirm active M365 seat count with IT",
            completed: true,
          },
          {
            id: "chk-2",
            text: "Request revised redline PDF from Legal",
            completed: false,
          },
        ],
        isPinned: true,
        reminderAt: new Date(now + 86400000).toISOString(),
        author: "Outlook User",
        createdAt: iso(35),
        updatedAt: iso(12),
      },
      {
        id: "seed-note-102",
        itemId: e1.itemId,
        conversationId: e1.conversationId,
        scope: "conversation",
        emailSubject: e1.subject,
        emailSender: e1.senderName + " <" + e1.senderEmail + ">",
        title: "PO Approval Reference #PO-88412",
        content: "Finance pre-approved budget code FIN-2026-Q4.",
        color: "green",
        priority: "medium",
        tags: ["decision", "billing"],
        checklist: [],
        isPinned: false,
        author: "Outlook User",
        createdAt: iso(90),
        updatedAt: iso(90),
      },
    ];
    var n2 = [
      {
        id: "seed-note-201",
        itemId: e2.itemId,
        conversationId: e2.conversationId,
        scope: "message",
        emailSubject: e2.subject,
        emailSender: e2.senderName + " <" + e2.senderEmail + ">",
        title: "Verify Conditional Access Evidence",
        content:
          "Auditor requested a screenshot of the MFA enforcement policy.",
        color: "pink",
        priority: "urgent",
        tags: ["action-required", "important"],
        checklist: [
          {
            id: "chk-201",
            text: "Export Conditional Access policy JSON",
            completed: false,
          },
        ],
        isPinned: true,
        author: "Outlook User",
        createdAt: iso(180),
        updatedAt: iso(180),
      },
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

  function loadCurrent(silent) {
    var seq = ++loadSeq;
    resolveContextAsync(function (ctx) {
      if (seq !== loadSeq) return;
      if (!ctx) {
        state.noItem = true;
        state.loading = false;
        state.emailNotes = [];
        state.allNotes = Core.readIndex();
        render();
        return;
      }
      var itemSwitched = !state.ctx || state.ctx.itemId !== ctx.itemId;
      state.noItem = false;
      state.ctx = ctx;

      if (!silent && itemSwitched) {
        state.loading = true;
        render();
      }

      ensureDemoSeed(function () {
        var item = currentItem();
        Core.loadItemNotes(
          item,
          ctx.itemId,
          function (notes) {
            if (seq !== loadSeq) return;
            var sorted = sortPinnedFirst(notes || []);
            var allIdx = Core.readIndex();
            var wasLoading = state.loading;
            state.loading = false;

            var changed =
              wasLoading ||
              itemSwitched ||
              JSON.stringify(sorted) !== JSON.stringify(state.emailNotes) ||
              JSON.stringify(allIdx) !== JSON.stringify(state.allNotes);

            state.emailNotes = sorted;
            state.allNotes = allIdx;

            if (changed && !state.editor) {
              render();
            }
            if (item) {
              Core.syncMarkers(
                item,
                ctx.itemId,
                state.emailNotes,
                null,
                function () {}
              );
            }
          },
          ctx.conversationId
        );
      });
    });
  }

  /**
   * Optimistically updates notes for the active email (or any target email in "All Emails")
   * immediately in memory + UI on the same frame, then persists in the background.
   */
  function persistNotesForTargetEmail(
    targetItemId,
    nextTargetNotes,
    toastMsg,
    withUndo
  ) {
    var isCurrentEmail = state.ctx && targetItemId === state.ctx.itemId;
    var sortedTarget = sortPinnedFirst(nextTargetNotes);

    if (isCurrentEmail) {
      state.emailNotes = sortedTarget;
    }

    // Update global index immediately in memory & localStorage
    var otherIndexNotes = state.allNotes.filter(function (n) {
      return n.itemId !== targetItemId;
    });
    var ownTargetNotes = sortedTarget.filter(function (n) {
      return !n.itemId || n.itemId === targetItemId;
    });
    state.allNotes = ownTargetNotes.concat(otherIndexNotes).sort(function (a, b) {
      return String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
    });

    // Save to local item storage immediately
    Core.writeLocalItemNotes(targetItemId, ownTargetNotes);

    // Render UI immediately (0ms delay!)
    render();
    if (toastMsg) {
      showToast(toastMsg, withUndo);
    }

    // Persist to Office CustomProperties & RoamingSettings in the background
    var item = isCurrentEmail ? currentItem() : null;
    Core.saveItemNotes(item, targetItemId, ownTargetNotes, function () {
      Core.replaceItemInIndex(targetItemId, ownTargetNotes, function () {
        if (isCurrentEmail && item) {
          Core.syncMarkers(
            item,
            targetItemId,
            state.emailNotes,
            null,
            function () {}
          );
        }
      });
    });
  }

  function persistCurrent(nextNotes, toastMsg, withUndo) {
    var key = (state.ctx && state.ctx.itemId) || DEMO_EMAILS[0].itemId;
    persistNotesForTargetEmail(key, nextNotes, toastMsg, withUndo);
  }

  function findNoteAnywhere(id) {
    return (
      state.emailNotes.find(function (n) {
        return n.id === id;
      }) ||
      state.allNotes.find(function (n) {
        return n.id === id;
      })
    );
  }

  function getNotesForTargetItem(targetItemId) {
    if (state.ctx && targetItemId === state.ctx.itemId) {
      return state.emailNotes.slice();
    }
    var local = Core.readLocalItemNotes(targetItemId);
    if (local && local.length > 0) {
      return local;
    }
    return state.allNotes.filter(function (n) {
      return n.itemId === targetItemId;
    });
  }

  // ---------------------------------------------------------------------------
  // Filtering
  // ---------------------------------------------------------------------------
  function sourceNotes() {
    return state.viewScope === "current_email"
      ? state.emailNotes
      : state.allNotes;
  }

  function filteredNotes() {
    var q = state.searchQuery.trim().toLowerCase();
    return sourceNotes()
      .filter(function (n) {
        if (state.onlyPinned && !n.isPinned) return false;
        if (state.selectedColor !== "all" && n.color !== state.selectedColor)
          return false;
        if (
          state.selectedPriority !== "all" &&
          n.priority !== state.selectedPriority
        )
          return false;
        if (
          state.selectedTag !== "all" &&
          (n.tags || []).indexOf(state.selectedTag) === -1
        )
          return false;
        if (!q) return true;
        var hay = [
          n.title,
          n.content,
          n.emailSubject,
          n.emailSender,
          (n.tags || []).join(" "),
          (n.checklist || [])
            .map(function (c) {
              return c.text;
            })
            .join(" "),
        ]
          .join(" ")
          .toLowerCase();
        return hay.indexOf(q) !== -1;
      })
      .sort(function (a, b) {
        if (Boolean(a.isPinned) !== Boolean(b.isPinned))
          return a.isPinned ? -1 : 1;
        if (state.sortBy === "priority_desc")
          return (
            (PRIORITY[b.priority] || PRIORITY.medium).rank -
            (PRIORITY[a.priority] || PRIORITY.medium).rank
          );
        if (state.sortBy === "created_desc")
          return String(b.createdAt).localeCompare(String(a.createdAt));
        if (state.sortBy === "created_asc")
          return String(a.createdAt).localeCompare(String(b.createdAt));
        return String(b.updatedAt).localeCompare(String(a.updatedAt));
      });
  }

  function hasActiveFilters() {
    return (
      Boolean(state.searchQuery) ||
      state.selectedColor !== "all" ||
      state.selectedPriority !== "all" ||
      state.selectedTag !== "all" ||
      state.onlyPinned
    );
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
    if (state.editor && !note && !state.editor.id) {
      // Editor for a new note is already open; keep user's typed content and focus title
      syncEditorInputs();
      var existingTitle = document.getElementById("ed-title");
      var existingCard = document.getElementById("note-editor");
      if (existingCard && existingCard.scrollIntoView) {
        existingCard.scrollIntoView({ block: "nearest" });
      }
      if (existingTitle) existingTitle.focus();
      return;
    }

    state.editor = note
      ? {
          id: note.id,
          targetItemId: note.itemId || (state.ctx && state.ctx.itemId),
          title: note.title || "",
          content: note.content || "",
          color: note.color || "yellow",
          priority: note.priority || "medium",
          scope: note.scope || "message",
          reminderAt: note.reminderAt ? String(note.reminderAt).slice(0, 16) : "",
          isPinned: Boolean(note.isPinned),
          tags: (note.tags || []).slice(),
          checklist: (note.checklist || []).map(function (c) {
            return {
              id: c.id,
              text: c.text,
              completed: Boolean(c.completed),
            };
          }),
          newChk: "",
          newTag: "",
        }
      : {
          id: null,
          targetItemId: (state.ctx && state.ctx.itemId) || DEMO_EMAILS[0].itemId,
          title: "",
          content: "",
          color: "yellow",
          priority: "medium",
          scope: "message",
          reminderAt: "",
          isPinned: false,
          tags: [],
          checklist: [],
          newChk: "",
          newTag: "",
        };

    render();

    var scrollArea = document.querySelector(".notes-scroll-area");
    if (scrollArea) {
      scrollArea.scrollTop = 0;
    }
    var t = document.getElementById("ed-title");
    if (t) {
      try {
        t.focus();
      } catch (e) {}
    }
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
    if ((v = get("ed-scope")) !== null) ed.scope = v;
    if ((v = get("ed-reminder")) !== null) ed.reminderAt = v;
    if ((v = get("ed-new-chk")) !== null) ed.newChk = v;
    if ((v = get("ed-new-tag")) !== null) ed.newTag = v;
  }

  function saveEditor() {
    syncEditorInputs();
    var ed = state.editor;
    if (!ed) return;

    // Auto-add any pending checklist item or custom tag the user typed without clicking "+ Add"
    if (ed.newChk && ed.newChk.trim()) {
      addChecklistItem();
    }
    if (ed.newTag && ed.newTag.trim()) {
      addCustomTag();
    }

    if (!ed.title.trim() && !ed.content.trim() && ed.checklist.length === 0) {
      showToast("Add a title, note text, or checklist item first.");
      var titleEl = document.getElementById("ed-title");
      if (titleEl) titleEl.focus();
      return;
    }

    var nowIso = new Date().toISOString();
    var reminderIso = undefined;
    if (ed.reminderAt && ed.reminderAt.trim()) {
      var parsedReminder = new Date(ed.reminderAt);
      if (!isNaN(parsedReminder.getTime())) {
        reminderIso = parsedReminder.toISOString();
      }
    }

    var fields = {
      title: ed.title.trim(),
      content: ed.content.trim(),
      color: ed.color,
      priority: ed.priority,
      scope: ed.scope || "message",
      reminderAt: reminderIso,
      tags: ed.tags.slice(),
      checklist: ed.checklist.slice(),
      isPinned: ed.isPinned,
      updatedAt: nowIso,
    };

    var targetItemId =
      ed.targetItemId || (state.ctx && state.ctx.itemId) || DEMO_EMAILS[0].itemId;
    var targetList = getNotesForTargetItem(targetItemId);
    var next;

    if (ed.id) {
      next = targetList.map(function (n) {
        return n.id === ed.id ? Object.assign({}, n, fields) : n;
      });
    } else {
      var activeCtx = state.ctx || DEMO_EMAILS[0];
      var created = Object.assign(
        {
          id: Core.newId("note"),
          itemId: targetItemId,
          conversationId: activeCtx.conversationId || targetItemId,
          emailSubject: activeCtx.subject || "(No Subject)",
          emailSender:
            activeCtx.senderName +
            (activeCtx.senderEmail ? " <" + activeCtx.senderEmail + ">" : ""),
          author: activeCtx.userDisplayName || "Outlook User",
          createdAt: nowIso,
          syncStatus: state.isInOutlook ? "synced_office" : "local_only",
        },
        fields
      );
      next = [created].concat(targetList);
    }

    var wasEdit = Boolean(ed.id);
    state.editor = null;
    persistNotesForTargetEmail(
      targetItemId,
      next,
      wasEdit ? "Sticky note updated" : "Sticky note saved to this email"
    );
  }

  function createNoteFromTemplate(tplKey) {
    var tpl = QUICK_TEMPLATES[tplKey];
    if (!tpl) return;
    var activeCtx = state.ctx || DEMO_EMAILS[0];
    var nowIso = new Date().toISOString();
    var created = {
      id: Core.newId("note"),
      itemId: activeCtx.itemId,
      conversationId: activeCtx.conversationId || activeCtx.itemId,
      scope: tpl.scope || "message",
      emailSubject: activeCtx.subject || "(No Subject)",
      emailSender:
        activeCtx.senderName +
        (activeCtx.senderEmail ? " <" + activeCtx.senderEmail + ">" : ""),
      title: tpl.title,
      content: tpl.content,
      color: tpl.color,
      priority: tpl.priority,
      tags: (tpl.tags || []).slice(),
      checklist: (tpl.checklist || []).map(function (c) {
        return {
          id: Core.newId("chk"),
          text: c.text,
          completed: Boolean(c.completed),
        };
      }),
      isPinned: Boolean(tpl.isPinned),
      author: activeCtx.userDisplayName || "Outlook User",
      createdAt: nowIso,
      updatedAt: nowIso,
      syncStatus: state.isInOutlook ? "synced_office" : "local_only",
    };
    persistCurrent(
      [created].concat(state.emailNotes),
      "Template note attached to this email"
    );
  }

  function addChecklistItem() {
    var ed = state.editor;
    if (!ed) return;
    var text = (ed.newChk || "").trim();
    if (!text) return;
    ed.checklist.push({ id: Core.newId("chk"), text: text, completed: false });
    ed.newChk = "";
  }

  function addCustomTag() {
    var ed = state.editor;
    if (!ed) return;
    var tag = (ed.newTag || "")
      .trim()
      .toLowerCase()
      .replace(/^#/, "")
      .replace(/\s+/g, "-");
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
      render();
      Core.syncBanner(item, key, state.emailNotes, null, function () {
        showToast(
          hidden
            ? "Banner removed above this email"
            : "Banner restored above this email"
        );
      });
    });
  }

  function setBannerGlobally(enabled) {
    var item = currentItem();
    Core.setBannerEnabled(enabled, function () {
      render();
      if (state.ctx) {
        Core.syncBanner(item, state.ctx.itemId, state.emailNotes, null, function () {
          showToast(
            enabled
              ? "Banners above emails turned on"
              : "Banners above emails turned off"
          );
        });
      }
    });
  }

  // ---------------------------------------------------------------------------
  // Export / import
  // ---------------------------------------------------------------------------
  function downloadFile(filename, content, mimeType, okMsg) {
    try {
      var blob = new Blob([content], { type: mimeType });
      if (window.navigator && window.navigator.msSaveBlob) {
        window.navigator.msSaveBlob(blob, filename);
      } else {
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(function () {
          URL.revokeObjectURL(url);
        }, 1000);
      }
      showToast(okMsg || "Download complete");
    } catch (e) {
      copyText(content, "Download blocked - copied to clipboard instead");
    }
  }

  function exportJson() {
    var notes = Core.readIndex();
    var payload = JSON.stringify(
      {
        application: "Email Sticky Notes for Outlook",
        version: "2.3",
        exportedAt: new Date().toISOString(),
        noteCount: notes.length,
        notes: notes,
      },
      null,
      2
    );
    var dateStr = new Date().toISOString().slice(0, 10);
    downloadFile(
      "outlook-sticky-notes-" + dateStr + ".json",
      payload,
      "application/json",
      "JSON backup downloaded"
    );
  }

  function exportMarkdown() {
    var notes = Core.readIndex();
    var lines = [
      "# Outlook Email Sticky Notes Export",
      "Exported: " + new Date().toLocaleString() + " (" + notes.length + " notes)",
      "",
    ];
    notes.forEach(function (n) {
      lines.push(
        "## " +
          (n.isPinned ? "[PINNED] " : "") +
          (n.title || "Untitled Sticky Note")
      );
      lines.push("- **Email Subject:** " + (n.emailSubject || "(No Subject)"));
      lines.push("- **From:** " + (n.emailSender || "Sender"));
      lines.push(
        "- **Priority:** " +
          String(n.priority || "medium").toUpperCase() +
          " | **Color:** " +
          (n.color || "yellow")
      );
      lines.push(
        "- **Created:** " +
          new Date(n.createdAt).toLocaleString() +
          " | **Updated:** " +
          new Date(n.updatedAt).toLocaleString()
      );
      if (n.tags && n.tags.length > 0) {
        lines.push(
          "- **Tags:** " +
            n.tags
              .map(function (t) {
                return "#" + t;
              })
              .join(" ")
        );
      }
      lines.push("");
      if (n.content) {
        lines.push(n.content);
        lines.push("");
      }
      if (n.checklist && n.checklist.length > 0) {
        lines.push("### Checklist");
        n.checklist.forEach(function (item) {
          lines.push("- [" + (item.completed ? "x" : " ") + "] " + item.text);
        });
        lines.push("");
      }
      lines.push("---");
      lines.push("");
    });
    var dateStr = new Date().toISOString().slice(0, 10);
    downloadFile(
      "outlook-sticky-notes-" + dateStr + ".md",
      lines.join("\n"),
      "text/markdown",
      "Markdown export downloaded"
    );
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
        showToast("That file isn't a valid JSON backup.");
        return;
      }
      var incoming = Array.isArray(parsed)
        ? parsed
        : parsed && Array.isArray(parsed.notes)
        ? parsed.notes
        : [];
      incoming = incoming.filter(function (n) {
        return n && n.id && n.itemId;
      });
      if (incoming.length === 0) {
        showToast("No notes found in that file.");
        return;
      }
      var byId = {};
      Core.readIndex()
        .concat(incoming)
        .forEach(function (n) {
          if (
            !byId[n.id] ||
            String(n.updatedAt) >= String(byId[n.id].updatedAt)
          ) {
            byId[n.id] = n;
          }
        });
      var merged = Object.keys(byId).map(function (k) {
        return byId[k];
      });

      // Also update per-item local storage caches for imported items
      var byItem = {};
      merged.forEach(function (n) {
        if (!byItem[n.itemId]) byItem[n.itemId] = [];
        byItem[n.itemId].push(n);
      });
      Object.keys(byItem).forEach(function (itemKey) {
        Core.writeLocalItemNotes(itemKey, byItem[itemKey]);
      });

      Core.writeIndex(merged, function () {
        var currentKey = state.ctx && state.ctx.itemId;
        var forThis = currentKey
          ? merged.filter(function (n) {
              return (
                n.itemId === currentKey ||
                (n.scope === "conversation" &&
                  n.conversationId === state.ctx.conversationId)
              );
            })
          : [];
        state.settingsOpen = false;
        state.allNotes = Core.readIndex();
        if (forThis.length > 0 && currentKey) {
          persistCurrent(forThis, "Imported " + incoming.length + " note(s)");
        } else {
          render();
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
    return (
      '<option value="' +
      value +
      '"' +
      (selected === value ? " selected" : "") +
      ">" +
      label +
      "</option>"
    );
  }

  function renderHeader(h) {
    h.push('<header class="app-header">');

    // Top row: Brand + clean utility icon buttons
    h.push('<div class="header-top-row">');
    h.push('<div class="brand-group">');
    h.push(
      '<img src="./assets/icon-64.png" alt="" class="brand-icon-img" width="34" height="34" />'
    );
    h.push(
      '<div><h1 class="brand-title">Email Sticky Notes</h1><div class="storage-pill-row">'
    );
    h.push(
      state.isInOutlook
        ? '<span class="status-pill status-office"><span class="status-dot"></span>Saved in Mailbox</span>'
        : '<span class="status-pill status-preview"><span class="status-dot"></span>Interactive Preview</span>'
    );
    h.push("</div></div></div>");

    h.push('<div class="header-utility-btns">');
    h.push(
      '<button type="button" class="icon-btn" data-action="refresh" id="btn-refresh" title="Refresh notes">' +
        ICONS.refresh +
        "</button>"
    );
    h.push(
      '<button type="button" class="icon-btn" data-action="toggle-theme" id="btn-toggle-theme" title="' +
        (state.isDark ? "Switch to Light Mode" : "Switch to Dark Mode") +
        '">' +
        (state.isDark ? ICONS.sun : ICONS.moon) +
        "</button>"
    );
    h.push(
      '<button type="button" class="icon-btn" data-action="open-settings" id="btn-open-settings" title="Settings & Backup">' +
        ICONS.settings +
        "</button>"
    );
    h.push("</div></div>");

    // Primary Action Row: Two spacious, easy-to-read buttons
    if (!state.noItem) {
      h.push('<div class="header-cta-row">');
      h.push(
        '<button type="button" class="btn-cta-new" data-action="new-note" id="btn-new-note" title="Create a new sticky note (Ctrl+N)">' +
          ICONS.plus +
          "<span>New Sticky Note</span></button>"
      );
      h.push(
        '<button type="button" class="btn-cta-quick" data-action="quick-note" id="btn-quick-note" title="Instantly stamp a follow-up sticky note with 1 click">' +
          ICONS.bolt +
          "<span>Quick Note</span></button>"
      );
      h.push("</div>");
    }

    // Segmented Scope Pill Tabs
    h.push('<div class="scope-segmented-bar" role="tablist">');
    h.push(
      '<button type="button" class="scope-tab ' +
        (state.viewScope === "current_email" ? "active" : "") +
        '" data-action="scope" data-value="current_email"><span>This Email</span><span class="tab-badge">' +
        state.emailNotes.length +
        "</span></button>"
    );
    h.push(
      '<button type="button" class="scope-tab ' +
        (state.viewScope === "all_emails" ? "active" : "") +
        '" data-action="scope" data-value="all_emails"><span>All Emails</span><span class="tab-badge">' +
        state.allNotes.length +
        "</span></button>"
    );
    h.push("</div></header>");
  }

  function renderEmailContext(h) {
    h.push('<section class="email-context-banner">');
    if (!state.isInOutlook) {
      h.push(
        '<div class="demo-switcher-bar"><span class="demo-label">Sample Email:</span><div class="demo-pills">'
      );
      DEMO_EMAILS.forEach(function (em, idx) {
        var count = state.allNotes.filter(function (n) {
          return (
            n.itemId === em.itemId ||
            (n.scope === "conversation" &&
              n.conversationId === em.conversationId)
          );
        }).length;
        h.push(
          '<button type="button" class="demo-pill ' +
            (idx === state.demoIndex ? "active" : "") +
            '" data-action="demo" data-value="' +
            idx +
            '">Email #' +
            (idx + 1) +
            (count ? " (" + count + ")" : "") +
            "</button>"
        );
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
    h.push(
      '<div class="email-subject" title="' +
        esc(c.subject) +
        '">' +
        esc(c.subject) +
        "</div>"
    );
    if (hasNotes) {
      h.push(
        '<div class="email-category-markers"><span class="outlook-category-pill cat-yellow" title="Category badge shown in your Outlook Inbox">' +
          Core.NOTE_CATEGORY +
          "</span>"
      );
      if (hasUrgent)
        h.push(
          '<span class="outlook-category-pill cat-red">' +
            Core.URGENT_CATEGORY +
            "</span>"
        );
      h.push("</div>");
    }
    h.push("</div>");
    if (c.senderName || c.senderEmail) {
      h.push(
        '<div class="email-sender">' +
          esc(c.senderName) +
          (c.senderEmail ? " &lt;" + esc(c.senderEmail) + "&gt;" : "") +
          "</div>"
      );
    }
    h.push("</div></div>");

    // Clean, pill-button Banner Control Bar
    h.push('<div class="banner-control-row">');
    if (!Core.isBannerEnabled()) {
      h.push(
        '<div class="banner-status-label">' +
          ICONS.bellOff +
          "<span>Email banner is off</span></div>"
      );
      h.push('<div class="banner-actions-group">');
      h.push(
        '<button type="button" class="banner-pill-btn" data-action="banner-clear-now" id="btn-banner-clear-now">Clear Now</button>'
      );
      h.push(
        '<button type="button" class="banner-pill-btn primary" data-action="banner-global-on" id="btn-banner-global-on">Turn On</button>'
      );
      h.push("</div>");
    } else if (Core.isBannerDismissed(c.itemId)) {
      h.push(
        '<div class="banner-status-label">' +
          ICONS.bellOff +
          "<span>Banner removed on this email</span></div>"
      );
      h.push('<div class="banner-actions-group">');
      h.push(
        '<button type="button" class="banner-pill-btn primary" data-action="banner-show" id="btn-banner-show">Restore Banner</button>'
      );
      h.push(
        '<button type="button" class="banner-pill-btn" data-action="banner-global-off" id="btn-banner-global-off">Turn Off Always</button>'
      );
      h.push("</div>");
    } else {
      h.push(
        '<div class="banner-status-label">' +
          ICONS.bell +
          "<span>Notification above email</span></div>"
      );
      h.push('<div class="banner-actions-group">');
      h.push(
        '<button type="button" class="banner-pill-btn danger" data-action="banner-hide" id="btn-banner-hide">' +
          ICONS.close +
          "<span>Remove Banner</span></button>"
      );
      h.push(
        '<button type="button" class="banner-pill-btn" data-action="banner-global-off" id="btn-banner-global-off">Turn Off Always</button>'
      );
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
    h.push(
      '<input type="text" id="input-search" class="search-input" placeholder="Search notes, tags, checklists, or subjects..." value="' +
        esc(state.searchQuery) +
        '" />'
    );
    h.push(
      '<button type="button" id="btn-clear-search" class="clear-search-btn" data-action="clear-search" title="Clear search" style="' +
        (state.searchQuery ? "" : "display:none") +
        '">' +
        ICONS.close +
        "</button>"
    );
    h.push("</div>");
    h.push(
      '<button type="button" class="filter-drawer-toggle ' +
        (showDrawer ? "active" : "") +
        '" data-action="toggle-filter-drawer" id="btn-toggle-filters" title="Show or hide color, priority, and tag filters">' +
        ICONS.sliders +
        "<span>Filters" +
        (activeCount > 0 ? " (" + activeCount + ")" : "") +
        "</span></button>"
    );
    h.push("</div>");

    if (showDrawer) {
      h.push('<div class="filter-drawer-body">');
      h.push('<div class="filter-controls-row"><div class="color-filter-dots">');
      h.push(
        '<button type="button" class="color-dot-all ' +
          (state.selectedColor === "all" ? "selected" : "") +
          '" data-action="filter-color" data-value="all">All Colors</button>'
      );
      Object.keys(PALETTE).forEach(function (k) {
        h.push(
          '<button type="button" class="color-dot ' +
            (state.selectedColor === k ? "selected" : "") +
            '" style="background-color:' +
            PALETTE[k].swatch +
            '" data-action="filter-color" data-value="' +
            k +
            '" title="' +
            PALETTE[k].label +
            '"></button>'
        );
      });
      h.push('</div><div class="filter-select-group">');
      h.push(
        '<button type="button" class="pin-filter-btn ' +
          (state.onlyPinned ? "active" : "") +
          '" data-action="toggle-only-pinned" id="btn-only-pinned" title="Show only pinned notes">' +
          ICONS.pin +
          "<span>Pinned</span></button>"
      );
      h.push(
        '<select id="select-priority-filter" class="compact-select" title="Priority">' +
          opt("all", "Any Priority", state.selectedPriority) +
          opt("urgent", "Urgent", state.selectedPriority) +
          opt("high", "High", state.selectedPriority) +
          opt("medium", "Medium", state.selectedPriority) +
          opt("low", "Low", state.selectedPriority) +
          "</select>"
      );
      h.push(
        '<select id="select-sort" class="compact-select" title="Sort">' +
          opt("updated_desc", "Latest Updated", state.sortBy) +
          opt("created_desc", "Newest Created", state.sortBy) +
          opt("created_asc", "Oldest Created", state.sortBy) +
          opt("priority_desc", "Highest Priority", state.sortBy) +
          "</select>"
      );
      h.push("</div></div>");

      var tagMap = {};
      sourceNotes().forEach(function (n) {
        (n.tags || []).forEach(function (t) {
          tagMap[t] = true;
        });
      });
      var tags = Object.keys(tagMap).sort();
      if (tags.length > 0 || hasActiveFilters()) {
        h.push('<div class="tag-filter-strip">');
        if (tags.length > 0) {
          h.push(
            '<button type="button" class="tag-chip-filter ' +
              (state.selectedTag === "all" ? "active" : "") +
              '" data-action="filter-tag" data-value="all">All Tags</button>'
          );
          tags.forEach(function (t) {
            h.push(
              '<button type="button" class="tag-chip-filter ' +
                (state.selectedTag === t ? "active" : "") +
                '" data-action="filter-tag" data-value="' +
                esc(t) +
                '">#' +
                esc(t) +
                "</button>"
            );
          });
        }
        if (hasActiveFilters()) {
          h.push(
            '<button type="button" class="reset-filters-link" data-action="reset-filters">Reset All</button>'
          );
        }
        h.push("</div>");
      }
      h.push("</div>");
    }
    h.push("</div>");
  }

  function renderEditor(h) {
    var ed = state.editor;
    h.push('<div class="note-editor-card" id="note-editor">');
    h.push(
      '<div class="editor-header"><span class="editor-heading">' +
        (ed.id ? "Edit Sticky Note" : "Create Sticky Note") +
        '</span><div class="editor-header-controls">'
    );
    h.push(
      '<button type="button" class="pill-action-btn ' +
        (ed.isPinned ? "active" : "") +
        '" data-action="ed-toggle-pin" id="ed-pin" title="Pin to top of email">' +
        ICONS.pin +
        "<span>" +
        (ed.isPinned ? "Pinned" : "Pin") +
        "</span></button>"
    );
    h.push(
      '<button type="button" class="icon-toggle-btn" data-action="ed-cancel" title="Close editor">' +
        ICONS.close +
        "</button></div></div>"
    );

    h.push(
      '<input type="text" id="ed-title" class="editor-title-input" placeholder="Note title (e.g. Follow up with Legal)..." value="' +
        esc(ed.title) +
        '" />'
    );
    h.push(
      '<textarea id="ed-content" class="editor-body-textarea" rows="4" placeholder="Write your sticky note details here... (Ctrl+Enter to save)">' +
        esc(ed.content) +
        "</textarea>"
    );

    h.push(
      '<div class="editor-options-grid"><div class="editor-field"><label>Sticky Color</label><div class="editor-color-swatches">'
    );
    Object.keys(PALETTE).forEach(function (k) {
      h.push(
        '<button type="button" class="swatch-btn ' +
          (ed.color === k ? "selected" : "") +
          '" style="background-color:' +
          PALETTE[k].swatch +
          '" data-action="ed-color" data-value="' +
          k +
          '" title="' +
          PALETTE[k].label +
          '"></button>'
      );
    });
    h.push("</div></div>");
    h.push(
      '<div class="editor-field"><label>Priority Level</label><select id="ed-priority" class="compact-select">' +
        opt("low", "Low Priority", ed.priority) +
        opt("medium", "Medium Priority", ed.priority) +
        opt("high", "High Priority", ed.priority) +
        opt("urgent", "Urgent Priority", ed.priority) +
        "</select></div>"
    );
    h.push(
      '<div class="editor-field"><label>Attach Scope</label><select id="ed-scope" class="compact-select">' +
        opt("message", "This Email Only", ed.scope || "message") +
        opt("conversation", "Entire Thread", ed.scope || "message") +
        "</select></div>"
    );
    h.push(
      '<div class="editor-field"><label>Follow-up Reminder</label><input type="datetime-local" id="ed-reminder" class="compact-text-input" value="' +
        esc(ed.reminderAt || "") +
        '" /></div></div>'
    );

    h.push(
      '<div class="editor-section"><div class="editor-section-label">Action Checklist</div>'
    );
    if (ed.checklist.length > 0) {
      h.push('<ul class="editor-checklist">');
      ed.checklist.forEach(function (c) {
        h.push(
          '<li class="editor-checklist-item"><input type="checkbox" data-action="ed-chk-toggle" data-value="' +
            esc(c.id) +
            '"' +
            (c.completed ? " checked" : "") +
            " /><span>" +
            esc(c.text) +
            '</span><button type="button" class="remove-chk-btn" data-action="ed-chk-remove" data-value="' +
            esc(c.id) +
            '" title="Remove item">' +
            ICONS.close +
            "</button></li>"
        );
      });
      h.push("</ul>");
    }
    h.push(
      '<div class="inline-add-row"><input type="text" id="ed-new-chk" class="compact-text-input" placeholder="Add a checklist item and press Enter..." value="' +
        esc(ed.newChk) +
        '" /><button type="button" class="btn-secondary-sm" data-action="ed-chk-add" id="ed-chk-add">+ Add Item</button></div></div>'
    );

    h.push(
      '<div class="editor-section"><div class="editor-section-label">Tags</div><div class="preset-tags-wrap">'
    );
    var allTags = PRESET_TAGS.concat(
      ed.tags.filter(function (t) {
        return PRESET_TAGS.indexOf(t) === -1;
      })
    );
    allTags.forEach(function (t) {
      var isCustom = PRESET_TAGS.indexOf(t) === -1;
      var isSelected = ed.tags.indexOf(t) !== -1;
      h.push(
        '<button type="button" class="tag-pill-selectable ' +
          (isSelected ? "selected" : "") +
          '" data-action="ed-tag" data-value="' +
          esc(t) +
          '">#' +
          esc(t) +
          (isCustom && isSelected ? " &times;" : "") +
          "</button>"
      );
    });
    h.push(
      '</div><div class="inline-add-row"><input type="text" id="ed-new-tag" class="compact-text-input" placeholder="Add custom tag..." value="' +
        esc(ed.newTag) +
        '" /><button type="button" class="btn-secondary-sm" data-action="ed-tag-add" id="ed-tag-add">+ Add Tag</button></div></div>'
    );

    h.push(
      '<div class="editor-footer"><button type="button" class="btn-ghost" data-action="ed-cancel">Cancel</button><button type="button" class="btn-primary" data-action="ed-save" id="btn-save-editor">' +
        (ed.id ? "Save Changes" : "Save Sticky Note") +
        "</button></div></div>"
    );
  }

  function renderCard(h, note) {
    var p = PALETTE[note.color] || PALETTE.yellow;
    var pr = PRIORITY[note.priority] || PRIORITY.medium;
    var dark = state.isDark;
    var id = esc(note.id);
    var wasEdited =
      note.updatedAt &&
      note.createdAt &&
      Math.abs(
        new Date(note.updatedAt).getTime() - new Date(note.createdAt).getTime()
      ) > 3000;

    h.push(
      '<article class="sticky-note-card" data-note-id="' +
        id +
        '" style="background-color:' +
        (dark ? p.dBg : p.lBg) +
        ";border-color:" +
        (dark ? p.dBd : p.lBd) +
        ";color:" +
        (dark ? p.dTx : p.lTx) +
        '">'
    );
    h.push(
      '<div class="sticky-note-top-strip" style="background-color:' +
        (dark ? p.dHd : p.lHd) +
        '"><div class="sticky-badges-left">'
    );
    if (note.isPinned) {
      h.push('<span class="pinned-indicator">' + ICONS.pin + " Pinned</span>");
    }
    h.push(
      '<span class="priority-badge" style="background-color:' +
        pr.color +
        '">' +
        pr.label +
        "</span>"
    );
    if (note.scope === "conversation") {
      h.push(
        '<span class="scope-badge" title="Attached to entire email conversation thread">' +
          ICONS.thread +
          " Thread</span>"
      );
    }
    h.push("</div>");

    h.push('<div class="sticky-card-actions">');
    h.push(
      '<button type="button" class="card-action-pill" data-action="copy-note" data-value="' +
        id +
        '" title="Copy note text">' +
        ICONS.copy +
        "<span>Copy</span></button>"
    );
    h.push(
      '<button type="button" class="card-action-pill ' +
        (note.isPinned ? "active" : "") +
        '" data-action="pin-note" data-value="' +
        id +
        '" title="' +
        (note.isPinned ? "Unpin note" : "Pin to top") +
        '">' +
        ICONS.pin +
        "</button>"
    );
    h.push(
      '<button type="button" class="card-action-pill" data-action="edit-note" data-value="' +
        id +
        '" title="Edit note">' +
        ICONS.edit +
        "<span>Edit</span></button>"
    );
    h.push(
      '<button type="button" class="card-action-pill delete-btn" data-action="delete-note" data-value="' +
        id +
        '" title="Delete note">' +
        ICONS.trash +
        "</button>"
    );
    h.push("</div></div>");

    h.push('<div class="sticky-note-body">');
    if (state.viewScope === "all_emails") {
      h.push(
        '<div class="cross-email-ref" title="' +
          esc(note.emailSender) +
          '">' +
          ICONS.mail +
          ' <span class="cross-email-subject">' +
          esc(note.emailSubject) +
          "</span></div>"
      );
    }
    if (note.title) {
      h.push('<h3 class="sticky-note-title">' + esc(note.title) + "</h3>");
    }
    if (note.content) {
      h.push('<p class="sticky-note-content">' + esc(note.content) + "</p>");
    }
    if (note.checklist && note.checklist.length > 0) {
      h.push('<div class="sticky-checklist-box">');
      note.checklist.forEach(function (c) {
        h.push(
          '<label class="sticky-checklist-row ' +
            (c.completed ? "completed" : "") +
            '"><input type="checkbox" data-action="chk-toggle" data-note="' +
            id +
            '" data-value="' +
            esc(c.id) +
            '"' +
            (c.completed ? " checked" : "") +
            " /><span>" +
            esc(c.text) +
            "</span></label>"
        );
      });
      h.push("</div>");
    }
    if (note.reminderAt) {
      var rDate = new Date(note.reminderAt);
      if (!isNaN(rDate.getTime())) {
        h.push(
          '<div class="sticky-reminder-pill">' +
            ICONS.bell +
            "<span>Reminder: " +
            esc(rDate.toLocaleString()) +
            "</span></div>"
        );
      }
    }
    if (note.tags && note.tags.length > 0) {
      h.push('<div class="sticky-tags-row">');
      note.tags.forEach(function (t) {
        h.push(
          '<button type="button" class="sticky-tag-chip" data-action="filter-tag" data-value="' +
            esc(t) +
            '">#' +
            esc(t) +
            "</button>"
        );
      });
      h.push("</div>");
    }
    h.push(
      '<div class="sticky-note-footer"><span class="timestamp-meta" title="Created: ' +
        esc(new Date(note.createdAt).toLocaleString()) +
        " | Updated: " +
        esc(new Date(note.updatedAt).toLocaleString()) +
        '">' +
        ICONS.clock +
        " <span>" +
        esc(formatTimestamp(note.updatedAt)) +
        (wasEdited ? " (edited)" : "") +
        "</span></span>"
    );
    h.push(
      '<span class="author-meta">' + esc(note.author) + "</span></div></div></article>"
    );
  }

  function renderNotesListRegion(h) {
    var list = filteredNotes();
    if (state.loading) {
      h.push(
        '<div class="loading-skeleton-stack"><div class="skeleton-card"></div></div>'
      );
    } else if (list.length === 0) {
      h.push('<div class="empty-state-box" id="empty-state">');
      if (hasActiveFilters()) {
        h.push(
          '<h3>No matching sticky notes</h3><p>Try clearing your search or active filter pills.</p><button type="button" class="btn-secondary-sm" data-action="reset-filters">Clear All Filters</button>'
        );
      } else if (state.viewScope === "all_emails") {
        h.push(
          "<h3>No annotated emails yet</h3><p>Sticky notes you attach to any email will appear here.</p>"
        );
      } else {
        h.push(
          '<h3>No sticky notes on this email</h3><p>Attach a private, color-coded note, checklist, or follow-up reminder to this message.</p><button type="button" class="btn-primary" data-action="new-note">' +
            ICONS.plus +
            "<span>Create Sticky Note</span></button>"
        );
        h.push('<div class="quick-templates-section">');
        h.push(
          '<span class="quick-templates-title">' +
            ICONS.sparkles +
            "<span>1-Click Quick Templates</span></span>"
        );
        h.push('<div class="quick-template-buttons">');
        h.push(
          '<button type="button" class="template-chip" data-action="quick-template" data-value="followup">+ Follow-up Checklist</button>'
        );
        h.push(
          '<button type="button" class="template-chip" data-action="quick-template" data-value="approval">+ Waiting on Approval</button>'
        );
        h.push(
          '<button type="button" class="template-chip" data-action="quick-template" data-value="meeting">+ Meeting Summary</button>'
        );
        h.push("</div></div>");
      }
      h.push("</div>");
    } else {
      h.push('<div class="sticky-notes-list">');
      list.forEach(function (n) {
        renderCard(h, n);
      });
      h.push("</div>");
    }
  }

  /**
   * Updates only #notes-list-region when filtering via search input so #input-search
   * never loses focus or cursor selection.
   */
  function updateNotesListRegionDom() {
    var region = document.getElementById("notes-list-region");
    if (!region) {
      render();
      return;
    }
    var h = [];
    renderNotesListRegion(h);
    region.innerHTML = h.join("");
  }

  function renderSettings(h) {
    var bannerOn = Core.isBannerEnabled();
    h.push(
      '<div class="modal-backdrop" data-action="close-settings-backdrop"><div class="settings-modal" id="settings-modal">'
    );
    h.push(
      '<div class="modal-header"><h2>Settings &amp; Backup</h2><button type="button" class="icon-btn" data-action="close-settings" title="Close">' +
        ICONS.close +
        '</button></div><div class="modal-body">'
    );

    h.push(
      '<div class="settings-group"><label class="settings-group-title">Notification Banner Above Emails</label>'
    );
    h.push(
      '<label class="radio-card ' +
        (bannerOn ? "selected" : "") +
        '"><input type="checkbox" id="cfg-banner-enabled"' +
        (bannerOn ? " checked" : "") +
        ' /><div><strong>Show summary banner at the top of annotated emails</strong><p>Uncheck this to stop the "Sticky Note" notification bar from appearing above your emails. The yellow category badge in your Inbox will still mark emails that have notes.</p></div></label></div>'
    );

    h.push(
      '<div class="settings-group"><label class="settings-group-title">Backup, Export &amp; Restore (' +
        state.allNotes.length +
        ' notes)</label><div class="backup-actions-row">'
    );
    h.push(
      '<button type="button" class="btn-secondary-sm" data-action="export-json" id="btn-export-json">' +
        ICONS.download +
        "<span>Export JSON</span></button>"
    );
    h.push(
      '<button type="button" class="btn-secondary-sm" data-action="export-markdown" id="btn-export-markdown">' +
        ICONS.fileText +
        "<span>Export Markdown</span></button>"
    );
    h.push(
      '<button type="button" class="btn-secondary-sm" data-action="copy-json">' +
        ICONS.copy +
        "<span>Copy Backup</span></button>"
    );
    h.push(
      '<label class="btn-secondary-sm file-upload-label">' +
        ICONS.upload +
        '<span>Import JSON</span><input type="file" id="cfg-import-file" accept=".json,application/json" hidden /></label>'
    );
    h.push("</div></div>");

    h.push(
      '<div class="settings-group"><label class="settings-group-title">Privacy &amp; Storage</label><p class="settings-hint">Your sticky notes are stored privately inside your own Microsoft 365 / Exchange mailbox on each email and never leave your account.</p></div>'
    );

    h.push(
      '<div class="modal-footer"><button type="button" class="btn-primary" data-action="close-settings">Done</button></div></div></div></div>'
    );
  }

  function render() {
    document.documentElement.setAttribute(
      "data-theme",
      state.isDark ? "dark" : "light"
    );
    var root = document.getElementById("root");
    if (!root) return;

    var h = [];
    h.push('<div class="addin-shell">');
    renderHeader(h);

    if (state.noItem && state.viewScope !== "all_emails") {
      h.push(
        '<main class="notes-scroll-area"><div class="empty-state-box" id="no-item-state"><h3>Select an email</h3><p>Pick any message in your Inbox to view or attach sticky notes, or switch to the <b>All Emails</b> tab above to view and search all notes.</p></div></main>'
      );
    } else {
      if (!state.noItem) renderEmailContext(h);
      renderFilters(h);
      h.push('<main class="notes-scroll-area">');
      if (state.editor) renderEditor(h);
      h.push('<div id="notes-list-region">');
      renderNotesListRegion(h);
      h.push("</div>");
      h.push("</main>");
    }

    h.push('<div id="toast-root"></div>');
    if (state.settingsOpen) renderSettings(h);
    h.push("</div>");

    root.innerHTML = h.join("");
    updateToastDom();
  }

  // ---------------------------------------------------------------------------
  // Events
  // ---------------------------------------------------------------------------
  function onClick(e) {
    var el = e.target;
    while (
      el &&
      el !== document &&
      !(el.getAttribute && el.getAttribute("data-action"))
    ) {
      el = el.parentNode;
    }
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
        persistCurrent(
          [qNote].concat(state.emailNotes),
          "Quick Note stamped on this email"
        );
        break;
      case "quick-template":
        createNoteFromTemplate(value);
        break;
      case "refresh":
        loadCurrent(false);
        showToast("Notes refreshed");
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
        loadCurrent(false);
        break;
      case "clear-search":
        state.searchQuery = "";
        render();
        var sInput = document.getElementById("input-search");
        if (sInput) sInput.focus();
        break;
      case "toggle-filter-drawer":
        state.filtersExpanded = !state.filtersExpanded;
        render();
        break;
      case "filter-color":
        state.selectedColor =
          state.selectedColor === value && value !== "all" ? "all" : value;
        render();
        break;
      case "filter-tag":
        state.selectedTag =
          state.selectedTag === value && value !== "all" ? "all" : value;
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
        var tagInp = document.getElementById("ed-new-tag");
        if (tagInp) tagInp.focus();
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
        var pn = findNoteAnywhere(value);
        if (!pn) return;
        var pinTargetId =
          pn.itemId || (state.ctx && state.ctx.itemId) || DEMO_EMAILS[0].itemId;
        var pinTargetList = getNotesForTargetItem(pinTargetId);
        persistNotesForTargetEmail(
          pinTargetId,
          pinTargetList.map(function (n) {
            return n.id === value
              ? Object.assign({}, n, {
                  isPinned: !n.isPinned,
                  updatedAt: new Date().toISOString(),
                })
              : n;
          }),
          pn.isPinned ? "Note unpinned" : "Note pinned"
        );
        break;
      case "edit-note":
        var en = findNoteAnywhere(value);
        if (en) {
          openEditor(en);
          var edEl = document.getElementById("note-editor");
          if (edEl && edEl.scrollIntoView) edEl.scrollIntoView();
        }
        break;
      case "delete-note":
        var dn = findNoteAnywhere(value);
        if (!dn) return;
        var delTargetId =
          dn.itemId || (state.ctx && state.ctx.itemId) || DEMO_EMAILS[0].itemId;
        state.lastDeleted = { note: dn, itemId: delTargetId };
        if (state.editor && state.editor.id === value) state.editor = null;
        var delTargetList = getNotesForTargetItem(delTargetId);
        persistNotesForTargetEmail(
          delTargetId,
          delTargetList.filter(function (n) {
            return n.id !== value;
          }),
          "Sticky note deleted",
          true
        );
        break;
      case "undo-delete":
        var ld = state.lastDeleted;
        state.lastDeleted = null;
        if (ld && ld.note) {
          var undoTargetId =
            ld.itemId ||
            ld.note.itemId ||
            (state.ctx && state.ctx.itemId) ||
            DEMO_EMAILS[0].itemId;
          var existingForUndo = getNotesForTargetItem(undoTargetId).filter(
            function (n) {
              return n.id !== ld.note.id;
            }
          );
          persistNotesForTargetEmail(
            undoTargetId,
            [ld.note].concat(existingForUndo),
            "Note restored"
          );
        }
        break;
      case "copy-note":
        var cn = findNoteAnywhere(value);
        if (cn) {
          var text = [cn.title, cn.content]
            .concat(
              (cn.checklist || []).map(function (c) {
                return (c.completed ? "[x] " : "[ ] ") + c.text;
              })
            )
            .filter(Boolean)
            .join("\n");
          copyText(text, "Note copied to clipboard");
        }
        break;
      case "export-json":
        exportJson();
        break;
      case "export-markdown":
        exportMarkdown();
        break;
      case "copy-json":
        copyText(
          JSON.stringify({ notes: Core.readIndex() }, null, 2),
          "Backup JSON copied"
        );
        break;
    }
  }

  function onChange(e) {
    var el = e.target;
    var action = el.getAttribute && el.getAttribute("data-action");
    if (action === "chk-toggle") {
      var noteId = el.getAttribute("data-note");
      var chkId = el.getAttribute("data-value");
      var targetNote = findNoteAnywhere(noteId);
      if (!targetNote) return;
      var chkTargetId =
        targetNote.itemId ||
        (state.ctx && state.ctx.itemId) ||
        DEMO_EMAILS[0].itemId;
      var chkTargetList = getNotesForTargetItem(chkTargetId);
      persistNotesForTargetEmail(
        chkTargetId,
        chkTargetList.map(function (n) {
          if (n.id !== noteId) return n;
          return Object.assign({}, n, {
            checklist: (n.checklist || []).map(function (c) {
              return c.id === chkId
                ? Object.assign({}, c, { completed: !c.completed })
                : c;
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
    if (
      state.editor &&
      (el.id === "ed-priority" || el.id === "ed-scope" || el.id === "ed-reminder")
    ) {
      syncEditorInputs();
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
    if (!el) return;

    // Live-sync editor fields without re-rendering DOM so typing is 100% smooth
    if (
      state.editor &&
      (el.id === "ed-title" ||
        el.id === "ed-content" ||
        el.id === "ed-new-chk" ||
        el.id === "ed-new-tag" ||
        el.id === "ed-reminder")
    ) {
      syncEditorInputs();
      return;
    }

    if (el.id === "input-search") {
      state.searchQuery = el.value;
      var clearBtn = document.getElementById("btn-clear-search");
      if (clearBtn) {
        clearBtn.style.display = state.searchQuery ? "" : "none";
      }
      updateNotesListRegionDom();
    }
  }

  function onKeyDown(e) {
    var el = e.target;

    // Global shortcut: Ctrl+N or Alt+N opens New Sticky Note composer
    if ((e.ctrlKey || e.altKey) && (e.key === "n" || e.key === "N")) {
      e.preventDefault();
      openEditor(null);
      return;
    }

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
      if (e.key === "Enter" && el && el.id === "ed-new-chk") {
        e.preventDefault();
        syncEditorInputs();
        addChecklistItem();
        render();
        var inp = document.getElementById("ed-new-chk");
        if (inp) inp.focus();
        return;
      }
      if (e.key === "Enter" && el && el.id === "ed-new-tag") {
        e.preventDefault();
        syncEditorInputs();
        addCustomTag();
        render();
        var tagInp = document.getElementById("ed-new-tag");
        if (tagInp) tagInp.focus();
        return;
      }
    } else if (e.key === "Escape" && state.settingsOpen) {
      state.settingsOpen = false;
      render();
    }
  }

  var eventsBound = false;
  function bindEvents() {
    if (eventsBound) return;
    eventsBound = true;
    var root = document.getElementById("root");
    root.addEventListener("mousedown", function () {
      lastPointerDownAt = Date.now();
    });
    root.addEventListener("touchstart", function () {
      lastPointerDownAt = Date.now();
    });
    root.addEventListener("click", onClick);
    root.addEventListener("change", onChange);
    root.addEventListener("input", onInput);
    window.addEventListener("keydown", onKeyDown);

    // Cross-frame sync when "Quick Note" ribbon command is executed while taskpane is open
    if (Core.onCrossFrameUpdate) {
      Core.onCrossFrameUpdate(function () {
        if (state.editor || Date.now() - lastPointerDownAt < 600) return;
        loadCurrent(true);
      });
    }
  }

  var itemChangedRegistered = false;
  function registerOutlookItemChanged() {
    if (itemChangedRegistered) return;
    if (
      state.isInOutlook &&
      Office.context &&
      Office.context.mailbox &&
      Office.context.mailbox.addHandlerAsync
    ) {
      itemChangedRegistered = true;
      try {
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.ItemChanged,
          function () {
            state.editor = null;
            state.lastDeleted = null;
            loadCurrent(false);
          }
        );
      } catch (e) {}
    }
  }

  function start(isOutlook) {
    state.isInOutlook = isOutlook;
    bindEvents();
    registerOutlookItemChanged();
    loadCurrent(false);
  }

  window.__stickyNotesApp = {
    state: state,
    reload: function () {
      loadCurrent(false);
    },
  };

  var started = false;
  if (window.Office && typeof Office.onReady === "function") {
    var fallbackTimer = setTimeout(function () {
      if (!started) {
        started = true;
        start(false);
      }
    }, 1000);

    Office.onReady(function (info) {
      clearTimeout(fallbackTimer);
      var inOutlook = Boolean(
        info &&
          info.host === Office.HostType.Outlook &&
          Office.context &&
          Office.context.mailbox
      );
      if (!started) {
        started = true;
        start(inOutlook);
      } else if (inOutlook && !state.isInOutlook) {
        // Upgraded from fallback preview to live Outlook host once Office.onReady completed
        state.isInOutlook = true;
        registerOutlookItemChanged();
        loadCurrent(false);
      }
    });
  } else {
    started = true;
    start(false);
  }
})();
