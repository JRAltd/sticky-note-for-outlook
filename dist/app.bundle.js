/**
 * Outlook Email Sticky Notes Add-in (v2.0.0)
 * Pre-compiled Standalone Production Bundle
 * Supports:
 *  - Office.js CustomProperties + RoamingSettings cross-email index
 *  - Microsoft Graph Open Extensions (/me/messages/{id}/extensions/com.jraltd.emailNotes)
 *  - Pinned Taskpane ItemChanged auto-switching
 *  - Multiple timestamped notes, 6 sticky colors, checklists, tags, search, dark mode, JSON/MD export
 */
(function () {
  "use strict";

  var CUSTOM_PROP_NOTES_KEY = "jraltd_email_sticky_notes_v2";
  var ROAMING_INDEX_KEY = "jraltd_email_notes_global_index_v2";
  var LOCAL_STORAGE_NOTES_PREFIX = "outlook_sticky_notes_item_";
  var LOCAL_STORAGE_INDEX_KEY = "outlook_sticky_notes_all_index";
  var LOCAL_STORAGE_AUTH_CONFIG_KEY = "outlook_sticky_notes_auth_config";
  var LOCAL_STORAGE_THEME_KEY = "outlook_sticky_notes_theme";
  var DEMO_SEED_FLAG_KEY = "outlook_sticky_notes_demo_seeded_v2";
  var OUTLOOK_NOTE_CATEGORY_NAME = "📝 Sticky Note";
  var OUTLOOK_URGENT_CATEGORY_NAME = "🔥 Urgent Note";

  var STICKY_COLOR_PALETTE = {
    yellow: {
      id: "yellow",
      label: "Classic Yellow",
      swatchHex: "#FFE255",
      lightBg: "#FFF9C4",
      lightHeader: "#FFF176",
      lightBorder: "#FBC02D",
      lightText: "#3E2723",
      darkBg: "#3D3514",
      darkHeader: "#4F4418",
      darkBorder: "#A68B22",
      darkText: "#FFF59D",
    },
    amber: {
      id: "amber",
      label: "Warm Amber",
      swatchHex: "#FFB74D",
      lightBg: "#FFE0B2",
      lightHeader: "#FFCC80",
      lightBorder: "#FB8C00",
      lightText: "#3E2723",
      darkBg: "#3E2914",
      darkHeader: "#523519",
      darkBorder: "#B26A1B",
      darkText: "#FFE0B2",
    },
    green: {
      id: "green",
      label: "Mint Green",
      swatchHex: "#81C784",
      lightBg: "#DCEDC8",
      lightHeader: "#C5E1A5",
      lightBorder: "#7CB342",
      lightText: "#1B5E20",
      darkBg: "#183622",
      darkHeader: "#20472D",
      darkBorder: "#3E8E53",
      darkText: "#C8E6C9",
    },
    blue: {
      id: "blue",
      label: "Sky Blue",
      swatchHex: "#64B5F6",
      lightBg: "#E1F5FE",
      lightHeader: "#B3E5FC",
      lightBorder: "#039BE5",
      lightText: "#0D47A1",
      darkBg: "#142B3E",
      darkHeader: "#1B3B54",
      darkBorder: "#2E75B6",
      darkText: "#BBDEFB",
    },
    pink: {
      id: "pink",
      label: "Rose Pink",
      swatchHex: "#F06292",
      lightBg: "#FCE4EC",
      lightHeader: "#F8BBD0",
      lightBorder: "#D81B60",
      lightText: "#4A148C",
      darkBg: "#3B1828",
      darkHeader: "#4F2036",
      darkBorder: "#A83868",
      darkText: "#F8BBD0",
    },
    purple: {
      id: "purple",
      label: "Lavender",
      swatchHex: "#BA68C8",
      lightBg: "#F3E5F5",
      lightHeader: "#E1BEE7",
      lightBorder: "#8E24AA",
      lightText: "#311B92",
      darkBg: "#2B193D",
      darkHeader: "#3A2252",
      darkBorder: "#7B46A8",
      darkText: "#E1BEE7",
    },
  };

  var PRIORITY_CONFIG = {
    low: { label: "Low", rank: 1, badgeColor: "#64748b" },
    medium: { label: "Medium", rank: 2, badgeColor: "#0284c7" },
    high: { label: "High", rank: 3, badgeColor: "#d97706" },
    urgent: { label: "Urgent", rank: 4, badgeColor: "#dc2626" },
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

  var DEMO_SAMPLE_EMAILS = [
    {
      itemId: "demo-msg-001-q4-contract",
      restId: "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OA001",
      conversationId: "conv-q4-enterprise-renewal",
      subject: "Q4 Enterprise License Renewal & Pricing Proposal",
      senderName: "Sarah Jenkins (VP Procurement)",
      senderEmail: "s.jenkins@contoso-enterprise.com",
      receivedDateTime: new Date(Date.now() - 1000 * 60 * 42).toISOString(),
      itemType: "message",
      mode: "read",
      hostEnvironment: "browser_preview",
      userEmail: "you@yourdomain.com",
      userDisplayName: "Outlook User",
    },
    {
      itemId: "demo-msg-002-security-audit",
      restId: "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OA002",
      conversationId: "conv-soc2-security-review",
      subject: "Action Required: SOC2 Type II Architecture Sign-off",
      senderName: "Marcus Vance (Security Engineering)",
      senderEmail: "mvance@cloudsec-ops.io",
      receivedDateTime: new Date(Date.now() - 1000 * 60 * 60 * 5).toISOString(),
      itemType: "message",
      mode: "read",
      hostEnvironment: "browser_preview",
      userEmail: "you@yourdomain.com",
      userDisplayName: "Outlook User",
    },
    {
      itemId: "demo-msg-003-board-deck",
      restId: "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OA003",
      conversationId: "conv-board-deck-prep",
      subject: "Draft Board Deck Slides & Financial Appendix",
      senderName: "Elena Rostova (CFO)",
      senderEmail: "elena@northwind-ventures.com",
      receivedDateTime: new Date(Date.now() - 1000 * 60 * 60 * 26).toISOString(),
      itemType: "message",
      mode: "read",
      hostEnvironment: "browser_preview",
      userEmail: "you@yourdomain.com",
      userDisplayName: "Outlook User",
    },
  ];

  // Application State
  var state = {
    isInOutlook: false,
    demoEmailIndex: 0,
    emailContext: DEMO_SAMPLE_EMAILS[0],
    emailNotes: [],
    allEmailsNotes: [],
    isDark: localStorage.getItem(LOCAL_STORAGE_THEME_KEY) === "dark",
    viewScope: "current_email",
    searchQuery: "",
    selectedColor: "all",
    selectedPriority: "all",
    selectedTag: "all",
    sortBy: "updated_desc",
    onlyPinned: false,
    isCreating: false,
    editingNote: null,
    isSettingsOpen: false,
    graphSynced: false,
    toastMessage: null,
    lastDeletedNote: null,
    editorState: null,
  };

  function escapeHtml(str) {
    return String(str || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function formatTimestamp(isoString) {
    try {
      var date = new Date(isoString);
      var diffMins = Math.floor((Date.now() - date.getTime()) / 60000);
      if (diffMins < 1) return "Just now";
      if (diffMins < 60) return diffMins + "m ago";
      var diffHours = Math.floor(diffMins / 60);
      if (diffHours < 24) return diffHours + "h ago";
      return date.toLocaleDateString(undefined, {
        month: "short",
        day: "numeric",
        hour: "2-digit",
        minute: "2-digit",
      });
    } catch (e) {
      return isoString;
    }
  }

  function getAuthConfig() {
    try {
      var raw = localStorage.getItem(LOCAL_STORAGE_AUTH_CONFIG_KEY);
      if (raw) return JSON.parse(raw);
    } catch (e) {}
    return {
      clientId: "",
      tenantId: "common",
      graphExtensionName: "com.jraltd.emailNotes",
      storageMode: "hybrid",
    };
  }

  function saveAuthConfig(cfg) {
    localStorage.setItem(LOCAL_STORAGE_AUTH_CONFIG_KEY, JSON.stringify(cfg));
  }

  function showToast(msg) {
    state.toastMessage = msg;
    render();
    setTimeout(function () {
      if (state.toastMessage === msg) {
        state.toastMessage = null;
        render();
      }
    }, 3500);
  }

  function getAllIndexedNotes() {
    var map = {};
    try {
      if (state.isInOutlook && window.Office && Office.context && Office.context.roamingSettings) {
        var rawR = Office.context.roamingSettings.get(ROAMING_INDEX_KEY);
        if (rawR) {
          var arrR = typeof rawR === "string" ? JSON.parse(rawR) : rawR;
          if (Array.isArray(arrR)) {
            arrR.forEach(function (n) {
              if (n && n.id) map[n.id] = n;
            });
          }
        }
      }
    } catch (e) {}

    try {
      var rawL = localStorage.getItem(LOCAL_STORAGE_INDEX_KEY);
      if (rawL) {
        var arrL = JSON.parse(rawL);
        if (Array.isArray(arrL)) {
          arrL.forEach(function (n) {
            if (!map[n.id] || n.updatedAt >= map[n.id].updatedAt) {
              map[n.id] = n;
            }
          });
        }
      }
    } catch (e) {}

    return Object.keys(map)
      .map(function (k) {
        return map[k];
      })
      .sort(function (a, b) {
        return b.updatedAt.localeCompare(a.updatedAt);
      });
  }

  function saveItemNotesToStorage(itemId, notes, cb) {
    try {
      localStorage.setItem(
        LOCAL_STORAGE_NOTES_PREFIX + itemId,
        JSON.stringify(notes)
      );
    } catch (e) {}

    var others = getAllIndexedNotes().filter(function (n) {
      return n.itemId !== itemId;
    });
    var allCombined = notes.concat(others).sort(function (a, b) {
      return b.updatedAt.localeCompare(a.updatedAt);
    });

    try {
      localStorage.setItem(LOCAL_STORAGE_INDEX_KEY, JSON.stringify(allCombined));
    } catch (e) {}

    if (
      state.isInOutlook &&
      window.Office &&
      Office.context &&
      Office.context.roamingSettings
    ) {
      try {
        Office.context.roamingSettings.set(
          ROAMING_INDEX_KEY,
          JSON.stringify(allCombined.slice(0, 150))
        );
        Office.context.roamingSettings.saveAsync(function () {});
      } catch (e) {}
    }

    if (
      state.isInOutlook &&
      window.Office &&
      Office.context &&
      Office.context.mailbox &&
      Office.context.mailbox.item &&
      Office.context.mailbox.item.loadCustomPropertiesAsync
    ) {
      Office.context.mailbox.item.loadCustomPropertiesAsync(function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          var cp = res.value;
          cp.set(CUSTOM_PROP_NOTES_KEY, JSON.stringify(notes));
          cp.saveAsync(function () {
            if (cb) cb();
          });
        } else if (cb) {
          cb();
        }
      });
    } else if (cb) {
      cb();
    }
  }

  function loadItemNotesFromStorage(itemId, cb) {
    var localArr = [];
    try {
      var raw = localStorage.getItem(LOCAL_STORAGE_NOTES_PREFIX + itemId);
      if (raw) localArr = JSON.parse(raw) || [];
    } catch (e) {}

    if (
      !state.isInOutlook ||
      !window.Office ||
      !Office.context ||
      !Office.context.mailbox ||
      !Office.context.mailbox.item ||
      !Office.context.mailbox.item.loadCustomPropertiesAsync
    ) {
      cb(localArr);
      return;
    }

    Office.context.mailbox.item.loadCustomPropertiesAsync(function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        var rawCp = res.value.get(CUSTOM_PROP_NOTES_KEY);
        if (rawCp) {
          try {
            var parsed = JSON.parse(rawCp);
            if (Array.isArray(parsed)) {
              cb(parsed);
              return;
            }
          } catch (e) {}
        }
      }
      cb(localArr);
    });
  }

  function ensureDemoSeed() {
    if (state.isInOutlook) return;
    if (localStorage.getItem(DEMO_SEED_FLAG_KEY) === "true") return;

    var e1 = DEMO_SAMPLE_EMAILS[0];
    var e2 = DEMO_SAMPLE_EMAILS[1];
    var now = Date.now();

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
          "Sarah confirmed flexibility on the 3-year enterprise renewal if we sign before quarter-end. Verify seat count with IT Ops before replying.",
        color: "yellow",
        priority: "high",
        tags: ["follow-up", "billing", "client"],
        checklist: [
          {
            id: "chk-1",
            text: "Confirm active M365 E5 seat count with IT",
            completed: true,
          },
          {
            id: "chk-2",
            text: "Request revised redline PDF from Legal",
            completed: false,
          },
          {
            id: "chk-3",
            text: "Send counter-proposal before Friday 3 PM",
            completed: false,
          },
        ],
        isPinned: true,
        author: e1.userDisplayName,
        createdAt: new Date(now - 1000 * 60 * 35).toISOString(),
        updatedAt: new Date(now - 1000 * 60 * 12).toISOString(),
        syncStatus: "synced_office",
      },
      {
        id: "seed-note-102",
        itemId: e1.itemId,
        conversationId: e1.conversationId,
        scope: "conversation",
        emailSubject: e1.subject,
        emailSender: e1.senderName + " <" + e1.senderEmail + ">",
        title: "PO Approval Reference #PO-88412",
        content:
          "Finance pre-approved budget code FIN-2026-Q4. Attach PO number when final MSA is countersigned.",
        color: "green",
        priority: "medium",
        tags: ["decision", "billing"],
        checklist: [],
        isPinned: false,
        author: e1.userDisplayName,
        createdAt: new Date(now - 1000 * 60 * 90).toISOString(),
        updatedAt: new Date(now - 1000 * 60 * 90).toISOString(),
        syncStatus: "synced_office",
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
        title: "Verify Azure AD Conditional Access Evidence",
        content:
          "Auditor requested screenshot of MFA enforcement policy and token lifetime settings for the Graph Open Extensions app registration.",
        color: "pink",
        priority: "urgent",
        tags: ["action-required", "important"],
        checklist: [
          {
            id: "chk-201",
            text: "Export Entra ID Conditional Access JSON policy",
            completed: false,
          },
        ],
        isPinned: true,
        author: e2.userDisplayName,
        createdAt: new Date(now - 1000 * 60 * 180).toISOString(),
        updatedAt: new Date(now - 1000 * 60 * 180).toISOString(),
        syncStatus: "synced_office",
      },
    ];

    saveItemNotesToStorage(e1.itemId, n1);
    saveItemNotesToStorage(e2.itemId, n2);
    localStorage.setItem(DEMO_SEED_FLAG_KEY, "true");
  }

  function syncOutlookEmailMarkers(notes) {
    if (
      !state.isInOutlook ||
      !window.Office ||
      !Office.context ||
      !Office.context.mailbox ||
      !Office.context.mailbox.item
    ) {
      return;
    }

    var item = Office.context.mailbox.item;
    var count = (notes || []).length;
    var hasUrgent = (notes || []).some(function (n) {
      return n.priority === "urgent";
    });
    var topNote =
      (notes || []).find(function (n) {
        return n.isPinned;
      }) || (notes || [])[0];

    // 1. Reading Pane Top Notification Banner
    try {
      if (item.notificationMessages) {
        if (count > 0 && topNote) {
          var preview =
            topNote.title || (topNote.content || "").slice(0, 50) || "Note attached";
          var msg =
            count === 1
              ? '📝 1 Sticky Note: "' + preview + '"'
              : "📝 " + count + ' Sticky Notes — Top: "' + preview + '"';
          item.notificationMessages.replaceAsync("EmailStickyNotesStatus", {
            type: Office.MailboxEnums.ItemNotificationMessageType
              .InformationalMessage,
            message: msg.slice(0, 150),
            icon: "Icon.16x16",
            persistent: false,
          });
        } else {
          item.notificationMessages.removeAsync("EmailStickyNotesStatus");
        }
      }
    } catch (e) {}

    // 2. Ensure Master Categories exist ("📝 Sticky Note" in Yellow, "🔥 Urgent Note" in Red)
    try {
      var mc = Office.context.mailbox.masterCategories;
      if (mc && typeof mc.getAsync === "function") {
        mc.getAsync(function (mcRes) {
          if (mcRes.status === Office.AsyncResultStatus.Succeeded) {
            var existingNames = (mcRes.value || []).map(function (c) {
              return c.displayName;
            });
            var toCreate = [];
            if (existingNames.indexOf(OUTLOOK_NOTE_CATEGORY_NAME) === -1) {
              toCreate.push({
                displayName: OUTLOOK_NOTE_CATEGORY_NAME,
                color: Office.MailboxEnums.CategoryColor.Preset3, // Yellow
              });
            }
            if (existingNames.indexOf(OUTLOOK_URGENT_CATEGORY_NAME) === -1) {
              toCreate.push({
                displayName: OUTLOOK_URGENT_CATEGORY_NAME,
                color: Office.MailboxEnums.CategoryColor.Preset0, // Red
              });
            }
            if (toCreate.length > 0) {
              mc.addAsync(toCreate, function () {});
            }
          }
        });
      }
    } catch (e) {}

    // 3. Stamp or remove Category badge on the Email item itself (shows in Inbox list!)
    try {
      if (item.categories && typeof item.categories.getAsync === "function") {
        item.categories.getAsync(function (catRes) {
          if (catRes.status !== Office.AsyncResultStatus.Succeeded) return;
          var currentCats = catRes.value || [];
          var hasSticky =
            currentCats.indexOf(OUTLOOK_NOTE_CATEGORY_NAME) !== -1;
          var hasUrg =
            currentCats.indexOf(OUTLOOK_URGENT_CATEGORY_NAME) !== -1;

          var toAdd = [];
          var toRemove = [];

          if (count > 0 && !hasSticky) toAdd.push(OUTLOOK_NOTE_CATEGORY_NAME);
          if (count === 0 && hasSticky)
            toRemove.push(OUTLOOK_NOTE_CATEGORY_NAME);

          if (count > 0 && hasUrgent && !hasUrg)
            toAdd.push(OUTLOOK_URGENT_CATEGORY_NAME);
          if ((count === 0 || !hasUrgent) && hasUrg)
            toRemove.push(OUTLOOK_URGENT_CATEGORY_NAME);

          var finishRemove = function () {
            if (toRemove.length > 0) {
              item.categories.removeAsync(toRemove, function () {});
            }
          };

          if (toAdd.length > 0) {
            item.categories.addAsync(toAdd, function () {
              finishRemove();
            });
          } else {
            finishRemove();
          }
        });
      }
    } catch (e) {}
  }

  function refreshCurrentEmailNotes() {
    ensureDemoSeed();
    loadItemNotesFromStorage(state.emailContext.itemId, function (loaded) {
      state.emailNotes = loaded;
      state.allEmailsNotes = getAllIndexedNotes();
      syncOutlookEmailMarkers(loaded);
      render();
    });
  }

  function readOutlookContext(cb) {
    if (
      !state.isInOutlook ||
      !window.Office ||
      !Office.context ||
      !Office.context.mailbox ||
      !Office.context.mailbox.item
    ) {
      cb(DEMO_SAMPLE_EMAILS[state.demoEmailIndex]);
      return;
    }

    var mailbox = Office.context.mailbox;
    var item = mailbox.item;
    var up = mailbox.userProfile || {};
    var rawId = item.itemId || item.conversationId || "outlook-item";
    var restId = rawId;
    try {
      if (mailbox.convertToRestId && item.itemId) {
        restId = mailbox.convertToRestId(
          item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
      }
    } catch (e) {}

    var fromObj = item.from || item.sender || {};
    cb({
      itemId: rawId,
      restId: restId,
      conversationId: item.conversationId || rawId,
      subject: item.subject || "(No Subject)",
      senderName: fromObj.displayName || up.displayName || "Sender",
      senderEmail: fromObj.emailAddress || up.emailAddress || "",
      receivedDateTime: new Date().toISOString(),
      itemType: "message",
      mode: "read",
      hostEnvironment: "outlook_desktop",
      userEmail: up.emailAddress || "user@outlook.com",
      userDisplayName: up.displayName || "Outlook User",
    });
  }

  function initEditorState(note) {
    if (note) {
      state.editorState = {
        id: note.id,
        title: note.title || "",
        content: note.content || "",
        color: note.color || "yellow",
        priority: note.priority || "medium",
        scope: note.scope || "message",
        isPinned: Boolean(note.isPinned),
        tags: (note.tags || []).slice(),
        checklist: (note.checklist || []).map(function (c) {
          return { id: c.id, text: c.text, completed: c.completed };
        }),
      };
    } else {
      state.editorState = {
        id: null,
        title: "",
        content: "",
        color: "yellow",
        priority: "medium",
        scope: "message",
        isPinned: false,
        tags: ["follow-up"],
        checklist: [],
      };
    }
  }

  function getFilteredNotes() {
    var src =
      state.viewScope === "current_email"
        ? state.emailNotes
        : state.allEmailsNotes;
    var q = (state.searchQuery || "").trim().toLowerCase();

    return src
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
        if (Boolean(a.isPinned) !== Boolean(b.isPinned)) {
          return a.isPinned ? -1 : 1;
        }
        if (state.sortBy === "priority_desc") {
          return (
            PRIORITY_CONFIG[b.priority].rank - PRIORITY_CONFIG[a.priority].rank
          );
        }
        if (state.sortBy === "created_asc") {
          return a.createdAt.localeCompare(b.createdAt);
        }
        return b.updatedAt.localeCompare(a.updatedAt);
      });
  }

  function render() {
    document.documentElement.setAttribute(
      "data-theme",
      state.isDark ? "dark" : "light"
    );
    var root = document.getElementById("root");
    if (!root) return;

    var filtered = getFilteredNotes();
    var allTagsMap = {};
    (state.viewScope === "current_email"
      ? state.emailNotes
      : state.allEmailsNotes
    ).forEach(function (n) {
      (n.tags || []).forEach(function (t) {
        allTagsMap[t] = true;
      });
    });
    var availableTags = Object.keys(allTagsMap).sort();

    var html = [];
    html.push('<div class="addin-shell">');

    // Header
    html.push('<header class="app-header">');
    html.push('<div class="header-top-row">');
    html.push('<div class="brand-group">');
    html.push(
      '<img src="./assets/icon-64.png" alt="Email Sticky Notes" class="brand-icon-img" width="32" height="32" />'
    );
    html.push("<div>");
    html.push('<h1 class="brand-title">Email Sticky Notes</h1>');
    html.push('<div class="storage-pill-row">');
    html.push(
      '<span class="status-pill ' +
        (state.isInOutlook ? "status-office" : "status-preview") +
        '">' +
        (state.isInOutlook
          ? "Office.js + Graph Ready"
          : "Interactive Preview Mode") +
        "</span>"
    );
    html.push("</div></div></div>");

    html.push('<div class="header-actions">');
    html.push(
      '<button type="button" class="btn-primary-compact" id="btn-new-note">+ New Note</button>'
    );
    html.push(
      '<button type="button" class="icon-btn" id="btn-toggle-theme" title="Toggle Dark/Light Theme">' +
        (state.isDark ? "☀️" : "🌙") +
        "</button>"
    );
    html.push(
      '<button type="button" class="icon-btn" id="btn-open-settings" title="Settings & Backup">⚙️</button>'
    );
    html.push("</div></div>");

    html.push('<div class="scope-tabs">');
    html.push(
      '<button type="button" class="scope-tab ' +
        (state.viewScope === "current_email" ? "active" : "") +
        '" data-scope="current_email"><span>This Email</span><span class="tab-badge">' +
        state.emailNotes.length +
        "</span></button>"
    );
    html.push(
      '<button type="button" class="scope-tab ' +
        (state.viewScope === "all_emails" ? "active" : "") +
        '" data-scope="all_emails"><span>All Annotated Emails</span><span class="tab-badge">' +
        state.allEmailsNotes.length +
        "</span></button>"
    );
    html.push("</div></header>");

    // Email Context Banner
    html.push('<section class="email-context-banner">');
    if (!state.isInOutlook) {
      html.push('<div class="demo-switcher-bar">');
      html.push('<span class="demo-label">Sandbox Email Switcher:</span>');
      html.push('<div class="demo-pills">');
      DEMO_SAMPLE_EMAILS.forEach(function (em, idx) {
        var emCount = state.allEmailsNotes.filter(function (n) {
          return (
            n.itemId === em.itemId ||
            (n.scope === "conversation" &&
              n.conversationId === em.conversationId)
          );
        }).length;
        html.push(
          '<button type="button" class="demo-pill ' +
            (idx === state.demoEmailIndex ? "active" : "") +
            '" data-demo-idx="' +
            idx +
            '">Email #' +
            (idx + 1) +
            (emCount > 0 ? " (📝 " + emCount + ")" : "") +
            "</button>"
        );
      });
      html.push("</div></div>");
    }
    var hasNotesOnCurrent = state.emailNotes.length > 0;
    var hasUrgentOnCurrent = state.emailNotes.some(function (n) {
      return n.priority === "urgent";
    });
    html.push('<div class="email-meta-row">');
    html.push('<div class="email-meta-text">');
    html.push('<div class="email-subject-row">');
    html.push(
      '<div class="email-subject">' +
        escapeHtml(state.emailContext.subject) +
        "</div>"
    );
    if (hasNotesOnCurrent) {
      html.push('<div class="email-category-markers">');
      html.push(
        '<span class="outlook-category-pill cat-yellow" title="Outlook Inbox Category Tag automatically applied to this email">' +
          OUTLOOK_NOTE_CATEGORY_NAME +
          "</span>"
      );
      if (hasUrgentOnCurrent) {
        html.push(
          '<span class="outlook-category-pill cat-red" title="Urgent Outlook Inbox Category Tag automatically applied">' +
            OUTLOOK_URGENT_CATEGORY_NAME +
            "</span>"
        );
      }
      html.push("</div>");
    }
    html.push("</div>");
    html.push(
      '<div class="email-sender">' +
        escapeHtml(
          state.emailContext.senderName +
            " (" +
            state.emailContext.senderEmail +
            ")"
        ) +
        "</div>"
    );
    html.push("</div></div></section>");

    // Search & Filters
    html.push('<div class="search-filter-container">');
    html.push('<div class="search-input-wrapper">');
    html.push(
      '<input type="text" id="input-search" class="search-input" style="padding-left:10px;" placeholder="Search sticky notes, tags, checklists, or subjects..." value="' +
        escapeHtml(state.searchQuery) +
        '" />'
    );
    html.push("</div>");

    html.push('<div class="filter-controls-row">');
    html.push('<div class="color-filter-dots">');
    html.push(
      '<button type="button" class="color-dot-all ' +
        (state.selectedColor === "all" ? "selected" : "") +
        '" data-filter-color="all">All</button>'
    );
    Object.keys(STICKY_COLOR_PALETTE).forEach(function (cKey) {
      var cTheme = STICKY_COLOR_PALETTE[cKey];
      html.push(
        '<button type="button" class="color-dot ' +
          (state.selectedColor === cKey ? "selected" : "") +
          '" style="background-color:' +
          cTheme.swatchHex +
          '" data-filter-color="' +
          cKey +
          '" title="' +
          cTheme.label +
          '"></button>'
      );
    });
    html.push("</div>");

    html.push('<div class="filter-select-group">');
    html.push(
      '<select id="select-priority-filter" class="compact-select">' +
        '<option value="all"' +
        (state.selectedPriority === "all" ? " selected" : "") +
        ">Any Priority</option>" +
        '<option value="urgent"' +
        (state.selectedPriority === "urgent" ? " selected" : "") +
        ">Urgent</option>" +
        '<option value="high"' +
        (state.selectedPriority === "high" ? " selected" : "") +
        ">High</option>" +
        '<option value="medium"' +
        (state.selectedPriority === "medium" ? " selected" : "") +
        ">Medium</option>" +
        '<option value="low"' +
        (state.selectedPriority === "low" ? " selected" : "") +
        ">Low</option></select>"
    );
    html.push("</div></div>");

    if (availableTags.length > 0) {
      html.push('<div class="tag-filter-strip">');
      html.push(
        '<button type="button" class="tag-chip-filter ' +
          (state.selectedTag === "all" ? "active" : "") +
          '" data-filter-tag="all">All Tags</button>'
      );
      availableTags.forEach(function (t) {
        html.push(
          '<button type="button" class="tag-chip-filter ' +
            (state.selectedTag === t ? "active" : "") +
            '" data-filter-tag="' +
            escapeHtml(t) +
            '">#' +
            escapeHtml(t) +
            "</button>"
        );
      });
      html.push("</div>");
    }
    html.push("</div>");

    // Main Notes Area
    html.push('<main class="notes-scroll-area">');

    if (state.editorState) {
      var ed = state.editorState;
      html.push('<div class="note-editor-card">');
      html.push('<div class="editor-header">');
      html.push(
        '<span class="editor-heading">' +
          (ed.id ? "Edit Sticky Note" : "New Sticky Note") +
          "</span>"
      );
      html.push(
        '<button type="button" class="icon-toggle-btn" id="btn-cancel-editor">✕</button>'
      );
      html.push("</div>");

      html.push(
        '<input type="text" id="ed-title" class="editor-title-input" placeholder="Sticky note title..." value="' +
          escapeHtml(ed.title) +
          '" />'
      );
      html.push(
        '<textarea id="ed-content" class="editor-body-textarea" rows="4" placeholder="Write your timestamped sticky note here...">' +
          escapeHtml(ed.content) +
          "</textarea>"
      );

      html.push('<div class="editor-options-grid">');
      html.push('<div class="editor-field"><label>Sticky Color</label><div class="editor-color-swatches">');
      Object.keys(STICKY_COLOR_PALETTE).forEach(function (cKey) {
        var cTheme = STICKY_COLOR_PALETTE[cKey];
        html.push(
          '<button type="button" class="swatch-btn ' +
            (ed.color === cKey ? "selected" : "") +
            '" style="background-color:' +
            cTheme.swatchHex +
            '" data-ed-color="' +
            cKey +
            '"></button>'
        );
      });
      html.push("</div></div>");

      html.push(
        '<div class="editor-field"><label>Priority</label><select id="ed-priority" class="compact-select">' +
          '<option value="low"' +
          (ed.priority === "low" ? " selected" : "") +
          ">Low</option>" +
          '<option value="medium"' +
          (ed.priority === "medium" ? " selected" : "") +
          ">Medium</option>" +
          '<option value="high"' +
          (ed.priority === "high" ? " selected" : "") +
          ">High</option>" +
          '<option value="urgent"' +
          (ed.priority === "urgent" ? " selected" : "") +
          ">Urgent</option></select></div>"
      );
      html.push("</div>");

      html.push('<div class="editor-section"><div class="editor-section-label">Tags</div><div class="preset-tags-wrap">');
      PRESET_TAGS.forEach(function (tag) {
        var isSel = ed.tags.indexOf(tag) !== -1;
        html.push(
          '<button type="button" class="tag-pill-selectable ' +
            (isSel ? "selected" : "") +
            '" data-ed-tag="' +
            tag +
            '">#' +
            tag +
            "</button>"
        );
      });
      html.push("</div></div>");

      html.push('<div class="editor-footer">');
      html.push(
        '<button type="button" class="btn-ghost" id="btn-cancel-editor-2">Cancel</button>'
      );
      html.push(
        '<button type="button" class="btn-primary" id="btn-save-editor">Save Sticky Note</button>'
      );
      html.push("</div></div>");
    }

    if (filtered.length === 0) {
      html.push('<div class="empty-state-box">');
      html.push("<h3>No sticky notes found</h3>");
      html.push(
        "<p>Attach color-coded, timestamped sticky notes directly to this email.</p>"
      );
      html.push(
        '<button type="button" class="btn-primary" id="btn-empty-new">+ Create Sticky Note</button>'
      );
      html.push("</div>");
    } else {
      html.push('<div class="sticky-notes-list">');
      filtered.forEach(function (note) {
        var cTheme =
          STICKY_COLOR_PALETTE[note.color] || STICKY_COLOR_PALETTE.yellow;
        var pInfo = PRIORITY_CONFIG[note.priority] || PRIORITY_CONFIG.medium;
        var bg = state.isDark ? cTheme.darkBg : cTheme.lightBg;
        var hdr = state.isDark ? cTheme.darkHeader : cTheme.lightHeader;
        var bdr = state.isDark ? cTheme.darkBorder : cTheme.lightBorder;
        var txt = state.isDark ? cTheme.darkText : cTheme.lightText;

        html.push(
          '<article class="sticky-note-card" style="background-color:' +
            bg +
            ";border-color:" +
            bdr +
            ";color:" +
            txt +
            '">'
        );
        html.push(
          '<div class="sticky-note-top-strip" style="background-color:' +
            hdr +
            '">'
        );
        html.push('<div class="sticky-badges-left">');
        if (note.isPinned) {
          html.push('<span class="pinned-indicator">📌 Pinned</span>');
        }
        html.push(
          '<span class="priority-badge" style="background-color:' +
            pInfo.badgeColor +
            '">' +
            pInfo.label +
            "</span>"
        );
        html.push("</div>");

        html.push('<div class="sticky-card-actions">');
        html.push(
          '<button type="button" class="card-icon-btn" data-pin-id="' +
            note.id +
            '" title="Pin/Unpin">📌</button>'
        );
        html.push(
          '<button type="button" class="card-icon-btn" data-edit-id="' +
            note.id +
            '" title="Edit">✏️</button>'
        );
        html.push(
          '<button type="button" class="card-icon-btn delete-btn" data-delete-id="' +
            note.id +
            '" title="Delete">🗑️</button>'
        );
        html.push("</div></div>");

        html.push('<div class="sticky-note-body">');
        if (state.viewScope === "all_emails") {
          html.push(
            '<div class="cross-email-ref">✉️ <span class="cross-email-subject">' +
              escapeHtml(note.emailSubject) +
              "</span></div>"
          );
        }
        if (note.title) {
          html.push(
            '<h3 class="sticky-note-title">' + escapeHtml(note.title) + "</h3>"
          );
        }
        if (note.content) {
          html.push(
            '<p class="sticky-note-content">' +
              escapeHtml(note.content) +
              "</p>"
          );
        }
        if (note.checklist && note.checklist.length > 0) {
          html.push('<div class="sticky-checklist-box">');
          note.checklist.forEach(function (chk) {
            html.push(
              '<label class="sticky-checklist-row ' +
                (chk.completed ? "completed" : "") +
                '"><input type="checkbox" data-chk-note="' +
                note.id +
                '" data-chk-id="' +
                chk.id +
                '" ' +
                (chk.completed ? "checked" : "") +
                " /><span>" +
                escapeHtml(chk.text) +
                "</span></label>"
            );
          });
          html.push("</div>");
        }
        if (note.tags && note.tags.length > 0) {
          html.push('<div class="sticky-tags-row">');
          note.tags.forEach(function (t) {
            html.push(
              '<button type="button" class="sticky-tag-chip" data-filter-tag="' +
                escapeHtml(t) +
                '">#' +
                escapeHtml(t) +
                "</button>"
            );
          });
          html.push("</div>");
        }
        html.push('<div class="sticky-note-footer">');
        html.push(
          '<span class="timestamp-meta">🕒 ' +
            escapeHtml(formatTimestamp(note.updatedAt)) +
            "</span>"
        );
        html.push(
          '<span class="author-meta">' + escapeHtml(note.author) + "</span>"
        );
        html.push("</div></div></article>");
      });
      html.push("</div>");
    }

    html.push("</main>");

    if (state.toastMessage) {
      html.push(
        '<div class="toast-bar"><span>' +
          escapeHtml(state.toastMessage) +
          "</span></div>"
      );
    }

    if (state.isSettingsOpen) {
      var cfg = getAuthConfig();
      html.push('<div class="modal-backdrop"><div class="settings-modal">');
      html.push(
        '<div class="modal-header"><h2>Storage, Azure AD & Export Settings</h2><button type="button" class="icon-btn" id="btn-close-settings">✕</button></div>'
      );
      html.push('<div class="modal-body">');
      html.push('<div class="settings-group">');
      html.push(
        '<label class="settings-group-title">Storage Backend Mode</label>'
      );
      html.push(
        '<select id="cfg-storage-mode" class="compact-select">' +
          '<option value="hybrid"' +
          (cfg.storageMode === "hybrid" ? " selected" : "") +
          ">Hybrid (Office.js CustomProperties + Microsoft Graph Open Extensions)</option>" +
          '<option value="office_only"' +
          (cfg.storageMode === "office_only" ? " selected" : "") +
          ">Office.js CustomProperties Only (Zero Azure Config)</option>" +
          '<option value="graph_only"' +
          (cfg.storageMode === "graph_only" ? " selected" : "") +
          ">Microsoft Graph Open Extensions Only</option></select>"
      );
      html.push("</div>");

      html.push('<div class="settings-group">');
      html.push(
        '<label class="settings-group-title">Microsoft Entra ID (Azure AD) App Registration</label>'
      );
      html.push(
        '<div class="settings-field"><label>Application (Client) ID</label><input type="text" id="cfg-client-id" class="compact-text-input" placeholder="00000000-0000-0000-0000-000000000000" value="' +
          escapeHtml(cfg.clientId) +
          '" /></div>'
      );
      html.push(
        '<div class="settings-field"><label>Directory (Tenant) ID</label><input type="text" id="cfg-tenant-id" class="compact-text-input" placeholder="common" value="' +
          escapeHtml(cfg.tenantId) +
          '" /></div>'
      );
      html.push("</div>");

      html.push('<div class="settings-group">');
      html.push('<label class="settings-group-title">Backup & Export</label>');
      html.push('<div class="backup-actions-row">');
      html.push(
        '<button type="button" class="btn-secondary-sm" id="btn-export-json">⬇️ Export All Notes (JSON)</button>'
      );
      html.push("</div></div>");

      html.push('<div class="modal-footer">');
      html.push(
        '<button type="button" class="btn-primary" id="btn-save-settings">Save Settings</button>'
      );
      html.push("</div></div></div></div>");
    }

    html.push("</div>");
    root.innerHTML = html.join("");

    bindEvents();
  }

  function syncEditorInputs() {
    if (!state.editorState) return;
    var tEl = document.getElementById("ed-title");
    var cEl = document.getElementById("ed-content");
    var pEl = document.getElementById("ed-priority");
    if (tEl) state.editorState.title = tEl.value;
    if (cEl) state.editorState.content = cEl.value;
    if (pEl) state.editorState.priority = pEl.value;
  }

  function bindEvents() {
    var btnNew = document.getElementById("btn-new-note");
    if (btnNew) {
      btnNew.onclick = function () {
        initEditorState(null);
        render();
      };
    }

    var btnEmptyNew = document.getElementById("btn-empty-new");
    if (btnEmptyNew) {
      btnEmptyNew.onclick = function () {
        initEditorState(null);
        render();
      };
    }

    var btnTheme = document.getElementById("btn-toggle-theme");
    if (btnTheme) {
      btnTheme.onclick = function () {
        state.isDark = !state.isDark;
        localStorage.setItem(
          LOCAL_STORAGE_THEME_KEY,
          state.isDark ? "dark" : "light"
        );
        render();
      };
    }

    var btnSettings = document.getElementById("btn-open-settings");
    if (btnSettings) {
      btnSettings.onclick = function () {
        state.isSettingsOpen = true;
        render();
      };
    }

    var btnCloseSettings = document.getElementById("btn-close-settings");
    if (btnCloseSettings) {
      btnCloseSettings.onclick = function () {
        state.isSettingsOpen = false;
        render();
      };
    }

    var btnSaveSettings = document.getElementById("btn-save-settings");
    if (btnSaveSettings) {
      btnSaveSettings.onclick = function () {
        var modeEl = document.getElementById("cfg-storage-mode");
        var cidEl = document.getElementById("cfg-client-id");
        var tidEl = document.getElementById("cfg-tenant-id");
        saveAuthConfig({
          storageMode: modeEl ? modeEl.value : "hybrid",
          clientId: cidEl ? cidEl.value.trim() : "",
          tenantId: tidEl ? tidEl.value.trim() : "common",
          graphExtensionName: "com.jraltd.emailNotes",
        });
        state.isSettingsOpen = false;
        showToast("Settings saved");
      };
    }

    var btnExportJson = document.getElementById("btn-export-json");
    if (btnExportJson) {
      btnExportJson.onclick = function () {
        var blob = new Blob(
          [JSON.stringify({ notes: state.allEmailsNotes }, null, 2)],
          { type: "application/json" }
        );
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = "outlook-sticky-notes-backup.json";
        a.click();
        URL.revokeObjectURL(url);
      };
    }

    document.querySelectorAll("[data-scope]").forEach(function (el) {
      el.onclick = function () {
        state.viewScope = el.getAttribute("data-scope");
        render();
      };
    });

    document.querySelectorAll("[data-demo-idx]").forEach(function (el) {
      el.onclick = function () {
        var idx = parseInt(el.getAttribute("data-demo-idx"), 10);
        state.demoEmailIndex = idx;
        state.emailContext = DEMO_SAMPLE_EMAILS[idx];
        refreshCurrentEmailNotes();
      };
    });

    var searchEl = document.getElementById("input-search");
    if (searchEl) {
      searchEl.oninput = function (e) {
        state.searchQuery = e.target.value;
        var pos = e.target.selectionStart;
        render();
        var reFocused = document.getElementById("input-search");
        if (reFocused) {
          reFocused.focus();
          reFocused.setSelectionRange(pos, pos);
        }
      };
    }

    document.querySelectorAll("[data-filter-color]").forEach(function (el) {
      el.onclick = function () {
        var c = el.getAttribute("data-filter-color");
        state.selectedColor =
          state.selectedColor === c && c !== "all" ? "all" : c;
        render();
      };
    });

    var prioSelect = document.getElementById("select-priority-filter");
    if (prioSelect) {
      prioSelect.onchange = function (e) {
        state.selectedPriority = e.target.value;
        render();
      };
    }

    document.querySelectorAll("[data-filter-tag]").forEach(function (el) {
      el.onclick = function () {
        var t = el.getAttribute("data-filter-tag");
        state.selectedTag = state.selectedTag === t && t !== "all" ? "all" : t;
        render();
      };
    });

    // Editor controls
    document.querySelectorAll("[data-ed-color]").forEach(function (el) {
      el.onclick = function () {
        syncEditorInputs();
        state.editorState.color = el.getAttribute("data-ed-color");
        render();
      };
    });

    document.querySelectorAll("[data-ed-tag]").forEach(function (el) {
      el.onclick = function () {
        syncEditorInputs();
        var t = el.getAttribute("data-ed-tag");
        var idx = state.editorState.tags.indexOf(t);
        if (idx === -1) state.editorState.tags.push(t);
        else state.editorState.tags.splice(idx, 1);
        render();
      };
    });

    var cancelEd1 = document.getElementById("btn-cancel-editor");
    var cancelEd2 = document.getElementById("btn-cancel-editor-2");
    if (cancelEd1)
      cancelEd1.onclick = function () {
        state.editorState = null;
        render();
      };
    if (cancelEd2)
      cancelEd2.onclick = function () {
        state.editorState = null;
        render();
      };

    var saveEd = document.getElementById("btn-save-editor");
    if (saveEd) {
      saveEd.onclick = function () {
        syncEditorInputs();
        var ed = state.editorState;
        if (!ed.title.trim() && !ed.content.trim()) return;
        var nowIso = new Date().toISOString();
        var next;
        if (ed.id) {
          next = state.emailNotes.map(function (n) {
            if (n.id !== ed.id) return n;
            return Object.assign({}, n, {
              title: ed.title.trim(),
              content: ed.content.trim(),
              color: ed.color,
              priority: ed.priority,
              tags: ed.tags,
              updatedAt: nowIso,
            });
          });
        } else {
          var created = {
            id: "note-" + Date.now(),
            itemId: state.emailContext.itemId,
            conversationId: state.emailContext.conversationId,
            scope: ed.scope,
            emailSubject: state.emailContext.subject,
            emailSender:
              state.emailContext.senderName +
              " <" +
              state.emailContext.senderEmail +
              ">",
            title: ed.title.trim(),
            content: ed.content.trim(),
            color: ed.color,
            priority: ed.priority,
            tags: ed.tags,
            checklist: ed.checklist || [],
            isPinned: ed.isPinned,
            author: state.emailContext.userDisplayName,
            createdAt: nowIso,
            updatedAt: nowIso,
            syncStatus: "synced_office",
          };
          next = [created].concat(state.emailNotes);
        }
        state.editorState = null;
        saveItemNotesToStorage(state.emailContext.itemId, next, function () {
          refreshCurrentEmailNotes();
          showToast("Sticky note saved to email");
        });
      };
    }

    // Card actions
    document.querySelectorAll("[data-pin-id]").forEach(function (el) {
      el.onclick = function () {
        var id = el.getAttribute("data-pin-id");
        var next = state.emailNotes.map(function (n) {
          if (n.id !== id) return n;
          return Object.assign({}, n, {
            isPinned: !n.isPinned,
            updatedAt: new Date().toISOString(),
          });
        });
        saveItemNotesToStorage(state.emailContext.itemId, next, function () {
          refreshCurrentEmailNotes();
        });
      };
    });

    document.querySelectorAll("[data-edit-id]").forEach(function (el) {
      el.onclick = function () {
        var id = el.getAttribute("data-edit-id");
        var found = state.emailNotes.find(function (n) {
          return n.id === id;
        });
        if (found) {
          initEditorState(found);
          render();
        }
      };
    });

    document.querySelectorAll("[data-delete-id]").forEach(function (el) {
      el.onclick = function () {
        var id = el.getAttribute("data-delete-id");
        var next = state.emailNotes.filter(function (n) {
          return n.id !== id;
        });
        saveItemNotesToStorage(state.emailContext.itemId, next, function () {
          refreshCurrentEmailNotes();
          showToast("Sticky note deleted");
        });
      };
    });

    document.querySelectorAll("[data-chk-note]").forEach(function (el) {
      el.onchange = function () {
        var noteId = el.getAttribute("data-chk-note");
        var chkId = el.getAttribute("data-chk-id");
        var next = state.emailNotes.map(function (n) {
          if (n.id !== noteId) return n;
          return Object.assign({}, n, {
            checklist: (n.checklist || []).map(function (c) {
              if (c.id !== chkId) return c;
              return Object.assign({}, c, { completed: !c.completed });
            }),
            updatedAt: new Date().toISOString(),
          });
        });
        saveItemNotesToStorage(state.emailContext.itemId, next, function () {
          refreshCurrentEmailNotes();
        });
      };
    });
  }

  function bootstrap() {
    if (window.Office && typeof Office.onReady === "function") {
      Office.onReady(function (info) {
        state.isInOutlook = Boolean(
          info &&
            info.host === Office.HostType.Outlook &&
            Office.context &&
            Office.context.mailbox
        );
        if (
          state.isInOutlook &&
          Office.context.mailbox.addHandlerAsync
        ) {
          try {
            Office.context.mailbox.addHandlerAsync(
              Office.EventType.ItemChanged,
              function () {
                readOutlookContext(function (ctx) {
                  state.emailContext = ctx;
                  refreshCurrentEmailNotes();
                });
              }
            );
          } catch (e) {}
        }
        readOutlookContext(function (ctx) {
          state.emailContext = ctx;
          refreshCurrentEmailNotes();
        });
      });
    } else {
      readOutlookContext(function (ctx) {
        state.emailContext = ctx;
        refreshCurrentEmailNotes();
      });
    }
  }

  bootstrap();
})();
