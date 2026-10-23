/**
 * Email Sticky Notes - shared core (used by taskpane AND the ribbon "Quick Note" command)
 * Plain ES5 so it also runs in Classic Outlook's older embedded browser.
 */
(function (global) {
  "use strict";

  // ---------------------------------------------------------------------------
  // Polyfills (Classic Outlook on Windows may use an IE11-era runtime)
  // ---------------------------------------------------------------------------
  if (!Array.prototype.find) {
    Array.prototype.find = function (fn, thisArg) {
      for (var i = 0; i < this.length; i++) {
        if (fn.call(thisArg, this[i], i, this)) return this[i];
      }
      return undefined;
    };
  }
  if (typeof Object.assign !== "function") {
    Object.assign = function (target) {
      for (var i = 1; i < arguments.length; i++) {
        var src = arguments[i];
        if (src) {
          for (var k in src) {
            if (Object.prototype.hasOwnProperty.call(src, k)) target[k] = src[k];
          }
        }
      }
      return target;
    };
  }
  if (global.NodeList && !global.NodeList.prototype.forEach) {
    global.NodeList.prototype.forEach = Array.prototype.forEach;
  }

  // ---------------------------------------------------------------------------
  // Constants
  // ---------------------------------------------------------------------------
  var KEYS = {
    customProp: "jraltd_email_sticky_notes_v2",
    roamingIndex: "jraltd_email_notes_global_index_v2",
    localNotesPrefix: "outlook_sticky_notes_item_",
    localIndex: "outlook_sticky_notes_all_index",
    bannerEnabled: "jraltd_email_notes_banner_enabled",
    bannerDismissed: "jraltd_email_notes_banner_dismissed",
  };

  var NOTE_CATEGORY = "\uD83D\uDCDD Sticky Note";
  var URGENT_CATEGORY = "\uD83D\uDD25 Urgent Note";
  var BANNER_ID = "EmailStickyNotesStatus";
  var SYNC_EVENT_KEY = "outlook_sticky_notes_last_sync_event";
  // RoamingSettings has a 32 KB total limit, so keep our share well below it.
  var ROAMING_INDEX_MAX_CHARS = 24000;
  var MAX_DISMISSED = 200;

  var bcChannel = null;
  try {
    if (typeof global.BroadcastChannel === "function") {
      bcChannel = new global.BroadcastChannel("jraltd_email_sticky_notes_bc");
    }
  } catch (e) {}

  function notifyCrossFrameUpdate(itemKey) {
    var payload = JSON.stringify({ itemKey: itemKey || "", ts: Date.now() });
    try {
      global.localStorage.setItem(SYNC_EVENT_KEY, payload);
    } catch (e) {}
    try {
      if (bcChannel) bcChannel.postMessage(payload);
    } catch (e2) {}
  }

  function onCrossFrameUpdate(callback) {
    if (typeof callback !== "function") return;
    if (bcChannel) {
      try {
        bcChannel.addEventListener("message", function () {
          callback();
        });
      } catch (e) {}
    }
    if (global.addEventListener) {
      global.addEventListener("storage", function (ev) {
        if (!ev || !ev.key || ev.key === SYNC_EVENT_KEY || ev.key.indexOf("outlook_sticky_notes_") === 0) {
          callback();
        }
      });
    }
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------
  function roaming() {
    return global.Office && Office.context && Office.context.roamingSettings
      ? Office.context.roamingSettings
      : null;
  }

  function parseJson(raw, fallback) {
    if (raw === null || raw === undefined || raw === "") return fallback;
    if (typeof raw !== "string") return raw;
    try {
      var v = JSON.parse(raw);
      return v === null || v === undefined ? fallback : v;
    } catch (e) {
      return fallback;
    }
  }

  function succeeded(res) {
    return res && res.status === Office.AsyncResultStatus.Succeeded;
  }

  /** Short stable hash so per-email preferences don't eat RoamingSettings space. */
  function hashKey(str) {
    var h = 5381;
    str = String(str || "");
    for (var i = 0; i < str.length; i++) {
      h = ((h << 5) + h + str.charCodeAt(i)) | 0;
    }
    return (h >>> 0).toString(36);
  }

  function getItemKey(item) {
    if (!item) return "";
    return item.itemId || item.conversationId || "outlook-item";
  }

  // ---------------------------------------------------------------------------
  // Preferences (roam with the mailbox; localStorage fallback for browser preview)
  // ---------------------------------------------------------------------------
  function getPref(key, fallback) {
    var rs = roaming();
    if (rs) {
      try {
        var v = rs.get(key);
        if (v !== undefined && v !== null) return v;
      } catch (e) {}
    }
    try {
      var raw = global.localStorage.getItem(key);
      if (raw !== null) return JSON.parse(raw);
    } catch (e2) {}
    return fallback;
  }

  function setPref(key, value, cb) {
    try {
      global.localStorage.setItem(key, JSON.stringify(value));
    } catch (e) {}
    var rs = roaming();
    if (rs) {
      try {
        rs.set(key, value);
        rs.saveAsync(function () {
          if (cb) cb();
        });
        return;
      } catch (e2) {}
    }
    if (cb) cb();
  }

  function isBannerEnabled() {
    return getPref(KEYS.bannerEnabled, true) !== false;
  }

  function setBannerEnabled(enabled, cb) {
    setPref(KEYS.bannerEnabled, Boolean(enabled), cb);
  }

  function getDismissedList() {
    var list = getPref(KEYS.bannerDismissed, []);
    return Array.isArray(list) ? list : [];
  }

  function isBannerDismissed(itemKey) {
    return getDismissedList().indexOf(hashKey(itemKey)) !== -1;
  }

  function setBannerDismissed(itemKey, dismissed, cb) {
    var h = hashKey(itemKey);
    var list = getDismissedList().filter(function (x) {
      return x !== h;
    });
    if (dismissed) list.unshift(h);
    setPref(KEYS.bannerDismissed, list.slice(0, MAX_DISMISSED), cb);
  }

  // ---------------------------------------------------------------------------
  // Per-email storage (Office.js CustomProperties, with localStorage cache)
  // ---------------------------------------------------------------------------
  function readLocalItemNotes(itemKey) {
    try {
      var arr = parseJson(global.localStorage.getItem(KEYS.localNotesPrefix + itemKey), []);
      return Array.isArray(arr) ? arr : [];
    } catch (e) {
      return [];
    }
  }

  function writeLocalItemNotes(itemKey, notes) {
    try {
      global.localStorage.setItem(KEYS.localNotesPrefix + itemKey, JSON.stringify(notes));
    } catch (e) {}
  }

  /** cb(notesArray) */
  function loadItemNotes(item, itemKey, cb) {
    var local = readLocalItemNotes(itemKey);
    if (!item || typeof item.loadCustomPropertiesAsync !== "function") {
      cb(local);
      return;
    }
    try {
      item.loadCustomPropertiesAsync(function (res) {
        if (succeeded(res)) {
          var raw = res.value.get(KEYS.customProp);
          if (raw) {
            var parsed = parseJson(raw, null);
            if (Array.isArray(parsed)) {
              writeLocalItemNotes(itemKey, parsed);
              cb(parsed);
              return;
            }
          } else {
            // Mailbox is the source of truth: no property means no notes.
            cb([]);
            return;
          }
        }
        cb(local);
      });
    } catch (e) {
      cb(local);
    }
  }

  /** cb(ok:boolean) */
  function saveItemNotes(item, itemKey, notes, cb) {
    writeLocalItemNotes(itemKey, notes);
    if (!item || typeof item.loadCustomPropertiesAsync !== "function") {
      cb(true);
      return;
    }
    try {
      item.loadCustomPropertiesAsync(function (res) {
        if (!succeeded(res)) {
          cb(false);
          return;
        }
        var cp = res.value;
        if (notes.length > 0) {
          cp.set(KEYS.customProp, JSON.stringify(notes));
        } else if (typeof cp.remove === "function") {
          cp.remove(KEYS.customProp);
        } else {
          cp.set(KEYS.customProp, "[]");
        }
        cp.saveAsync(function (saveRes) {
          cb(succeeded(saveRes));
        });
      });
    } catch (e) {
      cb(false);
    }
  }

  // ---------------------------------------------------------------------------
  // Cross-email index ("All Annotated Emails" tab)
  // ---------------------------------------------------------------------------
  function sortByUpdated(list) {
    return list.sort(function (a, b) {
      return String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
    });
  }

  function readIndex() {
    var map = {};
    var rs = roaming();
    if (rs) {
      try {
        var arrR = parseJson(rs.get(KEYS.roamingIndex), []);
        if (Array.isArray(arrR)) {
          arrR.forEach(function (n) {
            if (n && n.id) map[n.id] = n;
          });
        }
      } catch (e) {}
    }
    try {
      var arrL = parseJson(global.localStorage.getItem(KEYS.localIndex), []);
      if (Array.isArray(arrL)) {
        arrL.forEach(function (n) {
          if (!n || !n.id) return;
          // Prefer the full local copy when timestamps are equal (roaming copy is compacted).
          if (!map[n.id] || String(n.updatedAt) >= String(map[n.id].updatedAt)) {
            map[n.id] = n;
          }
        });
      }
    } catch (e2) {}
    return sortByUpdated(
      Object.keys(map).map(function (k) {
        return map[k];
      })
    );
  }

  function compactForRoaming(n) {
    return {
      id: n.id,
      itemId: n.itemId,
      conversationId: n.conversationId,
      scope: n.scope,
      emailSubject: String(n.emailSubject || "").slice(0, 120),
      emailSender: String(n.emailSender || "").slice(0, 120),
      title: String(n.title || "").slice(0, 120),
      content: String(n.content || "").slice(0, 160),
      color: n.color,
      priority: n.priority,
      tags: (n.tags || []).slice(0, 8),
      checklist: [],
      isPinned: Boolean(n.isPinned),
      author: n.author,
      createdAt: n.createdAt,
      updatedAt: n.updatedAt,
    };
  }

  function writeIndex(allNotes, cb) {
    var sorted = sortByUpdated(allNotes.slice());
    try {
      global.localStorage.setItem(KEYS.localIndex, JSON.stringify(sorted));
    } catch (e) {}

    var rs = roaming();
    if (!rs) {
      if (cb) cb();
      return;
    }
    var compact = sorted.map(compactForRoaming);
    var json = JSON.stringify(compact);
    while (json.length > ROAMING_INDEX_MAX_CHARS && compact.length > 0) {
      compact.pop();
      json = JSON.stringify(compact);
    }
    try {
      rs.set(KEYS.roamingIndex, json);
      rs.saveAsync(function () {
        if (cb) cb();
      });
    } catch (e2) {
      if (cb) cb();
    }
  }

  function replaceItemInIndex(itemKey, itemNotes, cb) {
    var others = readIndex().filter(function (n) {
      return n.itemId !== itemKey;
    });
    writeIndex(itemNotes.concat(others), function () {
      notifyCrossFrameUpdate(itemKey);
      if (cb) cb();
    });
  }

  // ---------------------------------------------------------------------------
  // Visual markers: banner above the email + Outlook color categories
  // ---------------------------------------------------------------------------
  function removeBanner(item, cb) {
    if (!item || !item.notificationMessages) {
      if (cb) cb();
      return;
    }
    try {
      item.notificationMessages.removeAsync(BANNER_ID, function () {
        if (cb) cb();
      });
    } catch (e) {
      if (cb) cb();
    }
  }

  function showBanner(item, text, isError, cb) {
    if (!item || !item.notificationMessages) {
      if (cb) cb();
      return;
    }
    var details = isError
      ? {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: String(text).slice(0, 150),
        }
      : {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: String(text).slice(0, 150),
          icon: "Icon.16x16",
          persistent: false,
        };
    try {
      item.notificationMessages.replaceAsync(BANNER_ID, details, function () {
        if (cb) cb();
      });
    } catch (e) {
      if (cb) cb();
    }
  }

  function bannerTextFor(notes) {
    var top =
      notes.find(function (n) {
        return n.isPinned;
      }) || notes[0];
    var preview = (top && (top.title || String(top.content || "").slice(0, 50))) || "Note attached";
    return notes.length === 1
      ? '\uD83D\uDCDD 1 Sticky Note: "' + preview + '"'
      : "\uD83D\uDCDD " + notes.length + ' Sticky Notes - Top: "' + preview + '"';
  }

  /** Applies (or clears) the banner according to the user's preferences. */
  function syncBanner(item, itemKey, notes, overrideText, cb) {
    if (notes.length > 0 && isBannerEnabled() && !isBannerDismissed(itemKey)) {
      showBanner(item, overrideText || bannerTextFor(notes), false, cb);
    } else {
      removeBanner(item, cb);
    }
  }

  function ensureMasterCategories(cb) {
    var mc = global.Office && Office.context && Office.context.mailbox
      ? Office.context.mailbox.masterCategories
      : null;
    if (!mc || typeof mc.getAsync !== "function") {
      cb();
      return;
    }
    try {
      mc.getAsync(function (res) {
        var names = succeeded(res)
          ? (res.value || []).map(function (c) {
              return c.displayName;
            })
          : [];
        var missing = [];
        if (names.indexOf(NOTE_CATEGORY) === -1) {
          missing.push({ displayName: NOTE_CATEGORY, color: Office.MailboxEnums.CategoryColor.Preset3 });
        }
        if (names.indexOf(URGENT_CATEGORY) === -1) {
          missing.push({ displayName: URGENT_CATEGORY, color: Office.MailboxEnums.CategoryColor.Preset0 });
        }
        if (missing.length === 0) {
          cb();
          return;
        }
        // Must finish BEFORE tagging the email, otherwise the item category call can fail.
        mc.addAsync(missing, function () {
          cb();
        });
      });
    } catch (e) {
      cb();
    }
  }

  function syncCategories(item, notes, cb) {
    if (!item || !item.categories || typeof item.categories.getAsync !== "function") {
      cb();
      return;
    }
    var count = notes.length;
    var hasUrgent = notes.some(function (n) {
      return n.priority === "urgent";
    });

    var apply = function () {
      try {
        item.categories.getAsync(function (res) {
          if (!succeeded(res)) {
            cb();
            return;
          }
          var current = res.value || [];
          var hasNoteCat = current.indexOf(NOTE_CATEGORY) !== -1;
          var hasUrgentCat = current.indexOf(URGENT_CATEGORY) !== -1;
          var toAdd = [];
          var toRemove = [];

          if (count > 0 && !hasNoteCat) toAdd.push(NOTE_CATEGORY);
          if (count === 0 && hasNoteCat) toRemove.push(NOTE_CATEGORY);
          if (count > 0 && hasUrgent && !hasUrgentCat) toAdd.push(URGENT_CATEGORY);
          if ((count === 0 || !hasUrgent) && hasUrgentCat) toRemove.push(URGENT_CATEGORY);

          var doRemove = function () {
            if (toRemove.length === 0) {
              cb();
              return;
            }
            item.categories.removeAsync(toRemove, function () {
              cb();
            });
          };
          if (toAdd.length > 0) {
            item.categories.addAsync(toAdd, function () {
              doRemove();
            });
          } else {
            doRemove();
          }
        });
      } catch (e) {
        cb();
      }
    };

    if (count > 0) {
      ensureMasterCategories(apply);
    } else {
      apply();
    }
  }

  function syncMarkers(item, itemKey, notes, overrideBannerText, cb) {
    syncBanner(item, itemKey, notes, overrideBannerText, function () {
      syncCategories(item, notes, function () {
        if (cb) cb();
      });
    });
  }

  // ---------------------------------------------------------------------------
  // Note factory
  // ---------------------------------------------------------------------------
  function newId(prefix) {
    return (prefix || "note") + "-" + Date.now() + "-" + Math.random().toString(36).slice(2, 7);
  }

  function senderOf(item) {
    var from = (item && (item.from || item.sender)) || null;
    if (from && from.emailAddress) {
      return (from.displayName || from.emailAddress) + " <" + from.emailAddress + ">";
    }
    return "Sender";
  }

  function subjectOf(item) {
    return item && typeof item.subject === "string" && item.subject ? item.subject : "(No Subject)";
  }

  function userDisplayName() {
    var up = global.Office && Office.context && Office.context.mailbox
      ? Office.context.mailbox.userProfile
      : null;
    return (up && up.displayName) || "Outlook User";
  }

  function createQuickNote(item, itemKey) {
    var nowIso = new Date().toISOString();
    return {
      id: newId("quick"),
      itemId: itemKey,
      conversationId: (item && item.conversationId) || itemKey,
      scope: "message",
      emailSubject: subjectOf(item),
      emailSender: senderOf(item),
      title: "Follow up",
      content: "Flagged for follow-up on " + new Date().toLocaleString() + ".",
      color: "yellow",
      priority: "high",
      tags: ["follow-up"],
      checklist: [],
      isPinned: true,
      author: userDisplayName(),
      createdAt: nowIso,
      updatedAt: nowIso,
      syncStatus: "synced_office",
    };
  }

  global.StickyNotesCore = {
    KEYS: KEYS,
    NOTE_CATEGORY: NOTE_CATEGORY,
    URGENT_CATEGORY: URGENT_CATEGORY,
    getItemKey: getItemKey,
    hashKey: hashKey,
    newId: newId,
    loadItemNotes: loadItemNotes,
    saveItemNotes: saveItemNotes,
    readIndex: readIndex,
    writeIndex: writeIndex,
    replaceItemInIndex: replaceItemInIndex,
    isBannerEnabled: isBannerEnabled,
    setBannerEnabled: setBannerEnabled,
    isBannerDismissed: isBannerDismissed,
    setBannerDismissed: setBannerDismissed,
    showBanner: showBanner,
    removeBanner: removeBanner,
    syncBanner: syncBanner,
    syncCategories: syncCategories,
    syncMarkers: syncMarkers,
    createQuickNote: createQuickNote,
    senderOf: senderOf,
    subjectOf: subjectOf,
    userDisplayName: userDisplayName,
    notifyCrossFrameUpdate: notifyCrossFrameUpdate,
    onCrossFrameUpdate: onCrossFrameUpdate,
  };
})(typeof window !== "undefined" ? window : this);
