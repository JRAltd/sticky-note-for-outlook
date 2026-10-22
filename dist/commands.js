/**
 * Ribbon "Quick Note" command (ExecuteFunction).
 * Stamps a pinned follow-up sticky note on the selected email without opening the pane.
 */
(function () {
  "use strict";

  var Core = window.StickyNotesCore;
  var SAFETY_TIMEOUT_MS = 20000;

  function runQuickNoteOnItem(item, complete) {
    var itemKey = Core.getItemKey(item);

    Core.loadItemNotes(item, itemKey, function (existing) {
      var quickNote = Core.createQuickNote(item, itemKey);
      var updated = [quickNote].concat(existing || []);

      Core.saveItemNotes(item, itemKey, updated, function (ok) {
        if (!ok) {
          Core.showBanner(
            item,
            "Quick Note could not be saved. Open Email Notes and try again.",
            true,
            complete
          );
          return;
        }
        // Keep the "All Annotated Emails" index and any open taskpane in sync, then tag the email.
        Core.replaceItemInIndex(itemKey, updated, function () {
          Core.syncMarkers(
            item,
            itemKey,
            updated,
            "📝 Quick Note added (" + updated.length + " on this email)",
            complete
          );
        });
      });
    });
  }

  function addQuickStickyNote(event) {
    var finished = false;
    var safetyTimer = null;

    // event.completed() MUST be called exactly once, or Outlook shows a stuck "working" state.
    function complete() {
      if (finished) return;
      finished = true;
      if (safetyTimer) clearTimeout(safetyTimer);
      try {
        if (event && typeof event.completed === "function") event.completed();
      } catch (e) {}
    }
    safetyTimer = setTimeout(complete, SAFETY_TIMEOUT_MS);

    try {
      var executeWhenReady = function () {
        try {
          var item =
            window.Office && Office.context && Office.context.mailbox
              ? Office.context.mailbox.item
              : null;
          if (!item) {
            complete();
            return;
          }
          runQuickNoteOnItem(item, complete);
        } catch (err) {
          complete();
        }
      };

      if (window.Office && typeof Office.onReady === "function") {
        Office.onReady(function () {
          executeWhenReady();
        });
      } else {
        executeWhenReady();
      }
    } catch (e) {
      complete();
    }
  }

  window.addQuickStickyNote = addQuickStickyNote;

  if (window.Office && typeof Office.onReady === "function") {
    Office.onReady(function () {
      try {
        if (Office.actions && typeof Office.actions.associate === "function") {
          Office.actions.associate("addQuickStickyNote", addQuickStickyNote);
        }
      } catch (e) {}
    });
  }
})();
