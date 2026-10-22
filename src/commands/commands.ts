import { officeContextService } from "../services/office/officeContextService";
import { storageManager } from "../services/storage/storageManager";
import { StickyNote } from "../types/note";

/**
 * Ribbon FunctionCommand: Stamps a quick timestamped follow-up sticky note
 * onto the currently selected Outlook email without even opening the taskpane.
 */
async function addQuickStickyNote(event: Office.AddinCommands.Event) {
  try {
    const ctx = await officeContextService.initialize();
    const { notes } = await storageManager.loadNotesForEmail(ctx, false);

    const nowIso = new Date().toISOString();
    const quickNote: StickyNote = {
      id: `quick-${Date.now()}`,
      itemId: ctx.itemId,
      conversationId: ctx.conversationId,
      scope: "message",
      emailSubject: ctx.subject,
      emailSender: `${ctx.senderName} <${ctx.senderEmail}>`,
      title: "Quick Follow-up Stamp",
      content: `Flagged for follow-up on ${new Date().toLocaleString()}.`,
      color: "yellow",
      priority: "high",
      tags: ["follow-up"],
      checklist: [],
      isPinned: true,
      author: ctx.userDisplayName,
      createdAt: nowIso,
      updatedAt: nowIso,
      syncStatus: "synced_office",
    };

    const nextNotes = [quickNote, ...notes];
    await storageManager.saveNotesForEmail(ctx, nextNotes, false);

    officeContextService.showEmailNotificationBanner(
      `Quick Sticky Note stamped (${nextNotes.length} total on this email).`
    );
  } catch (err) {
    console.error("Failed to stamp quick sticky note:", err);
  } finally {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }
}

if (typeof Office !== "undefined") {
  Office.onReady(() => {
    if (Office.actions) {
      Office.actions.associate("addQuickStickyNote", addQuickStickyNote);
    }
  });
}

(window as unknown as Record<string, unknown>).addQuickStickyNote =
  addQuickStickyNote;
