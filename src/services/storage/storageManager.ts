import { DEMO_SAMPLE_EMAILS } from "../../config/constants";
import { EmailContextInfo, StickyNote } from "../../types/note";
import { msalAuthService } from "../auth/msalAuthService";
import { officeContextService } from "../office/officeContextService";
import { graphOpenExtensionAdapter } from "./graphOpenExtensionAdapter";
import { officeCustomPropsAdapter } from "./officeCustomPropsAdapter";

const DEMO_SEED_FLAG_KEY = "outlook_sticky_notes_demo_seeded_v2";

export class StorageManager {
  public async ensureDemoSeedNotes(): Promise<void> {
    if (officeContextService.isInOutlook()) {
      return;
    }

    if (localStorage.getItem(DEMO_SEED_FLAG_KEY) === "true") {
      return;
    }

    const email1 = DEMO_SAMPLE_EMAILS[0];
    const email2 = DEMO_SAMPLE_EMAILS[1];

    const now = Date.now();
    const demoNotesEmail1: StickyNote[] = [
      {
        id: "seed-note-101",
        itemId: email1.itemId,
        conversationId: email1.conversationId,
        scope: "message",
        emailSubject: email1.subject,
        emailSender: `${email1.senderName} <${email1.senderEmail}>`,
        title: "Negotiate 12% Volume Discount Tier",
        content:
          "Sarah confirmed flexibility on the 3-year enterprise renewal if we sign before the end of the quarter. Verify seat count with IT Ops before replying.",
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
        reminderAt: new Date(now + 1000 * 60 * 60 * 24).toISOString(),
        author: email1.userDisplayName,
        createdAt: new Date(now - 1000 * 60 * 35).toISOString(),
        updatedAt: new Date(now - 1000 * 60 * 12).toISOString(),
        syncStatus: "synced_office",
      },
      {
        id: "seed-note-102",
        itemId: email1.itemId,
        conversationId: email1.conversationId,
        scope: "conversation",
        emailSubject: email1.subject,
        emailSender: `${email1.senderName} <${email1.senderEmail}>`,
        title: "PO Approval Reference #PO-88412",
        content:
          "Finance pre-approved budget code FIN-2026-Q4. Attach PO number when final MSA is countersigned.",
        color: "green",
        priority: "medium",
        tags: ["decision", "billing"],
        checklist: [],
        isPinned: false,
        author: email1.userDisplayName,
        createdAt: new Date(now - 1000 * 60 * 90).toISOString(),
        updatedAt: new Date(now - 1000 * 60 * 90).toISOString(),
        syncStatus: "synced_office",
      },
    ];

    const demoNotesEmail2: StickyNote[] = [
      {
        id: "seed-note-201",
        itemId: email2.itemId,
        conversationId: email2.conversationId,
        scope: "message",
        emailSubject: email2.subject,
        emailSender: `${email2.senderName} <${email2.senderEmail}>`,
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
        author: email2.userDisplayName,
        createdAt: new Date(now - 1000 * 60 * 180).toISOString(),
        updatedAt: new Date(now - 1000 * 60 * 180).toISOString(),
        syncStatus: "synced_office",
      },
    ];

    await officeCustomPropsAdapter.saveNotesForItem(
      email1.itemId,
      demoNotesEmail1
    );
    await officeCustomPropsAdapter.saveNotesForItem(
      email2.itemId,
      demoNotesEmail2
    );
    localStorage.setItem(DEMO_SEED_FLAG_KEY, "true");
  }

  /**
   * Loads notes for the active email from Office.js CustomProperties and (if configured)
   * Microsoft Graph Open Extensions, merging by most recent `updatedAt`.
   */
  public async loadNotesForEmail(
    emailContext: EmailContextInfo,
    interactiveGraph = false
  ): Promise<{
    notes: StickyNote[];
    graphSynced: boolean;
  }> {
    await this.ensureDemoSeedNotes();
    const mode = msalAuthService.getConfig().storageMode;

    let officeNotes: StickyNote[] = [];
    if (mode !== "graph_only") {
      officeNotes = await officeCustomPropsAdapter.loadNotesForItem(
        emailContext.itemId
      );
    }

    // Also include conversation-scoped notes matching this email's conversationId
    const threadNotes = officeCustomPropsAdapter
      .getAllIndexedNotes()
      .filter(
        (n) =>
          n.scope === "conversation" &&
          n.conversationId === emailContext.conversationId
      );

    let graphNotes: StickyNote[] | null = null;
    let graphSynced = false;

    if (mode !== "office_only" && msalAuthService.isGraphConfigured()) {
      try {
        graphNotes = await graphOpenExtensionAdapter.fetchNotesFromGraph(
          emailContext.restId,
          interactiveGraph
        );
        if (graphNotes !== null) {
          graphSynced = true;
        }
      } catch (err) {
        console.warn("Graph Open Extension fetch skipped/failed:", err);
      }
    }

    const merged = this.mergeNoteLists([
      ...officeNotes,
      ...threadNotes,
      ...(graphNotes || []),
    ]);

    // If Graph returned newer notes, persist back to Office CustomProperties
    if (graphSynced && graphNotes && graphNotes.length > 0) {
      await officeCustomPropsAdapter.saveNotesForItem(
        emailContext.itemId,
        merged
      );
    }

    // Update email category pill + notification banner in Outlook
    await officeContextService.syncEmailVisualMarkers(merged);

    return {
      notes: merged,
      graphSynced,
    };
  }

  /**
   * Persists the updated note list for an email to Office.js CustomProperties
   * and (when configured) Microsoft Graph Open Extensions.
   */
  public async saveNotesForEmail(
    emailContext: EmailContextInfo,
    notes: StickyNote[],
    interactiveGraph = false
  ): Promise<{
    savedOffice: boolean;
    savedGraph: boolean;
    updatedNotes: StickyNote[];
  }> {
    const mode = msalAuthService.getConfig().storageMode;

    let savedGraph = false;
    if (mode !== "office_only" && msalAuthService.isGraphConfigured()) {
      try {
        savedGraph = await graphOpenExtensionAdapter.saveNotesToGraph(
          emailContext.restId,
          notes,
          interactiveGraph
        );
      } catch (err) {
        console.warn("Graph Open Extension save failed:", err);
      }
    }

    const finalStatus: StickyNote["syncStatus"] = savedGraph
      ? mode === "graph_only"
        ? "synced_graph"
        : "synced_both"
      : officeContextService.isInOutlook()
      ? "synced_office"
      : "local_only";

    const updatedNotes = notes.map((n) => ({
      ...n,
      syncStatus: finalStatus,
    }));

    const savedOffice =
      mode !== "graph_only"
        ? await officeCustomPropsAdapter.saveNotesForItem(
            emailContext.itemId,
            updatedNotes
          )
        : true;

    await officeContextService.syncEmailVisualMarkers(updatedNotes);

    return {
      savedOffice,
      savedGraph,
      updatedNotes,
    };
  }

  public getAllEmailsNotes(): StickyNote[] {
    return officeCustomPropsAdapter.getAllIndexedNotes();
  }

  private mergeNoteLists(lists: StickyNote[]): StickyNote[] {
    const byId = new Map<string, StickyNote>();
    for (const note of lists) {
      if (!note || !note.id) continue;
      const existing = byId.get(note.id);
      if (!existing || note.updatedAt >= existing.updatedAt) {
        byId.set(note.id, note);
      }
    }

    return Array.from(byId.values()).sort((a, b) => {
      if (a.isPinned !== b.isPinned) {
        return a.isPinned ? -1 : 1;
      }
      return b.updatedAt.localeCompare(a.updatedAt);
    });
  }

  public exportNotesAsJson(notes: StickyNote[]): string {
    return JSON.stringify(
      {
        application: "Outlook Email Sticky Notes Add-in",
        version: "2.0.0",
        exportedAt: new Date().toISOString(),
        noteCount: notes.length,
        notes,
      },
      null,
      2
    );
  }

  public exportNotesAsMarkdown(notes: StickyNote[]): string {
    const lines: string[] = [
      `# Outlook Email Sticky Notes Export`,
      `Exported: ${new Date().toLocaleString()} (${notes.length} notes)`,
      "",
    ];

    for (const n of notes) {
      lines.push(`## ${n.isPinned ? "[PINNED] " : ""}${n.title || "Untitled Sticky Note"}`);
      lines.push(`- **Email Subject:** ${n.emailSubject}`);
      lines.push(`- **From:** ${n.emailSender}`);
      lines.push(`- **Priority:** ${n.priority.toUpperCase()} | **Color:** ${n.color}`);
      lines.push(`- **Created:** ${new Date(n.createdAt).toLocaleString()} | **Updated:** ${new Date(n.updatedAt).toLocaleString()}`);
      if (n.tags.length > 0) {
        lines.push(`- **Tags:** ${n.tags.map((t) => `#${t}`).join(" ")}`);
      }
      lines.push("");
      if (n.content) {
        lines.push(n.content);
        lines.push("");
      }
      if (n.checklist.length > 0) {
        lines.push("### Checklist");
        for (const item of n.checklist) {
          lines.push(`- [${item.completed ? "x" : " "}] ${item.text}`);
        }
        lines.push("");
      }
      lines.push("---");
      lines.push("");
    }

    return lines.join("\n");
  }

  public async importNotesFromJson(
    rawJson: string,
    currentEmail: EmailContextInfo
  ): Promise<number> {
    const parsed = JSON.parse(rawJson);
    const incoming: StickyNote[] = Array.isArray(parsed)
      ? parsed
      : Array.isArray(parsed?.notes)
      ? parsed.notes
      : [];

    if (incoming.length === 0) {
      return 0;
    }

    const allExisting = this.getAllEmailsNotes();
    const mergedGlobal = this.mergeNoteLists([...incoming, ...allExisting]);
    await officeCustomPropsAdapter.replaceGlobalIndex(mergedGlobal);

    // Also update current email's item storage if any imported notes belong to it
    const currentEmailNotes = mergedGlobal.filter(
      (n) =>
        n.itemId === currentEmail.itemId ||
        (n.scope === "conversation" &&
          n.conversationId === currentEmail.conversationId)
    );
    if (currentEmailNotes.length > 0) {
      await officeCustomPropsAdapter.saveNotesForItem(
        currentEmail.itemId,
        currentEmailNotes
      );
    }

    return incoming.length;
  }
}

export const storageManager = new StorageManager();
