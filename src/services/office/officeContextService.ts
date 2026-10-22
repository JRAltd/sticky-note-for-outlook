import {
  DEMO_SAMPLE_EMAILS,
  OUTLOOK_NOTE_CATEGORY_NAME,
  OUTLOOK_URGENT_CATEGORY_NAME,
} from "../../config/constants";
import { EmailContextInfo, StickyNote } from "../../types/note";

type ItemChangedCallback = (context: EmailContextInfo) => void;

class OfficeContextService {
  private isOfficeInitialized = false;
  private isRunningInOutlook = false;
  private currentDemoIndex = 0;
  private listeners: Set<ItemChangedCallback> = new Set();

  public async initialize(): Promise<EmailContextInfo> {
    if (this.isOfficeInitialized) {
      return this.getCurrentEmailContext();
    }

    return new Promise<EmailContextInfo>((resolve) => {
      const timeoutId = window.setTimeout(() => {
        // Fallback if Office.onReady doesn't fire in a standard web browser
        this.isOfficeInitialized = true;
        this.isRunningInOutlook = false;
        resolve(this.getCurrentEmailContext());
      }, 1200);

      if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
        Office.onReady((info) => {
          window.clearTimeout(timeoutId);
          this.isOfficeInitialized = true;
          this.isRunningInOutlook =
            info.host === Office.HostType.Outlook &&
            Boolean(Office.context?.mailbox);

          if (this.isRunningInOutlook) {
            this.registerItemChangedHandler();
          }

          this.getCurrentEmailContext().then(resolve);
        });
      } else {
        window.clearTimeout(timeoutId);
        this.isOfficeInitialized = true;
        this.isRunningInOutlook = false;
        resolve(this.getCurrentEmailContext());
      }
    });
  }

  public isInOutlook(): boolean {
    return this.isRunningInOutlook;
  }

  private registerItemChangedHandler(): void {
    try {
      if (Office.context?.mailbox?.addHandlerAsync) {
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.ItemChanged,
          async () => {
            const updatedContext = await this.getCurrentEmailContext();
            this.notifyListeners(updatedContext);
          }
        );
      }
    } catch (err) {
      console.warn("Could not register ItemChanged handler:", err);
    }
  }

  public onItemChanged(callback: ItemChangedCallback): () => void {
    this.listeners.add(callback);
    return () => {
      this.listeners.delete(callback);
    };
  }

  private notifyListeners(ctx: EmailContextInfo): void {
    this.listeners.forEach((cb) => cb(ctx));
  }

  /**
   * Switches between sample demo emails when testing in a regular browser outside Outlook.
   */
  public selectDemoEmail(index: number): EmailContextInfo {
    this.currentDemoIndex =
      ((index % DEMO_SAMPLE_EMAILS.length) + DEMO_SAMPLE_EMAILS.length) %
      DEMO_SAMPLE_EMAILS.length;
    const ctx = DEMO_SAMPLE_EMAILS[this.currentDemoIndex];
    this.notifyListeners(ctx);
    return ctx;
  }

  public getDemoEmails(): EmailContextInfo[] {
    return DEMO_SAMPLE_EMAILS;
  }

  public async getCurrentEmailContext(): Promise<EmailContextInfo> {
    if (!this.isRunningInOutlook || !Office.context?.mailbox?.item) {
      return DEMO_SAMPLE_EMAILS[this.currentDemoIndex];
    }

    const mailbox = Office.context.mailbox;
    const item = mailbox.item;
    const userProfile = mailbox.userProfile;

    // Determine Read vs Compose mode
    const isCompose = typeof (item as Office.MessageCompose).subject?.getAsync === "function";

    // Get itemId (in compose mode, itemId may require saveAsync if brand new draft)
    let rawItemId = item.itemId || "";
    if (!rawItemId && isCompose && typeof item.saveAsync === "function") {
      rawItemId = await new Promise<string>((resolve) => {
        item.saveAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            resolve(res.value || "compose-draft-item");
          } else {
            resolve("compose-draft-item");
          }
        });
      });
    }

    // Convert EWS itemId to Graph REST v2.0 ID
    let restId = rawItemId;
    try {
      if (rawItemId && mailbox.convertToRestId) {
        restId = mailbox.convertToRestId(
          rawItemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
      }
    } catch {
      restId = rawItemId;
    }

    // Read subject
    let subject = "(No Subject)";
    if (isCompose) {
      subject = await new Promise<string>((resolve) => {
        (item as Office.MessageCompose).subject.getAsync((res) => {
          resolve(
            res.status === Office.AsyncResultStatus.Succeeded && res.value
              ? res.value
              : "(Draft Message)"
          );
        });
      });
    } else {
      subject = (item as Office.MessageRead).subject || "(No Subject)";
    }

    // Read sender info
    let senderName = userProfile?.displayName || "Sender";
    let senderEmail = userProfile?.emailAddress || "";
    if (!isCompose) {
      const fromObj =
        (item as Office.MessageRead).from ||
        (item as Office.MessageRead).sender;
      if (fromObj) {
        senderName = fromObj.displayName || fromObj.emailAddress || senderName;
        senderEmail = fromObj.emailAddress || senderEmail;
      }
    }

    const receivedDateTime =
      (!isCompose &&
        (item as Office.MessageRead).dateTimeCreated?.toISOString?.()) ||
      new Date().toISOString();

    const conversationId =
      item.conversationId || restId || rawItemId || "default-conversation";

    const diagnosticsHost = mailbox.diagnostics?.hostName || "";
    const hostEnvironment =
      diagnosticsHost.toLowerCase().includes("outlookweb") ||
      window.location.ancestorOrigins?.[0]?.includes("office")
        ? "outlook_web"
        : "outlook_desktop";

    return {
      itemId: rawItemId || conversationId,
      restId: restId || rawItemId || conversationId,
      conversationId,
      subject,
      senderName,
      senderEmail,
      receivedDateTime,
      itemType:
        item.itemType === Office.MailboxEnums.ItemType.Appointment
          ? "appointment"
          : "message",
      mode: isCompose ? "compose" : "read",
      hostEnvironment,
      userEmail: userProfile?.emailAddress || "user@outlook.com",
      userDisplayName: userProfile?.displayName || "Outlook User",
    };
  }

  /**
   * Displays an informational banner at the top of the email inside Outlook when notes exist,
   * or removes the banner when all notes on the email are deleted.
   */
  public showEmailNotificationBanner(message: string | null): void {
    if (!this.isRunningInOutlook || !Office.context?.mailbox?.item?.notificationMessages) {
      return;
    }

    try {
      if (!message) {
        Office.context.mailbox.item.notificationMessages.removeAsync(
          "EmailStickyNotesStatus"
        );
        return;
      }

      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "EmailStickyNotesStatus",
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: message.slice(0, 150),
          icon: "Icon.16x16",
          persistent: false,
        }
      );
    } catch {
      // Ignore if notification messages are unsupported in current surface
    }
  }

  /**
   * Automatically marks (or unmarks) the email in the Outlook Inbox list using:
   *  1. Outlook Color Categories ("📝 Sticky Note" in Yellow, and "🔥 Urgent Note" in Red)
   *     so a visible badge appears right on the email row in the Inbox!
   *  2. Top-of-email Reading Pane Notification Banner showing the note count and top note title.
   */
  public async syncEmailVisualMarkers(notes: StickyNote[]): Promise<void> {
    const count = notes.length;
    const hasUrgent = notes.some((n) => n.priority === "urgent");
    const topNote = notes.find((n) => n.isPinned) || notes[0];

    // 1. Update Reading Pane info banner
    if (count > 0 && topNote) {
      const previewTitle = topNote.title || topNote.content.slice(0, 50) || "Note attached";
      const bannerText =
        count === 1
          ? `📝 1 Sticky Note: "${previewTitle}"`
          : `📝 ${count} Sticky Notes — Top: "${previewTitle}"`;
      this.showEmailNotificationBanner(bannerText);
    } else {
      this.showEmailNotificationBanner(null);
    }

    if (!this.isRunningInOutlook || !Office.context?.mailbox?.item) {
      return;
    }

    // 2. Ensure Master Categories exist in the user's mailbox ("📝 Sticky Note" & "🔥 Urgent Note")
    await this.ensureMasterCategoriesExist();

    // 3. Apply or remove category tags on the email item itself
    const itemCategories = Office.context.mailbox.item.categories;
    if (!itemCategories || typeof itemCategories.getAsync !== "function") {
      return;
    }

    await new Promise<void>((resolve) => {
      itemCategories.getAsync((res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        const existingCategories: string[] = res.value || [];
        const hasStickyCat = existingCategories.includes(OUTLOOK_NOTE_CATEGORY_NAME);
        const hasUrgentCat = existingCategories.includes(OUTLOOK_URGENT_CATEGORY_NAME);

        const toAdd: string[] = [];
        const toRemove: string[] = [];

        if (count > 0 && !hasStickyCat) {
          toAdd.push(OUTLOOK_NOTE_CATEGORY_NAME);
        } else if (count === 0 && hasStickyCat) {
          toRemove.push(OUTLOOK_NOTE_CATEGORY_NAME);
        }

        if (count > 0 && hasUrgent && !hasUrgentCat) {
          toAdd.push(OUTLOOK_URGENT_CATEGORY_NAME);
        } else if ((count === 0 || !hasUrgent) && hasUrgentCat) {
          toRemove.push(OUTLOOK_URGENT_CATEGORY_NAME);
        }

        const finishRemove = () => {
          if (toRemove.length > 0) {
            itemCategories.removeAsync(toRemove, () => resolve());
          } else {
            resolve();
          }
        };

        if (toAdd.length > 0) {
          itemCategories.addAsync(toAdd, () => finishRemove());
        } else {
          finishRemove();
        }
      });
    });
  }

  private async ensureMasterCategoriesExist(): Promise<void> {
    const masterCategories = Office.context?.mailbox?.masterCategories;
    if (!masterCategories || typeof masterCategories.getAsync !== "function") {
      return;
    }

    return new Promise<void>((resolve) => {
      masterCategories.getAsync((res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        const existingNames = new Set(
          (res.value || []).map((c: Office.CategoryDetails) => c.displayName)
        );
        const missing: Office.CategoryDetails[] = [];

        if (!existingNames.has(OUTLOOK_NOTE_CATEGORY_NAME)) {
          missing.push({
            displayName: OUTLOOK_NOTE_CATEGORY_NAME,
            color: Office.MailboxEnums.CategoryColor.Preset3, // Yellow
          });
        }

        if (!existingNames.has(OUTLOOK_URGENT_CATEGORY_NAME)) {
          missing.push({
            displayName: OUTLOOK_URGENT_CATEGORY_NAME,
            color: Office.MailboxEnums.CategoryColor.Preset0, // Red
          });
        }

        if (missing.length === 0) {
          resolve();
          return;
        }

        masterCategories.addAsync(missing, () => resolve());
      });
    });
  }
}

export const officeContextService = new OfficeContextService();
