import {
  CUSTOM_PROP_NOTES_KEY,
  LOCAL_STORAGE_INDEX_KEY,
  LOCAL_STORAGE_NOTES_PREFIX,
  ROAMING_INDEX_KEY,
} from "../../config/constants";
import { StickyNote } from "../../types/note";
import { officeContextService } from "../office/officeContextService";

export class OfficeCustomPropsAdapter {
  /**
   * Loads notes stored on the active Outlook email item via Office.js CustomProperties,
   * falling back to localStorage cache if outside Outlook or if item customProps are unavailable.
   */
  public async loadNotesForItem(itemId: string): Promise<StickyNote[]> {
    const localNotes = this.loadFromLocalStorage(itemId);

    if (
      !officeContextService.isInOutlook() ||
      !Office.context?.mailbox?.item?.loadCustomPropertiesAsync
    ) {
      return localNotes;
    }

    return new Promise<StickyNote[]>((resolve) => {
      Office.context.mailbox.item!.loadCustomPropertiesAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          resolve(localNotes);
          return;
        }

        const customProps = result.value;
        const rawJson = customProps.get(CUSTOM_PROP_NOTES_KEY);
        if (!rawJson) {
          resolve(localNotes);
          return;
        }

        try {
          const parsed = JSON.parse(rawJson) as StickyNote[];
          if (Array.isArray(parsed)) {
            this.saveToLocalStorage(itemId, parsed);
            resolve(parsed);
            return;
          }
        } catch (err) {
          console.warn("Failed to parse CustomProperties notes JSON:", err);
        }

        resolve(localNotes);
      });
    });
  }

  /**
   * Persists notes onto the active Outlook email item via Office.js CustomProperties
   * AND updates the cross-email RoamingSettings / localStorage index.
   */
  public async saveNotesForItem(
    itemId: string,
    notes: StickyNote[]
  ): Promise<boolean> {
    this.saveToLocalStorage(itemId, notes);
    await this.updateGlobalIndex(itemId, notes);

    if (
      !officeContextService.isInOutlook() ||
      !Office.context?.mailbox?.item?.loadCustomPropertiesAsync
    ) {
      return true;
    }

    return new Promise<boolean>((resolve) => {
      Office.context.mailbox.item!.loadCustomPropertiesAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          resolve(false);
          return;
        }

        const customProps = result.value;
        customProps.set(CUSTOM_PROP_NOTES_KEY, JSON.stringify(notes));
        customProps.saveAsync((saveRes) => {
          resolve(saveRes.status === Office.AsyncResultStatus.Succeeded);
        });
      });
    });
  }

  /**
   * Returns all known notes across all annotated emails from RoamingSettings + localStorage.
   */
  public getAllIndexedNotes(): StickyNote[] {
    const map = new Map<string, StickyNote>();

    // 1. Read from Office RoamingSettings if available
    try {
      if (
        officeContextService.isInOutlook() &&
        Office.context?.roamingSettings
      ) {
        const rawRoaming = Office.context.roamingSettings.get(ROAMING_INDEX_KEY);
        if (rawRoaming) {
          const parsed =
            typeof rawRoaming === "string" ? JSON.parse(rawRoaming) : rawRoaming;
          if (Array.isArray(parsed)) {
            parsed.forEach((n: StickyNote) => {
              if (n && n.id) map.set(n.id, n);
            });
          }
        }
      }
    } catch {
      // Ignore roaming settings parse errors
    }

    // 2. Merge with localStorage index
    try {
      const rawLocal = localStorage.getItem(LOCAL_STORAGE_INDEX_KEY);
      if (rawLocal) {
        const parsed = JSON.parse(rawLocal) as StickyNote[];
        if (Array.isArray(parsed)) {
          parsed.forEach((n) => {
            const existing = map.get(n.id);
            if (!existing || n.updatedAt >= existing.updatedAt) {
              map.set(n.id, n);
            }
          });
        }
      }
    } catch {
      // Ignore localStorage parse errors
    }

    return Array.from(map.values()).sort((a, b) =>
      b.updatedAt.localeCompare(a.updatedAt)
    );
  }

  public async replaceGlobalIndex(allNotes: StickyNote[]): Promise<void> {
    try {
      localStorage.setItem(LOCAL_STORAGE_INDEX_KEY, JSON.stringify(allNotes));
    } catch {
      // Ignore quota errors
    }

    if (
      officeContextService.isInOutlook() &&
      Office.context?.roamingSettings
    ) {
      await new Promise<void>((resolve) => {
        try {
          // Keep most recent 150 notes in RoamingSettings (32KB limit safeguard)
          const compactIndex = allNotes.slice(0, 150);
          Office.context.roamingSettings.set(
            ROAMING_INDEX_KEY,
            JSON.stringify(compactIndex)
          );
          Office.context.roamingSettings.saveAsync(() => resolve());
        } catch {
          resolve();
        }
      });
    }
  }

  private async updateGlobalIndex(
    itemId: string,
    itemNotes: StickyNote[]
  ): Promise<void> {
    const existingAll = this.getAllIndexedNotes().filter(
      (n) => n.itemId !== itemId
    );
    const combined = [...itemNotes, ...existingAll].sort((a, b) =>
      b.updatedAt.localeCompare(a.updatedAt)
    );
    await this.replaceGlobalIndex(combined);
  }

  private loadFromLocalStorage(itemId: string): StickyNote[] {
    try {
      const raw = localStorage.getItem(`${LOCAL_STORAGE_NOTES_PREFIX}${itemId}`);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {
      // Ignore error
    }
    return [];
  }

  public saveToLocalStorage(itemId: string, notes: StickyNote[]): void {
    try {
      localStorage.setItem(
        `${LOCAL_STORAGE_NOTES_PREFIX}${itemId}`,
        JSON.stringify(notes)
      );
    } catch {
      // Ignore quota errors
    }
  }
}

export const officeCustomPropsAdapter = new OfficeCustomPropsAdapter();
