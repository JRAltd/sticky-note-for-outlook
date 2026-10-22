import { GraphOpenExtensionPayload, StickyNote } from "../../types/note";
import { msalAuthService } from "../auth/msalAuthService";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export class GraphOpenExtensionAdapter {
  /**
   * Fetches Sticky Notes attached to an Exchange message via Microsoft Graph Open Extension:
   * GET /me/messages/{restId}/extensions/{extensionName}
   */
  public async fetchNotesFromGraph(
    restId: string,
    interactiveAuth = false
  ): Promise<StickyNote[] | null> {
    if (!restId || !msalAuthService.isGraphConfigured()) {
      return null;
    }

    const token = await msalAuthService.getGraphAccessToken(interactiveAuth);
    if (!token) {
      return null;
    }

    const extensionName =
      msalAuthService.getConfig().graphExtensionName || "com.jraltd.emailNotes";
    const encodedMessageId = encodeURIComponent(restId);
    const encodedExtId = encodeURIComponent(extensionName);

    const url = `${GRAPH_BASE_URL}/me/messages/${encodedMessageId}/extensions/${encodedExtId}`;

    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    });

    if (response.status === 404) {
      // Open extension has not been created on this message yet
      return [];
    }

    if (!response.ok) {
      throw new Error(
        `Graph Open Extension GET failed (${response.status}): ${response.statusText}`
      );
    }

    const data = (await response.json()) as Partial<GraphOpenExtensionPayload>;
    if (!data.notesJson) {
      return [];
    }

    try {
      const parsed = JSON.parse(data.notesJson) as StickyNote[];
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  /**
   * Creates or updates the Microsoft Graph Open Extension on an Exchange message:
   * PATCH /me/messages/{restId}/extensions/{extensionName}
   * (or POST /me/messages/{restId}/extensions if 404)
   */
  public async saveNotesToGraph(
    restId: string,
    notes: StickyNote[],
    interactiveAuth = false
  ): Promise<boolean> {
    if (!restId || !msalAuthService.isGraphConfigured()) {
      return false;
    }

    const token = await msalAuthService.getGraphAccessToken(interactiveAuth);
    if (!token) {
      return false;
    }

    const extensionName =
      msalAuthService.getConfig().graphExtensionName || "com.jraltd.emailNotes";
    const encodedMessageId = encodeURIComponent(restId);
    const encodedExtId = encodeURIComponent(extensionName);

    const payload: GraphOpenExtensionPayload = {
      "@odata.type": "microsoft.graph.openTypeExtension",
      extensionName,
      schemaVersion: "2.0",
      lastSyncedAt: new Date().toISOString(),
      notesJson: JSON.stringify(notes),
    };

    // Attempt PATCH first
    const patchUrl = `${GRAPH_BASE_URL}/me/messages/${encodedMessageId}/extensions/${encodedExtId}`;
    const patchRes = await fetch(patchUrl, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (patchRes.ok || patchRes.status === 204) {
      return true;
    }

    // If extension does not exist yet (404), create it via POST
    if (patchRes.status === 404) {
      const postUrl = `${GRAPH_BASE_URL}/me/messages/${encodedMessageId}/extensions`;
      const postRes = await fetch(postUrl, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });
      return postRes.ok;
    }

    return false;
  }
}

export const graphOpenExtensionAdapter = new GraphOpenExtensionAdapter();
