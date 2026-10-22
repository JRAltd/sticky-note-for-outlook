export type NoteColor =
  | "yellow"
  | "amber"
  | "green"
  | "blue"
  | "pink"
  | "purple";

export type NotePriority = "low" | "medium" | "high" | "urgent";

export type NoteScope = "message" | "conversation";

export type StorageBackendMode = "hybrid" | "office_only" | "graph_only";

export type ThemeMode = "system" | "light" | "dark";

export interface NoteChecklistItem {
  id: string;
  text: string;
  completed: boolean;
}

export interface StickyNote {
  id: string;
  /** Exchange REST / Immutable ID or Item ID of the email */
  itemId: string;
  /** Conversation ID for thread-level notes */
  conversationId: string;
  /** Whether this note is scoped to the specific email or entire conversation thread */
  scope: NoteScope;
  /** Email subject snapshot (for cross-email search & index display) */
  emailSubject: string;
  /** Email sender display name & address snapshot */
  emailSender: string;
  /** Main sticky note title (optional) */
  title: string;
  /** Sticky note body content */
  content: string;
  /** Visual color of the sticky note */
  color: NoteColor;
  /** Priority level */
  priority: NotePriority;
  /** Custom or preset tags */
  tags: string[];
  /** Optional interactive checklist items inside the note */
  checklist: NoteChecklistItem[];
  /** Whether the note is pinned to the top of the email's note list */
  isPinned: boolean;
  /** Optional follow-up reminder ISO timestamp */
  reminderAt?: string;
  /** Author display name or email */
  author: string;
  /** ISO 8601 creation timestamp */
  createdAt: string;
  /** ISO 8601 last modified timestamp */
  updatedAt: string;
  /** Sync status across backends */
  syncStatus: "synced_office" | "synced_graph" | "synced_both" | "local_only";
}

export interface EmailContextInfo {
  itemId: string;
  restId: string;
  conversationId: string;
  subject: string;
  senderName: string;
  senderEmail: string;
  receivedDateTime: string;
  itemType: "message" | "appointment";
  mode: "read" | "compose";
  hostEnvironment: "outlook_desktop" | "outlook_web" | "browser_preview";
  userEmail: string;
  userDisplayName: string;
}

export interface NoteFilterState {
  searchQuery: string;
  selectedColor: NoteColor | "all";
  selectedPriority: NotePriority | "all";
  selectedTag: string | "all";
  sortBy: "updated_desc" | "created_desc" | "created_asc" | "priority_desc";
  viewScope: "current_email" | "all_emails";
  onlyPinned: boolean;
}

export interface AzureAuthConfig {
  clientId: string;
  tenantId: string;
  redirectUri: string;
  graphExtensionName: string;
  storageMode: StorageBackendMode;
}

export interface GraphOpenExtensionPayload {
  "@odata.type": "microsoft.graph.openTypeExtension";
  extensionName: string;
  schemaVersion: string;
  lastSyncedAt: string;
  notesJson: string;
}

export interface CreateNoteInput {
  title: string;
  content: string;
  color: NoteColor;
  priority: NotePriority;
  tags: string[];
  checklist: NoteChecklistItem[];
  isPinned: boolean;
  scope: NoteScope;
  reminderAt?: string;
}
