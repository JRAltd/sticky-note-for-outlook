import { AzureAuthConfig, EmailContextInfo, NoteColor, NotePriority } from "../types/note";

export const CUSTOM_PROP_NOTES_KEY = "jraltd_email_sticky_notes_v2";
export const ROAMING_INDEX_KEY = "jraltd_email_notes_global_index_v2";
export const LOCAL_STORAGE_NOTES_PREFIX = "outlook_sticky_notes_item_";
export const LOCAL_STORAGE_INDEX_KEY = "outlook_sticky_notes_all_index";
export const LOCAL_STORAGE_AUTH_CONFIG_KEY = "outlook_sticky_notes_auth_config";
export const LOCAL_STORAGE_THEME_KEY = "outlook_sticky_notes_theme";

export const OUTLOOK_NOTE_CATEGORY_NAME = "📝 Sticky Note";
export const OUTLOOK_URGENT_CATEGORY_NAME = "🔥 Urgent Note";

export const DEFAULT_GRAPH_EXTENSION_NAME = "com.jraltd.emailNotes";

export const DEFAULT_AUTH_CONFIG: AzureAuthConfig = {
  clientId: import.meta.env?.VITE_AZURE_CLIENT_ID || "",
  tenantId: import.meta.env?.VITE_AZURE_TENANT_ID || "common",
  redirectUri:
    import.meta.env?.VITE_AZURE_REDIRECT_URI ||
    (typeof window !== "undefined" ? `${window.location.origin}/taskpane.html` : "https://localhost:3000/taskpane.html"),
  graphExtensionName:
    import.meta.env?.VITE_GRAPH_EXTENSION_NAME || DEFAULT_GRAPH_EXTENSION_NAME,
  storageMode: "hybrid",
};

export interface StickyColorTheme {
  id: NoteColor;
  label: string;
  swatchHex: string;
  lightBg: string;
  lightHeader: string;
  lightBorder: string;
  lightText: string;
  darkBg: string;
  darkHeader: string;
  darkBorder: string;
  darkText: string;
}

export const STICKY_COLOR_PALETTE: Record<NoteColor, StickyColorTheme> = {
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

export const PRIORITY_CONFIG: Record<
  NotePriority,
  { label: string; rank: number; badgeColor: string }
> = {
  low: { label: "Low", rank: 1, badgeColor: "#64748b" },
  medium: { label: "Medium", rank: 2, badgeColor: "#0284c7" },
  high: { label: "High", rank: 3, badgeColor: "#d97706" },
  urgent: { label: "Urgent", rank: 4, badgeColor: "#dc2626" },
};

export const PRESET_TAGS: string[] = [
  "follow-up",
  "waiting",
  "action-required",
  "decision",
  "client",
  "billing",
  "important",
  "meeting-notes",
];

/**
 * Realistic sample emails used when previewing the Add-in in a normal browser
 * outside Outlook so all features can be tested immediately.
 */
export const DEMO_SAMPLE_EMAILS: EmailContextInfo[] = [
  {
    itemId: "demo-msg-001-q4-contract",
    restId: "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9r7V8cBwAiIsqMbYjsT5e-T7KzVjKTAAAAAAEMAAAiIsqMbYjsT5e-T7KzVjKTAAADK912AAA=",
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
    restId: "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9r7V8cBwAiIsqMbYjsT5e-T7KzVjKTAAAAAAEMAAAiIsqMbYjsT5e-T7KzVjKTAAADK913AAA=",
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
    restId: "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9r7V8cBwAiIsqMbYjsT5e-T7KzVjKTAAAAAAEMAAAiIsqMbYjsT5e-T7KzVjKTAAADK914AAA=",
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
