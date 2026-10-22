import React, { useState } from "react";
import {
  CheckCircle2,
  Cloud,
  Download,
  FileText,
  KeyRound,
  LogOut,
  Save,
  Upload,
  X,
} from "lucide-react";
import { msalAuthService } from "../services/auth/msalAuthService";
import { storageManager } from "../services/storage/storageManager";
import {
  AzureAuthConfig,
  EmailContextInfo,
  StickyNote,
  StorageBackendMode,
} from "../types/note";

interface SettingsModalProps {
  emailContext: EmailContextInfo;
  allNotes: StickyNote[];
  onClose: () => void;
  onConfigSaved: () => void;
  onNotesImported: (count: number) => void;
}

export const SettingsModal: React.FC<SettingsModalProps> = ({
  emailContext,
  allNotes,
  onClose,
  onConfigSaved,
  onNotesImported,
}) => {
  const initialCfg = msalAuthService.getConfig();
  const [storageMode, setStorageMode] = useState<StorageBackendMode>(
    initialCfg.storageMode
  );
  const [clientId, setClientId] = useState<string>(initialCfg.clientId);
  const [tenantId, setTenantId] = useState<string>(initialCfg.tenantId);
  const [graphExtensionName, setGraphExtensionName] = useState<string>(
    initialCfg.graphExtensionName
  );
  const [authStatusText, setAuthStatusText] = useState<string>("");
  const [isTestingAuth, setIsTestingAuth] = useState<boolean>(false);

  const connectedAccount = msalAuthService.getConnectedAccount();

  const handleSaveSettings = (e: React.FormEvent) => {
    e.preventDefault();
    const updated: Partial<AzureAuthConfig> = {
      storageMode,
      clientId: clientId.trim(),
      tenantId: tenantId.trim() || "common",
      graphExtensionName: graphExtensionName.trim() || "com.jraltd.emailNotes",
    };
    msalAuthService.saveConfig(updated);
    onConfigSaved();
    onClose();
  };

  const handleTestGraphSignIn = async () => {
    msalAuthService.saveConfig({
      storageMode,
      clientId: clientId.trim(),
      tenantId: tenantId.trim() || "common",
      graphExtensionName: graphExtensionName.trim() || "com.jraltd.emailNotes",
    });

    if (!msalAuthService.isGraphConfigured()) {
      setAuthStatusText(
        "Please enter a valid Azure AD Application (client) ID first."
      );
      return;
    }

    setIsTestingAuth(true);
    setAuthStatusText("Connecting to Microsoft Entra ID / Graph...");
    try {
      const token = await msalAuthService.getGraphAccessToken(true);
      if (token) {
        const acct = msalAuthService.getConnectedAccount();
        setAuthStatusText(
          `Connected to Microsoft Graph${
            acct?.username ? ` as ${acct.username}` : ""
          }`
        );
      } else {
        setAuthStatusText("Could not acquire Microsoft Graph token.");
      }
    } catch (err) {
      setAuthStatusText(
        err instanceof Error ? err.message : "Authentication failed."
      );
    } finally {
      setIsTestingAuth(false);
    }
  };

  const handleSignOut = async () => {
    await msalAuthService.signOut();
    setAuthStatusText("Signed out of Microsoft Graph.");
  };

  const triggerDownload = (
    filename: string,
    content: string,
    mimeType: string
  ) => {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleExportJson = () => {
    const json = storageManager.exportNotesAsJson(allNotes);
    triggerDownload(
      `outlook-sticky-notes-${new Date().toISOString().slice(0, 10)}.json`,
      json,
      "application/json"
    );
  };

  const handleExportMarkdown = () => {
    const md = storageManager.exportNotesAsMarkdown(allNotes);
    triggerDownload(
      `outlook-sticky-notes-${new Date().toISOString().slice(0, 10)}.md`,
      md,
      "text/markdown"
    );
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const count = await storageManager.importNotesFromJson(
        text,
        emailContext
      );
      onNotesImported(count);
    } catch {
      setAuthStatusText("Invalid JSON backup file.");
    }
  };

  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true">
      <div className="settings-modal">
        <div className="modal-header">
          <h2>Storage, Graph & Backup Settings</h2>
          <button type="button" className="icon-btn" onClick={onClose}>
            <X size={16} />
          </button>
        </div>

        <form className="modal-body" onSubmit={handleSaveSettings}>
          <div className="settings-group">
            <label className="settings-group-title">
              Persistence Engine Mode
            </label>
            <div className="radio-cards">
              <label
                className={`radio-card ${
                  storageMode === "hybrid" ? "selected" : ""
                }`}
              >
                <input
                  type="radio"
                  name="storageMode"
                  checked={storageMode === "hybrid"}
                  onChange={() => setStorageMode("hybrid")}
                />
                <div>
                  <strong>Hybrid (Recommended)</strong>
                  <p>
                    Uses Office.js CustomProperties + RoamingSettings
                    immediately, and syncs with Microsoft Graph Open Extensions
                    when Azure AD is configured.
                  </p>
                </div>
              </label>

              <label
                className={`radio-card ${
                  storageMode === "office_only" ? "selected" : ""
                }`}
              >
                <input
                  type="radio"
                  name="storageMode"
                  checked={storageMode === "office_only"}
                  onChange={() => setStorageMode("office_only")}
                />
                <div>
                  <strong>Office.js CustomProperties Only (Zero Setup)</strong>
                  <p>
                    Stores notes directly inside Exchange item properties &
                    roaming settings with zero Azure AD registration needed.
                  </p>
                </div>
              </label>

              <label
                className={`radio-card ${
                  storageMode === "graph_only" ? "selected" : ""
                }`}
              >
                <input
                  type="radio"
                  name="storageMode"
                  checked={storageMode === "graph_only"}
                  onChange={() => setStorageMode("graph_only")}
                />
                <div>
                  <strong>Microsoft Graph Open Extensions Only</strong>
                  <p>
                    Persists notes exclusively via Graph API{" "}
                    <code>/me/messages/&#123;id&#125;/extensions</code>.
                  </p>
                </div>
              </label>
            </div>
          </div>

          <div className="settings-group">
            <label className="settings-group-title">
              <KeyRound size={14} /> Microsoft Entra ID (Azure AD) Configuration
            </label>
            <p className="settings-hint">
              Optional in Hybrid/Office.js modes. Required for Microsoft Graph
              Open Extensions (<code>Mail.ReadWrite</code>).
            </p>

            <div className="settings-field">
              <label>Application (Client) ID</label>
              <input
                type="text"
                className="compact-text-input"
                placeholder="e.g. 00000000-0000-0000-0000-000000000000"
                value={clientId}
                onChange={(e) => setClientId(e.target.value)}
              />
            </div>

            <div className="settings-row-2col">
              <div className="settings-field">
                <label>Directory (Tenant) ID</label>
                <input
                  type="text"
                  className="compact-text-input"
                  placeholder="common or tenant GUID"
                  value={tenantId}
                  onChange={(e) => setTenantId(e.target.value)}
                />
              </div>

              <div className="settings-field">
                <label>Open Extension Name</label>
                <input
                  type="text"
                  className="compact-text-input"
                  placeholder="com.jraltd.emailNotes"
                  value={graphExtensionName}
                  onChange={(e) => setGraphExtensionName(e.target.value)}
                />
              </div>
            </div>

            <div className="auth-test-row">
              <button
                type="button"
                className="btn-secondary-sm"
                onClick={handleTestGraphSignIn}
                disabled={isTestingAuth}
              >
                <Cloud size={13} />
                <span>
                  {connectedAccount
                    ? "Reconnect Microsoft Graph"
                    : "Sign In & Test Graph"}
                </span>
              </button>

              {connectedAccount && (
                <button
                  type="button"
                  className="btn-secondary-sm"
                  onClick={handleSignOut}
                >
                  <LogOut size={13} /> Sign Out ({connectedAccount.username})
                </button>
              )}
            </div>

            {authStatusText && (
              <div className="auth-status-banner">
                <CheckCircle2 size={13} />
                <span>{authStatusText}</span>
              </div>
            )}
          </div>

          <div className="settings-group">
            <label className="settings-group-title">
              Backup, Export & Import ({allNotes.length} total notes)
            </label>
            <div className="backup-actions-row">
              <button
                type="button"
                className="btn-secondary-sm"
                onClick={handleExportJson}
              >
                <Download size={13} /> Export JSON
              </button>
              <button
                type="button"
                className="btn-secondary-sm"
                onClick={handleExportMarkdown}
              >
                <FileText size={13} /> Export Markdown
              </button>
              <label className="btn-secondary-sm file-upload-label">
                <Upload size={13} /> Import JSON
                <input
                  type="file"
                  accept=".json,application/json"
                  onChange={handleImportFile}
                  hidden
                />
              </label>
            </div>
          </div>

          <div className="modal-footer">
            <button type="button" className="btn-ghost" onClick={onClose}>
              Cancel
            </button>
            <button type="submit" className="btn-primary">
              <Save size={14} /> Save Configuration
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
