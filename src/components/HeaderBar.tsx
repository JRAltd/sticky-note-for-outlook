import React from "react";
import {
  Cloud,
  Database,
  Moon,
  Plus,
  RefreshCw,
  Settings,
  Sun,
} from "lucide-react";

interface HeaderBarProps {
  currentEmailNoteCount: number;
  allEmailsNoteCount: number;
  viewScope: "current_email" | "all_emails";
  onChangeViewScope: (scope: "current_email" | "all_emails") => void;
  isDark: boolean;
  onToggleTheme: () => void;
  isSyncing: boolean;
  graphSynced: boolean;
  isInOutlook: boolean;
  onSyncNow: () => void;
  onOpenSettings: () => void;
  onOpenNewNote: () => void;
}

export const HeaderBar: React.FC<HeaderBarProps> = ({
  currentEmailNoteCount,
  allEmailsNoteCount,
  viewScope,
  onChangeViewScope,
  isDark,
  onToggleTheme,
  isSyncing,
  graphSynced,
  isInOutlook,
  onSyncNow,
  onOpenSettings,
  onOpenNewNote,
}) => {
  return (
    <header className="app-header">
      <div className="header-top-row">
        <div className="brand-group">
          <img
            src="./assets/icon-64.png"
            alt="Email Sticky Notes"
            className="brand-icon-img"
            width={32}
            height={32}
          />
          <div>
            <h1 className="brand-title">Email Sticky Notes</h1>
            <div className="storage-pill-row">
              <span
                className={`status-pill ${
                  graphSynced
                    ? "status-graph"
                    : isInOutlook
                    ? "status-office"
                    : "status-preview"
                }`}
              >
                {graphSynced ? (
                  <>
                    <Cloud size={11} /> Graph + Office.js
                  </>
                ) : isInOutlook ? (
                  <>
                    <Database size={11} /> Office.js Native
                  </>
                ) : (
                  <>
                    <Database size={11} /> Local Preview
                  </>
                )}
              </span>
            </div>
          </div>
        </div>

        <div className="header-actions">
          <button
            type="button"
            className="btn-primary-compact"
            onClick={onOpenNewNote}
            title="Add new sticky note (Ctrl+N)"
          >
            <Plus size={15} />
            <span>New Note</span>
          </button>

          <button
            type="button"
            className="icon-btn"
            onClick={onSyncNow}
            disabled={isSyncing}
            title="Refresh & Sync Notes"
          >
            <RefreshCw size={15} className={isSyncing ? "spin-anim" : ""} />
          </button>

          <button
            type="button"
            className="icon-btn"
            onClick={onToggleTheme}
            title={isDark ? "Switch to Light Mode" : "Switch to Dark Mode"}
          >
            {isDark ? <Sun size={15} /> : <Moon size={15} />}
          </button>

          <button
            type="button"
            className="icon-btn"
            onClick={onOpenSettings}
            title="Storage, Azure AD & Export Settings"
          >
            <Settings size={15} />
          </button>
        </div>
      </div>

      <div className="scope-tabs" role="tablist">
        <button
          type="button"
          role="tab"
          aria-selected={viewScope === "current_email"}
          className={`scope-tab ${
            viewScope === "current_email" ? "active" : ""
          }`}
          onClick={() => onChangeViewScope("current_email")}
        >
          <span>This Email</span>
          <span className="tab-badge">{currentEmailNoteCount}</span>
        </button>

        <button
          type="button"
          role="tab"
          aria-selected={viewScope === "all_emails"}
          className={`scope-tab ${viewScope === "all_emails" ? "active" : ""}`}
          onClick={() => onChangeViewScope("all_emails")}
        >
          <span>All Annotated Emails</span>
          <span className="tab-badge">{allEmailsNoteCount}</span>
        </button>
      </div>
    </header>
  );
};
