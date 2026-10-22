import React, { useState } from "react";
import {
  Bell,
  Check,
  Clock,
  Copy,
  Edit3,
  Mail,
  MessageSquare,
  Pin,
  Trash2,
} from "lucide-react";
import { PRIORITY_CONFIG, STICKY_COLOR_PALETTE } from "../config/constants";
import { StickyNote } from "../types/note";

interface StickyNoteCardProps {
  note: StickyNote;
  isDark: boolean;
  showEmailHeader: boolean;
  onTogglePin: (noteId: string) => void;
  onEdit: (note: StickyNote) => void;
  onDelete: (noteId: string) => void;
  onToggleChecklistItem: (noteId: string, checklistItemId: string) => void;
  onSelectTag?: (tag: string) => void;
}

function formatTimestamp(isoString: string): string {
  try {
    const date = new Date(isoString);
    const diffMs = Date.now() - date.getTime();
    const diffMins = Math.floor(diffMs / 60000);
    if (diffMins < 1) return "Just now";
    if (diffMins < 60) return `${diffMins}m ago`;
    const diffHours = Math.floor(diffMins / 60);
    if (diffHours < 24) return `${diffHours}h ago`;
    return date.toLocaleDateString(undefined, {
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  } catch {
    return isoString;
  }
}

export const StickyNoteCard: React.FC<StickyNoteCardProps> = ({
  note,
  isDark,
  showEmailHeader,
  onTogglePin,
  onEdit,
  onDelete,
  onToggleChecklistItem,
  onSelectTag,
}) => {
  const [copied, setCopied] = useState(false);

  const colorTheme =
    STICKY_COLOR_PALETTE[note.color] || STICKY_COLOR_PALETTE.yellow;
  const priorityInfo =
    PRIORITY_CONFIG[note.priority] || PRIORITY_CONFIG.medium;

  const cardStyle: React.CSSProperties = {
    backgroundColor: isDark ? colorTheme.darkBg : colorTheme.lightBg,
    borderColor: isDark ? colorTheme.darkBorder : colorTheme.lightBorder,
    color: isDark ? colorTheme.darkText : colorTheme.lightText,
  };

  const headerStyle: React.CSSProperties = {
    backgroundColor: isDark ? colorTheme.darkHeader : colorTheme.lightHeader,
  };

  const handleCopy = async () => {
    const text = [
      note.title ? `# ${note.title}` : "",
      note.content,
      ...note.checklist.map((c) => `- [${c.completed ? "x" : " "}] ${c.text}`),
    ]
      .filter(Boolean)
      .join("\n");

    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1800);
    } catch {
      // Clipboard fallback ignored
    }
  };

  const wasEdited =
    note.updatedAt &&
    note.createdAt &&
    Math.abs(
      new Date(note.updatedAt).getTime() - new Date(note.createdAt).getTime()
    ) > 3000;

  return (
    <article className="sticky-note-card" style={cardStyle}>
      <div className="sticky-note-top-strip" style={headerStyle}>
        <div className="sticky-badges-left">
          {note.isPinned && (
            <span className="pinned-indicator" title="Pinned Sticky Note">
              <Pin size={11} /> Pinned
            </span>
          )}
          <span
            className="priority-badge"
            style={{ backgroundColor: priorityInfo.badgeColor }}
          >
            {priorityInfo.label}
          </span>
          {note.scope === "conversation" && (
            <span
              className="scope-badge"
              title="Attached to entire email conversation thread"
            >
              <MessageSquare size={10} /> Thread
            </span>
          )}
        </div>

        <div className="sticky-card-actions">
          <button
            type="button"
            className={`card-icon-btn ${note.isPinned ? "pinned" : ""}`}
            onClick={() => onTogglePin(note.id)}
            title={note.isPinned ? "Unpin note" : "Pin note to top"}
          >
            <Pin size={13} />
          </button>
          <button
            type="button"
            className="card-icon-btn"
            onClick={handleCopy}
            title="Copy note text"
          >
            {copied ? <Check size={13} /> : <Copy size={13} />}
          </button>
          <button
            type="button"
            className="card-icon-btn"
            onClick={() => onEdit(note)}
            title="Edit sticky note"
          >
            <Edit3 size={13} />
          </button>
          <button
            type="button"
            className="card-icon-btn delete-btn"
            onClick={() => onDelete(note.id)}
            title="Delete sticky note"
          >
            <Trash2 size={13} />
          </button>
        </div>
      </div>

      <div className="sticky-note-body">
        {showEmailHeader && (
          <div className="cross-email-ref" title={note.emailSender}>
            <Mail size={12} />
            <span className="cross-email-subject">{note.emailSubject}</span>
          </div>
        )}

        {note.title && <h3 className="sticky-note-title">{note.title}</h3>}

        {note.content && (
          <p className="sticky-note-content">{note.content}</p>
        )}

        {note.checklist && note.checklist.length > 0 && (
          <div className="sticky-checklist-box">
            {note.checklist.map((item) => (
              <label
                key={item.id}
                className={`sticky-checklist-row ${
                  item.completed ? "completed" : ""
                }`}
              >
                <input
                  type="checkbox"
                  checked={item.completed}
                  onChange={() => onToggleChecklistItem(note.id, item.id)}
                />
                <span>{item.text}</span>
              </label>
            ))}
          </div>
        )}

        {note.reminderAt && (
          <div className="sticky-reminder-pill">
            <Bell size={11} />
            <span>
              Reminder: {new Date(note.reminderAt).toLocaleString()}
            </span>
          </div>
        )}

        {note.tags && note.tags.length > 0 && (
          <div className="sticky-tags-row">
            {note.tags.map((tag) => (
              <button
                key={tag}
                type="button"
                className="sticky-tag-chip"
                onClick={() => onSelectTag?.(tag)}
              >
                #{tag}
              </button>
            ))}
          </div>
        )}

        <div className="sticky-note-footer">
          <span
            className="timestamp-meta"
            title={`Created: ${new Date(
              note.createdAt
            ).toLocaleString()}\nLast Updated: ${new Date(
              note.updatedAt
            ).toLocaleString()}`}
          >
            <Clock size={11} />
            <span>
              {formatTimestamp(note.updatedAt)}
              {wasEdited ? " (edited)" : ""}
            </span>
          </span>

          <span className="author-meta">{note.author}</span>
        </div>
      </div>
    </article>
  );
};
