import React, { useState } from "react";
import {
  CheckSquare,
  Pin,
  Plus,
  Save,
  Tag,
  Trash2,
  X,
} from "lucide-react";
import { PRESET_TAGS, STICKY_COLOR_PALETTE } from "../config/constants";
import {
  CreateNoteInput,
  NoteChecklistItem,
  NoteColor,
  NotePriority,
  NoteScope,
  StickyNote,
} from "../types/note";

interface NoteEditorProps {
  initialNote?: StickyNote;
  onSave: (input: CreateNoteInput) => Promise<void>;
  onCancel: () => void;
}

const COLOR_KEYS = Object.keys(STICKY_COLOR_PALETTE) as NoteColor[];

export const NoteEditor: React.FC<NoteEditorProps> = ({
  initialNote,
  onSave,
  onCancel,
}) => {
  const [title, setTitle] = useState<string>(initialNote?.title || "");
  const [content, setContent] = useState<string>(initialNote?.content || "");
  const [color, setColor] = useState<NoteColor>(
    initialNote?.color || "yellow"
  );
  const [priority, setPriority] = useState<NotePriority>(
    initialNote?.priority || "medium"
  );
  const [scope, setScope] = useState<NoteScope>(
    initialNote?.scope || "message"
  );
  const [isPinned, setIsPinned] = useState<boolean>(
    initialNote?.isPinned ?? false
  );
  const [tags, setTags] = useState<string[]>(initialNote?.tags || []);
  const [customTagInput, setCustomTagInput] = useState<string>("");
  const [checklist, setChecklist] = useState<NoteChecklistItem[]>(
    initialNote?.checklist || []
  );
  const [newChecklistText, setNewChecklistText] = useState<string>("");
  const [reminderAt, setReminderAt] = useState<string>(
    initialNote?.reminderAt ? initialNote.reminderAt.slice(0, 16) : ""
  );
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);

  const toggleTag = (tag: string) => {
    const clean = tag.trim().toLowerCase().replace(/^#/, "");
    if (!clean) return;
    setTags((prev) =>
      prev.includes(clean) ? prev.filter((t) => t !== clean) : [...prev, clean]
    );
  };

  const handleAddCustomTag = () => {
    const clean = customTagInput.trim().toLowerCase().replace(/^#/, "");
    if (!clean) return;
    if (!tags.includes(clean)) {
      setTags((prev) => [...prev, clean]);
    }
    setCustomTagInput("");
  };

  const handleAddChecklistItem = () => {
    const text = newChecklistText.trim();
    if (!text) return;
    setChecklist((prev) => [
      ...prev,
      {
        id: `chk-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        text,
        completed: false,
      },
    ]);
    setNewChecklistText("");
  };

  const handleRemoveChecklistItem = (id: string) => {
    setChecklist((prev) => prev.filter((item) => item.id !== id));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!title.trim() && !content.trim() && checklist.length === 0) {
      return;
    }

    setIsSubmitting(true);
    try {
      await onSave({
        title,
        content,
        color,
        priority,
        tags,
        checklist,
        isPinned,
        scope,
        reminderAt: reminderAt ? new Date(reminderAt).toISOString() : undefined,
      });
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <form className="note-editor-card" onSubmit={handleSubmit}>
      <div className="editor-header">
        <span className="editor-heading">
          {initialNote ? "Edit Sticky Note" : "New Sticky Note"}
        </span>
        <div className="editor-header-controls">
          <button
            type="button"
            className={`icon-toggle-btn ${isPinned ? "active" : ""}`}
            onClick={() => setIsPinned((p) => !p)}
            title="Pin to top of email"
          >
            <Pin size={14} />
          </button>
          <button
            type="button"
            className="icon-toggle-btn"
            onClick={onCancel}
            title="Cancel"
          >
            <X size={15} />
          </button>
        </div>
      </div>

      <input
        type="text"
        className="editor-title-input"
        placeholder="Sticky note title (e.g. Follow up with Legal)..."
        value={title}
        onChange={(e) => setTitle(e.target.value)}
        autoFocus
      />

      <textarea
        className="editor-body-textarea"
        placeholder="Write your timestamped sticky note here... (Ctrl+Enter to save)"
        rows={4}
        value={content}
        onChange={(e) => setContent(e.target.value)}
        onKeyDown={(e) => {
          if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
            handleSubmit(e);
          }
        }}
      />

      {/* Checklist Builder */}
      <div className="editor-section">
        <div className="editor-section-label">
          <CheckSquare size={13} />
          <span>Action Checklist</span>
        </div>
        {checklist.length > 0 && (
          <ul className="editor-checklist">
            {checklist.map((item) => (
              <li key={item.id} className="editor-checklist-item">
                <input
                  type="checkbox"
                  checked={item.completed}
                  onChange={() =>
                    setChecklist((prev) =>
                      prev.map((c) =>
                        c.id === item.id ? { ...c, completed: !c.completed } : c
                      )
                    )
                  }
                />
                <span>{item.text}</span>
                <button
                  type="button"
                  className="remove-chk-btn"
                  onClick={() => handleRemoveChecklistItem(item.id)}
                >
                  <Trash2 size={12} />
                </button>
              </li>
            ))}
          </ul>
        )}
        <div className="inline-add-row">
          <input
            type="text"
            className="compact-text-input"
            placeholder="Add a checklist item and press Enter..."
            value={newChecklistText}
            onChange={(e) => setNewChecklistText(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                handleAddChecklistItem();
              }
            }}
          />
          <button
            type="button"
            className="btn-secondary-sm"
            onClick={handleAddChecklistItem}
          >
            <Plus size={13} /> Add
          </button>
        </div>
      </div>

      {/* Color, Priority, Scope & Reminder */}
      <div className="editor-options-grid">
        <div className="editor-field">
          <label>Sticky Color</label>
          <div className="editor-color-swatches">
            {COLOR_KEYS.map((cKey) => {
              const theme = STICKY_COLOR_PALETTE[cKey];
              return (
                <button
                  key={cKey}
                  type="button"
                  className={`swatch-btn ${color === cKey ? "selected" : ""}`}
                  style={{ backgroundColor: theme.swatchHex }}
                  onClick={() => setColor(cKey)}
                  title={theme.label}
                />
              );
            })}
          </div>
        </div>

        <div className="editor-field">
          <label>Priority</label>
          <select
            className="compact-select"
            value={priority}
            onChange={(e) => setPriority(e.target.value as NotePriority)}
          >
            <option value="low">Low</option>
            <option value="medium">Medium</option>
            <option value="high">High</option>
            <option value="urgent">Urgent</option>
          </select>
        </div>

        <div className="editor-field">
          <label>Attach Scope</label>
          <select
            className="compact-select"
            value={scope}
            onChange={(e) => setScope(e.target.value as NoteScope)}
          >
            <option value="message">This Email Only</option>
            <option value="conversation">Entire Thread</option>
          </select>
        </div>

        <div className="editor-field">
          <label>Follow-up Reminder</label>
          <input
            type="datetime-local"
            className="compact-text-input"
            value={reminderAt}
            onChange={(e) => setReminderAt(e.target.value)}
          />
        </div>
      </div>

      {/* Tags */}
      <div className="editor-section">
        <div className="editor-section-label">
          <Tag size={13} />
          <span>Tags</span>
        </div>
        <div className="preset-tags-wrap">
          {PRESET_TAGS.map((tag) => (
            <button
              key={tag}
              type="button"
              className={`tag-pill-selectable ${
                tags.includes(tag) ? "selected" : ""
              }`}
              onClick={() => toggleTag(tag)}
            >
              #{tag}
            </button>
          ))}
          {tags
            .filter((t) => !PRESET_TAGS.includes(t))
            .map((customTag) => (
              <button
                key={customTag}
                type="button"
                className="tag-pill-selectable selected"
                onClick={() => toggleTag(customTag)}
              >
                #{customTag} ×
              </button>
            ))}
        </div>
        <div className="inline-add-row">
          <input
            type="text"
            className="compact-text-input"
            placeholder="Custom tag (e.g. q4-audit)..."
            value={customTagInput}
            onChange={(e) => setCustomTagInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                handleAddCustomTag();
              }
            }}
          />
          <button
            type="button"
            className="btn-secondary-sm"
            onClick={handleAddCustomTag}
          >
            Add Tag
          </button>
        </div>
      </div>

      <div className="editor-footer">
        <button type="button" className="btn-ghost" onClick={onCancel}>
          Cancel
        </button>
        <button
          type="submit"
          className="btn-primary"
          disabled={
            isSubmitting ||
            (!title.trim() && !content.trim() && checklist.length === 0)
          }
        >
          <Save size={14} />
          <span>{initialNote ? "Update Note" : "Attach Sticky Note"}</span>
        </button>
      </div>
    </form>
  );
};
