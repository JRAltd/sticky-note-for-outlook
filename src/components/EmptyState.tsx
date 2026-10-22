import React from "react";
import { Plus, Sparkles, StickyNote as StickyNoteIcon } from "lucide-react";
import { CreateNoteInput } from "../types/note";

interface EmptyStateProps {
  hasFilters: boolean;
  onClearFilters: () => void;
  onOpenNewNote: () => void;
  onQuickTemplateCreate: (template: CreateNoteInput) => void;
}

export const EmptyState: React.FC<EmptyStateProps> = ({
  hasFilters,
  onClearFilters,
  onOpenNewNote,
  onQuickTemplateCreate,
}) => {
  if (hasFilters) {
    return (
      <div className="empty-state-box">
        <StickyNoteIcon size={32} className="empty-icon" />
        <h3>No matching sticky notes</h3>
        <p>Try adjusting your search query, color filter, or selected tag.</p>
        <button
          type="button"
          className="btn-secondary-sm"
          onClick={onClearFilters}
        >
          Clear All Filters
        </button>
      </div>
    );
  }

  return (
    <div className="empty-state-box">
      <StickyNoteIcon size={36} className="empty-icon" />
      <h3>No sticky notes on this email yet</h3>
      <p>
        Attach private, timestamped notes, action checklists, or follow-up
        reminders directly to this email.
      </p>

      <button type="button" className="btn-primary" onClick={onOpenNewNote}>
        <Plus size={15} /> Create Sticky Note
      </button>

      <div className="quick-templates-section">
        <span className="quick-templates-title">
          <Sparkles size={12} /> 1-Click Quick Templates
        </span>
        <div className="quick-template-buttons">
          <button
            type="button"
            className="template-chip"
            onClick={() =>
              onQuickTemplateCreate({
                title: "Follow-up Required",
                content: "Follow up with sender regarding open action items.",
                color: "yellow",
                priority: "high",
                tags: ["follow-up", "action-required"],
                checklist: [
                  {
                    id: `chk-${Date.now()}-1`,
                    text: "Review attachments & reply",
                    completed: false,
                  },
                ],
                isPinned: true,
                scope: "message",
              })
            }
          >
            + Follow-up Checklist
          </button>

          <button
            type="button"
            className="template-chip"
            onClick={() =>
              onQuickTemplateCreate({
                title: "Waiting on Stakeholder Approval",
                content:
                  "Holding reply until internal sign-off is confirmed.",
                color: "amber",
                priority: "medium",
                tags: ["waiting"],
                checklist: [],
                isPinned: false,
                scope: "conversation",
              })
            }
          >
            + Waiting on Approval
          </button>

          <button
            type="button"
            className="template-chip"
            onClick={() =>
              onQuickTemplateCreate({
                title: "Key Call / Meeting Takeaways",
                content: "Summary of key decisions discussed regarding this thread:",
                color: "blue",
                priority: "medium",
                tags: ["meeting-notes", "decision"],
                checklist: [],
                isPinned: false,
                scope: "conversation",
              })
            }
          >
            + Meeting Summary
          </button>
        </div>
      </div>
    </div>
  );
};
