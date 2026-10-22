import React from "react";
import { Mail, Tag, User } from "lucide-react";
import {
  OUTLOOK_NOTE_CATEGORY_NAME,
  OUTLOOK_URGENT_CATEGORY_NAME,
} from "../config/constants";
import { EmailContextInfo, StickyNote } from "../types/note";

interface EmailContextBannerProps {
  emailContext: EmailContextInfo;
  emailNotes: StickyNote[];
  allEmailsNotes: StickyNote[];
  isInOutlook: boolean;
  demoEmails: EmailContextInfo[];
  onSwitchDemoEmail: (index: number) => void;
}

export const EmailContextBanner: React.FC<EmailContextBannerProps> = ({
  emailContext,
  emailNotes,
  allEmailsNotes,
  isInOutlook,
  demoEmails,
  onSwitchDemoEmail,
}) => {
  const hasNotes = emailNotes.length > 0;
  const hasUrgent = emailNotes.some((n) => n.priority === "urgent");

  return (
    <section className="email-context-banner">
      {!isInOutlook && (
        <div className="demo-switcher-bar">
          <span className="demo-label">Sandbox Email Switcher:</span>
          <div className="demo-pills">
            {demoEmails.map((item, idx) => {
              const itemNoteCount = allEmailsNotes.filter(
                (n) =>
                  n.itemId === item.itemId ||
                  (n.scope === "conversation" &&
                    n.conversationId === item.conversationId)
              ).length;
              return (
                <button
                  key={item.itemId}
                  type="button"
                  className={`demo-pill ${
                    item.itemId === emailContext.itemId ? "active" : ""
                  }`}
                  onClick={() => onSwitchDemoEmail(idx)}
                >
                  Email #{idx + 1}
                  {itemNoteCount > 0 ? ` (📝 ${itemNoteCount})` : ""}
                </button>
              );
            })}
          </div>
        </div>
      )}

      <div className="email-meta-row">
        <Mail size={15} className="email-meta-icon" />
        <div className="email-meta-text">
          <div className="email-subject-row">
            <div className="email-subject" title={emailContext.subject}>
              {emailContext.subject}
            </div>
            {hasNotes && (
              <div className="email-category-markers">
                <span
                  className="outlook-category-pill cat-yellow"
                  title="Outlook Inbox Category Tag automatically applied to this email"
                >
                  <Tag size={10} />
                  {OUTLOOK_NOTE_CATEGORY_NAME}
                </span>
                {hasUrgent && (
                  <span
                    className="outlook-category-pill cat-red"
                    title="Urgent Outlook Inbox Category Tag automatically applied"
                  >
                    {OUTLOOK_URGENT_CATEGORY_NAME}
                  </span>
                )}
              </div>
            )}
          </div>
          <div className="email-sender">
            <User size={11} />
            <span>
              {emailContext.senderName} ({emailContext.senderEmail})
            </span>
          </div>
        </div>
      </div>
    </section>
  );
};
