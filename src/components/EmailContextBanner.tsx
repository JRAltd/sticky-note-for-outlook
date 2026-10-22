import React, { useState } from "react";
import { Mail, Tag, User } from "lucide-react";
import {
  OUTLOOK_NOTE_CATEGORY_NAME,
  OUTLOOK_URGENT_CATEGORY_NAME,
} from "../config/constants";
import { officeContextService } from "../services/office/officeContextService";
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
  const [, forceUpdate] = useState(0);
  const hasNotes = emailNotes.length > 0;
  const hasUrgent = emailNotes.some((n) => n.priority === "urgent");
  const bannerEnabled = officeContextService.isBannerEnabled();
  const bannerDismissed = officeContextService.isBannerDismissedForItem(
    emailContext.itemId
  );

  const handleHideBannerOnEmail = () => {
    officeContextService.setBannerDismissedForItem(emailContext.itemId, true);
    forceUpdate((n) => n + 1);
  };

  const handleRestoreBannerOnEmail = () => {
    officeContextService.setBannerDismissedForItem(emailContext.itemId, false);
    officeContextService.syncEmailVisualMarkers(emailNotes);
    forceUpdate((n) => n + 1);
  };

  const handleToggleGlobalBanner = (enabled: boolean) => {
    officeContextService.setBannerEnabled(enabled);
    if (enabled) {
      officeContextService.syncEmailVisualMarkers(emailNotes);
    }
    forceUpdate((n) => n + 1);
  };

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

      <div className="banner-control-row">
        {!bannerEnabled ? (
          <>
            <span>🔕 Top-of-email notification is turned off.</span>
            <div className="banner-actions-group">
              <button
                type="button"
                className="link-btn"
                onClick={() =>
                  officeContextService.showEmailNotificationBanner(null)
                }
              >
                Clear Email Banner
              </button>
              <button
                type="button"
                className="link-btn"
                onClick={() => handleToggleGlobalBanner(true)}
              >
                Turn On
              </button>
            </div>
          </>
        ) : bannerDismissed ? (
          <>
            <span>🔕 Notification above this email is removed.</span>
            <div className="banner-actions-group">
              <button
                type="button"
                className="link-btn"
                onClick={handleRestoreBannerOnEmail}
              >
                Restore on this email
              </button>
              <button
                type="button"
                className="link-btn"
                onClick={() => handleToggleGlobalBanner(false)}
              >
                Turn Off Always
              </button>
            </div>
          </>
        ) : (
          <>
            <span>🔔 Notification above email:</span>
            <div className="banner-actions-group">
              <button
                type="button"
                className="link-btn danger-link"
                onClick={handleHideBannerOnEmail}
              >
                ✕ Remove Above Email
              </button>
              <button
                type="button"
                className="link-btn"
                onClick={() => handleToggleGlobalBanner(false)}
              >
                Turn Off Always
              </button>
            </div>
          </>
        )}
      </div>
    </section>
  );
};
