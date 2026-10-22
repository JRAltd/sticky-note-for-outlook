import React, { useEffect, useState } from "react";
import { RotateCcw } from "lucide-react";
import { EmailContextBanner } from "./components/EmailContextBanner";
import { EmptyState } from "./components/EmptyState";
import { HeaderBar } from "./components/HeaderBar";
import { NoteEditor } from "./components/NoteEditor";
import { SearchAndFilterBar } from "./components/SearchAndFilterBar";
import { SettingsModal } from "./components/SettingsModal";
import { StickyNoteCard } from "./components/StickyNoteCard";
import { useEmailContext } from "./hooks/useEmailContext";
import { useNotes } from "./hooks/useNotes";
import { useTheme } from "./hooks/useTheme";
import { CreateNoteInput, StickyNote } from "./types/note";

export const App: React.FC = () => {
  const { isDark, cycleTheme } = useTheme();
  const {
    emailContext,
    isInOutlook,
    demoEmails,
    switchDemoEmail,
  } = useEmailContext();

  const {
    emailNotes,
    allEmailsNotes,
    filteredNotes,
    availableTags,
    filters,
    setFilters,
    isLoading,
    isSyncing,
    graphSynced,
    lastDeletedNote,
    statusMessage,
    showToast,
    createNote,
    updateNote,
    deleteNote,
    undoLastDelete,
    togglePin,
    toggleChecklistItem,
    refreshNotes,
    syncWithGraphInteractive,
  } = useNotes(emailContext);

  const [isCreating, setIsCreating] = useState<boolean>(false);
  const [editingNote, setEditingNote] = useState<StickyNote | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState<boolean>(false);

  // Global keyboard shortcut: Ctrl+N / Alt+N to open New Sticky Note composer
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.altKey) && e.key.toLowerCase() === "n") {
        e.preventDefault();
        setEditingNote(null);
        setIsCreating(true);
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, []);

  const handleSaveNew = async (input: CreateNoteInput) => {
    await createNote(input);
    setIsCreating(false);
  };

  const handleSaveEdit = async (input: CreateNoteInput) => {
    if (!editingNote) return;
    await updateNote(editingNote.id, {
      title: input.title,
      content: input.content,
      color: input.color,
      priority: input.priority,
      tags: input.tags,
      checklist: input.checklist,
      isPinned: input.isPinned,
      scope: input.scope,
      reminderAt: input.reminderAt,
    });
    setEditingNote(null);
  };

  const hasActiveFilters =
    Boolean(filters.searchQuery) ||
    filters.selectedColor !== "all" ||
    filters.selectedPriority !== "all" ||
    filters.selectedTag !== "all" ||
    filters.onlyPinned;

  return (
    <div className="addin-shell">
      <HeaderBar
        currentEmailNoteCount={emailNotes.length}
        allEmailsNoteCount={allEmailsNotes.length}
        viewScope={filters.viewScope}
        onChangeViewScope={(scope) =>
          setFilters((prev) => ({ ...prev, viewScope: scope }))
        }
        isDark={isDark}
        onToggleTheme={cycleTheme}
        isSyncing={isSyncing || isLoading}
        graphSynced={graphSynced}
        isInOutlook={isInOutlook}
        onSyncNow={syncWithGraphInteractive}
        onOpenSettings={() => setIsSettingsOpen(true)}
        onOpenNewNote={() => {
          setEditingNote(null);
          setIsCreating(true);
        }}
      />

      <EmailContextBanner
        emailContext={emailContext}
        emailNotes={emailNotes}
        allEmailsNotes={allEmailsNotes}
        isInOutlook={isInOutlook}
        demoEmails={demoEmails}
        onSwitchDemoEmail={switchDemoEmail}
      />

      <SearchAndFilterBar
        filters={filters}
        availableTags={availableTags}
        onChangeFilters={setFilters}
      />

      <main className="notes-scroll-area">
        {isCreating && (
          <NoteEditor
            onSave={handleSaveNew}
            onCancel={() => setIsCreating(false)}
          />
        )}

        {editingNote && (
          <NoteEditor
            initialNote={editingNote}
            onSave={handleSaveEdit}
            onCancel={() => setEditingNote(null)}
          />
        )}

        {isLoading ? (
          <div className="loading-skeleton-stack">
            <div className="skeleton-card" />
            <div className="skeleton-card" />
          </div>
        ) : filteredNotes.length === 0 ? (
          <EmptyState
            hasFilters={hasActiveFilters}
            onClearFilters={() =>
              setFilters((prev) => ({
                ...prev,
                searchQuery: "",
                selectedColor: "all",
                selectedPriority: "all",
                selectedTag: "all",
                onlyPinned: false,
              }))
            }
            onOpenNewNote={() => {
              setEditingNote(null);
              setIsCreating(true);
            }}
            onQuickTemplateCreate={async (template) => {
              await createNote(template);
            }}
          />
        ) : (
          <div className="sticky-notes-list">
            {filteredNotes.map((note) => (
              <StickyNoteCard
                key={note.id}
                note={note}
                isDark={isDark}
                showEmailHeader={filters.viewScope === "all_emails"}
                onTogglePin={togglePin}
                onEdit={(n) => {
                  setIsCreating(false);
                  setEditingNote(n);
                }}
                onDelete={deleteNote}
                onToggleChecklistItem={toggleChecklistItem}
                onSelectTag={(tag) =>
                  setFilters((prev) => ({
                    ...prev,
                    selectedTag: prev.selectedTag === tag ? "all" : tag,
                  }))
                }
              />
            ))}
          </div>
        )}
      </main>

      {(statusMessage || lastDeletedNote) && (
        <div className="toast-bar">
          <span>{statusMessage || "Note deleted"}</span>
          {lastDeletedNote && (
            <button
              type="button"
              className="undo-toast-btn"
              onClick={undoLastDelete}
            >
              <RotateCcw size={12} /> Undo
            </button>
          )}
        </div>
      )}

      {isSettingsOpen && (
        <SettingsModal
          emailContext={emailContext}
          allNotes={allEmailsNotes}
          onClose={() => setIsSettingsOpen(false)}
          onConfigSaved={() => {
            refreshNotes(false);
            showToast("Storage & Graph settings saved");
          }}
          onNotesImported={(count) => {
            refreshNotes(false);
            setIsSettingsOpen(false);
            showToast(`Imported ${count} sticky note(s)`);
          }}
        />
      )}
    </div>
  );
};
