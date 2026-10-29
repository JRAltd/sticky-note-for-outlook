import { useCallback, useEffect, useMemo, useState } from "react";
import { PRIORITY_CONFIG } from "../config/constants";
import { storageManager } from "../services/storage/storageManager";
import {
  CreateNoteInput,
  EmailContextInfo,
  NoteFilterState,
  StickyNote,
} from "../types/note";

export function useNotes(emailContext: EmailContextInfo) {
  const [emailNotes, setEmailNotes] = useState<StickyNote[]>([]);
  const [allEmailsNotes, setAllEmailsNotes] = useState<StickyNote[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isSyncing, setIsSyncing] = useState<boolean>(false);
  const [graphSynced, setGraphSynced] = useState<boolean>(false);
  const [lastDeletedNote, setLastDeletedNote] = useState<StickyNote | null>(
    null
  );
  const [statusMessage, setStatusMessage] = useState<string | null>(null);

  const [filters, setFilters] = useState<NoteFilterState>({
    searchQuery: "",
    selectedColor: "all",
    selectedPriority: "all",
    selectedTag: "all",
    sortBy: "updated_desc",
    viewScope: "current_email",
    onlyPinned: false,
  });

  const showToast = useCallback((msg: string) => {
    setStatusMessage(msg);
    window.setTimeout(() => {
      setStatusMessage((prev) => (prev === msg ? null : prev));
    }, 3800);
  }, []);

  const refreshNotes = useCallback(
    async (interactiveGraph = false) => {
      setIsLoading(true);
      try {
        const result = await storageManager.loadNotesForEmail(
          emailContext,
          interactiveGraph
        );
        setEmailNotes(result.notes);
        setGraphSynced(result.graphSynced);
        setAllEmailsNotes(storageManager.getAllEmailsNotes());
      } finally {
        setIsLoading(false);
      }
    },
    [emailContext]
  );

  useEffect(() => {
    refreshNotes(false);
  }, [refreshNotes]);

  const persistNotesForTargetEmail = useCallback(
    async (
      targetItemId: string,
      nextTargetNotes: StickyNote[],
      interactiveGraph = false
    ) => {
      const isCurrentEmail = targetItemId === emailContext.itemId;
      if (isCurrentEmail) {
        setEmailNotes(nextTargetNotes);
      }
      setIsSyncing(true);
      try {
        const targetCtx: EmailContextInfo = isCurrentEmail
          ? emailContext
          : {
              ...emailContext,
              itemId: targetItemId,
              restId: targetItemId,
            };
        const res = await storageManager.saveNotesForEmail(
          targetCtx,
          nextTargetNotes,
          interactiveGraph
        );
        if (isCurrentEmail) {
          setEmailNotes(res.updatedNotes);
        }
        if (res.savedGraph) {
          setGraphSynced(true);
        }
        setAllEmailsNotes(storageManager.getAllEmailsNotes());
      } finally {
        setIsSyncing(false);
      }
    },
    [emailContext]
  );

  const persistEmailNotes = useCallback(
    async (nextNotes: StickyNote[], interactiveGraph = false) => {
      await persistNotesForTargetEmail(
        emailContext.itemId,
        nextNotes,
        interactiveGraph
      );
    },
    [emailContext.itemId, persistNotesForTargetEmail]
  );

  const createNote = useCallback(
    async (input: CreateNoteInput) => {
      const nowIso = new Date().toISOString();
      const newNote: StickyNote = {
        id: `note-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        itemId: emailContext.itemId,
        conversationId: emailContext.conversationId,
        scope: input.scope,
        emailSubject: emailContext.subject,
        emailSender: `${emailContext.senderName} <${emailContext.senderEmail}>`,
        title: input.title.trim(),
        content: input.content.trim(),
        color: input.color,
        priority: input.priority,
        tags: input.tags,
        checklist: input.checklist,
        isPinned: input.isPinned,
        reminderAt: input.reminderAt,
        author: emailContext.userDisplayName,
        createdAt: nowIso,
        updatedAt: nowIso,
        syncStatus: "local_only",
      };

      const nextNotes = [newNote, ...emailNotes];
      await persistEmailNotes(nextNotes);
      showToast("Sticky note attached to email");
      return newNote;
    },
    [emailContext, emailNotes, persistEmailNotes, showToast]
  );

  const updateNote = useCallback(
    async (noteId: string, changes: Partial<StickyNote>) => {
      const target =
        emailNotes.find((n) => n.id === noteId) ||
        allEmailsNotes.find((n) => n.id === noteId);
      if (!target) return;

      const targetItemId = target.itemId || emailContext.itemId;
      const baseList =
        targetItemId === emailContext.itemId
          ? emailNotes
          : allEmailsNotes.filter((n) => n.itemId === targetItemId);

      const nowIso = new Date().toISOString();
      const nextNotes = baseList.map((n) =>
        n.id === noteId
          ? {
              ...n,
              ...changes,
              updatedAt: nowIso,
            }
          : n
      );
      await persistNotesForTargetEmail(targetItemId, nextNotes);
      showToast("Sticky note updated");
    },
    [
      allEmailsNotes,
      emailContext.itemId,
      emailNotes,
      persistNotesForTargetEmail,
      showToast,
    ]
  );

  const deleteNote = useCallback(
    async (noteId: string) => {
      const target =
        emailNotes.find((n) => n.id === noteId) ||
        allEmailsNotes.find((n) => n.id === noteId);
      if (!target) return;

      const targetItemId = target.itemId || emailContext.itemId;
      const baseList =
        targetItemId === emailContext.itemId
          ? emailNotes
          : allEmailsNotes.filter((n) => n.itemId === targetItemId);

      setLastDeletedNote(target);
      const nextNotes = baseList.filter((n) => n.id !== noteId);
      await persistNotesForTargetEmail(targetItemId, nextNotes);
      showToast("Sticky note deleted");
    },
    [
      allEmailsNotes,
      emailContext.itemId,
      emailNotes,
      persistNotesForTargetEmail,
      showToast,
    ]
  );

  const undoLastDelete = useCallback(async () => {
    if (!lastDeletedNote) return;
    const targetItemId = lastDeletedNote.itemId || emailContext.itemId;
    const baseList =
      targetItemId === emailContext.itemId
        ? emailNotes
        : allEmailsNotes.filter((n) => n.itemId === targetItemId);

    const restored = [
      lastDeletedNote,
      ...baseList.filter((n) => n.id !== lastDeletedNote.id),
    ];
    setLastDeletedNote(null);
    await persistNotesForTargetEmail(targetItemId, restored);
    showToast("Restored deleted sticky note");
  }, [
    allEmailsNotes,
    emailContext.itemId,
    emailNotes,
    lastDeletedNote,
    persistNotesForTargetEmail,
    showToast,
  ]);

  const togglePin = useCallback(
    async (noteId: string) => {
      const target =
        emailNotes.find((n) => n.id === noteId) ||
        allEmailsNotes.find((n) => n.id === noteId);
      if (!target) return;
      await updateNote(noteId, { isPinned: !target.isPinned });
    },
    [allEmailsNotes, emailNotes, updateNote]
  );

  const toggleChecklistItem = useCallback(
    async (noteId: string, checklistItemId: string) => {
      const target =
        emailNotes.find((n) => n.id === noteId) ||
        allEmailsNotes.find((n) => n.id === noteId);
      if (!target) return;
      const updatedChecklist = target.checklist.map((item) =>
        item.id === checklistItemId
          ? { ...item, completed: !item.completed }
          : item
      );
      await updateNote(noteId, { checklist: updatedChecklist });
    },
    [allEmailsNotes, emailNotes, updateNote]
  );

  const syncWithGraphInteractive = useCallback(async () => {
    setIsSyncing(true);
    try {
      await refreshNotes(true);
      await persistEmailNotes(emailNotes, true);
      showToast("Synced with Microsoft Graph Open Extensions");
    } catch (err) {
      showToast(
        err instanceof Error ? err.message : "Microsoft Graph sync failed"
      );
    } finally {
      setIsSyncing(false);
    }
  }, [emailNotes, persistEmailNotes, refreshNotes, showToast]);

  const sourceNotes =
    filters.viewScope === "current_email" ? emailNotes : allEmailsNotes;

  const availableTags = useMemo(() => {
    const tagSet = new Set<string>();
    sourceNotes.forEach((n) => n.tags.forEach((t) => tagSet.add(t)));
    return Array.from(tagSet).sort();
  }, [sourceNotes]);

  const filteredNotes = useMemo(() => {
    const q = filters.searchQuery.trim().toLowerCase();

    const matches = sourceNotes.filter((note) => {
      if (filters.onlyPinned && !note.isPinned) return false;
      if (
        filters.selectedColor !== "all" &&
        note.color !== filters.selectedColor
      ) {
        return false;
      }
      if (
        filters.selectedPriority !== "all" &&
        note.priority !== filters.selectedPriority
      ) {
        return false;
      }
      if (
        filters.selectedTag !== "all" &&
        !note.tags.includes(filters.selectedTag)
      ) {
        return false;
      }

      if (!q) return true;

      const haystack = [
        note.title,
        note.content,
        note.emailSubject,
        note.emailSender,
        ...note.tags,
        ...note.checklist.map((c) => c.text),
      ]
        .join(" ")
        .toLowerCase();

      return haystack.includes(q);
    });

    return matches.sort((a, b) => {
      if (a.isPinned !== b.isPinned) {
        return a.isPinned ? -1 : 1;
      }
      switch (filters.sortBy) {
        case "created_desc":
          return b.createdAt.localeCompare(a.createdAt);
        case "created_asc":
          return a.createdAt.localeCompare(b.createdAt);
        case "priority_desc":
          return (
            PRIORITY_CONFIG[b.priority].rank - PRIORITY_CONFIG[a.priority].rank
          );
        case "updated_desc":
        default:
          return b.updatedAt.localeCompare(a.updatedAt);
      }
    });
  }, [filters, sourceNotes]);

  return {
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
  };
}
