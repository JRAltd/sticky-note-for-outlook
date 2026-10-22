import React from "react";
import { Filter, Pin, Search, X } from "lucide-react";
import { STICKY_COLOR_PALETTE } from "../config/constants";
import { NoteColor, NoteFilterState, NotePriority } from "../types/note";

interface SearchAndFilterBarProps {
  filters: NoteFilterState;
  availableTags: string[];
  onChangeFilters: React.Dispatch<React.SetStateAction<NoteFilterState>>;
}

const COLOR_KEYS = Object.keys(STICKY_COLOR_PALETTE) as NoteColor[];

export const SearchAndFilterBar: React.FC<SearchAndFilterBarProps> = ({
  filters,
  availableTags,
  onChangeFilters,
}) => {
  const hasActiveFilters =
    Boolean(filters.searchQuery) ||
    filters.selectedColor !== "all" ||
    filters.selectedPriority !== "all" ||
    filters.selectedTag !== "all" ||
    filters.onlyPinned;

  const resetFilters = () => {
    onChangeFilters((prev) => ({
      ...prev,
      searchQuery: "",
      selectedColor: "all",
      selectedPriority: "all",
      selectedTag: "all",
      onlyPinned: false,
    }));
  };

  return (
    <div className="search-filter-container">
      <div className="search-input-wrapper">
        <Search size={15} className="search-icon" />
        <input
          type="text"
          className="search-input"
          placeholder={
            filters.viewScope === "current_email"
              ? "Search notes, tags, or checklist items..."
              : "Search across all emails, subjects, senders, or tags..."
          }
          value={filters.searchQuery}
          onChange={(e) =>
            onChangeFilters((prev) => ({
              ...prev,
              searchQuery: e.target.value,
            }))
          }
        />
        {filters.searchQuery && (
          <button
            type="button"
            className="clear-search-btn"
            onClick={() =>
              onChangeFilters((prev) => ({ ...prev, searchQuery: "" }))
            }
            title="Clear search"
          >
            <X size={14} />
          </button>
        )}
      </div>

      <div className="filter-controls-row">
        <div className="color-filter-dots" title="Filter by Sticky Note color">
          <button
            type="button"
            className={`color-dot-all ${
              filters.selectedColor === "all" ? "selected" : ""
            }`}
            onClick={() =>
              onChangeFilters((prev) => ({ ...prev, selectedColor: "all" }))
            }
          >
            All
          </button>
          {COLOR_KEYS.map((colorKey) => {
            const theme = STICKY_COLOR_PALETTE[colorKey];
            const isSelected = filters.selectedColor === colorKey;
            return (
              <button
                key={colorKey}
                type="button"
                className={`color-dot ${isSelected ? "selected" : ""}`}
                style={{ backgroundColor: theme.swatchHex }}
                title={`Filter: ${theme.label}`}
                onClick={() =>
                  onChangeFilters((prev) => ({
                    ...prev,
                    selectedColor: isSelected ? "all" : colorKey,
                  }))
                }
              />
            );
          })}
        </div>

        <div className="filter-select-group">
          <button
            type="button"
            className={`pin-filter-btn ${filters.onlyPinned ? "active" : ""}`}
            onClick={() =>
              onChangeFilters((prev) => ({
                ...prev,
                onlyPinned: !prev.onlyPinned,
              }))
            }
            title="Show only pinned notes"
          >
            <Pin size={12} />
          </button>

          <select
            className="compact-select"
            value={filters.selectedPriority}
            onChange={(e) =>
              onChangeFilters((prev) => ({
                ...prev,
                selectedPriority: e.target.value as NotePriority | "all",
              }))
            }
            title="Filter by priority"
          >
            <option value="all">Any Priority</option>
            <option value="urgent">Urgent</option>
            <option value="high">High</option>
            <option value="medium">Medium</option>
            <option value="low">Low</option>
          </select>

          <select
            className="compact-select"
            value={filters.sortBy}
            onChange={(e) =>
              onChangeFilters((prev) => ({
                ...prev,
                sortBy: e.target.value as NoteFilterState["sortBy"],
              }))
            }
            title="Sort notes"
          >
            <option value="updated_desc">Recently Updated</option>
            <option value="created_desc">Newest Created</option>
            <option value="created_asc">Oldest Created</option>
            <option value="priority_desc">Highest Priority</option>
          </select>
        </div>
      </div>

      {availableTags.length > 0 && (
        <div className="tag-filter-strip">
          <Filter size={12} className="tag-strip-icon" />
          <button
            type="button"
            className={`tag-chip-filter ${
              filters.selectedTag === "all" ? "active" : ""
            }`}
            onClick={() =>
              onChangeFilters((prev) => ({ ...prev, selectedTag: "all" }))
            }
          >
            All Tags
          </button>
          {availableTags.map((tag) => (
            <button
              key={tag}
              type="button"
              className={`tag-chip-filter ${
                filters.selectedTag === tag ? "active" : ""
              }`}
              onClick={() =>
                onChangeFilters((prev) => ({
                  ...prev,
                  selectedTag: prev.selectedTag === tag ? "all" : tag,
                }))
              }
            >
              #{tag}
            </button>
          ))}
          {hasActiveFilters && (
            <button
              type="button"
              className="reset-filters-link"
              onClick={resetFilters}
            >
              Reset
            </button>
          )}
        </div>
      )}
    </div>
  );
};
