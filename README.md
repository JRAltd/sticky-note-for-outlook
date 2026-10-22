# Outlook Email Sticky Notes Add-in (v2.0)

A production-grade Microsoft Outlook Add-in for attaching color-coded, timestamped sticky notes, interactive checklists, tags, and reminders directly to emails and conversation threads.

---

## Key Capabilities

- **Dual / Hybrid Storage Architecture**:
  1. **Office.js `CustomProperties` + `RoamingSettings` (Zero Setup Required)**: Works out-of-the-box in Outlook immediately without requiring an Azure AD App Registration. Stores notes directly on the Exchange message item and indexes them across emails in `RoamingSettings`.
  2. **Microsoft Graph Open Extensions (`com.jraltd.emailNotes`)**: Optionally connects via MSAL / Office SSO (`Mail.ReadWrite`) to persist structured JSON Open Extensions on `/me/messages/{id}/extensions/com.jraltd.emailNotes`.
- **Pinnable Taskpane (`VersionOverridesV1_1`)**: Pin the taskpane open in Outlook and click between emails in your inbox—notes automatically switch via `Office.EventType.ItemChanged`.
- **Multiple Timestamped Notes & Full CRUD**: Create unlimited sticky notes per email or conversation thread with 6 color themes, 4 priority levels, custom tags, interactive checklists, and follow-up reminders.
- **Cross-Email Global Search**: Search and filter notes within the current email or across **all annotated emails** by keyword, subject, sender, color, priority, or `#tag`.
- **Dark Mode**: Automatically detects Outlook's theme (`Office.context.officeTheme`) or toggles manually between Light and Dark mode.
- **Backup & Portability**: Export all email notes to JSON or Markdown, and import from JSON backups.

---

## Project Structure

```text
Sticky Note for Outlook/
├── manifest.xml                              # Local sideloading manifest (https://localhost:3000)
├── manifest.prod.xml                         # Enterprise / M365 Admin Center manifest template
├── package.json                              # Vite + React 18 + TypeScript + MSAL dependencies
├── vite.config.ts                            # Vite HTTPS multi-entry bundler configuration
├── tsconfig.json                             # Strict TypeScript configuration
├── taskpane.html                             # Taskpane HTML entry point
├── commands.html                             # Ribbon FunctionFile entry point
├── Start-EmailNotes.ps1                      # 1-Click local HTTPS server & SSL cert generator
├── server.py                                 # Zero-dependency HTTPS static server for ./dist
├── public/assets/                            # Add-in ribbon icons (16, 32, 64, 80, 128px)
├── dist/                                     # Pre-compiled production bundle (runs immediately)
├── scripts/
│   └── New-DeploymentPackage.ps1             # Generates email-notes-complete.zip
└── src/
    ├── main.tsx                              # Office.onReady + React 18 bootstrap
    ├── App.tsx                               # Main application shell & keyboard shortcuts
    ├── types/note.ts                         # Core domain & Graph Open Extension types
    ├── config/constants.ts                   # Color palettes, priority config, sample emails
    ├── commands/commands.ts                  # Ribbon 1-click "Quick Note" function handler
    ├── services/
    │   ├── office/officeContextService.ts    # Office.js item reader, REST ID converter & ItemChanged listener
    │   ├── auth/msalAuthService.ts           # OfficeRuntime SSO + MSAL Browser token manager
    │   └── storage/
    │       ├── officeCustomPropsAdapter.ts   # Exchange CustomProperties + RoamingSettings index
    │       ├── graphOpenExtensionAdapter.ts  # Microsoft Graph /me/messages/{id}/extensions API
    │       └── storageManager.ts             # Hybrid two-way sync, conflict resolution & export/import
    ├── hooks/
    │   ├── useEmailContext.ts                # Active email state & ItemChanged subscription
    │   ├── useNotes.ts                       # CRUD, checklist toggle, search/filter & undo
    │   └── useTheme.ts                       # Outlook theme detector & Dark Mode toggle
    ├── components/
    │   ├── HeaderBar.tsx                     # Storage status badge, view tabs & toolbar actions
    │   ├── EmailContextBanner.tsx            # Active email metadata & sandbox email switcher
    │   ├── SearchAndFilterBar.tsx            # Search bar, color dots, priority & tag filter strip
    │   ├── NoteEditor.tsx                    # Sticky note composer/editor with checklist builder
    │   ├── StickyNoteCard.tsx                # Sticky note card with pin, copy, edit, delete & tags
    │   ├── SettingsModal.tsx                 # Storage mode, Azure AD config & JSON/MD export/import
    │   └── EmptyState.tsx                    # Empty state with 1-click quick note templates
    └── styles/index.css                      # Fluent-inspired styling + Dark Mode variables
```

---

## Quick Start (Zero Build Step Needed)

Because a pre-compiled bundle is already included in `./dist`, you can start the local HTTPS server immediately:

```powershell
.\Start-EmailNotes.ps1
```

Then open **`https://localhost:3000/taskpane.html`** in your browser to test the interactive sandbox, or sideload `manifest.xml` into Outlook!

### Optional: Full Node.js / Vite Development Workflow

If you want to edit the React/TypeScript files in `src/` with hot module replacement:

1. Install Node.js LTS (`winget install OpenJS.NodeJS.LTS`)
2. Run:
   ```powershell
   npm install
   npm run start
   ```

---

## Sideloading into Outlook

1. Make sure the local server is running (`.\Start-EmailNotes.ps1` or `npm run start`).
2. **Outlook on the Web / New Outlook for Windows**:
   - Open `https://aka.ms/olksideload` (or click **Apps → Add apps → Manage your apps → Add a custom add-in → Add from file**).
   - Select [`manifest.xml`](./manifest.xml).
   - Open any email and click **Email Notes** (or **Quick Note**) in the ribbon or message actions bar.
3. **Pinning the Pane**:
   - Click the **Pin** icon in the top-right corner of the Email Sticky Notes taskpane so it stays open as you click through your inbox.
