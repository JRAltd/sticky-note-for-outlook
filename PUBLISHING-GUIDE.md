# How to Install & Publish Email Sticky Notes for Outlook

## Live Public URLs (GitHub + GitHub Pages)
- **GitHub Repository**: https://github.com/JRAltd/sticky-note-for-outlook
- **Public Install Page**: https://jraltd.github.io/sticky-note-for-outlook/
- **Live Public Taskpane**: https://jraltd.github.io/sticky-note-for-outlook/taskpane.html
- **Live Public Manifest**: https://jraltd.github.io/sticky-note-for-outlook/manifest.xml
- **Privacy Policy URL**: https://jraltd.github.io/sticky-note-for-outlook/privacy.html
- **Support URL**: https://jraltd.github.io/sticky-note-for-outlook/support.html

---

## Option A: How Anyone Can Install It in Outlook Right Now (No Waiting)

Anyone you share `https://jraltd.github.io/sticky-note-for-outlook/` with can install the add-in in 30 seconds:

1. **Download the Manifest**:
   - Visit https://jraltd.github.io/sticky-note-for-outlook/ and click **⬇️ 1. Download Add-in Manifest (manifest.xml)** *(or use `store-assets\manifest.public.xml` directly from this project folder)*.
2. **Open Outlook's Add-in Manager**:
   - Go to **https://aka.ms/olksideload** in your browser, **OR** inside **Outlook Web / New Outlook for Windows**:
     - Click the **Apps** icon on the far-left sidebar.
     - Click **Manage your apps** (bottom-left) or **Add apps → Manage your apps**.
3. **Upload the File**:
   - Click **My add-ins** (or scroll to **Custom Addins**).
   - Click **Add a custom add-in → Add from file...**
   - Select `email-sticky-notes-manifest.xml` (or `manifest.public.xml`) and click **Install**.
4. **Use & Pin It**:
   - Open any email in your Inbox, click **Email Notes** in the message ribbon/toolbar, and click the **Pin (📌)** icon in the top-right corner of the pane so it stays open as you click between emails.

---

## Option B: List It Publicly in the Official Outlook Add-in Store (Microsoft AppSource)

All required submission files (sized to exact Microsoft Partner Center pixel specifications) are saved in your **`store-assets/`** folder:
- `store-assets/manifest.public.xml`
- `store-assets/store-logo-300x300.png` (300×300 PNG Store Logo)
- `store-assets/screenshot-1-light-1366x768.png` (1366×768 PNG Light Mode Screenshot)
- `store-assets/screenshot-2-dark-1366x768.png` (1366×768 PNG Dark Mode Screenshot)

### Step 1: Enroll in Microsoft Partner Center
1. Go to **https://partner.microsoft.com/dashboard/commercial-marketplace**.
2. Sign in with a Microsoft 365 / Work account and complete the **Commercial Marketplace Publisher** enrollment (free for Office Add-in publishers).

### Step 2: Create the Office Add-in Offer
1. In the left menu of Partner Center, click **Marketplace offers → Overview**.
2. At the top, click **+ New offer → Office Store → Office Add-in**.
3. Fill in the dialog:
   - **Offer ID**: `jraltd-email-sticky-notes`
   - **Offer alias**: `Email Sticky Notes`
4. Click **Create**.

### Step 3: Offer Setup Tab
1. Under **Customer leads**, leave *"No CRM system connected"* selected.
2. Click **Save draft**.

### Step 4: Packages Tab (Upload Your Manifest)
1. Click **Packages** in the left sidebar.
2. Upload **`store-assets\manifest.public.xml`**.
3. Wait ~15 seconds for Microsoft's validator to verify your GitHub Pages URLs (`https://jraltd.github.io/sticky-note-for-outlook/...`) and show **Passed**.
4. Click **Save draft**.

### Step 5: Properties Tab (Copy & Paste These Exact Values)
1. Click **Properties** in the left sidebar.
2. **Categories**:
   - Primary category: **Productivity**
3. **Legal and support info**:
   - **Privacy policy link**: `https://jraltd.github.io/sticky-note-for-outlook/privacy.html`
   - **Support link**: `https://jraltd.github.io/sticky-note-for-outlook/support.html`
4. Click **Save draft**.

### Step 6: Marketplace Listings Tab (Copy & Paste + Upload Pre-Made Assets)
1. Click **Marketplace listings → English (United States)** in the left sidebar.
2. **Name**: `Email Sticky Notes`
3. **Search results summary**:
   ```text
   Attach color-coded, timestamped sticky notes, checklists, and Inbox category tags to any email in Outlook.
   ```
4. **Description**:
   ```text
   Email Sticky Notes lets you attach private, color-coded, timestamped sticky notes directly to any email or conversation thread in Microsoft Outlook.

   Key Features:
   • Automatic Inbox Category Marking: Automatically tags annotated emails with a yellow "Sticky Note" category badge right in your Inbox message list and shows a top-of-email notification banner.
   • Pinnable Taskpane: Pin the sticky notes pane open and click through your Inbox—notes automatically switch to match the selected email.
   • Rich Sticky Notes: Choose from 6 sticky note colors, 4 priority levels, interactive action checklists, custom #tags, and follow-up reminders.
   • Global Cross-Email Search: Search and filter notes within the current email or across every email you have ever annotated.
   • Privacy-First Hybrid Storage: Works out-of-the-box using Office.js Exchange CustomProperties and RoamingSettings, with optional Microsoft Graph Open Extensions sync.
   • Dark Mode & Backup: Includes automatic Outlook Dark Mode detection plus 1-click JSON and Markdown export/import.
   ```
5. **Search keywords**: `sticky notes`, `email notes`, `outlook notes`, `annotations`, `checklist`
6. **Marketplace media** (upload from `store-assets/`):
   - **300 x 300 px Store Logo**: Upload `store-assets\store-logo-300x300.png`
   - **Screenshots (1366 x 768 px PNG)**: Upload both:
     - `store-assets\screenshot-1-light-1366x768.png` *(Caption: "Attach color-coded sticky notes, checklists, and Inbox category markers to any email.")*
     - `store-assets\screenshot-2-dark-1366x768.png` *(Caption: "Full Dark Mode support with cross-email search and priority badges.")*
7. Click **Save draft**.

### Step 7: Review and Publish Tab (Certification Notes)
1. Click **Review and publish** in the left sidebar.
2. In the **Notes for certification** box, paste:
   ```text
   Testing Instructions for Microsoft Certification Reviewer:
   1. No external account or paid subscription is required. The add-in works out-of-the-box using Office.js CustomProperties and RoamingSettings.
   2. Open any email in Outlook (Web, New Outlook for Windows, or Desktop) and click "Email Notes" in the ribbon.
   3. Click "+ New Note" to attach a sticky note, checklist, or tag to the email.
   4. Reason for ReadWriteMailbox permission: Used to register and apply the "Sticky Note" color category (Office.context.mailbox.masterCategories and item.categories) so users can visually identify annotated emails in their Inbox list.
   ```
3. Click **Publish**.
