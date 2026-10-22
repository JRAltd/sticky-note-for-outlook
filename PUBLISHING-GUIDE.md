# How to Publish Email Sticky Notes So Anyone Can Install It

Right now, your add-in works locally via `https://localhost:3000`. Because Outlook add-ins are web applications loaded from the URL inside `manifest.xml`, making it installable by **anyone** (without them needing to install any software or run a local server) only requires hosting the `dist/` folder on any free `https://` static web host.

Below are the **3 ways** to make it available to others—from a 60-second free drag-and-drop share link to the official Microsoft Outlook Add-in Store (AppSource).

---

## Method 1: Share with Anyone via Free 60-Second Hosting (Netlify / Cloudflare / GitHub Pages)

This lets **anyone** install your add-in in 30 seconds by visiting your install page (`https://your-addin-url/install.html`) or uploading a single `manifest.public.xml` file.

### Step 1 — Host the `dist/` folder for free:
- **Easiest (Drag & Drop — 60 seconds)**:
  1. Go to **[app.netlify.com/drop](https://app.netlify.com/drop)** (or **Cloudflare Pages** / **GitHub Pages**).
  2. Drag and drop the **`dist`** folder from this project into the browser window.
  3. Netlify/Cloudflare immediately gives you a permanent HTTPS URL (for example: `https://email-sticky-notes.netlify.app`).

### Step 2 — Stamp your public HTTPS URL into the manifest:
Run this single PowerShell command in your project folder (replacing the URL with yours):
```powershell
.\scripts\Set-PublicHostUrl.ps1 -PublicUrl "https://email-sticky-notes.netlify.app"
```
Then drag the updated `dist` folder to Netlify once more so `/manifest.xml` on the site has your live URL.

### Step 3 — Share!
Send anyone either:
- Your self-service install link: **`https://email-sticky-notes.netlify.app/install.html`** (which has a 1-click download button + instructions), OR
- The generated **`manifest.public.xml`** file directly. They simply add it once in Outlook (**Apps → Manage Add-ins → Add from file**) and it works forever across Outlook Web, New Outlook for Windows, and Classic Outlook!

---

## Method 2: Deploy Automatically to Your Whole Organization (Microsoft 365 Admin Center)

If you want the **Email Notes** button to appear automatically in the Outlook ribbon for everyone in your company (with zero user setup):

1. Complete **Step 1 & Step 2** above so `manifest.public.xml` points to your hosted `https://` URL.
2. Open the **[Microsoft 365 Admin Center](https://admin.microsoft.com)**.
3. Navigate to **Settings → Integrated apps → Upload custom apps**.
4. Choose **Office Add-in** and upload **`manifest.public.xml`**.
5. Assign it to **Entire organization** (or specific users/groups) and click **Deploy**.
   - Within a few hours, the **Email Notes** and **Quick Note** buttons appear automatically in every user's Outlook!

---

## Method 3: Publish to the Official Microsoft Outlook Store (Microsoft AppSource)

If you want anyone in the world to find **"Email Sticky Notes"** by searching inside Outlook's built-in **Get Add-ins** store:

1. **Checklist Already Built Into This Project**:
   - Store-compliant PNG icons (`16x16`, `32x32`, `64x64`, `80x80`, `128x128`, and `300x300` Store Logo) are pre-generated in `dist/assets/`.
   - Store-required Privacy Policy (`/privacy.html`) and Support Page (`/support.html`) are pre-built in `dist/`.
   - Pinnable Taskpane (`VersionOverridesV1_1`) and `ReadWriteMailbox` permissions are configured.
2. **Submit via Microsoft Partner Center**:
   - Enroll in the **Microsoft Partner Center (Commercial Marketplace)** program at [partner.microsoft.com](https://partner.microsoft.com/).
   - Create a new **Office Add-in** offer.
   - Upload your `manifest.public.xml` (pointing to your HTTPS host), upload `dist/assets/icon-300.png` as the 300×300 Store Logo, and provide your hosted `/privacy.html` and `/support.html` URLs.
   - Click **Publish** for Microsoft AppSource certification review.
