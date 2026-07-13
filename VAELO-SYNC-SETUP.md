# Vaelo — Google Sheets sync setup

This connects all four logins (team, Dhruv, client staff, client head) to **one Google Sheet**, so everyone shares the same live data and it also lands in the Sheet in readable form.

You do this **once**, and it takes about 10 minutes. You need a Google account. Claude cannot do these steps for you — they happen inside your Google account.

---

## Step 1 — Create the Sheet
1. Go to https://sheets.google.com and create a **blank spreadsheet**.
2. Name it something like **Vaelo Ops Data**.

## Step 2 — Add the script
1. In the Sheet, open the menu **Extensions → Apps Script**.
2. Delete whatever code is there.
3. Open **`vaelo-apps-script.gs`** (in your VAELO folder), copy **all** of it, and paste it in.
4. Click the **Save** icon (💾).

## Step 3 — Deploy it as a Web app
1. Click **Deploy → New deployment**.
2. Click the gear icon ⚙ next to "Select type" and choose **Web app**.
3. Set:
   - **Description:** `Vaelo sync` (anything)
   - **Execute as:** **Me**
   - **Who has access:** **Anyone**
4. Click **Deploy**.
5. Google asks you to **authorise** — click through, choose your account, and on the "Google hasn't verified this app" screen click **Advanced → Go to (your project) → Allow**. (This is normal for your own script.)
6. Copy the **Web app URL** it shows. It ends with **`/exec`** and looks like:
   `https://script.google.com/macros/s/AKfy..../exec`

## Step 4 — Turn on sync in the app
1. Open **`vaelo-login.html`** in a text editor.
2. Near the top of the `<script>` section, find this line:
   ```js
   const SYNC_URL = '';
   ```
3. Paste your URL between the quotes:
   ```js
   const SYNC_URL = 'https://script.google.com/macros/s/AKfy..../exec';
   ```
4. Save the file.

Done. Now every login reads and writes through the Sheet.

---

## What you'll see in the Sheet
The script auto-creates three tabs:
- **store** — the raw synced data (don't edit by hand).
- **Calendar** — a readable view of the content pipeline (client, date, format, topic, product, stage).
- **Finance** — a readable view of billing (client, date, item, category, amount, paid, added by).

These refresh automatically whenever someone changes something in the app.

---

## Good to know (prototype-grade)
- **Shared, cross-device:** because everyone talks to the same Sheet, the four logins now genuinely share data across different computers.
- **A few seconds of latency**, and **last-write-wins** if two people save the same second — fine at your scale, not built for heavy concurrent editing.
- **Offline:** if the Sheet can't be reached, the app falls back to a local cache so it still opens; changes sync when it's reachable again.
- **Access:** "Who has access: Anyone" means anyone with the `/exec` URL can read/write the data. Keep the URL private. This is still prototype-level — for real security you'd move to the proper backend in your dev brief.
- **To update the script later:** edit it in Apps Script, then **Deploy → Manage deployments → edit (✏) → Version: New version → Deploy**. The URL stays the same.
