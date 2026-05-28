# Sales Pipeline Dashboard — SharePoint → GitHub Pages → TV

A zero-backend dashboard that reads the **Project Management List** from SharePoint via Microsoft Graph and displays a live sales-pipeline view on a TV: open enquiries, quotes out, orders won this month, pipeline value, a funnel chart, and a rotating table of quotes awaiting decision.

## How it works

```
 ┌──────────────┐    silent SSO    ┌──────────────────┐
 │   TV browser │ ───────────────▶ │  Microsoft Entra │
 │  (signed in) │ ◀─────token───── │   (Azure AD)     │
 └──────┬───────┘                  └──────────────────┘
        │ Bearer token
        ▼
 ┌──────────────────────────┐
 │  Microsoft Graph API     │
 │  /sites/.../lists/items  │
 └────────────┬─────────────┘
              │ JSON
              ▼
 ┌──────────────────────────┐
 │  Static HTML/JS page     │
 │  hosted on GitHub Pages  │
 └──────────────────────────┘
```

No server. No secrets in the code. Auth happens in the TV's browser using MSAL.js, then Graph API calls return the list data. Refreshes every 5 minutes.

## What the dashboard shows

**KPI tiles**
- **Open Enquiries** — enquiry received, no quote issued yet, not declined.
- **Quotes Out** — quote issued, no order placed, not declined. Subtitle shows total £.
- **Won This Month** — orders placed since the 1st of the current month. Subtitle shows total £.
- **Pipeline Value** — sum of Order Net Value across all open quotes.

**Funnel chart** — horizontal bar showing live counts at each stage: Open Enquiries → Quotes Out → Active Orders → Completed (last 30d).

**Table** — every open quote awaiting a decision, oldest first (so the team can see what to chase). Quotes 10+ days old are highlighted amber, 21+ days old red. Footer shows total open-quote count and value. If there are more rows than fit, the table auto-rotates pages every 15 seconds.

### How the pipeline stages are decided

A row falls into exactly one stage at a time, based on which dates are filled in. Newest stage wins:

| Stage | Rule |
|---|---|
| Declined | `Quote Declined Date` is set |
| Complete | `Production Complete Date` is set |
| Active Order | `Order Placed Date` is set, not yet complete |
| Quote Out | `Quote Issued Date` is set, no order, not declined |
| Enquiry | `Enquiry Received Date` is set, no quote, not declined |

So the dashboard is driven by the **dates** in the list, not by the Project Status choice field — that's more reliable because dates are set when something actually happens.

---

## What you need

1. A Microsoft 365 / SharePoint account that has read access to the Project Management List.
2. Permission to register an app in your Entra (Azure AD) tenant — i.e. you're a tenant admin.
3. A free GitHub account.
4. The TV's web browser (any modern Chrome/Edge works).

---

## Step 1 — Register an Entra (Azure AD) app

1. Go to https://entra.microsoft.com → sign in.
2. **Applications → App registrations → + New registration**.
3. Name: `Jobs Dashboard`. Account types: **Single tenant**. Leave Redirect URI blank for now. **Register**.
4. On the Overview page, copy the **Application (client) ID** and **Directory (tenant) ID**.
5. **API permissions → + Add a permission → Microsoft Graph → Delegated permissions**. Tick `Sites.Read.All`. **Add permissions** → **Grant admin consent**.
6. **Authentication → + Add a platform → Single-page application**. Enter `https://localhost` for now. **Configure**.

We update that Redirect URI once GitHub Pages is live.

---

## Step 2 — Confirm your SharePoint list

This dashboard is already configured for:

- **Site:** `https://mpepperjoinery.sharepoint.com/sites/MPJDev`
- **List:** `Project Management List`

If those change, edit `config.js`:

```js
sharePointHostname: "mpepperjoinery",
sitePath:           "/sites/MPJDev",
listName:           "Project Management List",
```

The column mapping in `config.js` is already populated with the real internal names from your list (including SharePoint's 32-character truncation and `_x0020_` encoding for spaces). You shouldn't need to touch the `columns` block unless someone renames or recreates a column in SharePoint.

---

## Step 3 — Create the GitHub repo and push

1. https://github.com → **New repository** → name `jobs-dashboard`, **Public**, no initial files. **Create**.
2. In a terminal opened in this folder:

   ```bash
   git init
   git add .
   git commit -m "Initial dashboard"
   git branch -M main
   git remote add origin https://github.com/<your-username>/jobs-dashboard.git
   git push -u origin main
   ```

GitHub Desktop (https://desktop.github.com) does the same thing with buttons if you'd rather avoid the terminal.

---

## Step 4 — Fill in `config.js`

Open `config.js` and replace the two placeholders:

```js
tenantId: "paste-the-Directory-(tenant)-ID-here",
clientId: "paste-the-Application-(client)-ID-here",
```

Save, commit, push:

```bash
git commit -am "Add tenant + client IDs"
git push
```

---

## Step 5 — Turn on GitHub Pages

1. In your GitHub repo, **Settings → Pages**.
2. **Build and deployment → Source:** *Deploy from a branch*. Branch: `main`, folder: `/ (root)`. **Save**.
3. After ~1 minute it'll show your site URL, e.g. `https://<your-username>.github.io/jobs-dashboard/`.
4. (If you set up a custom domain like `dashboard.mpepperjoinery.co.uk` — your URL is that one instead.)

---

## Step 6 — Tell Entra about the GitHub Pages URL

Back in Entra → your app → **Authentication → Single-page application**:

1. Remove the `https://localhost` placeholder.
2. Add the **exact** site URL as a Redirect URI, including the trailing slash, e.g. `https://<your-username>.github.io/jobs-dashboard/` or `https://dashboard.mpepperjoinery.co.uk/`.
3. Save.

---

## Step 7 — Open on the TV

1. Open the dashboard URL on the TV's browser.
2. Sign in with a Microsoft 365 account that can read the SharePoint list. Tick "Stay signed in".
3. Dashboard loads. Auto-refreshes every 5 minutes.

### Kiosk mode (recommended for TVs)

**Windows PC behind a TV** — shortcut target:

```
"C:\Program Files\Google\Chrome\Application\chrome.exe" --kiosk --app=https://dashboard.mpepperjoinery.co.uk/
```

---

## Security notes

- `Sites.Read.All` is **delegated** — the dashboard can only ever read what the signed-in user can already read.
- For best practice, sign the TV in with a **dedicated low-privilege user** (e.g. `tv-display@mpepperjoinery.co.uk`) that only has read access to the SharePoint site holding this list. Don't sign it in with a director / admin account.
- Client ID and Tenant ID in `config.js` are not secrets — they're public identifiers. Safe to commit to a public repo.
- Order Net Value will be visible on the TV. Position the TV where customers/visitors who shouldn't see those figures can't.

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| Stuck on "Sign in to view the dashboard" | Redirect URI in Entra doesn't match the GitHub Pages URL exactly — including trailing slash. |
| `AADSTS65001` consent error | Click "Grant admin consent" in Entra → API permissions. |
| `Graph 404` and "list not found" | `listName` in `config.js` must match the list's display name exactly. |
| KPIs show 0 | A column intern