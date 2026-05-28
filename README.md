# Sales Pipeline Dashboard

Reads SharePoint "Project Management List" via Microsoft Graph and shows a TV pipeline view: open enquiries, quotes out, won this month, pipeline value, a funnel chart, and a rotating table of quotes awaiting decision.

## Setup

1. **Entra app registration** — register a Single-page application named `Jobs Dashboard`, grant `Sites.Read.All` delegated permission with admin consent.
2. **`config.js`** — paste Tenant ID and Client ID.
3. **GitHub** — push this folder to a public repo.
4. **GitHub Pages** — Settings → Pages → Deploy from branch `main`, `/ (root)`.
5. **Entra redirect URI** — add the GitHub Pages URL (with trailing slash) as a Single-page application redirect URI.
6. **TV** — open the URL, sign in once, tick "Stay signed in".

## Pipeline logic

Each row is classified into exactly one stage by which dates are filled in (newest stage wins):

| Stage | Rule |
|---|---|
| Declined | Quote Declined Date set |
| Complete | Production Complete Date set |
| Active Order | Order Placed Date set |
| Quote Out | Quote Issued Date set, no order, not declined |
| Enquiry | Enquiry Received Date set, no quote, not declined |

## Files
- `index.html` — page layout
- `styles.css` — dark theme
- `config.js` — IDs and column mapping (edit this)
- `app.js` — auth, fetch, render
