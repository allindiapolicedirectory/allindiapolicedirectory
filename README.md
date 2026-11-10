# All India Police Directory — GitHub Pages

```
aipd-github-pages/
├── apps-script/Code.gs   ← paste into your Apps Script project (backend API)
└── docs/index.html       ← GitHub Pages front-end
```

How it works: GitHub Pages hosts the page only. Google Apps Script is the backend. It returns JSON and reads and writes the `AIPD` sheet. The page calls it with `fetch()` instead of `google.script.run`.

## Step 1: Update Apps Script (backend)

1. Open your existing Apps Script project and replace **Code.gs** with `apps-script/Code.gs`.
   You can keep your old `Index.html`. Opening the `/exec` URL with no `?action` still shows the old UI.
2. **Set the access key:** go to Project Settings (⚙) ▸ Script Properties ▸ Add property:
   - Property: `ACCESS_KEY`
   - Value: a strong secret, for example `Aipd@2026#Chittoor`
   The API refuses every request until this is set.
3. **Deploy:** click Deploy ▸ New deployment ▸ type **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**. This is required so fetch() from GitHub Pages works. The access key protects the data.
4. Copy the Web App URL (`https://script.google.com/macros/s/XXXX/exec`).
5. Test it in a browser: `…/exec?action=ping&key=YOUR_KEY` should return `{"success":true}`.

> After any change to Code.gs, go to Deploy ▸ Manage deployments ▸ ✏ Edit ▸ Version: **New version**. This keeps the same URL.

## Step 2: Configure the front-end

In `docs/index.html`, find this line:
```js
const API_URL = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
```
Replace the text with your `/exec` URL.

## Step 3: Publish on GitHub Pages

1. Create a GitHub repo, for example `aipd`. It can be under your `chittoorsmartpolicing` account.
2. Upload the `docs/` folder. The `apps-script/` folder is optional.
3. Go to Settings ▸ Pages ▸ Source: **Deploy from a branch** ▸ Branch: `main`, folder: **/docs** ▸ Save.
4. After about a minute the site is live at `https://<username>.github.io/aipd/`.

## Security notes

- The page asks for the access key, checks it on the server, and keeps it only for that browser tab (sessionStorage). Click 🔒 **Lock** to clear it.
- **Never put the key in index.html.** GitHub Pages sites are public.
- To revoke access, change `ACCESS_KEY` in Script Properties. No redeploy is needed.
- The page has `noindex` so search engines skip it.
