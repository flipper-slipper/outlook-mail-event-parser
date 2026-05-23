# Outlook Calendar Event Extractor

An Outlook add-in that reads an email, sends it to OpenAI GPT-4o, extracts all calendar events, and lets you add them to your Outlook calendar with one click.

## Setup

### 1. Generate SSL certificates (required by Office add-ins)

Office add-ins must be served over HTTPS. Run:

```powershell
npm run setup-certs
```

If that prints an `openssl` command instead, run that command in your terminal (Git Bash or WSL both work), then come back here.

After generation, **trust the certificate in your browser**:
1. Open https://localhost:3000 in Edge or Chrome
2. Click "Advanced" → "Proceed to localhost"

### 2. Start the server

```powershell
npm start
```

### 3. Sideload the add-in in Outlook

#### Outlook on the Web (recommended for testing)
1. Open [Outlook Web](https://outlook.live.com) or your org's OWA
2. Open any email → click the **"…"** (More actions) menu → **Get Add-ins**
3. Go to **My add-ins** → **Add a custom add-in** → **Add from file…**
4. Upload `manifest.xml`

#### Outlook Desktop (Windows)
1. Open Outlook → **File** → **Manage Add-ins** (opens OWA)
2. Follow the same steps as above

### 4. Use the add-in
1. Open any email
2. Click **Extract Events** in the ribbon (or from the "…" menu on web)
3. Enter your OpenAI API key and click **Save** (stored in localStorage, never sent anywhere except OpenAI)
4. Click **Extract Events from Email**
5. Review the detected events and click **Add to Calendar** on any you want

## File structure

```
outlook-calendar-extension/
├── manifest.xml        # Add-in registration for Office
├── taskpane.html       # Task pane UI
├── taskpane.css        # Styles
├── taskpane.js         # Logic: reads email, calls OpenAI, renders events
├── commands.html       # Required stub for ribbon button
├── server.js           # Local HTTPS server
├── scripts/
│   └── generate-certs.js
└── certs/              # Created by setup-certs (git-ignored)
    ├── server.key
    └── server.crt
```

## Notes

- Your OpenAI API key is stored only in your browser's `localStorage` and is sent directly from your browser to `api.openai.com`. It is never stored on any server.
- The add-in reads email body text only — no attachments.
- `gpt-4o` is used by default. Change the `model` field in `taskpane.js` if you prefer a different model.
