# Inbox Triage — Outlook Add-in

An Outlook task pane add-in that scores your unread emails by priority,
shows a ranked reply list, and marks low-priority emails as read.
Works on Outlook for Mac, Windows, and Outlook on the web.

---

## Files

```
outlook-triage/
├── manifest.xml      ← Add-in manifest (edit URLs before sideloading)
├── taskpane.html     ← Task pane UI
├── taskpane.js       ← Scoring engine + EWS communication
├── commands.html     ← Required stub file
└── README.md
```

---

## Step 1 — Host the files (free, ~5 minutes)

The add-in files must be served over HTTPS. The easiest free option
is GitHub Pages:

1. Create a free GitHub account at https://github.com if you don't
   have one.
2. Create a new **public** repository called `outlook-triage`.
3. Upload all four files (manifest.xml, taskpane.html, taskpane.js,
   commands.html) to the repository.
4. Go to the repository Settings → Pages → set Source to
   "Deploy from branch" → branch: main → folder: / (root) → Save.
5. GitHub will give you a URL like:
      https://YOUR-USERNAME.github.io/outlook-triage/
   It may take 1–2 minutes to go live.

> **Icon files**: The manifest references icon-16.png, icon-32.png,
> icon-64.png, icon-80.png, and icon-128.png. You can upload any small
> PNG images with those names, or just delete the icon references from
> manifest.xml — Outlook will use a default icon.

---

## Step 2 — Edit manifest.xml

Open manifest.xml and replace **every** occurrence of:
```
https://YOUR-GITHUB-USERNAME.github.io/outlook-triage
```
with your actual GitHub Pages URL, e.g.:
```
https://jsmith.github.io/outlook-triage
```

There are 8 places to update. Save the file and re-upload it to GitHub.

---

## Step 3 — Sideload the add-in into Outlook for Mac

1. Open Outlook for Mac.
2. Open any email message (double-click to open in its own window).
3. In the message window, click the **•••** (More actions) button in
   the toolbar, or look for **Get Add-ins** in the ribbon.
4. In the Add-ins dialog, click **My add-ins** in the left sidebar.
5. Scroll to the bottom and click **+ Add a custom add-in →
   Add from file…**
6. Select your local `manifest.xml` file and click Open.
7. Confirm any security prompts.

The "Triage Inbox" button will now appear in the toolbar when reading
any email. Click it to open the task pane.

---

## Using the add-in

**VIP senders** — comma-separated email addresses or partial strings
(e.g. `@yourcompany.com`) whose mail always scores HIGH.

**Noise senders** — partial strings that trigger auto-mark-as-read
(e.g. `noreply@`, `newsletter@`).

**Look back** — how many days of unread mail to scan (1–30).

| Button | What it does |
|---|---|
| 🔍 Dry run | Scores everything and shows the report, but changes nothing |
| ⚡ Triage | Scores, shows report, AND marks low-priority as read |
| ✓ Mark low-priority as read | Silently marks noise without showing report |

---

## Permissions note

The manifest requests `ReadWriteMailbox` permission. This is required
to both read inbox messages and mark them as read using the Exchange
Web Services (EWS) API. No data leaves your mailbox — all processing
happens locally in the task pane.

---

## Troubleshooting

**"Add-in could not be loaded"** — Make sure your GitHub Pages site is
live and the URLs in manifest.xml exactly match.

**"This add-in is not supported"** — You need an Exchange or Microsoft
365 mailbox. IMAP-only accounts (e.g. Gmail added via IMAP) won't have
EWS access.

**Empty results / EWS errors** — Your organisation's Exchange admin
may have restricted EWS access for add-ins. Contact IT.
