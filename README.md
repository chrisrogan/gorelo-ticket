# Gorelo Ticket — Outlook Add-in

Create [Gorelo PSA](https://gorelo.io) tickets directly from emails in Outlook.
Works across **Outlook Classic (Windows)**, **Outlook on the web**, **Outlook for Mac**, and **Outlook Mobile**.

Two modes:

| Button | What it does |
|--------|-------------|
| **Quick Create** | One click — creates a ticket from the current email using your saved defaults |
| **Create Ticket…** | Opens a full form to set client, contact, group, priority, tags and more |

---

## Setup

### 1. Generate a Gorelo API key

1. Log in to Gorelo
2. Go to **Settings → Integrated Apps → API Keys**
3. Click **Create API Key**, give it a name (e.g. *Outlook Add-in*)
4. Copy the key — you'll only see it once

### 2. Host the add-in files

The add-in is a set of static files. You can host them anywhere that serves HTTPS:

**GitHub Pages (recommended)**

```bash
git clone https://github.com/YOUR-ORG/gorelo-ticket.git
cd gorelo-ticket
```

1. Push the repo to GitHub
2. Go to **Settings → Pages** and set Source to `main` branch, `/ (root)` folder
3. Note your Pages URL, e.g. `https://your-org.github.io/gorelo-ticket`

### 3. Update the manifest

Edit `manifest.xml` and replace every occurrence of `REPLACE_WITH_YOUR_HOSTING_URL` with your hosting URL (no trailing slash):

```bash
# macOS / Linux
sed -i '' 's|REPLACE_WITH_YOUR_HOSTING_URL|https://your-org.github.io/gorelo-ticket|g' manifest.xml

# Windows PowerShell
(Get-Content manifest.xml) -replace 'REPLACE_WITH_YOUR_HOSTING_URL','https://your-org.github.io/gorelo-ticket' | Set-Content manifest.xml
```

Commit and push.

### 4. Add the icons

Place PNG icons in the `assets/` folder:

| File | Size |
|------|------|
| `icon-16.png` | 16 × 16 px |
| `icon-32.png` | 32 × 32 px |
| `icon-64.png` | 64 × 64 px |
| `icon-80.png` | 80 × 80 px |
| `icon-128.png` | 128 × 128 px |

Use the Gorelo green `#1DAF94` as the base colour.

### 5. Sideload the manifest

#### Outlook Classic (Windows)
1. Open Outlook → **File → Manage Add-ins** (or **Get Add-ins**)
2. Click **My add-ins → Add a custom add-in → Add from file…**
3. Select `manifest.xml`

#### Outlook on the web
1. Open [outlook.com](https://outlook.com) or your Microsoft 365 web mail
2. Open any email → click **…** → **Get add-ins**
3. **My add-ins → Add a custom add-in → Add from file…** → select `manifest.xml`

#### Admin deployment (for distributing to your whole team)
1. Microsoft 365 admin centre → **Settings → Integrated Apps**
2. Upload `manifest.xml`
3. Assign to users / groups

### 6. Enter your API key in the add-in

1. Open any email in Outlook
2. Click **Create Ticket…** in the Gorelo ribbon group
3. The task pane will ask for your API key — paste it in and click **Validate & Save**
4. Set your preferred defaults in **Settings** (group, priority, etc.)

---

## How it works

- Uses the **Gorelo Public API** (`api.usw.gorelo.io`) with `X-API-Key` authentication
- No Azure Active Directory registration required — just a Gorelo API key
- Settings are stored in **Office Roaming Settings** so they follow you across devices
- The original sender's name and email are preserved in the ticket description

### Quick Create defaults

Quick Create uses the defaults you set in the task pane Settings tab:

| Setting | Default |
|---------|---------|
| Priority | Normal (2) |
| Source | Email (2) |
| Group | (empty — set one in Settings) |
| Assignee | Unassigned |
| Inbound Mailbox | None |

---

## Development

No build step required — this is vanilla JS using ES modules.
To test locally, run any static file server from the project root:

```bash
npx serve .
# or
python3 -m http.server 3000
```

Then update `manifest.xml` with `https://localhost:3000` (requires HTTPS in most hosts — use a local certificate or ngrok).

---

## Gorelo API reference

- Base URL: `https://api.usw.gorelo.io`
- Auth: `X-API-Key: <your key>` header
- Docs: see your Gorelo instance under Settings → API

### Endpoints used

| Endpoint | Purpose |
|----------|---------|
| `GET /v1/clients` | Client list |
| `GET /v1/contacts` | Contact list (filterable by ClientId) |
| `GET /v1/technician-groups` | Group list |
| `GET /v1/technicians` | Technician/assignee list |
| `GET /v1/tags` | Tag list |
| `GET /v1/email-configurations` | Inbound mailboxes |
| `POST /v1/tickets` | Create ticket |

---

## Contributing

PRs welcome. Issues to [github.com/unl3ash3dit/gorelo-ticket/issues](https://github.com/unl3ash3dit/gorelo-ticket/issues).

> **Note:** This is a community project and is not affiliated with or supported by Gorelo.
