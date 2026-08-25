# Gorelo Ticket — Outlook Add-in

Create [Gorelo](https://gorelo.io) tickets directly from Outlook emails — one click for quick tickets, or a full task pane for complex ones.

Works across **Outlook Classic (Windows)**, **Outlook on the web**, **Outlook for Mac**, and the **new Outlook**.

---

## Features

- **Quick Create** — one-click ribbon button creates a ticket instantly using your default settings
- **Task Pane** — full form with client, contact, group, assignee, priority, tags, assets, and notes
- **Sender preservation** — original sender's email and name are embedded in the ticket so Gorelo can auto-associate it with the right client/contact
- **Auto-match** — if the sender's email matches a Gorelo contact, the client and contact fields pre-populate automatically
- **Settings** — configurable defaults for group, assignee, priority, and inbound mailbox

---

## How it works

This add-in uses Gorelo's internal API gateway (`gw.usw.gorelo.tech`) and authenticates via the same Azure AD B2C flow that Gorelo's web app uses. No separate API key is needed — users sign in with their existing Gorelo credentials.

> **Note:** Gorelo has not officially documented these API endpoints. They have been reverse-engineered from browser traffic and may change without notice. If the add-in stops working after a Gorelo update, please open an issue.

---

## Setup

### 1 — Host the files

The add-in is static files — no server-side code needed. Host them anywhere with HTTPS:

- [GitHub Pages](https://pages.github.com/) (free, recommended for testing)
- Azure Static Web Apps
- Cloudflare Pages
- Any web server

After hosting, note your URL (e.g. `https://yourusername.github.io/gorelo-ticket`).

### 2 — Update the manifest

Open `manifest.xml` and replace every instance of `REPLACE_WITH_YOUR_HOSTING_URL` with your hosting URL.

Also replace `REPLACE_WITH_YOUR_GITHUB` in the SupportUrl with your GitHub username.

Generate a new unique GUID for the `<Id>` field: https://guidgenerator.com

### 3 — Deploy to Outlook

**Option A — Sideload for testing (personal use)**

Follow Microsoft's guide to sideload the manifest:
- [Outlook on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing)
- [Outlook Classic (Windows)](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing#sideload-manually)

**Option B — Deploy via M365 Admin Centre (organisation-wide)**

1. Go to [M365 Admin Centre](https://admin.microsoft.com) → Settings → Integrated apps
2. Upload your `manifest.xml`
3. Assign to users or groups

### 4 — First use

1. Open an email in Outlook
2. You'll see a **Gorelo Ticket** group in the ribbon with two buttons:
   - **Quick Create** — creates a ticket immediately with defaults
   - **Create Ticket…** — opens the task pane for full options
3. On first use, a sign-in popup will appear — sign in with your Gorelo credentials (Microsoft/Azure AD login)
4. Open **Create Ticket…** → ⚙ Settings to configure your defaults

---

## Configuration (Settings pane)

| Setting | Description |
|---|---|
| Default Inbound Mailbox | Which Gorelo email configuration to attach tickets to |
| Default Group | Default technician group (required for ticket creation) |
| Default Assignee | Who tickets are assigned to by default (optional) |
| Default Priority | Normal by default |
| Always open task pane | Disables Quick Create — always shows full form |

---

## Deployment for Gorelo partners

Each MSP deploying this add-in needs their own:
1. Hosted copy of these files (or fork this repo and use GitHub Pages)
2. Updated `manifest.xml` with their hosting URL and a new GUID
3. M365 admin access to deploy the manifest

No Azure app registration is needed — authentication uses Gorelo's own Azure AD B2C application.

---

## Troubleshooting

**"Authentication failed" on sign-in**
- Ensure you're signing in with the same Microsoft account you use for Gorelo
- If you use Gorelo's native (non-Microsoft) login, authentication will not work with this add-in — this is a current limitation

**"Failed to create ticket" errors**
- Check your Default Group is set in Settings — Group is required by Gorelo
- Check your Default Inbound Mailbox is set
- Gorelo's internal API may have changed — check [Issues](../../issues) for updates

**Buttons not appearing in ribbon**
- Ensure the manifest was deployed correctly
- In Outlook Classic, try File → Manage Add-ins to verify it's listed and enabled

---

## Contributing

PRs welcome. Key areas for improvement:
- Support for Gorelo's native (non-Microsoft) login flow
- Asset/uptime association
- Ticket type support once the form loads groups/locations
- Gorelo official API support once ticket creation is added

---

## Licence

MIT — free to use, modify, and distribute. See [LICENSE](LICENSE).

---

## Disclaimer

This project is not affiliated with or endorsed by Gorelo. It uses internal APIs that are not officially documented and may change. Use at your own risk.
