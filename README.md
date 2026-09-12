# GPS QP Builder v9

School-wide Question Paper Builder for Goodwill Public School.

## Included

- Individual teacher/admin accounts
- Secure cookie sessions and role-based access
- Admin user management
- Admin impersonation with a clear return-to-admin control
- Account details and password changes
- Server-saved personal drafts, templates and settings
- Automatic marks blueprint and max-marks protection
- Section and question drag-and-drop ordering
- Better question numbering
- A4 page-aware preview and page-break controls
- English, Hindi and Sanskrit paper labels
- Rich formatting: bold, italic, underline, alignment, lists, superscript, subscript, symbols and tables
- Direct A4 PDF export
- DOCX export
- Dark mode and print font-size selector
- GPS branding
- No AI generation or AI API dependency
- No WhatsApp, Email or direct sharing controls

## Cloudflare setup

This version uses Cloudflare Workers + D1 for authentication and per-user data.

### 1. Create a D1 database

Create a Cloudflare D1 database named `gps-qp-builder`.

Run `migrations/0001_init.sql` against that database.

### 2. Bind the database

In the Worker settings, add a D1 binding:

- Variable name: `DB`
- Database: your `gps-qp-builder` D1 database

The Worker already expects the static asset binding named `ASSETS` from `wrangler.toml`.

### 3. Add the first-time setup secret

Create a secret environment variable:

- `GPS_SETUP_KEY` = a long random secret you choose

On first launch, when the database has no users, GPS QP Builder shows a School Setup screen. Enter that setup key and create the first admin accounts. After that, admins can create the remaining teacher and admin accounts from the Admin Panel.

### 4. Deploy

Build command:

```bash
npm run build
```

Deploy command:

```bash
npx wrangler deploy
```

Root directory: `/`

The included `wrangler.toml` points the Worker at `worker/index.js` and serves the Vite `dist` directory as static assets.

## Account model

Teachers have their own drafts, templates and settings. Admins can manage users and temporarily log in as another user without learning that user's password. Impersonation creates a separate short-lived session and is logged in the audit table.

For the initial rollout, create your three admin accounts in School Setup, then use Admin Panel to bulk-create the 30 teachers. The bulk format is one line per teacher:

```text
Name, username, temporaryPassword, email, teacher
```

Newly created users are flagged to change their password on first login. Passwords are stored as PBKDF2-SHA-256 hashes, never as plain text.

## Data

Workspace data is stored per user in D1. This means a teacher can sign in from another phone or computer and still access their saved work.

## AI

AI generation has been completely removed from the app, UI, package and backend.
