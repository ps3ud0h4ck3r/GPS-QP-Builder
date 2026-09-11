# GPS QP Builder v7

Final cleanup pass: the top navigation contains no left-side logo, PDF/DOCX filenames use `GPS-Class-Subject`, and direct sharing controls are removed.

All V6 features are retained, including blueprint/marks distribution, A4 page-break controls, language selection for English/Hindi/Sanskrit, rich formatting, drag-and-drop ordering, duplicate sections, PDF/DOCX export, dark mode, and preview zoom.
# GPS QP Builder

Mobile-friendly question paper builder for Goodwill Public School.

## Run locally

```bash
npm install
npm run dev
```

Open the local Vite URL shown in the terminal.

## Build

```bash
npm run build
```

Vite creates the production site in `dist`.

## Cloudflare Pages

- Connect this GitHub repository to Cloudflare Pages.
- Framework preset: Vite
- Build command: `npm run build`
- Output directory: `dist`

### AI setup

Add these Cloudflare environment variables:

- `ANTHROPIC_API_KEY` = your Anthropic API key
- `ANTHROPIC_MODEL` = optional model override

The API key is used only by the Cloudflare Function at `/api/generate`. Never put the key in browser code or GitHub.

The built-in 10/day browser counter is only an accidental-use guard, not real server-side security. For a large public rollout, add server-side rate limiting/authentication or disable public AI generation.

## Included features

- Question paper builder with local drafts
- Mobile layout and PWA install support
- GPS logo/icon branding
- Direct PDF export
- DOCX export
- WhatsApp, Email and native PDF sharing
- Max marks protection
- Section drag-and-drop and question drag-and-drop, with move buttons as fallback
- Duplicate sections
- Rich question editor with bold, italic, superscript, subscript, Greek symbols and tables
- Dark mode
- Print font size selector
- Preview zoom
- Diagram upload, drawing and reserved diagram space

## Data/privacy

Drafts and uploaded paper images are kept in the teacher's browser/local storage. They are not automatically stored in a database.

## V5 logo fix
The GPS/Goodwill logo used by the app UI is imported from `src/gps-logo.png` so Vite bundles it into the application. This avoids broken image paths when deploying the app through a Worker setup that does not expose arbitrary root files the way Pages does.
