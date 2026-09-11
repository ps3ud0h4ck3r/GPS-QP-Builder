# Question Paper Builder

A mobile-first React/Vite web app for teachers to create, preview, print, and download question papers. It supports multiple question types, diagrams, local drafts, PWA installation, and optional AI question generation.

## Included

- Mobile-friendly editor
- Paper details and school logo
- Fill in the blanks, MCQ (3/4 options), True/False, One Word, Short Answer (2/3 marks), Long Answer, Match the Following, Numerical, and Diagram Based questions
- Upload, draw, or reserve space for diagrams
- Automatic marks calculation and validation display
- Local draft save/load/delete
- Preview, browser Print, and direct PDF download
- PWA manifest, icons, and service worker
- AI generation through a server-side Cloudflare Function, never directly from the browser
- A small client-side AI usage guard of 10 generations/day to reduce accidental usage. This is not a security boundary.

## Local setup

1. Install Node.js LTS.
2. Open a terminal in this folder.
3. Run:

   npm install
   npm run dev

4. Open the local URL shown by Vite.

## Cloudflare Pages deployment

1. Put the entire project in a GitHub repository.
2. In Cloudflare, create a Pages project and connect the repository.
3. Framework preset: Vite.
4. Build command: `npm run build`.
5. Build output directory: `dist`.
6. Deploy.
7. Share the resulting `*.pages.dev` URL with teachers.

The `functions/api/generate.js` file is deployed as the `/api/generate` server-side endpoint.

## AI setup

AI generation is optional. The app works without it.

In the Cloudflare project, add:

`ANTHROPIC_API_KEY` = your Anthropic API key

Optional:

`ANTHROPIC_MODEL` = the model name you want to use

Never put the API key in `src/App.jsx`, `index.html`, or any other browser-side file.

### Important public-app warning

The site is intentionally usable without teacher accounts. That is ideal for the free V1, but a public AI endpoint can still be abused by someone who discovers it. The included 10/day browser limit only reduces accidental use and is not a real security control. For a large public rollout, either disable AI or add proper server-side authentication/rate limiting before sharing the AI-enabled URL widely.

## PDF

“Download PDF” uses `html2pdf.js` in the browser. No PDF server is required. “Print” opens the browser's print flow.

## PWA

On supported mobile browsers, teachers can use the browser's “Add to Home Screen” option. The app includes a web manifest, PNG icons, and a service worker for an app-like experience and basic asset caching.

## Data and privacy

Drafts and the selected school logo are stored in the teacher's browser using local storage. They are not automatically uploaded to a database. Clearing browser/site data can remove local drafts.
