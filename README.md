# Stockroom

Inventory, sales and returns ledger — originally a Claude.ai artifact,
now a standalone web app.

## What changed from the artifact

- `src/App.jsx` is your original component, unedited.
- `src/storage.js` now polyfills `window.storage` using **Supabase**
  (a free cloud database) instead of the browser's `localStorage`.
  This means everyone who opens the deployed app link — on any phone,
  tablet, or laptop — sees and edits the *same* inventory, invoices,
  returns, and login data in real time. Nothing is tied to one device
  anymore.
- Real npm packages (`react`, `xlsx`, `lucide-react`,
  `@supabase/supabase-js`) replace what Claude.ai injected
  automatically. `jsPDF` still loads from a CDN at runtime exactly as
  it did before — no change needed there.
- Added a PWA manifest + icon so the site can be "installed" to a
  phone or laptop home screen.

## Important: this app now depends on your Supabase project

`src/storage.js` is pre-configured with your Supabase Project URL and
public API key. As long as that Supabase project exists and its
`stockroom_kv` table is set up (see the SQL script you ran during
setup), the app will keep working for everyone using it.

If you ever need to change Supabase projects, just edit the two
constants (`SUPABASE_URL` and `SUPABASE_ANON_KEY`) near the top of
`src/storage.js`, then rebuild and redeploy.

## Run it locally

```bash
npm install
npm run dev
```

Opens at http://localhost:5173

## Deploy it online (free, ~2 minutes)

Easiest path — no server to manage:

1. Push this folder to a new GitHub repo.
2. Go to https://vercel.com (or https://netlify.com), sign in with
   GitHub, and "Import" the repo.
3. Framework preset: **Vite**. Build command: `npm run build`.
   Output directory: `dist`. Click Deploy.
4. You get a live URL (e.g. `stockroom.vercel.app`) immediately, and
   it redeploys automatically every time you push to GitHub.

You can also drag-and-drop the `dist/` folder (after running
`npm run build`) straight into https://app.netlify.com/drop for an
instant one-off deploy with no GitHub needed.

## Turning it into an installable "app"

Once it's live at a URL:

- **Phone home-screen app (no app store needed):** open the site in
  Chrome/Safari on a phone → "Add to Home Screen". Because of the PWA
  manifest already included, it installs with its own icon and opens
  without browser chrome, like a real app.
- **iOS/Android app-store app:** wrap the deployed site with
  [Capacitor](https://capacitorjs.com/) — it packages this same React
  app into a native iOS/Android project you can submit to the stores,
  with no app-code rewrite required.

## Data storage

Data is stored in your Supabase project's `stockroom_kv` table, shared
across every device and person using the app's live URL. There's one
honest trade-off worth knowing: the current setup lets anyone who has
your Supabase Project URL and API key read/write data directly,
bypassing the app's own login screen. That's a reasonable trade for a
small internal tool, but if it ever needs to be locked down further
(e.g. real per-user accounts with row-level permissions), that's a
follow-up upgrade to the Supabase policies — `App.jsx` still wouldn't
need to change.
