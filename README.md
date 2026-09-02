# HCV Program - Inventory & Asset Management System

This folder is a ready-to-deploy static site. No `npm install`, no build step needed -
just upload it as-is.

## Deploy to Netlify (drag-and-drop, easiest way)

1. Go to https://app.netlify.com/drop
2. Drag this entire folder (or a zip of it) onto the page
3. Netlify gives you a live URL immediately

## Deploy to Netlify (via a site you already have)

1. Netlify dashboard -> your site -> **Deploys** tab
2. Drag this folder onto the deploy area, OR
3. Connect this folder to a GitHub repo and Netlify will publish it automatically
   on every push (no build command needed - publish directory is `.`)

## What's inside

- `index.html` - loads Tailwind (via CDN) and the app bundle, with an import map
  that loads React, Recharts, lucide-react, and xlsx from esm.sh (a CDN) - so
  there is nothing to install.
- `bundle.js` - the compiled app (from `src/App.jsx` and `src/entry.jsx`)
- `src/App.jsx` - the actual application source, in case you want to edit it
- `src/entry.jsx` - mounts the app into the page

## Database

The app is already connected to your Supabase project (URL and anon key are in
`src/App.jsx`). Make sure you've run `supabase-app-state-setup.sql` in your
Supabase SQL editor before using the deployed site, if you haven't already.

## If you edit src/App.jsx

The deployed site uses the pre-built `bundle.js`, not the raw `src/App.jsx`
directly. If you change `src/App.jsx`, you need to rebuild `bundle.js` before
re-deploying. Ask Claude to regenerate it, or if you have Node.js installed:

```
npx esbuild src/entry.jsx --bundle --format=esm --jsx=automatic \
  --external:react --external:react-dom --external:react-dom/client \
  --external:lucide-react --external:recharts --external:xlsx \
  --outfile=bundle.js
```
