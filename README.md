# tpaulshippy.com

Paul Shippy's Personal Website, hosted as a **Cloudflare Worker with static
assets** at <https://tpaulshippy.com>.

This repo holds the static single-page homepage for tpaulshippy.com, the personal site of Paul Shippy — a Christian software developer and homeschooling dad. It features his projects, favorite links, and writings alongside links to his blog and GitHub, built as a self-contained, mobile-friendly page with system-aware dark mode and no runtime dependencies.

## How it works

- The site is plain static files in the repo root (`index.html`, `style.css`,
  writings like `McCain.html`, images, `resume.pdf`).
- `wrangler.toml` deploys them as Workers Static Assets (`directory = "./"`).
  `.assetsignore` keeps the legacy server-side app directories (classic
  ASP.NET/PHP, ~2 GB — they can't run on Workers) out of the upload.
- `src/index.js` is a tiny Worker that runs first (`run_worker_first`):
  canonical `www` → apex redirect, `/index.html` → `/`, the legacy
  `/mccain.html` → `/McCain.html` alias, plus security and caching headers.
- Pushes to `master` deploy via `.github/workflows/deploy.yml`, which then
  smoke-checks the `workers.dev` URL (home page, an essay, 404s, redirects).

## Local preview

```sh
npx wrangler@4.153.0 dev   # serves the worker + assets locally
```

## One-time Cloudflare setup

1. Add `tpaulshippy.com` to Cloudflare and change the domain's nameservers
   at the registrar to the ones Cloudflare assigns.
2. Create GitHub Actions secrets for this repo:
   - `CLOUDFLARE_API_TOKEN` — API token with **Workers Scripts: Edit** and
     **Workers Routes: Edit** (or the Workers template token).
   - `CLOUDFLARE_ACCOUNT_ID` — from the Cloudflare dashboard URL / sidebar.
3. Push to `master` (or run the Deploy workflow manually) and note the
   `*.workers.dev` URL in the workflow log.
4. In the Cloudflare dashboard go to the `tpaulshippy-com` Worker →
   **Settings → Domains & Routes → Add Custom Domain**, and add
   `tpaulshippy.com` (repeat for `www.tpaulshippy.com`; the Worker redirects
   it to the apex).
5. Recommended dashboard settings for the zone: **SSL/TLS → Full (strict)**,
   **Always Use HTTPS → On**, and enable **HSTS**.
6. Verify `https://tpaulshippy.com/` and a legacy essay
   (e.g. `/McCain.html`), then decommission the old DigitalOcean droplet and
   delete the now-unused `SERVER_HOST` / `SSH_KEY` repo secrets.

## Notes

- Legacy `FormProcessor/` (server-side form handler) is intentionally not
  deployed — it can't run on Workers. The homepage doesn't reference it.
- `RedirectDefault.aspx` / `web.config` are IIS leftovers kept in git for
  reference; the Worker replaces their behavior (directory index, 404 page).
- The old Google Analytics `UA-` tag in `index.html` is dead (Universal
  Analytics was shut down) — candidate for removal or a GA4 replacement.
