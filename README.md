# tpaulshippy.com

Paul Shippy's personal website — a static site with projects, links, writings,
and contact info.

## Contents

- `index.html` / `style.css` — the whole site (one page, no build step)
- `tpauls.htm` — legacy page, kept for old links
- `IMG_0568.JPG`, `E4oGfFRF_400x400.jpg`, `favicon.ico` — photos and icon
- `RedirectDefault.aspx`, `web.config` — legacy IIS bits
- `preview/` — design mockups (not served)
- `.github/workflows/deploy.yml` — deploy pipeline (see below)

## Preview locally

No toolchain needed — serve the folder statically, e.g.:

```sh
python3 -m http.server 8000
# open http://localhost:8000/
```

## Deploy

Pushing to `master` runs the `Deploy` workflow: it SSHes into the web server,
`git pull`s the site, and smoke-checks that https://tpaulshippy.com/ serves the
expected marker. The workflow can also be run by hand from the Actions tab
(`workflow_dispatch`). Branches other than `master` never deploy.
