# Allocate IPS Monitoring Dashboard

Client- and scheme-level asset-allocation monitoring for the Ionic Allocate PMS,
comparing each client's live allocation (from Portfolio Appraisal exports) against
their assigned Model Master target weights.

- **Tab 1 – Roof-level:** Total Equity, Fixed Income, Listed Shares (Stock Basket),
  Alternate Assets, Cash, Liquid.
- **Tab 2 – Category-level:** Domestic Equity, International Equity, Stock Basket,
  Fixed Income, Commodities, InVITs, Liquid ETFs, Cash.
- Compliant / breach flag per cell using a **relative ±20% of target** band.
- Filters (scheme / model / RM / status / search), breach charts, AUM-weighted
  scheme summary, and a **Download Excel** export button.

The dashboard is a single self-contained `index.html` — all data is embedded inline
and the only external dependencies (Chart.js, SheetJS) load from a public CDN over
HTTPS, so it works on GitHub Pages with no build step.

---

## What to deploy

Push the **contents of this `github-deploy` folder** to a GitHub repository:

```
index.html      <- the dashboard (this is all GitHub Pages needs)
.nojekyll       <- serves files as-is, skips Jekyll processing
README.md       <- this file
source/         <- scripts + data used to generate index.html (optional, for refresh)
```

> Only `index.html` and `.nojekyll` are required for hosting. `source/` is kept for
> reproducibility and is harmless to publish; **omit it if the underlying data is
> confidential** (the embedded data in `index.html` is the same data, so treat the
> whole repo as internal/private if client names should not be public).

---

## Deploy to GitHub Pages (step by step)

### Option A — GitHub website (no command line)
1. Go to https://github.com/new and create a repository, e.g. `allocate-ips-dashboard`.
   Set it to **Private** (recommended — it contains client names).
2. On the new repo page click **uploading an existing file**.
3. Drag in everything from this `github-deploy` folder (`index.html`, `.nojekyll`,
   `README.md`, and the `source` folder). Commit.
4. Go to **Settings → Pages**.
5. Under **Build and deployment → Source** choose **Deploy from a branch**.
6. Select branch **main** and folder **/ (root)**, click **Save**.
7. Wait ~1 minute. Your dashboard will be live at:
   `https://<your-username>.github.io/allocate-ips-dashboard/`

> Note: GitHub Pages on a **Private** repo requires a GitHub **Pro/Team/Enterprise**
> plan. On the free plan, Pages sites are public — so either upgrade, or only publish
> if you are comfortable with the data being public. (See "Keeping data private" below.)

### Option B — Git command line
```bash
cd "github-deploy"
git init
git add .
git commit -m "Allocate IPS dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/allocate-ips-dashboard.git
git push -u origin main
```
Then enable Pages via **Settings → Pages → Deploy from a branch → main → / (root)**.

---

## Refreshing the dashboard with new data

When you have new Portfolio Appraisal / Client Repository / Model Master files:

1. Run `source/build_ips.py` (edit the file paths at the top to point at the new
   exports). It re-parses the appraisals, classifies every holding, joins the model,
   applies the ±20% breach rule and writes `ips_data.json`.
2. Run `source/make_html.py` to embed the fresh `ips_data.json` into a new
   `index.html`.
3. Commit and push the updated `index.html` — Pages redeploys automatically.

---

## Keeping data private

This dashboard shows real client names and allocations. If it must not be public:

- Use a **Private** repo with a paid GitHub plan (Pages stays access-controlled), **or**
- Host on an internal/intranet static host instead of public GitHub Pages, **or**
- Keep the repo private and share the `index.html` file directly (it works by simply
  opening it in any browser — no server required).

---

*Data as of 16/06/2026. Breach band: relative ±20% of model target.
Coverage: all 304 accounts are scored against their scheme's model targets (asset-class
targets are identical across model variants — Aug / Feb / V2-Nov). Accounts matched in
the Client Repository carry their exact assigned model and RM; the rest are scored by
scheme default until the Client Repository is re-synced.*
