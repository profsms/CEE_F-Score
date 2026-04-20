# Changelog / Summary of Fixes

**MScFE 690 Capstone — pipeline.py**
April 2026

---

## Problems Found in the Original Codebase

### 1. No end-to-end pipeline on real data
The repository had two disconnected pieces:

- `scrape_biznesradar_production.py` — a Selenium scraper that writes
  `wse_piotroski_production.csv` to `output/`.
- `Downloads/Capstone_research_proposal.ipynb` — an exploratory notebook
  that loads the production CSV but cannot run because `ret` (stock returns)
  is empty.
- `Downloads/preliminary_results.py` — a clean analysis script, but uses
  **synthetic data only** (not the production CSV).

There was no single script that ran the full pipeline on the real data.

### 2. Returns column empty
The `ret` and `dlret` columns in `wse_piotroski_production.csv` were almost
entirely blank.  The notebook's `build_returns()` and portfolio construction
depended on `ret`; without it, `final_ret` was NaN for all firms, so no
portfolios could be formed.

**Fix:** `pipeline.py` fetches calendar-year returns from Yahoo Finance via
`yfinance` for all 11 tickers and attaches them to the panel.

### 3. `csho` (shares outstanding) missing
The `csho` column is blank in the production data (biznesradar.pl does not
expose shares outstanding in a machine-readable table).  The `F_EQ_OFFER`
signal (which checks whether new shares were issued) cannot be computed.

**Assumption made:** We use the 8-signal F-Score already computed by the
scraper (`N_SIGNALS = 8` rows).  The maximum F-Score is 8, not 9.  Portfolio
thresholds are adjusted accordingly (see point 5 below).

### 4. Portfolio thresholds too strict for a pilot of 11 firms
The canonical thresholds F≥8 (long) and F≤2 (short) produce 0–1 firms per
leg in most years of the pilot, making statistical tests degenerate.

**Fix:** Primary analysis uses **F≥7 / F≤3** (relaxed thresholds), consistent
with the proposal's Section 3.3 which explicitly anticipated this.  The
original F≥8 / F≤2 thresholds are run as a robustness check and saved with
the `_strict` filename suffix.

### 5. Notebook re-derived F-Score from raw columns instead of using pre-computed
The notebook's `compute_fscore()` function re-computed all nine signals from
`ni`, `at`, `cfo`, etc.  This was unnecessary and failed on rows where raw
columns were sparse (introducing NaN cascades).

**Fix:** `pipeline.py` uses the pre-computed `FSCORE` column from the
production CSV directly, retaining only rows where `N_SIGNALS >= 7`.

### 6. N_SIGNALS == 0 rows (placeholder years) were included
The production CSV contains placeholder rows for years in which a firm had no
financial data (all financial columns are blank but the firm-year row exists).
The notebook's `load_real_data()` filtered these out via a heuristic
`(status == "inactive") & (ni == 0) & (at == 0)` which missed some cases.

**Fix:** We filter `N_SIGNALS >= MIN_SIGNALS` (default 7).  Rows with
`N_SIGNALS == 0` (no data) and `N_SIGNALS == 3` (first observable year,
only non-lagged signals) are both excluded.

### 7. Survivorship-bias sample logic
The notebook filtered biased sample with:
```python
survivors = last_delist[last_delist.isna() | (last_delist > max_year)].index
```
This is correct in spirit but fragile: it depended on `delist_year` being
properly populated, which it sometimes wasn't.

**Fix:** `pipeline.py` filters on `status == "active"` for the biased sample,
which is set directly by the scraper and is reliable.

---

## Assumptions Made

| Assumption | Justification |
|---|---|
| Max F-Score = 8 (not 9) | `csho` unavailable; F_EQ_OFFER excluded |
| Primary thresholds: F≥7 long, F≤3 short | Only 11 firms; strict thresholds produce empty legs |
| Calendar-year returns from yfinance | No return data in production CSV; yfinance `.WA` tickers are supported |
| No –100% terminal return for delisted firms | All 3 delisted firms (Arteria, Bogdanka, Play) were acquired, not bankrupt |
| ENG.WA treated as "active" | Biznesradar publishes Energa data post-acquisition; treating it differently would require manual correction |
| `MIN_SIGNALS = 7` filter | Ensures at least 7 of 8 signals were computable; excludes placeholder rows |
| Random seed = 42 | For reproducibility of any stochastic elements (none in the current pipeline, but set for forward-compatibility) |

---

## Files Changed / Created

| File | Action |
|---|---|
| `pipeline.py` | **Created** — main analysis script |
| `README.md` | **Created** — usage and documentation |
| `CHANGELOG.md` | **Created** — this file |
| `scrape_biznesradar_production.py` | Unchanged |
| `output/wse_piotroski_production.csv` | Unchanged (read-only input) |
| `Downloads/preliminary_results.py` | Unchanged (kept as synthetic-data pilot) |
| `Downloads/Capstone_research_proposal.ipynb` | Unchanged (kept as exploratory notebook) |
