# WSE Piotroski F-Score — Survivorship Bias Analysis

**MScFE 690 Capstone** — Stanisław Halkiewicz, Orjika Blessing, Dmitrii Verdun
April 2026

---

## Overview

This project produces a survivorship-bias-corrected empirical evaluation of the
Piotroski F-Score on the Warsaw Stock Exchange (WSE) pilot sample (2010–2024).
It tests two hypotheses:

- **H1 (Bias):** F-Score long-short returns computed on survivorship-biased
  samples significantly overstate true historical performance.
- **H2 (Alpha):** After correcting for survivorship bias, a statistically
  significant Piotroski alpha still exists.

---

## Repository Structure

```
WQU_capstone/
├── pipeline.py                        Main analysis pipeline  <-- RUN THIS
├── scrape_biznesradar_production.py   Web scraper (pre-run; do not re-run
│                                      unless you want to refresh the data)
├── README.md                          This file
├── CHANGELOG.md                       Summary of fixes and assumptions
└── output/
    ├── cache/                         Per-ticker JSON cache (scraper output)
    ├── financials_raw_production.csv  Raw scraped financials
    ├── wse_piotroski_production.csv   Pre-computed Piotroski F-Score panel
    ├── wse_piotroski_production.xlsx  Excel copy of above
    ├── figures/                       Generated PNG figures (300 DPI)
    └── results/                       Generated CSV tables
```

---

## Required Packages

```bash
pip install pandas numpy matplotlib scipy yfinance openpyxl
```

| Package    | Role                                      | Version tested |
|------------|-------------------------------------------|----------------|
| pandas     | Data manipulation                         | ≥ 1.5          |
| numpy      | Numerical computing                       | ≥ 1.23         |
| matplotlib | Figures                                   | ≥ 3.5          |
| scipy      | t-tests                                   | ≥ 1.9          |
| yfinance   | Fetching WSE stock returns (internet req) | ≥ 0.2          |
| openpyxl   | Excel output (optional)                   | ≥ 3.0          |

---

## Input Files

| File | Description | Location |
|------|-------------|----------|
| `wse_piotroski_production.csv` | Pre-computed Piotroski F-Score panel (11 firms, 2010–2024) | `output/` |

The panel was produced by `scrape_biznesradar_production.py`, which scrapes
financial statements from biznesradar.pl and computes the nine Piotroski signals.

**Note on F-Score ceiling:** `csho` (shares outstanding) was unavailable from
the data source; `F_EQ_OFFER` is therefore excluded from all computations.
The maximum achievable F-Score is **8** (not 9).

---

## Order of Execution

1. **(Already done)** `scrape_biznesradar_production.py` — scrapes data and
   writes `wse_piotroski_production.csv`.  Do not re-run unless you need
   fresh data (requires Chrome and an internet connection).

2. **Run the pipeline:**
   ```bash
   cd path/to/WQU_capstone
   python pipeline.py
   ```
   Requires an internet connection to fetch stock returns from Yahoo Finance
   via `yfinance`.

---

## Output Files

### Figures (`output/figures/`)

| File | Description |
|------|-------------|
| `fig1_fscore_distribution.png` | F-Score distribution: active vs delisted firms |
| `fig2_annual_ls_returns.png` | Annual long-short returns: biased vs corrected |
| `fig3_cumulative_returns.png` | Cumulative long-short returns: biased vs corrected |
| `fig2_annual_ls_returns_strict.png` | As above, strict thresholds (F≥8/F≤2) |
| `fig3_cumulative_returns_strict.png` | As above, strict thresholds |

### Tables (`output/results/`)

| File | Description |
|------|-------------|
| `summary_stats.csv` | Mean, std, Sharpe, min, max for biased and corrected L/S |
| `hypothesis_tests.csv` | H1 and H2 t-statistics and p-values |
| `summary_stats_strict.csv` | As above, strict thresholds |
| `hypothesis_tests_strict.csv` | As above, strict thresholds |
| `portfolio_detail.csv` | Year-by-year portfolio composition (which firms, how many) |
| `portfolio_detail_strict.csv` | As above, strict thresholds |
| `fscore_by_survival_status.csv` | Cross-tabulation of last F-Score by survival status |
| `panel_with_returns.csv` | Full enriched panel (F-Scores + fetched returns, for audit) |

---

## Methodology Notes

### Portfolio formation thresholds

The pilot universe contains 11 firms (8 active, 3 delisted).  With the
canonical Piotroski thresholds (F≥8, F≤2) most years have 0–1 firms per
portfolio leg, making the test degenerate.  We therefore use:

- **Primary analysis:** Long F≥7, Short F≤3
- **Robustness check:** Long F≥8, Short F≤2 (original Piotroski thresholds)

This is consistent with the proposal's Section 3.3, which explicitly anticipates
relaxing thresholds for small CEE samples.

### Return assignment

Annual returns are calendar-year total returns (adjusted for dividends and
splits) fetched from Yahoo Finance via `yfinance`.  Year-end price is the last
available trading day in each December.

All three delisted firms in this pilot (Arteria SA, Bogdanka SA, Play
Communications) were acquired at a premium, not bankrupt.  We therefore use
the actual market return in the delisting year rather than a forced –100%
terminal return.  The proposal's –100% assumption applies to outright
bankruptcies, which are not represented in this pilot sample.

### Survivorship bias methodology

- **Biased sample:** Only firms that are marked `status = "active"` in the
  production CSV are included.
- **Corrected sample:** All firms (active and delisted) are included.  A
  delisted firm contributes its actual return in the delisting year and
  disappears from subsequent years.

---

## Limitations of the Pilot

1. **Small sample (11 firms):** Statistical tests are underpowered.  Results
   are preliminary.  The full project will cover all WSE/PSE/BSE/BVB firms
   via WRDS Compustat Global.

2. **F_EQ_OFFER missing:** The shares-outstanding signal cannot be computed
   from the available source.  F-Scores range 0–8.

3. **No delisted-bankrupt firms:** All delisted firms in the pilot were
   acquired.  The hypothesis-driven asymmetry (low-F-Score firms fail) is
   present but muted relative to a full-market sample with genuine bankruptcies.

4. **ENG.WA (Energa) data note:** Energa was acquired by PKN Orlen and
   technically delisted in 2020; the source data (biznesradar.pl) continues
   to publish consolidated group financial reports post-acquisition, which
   is why it appears with `status = "active"` through 2024.  This is treated
   as-is but acknowledged as a caveat.

5. **Fama-French factors:** Stage 4 of the full proposal requires F-F3 factor
   regressions.  These are not implemented in the pilot because WSE-specific
   factor series require Compustat Global data not yet available.  The pilot
   uses raw L/S returns and t-tests as a proxy.
