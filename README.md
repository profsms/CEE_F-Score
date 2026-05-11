# Piotroski F-Score Survivorship Bias Analysis — Visegrad 3

**MScFE 690 Capstone** — Stanisław Halkiewicz, Orjika Blessing, Dmitrii Verdun  
WorldQuant University · April 2026

---

## Overview

This project produces the first survivorship-bias-corrected empirical evaluation
of the Piotroski F-Score across three Central and Eastern European emerging
markets: the Warsaw Stock Exchange (WSE, Poland), the Prague Stock Exchange
(PSE, Czech Republic), and the Budapest Stock Exchange (BSE, Hungary).
The sample covers 31 firms and 301 firm-year observations over 2010–2024.

Two hypotheses are tested:

- **H1 (Bias):** F-Score long-short returns computed on survivorship-biased
  samples significantly overstate true historical performance relative to a
  corrected sample that includes all firms, including those subsequently
  delisted or bankrupt.
- **H2 (Alpha):** After correcting for survivorship bias, a statistically
  significant Piotroski F-Score alpha still exists in at least a subset of
  the CEE markets studied.

---

## Repository Structure

```
WQU_capstone/
├── piotroski_pipeline.py      Cleans data and computes F-Score panel  <-- run first
├── collect_returns.py         Downloads forward returns from Stooq / Yahoo Finance
├── portfolio_analysis.py      Main empirical analysis (portfolios, tests, figures)
├── data_audit.py              Data quality checks (signal coverage, taxonomy)
├── stooq_debug.py             Diagnostic tool for Stooq API connectivity
├── stooq_apikey.txt           YOUR Stooq API key (create this — see below)
├── README.md                  This file
├── data/
│   ├── wse_piotroski_production.xlsx   WSE panel (Poland, 11 firms)
│   ├── pse_piotroski_production.xlsx   PSE panel (Czech Republic, 10 firms)
│   └── bse_piotroski_production.xlsx   BSE panel (Hungary, 10 firms)
│   └── prices/                         Cached price CSVs (auto-created)
└── output/
    ├── tables/                          CSV result tables
    └── figures/                         PNG figures (150 DPI)
```

---

## Setup

### 1. Install required packages

```bash
pip install pandas numpy matplotlib scipy yfinance openpyxl requests
```

| Package      | Role                                          | Version tested |
|--------------|-----------------------------------------------|----------------|
| `pandas`     | Data manipulation and panel construction      | ≥ 2.0          |
| `numpy`      | Numerical computing                           | ≥ 1.23         |
| `matplotlib` | Figure generation                             | ≥ 3.5          |
| `scipy`      | t-tests, Wilcoxon tests, bootstrap CIs        | ≥ 1.9          |
| `yfinance`   | Fallback price source for PSE/BSE tickers     | ≥ 0.2          |
| `openpyxl`   | Reading and writing Excel files               | ≥ 3.0          |
| `requests`   | HTTP requests to Stooq API                    | ≥ 2.28         |

### 2. Create your Stooq API key file

`collect_returns.py` downloads historical price data from
[stooq.com](https://stooq.com). Stooq requires an API key for bulk CSV
downloads.

**To get a key:**
1. Go to `https://stooq.pl/q/d/?s=kgh&get_apikey` in your browser
2. Stooq will display your personal API key
3. Create a plain-text file named `stooq_apikey.txt` in the project root
4. Paste the key as a single line with no extra spaces or newlines

```
# stooq_apikey.txt (example — replace with your actual key)
aBcDeFgH1234567890xxxxxxxxxxxx
```

> **Do not share or commit `stooq_apikey.txt` to version control.**
> Add it to your `.gitignore`. The file is read automatically by
> `collect_returns.py`; no other configuration is needed.

**Troubleshooting:** If downloads return "No data", run `stooq_debug.py`
first. It tests six ticker formats and prints the raw Stooq response so you
can diagnose whether the key is valid or the ticker format has changed.

```bash
python stooq_debug.py
```

---

## Input Files

The three Excel files in `data/` contain the manually collected and
cleaned financial statement panel for each exchange. They were constructed
by sourcing annual report data directly from company websites and exchange
archives (see `*_URL_DATA_SOURCING.docx` files for the full source list).

Each file has a single sheet named **Panel Data** with the following key
columns:

| Column | Description |
|--------|-------------|
| `ticker` | Exchange ticker (e.g. `KGH.WA`, `CEZ`, `RICHTER`) |
| `company_name` | Full company name |
| `status` | `active` or `delisted` |
| `delist_year` | Year of delisting (if applicable) |
| `year` | Fiscal year of the financial statement |
| `ni` | Net income (local currency, thousands) |
| `at` | Total assets |
| `cfo` | Cash flow from operations |
| `lt` | Long-term debt |
| `act` | Current assets |
| `lct` | Current liabilities |
| `sale` | Revenue |
| `cogs` | Cost of goods sold |
| `csho` | Shares outstanding |
| `F_ROA … F_TURN` | Nine individual Piotroski binary signals |
| `FSCORE` | Composite F-Score (0–9) |
| `ret` | Forward 12-month return (filled by `collect_returns.py`) |
| `dlret` | Terminal return for delisted firms |

---

## Order of Execution

Run the scripts in this order. Each step writes output that the next
step reads.

### Step 1 — Compute F-Scores

```bash
python piotroski_pipeline.py
```

Reads the three Excel files from `data/`, recomputes all nine Piotroski
signals from scratch (overwriting any pre-existing computed columns), and
writes the cleaned, F-Score-enriched files back to `data/`. This step is
idempotent — safe to re-run.

**Output:** Updated `data/*.xlsx` files with `FSCORE` and signal columns
populated.

---

### Step 2 — Collect Forward Returns

```bash
python collect_returns.py
```

Requires an internet connection and a valid `stooq_apikey.txt` (see Setup).

For each firm, downloads daily price data from Stooq (primary) or Yahoo
Finance (fallback) and computes the 12-month equal-weighted holding-period
return from 1 July of year t+1 to 30 June of year t+2, where t is the
F-Score formation year. This convention embeds a **six-month reporting lag**
to avoid look-ahead bias.

Price data is cached in `data/prices/` as CSV files, keyed by ticker and
date range. Re-running the script uses the cache and does not make new HTTP
requests for tickers already downloaded.

**Ticker format notes:**
- WSE (Poland): bare Stooq ticker, no suffix — e.g. `kgh`, `cdr`
- PSE (Czech): bare Stooq ticker — e.g. `cez`; Yahoo fallback uses `.PR`
  suffix — e.g. `KOFOL.PR`, `TABAK.PR`
- BSE (Hungary): Stooq `.hu` suffix — e.g. `richter.hu`, `pannergy.hu`;
  Yahoo fallback uses `.BD` suffix

**Output:** Updated `data/*.xlsx` files with `ret` and `dlret` columns
populated.

---

### Step 3 — Main Portfolio Analysis

```bash
python portfolio_analysis.py
```

Constructs survivorship-biased and bias-corrected long-short portfolios,
runs all hypothesis tests, generates figures, and writes results to
`output/tables/` and `output/figures/`.

**Output tables (`output/tables/`):**

| File | Description |
|------|-------------|
| `tab0_coverage_summary.csv` | Sample coverage by exchange |
| `tab0_coverage_detail.csv` | Investable observations by exchange × year |
| `tab1_descriptive.csv` | Descriptive stats by exchange and score group (with bootstrap CIs) |
| `tab2_bias_comparison.csv` | Annual biased vs corrected L-S returns and bias |
| `tab3_portfolio_tests.csv` | t-test + Wilcoxon + bootstrap CI on L-S returns (H2) |
| `tab3b_h1_test.csv` | Paired tests: biased vs corrected (H1) |
| `tab3c_longonly.csv` | Annual long-only portfolio returns (FSCORE ≥ 7) |
| `tab4_country_subanalysis.csv` | WSE / PSE / BSE individually |
| `tab5_crisis_subanalysis.csv` | COVID-19 (2018–2021) vs Pre-COVID (2011–2017) |
| `tab6_fmb_betas.csv` | Fama-MacBeth annual cross-section betas |
| `tab7_missing_data.csv` | Firms with F-Score but no recoverable return |

**Output figures (`output/figures/`):**

| File | Description |
|------|-------------|
| `fig1_fscore_dist.png` | F-Score distribution by exchange (corrected sample) |
| `fig2_returns_by_score.png` | Mean return by F-Score value 1–9 (pooled) |
| `fig3_cumulative_ls.png` | Cumulative L-S returns: biased vs corrected |
| `fig4_annual_ls.png` | Annual L-S returns: biased vs corrected (bar chart) |
| `fig5_bias_magnitude.png` | Year-by-year bias (biased − corrected) |
| `fig6_country_returns.png` | Mean high/low returns by exchange |
| `fig7_longonly.png` | Long-only annual returns with 95% bootstrap CIs |

---

### Step 4 — Data Audit (optional but recommended before submission)

```bash
python data_audit.py
```

Runs three verification checks that directly address peer-review requirements:

1. **F_EQ_OFFER signal coverage** — confirms that csho data is available
   and that the full nine-signal F-Score is computed wherever possible.
   Identifies any rows forced to use a modified eight-signal score.

2. **PSE csho constancy check** — flags firms where shares outstanding
   are identical across all years. Determines whether this reflects genuine
   no-dilution (tightly held firms) or a data collection issue.

3. **Delisting taxonomy** — produces a formal D/A/U classification
   (Distress / Acquisition / Unverified) for each of the 13 delisted firms,
   with sources, terminal return recommendations, and verification flags.

**Output tables (`output/tables/`):**

| File | Description |
|------|-------------|
| `audit1_feqoffer_coverage.csv` | Signal coverage by exchange and signal |
| `audit2_pse_csho_check.csv` | csho constancy per firm |
| `audit3_delisting_taxonomy.csv` | Full D/A/U taxonomy with sources |

---

## Methodology Notes

### Portfolio thresholds

The combined panel contains 31 firms (10–18 investable per year). Canonical
Piotroski thresholds (F ≥ 8, F ≤ 2) produce zero or one firm in the short
leg in most years, making formal long-short tests degenerate. We therefore use:

- **Primary:** Long FSCORE ≥ 7, Short FSCORE ≤ 3
- **Strict robustness:** Long FSCORE ≥ 8, Short FSCORE ≤ 2

A portfolio year is reported only when ≥ 2 firms are in the long leg and
≥ 1 firm is in the short leg.

### Return convention (look-ahead bias safe)

| F-Score year | Portfolio formed | Held until | Holding period |
|---|---|---|---|
| Fiscal year t (Dec 31) | 1 July, year t+1 | 30 June, year t+2 | 12 months |

The six-month lag between fiscal year-end and portfolio formation reflects
the time required for annual reports to be publicly available and audited
in CEE markets.

### Survivorship bias methodology

- **Biased sample:** Only firms with `status = "active"` in the panel are
  included. This replicates the methodology of most prior F-Score studies.
- **Corrected sample:** All firms with a computable F-Score and a recoverable
  forward return are included, regardless of survival status.

### Terminal returns for delisted firms

All three WSE delisted firms (Arteria SA, Bogdanka SA, Play Communications)
were acquired at a premium — shareholders received cash consideration. We
therefore use their actual market return in the delisting year rather than
a forced −100% terminal return. The −100% assumption applies only to outright
equity wipeouts (e.g. NWR's 2016 restructuring), which are recorded explicitly
in the delisting taxonomy (`audit3_delisting_taxonomy.csv`).

### Missing price data

84 of 301 firm-year observations (28%) have no recoverable forward return.
The majority are PSE and BSE delisted firms whose historical prices are no
longer available on Stooq or Yahoo Finance. This data inaccessibility is
documented in `tab7_missing_data.csv` and is itself evidence of survivorship
bias — the firms most likely to lower corrected-sample performance are
precisely those for which data has been lost.

---

## Key Results (Summary)

| Finding | Result |
|---------|--------|
| H1: Biased sample overstates L-S returns | **Not supported** (mean bias = +0.4%, p = 0.951) |
| H2: Significant corrected alpha (long-short) | **Not supported** (mean = +1.4%, p = 0.927) |
| H2: Significant corrected alpha (long-only) | **Supported** (mean = +22.1%, t = 2.57, p = 0.023) |
| F-Score cross-sectional predictability | **Not supported** (FMB β = −0.017, p = 0.484) |
| Delisting composition (key auxiliary finding) | 54% acquisitions, 38% distress — explains near-zero bias |

---

## Known Limitations

1. **Small sample (31 firms, 217 investable obs.):** Statistical tests are
   underpowered. All results should be interpreted cautiously, particularly
   country-level subanalyses (PSE and BSE have zero viable long-short years).

2. **No local Fama-French factors:** WSE/PSE/BSE-specific SMB and HML factors
   require full-market data (all listed stocks, not just the 31 in this panel).
   Alpha tests use raw long-short returns and t-tests as a proxy. This is
   flagged as a limitation and direction for future research.

3. **No value-weighted portfolios:** Market capitalisation data was not
   collected for all firms. Equal-weighting is used throughout.

4. **GFC 2008-09 not testable:** The panel begins in 2010; the 2008-09 global
   financial crisis falls outside the sample period entirely.

5. **BSE sample skew:** Only three BSE firms have recoverable price data
   (Richter Gedeon, PannErgy, 4iG). BSE results reflect large-cap Hungarian
   firms, not the exchange as a whole.

6. **ENG.WA (Energa) caveat:** Energa was acquired by PKN Orlen and delisted
   in 2020, but Biznesradar.pl continues to publish group financials
   post-acquisition under the same identifier. Energa is treated as active
   through 2024 in the source data. This is acknowledged as a caveat.

7. **Short-selling constraints:** The long-short strategy is difficult or
   impossible to implement in practice in PSE and BSE, where securities
   lending markets are thin or non-existent. Long-only results are therefore
   the more practically relevant finding.

---

## Authors and Acknowledgements

Stanisław Halkiewicz — AGH University of Science and Technology Kraków /
WorldQuant University

Orjika Blessing — WorldQuant University

Dmitrii Verdun — WorldQuant University

Advisor: Iván Blanco (WorldQuant University)

---

*Last updated: May 2026*
