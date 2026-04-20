#!/usr/bin/env python3
"""
WSE Piotroski F-Score: Survivorship Bias Analysis Pipeline
============================================================
MScFE 690 Capstone — Stanisław Halkiewicz, Orjika Blessing, Dmitrii Verdun
April 2026

Five-stage pipeline that runs end-to-end on the pre-scraped production data:

  Stage 1  Load pre-computed Piotroski panel (wse_piotroski_production.csv)
  Stage 2  Fetch / attach annual stock returns via yfinance
  Stage 3  Build long-short portfolios (biased vs corrected)
  Stage 4  Statistical testing (H1 bias test, H2 alpha test)
  Stage 5  Generate figures and summary tables

Run
---
    pip install pandas numpy matplotlib scipy yfinance openpyxl
    python pipeline.py

All outputs are written to:
    output/figures/   PNG figures (300 DPI)
    output/results/   CSV tables
"""

import sys
import warnings
from pathlib import Path

# Force UTF-8 output on Windows consoles that default to cp1250/cp1252
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from scipy.stats import ttest_1samp, ttest_rel

try:
    import yfinance as yf
    _HAVE_YF = True
except ImportError:
    _HAVE_YF = False
    print("[warn] yfinance not installed — returns will not be fetched.\n"
          "       Run: pip install yfinance", file=sys.stderr)

warnings.filterwarnings("ignore")
np.random.seed(42)

# ── Paths ─────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
OUT_DIR  = BASE_DIR / "output"
FIG_DIR  = OUT_DIR / "figures"
RES_DIR  = OUT_DIR / "results"
for _d in (FIG_DIR, RES_DIR):
    _d.mkdir(parents=True, exist_ok=True)

PANEL_CSV = OUT_DIR / "wse_piotroski_production.csv"

# ── Analysis parameters ───────────────────────────────────────────────────
#
# The pilot dataset covers 11 WSE firms (8 active, 3 delisted).  With such
# a small universe the strict Piotroski thresholds (F>=8, F<=2) yield 0–1
# firms per portfolio leg in most years, making the regression meaningless.
# We therefore use relaxed thresholds as the PRIMARY analysis and report the
# strict thresholds as a ROBUSTNESS check.
#
# NOTE: csho (shares outstanding) is unavailable in the source data, so
# F_EQ_OFFER cannot be computed.  FSCORE ranges 0–8 (not 0–9).
# A firm scoring 7–8 is in the top tier; 0–3 in the bottom tier.

MIN_SIGNALS  = 7    # minimum non-null Piotroski signals; rows below this are dropped
HI           = 7    # primary long-leg threshold  (F >= HI)
LO           = 3    # primary short-leg threshold (F <= LO)
HI_STRICT    = 8    # robustness long-leg threshold
LO_STRICT    = 2    # robustness short-leg threshold
START_YR     = 2010
END_YR       = 2024

STYLE = {
    "biased":    "#2E75B6",
    "corrected": "#C0392B",
    "active":    "#2E75B6",
    "delisted":  "#C0392B",
}

try:
    plt.style.use("seaborn-v0_8")
except OSError:
    try:
        plt.style.use("seaborn")
    except OSError:
        pass  # fall back to matplotlib default


# ═══════════════════════════════════════════════════════════════════════════
# STAGE 1 — Load pre-computed panel
# ═══════════════════════════════════════════════════════════════════════════

def load_panel() -> pd.DataFrame:
    """
    Load wse_piotroski_production.csv and keep only rows with a valid
    F-Score.  Rows where N_SIGNALS == 0 are placeholder years (no
    financial data available for the firm that year); rows with
    N_SIGNALS in {1,2,...,MIN_SIGNALS-1} have too few signals for a
    reliable composite score and are excluded.

    The scraper already computed FSCORE (as sum of available signals);
    we reuse it rather than re-deriving it from raw columns.
    """
    if not PANEL_CSV.exists():
        sys.exit(f"ERROR: Panel CSV not found at {PANEL_CSV}\n"
                 "Run scrape_biznesradar_production.py first.")

    df = pd.read_csv(PANEL_CSV, low_memory=False)

    # Coerce numeric columns
    for col in ["N_SIGNALS", "FSCORE", "year", "delist_year", "ret", "dlret",
                "listing_year"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df["year"] = df["year"].astype("Int64").astype(int)

    # Remove placeholder years (no usable data)
    df = df[df["N_SIGNALS"] >= MIN_SIGNALS].copy()
    df = df.dropna(subset=["FSCORE"])

    # Normalise status
    df["status"] = df["status"].str.strip().str.lower()

    df = df.sort_values(["ticker", "year"]).reset_index(drop=True)
    return df


# ═══════════════════════════════════════════════════════════════════════════
# STAGE 2 — Fetch and attach annual returns
# ═══════════════════════════════════════════════════════════════════════════

def _download_ticker(ticker: str) -> pd.Series:
    """
    Download adjusted close prices for one ticker and return a Series of
    calendar-year returns indexed by integer year.  The year-end close is
    taken as the last available trading day in December.
    """
    try:
        raw = yf.download(
            ticker,
            start=f"{START_YR - 1}-12-15",
            end=f"{END_YR + 1}-01-15",
            auto_adjust=True,
            progress=False,
            repair=False,
        )
        if raw.empty:
            return pd.Series(dtype=float, name=ticker)

        close = raw["Close"]
        if isinstance(close, pd.DataFrame):
            close = close.squeeze()

        year_end = close.resample("YE").last().dropna()
        ret = year_end.pct_change().dropna()
        ret.index = ret.index.year
        ret.name = ticker
        return ret

    except Exception as exc:
        print(f"    [warn] {ticker}: {exc}")
        return pd.Series(dtype=float, name=ticker)


def fetch_returns(tickers: list) -> pd.DataFrame:
    """
    Download returns for all tickers from yfinance.
    Returns a (year × ticker) DataFrame; missing values are NaN.
    """
    if not _HAVE_YF:
        return pd.DataFrame()

    series = []
    for t in tickers:
        series.append(_download_ticker(t))

    df = pd.concat(series, axis=1)
    df.index.name = "year"
    return df


def attach_returns(panel: pd.DataFrame) -> pd.DataFrame:
    """
    Merge market returns onto the panel and create the `final_ret` column.

    Priority order for the return value assigned to each firm-year:
      1. Freshly fetched yfinance return (calendar year, adjusted close)
      2. Pre-existing `ret` column from the production CSV (near-empty)

    For delisted firms where `dlret` is non-null, compound the two:
        final_ret = (1 + annual_ret) × (1 + dlret) – 1

    Acquisition note: all three delisted firms in this pilot (Arteria SA,
    Bogdanka SA, Play Communications) were acquired at a premium and not
    bankrupt.  We therefore use the actual market return in the delisting
    year rather than forcing –100%.  A sensitivity run using a –100%
    terminal override is available via `terminal_override=True` in
    build_portfolios().
    """
    tickers = panel["ticker"].unique().tolist()
    print(f"  Fetching returns for {len(tickers)} tickers …")

    ret_df = fetch_returns(tickers)

    if not ret_df.empty:
        ret_long = (
            ret_df.reset_index()
            .melt(id_vars="year", var_name="ticker", value_name="mkt_ret")
        )
        panel = panel.merge(ret_long, on=["ticker", "year"], how="left")
        panel["final_ret"] = panel["mkt_ret"].combine_first(panel["ret"])
    else:
        panel = panel.copy()
        panel["final_ret"] = panel["ret"]

    # Compound with delisting-year return where available
    mask = panel["dlret"].notna() & panel["final_ret"].notna()
    panel.loc[mask, "final_ret"] = (
        (1 + panel.loc[mask, "final_ret"]) * (1 + panel.loc[mask, "dlret"]) - 1
    )

    return panel


# ═══════════════════════════════════════════════════════════════════════════
# STAGE 3 — Portfolio construction
# ═══════════════════════════════════════════════════════════════════════════

def build_portfolios(
    panel: pd.DataFrame,
    corrected: bool,
    hi: int = HI,
    lo: int = LO,
    terminal_override: bool = False,
) -> pd.DataFrame:
    """
    Build annual equal-weighted long-short portfolios.

    Formation convention (Piotroski 2000):
        F-Score from financial statements for year t–1
        → portfolio formed at start of year t
        → return measured over year t

    Parameters
    ----------
    corrected : bool
        False → survivorship-biased sample (active firms only).
        True  → full corrected sample (all firms, including delisted).
    terminal_override : bool
        If True, replace the delisting-year return with –100% for each
        delisted firm (sensitivity test; not used in the primary run).
    """
    p = panel.copy()
    if not corrected:
        p = p[p["status"] == "active"]

    if terminal_override:
        mask = (p["status"] == "delisted") & (p["year"] == p["delist_year"])
        p.loc[mask, "final_ret"] = -1.0

    p = p.dropna(subset=["final_ret", "FSCORE"])
    years = sorted(p["year"].unique())

    rows = []
    for yr in years:
        if (yr - 1) not in years:
            continue

        prior   = p[p["year"] == yr - 1][["ticker", "FSCORE"]].dropna()
        current = p[p["year"] == yr   ][["ticker", "final_ret"]].dropna()
        m = prior.merge(current, on="ticker")

        long_r  = m.loc[m["FSCORE"] >= hi, "final_ret"]
        short_r = m.loc[m["FSCORE"] <= lo, "final_ret"]

        if long_r.empty or short_r.empty:
            continue

        rows.append({
            "year"          : yr,
            "long"          : long_r.mean(),
            "short"         : short_r.mean(),
            "ls"            : long_r.mean() - short_r.mean(),
            "n_long"        : len(long_r),
            "n_short"       : len(short_r),
            "long_tickers"  : "|".join(m.loc[m["FSCORE"] >= hi, "ticker"].tolist()),
            "short_tickers" : "|".join(m.loc[m["FSCORE"] <= lo, "ticker"].tolist()),
        })

    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════════════════════
# STAGE 4 — Statistical testing
# ═══════════════════════════════════════════════════════════════════════════

def run_analysis(
    biased: pd.DataFrame,
    corrected: pd.DataFrame,
    label: str = "",
) -> dict:
    """
    Test H1 (survivorship bias inflates alpha) and H2 (corrected alpha ≠ 0).

    H1 — Paired t-test: mean(biased L/S) vs mean(corrected L/S).
    H2 — One-sample t-test: mean(corrected L/S) = 0.

    Sharpe ratio is annualised as (mean / std) × sqrt(T) where T is the
    number of annual observations (consistent with the pilot methodology).
    """
    sep = "=" * 60
    print(f"\n{sep}\n{label or 'RESULTS'}\n{sep}")

    if biased.empty or corrected.empty or "year" not in biased.columns or "year" not in corrected.columns:
        print("  No portfolio data available — skipping.")
        return {}

    merged = biased.merge(corrected, on="year", suffixes=("_b", "_c"))

    if len(merged) < 3:
        print("  Too few overlapping years for testing (need >= 3).")
        return {}

    T = len(merged)

    # H2: one-sample t-test
    t_c, p_c = ttest_1samp(merged["ls_c"], 0)
    print(f"\n[H2 Alpha test]  Corrected L/S mean = {merged['ls_c'].mean():.4f}")
    print(f"  t = {t_c:.3f}   p = {p_c:.4f}   (T = {T})")
    print("  => " + (
        "Preliminary support for H2 at 10% level"
        if p_c < 0.10 else "No significant alpha at 10% level"
    ))

    # H1: paired t-test, bias in basis points
    bias_bps = (merged["ls_b"].mean() - merged["ls_c"].mean()) * 10_000
    t_b, p_b = ttest_rel(merged["ls_b"], merged["ls_c"])
    print(f"\n[H1 Bias test]   Bias = {bias_bps:.1f} bps/year")
    print(f"  Biased mean    = {merged['ls_b'].mean():.4f}")
    print(f"  Corrected mean = {merged['ls_c'].mean():.4f}")
    print(f"  t = {t_b:.3f}   p = {p_b:.4f}")
    if p_b < 0.10:
        print("  => Significant bias detected at 10% level (supports H1)")

    sh_b = merged["ls_b"].mean() / merged["ls_b"].std() * np.sqrt(T)
    sh_c = merged["ls_c"].mean() / merged["ls_c"].std() * np.sqrt(T)
    print(f"\n  Sharpe (biased)    = {sh_b:.3f}")
    print(f"  Sharpe (corrected) = {sh_c:.3f}")

    return {
        "label"            : label,
        "T"                : T,
        "bias_bps"         : bias_bps,
        "t_bias"           : t_b,
        "p_bias"           : p_b,
        "t_alpha"          : t_c,
        "p_alpha"          : p_c,
        "sharpe_biased"    : sh_b,
        "sharpe_corrected" : sh_c,
        "mean_b"           : merged["ls_b"].mean(),
        "mean_c"           : merged["ls_c"].mean(),
        "std_b"            : merged["ls_b"].std(),
        "std_c"            : merged["ls_c"].std(),
        "min_b"            : merged["ls_b"].min(),
        "max_b"            : merged["ls_b"].max(),
        "min_c"            : merged["ls_c"].min(),
        "max_c"            : merged["ls_c"].max(),
    }


# ═══════════════════════════════════════════════════════════════════════════
# STAGE 5 — Figures
# ═══════════════════════════════════════════════════════════════════════════

def fig1_fscore_distribution(panel: pd.DataFrame) -> None:
    """
    Figure 1: Grouped bar chart comparing F-Score distribution between
    active and delisted firms.  Vertical dashed lines mark the portfolio
    formation cutoffs.  Expected pattern under H1: delisted firms are
    concentrated at low F-Scores, which inflates the apparent quality of
    the survivorship-biased short leg.
    """
    active   = (panel[panel["status"] == "active"]["FSCORE"]
                .value_counts(normalize=True).sort_index())
    delisted = (panel[panel["status"] != "active"]["FSCORE"]
                .value_counts(normalize=True).sort_index())

    if delisted.empty:
        print("  [skip] No delisted firm-years in panel — skipping Fig 1.")
        return

    max_score = int(panel["FSCORE"].max())
    scores = list(range(max_score + 1))
    x, w   = np.arange(len(scores)), 0.38

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(x - w / 2,
           [active.get(s, 0)   for s in scores], w,
           label="Active firms",   color=STYLE["active"],   alpha=0.85,
           edgecolor="white")
    ax.bar(x + w / 2,
           [delisted.get(s, 0) for s in scores], w,
           label="Delisted firms", color=STYLE["delisted"], alpha=0.85,
           edgecolor="white")

    ax.set_xlabel("Piotroski F-Score (max = 8; csho unavailable)",  fontsize=11)
    ax.set_ylabel("Proportion of Firm-Year Observations",            fontsize=11)
    ax.set_title(
        "Figure 1: F-Score Distribution by Survival Status\n"
        "WSE Pilot Sample (2010–2024)",
        fontsize=13, fontweight="bold",
    )
    ax.set_xticks(x)
    ax.set_xticklabels([str(s) for s in scores])
    ax.yaxis.set_major_formatter(mticker.PercentFormatter(xmax=1, decimals=0))
    ax.legend(fontsize=11, framealpha=0.9)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.spines[["top", "right"]].set_visible(False)

    # Mark portfolio cutoffs — place at the boundary between LO/LO+1 and HI-1/HI
    lo_x = LO + 0.5
    hi_x = HI - 0.5
    if lo_x < len(scores):
        ax.axvline(lo_x, color="grey", linestyle=":", lw=1.2, alpha=0.7)
        ax.text(LO / 2, ax.get_ylim()[1] * 0.92,
                f"Short\n(<={LO})", ha="center", fontsize=9, color="grey")
    if hi_x < len(scores):
        ax.axvline(hi_x, color="grey", linestyle=":", lw=1.2, alpha=0.7)
        ax.text((HI + max_score) / 2, ax.get_ylim()[1] * 0.92,
                f"Long\n(>={HI})", ha="center", fontsize=9, color="grey")

    plt.tight_layout()
    path = FIG_DIR / "fig1_fscore_distribution.png"
    plt.savefig(path, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {path.relative_to(BASE_DIR)}")


def fig2_annual_ls_returns(
    biased: pd.DataFrame,
    corrected: pd.DataFrame,
    suffix: str = "",
    title_suffix: str = "",
) -> None:
    """
    Figure 2: Annual long-short returns for biased vs corrected portfolios.
    The shaded area is the bias gap.  COVID-19 year (2020) is highlighted.
    """
    merged = biased.merge(corrected, on="year", suffixes=("_b", "_c"))
    if merged.empty:
        print(f"  [skip] No overlapping years for Fig 2{suffix}.")
        return

    yrs = merged["year"].values

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(yrs, merged["ls_b"] * 100, "o-",
            color=STYLE["biased"],    lw=2, ms=7,
            label="Biased (survivors only)")
    ax.plot(yrs, merged["ls_c"] * 100, "s--",
            color=STYLE["corrected"], lw=2, ms=7,
            label="Corrected (all firms incl. delisted)")
    ax.fill_between(yrs,
                    merged["ls_b"] * 100, merged["ls_c"] * 100,
                    alpha=0.12, color="grey", label="Bias gap")
    ax.axhline(0, color="black", lw=0.8)

    ax.set_xlabel("Year",                              fontsize=12)
    ax.set_ylabel("Annual Long-Short Return (%)",      fontsize=12)
    ax.set_title(
        f"Figure 2: Annual Long-Short Returns — Biased vs. Corrected\n"
        f"WSE Pilot{title_suffix}",
        fontsize=13, fontweight="bold",
    )
    ax.set_xticks(yrs)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter(decimals=0))
    ax.legend(fontsize=11, framealpha=0.9)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.spines[["top", "right"]].set_visible(False)

    if 2020 in yrs:
        ax.axvspan(2019.5, 2020.5, alpha=0.07, color="red")
        ylim = ax.get_ylim()
        ax.text(2020, ylim[0] + (ylim[1] - ylim[0]) * 0.05,
                "COVID-19", fontsize=8, color="darkred", ha="center")

    plt.tight_layout()
    path = FIG_DIR / f"fig2_annual_ls_returns{suffix}.png"
    plt.savefig(path, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {path.relative_to(BASE_DIR)}")


def fig3_cumulative_returns(
    biased: pd.DataFrame,
    corrected: pd.DataFrame,
    suffix: str = "",
    title_suffix: str = "",
) -> None:
    """
    Figure 3: Cumulative long-short returns.
    C_t = prod_{s=1}^{t}(1 + r_s) – 1.
    The growing gap between the two series illustrates the compounding
    effect of even a small annual bias.
    """
    merged = biased.merge(corrected, on="year", suffixes=("_b", "_c"))
    if merged.empty:
        print(f"  [skip] No overlapping years for Fig 3{suffix}.")
        return

    yrs   = merged["year"].values
    b_cum = (1 + merged["ls_b"]).cumprod() - 1
    c_cum = (1 + merged["ls_c"]).cumprod() - 1

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(yrs, b_cum * 100, "o-",
            color=STYLE["biased"],    lw=2, ms=6,
            label="Biased (survivors only)")
    ax.plot(yrs, c_cum * 100, "s--",
            color=STYLE["corrected"], lw=2, ms=6,
            label="Corrected (all firms incl. delisted)")
    ax.fill_between(yrs, b_cum * 100, c_cum * 100,
                    alpha=0.12, color="grey", label="Cumulative bias gap")
    ax.axhline(0, color="black", lw=0.8)

    ax.set_xlabel("Year",                                  fontsize=12)
    ax.set_ylabel("Cumulative Long-Short Return (%)",      fontsize=12)
    ax.set_title(
        f"Figure 3: Cumulative Long-Short Returns — Biased vs. Corrected\n"
        f"WSE Pilot{title_suffix}",
        fontsize=13, fontweight="bold",
    )
    ax.set_xticks(yrs)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter(decimals=0))
    ax.legend(fontsize=11, framealpha=0.9)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.spines[["top", "right"]].set_visible(False)

    plt.tight_layout()
    path = FIG_DIR / f"fig3_cumulative_returns{suffix}.png"
    plt.savefig(path, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {path.relative_to(BASE_DIR)}")


# ═══════════════════════════════════════════════════════════════════════════
# STAGE 5 — Tables
# ═══════════════════════════════════════════════════════════════════════════

def save_summary_table(
    biased: pd.DataFrame,
    corrected: pd.DataFrame,
    stats: dict,
    suffix: str = "",
) -> None:
    """Save performance summary and hypothesis test results as CSV."""
    if not stats:
        return

    summary = pd.DataFrame({
        "Metric": [
            "Mean L/S Return",
            "Std Dev",
            "Sharpe Ratio (annualised)",
            "Min Annual Return",
            "Max Annual Return",
            "Annual Observations (T)",
        ],
        "Biased": [
            f"{stats['mean_b']:.4f}",
            f"{stats['std_b']:.4f}",
            f"{stats['sharpe_biased']:.3f}",
            f"{stats['min_b']:.4f}",
            f"{stats['max_b']:.4f}",
            f"{stats['T']}",
        ],
        "Corrected": [
            f"{stats['mean_c']:.4f}",
            f"{stats['std_c']:.4f}",
            f"{stats['sharpe_corrected']:.3f}",
            f"{stats['min_c']:.4f}",
            f"{stats['max_c']:.4f}",
            f"{stats['T']}",
        ],
    })

    tests = pd.DataFrame({
        "Hypothesis": [
            "H1 — Bias inflates alpha (paired t)",
            "H2 — Corrected alpha ≠ 0 (one-sample t)",
        ],
        "Bias (bps/year)": [f"{stats['bias_bps']:.1f}", "—"],
        "t-statistic":     [f"{stats['t_bias']:.3f}",  f"{stats['t_alpha']:.3f}"],
        "p-value":         [f"{stats['p_bias']:.4f}",  f"{stats['p_alpha']:.4f}"],
        "Significant (10%)": [
            "Yes" if stats["p_bias"]  < 0.10 else "No",
            "Yes" if stats["p_alpha"] < 0.10 else "No",
        ],
    })

    p_sum = RES_DIR / f"summary_stats{suffix}.csv"
    p_tst = RES_DIR / f"hypothesis_tests{suffix}.csv"
    summary.to_csv(p_sum, index=False)
    tests.to_csv(p_tst,   index=False)
    print(f"  Saved: {p_sum.name}  |  {p_tst.name}")


def save_portfolio_detail(
    biased: pd.DataFrame,
    corrected: pd.DataFrame,
    suffix: str = "",
) -> None:
    """
    Save a year-by-year record of which firms entered each portfolio leg.
    Useful for auditing and for the proposal's methodology description.
    """
    cols = ["year", "n_long", "n_short", "long", "short", "ls",
            "long_tickers", "short_tickers"]

    def prep(df, label):
        out = df[[c for c in cols if c in df.columns]].copy()
        out.insert(0, "sample", label)
        return out

    detail = pd.concat(
        [prep(biased, "biased"), prep(corrected, "corrected")],
        ignore_index=True,
    ).sort_values(["year", "sample"])

    path = RES_DIR / f"portfolio_detail{suffix}.csv"
    detail.to_csv(path, index=False)
    print(f"  Saved: {path.name}")


def save_fscore_crosstab(panel: pd.DataFrame) -> None:
    """
    Table: last observed F-Score by firm vs survival status.
    Mirrors the crosstab produced in the original notebook.
    """
    last_obs = panel.sort_values("year").groupby("ticker").tail(1)
    if last_obs["status"].nunique() < 2:
        return

    table = pd.crosstab(
        last_obs["FSCORE"].astype(int),
        last_obs["status"],
        normalize="columns",
    ).round(3)

    path = RES_DIR / "fscore_by_survival_status.csv"
    table.to_csv(path)
    print(f"  Saved: {path.name}")


# ═══════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════

def main() -> None:
    # ── Stage 1 ──────────────────────────────────────────────────────────
    print("=" * 60)
    print("Stage 1: Loading Piotroski panel …")
    print("=" * 60)
    panel = load_panel()
    n_del = panel[panel["status"] != "active"]["ticker"].nunique()
    print(f"  {len(panel)} firm-year observations")
    print(f"  {panel['ticker'].nunique()} firms  "
          f"({panel['ticker'].nunique() - n_del} active, {n_del} delisted)")
    print(f"  Years: {panel['year'].min()}–{panel['year'].max()}")

    # ── Stage 2 ──────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("Stage 2: Attaching annual returns …")
    print("=" * 60)
    panel = attach_returns(panel)
    n_ret = panel["final_ret"].notna().sum()
    print(f"  {n_ret} / {len(panel)} firm-years have return data")

    # ── Stage 3 — Primary thresholds ─────────────────────────────────────
    print("\n" + "=" * 60)
    print(f"Stage 3: Portfolio construction  (primary: F>={HI} / F<={LO}) …")
    print("=" * 60)

    biased_p    = build_portfolios(panel, corrected=False, hi=HI, lo=LO)
    corrected_p = build_portfolios(panel, corrected=True,  hi=HI, lo=LO)

    print(f"  Biased:    {len(biased_p)} annual observations")
    print(f"  Corrected: {len(corrected_p)} annual observations")

    if not corrected_p.empty:
        print("\n  Corrected portfolio detail (primary thresholds):")
        print(corrected_p[["year", "n_long", "n_short", "ls",
                            "long_tickers", "short_tickers"]].to_string(index=False))

    # ── Stage 3 — Strict thresholds (robustness) ─────────────────────────
    print(f"\n  Robustness: strict thresholds (F>={HI_STRICT} / F<={LO_STRICT}) …")
    biased_s    = build_portfolios(panel, corrected=False,
                                   hi=HI_STRICT, lo=LO_STRICT)
    corrected_s = build_portfolios(panel, corrected=True,
                                   hi=HI_STRICT, lo=LO_STRICT)
    print(f"  Biased (strict):    {len(biased_s)} annual obs")
    print(f"  Corrected (strict): {len(corrected_s)} annual obs")

    # ── Stage 4 ──────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("Stage 4: Statistical testing …")
    print("=" * 60)

    stats_p = run_analysis(
        biased_p, corrected_p,
        label=f"PRIMARY — Relaxed Thresholds (F>={HI} / F<={LO})",
    )
    stats_s = run_analysis(
        biased_s, corrected_s,
        label=f"ROBUSTNESS — Strict Thresholds (F>={HI_STRICT} / F<={LO_STRICT})",
    )

    # ── Stage 5 ──────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("Stage 5: Generating figures and tables …")
    print("=" * 60)

    # Figure 1 (all firm-years with valid F-Score)
    fig1_fscore_distribution(panel)

    # F-Score crosstab
    save_fscore_crosstab(panel)

    # Primary thresholds
    ts_p = f", Long: F>={HI} | Short: F<={LO}"
    fig2_annual_ls_returns(biased_p,  corrected_p,  title_suffix=ts_p)
    fig3_cumulative_returns(biased_p, corrected_p,  title_suffix=ts_p)
    save_summary_table(biased_p, corrected_p, stats_p)
    save_portfolio_detail(biased_p, corrected_p)

    # Strict (robustness) thresholds — only if enough data
    ts_s = f", Long: F>={HI_STRICT} | Short: F<={LO_STRICT}"
    if (not biased_s.empty and not corrected_s.empty
            and "year" in biased_s.columns and "year" in corrected_s.columns):
        fig2_annual_ls_returns(biased_s,  corrected_s,
                               suffix="_strict", title_suffix=ts_s)
        fig3_cumulative_returns(biased_s, corrected_s,
                                suffix="_strict", title_suffix=ts_s)
        save_summary_table(biased_s, corrected_s, stats_s, suffix="_strict")
        save_portfolio_detail(biased_s, corrected_s, suffix="_strict")

    # Save enriched panel (for audit / proposal appendix)
    audit_path = RES_DIR / "panel_with_returns.csv"
    keep_cols = [c for c in [
        "ticker", "company_name", "sector", "status", "delist_year",
        "year", "FSCORE", "N_SIGNALS", "final_ret",
        "F_ROA", "F_DROA", "F_CFO", "F_ACCRUAL",
        "F_DLEVER", "F_DLIQUID", "F_MARGIN", "F_TURN",
    ] if c in panel.columns]
    panel[keep_cols].to_csv(audit_path, index=False)
    print(f"  Saved: panel_with_returns.csv")

    print("\n" + "=" * 60)
    print("Done.  All outputs in ./output/figures/ and ./output/results/")
    print("=" * 60)


if __name__ == "__main__":
    main()
