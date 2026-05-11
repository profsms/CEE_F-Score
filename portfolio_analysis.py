"""
portfolio_analysis.py — Piotroski F-Score Portfolio Analysis
=============================================================
Constructs biased vs. survivorship-bias-corrected F-Score portfolios
for the Visegrad 3 panel (WSE/PSE/BSE) and tests:

  H1: Survivorship-biased samples overstate F-Score long-short returns
  H2: After correction, statistically significant F-Score alpha remains

Peer-review additions (v2):
  - Bootstrapped confidence intervals on all key return estimates
  - Wilcoxon signed-rank non-parametric tests (robust to small n)
  - Country-level subsamples (WSE / PSE / BSE separately)
  - COVID-2020 crisis subanalysis (formation years 2018-2021)
  - Long-only formal test (annual high portfolio vs zero)
  - Data coverage table (firms, obs, returns by exchange × year)

Return convention (look-ahead bias safe):
  F-Score from fiscal year t → portfolio formed 1 Jul (t+1) → held 12m

Portfolio thresholds:
  Primary:  High ≥ 7, Low ≤ 3  (relaxed for thin sample)
  Strict:   High ≥ 8, Low ≤ 2  (robustness check)

Outputs: ./output/tables/  and  ./output/figures/

Authors: Stanisław Halkiewicz, Orjika Blessing, Dmitrii Verdun
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from scipy import stats
from scipy.stats import bootstrap as sp_bootstrap
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# ── Setup ─────────────────────────────────────────────────────────────────────

DATA_DIR   = Path('./data')
OUTPUT_DIR = Path('./output')
(OUTPUT_DIR / 'tables').mkdir(parents=True, exist_ok=True)
(OUTPUT_DIR / 'figures').mkdir(parents=True, exist_ok=True)

plt.rcParams.update({
    'font.family': 'Arial', 'font.size': 10,
    'axes.titlesize': 11, 'axes.labelsize': 10,
    'axes.spines.top': False, 'axes.spines.right': False,
    'axes.grid': True, 'grid.alpha': 0.3, 'figure.dpi': 150,
})

COLORS = {
    'WSE': '#2E5FA3', 'PSE': '#C0392B', 'BSE': '#27AE60',
    'high': '#27AE60', 'low': '#C0392B', 'ls': '#2E5FA3',
    'biased': '#E67E22', 'corrected': '#2E5FA3',
}

HIGH_THRESH  = 7
LOW_THRESH   = 3
HIGH_STRICT  = 8
LOW_STRICT   = 2
N_BOOTSTRAP  = 5000
CI_LEVEL     = 0.95

# ── Column name mapping ───────────────────────────────────────────────────────

DISPLAY_TO_CODE = {
    'ticker': 'ticker', 'company name': 'company_name', 'sector': 'sector',
    'country': 'country', 'exchange': 'exchange', 'listed': 'listing_year',
    'listing_year': 'listing_year', 'status': 'status', 'delisted': 'delist_year',
    'delist_year': 'delist_year', 'year': 'year',
    'net income': 'ni', 'ni': 'ni', 'total assets': 'at', 'at': 'at',
    'cfo': 'cfo', 'lt debt': 'lt', 'lt': 'lt',
    'curr assets': 'act', 'act': 'act', 'curr liab': 'lct', 'lct': 'lct',
    'revenue': 'sale', 'sale': 'sale', 'cogs': 'cogs',
    'shares out.': 'csho', 'csho': 'csho',
    'roa': 'roa', 'cfo/ta': 'cfo_ta', 'cfo_ta': 'cfo_ta',
    'leverage': 'lev', 'lev': 'lev', 'curr ratio': 'curr_ratio',
    'curr_ratio': 'curr_ratio', 'gross margin': 'gross_margin',
    'gross_margin': 'gross_margin', 'asset turn.': 'asset_turn',
    'asset_turn': 'asset_turn',
    'f_roa': 'F_ROA', 'f_droa': 'F_DROA', 'f_cfo': 'F_CFO',
    'f_accrual': 'F_ACCRUAL', 'f_dlever': 'F_DLEVER', 'f_dliquid': 'F_DLIQUID',
    'f_eq_off': 'F_EQ_OFFER', 'f_eq_offer': 'F_EQ_OFFER',
    'f_margin': 'F_MARGIN', 'f_turn': 'F_TURN',
    'n_sig': 'N_SIGNALS', 'n_signals': 'N_SIGNALS',
    'fscore': 'FSCORE', 'fwd return': 'ret', 'ret': 'ret',
    'delist ret': 'dlret', 'dlret': 'dlret',
    'ret start': 'ret_start', 'ret_start': 'ret_start',
    'ret end': 'ret_end', 'ret_end': 'ret_end',
}

# ── Load ──────────────────────────────────────────────────────────────────────

def load_panel(path):
    raw = pd.read_excel(path, sheet_name='Panel Data', header=None, dtype=str)
    for i, row in raw.iterrows():
        if any(str(v).strip().lower() == 'ticker'
               for v in row if v not in (None, 'nan')):
            headers = [
                DISPLAY_TO_CODE.get(str(v).strip().lower(),
                                    str(v).strip().lower().replace(' ', '_'))
                if v not in (None, 'nan') else f'col_{j}'
                for j, v in enumerate(raw.iloc[i])
            ]
            df = raw.iloc[i + 1:].copy()
            df.columns = headers
            df = df.dropna(how='all').reset_index(drop=True)
            num_cols = ['year', 'listing_year', 'delist_year', 'ni', 'at',
                        'cfo', 'lt', 'act', 'lct', 'sale', 'cogs', 'csho',
                        'roa', 'cfo_ta', 'lev', 'curr_ratio', 'gross_margin',
                        'asset_turn', 'F_ROA', 'F_DROA', 'F_CFO', 'F_ACCRUAL',
                        'F_DLEVER', 'F_DLIQUID', 'F_EQ_OFFER', 'F_MARGIN',
                        'F_TURN', 'N_SIGNALS', 'FSCORE', 'ret', 'dlret']
            for c in num_cols:
                if c in df.columns:
                    df[c] = pd.to_numeric(df[c], errors='coerce')
            return df
    raise ValueError(f"No header in {path}")


def load_all():
    dfs = []
    for fname, exch in [
        ('wse_piotroski_production.xlsx', 'WSE'),
        ('pse_piotroski_production.xlsx', 'PSE'),
        ('bse_piotroski_production.xlsx', 'BSE'),
    ]:
        df = load_panel(DATA_DIR / fname)
        df['exchange'] = exch
        dfs.append(df)
    all_df = pd.concat(dfs, ignore_index=True)
    all_df['year'] = all_df['year'].astype(int)
    return all_df


# ── Sample construction ───────────────────────────────────────────────────────

def make_samples(df):
    has_fscore = df['FSCORE'].notna()
    has_ret    = df['ret'].notna()
    biased     = df[has_fscore & (df['status'] == 'active')].copy()
    corrected  = df[has_fscore].copy()
    investable = df[has_fscore & has_ret].copy()
    return biased, corrected, investable


# ── Bootstrap CI ──────────────────────────────────────────────────────────────

def bootstrap_ci(series, n_boot=N_BOOTSTRAP, ci=CI_LEVEL):
    """Return (low, high) bootstrap CI for the mean. Returns (nan, nan) if n < 3."""
    clean = np.array(series.dropna())
    if len(clean) < 3:
        return np.nan, np.nan
    result = sp_bootstrap(
        (clean,), np.mean,
        n_resamples=n_boot,
        confidence_level=ci,
        method='percentile',
        random_state=42
    )
    return result.confidence_interval.low, result.confidence_interval.high


# ── Portfolio construction ────────────────────────────────────────────────────

def annual_portfolio_returns(df, hi=HIGH_THRESH, lo=LOW_THRESH, label=''):
    """Equal-weighted L-S portfolio returns by year. Viable if ≥2 high, ≥1 low."""
    rows = []
    for yr in sorted(df['year'].dropna().unique().astype(int)):
        sub  = df[(df['year'] == yr) & df['ret'].notna()]
        high = sub[sub['FSCORE'] >= hi]
        low  = sub[sub['FSCORE'] <= lo]
        n_h, n_l = len(high), len(low)
        viable = (n_h >= 2) and (n_l >= 1)
        rows.append({
            'year':    yr,
            'n_total': len(sub),
            'n_high':  n_h,
            'n_low':   n_l,
            'ret_high': high['ret'].mean() if n_h >= 1 else np.nan,
            'ret_low':  low['ret'].mean()  if n_l >= 1 else np.nan,
            'ret_ls':   (high['ret'].mean() - low['ret'].mean())
                        if viable else np.nan,
            'viable':  viable,
            'label':   label,
        })
    return pd.DataFrame(rows)


# ── Statistical tests ─────────────────────────────────────────────────────────

def test_ls_returns(ls_series, label=''):
    """One-sample t-test + Wilcoxon + bootstrap CI on L-S returns."""
    clean = ls_series.dropna()
    n = len(clean)
    if n < 3:
        return {
            'label': label, 'n': n,
            'mean': np.nan, 'median': np.nan, 'std': np.nan,
            'tstat': np.nan, 'pval_t': np.nan,
            'wilcoxon_stat': np.nan, 'pval_w': np.nan,
            'ci_low': np.nan, 'ci_high': np.nan,
            'sharpe': np.nan
        }
    mean   = clean.mean()
    median = clean.median()
    std    = clean.std(ddof=1)
    tstat, pval_t = stats.ttest_1samp(clean, 0)
    sharpe = mean / std if std > 0 else np.nan
    ci_low, ci_high = bootstrap_ci(clean)
    # Wilcoxon requires at least 4 non-zero obs
    try:
        w_stat, pval_w = stats.wilcoxon(clean)
    except Exception:
        w_stat, pval_w = np.nan, np.nan
    return {
        'label': label, 'n': n,
        'mean': mean, 'median': median, 'std': std,
        'tstat': tstat, 'pval_t': pval_t,
        'wilcoxon_stat': w_stat, 'pval_w': pval_w,
        'ci_low': ci_low, 'ci_high': ci_high,
        'sharpe': sharpe,
    }


def test_bias_difference(biased_ls, corrected_ls):
    """Paired t-test + Wilcoxon: biased vs corrected L-S returns."""
    merged = pd.DataFrame({'b': biased_ls, 'c': corrected_ls}).dropna()
    n = len(merged)
    if n < 3:
        return {'n': n, 'mean_diff': np.nan, 'median_diff': np.nan,
                'tstat': np.nan, 'pval_t': np.nan,
                'wilcoxon_stat': np.nan, 'pval_w': np.nan,
                'ci_low': np.nan, 'ci_high': np.nan}
    diff = merged['b'] - merged['c']
    tstat, pval_t = stats.ttest_rel(merged['b'], merged['c'])
    try:
        w_stat, pval_w = stats.wilcoxon(diff)
    except Exception:
        w_stat, pval_w = np.nan, np.nan
    ci_low, ci_high = bootstrap_ci(diff)
    return {
        'n': n, 'mean_diff': diff.mean(), 'median_diff': diff.median(),
        'tstat': tstat, 'pval_t': pval_t,
        'wilcoxon_stat': w_stat, 'pval_w': pval_w,
        'ci_low': ci_low, 'ci_high': ci_high,
    }


def test_longonly(port_returns, label=''):
    """One-sample t-test + bootstrap CI: high portfolio mean != 0."""
    return test_ls_returns(port_returns, label)


# ── Descriptive statistics ────────────────────────────────────────────────────

def descriptive_stats(df):
    rows = []
    for exch in ['WSE', 'PSE', 'BSE', 'ALL']:
        sub = df if exch == 'ALL' else df[df['exchange'] == exch]
        sub_ret = sub[sub['ret'].notna()]
        for grp, label in [
            (sub_ret[sub_ret['FSCORE'] >= HIGH_THRESH], f'High (≥{HIGH_THRESH})'),
            (sub_ret[sub_ret['FSCORE'] <= LOW_THRESH],  f'Low  (≤{LOW_THRESH})'),
            (sub_ret, 'All'),
        ]:
            if len(grp) == 0:
                continue
            ci_l, ci_h = bootstrap_ci(grp['ret'])
            rows.append({
                'Exchange':    exch,
                'Group':       label,
                'N':           len(grp),
                'Mean Ret':    grp['ret'].mean(),
                'Median':      grp['ret'].median(),
                'Std Dev':     grp['ret'].std(),
                'Min':         grp['ret'].min(),
                'Max':         grp['ret'].max(),
                f'CI_Low ({int(CI_LEVEL*100)}%)': ci_l,
                f'CI_High ({int(CI_LEVEL*100)}%)': ci_h,
                'Mean FSCORE': grp['FSCORE'].mean(),
            })
    return pd.DataFrame(rows)


# ── Data coverage table ───────────────────────────────────────────────────────

def data_coverage_table(df):
    """
    Produces two tables:
    - By exchange × year: obs, investable obs, high/low counts
    - Summary by exchange: total obs, investable, active/delisted firms
    """
    rows = []
    for exch in ['WSE', 'PSE', 'BSE']:
        sub = df[df['exchange'] == exch]
        for yr in sorted(sub['year'].dropna().unique().astype(int)):
            s   = sub[sub['year'] == yr]
            inv = s[s['ret'].notna()]
            rows.append({
                'Exchange': exch,
                'Year':     yr,
                'N_total':  len(s),
                'N_active': (s['status'] == 'active').sum(),
                'N_delisted': (s['status'] == 'delisted').sum(),
                'N_investable': len(inv),
                'N_high':   (inv['FSCORE'] >= HIGH_THRESH).sum(),
                'N_low':    (inv['FSCORE'] <= LOW_THRESH).sum(),
                'N_missing_ret': s['ret'].isna().sum(),
            })

    detail = pd.DataFrame(rows)

    summary_rows = []
    for exch in ['WSE', 'PSE', 'BSE']:
        sub = df[df['exchange'] == exch]
        inv = sub[sub['ret'].notna()]
        firms   = sub.groupby('ticker')['status'].first()
        summary_rows.append({
            'Exchange':          exch,
            'Total_firms':       len(firms),
            'Active_firms':      (firms == 'active').sum(),
            'Delisted_firms':    (firms == 'delisted').sum(),
            'Total_obs':         len(sub),
            'Investable_obs':    len(inv),
            'Coverage_pct':      len(inv) / len(sub) * 100,
            'Missing_obs':       sub['ret'].isna().sum(),
            'Year_range':        f"{int(sub['year'].min())}–{int(sub['year'].max())}",
        })

    summary = pd.DataFrame(summary_rows)
    return detail, summary


# ── Country-level subanalysis ─────────────────────────────────────────────────

def country_subanalysis(df, label_prefix=''):
    """Run portfolio + statistical tests for each exchange separately."""
    results = {}
    for exch in ['WSE', 'PSE', 'BSE']:
        sub = df[df['exchange'] == exch]
        biased_sub, corrected_sub, investable_sub = make_samples(sub)
        biased_inv = biased_sub[biased_sub['ret'].notna()]

        port_b = annual_portfolio_returns(biased_inv, HIGH_THRESH, LOW_THRESH,
                                          f'{exch} Biased')
        port_c = annual_portfolio_returns(investable_sub, HIGH_THRESH, LOW_THRESH,
                                          f'{exch} Corrected')

        test_b = test_ls_returns(port_b['ret_ls'], f'{exch} Biased')
        test_c = test_ls_returns(port_c['ret_ls'], f'{exch} Corrected')
        bias   = test_bias_difference(
            port_b.set_index('year')['ret_ls'],
            port_c.set_index('year')['ret_ls']
        )

        # Long-only: annual high-portfolio return
        lo_rows = []
        for yr in sorted(investable_sub['year'].dropna().unique().astype(int)):
            s = investable_sub[(investable_sub['year'] == yr) & investable_sub['ret'].notna()]
            hi = s[s['FSCORE'] >= HIGH_THRESH]
            if len(hi) >= 1:
                lo_rows.append({'year': yr, 'ret_high': hi['ret'].mean()})
        lo_df  = pd.DataFrame(lo_rows)
        lo_test = test_longonly(lo_df['ret_high'] if len(lo_df) > 0
                                else pd.Series(dtype=float),
                                f'{exch} Long-only')

        results[exch] = {
            'port_b': port_b, 'port_c': port_c,
            'test_b': test_b, 'test_c': test_c,
            'bias': bias, 'lo_test': lo_test,
            'n_firms': df[df['exchange'] == exch]['ticker'].nunique(),
            'n_investable': len(investable_sub),
        }
    return results


# ── Crisis subanalysis ────────────────────────────────────────────────────────

def crisis_subanalysis(investable, biased_inv):
    """
    COVID-19 subanalysis: F-Score formation years 2018-2021.
    Portfolio holding period Jul 2019–Jun 2022 spans the COVID shock.
    Note: GFC 2008-09 is outside our sample (data starts 2010-11).
    """
    covid_yrs = [2018, 2019, 2020, 2021]
    inv_covid = investable[investable['year'].isin(covid_yrs)]
    bia_covid = biased_inv[biased_inv['year'].isin(covid_yrs)]

    port_b_cov = annual_portfolio_returns(bia_covid, HIGH_THRESH, LOW_THRESH, 'COVID Biased')
    port_c_cov = annual_portfolio_returns(inv_covid, HIGH_THRESH, LOW_THRESH, 'COVID Corrected')

    test_b = test_ls_returns(port_b_cov['ret_ls'], 'COVID Biased')
    test_c = test_ls_returns(port_c_cov['ret_ls'], 'COVID Corrected')

    # Pre-COVID baseline: 2011-2017
    pre_yrs = list(range(2011, 2018))
    inv_pre = investable[investable['year'].isin(pre_yrs)]
    bia_pre = biased_inv[biased_inv['year'].isin(pre_yrs)]
    port_b_pre = annual_portfolio_returns(bia_pre, HIGH_THRESH, LOW_THRESH, 'Pre-COVID Biased')
    port_c_pre = annual_portfolio_returns(inv_pre, HIGH_THRESH, LOW_THRESH, 'Pre-COVID Corrected')
    test_b_pre = test_ls_returns(port_b_pre['ret_ls'], 'Pre-COVID Biased')
    test_c_pre = test_ls_returns(port_c_pre['ret_ls'], 'Pre-COVID Corrected')

    return {
        'covid': {'port_b': port_b_cov, 'port_c': port_c_cov,
                  'test_b': test_b, 'test_c': test_c},
        'pre_covid': {'port_b': port_b_pre, 'port_c': port_c_pre,
                      'test_b': test_b_pre, 'test_c': test_c_pre},
    }


# ── Fama-MacBeth cross-section ────────────────────────────────────────────────

def fama_macbeth(investable):
    from scipy.stats import linregress
    # Pooled OLS
    x = investable['FSCORE'].values
    y = investable['ret'].values
    mask = ~(np.isnan(x) | np.isnan(y))
    slope, intercept, r, p, se = linregress(x[mask], y[mask])
    pooled = {'beta': slope, 'intercept': intercept, 'r2': r**2, 'pval': p, 'n': mask.sum()}

    # Annual cross-sections
    betas = []
    for yr in sorted(investable['year'].unique()):
        sub = investable[investable['year'] == yr].dropna(subset=['FSCORE', 'ret'])
        if len(sub) >= 4:
            sl, ic, _, pv, _ = linregress(sub['FSCORE'], sub['ret'])
            betas.append({'year': yr, 'beta': sl, 'intercept': ic, 'n': len(sub)})
    bdf = pd.DataFrame(betas) if betas else pd.DataFrame()

    fmb = {}
    if len(bdf) >= 3:
        beta_t, beta_p = stats.ttest_1samp(bdf['beta'], 0)
        ci_l, ci_h = bootstrap_ci(bdf['beta'])
        fmb = {
            'n_years': len(bdf),
            'mean_beta': bdf['beta'].mean(),
            'std_beta': bdf['beta'].std(),
            'tstat': beta_t, 'pval': beta_p,
            'ci_low': ci_l, 'ci_high': ci_h,
            'pct_positive': (bdf['beta'] > 0).mean(),
        }
    return pooled, fmb, bdf


# ── Figures ───────────────────────────────────────────────────────────────────

def fig1_fscore_distribution(corrected):
    fig, axes = plt.subplots(1, 3, figsize=(12, 4), sharey=False)
    for ax, exch in zip(axes, ['WSE', 'PSE', 'BSE']):
        sub    = corrected[corrected['exchange'] == exch]
        counts = sub['FSCORE'].value_counts().sort_index()
        scores = range(0, 10)
        vals   = [counts.get(s, 0) for s in scores]
        bars   = ax.bar(scores, vals, color=COLORS[exch], alpha=0.8, edgecolor='white')
        ax.set_title(f'{exch} — F-Score Distribution\n(n={len(sub)} firm-years)')
        ax.set_xlabel('F-Score')
        ax.set_ylabel('Firm-year observations')
        ax.set_xticks(list(scores))
        for s in range(0, LOW_THRESH + 1):
            if s < len(vals): bars[s].set_color('#C0392B'); bars[s].set_alpha(0.7)
        for s in range(HIGH_THRESH, 10):
            if s < len(vals): bars[s].set_color('#27AE60'); bars[s].set_alpha(0.7)
    fig.suptitle('Piotroski F-Score Distribution — Visegrad 3 (Corrected Sample)',
                 fontweight='bold', y=1.01)
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'figures' / 'fig1_fscore_dist.png', bbox_inches='tight')
    plt.close()
    print("  Saved fig1_fscore_dist.png")


def fig2_returns_by_score(investable):
    means  = investable.groupby('FSCORE')['ret'].mean()
    counts = investable.groupby('FSCORE')['ret'].count()
    fig, ax = plt.subplots(figsize=(9, 4))
    scores = sorted(means.index)
    colors = ['#C0392B' if s <= LOW_THRESH
              else '#27AE60' if s >= HIGH_THRESH
              else '#7F8C8D' for s in scores]
    bars = ax.bar(scores, [means[s] for s in scores],
                  color=colors, alpha=0.85, edgecolor='white')
    for bar, s in zip(bars, scores):
        ax.text(bar.get_x() + bar.get_width()/2,
                bar.get_height() + 0.005, f'n={counts.get(s,0)}',
                ha='center', va='bottom', fontsize=8)
    ax.axhline(0, color='black', linewidth=0.8)
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
    ax.set_xlabel('Piotroski F-Score')
    ax.set_ylabel('Mean 12-Month Return')
    ax.set_xticks(scores)
    ax.set_title('Mean Forward Return by F-Score — Visegrad 3 Pooled\n'
                 '(Red = Low portfolio  |  Green = High portfolio)', fontweight='bold')
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'figures' / 'fig2_returns_by_score.png', bbox_inches='tight')
    plt.close()
    print("  Saved fig2_returns_by_score.png")


def fig3_cumulative_ls(port_b, port_c):
    fig, ax = plt.subplots(figsize=(10, 4))
    for port, label, color, ls in [
        (port_b, 'Biased (active only)',  COLORS['biased'],    '--'),
        (port_c, 'Corrected (all firms)', COLORS['corrected'], '-'),
    ]:
        viable = port[port['ret_ls'].notna()].sort_values('year')
        if len(viable) == 0: continue
        cumret = (1 + viable['ret_ls']).cumprod() - 1
        ax.plot(viable['year'], cumret, color=color, linestyle=ls,
                linewidth=1.8, marker='o', markersize=4, label=label)
    ax.axhline(0, color='black', linewidth=0.8)
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
    ax.set_xlabel('Year (F-Score formation year)')
    ax.set_ylabel('Cumulative Long-Short Return')
    ax.set_title(f'Cumulative L-S Returns: Biased vs Corrected\n'
                 f'(Long≥{HIGH_THRESH} | Short≤{LOW_THRESH} | Pooled V3)',
                 fontweight='bold')
    ax.legend()
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'figures' / 'fig3_cumulative_ls.png', bbox_inches='tight')
    plt.close()
    print("  Saved fig3_cumulative_ls.png")


def fig4_annual_ls(port_b, port_c):
    b = port_b[port_b['ret_ls'].notna()].set_index('year')
    c = port_c[port_c['ret_ls'].notna()].set_index('year')
    years = sorted(set(b.index) | set(c.index))
    x = np.arange(len(years)); w = 0.38
    fig, ax = plt.subplots(figsize=(12, 4))
    ax.bar(x - w/2, [b['ret_ls'].get(yr, np.nan) for yr in years],
           w, color=COLORS['biased'],    alpha=0.8, label='Biased',    edgecolor='white')
    ax.bar(x + w/2, [c['ret_ls'].get(yr, np.nan) for yr in years],
           w, color=COLORS['corrected'], alpha=0.8, label='Corrected', edgecolor='white')
    ax.axhline(0, color='black', linewidth=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(years, rotation=45, ha='right')
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
    ax.set_xlabel('F-Score Formation Year')
    ax.set_ylabel('Annual Long-Short Return')
    ax.set_title(f'Annual L-S Returns: Biased vs Corrected  (Long≥{HIGH_THRESH}/Short≤{LOW_THRESH})',
                 fontweight='bold')
    ax.legend()
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'figures' / 'fig4_annual_ls.png', bbox_inches='tight')
    plt.close()
    print("  Saved fig4_annual_ls.png")


def fig5_bias_magnitude(port_b, port_c):
    b    = port_b.set_index('year')['ret_ls']
    c    = port_c.set_index('year')['ret_ls']
    diff = (b - c).dropna()
    if len(diff) == 0:
        print("  Skipped fig5 — no overlapping years")
        return
    fig, ax = plt.subplots(figsize=(9, 3.5))
    colors = ['#C0392B' if v < 0 else '#27AE60' for v in diff.values]
    ax.bar(diff.index, diff.values, color=colors, alpha=0.85, edgecolor='white')
    ax.axhline(0, color='black', linewidth=0.8)
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
    ax.set_xlabel('F-Score Formation Year')
    ax.set_ylabel('Bias (Biased − Corrected Return)')
    ax.set_title('Survivorship Bias Magnitude: Overstatement in Biased L-S Returns\n'
                 '(Positive = biased overstates; Negative = biased understates)',
                 fontweight='bold')
    avg = diff.mean()
    ax.axhline(avg, color='navy', linestyle='--', linewidth=1,
               label=f'Mean bias: {avg:.1%}')
    ax.legend()
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'figures' / 'fig5_bias_magnitude.png', bbox_inches='tight')
    plt.close()
    print("  Saved fig5_bias_magnitude.png")


def fig6_country_returns(country_results):
    """Bar chart: mean high and low portfolio return by exchange."""
    fig, axes = plt.subplots(1, 3, figsize=(13, 4), sharey=False)
    for ax, exch in zip(axes, ['WSE', 'PSE', 'BSE']):
        r = country_results[exch]
        pc = r['port_c']
        inv_exch = pc  # port_c already filtered to exchange
        hi_mean = pc[pc['ret_high'].notna()]['ret_high'].mean()
        lo_mean = pc[pc['ret_low'].notna()]['ret_low'].mean()
        ls_vals = pc['ret_ls'].dropna()
        ax.bar(['High\n(≥7)', 'Low\n(≤3)'],
               [hi_mean if pd.notna(hi_mean) else 0,
                lo_mean if pd.notna(lo_mean) else 0],
               color=[COLORS['high'], COLORS['low']], alpha=0.8, edgecolor='white')
        ax.axhline(0, color='black', linewidth=0.8)
        ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
        ax.set_title(f'{exch}\n'
                     f'n={r["n_investable"]} obs  '
                     f'L-S years={r["port_c"]["viable"].sum()}',
                     fontweight='bold')
        ax.set_ylabel('Mean Annual Return')
    fig.suptitle('Mean Portfolio Returns by Exchange and Score Group (Corrected Sample)',
                 fontweight='bold', y=1.01)
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'figures' / 'fig6_country_returns.png', bbox_inches='tight')
    plt.close()
    print("  Saved fig6_country_returns.png")


def fig7_longonly(investable):
    """Annual long-only (high portfolio) returns with 95% CI band."""
    rows = []
    for yr in sorted(investable['year'].dropna().unique().astype(int)):
        s  = investable[(investable['year'] == yr) & investable['ret'].notna()]
        hi = s[s['FSCORE'] >= HIGH_THRESH]
        if len(hi) >= 2:
            m      = hi['ret'].mean()
            ci_l, ci_h = bootstrap_ci(hi['ret'], n_boot=2000)
            rows.append({'year': yr, 'mean': m, 'ci_low': ci_l, 'ci_high': ci_h, 'n': len(hi)})
    if not rows:
        print("  Skipped fig7 — no long-only data")
        return
    df_lo = pd.DataFrame(rows)
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.bar(df_lo['year'], df_lo['mean'], color=COLORS['high'], alpha=0.75,
           edgecolor='white', label='Mean return (High FSCORE ≥7)')
    ax.errorbar(df_lo['year'], df_lo['mean'],
                yerr=[df_lo['mean'] - df_lo['ci_low'],
                      df_lo['ci_high'] - df_lo['mean']],
                fmt='none', color='black', capsize=4, linewidth=1.2,
                label=f'{int(CI_LEVEL*100)}% bootstrap CI')
    ax.axhline(0, color='black', linewidth=0.8)
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
    ax.set_xlabel('F-Score Formation Year')
    ax.set_ylabel('Equal-Weighted Return')
    ax.set_title(f'Long-Only Portfolio (FSCORE≥{HIGH_THRESH}) — Annual Returns with '
                 f'{int(CI_LEVEL*100)}% Bootstrap CIs', fontweight='bold')
    ax.legend()
    for _, row in df_lo.iterrows():
        ax.text(row['year'], row['ci_high'] + 0.01, f"n={int(row['n'])}",
                ha='center', fontsize=7, color='gray')
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'figures' / 'fig7_longonly.png', bbox_inches='tight')
    plt.close()
    print("  Saved fig7_longonly.png")


def _fmt(v, pct=False, dec=4):
    if pd.isna(v): return "—"
    if pct: return f"{v:.1%}"
    return f"{v:.{dec}f}"


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("Piotroski Portfolio Analysis v2 — Visegrad 3")
    print("=" * 60)

    # ── Load ──────────────────────────────────────────────────────────────────
    df = load_all()
    biased, corrected, investable = make_samples(df)
    biased_inv = biased[biased['ret'].notna()]

    print(f"\nPanel: {len(corrected)} firm-years (corrected) | "
          f"{len(biased)} (biased) | {len(investable)} investable")

    # ── Data coverage table ───────────────────────────────────────────────────
    print("\n" + "─" * 60)
    print("TABLE 0: Data Coverage")
    cov_detail, cov_summary = data_coverage_table(df)
    print(cov_summary.to_string(index=False))
    cov_detail.to_csv(OUTPUT_DIR / 'tables' / 'tab0_coverage_detail.csv', index=False)
    cov_summary.to_csv(OUTPUT_DIR / 'tables' / 'tab0_coverage_summary.csv', index=False)

    # ── Descriptive statistics (with bootstrap CIs) ───────────────────────────
    print("\n" + "─" * 60)
    print(f"TABLE 1: Descriptive Statistics (with {int(CI_LEVEL*100)}% bootstrap CIs)")
    desc = descriptive_stats(investable)
    ci_cols = [c for c in desc.columns if 'CI' in c]
    show_cols = ['Exchange','Group','N','Mean Ret','Median','Std Dev'] + ci_cols + ['Mean FSCORE']
    print(desc[show_cols].to_string(index=False))
    desc.to_csv(OUTPUT_DIR / 'tables' / 'tab1_descriptive.csv', index=False)

    # ── Portfolio construction (pooled V3) ────────────────────────────────────
    print("\n" + "─" * 60)
    print("POOLED PORTFOLIO CONSTRUCTION")
    port_b       = annual_portfolio_returns(biased_inv,  HIGH_THRESH, LOW_THRESH, 'Biased')
    port_c       = annual_portfolio_returns(investable,  HIGH_THRESH, LOW_THRESH, 'Corrected')
    port_b_strict = annual_portfolio_returns(biased_inv, HIGH_STRICT, LOW_STRICT, 'Biased strict')
    port_c_strict = annual_portfolio_returns(investable, HIGH_STRICT, LOW_STRICT, 'Corrected strict')
    print(f"  Viable years — Biased: {port_b['viable'].sum()} | "
          f"Corrected: {port_c['viable'].sum()} / {len(port_c)}")

    # ── Table 2: Bias comparison ──────────────────────────────────────────────
    print("\n" + "─" * 60)
    print("TABLE 2: Annual Biased vs Corrected Returns")
    t2b = port_b[['year','n_total','n_high','n_low','ret_high','ret_low','ret_ls']].copy()
    t2b.columns = ['Year','N_tot_B','N_hi_B','N_lo_B','Ret_hi_B','Ret_lo_B','LS_B']
    t2c = port_c[['year','n_total','n_high','n_low','ret_high','ret_low','ret_ls']].copy()
    t2c.columns = ['Year','N_tot_C','N_hi_C','N_lo_C','Ret_hi_C','Ret_lo_C','LS_C']
    tab2 = t2b.merge(t2c, on='Year', how='outer').sort_values('Year')
    tab2['Bias'] = tab2['LS_B'] - tab2['LS_C']
    for c in ['Ret_hi_B','Ret_lo_B','LS_B','Ret_hi_C','Ret_lo_C','LS_C','Bias']:
        tab2[c] = tab2[c].map(lambda x: _fmt(x, pct=True))
    print(tab2.to_string(index=False))
    tab2.to_csv(OUTPUT_DIR / 'tables' / 'tab2_bias_comparison.csv', index=False)

    # ── Table 3: Statistical tests (t + Wilcoxon + bootstrap CI) ─────────────
    print("\n" + "─" * 60)
    print("TABLE 3: Portfolio Return Tests (t-test + Wilcoxon + Bootstrap CI)")
    test_rows = []
    for port, label in [
        (port_b,       f'Biased (≥{HIGH_THRESH}/≤{LOW_THRESH})'),
        (port_c,       f'Corrected (≥{HIGH_THRESH}/≤{LOW_THRESH})'),
        (port_b_strict, f'Biased strict (≥{HIGH_STRICT}/≤{LOW_STRICT})'),
        (port_c_strict, f'Corrected strict (≥{HIGH_STRICT}/≤{LOW_STRICT})'),
    ]:
        test_rows.append(test_ls_returns(port['ret_ls'], label))
    tab3 = pd.DataFrame(test_rows)
    print("\n  H2 — One-sample tests (H0: mean L-S return = 0)")
    for _, row in tab3.iterrows():
        print(f"  {row['label']:40s}  n={int(row['n']) if pd.notna(row['n']) else 0:2d}"
              f"  mean={_fmt(row['mean'],pct=True):8s}"
              f"  t={_fmt(row['tstat'],dec=3):7s}  p_t={_fmt(row['pval_t'],dec=4):7s}"
              f"  W={_fmt(row['wilcoxon_stat'],dec=1):8s}  p_W={_fmt(row['pval_w'],dec=4):7s}"
              f"  CI=[{_fmt(row['ci_low'],pct=True)}, {_fmt(row['ci_high'],pct=True)}]"
              f"  Sharpe={_fmt(row['sharpe'],dec=3):7s}")
    tab3.to_csv(OUTPUT_DIR / 'tables' / 'tab3_portfolio_tests.csv', index=False)

    # ── H1: Paired test biased vs corrected ───────────────────────────────────
    print("\n  H1 — Paired tests: Biased vs Corrected L-S returns")
    h1 = test_bias_difference(port_b.set_index('year')['ret_ls'],
                              port_c.set_index('year')['ret_ls'])
    print(f"  n={h1['n']}  mean_diff={_fmt(h1['mean_diff'],pct=True)}"
          f"  t={_fmt(h1['tstat'],dec=3)}  p_t={_fmt(h1['pval_t'],dec=4)}"
          f"  W={_fmt(h1['wilcoxon_stat'],dec=1)}  p_W={_fmt(h1['pval_w'],dec=4)}"
          f"  CI=[{_fmt(h1['ci_low'],pct=True)}, {_fmt(h1['ci_high'],pct=True)}]")
    verdict = ("H1 SUPPORTED" if (pd.notna(h1['pval_t']) and h1['pval_t'] < 0.10)
               else "H1 NOT SUPPORTED")
    print(f"  → {verdict}")
    pd.DataFrame([h1]).to_csv(OUTPUT_DIR / 'tables' / 'tab3b_h1_test.csv', index=False)

    # ── Long-only test ────────────────────────────────────────────────────────
    print("\n  Long-only test: annual high-portfolio returns vs zero")
    lo_rows = []
    for yr in sorted(investable['year'].dropna().unique().astype(int)):
        s  = investable[(investable['year'] == yr) & investable['ret'].notna()]
        hi = s[s['FSCORE'] >= HIGH_THRESH]
        if len(hi) >= 1:
            lo_rows.append({'year': yr, 'ret_high': hi['ret'].mean(), 'n': len(hi)})
    lo_df = pd.DataFrame(lo_rows)
    if len(lo_df) >= 3:
        lo_test = test_longonly(lo_df['ret_high'], 'Long-only (≥7)')
        print(f"  n={lo_test['n']}  mean={_fmt(lo_test['mean'],pct=True)}"
              f"  t={_fmt(lo_test['tstat'],dec=3)}  p_t={_fmt(lo_test['pval_t'],dec=4)}"
              f"  W={_fmt(lo_test['wilcoxon_stat'],dec=1)}  p_W={_fmt(lo_test['pval_w'],dec=4)}"
              f"  CI=[{_fmt(lo_test['ci_low'],pct=True)}, {_fmt(lo_test['ci_high'],pct=True)}]")
        lo_df.to_csv(OUTPUT_DIR / 'tables' / 'tab3c_longonly.csv', index=False)

    # High vs Rest (pooled)
    high_r = investable[investable['FSCORE'] >= HIGH_THRESH]['ret'].dropna()
    rest_r = investable[investable['FSCORE'] <  HIGH_THRESH]['ret'].dropna()
    t_hr, p_hr = stats.ttest_ind(high_r, rest_r, equal_var=False)
    print(f"  High vs Rest: mean_high={high_r.mean():.1%}  mean_rest={rest_r.mean():.1%}"
          f"  diff={high_r.mean()-rest_r.mean():.1%}"
          f"  t={t_hr:.3f}  p={p_hr:.4f}")

    # ── Country subanalysis ───────────────────────────────────────────────────
    print("\n" + "─" * 60)
    print("COUNTRY-LEVEL SUBANALYSIS")
    country_results = country_subanalysis(df)
    country_rows = []
    for exch, r in country_results.items():
        tb, tc, bias = r['test_b'], r['test_c'], r['bias']
        country_rows.append({
            'Exchange':       exch,
            'N_investable':   r['n_investable'],
            'Viable_yrs_B':   r['port_b']['viable'].sum(),
            'Viable_yrs_C':   r['port_c']['viable'].sum(),
            'Mean_LS_B':      _fmt(tb['mean'], pct=True),
            'Mean_LS_C':      _fmt(tc['mean'], pct=True),
            'p_t_B':          _fmt(tb['pval_t'], dec=4),
            'p_t_C':          _fmt(tc['pval_t'], dec=4),
            'p_W_B':          _fmt(tb['pval_w'], dec=4),
            'p_W_C':          _fmt(tc['pval_w'], dec=4),
            'Mean_bias':      _fmt(bias['mean_diff'], pct=True),
            'N_bias_pairs':   bias['n'],
            'LO_mean':        _fmt(r['lo_test']['mean'], pct=True),
            'LO_p':           _fmt(r['lo_test']['pval_t'], dec=4),
        })
        print(f"\n  {exch} (n_investable={r['n_investable']}):")
        print(f"    Biased   — mean L-S={_fmt(tb['mean'],pct=True)}  "
              f"p_t={_fmt(tb['pval_t'],dec=4)}  p_W={_fmt(tb['pval_w'],dec=4)}")
        print(f"    Corrected— mean L-S={_fmt(tc['mean'],pct=True)}  "
              f"p_t={_fmt(tc['pval_t'],dec=4)}  p_W={_fmt(tc['pval_w'],dec=4)}")
        print(f"    Bias     — mean={_fmt(bias['mean_diff'],pct=True)}  n={bias['n']}")
        print(f"    Long-only— mean={_fmt(r['lo_test']['mean'],pct=True)}  "
              f"p_t={_fmt(r['lo_test']['pval_t'],dec=4)}")
    tab_country = pd.DataFrame(country_rows)
    tab_country.to_csv(OUTPUT_DIR / 'tables' / 'tab4_country_subanalysis.csv', index=False)

    # ── Crisis subanalysis ────────────────────────────────────────────────────
    print("\n" + "─" * 60)
    print("CRISIS SUBANALYSIS: COVID-19 (formation years 2018-2021) vs Pre-COVID (2011-2017)")
    print("Note: GFC 2008-09 is outside our sample period (data starts 2010-11)")
    crisis = crisis_subanalysis(investable, biased_inv)
    crisis_rows = []
    for period, data in [('COVID (2018-2021)', crisis['covid']),
                          ('Pre-COVID (2011-2017)', crisis['pre_covid'])]:
        tb, tc = data['test_b'], data['test_c']
        print(f"\n  {period}:")
        print(f"    Biased   — n={tb['n']}  mean={_fmt(tb['mean'],pct=True)}"
              f"  p_t={_fmt(tb['pval_t'],dec=4)}  p_W={_fmt(tb['pval_w'],dec=4)}")
        print(f"    Corrected— n={tc['n']}  mean={_fmt(tc['mean'],pct=True)}"
              f"  p_t={_fmt(tc['pval_t'],dec=4)}  p_W={_fmt(tc['pval_w'],dec=4)}")
        crisis_rows.append({'Period': period, 'Sample': 'Biased',    **tb})
        crisis_rows.append({'Period': period, 'Sample': 'Corrected', **tc})
    pd.DataFrame(crisis_rows).to_csv(
        OUTPUT_DIR / 'tables' / 'tab5_crisis_subanalysis.csv', index=False)

    # ── Fama-MacBeth ──────────────────────────────────────────────────────────
    print("\n" + "─" * 60)
    print("CROSS-SECTIONAL REGRESSION (pooled OLS + Fama-MacBeth)")
    pooled, fmb, bdf = fama_macbeth(investable)
    print(f"\n  Pooled OLS: β={pooled['beta']:.4f}  R²={pooled['r2']:.4f}"
          f"  p={pooled['pval']:.4f}  n={pooled['n']}")
    if fmb:
        print(f"  Fama-MacBeth: mean_β={fmb['mean_beta']:.4f}  "
              f"t={fmb['tstat']:.3f}  p={fmb['pval']:.4f}  "
              f"CI=[{fmb['ci_low']:.4f}, {fmb['ci_high']:.4f}]  "
              f"n_years={fmb['n_years']}  pct_positive={fmb['pct_positive']:.0%}")
    if not bdf.empty:
        bdf.to_csv(OUTPUT_DIR / 'tables' / 'tab6_fmb_betas.csv', index=False)

    # ── Missing data summary ──────────────────────────────────────────────────
    print("\n" + "─" * 60)
    missing = df[df['ret'].isna() & df['FSCORE'].notna()][
        ['exchange','ticker','company_name','status','year','FSCORE']].copy()
    missing_firms = missing.groupby(
        ['exchange','ticker','company_name','status']).agg(
        missing_years=('year','count'),
        fscore_range=('FSCORE', lambda x: f"{x.min():.0f}–{x.max():.0f}")
    ).reset_index()
    print(f"Missing data: {missing_firms['ticker'].nunique()} firms, {len(missing)} firm-years")
    print(missing_firms.to_string(index=False))
    missing_firms.to_csv(OUTPUT_DIR / 'tables' / 'tab7_missing_data.csv', index=False)

    # ── Figures ───────────────────────────────────────────────────────────────
    print("\n" + "─" * 60)
    print("GENERATING FIGURES")
    fig1_fscore_distribution(corrected)
    fig2_returns_by_score(investable)
    fig3_cumulative_ls(port_b, port_c)
    fig4_annual_ls(port_b, port_c)
    fig5_bias_magnitude(port_b, port_c)
    fig6_country_returns(country_results)
    fig7_longonly(investable)

    # ── Final summary ─────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("COMPLETE — outputs in ./output/")
    print(f"  Tables: tab0 coverage, tab1 descriptive, tab2 bias,")
    print(f"          tab3 tests + H1 + long-only, tab4 country,")
    print(f"          tab5 crisis, tab6 FMB, tab7 missing")
    print(f"  Figures: fig1-7")


if __name__ == '__main__':
    main()
