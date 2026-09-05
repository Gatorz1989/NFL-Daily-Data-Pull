import re
"""
NFL Edge Model — Hybrid Data Pull Script
=========================================
Run weekly from Anaconda Prompt:
    cd "C:\\Users\\gator\\OneDrive\\Desktop\\NFL Edge Model"
    python nfl_edge_data_pull.py

Outputs: nfl_model_data.json  (load into NFL_Edge_Model.html)

Sources:
  - nflverse (player_stats, pfr_advstats, stats_team, rosters, injuries, depth_charts)
  - NFL Next Gen Stats (via nflverse parquet — requires pyarrow)
  - NFL Prospect Analyzer Excel (your file — rookie college fallback)
  - CFB Reference (auto-scraped — secondary college fallback)

Requirements:
    pip install pandas requests openpyxl pyarrow beautifulsoup4
"""

import os, sys, json, re, time, warnings
from datetime import datetime
from pathlib import Path

warnings.filterwarnings('ignore')

try:
    import pandas as pd
    import requests
except ImportError:
    print("Missing dependencies. Run: pip install pandas requests openpyxl")
    sys.exit(1)

# ─────────────────────────────────────────────
# CONFIG — edit these paths if needed
# ─────────────────────────────────────────────
# ── Season configuration ──────────────────────────────────────────────────
# BASELINE_SEASON: the completed NFL season to use as your prior-year data.
# For Week 1 of a new season, use the PREVIOUS year (e.g., 2025 for Week 1 2026).
# Once games are played in the new season, set CURRENT_SEASON = BASELINE_SEASON + 1
# and the script will auto-blend prior-year baseline with live current-year data.
BASELINE_SEASON = 2025      # ← Prior completed season  (change to 2025 for 2026 Week 1)
CURRENT_SEASON  = 2026      # ← Season being projected  (change to 2026 for 2026 Week 1)
SEASON          = BASELINE_SEASON  # used throughout script as the data pull year

# v2.36: was a hardcoded Windows path (C:\Users\gator\...) — worked fine on
# your laptop but crashed immediately (exit code 1) when this script ran
# on GitHub Actions, since that folder only ever existed on your machine.
# Path(__file__).parent resolves to wherever the script itself actually
# is, on any computer — same folder as today when run locally, and the
# repo root automatically when run by the GitHub Actions workflow (which
# is exactly where its commit step expects to find the output).
OUTPUT_DIR   = Path(__file__).parent
OUTPUT_FILE  = OUTPUT_DIR / "nfl_model_data.json"
# ── Prospect Analyzer — search multiple common locations ──────────────
_PA_SEARCH_DIRS = [
    Path(r'C:\Users\gator\OneDrive\Desktop\NFL Models\NFL Prospect Analyzer'),  # ← correct
    Path(r'C:\Users\gator\OneDrive\Desktop\NFL Models'),
    Path(r'C:\Users\gator\OneDrive\Desktop\NFL Prospect Analyzer'),
    Path(r'C:\Users\gator\OneDrive\Desktop'),
    Path(r'C:\Users\gator\Documents'),
]
PROSPECT_DIR = Path(r'C:\Users\gator\OneDrive\Desktop\NFL Models\NFL Prospect Analyzer')  # correct path

PROSPECT_FILENAMES = [
    # JSON exports (preferred — no engine issues)
    "NFL_Prospect_Analyzer_Model_v1.json",
    "NFL_Prospect_Analyzer_Model.json",
    "NFL Prospect Analyzer Model.json",
    "NFL_Prospect_Analyzer.json",
    "NFL Prospect Analyzer.json",
    "Prospect_Analyzer.json",
    # Excel files (fallback)
    "NFL_Prospect_Analyzer_Model.xlsx",
    "NFL Prospect Analyzer Model.xlsx",
    "NFL_Prospect_Analyzer.xlsx",
    "NFL Prospect Analyzer.xlsx",
    "Prospect_Analyzer.xlsx",
    "NFL Prospect Analyzer 2026.xlsx",
    "NFL_Prospect_Analyzer_2026.xlsx",
]
PROSPECT_FILE = None
for _search_dir in _PA_SEARCH_DIRS:
    if not _search_dir.exists():
        continue
    for _fn in PROSPECT_FILENAMES:
        _candidate = _search_dir / _fn
        if _candidate.exists():
            PROSPECT_FILE = _candidate
            PROSPECT_DIR  = _search_dir
            break
    if PROSPECT_FILE:
        break
    # Also do a glob scan for any .xlsx/.json with "prospect" or "analyzer" in the name
    for _f in list(_search_dir.glob("*.json")) + list(_search_dir.glob("*.xlsx")):
        if any(kw in _f.name.lower() for kw in ["prospect","analyzer","nfl_pa"]):
            PROSPECT_FILE = _f
            PROSPECT_DIR  = _search_dir
            break
    if PROSPECT_FILE:
        break
if PROSPECT_FILE is None:
    PROSPECT_FILE = _PA_SEARCH_DIRS[0] / "NFL_Prospect_Analyzer_Model.xlsx"  # default for error msg

NFLVERSE_BASE = "https://github.com/nflverse/nflverse-data/releases/download"

# ─────────────────────────────────────────────
# CONFERENCE TIERS (SEC-normalised)
# ─────────────────────────────────────────────
CONF_ADJ = {
    # Tier 1 — SEC baseline
    'sec': 1.00,
    # Tier 2 — near-SEC
    'big ten': 0.92, 'big 10': 0.92, 'b1g': 0.92,
    'big 12': 0.92, 'big xii': 0.92, 'big twelve': 0.92,
    # Tier 3 — mid-major power
    'acc': 0.82, 'pac-12': 0.82, 'pac 12': 0.82, 'pac-10': 0.82,
    # Tier 4 — group of 5
    'aac': 0.70, 'american athletic': 0.70,
    'mountain west': 0.70, 'mwc': 0.70,
    # Independent
    'fbs ind': 0.75, 'ind': 0.75,
    # Tier 5 — small conferences
    'sun belt': 0.58, 'sunbelt': 0.58,
    'mac': 0.58, 'mid-american': 0.58,
    'c-usa': 0.58, 'cusa': 0.58, 'conference usa': 0.58,
    # FCS
    'fcs': 0.48,
}
NFL_TRANS_FACTOR = 0.65   # college production → NFL baseline

# Rookie threshold: games before full NFL weighting
QB_THRESHOLD     = 5
SKILL_THRESHOLD  = 3

# Blend schedule: (college_weight, nfl_weight) by NFL games played
def get_blend(nfl_games, position):
    threshold = QB_THRESHOLD if position == 'QB' else SKILL_THRESHOLD
    if nfl_games == 0:           return (1.00, 0.00)
    elif nfl_games == 1:         return (0.70, 0.30)
    elif nfl_games == 2:         return (0.45, 0.55)
    elif nfl_games >= threshold: return (0.00, 1.00)
    else:   # QB games 3-4 (between skill threshold and QB threshold)
        pct = nfl_games / threshold
        return (round(1 - pct, 2), round(pct, 2))

# ─────────────────────────────────────────────
# UTILITY
# ─────────────────────────────────────────────
SESSION = requests.Session()
SESSION.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'text/csv,application/octet-stream,*/*',
})

def fetch_csv(url, label, silent_404=False):
    if not silent_404:
        print(f"  Fetching {label}...", end='', flush=True)
    try:
        r = SESSION.get(url, timeout=30, allow_redirects=True)
        r.raise_for_status()
        from io import StringIO
        df = pd.read_csv(StringIO(r.text), low_memory=False)
        if not silent_404:
            print(f" ✅ {len(df):,} rows")
        else:
            print(f"  Fetching {label}... ✅ {len(df):,} rows")
        return df
    except Exception as e:
        if not silent_404:
            print(f" ❌ {e}")
        return pd.DataFrame()

def fetch_parquet(url, label, silent=False):
    """Fetch parquet file — requires pyarrow"""
    if not silent:
        pass  # printing handled by caller for NGS
    try:
        import pyarrow.parquet as pq
        import io
        r = SESSION.get(url, timeout=45, allow_redirects=True)
        r.raise_for_status()
        buf = io.BytesIO(r.content)
        df = pq.read_table(buf).to_pandas()
        print(f" ✅ {len(df):,} rows")
        return df
    except ImportError:
        print(" ⚠ pyarrow not installed — skipping NGS parquet")
        return pd.DataFrame()
    except Exception as e:
        print(f" ❌ {e}")
        return pd.DataFrame()

def safe_float(val, default=0.0):
    try:    return float(val) if pd.notna(val) else default
    except: return default

def safe_int(val, default=0):
    try:    return int(val) if pd.notna(val) else default
    except: return default

def norm_name(s):
    """Normalise player names for matching.
    v2.29: also strips generational suffixes (Jr, Sr, II, III, IV, V) —
    different nflverse sources include/omit these inconsistently (e.g. a
    depth chart listing "James Cook III" against player stats listing plain
    "James Cook"), which was causing real players to be treated as two
    different people and silently falling back to a generic positional
    average instead of their own real stats."""
    s = re.sub(r"[^a-z ]", "", str(s).lower().strip())
    s = re.sub(r"\b(jr|sr|ii|iii|iv|v)\b", "", s)
    return re.sub(r"\s+", " ", s).strip()


# ─────────────────────────────────────────────
# PBP-BASED PLAYER STATS (fallback when season CSV not yet published)
# ─────────────────────────────────────────────
def build_player_stats_from_pbp(pbp, rosters):
    """
    Aggregate 2025 player season stats directly from play-by-play data.
    Called when nflverse hasn't yet published the season-level player_stats CSV.
    Returns a DataFrame matching the player_stats_season format expected by
    build_player_profiles().

    Stats computed:
      QB  — passing yards/TDs/INTs, EPA, CPOE (mean), rushing yards, games
      RB  — rushing yards/TDs/EPA, receiving yards/targets/receptions
      WR/TE — receiving yards/TDs/EPA, target share, air yards share, WOPR
    """
    print("  Building player stats from play-by-play data...")
    reg = pbp[pbp["season_type"] == "REG"].copy() if "season_type" in pbp.columns else pbp.copy()

    # Build gsis_id → full name + position lookup from rosters
    id_col   = next((c for c in rosters.columns if "gsis" in c.lower()), None)
    name_col = next((c for c in rosters.columns
                     if "full_name" in c.lower() or "display" in c.lower()), None)
    pos_col  = next((c for c in rosters.columns
                     if c.lower() in ["position", "pos"]), None)
    id_to_info = {}
    if id_col and name_col:
        for _, row in rosters.iterrows():
            gsis = str(row.get(id_col, "")).strip()
            name = str(row.get(name_col, "")).strip()
            pos  = str(row.get(pos_col, "")).upper().strip() if pos_col else ""
            if gsis and name and gsis != "nan":
                id_to_info[gsis] = {"display_name": name, "position": pos}

    def _resolve(player_id, abbr_name, fallback_pos=""):
        info = id_to_info.get(str(player_id), {})
        return (info.get("display_name", abbr_name),
                info.get("position", fallback_pos))

    pass_plays = reg[(reg.get("pass_attempt", pd.Series(dtype=float)) == 1) &
                     reg["passer_player_name"].notna()].copy()
    rush_plays = reg[(reg.get("rush_attempt", pd.Series(dtype=float)) == 1) &
                     reg["rusher_player_name"].notna()].copy()

    # ── QB stats ──────────────────────────────────────────────────────────
    if not pass_plays.empty:
        qb_ids = set(pass_plays["passer_player_id"].dropna())
        qb_agg = pass_plays.groupby(
            ["passer_player_id", "passer_player_name", "posteam"]
        ).agg(
            passing_yards    =("passing_yards",  "sum"),
            completions      =("complete_pass",  "sum"),
            attempts         =("pass_attempt",   "sum"),
            passing_tds      =("pass_touchdown", "sum"),
            interceptions    =("interception",   "sum"),
            passing_epa      =("qb_epa",         "sum"),
            cpoe             =("cpoe",            "mean"),
            passing_air_yards=("air_yards",      "sum"),
            games            =("week",           "nunique"),
        ).reset_index()

        # Add QB rushing
        qb_rush = rush_plays[
            rush_plays["rusher_player_id"].isin(qb_ids)].groupby(
            "rusher_player_id").agg(
            rush_yds_qb=("rushing_yards", "sum")).reset_index()
        qb_agg = qb_agg.merge(
            qb_rush.rename(columns={"rusher_player_id": "passer_player_id"}),
            on="passer_player_id", how="left")
        qb_agg["rushing_yards"] = qb_agg.get("rush_yds_qb", 0).fillna(0)
        qb_agg["dakota"] = qb_agg["passing_epa"] / qb_agg["attempts"].clip(lower=1)
    else:
        qb_agg = pd.DataFrame(); qb_ids = set()

    # ── RB stats ──────────────────────────────────────────────────────────
    rb_rush = rush_plays[~rush_plays["rusher_player_id"].isin(qb_ids)].copy()
    if not rb_rush.empty:
        rb_ids = set(rb_rush["rusher_player_id"].dropna())
        rb_agg = rb_rush.groupby(
            ["rusher_player_id", "rusher_player_name", "posteam"]
        ).agg(
            rushing_yards=("rushing_yards", "sum"),
            carries      =("rush_attempt",  "sum"),
            rushing_tds  =("rush_touchdown","sum"),
            rushing_epa  =("epa",           "sum"),
            games        =("week",          "nunique"),
        ).reset_index()

        rb_rec = pass_plays[
            pass_plays["receiver_player_id"].isin(rb_ids)].groupby(
            "receiver_player_id").agg(
            receiving_yards_rb=("receiving_yards","sum"),
            receptions_rb     =("complete_pass",  "sum"),
            targets_rb        =("pass_attempt",   "sum"),
            receiving_epa_rb  =("epa",            "sum"),
        ).reset_index()
        rb_agg = rb_agg.merge(
            rb_rec.rename(columns={"receiver_player_id": "rusher_player_id"}),
            on="rusher_player_id", how="left")
        for col in ["receiving_yards_rb","receptions_rb","targets_rb","receiving_epa_rb"]:
            rb_agg[col] = rb_agg.get(col, 0).fillna(0)
    else:
        rb_agg = pd.DataFrame(); rb_ids = set()

    # ── WR/TE stats ──────────────────────────────────────────────────────
    wr_plays = pass_plays[
        pass_plays["receiver_player_name"].notna() &
        ~pass_plays["receiver_player_id"].isin(rb_ids if not rb_rush.empty else set())
    ].copy()
    if not wr_plays.empty:
        # Fix: ALL team pass plays as denominator — not WR-filtered (avoids 2-3x inflation)
        team_tgt = pass_plays.groupby("posteam")["pass_attempt"].sum().rename("team_targets")
        team_ayd = pass_plays.groupby("posteam")["air_yards"].sum().rename("team_air_yards")
        wr_agg = wr_plays.groupby(
            ["receiver_player_id", "receiver_player_name", "posteam"]
        ).agg(
            targets         =("pass_attempt",   "sum"),
            receptions      =("complete_pass",  "sum"),
            receiving_yards =("receiving_yards","sum"),
            receiving_tds   =("pass_touchdown", "sum"),
            receiving_epa   =("epa",            "sum"),
            air_yards_sum   =("air_yards",      "sum"),
            games           =("week",           "nunique"),
        ).reset_index()
        wr_agg = wr_agg.merge(team_tgt, on="posteam", how="left")
        wr_agg = wr_agg.merge(team_ayd, on="posteam", how="left")
        wr_agg["target_share"]   = wr_agg["targets"] / wr_agg["team_targets"].clip(lower=1)
        wr_agg["air_yards_share"]= wr_agg["air_yards_sum"] / wr_agg["team_air_yards"].clip(lower=1)
        wr_agg["wopr"]           = (1.5 * wr_agg["target_share"] +
                                    0.7 * wr_agg["air_yards_share"])
    else:
        wr_agg = pd.DataFrame()

    # ── Assemble unified player_season DataFrame ──────────────────────────
    rows = []

    if not qb_agg.empty:
        for _, r in qb_agg[qb_agg["attempts"] >= 30].iterrows():
            full, pos = _resolve(r["passer_player_id"], r["passer_player_name"], "QB")
            rows.append({
                "player_display_name": full, "position": "QB",
                "recent_team": r["posteam"], "games": r["games"],
                "passing_yards": r["passing_yards"], "passing_tds": r["passing_tds"],
                "interceptions": r["interceptions"], "passing_epa": r["passing_epa"],
                "cpoe": r["cpoe"], "dakota": r["dakota"],
                "completions": r["completions"], "attempts": r["attempts"],
                "rushing_yards": r.get("rushing_yards", 0),
            })

    if not rb_agg.empty:
        for _, r in rb_agg[rb_agg["carries"] >= 10].iterrows():
            full, pos = _resolve(r["rusher_player_id"], r["rusher_player_name"], "RB")
            rows.append({
                "player_display_name": full, "position": pos or "RB",
                "recent_team": r["posteam"], "games": r["games"],
                "rushing_yards": r["rushing_yards"], "carries": r["carries"],
                "rushing_tds": r["rushing_tds"], "rushing_epa": r["rushing_epa"],
                "receiving_yards": r["receiving_yards_rb"],
                "receptions": r["receptions_rb"], "targets": r["targets_rb"],
                "receiving_epa": r["receiving_epa_rb"],
            })

    if not wr_agg.empty:
        for _, r in wr_agg[wr_agg["targets"] >= 5].iterrows():
            full, pos = _resolve(r["receiver_player_id"], r["receiver_player_name"], "WR")
            rows.append({
                "player_display_name": full, "position": pos or "WR",
                "recent_team": r["posteam"], "games": r["games"],
                "receiving_yards": r["receiving_yards"], "receptions": r["receptions"],
                "targets": r["targets"], "receiving_tds": r["receiving_tds"],
                "receiving_epa": r["receiving_epa"],
                "target_share": r["target_share"], "air_yards_share": r["air_yards_share"],
                "wopr": r["wopr"],
            })

    out = pd.DataFrame(rows)
    pos_ct = out["position"].value_counts().to_dict() if not out.empty else {}
    print(f"  ✅ PBP aggregation: {len(out)} players "
          f"({pos_ct.get('QB',0)} QB / {pos_ct.get('RB',0)} RB / "
          f"{pos_ct.get('WR',0)} WR / {pos_ct.get('TE',0)} TE)")
    return out


def fetch_pbp_player_stats(season, rosters):
    """
    Download play-by-play for `season` and compute player stats.
    Used when the season-level CSV hasn't been published yet by nflverse.
    """
    import gzip as _gz, io as _io
    url = f"{NFLVERSE_BASE}/pbp/play_by_play_{season}.csv.gz"
    print(f"  Downloading PBP {season} ({url.split('/')[-1]}) ...", end="", flush=True)
    try:
        r = SESSION.get(url, timeout=180, allow_redirects=True)
        r.raise_for_status()
        size_mb = len(r.content) / 1024 / 1024
        print(f" {size_mb:.1f}MB", end="", flush=True)
        with _gz.open(_io.BytesIO(r.content)) as gz:
            pbp = pd.read_csv(gz, low_memory=False)
        print(f" — {len(pbp):,} plays")
        # If rosters not yet loaded, fetch them now for name resolution
        _rosters = rosters if (rosters is not None and not rosters.empty) else pd.DataFrame()
        if _rosters.empty:
            _r2 = SESSION.get(f"{NFLVERSE_BASE}/rosters/roster_{season}.csv",
                              timeout=30, allow_redirects=True)
            if _r2.status_code == 200:
                from io import StringIO as _SI
                _rosters = pd.read_csv(_SI(_r2.text), low_memory=False)
        return build_player_stats_from_pbp(pbp, _rosters)
    except Exception as e:
        print(f" ❌ {e}")
        return pd.DataFrame()


# ─────────────────────────────────────────────
# 1. PULL NFLVERSE DATA
# ─────────────────────────────────────────────
def pull_nflverse(season):
    print(f"\n{'='*50}")
    print(f"PULLING NFLVERSE DATA — {season} SEASON")
    print('='*50)

    BASE = NFLVERSE_BASE
    dfs = {}

    # ── Helper: try current season first, fall back to prior year ──
    def fetch_with_fallback(url_template, label, fallback_season=None):
        """Try season URL, fall back to prior year if 404 (pre-season)."""
        df = fetch_csv(url_template.format(s=season), f"{label} {season}")
        if df.empty and fallback_season:
            print(f"    ↳ Falling back to {fallback_season} data for {label}")
            df = fetch_csv(url_template.format(s=fallback_season), f"{label} {fallback_season} (fallback)")
        return df

    prior = season - 1  # fallback year (e.g., 2024 when season=2025)

    # Player stats — direct nflverse CSV download (works on all Python versions)
    # nflverse stores player_stats by season. Priority order:
    #   1. Season CSV if published  2. PBP aggregation (same season, accurate)  3. Prior year CSV
    # ── Check for local combined CSV first (bypasses nflverse download) ──────
    import os as _os
    # Load weekly player stats if available (used for healthy roster shares)
    _WEEKLY_NAMES = [
        'nfl_player_stats_weekly.csv',
        f'stats_player_week_{season}.csv',
    ]
    for _wn in _WEEKLY_NAMES:
        _wp = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), _wn)
        if _os.path.exists(_wp):
            try:
                _wdf = pd.read_csv(_wp, low_memory=False)
                if 'week' in _wdf.columns:
                    dfs['player_stats_weekly'] = _wdf
                    print(f"  Weekly player stats: {_wn} ✅  {len(_wdf):,} rows")
            except Exception as _we:
                print(f"  Weekly CSV error: {_we}")
            break

    _LOCAL_NAMES = [
        'nfl_player_stats_combined.csv',
        f'nfl_player_stats_{season}.csv',
        f'stats_player_reg_{season}.csv',
    ]
    _local_found = False
    for _lcn in _LOCAL_NAMES:
        _lcp = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), _lcn)
        if _os.path.exists(_lcp):
            try:
                _lps = pd.read_csv(_lcp, low_memory=False)
                if 'team' in _lps.columns and 'recent_team' not in _lps.columns:
                    _lps = _lps.rename(columns={'team': 'recent_team'})
                if 'season' not in _lps.columns:
                    _lps['season'] = season
                dfs['player_stats']  = _lps
                dfs['player_season'] = _lps
                print(f"  Local player stats: {_lcn} ✅  {len(_lps):,} players loaded")
                _local_found = True
                break
            except Exception as _le:
                print(f"  Local CSV {_lcn} error: {_le}")

        _ps_raw = pd.DataFrame()
    # When local CSV found, assign it to _ps_raw so downstream
    # 'if not _ps_raw.empty:' uses the local data instead of triggering
    # the else: branch that wipes dfs['player_season']
    _ps_raw = dfs['player_stats'].copy() if _local_found else pd.DataFrame()
    if not _local_found:

        # Try current season CSV first (silently — 404 expected pre-season)
        _ps_try = fetch_csv(f"{BASE}/player_stats/player_stats_{season}.csv",
                            f"player_stats {season}", silent_404=True)
        if not _ps_try.empty:
            _ps_reg = _ps_try[_ps_try['season_type']=='REG'] if 'season_type' in _ps_try.columns else _ps_try
            if not _ps_reg.empty:
                _ps_raw = _ps_reg.copy()
                if 'dakota' in _ps_raw.columns and 'cpoe' not in _ps_raw.columns:
                    _ps_raw = _ps_raw.rename(columns={'dakota': 'cpoe'})
                elif 'passing_cpoe' in _ps_raw.columns:
                    _ps_raw = _ps_raw.rename(columns={'passing_cpoe': 'cpoe'})
                print(f"  player_stats {season}... ✅ {len(_ps_raw):,} rows")

        if _ps_raw.empty:
            # CSV not published yet — aggregate from play-by-play (accurate current season)
            print(f"  player_stats {season}... ❌ not published — aggregating from PBP")
            # Fetch rosters NOW so _resolve has gsis_id→full_name data during PBP aggregation
            # (previously rosters were fetched after PBP, leaving id_to_info empty)
            if dfs.get('rosters', pd.DataFrame()).empty:
                print(f"  Pre-fetching rosters {season} for PBP name resolution...", end='', flush=True)
                _early_roster = fetch_csv(
                    f"{BASE}/rosters/roster_{season}.csv",
                    f"rosters {season}")
                if _early_roster.empty:  # try prior year as fallback
                    _early_roster = fetch_csv(
                        f"{BASE}/rosters/roster_{prior}.csv",
                        f"rosters {prior}")
                dfs['rosters'] = _early_roster
                print(f" {len(_early_roster)} rows ✅")
            _rosters_tmp = dfs.get('rosters', pd.DataFrame())
            _ps_raw = fetch_pbp_player_stats(season, _rosters_tmp)
            if _ps_raw.empty:
                # Final fallback: prior year CSV
                print(f"  PBP failed — falling back to {prior} player stats")
                _ps_try2 = fetch_csv(f"{BASE}/player_stats/player_stats_{prior}.csv",
                                      f"player_stats {prior}")
                if not _ps_try2.empty:
                    _ps_reg2 = _ps_try2[_ps_try2['season_type']=='REG'] if 'season_type' in _ps_try2.columns else _ps_try2
                    if not _ps_reg2.empty:
                        _ps_raw = _ps_reg2.copy()
                        if 'dakota' in _ps_raw.columns and 'cpoe' not in _ps_raw.columns:
                            _ps_raw = _ps_raw.rename(columns={'dakota': 'cpoe'})

        dfs['player_stats'] = _ps_raw


    # ── PBP fallback: ensure dfs has 'pbp' for team stat aggregation ──
    if 'pbp' not in dfs or dfs.get('pbp', pd.DataFrame()).empty:
        print(f"  Fetching PBP {season} for team stat aggregation...", end='', flush=True)
        try:
            import gzip as _gz3, io as _io3
            _pbp_url3 = f"{NFLVERSE_BASE}/pbp/play_by_play_{season}.csv.gz"
            _pbp_r3 = SESSION.get(_pbp_url3, timeout=180, allow_redirects=True)
            if _pbp_r3.status_code == 200:
                with _gz3.open(_io3.BytesIO(_pbp_r3.content)) as _gz3f:
                    dfs['pbp'] = pd.read_csv(_gz3f, low_memory=False)
                print(f" {len(dfs['pbp']):,} plays ✅")
            else:
                print(f" HTTP {_pbp_r3.status_code} ❌")
        except Exception as _pbp_e3:
            print(f" {_pbp_e3} ❌")
    if not _ps_raw.empty:
        _SUM = [c for c in ['passing_yards','passing_tds','interceptions','passing_epa',
                             'rushing_yards','carries','rushing_tds','rushing_epa',
                             'receiving_yards','receptions','targets','receiving_tds',
                             'receiving_epa','completions','attempts','special_teams_tds']
                if c in _ps_raw.columns]
        _MEAN = [c for c in ['cpoe','target_share','air_yards_share','wopr','dakota']
                 if c in _ps_raw.columns]
        _AGG  = {c: 'sum' for c in _SUM}
        _AGG.update({c: 'mean' for c in _MEAN})
        # Flexible column detection — nflverse changes names between releases
        _name_col = next((c for c in ['player_display_name','player_name',
                          'full_name','display_name'] if c in _ps_raw.columns), None)
        _pos_col  = next((c for c in ['position','position_group','pos']
                          if c in _ps_raw.columns), None)
        _team_col = next((c for c in ['recent_team','team','posteam']
                          if c in _ps_raw.columns), None)
        _GRP = [c for c in [_name_col, _pos_col, _team_col] if c is not None]
        # Rename to standard column so downstream code always finds player_display_name
        if _name_col and _name_col != 'player_display_name' and _name_col in _ps_raw.columns:
            _ps_raw = _ps_raw.rename(columns={_name_col: 'player_display_name'})
            _GRP = ['player_display_name' if c == _name_col else c for c in _GRP]
        if 'week' in _ps_raw.columns:
            # Weekly data (from direct CSV) — aggregate and count games
            _AGG['week'] = 'nunique'
            _seas = (_ps_raw.groupby(_GRP).agg(_AGG)
                             .rename(columns={'week': 'games'})
                             .reset_index())
        else:
            # Already season-level (from PBP aggregation) — use directly
            _seas = _ps_raw.copy()
            if 'games' not in _seas.columns:
                _seas['games'] = 17  # default if not present
        dfs['player_season'] = _seas
        dfs['kicking'] = _ps_raw[_ps_raw['position'] == 'K'] if 'position' in _ps_raw.columns else pd.DataFrame()
        print(f"  Column detection: name={_name_col}, pos={_pos_col}, team={_team_col}")
        print(f"  Season totals built: {len(_seas):,} players  "
              f"(CPOE={'cpoe' in _seas.columns}, WOPR={'wopr' in _seas.columns})")
    else:
        dfs['player_season'] = dfs['kicking'] = pd.DataFrame()
        print("  ⚠ No player stats loaded — props will use static baselines")

    # Team stats (has EPA per play, off/def, pass/rush)
    dfs['team_stats'] = fetch_with_fallback(
        f"{BASE}/stats_team/stats_team_reg_{{s}}.csv",
        "team_stats_reg", prior)

    # PFR Advanced stats (broken tackles, YAC, pressure)
    print("  Fetching pfr_advstats (rush/rec/pass/def)...", end='', flush=True)
    _pfr_any = False
    for side in ['rush','rec','pass','def']:
        _pfr_url = f"{BASE}/pfr_advstats/advstats_season_{side}_{{s}}.csv"
        _pfr_df  = fetch_csv(_pfr_url.format(s=season), f"pfr_{side}", silent_404=True)
        if _pfr_df.empty:
            _pfr_df = fetch_csv(_pfr_url.format(s=prior), f"pfr_{side}", silent_404=True)
        dfs[f'pfr_{side}'] = _pfr_df
        if not _pfr_df.empty:
            _pfr_any = True
    if _pfr_any:
        print(" ✅ some data loaded")
    else:
        print(" ⚠ not available pre-season (will load once 2026 season starts)")

    # Rosters moved earlier — fetched before PBP aggregation for name resolution
    # Weekly rosters (depth chart / snap pct)
    dfs['weekly_rosters'] = fetch_csv(
        f"{BASE}/weekly_rosters/roster_weekly_{season}.csv",
        f"weekly_rosters {season}")

    # Injuries — CURRENT SEASON ONLY (prior-year injuries irrelevant)
    # Pre-season 404 is fine — ESPN auto-refresh handles live injury data
    # Injuries: skip nflverse pull — 2025 data is irrelevant for 2026 season.
    # Live 2026 injury statuses are pulled automatically from ESPN via
    # the model's EPA Update button (no file needed).
    print("  Injuries: skipping 2025 data — use ESPN auto-refresh for 2026 live injuries ✅")
    dfs['injuries'] = pd.DataFrame()

    # Depth charts
    # Depth charts: try current season first, fall back to baseline
    print("  Fetching depth_charts (current season)...", end='', flush=True)
    dfs['depth'] = pd.DataFrame()
    for _dc_yr in [CURRENT_SEASON, season]:
        _dc_df = fetch_csv(
            f"{BASE}/depth_charts/depth_charts_{_dc_yr}.csv",
            f"depth_charts {_dc_yr}", silent_404=True)
        if not _dc_df.empty:
            dfs['depth'] = _dc_df
            print(f" ✅ {_dc_yr} ({len(_dc_df):,} rows)")
            break
    if dfs['depth'].empty:
        print(f" ⚠ not available")

    # Schedules (upcoming games)
    dfs['schedules'] = fetch_csv(
        f"{BASE}/schedules/games.csv",
        "schedules (all seasons)")
    if not dfs['schedules'].empty and 'season' in dfs['schedules'].columns:
        dfs['schedules'] = dfs['schedules'][dfs['schedules']['season'].isin([season, CURRENT_SEASON])]

    # NGS — fetch .csv.gz files from nflverse (no nflreadpy required)
    # Note: separation & RYOE require full nflreadpy (Python >=3.10)
    # CPOE is already in player_stats above; other NGS fields loaded here when available
    print("  Fetching NGS (passing/rushing/receiving)...", end='', flush=True)
    import gzip as _gzip

    def _fetch_ngs_gz(year, stat_type):
        url = f"{BASE}/nextgen_stats/ngs_{year}_{stat_type}.csv.gz"
        try:
            r = SESSION.get(url, timeout=30, allow_redirects=True)
            r.raise_for_status()
            if len(r.content) < 2000:
                return pd.DataFrame()  # preseason stub — too small to be real data
            with _gzip.open(__import__('io').BytesIO(r.content)) as gz:
                df = pd.read_csv(gz, low_memory=False)
            # Keep full-season rows (week=0) and individual weeks
            return df
        except Exception:
            return pd.DataFrame()

    _ngs_loaded = False
    for _ngs_yr in [season, prior]:
        _ngs_p = _fetch_ngs_gz(_ngs_yr, 'passing')
        _ngs_r = _fetch_ngs_gz(_ngs_yr, 'rushing')
        _ngs_e = _fetch_ngs_gz(_ngs_yr, 'receiving')
        # Require at least 20 rows to consider it real season data
        if len(_ngs_p) >= 20 or len(_ngs_r) >= 20 or len(_ngs_e) >= 20:
            # Aggregate to season level
            def _agg_ngs(df, group_col='player_display_name', agg_cols=None):
                if df.empty or group_col not in df.columns: return df
                num_cols = [c for c in (agg_cols or df.select_dtypes('number').columns)
                            if c in df.columns and c != 'week']
                return df.groupby(group_col)[num_cols].mean().reset_index()
            dfs['ngs_pass'] = _agg_ngs(_ngs_p)
            dfs['ngs_rush'] = _agg_ngs(_ngs_r)
            dfs['ngs_rec']  = _agg_ngs(_ngs_e)
            print(f" ✅ {_ngs_yr} — pass:{len(dfs['ngs_pass'])} rush:{len(dfs['ngs_rush'])} rec:{len(dfs['ngs_rec'])}")
            _ngs_loaded = True
            break

    if not _ngs_loaded:
        print(f" ⚠ NGS not available via direct download — using player_stats proxies:")
        print(f"   QB: dakota (EPA-weighted CPOE) | WR: wopr, air_yards_share, target_share")
        print(f"   Missing: WR separation, RB RYOE (require Python >=3.10 + nflreadpy)")
        for k in ('ngs_pass', 'ngs_rush', 'ngs_rec'):
            dfs[k] = pd.DataFrame()

    return dfs

# ─────────────────────────────────────────────
# 2. BUILD TEAM EPA PROFILES
# ─────────────────────────────────────────────
def build_team_profiles(dfs, season):
    print(f"\n{'='*50}")
    print("BUILDING TEAM EPA PROFILES")
    print('='*50)
    teams = {}

    # Primary: stats_team has pre-computed EPA per play
    ts = dfs.get('team_stats', pd.DataFrame())
    if not ts.empty:
        # Column names from nflverse stats_team
        # Typical cols: team, season, games, off_epa, def_epa,
        #               off_pass_epa, off_rush_epa, def_pass_epa, def_rush_epa
        epa_cols = [c for c in ts.columns if 'epa' in c.lower()]
        print(f"  EPA columns found: {epa_cols}")

        # NFL season: ~1,050 offensive plays per team (65 plays/game × 17 games - ~1/3 pass neutral)
        # Per-play EPA = season_total_EPA / total_plays
        # nflverse stats_team may have season totals OR per-play depending on version
        PLAYS_PER_SEASON = 1050
        PASS_PLAYS = 600   # ~57% of plays are passes
        RUSH_PLAYS = 450   # ~43% of plays are rushes

        for _, row in ts.iterrows():
            abbr = str(row.get('team', row.get('team_abbr', row.get('recent_team', '')))).upper()
            if not abbr or abbr == 'NAN': continue
            # Normalise abbreviation variants (nflverse sometimes uses 'LA' for Rams)
            _TNORM = {'LA':'LAR','JAC':'JAX','KCC':'KC','SFO':'SF','NWE':'NE',
                      'NOR':'NO','GNB':'GB','TBB':'TB','SDG':'LAC','STL':'LAR'}
            abbr = _TNORM.get(abbr, abbr)
            games = safe_int(row.get('games', 17))
            if games == 0: games = 17

            # Try per-play columns first (preferred)
            off_epa_raw  = safe_float(row.get('offense_epa', row.get('off_epa',
                           row.get('passing_epa', 0) + row.get('rushing_epa', 0))))
            # nflverse team_stats_reg does NOT have defense_epa column directly
            # We derive defEPA from pts_against as a proxy:
            # Teams allowing 17 pts/game (elite) → negative defEPA; 28+ pts/game (bad) → positive
            # This gets replaced by PBP-computed defEPA after the team loop
            _pts_against = safe_float(row.get('pts_against', 0))
            _games_played = max(1, safe_int(row.get('games', 17)))
            _ptAllPG = _pts_against / _games_played
            # Convert pts/game to EPA/play proxy: (ptAllPG - 22.5) / 9.4 / 65 plays per game
            # Positive = bad defense (allows more than avg), Negative = good defense
            def_epa_raw  = (_ptAllPG - 22.5) / (9.4 * 65)  # ~0.015 per pt above avg
            pass_epa_raw = safe_float(row.get('offense_pass_epa', row.get('off_pass_epa',
                           row.get('passing_epa', 0))))
            rush_epa_raw = safe_float(row.get('offense_rush_epa', row.get('off_rush_epa',
                           row.get('rushing_epa', 0))))

            # If values look like season totals (abs > 10), convert to per-play
            def to_per_play(val, plays):
                # Always divide by plays — nflverse team_stats are season totals
                # Previously used abs>10 threshold which missed small totals (1-10)
                # causing per-play values to be wildly inflated (e.g. 8.0 vs 0.018)
                if plays and plays > 0:
                    result = val / plays
                else:
                    result = 0.0
                # Clamp to realistic per-play EPA range (-0.30 to +0.30)
                result = max(-0.30, min(0.30, result))
                return round(result, 4)

            teams[abbr] = {
                'games':      games,
                'offEPA':     to_per_play(off_epa_raw,  PLAYS_PER_SEASON),
                'defEPA':     round(max(-0.30, min(0.30, def_epa_raw)), 4),  # already per-play
                'offPassEPA': to_per_play(pass_epa_raw, PASS_PLAYS),
                'offRushEPA': to_per_play(rush_epa_raw, RUSH_PLAYS),
                # Validation: log any suspicious values
                # defPassEPA/defRushEPA: split total defEPA ~70/30 pass/rush
                'defPassEPA': round(max(-0.30, min(0.30, def_epa_raw * 0.70)), 4),
                'defRushEPA': round(max(-0.30, min(0.30, def_epa_raw * 0.30)), 4),
                'ptsPG':      round(safe_float(row.get('pts_for',     0)) / max(games,1), 1),
                'ptAllPG':    round(safe_float(row.get('pts_against',  0)) / max(games,1), 1),
                # passYdsPG / rushYdsPG populated after team loop from player_season
                'passYdsPG':  None,
                'rushYdsPG':  None,
                'source':     'stats_team'
            }
        print(f"  Built EPA profiles: {len(teams)} teams")

    # Fallback: aggregate from player_stats if team_stats empty
    if not teams and not dfs.get('player_stats', pd.DataFrame()).empty:
        ps = dfs['player_stats']
        ps = ps[ps['season'] == season] if 'season' in ps.columns else ps
        epa_agg = ps.groupby('recent_team').agg(
            offEPA=('offense_epa', 'mean') if 'offense_epa' in ps.columns else ('passing_epa', 'mean'),
            games=('week', 'nunique') if 'week' in ps.columns else ('passing_yards', 'count'),
        ).reset_index()
        for _, row in epa_agg.iterrows():
            abbr = str(row.get('recent_team', '')).upper()
            if not abbr or abbr == 'NAN': continue
            if abbr not in teams:
                teams[abbr] = {
                    'games': safe_int(row.get('games', 0)),
                    'offEPA': safe_float(row.get('offEPA', 0)),
                    'defEPA': 0, 'offPassEPA': 0, 'offRushEPA': 0,
                    'defPassEPA': 0, 'defRushEPA': 0,
                    'ptsPG': 0, 'ptAllPG': 0, 'source': 'player_stats_agg'
                }
        print(f"  Fallback: built {len(teams)} team profiles from player_stats")

    # ── Post-process: fill passYdsPG / rushYdsPG from PBP (not player_season) ──
    # PBP source is correct: CIN=249.6, ATL=217.8, not 530/370 from broken aggregation
    _TNORM_BTP2 = {'JAC':'JAX','LA':'LAR','WSH':'WAS','LVR':'LV','NWE':'NE','NOR':'NO',
                   'GNB':'GB','TBB':'TB','KCC':'KC','SFO':'SF'}
    def _np_btp(t): return _TNORM_BTP2.get(str(t).upper(), str(t).upper())
    _pbp_v = dfs.get('pbp', pd.DataFrame())
    if _pbp_v is not None and not _pbp_v.empty:
        _reg_v = _pbp_v[_pbp_v['season_type']=='REG'].copy() if 'season_type' in _pbp_v.columns else _pbp_v.copy()
        for _c in ['posteam','game_id']:
            if _c in _reg_v.columns: _reg_v[_c] = _reg_v[_c].astype(str)
        _reg_v['_t'] = _reg_v['posteam'].apply(_np_btp)
        # passYdsPG: average game-level passing yards per team
        _pass_v = _reg_v[_reg_v['pass_attempt']==1] if 'pass_attempt' in _reg_v.columns else _reg_v
        if 'passing_yards' in _pass_v.columns and 'game_id' in _pass_v.columns:
            _gm_pass = _pass_v.groupby(['_t','game_id'])['passing_yards'].sum().reset_index()
            _tm_pass = _gm_pass.groupby('_t')['passing_yards'].agg(['mean','count']).reset_index()
            _tm_pass.columns = ['team','passYdsPG','games']
            for _, _r in _tm_pass.iterrows():
                _a = _r['team']
                if _a in teams and _r['games'] >= 4:
                    teams[_a]['passYdsPG'] = round(float(_r['passYdsPG']), 1)
        # rushYdsPG: average game-level rushing yards per team
        _rush_v = _reg_v[_reg_v['rush_attempt']==1] if 'rush_attempt' in _reg_v.columns else _reg_v
        if 'rushing_yards' in _rush_v.columns and 'game_id' in _rush_v.columns:
            _gm_rush = _rush_v.groupby(['_t','game_id'])['rushing_yards'].sum().reset_index()
            _tm_rush = _gm_rush.groupby('_t')['rushing_yards'].agg(['mean','count']).reset_index()
            _tm_rush.columns = ['team','rushYdsPG','games']
            for _, _r in _tm_rush.iterrows():
                _a = _r['team']
                if _a in teams and _r['games'] >= 4:
                    teams[_a]['rushYdsPG'] = round(float(_r['rushYdsPG']), 1)
        _filled = sum(1 for t in teams.values() if t.get('passYdsPG') and t['passYdsPG'] != 225.0)
        print(f"  PBP-derived passYdsPG: {_filled}/32 teams filled (source: play_by_play)")
    else:
        print("  Warning: PBP not available for passYdsPG — using fallback")

    # Fill any remaining nulls with sensible league averages
    for _abbr, _t in teams.items():
        if _t.get('passYdsPG') is None: _t['passYdsPG'] = 225.0
        if _t.get('rushYdsPG') is None: _t['rushYdsPG'] = 112.0

    _pass_filled = sum(1 for t in teams.values() if t.get('passYdsPG') and t['passYdsPG'] != 225.0)
    print(f"  Team passYdsPG filled: {_pass_filled}/{len(teams)} teams from player stats")

    # ── PBP aggregation: rzOff, thirdDownPct, ptAllPG, per-team defEPA ──────
    pbp_raw = dfs.get('pbp', pd.DataFrame()) if dfs else pd.DataFrame()
    if pbp_raw is not None and not pbp_raw.empty:
        _TNORM2 = {'JAC':'JAX','LA':'LAR','WSH':'WAS','LVR':'LV','NWE':'NE',
                   'NOR':'NO','GNB':'GB','TBB':'TB','KCC':'KC','SFO':'SF'}
        def _np2(t): return _TNORM2.get(str(t).upper(), str(t).upper())
        try:
            reg_pbp = pbp_raw[pbp_raw['season_type']=='REG'].copy() if 'season_type' in pbp_raw.columns else pbp_raw.copy()
            for _col in ['posteam','home_team','away_team']:
                if _col in reg_pbp.columns: reg_pbp[_col] = reg_pbp[_col].astype(str)
            reg_pbp['_pos']  = reg_pbp['posteam'].apply(_np2)
            reg_pbp['_home'] = reg_pbp['home_team'].apply(_np2)
            reg_pbp['_away'] = reg_pbp['away_team'].apply(_np2)
            last_plays = reg_pbp.groupby('game_id').last().reset_index() if 'game_id' in reg_pbp.columns else pd.DataFrame()
            n_computed = 0
            for abbr, t in teams.items():
                t_norm = _np2(abbr)
                t_plays = reg_pbp[reg_pbp['_pos'] == t_norm]
                if t_plays.empty: continue
                # Red zone TD conversion rate
                if 'yardline_100' in reg_pbp.columns and 'drive' in reg_pbp.columns:
                    rz = t_plays[t_plays['yardline_100'] <= 20]
                    if len(rz) > 5:
                        # Composite TD flag: pass + rush + return touchdowns
                        # Verified against PBP: NE=43.9%, SEA=43.8%, KC=52.2% ✅
                        rz_copy = rz.copy()
                        rz_copy['_any_td'] = 0
                        for _tc in ['pass_touchdown','rush_touchdown','touchdown']:
                            if _tc in rz_copy.columns:
                                rz_copy['_any_td'] = rz_copy['_any_td'] + rz_copy[_tc].fillna(0)
                        rz_copy['_any_td'] = (rz_copy['_any_td'] > 0).astype(int)
                        rz_d = rz_copy.groupby(['game_id','drive'])['_any_td'].max().reset_index()
                        if not rz_d.empty:
                            t['rzOff'] = round(float(rz_d['_any_td'].sum()) / len(rz_d) * 100, 1)
                # Third down conversion rate
                if 'third_down_converted' in t_plays.columns and 'third_down_failed' in t_plays.columns:
                    td3c = float(t_plays['third_down_converted'].sum())
                    td3f = float(t_plays['third_down_failed'].sum())
                    if td3c + td3f > 0:
                        t['thirdDownPct'] = round(td3c / (td3c + td3f) * 100, 1)
                # Points for/against from final game scores
                if not last_plays.empty and 'home_score' in last_plays.columns:
                    hg = last_plays[last_plays['_home'] == t_norm]
                    ag = last_plays[last_plays['_away'] == t_norm]
                    pts_f = list(hg['home_score'].dropna()) + list(ag['away_score'].dropna())
                    pts_v = list(hg['away_score'].dropna()) + list(ag['home_score'].dropna())
                    if pts_f: t['ptsPG']   = round(sum(pts_f)/len(pts_f), 1)
                    if pts_v:
                        t['ptAllPG'] = round(sum(pts_v)/len(pts_v), 1)
                        # Override uniform defEPA with PBP-derived per-team value
                        _def_raw = (t['ptAllPG'] - 22.5) / (9.4 * 65)
                        t['defEPA'] = round(max(-0.30, min(0.30, _def_raw)), 4)
                n_computed += 1
            print(f"  PBP stats: {n_computed} teams — rzOff, thirdDownPct, ptsPG, ptAllPG, defEPA updated")
        except Exception as _pbp_err:
            print(f"  PBP aggregation error: {_pbp_err}")

    # ── v2.28: toMargin, kickerAdj, defSacks/defQBHits/defPassDef ─────────
    # These 5 fields existed in the output schema (save_output) but were
    # NEVER actually computed anywhere — every team silently got the
    # .get(field, 0/None) fallback. All the source data for these already
    # gets fetched elsewhere in this script (team_stats, kicking, pfr_def)
    # but was discarded unused. Wired up here instead of adding new fetches.
    #
    # Column names are detected flexibly (nflverse schema has varied
    # slightly between releases) and every miss prints which columns WERE
    # available, so a wrong guess here is visible in the console output
    # instead of silently producing another all-zero field.
    def _find_col(df, candidates):
        for c in candidates:
            if c in df.columns:
                return c
        return None

    # Self-contained abbreviation map — does NOT rely on the outer _TNORM,
    # which is only defined inside the `if not ts.empty:` branch above and
    # would raise NameError here if that branch didn't run this time.
    _ABBR_FIX = {'LA':'LAR','JAC':'JAX','KCC':'KC','SFO':'SF','NWE':'NE',
                 'NOR':'NO','GNB':'GB','TBB':'TB','SDG':'LAC','STL':'LAR'}
    def _norm_team_abbr(a):
        a = str(a).upper().strip()
        return _ABBR_FIX.get(a, a)

    # -- Turnover margin: takeaways (defense forced) minus giveaways (offense lost) --
    ts_df = dfs.get('team_stats', pd.DataFrame())
    if not ts_df.empty:
        team_c = _find_col(ts_df, ['team', 'team_abbr', 'recent_team'])
        int_thrown_c   = _find_col(ts_df, ['passing_interceptions', 'interceptions'])
        fum_lost_off_c = _find_col(ts_df, ['rushing_fumbles_lost', 'sack_fumbles_lost', 'receiving_fumbles_lost'])
        int_forced_c   = _find_col(ts_df, ['def_interceptions', 'defense_interceptions'])
        fum_forced_c   = _find_col(ts_df, ['def_fumbles', 'defense_fumbles', 'def_fumble_recovery_own'])
        if team_c and (int_thrown_c or int_forced_c):
            n_to = 0
            for _, row in ts_df.iterrows():
                abbr = _norm_team_abbr(row.get(team_c, ''))
                if abbr not in teams: continue
                giveaways = safe_float(row.get(int_thrown_c, 0)) + safe_float(row.get(fum_lost_off_c, 0))
                takeaways = safe_float(row.get(int_forced_c, 0)) + safe_float(row.get(fum_forced_c, 0))
                teams[abbr]['toMargin'] = round(takeaways - giveaways, 1)
                n_to += 1
            print(f"  Turnover margin: {n_to} teams computed (cols: giveaways={int_thrown_c}+{fum_lost_off_c}, takeaways={int_forced_c}+{fum_forced_c})")
        else:
            print(f"  ⚠ toMargin: no matching turnover columns in team_stats — available: {list(ts_df.columns)[:25]}")

        # -- Defensive pressure stats: sacks, QB hits, pass defense --
        sacks_c   = _find_col(ts_df, ['def_sacks', 'defense_sacks', 'sacks_suffered'])
        qbhits_c  = _find_col(ts_df, ['def_qb_hits', 'defense_qb_hits'])
        passdef_c = _find_col(ts_df, ['def_pass_defended', 'defense_pass_defended', 'passes_defended'])
        if team_c and (sacks_c or qbhits_c or passdef_c):
            n_def = 0
            for _, row in ts_df.iterrows():
                abbr = _norm_team_abbr(row.get(team_c, ''))
                if abbr not in teams: continue
                if sacks_c:   teams[abbr]['defSacks']   = safe_float(row.get(sacks_c, 0))
                if qbhits_c:  teams[abbr]['defQBHits']  = safe_float(row.get(qbhits_c, 0))
                if passdef_c: teams[abbr]['defPassDef'] = safe_float(row.get(passdef_c, 0))
                n_def += 1
            print(f"  Defensive pressure stats: {n_def} teams (cols: sacks={sacks_c}, qbHits={qbhits_c}, passDef={passdef_c})")
        else:
            print(f"  ⚠ defSacks/defQBHits/defPassDef: no matching columns in team_stats — trying pfr_def fallback")
            pfr_def_df = dfs.get('pfr_def', pd.DataFrame())
            if not pfr_def_df.empty:
                team_c2  = _find_col(pfr_def_df, ['team', 'tm', 'team_abbr'])
                sacks_c2 = _find_col(pfr_def_df, ['sacks', 'sk'])
                qbh_c2   = _find_col(pfr_def_df, ['qb_hits', 'qbhits', 'hits'])
                if team_c2 and (sacks_c2 or qbh_c2):
                    n_def2 = 0
                    for _, row in pfr_def_df.iterrows():
                        abbr = _norm_team_abbr(row.get(team_c2, ''))
                        if abbr not in teams: continue
                        if sacks_c2: teams[abbr]['defSacks']  = safe_float(row.get(sacks_c2, 0))
                        if qbh_c2:   teams[abbr]['defQBHits'] = safe_float(row.get(qbh_c2, 0))
                        n_def2 += 1
                    print(f"  Defensive pressure stats (pfr_def fallback): {n_def2} teams")
                else:
                    print(f"  ⚠ pfr_def fallback also missing expected columns — available: {list(pfr_def_df.columns)[:25]}")
            else:
                print(f"  ⚠ pfr_def is empty — defSacks/defQBHits/defPassDef will stay unset for this run")
    else:
        print("  ⚠ team_stats empty — toMargin/defSacks/defQBHits/defPassDef cannot be computed this run")

    # -- Kicker adjustment: team FG% vs. league average, small ±0.5 pt swing --
    # v2.29: switched from dfs['kicking'] (filtered from player_stats by
    # position=='K') to computing directly from play-by-play — confirmed via
    # a real run that dfs['kicking'] doesn't carry FG-specific columns at all
    # (player_stats is built around offensive skill positions, not kicking),
    # so that path always fell through to "no matching columns." PBP has a
    # real field_goal_result on every FG attempt and is already fetched
    # elsewhere in this function, so no new fetch is needed.
    pbp_df = dfs.get('pbp', pd.DataFrame())
    kicker_done = False
    if not pbp_df.empty and 'field_goal_result' in pbp_df.columns:
        team_c4 = _find_col(pbp_df, ['posteam', 'team'])
        if team_c4:
            fg_plays = pbp_df[pbp_df['field_goal_result'].notna()]
            if not fg_plays.empty:
                LEAGUE_AVG_FG_PCT = 0.85  # NFL long-run average FG%
                n_k = 0
                for abbr_raw, grp in fg_plays.groupby(team_c4):
                    abbr = _norm_team_abbr(abbr_raw)
                    if abbr not in teams: continue
                    att = len(grp)
                    if att < 5: continue  # too few attempts to trust a rate
                    made = (grp['field_goal_result'] == 'made').sum()
                    pct = made / att
                    teams[abbr]['kickerAdj'] = round(max(-0.5, min(0.5, (pct - LEAGUE_AVG_FG_PCT) * 5)), 2)
                    n_k += 1
                print(f"  Kicker adjustment: {n_k} teams computed from PBP field_goal_result ({len(fg_plays)} total FG attempts)")
                kicker_done = True

    if not kicker_done:
        # Fallback: try the old dfs['kicking'] path in case PBP is unavailable
        kicking_df = dfs.get('kicking', pd.DataFrame())
        if not kicking_df.empty:
            team_c3 = _find_col(kicking_df, ['recent_team', 'team', 'posteam'])
            fgm_c   = _find_col(kicking_df, ['fg_made', 'field_goals_made'])
            fga_c   = _find_col(kicking_df, ['fg_att', 'field_goals_attempted', 'fg_attempts'])
            if team_c3 and fgm_c and fga_c:
                LEAGUE_AVG_FG_PCT = 0.85
                n_k = 0
                for _, row in kicking_df.iterrows():
                    abbr = _norm_team_abbr(row.get(team_c3, ''))
                    if abbr not in teams: continue
                    att = safe_float(row.get(fga_c, 0))
                    if att < 5: continue
                    pct = safe_float(row.get(fgm_c, 0)) / att
                    teams[abbr]['kickerAdj'] = round(max(-0.5, min(0.5, (pct - LEAGUE_AVG_FG_PCT) * 5)), 2)
                    n_k += 1
                print(f"  Kicker adjustment (fallback via player_stats): {n_k} teams computed")
                kicker_done = n_k > 0
            else:
                print(f"  ⚠ kickerAdj: no FG columns in PBP or kicking data — kicking dataframe columns: {list(kicking_df.columns)[:25]}")
        if not kicker_done:
            print("  ⚠ kickerAdj: no usable field-goal data from PBP or kicking — will stay unset for this run "
                  "(fine per G-Money: only used for team-level scoring, no kicker props tracked)")

    return teams

# ─────────────────────────────────────────────
# 3. BUILD PLAYER PROFILES (NFL Data)
# ─────────────────────────────────────────────
def build_player_profiles(dfs, season):
    print(f"\n{'='*50}")
    print("BUILDING PLAYER PROFILES")
    print('='*50)
    players = {}

    # Main: player_season for season totals
    ps = dfs.get('player_season', dfs.get('player_stats', pd.DataFrame()))
    if not ps.empty:
        if 'season' in ps.columns:
            available = sorted(ps['season'].dropna().unique(), reverse=True)
            # Use BASELINE_SEASON if available, otherwise take the most recent season in the data
            best_season = next((s for s in [season, season-1, season-2] if s in available), available[0] if available else season)
            ps = ps[ps['season'] == best_season]
            print(f"  Using season {int(best_season)} player data ({len(ps):,} rows)")
        # Aggregate to true season totals — sum counting stats, keep max games
        if 'week' in ps.columns:
            _SUM_COLS = ['passing_yards','rushing_yards','receiving_yards','carries',
                         'completions','attempts','passing_tds','rushing_tds','receiving_tds',
                         'receptions','targets','interceptions','special_teams_tds']
            _MEAN_COLS = ['passing_epa','rushing_epa','receiving_epa','dakota',
                          'target_share','air_yards_share','wopr']
            _agg = {'position':'last','recent_team':'last','games':'max'}
            for c in _SUM_COLS:
                if c in ps.columns: _agg[c] = 'sum'
            for c in _MEAN_COLS:
                if c in ps.columns: _agg[c] = 'mean'
            ps = ps.groupby('player_display_name', as_index=False).agg(_agg)

        pos_map = {'QB':['passing_yards','completions','attempts','passing_tds',
                         'interceptions','passing_epa','dakota'],
                   'RB':['rushing_yards','carries','rushing_tds','rushing_epa',
                         'target_share','receiving_yards','receptions'],
                   'WR':['receiving_yards','receptions','targets','receiving_tds',
                         'receiving_epa','target_share','air_yards_share','wopr'],
                   'TE':['receiving_yards','receptions','targets','receiving_tds',
                         'receiving_epa','target_share'],
                   'K': ['special_teams_tds']}

        for _, row in ps.iterrows():
            name  = str(row.get('player_display_name') or 
                        row.get('player_name') or 
                        row.get('full_name') or '').strip()
            pos   = str(row.get('position') or row.get('position_group') or row.get('pos') or '').upper()
            team  = str(row.get('recent_team') or row.get('team') or row.get('posteam') or '').upper()
            games = safe_int(row.get('games', row.get('week', 1)))
            if not name or name == 'NAN' or pos not in pos_map: continue

            entry = {
                'name': name, 'pos': pos, 'team': team, 'games': games,
                'isRookie': False,  # filled later from rosters
                'nflGames': games,
                'source': 'nflverse_player_stats'
            }

            # Position-specific metrics
            if pos == 'QB':
                entry.update({
                    'passYdsPG':   safe_float(row.get('passing_yards', 0)) / max(games, 1),
                    'passTDsPG':   safe_float(row.get('passing_tds', 0)) / max(games, 1),
                    'intsPG':      safe_float(row.get('interceptions', 0)) / max(games, 1),
                    'cmp_pct':     safe_float(row.get('completions', 0)) / max(safe_float(row.get('attempts', 1)), 1) * 100,
                    'passEPA':     safe_float(row.get('passing_epa', 0)),
                    'dakota':      safe_float(row.get('dakota', 0)),
                    'rushYdsPG':   safe_float(row.get('rushing_yards', 0)) / max(games, 1),
                })
            elif pos == 'RB':
                entry.update({
                    'rushYdsPG':  safe_float(row.get('rushing_yards', 0)) / max(games, 1),
                    'rushTDsPG':  safe_float(row.get('rushing_tds', 0)) / max(games, 1),
                    'recYdsPG':   safe_float(row.get('receiving_yards', 0)) / max(games, 1),
                    'recsPG':     safe_float(row.get('receptions', 0)) / max(games, 1),
                    'targShare':  safe_float(row.get('target_share', 0)),
                    'rushEPA':    safe_float(row.get('rushing_epa', 0)),
                    'ypc':        safe_float(row.get('rushing_yards', 0)) / max(safe_float(row.get('carries', 1)), 1),
                    'catchRate':  safe_float(row.get('catchRate', None)),
                })
            elif pos in ('WR', 'TE'):
                entry.update({
                    'recYdsPG':   safe_float(row.get('receiving_yards', 0)) / max(games, 1),
                    'recsPG':     safe_float(row.get('receptions', 0)) / max(games, 1),
                    'recTDsPG':   safe_float(row.get('receiving_tds', 0)) / max(games, 1),
                    'targShare':  safe_float(row.get('target_share', 0)),
                    'airYdShare': safe_float(row.get('air_yards_share', 0)),
                    'wopr':       safe_float(row.get('wopr', 0)),
                    'recEPA':     safe_float(row.get('receiving_epa', 0)),
                    'ypr':        safe_float(row.get('receiving_yards', 0)) / max(safe_float(row.get('receptions', 1)), 1),
                    # Computed from weekly data — available when using nfl_player_stats_combined.csv
                    'catchRate':  safe_float(row.get('catchRate', None)),
                    'aDOT':       safe_float(row.get('aDOT', None)),
                    'yacPerRec':  safe_float(row.get('yacPerRec', None)),
                })
            elif pos == 'K':
                entry.update({
                    'fgPG': safe_float(row.get('special_teams_tds', 0)) / max(games, 1),
                })

            players[norm_name(name)] = entry

        # ── Build QB player stats from PBP passer data ─────────────────
        _pbp4 = dfs.get('pbp', pd.DataFrame())
        if _pbp4 is not None and not _pbp4.empty:
            try:
                _reg4 = _pbp4[_pbp4['season_type']=='REG'].copy() if 'season_type' in _pbp4.columns else _pbp4.copy()
                _reg4['passer_player_name'] = _reg4['passer_player_name'].astype(str)
                _pass4 = _reg4[_reg4['pass_attempt']==1] if 'pass_attempt' in _reg4.columns else _reg4
                _qb_grp = _pass4[_pass4['passer_player_name']!='nan'].groupby(
                    ['passer_player_name','posteam']).agg(
                    _py=('passing_yards','sum'),
                    _cmp=('complete_pass','sum'),
                    _att=('pass_attempt','sum'),
                    _gm=('game_id','nunique')
                ).reset_index()
                _qb_grp['_cpoe'] = _pass4.groupby('passer_player_name')['cpoe'].mean().reindex(_qb_grp['passer_player_name']).values if 'cpoe' in _pass4.columns else 0
                _qb_grp['_epa']  = _pass4.groupby('passer_player_name')['qb_epa'].mean().reindex(_qb_grp['passer_player_name']).values if 'qb_epa' in _pass4.columns else 0
                _TNORM4 = {'JAC':'JAX','LA':'LAR','WSH':'WAS','LVR':'LV','NWE':'NE','NOR':'NO',
                           'GNB':'GB','TBB':'TB','KCC':'KC','SFO':'SF'}
                def _np4(t): return _TNORM4.get(str(t).upper(), str(t).upper())
                _qb_added = 0
                for _, _r4 in _qb_grp.iterrows():
                    _nm4 = str(_r4['passer_player_name'])
                    _tm4 = _np4(_r4['posteam'])
                    _gm4 = int(_r4['_gm'])
                    if _gm4 < 2 or not _nm4: continue
                    _entry4 = {
                        'name': _nm4, 'pos': 'QB', 'team': _tm4, 'games': _gm4,
                        'passYdsPG': round(float(_r4['_py'])/max(_gm4,1), 1),
                        'cpoe':      round(float(_r4['_cpoe'] or 0), 4),
                        'passEPA':   round(float(_r4['_epa'] or 0), 4),
                        'cmpPct':    round(float(_r4['_cmp'])/max(float(_r4['_att']),1)*100, 1),
                        'isRookie': False, 'nflGames': _gm4, 'source': 'pbp_passer',
                    }
                    players[norm_name(_nm4)] = _entry4
                    _qb_added += 1
                print(f"  QB stats built from PBP: {_qb_added} passers (abbreviated names — resolved next step)")
            except Exception as _qbe:
                print(f"  QB PBP build error: {_qbe}")
        # ── Resolve abbreviated PBP names to full names via 2026 roster ─
        try:
            import urllib.request as _ur3, io as _io3, csv as _csv3
            _req3 = _ur3.Request(f"{NFLVERSE_BASE}/rosters/roster_2026.csv",
                                  headers={"User-Agent":"Mozilla/5.0"})
            with _ur3.urlopen(_req3, timeout=15) as _rr3:
                _r26 = _rr3.read().decode("utf-8", errors="replace")
            _gsis_full3 = {}
            for _row3 in _csv3.DictReader(_io3.StringIO(_r26)):
                _gid3  = (_row3.get("gsis_id","") or "").strip()
                _full3 = (_row3.get("full_name","") or "").strip()
                _pos3  = (_row3.get("depth_chart_position","") or _row3.get("position","")).strip().upper()
                _tm3   = (_row3.get("team","") or "").strip().upper()
                if _gid3 and _full3 and _pos3 in ("QB","RB","WR","TE"):
                    _gsis_full3[_gid3] = {"full": _full3, "pos": _pos3, "team": _tm3}
            _pbp3 = dfs.get("pbp", pd.DataFrame())
            _abbr_full3 = {}
            if not _pbp3.empty:
                for _cn3, _ci3 in [("receiver_player_name","receiver_player_id"),
                                    ("passer_player_name","passer_player_id"),
                                    ("rusher_player_name","rusher_player_id"),
                                    ("lateral_receiver_player_name","lateral_receiver_player_id"),
                                    ("kicker_player_name","kicker_player_id")]:
                    if _cn3 not in _pbp3.columns: continue
                    for _, _r3 in _pbp3[[_cn3,_ci3]].dropna().iterrows():
                        _gid4 = str(_r3[_ci3]).strip()
                        _nm4  = str(_r3[_cn3]).strip()
                        if _gid4 in _gsis_full3 and _nm4 not in _abbr_full3 and _nm4 != "nan":
                            _abbr_full3[_nm4] = _gsis_full3[_gid4]
            _renamed3 = 0; _new_pl = {}
            for _k3, _pd3 in players.items():
                # Check both exact key and abbreviated forms
                _fi3 = _abbr_full3.get(_k3) or _abbr_full3.get(
                    next((k for k in _abbr_full3 if norm_name(k)==_k3), ""))
                if _fi3 and _fi3["full"] != _k3:
                    _pd3["name"] = _fi3["full"]
                    if _fi3.get("team"): _pd3["team"] = _fi3["team"]
                    _new_pl[norm_name(_fi3["full"])] = _pd3; _renamed3 += 1
                else:
                    _new_pl[_k3] = _pd3
            players = _new_pl
            _added3 = 0
            for _gi5, _inf5 in _gsis_full3.items():
                _fn5     = _inf5["full"]
                _fn5_nrm = norm_name(_fn5)  # normalized key matches players dict
                if _fn5_nrm not in players and _fn5 not in players:
                    players[_fn5_nrm] = {"name":_fn5,"pos":_inf5["pos"],"team":_inf5["team"],
                                         "games":0,"isRookie":False,"nflGames":0,
                                         "source":"roster_2026_skeleton"}
                    _added3 += 1
            print(f"  Name resolution: {_renamed3} renamed, {_added3} skeleton entries added")
        except Exception as _ne3:
            print(f"  Warning: name resolution failed: {_ne3}")
        print(f"  Built {len(players)} player profiles from player_stats")

    # Augment with PFR advanced stats (broken tackles, YAC, pressure)
    pfr_rush = dfs.get('pfr_rush', pd.DataFrame())
    if not pfr_rush.empty:
        if 'season' in pfr_rush.columns:
            pfr_rush = pfr_rush[pfr_rush['season'] == season]
        for _, row in pfr_rush.iterrows():
            name = norm_name(row.get('player', row.get('pfr_player_name', '')))
            if name in players:
                players[name]['brokenTackles'] = safe_int(row.get('rushing_broken_tackles', 0))
                players[name]['yacPerCarry']   = safe_float(row.get('rushing_yards_after_contact_avg', 0))
                players[name]['pfrSource']     = True
        print(f"  PFR rush stats merged")

    pfr_rec = dfs.get('pfr_rec', pd.DataFrame())
    if not pfr_rec.empty:
        if 'season' in pfr_rec.columns:
            pfr_rec = pfr_rec[pfr_rec['season'] == season]
        for _, row in pfr_rec.iterrows():
            name = norm_name(row.get('player', row.get('pfr_player_name', '')))
            if name in players:
                players[name]['recBrokenTackles'] = safe_int(row.get('receiving_broken_tackles', 0))
                players[name]['dropPct']          = safe_float(row.get('receiving_drop_pct', 0))
        print(f"  PFR rec stats merged")

    pfr_pass = dfs.get('pfr_pass', pd.DataFrame())
    if not pfr_pass.empty:
        if 'season' in pfr_pass.columns:
            pfr_pass = pfr_pass[pfr_pass['season'] == season]
        for _, row in pfr_pass.iterrows():
            name = norm_name(row.get('player', row.get('pfr_player_name', '')))
            if name in players:
                players[name]['throwaways']  = safe_int(row.get('passing_throw_aways', 0))
                players[name]['spikes']      = safe_int(row.get('passing_spikes', 0))
                players[name]['drops']       = safe_int(row.get('passing_drops', 0))
                players[name]['pressurePct'] = safe_float(row.get('times_pressured_pct', 0))
        print(f"  PFR pass stats merged")

    # Augment with NGS data
    ngs_pass = dfs.get('ngs_pass', pd.DataFrame())
    if not ngs_pass.empty:
        # Aggregate to season level if weekly
        if 'week' in ngs_pass.columns:
            ngs_pass = ngs_pass.groupby('player_display_name').agg({
                'avg_time_to_throw': 'mean',
                'avg_completed_air_yards': 'mean',
                'avg_intended_air_yards': 'mean',
                'completion_percentage_above_expectation': 'mean',
                'aggressiveness': 'mean',
            }).reset_index()
        for _, row in ngs_pass.iterrows():
            name = norm_name(row.get('player_display_name', ''))
            if name in players:
                players[name]['cpoe']       = safe_float(row.get('completion_percentage_above_expectation', 0))
                players[name]['iay']        = safe_float(row.get('avg_intended_air_yards', 0))
                players[name]['aggPct']     = safe_float(row.get('aggressiveness', 0))
                players[name]['ngsSource']  = True
        print(f"  NGS passing stats merged")

    ngs_rush = dfs.get('ngs_rush', pd.DataFrame())
    if not ngs_rush.empty:
        if 'week' in ngs_rush.columns:
            ngs_rush = ngs_rush.groupby('player_display_name').agg({
                'rush_yards_over_expected_per_att': 'mean',
                'efficiency': 'mean',
                'percent_attempts_gte_8_defenders': 'mean',
            }).reset_index()
        for _, row in ngs_rush.iterrows():
            name = norm_name(row.get('player_display_name', ''))
            if name in players:
                players[name]['ryoe']       = safe_float(row.get('rush_yards_over_expected_per_att', 0))
                players[name]['rushEff']    = safe_float(row.get('efficiency', 0))
                players[name]['stackedPct'] = safe_float(row.get('percent_attempts_gte_8_defenders', 0))
        print(f"  NGS rushing stats merged")

    ngs_rec = dfs.get('ngs_rec', pd.DataFrame())
    if not ngs_rec.empty:
        if 'week' in ngs_rec.columns:
            ngs_rec = ngs_rec.groupby('player_display_name').agg({
                'avg_cushion': 'mean',
                'avg_separation': 'mean',
                'avg_intended_air_yards': 'mean',
                'avg_yac_above_expectation': 'mean',
                'percent_share_of_intended_air_yards': 'mean',
            }).reset_index()
        for _, row in ngs_rec.iterrows():
            name = norm_name(row.get('player_display_name', ''))
            if name in players:
                players[name]['separation'] = safe_float(row.get('avg_separation', 0))
                players[name]['cushion']    = safe_float(row.get('avg_cushion', 0))
                players[name]['yacAboveExp']= safe_float(row.get('avg_yac_above_expectation', 0))
                players[name]['airYdShare'] = safe_float(row.get('percent_share_of_intended_air_yards', 0))
        print(f"  NGS receiving stats merged")

    return players

# ─────────────────────────────────────────────
# 4. DETECT ROOKIES
# ─────────────────────────────────────────────
def detect_rookies(dfs, players, season):
    print(f"\n{'='*50}")
    print("DETECTING ROOKIES")
    print('='*50)
    rookies = set()

    rosters = dfs.get('rosters', pd.DataFrame())
    if not rosters.empty:
        entry_col = next((c for c in rosters.columns if 'entry_year' in c.lower() or 'rookie_year' in c.lower()), None)
        name_col  = next((c for c in rosters.columns if 'display_name' in c.lower() or 'full_name' in c.lower()), None)
        if entry_col and name_col:
            # Detect rookies for the CURRENT season being projected
            # CURRENT_SEASON is module-level global — must use globals(), not dir()
            # dir() only shows local vars; bug caused 2025 draftees flagged as 2026 rookies
            rook_season = str(CURRENT_SEASON) if 'CURRENT_SEASON' in globals() else str(int(season) + 1)
            rook_df = rosters[rosters[entry_col].astype(str) == rook_season]
            for _, row in rook_df.iterrows():
                rookies.add(norm_name(row.get(name_col, '')))
            print(f"  Detected {len(rookies)} rookies from rosters (entry_year={rook_season})")

    # Mark rookies in player profiles
    marked = 0
    for key, p in players.items():
        if key in rookies:
            p['isRookie'] = True
            marked += 1
    print(f"  Marked {marked} players as rookie in profiles")
    return rookies

# ─────────────────────────────────────────────
# 5. LOAD PROSPECT ANALYZER (College Fallback)
# ─────────────────────────────────────────────
def _load_prospect_json(filepath):
    """Load prospect data from JSON export (NFL_Prospect_Analyzer_Model_v1.json format)."""
    import json as _json
    college = {}
    with open(filepath, encoding='utf-8') as f:
        records = _json.load(f)
    if not isinstance(records, list):
        records = records.get('prospects', records.get('players', []))
    for p in records:
        name = str(p.get('name', '')).strip()
        if not name or name.lower() in ('nan', 'player', 'name'): continue
        pos  = str(p.get('pos', p.get('position', ''))).upper().strip()
        conf = str(p.get('conf', p.get('conference', ''))).lower().strip()
        conf_adj = CONF_ADJ.get(conf, CONF_ADJ.get(conf.split()[0] if conf else '', 0.70))
        entry = {
            'name':       name,
            'pos':        pos,
            'school':     str(p.get('college', p.get('school', ''))),
            'conf':       conf,
            'confAdj':    conf_adj,
            'nflTrans':   NFL_TRANS_FACTOR,
            'source':     'ProspectAnalyzer_JSON',
            # Core stats (already per-game from the JSON)
            'passYdsPG':  float(p.get('passYdsPG', 0) or 0),
            'rushYdsPG':  float(p.get('rushYdsPG', 0) or 0),
            'recYdsPG':   float(p.get('recYdsPG',  0) or 0),
            'recsPG':     float(p.get('recsPG',    0) or 0),
            'tdsPG':      float(p.get('tdsPG',     0) or 0),
            # Adjusted stats (conference × NFL translation factor)
            'passYdsPG_adj': round(float(p.get('passYdsPG', 0) or 0) * conf_adj * NFL_TRANS_FACTOR, 1),
            'rushYdsPG_adj': round(float(p.get('rushYdsPG', 0) or 0) * conf_adj * NFL_TRANS_FACTOR, 1),
            'recYdsPG_adj':  round(float(p.get('recYdsPG',  0) or 0) * conf_adj * NFL_TRANS_FACTOR, 1),
            # Bonus fields from NFL Prospect Analyzer
            'draftPick':     int(p.get('pick', 0)   or 0),
            'prospectScore': float(p.get('score', 0) or 0),
            'prospectTier':  str(p.get('tier', '')  or ''),
            'nflTeam':       str(p.get('team', '')  or '').upper(),
        }
        college[norm_name(name)] = entry
    print(f"  ✅ Loaded {len(college)} prospects from JSON "
          f"({sum(1 for v in college.values() if v['pos']=='QB')} QB / "
          f"{sum(1 for v in college.values() if v['pos']=='RB')} RB / "
          f"{sum(1 for v in college.values() if v['pos']=='WR')} WR / "
          f"{sum(1 for v in college.values() if v['pos']=='TE')} TE)")
    return college


def load_prospect_analyzer():
    print(f"\n{'='*50}")
    print("LOADING PROSPECT ANALYZER")
    print('='*50)
    college = {}

    if not PROSPECT_FILE.exists():
        print(f"  ⚠ File not found: {PROSPECT_FILE}")
        print(f"  Skipping Prospect Analyzer — will use CFB Reference only")
        return college

    # ── JSON path (preferred — no engine dependencies) ──────────────────────
    if PROSPECT_FILE.suffix.lower() == '.json':
        try:
            college = _load_prospect_json(PROSPECT_FILE)
            return college
        except Exception as e:
            print(f"  ❌ JSON read failed: {e}")
            return college

    # ── Excel path (fallback) ────────────────────────────────────────────────
    try:
        import openpyxl
        # Try multiple engines — openpyxl for .xlsx/.xlsm, xlrd for .xls
        xl = None
        xl = None
        for _engine in ['openpyxl', 'xlrd', 'calamine', None]:
            try:
                xl = pd.ExcelFile(str(PROSPECT_FILE), engine=_engine) if _engine \
                     else pd.ExcelFile(str(PROSPECT_FILE))
                print(f"  Opened with engine: {_engine or 'auto'}")
                break
            except Exception as _e:
                print(f"  ⚠ Engine '{_engine or 'auto'}' failed: {_e}")
                if _engine is None:
                    # Final fallback — try reading as CSV (some .xlsx are saved as CSV)
                    try:
                        df_csv = pd.read_csv(str(PROSPECT_FILE), nrows=5)
                        print(f"  ✅ File appears to be CSV — re-reading as CSV")
                        df_csv = pd.read_csv(str(PROSPECT_FILE))
                        xl = None   # signal CSV path
                        # Process CSV directly and return
                        name_col = next((c for c in df_csv.columns if 'name' in str(c).lower()), None)
                        if name_col:
                            for _, row in df_csv.iterrows():
                                name = str(row.get(name_col, '')).strip()
                                if name and name.lower() not in ('nan','player','name'):
                                    college[norm_name(name)] = {'name': name, 'source': 'ProspectAnalyzer_CSV'}
                            print(f"  Loaded {len(college)} prospects from CSV fallback")
                        return college
                    except Exception:
                        pass
                    print(f"  ❌ Cannot open Prospect Analyzer — check the file is not open in Excel")
                    print(f"     File: {PROSPECT_FILE}")
                    return college
                continue
        if xl is None:
            return college
        if xl is None:
            return college
        print(f"  Sheets: {xl.sheet_names}")

        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            if df.empty: continue

            # Find name column
            name_col = next((c for c in df.columns if 'name' in str(c).lower()), None)
            if not name_col: continue

            for _, row in df.iterrows():
                name = str(row.get(name_col, '')).strip()
                if not name or name.lower() in ('nan', 'player', 'name'): continue

                # Find conference and stats
                conf_col = next((c for c in df.columns if 'conf' in str(c).lower()), None)
                conf     = str(row.get(conf_col, '')).lower().strip() if conf_col else ''
                conf_adj = CONF_ADJ.get(conf, CONF_ADJ.get(conf.split()[0] if conf else '', 0.70))

                # Position detection
                pos_col = next((c for c in df.columns if 'pos' in str(c).lower()), None)
                pos     = str(row.get(pos_col, sheet[:2].upper())).upper().strip()[:2]

                entry = {
                    'name':      name,
                    'pos':       pos,
                    'school':    str(row.get(next((c for c in df.columns if 'school' in str(c).lower() or 'team' in str(c).lower()), name_col), '')),
                    'conf':      conf,
                    'confAdj':   conf_adj,
                    'nflTrans':  NFL_TRANS_FACTOR,
                    'source':    'ProspectAnalyzer',
                }

                # QB stats
                for col_hint, key in [('pass_yds','passYdsPG'),('yards','passYdsPG'),
                                       ('td','passTDsPG'),('att','attempts')]:
                    match = next((c for c in df.columns if col_hint in str(c).lower()), None)
                    if match: entry[key] = safe_float(row.get(match, 0))

                # RB stats
                for col_hint, key in [('rush','rushYdsPG'),('carry','ypc'),
                                       ('rec','recYdsPG'),('yac','yac')]:
                    match = next((c for c in df.columns if col_hint in str(c).lower()), None)
                    if match: entry[key] = safe_float(row.get(match, 0))

                # Apply adjustments
                for stat in ['passYdsPG','rushYdsPG','recYdsPG']:
                    if stat in entry:
                        entry[f'{stat}_adj'] = round(entry[stat] * conf_adj * NFL_TRANS_FACTOR, 1)

                college[norm_name(name)] = entry

        print(f"  Loaded {len(college)} prospects from Prospect Analyzer")
    except Exception as e:
        print(f"  ❌ Error reading Prospect Analyzer: {e}")

    return college

# ─────────────────────────────────────────────
# 6. CFB REFERENCE FALLBACK
# ─────────────────────────────────────────────
def fetch_cfb_ref_stats(name, pos, year=2024):
    """Scrape CFB Reference for a single player's college stats"""
    try:
        from bs4 import BeautifulSoup
        search_name = name.replace(' ', '+')
        url = f"https://www.sports-reference.com/cfb/search/search.fcgi?search={search_name}"
        r = SESSION.get(url, timeout=10)
        time.sleep(1.2)  # Respectful rate limiting

        soup = BeautifulSoup(r.text, 'html.parser')
        # Find player link
        player_link = soup.find('a', href=re.compile(r'/cfb/players/[a-z]+-\d+\.html'))
        if not player_link: return {}

        player_url = 'https://www.sports-reference.com' + player_link['href']
        r2 = SESSION.get(player_url, timeout=10)
        time.sleep(1.2)

        soup2 = BeautifulSoup(r2.text, 'html.parser')

        # Get conference
        conf = ''
        for td in soup2.find_all('td', {'data-stat': 'conf_abbr'}):
            conf = td.text.strip().lower()
        conf_adj = CONF_ADJ.get(conf, 0.70)

        # Get most recent season passing stats
        stats = {'conf': conf, 'confAdj': conf_adj, 'nflTrans': NFL_TRANS_FACTOR, 'source': 'CFBRef'}
        table = soup2.find('table', id='passing') or soup2.find('table', id='rushing') or soup2.find('table', id='receiving')
        if table:
            rows = table.find('tbody').find_all('tr')
            for row in reversed(rows):  # most recent season
                yr = row.find('td', {'data-stat': 'year_id'})
                if yr and str(year) in yr.text:
                    for stat, key, divisor in [
                        ('pass_yds','passYdsPG',None), ('rush_yds','rushYdsPG',None),
                        ('rec_yds','recYdsPG',None), ('pass_td','passTDsPG',None),
                        ('g','games',None)
                    ]:
                        td = row.find('td', {'data-stat': stat})
                        if td:
                            raw = safe_float(td.text.replace(',',''))
                            stats[key] = raw

                    # Per-game
                    games = stats.get('games', 1)
                    for stat_k in ['passYdsPG','rushYdsPG','recYdsPG','passTDsPG']:
                        if stat_k in stats and games > 0:
                            stats[stat_k] /= games
                            stats[f'{stat_k}_adj'] = round(stats[stat_k] * conf_adj * NFL_TRANS_FACTOR, 1)
                    break

        return stats
    except Exception as e:
        return {'error': str(e)}

# ─────────────────────────────────────────────
# 7. APPLY ROOKIE BLENDING
# ─────────────────────────────────────────────
def apply_rookie_blending(players, college_data, rookies):
    print(f"\n{'='*50}")
    print("APPLYING ROOKIE BLENDING (College → NFL Transition)")
    print('='*50)

    blended_count = 0
    for key, p in players.items():
        if not p.get('isRookie', False): continue

        nfl_games = p.get('nflGames', 0)
        pos       = p.get('pos', 'WR')
        cw, nw    = get_blend(nfl_games, pos)

        # Find college data
        col = college_data.get(key, {})

        # Try CFB Reference if not in Prospect Analyzer
        if not col and nfl_games == 0:
            print(f"  Fetching CFB Ref for rookie: {p['name']}...")
            col = fetch_cfb_ref_stats(p['name'], pos)
            if col:
                college_data[key] = col  # cache it

        if not col:
            p['rookieNote'] = f'Rookie — no college data found. Using league avg baseline.'
            continue

        # Blending
        p['collegeData'] = col
        p['blendWeights'] = {'college': cw, 'nfl': nw, 'games': nfl_games}

        # Apply blended projections per stat
        stat_pairs = {  # (college_key_adj, nfl_player_key, blend_result_key)
            'passYdsPG_adj': 'passYdsPG',
            'rushYdsPG_adj': 'rushYdsPG',
            'recYdsPG_adj':  'recYdsPG',
        }
        for col_key, nfl_key in stat_pairs.items():
            col_val = col.get(col_key, col.get(col_key.replace('_adj',''), 0))
            nfl_val = p.get(nfl_key, 0)
            if col_val or nfl_val:
                blended = round(cw * col_val + nw * nfl_val, 1)
                p[f'{nfl_key}_blended'] = blended

        conf_label = col.get('conf', 'Unknown').upper()
        conf_adj   = col.get('confAdj', 0.70)
        p['rookieNote'] = (
            f"ROOKIE — {nfl_games} NFL games played. "
            f"Blend: {int(cw*100)}% college ({conf_label}, {conf_adj:.2f}× adj) "
            f"+ {int(nw*100)}% NFL."
        )
        blended_count += 1

    print(f"  Applied blending to {blended_count} rookies")

# ─────────────────────────────────────────────
# 8. BUILD INJURY / DEPTH CHART LAYER
# ─────────────────────────────────────────────

def extract_starters(dfs):
    dc = dfs.get('depth', pd.DataFrame())
    if dc.empty: return {}
    starters = {}
    pos_map = {'QB':'qb','RB':'rbTop','WR':'wr1','TE':'te'}

    # Support both old nflverse format and new 2026 ESPN-based format:
    #   old: full_name | depth_chart_position | depth_team
    #   new: player_name | pos_abb             | pos_rank
    team_c = next((c for c in dc.columns if c in ['club_code','team','posteam']), None)
    pos_c  = next((c for c in dc.columns if c in ['position','pos','depth_chart_position','pos_abb']), None)
    name_c = next((c for c in dc.columns
                   if c == 'player_name'
                   or 'full_name' in c.lower()
                   or 'display_name' in c.lower()), None)
    # Prefer pos_rank over pos_slot — pos_rank=1 means starter at that position,
    # pos_slot is a formation-slot index (QB might be slot 9, not 1)
    dep_c  = (next((c for c in dc.columns if 'depth_team' in c.lower()), None)
              or ('pos_rank' if 'pos_rank' in dc.columns else None)
              or ('pos_slot' if 'pos_slot' in dc.columns else None))

    if not all([team_c, pos_c, name_c, dep_c]):
        print(f"  ⚠ Depth chart column mismatch — "
              f"team={team_c} pos={pos_c} name={name_c} depth={dep_c}")
        print(f"    Available columns: {list(dc.columns)}")
        return starters

    # Normalise nflverse abbreviation quirks (e.g. 'LA' -> 'LAR' for Rams)
    _DEPTH_ABBR = {'LA':'LAR','JAC':'JAX','KCC':'KC','SFO':'SF','NWE':'NE',
                   'NOR':'NO','GNB':'GB','TBB':'TB','SDG':'LAC','STL':'LAR'}
    try:
        # pos_rank/depth_team may be int, float, or string — normalise to numeric
        dep_numeric = pd.to_numeric(dc[dep_c], errors='coerce').fillna(999)
        starters_df = dc[dep_numeric == 1].copy()
    except Exception as e:
        print(f"  ⚠ Starters filter failed: {e}")
        return starters

    # For WR: collect unique starters per team before assigning wr1/wr2
    # (formation-based depth charts can have multiple pos_rank=1 WR rows per team)
    wr_per_team = {}
    for _, row in starters_df.iterrows():
        team = str(row.get(team_c,'')).upper().strip()
        team = _DEPTH_ABBR.get(team, team)
        pos  = str(row.get(pos_c, '')).upper().strip()
        name = str(row.get(name_c,'')).strip()
        if pos == 'WR' and team and name and name.lower() != 'nan':
            wr_per_team.setdefault(team, [])
            if name not in wr_per_team[team]:
                wr_per_team[team].append(name)

    for _, row in starters_df.iterrows():
        team = str(row.get(team_c,'')).upper().strip()
        team = _DEPTH_ABBR.get(team, team)
        pos  = str(row.get(pos_c, '')).upper().strip()
        name = str(row.get(name_c,'')).strip()
        if not team or not name or name.lower() == 'nan': continue
        if pos not in pos_map or pos == 'WR': continue   # WR handled separately
        if team not in starters: starters[team] = {}
        if pos_map[pos] not in starters[team]:
            starters[team][pos_map[pos]] = name

    for team, wr_names in wr_per_team.items():
        if team not in starters: starters[team] = {}
        if wr_names:               starters[team]['wr1'] = wr_names[0]
        if len(wr_names) > 1:      starters[team]['wr2'] = wr_names[1]
        if len(wr_names) > 2:      starters[team]['wr3'] = wr_names[2]

    n = sum(1 for v in starters.values() if v)
    print(f"  Extracted starters for {n} teams from depth charts")
    return starters

def build_injury_status(dfs, season):
    print(f"\n{'='*50}")
    print("BUILDING INJURY / DEPTH CHART STATUS")
    print('='*50)
    injury_map = {}

    inj = dfs.get('injuries', pd.DataFrame())
    if not inj.empty:
        if 'season' in inj.columns:
            inj = inj[inj['season'] == CURRENT_SEASON]  # current season only
        # Get most recent week per player
        if 'week' in inj.columns:
            inj = inj.sort_values('week', ascending=False).drop_duplicates(
                subset=['full_name'] if 'full_name' in inj.columns else ['player_id'])

        name_col   = next((c for c in inj.columns if 'full_name' in c.lower() or 'display_name' in c.lower()), None)
        status_col = next((c for c in inj.columns if 'status' in c.lower()), None)
        report_col = next((c for c in inj.columns if 'report' in c.lower() and 'status' in c.lower()), None)

        if name_col and (status_col or report_col):
            for _, row in inj.iterrows():
                name   = norm_name(row.get(name_col, ''))
                status = str(row.get(report_col or status_col, '')).upper()
                injury_map[name] = {
                    'status': status,
                    'week':   safe_int(row.get('week', 0))
                }
            print(f"  Loaded {len(injury_map)} injury statuses")

    return injury_map

# ─────────────────────────────────────────────
# 9. BUILD UPCOMING GAMES
# ─────────────────────────────────────────────
def build_upcoming_games(dfs, season):
    from datetime import datetime, timezone
    print(f"\n{'='*50}")
    print("BUILDING UPCOMING GAMES")
    print('='*50)

    games_out = []
    sched = dfs.get('schedules', pd.DataFrame())
    if sched.empty: return games_out

    now = datetime.now(timezone.utc)
    if 'game_type' in sched.columns:
        sched = sched[sched['game_type'] == 'REG']  # regular season only

    date_col = next((c for c in sched.columns if 'date' in c.lower() or 'gameday' in c.lower()), None)
    if date_col:
        sched[date_col] = pd.to_datetime(sched[date_col], errors='coerce', utc=True)
        # Include games from CURRENT_SEASON that are upcoming OR all current-season games
        # (for pre-season, show full upcoming schedule even if dates are future)
        curr_season_games = sched[sched['season'] == CURRENT_SEASON] if 'season' in sched.columns else sched
        upcoming_by_date  = sched[sched[date_col] >= now]
        # Combine: prefer upcoming by date, fill with current season games
        upcoming = pd.concat([upcoming_by_date, curr_season_games]).drop_duplicates().head(50)
    else:
        upcoming = sched[sched['season'] == CURRENT_SEASON].head(50) if 'season' in sched.columns else sched.tail(20)

    for _, row in upcoming.iterrows():
        home = str(row.get('home_team', '')).upper()
        away = str(row.get('away_team', '')).upper()
        if not home or not away: continue
        games_out.append({
            'home': home, 'away': away,
            'week': safe_int(row.get('week', 0)),
            'date': str(row.get(date_col, ''))[:10] if date_col else '',
            'gameId': str(row.get('game_id', '')),
        })
    print(f"  {len(games_out)} upcoming games found")
    return games_out

# ─────────────────────────────────────────────
# 9b. BUILD HEALTHY ROSTER SHARES (co-game target distribution)
# ─────────────────────────────────────────────
def build_healthy_roster_shares(dfs, teams):
    """
    Compute per-team target share distributions from 2025 weekly data
    across four roster states: full strength, WR1 out, WR2 out, TE out.
    These populate window._healthyShares in the model and enable
    context-aware projection adjustments based on active roster.
    """
    print(f"\n{'='*50}")
    print("BUILDING HEALTHY ROSTER SHARES")
    print('='*50)

    weekly = dfs.get('weekly_rosters', pd.DataFrame())
    # weekly_rosters has snap/depth data but not target data
    # We need the player_stats weekly file — check if it was loaded
    week_df = dfs.get('player_stats_weekly', pd.DataFrame())
    if week_df.empty:
        # Not in dfs — try loading from disk directly
        import os as _os2
        _script_dir = _os2.path.dirname(_os2.path.abspath(__file__))
        _week_candidates = [
            'stats_player_week_2025.csv',
            'stats_player_week_2024.csv',
            'nfl_player_stats_weekly.csv',
        ]
        for _wc in _week_candidates:
            _wp = _os2.path.join(_script_dir, _wc)
            if _os2.path.exists(_wp):
                try:
                    week_df = pd.read_csv(_wp, low_memory=False)
                    print(f"  Weekly stats loaded from disk: {_wc} ({len(week_df):,} rows)")
                    break
                except Exception as _we2:
                    print(f"  Could not load {_wc}: {_we2}")
    if week_df.empty:
        print("  ⚠ No weekly player stats found — place stats_player_week_2025.csv")
        print("    in the same folder as this script and re-run.")
        return {}

    # Filter to regular season
    if 'week' in week_df.columns:
        week_df = week_df[week_df['week'].between(1, 18)].copy()
    if week_df.empty:
        print("  ⚠ No regular season weekly data")
        return {}

    # Standardise team column
    if 'team' not in week_df.columns and 'recent_team' in week_df.columns:
        week_df = week_df.rename(columns={'recent_team': 'team'})

    skill_pos = ['WR', 'TE', 'RB']

    def _norm_pn(n):
        return re.sub(r'[^a-z ]', '', str(n).lower()).strip()

    def get_game_ids(tm_df, player_name):
        """Games a player appeared in — exact, normalized, and last-name fallbacks."""
        if not player_name: return set()
        # 1. Exact match
        rows = tm_df[tm_df['player_display_name'] == player_name]
        if not rows.empty:
            return set(rows['game_id']) if 'game_id' in rows.columns else set()
        # 2. Normalized match (handles D.J. Moore vs DJ Moore)
        pnorm = _norm_pn(player_name)
        rows = tm_df[tm_df['player_display_name'].apply(_norm_pn) == pnorm]
        if not rows.empty:
            return set(rows['game_id']) if 'game_id' in rows.columns else set()
        # 3. Last-name fallback — skip generational suffixes (Jr./Sr./III/II)
        _SUFFIXES = {'jr','jr.','sr','sr.','ii','iii','iv','v'}
        _parts = player_name.split()
        last = _parts[-1]
        if last.lower() in _SUFFIXES and len(_parts) > 2:
            last = _parts[-2]  # real last name before suffix
        if len(last) > 3:
            rows = tm_df[tm_df['player_display_name'].str.contains(last, case=False, na=False)]
        return set(rows['game_id']) if ('game_id' in rows.columns and not rows.empty) else set()

    def share_map(game_ids, tm_df, min_tgts=4, min_games=2):
        """Return {F.Lastname: avg_target_share} for given game_ids."""
        if not game_ids: return {}
        sub = tm_df[tm_df['game_id'].isin(game_ids)]
        if sub.empty: return {}
        t_tot = sub.groupby('game_id')['targets'].sum()
        p_tgt = sub[sub['position'].isin(skill_pos)].groupby(
            ['player_display_name', 'game_id'])['targets'].sum().reset_index()
        p_tgt = p_tgt.merge(t_tot.rename('tt'), on='game_id')
        p_tgt['ts'] = p_tgt['targets'] / p_tgt['tt'].where(p_tgt['tt'] > 0)
        agg = p_tgt.groupby('player_display_name').agg(
            avg_ts=('ts', 'mean'),
            games=('game_id', 'nunique'),
            total_tgts=('targets', 'sum'))
        agg = agg[(agg['total_tgts'] >= min_tgts) & (agg['games'] >= min_games)]
        out = {}
        for pname, row in agg.iterrows():
            parts = pname.split()
            if len(parts) >= 2:
                key = parts[0][0] + '.' + parts[-1]
                out[key] = round(float(row['avg_ts']), 4)
        return out

    healthy_shares = {}

    # nflverse uses different abbreviations in weekly vs season files
    _ABBR_MAP = {'LAR':'LA', 'JAC':'JAX', 'KCC':'KC', 'SFO':'SF',
                 'NWE':'NE', 'GNB':'GB', 'NOR':'NO', 'TBB':'TB',
                 'LVR':'LV', 'OAK':'LV'}

    for team_abbr, team_data in teams.items():
        wr1 = team_data.get('wr1', '')
        wr2 = team_data.get('wr2', '')
        wr3 = team_data.get('wr3', '')
        te  = team_data.get('te',  '')
        if not wr1: continue

        # Try both standard and nflverse-variant abbreviations
        _csv_abbr = _ABBR_MAP.get(team_abbr, team_abbr)
        tm = week_df[week_df['team'].isin([team_abbr, _csv_abbr])].copy()
        if tm.empty: continue

        all_g  = set(tm['game_id'].unique()) if 'game_id' in tm.columns else set()
        wr1_g  = get_game_ids(tm, wr1)
        wr2_g  = get_game_ids(tm, wr2) if wr2 else set()
        wr3_g  = get_game_ids(tm, wr3) if wr3 else set()
        te_g   = get_game_ids(tm, te)  if te  else set()

        if len(wr1_g) < 3:
            # 2026 WR1 wasn't on this team in 2025 (trade/FA) or has name mismatch
            # Fall back to the player with most targets on this team in 2025
            if 'targets' in tm.columns and 'player_display_name' in tm.columns:
                top_tgt = tm[tm['position'].isin(['WR','TE'])].groupby(
                    'player_display_name')['targets'].sum().sort_values(ascending=False)
                if not top_tgt.empty:
                    fallback_wr1 = top_tgt.index[0]
                    fallback_g   = get_game_ids(tm, fallback_wr1)
                    if len(fallback_g) >= 3:
                        wr1   = fallback_wr1
                        wr1_g = fallback_g
                        print(f"  {team_abbr}: 2026 WR1 not in 2025 data — "
                              f"using actual top target: {fallback_wr1}")
            if len(wr1_g) < 3: continue

        # ── Rookie injection: if a 2026 rookie is WR1/WR2/WR3, ──────────────
        # they have no 2025 weekly data. Assign positional default shares
        # so they appear in the share map with realistic starting estimates.
        # These get overridden once actual 2026 game data exists.
        _ROOKIE_DEFAULT_TS = {'wr1': 0.20, 'wr2': 0.14, 'wr3': 0.10, 'te': 0.13}
        _rookie_overrides  = {}  # pbp_key → default share for rookies
        for _rpos, _rname in [('wr1',wr1),('wr2',wr2),('wr3',wr3),('te',te)]:
            if not _rname: continue
            _rg = get_game_ids(tm, _rname)
            if len(_rg) < 2:  # <2 games in 2025 = likely true 2026 rookie/new
                _rparts = _rname.split()
                if len(_rparts) >= 2:
                    _rpbp = _rparts[0][0] + '.' + _rparts[-1]
                    _rookie_overrides[_rpbp] = _ROOKIE_DEFAULT_TS[_rpos]

        # Full strength: WR1 + WR2 + TE all played
        # Only intersect when all named players have real game data
        # Avoids empty full_g when a name mismatch gives 0 TE/WR2 games
        _use_wr2 = wr2 and len(wr2_g) >= 3
        _use_wr3 = wr3 and len(wr3_g) >= 3
        _use_te  = te  and len(te_g)  >= 3
        if _use_wr2 and _use_wr3 and _use_te: full_g = wr1_g & wr2_g & wr3_g & te_g
        elif _use_wr2 and _use_te:             full_g = wr1_g & wr2_g & te_g
        elif _use_wr2 and _use_wr3:            full_g = wr1_g & wr2_g & wr3_g
        elif _use_wr2:                         full_g = wr1_g & wr2_g
        elif _use_te:                          full_g = wr1_g & te_g
        else:                                  full_g = wr1_g

        # WR1 out: WR1 absent, WR2 (if exists) still active
        wr1_out_g = all_g - wr1_g
        if wr2: wr1_out_g &= wr2_g

        # WR2 out: WR2 absent, WR1 active
        wr2_out_g = (wr1_g - wr2_g) if wr2 else set()

        # WR3 out: WR3 absent, WR1+WR2 active
        wr3_out_g = (wr1_g - wr3_g) if wr3 else set()
        # TE out: TE absent, WR1 active
        te_out_g  = (wr1_g - te_g)  if te  else set()

        full_shares = share_map(full_g, tm)
        # Inject rookie defaults for players with no 2025 data
        for _rpbp, _rts in _rookie_overrides.items():
            if _rpbp not in full_shares:  # don't override real data
                full_shares[_rpbp] = _rts
        if not full_shares: continue

        healthy_shares[team_abbr] = {
            'anchor':  wr1,
            'wr2':     wr2 or None,
            'wr3':     wr3 or None,
            'te':      te  or None,
            'season':  2025,
            'source':  'co_game_weekly_2025',
            'shares':  full_shares,   # legacy key — full strength
            'full_strength': {
                'games':  len(full_g),
                'shares': full_shares
            },
            'wr1_out': {
                'games':  len(wr1_out_g),
                'shares': share_map(wr1_out_g, tm, min_games=1)
            },
            'wr2_out': {
                'games':  len(wr2_out_g),
                'shares': share_map(wr2_out_g, tm, min_games=1)
            },
            'wr3_out': {
                'games':  len(wr3_out_g),
                'shares': share_map(wr3_out_g, tm, min_games=1)
            },
            'te_out':  {
                'games':  len(te_out_g),
                'shares': share_map(te_out_g,  tm, min_games=1)
            },
        }

    print(f"  Healthy roster shares built: {len(healthy_shares)} teams")
    for t in ['DAL','CIN','PHI','DET']:
        if t in healthy_shares:
            hs = healthy_shares[t]
            full = hs['full_strength']
            wr1o = hs['wr1_out']
            teo  = hs['te_out']
            print(f"  {t}: full={full['games']}g, wr1_out={wr1o['games']}g, te_out={teo['games']}g")
    return healthy_shares


# ─────────────────────────────────────────────
# 10. SAVE JSON
# ─────────────────────────────────────────────
def save_output(teams, players, injury_map, games, college_data, season, roster_changes=None, starters=None, coaching_context=None, dfs=None, healthy_roster_shares=None):
    print(f"\n{'='*50}")
    print("SAVING OUTPUT JSON")
    print('='*50)

    # Convert player dict from norm_name keys to display names
    players_out = {p['name']: p for p in players.values() if p.get('name')}

    # Merge starters into teams dict so the model receives them
    if starters:
        for abbr, starter_data in starters.items():
            if abbr in teams:
                teams[abbr].update(starter_data)
            else:
                # Team had no EPA data but has depth-chart starters
                teams[abbr] = starter_data

    # Cap passYdsPG at 345 (NFL record pace) — fixes CIN=530/SF=479 doubling bug
    for _ac9, _tc9 in (teams or {}).items():
        if isinstance(_tc9, dict) and (_tc9.get('passYdsPG') or 0) > 345:
            _tc9['passYdsPG'] = 345.0

    output = {
        'generated':      datetime.now().isoformat(),
        'season':         season,
        'baseline_season':BASELINE_SEASON,
        'current_season': CURRENT_SEASON,
        'is_preseason':   CURRENT_SEASON > BASELINE_SEASON,
        'source':         'nflverse + PFR advstats + NGS + ProspectAnalyzer',
        'teams':      teams,
        'healthy_roster_shares': healthy_roster_shares or {},
        'players':    players_out,
        'injuries':   {players.get(k, {}).get('name', k): v
                       for k, v in injury_map.items()},
        'games':      games,
        'rookieData': {p.get('name',''):
                       {'collegeData': p.get('collegeData',{}),
                        'blendWeights': p.get('blendWeights',{}),
                        'rookieNote': p.get('rookieNote','')}
                       for p in players.values() if p.get('isRookie')},
        'coaching_context': coaching_context or {},
        # Player movement (FA/trade) impact per team, computed by build_roster_changes()
        # v2.26: this was computed and printed to console every run but never actually
        # written to the output JSON — the whole calculation was being discarded.
        'roster_changes': roster_changes or {},
        # Team stats derived from 2025 PBP — used by model for GI card scores
        'team_stats_2025': {
            abbr: {
                'rzOff':        t.get('rzOff'),
                'thirdDownPct': t.get('thirdDownPct'),
                'ptAllPG':      t.get('ptAllPG'),
                'ptsPG':        t.get('ptsPG'),
                'toMargin':     t.get('toMargin', 0),
                'kickerAdj':    t.get('kickerAdj', 0),
                'defSacks':     t.get('defSacks'),
                'defQBHits':    t.get('defQBHits'),
                'defPassDef':   t.get('defPassDef'),
            }
            for abbr, t in teams.items() if isinstance(t, dict)
        },
        'backtest_games': build_backtest_games(dfs, teams, BASELINE_SEASON) if dfs is not None else [],
        'meta': {
            'teamCount':    len(teams),
            'playerCount':  len(players_out),
            'rookieCount':  sum(1 for p in players.values() if p.get('isRookie')),
            'injuryCount':  len(injury_map),
            'gameCount':    len(games),
        }
    }

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=2, default=str)

    size_kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"  ✅ Saved: {OUTPUT_FILE}")
    print(f"  Size: {size_kb:.1f} KB")
    print(f"  Teams: {len(teams)} | Players: {len(players_out)} | Rookies: {output['meta']['rookieCount']}")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

# ══════════════════════════════════════════════════════════════════════
# ROSTER CHANGE DETECTION & SCORING IMPACT MODULE
# Added to nfl_edge_data_pull.py
# ══════════════════════════════════════════════════════════════════════

def fetch_roster_current(season):
    """Fetch the most recent available roster for CURRENT_SEASON."""
    urls = [
        f"{NFLVERSE_BASE}/rosters/roster_{season}.csv",
        f"{NFLVERSE_BASE}/rosters/roster_{season - 1}.csv",  # fallback
    ]
    for url in urls:
        yr = url.split('roster_')[1].split('.')[0]
        df = fetch_csv(url, f"rosters {yr}")
        if not df.empty:
            # Keep only the latest week available
            if 'week' in df.columns:
                df = df[df['week'] == df['week'].max()]
            return df
    return pd.DataFrame()


def build_prod_dict(player_stats_df):
    """
    Convert a player_season-style DataFrame into {norm_name: {pos, team,
    raw_name, <stat fields>, games}}.

    v2.26: extracted from build_roster_changes() so the exact same logic can
    build a production dict for ANY season's stats — 2025 (baseline) or 2024
    (fallback) — instead of duplicating this ~50 lines twice.
    """
    prod = {}
    if player_stats_df is None or player_stats_df.empty:
        return prod

    def get_pos_col(df):
        for c in ['position', 'pos', 'depth_chart_position']:
            if c in df.columns: return c
        return None

    def get_name_col(df):
        for c in ['full_name', 'player_display_name', 'player_name', 'name']:
            if c in df.columns: return c
        return None

    def get_team_col(df):
        for c in ['team', 'recent_team', 'team_abbr', 'club_code']:
            if c in df.columns: return c
        return None

    def norm_name(n):
        # v2.29: also strip generational suffixes (Jr/Sr/II/III/IV/V) —
        # see the module-level norm_name() for why this matters.
        s = re.sub(r'[^a-z ]', '', str(n).lower().strip()) if n else ''
        s = re.sub(r'\b(jr|sr|ii|iii|iv|v)\b', '', s)
        return re.sub(r'\s+', ' ', s).strip()

    ABBR_MAP = {
        'JAC':'JAX','KCC':'KC','LVR':'LV','SFO':'SF','NWE':'NE','NOR':'NO',
        'GNB':'GB','TBB':'TB','SDG':'LAC','STL':'LAR','LA':'LAR','LAR':'LAR',
    }
    def norm_abbr(a):
        a = str(a).upper().strip()
        return ABBR_MAP.get(a, a)

    SKILL = {'QB', 'RB', 'WR', 'TE'}
    name_c = get_name_col(player_stats_df)
    team_c = get_team_col(player_stats_df)
    pos_c  = get_pos_col(player_stats_df)
    if not all([name_c, team_c, pos_c]):
        return prod

    for _, row in player_stats_df.iterrows():
        name = norm_name(row.get(name_c, ''))
        pos  = str(row.get(pos_c, '')).upper().strip()
        if not name or pos not in SKILL: continue
        games = max(safe_int(row.get('games', 1)), 1)

        def pg(col, div=None):
            v = safe_float(row.get(col, 0))
            return round(v / (div or games), 2)

        if pos == 'QB':
            att = max(safe_int(row.get('attempts', 0)), 1)
            prod[name] = {
                'pos': 'QB', 'team': norm_abbr(row.get(team_c, '')),
                'raw_name': str(row.get(name_c, '')).strip(),
                'passYdsPG':  pg('passing_yards'),
                'passTDsPG':  pg('passing_tds'),
                'epaPerDB':   round(safe_float(row.get('passing_epa', 0)) / att, 4),
                'cpoe':       safe_float(row.get('cpoe', 0)),
                'games':      games,
            }
        elif pos == 'RB':
            carries = max(safe_int(row.get('carries', 0)), 1)
            prod[name] = {
                'pos': 'RB', 'team': norm_abbr(row.get(team_c, '')),
                'raw_name': str(row.get(name_c, '')).strip(),
                'rushYdsPG': pg('rushing_yards'),
                'recYdsPG':  pg('receiving_yards'),
                'recsPG':    pg('receptions'),
                'rushEPA':   round(safe_float(row.get('rushing_epa', 0)) / carries, 4),
                'games':     games,
            }
        elif pos in ('WR', 'TE'):
            tgt = max(safe_int(row.get('targets', 0)), 1)
            prod[name] = {
                'pos': pos, 'team': norm_abbr(row.get(team_c, '')),
                'raw_name': str(row.get(name_c, '')).strip(),
                'recYdsPG':  pg('receiving_yards'),
                'recsPG':    pg('receptions'),
                'targetsPG': pg('targets'),
                'recEPA':    round(safe_float(row.get('receiving_epa', 0)) / tgt, 4),
                'games':     games,
            }
    return prod


def fetch_player_season_stats_for_year(year):
    """
    Fetch and aggregate ONE specific year's player_stats CSV to season
    totals, independent of BASELINE_SEASON/CURRENT_SEASON.

    v2.26: added for the 2024 fallback in build_roster_changes() — when a
    2026 confirmed starter didn't play enough 2025 games to trust (e.g. a
    season-ending injury), this gives their most recent trustworthy season
    instead of silently falling all the way to a generic positional average.

    Reuses the exact fetch URL pattern already used elsewhere in this script
    (see the BASELINE_SEASON fetch above) — same source, just parameterized
    by year so it can be called for any season on demand.
    """
    df = fetch_csv(f"{NFLVERSE_BASE}/player_stats/player_stats_{year}.csv",
                    f"player_stats {year} (fallback)", silent_404=True)
    if df.empty:
        return pd.DataFrame()
    df_reg = df[df['season_type'] == 'REG'] if 'season_type' in df.columns else df
    if df_reg.empty:
        return pd.DataFrame()
    if 'dakota' in df_reg.columns and 'cpoe' not in df_reg.columns:
        df_reg = df_reg.rename(columns={'dakota': 'cpoe'})
    elif 'passing_cpoe' in df_reg.columns:
        df_reg = df_reg.rename(columns={'passing_cpoe': 'cpoe'})

    SUM = [c for c in ['passing_yards','passing_tds','interceptions','passing_epa',
                        'rushing_yards','carries','rushing_tds','rushing_epa',
                        'receiving_yards','receptions','targets','receiving_tds',
                        'receiving_epa','completions','attempts']
           if c in df_reg.columns]
    MEAN = [c for c in ['cpoe'] if c in df_reg.columns]
    AGG = {c: 'sum' for c in SUM}
    AGG.update({c: 'mean' for c in MEAN})

    name_c = next((c for c in ['player_display_name','player_name','full_name','display_name']
                   if c in df_reg.columns), None)
    pos_c  = next((c for c in ['position','position_group','pos'] if c in df_reg.columns), None)
    team_c = next((c for c in ['recent_team','team','posteam'] if c in df_reg.columns), None)
    if not all([name_c, pos_c, team_c]):
        return pd.DataFrame()

    GRP = [name_c, pos_c, team_c]
    if name_c != 'player_display_name':
        df_reg = df_reg.rename(columns={name_c: 'player_display_name'})
        GRP = ['player_display_name' if c == name_c else c for c in GRP]

    if 'week' in df_reg.columns:
        AGG['week'] = 'nunique'
        seas = df_reg.groupby(GRP).agg(AGG).rename(columns={'week': 'games'}).reset_index()
    else:
        seas = df_reg.copy()
        if 'games' not in seas.columns:
            seas['games'] = 17
    return seas


def build_roster_changes(dfs, players, season, starters_2026=None):
    """
    Compare BASELINE_SEASON rosters vs CURRENT_SEASON rosters.
    Identify key player movements (FA signings, trades, cuts).
    Compute per-team offensive scoring impact of those changes.

    Returns dict: {team_abbr: {offAdj, defAdj, changes: [...], summary}}
    """
    print(f"\n{'='*50}")
    print("BUILDING ROSTER CHANGES & SCORING IMPACT")
    print('='*50)

    roster_base = dfs.get('rosters', pd.DataFrame())     # BASELINE_SEASON
    roster_curr = dfs.get('roster_curr', pd.DataFrame()) # CURRENT_SEASON
    player_stats = dfs.get('player_season', pd.DataFrame())

    # v2.27 fix: these used to be hard-exit guards that returned {} if either
    # roster snapshot was empty. That was correct for the OLD design, where
    # comparing roster_base vs roster_curr WAS the core mechanism. It's no
    # longer accurate — the core logic below runs on the depth chart
    # (extract_starters) plus player production data, neither of which
    # needs these two dataframes. Left as hard exits, this was silently
    # discarding the entire roster_changes computation whenever either
    # roster fetch happened to fail, even though everything the real logic
    # needs was still available. Now: missing rosters only disables the
    # supplementary "movers" list, not the core starter-vs-starter analysis.
    can_detect_movers = not roster_base.empty and not roster_curr.empty
    if not can_detect_movers:
        print("  ⚠ Roster snapshot(s) missing — skipping the movers list, "
              "but starter-vs-starter comparison will still run from depth charts")

    # ── Normalize team abbreviations ──────────────────────────────────
    ABBR_MAP = {
        'JAC':'JAX','KCC':'KC','LVR':'LV',
        'SFO':'SF','NWE':'NE','NOR':'NO','GNB':'GB',
        'TBB':'TB','SDG':'LAC','STL':'LAR',
        'LA':'LAR',   # nflverse depth chart uses 'LA' for Rams
        'LAR':'LAR',
    }
    def norm_abbr(a):
        a = str(a).upper().strip()
        return ABBR_MAP.get(a, a)

    # ── Name normalizer ───────────────────────────────────────────────
    def norm_name(n):
        # v2.29: also strip generational suffixes (Jr/Sr/II/III/IV/V) —
        # see the module-level norm_name() for why this matters.
        s = re.sub(r'[^a-z ]', '', str(n).lower().strip()) if n else ''
        s = re.sub(r'\b(jr|sr|ii|iii|iv|v)\b', '', s)
        return re.sub(r'\s+', ' ', s).strip()

    # ── Build position column (handle multiple column names) ──────────
    def get_pos_col(df):
        for c in ['position','pos','depth_chart_position']:
            if c in df.columns: return c
        return None

    # ── Build name column ─────────────────────────────────────────────
    def get_name_col(df):
        for c in ['full_name','player_display_name','player_name','name']:
            if c in df.columns: return c
        return None

    # ── Build team column ─────────────────────────────────────────────
    def get_team_col(df):
        for c in ['team','recent_team','team_abbr','club_code']:
            if c in df.columns: return c
        return None

    SKILL = {'QB','RB','WR','TE'}

    # ── Extract skill position players from each roster ───────────────
    def extract_skill(df, label):
        pos_c  = get_pos_col(df)
        name_c = get_name_col(df)
        team_c = get_team_col(df)
        if not all([pos_c, name_c, team_c]):
            print(f"  ⚠ {label}: missing columns (pos={pos_c} name={name_c} team={team_c})")
            return {}
        skill_df = df[df[pos_c].str.upper().isin(SKILL)].copy()
        result = {}
        for _, row in skill_df.iterrows():
            name = norm_name(row[name_c])
            team = norm_abbr(row[team_c])
            pos  = str(row[pos_c]).upper().strip()
            if name and team:
                result[name] = {'team': team, 'pos': pos, 'raw_name': str(row[name_c]).strip()}
        print(f"  {label}: {len(result)} skill position players")
        return result

    base_players = extract_skill(roster_base, f"Baseline ({BASELINE_SEASON})")
    curr_players = extract_skill(roster_curr, f"Current ({CURRENT_SEASON})")

    # ── Build player production dicts: 2025 (baseline) + 2024 (fallback) ──
    # v2.26: refactored into build_prod_dict() so both seasons use identical
    # logic; 2024 is fetched fresh here since nothing else in this script
    # pulls that specific year (verified no other fetch touches player_stats
    # for CURRENT_SEASON - 2, so this can't collide/conflict with anything).
    prod = build_prod_dict(player_stats)
    print(f"  Production data (2025): {len(prod)} skill players")

    prod_2024_raw = fetch_player_season_stats_for_year(CURRENT_SEASON - 2)
    prod_2024 = build_prod_dict(prod_2024_raw)
    print(f"  Production data (2024 fallback): {len(prod_2024)} skill players")

    # ── Identify movers: players whose team changed (context/display only —
    # the offAdj calculation below no longer keys off this list directly) ──
    movers = []
    for norm, curr_info in curr_players.items():
        if norm in base_players:
            base_info = base_players[norm]
            if base_info['team'] != curr_info['team'] and base_info['team'] and curr_info['team']:
                movers.append({
                    'norm_name':  norm,
                    'raw_name':   curr_info['raw_name'],
                    'pos':        curr_info['pos'],
                    'from_team':  base_info['team'],
                    'to_team':    curr_info['team'],
                    'prod':       prod.get(norm, {}),
                })
    print(f"  Roster movers detected: {len(movers)}")

    # ── 2025 "actual" baseline per team/position, RANKED ────────────────
    # What each position actually produced for that team in 2025. Ranked
    # (not just single-highest) so WR1 and WR2 each compare against their
    # own 2025 counterpart by rank — comparing both 2026 WR slots against
    # the SAME single top-2025-producer was a real bug caught in testing:
    # it made a healthy WR2 signing look like a big loss just because the
    # team's 2025 WR1 happened to be excellent.
    starters_2025_ranked = {}  # team → pos → [entries sorted desc by metric]
    pos_metric = {'QB':'passYdsPG', 'RB':'rushYdsPG', 'WR':'recYdsPG', 'TE':'recYdsPG'}
    for norm, p in prod.items():
        team = p.get('team', '')
        pos  = p.get('pos', '')
        if not team or pos not in SKILL: continue
        metric = pos_metric.get(pos, 'recYdsPG')
        val    = p.get(metric, 0)
        starters_2025_ranked.setdefault(team, {}).setdefault(pos, []).append(
            {'norm_name': norm, 'val': val, 'prod': p})
    for team in starters_2025_ranked:
        for pos in starters_2025_ranked[team]:
            starters_2025_ranked[team][pos].sort(key=lambda e: e['val'], reverse=True)

    def old_baseline_for_slot(team, pos, rank_idx):
        """rank_idx: 0 for the primary slot (qb/rbTop/wr1/te), 1 for wr2, etc."""
        entries = starters_2025_ranked.get(team, {}).get(pos, [])
        return entries[rank_idx]['prod'] if rank_idx < len(entries) else None

    # ── 2026 confirmed starters, from the REAL depth chart ─────────────
    # v2.26: this is the core fix — previously "who replaced a departed
    # player" was guessed at via a positional average. Now we look up the
    # actual 2026 depth-chart starter for each slot and use THEIR OWN
    # trusted history, never blending in a different player's stats (e.g.
    # an injury fill-in who isn't part of the 2026 plan).
    # v2.26: accept the already-computed starters_map from main() instead of
    # recomputing it — main() already calls extract_starters(dfs) once,
    # calling it again here produced an identical result but wasted a pass
    # over the depth chart data. Still falls back to computing it directly
    # if this function is ever called standalone without that parameter.
    starters_2026 = starters_2026 or extract_starters(dfs)  # {team: {qb, rbTop, wr1, wr2, wr3, te}}

    MIN_GAMES_TRUST = 3  # own-2025-games threshold before falling back to 2024

    def resolve_2026_starter_prod(player_name):
        """Own 2025 stats if trusted (>= MIN_GAMES_TRUST games) -> own 2024 stats -> None.
        Deliberately never substitutes a DIFFERENT player's stats."""
        norm = norm_name(player_name)
        p25 = prod.get(norm)
        if p25 and p25.get('games', 0) >= MIN_GAMES_TRUST:
            return p25, '2025'
        p24 = prod_2024.get(norm)
        if p24:
            return p24, '2024'
        return None, 'estimate'

    # ── Scoring impact coefficients (pts per unit) ────────────────────
    # Calibrated from NFL analytics research
    IMPACT = {
        'QB': {
            # EPA per dropback: each +0.01 delta → ~+0.35 pts/game
            'epaPerDB':   35.0,
            # CPOE: each +1% → +0.18 pts
            'cpoe':        0.18,
            # Backup QB (no data): assume –6 pts relative to starter
            'nodata':     -6.0,
        },
        'RB': {
            # Rush yards/game: each +10 → +0.8 pts
            'rushYdsPG':  0.08,
            # Rush EPA/carry: each +0.01 → +0.4 pts
            'rushEPA':     4.0,
            'nodata':     -2.0,
        },
        'WR': {
            # Rec yards/game: each +10 → +0.6 pts
            'recYdsPG':   0.06,
            'nodata':     -1.5,
        },
        'TE': {
            'recYdsPG':   0.06,
            'nodata':     -1.0,
        },
    }

    def player_pts(p, pos):
        """Convert a player's stats dict to approximate pts/game contribution."""
        if not p or not pos: return 0
        imp = IMPACT.get(pos, {})
        pts = 0
        if pos == 'QB':
            pts += p.get('epaPerDB', 0) * imp.get('epaPerDB', 0)
            pts += p.get('cpoe', 0) * imp.get('cpoe', 0)
        elif pos == 'RB':
            pts += p.get('rushYdsPG', 0) * imp.get('rushYdsPG', 0)
            pts += p.get('rushEPA', 0) * imp.get('rushEPA', 0)
        else:
            pts += p.get('recYdsPG', 0) * imp.get('recYdsPG', 0)
        return round(pts, 2)

    avg_pos_pts = {'QB': 4.2, 'RB': 2.1, 'WR': 1.8, 'TE': 1.2}

    # ── Compute team-level roster adjustments: 2026 confirmed starter's
    # own trusted rate vs. what the position actually produced in 2025 ────
    SLOT_POS = {'qb': 'QB', 'rbTop': 'RB', 'wr1': 'WR', 'wr2': 'WR', 'te': 'TE'}
    SLOT_RANK = {'qb': 0, 'rbTop': 0, 'wr1': 0, 'wr2': 1, 'te': 0}
    team_adj = {}

    for team, slots in starters_2026.items():
        for slot_key, pos in SLOT_POS.items():
            new_starter_name = slots.get(slot_key)
            if not new_starter_name:
                continue

            new_prod, new_source = resolve_2026_starter_prod(new_starter_name)
            new_pts = player_pts(new_prod, pos) if new_prod else avg_pos_pts.get(pos, 2.0)

            old_prod  = old_baseline_for_slot(team, pos, SLOT_RANK[slot_key])
            old_pts   = player_pts(old_prod, pos) if old_prod else avg_pos_pts.get(pos, 2.0)
            old_name  = old_prod.get('raw_name') if old_prod else '(no 2025 data)'

            delta = round(new_pts - old_pts, 2)

            if team not in team_adj:
                team_adj[team] = {'offAdj': 0.0, 'changes': []}
            team_adj[team]['offAdj'] = round(team_adj[team]['offAdj'] + delta, 2)

            # Only log a change entry when the starter is actually different,
            # or the position's expected output shifted meaningfully even
            # with the same player (e.g. falling back to 2024 data).
            same_player = old_prod and norm_name(old_prod.get('raw_name', '')) == norm_name(new_starter_name)
            if not same_player or abs(delta) >= 0.5:
                team_adj[team]['changes'].append({
                    'slot':           slot_key,
                    'pos':            pos,
                    'old_starter':    old_name,
                    'new_starter':    new_starter_name,
                    'new_data_source': new_source,   # '2025' | '2024' | 'estimate'
                    'impact_pts':     delta,
                })

    # ── Cap adjustments at ±8 pts (prevent extreme outliers) ──────────
    # This is a proportionality guard, not a real ceiling on how much roster
    # turnover COULD matter — every other team-scoring factor in this model
    # (QB efficiency, kicker quality, red zone, turnovers, special teams,
    # 3rd down) tops out well under this, so an uncapped sum across 5
    # position slots could otherwise dominate the whole scoring formula.
    for team, data in team_adj.items():
        data['offAdj'] = round(max(-8.0, min(8.0, data['offAdj'])), 2)

    # ── Print summary ─────────────────────────────────────────────────
    significant = {t: d for t, d in team_adj.items() if abs(d['offAdj']) >= 1.0}
    print(f"\n  Teams with significant roster impact (±1+ pts):")
    for team, data in sorted(significant.items(), key=lambda x: abs(x[1]['offAdj']), reverse=True):
        sign = '+' if data['offAdj'] > 0 else ''
        changes_str = ', '.join(
            f"{c['slot']} {c['new_starter'].split()[-1]} ({c['new_data_source']})"
            for c in data['changes'][:3])
        print(f"    {team:<4} {sign}{data['offAdj']:.1f} pts | "
              f"{len(data['changes'])} changes: {changes_str}")

    return team_adj




# ─────────────────────────────────────────────
# COACHING OVERRIDES (mid-season adjustments)
# ─────────────────────────────────────────────
# Update this dict when a coordinator is fired or scheme changes dramatically mid-season.
# Leave empty {} at season start — EPA data self-corrects within 2-3 weeks naturally.
# adj_pass_rate: multiplier on expected pass rate  (1.15 = +15% vs baseline, 0.85 = run-heavy shift)
# adj_epa:       flat EPA per play overlay          (+0.03 = new scheme estimated uplift)
# week_of_change: when to start applying the override (0 = from the start of season)
COACHING_OVERRIDES_2026 = {
    # Example — uncomment and edit when a real mid-season change happens:
    # 'MIN': {'type': 'OC', 'week_of_change': 9,
    #         'adj_pass_rate': 1.12, 'adj_epa': 0.02,
    #         'note': 'New OC hired Week 9 — more aggressive passing scheme'},
    # 'NYJ': {'type': 'HC', 'week_of_change': 6,
    #         'adj_pass_rate': 0.88, 'adj_epa': -0.01,
    #         'note': 'HC fired Week 6 — interim running more conservative offense'},
}

def build_coaching_context(dfs):
    """
    Build coaching context for the JSON output.
    Combines static COACHING_OVERRIDES_2026 with any auto-detected indicators.
    The model reads 'coaching_context' from the JSON and applies it as an overlay
    on top of EPA-based projections via computeSchemeMatchupAdj().
    """
    coaching = {}

    for abbr, override in COACHING_OVERRIDES_2026.items():
        coaching[abbr] = {
            'type':            override.get('type', 'OC'),
            'week_of_change':  override.get('week_of_change', 0),
            'adj_pass_rate':   override.get('adj_pass_rate', 1.0),
            'adj_epa':         override.get('adj_epa', 0.0),
            'note':            override.get('note', ''),
            'source':          'manual_override',
        }

    if coaching:
        print(f"  Coaching overrides applied: {list(coaching.keys())}")
    else:
        print("  No mid-season coaching overrides — EPA data handles adjustments naturally")

    return coaching


# ─────────────────────────────────────────────
# IN-SEASON NGS UPDATE FUNCTION
# ─────────────────────────────────────────────
def update_inseason_ngs(season=None, output_file=None):
    """
    In-season weekly update: pulls the latest NGS and player stats
    for the CURRENT season and patches the existing nfl_model_data.json.

    Run every Tuesday after Monday Night Football:
        python nfl_edge_data_pull.py --update

    What it updates vs the full pull:
        UPDATES:  player EPA, CPOE, RYOE, separation, target share, air yards share
                  team EPA profiles, coaching context
        SKIPS:    rosters, depth charts, schedules, prospect analyzer, rookie blending
                  (those change rarely and are covered by the weekly full pull)

    The update patches ONLY players and teams that have new data — it does not
    wipe data for players who haven't played yet this season.
    """
    import json as _json
    _season = season or CURRENT_SEASON
    _outfile = output_file or OUTPUT_FILE

    print(f"\n{'='*50}")
    print(f"IN-SEASON NGS UPDATE — {_season} Season")
    print(f"{'='*50}")

    # ── Load existing JSON ──────────────────────────────────────────────────
    existing = {}
    if Path(_outfile).exists():
        with open(_outfile, encoding='utf-8') as f:
            existing = _json.load(f)
        print(f"  Loaded existing JSON: {len(existing.get('teams',{}))} teams, "
              f"{len(existing.get('players',{}))} players")
    else:
        print(f"  ⚠ No existing JSON at {_outfile} — run full pull first")
        return

    import gzip as _gzip_upd

    def _nfl_fetch_ps(yr):
        """Fetch player stats for a season via direct nflverse URL."""
        url = f"{NFLVERSE_BASE}/player_stats/player_stats_{yr}.csv"
        try:
            r = SESSION.get(url, timeout=30, allow_redirects=True)
            r.raise_for_status()
            from io import StringIO
            df = pd.read_csv(StringIO(r.text), low_memory=False)
            df = df[df['season_type']=='REG'] if 'season_type' in df.columns else df
            if 'passing_cpoe' in df.columns:
                df = df.rename(columns={'passing_cpoe':'cpoe'})
            return df
        except Exception:
            return pd.DataFrame()

    def _nfl_fetch_ngs(yr, stat):
        url = f"{NFLVERSE_BASE}/nextgen_stats/ngs_{yr}_{stat}.csv.gz"
        try:
            r = SESSION.get(url, timeout=30, allow_redirects=True)
            r.raise_for_status()
            if len(r.content) < 2000: return pd.DataFrame()
            with _gzip_upd.open(__import__('io').BytesIO(r.content)) as gz:
                return pd.read_csv(gz, low_memory=False)
        except Exception:
            return pd.DataFrame()

    # ── Pull current season NGS ─────────────────────────────────────────────
    print(f"  Pulling {_season} NGS (passing/rushing/receiving)...")
    ngs_updated = 0
    players_out = existing.get('players', {})

    try:
        # Passing NGS — CPOE, intended air yards, aggressiveness
        _ngs_p = _nfl_fetch_ngs(_season, 'passing')
        _ngs_p_seas = _ngs_p.groupby('player_display_name').mean(numeric_only=True).reset_index() if not _ngs_p.empty else _ngs_p
        for _, row in _ngs_p_seas.iterrows():
            name = str(row.get('player_display_name', '')).strip()
            if not name or name == 'nan': continue
            if name not in players_out:
                players_out[name] = {'name': name, 'pos': 'QB', 'source': 'ngs_update'}
            players_out[name].update({
                'cpoe':        safe_float(row.get('completion_percentage_above_expectation', 0)),
                'iay':         safe_float(row.get('avg_intended_air_yards', 0)),
                'aggPct':      safe_float(row.get('aggressiveness', 0)),
                'timeToThrow': safe_float(row.get('avg_time_to_throw', 0)),
                'ngsSource':   True,
                'ngsWeek':     _season,
            })
            ngs_updated += 1
        print(f"    QB NGS: {len(_ngs_p_seas)} players updated")
    except Exception as _e:
        print(f"    QB NGS: ⚠ {_e}")

    try:
        # Rushing NGS — RYOE, efficiency, stacked box rate
        _ngs_r = _nfl_fetch_ngs(_season, 'rushing')
        _ngs_r_seas = _ngs_r.groupby('player_display_name').mean(numeric_only=True).reset_index() if not _ngs_r.empty else _ngs_r
        for _, row in _ngs_r_seas.iterrows():
            name = str(row.get('player_display_name', '')).strip()
            if not name or name == 'nan': continue
            if name not in players_out:
                players_out[name] = {'name': name, 'pos': 'RB', 'source': 'ngs_update'}
            players_out[name].update({
                'ryoe':       safe_float(row.get('rush_yards_over_expected_per_att', 0)),
                'rushEff':    safe_float(row.get('efficiency', 0)),
                'stackedPct': safe_float(row.get('percent_attempts_gte_eight_defenders',
                              row.get('percent_attempts_gte_eight_defenders', 0))),
                'ngsSource':  True,
                'ngsWeek':    _season,
            })
            ngs_updated += 1
        print(f"    RB NGS: {len(_ngs_r_seas)} players updated")
    except Exception as _e:
        print(f"    RB NGS: ⚠ {_e}")

    try:
        # Receiving NGS — separation, YAC above expected, air yards share
        _ngs_e = _nfl_fetch_ngs(_season, 'receiving')
        _ngs_e_seas = _ngs_e.groupby('player_display_name').mean(numeric_only=True).reset_index() if not _ngs_e.empty else _ngs_e
        for _, row in _ngs_e_seas.iterrows():
            name = str(row.get('player_display_name', '')).strip()
            if not name or name == 'nan': continue
            if name not in players_out:
                players_out[name] = {'name': name, 'pos': 'WR', 'source': 'ngs_update'}
            players_out[name].update({
                'separation':  safe_float(row.get('avg_separation', 0)),
                'cushion':     safe_float(row.get('avg_cushion', 0)),
                'yacAboveExp': safe_float(row.get('avg_yac_above_expectation', 0)),
                'airYdShare':  safe_float(row.get('percent_share_of_intended_air_yards', 0)),
                'ngsSource':   True,
                'ngsWeek':     _season,
            })
            ngs_updated += 1
        print(f"    WR/TE NGS: {len(_ngs_e_seas)} players updated")
    except Exception as _e:
        print(f"    WR/TE NGS: ⚠ {_e}")

    # ── Pull current season player stats ────────────────────────────────────
    print(f"  Pulling {_season} player stats...")
    try:
        _ps_raw = _nfl_fetch_ps(_season)
        _ps_reg = _ps_raw  # already filtered to REG and cpoe renamed
        if not _ps_reg.empty:
            if 'dakota' in _ps_reg.columns and 'cpoe' not in _ps_reg.columns:
                _ps_reg = _ps_reg.rename(columns={'dakota': 'cpoe'})
            elif 'passing_cpoe' in _ps_reg.columns:
                _ps_reg = _ps_reg.rename(columns={'passing_cpoe': 'cpoe'})
            if 'team' in _ps_reg.columns and 'recent_team' not in _ps_reg.columns:
                _ps_reg = _ps_reg.rename(columns={'team':'recent_team'})
            # Build season totals
            _SUM  = [c for c in ['passing_yards','passing_tds','interceptions','passing_epa',
                                  'rushing_yards','carries','rushing_tds','rushing_epa',
                                  'receiving_yards','receptions','targets','receiving_tds',
                                  'receiving_epa','completions','attempts'] if c in _ps_reg.columns]
            _MEAN = [c for c in ['cpoe','target_share','air_yards_share','wopr'] if c in _ps_reg.columns]
            _GRP  = [c for c in ['player_display_name','position','recent_team'] if c in _ps_reg.columns]
            _AGG  = {c:'sum' for c in _SUM}; _AGG.update({c:'mean' for c in _MEAN}); _AGG['week']='nunique'
            _seas = _ps_reg.groupby(_GRP).agg(_AGG).rename(columns={'week':'games'}).reset_index()

            updated_ps = 0
            for _, row in _seas.iterrows():
                name   = str(row.get('player_display_name','')).strip()
                pos    = str(row.get('position','')).upper()
                team   = str(row.get('recent_team','')).upper()
                games  = safe_int(row.get('games', 1))
                if not name or name == 'nan': continue
                if name not in players_out:
                    players_out[name] = {'name': name, 'pos': pos, 'team': team, 'source': 'ps_update'}
                p = players_out[name]
                p['team'] = team; p['games'] = games
                if pos == 'QB':
                    att = max(safe_int(row.get('attempts',0)), 1)
                    p.update({'passYdsPG': safe_float(row.get('passing_yards',0))/max(games,1),
                              'passTDsPG': safe_float(row.get('passing_tds',0))/max(games,1),
                              'passEPA':   safe_float(row.get('passing_epa',0)),
                              'cpoe':      safe_float(row.get('cpoe',0))})
                elif pos == 'RB':
                    p.update({'rushYdsPG':  safe_float(row.get('rushing_yards',0))/max(games,1),
                              'recYdsPG':   safe_float(row.get('receiving_yards',0))/max(games,1),
                              'targShare':  safe_float(row.get('target_share',0))})
                elif pos in ('WR','TE'):
                    p.update({'recYdsPG':   safe_float(row.get('receiving_yards',0))/max(games,1),
                              'targShare':  safe_float(row.get('target_share',0)),
                              'airYdShare': safe_float(row.get('air_yards_share',0)),
                              'wopr':       safe_float(row.get('wopr',0))})
                updated_ps += 1
            print(f"    Player stats: {updated_ps} players patched")
    except Exception as _e:
        print(f"    Player stats: ⚠ {_e}")

    # ── Pull current season team EPA ────────────────────────────────────────
    print(f"  Pulling {_season} team EPA...")
    try:
        _ts = fetch_csv(f"{NFLVERSE_BASE}/stats_team/stats_team_reg_{_season}.csv",
                        f"team_stats_reg {_season}", silent_404=True)
        if not _ts.empty:
            teams_out  = existing.get('teams', {})
            _TNORM = {'LA':'LAR','JAC':'JAX','KCC':'KC','SFO':'SF','NWE':'NE',
                      'NOR':'NO','GNB':'GB','TBB':'TB','SDG':'LAC','STL':'LAR'}
            PLAYS_PER_SEASON = 1050; PASS_PLAYS = 600; RUSH_PLAYS = 450
            updated_t = 0
            for _, row in _ts.iterrows():
                abbr  = _TNORM.get(str(row.get('team','')).upper(), str(row.get('team','')).upper())
                games = safe_int(row.get('games', 17)) or 17
                def _pp(val, plays): return round(val/plays,4) if abs(val)>10 else round(val,4)
                off_r = safe_float(row.get('passing_epa',0)+row.get('rushing_epa',0))
                def_r = safe_float(row.get('defense_epa', row.get('def_epa',0)))
                if abbr not in teams_out: teams_out[abbr] = {}
                teams_out[abbr].update({
                    'offEPA':     _pp(off_r, PLAYS_PER_SEASON),
                    'defEPA':     _pp(def_r, PLAYS_PER_SEASON),
                    'offPassEPA': _pp(safe_float(row.get('passing_epa',0)), PASS_PLAYS),
                    'offRushEPA': _pp(safe_float(row.get('rushing_epa',0)), RUSH_PLAYS),
                    'ptsPG':      round(safe_float(row.get('pts_for',0))/max(games,1),1),
                    'ptAllPG':    round(safe_float(row.get('pts_against',0))/max(games,1),1),
                    'games':      games,
                    'source':     f'stats_team_{_season}',
                })
                updated_t += 1
            existing['teams'] = teams_out
            print(f"    Team EPA: {updated_t} teams updated")
    except Exception as _e:
        print(f"    Team EPA: ⚠ {_e}")

    # ── Apply coaching overrides ────────────────────────────────────────────
    coaching = build_coaching_context({})
    if coaching:
        for abbr, ctx in coaching.items():
            if abbr in existing.get('teams', {}):
                existing['teams'][abbr]['coaching_override'] = ctx
                adj = ctx.get('adj_epa', 0)
                if adj:
                    existing['teams'][abbr]['offEPA'] = round(
                        existing['teams'][abbr].get('offEPA', 0) + adj, 4)

    # ── Save patched JSON ───────────────────────────────────────────────────
    existing['players'] = players_out
    existing['meta']['ngs_updated']     = datetime.now().isoformat()
    existing['meta']['ngs_season']      = _season
    existing['meta']['ngs_player_count']= ngs_updated
    existing['coaching_context']        = coaching

    with open(_outfile, 'w', encoding='utf-8') as f:
        _json.dump(existing, f, indent=2, default=str)

    size_kb = Path(_outfile).stat().st_size / 1024
    print(f"\n  ✅ Patched JSON saved: {_outfile}")
    print(f"  Size: {size_kb:.1f} KB  |  NGS players updated: {ngs_updated}")
    print(f"\n  📋 What was updated:")
    print(f"     CPOE / IAY / aggressiveness / time-to-throw  (QBs)")
    print(f"     RYOE / efficiency / stacked-box rate          (RBs)")
    print(f"     Separation / YAC above exp / air yards share  (WR/TE)")
    print(f"     Team EPA / pts scored / pts allowed           (all 32 teams)")
    if coaching:
        print(f"     Coaching overrides applied: {list(coaching.keys())}")
    print(f"\n  Load nfl_model_data.json in the model to activate.")



# ─────────────────────────────────────────────
# BACKTEST GAMES BUILDER
# ─────────────────────────────────────────────
def build_backtest_games(dfs, teams, season):
    """
    Build backtest records for every completed game in `season`.
    Merges actual scores from nflverse games.csv with:
      - Model projections (EPA-based scoring formula)
      - Vegas lines (spread_line, total_line from nflverse)
    Returns list of game dicts compatible with the model's BT_ENGINE format.
    """
    sched = dfs.get('schedules', pd.DataFrame())
    if sched.empty:
        print(f"  ⚠ No schedule data for backtest")
        return []

    # Filter to completed regular-season games for the target season
    completed = sched[
        (sched['season'] == season) &
        (sched['game_type'] == 'REG') &
        sched['home_score'].notna() &
        sched['away_score'].notna()
    ].copy()

    if completed.empty:
        print(f"  ⚠ No completed games found for {season}")
        return []

    TNORM = {'JAC':'JAX','LA':'LAR','WSH':'WAS','LVR':'LV'}
    DIV_MAP = {
        'AFC East':  ['BUF','MIA','NE','NYJ'],
        'AFC North': ['BAL','CIN','CLE','PIT'],
        'AFC South': ['HOU','IND','JAX','TEN'],
        'AFC West':  ['DEN','KC','LAC','LV'],
        'NFC East':  ['DAL','NYG','PHI','WAS'],
        'NFC North': ['CHI','DET','GB','MIN'],
        'NFC South': ['ATL','CAR','NO','TB'],
        'NFC West':  ['ARI','LAR','SEA','SF'],
    }
    def same_div(a, b):
        for teams_list in DIV_MAP.values():
            if a in teams_list and b in teams_list:
                return True
        return False

    # Scoring constants (match JS projectGamePure)
    MR=0.88; AWAY_DISC=0.83; HFA=1.7; TO_C=0.10; BASE=20.0

    results = []
    for _, g in completed.iterrows():
        away = TNORM.get(str(g['away_team']).upper(), str(g['away_team']).upper())
        home = TNORM.get(str(g['home_team']).upper(), str(g['home_team']).upper())
        A = teams.get(away, {}); H = teams.get(home, {})

        # EPA-based model projection
        if A and H:
            off_a = (A.get('offEPA',0)*MR*AWAY_DISC - H.get('defEPA',0)*MR)*0.28
            off_h = (H.get('offEPA',0)*MR           - A.get('defEPA',0)*MR*AWAY_DISC)*0.28
            to_a  = min(6, max(-6, A.get('toMargin', 0))) * TO_C
            to_h  = min(6, max(-6, H.get('toMargin', 0))) * TO_C
            div   = -1.0 if same_div(away, home) else 0.0
            s_a   = max(10, min(42, BASE + off_a + to_a + div))
            s_h   = max(10, min(42, BASE + off_h + to_h + HFA + div))
        else:
            s_a = s_h = BASE

        away_sc = float(g['away_score'])
        home_sc = float(g['home_score'])
        act_total = away_sc + home_sc
        mod_total = s_a + s_h

        # Winner
        winner_ok = None
        if home_sc != away_sc:
            winner_ok = (s_h > s_a) == (home_sc > away_sc)

        # Vegas lines
        spread = float(g['spread_line']) if pd.notna(g.get('spread_line')) else None
        total  = float(g['total_line'])  if pd.notna(g.get('total_line'))  else None

        # ATS — spread_line is home-team perspective (negative = home fav)
        ats_ok = None
        if spread is not None:
            actual_margin = home_sc - away_sc   # positive = home won by X
            ats_ok = actual_margin > spread       # home covers if won by > spread

        # O/U
        ou_ok = None
        if total is not None:
            ou_ok = act_total > total

        gameday = str(g.get('gameday', '')) if pd.notna(g.get('gameday','')) else ''

        results.append({
            'season':        int(season),
            'wk':            int(g['week']),
            'gameday':       gameday,
            'away':          away,
            'home':          home,
            'awayScore':     away_sc,
            'homeScore':     home_sc,
            'actMargin':     round(home_sc - away_sc, 1),
            'actTotal':      round(act_total, 1),
            'modScoreA':     round(s_a, 1),
            'modScoreH':     round(s_h, 1),
            'modMargin':     round(s_h - s_a, 1),
            'modTotal':      round(mod_total, 1),
            'winnerCorrect': winner_ok,
            'atsCorrect':    ats_ok,
            'ouCorrect':     ou_ok,
            'spread':        spread,
            'total':         total,
            'marginErr':     round(abs((s_h - s_a) - (home_sc - away_sc)), 2),
            'totalErr':      round(abs(mod_total - act_total), 2),
            'scaleBias':     round(mod_total / act_total, 3) if act_total > 0 else 1.0,
            'isDivGame':     same_div(away, home),
            'note':          '',
        })

    print(f"  Built {len(results)} backtest games for {season} "
          f"({sum(1 for g in results if g['atsCorrect'] is not None)} with Vegas lines)")
    return results


def main():
    import argparse
    parser = argparse.ArgumentParser(description="NFL Edge Model data pull")
    parser.add_argument('--update', action='store_true',
        help='In-season weekly NGS patch (fast — skips roster/depth/rookie steps)')
    args, _ = parser.parse_known_args()

    if args.update:
        update_inseason_ngs()
        return

    print("="*50)
    print("NFL EDGE MODEL — HYBRID DATA PULL")
    print(f"Season: {SEASON}")
    print(f"Output: {OUTPUT_FILE}")
    print("="*50)

    # Pull all nflverse data
    dfs = pull_nflverse(SEASON)

    # Build team EPA profiles
    teams = build_team_profiles(dfs, SEASON)

    # Build player profiles
    players = build_player_profiles(dfs, SEASON)

    # Detect rookies
    rookies = detect_rookies(dfs, players, SEASON)

    # Load college fallback data
    college_data = load_prospect_analyzer()

    # Apply rookie blending (college → NFL transition)
    apply_rookie_blending(players, college_data, rookies)

    # Extract starters from depth charts (must come before injury / roster calls)
    starters_map = extract_starters(dfs)

    # Merge starters into teams NOW so WR1/WR2/WR3/TE are available
    # when build_healthy_roster_shares runs (previously merged only in save_output)
    if starters_map:
        for _abbr, _sd in starters_map.items():
            if _abbr in teams:
                teams[_abbr].update(_sd)
            else:
                teams[_abbr] = _sd

    # Fetch current-season roster for roster-change comparison
    dfs['roster_curr'] = fetch_roster_current(CURRENT_SEASON)

    # Build injury/depth chart status
    injury_map = build_injury_status(dfs, SEASON)

    # Build upcoming games list
    games = build_upcoming_games(dfs, SEASON)

    # Detect roster changes and compute scoring impact
    # Build healthy roster shares (co-game target distributions)
    healthy_roster_shares = build_healthy_roster_shares(dfs, teams)

    roster_changes = build_roster_changes(dfs, players, SEASON, starters_2026=starters_map)

    # Build coaching context (mid-season overrides + scheme flags)
    coaching_context = build_coaching_context(dfs)

    # Save everything to JSON
    save_output(teams, players, injury_map, games, college_data, SEASON,
                roster_changes, starters_map, coaching_context, dfs=dfs,
                healthy_roster_shares=healthy_roster_shares)

    print("\n" + "="*50)
    print("✅ COMPLETE — Load nfl_model_data.json into NFL_Edge_Model.html")
    print(f"\n📊 DATA SUMMARY:")
    print(f"   Teams with EPA:  {len(teams)} / 32")
    print(f"   Players loaded:  {len(players)}")
    print(f"   Injuries:        {len(injury_map)}")
    print(f"   Upcoming games:  {len(games)}")
    n_adj = len([t for t, d in (roster_changes or {}).items() if abs(d.get('offAdj', 0)) >= 0.5])
    print(f"   Roster adjustments: {n_adj} teams affected")
    print(f"   Starters extracted: {len([t for t in (starters_map or {}).values() if t])} teams")
    if len(players) == 0:
        print(f"\n⚠  No player stats loaded — normal pre-season behavior.")
        print(f"   Props will use static baselines until {CURRENT_SEASON} season starts.")
    if len(games) == 0:
        print(f"\n⚠  No upcoming games — use ESPN fetch in the model Date Range bar.")
    print("="*50)
    if CURRENT_SEASON > BASELINE_SEASON:
        print(f"\n📋 PRE-SEASON MODE:")
        print(f"   Baseline year: {BASELINE_SEASON} (prior season stats loaded as team baseline)")
        print(f"   Current year:  {CURRENT_SEASON} (season being projected)")
        print(f"   Week 1 note:   {CURRENT_SEASON} nflverse files are empty until games are played.")
        print(f"   After Week 1:  Set BASELINE_SEASON = {CURRENT_SEASON} and re-run for live data.")
    else:
        print(f"\n📊 LIVE MODE: Pulling {BASELINE_SEASON} current-season data")

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='NFL Edge Model data pull')
    parser.add_argument('--update',  action='store_true',
                        help='In-season update: patch existing JSON with latest NGS + recency EPA')
    parser.add_argument('--weeks',   type=int, default=4,
                        help='Recency window in weeks for --update (default: 4)')
    parser.add_argument('--season',  type=int, default=None,
                        help='Override season for --update (default: CURRENT_SEASON)')
    args = parser.parse_args()

    if args.update:
        update_inseason_ngs(
            target_season=args.season or CURRENT_SEASON,
            weeks_back=args.weeks,
        )
    else:
        main()

# ══════════════════════════════════════════════════════════════════════
# IN-SEASON UPDATE ENGINE
# ══════════════════════════════════════════════════════════════════════
# Run weekly (Tuesday after games) to refresh the model with current-season
# NGS, player stats, and recency-weighted EPA that automatically captures
# coaching changes and system shifts without manual date entry.
#
# Usage:
#   python nfl_edge_data_pull.py --update              # patch existing JSON
#   python nfl_edge_data_pull.py --update --weeks 4   # last N weeks for recency
#   python nfl_edge_data_pull.py --update --week 12   # through a specific week
# ══════════════════════════════════════════════════════════════════════

def compute_recency_weighted_epa(weekly_df, weeks_back=4):
    """
    Compute per-team EPA weighted toward recent weeks.

    Weight scheme: most recent week = weeks_back, oldest included = 1.
    E.g. for weeks_back=4 and max completed week=12:
      Week 12 weight=4, Week 11=3, Week 10=2, Week 9=1 (weeks ≤8 excluded).

    Returns dict: {team_abbr: {'offEPA_L4': float, 'offPassEPA_L4': float,
                               'offRushEPA_L4': float, 'games_L4': int}}
    Automatically captures coaching/scheme changes without needing change dates.
    A team that fired their OC in Week 9 will show a different L4 vs season avg.
    """
    if weekly_df.empty:
        return {}

    # Determine columns available
    team_col = next((c for c in ['recent_team','team'] if c in weekly_df.columns), None)
    week_col  = 'week' if 'week' in weekly_df.columns else None
    if not team_col or not week_col:
        return {}

    reg = weekly_df.copy()
    if 'season_type' in reg.columns:
        reg = reg[reg['season_type'] == 'REG']

    max_week = int(reg[week_col].max()) if not reg.empty else 0
    if max_week == 0:
        return {}

    cutoff = max(1, max_week - weeks_back + 1)
    recent = reg[reg[week_col] >= cutoff].copy()

    # Recency weight: week=max_week → weight=weeks_back, week=cutoff → weight=1
    recent['_wt'] = recent[week_col] - cutoff + 1   # 1 … weeks_back

    result = {}
    _TNORM = {'LA':'LAR','JAC':'JAX','KCC':'KC','SFO':'SF','NWE':'NE',
              'NOR':'NO','GNB':'GB','TBB':'TB','SDG':'LAC','STL':'LAR'}

    for team, grp in recent.groupby(team_col):
        abbr = _TNORM.get(str(team).upper(), str(team).upper())
        w    = grp['_wt']

        # Pass EPA per dropback
        pass_epa_col  = next((c for c in ['passing_epa'] if c in grp.columns), None)
        rush_epa_col  = next((c for c in ['rushing_epa'] if c in grp.columns), None)
        attempts_col  = next((c for c in ['attempts','pass_attempts'] if c in grp.columns), None)
        carries_col   = next((c for c in ['carries','rush_attempts'] if c in grp.columns), None)

        def weighted_per_play(epa_col, plays_col):
            if not epa_col or not plays_col: return 0.0
            epa   = pd.to_numeric(grp[epa_col],  errors='coerce').fillna(0)
            plays = pd.to_numeric(grp[plays_col], errors='coerce').fillna(0)
            wt_epa   = (epa   * plays * w).sum()
            wt_plays = (plays * w).sum()
            return round(float(wt_epa / wt_plays), 4) if wt_plays > 0 else 0.0

        pass_epa_l4 = weighted_per_play(pass_epa_col, attempts_col)
        rush_epa_l4 = weighted_per_play(rush_epa_col, carries_col)
        # Combined off EPA (weighted average of pass and rush contributions)
        wt_pass = (pd.to_numeric(grp.get(attempts_col, pd.Series([0])*len(grp)), errors='coerce').fillna(0) * w).sum()
        wt_rush = (pd.to_numeric(grp.get(carries_col,  pd.Series([0])*len(grp)), errors='coerce').fillna(0) * w).sum()
        total_plays = wt_pass + wt_rush
        off_epa_l4 = round(
            (pass_epa_l4 * wt_pass + rush_epa_l4 * wt_rush) / total_plays, 4
        ) if total_plays > 0 else 0.0

        result[abbr] = {
            'offEPA_L4':     off_epa_l4,
            'offPassEPA_L4': pass_epa_l4,
            'offRushEPA_L4': rush_epa_l4,
            'games_L4':      int(grp[week_col].nunique()),
            'weeks_included': sorted(grp[week_col].unique().tolist()),
        }

    return result


def compute_scheme_trend(full_season_df, recent_df, weeks_back=4):
    """
    Detect teams where recent execution diverges from season baseline.
    Returns dict: {abbr: {trend_flag, offEPA_trend, passRate_trend, trend_reason}}

    Large positive trend  → improving / new system clicking
    Large negative trend  → regressing / coaching change hurting performance
    Threshold ±0.04 EPA/play (~1.5 pts/game) flags significant change.
    """
    if full_season_df.empty or recent_df.empty:
        return {}

    TREND_THRESHOLD = 0.04   # EPA/play delta that signals meaningful shift
    PASS_RATE_THRESHOLD = 0.06  # pass rate delta (6% shift = scheme change signal)

    team_col = next((c for c in ['recent_team','team'] if c in full_season_df.columns), None)
    if not team_col: return {}
    _TNORM = {'LA':'LAR','JAC':'JAX','KCC':'KC','SFO':'SF','NWE':'NE',
              'NOR':'NO','GNB':'GB','TBB':'TB','SDG':'LAC','STL':'LAR'}

    def team_summary(df):
        """Per-team season summary from weekly data."""
        out = {}
        for team, grp in df.groupby(team_col):
            abbr = _TNORM.get(str(team).upper(), str(team).upper())
            att  = pd.to_numeric(grp.get('attempts', pd.Series([0]*len(grp))), errors='coerce').fillna(0).sum()
            car  = pd.to_numeric(grp.get('carries',  pd.Series([0]*len(grp))), errors='coerce').fillna(0).sum()
            plays = att + car
            if plays == 0: continue
            pepa = pd.to_numeric(grp.get('passing_epa', pd.Series([0]*len(grp))), errors='coerce').fillna(0).sum()
            repa = pd.to_numeric(grp.get('rushing_epa', pd.Series([0]*len(grp))), errors='coerce').fillna(0).sum()
            out[abbr] = {
                'offEPA':    round((pepa + repa) / plays, 4),
                'passRate':  round(float(att / plays), 3) if plays > 0 else 0.5,
            }
        return out

    full = team_summary(full_season_df)
    recn = team_summary(recent_df)

    trends = {}
    for abbr in set(full) | set(recn):
        f = full.get(abbr, {})
        r = recn.get(abbr, {})
        if not f or not r: continue

        epa_delta       = round(r.get('offEPA',0) - f.get('offEPA',0), 4)
        pass_rate_delta = round(r.get('passRate',0) - f.get('passRate',0), 3)

        reasons = []
        if abs(epa_delta) >= TREND_THRESHOLD:
            direction = 'Improving' if epa_delta > 0 else 'Declining'
            reasons.append(f"{direction} efficiency ({epa_delta:+.3f} EPA/play vs season avg)")
        if abs(pass_rate_delta) >= PASS_RATE_THRESHOLD:
            direction = 'More pass-heavy' if pass_rate_delta > 0 else 'More run-heavy'
            reasons.append(f"{direction} ({pass_rate_delta:+.1%} pass rate shift)")

        trends[abbr] = {
            'offEPA_trend':    epa_delta,
            'passRate_trend':  pass_rate_delta,
            'trend_flag':      len(reasons) > 0,
            'trend_reason':    '; '.join(reasons) if reasons else 'Stable',
        }

    return trends


def update_inseason_ngs(target_season=None, weeks_back=4, output_file=None):
    """
    Weekly in-season update: pull current season NGS + player stats,
    compute recency-weighted EPA and scheme trends, patch existing JSON.

    Automatically detects coaching/system changes via EPA and pass-rate
    divergence — no manual change dates needed.

    Args:
        target_season: Season to pull (default: CURRENT_SEASON)
        weeks_back:    Weeks to use for recency window (default: 4)
        output_file:   JSON path to patch (default: OUTPUT_FILE)
    """
    import json
    from datetime import datetime, timezone

    season = target_season or CURRENT_SEASON
    out_path = Path(output_file) if output_file else OUTPUT_FILE

    print(f"\n{'='*50}")
    print(f"IN-SEASON NGS UPDATE — {season} Season (L{weeks_back} Recency Window)")
    print('='*50)

    # ── Load existing JSON to patch ──────────────────────────────────
    existing = {}
    if out_path.exists():
        with open(out_path, encoding='utf-8') as f:
            existing = json.load(f)
        print(f"  Loaded existing JSON: {out_path.name} "
              f"({len(existing.get('teams',{}))} teams, "
              f"{len(existing.get('players',{}))} players)")
    else:
        print(f"  ⚠ No existing JSON found — run full pull first")
        return

    # ── Pull current season player stats (all weeks) ─────────────────
    print(f"\n  Pulling {season} player stats...", end='', flush=True)
    try:
        ps_raw = _nfl.load_player_stats([season]).to_pandas()
        ps_reg = ps_raw[ps_raw['season_type'] == 'REG'].copy() if 'season_type' in ps_raw.columns else ps_raw
        ps_reg = ps_reg.rename(columns={'team': 'recent_team', 'passing_cpoe': 'cpoe'})
        max_wk = int(ps_reg['week'].max()) if 'week' in ps_reg.columns and not ps_reg.empty else 0
        print(f" ✅ {len(ps_reg):,} rows through Week {max_wk}")
    except Exception as e:
        print(f" ❌ {e}")
        ps_reg = pd.DataFrame()

    # ── Recency-weighted EPA ─────────────────────────────────────────
    if not ps_reg.empty and max_wk >= weeks_back:
        print(f"  Computing L{weeks_back} recency-weighted EPA (Weeks {max(1,max_wk-weeks_back+1)}–{max_wk})...")
        l4_epa = compute_recency_weighted_epa(ps_reg, weeks_back=weeks_back)

        # Recent window for trend comparison
        cutoff = max(1, max_wk - weeks_back + 1)
        recent_only = ps_reg[ps_reg['week'] >= cutoff] if 'week' in ps_reg.columns else ps_reg
        trends     = compute_scheme_trend(ps_reg, recent_only, weeks_back=weeks_back)

        # Merge into existing teams
        updated = 0
        flagged = []
        for abbr, data in l4_epa.items():
            if abbr in existing.get('teams', {}):
                existing['teams'][abbr].update(data)
                if abbr in trends:
                    existing['teams'][abbr].update(trends[abbr])
                    if trends[abbr].get('trend_flag'):
                        flagged.append((abbr, trends[abbr]['trend_reason']))
                updated += 1
        print(f"  L{weeks_back} EPA updated: {updated} teams")

        if flagged:
            print(f"\n  ⚡ SCHEME/SYSTEM CHANGE FLAGS ({len(flagged)} teams):")
            for abbr, reason in sorted(flagged):
                print(f"    {abbr:<4} {reason}")
        else:
            print(f"  No significant scheme shifts detected this window")
    else:
        print(f"  ⚠ Fewer than {weeks_back} weeks completed — skipping recency window")
        l4_epa = {}

    # ── Pull current season NGS ──────────────────────────────────────
    print(f"\n  Pulling {season} NGS...", end='', flush=True)
    ngs_updated = 0
    try:
        ngs_pass = _nfl.load_nextgen_stats([season], stat_type='passing').to_pandas()
        ngs_rush = _nfl.load_nextgen_stats([season], stat_type='rushing').to_pandas()
        ngs_rec  = _nfl.load_nextgen_stats([season], stat_type='receiving').to_pandas()

        # Use season aggregates (week=0)
        ngs_p_agg = ngs_pass[ngs_pass['week'] == 0] if 'week' in ngs_pass.columns else ngs_pass
        ngs_r_agg = ngs_rush[ngs_rush['week'] == 0] if 'week' in ngs_rush.columns else ngs_rush
        ngs_e_agg = ngs_rec[ngs_rec['week']  == 0] if 'week' in ngs_rec.columns  else ngs_rec

        def norm(n): return re.sub(r"[^a-z ]", "", str(n).lower().strip())
        players_out = existing.get('players', {})

        # Merge NGS passing
        for _, row in ngs_p_agg.iterrows():
            name = norm(row.get('player_display_name', ''))
            if name in players_out:
                players_out[name].update({
                    'cpoe':       safe_float(row.get('completion_percentage_above_expectation', 0)),
                    'iay':        safe_float(row.get('avg_intended_air_yards', 0)),
                    'aggPct':     safe_float(row.get('aggressiveness', 0)),
                    'timeToThrow':safe_float(row.get('avg_time_to_throw', 0)),
                    'ngsSource':  True,
                    'ngsSeason':  season,
                })
                ngs_updated += 1

        # Merge NGS rushing (RYOE)
        for _, row in ngs_r_agg.iterrows():
            name = norm(row.get('player_display_name', ''))
            if name in players_out:
                players_out[name].update({
                    'ryoe':       safe_float(row.get('rush_yards_over_expected_per_att', 0)),
                    'rushEff':    safe_float(row.get('efficiency', 0)),
                    'stackedPct': safe_float(row.get('percent_attempts_gte_eight_defenders', 0)),
                    'ngsSource':  True,
                    'ngsSeason':  season,
                })
                ngs_updated += 1

        # Merge NGS receiving (separation, YAC above expected)
        for _, row in ngs_e_agg.iterrows():
            name = norm(row.get('player_display_name', ''))
            if name in players_out:
                players_out[name].update({
                    'separation':  safe_float(row.get('avg_separation', 0)),
                    'cushion':     safe_float(row.get('avg_cushion', 0)),
                    'yacAboveExp': safe_float(row.get('avg_yac_above_expectation', 0)),
                    'airYdShare':  safe_float(row.get('percent_share_of_intended_air_yards', 0)),
                    'ngsSource':   True,
                    'ngsSeason':   season,
                })
                ngs_updated += 1

        existing['players'] = players_out
        print(f" ✅ {ngs_updated} player NGS fields updated (pass/rush/rec)")

    except Exception as e:
        print(f" ❌ {e}")

    # ── Pull current season team stats (for updated team EPA) ─────────
    print(f"  Pulling {season} team stats...", end='', flush=True)
    try:
        ts_url = f"{NFLVERSE_BASE}/stats_team/stats_team_reg_{season}.csv"
        ts_df  = fetch_csv(ts_url, f"team_stats_reg {season}", silent_404=True)
        if not ts_df.empty:
            _TNORM = {'LA':'LAR','JAC':'JAX','KCC':'KC','SFO':'SF','NWE':'NE',
                      'NOR':'NO','GNB':'GB','TBB':'TB','SDG':'LAC','STL':'LAR'}
            PLAYS = 1050; PASS_P = 600; RUSH_P = 450
            for _, row in ts_df.iterrows():
                abbr = str(row.get('team', row.get('team_abbr', ''))).upper().strip()
                abbr = _TNORM.get(abbr, abbr)
                if not abbr or abbr == 'NAN' or abbr not in existing.get('teams', {}): continue
                def tpp(val, plays):
                    v = safe_float(val)
                    return round(v / plays, 4) if abs(v) > 10 else round(v, 4)
                existing['teams'][abbr].update({
                    'offEPA':     tpp(row.get('offense_epa', row.get('passing_epa',0))
                                      + row.get('rushing_epa',0), PLAYS),
                    'offPassEPA': tpp(row.get('passing_epa', 0), PASS_P),
                    'offRushEPA': tpp(row.get('rushing_epa', 0), RUSH_P),
                    'defEPA':     tpp(row.get('defense_epa', 0), PLAYS),
                    'games':      safe_int(row.get('games', 0)),
                    'ptsPG':      round(safe_float(row.get('pts_for',0)) / max(safe_int(row.get('games',1)),1), 1),
                    'ptAllPG':    round(safe_float(row.get('pts_against',0)) / max(safe_int(row.get('games',1)),1), 1),
                    'epaSource':  f'nflverse_{season}',
                })
            print(f" ✅ {len(ts_df)} teams updated")
    except Exception as e:
        print(f" ❌ {e}")

    # ── Update metadata and save ─────────────────────────────────────
    existing.setdefault('meta', {}).update({
        'inseason_updated':     datetime.now(timezone.utc).isoformat(),
        'inseason_season':      season,
        'inseason_weeks_back':  weeks_back,
        'inseason_max_week':    max_wk if not ps_reg.empty else 0,
        'ngs_player_count':     ngs_updated,
        'trend_flags':          len(flagged) if 'flagged' in dir() else 0,
    })

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(existing, f, indent=2, default=str)

    size_kb = out_path.stat().st_size / 1024
    print(f"\n{'='*50}")
    print(f"✅ IN-SEASON UPDATE COMPLETE — {out_path.name} ({size_kb:.1f} KB)")
    if 'flagged' in dir() and flagged:
        print(f"   {len(flagged)} teams flagged for scheme shifts — review before publishing picks")
    print('='*50)

