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

import os, sys, json, re, time, warnings, math
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

import os as _os

# v4.0 REAL BUG FIXED — OUTPUT_DIR was an unconditional hardcoded Windows
# path. Confirmed via a real GitHub Actions failure: the pipeline ran
# completely successfully end-to-end (real opportunity shares, snap counts,
# coverage/red-zone proxies, 1000 player profiles built) and only failed on
# the very last line — writing the output JSON — because GitHub Actions
# runs on a Linux server, where a `C:\Users\...` path doesn't exist at all.
# This had nothing to do with anything else changed this round; it's a
# pre-existing gap that simply hadn't been reached by a successful run in
# this environment before. Fixed by detecting GitHub Actions' own
# `GITHUB_ACTIONS` environment variable (automatically set to "true" in
# every workflow run — no configuration needed) and writing to the current
# working directory in that case, which is the repo root right after
# actions/checkout — exactly where the workflow's own commit step expects
# to find it. Local Windows runs are completely unaffected.
if _os.environ.get('GITHUB_ACTIONS') == 'true':
    OUTPUT_DIR = Path('.')
else:
    OUTPUT_DIR = Path(r'C:\Users\gator\OneDrive\Desktop\NFL Models\NFL Edge Model')
OUTPUT_FILE  = OUTPUT_DIR / "nfl_model_data.json"

# v4.0 NEW — optional automatic GitHub push after a successful pull. This
# needs the LOCAL PATH to your actual git clone of Gatorz1989/NFL-Daily-Data-Pull
# — NOT the same folder as OUTPUT_DIR above, which is the model's own working
# directory, not a git repo. Set this to wherever you've cloned that repo
# locally (e.g. r'C:\Users\gator\OneDrive\Desktop\NFL Models\NFL-Daily-Data-Pull').
# Leave as None to disable — the script will just skip the push step with a
# clear message, exactly like it does today, rather than fail.
GITHUB_REPO_DIR = None  # ← SET THIS to your local repo clone path to enable --push

def push_output_to_github(repo_dir, output_file, commit_message=None):
    """v4.0 NEW — copies output_file into repo_dir (if not already there) and
    runs a real git add/commit/push. Requires repo_dir to already be a valid
    local git clone with push access already configured (SSH key or a stored
    HTTPS credential) — this function does not handle interactive auth; if
    git prompts for a password, this will hang or fail, and prints a clear
    message either way rather than crashing the whole pull. Non-fatal on any
    failure: the local JSON write in save_output() already succeeded before
    this ever runs, so a git problem here never loses that output."""
    import subprocess
    repo_dir = Path(repo_dir)
    if not repo_dir.exists():
        print(f"  ⚠ GitHub push skipped — repo directory not found: {repo_dir}")
        return False
    if not (repo_dir / '.git').exists():
        print(f"  ⚠ GitHub push skipped — {repo_dir} is not a git repository "
              f"(no .git folder). Clone the repo there first.")
        return False

    dest = repo_dir / output_file.name
    try:
        if dest.resolve() != Path(output_file).resolve():
            import shutil
            shutil.copy2(output_file, dest)
            print(f"  Copied {output_file.name} into {repo_dir}")
    except Exception as e:
        print(f"  ⚠ GitHub push skipped — couldn't copy JSON into repo dir: {e}")
        return False

    msg = commit_message or f"Automated data pull — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    try:
        subprocess.run(['git', 'add', dest.name], cwd=repo_dir, check=True,
                        capture_output=True, text=True, timeout=30)
        _commit = subprocess.run(['git', 'commit', '-m', msg], cwd=repo_dir,
                        capture_output=True, text=True, timeout=30)
        if _commit.returncode != 0 and 'nothing to commit' not in (_commit.stdout + _commit.stderr).lower():
            print(f"  ⚠ git commit reported an issue: {_commit.stderr.strip() or _commit.stdout.strip()}")
            return False
        if 'nothing to commit' in (_commit.stdout + _commit.stderr).lower():
            print(f"  ℹ No changes to push — JSON is identical to what's already in the repo.")
            return True
        _push = subprocess.run(['git', 'push'], cwd=repo_dir, check=True,
                        capture_output=True, text=True, timeout=60)
        print(f"  ✅ Pushed to GitHub: {msg}")
        return True
    except subprocess.TimeoutExpired:
        print(f"  ⚠ GitHub push timed out — likely waiting on a credential prompt. "
              f"Configure a stored credential (SSH key or git credential helper) so this can run unattended.")
        return False
    except subprocess.CalledProcessError as e:
        print(f"  ⚠ GitHub push failed: {(e.stderr or e.stdout or str(e)).strip()}")
        return False
    except Exception as e:
        print(f"  ⚠ GitHub push failed with an unexpected error: {e}")
        return False

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

# v3.31 NEW — CFBD (College Football Data) integration, per G-Money's
# explicit request: real per-player college production + games-played,
# used only for rookie blending until a player has enough real NFL games
# to no longer need a college baseline. Same env-var pattern already used
# in the companion cfb_edge_data_pull.py script, for consistency.
#   set CFBD_API_KEY=your_key_here        (Windows)
#   export CFBD_API_KEY=your_key_here     (Mac/Linux)
# Get a free key at https://collegefootballdata.com/key
#   NOTE: env var (above) always wins if set — this only matters when it's NOT set,
#   so GitHub Actions secrets keep working exactly as before. For local runs, set
#   your real key on the line below ONCE and stop re-entering it every session.
CFBD_API_KEY_LOCAL_FALLBACK = ""  # <-- paste your real CFBD key here, once
CFBD_API_KEY = os.environ.get("CFBD_API_KEY", "") or CFBD_API_KEY_LOCAL_FALLBACK
CFBD_BASE = "https://api.collegefootballdata.com"
_cfbd_auth_failed = False
_cfbd_last_call_time = 0.0
_cfbd_quota_exhausted = False

def cfbd_get(path, params=None, max_retries=3, base_delay=1.0, min_interval=0.6):
    """Same pattern as the working helper in cfb_edge_data_pull.py.
    v3.64 FIX — this had no rate limiting or 429 handling at all: with
    ~200 rookies each needing 1-2 calls, requests fired back-to-back and
    burned through CFBD's rate limit almost immediately, so most rookies
    fell back to the CFB Reference scrape unnecessarily even with a valid
    key. Added a minimum courtesy interval between calls plus retry with
    exponential backoff (1s, 2s, 4s) on 429, before giving up.
    v3.65 FIX — that backoff helps a brief burst, but a real run showed
    EVERY single rookie hitting 429 even after full retries — a sustained
    quota exhaustion (the free tier's hourly/daily cap already used up
    from earlier runs this session), not a quick blip. Retrying the same
    7-second dance for all ~200 rookies in that case wastes 20+ minutes
    for zero benefit. Now: the first time a call exhausts every retry, that's
    treated as confirmed quota exhaustion — CFBD is skipped for the rest
    of this run (straight to None, no delay, no attempt), same as the
    existing _cfbd_auth_failed short-circuit for a bad key."""
    global _cfbd_auth_failed, _cfbd_last_call_time, _cfbd_quota_exhausted
    if not CFBD_API_KEY or _cfbd_auth_failed or _cfbd_quota_exhausted:
        return None
    headers = {"Authorization": f"Bearer {CFBD_API_KEY}"}
    for attempt in range(max_retries + 1):
        elapsed = time.time() - _cfbd_last_call_time
        if elapsed < min_interval:
            time.sleep(min_interval - elapsed)
        try:
            r = requests.get(f"{CFBD_BASE}{path}", headers=headers, params=params or {}, timeout=20)
            _cfbd_last_call_time = time.time()
            if r.status_code == 429:
                if attempt < max_retries:
                    wait = base_delay * (2 ** attempt)
                    print(f"  ! CFBD rate limited on {path} — waiting {wait:.0f}s and retrying "
                          f"({attempt + 1}/{max_retries})...")
                    time.sleep(wait)
                    continue
                _cfbd_quota_exhausted = True
                print(f"  ! CFBD still rate limited on {path} after {max_retries} retries — "
                      f"treating this as quota exhaustion for the hour, not a brief blip. "
                      f"Skipping CFBD for the rest of this run (straight to CFB Reference scrape).")
                return None
            r.raise_for_status()
            return r.json()
        except requests.exceptions.HTTPError as e:
            _cfbd_last_call_time = time.time()
            if e.response is not None and e.response.status_code == 401:
                _cfbd_auth_failed = True
                print("  ! CFBD_API_KEY was rejected (401) — skipping further CFBD calls this run.")
            else:
                print(f"  ! CFBD error on {path}: {e}")
            return None
        except Exception as e:
            _cfbd_last_call_time = time.time()
            print(f"  ! CFBD error on {path}: {e}")
        return None

# Rookie threshold: games before full NFL weighting
QB_THRESHOLD     = 5
SKILL_THRESHOLD  = 3

# Blend schedule: (college_weight, nfl_weight) by NFL games played
def get_season_blend_weight(current_season_games, k=6):
    """v3.39 NEW — weight for CURRENT_SEASON data in the season-transition
    blend, scaling from 0 (no current-season games yet) to 1 (full weight
    on current season) as games accumulate. k = games needed to reach full
    weight; k=6 (~1/3 of a season) is the starting default — tune by
    changing SEASON_BLEND_K below, no code change needed."""
    if current_season_games <= 0:
        return 0.0
    return min(1.0, current_season_games / float(k))


SEASON_BLEND_K = 6  # games until a player's stat line is 100% CURRENT_SEASON

# v4.4 NEW — games until a player's ROLE-usage fields (target/carry/snap/air-yard share,
# WOPR, and the efficiency fields gated on top of it) are 100% CURRENT_SEASON. Set from a
# real held-out backtest, not a guess: predict each player's FINAL 2025 target/carry share
# from their 2024 share blended with their first N 2025 games (n = 34-124 players per
# checkpoint). At N=3 games: target-share MAE 0.0298 (k=4) vs 0.0316 (k=12) vs 0.0377
# (prior-year only); carry-share MAE 0.0879 (k=4) vs 0.1086 (k=12) vs 0.1304 (prior only).
# k=4 was within ~3% of the best k at every checkpoint (2/3/4/6 games) for BOTH metrics.
# Role usage carries over inside a season far better than it carries over between seasons,
# so it deserves faster trust in new data than the per-game yardage blend above (whose
# k comes from the joint optimizer and is left alone). Tunable here, no other code change.
USAGE_ROLE_K = 4

# v3.55 NEW — multi-year recency weighting. Named, tunable constant same
# as SEASON_BLEND_K above: each prior year back gets this fraction of the
# weight the year after it got. 0.5 means BASELINE_SEASON-1 gets half of
# whatever weight BASELINE_SEASON gets, once current-season weight is
# subtracted out.
RECENCY_DECAY_RATE = 0.5


def fetch_prior2_season_players(season):
    """v3.55 NEW — fetches a SECOND prior season's full-season per-game
    stats (BASELINE_SEASON-1), the third tier of the multi-year
    recency-weighted blend. Mirrors fetch_current_season_partial_stats()'s
    simple, self-contained style below — a completed season, so this is a
    full-season total divided by games played, no partial-season
    complexity. Returns {} on any fetch failure (older seasons
    occasionally get reorganized in nflverse's release history, and a
    player who wasn't in the league that year legitimately won't be
    present) rather than raising — this tier is a genuine enhancement on
    top of the existing 2-season blend, never something projections
    should hard-fail without.
    v3.67 FIX — confirmed directly against nflverse's real current
    release assets: the old player_stats_{season}.csv this requested
    doesn't exist anymore — nflverse renamed it to
    stats_player_reg_{season}.csv before the 2025 season. That's why a
    fully completed, long-available season (2024) was silently 404ing
    here even though it's genuinely on nflverse — wrong filename, not
    missing data. The team-stats fetch elsewhere in this file already
    uses the equivalent new stats_team_reg_{season} pattern, which is
    how this got caught.
    v3.68 FIX — that rename alone didn't fix it: a real run still showed
    zero 2024 data after the rename, with silent_404=True hiding the
    actual reason why. This now fetches directly (bypassing fetch_csv's
    silent error-swallowing) and prints the real HTTP status code or
    exception on failure — no more guessing blind at what's actually
    going wrong.
    v3.69 FIX — the v3.68 diagnostic paid off: a real run showed the
    actual error is ConnectionResetError (WinError 10054) — GitHub's CDN
    dropping the connection mid-request. That confirms the v3.67 filename
    fix was correct all along; this was a transient network drop, not a
    wrong URL or a logic bug. Added retry-with-backoff (3 attempts, 2s/4s/6s)
    since that's exactly the kind of failure a retry resolves — a single
    dropped connection shouldn't cost this entire tier of the blend for
    the whole run."""
    url = f"{NFLVERSE_BASE}/player_stats/stats_player_reg_{season}.csv"
    df = pd.DataFrame()
    max_retries = 3
    for attempt in range(max_retries + 1):
        try:
            r = SESSION.get(url, timeout=30, allow_redirects=True)
            print(f"  Prior2-season fetch ({season}): HTTP {r.status_code}, "
                  f"{len(r.content):,} bytes, content-type={r.headers.get('content-type')}")
            r.raise_for_status()
            from io import StringIO
            df = pd.read_csv(StringIO(r.text), low_memory=False)
            break
        except Exception as e:
            if attempt < max_retries:
                wait = 2.0 * (attempt + 1)
                print(f"  Prior2-season fetch ({season}) attempt {attempt + 1}/{max_retries + 1} "
                      f"failed ({type(e).__name__}: {e}) — retrying in {wait:.0f}s...")
                time.sleep(wait)
                continue
            print(f"  Prior2-season fetch ({season}) FAILED after {max_retries + 1} attempts: "
                  f"{type(e).__name__}: {e}")
            df = pd.DataFrame()
    if df is None or df.empty:
        print(f"  Prior2-season blend: no {season} data found — this tier of the blend will be skipped (players fall back to the existing 2-season blend, unchanged from today's behavior).")
        return {}
    if 'season_type' in df.columns:
        df = df[df['season_type'] == 'REG']
    if df.empty:
        return {}
    name_col = next((c for c in ['player_display_name', 'player_name',
                                  'full_name', 'display_name'] if c in df.columns), None)
    if not name_col:
        return {}
    stat_cols = {
        'passing_yards':   'passYdsPG',
        'rushing_yards':   'rushYdsPG',
        'receiving_yards': 'recYdsPG',
        'receptions':      'recsPG',
    }
    present = {raw: pg for raw, pg in stat_cols.items() if raw in df.columns}
    if not present:
        return {}
    agg = {raw: 'sum' for raw in present}
    # v3.71 FIX — the full optimizer grid showed MAE exploding from 2.165
    # to 75+ as decay_rate gave prior2 more weight, which meant prior2's
    # per-game values were badly wrong, not just unhelpful. Root cause:
    # stats_player_reg_{season}.csv is very likely a season-summary file
    # (one row per player, real season totals) with no 'week' column —
    # the old code's fallback then counted ROWS as games (always 1 for a
    # season-summary file), turning a full-season total into a per-game
    # value ~17x too large. Now prefers an explicit 'games' column first
    # — the same pattern stats_team_reg already uses successfully
    # elsewhere in this file — before falling back to week-nunique or,
    # as a last resort, row count (with a visible warning, since that
    # path is exactly what produced the bad data this run).
    if 'games' in df.columns:
        agg['games'] = 'max'
        grouped = df.groupby(name_col).agg(agg)
        games = grouped['games']
    elif 'week' in df.columns:
        agg['week'] = 'nunique'
        grouped = df.groupby(name_col).agg(agg)
        games = grouped['week']
    else:
        grouped = df.groupby(name_col).agg(agg)
        games = df.groupby(name_col).size()
        print(f"  Prior2-season blend ({season}): no 'games' or 'week' column found — "
              f"falling back to row count, which is unreliable for a season-summary file. "
              f"Sanity-checking resulting values before trusting them.")
    out = {}
    dropped_implausible = 0
    # A real per-game value for any of these stats should never realistically
    # exceed this — catches the exact failure mode above (a season total
    # divided by 1 "game") before it can poison the blend again.
    PLAUSIBLE_MAX_PG = {'passYdsPG': 500, 'rushYdsPG': 250, 'recYdsPG': 250, 'recsPG': 20}
    for raw_name, row in grouped.iterrows():
        g = int(games.get(raw_name, 0))
        if g <= 0:
            continue
        key = norm_name(raw_name)
        vals = {pg_field: round(row[raw] / g, 1) for raw, pg_field in present.items()}
        if any(vals.get(pg) is not None and vals[pg] > PLAUSIBLE_MAX_PG.get(pg, 1e9) for pg in vals):
            dropped_implausible += 1
            continue
        out[key] = vals
        out[key]['games'] = g
    if dropped_implausible:
        print(f"  Prior2-season blend ({season}): dropped {dropped_implausible} players with "
              f"implausible per-game values (likely a games-count mismatch) rather than risk "
              f"corrupting the blend with them.")
    print(f"  Prior2-season blend: {len(out)} players found in {season} data")
    return out


def build_gsis_name_map(*frames):
    """v4.4 NEW — gsis_id -> full display name, built from any roster frames
    (pass older frames first; later frames override, so the most current
    roster's spelling wins). nflverse PBP identifies every rusher/receiver by
    a stable gsis_id alongside an ABBREVIATED name ('J.Cook'); the player
    profiles here are keyed by FULL names ('James Cook'). Matching on the
    abbreviated name can't work (see compute_opportunity_shares_from_pbp), so
    the id is the bridge. Returns {} if no usable frame is supplied."""
    m = {}
    for df in frames:
        if df is None or getattr(df, 'empty', True):
            continue
        idc = next((c for c in ('gsis_id', 'player_id') if c in df.columns), None)
        nmc = next((c for c in ('full_name', 'player_name', 'display_name') if c in df.columns), None)
        if not idc or not nmc:
            continue
        for gid, nm in zip(df[idc], df[nmc]):
            if isinstance(gid, str) and gid.strip() and isinstance(nm, str) and nm.strip():
                m[gid.strip()] = nm.strip()
    return m


def compute_opportunity_shares_from_pbp(pbp, id_to_name=None, counts_out=None):
    """v3.63 NEW / v4.4 FIXED — real per-player opportunity-share metrics
    computed from play-by-play data: carryShare, breakawayRate, rzCarryShare
    (RB) and rzTargetShare (WR/TE/RB).

    v4.4 REAL BUG FIXED — this used to key every result by
    norm_name(<PBP abbreviated name>) e.g. 'aabdullah' (from 'A.Abdullah'),
    but the player profiles it merges into are keyed by FULL names
    ('ameer abdullah'). Measured against the real 2025 PBP and the live
    JSON: 520 players computed, 4 matched (0.8%) — so carryShare,
    breakawayRate, rzCarryShare and rzTargetShare were empty for every
    player in production, and every RB outside the ~30 hand-typed starters
    silently fell back to the RB1 default carry share (0.55). Fixed by
    grouping on the stable gsis player id and translating it to a full name
    via `id_to_name` (see build_gsis_name_map). Grouping by id also stops two
    same-abbreviation teammates (e.g. two 'J.Smith') from being merged into
    one player. A player who appears for two teams keeps the higher-volume
    team's row. Falls back to the old abbreviated-name key when no id or no
    mapping exists, so nothing regresses. If `counts_out` (a dict) is passed it is
    filled with per-player sample sizes {pbp_carries, rz_carries, rz_tgts} so callers
    can weight small samples. Returns {} on missing/empty PBP."""
    if pbp is None or pbp.empty:
        return {}
    id_to_name = id_to_name or {}

    def safe_round(v, nd=3):
        try:
            f = float(v)
            if math.isnan(f) or math.isinf(f):
                return 0.0
            return round(f, nd)
        except (TypeError, ValueError):
            return 0.0

    def _key(pid, abbr):
        nm = id_to_name.get(str(pid).strip()) if isinstance(pid, str) else None
        return norm_name(nm) if nm else norm_name(abbr)

    reg = pbp[pbp["season_type"] == "REG"].copy() if "season_type" in pbp.columns else pbp.copy()
    if "rusher_player_name" not in reg.columns or "passer_player_name" not in reg.columns:
        return {}
    rid = "rusher_player_id" if "rusher_player_id" in reg.columns else "rusher_player_name"
    vid = "receiver_player_id" if "receiver_player_id" in reg.columns else "receiver_player_name"
    rush = reg[(reg.get("rush_attempt", pd.Series(dtype=float)) == 1) & reg["rusher_player_name"].notna()].copy()
    passp = reg[(reg.get("pass_attempt", pd.Series(dtype=float)) == 1) & reg["passer_player_name"].notna()].copy()

    out = {}
    if not rush.empty and "yardline_100" in rush.columns:
        team_carries = rush.groupby("posteam")["rush_attempt"].sum().rename("team_carries")
        rb_agg = rush.groupby([rid, "posteam"]).agg(
            abbr=("rusher_player_name", "first"),
            carries=("rush_attempt", "sum"),
            breakaways=("rushing_yards", lambda s: (s >= 15).sum()),
        ).reset_index()
        rb_agg = rb_agg.merge(team_carries, on="posteam", how="left")
        rb_agg["carryShare"] = rb_agg["carries"] / rb_agg["team_carries"].clip(lower=1)
        rb_agg["breakawayRate"] = rb_agg["breakaways"] / rb_agg["carries"].clip(lower=1)

        rz_rush = rush[rush["yardline_100"] <= 20]
        if not rz_rush.empty:
            team_rz_carries = rz_rush.groupby("posteam")["rush_attempt"].sum().rename("team_rz_carries")
            rz_agg = rz_rush.groupby([rid, "posteam"]).agg(rz_carries=("rush_attempt", "sum")).reset_index()
            rz_agg = rz_agg.merge(team_rz_carries, on="posteam", how="left")
            rz_agg["rzCarryShare"] = rz_agg["rz_carries"] / rz_agg["team_rz_carries"].clip(lower=1)
            rb_agg = rb_agg.merge(rz_agg[[rid, "posteam", "rzCarryShare", "rz_carries"]], on=[rid, "posteam"], how="left")

        for _, r in rb_agg.sort_values("carries", ascending=False).iterrows():
            key = _key(r[rid], r["abbr"])
            if key in out:
                continue  # traded player: keep the higher-volume team's row
            out[key] = {
                "carryShare":    safe_round(r["carryShare"]),
                "breakawayRate": safe_round(r["breakawayRate"]),
                "rzCarryShare":  safe_round(r.get("rzCarryShare", 0)),
            }
            if counts_out is not None:
                _rzc = r.get("rz_carries", 0)
                counts_out[key] = {"pbp_carries": int(r["carries"]),
                                   "rz_carries": 0 if (_rzc is None or (isinstance(_rzc, float) and math.isnan(_rzc))) else int(_rzc)}

    if not passp.empty and "yardline_100" in passp.columns and "receiver_player_name" in passp.columns:
        rz_pass = passp[passp["yardline_100"] <= 20]
        if not rz_pass.empty:
            team_rz_tgts = rz_pass.groupby("posteam")["pass_attempt"].sum().rename("team_rz_tgts")
            rz_wr = rz_pass.groupby([vid, "posteam"]).agg(
                abbr=("receiver_player_name", "first"),
                rz_tgts=("pass_attempt", "sum")).reset_index()
            rz_wr = rz_wr.merge(team_rz_tgts, on="posteam", how="left")
            rz_wr["rzTargetShare"] = rz_wr["rz_tgts"] / rz_wr["team_rz_tgts"].clip(lower=1)
            seen_rz = set()
            for _, r in rz_wr.sort_values("rz_tgts", ascending=False).iterrows():
                key = _key(r[vid], r["abbr"])
                if key in seen_rz:
                    continue
                seen_rz.add(key)
                out.setdefault(key, {})
                out[key]["rzTargetShare"] = safe_round(r["rzTargetShare"])
                if counts_out is not None:
                    counts_out.setdefault(key, {})["rz_tgts"] = int(r["rz_tgts"])

    _mapped = sum(1 for k in out if k in {norm_name(v) for v in id_to_name.values()}) if id_to_name else 0
    print(f"  Opportunity shares (from PBP): {len(out)} players "
          f"(carryShare/breakawayRate/rzCarryShare/rzTargetShare); "
          f"{_mapped} resolved to full names via gsis id")
    return out


def fetch_snap_counts(season):
    """v3.63 NEW — real per-player snap % from nflverse's dedicated
    snap_counts dataset (offense_pct column), season-averaged. This is a
    genuinely new data source for this pipeline — weekly_rosters (already
    fetched) carries roster/depth status but not snap percentage.
    Returns {} on fetch failure, same convention as fetch_prior2_season_players."""
    url = f"{NFLVERSE_BASE}/snap_counts/snap_counts_{season}.csv"
    df = fetch_csv(url, f"snap_counts ({season})", silent_404=True)
    if df is None or df.empty:
        print(f"  Snap counts: no {season} data found — snapShare will be skipped this run.")
        return {}
    name_col = next((c for c in ['player', 'full_name', 'player_display_name'] if c in df.columns), None)
    if not name_col or 'offense_pct' not in df.columns:
        print(f"  Snap counts: expected columns not found (have: {list(df.columns)[:8]}...) — skipping.")
        return {}
    grouped = df.groupby(name_col)['offense_pct'].mean()
    out = {}
    for raw_name, pct in grouped.items():
        try:
            val = float(pct)
        except (TypeError, ValueError):
            continue
        if math.isnan(val):
            continue
        out[norm_name(raw_name)] = round(val, 3)
    print(f"  Snap counts: {len(out)} players found in {season} data")
    return out


# ══════════════════════════════════════════════════════════════════════════
# v3.63 NEW — JOINT RECENCY_DECAY_RATE / SEASON_BLEND_K OPTIMIZER
# ══════════════════════════════════════════════════════════════════════════
#
# Ground truth: BASELINE_SEASON's own real weekly data. At each of several
# "as-of" checkpoints (game 3, 6, 9 of that season), compute what the blend
# WOULD have projected using only games up to that checkpoint plus real
# BASELINE_SEASON-1/-2 season averages — then compare against BASELINE_SEASON's
# actual, now-known final full-season per-game average. No lookahead: the
# checkpoint's partial data is real, in-order, played-so-far data, never
# information from later in that same season. This uses only data the
# pipeline already fetches (fetch_current_season_partial_stats works for
# any season's weekly file, not just the current one; fetch_prior2_season_players
# for the two prior tiers) — no external CSV or extra network source needed.
#
# Metric: mean absolute error (MAE) between projected and real final
# per-game average, across all eligible players/stats — same standard used
# throughout this pipeline and the HTML model's own backtest tooling this
# session. Also reports % of projections within 20% of the real value, for
# a second, more interpretable view of the same result.

def _blend_at_checkpoint(partial_avg, prior1_pg, prior2_pg, games_so_far, k, decay_rate):
    w_current = min(1.0, games_so_far / float(k)) if games_so_far > 0 else 0.0
    remaining = 1 - w_current
    if prior2_pg is not None:
        w_prior1 = remaining / (1 + decay_rate)
        w_prior2 = remaining * decay_rate / (1 + decay_rate)
        return w_current * partial_avg + w_prior1 * prior1_pg + w_prior2 * prior2_pg
    return w_current * partial_avg + remaining * prior1_pg


# v3.86 NEW — translates this pipeline's internal per-game-average field
# names to the JS model's own propType strings, needed so the exported
# optimizer cases (below) can be joined there against real historical
# prop lines by player+propType+week.
PG_FIELD_TO_PROPTYPE = {
    'passYdsPG': 'pass_yds', 'rushYdsPG': 'rush_yds',
    'recYdsPG': 'rec_yds', 'recsPG': 'recs',
}


def optimize_blend_params(baseline_weekly_df, prior1_data, prior2_data, season, checkpoints=(3, 6, 9),
                           decay_grid=(0.0, 0.25, 0.5, 0.75, 1.0), k_grid=(3, 5, 6, 8, 10, 12)):
    """v3.63 NEW — grid-searches (RECENCY_DECAY_RATE, SEASON_BLEND_K) jointly,
    scored by real held-out accuracy on BASELINE_SEASON's own weekly data
    (see module docstring above for the full method). Returns
    {best: {decay_rate, k, mae, pct_within_20}, grid: [...all results...]}
    so the choice is auditable, not a black box. Safe no-op (returns None)
    if baseline_weekly_df is empty — same "no data yet" convention as the
    rest of this pipeline's optional enhancements."""
    from datetime import datetime, timezone  # local import matches this file's existing pattern
    if baseline_weekly_df is None or baseline_weekly_df.empty:
        print("  Blend optimizer: no baseline weekly data available — skipping.")
        return None

    name_col = next((c for c in ['player_display_name', 'player_name', 'full_name', 'display_name']
                      if c in baseline_weekly_df.columns), None)
    if not name_col or 'week' not in baseline_weekly_df.columns:
        print("  Blend optimizer: couldn't find name/week columns — skipping.")
        return None

    stat_cols = {'passing_yards': 'passYdsPG', 'rushing_yards': 'rushYdsPG',
                 'receiving_yards': 'recYdsPG', 'receptions': 'recsPG'}
    present = {raw: pg for raw, pg in stat_cols.items() if raw in baseline_weekly_df.columns}
    if not present:
        print("  Blend optimizer: no usable stat columns found — skipping.")
        return None

    df = baseline_weekly_df.copy()
    df['_key'] = df[name_col].map(norm_name)
    max_week = int(df['week'].max())
    eligible_checkpoints = [c for c in checkpoints if c < max_week]
    if not eligible_checkpoints:
        print(f"  Blend optimizer: season only has {max_week} weeks — no checkpoint leaves room "
              f"for a held-out comparison. Skipping.")
        return None

    # Real, final full-season per-game average per player/stat (the answer
    # each checkpoint's blend is trying to predict).
    final_avg = df.groupby('_key')[list(present.keys())].mean()
    games_played = df.groupby('_key')['week'].nunique()

    # Build one evaluation case per (player, stat, checkpoint) with enough
    # real prior-season data to blend and enough real final-season games to
    # trust as ground truth (>=3, matching this pipeline's existing
    # small-sample guards elsewhere).
    cases = []
    for key, g_final in games_played.items():
        if g_final < 3:
            continue
        p1 = prior1_data.get(key)
        if not p1:
            continue
        p2 = prior2_data.get(key) if prior2_data else None
        for raw, pg_field in present.items():
            prior1_val = p1.get(pg_field)
            if prior1_val is None:
                continue
            prior2_val = p2.get(pg_field) if p2 else None
            true_final = final_avg.loc[key, raw]
            if pd.isna(true_final):
                continue
            for cp in eligible_checkpoints:
                partial = df[(df['_key'] == key) & (df['week'] <= cp)]
                g_partial = partial['week'].nunique()
                if g_partial < 1:
                    continue
                partial_avg = partial[raw].mean()
                if pd.isna(partial_avg):
                    continue
                cases.append({
                    'key': key, 'propType': PG_FIELD_TO_PROPTYPE.get(pg_field, pg_field),
                    'checkpoint': cp, 'season': season,
                    'partial_avg': partial_avg, 'games_partial': g_partial,
                    'prior1': prior1_val, 'prior2': prior2_val, 'true_final': true_final,
                })

    if not cases:
        print("  Blend optimizer: no eligible player/stat/checkpoint cases found — skipping.")
        return None

    grid_results = []
    for decay_rate in decay_grid:
        for k in k_grid:
            errors, within20 = [], 0
            for c in cases:
                proj = _blend_at_checkpoint(c['partial_avg'], c['prior1'], c['prior2'],
                                             c['games_partial'], k, decay_rate)
                err = abs(proj - c['true_final'])
                errors.append(err)
                if c['true_final'] > 0 and err / c['true_final'] <= 0.20:
                    within20 += 1
            mae = sum(errors) / len(errors)
            pct20 = round(within20 / len(errors) * 100, 1)
            grid_results.append({'decay_rate': decay_rate, 'k': k, 'mae': round(mae, 3),
                                  'pct_within_20': pct20, 'n_cases': len(errors)})

    grid_results.sort(key=lambda r: r['mae'])
    best = grid_results[0]
    print(f"  Blend optimizer: tested {len(grid_results)} (decay_rate, k) combinations "
          f"across {len(cases)} real held-out player/stat/checkpoint cases")
    print(f"  Best: decay_rate={best['decay_rate']}, k={best['k']} "
          f"(MAE={best['mae']}, {best['pct_within_20']}% within 20% of real final average)")
    # v3.71 NEW — this is what caught the games-count bug above: a jump
    # here across decay_rate values means the extra data source (2024)
    # is behaving badly, not just failing to help. Shows the best k's MAE
    # at each decay_rate so that's visible directly in the console instead
    # of needing to open blend_params.json to see the full grid.
    by_decay = {}
    for r in grid_results:
        dr = r['decay_rate']
        if dr not in by_decay or r['mae'] < by_decay[dr]['mae']:
            by_decay[dr] = r
    print(f"  Decay-rate summary (best k at each decay_rate — a big jump here means the "
          f"2024 data is hurting, not just failing to help):")
    for dr in sorted(by_decay.keys()):
        r = by_decay[dr]
        print(f"    decay_rate={dr}: best MAE={r['mae']} (k={r['k']}, {r['pct_within_20']}% within 20%)")
    return {'best': best, 'grid': grid_results, 'n_cases': len(cases), 'cases': cases,
            'evaluated_at': datetime.now(timezone.utc).isoformat()}


def load_or_optimize_blend_params(script_dir, season, baseline_weekly_df=None, prior1_data=None, prior2_data=None,
                                   force=False):
    """v3.63 NEW — originally the persisted-params half of the optimizer,
    reading blend_params.json if present to avoid recomputing every run.
    v3.70 FIX — removed the cache-read per explicit request: the grid
    search itself is cheap (well under a second over a few thousand
    already-in-memory rows), so there was no real performance reason to
    cache it — and the caching was the direct cause of several rounds of
    'stale results' confusion while debugging the 2024-data fetch above
    (this kept loading an old decay_rate/k computed before that fix
    landed, making it look like nothing had changed). Now always
    recomputes fresh from whatever real data is available this run. The
    `force` parameter and blend_params.json write are both kept — the
    file still gets saved each run as a human-readable record of the
    latest decision, it's just never read back as a cache. Falls back to
    the original hardcoded RECENCY_DECAY_RATE/SEASON_BLEND_K defaults
    only if no baseline data was passed in to compute one at all."""
    params_path = os.path.join(script_dir, 'blend_params.json')

    if baseline_weekly_df is None:
        print(f"  Blend params: no baseline data to compute one — "
              f"using defaults (decay_rate={RECENCY_DECAY_RATE}, k={SEASON_BLEND_K}).")
        return RECENCY_DECAY_RATE, SEASON_BLEND_K, None

    result = optimize_blend_params(baseline_weekly_df, prior1_data or {}, prior2_data or {}, season)
    if result is None:
        return RECENCY_DECAY_RATE, SEASON_BLEND_K, None

    try:
        with open(params_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2, default=str)
        print(f"  Blend params: saved to {params_path}")
    except Exception as e:
        print(f"  Blend params: couldn't save {params_path} ({e}) — using this run's result anyway.")

    return result['best']['decay_rate'], result['best']['k'], result


def export_blend_optimizer_cases(output_dir, season, opt_result):
    """v3.86 NEW — writes the raw per-player/stat/checkpoint cases the
    decay-rate optimizer above already builds internally to their own
    JSON file, separate from nfl_model_data.json (this is analysis input
    for a specific tool, not live model data). Each case already carries
    everything needed to join it against a real historical prop line on
    the JS side: player key, propType (translated to the JS model's own
    naming via PG_FIELD_TO_PROPTYPE), the checkpoint week, and the season.
    Safe no-op if there are no cases (e.g. optimizer skipped this run for
    lack of data) — writes nothing rather than an empty/misleading file."""
    cases = opt_result.get('cases') if opt_result else None
    if not cases:
        print(f"  Blend optimizer cases: nothing to export for {season} (optimizer had no cases this run).")
        return
    out_path = os.path.join(output_dir, f'nfl_blend_optimizer_cases_{season}.json')
    payload = {'season': season, 'n_cases': len(cases), 'cases': cases,
               'evaluated_at': opt_result.get('evaluated_at')}
    try:
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, indent=2, default=str)
        print(f"  Blend optimizer cases: exported {len(cases)} cases to {out_path}")
    except Exception as e:
        print(f"  Blend optimizer cases: couldn't save {out_path} ({e}).")


def fetch_current_season_partial_stats(current_season):
    """v3.39 NEW / v4.4 FIXED — REG-season weekly player stats for `current_season`
    (works for any season). Delegates to fetch_weekly_stats(): local copy first,
    then the CORRECT nflverse path (`stats_player/…`). The path this used before
    (`player_stats/…`) 404s for every season, which was swallowed as "no data
    published yet" and silently disabled the season-transition blend, the blend
    optimizer and the healthy-roster shares — while 2026 weeks 1-3 were public."""
    return fetch_weekly_stats(current_season)


def apply_season_transition_blend(players, current_season, k=None, prior2_data=None, decay_rate=None, resolver=None):
    """v3.39 original 2-season blend, v3.55 EXTENDED to a real multi-year
    blend. Still blends each player's BASELINE_SEASON stat with real
    CURRENT_SEASON production as it becomes available (w_current =
    get_season_blend_weight(games) — unchanged from v3.39, already
    correctly distrusts a tiny early-season sample). NEW: whatever weight
    current-season doesn't use (1-w_current) now splits across
    BASELINE_SEASON and BASELINE_SEASON-1 (when prior2_data has that
    player) using RECENCY_DECAY_RATE — BASELINE_SEASON gets
    remaining/(1+decay), BASELINE_SEASON-1 gets
    remaining*decay/(1+decay). A player with no prior2 data (e.g. a
    2nd-year player with no BASELINE_SEASON-1 in the league) gets the
    full remaining weight on BASELINE_SEASON alone — same normalization
    principle the existing rookie college/NFL blend already uses, not a
    new special case. `prior2_data` is the dict fetch_prior2_season_players()
    returns; pass None to skip this tier entirely and fall back to the
    exact original v3.39 2-season behavior (e.g. if that fetch failed).
    Sets `<stat>_seasonBlended` exactly as before — no JS-side changes
    needed, that field is already preferred when present. A player with 0
    CURRENT_SEASON games gets no blended fields at all (pure no-op, the
    existing baseline is already the right answer)."""
    k = k or SEASON_BLEND_K
    decay_rate = RECENCY_DECAY_RATE if decay_rate is None else decay_rate
    df = fetch_current_season_partial_stats(current_season)
    if df is None or df.empty:
        print(f"  Season-transition blend: no {current_season} data published yet — "
              f"no-op (expected pre-Week-1; every player keeps their {current_season-1} baseline).")
        return 0

    name_col = next((c for c in ['player_display_name', 'player_name',
                                  'full_name', 'display_name'] if c in df.columns), None)
    if not name_col:
        print(f"  Season-transition blend: couldn't find a name column on the "
              f"{current_season} weekly file (columns: {list(df.columns)[:10]}...) — skipping.")
        return 0

    stat_cols = {
        'passing_yards':   'passYdsPG',
        'rushing_yards':   'rushYdsPG',
        'receiving_yards': 'recYdsPG',
        'receptions':      'recsPG',
    }
    present_stats = {raw: pg for raw, pg in stat_cols.items() if raw in df.columns}
    if not present_stats:
        print(f"  Season-transition blend: none of {list(stat_cols.keys())} found on "
              f"the {current_season} weekly file — skipping.")
        return 0

    agg = {raw: 'sum' for raw in present_stats}
    grouped = df.groupby(name_col).agg(agg)
    games_played = df.groupby(name_col)['week'].nunique() if 'week' in df.columns \
        else df.groupby(name_col).size()

    prior2_data = prior2_data or {}
    blended_count = 0
    multiyear_count = 0
    for raw_name, row in grouped.iterrows():
        # v4.4: resolve through the id/full-name resolver so 'Travis Etienne' (weekly file)
        # finds 'travis etienne jr' (roster-keyed profile) — the plain norm_name lookup
        # missed every player whose two sources disagree on a generational suffix.
        key = (resolver.resolve(name=raw_name) if resolver is not None else None) or norm_name(raw_name)
        p = players.get(key)
        if not p:
            continue
        g = int(games_played.get(raw_name, 0))
        if g <= 0:
            continue
        w_current = get_season_blend_weight(g, k)
        remaining = 1 - w_current
        p2 = prior2_data.get(key)
        has_prior2 = p2 is not None
        if has_prior2:
            w_prior1 = remaining / (1 + decay_rate)
            w_prior2 = remaining * decay_rate / (1 + decay_rate)
        else:
            w_prior1 = remaining
            w_prior2 = 0.0
        weights = {
            'current': round(w_current, 3), 'prior1': round(w_prior1, 3),
            'prior2': round(w_prior2, 3), 'games': g, 'hasPrior2Data': has_prior2,
        }
        did_blend = False
        for raw, pg_field in present_stats.items():
            prior1_val = p.get(pg_field)
            if prior1_val is None:
                continue
            current_pg = row[raw] / g
            prior2_val = p2.get(pg_field) if (has_prior2 and pg_field in p2) else None
            if prior2_val is None:
                # No real prior2 value for this specific stat even though
                # the player has SOME prior2 data (e.g. a pure rusher with
                # no receiving stats that year) — fold that share back
                # onto prior1 rather than silently dropping it to 0.
                blended = round(w_current * current_pg + (w_prior1 + w_prior2) * prior1_val, 1)
            else:
                blended = round(w_current * current_pg + w_prior1 * prior1_val + w_prior2 * prior2_val, 1)
            p[f'{pg_field}_seasonBlended'] = blended
            did_blend = True
        if did_blend:
            p['seasonBlendWeights'] = weights
            p['seasonBlendNote'] = (
                f"{current_season}: {g}g played — "
                f"{int(round(weights['current']*100))}% {current_season} + "
                f"{int(round(weights['prior1']*100))}% {current_season-1}" +
                (f" + {int(round(weights['prior2']*100))}% {current_season-2}" if has_prior2 else "")
            )
            blended_count += 1
            if has_prior2:
                multiyear_count += 1

    print(f"  Season-transition blend: {blended_count} players blended with real "
          f"{current_season} data (k={k} games to full weight), "
          f"{multiyear_count} of those using a real {current_season-2} 3rd tier too")
    return blended_count


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
    """Normalise player names for matching"""
    return re.sub(r"[^a-z ]", "", str(s).lower().strip())


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
        # v3.67 FIX — renamed by nflverse (see fetch_prior2_season_players
        # docstring for the confirmed real current filename).
        _ps_try = fetch_csv(f"{BASE}/player_stats/stats_player_reg_{season}.csv",
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
                # v3.67 FIX — same rename as above.
                print(f"  PBP failed — falling back to {prior} player stats")
                _ps_try2 = fetch_csv(f"{BASE}/player_stats/stats_player_reg_{prior}.csv",
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

            # v4.0 REAL FIX — no minimum-activity threshold existed here at all;
            # every player with any row in player_season/player_stats was
            # included regardless of real playing time, meaning a single
            # garbage-time snap got the same treatment as a real contributor.
            # This wasn't causing under-coverage (the HTML model's separate,
            # hand-maintained PLAYER_PROPS_2026 static block was the actual
            # bottleneck — see findPlayerPropsData fix), but it's a real gap
            # worth closing now that this fuller player pool is being wired
            # into the live model: a meaningful floor keeps noisy, unreliable
            # single-game entries out rather than letting them surface as if
            # they were trustworthy. Floor: 2+ games AND real position-
            # appropriate volume (targets for WR/TE, carries+targets for RB,
            # attempts for QB) — low enough to keep real WR3/WR4/backup-TE
            # usage, high enough to drop pure one-off garbage-time entries.
            _targets  = safe_float(row.get('targets', 0))
            _carries  = safe_float(row.get('carries', 0))
            _attempts = safe_float(row.get('attempts', 0))
            if games < 2:
                continue
            if pos in ('WR', 'TE') and _targets < 8:
                continue
            if pos == 'RB' and (_carries + _targets) < 8:
                continue
            if pos == 'QB' and _attempts < 10:
                continue

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

                # v3.30 REAL BUG FIXED — confirmed via the NFL Edge Model chat:
                # this step used to build a brand-new, PASSING-ONLY record and
                # assign it with `players[name] = entry4`, which OVERWRITES
                # (not merges) whatever record already existed for that name —
                # including a real, rushing-inclusive record the earlier
                # nflverse_player_stats loop above had just built. Since every
                # real starting QB throws enough passes to clear this step's
                # `_gm4 < 2` filter, virtually every starter's real rushYdsPG
                # was being silently destroyed here, every run — only
                # low-snap backups who never cleared that filter kept their
                # rushing data by accident. Confirmed empirically: Lamar
                # Jackson, Josh Allen, Jalen Hurts, Caleb Williams all ended up
                # with source='pbp_passer' and NO rushYdsPG field at all.
                # Fixed two ways: (1) compute a real, PBP-derived rushYdsPG
                # for every passer from the same _pbp4 dataframe (rushing
                # plays are in the same play-by-play data, just under
                # rusher_player_name instead of passer_player_name — no new
                # data source needed), and (2) MERGE into any existing entry
                # instead of replacing it outright, so no field computed
                # elsewhere for this player is ever silently lost again.
                _rush4 = _reg4[_reg4['rush_attempt']==1] if 'rush_attempt' in _reg4.columns else _reg4.iloc[0:0]
                _rush4 = _rush4[_rush4['rusher_player_name'].astype(str)!='nan'] if 'rusher_player_name' in _rush4.columns else _rush4.iloc[0:0]
                _qb_rush_yds = _rush4.groupby('rusher_player_name')['rushing_yards'].sum() if not _rush4.empty else pd.Series(dtype=float)
                _qb_rush_games = _rush4.groupby('rusher_player_name')['game_id'].nunique() if not _rush4.empty else pd.Series(dtype=int)

                _TNORM4 = {'JAC':'JAX','LA':'LAR','WSH':'WAS','LVR':'LV','NWE':'NE','NOR':'NO',
                           'GNB':'GB','TBB':'TB','KCC':'KC','SFO':'SF'}
                def _np4(t): return _TNORM4.get(str(t).upper(), str(t).upper())
                _qb_added = 0
                for _, _r4 in _qb_grp.iterrows():
                    _nm4 = str(_r4['passer_player_name'])
                    _tm4 = _np4(_r4['posteam'])
                    _gm4 = int(_r4['_gm'])
                    if _gm4 < 2 or not _nm4: continue
                    # Real, per-QB rushing from this same PBP dataset — uses
                    # the QB's OWN games-played-as-a-passer as the games
                    # denominator (not the rushing-specific game count),
                    # since a QB's rushing volume is a rate relative to their
                    # real playing time, matching how every other position's
                    # rushYdsPG is already computed in this file.
                    _rush_yds_total = float(_qb_rush_yds.get(_nm4, 0.0))
                    _entry4 = {
                        'name': _nm4, 'pos': 'QB', 'team': _tm4, 'games': _gm4,
                        'passYdsPG': round(float(_r4['_py'])/max(_gm4,1), 1),
                        'cpoe':      round(float(_r4['_cpoe'] or 0), 4),
                        'passEPA':   round(float(_r4['_epa'] or 0), 4),
                        'cmpPct':    round(float(_r4['_cmp'])/max(float(_r4['_att']),1)*100, 1),
                        'rushYdsPG': round(_rush_yds_total/max(_gm4,1), 1),
                        'isRookie': False, 'nflGames': _gm4, 'source': 'pbp_passer',
                    }
                    _existing4 = players.get(norm_name(_nm4))
                    if _existing4:
                        _existing4.update(_entry4)
                        players[norm_name(_nm4)] = _existing4
                    else:
                        players[norm_name(_nm4)] = _entry4
                    _qb_added += 1
                print(f"  QB stats built from PBP: {_qb_added} passers (abbreviated names — resolved next step); rushYdsPG now computed for all, not just games<2 backups")
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

    # v3.33 REAL BUG FIXED — confirmed via a real run showing "Marked 0
    # players as rookie" despite a full incoming draft class. Root cause:
    # dfs['rosters'] is populated elsewhere in this file purely as a
    # name-resolution fallback during PBP processing, using whatever
    # `season` is currently being iterated (2025, in preseason baseline
    # mode) — NOT the current CURRENT_SEASON being projected (2026). A
    # player who entered the league in 2026 cannot possibly appear on a
    # roster snapshot dated 2025, so filtering dfs['rosters'] for
    # entry_year==2026 was structurally guaranteed to return zero matches
    # every time, regardless of how many real rookies exist. A completely
    # separate, independent fetch elsewhere in this file already pulls a
    # real roster_2026.csv successfully (confirmed working — it's what
    # powers PBP name resolution), but that flow hardcodes isRookie=False
    # on every entry and never populates this `rookies` set. Fixed by
    # giving detect_rookies() its own direct fetch of the CURRENT_SEASON
    # roster file, independent of whatever dfs['rosters'] happens to hold.
    rook_season = str(CURRENT_SEASON) if 'CURRENT_SEASON' in globals() else str(int(season) + 1)
    rosters = fetch_csv(f"{NFLVERSE_BASE}/rosters/roster_{rook_season}.csv", f"rosters {rook_season} (rookie detection)")

    if not rosters.empty:
        entry_col = next((c for c in rosters.columns if 'entry_year' in c.lower() or 'rookie_year' in c.lower()), None)
        name_col  = next((c for c in rosters.columns if 'display_name' in c.lower() or 'full_name' in c.lower()), None)
        if entry_col and name_col:
            rook_df = rosters[rosters[entry_col].astype(str) == rook_season]
            for _, row in rook_df.iterrows():
                rookies.add(norm_name(row.get(name_col, '')))
            print(f"  Detected {len(rookies)} rookies from rosters (entry_year={rook_season})")
        else:
            print(f"  ! Could not find entry_year/name columns on roster_{rook_season}.csv — 0 rookies detected this run")
    else:
        print(f"  ! roster_{rook_season}.csv came back empty — 0 rookies detected this run")

    # Mark rookies in player profiles
    # v3.39 REAL BUG FIXED — confirmed via direct inspection of a real
    # nfl_model_data.json export: established veterans (Justin Jefferson,
    # DeVonta Smith, 216 players total) came back isRookie=True. The
    # matched entry_year/rookie_year column on roster_{season}.csv can't
    # be schema-verified without a live key, so it's now trusted only
    # when it doesn't contradict production this player already has on
    # record — a genuine incoming rookie cannot already have a real,
    # non-skeleton season of games. This is a safety net, not a fix to
    # the underlying column match itself; if `overridden` below is large
    # on a live run, the actual entry_col being matched needs a direct
    # look (print the column name/values for a few known veterans).
    marked = 0
    overridden = 0
    for key, p in players.items():
        if key in rookies:
            has_real_prod = (
                p.get('source') not in (None, 'roster_2026_skeleton')
                and (p.get('games', 0) or 0) >= 3
            )
            if has_real_prod:
                overridden += 1
                continue
            p['isRookie'] = True
            marked += 1
    print(f"  Marked {marked} players as rookie in profiles")
    if overridden:
        print(f"  ⚠ {overridden} players matched the rookie roster filter but already "
              f"have real, non-skeleton season production — treated as a roster "
              f"entry_year/rookie_year mismatch, not a real rookie, and left alone.")
    return rookies

# ─────────────────────────────────────────────
# 5. LOAD PROSPECT ANALYZER (College Fallback)
# ─────────────────────────────────────────────
def _load_prospect_json(filepath_or_dict):
    """Load prospect data from JSON export. Handles two real, confirmed
    shapes seen in the wild: a flat list of records, OR nfl_prospects_2026.json's
    real shape — {"prospects": {"QB": [...], "RB": [...], "WR": [...], "TE": [...]}}
    (a dict KEYED BY POSITION, not a flat list).

    v3.31 REAL BUG FIXED — confirmed via a real run's error log
    ("'str' object has no attribute 'get'") and by directly downloading and
    inspecting the real nfl_prospects_2026.json file: the old code did
    `records = records.get('prospects', ...)` then `for p in records:` —
    but since the extracted `prospects` value is ITSELF a dict keyed by
    position, iterating it with a plain for-loop yields the STRING KEYS
    ("QB", "RB", "WR", "TE"), not the actual prospect records. Calling
    `p.get('name', '')` on the string "QB" is exactly what threw the
    AttributeError. Fixed to flatten the position-keyed dict into one list
    first, tagging each record with its position from the dict key (in
    case a record's own `pos` field is missing).

    Also fixed field mapping to match what this real file actually
    contains — confirmed by inspection there is NO passYdsPG/rushYdsPG/
    recYdsPG/recsPG/tdsPG anywhere in it (this file is a draft-grade/
    landing-spot export, not a per-game-production export). Those stay at
    0 here and are meant to be filled in by fetch_cfbd_prospect_stats()
    (real per-game college production from CFBD) — this function now only
    captures the real DRAFT-CONTEXT fields this file actually has:
    draftPick (string form, e.g. "1.01"), pick (numeric), tier,
    finalRanking, landingMult, nflTeam.
    """
    import json as _json
    college = {}
    if isinstance(filepath_or_dict, dict):
        records_raw = filepath_or_dict
    else:
        with open(filepath_or_dict, encoding='utf-8') as f:
            records_raw = _json.load(f)

    flat_records = []
    if isinstance(records_raw, list):
        flat_records = records_raw
    else:
        prospects_val = records_raw.get('prospects', records_raw.get('players', []))
        if isinstance(prospects_val, list):
            flat_records = prospects_val
        elif isinstance(prospects_val, dict):
            # Real shape: {"QB": [...], "RB": [...], "WR": [...], "TE": [...]}
            for pos_key, pos_records in prospects_val.items():
                if not isinstance(pos_records, list):
                    continue
                for rec in pos_records:
                    if isinstance(rec, dict) and not rec.get('pos'):
                        rec = {**rec, 'pos': pos_key}
                    flat_records.append(rec)

    for p in flat_records:
        if not isinstance(p, dict):
            continue  # defensive — never crash on an unexpected shape again
        name = str(p.get('name', '')).strip()
        if not name or name.lower() in ('nan', 'player', 'name'): continue
        pos  = str(p.get('pos', p.get('position', ''))).upper().strip()
        conf = str(p.get('conf', p.get('conference', ''))).lower().strip()
        conf_adj = CONF_ADJ.get(conf, CONF_ADJ.get(conf.split()[0] if conf else '', 0.70))
        # draftPick in this real file is a STRING like "1.01" or "UDFA" —
        # `pick` (numeric overall pick) is the separate, real numeric field.
        entry = {
            'name':       name,
            'pos':        pos,
            'school':     str(p.get('college', p.get('school', ''))),
            'conf':       conf,
            'confAdj':    conf_adj,
            'nflTrans':   NFL_TRANS_FACTOR,
            'source':     'ProspectAnalyzer_JSON',
            # Core stats — genuinely not present in this file's real schema;
            # left at 0 here on purpose. fetch_cfbd_prospect_stats() is the
            # real source for these, called separately in apply_rookie_blending().
            'passYdsPG':  float(p.get('passYdsPG', 0) or 0),
            'rushYdsPG':  float(p.get('rushYdsPG', 0) or 0),
            'recYdsPG':   float(p.get('recYdsPG',  0) or 0),
            'recsPG':     float(p.get('recsPG',    0) or 0),
            'tdsPG':      float(p.get('tdsPG',     0) or 0),
            'passYdsPG_adj': round(float(p.get('passYdsPG', 0) or 0) * conf_adj * NFL_TRANS_FACTOR, 1),
            'rushYdsPG_adj': round(float(p.get('rushYdsPG', 0) or 0) * conf_adj * NFL_TRANS_FACTOR, 1),
            'recYdsPG_adj':  round(float(p.get('recYdsPG',  0) or 0) * conf_adj * NFL_TRANS_FACTOR, 1),
            # Real draft-context fields this file actually has
            'draftPickStr':  str(p.get('draftPick', '') or ''),
            'draftPick':     int(p.get('pick', 0) or 0),
            'prospectScore': float(p.get('finalRanking', p.get('score', 0)) or 0),
            'prospectTier':  str(p.get('tier', '')  or ''),
            'landingMult':   float(p.get('landingMult', 1.0) or 1.0),
            'nflTeam':       str(p.get('nflTeam', p.get('team', '')) or '').upper(),
        }
        college[norm_name(name)] = entry
    print(f"  ✅ Loaded {len(college)} prospects from JSON "
          f"({sum(1 for v in college.values() if v['pos']=='QB')} QB / "
          f"{sum(1 for v in college.values() if v['pos']=='RB')} RB / "
          f"{sum(1 for v in college.values() if v['pos']=='WR')} WR / "
          f"{sum(1 for v in college.values() if v['pos']=='TE')} TE)")
    return college


def _fetch_prospect_json_from_github():
    """v3.31 NEW — real GitHub fallback for the prospect JSON, per G-Money's
    explicit question ("why isn't it pulling from GitHub, wouldn't that make
    more sense?"). Confirmed via a direct GitHub API check that
    nfl_prospects_2026.json genuinely exists in the Gatorz1989/
    NFL-Daily-Data-Pull repo, but this script previously had ZERO code to
    fetch anything from it for the prospect step — it only ever checked
    local Windows folders, which is also why this step would silently find
    nothing at all if this script were ever run via the repo's own GitHub
    Actions automation (no local Windows paths exist on that runner).
    Only used when no local file is found — local files still take
    priority so nothing changes for a normal local run.
    """
    url = "https://raw.githubusercontent.com/Gatorz1989/NFL-Daily-Data-Pull/main/nfl_prospects_2026.json"
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"  ⚠ GitHub prospect JSON fetch failed: {e}")
        return None


def load_prospect_analyzer():
    print(f"\n{'='*50}")
    print("LOADING PROSPECT ANALYZER")
    print('='*50)
    college = {}

    if not PROSPECT_FILE.exists():
        print(f"  ⚠ Local file not found: {PROSPECT_FILE}")
        print(f"  Trying GitHub fallback (Gatorz1989/NFL-Daily-Data-Pull)...")
        gh_data = _fetch_prospect_json_from_github()
        if gh_data is not None:
            try:
                college = _load_prospect_json(gh_data)
                print(f"  ✅ Loaded from GitHub fallback")
                return college
            except Exception as e:
                print(f"  ❌ GitHub JSON parse failed: {e}")
        print(f"  Skipping Prospect Analyzer — will use CFBD/CFB Reference only")
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
# 6b. CFBD PROSPECT STATS — real per-game college production + real
#     games-played, per G-Money's explicit request
# ─────────────────────────────────────────────
def fetch_cfbd_prospect_stats(name, pos, college_school='', year=2025):
    """Real per-player college production from CFBD, used as the PRIMARY
    rookie-blending data source (tried before the CFB Reference scrape
    fallback below) — a real, structured API beats scraping HTML.

    v3.34 REAL BUG FIXED — confirmed via a real Anaconda Prompt log showing
    "400 Client Error: Bad Request" on every single /player/search call,
    for every rookie, with a real CFBD_API_KEY now correctly set (ruling
    out an auth/key problem — 400 means the request itself was malformed,
    not rejected credentials, which would be 401/403). Confirmed root
    cause by inspecting the exact raw HTTP query-parameter names the
    official cfbd Python SDK actually builds (not just its Python
    attribute names): the SDK's own internal attribute is `search_term`
    (Python snake_case convention), but the REAL, raw HTTP query parameter
    it sends is `searchTerm` (camelCase) — confirmed directly from the
    SDK's request-building source code. This script makes raw REST calls
    (no SDK dependency), and used `search_term` directly as the literal
    query key, which the real API never recognized, hence Bad Request on
    every call. Fixed to send `searchTerm`.

    v3.31 NEW. Two real CFBD calls per prospect (confirmed via direct
    inspection of the official cfbd Python SDK's generated models — this
    script makes raw REST calls, not via the SDK itself, since the SDK
    isn't a project dependency, but the endpoints/schemas below are
    confirmed real from it):
      1. GET /player/search?searchTerm=...&team=...&year=...
         -> resolves name+school to a real numeric player id
      2. GET /player/season/overview?year=...&player_id=...
         -> returns {"games": <int>, "boxScoreStats": {"categories": [...]}}
         confirmed to include a genuine games-played count, which the
         bulk /stats/player/season endpoint this script already uses
         elsewhere for team-level CFB stats does NOT have.

    Honest limitation flagged to G-Money: the exact stat-name strings
    inside each category (e.g. whether passing yards is labeled "YDS" or
    something else) could not be confirmed without a live API key —
    matching below is deliberately fuzzy (substring, case-insensitive) to
    be resilient to this, but should be spot-checked against a real
    response on first use.
    """
    if not CFBD_API_KEY:
        return {}
    try:
        search_params = {'searchTerm': name, 'year': year}
        if college_school:
            search_params['team'] = college_school
        results = cfbd_get('/player/search', search_params)
        if not results:
            return {}
        # Prefer an exact name match if multiple results come back
        match = next((r for r in results if r.get('name', '').lower() == name.lower()), results[0])
        player_id = match.get('id')
        if not player_id:
            return {}

        overview = cfbd_get('/player/season/overview', {'year': year, 'playerId': player_id})
        if not overview:
            return {}
        games = int(overview.get('games', 0) or 0)
        if games <= 0:
            return {}

        cat_map = {'QB': 'passing', 'RB': 'rushing', 'WR': 'receiving', 'TE': 'receiving'}
        target_cat = cat_map.get(pos, 'receiving')
        categories = (overview.get('boxScoreStats') or {}).get('categories') or []
        cat = next((c for c in categories if target_cat in str(c.get('name', '')).lower()), None)

        def _find_stat(stats_list, *name_fragments):
            for s in (stats_list or []):
                sname = str(s.get('name', '')).lower()
                if any(frag in sname for frag in name_fragments):
                    try:
                        return float(s.get('value', 0) or 0)
                    except (ValueError, TypeError):
                        return 0.0
            return 0.0

        yds_total = _find_stat(cat.get('stats') if cat else None, 'yd')
        conf_raw = str(match.get('conference', '') or overview.get('conference', '') or '').lower().strip()
        conf_adj = CONF_ADJ.get(conf_raw, CONF_ADJ.get(conf_raw.split()[0] if conf_raw else '', 0.70))

        stat_key = {'QB': 'passYdsPG', 'RB': 'rushYdsPG', 'WR': 'recYdsPG', 'TE': 'recYdsPG'}.get(pos, 'recYdsPG')
        per_game = round(yds_total / games, 1)

        return {
            'conf': conf_raw, 'confAdj': conf_adj, 'nflTrans': NFL_TRANS_FACTOR,
            'source': 'CFBD', 'games': games,
            stat_key: per_game,
            f'{stat_key}_adj': round(per_game * conf_adj * NFL_TRANS_FACTOR, 1),
        }
    except Exception as e:
        print(f"    CFBD prospect stats error for {name}: {e}")
        return {}

# ─────────────────────────────────────────────
# 7. CFB REFERENCE FALLBACK — secondary fallback if CFBD has no key/fails
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

        # Find college data (draft context: school/tier/nflTeam from the
        # Prospect Analyzer JSON, if a real prospect record exists there)
        col = college_data.get(key, {})

        # v3.31 REWORKED — real per-game production fallback chain, per
        # G-Money's explicit request: the Prospect Analyzer JSON has real
        # draft context (school, tier, nflTeam) but genuinely NO per-game
        # production stats (confirmed by inspecting the real file) — so
        # `col` can exist with real context but all-zero stats. Checking
        # only `if not col` (the old condition) would skip the real fix
        # entirely whenever a prospect record existed with empty stats,
        # which is every single prospect in that file. Now checks whether
        # the position-relevant stat is actually present, tries CFBD
        # first (real, structured API), and MERGES the result into any
        # existing `col` so draft-context fields are never lost — rather
        # than replacing `col` outright.
        pos_stat_key = {'QB': 'passYdsPG', 'RB': 'rushYdsPG', 'WR': 'recYdsPG', 'TE': 'recYdsPG'}.get(pos, 'recYdsPG')
        has_real_stats = bool(col.get(pos_stat_key))
        if not has_real_stats and nfl_games == 0:
            print(f"  Fetching real college stats for rookie: {p['name']} ({pos})...")
            cfbd_stats = fetch_cfbd_prospect_stats(p['name'], pos, col.get('school', ''))
            if cfbd_stats.get(pos_stat_key):
                col = {**col, **cfbd_stats}
                print(f"    ✅ CFBD: {cfbd_stats[pos_stat_key]} {pos_stat_key} over {cfbd_stats.get('games','?')} games")
            else:
                print(f"    CFBD had nothing (no key, or no match) — trying CFB Reference scrape...")
                ref_stats = fetch_cfb_ref_stats(p['name'], pos)
                if ref_stats.get(pos_stat_key):
                    col = {**col, **ref_stats}
            college_data[key] = col  # cache whatever we ended up with

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

def _extract_starters_impl(dfs):
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

    # v3.39 REAL BUG FIXED — confirmed via direct inspection of a real
    # nfl_model_data.json export: wr3 was None for every team. Root cause:
    # most teams' raw depth charts only carry 2 distinct pos_rank==1 WR
    # formation slots (e.g. X/Z), not 3 — so the "multiple rank=1 rows per
    # team" trick above genuinely can't surface a 3rd starter for most
    # teams; it only ever worked for the handful of teams whose depth
    # chart happens to list 3 separate rank-1 WR formation rows. The real
    # WR3 (primary slot receiver in 3-WR sets) is usually listed at
    # pos_rank==2 in the raw data instead of its own rank==1 row. Fixed by
    # falling back to rank-2 WR rows, in depth order, for any team still
    # short of 3 unique names — this only fills a gap, it never overwrites
    # a real rank-1 name already captured above.
    rank2_df = dc[dep_numeric == 2]
    wr2_per_team = {}
    for _, row in rank2_df.iterrows():
        team = str(row.get(team_c, '')).upper().strip()
        team = _DEPTH_ABBR.get(team, team)
        pos  = str(row.get(pos_c, '')).upper().strip()
        name = str(row.get(name_c, '')).strip()
        if pos == 'WR' and team and name and name.lower() != 'nan':
            wr2_per_team.setdefault(team, [])
            if name not in wr2_per_team[team]:
                wr2_per_team[team].append(name)

    for team, names in wr2_per_team.items():
        existing = wr_per_team.setdefault(team, [])
        for name in names:
            if len(existing) >= 3:
                break
            if name not in existing:
                existing.append(name)

    for team, wr_names in wr_per_team.items():
        if team not in starters: starters[team] = {}
        if wr_names:               starters[team]['wr1'] = wr_names[0]
        if len(wr_names) > 1:      starters[team]['wr2'] = wr_names[1]
        if len(wr_names) > 2:      starters[team]['wr3'] = wr_names[2]

    n = sum(1 for v in starters.values() if v)
    n_wr3 = sum(1 for v in starters.values() if v.get('wr3'))
    print(f"  Extracted starters for {n} teams from depth charts (wr3 found for {n_wr3} teams)")
    return starters

def receiver_key(name):
    """
    Normalize a player name to 'firstinitial_lastname' for matching between
    nflverse PBP's abbreviated format ('T.McBride') and full roster names
    ('Trey McBride Jr.'). Strips suffixes and hyphens for consistent matching
    on both sides. Discovered necessary via a live PBP pull test — the
    original norm_name()-based matching produced zero role matches against
    real nflverse data because PBP receiver names are abbreviated.
    """
    name = str(name).strip()
    if not name:
        return ''
    if '.' in name.split(' ')[0]:
        # PBP format: "T.McBride" or "J.Smith-Schuster"
        parts = name.split('.', 1)
        initial = parts[0][:1].lower()
        last = parts[1] if len(parts) > 1 else ''
    else:
        # Full name format: "Trey McBride" / "JuJu Smith-Schuster" / "Michael Pittman Jr."
        tokens = name.split()
        tokens = [t for t in tokens if t.lower().rstrip('.') not in
                  ('jr', 'sr', 'ii', 'iii', 'iv', 'v')]
        if not tokens:
            return ''
        initial = tokens[0][:1].lower()
        last = tokens[-1]
    last = re.sub(r'[^a-zA-Z]', '', last).lower()
    return f"{initial}_{last}"


def build_coverage_proxy(pbp, starters):
    """
    NFLVERSE PBP-DERIVED PASS-DEFENSE-BY-ROLE PROXY.

    Free, TOS-clean alternative to PFF coverage grades (PFF's terms prohibit
    automated or manual extraction of their data — see project notes).

    For each team's current WR1/WR2/TE (from depth-chart starters), finds every
    pass play where that player was targeted, groups by the OPPOSING defense
    (defteam), and computes EPA/target allowed against that specific role.
    Converts to a 0-100 grade via percentile rank across all 32 defenses so it
    displays on the same scale as the existing static cb1Grade/cb2Grade fields
    (higher = tougher defense).

    LIMITATION: this is a role-level proxy, not a true charted man-coverage
    grade. It answers "how tough is this defense against a team's WR1-role
    receiver", not "how good is this specific cornerback" — nflverse's public
    play-by-play has no defender-assignment charting. Updates weekly whenever
    fresh PBP data is pulled.
    """
    print("  Building nflverse coverage-by-role proxy (WR1/WR2/TE)...")
    if pbp is None or pbp.empty or not starters:
        print("    ⚠ No PBP data or starters — skipping coverage proxy")
        return {}

    reg = pbp[pbp["season_type"] == "REG"].copy() if "season_type" in pbp.columns else pbp.copy()
    needed = ["defteam", "posteam", "receiver_player_name", "complete_pass", "epa", "pass_attempt"]
    missing = [c for c in needed if c not in reg.columns]
    if missing:
        print(f"    ⚠ PBP missing columns {missing} — skipping coverage proxy")
        return {}

    pass_plays = reg[(reg["pass_attempt"] == 1) & reg["receiver_player_name"].notna()].copy()
    if pass_plays.empty:
        print("    ⚠ No pass plays found — skipping coverage proxy")
        return {}
    pass_plays["_rname_norm"] = pass_plays["receiver_player_name"].apply(receiver_key)

    # Build (offense team, receiver_key) -> role lookup
    role_lookup = {}
    for team, roles in starters.items():
        for role in ("wr1", "wr2", "te"):
            nm = roles.get(role)
            if nm:
                role_lookup[(team, receiver_key(nm))] = role

    def _lookup_role(row):
        return role_lookup.get((str(row["posteam"]).upper(), row["_rname_norm"]))

    pass_plays["_role"] = pass_plays.apply(_lookup_role, axis=1)
    role_plays = pass_plays[pass_plays["_role"].notna()].copy()

    if role_plays.empty:
        print("    ⚠ No role-matched targets found — skipping coverage proxy")
        return {}

    grp = role_plays.groupby(["defteam", "_role"]).agg(
        targets=("pass_attempt", "sum"),
        completions=("complete_pass", "sum"),
        epa_allowed=("epa", "mean"),
    ).reset_index()

    out = {}
    MIN_SAMPLE = 8  # minimum targets before trusting the number (early-season noise guard)
    for role in ("wr1", "wr2", "te"):
        sub = grp[grp["_role"] == role].copy()
        sub = sub[sub["targets"] >= MIN_SAMPLE]
        if sub.empty:
            continue
        # Percentile rank on EPA allowed, inverted so low EPA allowed = high grade
        sub["_pct"] = sub["epa_allowed"].rank(pct=True, ascending=False)
        sub["_grade"] = (40 + sub["_pct"] * 55).round(1)  # scale ~40-95, matches cbGrade spread
        for _, row in sub.iterrows():
            team = row["defteam"]
            out.setdefault(team, {})
            out[team][f"covGradeVs{role.upper()}"] = float(row["_grade"])
            out[team][f"covSampleVs{role.upper()}"] = int(row["targets"])

    print(f"    ✅ Coverage proxy computed for {len(out)} defenses")
    return out


def build_redzone_defense_proxy(pbp):
    """
    NFLVERSE PBP-DERIVED RED-ZONE DEFENSE PROXY.

    Same methodology/rationale as build_coverage_proxy() — a free, TOS-clean,
    weekly-refreshing alternative to a static red-zone-defense rating (the
    existing static `rzDef` field in the model is hardcoded and never
    refreshes). Powers "VS OPP RZ DEF" on the Receiving TDs slide.

    For each unique red-zone possession (any drive that had at least one play
    inside the opponent's 20-yard line — deduped by game_id + drive), records
    whether that drive ended in a touchdown (via nflverse's fixed_drive_result
    column) and which team was on defense. Aggregates a TD-rate-allowed per
    defense, then converts to a 0-100 grade via percentile rank across all 32
    teams (lower TD rate allowed = tougher defense = higher grade), matching
    the direction and rough scale of covGradeVs* from build_coverage_proxy.
    """
    print("  Building nflverse red-zone-defense proxy...")
    if pbp is None or pbp.empty:
        print("    ⚠ No PBP data — skipping red-zone-defense proxy")
        return {}

    reg = pbp[pbp["season_type"] == "REG"].copy() if "season_type" in pbp.columns else pbp.copy()
    needed = ["yardline_100", "drive", "game_id", "defteam", "fixed_drive_result"]
    missing = [c for c in needed if c not in reg.columns]
    if missing:
        print(f"    ⚠ PBP missing columns {missing} — skipping red-zone-defense proxy")
        return {}

    rz = reg[reg["yardline_100"] <= 20].dropna(subset=["drive", "game_id", "defteam"]).copy()
    if rz.empty:
        print("    ⚠ No red-zone plays found — skipping red-zone-defense proxy")
        return {}

    # One row per unique red-zone possession (a drive counts once even if it
    # had multiple plays inside the 20)
    possessions = rz.drop_duplicates(subset=["game_id", "drive"])[
        ["game_id", "drive", "defteam", "fixed_drive_result"]
    ].copy()
    possessions["is_td"] = (possessions["fixed_drive_result"] == "Touchdown").astype(int)

    grp = possessions.groupby("defteam").agg(
        rz_possessions=("is_td", "count"),
        rz_tds_allowed=("is_td", "sum"),
    ).reset_index()
    grp["rz_td_rate"] = grp["rz_tds_allowed"] / grp["rz_possessions"]

    MIN_SAMPLE = 15  # minimum red-zone possessions faced before trusting the number
    grp = grp[grp["rz_possessions"] >= MIN_SAMPLE].copy()
    if grp.empty:
        print("    ⚠ No defense met the minimum red-zone sample size — skipping")
        return {}

    grp["_pct"] = grp["rz_td_rate"].rank(pct=True, ascending=False)  # low TD rate -> high grade
    grp["_grade"] = (40 + grp["_pct"] * 55).round(1)

    out = {}
    for _, row in grp.iterrows():
        out[row["defteam"]] = {
            "rzDefGrade": float(row["_grade"]),
            "rzDefSample": int(row["rz_possessions"]),
            "rzDefTdRate": round(float(row["rz_td_rate"]), 3),
        }

    print(f"    ✅ Red-zone-defense proxy computed for {len(out)} defenses")
    return out


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
# ══════════════════════════════════════════════════════════════════════════
# v4.4 NEW — CURRENT-SEASON (2026) USAGE INGESTION, DEPTH PRIORS, HARDENED
# INJURY-STATE DATA. Everything below joins nflverse data by stable player id
# (gsis_id) first, then by FULL normalized name (suffix-insensitive), and
# NEVER by last name alone — the last-name-only matching this replaces was the
# root of several silent mis-attributions found while building this.
# ══════════════════════════════════════════════════════════════════════════
_NAME_SUFFIXES = ('jr', 'sr', 'ii', 'iii', 'iv', 'v')
_DEPTH_TEAM_FIX = {'LA': 'LAR', 'JAC': 'JAX', 'KCC': 'KC', 'SFO': 'SF', 'NWE': 'NE',
                   'NOR': 'NO', 'GNB': 'GB', 'TBB': 'TB', 'SDG': 'LAC', 'STL': 'LAR',
                   'LVR': 'LV', 'OAK': 'LV', 'WSH': 'WAS'}


def _fix_team(t):
    t = str(t).upper().strip()
    return _DEPTH_TEAM_FIX.get(t, t)


def strip_suffix_norm(nk):
    """'travis etienne jr' -> 'travis etienne' (normalized name, generational
    suffix removed). Only ever strips a TRAILING suffix and keeps >= 2 tokens."""
    parts = str(nk).split()
    while len(parts) > 2 and parts[-1] in _NAME_SUFFIXES:
        parts.pop()
    return ' '.join(parts)


class PlayerResolver:
    """Maps (gsis_id | display name [+ team/pos]) -> the key used in `players`.
    Order: 1) gsis id via the roster id map, 2) exact normalized full name,
    3) suffix-insensitive full name — if that is ambiguous, disambiguate by
    team then position, and if STILL ambiguous return None (never guess).
    Deliberately has no last-name-only or substring fallback."""

    def __init__(self, players, id_map=None):
        self.players = players
        self.by_norm = {}
        self.by_stripped = {}
        for k, p in players.items():
            nm = norm_name(p.get('name') or k)
            self.by_norm.setdefault(nm, k)
            self.by_norm.setdefault(k, k)
            self.by_stripped.setdefault(strip_suffix_norm(nm), set()).add(k)
            self.by_stripped.setdefault(strip_suffix_norm(k), set()).add(k)
        self.by_id = {}
        for gid, full in (id_map or {}).items():
            k = self._lookup(norm_name(full))
            if k:
                self.by_id[gid] = k

    def _lookup(self, nk, team=None, pos=None):
        if nk in self.by_norm:
            return self.by_norm[nk]
        cands = self.by_stripped.get(strip_suffix_norm(nk), set())
        if len(cands) == 1:
            return next(iter(cands))
        if len(cands) > 1:
            if team:
                t = _fix_team(team)
                narrowed = [c for c in cands if _fix_team(self.players[c].get('team', '')) == t
                            or _fix_team(self.players[c].get('team2026', '')) == t]
                if len(narrowed) == 1:
                    return narrowed[0]
                cands = set(narrowed) or cands
            if pos:
                narrowed = [c for c in cands if str(self.players[c].get('pos', '')).upper() == str(pos).upper()]
                if len(narrowed) == 1:
                    return narrowed[0]
        return None

    def resolve(self, gsis_id=None, name=None, team=None, pos=None):
        if isinstance(gsis_id, str) and gsis_id.strip() in self.by_id:
            return self.by_id[gsis_id.strip()]
        if isinstance(name, str) and name.strip():
            return self._lookup(norm_name(name), team, pos)
        return None

    def resolve_norm(self, nk, team=None, pos=None):
        return self._lookup(nk, team, pos)


# ── Depth-chart context ────────────────────────────────────────────────────
def build_depth_context(depth_df):
    """v4.4 NEW — {TEAM: {'QB'|'RB'|'WR'|'TE': [ {rank, name, gsis}, ... ]}} built
    from ONLY each team's most recent depth-chart snapshot (the nflverse file
    is a daily history — 202 snapshots since March — so anything that scans
    the whole file sees stale offseason charts). Rank is the 1-based position
    in the chart's own pos_rank ordering with a player who appears in several
    formation slots counted once. Returns None if the frame isn't in the
    daily-snapshot format (caller falls back to the legacy extractor)."""
    if depth_df is None or getattr(depth_df, 'empty', True):
        return None
    need = {'dt', 'team', 'player_name', 'pos_abb', 'pos_rank'}
    if not need.issubset(depth_df.columns):
        return None
    df = depth_df[depth_df['pos_abb'].isin(['QB', 'RB', 'WR', 'TE'])].copy()
    if df.empty:
        return None
    df = df[df['dt'] == df.groupby('team')['dt'].transform('max')]
    ctx = {}
    for team, g in df.groupby('team'):
        tm = _fix_team(team)
        for pos in ('QB', 'RB', 'WR', 'TE'):
            gp = g[g['pos_abb'] == pos].sort_values('pos_rank')
            seen, lst = set(), []
            for _, r in gp.iterrows():
                gid = r.get('gsis_id') if isinstance(r.get('gsis_id'), str) else None
                ident = gid or str(r['player_name'])
                if ident in seen:
                    continue
                seen.add(ident)
                lst.append({'rank': len(lst) + 1, 'name': str(r['player_name']), 'gsis': gid})
            if lst:
                ctx.setdefault(tm, {})[pos] = lst
    return ctx


def _extract_starters_legacy(dfs):
    """The pre-v4.4 extractor, kept verbatim as a fallback for depth-chart
    files that are NOT in the daily-snapshot format."""
    return _extract_starters_impl(dfs)


def extract_starters(dfs):
    """v4.4 REAL BUG FIXED — this used to filter pos_rank == 1 across EVERY
    snapshot in the depth-chart file and take names in file order, i.e. it
    effectively read the oldest (March, pre-draft/free-agency) charts. Measured
    against the real 2026 file: it disagreed with the current depth chart on
    most teams (e.g. PHI WR2 'A.J. Brown', NO WR3 'Jordyn Tyson' where the
    current charts have Dontayvion Wicks / Bryce Lance). These names anchor
    TEAMS.wr1/wr2/wr3/te/rbTop AND the injury-state scenarios, so stale names
    meant scenarios keyed on the wrong players. Now uses each team's latest
    snapshot and its own rank ordering."""
    ctx = build_depth_context(dfs.get('depth', pd.DataFrame()))
    if not ctx:
        return _extract_starters_legacy(dfs)
    starters = {}
    for tm, pos_map in ctx.items():
        s = {}
        if pos_map.get('QB'):
            s['qb'] = pos_map['QB'][0]['name']
        if pos_map.get('RB'):
            s['rbTop'] = pos_map['RB'][0]['name']
        if pos_map.get('TE'):
            s['te'] = pos_map['TE'][0]['name']
        for i, w in enumerate(pos_map.get('WR', [])[:3]):
            s[f'wr{i + 1}'] = w['name']
        if s:
            starters[tm] = s
    n_wr3 = sum(1 for v in starters.values() if v.get('wr3'))
    print(f"  Extracted starters for {len(starters)} teams from the LATEST depth-chart "
          f"snapshot (wr3 found for {n_wr3} teams)")
    return starters

# ── Cached nflverse frame fetchers (weekly stats / snaps / PBP) ────────────
_FRAME_CACHE = {}


def _download_frame(urls, label, compression=None, timeout=120, usecols=None):
    """Try each URL in order; return a DataFrame ('' -> empty on total failure)."""
    from io import BytesIO
    for url in urls:
        try:
            r = SESSION.get(url, timeout=timeout, allow_redirects=True)
            r.raise_for_status()
            df = pd.read_csv(BytesIO(r.content), compression=compression,
                             low_memory=False, usecols=usecols)
            if not df.empty:
                print(f"  {label}: {len(df):,} rows ✅")
                return df
        except Exception as e:
            last = e
            continue
    print(f"  {label}: not available")
    return pd.DataFrame()


def fetch_weekly_stats(season):
    """v4.4 NEW — REG-season weekly player stats for `season`. Local copy first
    (same file names the older loaders used), then nflverse. v4.4 REAL BUG
    FIXED: the network path used here previously was
    `player_stats/stats_player_week_YYYY.csv`, which 404s for EVERY season —
    nflverse publishes these under the `stats_player` release tag. The 404 was
    swallowed as 'no data published yet', so the season-transition blend, the
    blend optimizer and the healthy-roster-share builder all silently no-op'd
    even though weeks 1-3 of 2026 were already public."""
    key = ('weekly', season)
    if key in _FRAME_CACHE:
        return _FRAME_CACHE[key]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    names = [f'stats_player_week_{season}.csv']
    if season == BASELINE_SEASON:
        names.append('nfl_player_stats_weekly.csv')
    df = pd.DataFrame()
    for nm in names:
        path = os.path.join(script_dir, nm)
        if os.path.exists(path):
            try:
                df = pd.read_csv(path, low_memory=False)
                if not df.empty:
                    print(f"  Weekly stats {season}: loaded locally from {nm} ({len(df):,} rows)")
                    break
            except Exception as e:
                print(f"  Local weekly stats {nm} failed to load ({e})")
    if df.empty:
        df = _download_frame([
            f"{NFLVERSE_BASE}/stats_player/stats_player_week_{season}.csv",
            f"{NFLVERSE_BASE}/player_stats/stats_player_week_{season}.csv",  # legacy path
        ], f"Weekly stats {season}")
    if not df.empty:
        if 'season' in df.columns:
            df = df[df['season'] == season]
        if 'season_type' in df.columns:
            df = df[df['season_type'] == 'REG']
        if 'team' not in df.columns and 'recent_team' in df.columns:
            df = df.rename(columns={'recent_team': 'team'})
        df = df.copy()
    _FRAME_CACHE[key] = df
    return df


def fetch_snap_frame(season):
    key = ('snaps', season)
    if key not in _FRAME_CACHE:
        df = _download_frame([f"{NFLVERSE_BASE}/snap_counts/snap_counts_{season}.csv"], f"Snap counts {season}")
        if not df.empty and 'game_type' in df.columns:
            df = df[df['game_type'] == 'REG']
        _FRAME_CACHE[key] = df
    return _FRAME_CACHE[key]


_PBP_USECOLS = ['season_type', 'posteam', 'rusher_player_id', 'rusher_player_name', 'rush_attempt',
                'rushing_yards', 'yardline_100', 'pass_attempt', 'passer_player_name',
                'receiver_player_id', 'receiver_player_name']


def fetch_pbp_frame(season):
    key = ('pbp', season)
    if key not in _FRAME_CACHE:
        _FRAME_CACHE[key] = _download_frame(
            [f"{NFLVERSE_BASE}/pbp/play_by_play_{season}.csv.gz"], f"PBP {season}",
            compression='gzip', usecols=lambda c: c in _PBP_USECOLS, timeout=180)
    return _FRAME_CACHE[key]


# ── Depth-level priors ─────────────────────────────────────────────────────
PRIOR_FIELDS = {
    'WR': ['targShare', 'airYdShare', 'wopr', 'aDOT', 'catchRate', 'yacPerRec', 'snapShare',
           'rzTargetShare', 'recYdsPG', 'recsPG'],
    'TE': ['targShare', 'airYdShare', 'wopr', 'aDOT', 'catchRate', 'yacPerRec', 'snapShare',
           'rzTargetShare', 'recYdsPG', 'recsPG'],
    'RB': ['carryShare', 'rzCarryShare', 'rzTargetShare', 'targShare', 'breakawayRate', 'ypc',
           'catchRate', 'snapShare', 'rushYdsPG', 'recYdsPG', 'recsPG'],
}
# fields where "not recorded" genuinely means zero (a real player with no red-zone
# targets has rzTargetShare 0, not "unknown"), so a missing value counts as 0 in medians
_ZERO_OK = {'rzTargetShare', 'rzCarryShare', 'breakawayRate', 'carryShare'}
_PRIOR_MIN_N = 5


def depth_bucket(pos, rank):
    """WR: 1/2/3/4+ ; TE: 1/2+ ; RB: 1/2/3+ (unknown rank -> the lowest bucket)."""
    r = rank if isinstance(rank, int) and rank > 0 else 99
    if pos == 'WR':
        return str(r) if r in (1, 2, 3) else '4+'
    if pos == 'TE':
        return '1' if r == 1 else '2+'
    return str(r) if r in (1, 2) else '3+'


def _is_real_usage(p):
    """A player counts as having REAL prior usage if they came out of the season
    stats (not a roster-only skeleton) with a real sample."""
    return p.get('source') != 'roster_2026_skeleton' and (p.get('games') or 0) >= 2


def assign_depth_ranks(players, depth_ctx, resolver):
    """Stamp depthRank/depthPos onto players that appear on the current depth chart."""
    n = 0
    for tm, pos_map in (depth_ctx or {}).items():
        for pos in ('WR', 'TE', 'RB'):
            for ent in pos_map.get(pos, []):
                k = resolver.resolve(gsis_id=ent.get('gsis'), name=ent['name'], team=tm, pos=pos)
                if k:
                    players[k]['depthRank'] = ent['rank']
                    players[k]['depthPos'] = pos
                    players[k]['team2026'] = tm
                    n += 1
    return n


def compute_depth_priors(players):
    """Median usage by (position, depth bucket) over players with REAL data.
    Buckets thinner than _PRIOR_MIN_N fall back to the position's lowest bucket."""
    import statistics
    vals = {}
    for k, p in players.items():
        pos = p.get('pos')
        if pos not in PRIOR_FIELDS or not _is_real_usage(p):
            continue
        b = depth_bucket(pos, p.get('depthRank'))
        for f in PRIOR_FIELDS[pos]:
            v = p.get(f)
            if v is None and f in _ZERO_OK:
                v = 0.0
            if v is None:
                continue
            try:
                v = float(v)
            except (TypeError, ValueError):
                continue
            if math.isnan(v) or (v <= 0 and f not in _ZERO_OK):
                continue
            vals.setdefault((pos, b, f), []).append(v)
    priors = {}
    for (pos, b, f), lst in vals.items():
        if len(lst) >= _PRIOR_MIN_N:
            priors.setdefault(pos, {}).setdefault(b, {})[f] = round(statistics.median(lst), 4)
            priors[pos][b].setdefault('n', len(lst))
    low = {'WR': '4+', 'TE': '2+', 'RB': '3+'}
    for pos, buckets in priors.items():
        base = buckets.get(low[pos], {})
        for b, fields in buckets.items():
            for f, v in base.items():
                fields.setdefault(f, v)
    return priors


def apply_depth_priors(players, priors):
    """Fill fields for players with NO real usage from their depth bucket, and
    stamp usageSource so downstream code can tell an estimate from real data.
    Players WITH real usage are never touched here."""
    filled = 0
    for k, p in players.items():
        pos = p.get('pos')
        if pos not in PRIOR_FIELDS:
            continue
        if _is_real_usage(p):
            p.setdefault('usageSource', 'baseline')
            continue
        b = depth_bucket(pos, p.get('depthRank'))
        pr = (priors.get(pos, {}).get(b)) or {}
        if not pr:
            continue
        for f in PRIOR_FIELDS[pos]:
            if p.get(f) is None and f in pr:
                p[f] = pr[f]
        p['usageSource'] = 'depth_prior'
        p['usageNote'] = f"no prior-year usage — {pos} depth-{b} median"
        filled += 1
    return filled

# ── Real 2026 usage from weekly stats + snap counts ────────────────────────
def compute_current_usage(weekly, snaps, resolver):
    """v4.4 NEW — per-player CURRENT-season usage from nflverse weekly stats and
    snap counts, keyed by `players` key. 'Games' = weeks the player actually
    played (snap-count offense_snaps > 0, or a weekly stat row with any
    target/carry/reception), and per-game shares are averaged over THOSE games,
    with a played-but-no-stat week counting as a genuine 0 — so a player who
    missed a game isn't penalised, and one who played without a target isn't
    inflated. No minimum-games threshold: even a single game is included (its
    influence is set by the blend weight, not by excluding it)."""
    usage = {}
    if weekly is None or weekly.empty:
        return usage
    num = ['targets', 'receptions', 'receiving_yards', 'receiving_air_yards',
           'receiving_yards_after_catch', 'carries', 'rushing_yards',
           'target_share', 'air_yards_share', 'wopr']
    w = weekly.copy()
    for c in num:
        if c not in w.columns:
            w[c] = 0.0
        w[c] = pd.to_numeric(w[c], errors='coerce').fillna(0.0)
    team_carries = w.groupby(['team', 'week'])['carries'].sum()

    # snap-count evidence of who actually played, per resolved player
    snap_weeks, snap_pct = {}, {}
    if snaps is not None and not snaps.empty and 'offense_snaps' in snaps.columns:
        sn = snaps[snaps['position'].isin(['WR', 'TE', 'RB', 'FB'])]
        seen = {}
        for (nm, tm, ps), grp in sn.groupby(['player', 'team', 'position']):
            k = resolver.resolve(name=nm, team=tm, pos='RB' if ps == 'FB' else ps)
            if not k:
                continue
            played = grp[grp['offense_snaps'] > 0]
            snap_weeks.setdefault(k, set()).update(int(x) for x in played['week'])
            for wk, pct in zip(played['week'], played['offense_pct']):
                snap_pct.setdefault(k, {})[int(wk)] = float(pct)

    sk = w[w['position'].isin(['WR', 'TE', 'RB'])]
    for pid, grp in sk.groupby('player_id'):
        first = grp.iloc[-1]
        k = resolver.resolve(gsis_id=pid, name=first.get('player_display_name'),
                             team=first.get('team'), pos=first.get('position'))
        if not k:
            continue
        stat_weeks = set(int(x) for x in grp[(grp['targets'] + grp['carries'] + grp['receptions']) > 0]['week'])
        weeks = sorted(snap_weeks.get(k, set()) | stat_weeks)
        g = len(weeks)
        if g == 0:
            continue
        gw = grp[grp['week'].isin(weeks)]
        car_share = 0.0
        for _, r in gw.iterrows():
            tcar = team_carries.get((r['team'], r['week']), 0.0)
            car_share += (r['carries'] / tcar) if tcar > 0 else 0.0
        tg, rc, cr = gw['targets'].sum(), gw['receptions'].sum(), gw['carries'].sum()
        u = {
            'games': g, 'team': _fix_team(first.get('team')),
            'targets': int(tg), 'receptions': int(rc), 'carries': int(cr),
            'recYds': round(float(gw['receiving_yards'].sum()), 1),
            'rushYds': round(float(gw['rushing_yards'].sum()), 1),
            'targShare': float(gw['target_share'].sum()) / g,
            'airYdShare': float(gw['air_yards_share'].sum()) / g,
            'wopr': float(gw['wopr'].sum()) / g,
            'carryShare': car_share / g,
            'recYdsPG': float(gw['receiving_yards'].sum()) / g,
            'recsPG': float(rc) / g,
            'rushYdsPG': float(gw['rushing_yards'].sum()) / g,
        }
        if tg > 0:
            u['aDOT'] = float(gw['receiving_air_yards'].sum()) / tg
            u['catchRate'] = float(rc) / tg
        if rc > 0:
            u['yacPerRec'] = float(gw['receiving_yards_after_catch'].sum()) / rc
        if cr > 0:
            u['ypc'] = float(gw['rushing_yards'].sum()) / cr
        sp = snap_pct.get(k)
        if sp:
            u['snapShare'] = sum(sp.values()) / len(sp)
        usage[k] = u
    return usage


# (player field, sample-size counter, N0) — N0 None => role/share field weighted by games only
_BLEND_FIELDS = [
    ('targShare', None, 0), ('airYdShare', None, 0), ('wopr', None, 0),
    ('carryShare', None, 0), ('snapShare', None, 0),
    ('aDOT', 'targets', 25), ('catchRate', 'targets', 25), ('yacPerRec', 'receptions', 25),
    ('ypc', 'carries', 40), ('breakawayRate', 'pbp_carries', 40),
    ('rzTargetShare', 'rz_tgts', 8), ('rzCarryShare', 'rz_carries', 8),
]
_PG_FIELDS = [('recYdsPG', 'recYdsPG'), ('recsPG', 'recsPG'), ('rushYdsPG', 'rushYdsPG')]
_POSITIVE_ONLY = {'aDOT', 'catchRate', 'yacPerRec', 'ypc'}   # a stored 0.0 here means "missing"
# Which positions each field is meaningful for. Measured on the real 2026 data:
# RBs' receiving air yards are NEGATIVE on average (screens/check-downs behind the
# line), so a raw RB aDOT (-0.7, -1.7) would push receiving projections negative;
# and a WR's carry share is noise. Blending is limited to where the metric means
# something, and a non-positive 2026 aDOT/YAC is treated as unusable, not blended.
_FIELD_POS = {
    'targShare': {'WR', 'TE', 'RB'}, 'airYdShare': {'WR', 'TE'}, 'wopr': {'WR', 'TE'},
    'carryShare': {'RB'}, 'snapShare': {'WR', 'TE', 'RB'},
    'aDOT': {'WR', 'TE'}, 'catchRate': {'WR', 'TE', 'RB'}, 'yacPerRec': {'WR', 'TE', 'RB'},
    'ypc': {'RB'}, 'breakawayRate': {'RB'},
    'rzTargetShare': {'WR', 'TE', 'RB'}, 'rzCarryShare': {'RB'},
}
_MUST_BE_POSITIVE_26 = {'aDOT', 'yacPerRec', 'ypc'}


def apply_current_usage_blend(players, usage, priors, k, k_role=None):
    """v4.4 NEW — blend REAL current-season usage into every player it exists for.
    Weight = get_season_blend_weight(games, k) (the same tunable/optimized k the
    per-game blend already uses) for ROLE fields (target/carry/snap share); for
    EFFICIENCY fields (aDOT, catch rate, YAC, ypc, breakaway, red-zone share)
    the weight is additionally capped by sample size n/(n+N0), so 3 targets at a
    100% catch rate can't drag a profile around. The prior it blends toward is
    the player's own prior-year value when they have one, else their depth-
    bucket median — so a rookie/depth player with no history is anchored to
    what a player at that depth typically does, then moves toward real 2026
    usage as it accrues. The raw 2026 numbers are kept in `usage2026` and the
    pre-blend values in `usageBaseline` so real usage stays visible."""
    blended_n = 0
    for key, u in usage.items():
        p = players.get(key)
        if not p or p.get('pos') not in PRIOR_FIELDS:
            continue
        g = u['games']
        w_pg = get_season_blend_weight(g, k)                      # per-game yardage/count fields
        w_g = get_season_blend_weight(g, k_role or k)             # role-usage fields (USAGE_ROLE_K)
        if w_g <= 0:
            continue
        pos = p['pos']
        bucket = depth_bucket(pos, p.get('depthRank'))
        pr = priors.get(pos, {}).get(bucket, {})
        had_baseline = _is_real_usage(p)
        baseline_used = {}
        for field, ncnt, n0 in _BLEND_FIELDS:
            r26 = u.get(field)
            if r26 is None or pos not in _FIELD_POS.get(field, ()):
                continue
            if field in _MUST_BE_POSITIVE_26 and r26 <= 0:
                continue
            n = None
            if ncnt is not None:
                n = u.get(ncnt)
                if n is None or n <= 0:
                    continue
            w = w_g if n0 == 0 else min(w_g, n / float(n + n0))
            base = p.get(field)
            if base is None or (field in _POSITIVE_ONLY and base <= 0):
                base = pr.get(field)
            if base is None:
                new = r26
            else:
                new = w * r26 + (1.0 - w) * base
                baseline_used[field] = round(float(base), 4)
            p[field] = round(float(new), 4)
        # per-game yardage/receptions (the recs prop reads recsPG directly)
        for pf, uf in _PG_FIELDS:
            cur = u.get(uf)
            if cur is None or p.get(f'{pf}_seasonBlended') is not None:
                continue   # the existing blend already handled this player/stat
            base = p.get(pf)
            if base is None:
                base = pr.get(pf)
            if base is None:
                base = 0.0
            p[f'{pf}_seasonBlended'] = round(w_pg * cur + (1.0 - w_pg) * base, 1)
        p['games2026'] = g
        p['blendWeight2026'] = round(w_g, 3)
        p['blendWeight2026PG'] = round(w_pg, 3)
        if not p.get('team2026'):
            p['team2026'] = u['team']
        p['usage2026'] = {kx: (round(vx, 4) if isinstance(vx, float) else vx)
                          for kx, vx in u.items() if kx != 'team'}
        if baseline_used:
            p['usageBaseline'] = baseline_used
        p['usageSource'] = 'blended_2026'
        src = f"{BASELINE_SEASON} baseline" if had_baseline else f"depth-{bucket} prior"
        p['usageNote'] = (f"{CURRENT_SEASON}: {g}g — {int(round(w_g * 100))}% {CURRENT_SEASON} + "
                          f"{int(round((1 - w_g) * 100))}% {src}")
        blended_n += 1
    return blended_n

# ══════════════════════════════════════════════════════════════════════════
# v4.4 — INJURY DATA (nflverse backup) + HARDENED HEALTHY-ROSTER SHARES
# ══════════════════════════════════════════════════════════════════════════
_INJ_LABEL = {'out': 'Out', 'doubtful': 'Doubtful', 'questionable': 'Questionable'}
_INJ_POS = {'QB', 'RB', 'WR', 'TE'}


def build_injury_status(dfs, season, resolver=None):
    """v4.4 REWRITTEN — nflverse injury data as a BACKUP to the live in-browser
    fetch. (Previously this returned {} unconditionally: the 2025 file was
    deliberately skipped and nothing read 2026, so the JSON's `injuries` was
    always empty and the model had no injury information when the live feed
    was unavailable.) Two sources, skill positions only:
      1) the most recent weekly game-status report (Out / Doubtful /
         Questionable). Only rows from the LATEST report week are used — a
         player who recovered simply stops appearing, so older rows are stale.
      2) the current roster's reserve list (status 'RES': injured reserve and
         the other unavailable reserve designations), reported as 'IR'.
    Each entry carries source='nflverse' and the report week so the model can
    prefer live data and can see how old this is."""
    print(f"\n{'='*50}\nBUILDING INJURY STATUS (nflverse backup)\n{'='*50}")
    injury_map = {}

    def _put(name, gsis, team, pos, status, detail, practice, week, origin):
        key = None
        if resolver is not None:
            key = resolver.resolve(gsis_id=gsis, name=name, team=team, pos=pos)
        key = key or norm_name(name)
        if key in injury_map and origin == 'roster':
            return   # a game-week designation is more specific than a bare roster flag
        injury_map[key] = {
            'name': name, 'team': _fix_team(team), 'pos': pos, 'status': status,
            'detail': detail or '', 'practice': practice or '', 'week': int(week or 0),
            'source': 'nflverse', 'origin': origin,
        }

    inj = fetch_nflverse_injuries_frame(CURRENT_SEASON)
    if not inj.empty and {'week', 'report_status', 'full_name'}.issubset(inj.columns):
        if 'game_type' in inj.columns:
            inj = inj[inj['game_type'] == 'REG']
        if not inj.empty:
            latest = int(inj['week'].max())
            cur = inj[(inj['week'] == latest) & inj['report_status'].notna()
                      & inj['position'].isin(_INJ_POS)]
            for _, r in cur.iterrows():
                label = _INJ_LABEL.get(str(r['report_status']).strip().lower())
                if not label:
                    continue
                _put(r['full_name'], r.get('gsis_id'), r.get('team'), r.get('position'), label,
                     r.get('report_primary_injury'), r.get('practice_status'), latest, 'report')
            print(f"  Game-status report week {latest}: {len(injury_map)} skill-position designations")

    roster = dfs.get('roster_curr', pd.DataFrame())
    if roster is not None and not roster.empty and 'status' in roster.columns:
        res = roster[(roster['status'] == 'RES') & roster['position'].isin(_INJ_POS)]
        before = len(injury_map)
        for _, r in res.iterrows():
            code = r.get('status_description_abbr')
            _put(r['full_name'], r.get('gsis_id'), r.get('team'), r.get('position'), 'IR',
                 f"Reserve list ({code})" if isinstance(code, str) else 'Reserve list',
                 '', roster['week'].max() if 'week' in roster.columns else 0, 'roster')
        print(f"  Reserve list (IR) skill players added: {len(injury_map) - before}")
    print(f"  Total nflverse injury entries: {len(injury_map)}")
    return injury_map


def fetch_nflverse_injuries_frame(season):
    key = ('injuries', season)
    if key not in _FRAME_CACHE:
        _FRAME_CACHE[key] = _download_frame([f"{NFLVERSE_BASE}/injuries/injuries_{season}.csv"],
                                            f"Injury reports {season}")
    return _FRAME_CACHE[key]


def build_healthy_roster_shares(dfs, teams, depth_ctx=None, resolver=None):
    """v4.4 HARDENED REWRITE. Per-team target-share distributions across roster
    states (full strength / WR1 out / WR2 out / WR3 out / TE out), built from
    BOTH the baseline and the current season (current-season games count 2x).
    Fixes, each of which was a real defect in the version this replaces:
      * matched players with `str.contains(lastName)` — 'Johnson' matched every
        Johnson — now id/full-name resolution only (PlayerResolver);
      * share keys were 'F.LastToken' ('Michael Pittman Jr.' -> 'M.Jr.', and two
        J.Williams overwrote each other) — now the normalized, suffix-stripped
        FULL name;
      * 'out' = ANY game without a stat row, which also counted games before a
        player joined the team (trades/rookies) and games where a player played
        but had no stats — now presence uses snap counts + stat rows, and a
        player counts as OUT only for games inside their own tenure on the team
        (between their first and last appearance), so a mid-year arrival isn't
        'out' for the games before he got there;
      * anchors came from the stale-snapshot extract_starters — now the latest
        depth chart;
      * fixed 'rookie default' shares silently overrode real data — removed;
        players without history simply fall back to their own blended share.
    Returns {} (with a printed reason) if the inputs aren't available."""
    print(f"\n{'='*50}\nBUILDING HEALTHY ROSTER SHARES (hardened)\n{'='*50}")
    if not depth_ctx or resolver is None:
        print("  ⚠ no depth context/resolver — skipping")
        return {}
    seasons = [(BASELINE_SEASON, 1.0)]
    if CURRENT_SEASON != BASELINE_SEASON:
        seasons.append((CURRENT_SEASON, 2.0))
    frames, snap_frames = [], []
    for s, wt in seasons:
        wdf = fetch_weekly_stats(s)
        if wdf is None or wdf.empty or 'game_id' not in wdf.columns:
            continue
        wdf = wdf.copy()
        wdf['season_w'] = wt
        wdf['team'] = wdf['team'].map(_fix_team)
        frames.append(wdf)
        sn = fetch_snap_frame(s)
        if sn is not None and not sn.empty:
            sn = sn.copy()
            sn['team'] = sn['team'].map(_fix_team)
            snap_frames.append(sn)
    if not frames:
        print("  ⚠ no weekly player stats with game_id available — skipping")
        return {}
    wk = pd.concat(frames, ignore_index=True)
    for c in ('targets', 'carries', 'receptions'):
        wk[c] = pd.to_numeric(wk.get(c, 0), errors='coerce').fillna(0)

    # identity: resolved player key for every weekly row / snap row
    pid_key = {}
    for pid, grp in wk.groupby('player_id'):
        r = grp.iloc[-1]
        pid_key[pid] = resolver.resolve(gsis_id=pid, name=r.get('player_display_name'),
                                        team=r.get('team'), pos=r.get('position'))
    wk['key'] = wk['player_id'].map(pid_key)
    wk = wk[wk['key'].notna()]
    _gl = wk[['game_id', 'season', 'week']].drop_duplicates().sort_values(['season', 'week'])
    game_order = {g: i for i, g in enumerate(_gl['game_id'])}
    season_w = {sv: wt for sv, wt in seasons}

    _act = wk[(wk['targets'] + wk['carries'] + wk['receptions']) > 0]
    # presence is keyed (game, TEAM, player): a player also appears in games he played
    # AGAINST his former team, and that must not count as being present for the new one
    present = set(zip(_act['game_id'], _act['team'], _act['key']))
    snap_only = []          # (game_id, team, key) — played per snap counts, no stat row
    have_row = set(zip(wk['game_id'], wk['team'], wk['key']))
    for sn in snap_frames:
        sn = sn[sn['position'].isin(['WR', 'TE', 'RB', 'FB']) & (sn['offense_snaps'] > 0)]
        cache = {}
        for gid, nm, tm, ps in zip(sn['game_id'], sn['player'], sn['team'], sn['position']):
            ck = (nm, tm, ps)
            if ck not in cache:
                cache[ck] = resolver.resolve(name=nm, team=tm, pos='RB' if ps == 'FB' else ps)
            k = cache[ck]
            if k and (gid, tm, k) not in present:
                present.add((gid, tm, k))
                if (gid, tm, k) not in have_row:
                    snap_only.append((gid, tm, k))
    print(f"  Presence evidence: {len(present):,} player-games ({len(snap_only):,} from snap counts alone)")

    team_tgts = wk.groupby(['game_id', 'team'])['targets'].sum()
    skill = wk[wk['position'].isin(['WR', 'TE', 'RB'])]
    disp = {}
    for k, nm in zip(skill['key'], skill['player_display_name']):
        disp.setdefault(k, nm)
    # per (game, team, key): targets (0 if present via snaps only)
    rowt = skill.groupby(['game_id', 'team', 'key']).agg(targets=('targets', 'sum'), sw=('season_w', 'max')).reset_index()
    if snap_only:   # a game a player PLAYED but had no stat row in is a genuine 0-target game
        so = pd.DataFrame(snap_only, columns=['game_id', 'team', 'key']).drop_duplicates()
        so['targets'] = 0.0
        so['sw'] = so['game_id'].map(lambda g: season_w.get(int(str(g)[:4]), 1.0))
        rowt = pd.concat([rowt, so], ignore_index=True).groupby(['game_id', 'team', 'key']).agg(
            targets=('targets', 'sum'), sw=('sw', 'max')).reset_index()

    def _skey(name_or_key):
        return strip_suffix_norm(norm_name(name_or_key))

    def share_map(game_ids, team, min_tgts=4, min_games=2):
        if not game_ids:
            return {}
        gset = set(game_ids)
        sub = rowt[(rowt['team'] == team) & rowt['game_id'].isin(gset)]
        acc = {}
        for gid, k, tg, sw in zip(sub['game_id'], sub['key'], sub['targets'], sub['sw']):
            tt = team_tgts.get((gid, team), 0)
            if tt <= 0:
                continue
            a = acc.setdefault(k, [0.0, 0.0, 0, 0])   # weighted share, weight, games, targets
            a[0] += (tg / tt) * sw; a[1] += sw; a[2] += 1; a[3] += tg
        out = {}
        for k, (ws, w, n, tg) in acc.items():
            if tg >= min_tgts and n >= min_games and w > 0:
                out[_skey(disp.get(k, k))] = round(float(ws / w), 4)
        return out

    healthy = {}
    for team, pos_map in depth_ctx.items():
        wr = pos_map.get('WR', [])
        te = pos_map.get('TE', [])
        if not wr:
            continue
        team_games = sorted(set(wk[wk['team'] == team]['game_id']), key=lambda g: game_order.get(g, 0))
        if not team_games:
            continue
        idx = {g: i for i, g in enumerate(team_games)}

        def appearances(ent):
            k = resolver.resolve(gsis_id=ent.get('gsis'), name=ent['name'], team=team)
            if not k:
                return k, set()
            return k, {g for g in team_games if (g, team, k) in present}

        anchors = {}
        for label, ent in (('wr1', wr[0] if len(wr) > 0 else None), ('wr2', wr[1] if len(wr) > 1 else None),
                           ('wr3', wr[2] if len(wr) > 2 else None), ('te', te[0] if te else None)):
            if ent:
                k, appear = appearances(ent)
                anchors[label] = {'name': ent['name'], 'key': k, 'appear': appear}
        usable = {l: a for l, a in anchors.items() if len(a['appear']) >= 3}

        def tenure_out(a):
            if len(a['appear']) < 3:
                return set()
            lo = min(idx[g] for g in a['appear']); hi = max(idx[g] for g in a['appear'])
            return {g for g in team_games if lo <= idx[g] <= hi} - a['appear']

        all_g = set(team_games)
        # FULL STRENGTH = no anchor was absent WITHIN his own tenure. A game from before
        # an anchor joined the team (rookie, trade, free agent) is simply outside his
        # window and must not disqualify the game.
        full_g = set(all_g)
        for a in usable.values():
            full_g -= tenure_out(a)
        w1 = usable.get('wr1'); w2 = usable.get('wr2'); w3 = usable.get('wr3'); t_ = usable.get('te')

        def scen(target, *others):
            if not target:
                return set()
            g = set(tenure_out(target))
            for m in others:            # the other named anchors must not have been out either
                if m:
                    g -= tenure_out(m)
            return g

        wr1_out = scen(w1, w2)
        wr2_out = scen(w2, w1)
        wr3_out = scen(w3, w1)
        te_out = scen(t_, w1)
        full_shares = share_map(full_g, team)
        if not full_shares:
            continue
        healthy[team] = {
            'anchor': anchors['wr1']['name'] if 'wr1' in anchors else None,
            'wr2': anchors['wr2']['name'] if 'wr2' in anchors else None,
            'wr3': anchors['wr3']['name'] if 'wr3' in anchors else None,
            'te': anchors['te']['name'] if 'te' in anchors else None,
            'season': CURRENT_SEASON,
            'source': f'co_game_weekly_{BASELINE_SEASON}_{CURRENT_SEASON}',
            'keyFormat': 'norm_fullname_nosuffix',
            'shares': full_shares,
            'full_strength': {'games': len(full_g), 'shares': full_shares},
            'wr1_out': {'games': len(wr1_out), 'shares': share_map(wr1_out, team, min_games=1)},
            'wr2_out': {'games': len(wr2_out), 'shares': share_map(wr2_out, team, min_games=1)},
            'wr3_out': {'games': len(wr3_out), 'shares': share_map(wr3_out, team, min_games=1)},
            'te_out':  {'games': len(te_out),  'shares': share_map(te_out,  team, min_games=1)},
        }
    print(f"  Built roster-state share maps for {len(healthy)} teams "
          f"(scenario game counts include current-season games)")
    return healthy


def save_output(teams, players, injury_map, games, college_data, season, roster_changes=None, starters=None, coaching_context=None, dfs=None, healthy_roster_shares=None, coverage_proxy=None, redzone_defense_proxy=None, depth_priors=None):
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

    # v3.39 NEW — dynamic defensive rank computation. Previously
    # passDefRank/runDefRank/wr1vsRank/teVsRank/rbVsRank were hardcoded
    # static values baked directly into the HTML's TEAMS object, frozen
    # at whatever they were when the model was first built — confirmed
    # via direct inspection that a real JSON export carried none of these
    # fields for any of the 32 teams, while real, current
    # defPassEPA/defRushEPA were already being computed here every run.
    # Fixed by ranking all 32 teams on their real, current EPA allowed
    # (lower/more negative EPA allowed = better defense = rank 1) and
    # shipping the result so the model can use live ranks instead of
    # stale hardcoded ones.
    # wr1vsRank/teVsRank reuse passDefRank as a team-level pass-defense
    # proxy — real position-specific vs-WR/vs-TE charting data isn't
    # available here (coverage_proxy already covers WR/TE matchup grades
    # separately and takes priority over these rank fields in the model);
    # rbVsRank reuses runDefRank. Coarser than true position-specific
    # matchup data, but real and current instead of frozen.
    def _rank_teams_by(stat_key, teams_dict):
        valid = [(abbr, t.get(stat_key)) for abbr, t in teams_dict.items()
                 if isinstance(t, dict) and isinstance(t.get(stat_key), (int, float))]
        valid.sort(key=lambda x: x[1])
        return {abbr: i + 1 for i, (abbr, _) in enumerate(valid)}

    _pass_def_ranks = _rank_teams_by('defPassEPA', teams)
    _run_def_ranks  = _rank_teams_by('defRushEPA', teams)
    print(f"  Computed live defensive ranks: {len(_pass_def_ranks)}/32 teams (pass), "
          f"{len(_run_def_ranks)}/32 teams (run)")

    output = {
        'generated':      datetime.now().isoformat(),
        'season':         season,
        'baseline_season':BASELINE_SEASON,
        'current_season': CURRENT_SEASON,
        'is_preseason':   CURRENT_SEASON > BASELINE_SEASON,
        'source':         'nflverse + PFR advstats + NGS + ProspectAnalyzer',
        'teams':      teams,
        'healthy_roster_shares': healthy_roster_shares or {},
        # nflverse PBP-derived pass-defense-by-role proxy (WR1/WR2/TE vs each
        # defense). Free, TOS-clean stand-in for PFF coverage grades — see
        # build_coverage_proxy() docstring for methodology and limitations.
        'coverage_proxy': coverage_proxy or {},
        # nflverse PBP-derived red-zone-defense proxy (TD rate allowed per
        # team, converted to a 0-100 grade). Powers "VS OPP RZ DEF" on the
        # Receiving TDs slide — see build_redzone_defense_proxy() docstring.
        'redzone_defense_proxy': redzone_defense_proxy or {},
        'players':    players_out,
        'injuries':   {(players.get(k, {}).get('name') or (v.get('name') if isinstance(v, dict) else None) or k): v
                       for k, v in injury_map.items()},
        # v4.4: depth-level usage priors (position x depth bucket medians) so the model's
        # fallback for a player with no data is derived from real depth-level usage
        'depth_priors': depth_priors or {},
        'games':      games,
        'rookieData': {p.get('name',''):
                       {'collegeData': p.get('collegeData',{}),
                        'blendWeights': p.get('blendWeights',{}),
                        'rookieNote': p.get('rookieNote','')}
                       for p in players.values() if p.get('isRookie')},
        'coaching_context': coaching_context or {},
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
                'passDefRank':  _pass_def_ranks.get(abbr),
                'runDefRank':   _run_def_ranks.get(abbr),
                'wr1vsRank':    _pass_def_ranks.get(abbr),
                'teVsRank':     _pass_def_ranks.get(abbr),
                'rbVsRank':     _run_def_ranks.get(abbr),
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
            'currentSeasonPlayers': sum(1 for p in players.values() if p.get('games2026')),
            'depthPriorPlayers':    sum(1 for p in players.values() if p.get('usageSource') == 'depth_prior'),
            'injuryBackupCount':    len(injury_map),
        }
    }

    def _json_safe(o):
        # The browser's JSON.parse REJECTS NaN/Infinity, and numpy ints are not int
        # subclasses (default=str would silently turn them into strings) — clean both.
        if isinstance(o, dict):
            return {str(k): _json_safe(v) for k, v in o.items()}
        if isinstance(o, (list, tuple, set)):
            return [_json_safe(v) for v in o]
        try:
            import numpy as _np
            if isinstance(o, _np.generic):
                o = o.item()
        except ImportError:
            pass
        if isinstance(o, float) and (math.isnan(o) or math.isinf(o)):
            return None
        return o

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(_json_safe(output), f, indent=2, default=str)

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


def build_roster_changes(dfs, players, season):
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

    if roster_base.empty:
        print("  ⚠ No baseline rosters — skipping roster change analysis")
        return {}

    if roster_curr.empty:
        print("  ⚠ No current rosters — skipping roster change analysis")
        return {}

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
        return str(n).lower().strip() if n else ''

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

    # ── Build player production from 2025 season stats ────────────────
    # Key metrics per position for scoring impact
    prod = {}  # norm_name → stats dict
    if not player_stats.empty:
        name_c = get_name_col(player_stats)
        team_c = get_team_col(player_stats)
        pos_c  = get_pos_col(player_stats)
        if name_c and team_c and pos_c:
            for _, row in player_stats.iterrows():
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
                        'pos': 'QB',
                        'team': norm_abbr(row.get(team_c, '')),
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
                        'pos': 'RB',
                        'team': norm_abbr(row.get(team_c, '')),
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
                        'pos': pos,
                        'team': norm_abbr(row.get(team_c, '')),
                        'raw_name': str(row.get(name_c, '')).strip(),
                        'recYdsPG':  pg('receiving_yards'),
                        'recsPG':    pg('receptions'),
                        'targetsPG': pg('targets'),
                        'recEPA':    round(safe_float(row.get('receiving_epa', 0)) / tgt, 4),
                        'games':     games,
                    }
        print(f"  Production data: {len(prod)} skill players")

    # ── Identify movers: players whose team changed ───────────────────
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

    # ── Build 2025 starters per team per position ─────────────────────
    # "Starter" = highest production player at each position per team
    starters_2025 = {}  # team → pos → {norm_name, prod}
    pos_metric = {'QB':'passYdsPG', 'RB':'rushYdsPG', 'WR':'recYdsPG', 'TE':'recYdsPG'}
    for norm, p in prod.items():
        team = p.get('team', '')
        pos  = p.get('pos', '')
        if not team or pos not in SKILL: continue
        metric = pos_metric.get(pos, 'recYdsPG')
        val    = p.get(metric, 0)
        if team not in starters_2025:
            starters_2025[team] = {}
        if pos not in starters_2025[team] or val > starters_2025[team][pos]['val']:
            starters_2025[team][pos] = {'norm_name': norm, 'val': val, 'prod': p}

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
        """Convert player's 2025 stats to approximate pts/game contribution."""
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

    # ── Compute team-level roster adjustments ─────────────────────────
    team_adj = {}  # team → {offAdj, changes}

    for mover in movers:
        from_team = mover['from_team']
        to_team   = mover['to_team']
        pos       = mover['pos']
        name      = mover['raw_name']
        p_prod    = mover['prod']

        # Skip if no production data (camp bodies, practice squad)
        games_played = p_prod.get('games', 0) if p_prod else 0
        if games_played < 3 and pos != 'QB':
            continue  # too small a sample for non-QBs

        mover_pts = player_pts(p_prod, pos) if p_prod else 0

        # Who was the starter at this pos for each team?
        old_starter = starters_2025.get(from_team, {}).get(pos, {})
        new_starter = starters_2025.get(to_team, {}).get(pos, {})

        old_pts  = player_pts(old_starter.get('prod', {}), pos)
        new_pts  = player_pts(new_starter.get('prod', {}), pos)

        # Impact on the LOSING team (from_team): lost mover, who replaces?
        # Lost pts = mover_pts - replacement_pts
        # If we don't know who replaces, use positional average
        avg_pos_pts = {'QB': 4.2, 'RB': 2.1, 'WR': 1.8, 'TE': 1.2}
        replacement_pts = max(old_pts - mover_pts, 0) if old_pts > 0 else avg_pos_pts.get(pos, 2.0)

        loss_impact = round(replacement_pts - mover_pts, 2)  # negative = loss

        # Impact on GAINING team (to_team): added mover vs what they had
        # Gained pts = mover_pts - what_they_had
        they_had_pts = new_pts if new_starter else avg_pos_pts.get(pos, 2.0)
        gain_impact  = round(mover_pts - they_had_pts, 2)    # positive = gain

        # Record changes
        for team, impact, direction in [
            (from_team, loss_impact, 'OUT'),
            (to_team,   gain_impact, 'IN'),
        ]:
            if not team or team == 'FA': continue
            if team not in team_adj:
                team_adj[team] = {'offAdj': 0.0, 'changes': []}
            team_adj[team]['offAdj'] = round(team_adj[team]['offAdj'] + impact, 2)
            team_adj[team]['changes'].append({
                'player':     name,
                'pos':        pos,
                'direction':  direction,
                'from_team':  from_team,
                'to_team':    to_team,
                'impact_pts': impact if direction == 'IN' else -abs(loss_impact),
                'stats_2025': {
                    k: v for k, v in p_prod.items()
                    if k not in ('pos','team','raw_name','games')
                } if p_prod else {},
                'games_2025': games_played,
            })

    # ── Cap adjustments at ±8 pts (prevent extreme outliers) ──────────
    for team, data in team_adj.items():
        data['offAdj'] = round(max(-8.0, min(8.0, data['offAdj'])), 2)

    # ── Print summary ─────────────────────────────────────────────────
    significant = {t: d for t, d in team_adj.items() if abs(d['offAdj']) >= 1.0}
    print(f"\n  Teams with significant roster impact (±1+ pts):")
    for team, data in sorted(significant.items(), key=lambda x: abs(x[1]['offAdj']), reverse=True):
        sign = '+' if data['offAdj'] > 0 else ''
        changes_str = ', '.join(
            f"{c['direction']} {c['pos']} {c['player'].split()[-1]}"
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
        """Fetch player stats for a season via direct nflverse URL.
        v3.67 FIX — same rename as fetch_prior2_season_players above."""
        url = f"{NFLVERSE_BASE}/player_stats/stats_player_reg_{yr}.csv"
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

    # ══ v4.4: shared identity, baseline-season shares, depth-level priors ═══════
    # One resolver (gsis id -> full name -> suffix-insensitive name; never last-name
    # only) is used for every join below, so 2025 PBP shares, 2026 weekly stats,
    # snap counts, the depth chart and injury reports all land on the SAME profile.
    try:
        _cur_roster = fetch_roster_current(CURRENT_SEASON)
    except Exception as _cre:
        print(f"  ⚠ current-season roster unavailable ({_cre}) — id mapping uses weekly rosters only")
        _cur_roster = pd.DataFrame()
    _id_map = build_gsis_name_map(dfs.get('weekly_rosters'), dfs.get('rosters'), _cur_roster)
    resolver = PlayerResolver(players, _id_map)
    depth_ctx = build_depth_context(dfs.get('depth', pd.DataFrame()))

    # Baseline-season (2025) opportunity + snap shares. v4.4 REAL BUG FIXED — the
    # opportunity-share merge matched ~0.8% of players (PBP abbreviated names vs
    # full-name profile keys); merged here by id/full name, with the rate reported.
    _opp_shares = compute_opportunity_shares_from_pbp(dfs.get('pbp', pd.DataFrame()), _id_map)
    _opp_merged = 0
    for _nk, _vals in _opp_shares.items():
        _pk = resolver.resolve_norm(_nk)
        if _pk:
            players[_pk].update(_vals)
            _opp_merged += 1
    print(f"  Opportunity shares merged into player profiles: {_opp_merged}/{len(_opp_shares)}")
    if _opp_shares and _opp_merged < 0.5 * len(_opp_shares):
        print("  ⚠ WARNING: fewer than half of the computed opportunity shares matched a player "
              "profile — carryShare/rzTargetShare will be missing for most players. Check name/id mapping.")
    for _nk, _pct in fetch_snap_counts(BASELINE_SEASON).items():
        _pk = resolver.resolve_norm(_nk)
        if _pk:
            players[_pk]['snapShare'] = _pct

    # Depth ranks (from the LATEST depth chart) -> depth-level priors -> fill roster-only
    # players so a rookie with no history is anchored to what a player at that depth does
    # (e.g. a WR3's ~9-10% target share), not a WR1 default.
    depth_priors = {}
    if depth_ctx:
        _ranked = assign_depth_ranks(players, depth_ctx, resolver)
        depth_priors = compute_depth_priors(players)
        _filled = apply_depth_priors(players, depth_priors)
        print(f"  Depth ranks assigned to {_ranked} players; {_filled} roster-only players "
              f"given depth-level priors ({sum(len(v) for v in depth_priors.values())} position/depth buckets)")
    else:
        print("  ⚠ Depth chart not in daily-snapshot format — depth priors skipped this run")

    # v3.39 NEW — Bayesian season-transition blend. Previously the only
    # way to move from BASELINE_SEASON's full-season stats to real
    # CURRENT_SEASON data was flipping BASELINE_SEASON=CURRENT_SEASON by
    # hand and re-running — meaning Week 1 alone would instantly become
    # a player's entire baseline the moment it's played, which is
    # extremely noisy on 1 game of data. This blends BASELINE_SEASON
    # (the real, completed prior season) with whatever real CURRENT_SEASON
    # games exist so far, weighted toward CURRENT_SEASON as more of it is
    # played. Safe to leave on every run: pre-Week-1 the current-season
    # file is empty/unpublished and this is a confirmed no-op (see
    # fetch_current_season_partial_stats docstring).
    # v3.55: fetches the real BASELINE_SEASON-1 data for the new 3rd
    # blend tier. Safe to leave on every run — fetch_prior2_season_players
    # returns {} (not an error) if that season's file isn't available,
    # and apply_season_transition_blend() falls back to the original
    # 2-season behavior exactly when prior2_data is empty for a player.
    # v3.63 NEW — joint (RECENCY_DECAY_RATE, SEASON_BLEND_K) optimizer.
    # Uses BASELINE_SEASON's OWN real weekly data as ground truth (see the
    # optimizer's module docstring above for the full no-lookahead method)
    # — no external file needed.
    # v3.70 FIX — every run recomputes fresh now (the grid search is fast
    # enough that caching wasn't buying anything real, and it was the
    # direct cause of the "stale result" confusion this got debugged
    # through). blend_params.json is still written each run as a
    # human-readable record of the latest decision — just never read back.
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    _baseline_weekly = fetch_current_season_partial_stats(BASELINE_SEASON)
    prior2_players = fetch_prior2_season_players(BASELINE_SEASON - 1)
    _decay_rate, _k, _opt_result = load_or_optimize_blend_params(
        _script_dir, BASELINE_SEASON, _baseline_weekly, prior1_data=players, prior2_data=prior2_players)
    apply_season_transition_blend(players, CURRENT_SEASON, k=_k, decay_rate=_decay_rate,
                                   prior2_data=prior2_players, resolver=resolver)
    # v3.86 NEW — exports the same raw per-player/stat/checkpoint cases the
    # optimizer above already builds internally, so the JS model can cross-
    # reference them against real historical prop lines for the hit-rate
    # half of the decay-rate optimizer (this Python side only ever measures
    # projection error against the real final season average — it has no
    # access to real betting lines at all, that lives entirely in the JS
    # model's Odds API integration). Safe no-op if the optimizer didn't run
    # (e.g. no baseline data yet) — same convention as its other outputs.
    if _opt_result is not None:
        export_blend_optimizer_cases(OUTPUT_DIR, BASELINE_SEASON, _opt_result)

    # ══ v4.4: CURRENT-SEASON (2026) USAGE INGESTION ═════════════════════════════
    # Usage changes year over year (roles, new teams, injuries), so every 2026 game
    # already played is ingested — no minimum-games threshold. How much each game
    # counts is set by the blend weight (games/k, k tuned by the optimizer above),
    # not by excluding early data. Real 2026 numbers stay visible in `usage2026`.
    if CURRENT_SEASON > BASELINE_SEASON:
        weekly26 = fetch_weekly_stats(CURRENT_SEASON)
        snaps26 = fetch_snap_frame(CURRENT_SEASON)
        usage26 = compute_current_usage(weekly26, snaps26, resolver)
        pbp26 = fetch_pbp_frame(CURRENT_SEASON)
        if pbp26 is not None and not pbp26.empty and usage26:
            _cnt26 = {}
            _opp26 = compute_opportunity_shares_from_pbp(pbp26, _id_map, counts_out=_cnt26)
            for _nk, _vals in _opp26.items():
                _pk = resolver.resolve_norm(_nk)
                if _pk in usage26:
                    for _f in ('rzTargetShare', 'rzCarryShare', 'breakawayRate'):
                        if _f in _vals:
                            usage26[_pk][_f] = _vals[_f]
                    usage26[_pk].update(_cnt26.get(_nk, {}))
        _u_blended = apply_current_usage_blend(players, usage26, depth_priors, _k, k_role=USAGE_ROLE_K)
        print(f"  {CURRENT_SEASON} usage: {len(usage26)} players with games played, "
              f"{_u_blended} blended into profiles (role k={USAGE_ROLE_K}, per-game k={_k})")
    else:
        print("  (BASELINE_SEASON == CURRENT_SEASON — current-season usage blend not needed)")

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

    # Build nflverse-derived coverage-by-role proxy (WR1/WR2/TE vs each defense)
    coverage_proxy = build_coverage_proxy(dfs.get('pbp'), starters_map)

    # Build nflverse-derived red-zone-defense proxy (TD rate allowed per team)
    redzone_defense_proxy = build_redzone_defense_proxy(dfs.get('pbp'))

    # Fetch current-season roster for roster-change comparison
    dfs['roster_curr'] = _cur_roster if (_cur_roster is not None and not _cur_roster.empty) else fetch_roster_current(CURRENT_SEASON)

    # Build injury/depth chart status
    injury_map = build_injury_status(dfs, SEASON, resolver)

    # Build upcoming games list
    games = build_upcoming_games(dfs, SEASON)

    # Detect roster changes and compute scoring impact
    # Build healthy roster shares (co-game target distributions)
    healthy_roster_shares = build_healthy_roster_shares(dfs, teams, depth_ctx, resolver)

    roster_changes = build_roster_changes(dfs, players, SEASON)

    # Build coaching context (mid-season overrides + scheme flags)
    coaching_context = build_coaching_context(dfs)

    # Save everything to JSON
    save_output(teams, players, injury_map, games, college_data, SEASON,
                roster_changes, starters_map, coaching_context, dfs=dfs,
                healthy_roster_shares=healthy_roster_shares,
                coverage_proxy=coverage_proxy,
                redzone_defense_proxy=redzone_defense_proxy,
                depth_priors=depth_priors)

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
# ══════════════════════════════════════════════════════════════════════
# v3.88 NEW — ONE-TIME HISTORICAL BACKFILL: real red-zone-defense grades
# for BT_BASELINE_2023 / BT_BASELINE_2024 (2026-09-25)
# ══════════════════════════════════════════════════════════════════════
# Why this exists: BT_BASELINE_2023 and BT_BASELINE_2024 (the historical
# snapshots the HTML's real Game Backtest evaluates 2024/2025 games
# against) currently have every single team hardcoded at rzDef=57.0 —
# confirmed by direct inspection, no exceptions, either season. That
# means getTeamOffScore()'s opponent-red-zone-defense term (rzAdj, fixed
# in v3.88) evaluates to exactly zero for every historical backtest game,
# so a red-zone-defense formula variant can't be meaningfully tested
# against real 2024/2025 outcomes yet — there's no real per-team
# variation in the data those two baselines actually use.
#
# This does NOT touch the main daily pipeline. It's a standalone, opt-in
# backfill: fetches PBP for each requested past season directly from
# nflverse (same public source, same URL pattern as fetch_pbp_player_stats
# above) and runs it through the EXISTING, already-verified
# build_redzone_defense_proxy() — not a reimplementation, the literal same
# function used for the current season's live "VS OPP RZ DEF" prop stat.
#
# Output: nfl_historical_rzdef_proxy.json, shaped as
#   { "2023": { "ARI": {rzDefGrade, rzDefSample, rzDefTdRate}, ... },
#     "2024": { ... } }
# Run once: python nfl_edge_data_pull.py --backfill-rzdef
# Then hand the JSON back so it can be merged into BT_BASELINE_2023/2024.
# ══════════════════════════════════════════════════════════════════════
def backfill_historical_rzdef_proxy(seasons):
    import gzip as _gz, io as _io, json as _json

    out = {}
    for season in seasons:
        print(f"\n{'='*50}")
        print(f"BACKFILL: red-zone-defense proxy for {season}")
        print('='*50)
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
        except Exception as e:
            print(f" ❌ {e} — skipping {season}")
            continue

        proxy = build_redzone_defense_proxy(pbp)
        if not proxy:
            print(f"  ⚠ No red-zone-defense proxy computed for {season} — skipping")
            continue
        out[str(season)] = proxy
        print(f"  ✅ {season}: {len(proxy)} teams graded")

    if not out:
        print("\n❌ Backfill produced no data for any requested season — nothing written.")
        return

    out_path = os.path.join(OUTPUT_DIR, 'nfl_historical_rzdef_proxy.json')
    with open(out_path, 'w') as f:
        _json.dump(out, f, indent=2)
    print(f"\n✅ Wrote {out_path}")
    print("   Next: push this to the GitHub repo (or hand it back) so it can be merged")
    print("   into BT_BASELINE_2023/BT_BASELINE_2024 in the HTML — no formula change yet,")
    print("   this only supplies the real historical data those two years were missing.")


# ══════════════════════════════════════════════════════════════════════
# BACKFILL: pace/efficiency proxy (ptsPerDrive, successRate,
# pressureRatePerGame) for past seasons — same one-time pattern as
# backfill_historical_rzdef_proxy() above. Supplies the three real fields
# the validated market-residual O/U total formula (2026-09-26) needs but
# that BT_BASELINE_2023/2024 don't currently have.
#
# Run once: python nfl_edge_data_pull.py --backfill-pace-efficiency
# Then hand the JSON back so it can be merged into BT_BASELINE_2023/2024.
# ══════════════════════════════════════════════════════════════════════
def build_pace_efficiency_proxy(pbp, pfr_def):
    reg = pbp[pbp['season_type'] == 'REG'].copy()
    out = {}

    # ptsPerDrive (offense): real per-drive scoring via score change across
    # each drive's plays — posteam_score_post minus posteam_score summed
    # per (game, team, drive), divided by real drive count.
    reg['drive_pts'] = reg['posteam_score_post'] - reg['posteam_score']
    drive_scoring = reg.groupby(['game_id', 'posteam', 'drive'])['drive_pts'].sum().reset_index()
    drive_scoring = drive_scoring[drive_scoring['posteam'].notna()]
    totals = drive_scoring.groupby('posteam').agg(
        total_pts=('drive_pts', 'sum'), total_drives=('drive_pts', 'count'))

    # successRate (offense): real series_success column, nflverse's own
    # down/distance success definition.
    succ = reg[reg['series_success'].notna() & reg['posteam'].notna()]
    succ_by_team = succ.groupby('posteam')['series_success'].mean()

    for team in totals.index:
        out[team] = {
            'ptsPerDrive': round(float(totals.loc[team, 'total_pts'] / totals.loc[team, 'total_drives']), 3),
            'successRate': round(float(succ_by_team.get(team, float('nan'))), 4) if team in succ_by_team.index else None,
        }

    # pressureRatePerGame (defense): real PFR def_pressures, summed per
    # team+week then averaged across weeks played — a real per-game rate.
    if 'team' in pfr_def.columns and 'def_pressures' in pfr_def.columns:
        team_week = pfr_def.groupby(['team', 'week'])['def_pressures'].sum().reset_index()
        pressure_pg = team_week.groupby('team')['def_pressures'].mean()
        for team, val in pressure_pg.items():
            out.setdefault(team, {})['pressureRatePerGame'] = round(float(val), 2)

    return out


def backfill_pace_efficiency_proxy(seasons):
    import gzip as _gz, io as _io, json as _json

    out = {}
    for season in seasons:
        print(f"\n{'='*50}")
        print(f"BACKFILL: pace/efficiency proxy for {season}")
        print('='*50)
        pbp_url = f"{NFLVERSE_BASE}/pbp/play_by_play_{season}.csv.gz"
        print(f"  Downloading PBP {season} ({pbp_url.split('/')[-1]}) ...", end="", flush=True)
        try:
            r = SESSION.get(pbp_url, timeout=180, allow_redirects=True)
            r.raise_for_status()
            size_mb = len(r.content) / 1024 / 1024
            print(f" {size_mb:.1f}MB", end="", flush=True)
            with _gz.open(_io.BytesIO(r.content)) as gz:
                pbp = pd.read_csv(gz, low_memory=False)
            print(f" — {len(pbp):,} plays")
        except Exception as e:
            print(f" ❌ {e} — skipping {season}")
            continue

        pfr_url = f"{NFLVERSE_BASE}/pfr_advstats/advstats_week_def_{season}.csv"
        print(f"  Downloading PFR defensive advstats {season} ...", end="", flush=True)
        try:
            r2 = SESSION.get(pfr_url, timeout=120, allow_redirects=True)
            r2.raise_for_status()
            pfr_def = pd.read_csv(_io.BytesIO(r2.content), low_memory=False)
            print(f" — {len(pfr_def):,} player-weeks")
        except Exception as e:
            print(f" ❌ {e} — proceeding without pressure data for {season}")
            pfr_def = pd.DataFrame(columns=['team', 'week', 'def_pressures'])

        proxy = build_pace_efficiency_proxy(pbp, pfr_def)
        if not proxy:
            print(f"  ⚠ No pace/efficiency proxy computed for {season} — skipping")
            continue
        out[str(season)] = proxy
        print(f"  ✅ {season}: {len(proxy)} teams computed")

    if not out:
        print("\n❌ Backfill produced no data for any requested season — nothing written.")
        return

    out_path = os.path.join(OUTPUT_DIR, 'nfl_historical_pace_efficiency_proxy.json')
    with open(out_path, 'w') as f:
        _json.dump(out, f, indent=2)
    print(f"\n✅ Wrote {out_path}")
    print("   Next: hand this back so it can be merged into BT_BASELINE_2023/BT_BASELINE_2024")
    print("   in the HTML — supplies ptsPerDrive/successRate/pressureRatePerGame, the three")
    print("   fields the validated market-residual O/U total formula needs but doesn't have yet.")


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='NFL Edge Model data pull')
    parser.add_argument('--update',  action='store_true',
                        help='In-season update: patch existing JSON with latest NGS + recency EPA')
    parser.add_argument('--weeks',   type=int, default=4,
                        help='Recency window in weeks for --update (default: 4)')
    parser.add_argument('--season',  type=int, default=None,
                        help='Override season for --update (default: CURRENT_SEASON)')
    parser.add_argument('--backfill-rzdef', action='store_true',
                        help='One-time historical backfill: compute real per-team red-zone-'
                             'defense grades (rzDefGrade) for past seasons and export a JSON '
                             'ready to merge into BT_BASELINE_2023/BT_BASELINE_2024 in the HTML.')
    parser.add_argument('--backfill-seasons', type=str, default='2023,2024',
                        help='Comma-separated seasons for --backfill-rzdef/--backfill-pace-efficiency '
                             '(default: 2023,2024, matching BT_BASELINE_2023/BT_BASELINE_2024)')
    parser.add_argument('--backfill-pace-efficiency', action='store_true',
                        help='One-time historical backfill: compute real per-team ptsPerDrive, '
                             'successRate, and pressureRatePerGame for past seasons and export a '
                             'JSON ready to merge into BT_BASELINE_2023/BT_BASELINE_2024 in the '
                             'HTML — supplies the fields the market-residual O/U total formula needs.')
    parser.add_argument('--push', action='store_true',
                        help='After a successful pull, automatically copy the output JSON into '
                             'your local GitHub repo clone and commit+push it — so the Data tab\'s '
                             'fetch always sees this run\'s output without a manual git step. '
                             'Requires GITHUB_REPO_DIR to be set near the top of this file (or '
                             'passed via --repo-dir), pointing at your actual local clone of '
                             'Gatorz1989/NFL-Daily-Data-Pull, with push access already configured '
                             '(SSH key or stored credential — this does not handle interactive login).')
    parser.add_argument('--repo-dir', type=str, default=None,
                        help='Overrides GITHUB_REPO_DIR for this run only, without editing the file.')
    args = parser.parse_args()

    if args.backfill_rzdef:
        seasons = [int(s.strip()) for s in args.backfill_seasons.split(',') if s.strip()]
        backfill_historical_rzdef_proxy(seasons)
    elif args.backfill_pace_efficiency:
        seasons = [int(s.strip()) for s in args.backfill_seasons.split(',') if s.strip()]
        backfill_pace_efficiency_proxy(seasons)
    elif args.update:
        update_inseason_ngs(
            target_season=args.season or CURRENT_SEASON,
            weeks_back=args.weeks,
        )
        if args.push:
            _repo = args.repo_dir or GITHUB_REPO_DIR
            if _repo:
                push_output_to_github(_repo, OUTPUT_FILE)
            else:
                print("  ⚠ --push given but no repo path set — use --repo-dir or set "
                      "GITHUB_REPO_DIR near the top of this file.")
    else:
        main()
        if args.push:
            _repo = args.repo_dir or GITHUB_REPO_DIR
            if _repo:
                push_output_to_github(_repo, OUTPUT_FILE)
            else:
                print("  ⚠ --push given but no repo path set — use --repo-dir or set "
                      "GITHUB_REPO_DIR near the top of this file.")


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

