#!/usr/bin/env python3
"""
NFL Edge Model — weekly line-snapshot collector (v1.0, new in v7.42)
=====================================================================
Saves what the sportsbooks are posting RIGHT NOW so the model can later be checked against opening
and closing numbers (line movement, closing-line value, calibrated Over/Under and median projections).

What it saves, every run, into the line_snapshots/ folder of the repo:
  * games_<UTC stamp>.csv.gz   spread / total / moneyline per game per book
  * props_<UTC stamp>.csv.gz   one row per player, market and book: line, Over price, Under price
  * manifest.csv               one line per run: when, how many games/props, how many API credits were left

Needs one secret in GitHub (Settings > Secrets and variables > Actions): ODDS_API_KEY (your The Odds API key).
Run by the "Weekly Line Snapshot" workflow (Tuesday and Sunday mornings) or by hand:
    ODDS_API_KEY=xxxx python nfl_line_snapshot.py
Optional settings (environment variables):
    SNAPSHOT_PROPS=1        include player props (about 7 API credits per game). Set 0 for game lines only.
    SNAPSHOT_BOOKS=draftkings,fanduel,betmgm,williamhill_us   which books to keep
    SNAPSHOT_DIR=line_snapshots
"""
import csv, gzip, io, json, os, sys, time
from datetime import datetime, timezone
import requests

KEY = os.environ.get("ODDS_API_KEY", "").strip()
SPORT = "americanfootball_nfl"
BASE = "https://api.the-odds-api.com/v4"
BOOKS = [b.strip() for b in os.environ.get("SNAPSHOT_BOOKS", "draftkings,fanduel,betmgm,williamhill_us").split(",") if b.strip()]
WITH_PROPS = os.environ.get("SNAPSHOT_PROPS", "1").strip() not in ("0", "false", "no", "")
OUT = os.environ.get("SNAPSHOT_DIR", "line_snapshots")
PROP_MARKETS = ["player_pass_yds", "player_pass_tds", "player_rush_yds", "player_reception_yds", "player_receptions", "player_rush_tds", "player_anytime_td"]
STAMP = datetime.now(timezone.utc).strftime("%Y-%m-%d_%H%M")
NOW = datetime.now(timezone.utc).isoformat(timespec="seconds")


def get(url, params, tries=3):
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, params=params, timeout=40)
            remaining = r.headers.get("x-requests-remaining")
            if r.status_code == 200:
                return r.json(), remaining
            last = f"HTTP {r.status_code}: {r.text[:200]}"
            if r.status_code in (401, 403, 422, 429):
                break
        except Exception as e:  # network trouble: try again
            last = str(e)
        time.sleep(2 * (i + 1))
    print(f"  request failed: {url.split('/v4')[-1]} -> {last}")
    return None, None


def write_gz(path, header, rows):
    buf = io.StringIO()
    w = csv.writer(buf); w.writerow(header); w.writerows(rows)
    with gzip.open(path, "wt", encoding="utf-8", newline="") as f:
        f.write(buf.getvalue())


def main():
    if not KEY:
        print("::warning::ODDS_API_KEY is not set, so no line snapshot was taken. Add it under Settings > Secrets and variables > Actions.")
        return 0
    os.makedirs(OUT, exist_ok=True)
    credits = None
    bm = ",".join(BOOKS)
    # 1. game lines for every upcoming game (3 credits)
    data, rem = get(f"{BASE}/sports/{SPORT}/odds", {"apiKey": KEY, "regions": "us", "markets": "h2h,spreads,totals", "oddsFormat": "american", "bookmakers": bm})
    if rem: credits = rem
    if data is None:
        print("No game lines came back; nothing saved.")
        return 0
    game_rows, events = [], []
    for ev in data:
        events.append({"id": ev.get("id"), "home": ev.get("home_team"), "away": ev.get("away_team"), "commence": ev.get("commence_time")})
        for b in ev.get("bookmakers", []):
            for m in b.get("markets", []):
                for o in m.get("outcomes", []):
                    game_rows.append([NOW, ev.get("id"), ev.get("commence_time"), ev.get("away_team"), ev.get("home_team"), b.get("key"), m.get("key"), o.get("name"), o.get("point"), o.get("price"), m.get("last_update")])
    write_gz(os.path.join(OUT, f"games_{STAMP}.csv.gz"), ["snapshot_utc", "event_id", "commence", "away", "home", "book", "market", "outcome", "point", "price", "book_updated"], game_rows)
    print(f"Game lines: {len(events)} games, {len(game_rows)} rows")
    # 2. player props, one request per game
    prop_rows = []
    if WITH_PROPS:
        for ev in events:
            d, rem = get(f"{BASE}/sports/{SPORT}/events/{ev['id']}/odds", {"apiKey": KEY, "regions": "us", "markets": ",".join(PROP_MARKETS), "oddsFormat": "american", "bookmakers": bm})
            if rem: credits = rem
            if not d:
                continue
            for b in d.get("bookmakers", []):
                for m in b.get("markets", []):
                    pts = {}
                    for o in m.get("outcomes", []):
                        who = o.get("description") or ""
                        side = str(o.get("name", "")).lower()
                        k = (who, o.get("point"))
                        slot = pts.setdefault(k, {"over": None, "under": None, "yes": None})
                        if side in slot:
                            slot[side] = o.get("price")
                    for (who, pt), v in pts.items():
                        prop_rows.append([NOW, ev["id"], ev["commence"], ev["away"], ev["home"], b.get("key"), m.get("key"), who, pt, v["over"] if v["over"] is not None else v["yes"], v["under"], m.get("last_update")])
            time.sleep(0.4)
        write_gz(os.path.join(OUT, f"props_{STAMP}.csv.gz"), ["snapshot_utc", "event_id", "commence", "away", "home", "book", "market", "player", "line", "over_price", "under_price", "book_updated"], prop_rows)
        print(f"Player props: {len(prop_rows)} rows")
    # 3. manifest
    man = os.path.join(OUT, "manifest.csv")
    new = not os.path.exists(man)
    with open(man, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["snapshot_utc", "games", "game_rows", "prop_rows", "books", "credits_remaining"])
        w.writerow([NOW, len(events), len(game_rows), len(prop_rows), bm, credits])
    print(f"Saved snapshot {STAMP}. API credits remaining: {credits}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
