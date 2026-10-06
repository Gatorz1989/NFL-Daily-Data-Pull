#!/usr/bin/env python3
"""
Model Review — The Edge Report / NFL Edge Model
================================================
Reads the saved season archives in this folder (nfl_backtest_2024.json, nfl_backtest_2025.json, nfl_backtest_2026.json ...),
re-runs every check we do by hand, and writes:

  model_review.json   what the Model Review tab in the model reads
  model_review.md     the same report as plain text

It needs no internet, no API key and no extra packages (standard library only), so it runs on GitHub Actions for free.
It RECOMMENDS only: it never changes the model. Every finding carries a confidence label:

  High    both seasons agree, at least 500 picks in each, and the result clearly clears or misses the bar
  Medium  both seasons agree, at least 150 picks in each
  Low     everything else (small sample, or the seasons disagree)

Tuning is always done on one season and tested on the other (both directions). Treat anything found by looking at the same
data it was tuned on as a candidate, not a fact.
"""
import datetime
import glob
import json
import math
import os
import re
import statistics as st
import sys

BREAK_EVEN = 52.38   # hit rate needed at -110
PROPS = ['pass_yds', 'pass_tds', 'rush_yds', 'rec_yds', 'recs']
LABEL = {'pass_yds': 'Passing yards', 'pass_tds': 'Passing TDs', 'rush_yds': 'Rushing yards', 'rec_yds': 'Receiving yards', 'recs': 'Receptions'}
# The weights the model uses today (V721_CAL_WEIGHTS in the model). The review compares its own recommendation with these.
CURRENT_WEIGHTS = {'pass_yds': 0.0, 'rec_yds': 0.0, 'rush_yds': 0.05, 'recs': 0.25, 'pass_tds': 0.5}
MIN_SEASON_ROWS = 300


# ───────────────────────────── helpers ─────────────────────────────
def wilson(k, n, z=1.96):
    if not n:
        return (None, None)
    p = k / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return (100 * (c - h), 100 * (c + h))


def rate(rows):
    n = len(rows)
    k = sum(1 for x in rows if x['hit'])
    lo, hi = wilson(k, n)
    return {'n': n, 'hits': k, 'rate': round(100 * k / n, 1) if n else None,
            'lo': round(lo, 1) if lo is not None else None, 'hi': round(hi, 1) if hi is not None else None}


def quantile(a, p):
    n = len(a)
    if not n:
        return None
    if n == 1:
        return a[0]
    i = p * (n - 1)
    lo = int(math.floor(i))
    hi = min(lo + 1, n - 1)
    return a[lo] + (a[hi] - a[lo]) * (i - lo)


def picks_of(rows):
    return [x for x in rows if x.get('hit') is not None and x.get('modelDir') and x.get('edgePct') is not None]


def paired(rows, prop):
    return [x for x in rows if x.get('propType') == prop and x.get('modelProj') is not None and x.get('line') is not None and x.get('actual') is not None]


def mae(rows, f):
    return st.mean(abs(f(x) - x['actual']) for x in rows) if rows else None


def conf(n_min, consistent, clear):
    if consistent and clear and n_min >= 500:
        return 'High'
    if consistent and n_min >= 150:
        return 'Medium'
    return 'Low'


def load_seasons(folder):
    out = {}
    for p in sorted(glob.glob(os.path.join(folder, 'nfl_backtest_*.json'))):
        m = re.search(r'nfl_backtest_(\d{4})\.json$', p)
        if not m:
            continue
        try:
            d = json.load(open(p, encoding='utf-8'))
        except Exception as e:  # a broken file must never stop the review
            print('skipping', p, e, file=sys.stderr)
            continue
        rows = d.get('graded') or []
        if len(rows) >= MIN_SEASON_ROWS:
            out[int(m.group(1))] = rows
    return out


def tier_cuts(rows):
    groups = {}
    for x in picks_of(rows):
        groups.setdefault('%s|%s' % (x['propType'], x['modelDir']), []).append(abs(x['edgePct']))
    cuts = {}
    for k, v in groups.items():
        v.sort()
        if len(v) >= 10:
            cuts[k] = {'n': len(v), 'cuts': [round(quantile(v, p), 3) for p in (.3, .5, .7, .9)]}
    return cuts


def tier_of(x, cuts):
    c = cuts.get('%s|%s' % (x['propType'], x['modelDir']))
    if not c:
        return None
    a = abs(x['edgePct'])
    q = c['cuts']
    return 'ELITE' if a >= q[3] else 'STRONG' if a >= q[2] else 'MED' if a >= q[1] else 'LEAN' if a >= q[0] else 'SKIP'


def best_lambda(train, prop):
    rows = paired(train, prop)
    grid = [i / 100 for i in range(0, 101)]
    return min(grid, key=lambda l: mae(rows, lambda x: x['line'] + l * (x['modelProj'] - x['line'])))


def ols_slope(rows):
    xs = [x['modelProj'] for x in rows]
    ys = [x['actual'] for x in rows]
    mx, my = st.mean(xs), st.mean(ys)
    sxx = sum((a - mx) ** 2 for a in xs)
    return (sum((a - mx) * (b - my) for a, b in zip(xs, ys)) / sxx) if sxx else None


# ───────────────────────────── the review ─────────────────────────────
def review(folder):
    seasons = load_seasons(folder)
    ys = sorted(seasons)
    if not ys:
        raise SystemExit('No nfl_backtest_YYYY.json files with at least %d rows found in %s' % (MIN_SEASON_ROWS, folder))
    ref = ys[0]
    ref_cuts = tier_cuts(seasons[ref])
    R = {'generated': datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).strftime('%Y-%m-%dT%H:%M:%SZ'), 'seasons': ys, 'reference_season': ref,
         'rowCounts': {str(y): len(seasons[y]) for y in ys}, 'break_even': BREAK_EVEN}
    allp = {y: picks_of(seasons[y]) for y in ys}

    # coverage
    R['coverage'] = {str(y): {'rows': len(seasons[y]), 'with_projection': sum(1 for x in seasons[y] if x.get('modelProj') is not None),
                              'picks': len(allp[y])} for y in ys}
    for y in ys:
        c = R['coverage'][str(y)]
        c['pct'] = round(100 * c['with_projection'] / c['rows'], 1)

    # hit rates
    def all_rows():
        return [x for y in ys for x in allp[y]]
    R['hit'] = {'overall': {str(y): rate(allp[y]) for y in ys}}
    R['hit']['overall']['all'] = rate(all_rows())
    R['hit']['by_prop'] = {}
    R['hit']['by_prop_side'] = {}
    for pr in PROPS:
        R['hit']['by_prop'][pr] = {str(y): rate([x for x in allp[y] if x['propType'] == pr]) for y in ys}
        R['hit']['by_prop'][pr]['all'] = rate([x for x in all_rows() if x['propType'] == pr])
        for sd in ('over', 'under'):
            k = '%s|%s' % (pr, sd)
            R['hit']['by_prop_side'][k] = {str(y): rate([x for x in allp[y] if x['propType'] == pr and x['modelDir'] == sd]) for y in ys}
            R['hit']['by_prop_side'][k]['all'] = rate([x for x in all_rows() if x['propType'] == pr and x['modelDir'] == sd])
    R['hit']['by_side'] = {}
    for sd in ('over', 'under'):
        R['hit']['by_side'][sd] = {str(y): rate([x for x in allp[y] if x['modelDir'] == sd]) for y in ys}
        R['hit']['by_side'][sd]['all'] = rate([x for x in all_rows() if x['modelDir'] == sd])

    # the market baseline: how often does ANY prop simply go under the line, with no model at all
    R['baseline'] = {'always_under': {}, 'by_prop': {}}
    for y in ys:
        rs = [x for x in seasons[y] if x.get('actual') is not None and x.get('line') is not None]
        R['baseline']['always_under'][str(y)] = round(100 * sum(1 for x in rs if x['actual'] < x['line']) / len(rs), 1)
    for pr in PROPS:
        R['baseline']['by_prop'][pr] = {}
        for y in ys:
            rs = [x for x in seasons[y] if x.get('propType') == pr and x.get('actual') is not None and x.get('line') is not None]
            if rs:
                R['baseline']['by_prop'][pr][str(y)] = round(100 * sum(1 for x in rs if x['actual'] < x['line']) / len(rs), 1)

    # accuracy against the actual result and against the line
    R['accuracy'] = {}
    R['calibration'] = {}
    for pr in PROPS:
        acc = {}
        for y in ys:
            rs = paired(seasons[y], pr)
            if len(rs) < 30:
                continue
            m_model = mae(rs, lambda x: x['modelProj'])
            m_line = mae(rs, lambda x: x['line'])
            acc[str(y)] = {'n': len(rs), 'model_mae': round(m_model, 3), 'line_mae': round(m_line, 3),
                           'bias': round(st.mean(x['modelProj'] - x['actual'] for x in rs), 3),
                           'mean_actual': round(st.mean(x['actual'] for x in rs), 3),
                           'closer_pct': round(100 * sum(1 for x in rs if abs(x['modelProj'] - x['actual']) < abs(x['line'] - x['actual'])) / len(rs), 1),
                           'slope': round(ols_slope(rs), 3) if ols_slope(rs) is not None else None}
        R['accuracy'][pr] = acc
        cal = {'pairs': []}
        lams = []
        for a in ys:
            for b in ys:
                if a == b or len(paired(seasons[a], pr)) < 100 or len(paired(seasons[b], pr)) < 100:
                    continue
                lam = best_lambda(seasons[a], pr)
                test = paired(seasons[b], pr)
                cal['pairs'].append({'tuned_on': a, 'tested_on': b, 'model_weight': lam, 'n_test': len(test),
                                     'raw_mae': round(mae(test, lambda x: x['modelProj']), 3), 'line_mae': round(mae(test, lambda x: x['line']), 3),
                                     'blend_mae': round(mae(test, lambda x: x['line'] + lam * (x['modelProj'] - x['line'])), 3)})
                lams.append(lam)
        if lams:
            cal['recommended_weight'] = round(round(st.mean(lams) / 0.05) * 0.05, 2)
            cal['current_weight'] = CURRENT_WEIGHTS.get(pr)
            cal['beats_line_everywhere'] = all(p['blend_mae'] < p['line_mae'] * 0.995 for p in cal['pairs'])
        R['calibration'][pr] = cal

    # bias by projection size for yardage props (does the error grow with the projection?)
    R['bias_by_size'] = {}
    for pr in ('pass_yds', 'rec_yds', 'rush_yds'):
        rs = sorted([x for y in ys for x in paired(seasons[y], pr)], key=lambda x: x['modelProj'])
        n = len(rs)
        if n < 90:
            continue
        thirds = []
        for i, nm in enumerate(('low', 'mid', 'high')):
            c = rs[i * n // 3:(i + 1) * n // 3]
            thirds.append({'third': nm, 'avg_projection': round(st.mean(x['modelProj'] for x in c), 1), 'avg_actual': round(st.mean(x['actual'] for x in c), 1),
                           'bias': round(st.mean(x['modelProj'] - x['actual'] for x in c), 1)})
        R['bias_by_size'][pr] = thirds

    # tiers: the saved reference season sets the cutoffs, so every other season is an honest out-of-sample test
    R['tier_cuts'] = ref_cuts
    R['tiers'] = {}
    for y in ys:
        tab = {}
        for sd in ('over', 'under'):
            tab[sd] = {}
            for t in ('ELITE', 'STRONG', 'MED', 'LEAN', 'SKIP'):
                rows = [x for x in allp[y] if x['modelDir'] == sd and tier_of(x, ref_cuts) == t]
                r = rate(rows) if rows else {'n': 0}
                if rows:
                    p = r['hits'] / r['n']
                    r['roi_pct'] = round((p * (100 / 110) - (1 - p)) * 100, 1)
                    r['verdict'] = 'above break-even' if r['lo'] > BREAK_EVEN else ('below break-even' if r['hi'] < BREAK_EVEN else 'no proven edge')
                tab[sd][t] = r
        R['tiers'][str(y)] = {'sides': tab, 'in_sample': (y == ref)}

    # ───────── where to adjust ─────────
    items = []

    def add(area, title, finding, suggestion, confidence, effect, n_min, evidence=None):
        items.append({'area': area, 'title': title, 'finding': finding, 'suggestion': suggestion, 'confidence': confidence,
                      'effect': round(effect, 3), 'support_n': n_min, 'evidence': evidence or {}})

    for pr in PROPS:
        acc = R['accuracy'].get(pr, {})
        if len(acc) < 1:
            continue
        ns = [a['n'] for a in acc.values()]
        ratios = [a['model_mae'] / a['line_mae'] for a in acc.values()]
        worse = all(r >= 1.03 for r in ratios)
        better = all(r <= 0.97 for r in ratios)
        cal = R['calibration'].get(pr, {})
        w = cal.get('recommended_weight')
        pairs = cal.get('pairs', [])
        if worse and pairs:
            avg_raw = st.mean(p['raw_mae'] for p in pairs)
            avg_bl = st.mean(p['blend_mae'] for p in pairs)
            avg_ln = st.mean(p['line_mae'] for p in pairs)
            add(LABEL[pr], '%s: the projection is less accurate than the Vegas line' % LABEL[pr],
                'Average error is %.1f%% worse than the line in every season (%s). Blending toward the line at a %d%% model weight cut the average error from %.2f to %.2f in out-of-sample tests, against %.2f for the line alone.'
                % (100 * (st.mean(ratios) - 1), ', '.join('%s: %.2f vs %.2f' % (y, a['model_mae'], a['line_mae']) for y, a in acc.items()), round((w or 0) * 100), avg_raw, avg_bl, avg_ln),
                'Show the calibrated projection (line + %d%% × (raw − line)) next to the raw one. %s'
                % (round((w or 0) * 100), 'The blend only matches the line, so do not present large raw-versus-line gaps as high-confidence edges.' if not cal.get('beats_line_everywhere') else 'The blend beats the line, so this prop is where the model adds value.'),
                conf(min(ns), True, st.mean(ratios) >= 1.10), st.mean(ratios) - 1, min(ns), {'weight': w, 'pairs': pairs})
        elif better:
            add(LABEL[pr], '%s: the projection beats the Vegas line' % LABEL[pr],
                'Average error is %.1f%% lower than the line in every season.' % (100 * (1 - st.mean(ratios))),
                'Keep the model weight high for this prop and look for what drives the edge.', conf(min(ns), True, st.mean(ratios) <= 0.90), 1 - st.mean(ratios), min(ns))
        elif cal.get('beats_line_everywhere'):
            avg_bl = st.mean(p['blend_mae'] for p in pairs)
            avg_ln = st.mean(p['line_mae'] for p in pairs)
            add(LABEL[pr], '%s: a blend with the line beats the line itself' % LABEL[pr],
                'At a %d%% model weight the blended projection beat the line in every out-of-sample test (average error %.3f against %.3f).' % (round((w or 0) * 100), avg_bl, avg_ln),
                'Use a %d%% model weight for the calibrated projection.' % round((w or 0) * 100), conf(min(ns), True, False), (avg_ln - avg_bl) / avg_ln, min(ns), {'weight': w, 'pairs': pairs})
        # level bias
        biases = [(a['bias'] / a['mean_actual'] * 100) if a['mean_actual'] else 0 for a in acc.values()]
        if len(biases) >= 1 and all(abs(b) >= 5 for b in biases) and (all(b > 0 for b in biases) or all(b < 0 for b in biases)):
            add(LABEL[pr], '%s: projections run %s' % (LABEL[pr], 'high' if biases[0] > 0 else 'low'),
                'The average projection is %s than the average result by %s in every season.' % ('above' if biases[0] > 0 else 'below', ', '.join('%.1f%%' % abs(b) for b in biases)),
                'Trace this prop through its components (volume × efficiency, then each adjustment) to find where the gap is added. Do not apply one flat correction: the size changes by season.',
                conf(min(ns), True, all(abs(b) >= 8 for b in biases)), abs(st.mean(biases)) / 100, min(ns), {'bias_pct_by_season': biases})
        # spread
        slopes = [a['slope'] for a in acc.values() if a['slope'] is not None]
        if pr in ('pass_yds', 'rec_yds', 'rush_yds') and slopes and all(s < 0.6 for s in slopes):
            add(LABEL[pr], '%s: projections are spread too wide' % LABEL[pr],
                'Actual results rise only %s yard(s) for every projected yard (1.0 would be a perfect match).' % ', '.join('%.2f' % s for s in slopes),
                'Pull projections toward the typical level for the position before the line comparison (the biggest errors are at the top of the range, see bias_by_size).',
                conf(min(ns), True, all(s < 0.4 for s in slopes)), 1 - st.mean(slopes), min(ns), {'slope_by_season': slopes, 'bias_by_size': R['bias_by_size'].get(pr)})

    # sides
    for sd in ('over', 'under'):
        per = [R['hit']['by_side'][sd][str(y)] for y in ys if R['hit']['by_side'][sd][str(y)]['n']]
        al = R['hit']['by_side'][sd]['all']
        if len(per) < 1:
            continue
        nmin = min(p['n'] for p in per)
        if all(p['rate'] < BREAK_EVEN for p in per) and al['hi'] < BREAK_EVEN:
            add('Picks', '%s picks lose money' % sd.title(),
                '%s picks hit %.1f%% over %d picks (95%% range %.1f–%.1f%%), below the %.1f%% needed at -110, and below it in every season (%s).'
                % (sd.title(), al['rate'], al['n'], al['lo'], al['hi'], BREAK_EVEN, ', '.join('%s: %.1f%%' % (y, R['hit']['by_side'][sd][str(y)]['rate']) for y in ys)),
                'Treat every %s pick with extra skepticism: raise the bar for surfacing one, or require support from a second signal.' % sd,
                conf(nmin, True, True), (BREAK_EVEN - al['rate']) / 100, nmin, {'by_season': {str(y): R['hit']['by_side'][sd][str(y)] for y in ys}})
        elif all(p['rate'] > BREAK_EVEN for p in per) and al['lo'] > BREAK_EVEN:
            add('Picks', '%s picks beat break-even' % sd.title(),
                '%s picks hit %.1f%% over %d picks (95%% range %.1f–%.1f%%), above %.1f%% in every season.' % (sd.title(), al['rate'], al['n'], al['lo'], al['hi'], BREAK_EVEN),
                'Make sure these picks are surfaced; check how much of this is the market (see the baseline).', conf(nmin, True, True), (al['rate'] - BREAK_EVEN) / 100, nmin)
        elif any(p['rate'] > BREAK_EVEN for p in per) and any(p['rate'] < BREAK_EVEN for p in per):
            add('Picks', '%s picks: the seasons disagree' % sd.title(),
                '%s picks cleared break-even in some seasons and not others (%s).' % (sd.title(), ', '.join('%s: %.1f%%' % (y, R['hit']['by_side'][sd][str(y)]['rate']) for y in ys)),
                'Do not act on this side until more weeks are in.', 'Low', 0.0, nmin)

    # prop x side
    for pr in PROPS:
        for sd in ('over', 'under'):
            k = '%s|%s' % (pr, sd)
            per = [R['hit']['by_prop_side'][k][str(y)] for y in ys if R['hit']['by_prop_side'][k][str(y)]['n'] >= 30]
            al = R['hit']['by_prop_side'][k]['all']
            if len(per) < len(ys) or not al['n']:
                continue
            nmin = min(p['n'] for p in per)
            base = [R['baseline']['by_prop'][pr].get(str(y)) for y in ys if R['baseline']['by_prop'][pr].get(str(y)) is not None]
            naive = (st.mean(base) if sd == 'under' else 100 - st.mean(base)) if base else None
            if all(p['rate'] > BREAK_EVEN for p in per) and al['lo'] > BREAK_EVEN:
                add(LABEL[pr], '%s %ss: above break-even in every season' % (LABEL[pr], sd),
                    '%.1f%% over %d picks (95%% range %.1f–%.1f%%).%s' % (al['rate'], al['n'], al['lo'], al['hi'],
                                                                      (' Betting every %s with no model at all hits %.1f%%, so the model adds about %+.1f points.' % (sd, naive, al['rate'] - naive)) if naive is not None else ''),
                    'Keep surfacing these picks. %s' % ('Most of the result is the market, not the model.' if naive is not None and al['rate'] - naive < 2 else ''),
                    conf(nmin, True, True), (al['rate'] - BREAK_EVEN) / 100, nmin, {'cell': al, 'naive_rate': naive})
            elif all(p['rate'] < BREAK_EVEN for p in per) and al['hi'] < BREAK_EVEN:
                add(LABEL[pr], '%s %ss: below break-even in every season' % (LABEL[pr], sd),
                    '%.1f%% over %d picks (95%% range %.1f–%.1f%%).' % (al['rate'], al['n'], al['lo'], al['hi']),
                    'Stop surfacing these as picks, or label them clearly as low confidence.', conf(nmin, True, True), (BREAK_EVEN - al['rate']) / 100, nmin, {'cell': al})

    # tier ranking out of sample
    for y in ys:
        if y == ref:
            continue
        for sd in ('over', 'under'):
            t = R['tiers'][str(y)]['sides'][sd]
            rates = [t[x].get('rate') for x in ('ELITE', 'STRONG', 'MED', 'LEAN', 'SKIP')]
            ns_ = [t[x].get('n', 0) for x in ('ELITE', 'STRONG', 'MED', 'LEAN', 'SKIP')]
            if None in rates or min(ns_) < 30:
                continue
            inversions = sum(1 for i in range(4) if rates[i] < rates[i + 1])
            if rates[0] <= rates[4] or inversions >= 2:
                add('Tiers', 'Tier ranking does not hold for %s picks in %d' % (sd, y),
                    'Using %d cutoffs, %s picks in %d hit ELITE %.1f%% / STRONG %.1f%% / MED %.1f%% / LEAN %.1f%% / SKIP %.1f%%: the order is not clean.' % ((ref, sd, y) + tuple(rates)),
                    'Do not treat a higher tier as more reliable for %s picks until this ranks cleanly.' % sd, conf(min(ns_), False, False), 0.0, min(ns_), {'tiers': t})
            elif rates[0] - rates[4] >= 8:
                add('Tiers', 'Tier ranking holds for %s picks in %d' % (sd, y),
                    'Using %d cutoffs, %s picks in %d hit ELITE %.1f%% down to SKIP %.1f%%: the model ranks them well.' % (ref, sd, y, rates[0], rates[4]),
                    'Keep the percentile tiers; they separate the better picks from the weaker ones.', conf(min(ns_), len(ys) > 2, rates[0] - rates[4] >= 12), (rates[0] - rates[4]) / 100, min(ns_), {'tiers': t})

    # coverage
    for y in ys:
        c = R['coverage'][str(y)]
        if c['pct'] < 85:
            add('Data', '%d: %.0f%% of props have a projection' % (y, c['pct']),
                '%d of %d props could not be projected, so every hit rate describes only the other %.0f%% (mostly early-season weeks and unmatched player names).' % (c['rows'] - c['with_projection'], c['rows'], c['pct']),
                'Compare any change on identical rows only, and look at which players or weeks are missing.', 'Medium', (100 - c['pct']) / 100, c['rows'])

    # recommended model settings vs what the model uses now
    settings = {}
    for pr in PROPS:
        cal = R['calibration'].get(pr, {})
        if 'recommended_weight' in cal:
            settings[pr] = {'recommended_weight': cal['recommended_weight'], 'current_weight': cal.get('current_weight')}
            if cal.get('current_weight') is not None and abs(cal['recommended_weight'] - cal['current_weight']) >= 0.1:
                add(LABEL[pr], '%s: update the model weight' % LABEL[pr],
                    'This review recommends a %d%% model weight; the model is using %d%%.' % (round(cal['recommended_weight'] * 100), round(cal['current_weight'] * 100)),
                    'Change V721_CAL_WEIGHTS for this prop.', 'Medium', abs(cal['recommended_weight'] - cal['current_weight']), 0)
    R['recommended_settings'] = settings

    order = {'High': 0, 'Medium': 1, 'Low': 2}
    items.sort(key=lambda it: (order[it['confidence']], -it['effect']))
    for i, it in enumerate(items, 1):
        it['rank'] = i
    R['where_to_adjust'] = items
    R['notes'] = [
        'Recommendations only: this review never changes the model.',
        'Every weight and cutoff here was tuned on one season and tested on the other, in both directions. Two seasons is a small sample.',
        'Both seasons have now been looked at many times, so they are development data. The clean test is live picks frozen before kickoff.',
        'The hit-rate tables use -110 as the price; real prices are not stored in the archives yet.',
    ]
    return R


def write_markdown(R, path):
    L = ['# Model Review', '', 'Generated %s from seasons %s.' % (R['generated'], ', '.join(str(y) for y in R['seasons'])), '']
    L += ['## Scorecard', '', '| | ' + ' | '.join(str(y) for y in R['seasons']) + ' | All |', '|---|' + '---|' * (len(R['seasons']) + 1)]
    o = R['hit']['overall']
    L.append('| Hit rate | ' + ' | '.join('%.1f%% (n=%d)' % (o[str(y)]['rate'], o[str(y)]['n']) for y in R['seasons']) + ' | %.1f%% (n=%d) |' % (o['all']['rate'], o['all']['n']))
    for pr in PROPS:
        b = R['hit']['by_prop'][pr]
        L.append('| %s | ' % LABEL[pr] + ' | '.join('%.1f%% (n=%d)' % (b[str(y)]['rate'], b[str(y)]['n']) if b[str(y)]['n'] else '—' for y in R['seasons']) + ' | %.1f%% (n=%d) |' % (b['all']['rate'], b['all']['n']))
    L += ['', 'Break-even at -110 is %.1f%%.' % R['break_even'], '', '## Where to adjust (best-supported first)', '']
    for it in R['where_to_adjust']:
        L += ['%d. **%s** [%s confidence]' % (it['rank'], it['title'], it['confidence']), '   - Finding: %s' % it['finding'], '   - Suggestion: %s' % it['suggestion'], '']
    L += ['## Notes', ''] + ['- %s' % n for n in R['notes']]
    open(path, 'w', encoding='utf-8').write('\n'.join(L) + '\n')


if __name__ == '__main__':
    folder = sys.argv[1] if len(sys.argv) > 1 else '.'
    out_dir = sys.argv[2] if len(sys.argv) > 2 else folder
    R = review(folder)
    json.dump(R, open(os.path.join(out_dir, 'model_review.json'), 'w', encoding='utf-8'), separators=(',', ':'))
    write_markdown(R, os.path.join(out_dir, 'model_review.md'))
    print('Model Review written: %d findings (%s)' % (len(R['where_to_adjust']), ', '.join('%s %d' % (k, sum(1 for i in R['where_to_adjust'] if i['confidence'] == k)) for k in ('High', 'Medium', 'Low'))))
