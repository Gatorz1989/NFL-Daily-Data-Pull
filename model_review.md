# Model Review

Generated 2026-10-27T19:03:05Z from seasons 2024, 2025, 2026.

## Scorecard

| | 2024 | 2025 | 2026 | All |
|---|---|---|---|---|
| Hit rate | 50.0% (n=6232) | 51.0% (n=6891) | 51.0% (n=804) | 50.5% (n=13927) |
| Passing yards | 48.0% (n=377) | 49.1% (n=395) | 36.2% (n=47) | 47.9% (n=819) |
| Passing TDs | 48.0% (n=377) | 51.1% (n=395) | 55.6% (n=45) | 49.9% (n=817) |
| Rushing yards | 50.0% (n=1106) | 51.6% (n=1222) | 51.7% (n=143) | 50.9% (n=2471) |
| Receiving yards | 49.3% (n=2294) | 49.9% (n=2525) | 50.3% (n=288) | 49.7% (n=5107) |
| Receptions | 51.4% (n=2078) | 52.1% (n=2354) | 53.0% (n=281) | 51.8% (n=4713) |

Break-even at -110 is 52.4%.

## Where to adjust (best-supported first)

1. **Tier ranking holds for over picks in 2025** [High confidence]
   - Finding: Using 2024 cutoffs, over picks in 2025 hit ELITE 56.3% down to SKIP 43.6%: the model ranks them well.
   - Suggestion: Keep the percentile tiers; they separate the better picks from the weaker ones.

2. **2026: 32% of props have a projection** [Medium confidence]
   - Finding: 1689 of 2499 props could not be projected, so every hit rate describes only the other 32% (mostly early-season weeks and unmatched player names).
   - Suggestion: Compare any change on identical rows only, and look at which players or weeks are missing.

3. **2024: 77% of props have a projection** [Medium confidence]
   - Finding: 1884 of 8179 props could not be projected, so every hit rate describes only the other 77% (mostly early-season weeks and unmatched player names).
   - Suggestion: Compare any change on identical rows only, and look at which players or weeks are missing.

4. **2025: 79% of props have a projection** [Medium confidence]
   - Finding: 1894 of 8880 props could not be projected, so every hit rate describes only the other 79% (mostly early-season weeks and unmatched player names).
   - Suggestion: Compare any change on identical rows only, and look at which players or weeks are missing.

5. **Receiving yards: the projection is less accurate than the Vegas line** [Medium confidence]
   - Finding: Average error is 11.8% worse than the line in every season (2024: 22.27 vs 19.88, 2025: 21.35 vs 18.66, 2026: 25.65 vs 23.54). Blending toward the line at a 0% model weight cut the average error from 23.09 to 20.69 in out-of-sample tests, against 20.69 for the line alone.
   - Suggestion: Show the calibrated projection (line + 0% × (raw − line)) next to the raw one. The blend only matches the line, so do not present large raw-versus-line gaps as high-confidence edges.

6. **Receptions: the projection is less accurate than the Vegas line** [Medium confidence]
   - Finding: Average error is 6.5% worse than the line in every season (2024: 1.65 vs 1.55, 2025: 1.53 vs 1.44, 2026: 1.83 vs 1.70). Blending toward the line at a 25% model weight cut the average error from 1.67 to 1.55 in out-of-sample tests, against 1.57 for the line alone.
   - Suggestion: Show the calibrated projection (line + 25% × (raw − line)) next to the raw one. The blend only matches the line, so do not present large raw-versus-line gaps as high-confidence edges.

7. **Over picks lose money** [Medium confidence]
   - Finding: Over picks hit 49.1% over 9000 picks (95% range 48.1–50.1%), below the 52.4% needed at -110, and below it in every season (2024: 49.4%, 2025: 48.8%, 2026: 49.7%).
   - Suggestion: Treat every over pick with extra skepticism: raise the bar for surfacing one, or require support from a second signal.

8. **Receiving yards overs: below break-even in every season** [Medium confidence]
   - Finding: 49.2% over 3867 picks (95% range 47.7–50.8%).
   - Suggestion: Stop surfacing these as picks, or label them clearly as low confidence.

9. **Passing yards: projections are spread too wide** [Low confidence]
   - Finding: Actual results rise only 0.08, 0.37, 0.36 yard(s) for every projected yard (1.0 would be a perfect match).
   - Suggestion: Pull projections toward the typical level for the position before the line comparison (the biggest errors are at the top of the range, see bias_by_size).

10. **Tier ranking holds for over picks in 2026** [Low confidence]
   - Finding: Using 2024 cutoffs, over picks in 2026 hit ELITE 63.6% down to SKIP 38.3%: the model ranks them well.
   - Suggestion: Keep the percentile tiers; they separate the better picks from the weaker ones.

11. **Passing yards: the projection is less accurate than the Vegas line** [Low confidence]
   - Finding: Average error is 17.6% worse than the line in every season (2024: 68.43 vs 59.87, 2025: 65.77 vs 55.10, 2026: 71.23 vs 59.76). Blending toward the line at a 0% model weight cut the average error from 67.10 to 57.49 in out-of-sample tests, against 57.49 for the line alone.
   - Suggestion: Show the calibrated projection (line + 0% × (raw − line)) next to the raw one. The blend only matches the line, so do not present large raw-versus-line gaps as high-confidence edges.

12. **Rushing yards: the projection is less accurate than the Vegas line** [Low confidence]
   - Finding: Average error is 17.2% worse than the line in every season (2024: 22.36 vs 17.83, 2025: 21.09 vs 17.83, 2026: 19.89 vs 18.43). Blending toward the line at a 5% model weight cut the average error from 21.11 to 18.03 in out-of-sample tests, against 18.03 for the line alone.
   - Suggestion: Show the calibrated projection (line + 5% × (raw − line)) next to the raw one. The blend only matches the line, so do not present large raw-versus-line gaps as high-confidence edges.

13. **Rushing yards overs: below break-even in every season** [Low confidence]
   - Finding: 46.9% over 1395 picks (95% range 44.3–49.5%).
   - Suggestion: Stop surfacing these as picks, or label them clearly as low confidence.

14. **Rushing yards unders: above break-even in every season** [Low confidence]
   - Finding: 56.1% over 1076 picks (95% range 53.2–59.1%). Betting every under with no model at all hits 54.6%, so the model adds about +1.5 points.
   - Suggestion: Keep surfacing these picks. Most of the result is the market, not the model.

15. **Passing TDs: a blend with the line beats the line itself** [Low confidence]
   - Finding: At a 50% model weight the blended projection beat the line in every out-of-sample test (average error 0.966 against 0.980).
   - Suggestion: Use a 50% model weight for the calibrated projection.

16. **Under picks: the seasons disagree** [Low confidence]
   - Finding: Under picks cleared break-even in some seasons and not others (2024: 51.0%, 2025: 55.2%, 2026: 52.6%).
   - Suggestion: Do not act on this side until more weeks are in.

17. **Tier ranking does not hold for under picks in 2025** [Low confidence]
   - Finding: Using 2024 cutoffs, under picks in 2025 hit ELITE 55.7% / STRONG 57.5% / MED 54.5% / LEAN 55.4% / SKIP 53.5%: the order is not clean.
   - Suggestion: Do not treat a higher tier as more reliable for under picks until this ranks cleanly.

18. **Tier ranking does not hold for under picks in 2026** [Low confidence]
   - Finding: Using 2024 cutoffs, under picks in 2026 hit ELITE 55.3% / STRONG 53.5% / MED 41.6% / LEAN 55.3% / SKIP 58.4%: the order is not clean.
   - Suggestion: Do not treat a higher tier as more reliable for under picks until this ranks cleanly.

## Notes

- Recommendations only: this review never changes the model.
- Every weight and cutoff here was tuned on one season and tested on the other, in both directions. Two seasons is a small sample.
- Both seasons have now been looked at many times, so they are development data. The clean test is live picks frozen before kickoff.
- The hit-rate tables use -110 as the price; real prices are not stored in the archives yet.
