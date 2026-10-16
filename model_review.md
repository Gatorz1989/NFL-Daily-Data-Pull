# Model Review

Generated 2026-10-16T22:03:55Z from seasons 2024, 2025.

## Scorecard

| | 2024 | 2025 | All |
|---|---|---|---|
| Hit rate | 50.0% (n=6232) | 51.0% (n=6891) | 50.5% (n=13123) |
| Passing yards | 48.0% (n=377) | 49.1% (n=395) | 48.6% (n=772) |
| Passing TDs | 48.0% (n=377) | 51.1% (n=395) | 49.6% (n=772) |
| Rushing yards | 50.0% (n=1106) | 51.6% (n=1222) | 50.9% (n=2328) |
| Receiving yards | 49.3% (n=2294) | 49.9% (n=2525) | 49.6% (n=4819) |
| Receptions | 51.4% (n=2078) | 52.1% (n=2354) | 51.8% (n=4432) |

Break-even at -110 is 52.4%.

## Where to adjust (best-supported first)

1. **Rushing yards: the projection is less accurate than the Vegas line** [High confidence]
   - Finding: Average error is 21.8% worse than the line in every season (2024: 22.36 vs 17.83, 2025: 21.09 vs 17.83). Blending toward the line at a 5% model weight cut the average error from 21.73 to 17.84 in out-of-sample tests, against 17.83 for the line alone.
   - Suggestion: Show the calibrated projection (line + 5% × (raw − line)) next to the raw one. The blend only matches the line, so do not present large raw-versus-line gaps as high-confidence edges.

2. **Receiving yards: the projection is less accurate than the Vegas line** [High confidence]
   - Finding: Average error is 13.2% worse than the line in every season (2024: 22.27 vs 19.88, 2025: 21.35 vs 18.66). Blending toward the line at a 0% model weight cut the average error from 21.81 to 19.27 in out-of-sample tests, against 19.27 for the line alone.
   - Suggestion: Show the calibrated projection (line + 0% × (raw − line)) next to the raw one. The blend only matches the line, so do not present large raw-versus-line gaps as high-confidence edges.

3. **Rushing yards overs: below break-even in every season** [High confidence]
   - Finding: 47.1% over 1315 picks (95% range 44.4–49.8%).
   - Suggestion: Stop surfacing these as picks, or label them clearly as low confidence.

4. **Over picks lose money** [High confidence]
   - Finding: Over picks hit 49.1% over 8559 picks (95% range 48.0–50.1%), below the 52.4% needed at -110, and below it in every season (2024: 49.4%, 2025: 48.8%).
   - Suggestion: Treat every over pick with extra skepticism: raise the bar for surfacing one, or require support from a second signal.

5. **Receiving yards overs: below break-even in every season** [High confidence]
   - Finding: 49.2% over 3696 picks (95% range 47.6–50.8%).
   - Suggestion: Stop surfacing these as picks, or label them clearly as low confidence.

6. **Receptions overs: below break-even in every season** [High confidence]
   - Finding: 49.4% over 2388 picks (95% range 47.4–51.4%).
   - Suggestion: Stop surfacing these as picks, or label them clearly as low confidence.

7. **Passing yards: projections are spread too wide** [Medium confidence]
   - Finding: Actual results rise only 0.08, 0.37 yard(s) for every projected yard (1.0 would be a perfect match).
   - Suggestion: Pull projections toward the typical level for the position before the line comparison (the biggest errors are at the top of the range, see bias_by_size).

8. **2024: 77% of props have a projection** [Medium confidence]
   - Finding: 1884 of 8179 props could not be projected, so every hit rate describes only the other 77% (mostly early-season weeks and unmatched player names).
   - Suggestion: Compare any change on identical rows only, and look at which players or weeks are missing.

9. **2025: 79% of props have a projection** [Medium confidence]
   - Finding: 1894 of 8880 props could not be projected, so every hit rate describes only the other 79% (mostly early-season weeks and unmatched player names).
   - Suggestion: Compare any change on identical rows only, and look at which players or weeks are missing.

10. **Passing yards: the projection is less accurate than the Vegas line** [Medium confidence]
   - Finding: Average error is 16.8% worse than the line in every season (2024: 68.43 vs 59.87, 2025: 65.77 vs 55.10). Blending toward the line at a 0% model weight cut the average error from 67.10 to 57.49 in out-of-sample tests, against 57.49 for the line alone.
   - Suggestion: Show the calibrated projection (line + 0% × (raw − line)) next to the raw one. The blend only matches the line, so do not present large raw-versus-line gaps as high-confidence edges.

11. **Passing yards: projections run high** [Medium confidence]
   - Finding: The average projection is above than the average result by 7.4%, 12.2% in every season.
   - Suggestion: Trace this prop through its components (volume × efficiency, then each adjustment) to find where the gap is added. Do not apply one flat correction: the size changes by season.

12. **Receiving yards: projections run high** [Medium confidence]
   - Finding: The average projection is above than the average result by 6.0%, 13.1% in every season.
   - Suggestion: Trace this prop through its components (volume × efficiency, then each adjustment) to find where the gap is added. Do not apply one flat correction: the size changes by season.

13. **Receptions: the projection is less accurate than the Vegas line** [Medium confidence]
   - Finding: Average error is 6.0% worse than the line in every season (2024: 1.65 vs 1.55, 2025: 1.53 vs 1.44). Blending toward the line at a 25% model weight cut the average error from 1.59 to 1.49 in out-of-sample tests, against 1.50 for the line alone.
   - Suggestion: Show the calibrated projection (line + 25% × (raw − line)) next to the raw one. The blend only matches the line, so do not present large raw-versus-line gaps as high-confidence edges.

14. **Rushing yards unders: above break-even in every season** [Medium confidence]
   - Finding: 55.8% over 1013 picks (95% range 52.7–58.8%). Betting every under with no model at all hits 53.9%, so the model adds about +1.9 points.
   - Suggestion: Keep surfacing these picks. Most of the result is the market, not the model.

15. **Passing TDs: a blend with the line beats the line itself** [Medium confidence]
   - Finding: At a 50% model weight the blended projection beat the line in every out-of-sample test (average error 0.966 against 0.980).
   - Suggestion: Use a 50% model weight for the calibrated projection.

16. **Tier ranking holds for over picks in 2025** [Low confidence]
   - Finding: Using 2024 cutoffs, over picks in 2025 hit ELITE 56.3% down to SKIP 43.6%: the model ranks them well.
   - Suggestion: Keep the percentile tiers; they separate the better picks from the weaker ones.

17. **Passing yards unders: below break-even in every season** [Low confidence]
   - Finding: 43.8% over 176 picks (95% range 36.6–51.1%).
   - Suggestion: Stop surfacing these as picks, or label them clearly as low confidence.

18. **Under picks: the seasons disagree** [Low confidence]
   - Finding: Under picks cleared break-even in some seasons and not others (2024: 51.0%, 2025: 55.2%).
   - Suggestion: Do not act on this side until more weeks are in.

19. **Tier ranking does not hold for under picks in 2025** [Low confidence]
   - Finding: Using 2024 cutoffs, under picks in 2025 hit ELITE 55.7% / STRONG 57.5% / MED 54.5% / LEAN 55.4% / SKIP 53.5%: the order is not clean.
   - Suggestion: Do not treat a higher tier as more reliable for under picks until this ranks cleanly.

## Notes

- Recommendations only: this review never changes the model.
- Every weight and cutoff here was tuned on one season and tested on the other, in both directions. Two seasons is a small sample.
- Both seasons have now been looked at many times, so they are development data. The clean test is live picks frozen before kickoff.
- The hit-rate tables use -110 as the price; real prices are not stored in the archives yet.
