# Corrected Dataset Ranking Diagnostics

Ranked members: 1,234
Events analyzed: 53
Date range: 2025-04-04 to 2026-05-16

## Highest Correlations With JPAR

- `mean_adjusted_event_jpar`: 0.993
- `trueskill_conservative`: 0.882
- `best3_mean_log_zscore`: 0.794
- `conservative_log_zscore`: 0.791
- `mean_event_percentile`: 0.787
- `mean_normalized_rank`: 0.781
- `robust_log_zscore`: 0.780
- `elo_rating`: 0.775
- `mean_log_zscore`: 0.767
- `weighted_event_percentile`: 0.767

## Notes

- `weighted_log_zscore`, `weighted_event_percentile`, Elo, and TrueSkill are useful comparison systems because they reduce or avoid direct dependence on JPAR's event calibration multiplier.
- Positive `rank_delta_vs_jpar` in the disagreement CSV means the alternative system ranks someone worse than JPAR; negative means it ranks them better.
- The PDF is the primary review artifact; CSVs are included only for drill-down.
