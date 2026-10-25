# Findings

Synthetic case: every figure, plan, actual and outcome below is illustrative. No real company, deal or realized value.

## Results (Base / Downside / Upside)

| Case | Exit EBITDA ($m) | Exit multiple | MOIC | IRR |
|---|---:|---:|---:|---:|
| Base | 22.76 | 10.0x | 2.42x | 19.3% |
| Downside | 17.24 | 8.5x | 1.35x | 6.1% |
| Upside | 25.67 | 11.0x | 3.13x | 25.6% |

## What the model shows

- Core growth and margin drive about two-thirds of the Base EBITDA increase ($5.84m of $8.76m). The three initiatives add $2.92m; AI finance automation adds $0.18m.
- The AI lever is cash-negative over the hold (about -$0.44m cumulative net EBITDA after the $0.85m implementation cost). Its value to equity comes only through the exit multiple on run-rate savings (about $1.83m of exit EV).
- Across the full AI-realization range, Base MOIC moves by less than 0.1x; a 1.0x change in exit multiple moves it by about 0.26x.
- Base IRR sits just below an illustrative 20% hurdle; closing the gap needs roughly $0.6m more exit EBITDA or about 0.26x on the exit multiple.

## Verification

- 17 pytest tests pass: scenario switching, toggle-off removes each lever's net contribution, Python returns check matches the recalculated workbook for all three scenarios, no formula errors, no literals inside formulas, tracker plan non-decreasing, sensitivity grids monotonic, content restrictions.
- All 12 workbook checks PASS in Base. Returns independently recomputed in plain Python from the workbook inputs.

## Review log

- Round 0 review found: quarterly tracker plan reset to zero each year; case study had no returns table, bridge, sensitivity or returns effect for the worked example; no monotonicity test.
- Round 1 fixed all of the above, plus an AI sensitivity block that overlapped and read forecast years one column late (sensitivity only; the main cash-flow schedule was already aligned). The expand/fix/pause gate is now stated as a governance signal that does not change the forecast, with Pause taking precedence.

## Limitations

Simple tax, working-capital, fee and debt assumptions; interest on opening debt (no circularity); no market research, diligence validation or financing commitment. Not an investment recommendation.
