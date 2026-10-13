# QA Summary — Sola Kitchen Marketplace P&L

**Status:** The model was recalculated with LibreOffice (headless): 105,405 formulas, **0 errors**, and all 29 checks on `QA_Checks` read TRUE. Key outputs were also recomputed independently in pandas, outside the workbook, and matched to the cent. **No human has reviewed this work yet.**

## 1. Data audit

| Tab | Grain / period | Result |
|---|---|---|
| Locations_Master | 122 stores (110 Established, 12 "New (East Coast launch)") | Clean. **Issue:** `state` holds a region code, not a real state (Salt Lake City, Tempe and Las Vegas are "CO"; Portland and Bend are "WA"; Boston, Hoboken and Stamford are "NY"). `months_open` doesn't line up with any single as-of date. Neither field is used in calculations. |
| Marketplace_Payouts | 17,934 rows, weekly × store × platform | No nulls, no duplicate keys, no negative values. Every row reconciles: gross − 6 deductions = net within $0.03. **51 weeks** (Jul 7 2025 – Jun 22 2026), not 52. 10 East Coast stores open mid-window. |
| Channel_Revenue_Summary | 1,408 rows, monthly × store | Components sum to total. **The 3PD gross column equals payout gross exactly by store and month**, so it is the same data and not an independent check. Months contain 4 or 5 payout weeks (Sep, Dec and Mar have 5), so monthly figures must be normalized per week. |
| Cohort_Retention | 2,156 rows, 44 stores, Nov-25 to May-26 cohorts | Retention falls from 30d to 90d in every row. Grubhub has no promo cohort. **The definition is not stated:** "retention" could mean a return to Sola through any channel, or a return to the same platform. |
| Daypart_Patterns | 6 regions × 4 dayparts | **Issue:** order shares sum to 0.92–1.03 per region and platform, not 1.00. The model normalizes them. |
| Markup_Test | 32 DoorDash stores, 8 per arm at 0/5/10/15% markup | Arms are contiguous blocks (checked). The test is DoorDash-only, and the README describes it as noisy. |
| Contract_Terms | 9 terms | All values are stored as text ("27.0%", "~$39.5M"). They were re-entered as numeric inputs and checked against the payouts: DoorDash commission = 27.0%, marketing commission = 1.5%, total marketing spend = $2.93M (contract: "~$2.9M"), ad share = 27.4% (contract: "~27%"), DoorDash ad share = 31.0% (contract: "~30%"). |

## 2. Calculations checked

- **P&L reconciliation.** Computed net revenue is within $1 of the source `net_bank_deposit` total over 51 weeks; the gap comes from cent rounding in the source rows. CM2 matches an independent recompute from the source net within $1.48.
- **Segments sum to the total.** Platforms, regions, market status, and platform × status all sum to total gross and CM2, with zero difference.
- **Fee definitions and signs.** The six deductions match the README definition. Tax is excluded (memo only, about $6.3M a year). Promo and ads are each counted once (no double counting).
- **Observed commission rates match the contract:** DoorDash 27.0%, UberEats 30.5%, Grubhub 22.5%.
- **Annualization.** FY figures = 51-week actuals × 52/51, using 13 periods × 28 days. The exit run-rate is the latest 4-week period × 13. The Periods sheet sums to 51 weeks, and the latest period has 4 weeks.
- **DoorDash scenarios.** Year 1 uses the model-derived $40.6M of DoorDash gross. This is 2.8% above the contract's "~$39.5M", which matches the 51-week actual of $39.8M. Year 2 adds 8% growth. 24-month cost = Y1 + Y2.
- **Weighted averages.** Cohort retention is weighted by new customers. Daypart shares are weighted by actual orders per region and platform. Same-store trends are per store-week, weighted by store-weeks.
- **Deck vs workbook.** Every chart and number in the PDF is read from the recalculated workbook (`build_scripts/charts.py`). The dollar figures in the PDF were extracted and checked against the model.

## 3. Key findings (with sheet references)

1. **Standalone, the channel makes money.** FY CM2 is $13.7M on $74.4M of gross (18.5%), against a 48% dine-in benchmark. The all-in take is 35.5%. For the channel to beat dine-in, at least **62%** of its sales must be truly incremental (`PL_Waterfall`).
2. **Restaurant-funded promo rose from 1.6% to 7.0% of gross after Feb 9 2026.** It happened on every platform and in every region at once, with no volume response: established gross per store-week is −2.7% vs late January. At the exit run-rate this costs **$4.4M a year, 2.7× the DoorDash ask**. Exit CM2 is 14.4% (`Promo_Escalation`, `Weeks`).
3. **DoorDash.** Accepting 31% costs $1.6M a year and $3.4M over 24 months, or up to $5.6M if the ads mandate forces a top-up. The counter (28.5%) costs $1.3M over 24 months. Walking costs $3.2M a year under base recapture assumptions (`DoorDash`).
4. **Margin variation comes from platform and maturity, not geography.** UberEats runs at 14.2% CM2 (8.9% at exit) and the East Coast at 13.5% (8.4% at exit). All established regions are about 18.7% (`Platform_Market`).

## 4. Assumptions (most material first)

| Assumption | Type | Value | Sensitivity |
|---|---|---|---|
| Incrementality of 3PD sales | Analyst | Bear 35% / Base 55% / Bull 80%. Evidence-derived estimate is 52%. | Economic value ranges from −$9.5M to +$6.6M a year (`Sensitivity` §2) |
| Promo rollback volume loss | Analyst | 3% | 0–15% → rollback value of $4.4M to $2.0M (`Sensitivity` §4) |
| Markup response extends to UberEats/Grubhub, and haircut | Analyst | Yes; 50% haircut | Test-implied $3.3M; mechanical cross-check $2.3M; banked $1.7M |
| Walk recapture | Analyst | 30% to UberEats/Grubhub, 15% to own channels | Y1 walk cost ranges from −$4.7M to +$8.1M across the grid |
| COGS 30% / labor 16% / dine-in 48% | Assignment | — | CM2% sensitivity grid (`Sensitivity` §1) |
| Cannibalized sales would earn 48% (pickup and 1P treated like dine-in) | Analyst | — | Grid in `Sensitivity` §2 |
| COGS applies to the un-marked-up price | Analyst | — | Raises the measured markup gain |

## 5. Limitations

- **Incrementality is inferred, not measured.** The 52% estimate assumes that, without 3PD, in-store sales would have grown like pickup, 1P and catering. With only one year of data, seasonality cannot be separated from trend. The store-level correlation between 3PD growth and in-store growth is **+0.13**, which contradicts store-level cannibalization. The estimate is treated as indicative only.
- **The promo-ramp conclusion rests on time-series evidence.** Dine-in also fell about 10% per store-week from February to June, partly seasonal. One could argue the promo protected 3PD against that decline. The counter-evidence is that the 3PD-to-dine-in ratio grew more slowly after the ramp (1.5% per month) than before it (2.6% per month).
- **Markup test.** Only the 5% arm is close to significant (t≈2.1). Net revenue per order rose more than mechanically plausible (+7.4% on a 5% price increase), which suggests noise. That is why the mechanical cross-check and the 50% haircut are used.
- **Plan totals ignore interactions** between actions, for example markup applied on a lower-promo base.
- **Mismatched bases.** The DoorDash contract impact uses FY-basis gross, while the plan is measured against the exit run-rate.
- **1P delivery economics are not in the data.** The model therefore cannot say whether shifting customers to 1P is margin-accretive.

## 6. Unresolved questions to raise with management

1. What drove the February 2026 promo ramp? Was it a deliberate campaign, a platform auto-enrollment, or a classification change in the payout feed? **This is the highest-priority check, because the recommendation's largest lever depends on it.**
2. Does cohort "retention" mean a return to Sola or a return to the platform?
3. What does "walk" mean contractually? Does DoorDash revert to rack rates, or delist?
4. Is there a pre-July 2025 history that would let seasonality be separated from trend?
5. What are 1P delivery's unit economics (driver cost, fees)?
6. Why don't the daypart order shares sum to 100%?
