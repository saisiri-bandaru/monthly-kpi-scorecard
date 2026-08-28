# Monthly ops and finance KPI scorecard

Design note for a one-page scorecard operators can run in standup. Same fictional company as the rest of the profile (**Northline**).

The live Excel file is not in this repo yet. Working files to open instead:

- [FP&A variance dashboard](https://github.com/saisiri-bandaru/fpna-variance-dashboard)
- [Cost accounting workbook](https://github.com/saisiri-bandaru/cost-accounting-workbook)

## Business question

Which five numbers does the plant or shared-services lead need this week — and which ones are actually off plan?

A scorecard that lists 40 KPIs is a data dump. This design keeps a short stack with an owner and a RAG rule.

## Sample card (illustrative July)

| KPI | Owner | Target | Actual | Var | RAG |
| --- | --- | ---: | ---: | ---: | --- |
| Units produced | Ops | 100,000 | 97,200 | −2.8% | Amber |
| COGS % sales | Plant controller | 60.0% | 61.4% | +140 bp | Red |
| OT hours / total hours | HR / scheduling | 6.0% | 7.1% | +110 bp | Amber |
| On-time ship | Warehouse | 98.0% | 98.4% | +40 bp | Green |
| OpEx vs budget | FP&A | $0 | +$40k | Unfav | Amber |

RAG rule used here: green within 1% or 20 bp of target, amber inside 3% / 100 bp, red beyond that. Pick one rule and do not change it mid-year.

## Design rules I use

- Every KPI has one owner, not a committee
- Actual and target use the same definition as the variance pack
- Red requires a one-line comment, not a color
- Finance KPIs and ops KPIs sit on the same page so the meeting does not split
