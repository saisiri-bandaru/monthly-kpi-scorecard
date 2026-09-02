# Monthly KPI scorecard

Ops + finance RAG scorecard for a fictional CPG company (**Northline Consumer Products**): twelve monthly KPIs, actual vs target, YTD, trends, and a one-page dashboard.

**File to open:** `Northline_Monthly_KPI_Scorecard.xlsx`

## What you will see

- KPI register with targets, direction (higher-better / lower-better / range), and RAG bands
- Monthly actuals for revenue, margin, EBITDA, volume, OTIF, scrap, DSO, CCC, headcount, attrition, cash, SG&A
- Scorecard: as-of actual vs target, variance, Green / Amber / Red — all formulas
- A 12-month RAG heatmap, MoM / rolling-3 / YTD / YoY trends
- A one-pager with traffic-light tiles and charts

Yellow cells with blue font are inputs. Black font is formulas.

## How to use

1. Open `01_Assumptions` and change the yellow cells (targets, RAG bands, reporting month).
2. Type or paste monthly actuals on `02_Actuals`.
3. Read actual vs target, variance, and RAG on `03_Scorecard`.
4. Read MoM, rolling 3-month, YTD, and YoY on `04_Trends`.
5. Use `05_Dashboard` as the one-pager.

## Tabs

| Tab | Role |
| --- | --- |
| `00_Cover` | Purpose and how to use |
| `01_Assumptions` | KPI targets, RAG thresholds, reporting month |
| `02_Actuals` | Monthly actuals (yellow inputs) |
| `03_Scorecard` | Actual vs target, variance, RAG (formulas) |
| `04_Trends` | MoM, rolling 3-month, YTD, YoY |
| `05_Dashboard` | One-pager: RAG tiles and charts |
| `06_Data_Dictionary` | Field definitions |

## Stack

Excel (formulas only — no VBA). Built so another analyst can inherit the file from the data dictionary.

## Not included on purpose

- Live data warehouse pulls
- Confidential employer data
- A full P&L or three-statement pack (this is the *monthly operating scorecard*)

All sample numbers are fictional.

## Profile

Sai Siri Bandaru — Financial Analyst | FP&A | forecasting, variance analysis, Excel
