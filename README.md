# Coca-Cola Sales Analysis Dashboard

An end-to-end Excel analysis of one full year (2021) of Coca-Cola sales across four American retailers. The project combines data validation, native pivot tables, formula-driven chart sheets and an interactive dashboard, all in a single workbook.

![Excel](https://img.shields.io/badge/Built%20with-Microsoft%20Excel-217346?logo=microsoftexcel&logoColor=white)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen)
![Formula errors](https://img.shields.io/badge/Formula%20errors-0-brightgreen)

---

## Table of Contents

- [Project Overview](#project-overview)
- [Business Questions](#business-questions)
- [Dataset](#dataset)
- [Workbook Structure](#workbook-structure)
- [Data Preparation and Methodology](#data-preparation-and-methodology)
- [Key Findings](#key-findings)
- [Interactive Dashboard](#interactive-dashboard)
- [Recommendations](#recommendations)
- [Limitations](#limitations)
- [Validation](#validation)
- [Excel Techniques Used](#excel-techniques-used)
- [How to Use](#how-to-use)
- [Future Improvements](#future-improvements)

---

## Project Overview

This project analyses **3,888 invoice lines** covering **50 states, 5 regions, 6 beverage brands and 4 retailers** (Sodapop, FizzySip, BevCo, DreamCo). The deliverable is a 14-sheet Excel workbook:

| Layer | Sheets | Purpose |
|---|---|---|
| Data | `Data` | Cleaned single source of truth |
| Exploration | `PT1` to `PT6` | Six native pivot tables with report filters |
| Analysis | `CH1` to `CH6` | Six chart sheets (24 charts) driven by formulas |
| Summary | `Dashboard` | One-page interactive view with KPIs, charts, heatmaps and live insights |

**Headline numbers**

| Metric | Value |
|---|---|
| Total sales | $8,684,028 |
| Operating profit | $3,173,632 |
| Operating margin | 36.5% |
| Units sold | 17,148,250 |
| Average price per unit | $0.506 |

---

## Business Questions

1. How do sales vary by retailer, region and time?
2. How do price, volume and profitability relate to each other?
3. Where are operating margins strong, and where are they weak?

---

## Dataset

One Excel table of invoice lines. Each row is one beverage brand sold by one retailer in one city on one invoice date.

| Column | Description |
|---|---|
| Retailer, Retailer ID | Retailer name and identifier |
| Invoice Date | Date of the invoice (dd-mmm-yyyy) |
| Region, State, City | Location |
| Beverage Brand | Coca-Cola, Dasani Water, Diet Coke, Sprite, Powerade, Fanta |
| Price per Unit | Unit price |
| Units Sold | Quantity sold |
| Total Sales | `Price per Unit x Units Sold` (formula) |
| Operating Profit | `Total Sales x Operating Margin` (formula) |
| Operating Margin | Margin for the line |
| Month | `MONTH(Invoice Date)` (formula) |

---

## Workbook Structure

### Layer 1: Pivot tables (exploration)

Native Excel pivot tables, each with two report filters and a companion formula table showing share of sales and correctly weighted margin.

| Sheet | Rows | Values | Question answered |
|---|---|---|---|
| `PT1_Retailer` | Retailer | Sales, Units, Profit | Which retailer is biggest and most profitable? |
| `PT2_Region` | Region | Sales, Units, Profit | Which regions drive sales and profit? |
| `PT3_Brand` | Brand | Sales, Units, Profit, Avg Price | Which brands lead, and at what price? |
| `PT4_Monthly` | Month | Sales, Units, Profit | What does seasonality look like? |
| `PT5_Region_x_Retailer` | Region x Retailer | Sales | Which retailer sells where? |
| `PT6_States` | State | Sales, Profit | Which states are top and bottom performers? |

### Layer 2: Chart sheets (formula-driven)

Each sheet has a title banner, a live-formula data table, a Key Insights box and a 2x2 grid of charts.

| Sheet | Focus | Highlights |
|---|---|---|
| `CH1_Retailer` | Retailer comparison | Sales vs profit, margin, units vs price combo, share doughnut |
| `CH2_Monthly` | Trends and seasonality | Monthly and quarterly tables, Retailer x Month heatmap |
| `CH3_Region` | Regional performance | Stacked sales by retailer, Region x Month and Region x Brand heatmaps |
| `CH4_Brand` | Brand mix and profitability | Brand x Retailer margin heatmap |
| `CH5_Price_Volume` | Price, volume, profitability | Seven editable price bands, CORREL block, scatter plots with trend lines |
| `CH6_States` | Leaders and laggards | Dynamic Top 10 / Bottom 10 rankings across all 50 states |

### Layer 3: Dashboard

A one-page manager view. See [Interactive Dashboard](#interactive-dashboard).

---

## Data Preparation and Methodology

**Quality checks**

- **Completeness:** no blank cells and no duplicate rows.
- **Consistency:** Total Sales, Operating Profit and Month were recomputed from their underlying columns and reconciled exactly.
- **Cleanup:** stray marker cells and scratch cells beside the table were removed so they could not leak into a pivot.
- **Zero-unit rows:** two Fanta lines in Omaha, Nebraska (March and April) with zero units were deliberately kept, as they look like genuine stock-outs and removing them would hide an operational signal.

**Margin definition**

Overall margin is calculated as **total profit divided by total sales**, not the average of each line's margin. A simple average would give a $500 line the same weight as a $5,000 line and misstate the true margin. This definition is applied consistently across the whole workbook.

```
Overall margin = SUM(Operating Profit) / SUM(Total Sales)
               = $3,173,632 / $8,684,028 = 36.5%
```

**Named ranges**

Every data column has a named range (`d_Sales`, `d_Profit`, `d_Units`, `d_Region`, `d_Retailer`, `d_Brand`, `d_State`, `d_Month`, `d_Price`), so formulas read like `SUMIFS(d_Sales, d_Region, "West")` and are easy to audit and maintain.

---

## Key Findings

### Size is not profitability: retailers

| Retailer | Sales | Share | Operating Profit | Margin | Avg Price/Unit |
|---|---|---|---|---|---|
| Sodapop | $4,403,630 | 50.7% | $1,647,936 | 37.4% | $0.481 |
| FizzySip | $2,584,450 | 29.8% | $835,759 | 32.3% | $0.629 |
| BevCo | $1,291,535 | 14.9% | $542,046 | 42.0% | $0.421 |
| DreamCo | $404,413 | 4.7% | $147,890 | 36.6% | $0.493 |

- **Sodapop** is the volume and profit anchor.
- **FizzySip** charges the highest prices but earns the lowest margin. Matching Sodapop's 37.4% margin on its current sales would add roughly **$130,000** in profit (illustrative, not a forecast).
- **BevCo** is small but the most efficient.
- **DreamCo** is the clearest room for growth.

### Regions

| Region | Sales | Share | Margin | Avg Price/Unit |
|---|---|---|---|---|
| West | $2,835,100 | 32.6% | 32.6% | $0.616 |
| Northeast | $1,788,348 | 20.6% | 35.2% | $0.527 |
| Southeast | $1,618,935 | 18.6% | 39.6% | $0.490 |
| South | $1,291,535 | 14.9% | 42.0% | $0.421 |
| Midwest | $1,150,110 | 13.2% | 38.0% | $0.413 |

The two highest-priced regions (West, Northeast) have the two lowest margins. Retailers barely overlap geographically (BevCo sells only in the South, FizzySip only in the West, accounting for about 91% of West sales), so **retailer and region effects are mixed together** in this data.

### Seasonality

- Sales sit around **$490K per month** from January to April, ramp sharply through May and June, peak in **July ($1.04M)**, dip in autumn, and peak again in **December ($1.05M)**.
- March is the weakest month; the peak-to-trough swing is about **2.2x**.
- Quarterly sales: Q1 $1.48M, Q2 $2.07M, Q3 $2.67M, Q4 $2.46M.
- **Margin stays flat** (about 35.8% to 37.0%) while average price rises from $0.38 in January to $0.62 in December. The seasonal lift is driven by volume and price, not better efficiency.

### Price, volume and profitability

| Price per Unit | Invoice Lines | Avg Units/Line | Total Sales | Margin |
|---|---|---|---|---|
| Under $0.30 | 301 | 3,586 | $235,853 | 37.8% |
| $0.30 to $0.39 | 573 | 3,811 | $721,085 | 38.4% |
| $0.40 to $0.49 | 957 | 4,087 | $1,673,198 | 37.4% |
| $0.50 to $0.59 | 1,029 | 4,551 | $2,453,508 | 36.7% |
| $0.60 to $0.69 | 667 | 5,002 | $2,074,948 | 37.3% |
| $0.70 to $0.79 | 225 | 5,087 | $820,875 | 35.6% |
| $0.80 and above | 136 | 5,960 | $704,563 | 30.6% |

| Correlation (all 3,888 lines) | Value | Strength |
|---|---|---|
| Price vs Units | +0.25 | Moderate |
| Price vs Total Sales | +0.65 | Strong |
| Price vs Operating Margin | -0.13 | Weak |

> **Note:** this is correlation, not causation. High prices cluster in the West, in summer and in December, where volumes are also higher. Premium pricing is associated with more revenue, but not a better profit rate.

### Brands

| Brand | Sales | Share | Margin |
|---|---|---|---|
| Coca-Cola | $2,015,890 | 23.2% | 39.3% |
| Dasani Water | $1,725,838 | 19.9% | 38.0% |
| Diet Coke | $1,481,425 | 17.1% | 33.4% |
| Sprite | $1,235,588 | 14.2% | 34.8% |
| Powerade | $1,193,638 | 13.7% | 35.8% |
| Fanta | $1,031,650 | 11.9% | 36.1% |

Dasani earns about 51% margin at BevCo but only about 22% at DreamCo, which suggests different pricing or cost conditions by account.

### States

- **Top by sales:** New York ($582,675), California ($582,400), Florida ($561,850), Texas ($440,988). The top four generate about 25% of sales.
- **Best margins:** Alabama (51.5%), Tennessee (46.4%), Virginia (43.6%).
- **Weakest margins:** Alaska (25.6%), Washington (25.8%), Rhode Island (28.3%).
- **California** is the only top-four state with a weak margin (31.6%).

---

## Interactive Dashboard

The final sheet is built for a manager with 30 seconds.

1. **Filter dropdowns:** Retailer, Region and Brand (Data Validation, default "All"). They behave like slicers.
2. **KPI cards:** Total Sales, Units Sold, Operating Profit, Operating Margin, Average Price per Unit, Peak Month.
3. **Charts:** monthly sales with margin line, brand share doughnut, sales by retailer, sales vs profit by region, top 10 states.
4. **Heatmaps:** Region x Month sales and Region x Brand operating margin.
5. **Live insights:** five formula-generated sentences that update with the filters.

**How the filters work.** A hidden "engine" area converts each dropdown into a criterion: if the dropdown says `All`, the criterion is the wildcard `"*"`, otherwise it is the selected value. Every formula is a `SUMIFS` using all three criteria.

**Edge cases.** If a combination has no sales (for example BevCo in the West), the KPIs show zero and the insight line says there are no sales for that combination. Nothing shows an error.

---

## Recommendations

1. **Review FizzySip's pricing and cost structure.** Highest prices but lowest margin points to cost-to-serve or discounting issues.
2. **Protect and learn from the margin leaders** (BevCo and the South), and test whether their practices can be replicated elsewhere.
3. **Plan around seasonality.** Align inventory, promotions and staffing to the May to July ramp and the December peak, and use February to April for activity that does not depend on volume.
4. **Investigate Diet Coke and low-margin West states** such as Alaska and Washington, where high prices are not converting to profit.

---

## Limitations

- **Single year of data:** trend cannot be separated from seasonality, and year-over-year growth cannot be measured.
- **Correlation only:** price relationships are not causal effects.
- **Retailer and region overlap:** retailers barely overlap by region, so their effects are entangled.
- **No cost data:** margins are visible, but the reasons they differ are not.
- **Pivot-adjacent tables and insight text:** companion margin tables and written insights on the pivot and chart sheets use the full dataset and do not react to pivot filters. Only the Dashboard insights are live formulas.

---

## Validation

- Every headline number (retailers, regions, brands, months, quarters, price bands, correlations, state rankings) was recomputed independently in **Python (pandas)** and matched the workbook exactly.
- The workbook was recalculated with **zero formula errors** across about 13,000 formulas.
- Dashboard filters were tested, including a single-retailer view (FizzySip: $2,584,450 sales, 32.3% margin) and an empty combination.
- Pivot-table files were validated against the Office XML schema.

---

## Excel Techniques Used

- Excel Tables and named ranges
- Pivot tables with report filters
- `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`
- `INDEX` / `MATCH` with `LARGE` and `SMALL` for dynamic top-N and bottom-N rankings (compatible with older Excel versions)
- `CORREL` for correlation analysis
- `IFERROR` and `IF` guards against divide-by-zero errors
- Wildcard criteria (`"*"`) for an "All" option in `SUMIFS`
- Data Validation dropdowns
- Conditional formatting: colour scales (heatmaps) and data bars
- Combo charts with secondary axes, doughnut charts, scatter plots with trend lines
- `TEXT` function inside formulas for live insight sentences

---

## How to Use

1. Download the workbook from this repository and open it in Microsoft Excel. If prompted, click **Enable Editing** so the pivot tables can refresh on open.
2. Go to the **`Dashboard`** sheet (the last tab) and use the Retailer, Region and Brand dropdowns to filter everything.
3. Open the **`PT1` to `PT6`** sheets to explore the pivots, or the **`CH1` to `CH6`** sheets for the detailed chart analysis.

**Adding new data**

1. Paste new rows into the `Data` table.
2. The formulas currently cover Data rows 6 to 3893, so extend the named ranges to the new last row.
3. Refresh the pivot tables (Data > Refresh All). Everything else recalculates automatically.

**Optional:** native slicers can be added to any pivot sheet via *Insert > Slicer*. The Dashboard uses dropdowns instead because slicers cannot drive formula-based charts.

---

## Future Improvements

- Add prior-year data for year-over-year growth analysis
- Add cost drivers to explain low-margin states
- Run a price-sensitivity test within a single retailer and region to control for confounding factors
- Automate the monthly data refresh

---

## Author

**Vaibhav Nagdeve**
[LinkedIn](www.linkedin.com/in/vaibhav-nagdeve) | [Email](mailto:vnagdeve6@gmail.com)
