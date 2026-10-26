# Shipping Data Analysis

An Excel-based sales and shipping analysis of **1,039 order lines** from an Australian office-products business (Feb 2013 – Feb 2017). The repository contains the cleaned source data and a pivot-table and dashboard workbook that answers questions about sales by product, region, customer segment, and account manager.

![Excel](https://img.shields.io/badge/built%20with-Microsoft%20Excel-217346?logo=microsoftexcel&logoColor=white)
![Data](https://img.shields.io/badge/data-1%2C039%20order%20lines-blue)
![Period](https://img.shields.io/badge/period-2013--2017-lightgrey)

---

## Table of Contents

- [Overview](#overview)
- [Headline Findings](#headline-findings)
- [Repository Structure](#repository-structure)
- [Dataset](#dataset)
  - [Data Dictionary](#data-dictionary)
  - [Calculated Columns](#calculated-columns)
  - [Named Ranges](#named-ranges)
- [Analysis Workbook](#analysis-workbook)
  - [Sheet1: Category and State Pivots](#sheet1-category-and-state-pivots)
  - [Sheet2: Customer Type and Account Manager Pivots](#sheet2-customer-type-and-account-manager-pivots)
  - [Sales Dash: Dashboard](#sales-dash-dashboard)
- [Getting Started](#getting-started)
- [Methodology](#methodology)
- [Data Quality Notes and Limitations](#data-quality-notes-and-limitations)
- [Ideas for Extension](#ideas-for-extension)
- [Author](#author)
- [License](#license)

---

## Overview

The project takes a flat table of customer orders and turns it into an interactive sales view, entirely in Excel using structured tables, named ranges, formulas, pivot tables, charts, and sparklines.

**Questions the analysis answers**

1. Which product categories drive revenue?
2. How do sales and order volumes differ across states (NSW, VIC, WA)?
3. Which customer segments (Corporate, Small Business, Home Office, Consumer) are most valuable?
4. How is each account manager performing, year over year?
5. How do orders break down by priority, ship mode, and order size?

**Skills demonstrated:** Excel Tables, calculated columns, named ranges, `SUMIFS` / `COUNTIFS` / `COUNTA` / `COUNTBLANK`, PivotTables, PivotCharts-style dashboard charts (pie, clustered column, combo), and in-cell column sparklines.

---

## Headline Findings

All figures are computed from the `Total` column (order total after discount, plus shipping). Currency is shown as `$` in the workbooks; the customer cities (Sydney, Melbourne, Perth) suggest Australian dollars, though the source does not state it.

| Metric | Value |
|---|---|
| Total sales | **$1,138,807** across 1,039 order lines (1,035 unique order numbers) |
| Average vs median order line | $1,096 vs $160. A few very large orders skew the average; the top 5 lines alone account for about 11% of sales |
| Top category | **Technology: $784,366 (68.9%)** from only 207 lines. Office Supplies is 28.5% from 805 lines; Furniture is 2.7% from 27 lines |
| Top state | **NSW: $768,723 (67.5%)**, then VIC $283,641 (24.9%) and WA $86,443 (7.6%) |
| Top customer segment | **Corporate: $389,261 (34.2%)**. Small Business has the highest average line value ($1,290) |
| Top account managers | Yvette Biti ($148,147), Connor Betts ($135,494), Natasha Song ($120,790) |
| Lowest account managers | Stevie Bacata ($6,771 from 9 lines) and Preston Senome ($18,350) |
| Sales trend | $171,751 (2013) → $319,232 (2014, +85.9%) → $352,763 (2015) → $272,289 (2016, **−22.8%**). 2017 covers only 7 Feb and is not comparable |
| Shipping | Regular Air carries 83.9% of lines. Delivery Truck is only 2.8% of lines but averages **$40.61** shipping per line vs about $5 for air. Shipping is only 0.55% of total sales |
| Dispatch speed | Median of **2 days** from order to ship date. 120 lines shipped same-day; 15 took more than 7 days (max 27) |
| Large orders | 238 lines have a quantity above 40 |

---

## Repository Structure

```
Shipping-Data-Analysis/
├── Shipping_Data.xlsx            # Source data: the "Orders" sheet with an Excel Table named "sales"
├── Shipping_Data_Analysis.xlsx   # Analysis: pivot tables (Sheet1, Sheet2) and the "Sales Dash" dashboard
└── README.md
```

| File | Purpose | Sheets |
|---|---|---|
| `Shipping_Data.xlsx` | Cleaned transactional data and calculated columns | `Orders` |
| `Shipping_Data_Analysis.xlsx` | Pivot tables, KPI tables, and dashboard charts built on the data | `Sheet1`, `Sheet2`, `Sales Dash` |

---

## Dataset

**File:** `Shipping_Data.xlsx` → sheet `Orders`

- **Grain:** one row per product line on an order (1,039 rows, 25 columns, header on row 5, data in `A6:Y1044`)
- **Table name:** `sales`
- **Date range:** orders from 11 Feb 2013 to 7 Feb 2017
- **Coverage:** 538 customers, 147 products, 14 account managers, 3 cities/states

A small summary strip above the table (row 3) uses formulas to show **Orders** (`=COUNTA(Order_No)`), **Missing Priorities** (`=COUNTBLANK(Order_Priority)`), and **Shipped** (`=COUNT(Ship_Date)`), currently 1,039 / 2 / 1,037.

### Data Dictionary

| # | Column | Type | Description |
|---|---|---|---|
| A | Order No | Text | Order identifier, e.g. `5014-1`. An order number can span several lines |
| B | Order Date | Date | Date the order was placed |
| C | Order Year | Calculated | Year of the order date (`=TEXT(B6,"yyyy")`) |
| D | Customer Name | Text | Customer who placed the order |
| E | Address | Text | Delivery address |
| F | City | Text | Sydney, Melbourne, or Perth |
| G | State | Text | NSW, VIC, or WA (maps one-to-one to city) |
| H | Customer Type | Category | Corporate, Home Office, Small Business, or Consumer |
| I | Account Manager | Text | Salesperson responsible (14 people) |
| J | Order Priority | Category | Critical, High, Medium, Low, or Not Specified |
| K | Product Name | Text | Product description (147 distinct) |
| L | Product Category | Category | Office Supplies, Technology, or Furniture |
| M | Product Container | Category | Small Box, Wrap Bag, Small Pack, Jumbo Drum, Large Box, or Medium Box |
| N | Ship Mode | Category | Regular Air, Express Air, or Delivery Truck |
| O | Ship Date | Date | Date the order shipped (blank if not yet shipped) |
| P | Cost Price | Currency | Unit cost to the business |
| Q | Retail Price | Currency | Unit selling price |
| R | Profit Margin | Calculated | **Per-unit profit in dollars**, `Retail Price − Cost Price` (not a percentage) |
| S | Order Quantity | Integer | Units ordered (1–50) |
| T | Sub Total | Calculated | `Retail Price × Order Quantity` |
| U | Discount % | Percent | Discount rate, 0%–10% |
| V | Discount $ | Calculated | `Sub Total × Discount %` |
| W | Order Total | Calculated | `Sub Total − Discount $` |
| X | Shipping Cost | Currency | Shipping charge for the line |
| Y | Total | Calculated | `Order Total + Shipping Cost`. This is the sales figure used throughout the analysis |

### Calculated Columns

Six columns are live Excel formulas inside the table, so they recalculate if inputs change:

```
Order Year   = TEXT([Order Date], "yyyy")
Profit Margin = Retail Price − Cost Price
Sub Total     = Retail Price × Order Quantity
Discount $    = Sub Total × Discount %
Order Total   = Sub Total − Discount $
Total         = Order Total + Shipping Cost
```

### Named Ranges

Every column has a workbook-level named range (for example `Total`, `State`, `Customer_Type`, `Account_Manager`, `Order_Year`, `Order_Quantity`) pointing at rows 6–1044. This lets the dashboard use readable formulas such as `=SUMIFS(Total, State, A5)` instead of cell references.

---

## Analysis Workbook

**File:** `Shipping_Data_Analysis.xlsx`

### Sheet1: Category and State Pivots

| Pivot | Rows | Values | Shows |
|---|---|---|---|
| PivotTable1 (`A3:B7`) | Product Category | Sum of Total | Revenue by category |
| PivotTable2 (`A16:C35`) | State → Order Year | Count of Total, Count of Account Manager | Order-line volume by state and year |

### Sheet2: Customer Type and Account Manager Pivots

| Pivot | Rows | Columns | Values | Shows |
|---|---|---|---|---|
| PivotTable1 (`A3:B8`) | Customer Type | n/a | Sum of Total | Revenue by customer segment |
| PivotTable2 (`A10:G97`) | Account Manager → Order Priority | Order Year | Sum of Total | Each manager's sales by priority level, by year |

### Sales Dash: Dashboard

A formula-driven dashboard. The tables feeding it use named-range `SUMIFS` / `COUNTIFS`, so every number traces back to the `Orders` data.

**KPI tables**

| Range | Content |
|---|---|
| `A4:C7` | Sales (`Total`) and order-line count by **state** |
| `A10:C14` | Sales and order-line count by **customer type** |
| `A17:B17` | Count of lines with **Order Quantity over 40** (`=COUNTIFS(Order_Quantity,">40")`) |
| `A20:G34` | **Sales by account manager**: overall total, 2013–2016 yearly columns, and a column-chart **sparkline** per manager showing the trend |
| `A37:E42` | **Orders by year and state** (NSW, VIC, WA), plus total sales per year |

**Charts**

| Chart | Type | Source |
|---|---|---|
| Sales by Account Manager | Clustered column | `A21:B34` |
| Orders by Customer Type | Pie | `A11:C14` |
| Orders by State | Pie | `A5:C7` |
| Sales Trends (×2) | Column + line combo | `A37:C41` (orders by year for NSW and VIC) |

---

## Getting Started

### Prerequisites

- **Microsoft Excel** 2016 or later (Excel 365 recommended). Sparklines and pivot features are Excel-native; LibreOffice and Google Sheets may not render them faithfully.

### Open the files

1. Clone or download the repository:
   ```bash
   git clone https://github.com/AshishSiwach/Shipping-Data-Analysis.git
   cd Shipping-Data-Analysis
   ```
2. **Keep both `.xlsx` files in the same folder.**
3. Open `Shipping_Data_Analysis.xlsx` and start on the **Sales Dash** sheet.

### Re-linking the data source

The dashboard formulas and named ranges in `Shipping_Data_Analysis.xlsx` reference `Shipping_Data.xlsx` as an **external workbook**, and the link was saved with the original author's path (`C:\Data Science and Analysis\Shipping_Data.xlsx`). On your machine:

1. Excel will show a security warning about links. Choose **Update** (or **Enable Content**) if you want live values.
2. If Excel cannot find the file, go to **Data → Queries & Connections → Edit Links to Files**, select `Shipping_Data.xlsx`, click **Change Source…**, and browse to your local copy.
3. Until the link is repointed, the workbook still displays the last saved values.

The pivot tables carry a stored snapshot of the 1,039 records, so they display correctly without the link.

### Explore the source data

Open `Shipping_Data.xlsx` and use the filter buttons on the `sales` table. Add your own pivots with **Insert → PivotTable** and select the table name `sales`.

---

## Methodology

1. **Structure the data.** The raw orders are held in an Excel Table (`sales`), so formulas and pivots expand automatically as rows are added.
2. **Derive financial fields.** Sub total, discount, order total, and total (including shipping) are computed as formulas rather than pasted values.
3. **Name the columns.** Named ranges make dashboard formulas self-documenting and keep them independent of cell addresses.
4. **Summarise with pivots.** PivotTables give the category, segment, state-by-year, and manager-by-priority views.
5. **Build the dashboard.** `SUMIFS` / `COUNTIFS` tables sit behind each chart, and sparklines add a compact per-manager trend.
6. **Define sales consistently.** "Sales" always means the `Total` column (net of discount, including shipping) unless a chart is explicitly labelled as order counts.

---

## Data Quality Notes and Limitations

- **2017 is a partial year.** Data ends on 7 February 2017 (25 lines). Do not compare 2017 with full years. The dashboard's per-manager yearly columns cover 2013–2016 only, so they sum to slightly less than each manager's overall total.
- **"Orders" means order lines.** The dashboard counts rows, not distinct order numbers (1,039 lines vs 1,035 order numbers; four order numbers contain two lines).
- **Missing values.** Two lines have no `Order Priority` and two recent lines have no `Ship Date` (not yet shipped). A further 213 lines are labelled `Not Specified` for priority.
- **`Profit Margin` is a per-unit dollar amount**, not a percentage margin. Total profit per line is `Profit Margin × Order Quantity`.
- **Broken named ranges.** `Order_Month` and `Ship_Time` are defined in both workbooks but point to `#REF!`, as the columns they referred to no longer exist. They are unused by any formula. `Ship_Time` is a natural column to restore (`Ship Date − Order Date`).
- **Hard-coded dimension lists.** The dashboard's state, customer type, and account manager labels are typed in, so a new category in the data will not appear automatically.
- **Single currency, no tax.** Currency is not stated in the source, and the data has no tax field.
- **Data provenance.** The source of the underlying orders is not documented in the repository. Treat the dataset as illustrative.

---

## Ideas for Extension

- Add a `Ship Time` calculated column and analyse delivery performance by ship mode and priority
- Report **profit** (not just revenue) by category, state, and account manager using `Profit Margin × Order Quantity`
- Add slicers and a timeline to the pivots for interactive filtering
- Replace hard-coded dashboard labels with pivot-driven or dynamic-array lists
- Add a monthly sales trend and year-over-year growth view
- Investigate the Delivery Truck shipping cost premium and the Technology revenue concentration
- Recreate the analysis in Python/pandas or Power BI for comparison

---

## Author

**Ashish Siwach**
GitHub: [@AshishSiwach](https://github.com/AshishSiwach)

---

## License

No license file is currently included in this repository, which means all rights are reserved by default. If you want others to reuse the data or analysis, add a license such as [MIT](https://choosealicense.com/licenses/mit/) (code and workbooks) or [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/) (data and documentation).

