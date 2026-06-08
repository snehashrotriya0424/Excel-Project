# 📊 E-Commerce Sales Analytics Dashboard — FY 2024

> A hiring-level Microsoft Excel analytics project demonstrating end-to-end data analysis across 1,913 transactions, 10 product categories, 5 regions, and 4 sales channels — built entirely with advanced Excel formulas, dynamic dashboards, and professional data visualization.

---

## 📁 File

```
Ecommerce_Sales_Analytics_Dashboard_FY2024.xlsx
```

---

## 🎯 Project Objective

To design and build a production-ready sales analytics workbook that mirrors the work a Data Analyst does at an e-commerce company — transforming raw transactional data into executive-level insights using only Microsoft Excel.

This project answers real business questions:
- Which product category generates the most revenue and highest profit margin?
- Which month had the strongest and weakest performance?
- Which region and sales channel is most profitable?
- How did the business perform quarter-over-quarter?
- Which categories need strategic attention vs. which are star performers?

---

## 🗂️ Workbook Structure — 9 Sheets

| # | Sheet Name | Description |
|---|---|---|
| 1 | `KPI_Dashboard` | Executive dark-theme dashboard with KPI cards, top-10 category ranking, monthly trend table, and regional + channel summary |
| 2 | `Monthly_Summary` | 12-month performance table with revenue, GP margin, orders, MoM growth — conditional formatting highlights positive/negative trends |
| 3 | `Category_Analysis` | Deep-dive into 10 product categories with data bars (revenue share) and color scales (GP margin) |
| 4 | `Regional_Analysis` | 5-region breakdown with revenue heatmap using color scale conditional formatting |
| 5 | `Channel_Analysis` | Online vs Retail Store vs Wholesale vs Mobile App performance comparison |
| 6 | `Advanced_Analytics` | Performance tier classification using Nested IFs, Quarter-over-Quarter analysis via SUMPRODUCT, and statistical summary |
| 7 | `Charts_Visualizations` | 4 linked charts: monthly revenue bar, GP margin line trend, category horizontal bar, region pie chart |
| 8 | `Raw_Data` | 1,913 rows of source transaction data with auto-filter and freeze panes |
| 9 | `Documentation` | Full data dictionary, formula reference guide, and sheet navigation guide |

---

## 📐 Dataset Overview

| Attribute | Details |
|---|---|
| **Total Records** | 1,913 orders |
| **Time Period** | January – December 2024 (Full Year) |
| **Product Categories** | Electronics, Apparel, Home & Kitchen, Sports, Beauty, Books, Toys, Automotive, Food & Grocery, Health |
| **Regions** | North, South, East, West, Central |
| **Sales Channels** | Online, Retail Store, Wholesale, Mobile App |
| **Fields per Record** | 17 columns (Order ID, Month, Category, Product, Region, Channel, Units, Price, Revenue, COGS, Gross Profit, Discount, Returns, Rating, Delivery Days, and more) |

---

## 🧠 Excel Skills & Formulas Demonstrated

### Aggregation & Lookup
| Formula | Used For |
|---|---|
| `SUMIF` | Total revenue, units, gross profit by month / category / region / channel |
| `COUNTIF` | Order count by month, category, region |
| `AVERAGEIF` | Average rating and delivery days by region/channel |
| `INDEX / MATCH` | Dynamic lookup — best category, peak revenue month, top region |
| `SUMPRODUCT` | Quarterly aggregation with multi-condition array logic |
| `LARGE / MATCH` | Top-N category ranking logic in KPI Dashboard |

### Logic & Classification
| Formula | Used For |
|---|---|
| `Nested IF (4-level)` | Performance tier: 🥇 STAR → 🥈 PERFORMER → 🎯 HIGH MARGIN → 📈 HIGH VOLUME → ⚠ NEEDS ATTENTION |
| `IFERROR` | Clean error handling on all division and lookup formulas |
| `IF + AND` | Multi-condition logic for tier classification |

### Statistical Analysis
| Formula | Used For |
|---|---|
| `STDEV` | Revenue volatility across months |
| `AVERAGE` | Mean calculations — rating, price, delivery time |
| `MAX / MIN` | Peak and lowest performance identification |
| Coefficient of Variation | `STDEV / AVERAGE` — measures revenue consistency |

### Formatting & Visualization
- **Conditional Formatting** — Color scales (revenue heatmap), Data Bars (revenue share %), Green/Red rules (MoM growth)
- **Number Formatting** — Currency `$#,##0`, Percentages `0.0%`, Thousands separator
- **Charts** — Bar (monthly revenue), Line (GP margin trend), Horizontal Bar (category comparison), Pie (regional share)
- **Freeze Panes** — Applied on all data sheets for usability
- **Auto Filter** — Enabled on Raw_Data for quick slicing

---

## 📊 Key Metrics Calculated

```
Total Revenue          →  SUMIF across all 12 months from Raw_Data
Gross Profit           →  Revenue − COGS, aggregated by category
GP Margin %            →  Gross Profit / Revenue (formatted as %)
Month-over-Month Δ     →  (Current Month − Previous Month) / Previous Month
Quarter-over-Quarter   →  SUMPRODUCT with month-number array conditions
Performance Tier       →  Nested IF with LARGE() benchmark thresholds
Revenue Std Deviation  →  STDEV(Monthly Revenue) — volatility measure
Return Rate            →  Returns / Units Sold per category
```

---

## 🏗️ Project Architecture

```
Raw_Data (source)
    │
    ├──► Monthly_Summary    (SUMIF by Month_Num)
    ├──► Category_Analysis  (SUMIF by Category)
    ├──► Regional_Analysis  (SUMIF by Region)
    └──► Channel_Analysis   (SUMIF by Channel)
              │
              ├──► KPI_Dashboard        (pulls from all 4 summary sheets)
              ├──► Advanced_Analytics   (SUMPRODUCT, Nested IF, STDEV)
              └──► Charts_Visualizations (linked to Monthly + Category sheets)
```

All formulas are **live and dynamic** — changing any value in `Raw_Data` cascades updates across every sheet automatically.

---

## 💼 Business Insights This Project Can Surface

1. **Seasonality Detection** — Monthly revenue trend reveals Q4 spike (Oct–Dec) from holiday demand
2. **Category Profitability** — Electronics drives highest revenue; Beauty/Health show best GP margins
3. **Channel Efficiency** — Online vs Mobile App conversion and margin comparison
4. **Regional Gaps** — Color heatmap surfaces underperforming regions for targeted strategy
5. **QoQ Growth** — Quarter-over-quarter table reveals acceleration or deceleration in growth
6. **Performance Tiers** — Automatic classification flags categories needing management attention

---

## 🛠️ How to Use

1. **Open** `Ecommerce_Sales_Analytics_Dashboard_FY2024.xlsx` in Microsoft Excel (2016 or later recommended)
2. **Start** on the `KPI_Dashboard` sheet for the executive overview
3. **Drill down** into `Monthly_Summary`, `Category_Analysis`, `Regional_Analysis`, or `Channel_Analysis` for detailed breakdowns
4. **Explore** `Advanced_Analytics` for classification logic and statistical insights
5. **View** `Charts_Visualizations` for all 4 linked charts
6. **Reference** `Documentation` for the full data dictionary and formula guide
7. **Filter** raw transactions on `Raw_Data` using the auto-filter dropdowns

> ⚠️ Enable macros is **not required** — this workbook uses formulas only, no VBA.

---

## 📌 Resume / Portfolio Description

> **E-Commerce Sales Analytics Dashboard (Excel) — FY 2024**
> Built a 9-sheet analytics workbook analyzing 1,913 transactions across 10 product categories, 5 regions, and 4 sales channels. Implemented 529 live formulas including SUMIF, INDEX/MATCH, SUMPRODUCT, and 4-level Nested IFs for automated performance classification. Designed an executive KPI dashboard with conditional formatting, color-scale heatmaps, data bars, and 4 linked charts. Applied statistical analysis (STDEV, CV) to measure revenue volatility. All sheets dynamically linked to a single source data table.

---

## 🧰 Tools & Requirements

| Tool | Version |
|---|---|
| Microsoft Excel | 2016 / 2019 / 2021 / Microsoft 365 (recommended) |
| Google Sheets | Partially compatible (some formatting may differ) |
| LibreOffice Calc | Basic compatible |

> For the best experience and full conditional formatting + chart rendering, use **Microsoft Excel 2019 or Microsoft 365**.

---

## 👤 Author

**[ Sneha Kumari]**
Aspiring Data Analyst | Excel · SQL · Python · Power BI

- 🔗 LinkedIn: [https://www.linkedin.com/in/sneha-sahu-427877320/]
- 📧 Email: [snehasahu909@gmail.com]

---

## 📄 License

This project is created for portfolio and educational purposes.
Feel free to use it as inspiration — give credit if you share it publicly.

---

*Built with 💙 using Microsoft Excel — no plugins, no macros, just pure formula-driven analytics.*
