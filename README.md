
# 📊 Sales Performance Dashboard — Excel

> An interactive, multi-year Sales Performance Dashboard built entirely in Microsoft Excel using PivotTables, Slicers, and advanced formulas — no macros, no Power BI.

---

## 🖼️ Preview

Sales Performance Dashboard

<img width="707" height="740" alt="image" src="https://github.com/Shobhith01/sales-performance-dashboard-excel/blob/main/Dashboard-Preview/Screenshot%202026-05-11%20193755.png" />

---

## 📌 Overview

This dashboard provides a 360° view of sales performance across **4 years (2018–2021)**, covering revenue trends, customer insights, product performance, and category breakdowns — all filterable interactively using slicers.

| Metric | Value |
|---|---|
| 💰 Total Revenue | $1,92,83,112 |
| 📦 Total Orders | 6,820 |
| 👥 Total Customers | 657 |
| 🧮 Total Units Sold | 8,64,573 |

---

## ✨ Features

- **KPI Summary Cards** — Instant snapshot of Revenue, Orders, Customers & Units Sold
- **Top 10 Customers by Revenue** — Bar chart highlighting highest-value accounts
- **Revenue Trend (Multi-Year)** — Line chart comparing monthly trends across 2018–2021
- **Revenue by Category** — Donut chart across 7 categories: Apparel, Food, Footwear, Mugs, Other, Packaging, Toys
- **Top Products** — Combined bar chart showing revenue and quantity for best-selling SKUs
- **Interactive Slicers** — Filter everything by Year, Month, Category, and Product simultaneously

---

## 🗂️ File Structure

```
Sales_Performance_Dashboard.xlsx
├── DataForPivotsRaw   ← Raw transactional data (21,988 rows)
├── Data               ← Enriched data with Month, MonthNum, Category columns
├── Pivot_KPIs         ← PivotTables powering KPI cards
└── Pivot_Chart        ← PivotTables powering all charts + slicer connections
```

---

## 🔧 Month Slicer Fix — What Was Wrong & How It Was Fixed

### The Problem
The `Month` slicer was displaying months in **alphabetical order** (Apr, Aug, Dec, Feb...) instead of **calendar order** (Jan, Feb, Mar...).

**Root cause:** The `Month` column used `=TEXT(InvoiceDate,"MMM")` which outputs plain text abbreviations. Excel's slicer sorts these alphabetically, not chronologically.

### The Fix
Updated the `Month` column formula to prefix each month name with its two-digit numeric equivalent:

```excel
-- Before (alphabetical sort)
=TEXT(InvoiceDate,"MMM")
-- Outputs: "Jan", "Feb", "Mar" → sorted as Apr, Aug, Dec...

-- After (chronological sort)
=TEXT(MONTH(InvoiceDate),"00")&"-"&TEXT(InvoiceDate,"MMM")
-- Outputs: "01-Jan", "02-Feb", "03-Mar" → sorted correctly!
```

This forces alphabetical sorting to align with calendar ordering — no VBA, no helper columns, no Power Query required.

---

## 🛠️ Tools & Techniques Used

| Tool / Feature | Usage |
|---|---|
| **Excel PivotTables** | Aggregating 21,988 rows for all charts |
| **Excel Slicers** | Cross-filtering across all charts simultaneously |
| **TEXT / MONTH formulas** | Date parsing and month ordering |
| **IF + SEARCH formulas** | Dynamic category classification |
| **Named Tables** | Structured references (`SalesData3`) for formula reliability |
| **Conditional Formatting** | KPI card color highlights |
| **Chart Types** | Bar, Line (multi-series), Donut, Combo (bar + line) |

---

## 📂 Dataset

- **Source:** Adapted from the [Wide World Importers](https://learn.microsoft.com/en-us/sql/samples/wide-world-importers-what-is) sample dataset (Microsoft)
- **Period:** January 2018 – March 2021
- **Records:** 21,988 invoice line items
- **Fields:** CustomerID, CustomerName, InvoiceID, InvoiceDate, StockItemID, Description, Quantity, UnitPrice, TotalLine, TaxAmount, ExtendedPrice

---

## 🚀 How to Use

1. **Download** `Sales_Performance_Dashboard.xlsx`
2. **Open** in Microsoft Excel (2016 or later recommended)
3. Click **Enable Editing** if prompted
4. Use the **slicers on the right panel** to filter by:
   - Year (2018, 2019, 2020, 2021)
   - Month (01-Jan through 12-Dec, in order ✅)
   - Category (Apparel, Food, Footwear, Mugs, Other, Packaging, Toys)
   - Product (individual SKU level)
5. All 4 charts and KPI cards update **instantly** with your selection

---

## 📁 Repository Contents

```
├── README.md
├── Sales_Performance_Dashboard.xlsx    ← Main dashboard file
└── Screenshot_2026-05-11_193755.png    ← Dashboard preview
```

---

## 🙋 About

Built as a data analytics portfolio project to demonstrate proficiency in:
- Excel dashboard design and layout
- PivotTable architecture for interactive reporting
- Data modelling and formula-based ETL in Excel
- UX thinking for business intelligence tools

---
## 👤 Author

**Shobhith S Kounder**
Aspiring Data Analyst · SQL · Excel · Power BI
📍 Sringeri, Karnataka — Open to relocation anywhere in India

<p>
  <a href="https://www.linkedin.com/in/shobhith-kounder-768859264">
    <img src="https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white"/>
  </a>
  <a href="https://github.com/Shobhith01">
    <img src="https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white"/>
  </a>
</p>

---

<p align="center">
  <i>This project is part of a 14-day data analyst portfolio challenge —
  building job-ready projects in SQL, Excel and Power BI to land
  a data analyst role in Bengaluru.</i>
</p>

