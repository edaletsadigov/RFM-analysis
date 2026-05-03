# 🛒 RFM Customer Segmentation Analysis
### UK-Based E-Commerce | Dec 2010 – Dec 2011

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![Power Query](https://img.shields.io/badge/ETL-Power%20Query-2C3E50?style=flat)
![Pivot Tables](https://img.shields.io/badge/Analysis-Pivot%20Tables-E67E22?style=flat)
![RFM](https://img.shields.io/badge/Method-RFM%20Segmentation-27AE60?style=flat)

---

## 📌 Project Overview

End-to-end customer segmentation project applied to a UK-based online retail dataset.
The goal was to identify distinct customer groups using the RFM (Recency, Frequency, Monetary)
framework — enabling targeted retention and growth strategies per segment.

The entire workflow — from raw data ingestion to final dashboard — was built in **Microsoft Excel**
using **Power Query** for cleaning and **Pivot Tables** for analysis.

---

## 📁 Dataset

| Attribute | Value |
|---|---|
| Source | UCI Machine Learning Repository — [Online Retail Dataset](https://archive.ics.uci.edu/ml/datasets/online+retail) |
| Raw rows | 541,909 |
| Clean rows | 391,153 |
| Customers | 4,334 |
| Invoices | 18,402 |
| Period | 01 Dec 2010 → 09 Dec 2011 |
| Countries | 38 |

---

## 🧹 Data Cleaning (Power Query)

Raw data contained multiple quality issues that required systematic treatment before any analysis.
All cleaning steps were applied in **Power Query** — no data was modified manually.

| Step | Issue | Action | Rows Removed |
|---|---|---|---|
| 1 | Full duplicate rows | `Table.Distinct()` | ~5,270 |
| 2 | Service/non-product StockCodes | Exclude: POST, DOT, M, C2, D, S, BANK CHARGES, AMAZONFEE, CRUK | ~2,900 |
| 3 | Cancelled invoices (C-prefix) | `Text.StartsWith([InvoiceNo], "C")` filter | ~9,288 |
| 4 | Accounting adjustments (A-prefix) | `Text.StartsWith([InvoiceNo], "A")` filter | 3 |
| 5 | Negative / zero Quantity | Keep `[Quantity] > 0` only | ~10,624 |
| 6 | Zero / negative UnitPrice | Keep `[UnitPrice] > 0` only | ~2,517 |
| 7 | Null CustomerID | Remove — RFM requires customer identity | ~135,080 |
| 8 | Unparseable / null InvoiceDate | `Remove Errors` + null filter | ~rest |
| 9 | CustomerID type | Float64 → Integer (post-null removal) | — |
| + | Engineered feature | `TotalRevenue = Quantity × UnitPrice` | — |

> **Note on A-prefix invoices:** Three rows (`A563185–A563187`) were "Adjust bad debt" entries
> with offsetting values (+£1.1M and −£1.1M). These are accounting corrections, not sales transactions.

---

## 📐 RFM Methodology

RFM scores each customer on three behavioural dimensions:

| Dimension | Definition | Scoring Logic |
|---|---|---|
| **Recency (R)** | Days since last purchase (as of 2011-12-10) | Lower recency = higher score (1–5) |
| **Frequency (F)** | Number of unique invoices | Higher frequency = higher score (1–5) |
| **Monetary (M)** | Total spend (£) | Higher spend = higher score (1–5) |

Scores were assigned using **quintile-based binning** (equal-frequency intervals).
Each customer receives a composite `RFM_CATEGORY` label based on their R/F/M score combination.

### Segment Definitions

| Segment | R | F | M | Description |
|---|---|---|---|---|
| VİP Customer | High | High | High | Best customers — buy often, recently, and spend most |
| Loyal Customer | Medium–High | High | High | Consistent buyers, slightly less recent |
| New Customer | High | Low | Low–Mid | Recent first-time or early buyers |
| Mid Value Customer | Medium | Medium | Medium | Stable but underdeveloped customers |
| At Risk Customer | Low | High | High | Previously valuable, now disengaging |
| Lost Customer | Low | Low | Low | Haven't purchased in a long time |

---

## 📊 Key Results

### Segment Distribution

| Segment | Customers | % of Total | Revenue (£) | Rev % |
|---|---|---|---|---|
| 🏆 VİP Customer | 325 | 7.5% | 365,561,120 | **45.0%** |
| 💙 Loyal Customer | 531 | 12.3% | 169,771,637 | 20.9% |
| ⚰️ Lost Customer | 1,916 | **44.2%** | 118,465,269 | 14.6% |
| 🌱 Mid Value Customer | 845 | 19.5% | 61,949,156 | 7.6% |
| 🆕 New Customer | 555 | 12.8% | 60,775,345 | 7.5% |
| ⚠️ At Risk Customer | 162 | 3.7% | 35,538,160 | 4.4% |
| **Total** | **4,334** | 100% | **812,060,687** | 100% |

### Business KPIs

| Metric | Value |
|---|---|
| Total Revenue | £812,060,687 |
| Total Orders | 18,402 |
| Unique Customers | 4,334 |
| Average Transaction Value | £44,129 |
| Average Items per Order | 21.3 |
| Peak Month | November 2011 — £104,940,239 |

---

## 💡 Key Insights

**1. 325 customers drive 45% of total revenue**
VİP customers (7.5% of the base) generate £365M. Losing even 50 of them has a measurable P&L impact.

**2. 44% of customers are Lost — but they weren't always low-value**
Lost customers contributed £118M historically. A targeted win-back campaign with even 10% conversion
would recover ~£12M.

**3. At Risk customers represent £35.5M in endangered revenue**
162 customers who were once frequent, high-spend buyers are now disengaging. Average days since last
purchase: ~118 days. Time-sensitive reactivation window.

**4. UK concentration risk: 82.4% of revenue from one market**
United Kingdom: £669M. All other 37 countries combined: £143M. Single-market dependency is a
structural vulnerability. Notably, the Netherlands (9 customers) generates £27.2M — £3M per customer
vs £171K per UK customer.

**5. November is peak month — Q4 accounts for 35% of annual revenue**
Sep–Nov 2011 combined: £284M (35% of annual total). Inventory, logistics, and staffing should be
planned with this seasonal curve in mind.

**6. One-time buyers (34.7%) leave £59M on the table**
1,505 customers purchased exactly once. Converting 20% of them to repeat buyers would add ~£34M
based on the observed repeat-buyer average spend differential (6.8×).

**7. Thursday is peak sales day — Saturday has zero transactions**
The absence of Saturday data confirms this is a **B2B wholesale** platform (not consumer retail).
Thursday peaks suggest customers place orders ahead of the weekend for early-week delivery.

---

## 🗂️ File Structure

```
rfm-analysis/
│
├── RFM_PROJECT_FINAL.xlsx          ← Main workbook
│   ├── RFM Project                 ← Cleaned data + RFM scores + segments
│   ├── pivots                      ← Pivot tables (segment counts, KPIs)
│   └── dashboard 2                 ← Dashboard sheet
│
└── README.md                       ← This file
```

---

## 🛠️ Tools Used

| Tool | Purpose |
|---|---|
| Microsoft Excel | Main environment |
| Power Query (M Language) | Data ingestion, cleaning, transformation |
| Pivot Tables | Aggregation and segment analysis |
| Excel Charts | Visualization |
| Custom Column Formulas | RFM scoring and segmentation logic |

---

## 🚀 How to Use

1. Download `RFM_PROJECT_FINAL.xlsx`
2. Open in Microsoft Excel (2016 or later recommended)
3. Navigate to the **`RFM Project`** sheet to explore the cleaned dataset with RFM scores
4. Navigate to **`pivots`** for aggregated segment metrics
5. Use **Slicers** on the dashboard to filter by Country or Month

---

## 👤 Author Sadıqov Ədalət

**Ədalət**
Data Analytics Intern @ Codveda Technologies & SaiKet Systems
Baku, Azerbaijan

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=flat&logo=linkedin)](https://linkedin.com)
[![GitHub](https://img.shields.io/badge/GitHub-Portfolio-181717?style=flat&logo=github)](https://github.com)
