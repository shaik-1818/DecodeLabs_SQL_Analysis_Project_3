# 🗄️ Project 3: SQL Data Analysis
**DecodeLabs Industrial Training | Batch 2026**
**Author:** Shaik | Mohan Babu University


---

## 📌 Problem Statement

Raw data in spreadsheets can't answer specific business questions at scale. This project loads the cleaned sales dataset into a relational database (SQLite) and uses 15 structured SQL queries to extract actionable business intelligence — proving the ability to filter, group, and aggregate data with precision.

> *"Think like the database, filter early, and build your queries logically."*

---

## 🎯 Goal

Use SQL queries to extract insights from the 1,200-row sales dataset covering revenue trends, product performance, customer behaviour, and loss analysis.

---

## 📁 Project Structure

```
Project 3/
├── sql_analysis_project3.py       ← Main SQL analysis script
├── Dataset for Data Analytics.xlsx ← Input dataset
└── SQL_Analysis_Project3.xlsx      ← Output (16 formatted sheets)
    ├── Summary                     ← Executive summary of all 15 queries
    ├── Overall_Summary             ← Q01 results
    ├── Revenue_By_Product          ← Q02 results
    ├── Orders_By_Status            ← Q03 results
    ├── Annual_Revenue_Trend        ← Q04 results
    ├── Monthly_Revenue_2024        ← Q05 results
    ├── Payment_Method_Analysis     ← Q06 results
    ├── Top10_Highest_Orders        ← Q07 results
    ├── Coupon_Code_Impact          ← Q08 results
    ├── Referral_Source_Performance ← Q09 results
    ├── Delivered_Above_Avg         ← Q10 results
    ├── Cancelled_Returned_Losses   ← Q11 results
    ├── High_Value_Products_HAVING  ← Q12 results
    ├── Bulk_Orders_Qty5            ← Q13 results
    ├── Instagram_Product_Perf      ← Q14 results
    └── Full_Monthly_Trend          ← Q15 results
```

---

## 🧠 SQL Execution Order (How the Engine Reads It)

> Humans write SQL top-to-bottom. The database engine executes in a different order:

| Step | Clause | Purpose |
|---|---|---|
| 1 | `FROM` | Locates the data source (table) |
| 2 | `WHERE` | Filters individual rows |
| 3 | `GROUP BY` | Categorises rows into buckets |
| 4 | `HAVING` | Filters aggregated buckets |
| 5 | `SELECT` | Picks columns and computes metrics |
| 6 | `ORDER BY` | Sorts the final output |

> ⚠️ **Alias Trap**: You cannot use a SELECT alias in a WHERE clause — WHERE runs before SELECT. Use the original column name instead.

---

## 📋 All 15 SQL Queries

### Q01 — Overall Business Summary
```sql
SELECT
    COUNT(*)                  AS Total_Orders,
    ROUND(SUM(TotalPrice), 2) AS Total_Revenue,
    ROUND(AVG(TotalPrice), 2) AS Avg_Order_Value,
    ROUND(MIN(TotalPrice), 2) AS Min_Order,
    ROUND(MAX(TotalPrice), 2) AS Max_Order
FROM orders;
```
**Result:** 1,200 orders | Rs 12,64,762 revenue | Rs 1,053.97 avg

---

### Q02 — Revenue by Product
```sql
SELECT Product, COUNT(*) AS Order_Count,
    ROUND(SUM(TotalPrice), 2) AS Total_Revenue,
    ROUND(AVG(TotalPrice), 2) AS Avg_Order_Value
FROM orders
GROUP BY Product
ORDER BY Total_Revenue DESC;
```
**Result:** Chair (Rs 1,95,620) leads; Laptop has highest AOV (Rs 1,110.56)

---

### Q03 — Orders by Status (with % share)
```sql
SELECT OrderStatus, COUNT(*) AS Order_Count,
    ROUND(SUM(TotalPrice), 2) AS Total_Revenue,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM orders), 2) AS Pct_of_Total
FROM orders
GROUP BY OrderStatus
ORDER BY Order_Count DESC;
```
**Result:** Cancelled+Returned = 41.4% — critical Rs 5,19,673 revenue at risk

---

### Q04 — Annual Revenue Trend
```sql
SELECT STRFTIME('%Y', Date) AS Year,
    COUNT(*) AS Orders,
    ROUND(SUM(TotalPrice), 2) AS Revenue
FROM orders
GROUP BY Year
ORDER BY Year;
```
**Result:** 2023: Rs 5,52,643 → 2024: Rs 4,80,235 (−13.1% decline)

---

### Q05 — Monthly Revenue 2024 (WHERE filter)
```sql
SELECT STRFTIME('%Y-%m', Date) AS Month,
    COUNT(*) AS Orders,
    ROUND(SUM(TotalPrice), 2) AS Revenue
FROM orders
WHERE STRFTIME('%Y', Date) = '2024'
GROUP BY Month
ORDER BY Month;
```
**Result:** June peak (Rs 68,068); May trough (Rs 27,909)

---

### Q06 — Payment Method Analysis
```sql
SELECT PaymentMethod, COUNT(*) AS Order_Count,
    ROUND(AVG(TotalPrice), 2) AS Avg_Order_Value
FROM orders
GROUP BY PaymentMethod
ORDER BY Order_Count DESC;
```
**Result:** Online most popular (258); Credit Card highest avg (Rs 1,127.55)

---

### Q07 — Top 10 Highest Value Orders (LIMIT)
```sql
SELECT OrderID, Date, Product, Quantity, TotalPrice, OrderStatus
FROM orders
ORDER BY TotalPrice DESC
LIMIT 10;
```
**Result:** All top 10 are Qty=5 bulk orders — confirmed VIP signals

---

### Q08 — Coupon Code Impact
```sql
SELECT CouponCode, COUNT(*) AS Usage_Count,
    ROUND(SUM(TotalPrice), 2) AS Total_Revenue
FROM orders
GROUP BY CouponCode
ORDER BY Usage_Count DESC;
```
**Result:** FREESHIP leads (313 uses, Rs 3,35,037 revenue)

---

### Q09 — Referral Source Performance (% share subquery)
```sql
SELECT ReferralSource, COUNT(*) AS Orders,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM orders), 2) AS Pct_Share
FROM orders
GROUP BY ReferralSource
ORDER BY Orders DESC;
```
**Result:** Instagram #1 channel (21.6%); Facebook highest avg order

---

### Q10 — Delivered Orders Above Average (Subquery in WHERE)
```sql
SELECT OrderID, Product, TotalPrice, Date
FROM orders
WHERE OrderStatus = 'Delivered'
  AND TotalPrice > (SELECT AVG(TotalPrice) FROM orders)
ORDER BY TotalPrice DESC
LIMIT 15;
```
**Result:** Demonstrates correlated subquery — filters dynamically against dataset average

---

### Q11 — Cancelled & Returned Loss Analysis (WHERE IN)
```sql
SELECT OrderStatus, Product, COUNT(*) AS Count,
    ROUND(SUM(TotalPrice), 2) AS Lost_Revenue
FROM orders
WHERE OrderStatus IN ('Cancelled', 'Returned')
GROUP BY OrderStatus, Product
ORDER BY OrderStatus, Lost_Revenue DESC;
```
**Result:** Chair highest cancellation loss (Rs 48,660); Tablet highest return loss (Rs 42,525)

---

### Q12 — High Value Products using HAVING
```sql
SELECT Product, COUNT(*) AS Orders,
    ROUND(AVG(TotalPrice), 2) AS Avg_Revenue
FROM orders
GROUP BY Product
HAVING AVG(TotalPrice) > 1000
ORDER BY Avg_Revenue DESC;
```
**Result:** 5 of 7 products qualify. Desk & Phone excluded (avg < Rs 1,000)

---

### Q13 — Bulk Orders Profile (Quantity = 5)
```sql
SELECT Product, OrderStatus, PaymentMethod,
    COUNT(*) AS Bulk_Orders, ROUND(SUM(TotalPrice), 2) AS Revenue
FROM orders
WHERE Quantity = 5
GROUP BY Product, OrderStatus, PaymentMethod
ORDER BY Revenue DESC
LIMIT 20;
```
**Result:** Confirms all IQR outliers are genuine bulk/VIP orders

---

### Q14 — Instagram Channel Product Performance
```sql
SELECT Product, COUNT(*) AS Orders,
    ROUND(SUM(TotalPrice), 2) AS Revenue
FROM orders
WHERE ReferralSource = 'Instagram'
GROUP BY Product
ORDER BY Revenue DESC;
```
**Result:** Laptop leads Instagram revenue (Rs 48,453)

---

### Q15 — Full Monthly Revenue Trend (All Years)
```sql
SELECT STRFTIME('%Y-%m', Date) AS Month,
    COUNT(*) AS Orders, ROUND(SUM(TotalPrice), 2) AS Revenue
FROM orders
GROUP BY Month
ORDER BY Month;
```
**Result:** 30 months of data. May 2023 peak (Rs 63,836); May 2024 worst (Rs 27,909)

---

## 📊 SQL Clauses Coverage

| Clause | Queries Using It |
|---|---|
| `SELECT` + Aggregations (`COUNT`, `SUM`, `AVG`, `MIN`, `MAX`) | Q01–Q15 |
| `WHERE` (equality, comparison, pattern) | Q05, Q07, Q10, Q11, Q13, Q14 |
| `GROUP BY` | Q02–Q06, Q08, Q09, Q11–Q15 |
| `ORDER BY` (ASC / DESC) | Q02–Q15 |
| `HAVING` | Q12 |
| `LIMIT` | Q07, Q10, Q13 |
| `Subquery` in WHERE | Q10 |
| `Subquery` in SELECT | Q03, Q09 |
| `STRFTIME` (date functions) | Q04, Q05, Q15 |
| `WHERE IN` | Q11 |

---

## 💼 Key Business Insights from SQL Analysis

| Finding | Business Action |
|---|---|
| Chair leads total revenue (Rs 1,95,620) | Prioritise Chair stock and marketing |
| Laptop has highest AOV (Rs 1,110.56) | Feature Laptop in premium campaigns |
| 41.4% cancel+return rate (Rs 5.19L at risk) | Investigate root causes urgently |
| Revenue declined 13.1% (2023→2024) | Deep-dive into 2024 monthly data |
| June is consistently the best month | Run promotions in June every year |
| FREESHIP coupon is most used (313) | Expand free shipping offers |
| Instagram drives 21.6% of acquisitions | Increase Instagram ad budget |
| Credit Card orders have highest AOV | Target credit card users for upsell |

---

## 🛠️ Tech Stack

| Tool | Purpose |
|---|---|
| Python 3.x | Script runner and data pipeline |
| pandas | Loading Excel data, executing queries via `read_sql_query` |
| sqlite3 | In-memory relational database engine |
| openpyxl | Formatted Excel report with 16 styled sheets |
| os | Cross-platform file path handling |

---

## ▶️ How to Run

1. Place `sql_analysis_project3.py` and `Dataset for Data Analytics.xlsx` in the **same folder**
2. Open terminal in that folder
3. Install dependencies (first time only):
```bash
pip install pandas openpyxl
```
4. Run the script:
```bash
python sql_analysis_project3.py
```
5. `SQL_Analysis_Project3.xlsx` with 16 formatted sheets appears in the same folder

---

## 💡 Key Learnings

- **SQL is declarative** — you describe WHAT you want, the engine decides HOW to get it
- **Execution order ≠ writing order** — FROM → WHERE → GROUP BY → HAVING → SELECT → ORDER BY
- **HAVING vs WHERE** — WHERE filters rows before grouping; HAVING filters groups after aggregation
- **Subqueries in WHERE** — powerful for dynamic filtering (e.g. above average)
- **STRFTIME** — essential for time-series analysis in SQLite
- **Every query answers a business question** — not just an exercise in syntax

---

*DecodeLabs | Professional Standard | Batch 2026*
