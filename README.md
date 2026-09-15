# Task 8 — Sales Summary Using SUMIFS

**Internship:** Veda Technology — Data Analytics Track
**Level:** 1 · Day 8 of 45
**Submitted by:** Akshat Srivastava

## Overview

This task is about using the Excel `SUMIFS` function to turn transaction-level
sales data into useful daily, weekly, and monthly sales summaries. The
workbook uses live formulas so the summary recalculates automatically from the
underlying sales dataset.

## Objective

Build a formula-driven sales summary that calculates daily, weekly, and
monthly totals from the raw transaction data, while also verifying the results
with manual checks.

## Dataset

A custom sales transaction dataset created for this task — **49 transactions,
7 fields**: Date, Product, Category, Qty, Price, and Amount, with one additional
blank column in the source sheet. The dataset covers **1 August 2026 to
21 August 2026**.

The `Amount` column represents the total value of each transaction
(`Qty × Price`). The dataset contains **286 total units** and **₹67,100 total
sales**.

## What's in the workbook

| Sheet                    | What it demonstrates                                                                       |
| ------------------------ | ------------------------------------------------------------------------------------------ |
| `raw_sales_dataset_.csv` | Base transaction-level sales dataset used for the analysis                                 |
| `Summary`                | Daily, weekly, and monthly sales totals calculated using `SUMIFS`, plus verification notes |

## How SUMIFS is used

The `Summary` sheet uses `SUMIFS` with date criteria to calculate sales for
different time periods.

### Daily Sales

```excel
=SUMIFS(raw_sales_dataset_.csv!F:F,raw_sales_dataset_.csv!A:A,A2)
```

Adds the sales amount for transactions matching the selected date.

### Weekly Sales

```excel
=SUMIFS(raw_sales_dataset_.csv!F:F,raw_sales_dataset_.csv!A:A,">="&A2,raw_sales_dataset_.csv!A:A,"<="&A2+6)
```

Adds sales for a seven-day period starting from the date in the summary row.

### Monthly Sales

```excel
=SUMIFS(raw_sales_dataset_.csv!F:F,raw_sales_dataset_.csv!A:A,">="&DATE(2026,8,1),raw_sales_dataset_.csv!A:A,"<="&DATE(2026,8,31))
```

Calculates the total sales for the month using start-date and end-date
criteria.

## Key outputs

* Total sales across all 49 transactions were **₹67,100**.
* Total quantity sold was **286 units**.
* The highest daily sales were **₹6,150 on 17 August 2026**.
* The largest category by sales was **Electronics (₹23,600)**.
* The monthly total was manually verified as **₹67,100**.
* The first-day total was manually checked as **₹3,250** from ₹400 + ₹900 + ₹1,950.

## Challenges & learnings

* **Date-based SUMIFS criteria** — learned how to use `>=` and `<=` conditions
  with dates to calculate weekly and monthly totals.
* **Dynamic summaries** — keeping calculations formula-driven means the
  summary updates when the underlying dataset changes.
* **Verification matters** — manually checking selected totals helped confirm
  that the formulas were returning the expected values.
* `SUMIFS` is useful for turning a transaction table into a simple reporting
  view without manually calculating each total.

## Files in this repo

| File                         | Description                                                                  |
| ---------------------------- | ---------------------------------------------------------------------------- |
| `Task8_Veda_Technology.xlsx` | The workbook containing the raw sales dataset and formula-based summary      |
| `Task8_Report.pdf`           | Full write-up covering the task approach, outputs, challenges, and learnings |
| `README.md`                  | This file                                                                    |

## Tools

Microsoft Excel (workbook uses formula-based analysis with `SUMIFS`).
