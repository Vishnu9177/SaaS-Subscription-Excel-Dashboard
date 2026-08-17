# SaaS-Subscription-Excel-Dashboard
SaaS subscription dashboard built from 1,000-customer data (18 months). Caught a data quality issue where billing records only covered churned customers, fixed MRR calculation accordingly. Surfaced a hidden churn risk: net-new customers dropped 80% while headline MRR/growth metrics still looked healthy — a leading vs. lagging indicator insight.

# SaaS Subscription Business — Executive Dashboard

An Excel dashboard built entirely with live formulas (no hardcoded values) to analyze customer growth, revenue, and churn for a subscription-based SaaS business, using 18 months of raw operational data (Jan 2024 – Jun 2025).

---

## 📁 Repository Contents

| File | Description |
|---|---|
| `SaaS_Subscription_Dashboard.xlsx` | Final deliverable — 6-tab Excel workbook with dashboard, calculation engine, and raw data |
| `customers.csv` | Raw customer roster (signup date, plan, fee, CAC, churn date) |
| `revenue.csv` | Raw monthly billing records |
| `subscriptions.csv` | Raw monthly subscription records |

---

## 📊 About the Dataset

| Attribute | Detail |
|---|---|
| Customers | 1,000 total |
| Time period | Jan 2024 – Jun 2025 (18 months) |
| Plans | Basic ($50/mo), Pro ($200/mo), Enterprise ($500/mo) |
| Fields (customers.csv) | `customer_id`, `signup_date`, `plan_type`, `monthly_fee`, `acquisition_cost`, `churn_date` |
| Fields (revenue.csv / subscriptions.csv) | `subscription_id`, `customer_id`, `month`, `monthly_fee`, `revenue_type`, `amount` |

### ⚠️ Data Quality Finding

`revenue.csv` and `subscriptions.csv` were found to contain billing rows **only for the 168 customers who eventually churned** — verified by matching unique `customer_id` counts against the churned segment in `customers.csv`. Active (never-churned) customers are absent from both files, so they cannot be summed directly to calculate company-wide MRR.

**Resolution:** MRR, ARR, and ARPU were instead derived from the full `customers.csv` roster (`signup_date` / `churn_date` / `monthly_fee`), which covers all 1,000 customers. A reconciliation check (implied vs. reported billing for the churned cohort) is included on the `Monthly_Trends` tab to document and validate this decision.

---

## 🧮 Workbook Structure

1. **Dashboard** — KPI cards, 6 charts, and an auto-updating insights panel
2. **Monthly_Trends** — calculation engine: New/Churned/Active customers, MRR, ARR, ARPU, churn rate per month (all formula-driven)
3. **Plan_Analysis** — churn rate, revenue share, and CAC payback broken out by plan
4. **Raw_Customers / Raw_Revenue / Raw_Subscriptions** — original data as native Excel Tables, retained for traceability

All figures are live formulas referencing the raw tables — updating the source data recalculates the entire workbook.

---

## 🔑 Key Metrics (as of Jun-2025)

| Metric | Value |
|---|---|
| Active Customers | 832 |
| Current MRR | $207,950 |
| Current ARR | $2,495,400 |
| Average Monthly Churn Rate | 1.7% |
| Blended LTV : CAC | 129.65x |
| Average ARPU | $246.75 |

---

## 💡 Key Insights & Patterns

**1. Churn rate is the leading warning signal**
- Held steady near ~1% from Jan-24 through Feb-25
- Spiked to 3.7% in Mar-25, climbing to 5.0% by Jun-25
- The earliest chart to reflect the emerging problem

**2. New vs. Churned Customers confirms the trend**
- Churned customers rose from 6 (Feb-25) → 26 (Mar-25) → 41 (Jun-25)
- New customer additions declined in parallel (60 → 50 by Jun-25)

**3. Net-new customers collapsed in the most recent month**

| Month | Net New |
|---|---|
| Jan-25 | 47 |
| Feb-25 | 50 |
| Mar-25 | 39 |
| Apr-25 | 34 |
| May-25 | 37 |
| Jun-25 | 9 |

- ~80% drop in a single month — the earliest hard evidence of deceleration

**4. Active Customer Growth looks healthy but is a lagging metric**
- Cumulative line still trends up smoothly (57 → 832)
- Masks that growth is now being driven by a shrinking margin
- One more weak month of net-new could visibly flatten the curve

**5. MRR Trend is the most lagging chart — currently hides the risk**
- Being cumulative, it hasn't reacted to the churn spike yet
- Growth deceleration expected within 1–2 months if churn persists

**6. Revenue concentration amplifies churn risk**
- Customer mix is balanced by plan: Basic 32%, Pro 35%, Enterprise 33%
- MRR contribution is not: Enterprise 65.6% ($136.5K), Pro 27.9% ($58K), Basic 6.5% ($13.5K)
- Lifetime churn rate is similar across plans (Basic 17.7%, Pro 15.5%, Enterprise 17.3%)
- Since churn isn't concentrated in low-value tiers, the same churn rate hitting Enterprise accounts carries ~10x the MRR impact of Basic accounts churning

---

## ✅ End Result

The dashboard confirms healthy top-line growth (MRR up ~16x, active customers up from 57 to 832) but surfaces a retention problem forming in Q1–Q2 2025 that is **not yet visible in the headline MRR and active-customer numbers**, because those are lagging, cumulative metrics. The clearest early-warning indicator identified is **net-new customer count**, which dropped ~80% in the final month of the dataset — a signal worth tracking going forward, alongside monthly churn rate.

---

## 🛠️ Built With

- Microsoft Excel (formulas: `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS`, `SUMPRODUCT`, `EOMONTH`, `DATEDIF`)
- Native Excel Tables and Charts
- Python (`openpyxl`, `pandas`) for workbook generation
