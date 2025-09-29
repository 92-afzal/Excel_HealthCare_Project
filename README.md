# Excel Dashboard Project - Healthcare Emergency Room Data

## 📌 Project Overview
This project demonstrates how to build an **interactive Excel dashboard** using dummy healthcare emergency room data. The dataset includes patient admission details along with a generated calendar table, and the dashboard shows insights such as patient trends and KPIs.

The project also illustrates how to handle **data modeling** in Power Query and Power Pivot, including creating relationships between fact and dimension tables.

---

## 📂 Files Included
- **dashboard HC.xlsx** → Main Excel file with data model, pivot tables, and dashboard visualization.
---

## 🛠️ Steps in the Project

### 1. Data Preparation
- Created a **Calendar Table** using Power Query:
  ```m
  = List.Dates(#date(2023,01,01), 731, #duration(1,0,0,0))
  ```
- Cleaned and transformed **Patient Admission Date** to ensure proper `Date` type using *Change Type with Locale* to resolve text-format issues.

### 2. Data Modeling
- Built relationships between:
  - **Fact Table (Hospital Emergency Room data)**  
  - **Dimension Table (Calendar)**  
- Ensured both sides used the correct `Date` data type.

### 3. Pivot Tables & Measures
- Created Pivot Tables to analyze:
  - Patient counts by Year/Month.
  - Channel performance (Impressions, CTR, %Δ).
  - Data source and campaign breakdowns (dummy marketing dataset).

- Used calculated fields and percentage change formula:
  ```
  % Change = (Current - Previous) / Previous * 100
  ```

### 4. Dashboard Design
- Built an **Excel dashboard** with:
  - KPIs (Spend, CPM, CTR, CPC, Conversions).  
  - Channel performance summary.  
  - Trend line charts.  
  - Slicers for interactivity (custom styled with transparent background).  

### 5. Dummy Data for Simulation
- Since no live connection to BigQuery/SQL was used, dummy CSVs were generated with random values.
- Extra CSV created to pre-compute `%Δ` differences, mimicking real SQL/BigQuery queries.

---

## 🔄 How It Would Work in SQL/BigQuery
In production, percentage differences would be computed dynamically. Example:

```sql
SELECT
  channel,
  month,
  impressions,
  ctr,
  SAFE_DIVIDE(impressions - LAG(impressions) OVER (PARTITION BY channel ORDER BY month), 
              LAG(impressions) OVER (PARTITION BY channel ORDER BY month)) * 100 AS impressions_pct_change
FROM my_project.my_dataset.channel_performance
```

---

## 🚀 How to Use
1. Open `dashboard HC.xlsx` in Excel.  
2. Refresh Pivot Tables to update metrics.  
3. Use slicers to filter by year, month, or channel.  
4. Explore KPIs and performance comparisons.

---

## 📖 Key Learnings
- How to generate and clean data with Power Query.  
- How to build relationships in a data model.  
- How to calculate percentage differences manually when live SQL connections are unavailable.  
- Dashboard design best practices (transparent slicers, KPI cards, trend visuals).  

---

## ✅ Next Steps
- Connect Excel/Power BI directly to **BigQuery or SQL** for real-time refresh.  
- Automate percentage change calculations in the backend (instead of static CSV).  
- Enhance visuals with conditional formatting and interactive charts.  

---

**Author:** Muhammad Afzal

**Email:** hmafzal@gmail.com

**Contact:** +923225353731