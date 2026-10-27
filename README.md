<div align="center">

# 📊 PR. 2 Analyzer — Advanced Excel Analytics & Executive Dashboard

### A comprehensive business intelligence project featuring Descriptive Statistics, Linear Regression Modeling, What-If Sensitivity Analysis, Dynamic Lookups, Advanced Pivot Tables, and a Fully Interactive Executive Dashboard.

<p>
  <img src="https://img.shields.io/badge/Author-Priyanshu%20Buhecha-1f6feb?style=for-the-badge" alt="Author" />
  <img src="https://img.shields.io/badge/Microsoft%20Excel-Advanced%20Analytics-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" alt="Microsoft Excel" />
  <img src="https://img.shields.io/badge/Project-PR.%202%20Analyzer-1f6feb?style=for-the-badge" alt="Project" />
  <img src="https://img.shields.io/badge/Data%20Analysis-Toolpak%20Enabled-8957e5?style=for-the-badge" alt="Data Analysis Toolpak" />
  <img src="https://img.shields.io/badge/Status-Completed-2ea44f?style=for-the-badge" alt="Completed" />
</p>

<p><b>📁 Workbook:</b> <code>PR. 2.xlsx</code> &nbsp;•&nbsp; <b>👤 Analyst:</b> Priyanshu Buhecha &nbsp;•&nbsp; <b>🎯 Focus:</b> Advanced Statistical Modeling, What-If Scenarios & Executive Reporting</p>

</div>

---

## 📌 Project Overview

**PR. 2 Analyzer** is an enterprise-grade Excel practical project developed by **Priyanshu Buhecha**, designed to extract actionable business insights from retail sales transactions using advanced analytical tools, statistical methods, dynamic formula modeling, and visual storytelling.

The project utilizes a rich 200-transaction retail dataset spanning five product categories and five geographic sales regions over a full annual cycle (April 2024 – April 2025). Each task is structured within its own dedicated worksheet to provide modular, transparent, and auditable analytical workflows.

### 🎥 Project Walkthrough Video

Watch the comprehensive video demonstration of this project by **Priyanshu Buhecha**:  
🔗 **[Google Drive Walkthrough Video Link](https://drive.google.com/file/d/1Hl3__nBMU5PULnsr6CPDehda_RTxAoo5/view?usp=sharing)** *(Insert your recording link here)*

---

## 🎯 Main Objectives & Key Highlights

- **Dynamic Visual Cues:** Apply rule-based Conditional Formatting to instantly spotlight top-performing customers and monthly growth direction using 3-arrow icon sets.
- **Scenario & What-If Modeling:** Measure the financial sensitivity of gross profit margins against varying promotional discount structures (0% to 20%).
- **Predictive Statistical Modeling:** Perform Ordinary Least Squares (OLS) Linear Regression using the Excel **Data Analysis Toolpak** to quantify the predictive relationship between Sales and Profit ($R^2 = 59.83\%$).
- **Parametric Distribution Analysis:** Run multi-variable **Descriptive Statistics** across Sales, Quantity, Discount, and Profit to assess central tendency, dispersion, skewness, and kurtosis.
- **Dynamic Array & Lookup Formulas:** Dynamically rank and isolate high-value customers using modern combinations of `INDEX()`, `MATCH()`, and `LARGE()`.
- **Dimensional Cross-Tabulation:** Construct multi-dimensional Pivot Tables to cross-examine sales performance by Region and Product Category.
- **Executive KPI Dashboard:** Build an interactive executive dashboard with synchronized KPI summary metric cards, clustered bar charts, line trends, pie proportions, connected slicers, and a native timeline filter.

---

## 📑 Workbook Architecture & Sheet Navigation

The workbook is organized into 12 dedicated worksheets, separating raw transactions, analytical tasks, visualizations, and dashboard pivot engines:

| 📄 Sheet Name | 🏷️ Task | 🔍 Focus & Core Technique | 💡 Key Analytical Output |
|:---|:---:|:---|:---|
| **`Sheet1`** | — | Master Retail Transaction Dataset | 200 retail records across 5 regions and 5 categories |
| **`Task 1`** | Task 1 | Conditional Formatting: Top 10 Customers | Automated highlighting of the top 10 revenue accounts |
| **`Task 2`** | Task 2 | WHAT-IF Sensitivity Analysis | Profit impact simulation across 0%, 5%, 10%, 15%, 20% discounts |
| **`Task 3`** | Task 3 | Linear Regression (Data Analysis Toolpak) | Model: $\text{Profit} = 0.3511 \times \text{Sales} - 0.8895$ ($p < 0.001$) |
| **`Task 4`** | Task 4 | Advanced Descriptive Statistics | 14 summary statistics across Sales, Qty, Discount, Profit |
| **`Task 5`** | Task 5 | Custom Formatting & Directional Trend Arrows | 3-Arrow Icon Sets indicating top, mid, and low sales growth tiers |
| **`Task 6`** | Task 6 | Real-Time Dynamic Audit Timestamp | Live timestamp generation using `=NOW()` |
| **`Task 7`** | Task 7 | High-Value Customer Identification | Top 10 customer ranking via `=INDEX(..., MATCH(LARGE(...)))` |
| **`Task 8`** | Task 8 | Regional & Product Cross-Tabulation | Two-dimensional matrix of Total Sales by Region $\times$ Category |
| **`Task 9`** | Task 9 | KPI Visualizations & Data Storytelling | Clustered Column/Bar, Line Trend, and Proportion Pie Charts |
| **`Task 10`** | Task 10 | Interactive Executive Dashboard | Cockpit with 4 KPI Cards, 5 Charts, 3 Slicers & 1 Timeline |
| **`pivot tables`** | Engine | Centralized Analytical Pivot Engine | 5 Specialized Pivot Tables feeding dashboard components |

---

## 🚀 Quick Navigation — Core Tasks Summary

| # | Task Description | Primary Tool / Formulas | Primary Business Output |
|:---:|:---|:---|:---|
| **1** | Top 10 Customers Highlight | Pivot Table, Conditional Formatting (`Top 10 Items`) | Top 10 customers identified ($11.6K to $7.8K sales) |
| **2** | What-If Discount Sensitivity | What-If Data Table, `=SUM(Sheet1!I2:I201)` | $10,901.36 profit reduction per 5% discount increase |
| **3** | Linear Regression (Profit vs Sales) | Data Analysis Toolpak (Regression) | $R^2 = 0.5983$, Slope = $+0.3511$, $F = 294.85$ |
| **4** | Descriptive Statistics Summary | Data Analysis Toolpak (Descriptive Stats) | Sales Mean = $976.09, Profit Mean = $341.85 |
| **5** | Directional Up/Down Arrows | Pivot Table, Conditional Formatting (`3Arrows` Icon Set) | Instant visual categorization of monthly sales |
| **6** | Dynamic System Timestamp | Date & Time Formula: `=NOW()` | Live execution timestamp for audit trail |
| **7** | High-Value Customer Lookup | `=INDEX()`, `=MATCH()`, `=LARGE()` | Top individual transactions ranked #1 to #10 |
| **8** | Region $\times$ Product Matrix | Two-Way Pivot Table (Region $\times$ Product Category) | West leads ($44.1K), Books leads ($47.0K) |
| **9** | Multi-Chart Data Storytelling | Clustered Bar, Line, and Pie Visualizations | Comprehensive visual breakdown of retail KPIs |
| **10** | Interactive Executive Dashboard | 4 Metric Cards, 5 Charts, 3 Slicers, 1 Date Timeline | Centralized business intelligence command center |

---

# 🔬 Detailed Task Implementation & Analytical Findings

---

### Task 1: Conditional Formatting — Top 10 Customers

#### 📌 Objective
Automatically highlight the top 10 revenue-generating customer accounts out of all 30 corporate accounts to streamline VIP customer management and resource allocation.

#### ⚙️ Technical Implementation
1. Built an aggregated Pivot Table on `Task 1` drawing from `Sheet1`:
   - **Row Field:** `Customer_ID` (30 distinct customer entities: `CUST001` through `CUST030`).
   - **Values Field:** `Sum of Sales`.
2. Applied **Conditional Formatting** to the revenue range `B4:B33`:
   - **Rule Type:** Top/Bottom Rules $\rightarrow$ `Top 10 Items...` (`type="top10"`, `rank="10"`).
   - **Formatting:** Light Red / Soft Green fill with dark accent text to clearly differentiate elite accounts.

#### 📊 Top 10 Customer Accounts Identified

| Rank | Customer ID | Total Revenue ($) | % of Total Portfolio |
|:---:|:---:|:---:|:---:|
| **1** | `CUST015` | **$11,611.06** | 5.95% |
| **2** | `CUST028` | **$10,293.87** | 5.27% |
| **3** | `CUST007` | **$9,941.88** | 5.09% |
| **4** | `CUST008` | **$9,760.80** | 5.00% |
| **5** | `CUST026` | **$9,371.81** | 4.80% |
| **6** | `CUST021` | **$8,740.18** | 4.48% |
| **7** | `CUST003` | **$8,490.18** | 4.35% |
| **8** | `CUST012` | **$8,339.00** | 4.27% |
| **9** | `CUST030` | **$7,928.71** | 4.06% |
| **10** | `CUST004` | **$7,813.59** | 4.00% |

> **💡 Key Business Insight:** The top 10 customers contribute **$92,291.08**, representing **47.28%** of the entire business revenue ($195,217.76) from just one-third of the total customer base. These 10 accounts represent the core of the business and require dedicated account management.

---

### Task 2: What-If Analysis — Discount Impact on Profit

#### 📌 Objective
Assess how promotional discounting impacts net profitability across the organization and establish clear discount ceiling guidelines.

#### ⚙️ Technical Implementation
1. **Base Parameters:**
   - Base Test Discount Rate (Cell `A2`): `10.0%` (`0.10`).
   - Baseline Total Net Profit (Cell `A5`):
     ```excel
     =SUM(Sheet1!I2:I201)
     ```
     Yielding a baseline net profit of **$68,369.90**.
2. **What-If Scenario Simulation Table:**
   - Configured a single-variable sensitivity table testing discount rates from 0% to 20% in 5% increments:

| Scenario | Discount Rate | Simulated Total Profit ($) | Absolute Change ($) | % Margin Shift vs Base |
|:---:|:---:|:---:|:---:|:---:|
| **Zero Discount** | `0.0%` | **$91,179.33** | +$22,809.43 | **+33.36%** |
| **Conservative** | `5.0%` | **$80,277.97** | +$11,908.07 | **+17.42%** |
| **Baseline** | `10.0%` | **$69,376.61** | Baseline | **0.00%** |
| **Moderate** | `15.0%` | **$58,475.25** | -$10,901.36 | **-15.71%** |
| **Aggressive** | `20.0%` | **$47,573.89** | -$21,802.72 | **-31.42%** |

> **💡 Key Business Insight:** Profitability degrades strictly linearly by **-$10,901.36** for every **5%** increase in average discount rate. Moving from 0% discount to 20% discount eliminates **$43,605.44** in net profit (a **47.8% profit reduction**). Unless higher discounts drive substantial volume increases, discounts greater than 10% erode overall earnings.

---

### Task 3: Linear Regression — Profit vs Sales (Data Analysis Toolpak)

#### 📌 Objective
Perform Ordinary Least Squares (OLS) linear regression to quantify the relationship between **Sales** ($X$, Independent Variable) and **Profit** ($Y$, Dependent Variable) and calculate expected margins.

#### ⚙️ Technical Implementation
- Tool: **Excel Data Analysis Toolpak** $\rightarrow$ `Regression`.
- Input $Y$ Range: `Sheet1!$I$1:$I$201` (Profit).
- Input $X$ Range: `Sheet1!$E$1:$E$201` (Sales).
- Options: Labels checked, 95% Confidence Level, Output Range: `Task 3!$A$1`.

#### 📊 Regression Statistics Output

| Metric | Output Value | Statistical Interpretation |
|:---|:---:|:---|
| **Multiple R** | **0.77347** | Strong positive linear correlation between Sales and Profit. |
| **R Square ($R^2$)** | **0.59826** | **59.83%** of total profit variation is explained by sales volume. |
| **Adjusted R Square** | **0.59623** | Adjusted for degrees of freedom, confirming strong model stability. |
| **Standard Error** | **86.3312** | Average deviation of observed values from the regression line is $86.33. |
| **Observations** | **200** | Full statistical sample size analyzed. |

#### 📈 ANOVA Table (Analysis of Variance)

| Source | df | Sum of Squares (SS) | Mean Square (MS) | F-Statistic | Significance F |
|:---|:---:|:---:|:---:|:---:|:---:|
| **Regression** | 1 | 2,197,549.54 | 2,197,549.54 | **294.85** | **$4.51 \times 10^{-41}$** |
| **Residual** | 198 | 1,475,707.55 | 7,453.07 | — | — |
| **Total** | 199 | 3,673,257.08 | — | — | — |

#### 🧮 Model Parameter Estimates

| Parameter | Coefficient | Std Error | t-Statistic | P-Value | 95% Lower | 95% Upper |
|:---|:---:|:---:|:---:|:---:|:---:|:---:|
| **Intercept ($\beta_0$)** | **-0.8895** | 20.8727 | -0.0426 | 0.9661 | -42.0508 | +40.2718 |
| **Sales ($\beta_1$)** | **+0.3511** | 0.0204 | **17.1712** | **$4.51 \times 10^{-41}$** | +0.3108 | +0.3915 |

#### 📐 Mathematical Formulation & Business Meaning

$$\widehat{\text{Profit}} = 0.3511 \times \text{Sales} - 0.8895$$

1. **Marginal Contribution Rate ($\beta_1 = 0.3511$):** For every additional **$100.00** generated in top-line sales, operating profit increases by **$35.11**. The 95% confidence interval shows the true marginal contribution lies reliably between 31.1% and 39.2%.
2. **Baseline Fixed Overhead ($\beta_0 = -0.8895$):** The intercept is statistically indistinguishable from zero ($p = 0.966 > 0.05$). This indicates that the business has virtually zero fixed costs tied to zero-sales conditions, behaving like a variable-cost retail model.
3. **Model Significance ($p = 4.51 \times 10^{-41}$):** The $F$-test value of 294.85 confirms an established, non-random linear dependency.

---

### Task 4: Advanced Descriptive Statistics (Data Analysis Toolpak)

#### 📌 Objective
Perform a comprehensive parametric and non-parametric statistical examination of all numerical fields in the transaction ledger.

#### ⚙️ Technical Implementation
- Tool: **Excel Data Analysis Toolpak** $\rightarrow$ `Descriptive Statistics`.
- Input Range: `Sheet1!$E$1:$I$201` (Sales, Quantity, Discount, Profit).
- Grouped By: Columns, Summary Statistics checked, Output Range: `Task 4!$A$1`.

#### 📊 Complete Summary Statistics Table

| Statistical Metric | 💵 Sales ($) | 📦 Quantity (Units) | 🏷️ Discount (%) | 📈 Profit ($) |
|:---|:---:|:---:|:---:|:---:|
| **Mean** | **$976.09** | **9.99** (~10 units) | **9.73%** (0.0973) | **$341.85** |
| **Standard Error** | $21.16 | 0.40 | 0.51% (0.0051) | $9.61 |
| **Median** | **$976.72** | **10.00** | **10.00%** (0.10) | **$332.69** |
| **Mode** | *#N/A* (Continuous) | **8** | **20.00%** (0.20) | **$414.44** |
| **Standard Deviation** | **$299.27** | **5.65** | **7.23%** (0.0723) | **$135.86** |
| **Sample Variance** | 89,564.78 | 31.88 | 0.0052 | 18,458.58 |
| **Kurtosis** | **-0.0886** | **-1.2091** | **-1.3257** | **+0.2485** |
| **Skewness** | **+0.1469** | **+0.0608** | **+0.0968** | **+0.5637** |
| **Range** | **$1,665.15** | **18** | **20.00%** (0.20) | **$704.83** |
| **Minimum** | **$258.51** | **1** | **0.00%** (0.00) | **$64.63** |
| **Maximum** | **$1,923.66** | **19** | **20.00%** (0.20) | **$769.46** |
| **Sum** | **$195,217.76** | **1,999** | **19.45** | **$68,369.90** |
| **Count** | **200** | **200** | **200** | **200** |

#### 🔎 Key Statistical Observations
1. **Sales Symmetry:** Mean ($976.09) and Median ($976.72) are nearly identical, with a skewness of +0.15, demonstrating an evenly balanced bell curve across transaction sizes.
2. **Profit Skewness:** Profit exhibits a moderate positive skew (+0.56) and higher median-to-maximum spread ($332.69 to $769.46), reflecting a tail of high-margin transactions.
3. **Discount Distribution:** While the average discount is 9.73%, the mode is 20.0%, indicating that when sales reps apply discounts, they tend to offer the maximum permitted rate.

---

### Task 5: Monthly Sales Growth with Directional Trend Indicators

#### 📌 Objective
Provide visual indicators showing monthly sales performance and growth trajectory using icon sets and custom formatting.

#### ⚙️ Technical Implementation
1. Configured a monthly timeline Pivot Table on `Task 5`:
   - **Report Filter:** `Order_Date` (Grouped by Year & Month).
   - **Values:** Monthly aggregate transaction volume and order counts.
2. Applied **3-Arrows Icon Set** Conditional Formatting (`iconSet: 3Arrows`):
   - 🟢 **Green Up Arrow (▲):** Values $\ge 67\text{th}$ Percentile (High Performance / Top Third).
   - 🟡 **Yellow Side Arrow (➔):** Values $\ge 33\text{rd}$ and $< 67\text{th}$ Percentile (Moderate Stability).
   - 🔴 **Red Down Arrow (▼):** Values $< 33\text{rd}$ Percentile (Low / Contraction Zone).
3. **Alternative Custom Number Formatting Rule:**
   ```excel
   [Color10]▲ #,##0.00%;[Red]▼ -#,##0.00%;[Blue]– 0.00%
   ```

> **💡 Visual Impact:** Enables leadership to spot cyclical dips (e.g., July/August seasonal troughs) and surge periods (e.g., September/November quarter-end spikes) at a glance without reading raw numeric columns.

---

### Task 6: Dynamic System Timestamp

#### 📌 Objective
Incorporate an automated, real-time workbook timestamp to verify audit currency and ensure data freshness during executive presentations.

#### ⚙️ Formula Applied

```excel
=NOW()
```

#### ⚙️ Implementation & Format Details
- **Cell Range:** `Task 6!A1:A8`.
- **Formatting:** Configured with custom date-time format:
  ```excel
  yyyy-mm-dd hh:mm:ss AM/PM
  ```
- **Execution Output:** Dynamic serial value `46279.57`, automatically re-evaluating on sheet calculation or workbook opening to record the precise date and time of analysis.

> **💡 Best Practice Note:** For fixed historical audit freezes, copy and `Paste Special > Values` to prevent unwanted updates, while keeping `=NOW()` live on active operational dashboards.

---

### Task 7: High-Value Customer Identification (`INDEX`, `MATCH`, `LARGE`)

#### 📌 Objective
Dynamically locate, rank, and retrieve top individual customer transactions without manually sorting the master dataset, preserving full formula reactivity.

#### ⚙️ Formulas Applied

In Cell `C2` (Extracting $k$-th Largest Sales Value):
```excel
=LARGE(Sheet1!E:E, A2)
```

In Cell `B2` (Dynamic Two-Way Lookup of Corresponding Customer Name):
```excel
=INDEX(Sheet1!B:B, MATCH(C2, Sheet1!E:E, 0))
```

#### 📊 Formula Breakdown
- **`LARGE(Sheet1!E:E, A2)`**: Scans the entire Sales column (`E:E`) on `Sheet1` and returns the numerical value corresponding to rank `A2` (1st largest, 2nd largest, etc.).
- **`MATCH(C2, Sheet1!E:E, 0)`**: Locates the exact row index of the sales figure returned in `C2`.
- **`INDEX(Sheet1!B:B, ...)`**: Retrieves the customer's full name from column `B` at that matching row position.

#### 🏆 Top 10 High-Value Transactions

| Rank (`A`) | Customer Name (`B`) | Dynamic Formula in `B` | Transaction Amount (`C`) | Dynamic Formula in `C` |
|:---:|:---|:---|:---:|:---|
| **1** | **James Wilson** | `=INDEX(Sheet1!B:B, MATCH(C2, Sheet1!E:E, 0))` | **$1,923.66** | `=LARGE(Sheet1!E:E, A2)` |
| **2** | **James Wilson** | `=INDEX(Sheet1!B:B, MATCH(C3, Sheet1!E:E, 0))` | **$1,681.21** | `=LARGE(Sheet1!E:E, A3)` |
| **3** | **Linda Taylor** | `=INDEX(Sheet1!B:B, MATCH(C4, Sheet1!E:E, 0))` | **$1,622.62** | `=LARGE(Sheet1!E:E, A4)` |
| **4** | **Mary Jackson** | `=INDEX(Sheet1!B:B, MATCH(C5, Sheet1!E:E, 0))` | **$1,618.22** | `=LARGE(Sheet1!E:E, A5)` |
| **5** | **Mary Davis** | `=INDEX(Sheet1!B:B, MATCH(C6, Sheet1!E:E, 0))` | **$1,589.42** | `=LARGE(Sheet1!E:E, A6)` |
| **6** | **Emma Davis** | `=INDEX(Sheet1!B:B, MATCH(C7, Sheet1!E:E, 0))` | **$1,572.82** | `=LARGE(Sheet1!E:E, A7)` |
| **7** | **Emma Davis** | `=INDEX(Sheet1!B:B, MATCH(C8, Sheet1!E:E, 0))` | **$1,526.60** | `=LARGE(Sheet1!E:E, A8)` |
| **8** | **Mary Jackson** | `=INDEX(Sheet1!B:B, MATCH(C9, Sheet1!E:E, 0))` | **$1,506.14** | `=LARGE(Sheet1!E:E, A9)` |
| **9** | **Sarah Anderson** | `=INDEX(Sheet1!B:B, MATCH(C10, Sheet1!E:E, 0))` | **$1,502.93** | `=LARGE(Sheet1!E:E, A10)` |
| **10** | **David Brown** | `=INDEX(Sheet1!B:B, MATCH(C11, Sheet1!E:E, 0))` | **$1,493.49** | `=LARGE(Sheet1!E:E, A11)` |

> **💡 Key Finding:** James Wilson holds the top two single transactions in company history ($1,923.66 and $1,681.21). Repeat appearances by Emma Davis and Mary Jackson highlight strong multi-order loyalty among elite buyers.

---

### Task 8: Regional & Product Cross-Tabulation Pivot Table

#### 📌 Objective
Construct a two-dimensional matrix to analyze total sales performance across the 5 geographic regions and 5 product lines.

#### ⚙️ Technical Implementation
- **Data Source:** `Sheet1!$A$1:$I$201`.
- **Row Labels:** `Region` (Central, East, North, South, West).
- **Column Labels:** `Product_Category` (Books, Clothing, Electronics, Furniture, Office Supplies).
- **Values:** `Sum of Sales` (formatted as Currency `$#,##0.00`).

#### 📊 Two-Dimensional Sales Matrix Output ($)

| Region | 📚 Books | 👕 Clothing | 💻 Electronics | 🪑 Furniture | 📎 Office Supplies | 🌐 Grand Total | % Share |
|:---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| **Central** | $11,297.15 | **$12,200.26** | $6,621.24 | $4,876.96 | $8,472.72 | **$43,468.33** | 22.27% |
| **East** | **$11,389.09** | $2,401.45 | $2,869.54 | $7,592.34 | $8,189.22 | **$32,441.64** | 16.62% |
| **North** | $8,037.01 | $8,422.82 | $8,191.71 | $6,978.08 | $3,345.80 | **$34,975.42** | 17.92% |
| **South** | $9,588.04 | $8,393.63 | **$11,989.07** | $7,982.73 | $2,295.32 | **$40,248.79** | 20.62% |
| **West** | $6,691.17 | $10,509.43 | **$11,372.44** | **$9,182.78** | $6,327.76 | **$44,083.58** | **22.58%** |
| **Grand Total** | **$47,002.46** | **$41,927.59** | **$41,044.00** | **$36,612.89** | **$28,630.82** | **$195,217.76** | **100.0%** |
| **Category %** | **24.08%** | **21.48%** | **21.02%** | **18.75%** | **14.67%** | **100.0%** | — |

> **💡 Key Takeaways:**
> - **Top Regional Market:** **West** ($44,083.58) and **Central** ($43,468.33) drive a combined **44.85%** of all revenue.
> - **Top Category:** **Books** leads all categories ($47,002.46), driven by strong sales in East ($11.4K) and Central ($11.3K).
> - **Category Hotspots:** **Electronics** dominates the South ($11,989.07) and West ($11,372.44), while **Clothing** peaks in Central ($12,200.26).
> - **Vulnerability:** East Clothing ($2,401.45) and South Office Supplies ($2,295.32) represent notable underperformance requiring targeted promotions.

---

### Task 9: KPI Visualizations & Multi-Chart Storytelling

#### 📌 Objective
Create bar, line, and pie charts to visualize key performance indicators (KPIs) across regions and product categories.

<div align="center">

![Task 9 KPI Visualizations](assets/task9_visualizations.png)

</div>

> **💡 Visualization Breakdown:**
> - **Clustered Column Chart (Left):** Total sales by Region broken down across all 5 Product Categories.
> - **Line Trend Chart (Middle):** Regional sales trajectory comparison across product lines.
> - **Pie Chart (Right):** Regional market share contribution for the top revenue category (**Books**).

---

### Task 10: Interactive Executive Dashboard

#### 📌 Objective
Summarize organizational performance and business insights using a centralized, interactive executive dashboard equipped with KPI cards, dynamic charts, slicers, and a date timeline.

<div align="center">

![Task 10 Executive Interactive Dashboard](assets/task10_dashboard.png)

</div>

> **💡 Executive Dashboard Highlights:**
> - **KPI Scorecards:** Total Sales (**$195,217.76**), Total Profit (**$68,287.01**), Units Sold (**1,999**), and Profit Margin (**34.98%**).
> - **Interactive Controls:** Dynamic slicers for **Region**, **Customer Name**, **Product Category**, and an **Order Date Timeline**.
> - **Synchronized Analytics Charts:** Profit by Category, Sales by Region, Sales by Category, Regional Sales Distribution, and Monthly Sales Trend.

---

# 🧠 Strategic Insights & Business Recommendations

### 1. Protect Pricing & Discount Margins
- **Insight:** Linear regression shows an operating contribution rate of **35.11%**, but the What-If analysis demonstrates that raising discounts from 10% to 20% slashes net profits by **$21,802.72** (nearly a 31.4% profit drop).
- **Action:** Implement a hard discount ceiling of **10%** for standard retail reps. Discounts between 10% and 20% should require managerial sign-off and minimum volume commitments ($\ge 15$ units).

### 2. Double Down on High-Yield Regional Lines
- **Insight:** **West** ($44,083.58) and **Central** ($43,468.33) generate nearly half of all company sales. Electronics is strongest in South and West ($23.3K combined), while Clothing delivers highest sales in Central ($12.2K) and top overall profit ($18.8K).
- **Action:** Shift inventory allocations for Electronics to Western and Southern distribution hubs, and target Clothing promotions in Central territories where conversion is highest.

### 3. VIP Customer Retention & Loyalty Programs
- **Insight:** The top 10 accounts generate **$92,291.08** (**47.28%** of total business revenue). James Wilson alone accounts for multiple top-tier transactions ($1,923.66 and $1,681.21).
- **Action:** Establish a Dedicated Account Management tier for the top 10 accounts (e.g., `CUST015`, `CUST028`, `CUST007`, `CUST008`, `CUST026`). Offer volume rebates rather than front-end discounts to preserve operational margins.

### 4. Turn Around Underperforming Product Lines
- **Insight:** East Clothing generated only **$2,401.45** and South Office Supplies generated only **$2,295.32** across the entire 12-month period.
- **Action:** Audit local marketing, inventory assortments, and competitor pricing in the East and South to address these regional performance gaps.

---

# 📁 Project Directory Structure

```text
Excel PR. 2 Analyzer/
│
├── assets/
│   ├── task9_visualizations.png   # Task 9 KPI charts screenshot
│   └── task10_dashboard.png       # Task 10 Executive interactive dashboard screenshot
├── PR. 2 Analyzer Tasks.xlsx
├── PR. 2.xlsx                     # Complete practical workbook with all 12 sheets & dashboard
└── README.md                      # Comprehensive project documentation & analytical report
```

---

# 💻 How to Use & Review This Project

1. **Open `PR. 2.xlsx` in Microsoft Excel 2016, 2019, 2021, or Microsoft 365** (recommended for full Timeline and Slicer interactivity).
2. **Review Task Sheets (`Task 1` to `Task 10`):**
   - Click on any calculation cell (e.g., `Task 2!A5`, `Task 6!A1`, `Task 7!B2:C11`) to inspect the underlying Excel formulas.
3. **Explore the Regression & Stats (`Task 3` & `Task 4`):**
   - Review the complete ANOVA tables, $p$-values, confidence intervals, and descriptive statistics generated via the Data Analysis Toolpak.
4. **Interact with the Executive Dashboard (`Task 10`):**
   - Click any button on the **Region**, **Category**, or **Customer** Slicers to filter the entire dashboard in real time.
   - Adjust the **Order_Date Timeline** bar to analyze performance by month or quarter.
   - Note how all 4 KPI cards and 5 charts update simultaneously across linked pivot caches.

---

<div align="center">

### 🏆 ALL 10 MAIN TASKS COMPLETED SUCCESSFULLY

**PR. 2 Analyzer — Advanced Excel Business Analytics & Executive Dashboard**  
*Prepared by Priyanshu Buhecha*

---

### 🚀 Data Analysis • Statistical Modeling • Visual Storytelling • Business Intelligence

</div>
