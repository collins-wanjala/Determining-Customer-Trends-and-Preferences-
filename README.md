# Determining-Customer-Trends-and-Preferences-

Determining Customer Trends  and Preferences , using customer dataset, appended all the reports



\# 📊 Customer Analytics — Python \& Excel Data Analysis Project



> A full end-to-end data analytics pipeline built with \*\*Python\*\*, \*\*pandas\*\*, \*\*matplotlib\*\*, and \*\*openpyxl\*\* —

&nbsp;from raw CSV upload to a styled multi-sheet Excel report, embedded charts, and a professional PDF Q\&A document.



---



\## 🗂️ Table of Contents



\- \[Project Overview](#-project-overview)

\- \[Dataset](#-dataset)

\- \[Features](#-features)

\- \[Project Structure](#-project-structure)

\- \[Getting Started](#-getting-started)

\- \[Step-by-Step Workflow](#-step-by-step-workflow)

\- \[Output Files](#-output-files)

\- \[Charts Generated](#-charts-generated)

\- \[Excel Report Sheets](#-excel-report-sheets)

\- \[Tech Stack](#-tech-stack)

\- \[Key Findings](#-key-findings)

\- \[Questions \& Answers PDF](#-questions--answers-pdf)

\- \[Contributing](#-contributing)

\- \[License](#-license)



---



**## 🔍 Project Overview**



This project demonstrates a complete data analytics workflow applied to a customer subscription dataset. It covers every stage of the analytics lifecycle:



1\. \*\*Data loading \& inspection\*\* — shape, dtypes, summary statistics

2\. \*\*Missing value detection\*\* — zero nulls confirmed across all columns

3\. \*\*Duplicate checking\*\* — full-row and key-column deduplication

4\. \*\*Date parsing \& feature engineering\*\* — year, month, quarter, day-of-week extraction

5\. \*\*Aggregation \& trend analysis\*\* — monthly, quarterly, and annual subscription trends

6\. \*\*Geographic analysis\*\* — customer distribution across 224 countries

7\. \*\*Visualisation\*\* — 7 publication-quality matplotlib charts

8\. \*\*Excel export\*\* — styled multi-sheet workbook with embedded charts and a dashboard

9\. \*\*PDF report\*\* — 25 Q\&A questions with charts, tables, and business insights



---



**## 📁 Dataset**



| Property | Value |

|---|---|

| \*\*File\*\* | `customers-1000.csv` |

| \*\*Rows\*\* | 1,000 |

| \*\*Columns\*\* | 12 |

| \*\*Date range\*\* | January 2020 – May 2022 |

| \*\*Unique countries\*\* | 224 |

| \*\*Missing values\*\* | 0 |

| \*\*Duplicate records\*\* | 0 |



**### Columns**



| Column | Type | Description |

|---|---|---|

| `Index` | int | Sequential row number |

| `Customer Id` | string | Unique customer identifier |

| `First Name` | string | Customer first name |

| `Last Name` | string | Customer last name |

| `Company` | string | Company name |

| `City` | string | Customer city |

| `Country` | string | Customer country |

| `Phone 1` | string | Primary phone number |

| `Phone 2` | string | Secondary phone number |

| `Email` | string | Unique email address |

| `Subscription Date` | datetime | Date of subscription |

| `Website` | string | Customer website URL |



---



**## ✨ Features**



\- ✅ Full data quality audit (missing values + duplicates)

\- ✅ Automated date feature extraction (year, month, quarter, day of week)

\- ✅ Multi-year trend comparison with line charts

\- ✅ Geographic distribution analysis across 224 countries

\- ✅ 7 high-resolution matplotlib charts (160 DPI)

\- ✅ Professional Excel workbook with 11 sheets, auto-styling, and frozen panes

\- ✅ Embedded charts in every relevant Excel sheet + a master dashboard

\- ✅ 25-question analytics Q\&A PDF with a cover page, table of contents, and colour-coded sections

\- ✅ All code is modular and reusable on any similar CSV dataset



---



**## 📂 Project Structure**



**```**

customer-analytics/

│

├── data/

│   └── customers-1000.csv          # Raw input dataset

│

├── charts/

│   ├── chart1\_monthly.png          # Multi-year monthly line chart

│   ├── chart2\_year.png             # Subscriptions by year (bar)

│   ├── chart3\_month.png            # Subscriptions by calendar month (bar)

│   ├── chart4\_country.png          # Top 15 countries (horizontal bar)

│   ├── chart5\_dow.png              # Subscriptions by day of week (bar)

│   ├── chart6\_quarterly.png        # Quarterly breakdown (bar)

│   └── chart7\_missing.png          # Missing values audit (bar)

│

├── outputs/

│   ├── analytics\_full\_with\_charts.xlsx   # Styled Excel report with charts

│   └── customer\_analytics\_qa.pdf         # 25-question Q\&A PDF report

│

├── notebooks/

│   └── analysis.ipynb              # Jupyter notebook version (optional)

│

├── analytics.py                    # Main analysis script

├── charts.py                       # Chart generation script

├── export\_excel.py                 # Excel export \& styling script

├── export\_pdf.py                   # PDF report generation script

├── requirements.txt                # Python dependencies

└── README.md                       # This file

```



---



\## 🚀 Getting Started



\### Prerequisites



\- Python 3.8 or higher

\- pip



\### Installation



```bash

\# 1. Clone the repository

git clone https://github.com/your-username/customer-analytics.git

cd customer-analytics



\# 2. Create a virtual environment (recommended)

python -m venv venv

source venv/bin/activate        # macOS/Linux

venv\\Scripts\\activate           # Windows



\# 3. Install dependencies

pip install -r requirements.txt

```



\### Requirements



```txt

pandas>=1.5.0

numpy>=1.23.0

matplotlib>=3.6.0

seaborn>=0.12.0

openpyxl>=3.0.10

reportlab>=3.6.0

Pillow>=9.0.0

```



\### Run the full pipeline



```bash

\# Run everything in one go

python analytics.py



\# Or run each step individually

python charts.py          # Generate charts only

python export\_excel.py    # Build Excel report

python export\_pdf.py      # Build PDF report

```



---



**## 🔄 Step-by-Step Workflow**



\### Step 1 — Load \& Inspect



```python

import pandas as pd



df = pd.read\_csv("data/customers-1000.csv")

print(df.shape)           # (1000, 12)

print(df.dtypes)

print(df.describe(include="all"))

```



\### Step 2 — Find Missing Values



```python

missing = df.isnull().sum()

pct     = (df.isnull().sum() / len(df) \* 100).round(2)

print(pd.concat(\[missing, pct], axis=1, keys=\["Count", "%"]))

\# Result: all zeros — dataset is clean

```



\### Step 3 — Parse Dates \& Extract Features



```python

df\["Subscription Date"] = pd.to\_datetime(df\["Subscription Date"])

df\["Year"]      = df\["Subscription Date"].dt.year

df\["Month"]     = df\["Subscription Date"].dt.month

df\["MonthName"] = df\["Subscription Date"].dt.strftime("%b")

df\["DayOfWeek"] = df\["Subscription Date"].dt.day\_name()

df\["Quarter"]   = df\["Subscription Date"].dt.to\_period("Q").astype(str)

```



\### Step 4 — Subscription Trends



```python

monthly = df.groupby("YearMonth").size().reset\_index(name="Count")

pivot   = df.pivot\_table(index="Year", columns="Month",

&nbsp;                         aggfunc="size", fill\_value=0)

```



**### Step 5 — Geographic Analysis**



```python

top\_countries = df\["Country"].value\_counts().head(10)

print(f"Unique countries: {df\['Country'].nunique()}")  # 224

```



**### Step 6 — Duplicate Detection**



```python

print(df.duplicated().sum())                           # 0

print(df.duplicated(subset=\["Email"]).sum())           # 0

print(df.duplicated(subset=\["Customer Id"]).sum())     # 0

```



**### Step 7 — Export Styled Excel**



```python

from openpyxl import load\_workbook

from openpyxl.styles import Font, PatternFill, Alignment



\# Write sheets

with pd.ExcelWriter("outputs/report.xlsx", engine="openpyxl") as writer:

&nbsp;   df.to\_excel(writer, sheet\_name="Raw Data", index=False)

&nbsp;   monthly.to\_excel(writer, sheet\_name="Monthly Trend", index=False)

&nbsp;   # ... additional sheets



\# Style with openpyxl

wb = load\_workbook("outputs/report.xlsx")

ws = wb\["Raw Data"]

ws.freeze\_panes = "A2"

ws.auto\_filter.ref = ws.dimensions

wb.save("outputs/report.xlsx")

```

**### Step 8 — Embed Charts \& Build Dashboard**



```python

from openpyxl.drawing.image import Image as XLImage



img = XLImage("charts/chart1\_monthly.png")

img.width, img.height = 700, 270

ws\_dashboard.add\_image(img, "A4")

```



---



\## 📤 Output Files



| File | Description | Size |

|---|---|---|

| `analytics\_full\_with\_charts.xlsx` | 11-sheet styled Excel report with 7 embedded charts | ~886 KB |

| `customer\_analytics\_qa.pdf` | 25-question Q\&A PDF with cover, TOC, 5 sections, 5 charts | ~211 KB |



---



\## 📈 Charts Generated



| # | Chart | Type | Key Insight |

|---|---|---|---|

| 1 | Monthly subscriptions 2020–2022 | Multi-series line | Consistent volume; April peaks |

| 2 | Subscriptions by year | Bar | 2020: 426, 2021: 404, 2022: 170 (partial) |

| 3 | Subscriptions by calendar month | Bar | April is the peak month (127 total) |

| 4 | Top 15 countries | Horizontal bar | Liechtenstein leads with 12 customers |

| 5 | Subscriptions by day of week | Bar | Mid-week outperforms weekends |

| 6 | Quarterly breakdown | Grouped bar | Stable ~100/quarter across 2020–2021 |

| 7 | Missing values audit | Bar | All columns: 0% missing |



---



\## 📋 Excel Report Sheets



| Sheet | Contents |

|---|---|

| 📊 Dashboard | All 7 charts on one page with summary header |

| Raw Data | 1,000 rows — styled, filtered, frozen header, alternating rows |

| Missing Values | Null counts per column + chart |

| Duplicates | 3-check deduplication audit table |

| Monthly Trend | Count per month + embedded line chart |

| By Year | Annual totals + bar chart |

| By Month | Jan–Dec totals + bar chart |

| By Country | Top 15 countries + data bars + horizontal bar chart |

| By Day of Week | Weekday subscription counts + bar chart |

| By Quarter | Quarterly counts + bar chart |

| Pivot Year x Month | Year × Month cross-tab with amber heat-map |



---



\## 🛠️ Tech Stack



| Tool | Purpose |

|---|---|

| `pandas` | Data loading, cleaning, groupby, pivot tables |

| `numpy` | Numerical operations |

| `matplotlib` | Chart generation (7 charts at 160 DPI) |

| `seaborn` | Statistical visualisation styling |

| `openpyxl` | Excel file creation, styling, chart embedding |

| `reportlab` | PDF generation (cover page, tables, images, TOC) |

| `Pillow` | Image handling for Excel/PDF embedding |



---



\## 💡 Key Findings



\- 📅 \*\*April\*\* is the strongest subscription month across all years (127 total, peak of 51 in April 2022)

\- 📆 \*\*June and August\*\* are the quietest months (57 subscriptions each)

\- 🌍 \*\*224 unique countries\*\* — no dominant market; max 12 customers from any single country (Liechtenstein)

\- 📊 \*\*2020\*\* had the highest annual subscriptions (426), closely followed by 2021 (404)

\- ✅ \*\*Zero data quality issues\*\* — no missing values, no duplicates across any column

\- 📆 \*\*Mid-week days\*\* show higher subscription rates than weekends



---



\## ❓ Questions \& Answers PDF



The PDF report covers 25 analytical questions across 5 sections:



| Section | Topic | Questions |

|---|---|---|

| 1 | Dataset Overview | Q1–5 |

| 2 | Temporal \& Trend Analysis | Q6–12 |

| 3 | Geographic Analysis | Q13–17 |

| 4 | Data Quality | Q18–21 |

| 5 | Business Insights | Q22–25 |



Each answer references real numbers from the dataset and includes \*\*Python code tips\*\* and \*\*Excel equivalents\*\* for every analytical technique used.



---

\## 🤝 Contributing

Contributions are welcome! To contribute:



1\. Fork the repository

2\. Create a feature branch: `git checkout -b feature/your-feature-name`

3\. Commit your changes: `git commit -m "Add: your feature description"`

4\. Push to the branch: `git push origin feature/your-feature-name`

5\. Open a Pull Request



Please ensure your code follows PEP 8 style guidelines and includes comments for any new analytical steps.



---

\## 📄 License



This project is licensed under the \*\*MIT License\*\* — see the \[LICENSE](LICENSE) file for details.

---



**## 👤 Author(Collins Wanjala)**



Built as a full data analytics showcase using Python and Excel techniques — from raw CSV to professional reports.



> ⭐ If you found this project useful, consider giving it a star on GitHub!



---



