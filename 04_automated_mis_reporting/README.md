# 📋 Automated MIS Reporting System

## Overview
This project simulates an automated Management Information System (MIS) that pulls data
from multiple Excel source files (Sales, HR, Finance), performs aggregations, and generates
a consolidated MIS report with a cover sheet, KPI summary, and 4 department dashboards.

## Features
- Auto-generates 3 source Excel files (Sales, HR, Finance) — simulating real data feeds
- Reads all sources and computes KPIs automatically
- Sales: Revenue vs Target, Achievement %, Deals Won, Lead Count
- HR: Headcount, Attendance %, Attrition Rate, Net Joiner/Exits
- Finance: YTD Revenue, Profit, Margins, Outstanding AR
- 4-chart MIS Dashboard (PNG): Revenue vs Target, Attendance, Profit Margin, YTD Finance
- Consolidated Excel MIS Report with Cover Sheet + 4 department sheets

## Technologies Used
| Tool | Purpose |
|------|---------|
| Python 3 | Core scripting & automation |
| Pandas | Data pull, merge, aggregation |
| NumPy | Data generation |
| Matplotlib | MIS Dashboard chart |
| OpenPyXL | Excel report creation |

## Project Structure
```
04_automated_mis_reporting/
│
├── mis_reporting.py                ← Main script
├── source_sales.xlsx               ← Auto-generated sales source
├── source_hr.xlsx                  ← Auto-generated HR source
├── source_finance.xlsx             ← Auto-generated finance source
├── MIS_Report_Consolidated.xlsx    ← Final consolidated MIS report
├── mis_dashboard.png               ← 4-panel dashboard chart
└── README.md
```

## How to Run

### Prerequisites
```bash
pip install pandas numpy matplotlib openpyxl
```

### Run
```bash
cd 04_automated_mis_reporting
python mis_reporting.py
```

## Output Files
| File | Description |
|------|-------------|
| `MIS_Report_Consolidated.xlsx` | 5-sheet Excel MIS report with cover |
| `mis_dashboard.png` | 4-panel KPI dashboard chart |
| `source_*.xlsx` | Auto-generated source data files |

## Real-World Integration
To connect to real data sources:
1. Replace the source generation section with `pd.read_excel("your_file.xlsx")`
2. Schedule the script using Windows Task Scheduler or cron (Linux/Mac)
3. Import the output into Power BI with auto-refresh enabled
4. Use Power Query to link directly to the output Excel file
