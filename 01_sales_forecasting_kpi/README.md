# 📊 Sales Forecasting & KPI Dashboard

## Overview
This project analyzes 24 months of historical sales data across 5 products and 4 regions,
builds a 6-month sales forecast using Linear Regression, computes key KPIs,
and exports a fully styled Excel KPI Dashboard with charts.

## Features
- Realistic synthetic sales data generation (5 products × 4 regions × 24 months)
- Monthly aggregation: Revenue, Profit, Units Sold, Profit Margin
- 6-Month Sales Forecast with confidence bands (Linear Regression)
- Model evaluation: MAE, RMSE, R² Score
- Excel KPI Dashboard with KPI cards, data tables, embedded charts
- Matplotlib forecast chart (PNG export)

## Technologies Used
| Tool | Purpose |
|------|---------|
| Python 3 | Core scripting |
| Pandas | Data manipulation & aggregation |
| NumPy | Numerical computations |
| Scikit-learn | Linear Regression forecasting |
| Matplotlib | Forecast visualization chart |
| OpenPyXL | Excel KPI Dashboard creation |

## Project Structure
```
01_sales_forecasting_kpi/
│
├── sales_forecasting.py      ← Main script
├── KPI_Dashboard.xlsx        ← Generated Excel Dashboard (after running)
├── sales_forecast_chart.png  ← Forecast chart (after running)
└── README.md
```

## How to Run

### Prerequisites
```bash
pip install pandas numpy scikit-learn matplotlib openpyxl
```

### Run the Script
```bash
cd 01_sales_forecasting_kpi
python sales_forecasting.py
```

### Expected Output
```
Total Revenue (24 months)  : ₹ 1,23,45,67,890
Total Profit               : ₹  45,67,89,012
...
✅  Excel KPI Dashboard saved → KPI_Dashboard.xlsx
✅  Forecast chart saved → sales_forecast_chart.png
```

## Output Files
| File | Description |
|------|-------------|
| `KPI_Dashboard.xlsx` | 4-sheet Excel dashboard with KPI cards & charts |
| `sales_forecast_chart.png` | Line chart showing historical + forecast |

## KPI Metrics Computed
- Total Revenue & Profit (24 months)
- Average Profit Margin (%)
- Month-over-Month Growth Rate
- Best Product & Best Region by Sales
- 6-Month Forecast with Lower/Upper Bounds

## Power BI Integration
1. Open Power BI Desktop
2. Get Data → Excel → select `KPI_Dashboard.xlsx`
3. Load "Raw Data" sheet
4. Use DAX to create KPI measures
5. Add AI Forecasting visual from Visualizations panel
