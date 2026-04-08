# 📦 Inventory Discrepancy Detection System

## Overview
This project compares warehouse stock levels with actual sales records to detect
inventory mismatches. It uses Z-score based anomaly detection to flag discrepancies
and generates color-coded alerts (Critical / Warning / Monitor / Normal).

## Features
- Warehouse data: Opening stock, received stock, closing stock per SKU
- Sales records linked to each SKU
- Discrepancy calculation: Expected vs Actual closing stock
- Z-score statistical anomaly detection
- 4-level alert classification: Critical, Warning, Monitor, Normal
- Bar charts: Discrepancy per SKU + Alert distribution
- Excel report with color-coded rows by alert level (4 sheets)

## Technologies Used
| Tool | Purpose |
|------|---------|
| Python 3 | Core scripting |
| Pandas | Data merging & calculation |
| NumPy | Z-score computation & data generation |
| Matplotlib | Visualization charts |
| OpenPyXL | Excel report with conditional formatting |

## Project Structure
```
03_inventory_discrepancy/
│
├── inventory_discrepancy.py            ← Main script
├── Inventory_Discrepancy_Report.xlsx   ← Excel report (after running)
├── inventory_discrepancy_chart.png     ← Bar charts
└── README.md
```

## How to Run

### Prerequisites
```bash
pip install pandas numpy matplotlib openpyxl
```

### Run
```bash
cd 03_inventory_discrepancy
python inventory_discrepancy.py
```

## Alert Levels
| Level | Condition |
|-------|-----------|
| 🔴 CRITICAL | Z-score > 2.0 OR Discrepancy > 20% |
| 🟠 WARNING | Z-score > 1.5 OR Discrepancy > 10% |
| 🟡 MONITOR | Z-score > 1.0 OR Discrepancy > 5% |
| 🟢 NORMAL | Within acceptable range |

## Output Files
| File | Description |
|------|-------------|
| `Inventory_Discrepancy_Report.xlsx` | 4-sheet report with color-coded alerts |
| `inventory_discrepancy_chart.png` | Discrepancy + alert distribution charts |
