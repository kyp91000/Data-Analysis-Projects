# 🏭 Project 3: Supply Chain Analytics — Inventory Optimization & Supplier Performance

## Overview
A comprehensive supply chain analytics platform that processes multi-year demand history across 500 SKUs and 25 suppliers. The system computes science-based inventory parameters (EOQ, safety stock, reorder points), performs ABC-XYZ classification, generates supplier scorecards, and identifies multi-million dollar cost reduction opportunities from excess inventory.

## Business Problem
Poor inventory management costs companies 25–30% of total inventory value annually through excess holding costs, stockouts causing lost sales, and supplier quality failures. This project replaces spreadsheet-based guesswork with statistically rigorous, reproducible analytics.

## What Was Done
- **Data Generation**: 500 SKUs × 25 suppliers × 24 months of demand history with realistic variability, seasonality, and trend patterns
- **EOQ Optimization**: Economic Order Quantity computed for each SKU using classical EOQ formula balancing holding cost vs. ordering cost
- **Safety Stock Calculation**: Statistically derived safety stock using demand variance + lead-time variance, calibrated to product-specific service level targets (90%, 95%, 99%)
- **Reorder Point**: Average demand during lead time + safety stock buffer
- **ABC Analysis**: Pareto classification — A items (80% of spend), B items (15%), C items (5%) — for differentiated management
- **XYZ Analysis**: Demand variability classification (X=stable CV<25%, Y=moderate, Z=erratic) combined with ABC into ABC-XYZ matrix
- **Supplier Scorecard**: Composite 0–100 weighted score across On-Time Delivery (30%), Quality Acceptance (30%), Lead Time Consistency (20%), Defect Rate (20%)
- **Supplier Tiering**: Strategic / Preferred / Development / At Risk segmentation
- **9-Panel Dashboard**: Stock status, ABC Pareto, supplier heatmap, EOQ scatter, excess cost by category, tier distribution, ABC-XYZ matrix, lead time histogram, safety stock scatter

## Tech Stack
```
Python 3.x | Pandas | NumPy | Scikit-Learn | SciPy | Matplotlib | Seaborn
```

## Project Structure
```
project_3_supply_chain/
├── supply_chain_analytics.py         # Complete pipeline
├── data/
│   ├── products_master.csv           # 500 SKUs with all calculated parameters
│   ├── demand_history.csv            # 24-month demand history
│   └── supplier_data.csv             # 25 supplier records
├── outputs/
│   ├── plots/
│   │   └── 01_supply_chain_dashboard.png
│   └── reports/
│       ├── inventory_optimization.csv  # EOQ, ROP, safety stock per SKU
│       └── supplier_scorecard.csv      # Ranked supplier performance
└── README.md
```

## How to Run

### Prerequisites
```bash
pip install pandas numpy scikit-learn scipy matplotlib seaborn
```

### Execute
```bash
cd project_3_supply_chain
python supply_chain_analytics.py
```

## Key Results
| Metric | Value |
|---|---|
| SKUs Analyzed | 500 |
| Critical Understock SKUs | ~50–80 |
| Overstocked SKUs | ~100–150 |
| Annual Excess Holding Cost | $500K–$2M+ |
| Potential Savings (60% reduction) | $300K–$1.2M |
| At-Risk Suppliers | ~3–5 |

- **Top Insight**: A items (typically ~20% of SKUs) drive 80% of spend — these receive tightest inventory controls and safety stock
- **Supplier Action**: At-Risk suppliers flagged for immediate review; Strategic suppliers identified for partnership development

## GitHub Topics
`supply-chain` `inventory-optimization` `eoq` `safety-stock` `abc-analysis` `supplier-management` `operations-analytics` `python` `data-analysis`
