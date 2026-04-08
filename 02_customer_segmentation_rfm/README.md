# 👥 Customer Segmentation — RFM Analysis

## Overview
This project segments customers using **RFM (Recency, Frequency, Monetary)** analysis
combined with **K-Means clustering**. It identifies high-value customers, at-risk customers,
and churned customers — enabling targeted marketing strategies.

## Features
- Synthetic transaction dataset: 500 customers, 3,000 transactions
- RFM metric computation per customer
- Quintile-based RFM scoring (1–5 scale)
- Rule-based segment labeling (Champions, Loyal, At Risk, Lost, etc.)
- K-Means clustering (4 clusters) for data-driven grouping
- PCA 2D scatter plot for segment visualization
- Pie chart + bar chart for segment distribution & spend analysis
- Fully styled Excel report with 3 sheets

## Technologies Used
| Tool | Purpose |
|------|---------|
| Python 3 | Core scripting |
| Pandas | Data processing & aggregation |
| NumPy | Random data generation |
| Scikit-learn | KMeans clustering, PCA, StandardScaler |
| Matplotlib | Charts (pie, bar, scatter) |
| OpenPyXL | Excel report export |

## Project Structure
```
02_customer_segmentation_rfm/
│
├── customer_segmentation.py          ← Main script
├── Customer_Segmentation_RFM.xlsx    ← Excel report (after running)
├── rfm_segment_charts.png            ← Pie + bar charts
├── rfm_pca_scatter.png               ← PCA cluster scatter
└── README.md
```

## How to Run

### Prerequisites
```bash
pip install pandas numpy scikit-learn matplotlib openpyxl
```

### Run
```bash
cd 02_customer_segmentation_rfm
python customer_segmentation.py
```

## RFM Segments Explained
| Segment | Meaning |
|---------|---------|
| Champions | Bought recently, buy often, spend the most |
| Loyal Customers | Consistent buyers with good spend |
| Potential Loyalists | Recent customers with some frequency |
| New Customers | Bought recently, very few purchases |
| Promising | Recent shoppers but not frequent |
| At Risk | Used to buy frequently but haven't recently |
| Lost Customers | Low recency, low frequency, low spend |

## Output Files
| File | Description |
|------|-------------|
| `Customer_Segmentation_RFM.xlsx` | 3-sheet report: RFM data, summary, transactions |
| `rfm_segment_charts.png` | Segment pie chart + avg spend bar chart |
| `rfm_pca_scatter.png` | PCA scatter showing customer clusters |
