# 👥 HR Attrition Analysis Dashboard

## Overview
This project analyzes employee attrition data to identify key drivers of turnover,
compute retention metrics by department/age/satisfaction, and build a Random Forest
model to rank the most important factors influencing employee exits.

## Features
- Synthetic HR dataset: 800 employees with 15+ attributes
- Overall attrition & retention rate calculation
- Attrition breakdown: by Department, Age Group, Job Satisfaction
- Random Forest classifier to identify top attrition drivers
- 4-panel dashboard chart (PNG)
- Fully styled Excel report with 4 sheets

## Technologies Used
| Tool | Purpose |
|------|---------|
| Python 3 | Core scripting |
| Pandas | Data processing |
| NumPy | Data generation |
| Scikit-learn | Random Forest, Label Encoding |
| Matplotlib | 4-panel dashboard chart |
| OpenPyXL | Excel report export |

## Project Structure
```
05_hr_attrition_analysis/
│
├── hr_attrition.py                ← Main script
├── HR_Attrition_Analysis.xlsx     ← Excel report (after running)
├── hr_attrition_dashboard.png     ← 4-panel dashboard chart
└── README.md
```

## How to Run

### Prerequisites
```bash
pip install pandas numpy scikit-learn matplotlib openpyxl
```

### Run
```bash
cd 05_hr_attrition_analysis
python hr_attrition.py
```

## Output Files
| File | Description |
|------|-------------|
| `HR_Attrition_Analysis.xlsx` | 4-sheet Excel report |
| `hr_attrition_dashboard.png` | Charts: Dept, Age, Satisfaction, Feature Importance |

## Key Metrics Computed
- Overall Attrition Rate & Retention Rate
- Average Tenure of Employees Who Left
- Department-wise Attrition Rate
- Attrition vs Job Satisfaction (1–4 scale)
- Attrition vs Age Group
- Top 8 Attrition Drivers (Random Forest Feature Importance)
- Model Accuracy Score

## Power BI Integration
1. Import `HR_Attrition_Analysis.xlsx` into Power BI
2. Use "Key Drivers" sheet for AI-powered visual
3. Enable Power BI AI Insights → Key Influencers visual
4. Set target field to "Attrition" column
