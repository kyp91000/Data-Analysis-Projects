"""
==========================================================================
 SUPPLY CHAIN ANALYTICS — INVENTORY OPTIMIZATION & SUPPLIER PERFORMANCE
 Author  : Senior Data Analyst
 Stack   : Python · Pandas · NumPy · Scikit-Learn · Matplotlib · Seaborn
==========================================================================
 Business Problem:
   Poor inventory management costs companies 25-30% of total inventory
   value annually through overstock, stockouts, and spoilage. This
   project builds a data-driven supply chain analytics system that
   calculates optimal reorder points, safety stock levels, supplier
   scorecards, and identifies $2M+ in cost-saving opportunities.
==========================================================================
"""

import os, warnings
warnings.filterwarnings('ignore')

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import seaborn as sns
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
import pickle

os.makedirs('data', exist_ok=True)
os.makedirs('outputs/plots', exist_ok=True)
os.makedirs('outputs/reports', exist_ok=True)

print("=" * 65)
print("  SUPPLY CHAIN ANALYTICS PIPELINE")
print("=" * 65)

# ──────────────────────────────────────────────────────────────
# 1. DATA GENERATION
# ──────────────────────────────────────────────────────────────
print("\n[1/6] Generating supply chain dataset …")

np.random.seed(42)
n_products  = 500
n_suppliers = 25
n_months    = 24  # 2 years

categories = ['Raw Materials', 'Components', 'Packaging', 'MRO', 'Finished Goods']
suppliers  = [f'SUP_{str(i).zfill(3)}' for i in range(1, n_suppliers + 1)]

# Products
product_ids      = [f'SKU-{str(i).zfill(4)}' for i in range(1, n_products + 1)]
product_cats     = np.random.choice(categories, n_products, p=[0.20, 0.30, 0.15, 0.15, 0.20])
product_supplier = np.random.choice(suppliers, n_products)
unit_cost        = np.random.lognormal(mean=3.5, sigma=1.2, size=n_products).round(2)
holding_rate     = np.random.uniform(0.18, 0.28, n_products)   # 18-28% of unit cost / year
order_cost       = np.random.uniform(50, 500, n_products).round(2)
lead_time_mean   = np.random.uniform(5, 45, n_products).round(1)
lead_time_std    = lead_time_mean * np.random.uniform(0.10, 0.30, n_products)
service_level_target = np.random.choice([0.90, 0.95, 0.99], n_products, p=[0.30, 0.50, 0.20])

# Demand history (monthly)
demand_records = []
for i, pid in enumerate(product_ids):
    base_demand = np.random.lognormal(mean=4, sigma=1.0)
    for m in range(n_months):
        seasonal = 1 + 0.20 * np.sin(2 * np.pi * m / 12 + np.random.uniform(0, np.pi))
        trend    = 1 + 0.02 * m / 12
        demand   = max(0, np.random.poisson(base_demand * seasonal * trend))
        demand_records.append({'product_id': pid, 'month': m + 1, 'demand': demand})

df_demand = pd.DataFrame(demand_records)

# Aggregate demand stats per product
demand_stats = df_demand.groupby('product_id').agg(
    avg_monthly_demand=('demand', 'mean'),
    std_monthly_demand=('demand', 'std'),
    total_demand=('demand', 'sum'),
    max_monthly_demand=('demand', 'max'),
    min_monthly_demand=('demand', 'min'),
).reset_index()
demand_stats['demand_cv'] = demand_stats['std_monthly_demand'] / demand_stats['avg_monthly_demand'].clip(lower=0.01)

# Supplier performance data
supplier_records = []
for sup in suppliers:
    on_time_rate   = np.random.beta(8, 2)       # ~80% avg
    quality_rate   = np.random.beta(9, 1)       # ~90% avg
    lead_time_avg  = np.random.uniform(7, 40)
    lead_time_var  = np.random.uniform(0.5, 8)
    defect_rate    = np.random.uniform(0.005, 0.08)
    resp_score     = np.random.uniform(60, 100)
    n_products_sup = (product_supplier == sup).sum()
    supplier_records.append({
        'supplier_id': sup, 'on_time_delivery_rate': on_time_rate,
        'quality_acceptance_rate': quality_rate, 'avg_lead_time_days': lead_time_avg,
        'lead_time_variability': lead_time_var, 'defect_rate': defect_rate,
        'responsiveness_score': resp_score, 'n_products_supplied': n_products_sup,
    })
df_suppliers = pd.DataFrame(supplier_records)

# Products master table
df_products = pd.DataFrame({
    'product_id': product_ids, 'category': product_cats,
    'supplier_id': product_supplier, 'unit_cost': unit_cost,
    'holding_rate_annual': holding_rate, 'order_cost': order_cost,
    'lead_time_mean_days': lead_time_mean, 'lead_time_std_days': lead_time_std,
    'service_level_target': service_level_target,
})
df_products = df_products.merge(demand_stats, on='product_id')
df_products = df_products.merge(df_suppliers[['supplier_id','on_time_delivery_rate',
                                               'lead_time_variability']], on='supplier_id')

# Convert demand to daily
df_products['avg_daily_demand'] = df_products['avg_monthly_demand'] / 30
df_products['std_daily_demand'] = df_products['std_monthly_demand'] / np.sqrt(30)

print(f"      Products: {n_products} | Suppliers: {n_suppliers} | Months: {n_months}")

# ──────────────────────────────────────────────────────────────
# 2. INVENTORY OPTIMIZATION (EOQ + SAFETY STOCK)
# ──────────────────────────────────────────────────────────────
print("\n[2/6] Computing EOQ, Safety Stock, and Reorder Points …")

from scipy.stats import norm

# Economic Order Quantity (EOQ)
annual_demand = df_products['avg_daily_demand'] * 365
holding_cost  = df_products['unit_cost'] * df_products['holding_rate_annual']
df_products['eoq'] = np.sqrt(
    (2 * annual_demand * df_products['order_cost']) / holding_cost
).round(0)
df_products['eoq'] = df_products['eoq'].clip(lower=1)

# Safety Stock = Z * sqrt(lead_time * sigma_demand^2 + demand^2 * sigma_lead^2)
df_products['z_score'] = norm.ppf(df_products['service_level_target'])
demand_variance    = df_products['std_daily_demand'] ** 2
lead_time_variance = (df_products['lead_time_variability'] / np.sqrt(30)) ** 2

df_products['safety_stock'] = (
    df_products['z_score'] *
    np.sqrt(df_products['lead_time_mean_days'] * demand_variance +
            (df_products['avg_daily_demand'] ** 2) * lead_time_variance)
).round(0)
df_products['safety_stock'] = df_products['safety_stock'].clip(lower=0)

# Reorder Point = (avg daily demand * lead time) + safety stock
df_products['reorder_point'] = (
    df_products['avg_daily_demand'] * df_products['lead_time_mean_days']
    + df_products['safety_stock']
).round(0)

# Current stock simulation (some understocked, some overstocked)
df_products['current_stock'] = (
    df_products['reorder_point'] * np.random.uniform(0.3, 3.5, n_products)
).round(0)

df_products['stock_status'] = pd.cut(
    df_products['current_stock'] / df_products['reorder_point'],
    bins=[0, 0.5, 0.9, 1.5, 999],
    labels=['Critical Understock', 'Low Stock', 'Optimal', 'Overstock']
)

# Holding cost for overstock
df_products['overstock_qty'] = (df_products['current_stock'] - df_products['eoq']).clip(lower=0)
df_products['annual_holding_cost'] = df_products['current_stock'] * holding_cost
df_products['excess_holding_cost'] = df_products['overstock_qty'] * holding_cost

# ──────────────────────────────────────────────────────────────
# 3. ABC ANALYSIS
# ──────────────────────────────────────────────────────────────
print("\n[3/6] Performing ABC-XYZ Classification …")

df_products['annual_spend'] = df_products['avg_daily_demand'] * 365 * df_products['unit_cost']
df_products = df_products.sort_values('annual_spend', ascending=False).reset_index(drop=True)
df_products['cumulative_spend_pct'] = df_products['annual_spend'].cumsum() / df_products['annual_spend'].sum()

df_products['abc_class'] = np.where(df_products['cumulative_spend_pct'] <= 0.80, 'A',
                            np.where(df_products['cumulative_spend_pct'] <= 0.95, 'B', 'C'))

# XYZ (demand variability)
df_products['xyz_class'] = np.where(df_products['demand_cv'] < 0.25, 'X',
                            np.where(df_products['demand_cv'] < 0.50, 'Y', 'Z'))

df_products['abc_xyz'] = df_products['abc_class'] + df_products['xyz_class']

abc_summary = df_products.groupby('abc_class').agg(
    n_products=('product_id', 'count'),
    total_annual_spend=('annual_spend', 'sum'),
    avg_unit_cost=('unit_cost', 'mean'),
).round(2)
print(f"\n  ABC Summary:\n{abc_summary.to_string()}\n")

# ──────────────────────────────────────────────────────────────
# 4. SUPPLIER SCORECARD
# ──────────────────────────────────────────────────────────────
print("[4/6] Generating Supplier Scorecards …")

df_suppliers['delivery_score']  = df_suppliers['on_time_delivery_rate'] * 100
df_suppliers['quality_score']   = df_suppliers['quality_acceptance_rate'] * 100
df_suppliers['lead_time_score'] = (1 - df_suppliers['lead_time_variability'] /
                                   df_suppliers['lead_time_variability'].max()) * 100
df_suppliers['defect_score']    = (1 - df_suppliers['defect_rate'] / df_suppliers['defect_rate'].max()) * 100

# Weighted composite score
df_suppliers['composite_score'] = (
    0.30 * df_suppliers['delivery_score']
    + 0.30 * df_suppliers['quality_score']
    + 0.20 * df_suppliers['lead_time_score']
    + 0.20 * df_suppliers['defect_score']
)

df_suppliers['supplier_tier'] = pd.cut(df_suppliers['composite_score'],
    bins=[0, 60, 75, 88, 100], labels=['At Risk', 'Development', 'Preferred', 'Strategic'])

# ──────────────────────────────────────────────────────────────
# 5. VISUALIZATIONS
# ──────────────────────────────────────────────────────────────
print("[5/6] Generating supply chain analytics dashboard …")

fig = plt.figure(figsize=(22, 16))
gs  = gridspec.GridSpec(3, 3, figure=fig, hspace=0.40, wspace=0.35)
fig.suptitle('Supply Chain Analytics Dashboard', fontsize=16, fontweight='bold')

# Plot 1: Stock Status Distribution
ax1 = fig.add_subplot(gs[0, 0])
status_counts = df_products['stock_status'].value_counts()
colors_status = {'Critical Understock': '#c0392b', 'Low Stock': '#f39c12',
                 'Optimal': '#27ae60', 'Overstock': '#2980b9'}
bars = ax1.bar(status_counts.index, status_counts.values,
               color=[colors_status[k] for k in status_counts.index], edgecolor='black', linewidth=0.5)
ax1.set_title('Inventory Status Distribution', fontweight='bold')
ax1.set_ylabel('Number of SKUs')
ax1.tick_params(axis='x', rotation=20)
for bar, v in zip(bars, status_counts.values):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
             str(v), ha='center', fontsize=10, fontweight='bold')

# Plot 2: ABC Analysis - Pareto
ax2 = fig.add_subplot(gs[0, 1])
abc_data = df_products.groupby('abc_class')['annual_spend'].sum()
colors_abc = ['#e74c3c', '#f39c12', '#3498db']
ax2.bar(abc_data.index, abc_data.values/1e6, color=colors_abc, edgecolor='black', linewidth=0.5)
ax2_twin = ax2.twinx()
abc_pct = (abc_data / abc_data.sum() * 100).cumsum()
ax2_twin.plot(abc_data.index, abc_pct.values, 'k-o', linewidth=2)
ax2.set_title('ABC Analysis — Pareto', fontweight='bold')
ax2.set_ylabel('Annual Spend ($M)')
ax2_twin.set_ylabel('Cumulative % Spend')
ax2_twin.set_ylim(0, 110)

# Plot 3: Supplier Scorecard Heatmap
ax3 = fig.add_subplot(gs[0, 2])
top_sups = df_suppliers.nlargest(12, 'n_products_supplied')[
    ['supplier_id', 'delivery_score', 'quality_score', 'lead_time_score', 'defect_score']
].set_index('supplier_id')
sns.heatmap(top_sups, ax=ax3, cmap='RdYlGn', vmin=50, vmax=100, annot=True, fmt='.0f',
            cbar_kws={'label': 'Score'}, linewidths=0.5, annot_kws={'size': 7})
ax3.set_title('Supplier Scorecard (Top 12)', fontweight='bold')
ax3.tick_params(axis='x', rotation=20)
ax3.tick_params(axis='y', rotation=0)

# Plot 4: EOQ vs Current Stock scatter
ax4 = fig.add_subplot(gs[1, 0])
color_map = {'Critical Understock': '#c0392b', 'Low Stock': '#f39c12',
             'Optimal': '#27ae60', 'Overstock': '#2980b9'}
for status, grp in df_products.groupby('stock_status'):
    ax4.scatter(grp['eoq'], grp['current_stock'],
                alpha=0.5, s=30, color=color_map[status], label=status)
ax4.plot([0, df_products['eoq'].max()], [0, df_products['eoq'].max()],
         'k--', alpha=0.5, label='EOQ = Current Stock')
ax4.set_title('Current Stock vs EOQ', fontweight='bold')
ax4.set_xlabel('Economic Order Quantity (EOQ)')
ax4.set_ylabel('Current Stock Level')
ax4.legend(fontsize=7)
ax4.set_xlim(left=0)
ax4.set_ylim(bottom=0)

# Plot 5: Excess Holding Cost by Category
ax5 = fig.add_subplot(gs[1, 1])
excess_by_cat = df_products.groupby('category')['excess_holding_cost'].sum().sort_values(ascending=True)
colors_cat = ['#c0392b' if v > 100000 else '#f39c12' if v > 50000 else '#3498db'
              for v in excess_by_cat.values]
excess_by_cat.plot(kind='barh', ax=ax5, color=colors_cat, edgecolor='black', linewidth=0.5)
ax5.set_title('Excess Holding Cost by Category', fontweight='bold')
ax5.set_xlabel('Annual Excess Cost ($)')
ax5.grid(True, axis='x', alpha=0.3)

# Plot 6: Supplier Tier Distribution
ax6 = fig.add_subplot(gs[1, 2])
tier_counts = df_suppliers['supplier_tier'].value_counts()
colors_tier = ['#27ae60', '#3498db', '#f39c12', '#e74c3c']
ax6.pie(tier_counts.values, labels=tier_counts.index, colors=colors_tier,
        autopct='%1.0f%%', startangle=90)
ax6.set_title('Supplier Tier Distribution', fontweight='bold')

# Plot 7: ABC-XYZ Matrix heatmap
ax7 = fig.add_subplot(gs[2, 0])
xyz_matrix = df_products.groupby(['abc_class', 'xyz_class'])['product_id'].count().unstack(fill_value=0)
sns.heatmap(xyz_matrix, ax=ax7, cmap='Blues', annot=True, fmt='d',
            linewidths=0.5, cbar_kws={'label': 'SKU Count'})
ax7.set_title('ABC-XYZ Classification Matrix', fontweight='bold')

# Plot 8: Lead Time Distribution by Category
ax8 = fig.add_subplot(gs[2, 1])
for i, cat in enumerate(categories):
    subset = df_products[df_products['category'] == cat]['lead_time_mean_days']
    ax8.hist(subset, bins=20, alpha=0.5, label=cat)
ax8.set_title('Lead Time Distribution by Category', fontweight='bold')
ax8.set_xlabel('Lead Time (Days)')
ax8.set_ylabel('Count')
ax8.legend(fontsize=7)

# Plot 9: Safety Stock vs Lead Time
ax9 = fig.add_subplot(gs[2, 2])
scatter = ax9.scatter(df_products['lead_time_mean_days'], df_products['safety_stock'],
                      c=df_products['demand_cv'], cmap='RdYlGn_r', alpha=0.6, s=25)
plt.colorbar(scatter, ax=ax9, label='Demand CV (Variability)')
ax9.set_title('Safety Stock vs Lead Time\n(colored by demand variability)', fontweight='bold')
ax9.set_xlabel('Lead Time (Days)')
ax9.set_ylabel('Safety Stock (Units)')

plt.savefig('outputs/plots/01_supply_chain_dashboard.png', dpi=150, bbox_inches='tight')
plt.close()

# ──────────────────────────────────────────────────────────────
# 6. SAVE REPORTS
# ──────────────────────────────────────────────────────────────
print("[6/6] Saving reports …")

df_products.to_csv('outputs/reports/inventory_optimization.csv', index=False)
df_suppliers.sort_values('composite_score', ascending=False).to_csv(
    'outputs/reports/supplier_scorecard.csv', index=False)
df_demand.to_csv('data/demand_history.csv', index=False)
df_suppliers.to_csv('data/supplier_data.csv', index=False)
df_products.to_csv('data/products_master.csv', index=False)

# Executive Summary
critical = (df_products['stock_status'] == 'Critical Understock').sum()
overstock = (df_products['stock_status'] == 'Overstock').sum()
total_excess_cost = df_products['excess_holding_cost'].sum()
at_risk_sups = (df_suppliers['supplier_tier'] == 'At Risk').sum()

print(f"""
  ┌──────────────────────────────────────────────────┐
  │         SUPPLY CHAIN EXECUTIVE SUMMARY           │
  ├──────────────────────────────────────────────────┤
  │  SKUs Analyzed         : {n_products:,}               │
  │  Critical Understock   : {critical} SKUs              │
  │  Overstocked SKUs      : {overstock}                │
  │  Annual Excess Hold Cost: ${total_excess_cost:,.0f}   │
  │  At-Risk Suppliers     : {at_risk_sups}                  │
  │  Potential Savings     : ${total_excess_cost * 0.60:,.0f}   │
  └──────────────────────────────────────────────────┘

  Output files:
  ├── data/products_master.csv
  ├── data/demand_history.csv
  ├── data/supplier_data.csv
  ├── outputs/plots/01_supply_chain_dashboard.png
  ├── outputs/reports/inventory_optimization.csv
  └── outputs/reports/supplier_scorecard.csv
""")
