#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Enhanced summary of the estimation report"""

import openpyxl
from datetime import datetime

wb = openpyxl.load_workbook('HMA_開發人天評估表.xlsx')
ws = wb['開發人天評估']

print('\n' + '='*80)
print('📊 HMA 醫務管理系統 - 第二階段開發人天評估報告')
print('='*80)
print(f'📅 評估日期: {datetime.now().strftime("%Y-%m-%d")}')
print('='*80 + '\n')

# Collect all category subtotals
categories = []
feature_counts = {}
current_category = ''
feature_count_in_category = 0

for row in ws.iter_rows(min_row=2, values_only=True):
    if not row[1]:
        continue
    
    cell_value = str(row[1])
    
    # Category header (no complexity value)
    if row[4] is None and '小計' not in cell_value and '總計' not in cell_value and '其他工作項目' not in cell_value:
        if feature_count_in_category > 0 and current_category:
            feature_counts[current_category] = feature_count_in_category
        current_category = cell_value
        feature_count_in_category = 0
    # Regular feature (has complexity)
    elif row[4] is not None and row[0] != '序號':
        feature_count_in_category += 1
    # Subtotal row
    elif '小計' in cell_value and '總計' not in cell_value:
        cat_name = cell_value.replace(' 小計', '')
        if '其他工作項目' not in cat_name:
            categories.append({
                'name': cat_name,
                'count': feature_count_in_category,
                'days': row[5],
                'cost': row[6]
            })
            feature_counts[cat_name] = feature_count_in_category
        else:
            # Additional items
            categories.append({
                'name': '其他工作項目 (6項)',
                'count': 6,
                'days': row[5],
                'cost': row[6]
            })

# Print category breakdown
print('📋 功能分類明細:\n')
print(f"{'分類':<20} {'功能數':<10} {'人天':<12} {'費用 (NT$)':<20}")
print('-'*80)

total_features = 0
total_dev_days = 0
total_dev_cost = 0

for cat in categories:
    if '其他工作項目' not in cat['name']:
        print(f"{cat['name']:<20} {cat['count']:<10} {cat['days']:<12} {cat['cost']:>18,}")
        total_features += cat['count']
        total_dev_days += cat['days']
        total_dev_cost += cat['cost']

print('-'*80)
print(f"{'功能開發小計':<20} {total_features:<10} {total_dev_days:<12} {total_dev_cost:>18,}")
print()

# Additional items
for cat in categories:
    if '其他工作項目' in cat['name']:
        print(f"{cat['name']:<20} {cat['count']:<10} {cat['days']:<12} {cat['cost']:>18,}")
        print()

# Grand total
grand_total_days = total_dev_days + 54
grand_total_cost = total_dev_cost + 518400

print('='*80)
print(f"{'總計 (GRAND TOTAL)':<20} {total_features+6:<10} {grand_total_days:<12} {grand_total_cost:>18,}")
print('='*80)

# Project timeline estimates
print('\n⏱️  預估工期:\n')
print(f"  • 以 1 人開發: {grand_total_days} 天 (約 {grand_total_days/20:.1f} 個月)")
print(f"  • 以 2 人開發: {grand_total_days/2:.1f} 天 (約 {grand_total_days/40:.1f} 個月)")
print(f"  • 以 3 人開發: {grand_total_days/3:.1f} 天 (約 {grand_total_days/60:.1f} 個月)")
print(f"  • 以 5 人開發: {grand_total_days/5:.1f} 天 (約 {grand_total_days/100:.1f} 個月)")

print('\n' + '='*80)
print('✅ 評估報告已生成: HMA_開發人天評估表.xlsx')
print('   包含兩個工作表:')
print('   1. 開發人天評估 - 詳細功能清單及估算')
print('   2. 專案摘要 - 總體評估數據')
print('='*80 + '\n')
