#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Final summary report with correct category counts"""

import openpyxl
from datetime import datetime

wb = openpyxl.load_workbook('HMA_開發人天評估表.xlsx')
ws = wb['開發人天評估']

print('\n' + '='*90)
print('📊 HMA 醫務管理系統 - 第二階段開發人天評估報告')
print('='*90)
print(f'📅 評估日期: {datetime.now().strftime("%Y-%m-%d")}')
print('='*90 + '\n')

# Collect category data properly
categories = []
current_category = None
features_in_category = []

for row in ws.iter_rows(min_row=2, values_only=True):
    if not row[1]:
        continue
    
    cell_value = str(row[1])
    
    # Category header (merged cell, no complexity)
    if row[4] is None and '小計' not in cell_value and '總計' not in cell_value and '其他工作項目' != cell_value:
        # Save previous category if exists
        if current_category and features_in_category:
            pass  # Will be saved when we hit the subtotal
        
        current_category = cell_value
        features_in_category = []
    
    # Regular feature row (has complexity and index)
    elif row[0] and row[4] is not None and row[0] != '序號':
        if current_category:
            features_in_category.append({
                'index': row[0],
                'name': row[1],
                'code': row[2],
                'complexity': row[4],
                'days': row[5],
                'cost': row[6]
            })
    
    # Subtotal row
    elif '小計' in cell_value and '總計' not in cell_value:
        cat_name = cell_value.replace(' 小計', '')
        
        if '其他工作項目' in cat_name:
            categories.append({
                'name': '其他工作項目',
                'count': 6,
                'days': row[5],
                'cost': row[6],
                'is_additional': True
            })
        else:
            categories.append({
                'name': cat_name,
                'count': len(features_in_category),
                'days': row[5],
                'cost': row[6],
                'is_additional': False
            })
        
        features_in_category = []
        current_category = None

# Print header
print('📋 功能分類明細:\n')
print(f"{'分類':<25} {'功能數':<10} {'人天':<12} {'費用 (NT$)':<20} {'占比':<10}")
print('-'*90)

# Calculate totals
total_features = sum(cat['count'] for cat in categories if not cat['is_additional'])
total_dev_days = sum(cat['days'] for cat in categories if not cat['is_additional'])
total_dev_cost = sum(cat['cost'] for cat in categories if not cat['is_additional'])

# Sort categories by man-days (descending)
feature_cats = [cat for cat in categories if not cat['is_additional']]
feature_cats.sort(key=lambda x: x['days'], reverse=True)

# Print feature categories
for cat in feature_cats:
    percentage = (cat['days'] / total_dev_days * 100) if total_dev_days > 0 else 0
    print(f"{cat['name']:<25} {cat['count']:<10} {cat['days']:<12} {cat['cost']:>18,}  {percentage:>6.1f}%")

print('-'*90)
print(f"{'功能開發小計':<25} {total_features:<10} {total_dev_days:<12} {total_dev_cost:>18,}  100.0%")
print()

# Additional items
for cat in categories:
    if cat['is_additional']:
        print(f"{cat['name']:<25} {cat['count']:<10} {cat['days']:<12} {cat['cost']:>18,}")

print()
print('='*90)

# Grand total
grand_total_days = total_dev_days + 54
grand_total_cost = total_dev_cost + 518400

print(f"{'總計 (GRAND TOTAL)':<25} {total_features+6:<10} {grand_total_days:<12} {grand_total_cost:>18,}")
print('='*90)

# Breakdown by complexity
print('\n📊 複雜度分布:\n')
print(f"{'複雜度':<15} {'人天單價':<12} {'說明':<40}")
print('-'*90)
print(f"{'簡單':<15} {'2-3 天':<12} {'查詢、列印等基本功能':<40}")
print(f"{'中等':<15} {'5 天':<12} {'一般作業、申請功能':<40}")
print(f"{'複雜':<15} {'10 天':<12} {'批次處理、整合、匯入/上傳功能':<40}")
print(f"{'非常複雜':<15} {'15 天':<12} {'大規模系統整合、複雜工作流程':<40}")

# Project timeline estimates
print('\n⏱️  預估工期:\n')
print(f"  • 以 1 人開發: {grand_total_days:>6} 天 (約 {grand_total_days/20:.1f} 個月)")
print(f"  • 以 2 人開發: {grand_total_days/2:>6.1f} 天 (約 {grand_total_days/40:.1f} 個月)")
print(f"  • 以 3 人開發: {grand_total_days/3:>6.1f} 天 (約 {grand_total_days/60:.1f} 個月)")
print(f"  • 以 5 人開發: {grand_total_days/5:>6.1f} 天 (約 {grand_total_days/100:.1f} 個月)")
print(f"  • 以 8 人開發: {grand_total_days/8:>6.1f} 天 (約 {grand_total_days/160:.1f} 個月)")

print('\n💡 建議開發團隊規模: 5-8 人 (工期約 5.6-7.0 個月)')

# Cost breakdown
print('\n💰 成本明細:\n')
print(f"  • 每日工時: 8 小時")
print(f"  • 時薪: NT$ 1,200")
print(f"  • 日薪: NT$ 9,600")
print(f"  • 月薪(20工作天): NT$ 192,000")
print()
print(f"  • 功能開發費用: NT$ {total_dev_cost:,}")
print(f"  • 其他工作項目費用: NT$ 518,400")
print(f"  • 總開發費用: NT$ {grand_total_cost:,}")

print('\n' + '='*90)
print('✅ 完整評估報告: HMA_開發人天評估表.xlsx')
print('   工作表1: 開發人天評估 - 229個功能詳細清單')
print('   工作表2: 專案摘要 - 總體評估數據')
print('='*90 + '\n')

print('📄 說明文件: HMA_評估說明.md')
print()
