#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Quick summary of the estimation report"""

import openpyxl

wb = openpyxl.load_workbook('HMA_開發人天評估表.xlsx')
ws = wb['開發人天評估']

print('\n' + '='*70)
print('📊 HMA 醫務管理系統 - 第二階段開發人天評估')
print('='*70 + '\n')

categories = {}
feature_count = 0
in_additional = False
current_category = ''  # Initialize current_category

for row in ws.iter_rows(min_row=2, values_only=True):
    if not row[1]:  # Empty row
        continue
    
    cell_value = str(row[1])
    
    # Check if this is a category header or subtotal
    if '小計' in cell_value:
        if '其他工作項目' in cell_value:
            in_additional = True
            print(f"\n{'='*70}")
            print('📋 其他工作項目')
            print(f"{'='*70}")
            print(f"  人天: {row[5]}")
            print(f"  費用: NT$ {row[6]:,}")
        elif '總計' not in cell_value:
            cat_name = cell_value.replace(' 小計', '')
            if cat_name in categories:
                categories[cat_name]['subtotal_days'] = row[5]
                categories[cat_name]['subtotal_cost'] = row[6]
                print(f"  --> 小計: {row[5]} 人天, NT$ {row[6]:,}\n")
    elif cell_value == '總計 (GRAND TOTAL)':
        print(f"\n{'='*70}")
        print('💰 總計 (GRAND TOTAL)')
        print(f"{'='*70}")
        print(f"  總人天: {row[5]}")
        print(f"  總費用: NT$ {row[6]:,}")
        print(f"{'='*70}\n")
    elif not any(x in cell_value for x in ['小計', '總計', '其他工作項目']) and row[0] != '序號':
        # This is either a category header or a feature
        if row[4] is None or row[4] == '':  # Category header (no complexity)
            current_category = cell_value
            if current_category not in categories:
                categories[current_category] = {
                    'features': [],
                    'subtotal_days': 0,
                    'subtotal_cost': 0
                }
            print(f"\n【{current_category}】")
            print('-'*70)
        else:
            # Regular feature
            feature_count += 1
            if current_category in categories:
                categories[current_category]['features'].append({
                    'index': row[0],
                    'name': row[1],
                    'code': row[2],
                    'complexity': row[4],
                    'days': row[5],
                    'cost': row[6]
                })

# Print summary
print('\n' + '='*70)
print('📈 分類摘要')
print('='*70)
for cat_name, cat_data in categories.items():
    if cat_name and cat_data['subtotal_days']:
        print(f"\n{cat_name}:")
        print(f"  功能數量: {len(cat_data['features'])}")
        print(f"  人天: {cat_data['subtotal_days']}")
        print(f"  費用: NT$ {cat_data['subtotal_cost']:,}")

print(f'\n總功能數: {feature_count}')
print('='*70)
