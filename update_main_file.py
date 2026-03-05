import os
import shutil

# 刪除舊檔案
try:
    os.remove('HMA_開發人天評估表.xlsx')
    print('✓ 舊檔案已刪除')
except:
    print('舊檔案不存在或正在使用中')

# 複製新檔案
shutil.copy('HMA_開發人天評估表_含維護.xlsx', 'HMA_開發人天評估表.xlsx')
print('✓ 已更新主檔案: HMA_開發人天評估表.xlsx')
print('\n現在有兩個檔案:')
print('  1. HMA_開發人天評估表.xlsx (主檔案,含保固維護)')
print('  2. HMA_開發人天評估表_含維護.xlsx (備份)')
