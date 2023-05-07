import pandas as pd
import csv

庫存報表 = pd.read_excel('csfp422-230505-094209(已排).xls')  # '研騰件號\nPSI Part No.'
銷貨報表 = pd.read_excel('庫存列印 0505.xls', header=1)  # '料件編號'
輸出用檔案 = pd.DataFrame()
x = 0
# 建立一個新的excel檔案


for index1, value1 in 銷貨報表.iterrows():
    for index2, value2 in 庫存報表.iterrows():
        # 若 'MIS Ship Remark' 銷貨紀錄 和 'GDS' 庫存皆為空
        # 則跳過此資料
        if pd.isna(value2['MIS Ship Remark']) and pd.isna(value2['GDS']):
            continue
        elif value1['料件編號'] == value2['研騰件號\nPSI Part No.'] \
                and not pd.isna(value2['MIS Ship Remark']) \
                and not pd.isna(value2['GDS']):
            輸出用檔案.at[x, ['料件編號']] = value1['料件編號']
            輸出用檔案.at[x, ['銷貨紀錄']] = value2['MIS Ship Remark']
            輸出用檔案.at[x, ['GDS']] = value2['GDS']
            x += 1
        elif value1['料件編號'] == value2['研騰件號\nPSI Part No.'] \
                and not pd.isna(value2['MIS Ship Remark']) \
                and pd.isna(value2['GDS']):
            輸出用檔案.at[x, ['料件編號']] = value1['料件編號']
            輸出用檔案.at[x, ['銷貨紀錄']] = value2['MIS Ship Remark']
            x += 1
            # print(value1['料件編號'])
            # print()
            # print('銷貨紀錄')
            # print(value2['MIS Ship Remark'])
            # print()
        elif value1['料件編號'] == value2['研騰件號\nPSI Part No.'] \
                and pd.isna(value2['MIS Ship Remark']) \
                and not pd.isna(value2['GDS']):
            輸出用檔案.at[x, ['料件編號']] = value1['料件編號']
            輸出用檔案.at[x, ['GDS']] = value2['GDS']
            x += 1
            # print(value1['料件編號'])
            # print()
            # print('GDS:', value2['GDS'])
            # print()
    # if index1 == 5:
    #     break

writer = pd.ExcelWriter('銷貨核對.xlsx', engine='xlsxwriter')
輸出用檔案.to_excel(writer, index=False, sheet_name='銷貨核對')

worksheet = writer.sheets['銷貨核對']

銷貨紀錄格式 = writer.book.add_format({'font_size': 11,'text_wrap': True})
料件編號格式 = writer.book.add_format({'font_size': 11, 'valign': 'vcenter'})
# 欄寬設置
# 料件編號 18
# 銷貨紀錄 75
worksheet.set_column(0, 0, 18, 料件編號格式)
worksheet.set_column(1, 1, 75, 銷貨紀錄格式)
writer.save()
