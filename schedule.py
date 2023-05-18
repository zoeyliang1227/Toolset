import openpyxl

wb = openpyxl.load_workbook('專櫃值班表112年3月.xlsx' , data_only=True)     # 開啟 Excel 檔案, 設定 data_only=True 只讀取計算後的數值

def get_values():
    for sheet in wb.worksheets:
        print(sheet)
        sh = sheet.max_row
        for column in range(1, sheet.max_column+1):
            
            data=[sheet.cell(row=i,column=column).value for i in range(1, sheet.max_row+1)]
            count = 0
            for i in data:
                if i == '全':
                    count += 1
            
            sheet.cell(row=sh, column=column).value = count
        wb.save('專櫃值班表112年3月.xlsx')

get_values()