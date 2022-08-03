from openpyxl import load_workbook

# read_only=True 無法進行 tuple(ws.columns) 操作，要跑column列表必須 read_only=False
wb = load_workbook(filename='readtest.xlsx', read_only=False)  # 輸入檔案名稱
ws = wb['sheet1']  # 輸入分頁名稱

ROW = tuple(ws.rows)  # 依序列出同列(row)表格資料
row_designate = ROW[8]
for row_result in row_designate:
    print(row_result.value)

print("Row finish. Show column.")

COL = tuple(ws.columns)  # 依序列出同欄(column)表格資料
column_designate = COL[1]
for column_result in column_designate:
    print(column_result.value)

wb.close()

print("All finish!")
