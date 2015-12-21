#https://automatetheboringstuff.com/chapter12/
#http://scrolltest.com/python/working-with-excel-files-in-python/

import openpyxl

wb = openpyxl.load_workbook('workbookname.xlsx')
#print wb.get_sheet_names()



sheet = wb.get_sheet_by_name('Sheet2')


# for i in range(0, sheet.get_highest_row()):
#     a=(i, sheet.cell(row=i, column=1).value,)
#     print a

# 
#rowOfCellObjects=tuple(sheet['A1':'C3'])
# #print rowOfCellObjects 
# for rowOfCellObjects in sheet['A1':'C3']:
#     for cellObj in rowOfCellObjects:
#         print(cellObj, cellObj.value)
#     print('--- END OF ROW ---')
    
a={}    
i =1
for rowOfCellObj in sheet['A1':'L3']:
    a[i]={}
    for cellObj in rowOfCellObj:
        #print(cellObj, cellObj.value)
        stra = str(cellObj).split('.')
        a[i][stra[1].replace(">", "")]=cellObj.value
        #print a
        #print('--- ROW END ---')
    i=i+1
    

print a[1] #row1
print a[2] #row1
print a[3] #row






