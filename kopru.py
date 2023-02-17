import xlsxwriter
import pylightxl as xl

db = xl.readxl(fn='datasheet.xlsx')

eslesmis_row = db.ws(ws="Eşleşmiş Liste").maxrow

result = []

for i in range(1, eslesmis_row + 1):
    result.append(db.ws(ws="Eşleşmiş Liste").index(row=i, col=1))
    
print(result)