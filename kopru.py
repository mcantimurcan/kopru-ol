import xlsxwriter
import pylightxl as xl

db = xl.readxl(fn='datasheet.xlsx')

gonder_row = db.ws(ws="Kitap Gönder").maxrow
al_row = db.ws(ws="Kitap Al").maxrow

result = []
gonder_liste = []
alici_liste = []

for i in range(2, gonder_row + 1):
    gonder_liste.append(db.ws(ws="Kitap Gönder").index(row=i, col=3))
    
for i in range(2, al_row + 1):
    alici_liste.append(db.ws(ws="Kitap Al").index(row=i, col=3))

for i in range(2, al_row + 1):
    alici_il = db.ws(ws="Kitap Al").index(row=i, col=7)
    for j in range(2, gonder_row + 1):
        gonderen_il = db.ws(ws="Kitap Gönder").index(row=j, col=6)
        if alici_il == gonderen_il:
            result.append([db.ws(ws="Kitap Al").index(row=i, col=3), db.ws(ws="Kitap Gönder").index(row=j, col=3)])
    
    
print(result)