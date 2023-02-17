import xlsxwriter
import pylightxl as xl

db = xl.readxl(fn='datasheet.xlsx')

gonder_row = db.ws(ws="Kitap Gönder").maxrow
gonder_col = db.ws(ws="Kitap Gönder").maxcol
al_row = db.ws(ws="Kitap Al").maxrow

result = []
gonder_liste = []
alici_liste = []

for i in range(2, gonder_row + 1):
    new = []
    for j in range(1, gonder_col + 1):
        new.append(db.ws(ws="Kitap Gönder").index(row=i, col=j))
    gonder_liste.append(new)
    
for i in range(2, al_row + 1):
    new = []
    for j in range(1, gonder_col + 1):
        new.append(db.ws(ws="Kitap Al").index(row=i, col=j))
    alici_liste.append(new)

for i in gonder_liste:
    gonderen_il = i[6]
    for j in alici_liste:
        alici_il = j[5]
        if alici_il == gonderen_il:
            result.append([i[2], j[2]])
            alici_liste.remove(j)
            gonder_liste.remove(i)
    
    
print(result)