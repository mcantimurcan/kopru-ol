import xlsxwriter
import pylightxl as xl

db = xl.readxl(fn='datasheet.xlsx')

gonder_row = db.ws(ws="Kitap Gönder").maxrow
gonder_col = db.ws(ws="Kitap Gönder").maxcol
al_row = db.ws(ws="Kitap Al").maxrow
al_col = db.ws(ws="Kitap Al").maxcol

result = []
eslesmeyen = []
gonderen_liste = []
alici_liste = []

#gönderen listesi oluşturma
for i in range(2, gonder_row + 1):
    new = []
    for j in range(1, gonder_col + 1):
        new.append(db.ws(ws="Kitap Gönder").index(row=i, col=j))
    gonderen_liste.append(new)
    
#alıcı listesi oluşturma   
for i in range(2, al_row + 1):
    new = []
    for j in range(1, al_col + 1):
        new.append(db.ws(ws="Kitap Al").index(row=i, col=j))
    alici_liste.append(new)

#teyitsiz olanları silme
for i in alici_liste:
    if i[::-2] == "Doğru Değil":
        alici_liste.remove(i)

#eşleştirme
m = 0
while m < len(gonderen_liste):
    gonderen_il = gonderen_liste[m][5].lower()
    for i in alici_liste:
        alici_il = i[6].lower()
        if alici_il == gonderen_il:
            result.append([gonderen_liste[m][2], i[2]])
            alici_liste.remove(i)
            break
    else:
        eslesmeyen.append(gonderen_liste[m])
    gonderen_liste.remove(gonderen_liste[m])

   
