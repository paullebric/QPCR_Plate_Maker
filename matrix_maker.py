import numpy as np
import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font
from datetime import datetime

path="C:\\Users\\paulo\\OneDrive\\Bureau\\"
def main_input(char):
    échantillon = []
    while True :
        échantillon += [input(char)]
        print(échantillon[-1])
        if échantillon[-1] == "":
            échantillon.pop(-1)
            break
    return échantillon

def apply_method(S,T,mode):
    fill=[]
    for s in S :
        for t in T :
            for i in range(mode):
                fill +=[s+" "+t]
    return fill

def simple_plate(plate):
    for x in range(9) :
        for y in range(13):
            left = plate[x][0]
            top = plate[0][y]
            if left and top !="": plate[x][y] = left + " " + top
    return plate

def ecrire_matrice_excel(matrice,Sample,Target):
    couleurs_fond = [
    "F0F8FF",  # Alice Blue
    "E6E6FA",  # Lavender
    "F5F5DC",  # Beige
    "FFFACD",  # Lemon Chiffon
    "FDF5E6",  # Old Lace
    "FAEBD7",  # Antique White
    "F0FFF0",  # Honeydew
    "F5FFFA",  # Mint Cream
    "F0FFFF",  # Azure
    "F5F5F5",  # White Smoke
    "FFFFF0",  # Ivory
    "F0F0F0"   # Gainsboro
]

    couleurs_texte = [
    "8B0000",  # Dark Red
    "2F4F4F",  # Dark Slate Gray
    "3A3A3A",  # Dark Gray
    "4B0082",  # Indigo
    "5F9EA0",  # Cadet Blue
    "556B2F",  # Dark Olive Green
    "708090",  # Slate Gray
    "2E8B57",  # Sea Green
    "6A5ACD",  # Slate Blue
    "8B4513",  # Saddle Brown
    "800080",  # Purple
    "483D8B"   # Dark Slate Blue
]

    nom_fichier = "QPCR_plate_"+str(datetime.now().date())+".xlsx"
    wb = openpyxl.Workbook()
    sheet = wb.active

    max_length =0
    for x in range(np.shape(matrice)[0]):
            for y in range(np.shape(matrice)[1]):
                if len(matrice[x][y]) > max_length:
                    max_length = len(matrice[x][y])
    
    for x in range(42):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(x+1)].width = max_length
        
    for i in range(matrice.shape[0]):  # Pour chaque ligne
        for j in range(matrice.shape[1]):  # Pour chaque colonne
            sheet.cell(row=i+1, column=j+1, value=matrice[i, j])  # Insérer la valeur dans Excel

    for x in range(len(Sample)) :
        couleur = couleurs_fond[x % len(couleurs_fond)]  # Assurer que les couleurs sont cyclées si plus d'échantillons que de couleurs
        fill = PatternFill(start_color=couleur, end_color=couleur, fill_type="solid")
        formula = f'ISNUMBER(SEARCH("{Sample[x]}", A1))'
        rule = FormulaRule(formula=[formula], fill=fill)
        sheet.conditional_formatting.add('A1:BZ10', rule)

    for x in range(len(Target)) :
        couleur_texte = couleurs_texte[x % len(couleurs_texte)]  # Assurer que les couleurs sont cyclées si plus d'échantillons que de couleurs
        font = Font(color=couleur_texte)
        formula = f'ISNUMBER(SEARCH("{Target[x]}", A1))'
        rule = FormulaRule(formula=[formula], font=font)
        sheet.conditional_formatting.add('A1:BZ10', rule)
    wb.save(path+nom_fichier)
    print(f"Fichier {nom_fichier} créé avec succès!")

def plate_matrixer(S,T,mode):
    plate = np.zeros((9,13),dtype='U60')
    if len(S)<=8 and len(T)*mode<=12:
        left = S ;top = T
    elif len(T)<8 and len(S)*mode<=12:
        left = T; top = S
    else :
        return False
    for x in range(len(left)) :
        plate[x+1][0] = left[x]
    ntop=[]
    for x in top :
        for y in range(mode) : ntop += [x]
    for y in range(len(ntop)) : plate[0][y+1] = ntop[y]

    return simple_plate(plate)

def main():
    Sample= main_input("Noms des écchantillons/controles :\nSi finit just clic enter :\n")
    Target= main_input("Noms des ammorces/targets :\nSi finit just clic enter :\n")
    mode = int(input("Simplicat [1], Duplicat [2], Triplicat [3]"))
    Fillers = apply_method(Sample,Target,mode)
    plate = plate_matrixer(Sample,Target,mode)
    if plate != False:
        ecrire_matrice_excel(plate,Sample,Target)

main()
#plate_matrixer(["S12Allprep","S14Allprep","S12RNeasy","S14RNeasy"],["423","669","16-1","103","16-5","223"],3)


    