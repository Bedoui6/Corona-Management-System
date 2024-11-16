from openpyxl import Workbook
import openpyxl
personnes=openpyxl.Workbook()
Maladies=openpyxl.Workbook()


sheet1["A1"].value="CIN"
sheet1["B1"].value="NOM"
sheet1["C1"].value="PRENOM"
sheet1["D1"].value="TEL"
sheet1["E1"].value="NATIONALITE"
sheet1["F1"].value="ADRESSE"
sheet1["G1"].value="AGE"
sheet1["H1"].value="JOUR"
sheet1["I1"].value="MOIS"
sheet1["J1"].value="ANNEE"
sheet1["K1"].value="DECEDE"
personnes.save("personnes.xlsx")



sheet2["A1"].value="CODE"
sheet2["B1"].value="CIN"
sheet2["C1"].value="MALADIE"
sheet2["D1"].value="NOMBREANNEE"
Maladies.save("Maladies.xlsx")