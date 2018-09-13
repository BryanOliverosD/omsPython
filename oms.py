import pandas as pd

class Propuesta:
	comuna = ""
	tipo = ""
	CostoFalabella_MT=""
	DiasFalabella_MT = ""
	CostoRipley_MT = ""
	DiasRipley_MT = ""
	CostoParis_MT = ""
	DiasParis_MT =""

lista_propuestas = []

file_name = "propuesta.xlsm" # path to file + file name
sheet = "Detalle" # sheet name or sheet number or list of sheet numbers and names
excelfile = pd.read_excel(header=None, skiprows=3,io=file_name, sheet_name=sheet, usecols=18)
print(excelfile.head(20))  # print first 5 rows of the datafram

for i in range(3,96):

	propuesta = Propuesta()
	propuesta.comuna = excelfile[0][i]
	propuesta.CostoFalabella_MT = excelfile[1][i]
	propuesta.DiasFalabella_MT = excelfile[2][i]

	propuesta.CostoRipley_MT = excelfile[3][i]
	propuesta.DiasRipley_MT = excelfile[4][i]

	propuesta.CostoParis_MT = excelfile[5][i]
	propuesta.DiasParis_MT = excelfile[6][i]

	lista_propuestas.append(propuesta)

for a in lista_propuestas:
	print(a.comuna, a.CostoFalabella_MT, a.DiasFalabella_MT, a.CostoRipley_MT, 
			a.DiasRipley_MT, a.CostoParis_MT, a.DiasParis_MT)
print(len(excelfile.index))
