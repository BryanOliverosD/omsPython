###################### leer hoja por hoja #######################

#excelfile1_propuesta = pd.read_excel(archivopropuesta,'Instrucciones')
#excelfile2_propuesta = pd.read_excel(archivopropuesta,'Detalle')
#excelfile3_propuesta = pd.read_excel(archivopropuesta,'Base')
#excelfile4_propuesta = pd.read_excel(archivopropuesta,'Análisis')
#excelfile5_propuesta = pd.read_excel(archivopropuesta,'Resumen')
#excelfile6_propuesta = pd.read_excel(archivopropuesta,'4 bandas')
# Create a Pandas dataframe from the data.
#df1 = pd.DataFrame(excelfile1_propuesta)
#df2 = pd.DataFrame(excelfile2_propuesta)
#df3 = pd.DataFrame(excelfile3propuesta)
#df4 = pd.DataFrame(excelfile4_propuesta)
#df5 = pd.DataFrame(excelfile5_propuesta)
#df6 = pd.DataFrame(excelfile6_propuesta)

################ ESCRIBIR EN EXCEL POR HOJA ######################
#df2.to_excel(newexcel, sheet_name='Detalle')
#df3.to_excel(newexcel, sheet_name='Base')
#df4.to_excel(newexcel, sheet_name='Análisis')
#df5.to_excel(newexcel, sheet_name='Resumen')
#df6.to_excel(newexcel, sheet_name='4 bandas')

###################### RECORRER OBJETOS ########################

"""class Propuesta:
	comuna = ""
	tipo = ""
	CostoFalabella_MT=""
	DiasFalabella_MT = ""
	CostoRipley_MT = ""
	DiasRipley_MT = ""
	CostoParis_MT = ""
	DiasParis_MT =""
"""
"""
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
print(len(excelfile.index))"""