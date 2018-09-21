import pandas as pd
from openpyxl import load_workbook

class ShippingFalabella():

	nombreComuna = ""

	precio_MT = ""
	precio_BT = ""
	precio_SBT = ""
	dias_MT = ""
	dias_BT = ""
	dias_SBT = ""

class ShippingRipley():

	nombreComuna =""

	precio_MT = ""
	precio_BT = ""
	precio_SBT = ""
	dias_MT = ""
	dias_BT = ""
	dias_SBT = ""

class ShippingParis():

	nombreComuna = ""

	precio_MT = ""
	precio_BT = ""
	precio_SBT = ""
	dias_MT = ""
	dias_BT = ""
	dias_SBT = ""


##### Se copia información desde archivo shipping a propuesta #####
def CopiarHojaDetalle(name_propuesta,name_shipping):

	sheet_shipping = "Hoja 1"

	archivopropuesta = pd.ExcelFile(name_propuesta)

	excelfile_shipping = pd.read_excel(skiprows=0,io=name_shipping, sheet_name=sheet_shipping, usecols='A,C:I', skipcols="tamaño")

	# Create a Pandas dataframe from the data.
	df = pd.DataFrame(excelfile_shipping)

	# Se crea un nuevo archivo en el cual se pegará lo tomado desde la hoja detalles y posteriormente guardamos.
	book = load_workbook(name_propuesta)
	writer = pd.ExcelWriter(name_propuesta, engine='openpyxl')
	writer.book = book
	writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

	df.to_excel(writer, sheet_name='Shipping.csv', index=False, startcol=1)
	writer.save() 

## Validamos parámetros hoja base archivo propuesta actualizado ##
def ValidarParametrosBase(name_propuesta):

	sheet_base = "Base"

	excelfile = pd.read_excel(header=None,skiprows=1,io=name_propuesta, sheet_name=sheet_base, usecols='B',nrows=5)
	#posiciones variables a validar
	df = pd.DataFrame(excelfile)
	y = str(df.at[0,0])
	x = str(df.at[1,0])
	a = str(df.at[2,0])
	b = str(df.at[3,0])
	c = str(df.at[4,0])

	book = load_workbook(name_propuesta)
	#Utilizamos para poder reemplazar un valor, reescribimos todo de la misma manera salvo la variable en cuestion
	writer = pd.ExcelWriter(name_propuesta, engine='openpyxl')
	writer.book = book
	writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

	if (y != "0.1"):
		df.at[0,0] = 0.1
	if (x != "0.0"):
		df.at[1,0] = 0
	if (a != "500.0"):
		df.at[2,0] = 500
	if (b != "750.0"):
		df.at[3,0] = 750
	if (c != "1000.0"):
		df.at[4,0] = 1000
	
	df.to_excel(writer, sheet_name='Base', index=False, startcol=1)
	writer.save()

####### Almacenar datos ##############
def AlmacenarDatos(name_propuesta):
	#creamos diccionario
	almacenador = {}
	
	excelfile = pd.read_excel(header=None,skiprows=1,io=name_propuesta, sheet_name="Shipping.csv", usecols='D,F,G,H,J')
	df = pd.DataFrame(excelfile)

	for fila in range(0,len(df)):

		#lista almacenador de objetos
		objetos = []

		producto = str(df.iloc[fila,1])

		if producto[0].lower() == "f":

			shippingF = ShippingFalabella()
			shippingF.nombreComuna = str(df.iloc[fila,0])

			if df.iloc[fila,4] == "MT":
				shippingF.precio_MT = df.iloc[fila,2]
				shippingF.dias_MT = df.iloc[fila,3]
			elif df.iloc[fila,4] == "BT":
				shippingF.precio_BT = df.iloc[fila,2]
				shippingF.dias_BT = df.iloc[fila,3]
			elif df.iloc[fila,4] == "SBT":
				shippingF.precio_SBT = df.iloc[fila,2]
				shippingF.dias_SBT = df.iloc[fila,3]

			objetos.append(shippingF)

		elif producto[0].lower() == "r":

			shippingR = ShippingRipley()
			shippingR.nombreComuna = str(df.iloc[fila,0])

			if df.iloc[fila,4] == "MT":
				shippingR.precio_MT = df.iloc[fila,2]
				shippingR.dias_MT = df.iloc[fila,3]
			elif df.iloc[fila,4] == "BT":
				shippingR.precio_BT = df.iloc[fila,2]
				shippingR.dias_BT = df.iloc[fila,3]
			elif df.iloc[fila,4] == "SBT":
				shippingR.precio_SBT = df.iloc[fila,2]
				shippingR.dias_SBT = df.iloc[fila,3]

			objetos.append(shippingR)

		elif producto[0].lower() == "p":

			shippingP = ShippingParis()

			shippingP.nombreComuna = str(df.iloc[fila,0])

			if df.iloc[fila,4] == "MT":
				shippingP.precio_MT = df.iloc[fila,2]
				shippingP.dias_MT = df.iloc[fila,3]
			elif df.iloc[fila,4] == "BT":
				shippingP.precio_BT = df.iloc[fila,2]
				shippingP.dias_BT = df.iloc[fila,3]
			elif df.iloc[fila,4] == "SBT":
				shippingP.precio_SBT = df.iloc[fila,2]
				shippingP.dias_SBT = df.iloc[fila,3]

			objetos.append(shippingP)
		
		# Validamos si la ciudad se encuentra en el diccionario, si está se agrega a la lista que arrastra, si no se agrega al diccionario.		
		if str(df.iloc[fila,0]) not in almacenador:
			almacenador[str(df.iloc[fila,0])] = objetos
		else:
			almacenador[str(df.iloc[fila,0])] = objetos + almacenador[str(df.iloc[fila,0])]

	for dia in almacenador:
		for i in almacenador[dia]:
			if i.precio_MT != "" and i.dias_MT != "":
				print(i.nombreComuna,i.precio_MT,i.dias_MT)


######## Almacenar comunas #############
def AlmacenarComunas(name_propuesta):

	comunas = []

	excelfile = pd.read_excel(header=None,skiprows=1,io=name_propuesta, sheet_name="Shipping.csv", usecols='D,F,G,H,J')
	df = pd.DataFrame(excelfile)

	excelfile_detalle = pd.read_excel(header=None,skiprows=6,io=name_propuesta, sheet_name="Detalle", usecols='A:AA')
	df1 = pd.DataFrame(excelfile_detalle)

	print(df1.iloc[0][0])

	for fila in range(0,len(df1)):

		nombreComuna = str(df1.iloc[fila,0])
		print(nombreComuna)


	"""for fila in range(0,len(df)):

		nombreComuna = str(df.iloc[fila,0])

		producto = str(df.iloc[fila,1])
		precio = str(df.iloc[fila,2])
		dias = str(df.iloc[fila,3])
		ticket = str(df.iloc[fila,4])

		if nombreComuna == "ANTOFAGASTA" and producto[0] == 'F' and ticket == "MT":
			print(precio,dias)
		#if str(df.iloc[fila,0]) not in comunas:
			comunas.append(nombreComuna)
			print(producto, producto[0])"""		


##################### MAIN ###################

file_name_shipping = "shipping_Falabella.xls"
file_name_propuesta = "propuesta2.xlsx"
#CopiarHojaDetalle(file_name_propuesta,file_name_shipping)
#ValidarParametrosBase(file_name_propuesta)
#AlmacenarComunas(file_name_propuesta)
AlmacenarDatos(file_name_propuesta)