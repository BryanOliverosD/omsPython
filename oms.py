import funciones

def CallOMS(nameShipping,namePropuesta):
		####################### MAIN ###########################
		#diccionario para guardar objetos.
		almacenador = {}
		#diccionario para ordenar objetos y dejar similar hoja detalle
		detalle = {}
		#diccionario en donde se calculan los escenarios, tarifa sugerida y nueva tarifa. similar hoja análisis
		analisis = {}
		#print(y,x,a,b,c)
		fileNameShipping = nameShipping#"input/shipping_Falabella.xls"
		fileNamePropuesta = namePropuesta#"input/propuesta2.xlsm"
		funciones.leerParametros()
		#parametros basicos
		y = funciones.y
		x = funciones.x
		a = funciones.a
		b = funciones.b
		c = funciones.c
		d = funciones.d
		funciones.copiarHoja(fileNameShipping,fileNamePropuesta)
		#funciones.ValidarParametros(fileNamePropuesta)
		almacenador = funciones.almacenarDatos(fileNamePropuesta)
		#funciones.ActualizarDetalle(fileNamePropuesta,almacenador)
		detalle = funciones.reordenarDiccionario(almacenador)
		# actualizamos MT;BT;SBT
		analisis = funciones.generarAnalisis(detalle)
		analisis = funciones.aproximarValores(analisis)
		analisis = funciones.calcularNuevaTarifa(analisis)
		analisis = funciones.restriccionesTickets(analisis)
		funciones.generarResumen(analisis)
