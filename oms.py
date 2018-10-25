import funciones
####################### MAIN ###########################
#diccionario para guardar objetos.
almacenador = {}
#diccionario para ordenar objetos y dejar similar hoja detalle
detalle = {}
#diccionario en donde se calculan los escenarios, tarifa sugerida y nueva tarifa. similar hoja análisis
analisis = {}
#print(y,x,a,b,c)
file_name_shipping = "input/shipping_Falabella.xls"
file_name_propuesta = "input/propuesta2.xlsm"
funciones.LeerParametros()
#parametros basicos
y = funciones.y
x = funciones.x
a = funciones.a
b = funciones.b
c = funciones.c
funciones.CopiarHojaDetalle(file_name_propuesta,file_name_shipping)
#funciones.ValidarParametros(file_name_propuesta)
almacenador = funciones.AlmacenarDatos(file_name_propuesta)
#funciones.ActualizarDetalle(file_name_propuesta,almacenador)
detalle = funciones.ReordenarDiccionario(almacenador)
# actualizamos MT;BT;SBT
analisis = funciones.GenerarAnalisis(detalle)
analisis = funciones.calcularNuevaTarifa(analisis)
analisis = funciones.RestriccionesTickets(analisis)
funciones.GenerarResumen(analisis)