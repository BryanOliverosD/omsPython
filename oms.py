import funciones

########################################################
####################### MAIN ###########################
#########################################################

almacenador = {}
detalle = {}
analisis = {}

file_name_shipping = "input/shipping_Falabella.xls"
file_name_propuesta = "input/propuesta2.xlsm"
#CopiarHojaDetalle(file_name_propuesta,file_name_shipping)
#ValidarParametrosBase(file_name_propuesta)
almacenador = funciones.AlmacenarDatos(file_name_propuesta)
#ActualizarDetalle(file_name_propuesta,almacenador)
detalle = funciones.ReordenarDiccionario(almacenador)
# actualizamos MT;BT;SBT
analisis = funciones.GenerarAnalisis(detalle)
analisis = funciones.calcularNuevaTarifa(analisis)
analisis = funciones.RestriccionesTickets(analisis)
funciones.GenerarResumen(analisis)