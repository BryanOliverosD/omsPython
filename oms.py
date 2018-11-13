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
    file_name_shipping = nameShipping#"input/shipping_Falabella.xls"
    file_name_propuesta = namePropuesta#"input/propuesta2.xlsm"
    funciones.leerParametros()
    #parametros basicos
    y = funciones.y
    x = funciones.x
    a = funciones.a
    b = funciones.b
    c = funciones.c
    funciones.copiarHojaDetalle(file_name_propuesta,file_name_shipping)
    #funciones.ValidarParametros(file_name_propuesta)
    almacenador = funciones.almacenarDatos(file_name_propuesta)
    #funciones.ActualizarDetalle(file_name_propuesta,almacenador)
    detalle = funciones.reordenarDiccionario(almacenador)
    # actualizamos MT;BT;SBT
    analisis = funciones.generarAnalisis(detalle)
    analisis = funciones.aproximarValores(analisis)
    analisis = funciones.calcularNuevaTarifa(analisis)
    analisis = funciones.restriccionesTickets(analisis)
    funciones.generarResumen(analisis)
