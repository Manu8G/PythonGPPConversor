"""
Buenas si estás leyendo esto significa que te ha tocado cambiar algo de limesurvey y que por consiguiente vas a tener
que cambiar este código, por lo que te lo voy a intentar comentar lo mejor posible... Antes de nada te recomiendo
descargar Pycharm, puesto que es increíblemente util hazme caso.
Para que este código funcione es necesario realizar antes los siguientes comandos:
    pip install docxtpl
    pip install docx
    pip install openpyxl
"""

import os
import shutil
from docxtpl import DocxTemplate
import packaging
import packaging.version
import packaging.specifiers
import packaging.requirements

from Extractor import extraer_datos
from traductor_de_campos import traducir
from datetime import date

# Rutas
OUTPUT_PATH = 'Outputs'
CSV_FILE = 'AACSV/results-survey346241.csv'
PIP_PATH = 'Templates/PLANTILLA FINAL.docx'
SESIONES_PATH = 'Templates/Plantilla de sesiones.docx'
DIRECTORY = os.getcwd()

'''
El comentario de debajo se refiere a que en el test SDQ hay dos formas de actuar tomando los valores de 
"Verdaderamente sí" como 3 hasta "No es verdad" como 0, pero hay preguntas donde estos valores se cambian 
completamente, pasando a ser "Verdaderamente sí" como 0 hasta "No es verdad" como 3. Esto pasa en el resto 
de matrices pero como sus valores ya son mas variables se ha especificado de otra forma, este al ser siempre 
el mismo tipo de respuesta se ha dejado asi.  
'''
# Preguntas cuyo valor no es positivo en la pregunta de la matriz del SDQ
ESPECIALES = [7, 11, 14, 21, 25]
# Categorías matriz 1 (SQD)
EMOCIONALES = [3, 8, 13, 16, 24]
CONDUCTA = [5, 7, 12, 18, 22]
HIPERACTIVIDAD = [2, 10, 15, 21, 25]
PROBLEMAS = [6, 11, 14, 19, 23]
PROSOCIAL = [1, 4, 9, 17, 20]
# Categorías matriz 2 (EFECO)
FLEXIBILIDAD = [4, 23, 27, 32, 49, 59]
PLANIFICACION = [22, 28, 39, 44, 58, 61, 62]
CONTROL = [7, 19, 48, 50, 55, 63, 67]
ORGANIZACION = [1, 9, 10, 26, 30, 45, 51, 60]
MONITORIZACION = [2, 6, 11, 12, 25, 29, 31, 35, 43]
INHIBICION = [3, 14, 15, 18, 21, 33, 34, 37, 42, 46]
INICIATIVA = [8, 17, 20, 36, 40, 47, 53, 56, 64, 65]
MEMORIA = [5, 13, 16, 24, 38, 41, 52, 54, 57, 66]
# Categorias matriz 3 (HADs)
NEGATIVOS = [1, 3, 5, 6, 11, 13]
ANSIEDAD = [0, 2, 4, 6, 8, 10, 12]
DEPRESION = [1, 3, 5, 7, 9, 11, 13]


def eliminar_y_crear_carpeta(path):
    # Eliminamos la carpeta y su contenido
    if os.path.exists(path):
        shutil.rmtree(path)

    # Creamos la carpeta
    os.mkdir(path)


def por_defecto(campos, aux, i, j, diccionario):
    igual = True
    encontrado = False
    while igual and encontrado != True:  # Sea el próximo campo diferente al campo en el que estamos
        aux2 = campos[j + 1][0:(campos[j].find("_["))] # Guardamos el siguiente campo
        if aux == aux2: # Comprobamos si el primer campo y el aux2 son iguales
            if i[j] != "No" and i[j] != "N/A" and i[j] != "":
                # Si el contenido es distinto de estos significa que esta respondido
                if campos[j][(campos[j].find("_[")):(len(campos[j + 1]))] == "_[Otro]":
                    # Si el campo rellenado es el campo de "Otro", guardamos su solucion
                    guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], i[j])

                else:
                    # Si es un campo distinto de "Otro" guardamos el campo
                    guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))],
                                           campos[j][(campos[j].find("_[") + 2):(len(campos[j]) - 1)])
                encontrado = True   # Como es solucion unica rompemos el bucle
            j += 1
        else:  # Si dejan de esr iguales los campos rompemos el bucle
            igual = False
    # Nos queda comprobar el ultimo elemento, asi que lo hacemos
    if i[j] != "No" and i[j] != "N/A" and i[j] != "" and encontrado != True:
        # Si el contenido es distinto de estos significa que esta respondido
        if campos[j][(campos[j].find("_[")):(len(campos[j + 1]))] == "_[Otro]":
            # Si el campo rellenado es el campo de "Otro", guardamos su solucion
            guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], i[j])

        else:
            # Si es un campo distinto de "Otro" guardamos el campo
            guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))],
                                   campos[j][(campos[j].find("_[") + 2):(len(campos[j]) - 1)])
    # Devolvemos la j para no quedarnos estancados en la misma pregunta, ni repetir iteraciones innecesarias del bucle
    return j


def funcion_sdq(diccionario, aux, aux2, cont, campos, i, j):
    # En esta funcion vamos a realizar todos los calculos referentes al test SDQ
    sdq = []
    # Primero vamos a guardar todos los valores correspondientes de las respuestas por orden, esto para facilitar
    # el calculo de los parametros del test
    while aux == aux2:
        aux2 = campos[j + 1][0:(campos[j].find("_["))]
        guardar_en_diccionario(diccionario, campos[j][(campos[j].find("_[") + 2):(len(campos[j]) - 1)], i[j])
        if i[j] == "Verdaderamente_sí" and cont not in ESPECIALES:
            sdq.append((cont, 2))
        elif i[j] == "Es_verdad_a_medias" and cont not in ESPECIALES:
            sdq.append((cont, 1))
        elif i[j] == "No_es_verdad" and cont not in ESPECIALES:
            sdq.append((cont, 0))
        elif cont in ESPECIALES:
            if i[j] == "Verdaderamente_sí":
                sdq.append((cont, 0))
            elif i[j] == "Es_verdad_a_medias":
                sdq.append((cont, 1))
            elif i[j] == "No_es_verdad":
                sdq.append((cont, 2))
        cont += 1
        j += 1

    # Ahora pasamos a calcular los diferentes parametros que mide el test SDQ
    val_emo = val_conducta = val_hiper = val_problemas = val_prosocial = 0

    for x, y in sdq:
        if x in EMOCIONALES:
            val_emo += y

        elif x in CONDUCTA:
            val_conducta += y

        elif x in HIPERACTIVIDAD:
            val_hiper += y

        elif x in PROBLEMAS:
            val_problemas += y

        elif x in PROSOCIAL:
            val_prosocial += y

    # Guardamos los valores obtenidos y ponemos marcas de aviso (*) en aquellos que lleguen a valores anormales
    if val_emo >= 7:
        guardar_en_diccionario(diccionario, 'Emocionales', "{} *".format(val_emo))
    else:
        guardar_en_diccionario(diccionario, 'Emocionales', "{}".format(val_emo))

    if val_conducta >= 5:
        guardar_en_diccionario(diccionario, 'conducta', "{} *".format(val_conducta))
    else:
        guardar_en_diccionario(diccionario, 'conducta', "{}".format(val_conducta))

    if val_hiper >= 7:
        guardar_en_diccionario(diccionario, 'hiper', "{} *".format(val_hiper))
    else:
        guardar_en_diccionario(diccionario, 'hiper', "{}".format(val_hiper))

    if val_problemas >= 6:
        guardar_en_diccionario(diccionario, 'problemas', "{} *".format(val_problemas))
    else:
        guardar_en_diccionario(diccionario, 'problemas', "{}".format(val_problemas))

    if 0 <= val_prosocial < 5:
        guardar_en_diccionario(diccionario, 'prosocial', "{} *".format(val_prosocial))
    else:
        guardar_en_diccionario(diccionario, 'prosocial', "{}".format(val_prosocial))
    return cont, j, aux2


def funcion_efeco(diccionario, aux, aux2, cont, campos, i, j):
    # Vamos a realizar los calculos necesarios para el test EFECO
    efeco = []
    # Primero vamos a guardar todos los valores correspondientes de las respuestas por orden, esto para facilitar
    # el calculo de los parametros del test
    while aux == aux2:
        aux2 = campos[j + 1][0:(campos[j].find("_["))]
        guardar_en_diccionario(diccionario, campos[j][(campos[j].find("_[")):(len(campos[j]) - 1)], i[j])
        if i[j] == "Con_mucha_frecuencia":
            efeco.append((cont, 3))
        elif i[j] == "Con_frecuencia":
            efeco.append((cont, 2))
        elif i[j] == "A_veces":
            efeco.append((cont, 1))
        elif i[j] == "Nunca":
            efeco.append((cont, 0))
        cont += 1
        j += 1
    j -= 1
    # Ahora vamos a calcular el total de cada parametro del test
    val_flexi = val_plani = val_control = val_organi = val_monito = val_inhibi = val_inici = val_memoria = 0

    for x, y in efeco:
        if x in FLEXIBILIDAD:
            val_flexi += y

        elif x in PLANIFICACION:
            val_plani += y

        elif x in CONTROL:
            val_control += y

        elif x in ORGANIZACION:
            val_organi += y

        elif x in MONITORIZACION:
            val_monito += y

        elif x in INHIBICION:
            val_inhibi += y

        elif x in INICIATIVA:
            val_inici += y

        elif x in MEMORIA:
            val_memoria += y

    # Por ultimo vamos a obtener el porcentaje que ha obtenido el usuario en cada uno de los parametros
    guardar_en_diccionario(diccionario, 'flexi', "{:.2f}%".format(((val_flexi / 18) * 100)))
    guardar_en_diccionario(diccionario, 'plani', "{:.2f}%".format(((val_plani / 21) * 100)))
    guardar_en_diccionario(diccionario, 'control', "{:.2f}%".format(((val_control / 21) * 100)))
    guardar_en_diccionario(diccionario, 'organi', "{:.2f}%".format(((val_organi / 24) * 100)))
    guardar_en_diccionario(diccionario, 'monito', "{:.2f}%".format(((val_monito / 27) * 100)))
    guardar_en_diccionario(diccionario, 'inhibi', "{:.2f}%".format(((val_inhibi / 30) * 100)))
    guardar_en_diccionario(diccionario, 'inici', "{:.2f}%".format(((val_inici / 30) * 100)))
    guardar_en_diccionario(diccionario, 'memoria', "{:.2f}%".format(((val_memoria / 30) * 100)))

    return cont, j, aux2


def funcion_hads(diccionario, campos, i, j):
    # Vamos a realizar lo necesario para calcular los valores del test HADs
    hads = []
    j += 2
    # Primero voy a guardar todos los valores correspondientes de las respuestas por orden, esto para facilitar
    # el calculo de los parametros del test
    cont = 0
    while cont < 14:
        if cont in NEGATIVOS:
            if i[j] == "Como_siempre":
                hads.append((cont, 0))
            elif i[j] == "No_lo_bastante":
                hads.append((cont, 1))
            elif i[j] == "Sólo_un_poco":
                hads.append((cont, 2))
            elif i[j] == "Nada":
                hads.append((cont, 3))
            elif i[j] == "Al_igual_que_siempre_lo_hice":
                hads.append((cont, 0))
            elif i[j] == "No_tanto_ahora":
                hads.append((cont, 1))
            elif i[j] == "Casi_nunca":
                hads.append((cont, 2))
            elif i[j] == "Nunca":
                hads.append((cont, 3))
            elif i[j] == "Casi_siempre":
                hads.append((cont, 0))
            elif i[j] == "A_veces":
                hads.append((cont, 1))
            elif i[j] == "No_muy_a_menudo":
                hads.append((cont, 2))
            elif i[j] == "Rara_vez":
                hads.append((cont, 3))
            elif i[j] == "Siempre":
                hads.append((cont, 0))
            elif i[j] == "Igual_que_siempre":
                hads.append((cont, 0))
            elif i[j] == "A_menudo":
                hads.append((cont, 0))
            elif i[j] == "Por_lo_general":
                hads.append((cont, 1))
            elif i[j] == "Menos_de_lo_que_acostumbraba":
                hads.append((cont, 1))
            elif i[j] == "A_veces":
                hads.append((cont, 1))
            elif i[j] == "Mucho_menos_de_lo_que_acostumbraba":
                hads.append((cont, 2))

        else:
            if i[j] == "Todos_los_días":
                hads.append((cont, 3))
            elif i[j] == "Muchas_veces":
                hads.append((cont, 2))
            elif i[j] == "A_veces":
                hads.append((cont, 1))
            elif i[j] == "Nunca":
                hads.append((cont, 0))
            elif i[j] == "Definitivamente_y_es_muy_fuerte":
                hads.append((cont, 3))
            elif i[j] == "Sí,_pero_no_es_muy_fuerte":
                hads.append((cont, 2))
            elif i[j] == "Un_poco,_pero_no_me_preocupa":
                hads.append((cont, 1))
            elif i[j] == "Nada":
                hads.append((cont, 0))
            elif i[j] == "La_mayoría_de_las_veces":
                hads.append((cont, 3))
            elif i[j] == "Con_bastante_frecuencia":
                hads.append((cont, 2))
            elif i[j] == "A_veces,_aunque_no_muy_a_menudo":
                hads.append((cont, 1))
            elif i[j] == "Sólo_en_ocasiones":
                hads.append((cont, 0))
            elif i[j] == "Por_lo_general,_en_todo_momento":
                hads.append((cont, 3))
            elif i[j] == "Muy_a_menudo" and campos[j] == "Me_siento_como_si_cada_día_estuviera_más_lento(a)":
                hads.append((cont, 2))
            elif i[j] == "A_veces":
                hads.append((cont, 1))
            elif i[j] == "Me_preocupo_al_igual_que_siempre":
                hads.append((cont, 0))
            elif i[j] == "Rara_vez":
                hads.append((cont, 0))
            elif i[j] == "En_ciertas_ocasiones":
                hads.append((cont, 1))
            elif i[j] == "Podría_tener_un_poco_más_de_cuidado":
                hads.append((cont, 1))
            elif i[j] == "No_mucho":
                hads.append((cont, 1))
            elif i[j] == "No_muy_a_menudo":
                hads.append((cont, 1))
            elif i[j] == "Con_bastante_frecuencia":
                hads.append((cont, 2))
            elif i[j] == "No_me_preocupo_tanto_como_debiera":
                hads.append((cont, 2))
            elif i[j] == "Bastante":
                hads.append((cont, 2))
            elif i[j] == "Bastante_a_menudo":
                hads.append((cont, 2))
            elif i[j] == "Muy_a_menudo" and campos[j] == "Tengo_una_sensación_extraña,_como_si_tuviera_" \
                                                         "mariposas_en_el_estómago":
                hads.append((cont, 3))
            elif i[j] == "Totalmente":
                hads.append((cont, 3))
            elif i[j] == "Mucho":
                hads.append((cont, 3))
            elif i[j] == "Muy_frecuentemente":
                hads.append((cont, 3))
        guardar_en_diccionario(diccionario, campos[j], i[j])
        j += 1
        cont += 1

    # Vamos a calcular el total de los parametros que mide el test
    val_ansiedad = val_depresion = 0

    for x, y in hads:
        if x in ANSIEDAD:
            val_ansiedad += y
        elif x in DEPRESION:
            val_depresion += y

    # Por ultimo vamos a guardar lo valores obtenidos, y si hay algun valor anormal vamos a ponerle un aviso (*)
    if val_ansiedad >= 11:
        guardar_en_diccionario(diccionario, 'Ansiedad', "{}*".format(val_ansiedad))
    else:
        guardar_en_diccionario(diccionario, 'Ansiedad', val_ansiedad)

    if val_depresion >= 11:
        guardar_en_diccionario(diccionario, 'Depresion', "{}*".format(val_depresion))
    else:
        guardar_en_diccionario(diccionario, 'Depresion', val_depresion)

    return j


def matriz(campos, aux, i, j, diccionario):
    # Vamos a realizar los calculos pertinentes a los diferentes test (SDQ, EFECO, HADs...)

    cont = 1
    aux2 = campos[j + 1][0:(campos[j].find("_["))]

    # Apartado del SDQ
    cont, j, aux2 = funcion_sdq(diccionario, aux, aux2, cont, campos, i, j)

    # Apartado del EFECO
    cont = 1
    aux = campos[j][0:(campos[j].find("_["))]
    aux2 = campos[j + 1][0:(campos[j].find("_["))]
    cont, j, aux2 = funcion_efeco(diccionario, aux, aux2, cont, campos, i, j)

    # Apartado del HADs
    j = funcion_hads(diccionario, campos, i, j)

    # Ultimo apartado
    for h in range(18):
        # Si es un valor muy anormal ponemos un aviso importante
        if i[j] == "Necesito_mejorar_en_gran_medida" or i[j] == "Me_preocupa_en_gran_medida":
            guardar_en_diccionario(diccionario, campos[j][(campos[j].find("_[")):(len(campos[j]) - 1)], "**{}"
                                   .format(i[j]))
        elif i[j] == "Necesito_mejorar_bastante" or i[j] == "Me_preocupa_bastante":
            # Si es un valor anormal ponemos un aviso
            guardar_en_diccionario(diccionario, campos[j][(campos[j].find("_[")):(len(campos[j]) - 1)], "*{}"
                                   .format(i[j]))
        else:
            # Si es un valor normal lo guardamos sin mas
            guardar_en_diccionario(diccionario, campos[j][(campos[j].find("_[")):(len(campos[j]) - 1)], i[j])
        j += 1

    return j


def calcular_edad(diccionario):
    # Obtenemos la fecha de nacimiento, solo la fecha, quitamos las horas, minutos y segundos del final
    fecha = diccionario['Fecha_de_nacimiento']
    fecha = fecha[0:fecha.find(" ") + 1]
    # Obtenemos la fecha actual
    actual = date.today()
    # Calculamos en una sola cifra la fecha de nacimiento, es decir transformamos su fecha de nacimiento
    # en una cifra que seran los dias desde el año 0 hasta su nacimiento (si es un poco loco pero asi evitamos)
    # problemas por si hubiera alguna persona muy mayor, ¿porque donde ponemos el limite?
    nacimiento = fecha[0:(fecha.find("_"))]
    dias_nacimiento = 0
    dias_nacimiento += (int(nacimiento[0:(nacimiento.find("-"))])) * 30 * 12
    nacimiento = nacimiento.replace(nacimiento[0:(nacimiento.find("-") + 1)], '')
    dias_nacimiento += ((int(nacimiento[0:(nacimiento.find("-"))])) * 30) - 30
    nacimiento = nacimiento.replace(nacimiento[0:(nacimiento.find("-") + 1)], '')
    dias_nacimiento += int(nacimiento)
    # Hacemos lo mismo de antes pero con la fecha actual
    dias_hoy = 0
    actual = str(actual)
    dias_hoy += (int(actual[0:(actual.find("-"))])) * 30 * 12
    actual = actual.replace(actual[0:(actual.find("-") + 1)], '')
    dias_hoy += ((int(actual[0:(actual.find("-"))])) * 30) - 30
    actual = actual.replace(actual[0:(actual.find("-") + 1)], '')
    dias_hoy += int(actual)
    # Restamos ambas cantidades y lo transformamos en años dividiendolo entre 360
    dias = dias_hoy - dias_nacimiento
    dias = int(dias / 360)
    # Guardamos la edad en el diccionario
    guardar_en_diccionario(diccionario, "Edad", dias)


def calcular_tiempo(diccionario):
    # Vamos a calcular el tiempo que tarda en cada apartado, ya que LimeSurvey lo da de una forma un tanto
    # extraña por eso lo vamos a poner en formato minutos:segundos, no hemos incluido las horas puesto que
    # esperamos que nadie tenga que tardar tanto en rellenar un cuestionario, las personas que mas han tardado
    # con la version larga (todas las matrices activas) lo han realizado en 25 minutos
    lista_campos = ['Tiempo_total', 'Temporización_del_grupo_Consentimiento',
                    'Temporización_del_grupo_Datos_Personales', 'Temporización_del_grupo_Datos_Académicos',
                    'Temporización_del_grupo_Matrices', 'Temporización_del_grupo_Atención_Psicológica_Previa']
    # La lista anterior son los valores que vamos a calcular para meterlos en el diccionario
    for i in lista_campos:
        segundos = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
        tiempo_total = int(int(diccionario[i][0:diccionario[i].find('.')]) / 60)
        tiempo_total_segundos = (int(diccionario[i][0:diccionario[i].find('.')]) % 60)
        # Si sobran de 0 a 9 segundos le ponemos un 0 delante a ese numero
        if tiempo_total_segundos in segundos:
            tiempo_total_segundos = "0{}".format(tiempo_total_segundos)
        resultado_total = "{}:{}".format(str(tiempo_total), str(tiempo_total_segundos))
        guardar_en_diccionario(diccionario, i, resultado_total)


def guardar_en_diccionario(diccionario, campo, valor):
    # Esta funcion se encarga de guardar el valor en el diccionario en su campo correspondiente
    campo = campo.replace(":", "").replace("\"", "").replace("(", "").replace(")", "").replace("/", "").replace("[", "") \
        .replace("]", "").replace("...", "").replace(",", "").replace(".", "")
    # Si el valor a guardar es de tipo int dara un error mas tarde por lo cual le hacemos los cambios necesarios
    var = 12.0
    if type(valor) == type(int(var)):
        pass
    else:
        valor = valor.replace("_", " ").replace("&lt;", ">").replace("00:00:00", "") # .replace(">", "menores")
    diccionario[campo] = valor


def crear_words(datos):
    # Creamos el diccionario que contendra los valores que vamos a meter en el word
    diccionario = {}
    # Como hemos explicado ya, la primera lista de datos contiene los campos de los valores
    campos = datos[0]
    # Eliminamos los campos puesto que ya los tenemos en otra variables, y la ultima linea que no guarda nada
    del datos[0]
    del datos[-1]
    numero_usuario = 0
    for i in datos:
        # Cada "i" es un nuevo usuario, por lo cual con cada iteracion del bucle recorremos un nuevo usuario
        pip = DocxTemplate(PIP_PATH)
        sesiones = DocxTemplate(SESIONES_PATH)
        j = 0
        while j < len(campos):  # Cada "j" es el número de un campo
            if campos[j].find("_[") != -1:
                # Si es un campo compuesto entra aquí, para entender esto mirar un archivo tipo ".xlsx" y ver
                # como las respuestas de selección poseen esta estructura (teniendo en cuenta que los espacios
                # ahora son "_"
                aux = campos[j][0:(campos[j].find("_["))]
                # En las siguientes lineas comprobamos si es alguna de las preguntas especiales, las cuales
                # se debe hacer un procesamiento extra o diferente al comun
                if campos[j][0:(campos[j].find("_["))] == "Eres_de_España":
                    # Esto corresponde al campo de "País de origen" en la "PLANTILLA FINAL.docx" de la carpeta
                    # "Templates", si es español se pondra España si no se pondra el pais del que porviene
                    if i[j] == "Sí":
                        guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], "España")
                    else:
                        guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], i[j + 2])
                    j += 2
                elif campos[j][0:(campos[j].find("_["))] == "Eres_de_Granada":
                    # Si es de España habra tenido que responder si es o no de Granada, en cuyo caso
                    # en la pregunta de Comunidad Autonoma podra de cual es, y si respondio que es de Granada
                    # pondra que es de Andalucia -> Granada
                    if i[j] == "Sí":
                        guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], "Andalucía -> Granada")
                    else:
                        guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], i[j + 2])
                    j += 2
                elif campos[j][0:(campos[j].find("_["))] == "Titulación":
                    # Guarda la titulacion que se esta estudiando, y para no generar duda se pone si es un martes
                    # doctorado (en el caso de grado no se pone nada porque en todos menos en el curso de puentes
                    # pone que es un grado)
                    # Primero vemos cual es el tipo de titulacion
                    if i[j] == "Sí":        # Grado
                        guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], i[j + 5])
                    elif i[j + 1] == "Sí":  # Master
                        guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))],
                                               "Master -> {}".format(i[j + 6]))
                    else:                   # Doctorado
                        guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))],
                                               "Doctorado -> {}".format(i[j + 7]))

                    # Guardamos el tipo de matricula
                    if i[j + 3] == 'Sí':
                        guardar_en_diccionario(diccionario, campos[j + 3][0:(campos[j + 3].find("_["))],
                                               campos[j + 3][(campos[j + 3].find("_[") + 2):len(campos[j + 3]) - 1])
                    else:
                        guardar_en_diccionario(diccionario, campos[j + 4][0:(campos[j + 4].find("_["))],
                                               campos[j + 4][(campos[j + 4].find("_[") + 2):len(campos[j + 4]) - 1])
                    j += 8

                elif campos[j][0:(campos[j].find("_["))] == "De_las_asignaturas_NO_superadas...":
                    # En esta pregunta se necesita el campo de la pregunta y el resultado del comentario
                    # Para mayor comprension busca "De_las_asignaturas_NO_superadas..." en el archivo
                    # de tipo ".xlsx" la estructura de esta pregunta
                    if campos[j][(campos[j].find("_[")):(len(campos[j]) - 1)] == "_[Comentario":
                        pass
                    else:
                        guardar_en_diccionario(diccionario, campos[j], i[j + 1])
                    j += 2

                elif campos[j][0:(campos[j].find("_["))] == "Estas_cursando_la_titulación_que_querías":
                    # Si queria otra titulación tenemos que poner cual era la que queria, si no poner que no
                    if i[j] == "No":
                        if i[j + 2] == "Sí":    # Grado
                            guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], i[j + 5])
                        elif i[j + 3] == "Sí":  # Master
                            guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))],
                                                   "Master -> {}".format(i[j + 6]))
                        else:                   # Doctorado
                            guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))],
                                                   "Doctorado -> {}".format(i[j + 7]))
                    else:
                        guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], i[j])
                    j += 8

                elif campos[j][0:(campos[j].find("_["))] == "Tiene_alguna_otra_titulación":
                    # Si tiene otra titulacion se comrpueba si es universitaria o no, si lo es hacemos
                    # lo mismo de siempre vemos cual es, si es otro tipo la guardamos sin mas

                    solucion = ""
                    cont = j
                    while cont <= j + 4:  # Sea el próximo campo diferente al campo en el que estamos
                        if campos[cont] == "Tiene_alguna_otra_titulación_[Otra_Titulación_Universitaria]" and \
                                i[cont] != "No":
                            if i[cont + 1] == "Sí":  # Grado
                                if solucion != "":
                                    solucion = "{} \n -{}".format(solucion, i[cont + 4])
                                else:
                                    solucion = i[cont + 4]
                            elif i[cont + 2] == "Sí":  # Master
                                if solucion != "":
                                    solucion = "{} \n -{}".format(solucion, "Master -> {}".format(i[cont + 6]))
                                else:
                                    solucion = "Master -> {}".format(i[cont + 6])
                            else:  # Doctorado
                                if solucion != "":
                                    solucion = "{} \n -{}".format(solucion, "Doctorado -> {}".format(i[cont + 5]))
                                else:
                                    solucion = "Doctorado -> {}".format(i[cont + 5])

                        elif i[cont] != "No" and i[cont] != "N/A" and i[cont] != "":
                            if solucion != "":
                                solucion = "{} \n -{}".format(solucion, campos[cont][(campos[cont].find("_[") + 2)
                                                                                     :(len(campos[cont]) - 1)])
                            else:
                                solucion = "-{}".format(campos[cont][(campos[cont].find("_[") + 2)
                                                                     :(len(campos[cont]) - 1)])

                        cont += 1
                    guardar_en_diccionario(diccionario, campos[j][0:campos[j].find("_[")], solucion)
                    j += 11

                elif campos[j][0:(campos[j].find("_["))] == "Has_cambiado_alguna_vez_de_titulación":
                    # Si ha cambiado de titulacion se pone la que tenia antes de cambiar, si no se pone no
                    if campos[j + 1] == "Has_cambiado_alguna_vez_de_titulación_[Si]" and i[j + 1] == "Sí":
                        if i[j + 2] == "Sí":    # Grado
                            guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], i[j + 5])
                        elif i[j + 3] == "Sí":  # Master
                            guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))],
                                                   "Master -> {}".format(i[j + 6]))
                        else:                   # Doctorado
                            guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))],
                                                   "Doctorado -> {}".format(i[j + 7]))
                        j += 8
                    else:
                        guardar_en_diccionario(diccionario, campos[j][0:(campos[j].find("_["))], "No")
                        j += 8

                elif campos[j][0:(campos[j].find("Es_importante"))] == "Por_favor,_marca_con una_cruz_en_el_" \
                                                                       "cuadro_que_crea_que_se_corresponde_a_cada_" \
                                                                       "pregunta:_No_es_verdad,_Es_verdad_a_medias," \
                                                                       "_Verdaderamente_sí._":
                    # Este apartado corresponde con los test de SDQ, EFECO y HADs
                    j = matriz(campos, aux, i, j, diccionario)

                elif campos[j][0:(campos[j].find("_["))] == "Cómo_has_conocido_este_servicio":
                    solucion = ""
                    cont = j
                    # Comprobamos las soluciones a ver cuales son las seleccionadas
                    while cont <= j + 12:
                        if i[cont] != "No" and i[cont] != "N/A" and i[cont] != "":
                            if solucion != "":  # Si la variable "solucion" no esta vacia quiere decir que ya
                                # hay un elemento
                                # Comprobamos si ese elemento es el del campo "Otro" o no
                                if campos[cont][(campos[cont].find("_[") + 2):(len(campos[cont]) - 1)] == "Otro":
                                    solucion = "{} \n -{}".format(solucion, i[cont])
                                else:
                                    solucion = "{} \n -{}".format(solucion, campos[cont][(campos[cont].find("_[") + 2)
                                                                                          :(len(campos[cont]) - 1)])
                            else:  # Si la variable "solucion esta vacia introducimos el primer elemento"
                                # Comprobamos si el primer elemento es el del campo "Otro" o no
                                if campos[cont][(campos[cont].find("_[") + 2):(len(campos[cont]) - 1)] == "Otro":
                                    solucion = "-{}".format(i[cont])
                                else:
                                    solucion = "-{}".format(campos[cont][(campos[cont].find("_[") + 2)
                                                                         :(len(campos[cont]) - 1)])

                        cont += 1
                    # Guardamos la solucion en la variable solucion
                    guardar_en_diccionario(diccionario, campos[j][0:campos[j].find("_[")], solucion)
                    j += 13

                else:
                    # Si no es una pregunta especial, ponemos la opcion por defecto
                    j = por_defecto(campos, aux, i, j, diccionario)
                    j += 1

            elif campos[j][0:(campos[j].find("institucional"))] == "Introduce_tu_direccion_de_correo_":
                # Esta es la unica pregunta especial que no cumple la condicion anterior y que no se puede poner
                # como pregunta normal
                guardar_en_diccionario(diccionario, 'correo', i[j])
                j += 1

            else:
                # Preguntas normales, las que estan vacias o con N/A (no respondido) no se ponen y las que
                # si lo estan se guardan
                if i[j] == "N/A" or i[j] == "":
                    j += 1
                else:
                    guardar_en_diccionario(diccionario, campos[j], i[j])
                    j += 1
        # Tras guardar los valores pasamos a las secciones de calculo
        calcular_edad(diccionario)
        calcular_tiempo(diccionario)
        # Traducimos los campos y los guadamos en un nuevo diciconario
        nuevo_diccionario = traducir(diccionario)
        # Con este nuevo diccionario empezamos crear un nuevo docx
        pip.render(nuevo_diccionario)
        sesiones.render(nuevo_diccionario)

        os.chdir('{}/{}'.format(os.getcwd(), OUTPUT_PATH))
        usuario = '{} {}'.format(diccionario['Nombre'], diccionario['Apellidos'])

        if os.path.isdir(usuario):
            # Guardamos el word con el nombre de los usuarios
            os.chdir('{}/{}'.format(os.getcwd(), usuario))
            pip.save('{} {} V{}.docx'.format(diccionario['Nombre'], diccionario['Apellidos'],
                                                     numero_usuario))
            os.chdir('{}/{}'.format(DIRECTORY, OUTPUT_PATH))
            os.rename(usuario, '@{} {}'.format(diccionario['Nombre'], diccionario['Apellidos']))
            numero_usuario += 1
        else:
            # Guardamos el word con el nombre de los usuarios
            os.mkdir(usuario)
            os.chdir('{}/{}'.format(os.getcwd(), usuario))
            pip.save('{} {}.docx'.format(diccionario['Nombre'], diccionario['Apellidos']))
            sesiones.save('Sesiones {}.docx'.format(diccionario['Nombre']))
        # Vaciamos el diccionario para poder guardar los datos del nuevo usuario
        diccionario = {}
        os.chdir(DIRECTORY)


def main():
    datos = extraer_datos(CSV_FILE)
    eliminar_y_crear_carpeta(OUTPUT_PATH)
    crear_words(datos)


if __name__ == '__main__':
    main()
