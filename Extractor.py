"""
**Si tienes que añadir o eliminar alguna respuesta vas a tener que cambiar un par de cosas de las lineas (26 a 37),
para esto te recomiendo que en PyCharm uses el debuger y veas el contenido de la variable (---datos---) este
deberia de ser tres listas las caules son la 0 que contiene los campos (enunciados) de las preguntas y la 1 que contiene
los valores. Si las despliegas podras ver unos 100 y si le das doble click al mensajito que aparece abajo del todo
podras ver mas. Te recomiendo que vayas probando a cambiar los numeros del "del" y compruebes con el debuger hasta que
obtengas lo que quieres.

**
"""

import csv


def extraer_datos(csv_file):
    with open(csv_file, encoding="utf8") as csvfile:
        resultados = csv.reader(csvfile, delimiter=',')
        datos = list(resultados)

        for i in range(len(datos) - 1):
            # Seccion 1
            # A continuacion vamos a limpiar campos inutiles, esta parte realmente se podria obviar (en parte) con
            # el codigo de la siguiente seccion

            # Eliminamos la informacion de la ID, ultima pagina, lenguaje, semilla, en la lista de los campos
            del datos[i][0]
            del datos[i][1:9]

            # Eliminamos los temporizadores del final, en la lista de los campos
            del datos[i][356:363]
            del datos[i][334:355]
            del datos[i][286:333]
            del datos[i][264:285]
            del datos[i][261:263]

            # Eliminamos la informacion de los creditos, en la lista de los valores y los campos
            del datos[i][57:76]

        for i in range(len(datos) - 1):
            for j in range(len(datos[i])):
                datos[i][j] = datos[i][j].replace(" ", "_").replace("¿", "").replace("?", "")

    return datos
