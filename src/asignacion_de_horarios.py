from tabulate import tabulate
from saveHorario import Horario
from asignar import Recuperar_Datos
import random

# Container de Horario Actual:  dict(Carrera)<dict(Semestre)<list(Seccion)<dict(dia)<list(bloque)<dict>>>>>

datos = Recuperar_Datos()
datos.insertar_aulas()
datos.insertar_bloque()
datos.insertar_carreras()
datos.insertar_materias()
datos.insertarProfesores()
datos.insertar_globarVar()

horario = dict()

profesores = datos.profesores
materias = datos.materias
aulas = datos.aulas
bloques = datos.bloques
carreras = datos.carreras
globalVar = datos.globalVar


def profesor_ocupado(dia, bloque_inicio, bloque_fin, cedula): # se buscara al profesor en la lista de profesores y se verificara si esta ocupado o no
    profesor = buscar_profesor(cedula)
    
    horario_profesor = profesor["horario"]
    for bloque in range(bloque_inicio, bloque_fin + 1):
        if horario_profesor[dia][bloque] is not None:
            return True
    return False


# Busca en la lista de dias de disponibilidad del profesor si el dia esta disponible o no, dia es un int
def dia_disp_profesor(dia, cedula_profesor):
    profesor = buscar_profesor(cedula_profesor)
    dias_disp = profesor["dias_disp"]
    
    return dia in dias_disp

def aula_ocupada(dia, bloque_inicio, bloque_fin, id_aula):
    for aula in aulas:
        if aula["id_aula"] == id_aula:
            horario_aula = aula["horario"]
            if dia in horario_aula and bloque_inicio in horario_aula[dia]:
                for bloque in range(bloque_inicio, bloque_fin + 1):
                    if horario_aula[dia][bloque] is not None:
                        return True
            break
    return False

def verificar_materia_asignada(seccion, materia_nombre):
    for dia in seccion:
        for bloque in seccion[dia]:
            if bloque and bloque.get("materia") == materia_nombre:
                return True
    return False


def filtro_bloque_por_prioridad(prioridad, lista):
    lista_filtrada = [bloque for bloque in lista if bloque["prioridad"] == prioridad]
    return lista_filtrada

def obtener_bloques_inicio_final_materia(bloques):
    if not bloques:
        return None, None
    # print(bloques)
    return (int(bloques[0]["id"]), int(bloques[-1]["id"]))

def buscar_profesor(cedula):
    for profesor in profesores:
        if profesor["cedula"] == cedula:
            return profesor
    
    raise Exception(f"No se encontró el profesor con cedula {cedula}")

def buscar_aula(codigo):
    for aula in aulas:
        if aula["codigo"] == codigo:
            return aula
    
    raise Exception(f"No se encontró la aula con codigo {codigo}")

def asignar_bloques_horario(materia, inicio, final, dia, prof, seccion, aula_temp):
    profesor = buscar_profesor(prof)
    
    
    for bloque in range(inicio, final + 1):
        
        if aula_temp != {  "id_aula": "Virtual" }:
            aula = buscar_aula(aula_temp["codigo"])
            aula["horario"][dia][bloque-1] = {
                "materia": materia["nombre"],
                "profesor": profesor['nombre'],
                "cedula_profesor": prof,
                "bloque_inicial": inicio,
                "bloque_final": final,
                "aula": aula["id_aula"],
                "modalidad": materia["modalidad"],
                "codigo": materia["codigo"]
            }
        else:
            aula = aula_temp
        
        profesor["horario"][dia][bloque-1] = {
            "materia": materia["nombre"],
            "profesor": profesor['nombre'],
            "cedula_profesor": prof,
            "bloque_inicial": inicio,
            "bloque_final": final,
            "aula": aula["id_aula"],
            "modalidad": materia["modalidad"],
            "codigo": materia["codigo"]
        }
        
        horario[materia["carrera"]]["semestre " + str(materia["semestre"])][seccion][dia][bloque-1] = {
            "materia": materia["nombre"],
            "profesor": profesor['nombre'],
            "cedula_profesor": prof,
            "bloque_inicial": inicio,
            "bloque_final": final,
            "aula": aula["id_aula"],
            "modalidad": materia["modalidad"],
            "codigo": materia["codigo"]
        }

def obtener_cantidad_secciones_semestre(carrera, semestre):
    secciones = 1
    for materia in materias:
        if materia["carrera"] == carrera and materia["semestre"] == semestre and materia["secciones"] > secciones:
            secciones = materia["secciones"]
    return secciones

def obtener_bloques_trabajables(prioridad, profesor):
    horario_profesor = buscar_profesor(profesor[0])["tiempo"]

    inicio = None
    final = None

    if horario_profesor == "Mañana":
        inicio = 1
        final = globalVar['Bloques maximos para el horario de la mañana']
        return inicio, final, None
    elif horario_profesor == "Tarde":
        inicio = globalVar['Bloques maximos para el horario de la mañana'] + 1
        final = globalVar['Maximo de Bloques por dia']
        return inicio, final, None
    else:
        if prioridad == "Media" or prioridad == "Baja":
            inicio = 1
            final = globalVar['Maximo de Bloques por dia']
        else:
            inicio = 1
            final = globalVar['Bloques maximos para el horario de la mañana']
        return inicio, final, None

def imprimir_horarios():
    for carrera, semestres in horario.items():
        print(f"Carrera: {carrera}")
        for semestre, secciones in semestres.items():
            print(f"  Semestre: {semestre}")
            for i, seccion in enumerate(secciones):
                print(f"    Sección {i + 1}:")
                for dia, bloques in seccion.items():
                    print(f"      {dia.capitalize()}:")
                    for j, bloque in enumerate(bloques):
                        if bloque:
                            print(f"        Bloque {j + 1}: {bloque}")

def cantidad_de_materias_por_dia(carrera, semestre, seccion, dia):
    cantidad_materias = 0
    
    ultimo_bloque = None
    for bloque in horario[carrera]["semestre " + str(semestre)][seccion][dia]:
        if ultimo_bloque != bloque and bloque != None:
            cantidad_materias += 1
        ultimo_bloque = bloque
        
    return cantidad_materias

def main():
    for materia in materias:    
        if materia["carrera"] not in horario:
            horario[materia["carrera"]] = dict()
        
        if ("semestre " + str(materia["semestre"])) not in horario[materia["carrera"]]:
            semestre = materia["semestre"]
            cantidad_secciones = obtener_cantidad_secciones_semestre(materia["carrera"], materia["semestre"])
            horario[materia["carrera"]][f"semestre {semestre}"] = [dict({ "lunes": [None] * globalVar['Maximo de Bloques por dia'], "martes": [None] * globalVar['Maximo de Bloques por dia'], "miercoles": [None] * globalVar['Maximo de Bloques por dia'], "jueves": [None] * globalVar['Maximo de Bloques por dia'], "viernes": [None] * globalVar['Maximo de Bloques por dia'] }) for _ in range(cantidad_secciones)]
        
        seccion = horario[materia["carrera"]]["semestre " + str(materia["semestre"])]
        
        seccion_actual = -1

        for profesor in materia["profesor"]:
            for sec in range(int(profesor[1])):
                seccion_actual += 1
                
                if verificar_materia_asignada(seccion[seccion_actual], materia["nombre"]):
                    continue

                intentos = 0
                prioridad_interna = materia["prioridad"]

                if materia["UC"] >= 4:  # 3 bloques
                    i_dia = 0
                    
                    while i_dia < 5:
                        dia_exitoso = False
                        
                        # Hace el recorrido de la lista de dias prioritarios del profesor, en caso de que no se encuentre disponible, se intenta con la siguiente dia
                        if dia_disp_profesor(i_dia, profesor[0]) is False and intentos < 2:
                            i_dia += 1
                            continue
                        
                        inicio, final, err = obtener_bloques_trabajables(prioridad_interna, profesor)

                        if inicio is None or final is None:
                            print(f"No se pudo asignar {materia['nombre']} en ninguno de los días para sección {seccion_actual}")
                            continue
                        
                        if i_dia == 0:
                            dia = seccion[seccion_actual]["lunes"]
                        elif i_dia == 1:
                            dia = seccion[seccion_actual]["martes"]
                        elif i_dia == 2:
                            dia = seccion[seccion_actual]["miercoles"]
                        elif i_dia == 3:
                            dia = seccion[seccion_actual]["jueves"]
                        elif i_dia == 4:
                            dia = seccion[seccion_actual]["viernes"]
                        
                        if cantidad_de_materias_por_dia(materia["carrera"], materia["semestre"], seccion_actual, list(seccion[seccion_actual].keys())[i_dia]) > globalVar['Maximo de Materias Por Dia']:
                            i_dia += 1
                            continue
                        
                        for i_bloque in range(inicio, final - 2):  # Ajustar rango para no exceder
                            # Comprueba si el profesor está ocupado en el bloque actual, en caso de que esté ocupado, se intenta con el siguiente bloque
                            if profesor_ocupado(list(seccion[seccion_actual].keys())[i_dia], i_bloque, i_bloque + 2, profesor[0]) is True:
                                continue
                            
                            aula = None
                            err = False
                            while True:
                                print(materia['nombre'], materia['codigo'], materia['modalidad'])
                                if materia['modalidad'] == "Virtual":
                                    aula = { "id_aula": "Virtual" }
                                    break
                                
                                if materia['aula'] is not None:
                                    temp = buscar_aula(int(materia['aula']))
                                    
                                    if aula_ocupada(list(seccion[seccion_actual].keys())[i_dia], inicio, final, temp["id_aula"]) is False:
                                        aula = temp
                                        break
                                    else:
                                        err = True
                                        break
                                else:
                                    temp = random.choice(aulas)
                                        
                                    if aula_ocupada(list(seccion[seccion_actual].keys())[i_dia], inicio, final, temp["id_aula"]) is False:
                                        aula = temp
                                        break
                            
                            if err is True:
                                continue
                            
                            if dia[i_bloque] is None and dia[i_bloque + 1] is None and dia[i_bloque + 2] is None:
                                print(f"Asignando {materia['nombre']} en {list(seccion[seccion_actual].keys())[i_dia]} para bloques {i_bloque} y {i_bloque + 2} de sección {seccion_actual}")
                                asignar_bloques_horario(materia, i_bloque, i_bloque + 2, list(seccion[seccion_actual].keys())[i_dia], profesor[0], seccion_actual, aula)
                                dia_exitoso = True
                                break

                        if dia_exitoso:
                            break
                        
                        i_dia += 1
                        
                        if i_dia == 5:
                            if intentos == 0:
                                prioridad_interna == "Media"
                                intentos += 1
                            elif intentos == 1:
                                prioridad_interna == "Baja"
                                intentos += 1
                            else:
                                print(f"No se pudo asignar {materia['nombre']} en ninguno de los días para sección {seccion_actual}")
                                break
                else:  # 2 bloques
                    i_dia = 0

                    while i_dia < 5:
                        dia_exitoso = False
                        
                        inicio, final, err = obtener_bloques_trabajables(prioridad_interna, profesor)

                        # Hace el recorrido de la lista de dias prioritarios del profesor, en caso de que no se encuentre disponible, se intenta con la siguiente dia
                        if dia_disp_profesor(i_dia, profesor[0]) is False and intentos < 2:
                            i_dia += 1
                            continue

                        if inicio is None or final is None:
                            print(f"No se pudo asignar {materia['nombre']} en ninguno de los días para sección {seccion_actual}")
                            continue
                        
                        if i_dia == 0:
                            dia = seccion[seccion_actual]["lunes"]
                        elif i_dia == 1:
                            dia = seccion[seccion_actual]["martes"]
                        elif i_dia == 2:
                            dia = seccion[seccion_actual]["miercoles"]
                        elif i_dia == 3:
                            dia = seccion[seccion_actual]["jueves"]
                        elif i_dia == 4:
                            dia = seccion[seccion_actual]["viernes"]
                        
                        print(materia["carrera"], materia["semestre"], seccion_actual, list(seccion[seccion_actual].keys())[i_dia])
                        if cantidad_de_materias_por_dia(materia["carrera"], materia["semestre"], seccion_actual, list(seccion[seccion_actual].keys())[i_dia]) == globalVar['Maximo de Materias Por Dia']:
                            i_dia += 1
                            continue
                        
                        for i_bloque in range(inicio, final - 1):  # Ajustar rango para no exceder
                            
                            # Comprueba si el profesor está ocupado en el bloque actual, en caso de que esté ocupado, se intenta con el siguiente bloque    
                            if profesor_ocupado(list(seccion[sec].keys())[i_dia], i_bloque, i_bloque + 1, profesor[0]) is True:
                                print(f"No se pudo asignar {materia['nombre']} en ninguno de los días para sección {sec}")
                                continue
                            
                            aula = None
                            err = False
                            
                            while True:
                                
                                print(materia['nombre'], materia['codigo'], materia['modalidad'])
                                if materia['modalidad'] == "Virtual":
                                    aula = { "id_aula": "Virtual" }
                                    break
                                
                                if materia['aula'] is not None:
                                    temp = buscar_aula(int(materia['aula']))
                                    
                                    if aula_ocupada(list(seccion[seccion_actual].keys())[i_dia], inicio, final, temp["id_aula"]) is False:
                                        aula = temp
                                        break
                                    else:
                                        err = True
                                        break
                                else:
                                    temp = random.choice(aulas)
                                        
                                    if aula_ocupada(list(seccion[seccion_actual].keys())[i_dia], inicio, final, temp["id_aula"]) is False:
                                        aula = temp
                                        break
                            
                            if err is True:
                                continue
                            
                            if dia[i_bloque] is None and dia[i_bloque + 1] is None:
                                asignar_bloques_horario(materia, i_bloque, i_bloque + 1, list(seccion[sec].keys())[i_dia], profesor[0], seccion_actual, aula)
                                dia_exitoso = True
                                break

                        if dia_exitoso:
                            break
                        
                        i_dia += 1
                        
                        if i_dia == 5:
                            if intentos == 0:
                                prioridad_interna == "Media"
                                intentos += 1
                            elif intentos == 1:
                                prioridad_interna == "Baja"
                                intentos += 1
                            else:
                                print(f"No se pudo asignar {materia['nombre']} en ninguno de los días para sección {seccion_actual}")
                                break
            
    # imprimir_horarios()  # Llamar a la función para imprimir los horarios
    # print(buscar_profesor("V10929321"))
    return
    
        
main()


print(globalVar['Maximo de Materias Por Dia'])
creador = Horario()
creador.guardar_en_excel(horario)
