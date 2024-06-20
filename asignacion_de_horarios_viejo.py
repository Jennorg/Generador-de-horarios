from tabulate import tabulate
from saveHorario import Horario
from asignar import Recuperar_Datos

# Container de Horario Actual:  dict(Carrera)<dict(Semestre)<list(Seccion)<dict(dia)<list(bloque)<dict>>>>>

datos = Recuperar_Datos()

datos.insertar_aulas()
datos.insertar_bloque()
datos.insertar_carreras()
datos.insertar_materias()
datos.insertarProfesores()

horario = dict()

profesores = datos.profesores

materias = datos.materias

aulas = datos.aulas

bloques = datos.bloques

carreras = datos.carreras

def asignar_tiempo_horario():
    return

def insertar_profesores():
    return

def insertar_materias():
    return

def insertar_aulas():
    return

def filtro_bloque_por_prioridad(prioridad, lista):
    lista_filtrada = []
    for bloque in lista:
        if bloque["prioridad"] == prioridad:
            lista_filtrada.append(bloque)
    return lista_filtrada

def obtener_bloques_inicio_final_materia(bloques):
    return (int(bloques[0]["id"]), int(bloques[-1]["id"]))

def buscar_profesor(cedula):
    for profesor in profesores:
        if profesor["cedula"] == cedula:
            return profesor

def asignar_bloques_horario(materia, inicio, final, dia, prof, seccion):
    profesor = buscar_profesor(prof)
    
    for bloque in range(inicio, final):
        horario[materia["carrera"]]["semestre " + str(materia["semestre"])][seccion][dia][bloque] = dict({"materia": materia["nombre"], "profesor": profesor["nombre"], "cedula_profesor": prof,"bloque_inicial": inicio, "bloque_final": final, "aula": "12", "modalidad": materia["modalidad"], "codigo": materia["codigo"]})

def obtener_cantidad_secciones_semestre(carrera, semestre):
    secciones = 1
    
    for materia in materias:
        if materia["carrera"] == carrera and materia["semestre"] == semestre and materia["secciones"] > secciones:
            secciones = materia["secciones"]
    
    return secciones

def main():
    for materia in materias:    
        if materia["carrera"] not in horario:
            horario[materia["carrera"]] = dict()
        
        if ("semestre " + str(materia["semestre"])) not in horario[materia["carrera"]]:
            vacio = { "lunes": [None] * 9, "martes": [None] * 9, "miercoles": [None] * 9, "jueves": [None] * 9, "viernes": [None] * 9 }
            semestre = materia["semestre"]
            horario[materia["carrera"]][f"semestre {semestre}"] = [dict(vacio)] * obtener_cantidad_secciones_semestre(materia["carrera"], materia["semestre"])
        
        seccion = horario[materia["carrera"]]["semestre " + str(materia["semestre"])]
        
        for profesor in materia["profesor"]:
            for sec in range(0, int(profesor[1])):
                if materia["UC"] >= 4: # 3 bloques
                    i_dia = 0
                    prioridad_interna = materia["prioridad"]
                    
                    temp = obtener_bloques_inicio_final_materia(filtro_bloque_por_prioridad(prioridad_interna, bloques))
                    inicio = temp[0]
                    final = temp[1]
                    
                    
                    while i_dia < 5:
                        dia_exitoso = False
                        
                        if i_dia == 0: # Lunes
                            dia = seccion[sec]["lunes"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None and dia[i_bloque] is None:
                                        asignar_bloques_horario(materia, inicio, final, "lunes", profesor[0], sec)
                                        dia_exitoso = True
                                        break
                        elif i_dia == 1:
                            dia = seccion["martes"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None and dia[i_bloque] is None:
                                        asignar_bloques_horario(materia, inicio, final, "martes", profesor[0], sec)
                                        dia_exitoso = True
                                        break
                        elif i_dia == 2:
                            dia = seccion["miercoles"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None and dia[i_bloque] is None:
                                        asignar_bloques_horario(materia, inicio, final, "miercoles", profesor[0], sec)
                                        dia_exitoso = True
                                        break
                        elif i_dia == 3:
                            dia = seccion["jueves"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None and dia[i_bloque] is None:
                                        asignar_bloques_horario(materia, inicio, final, "jueves", profesor[0], sec)
                                        dia_exitoso = True
                                        break
                        elif i_dia == 4:
                            dia = seccion["viernes"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None and dia[i_bloque] is None:
                                        asignar_bloques_horario(materia, inicio, final, "viernes", profesor[0], sec)
                                        dia_exitoso = True
                                        break
                        elif i_dia == 5:
                            dia = seccion["sabado"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None and dia[i_bloque] is None:
                                        asignar_bloques_horario(materia, inicio, final, "sabado", profesor[0], sec)
                                        dia_exitoso = True
                                        break

                        if dia_exitoso == True:
                            break
                        
                        i_dia = i_dia + 1

                        if i_dia == 4:
                            i_dia = 0
                            inicio = 0
                            final = len(bloques) - 1     
                else: # 2 bloques
                    i_dia = 0
                    prioridad_interna = materia["prioridad"]
                    
                    temp = obtener_bloques_inicio_final_materia(filtro_bloque_por_prioridad(prioridad_interna, bloques))
                    inicio = temp[0]
                    final = temp[1]
                    
                    
                    while i_dia < 5:
                        dia_exitoso = False
                        
                        if i_dia == 0: # Lunes
                            dia = seccion[sec]["lunes"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None:
                                        asignar_bloques_horario(materia, inicio, final, "lunes", profesor[0], sec)
                                        dia_exitoso = True
                                        break
                        elif i_dia == 1:
                            dia = seccion[sec]["martes"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None:
                                        asignar_bloques_horario(materia, inicio, final, "martes", profesor[0], sec)
                                        dia_exitoso = True
                                        break
                        elif i_dia == 2:
                            dia = seccion[sec]["miercoles"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None:
                                        asignar_bloques_horario(materia, inicio, final, "miercoles", profesor[0], sec)
                                        dia_exitoso = True
                                        break
                        elif i_dia == 3:
                            dia = seccion[sec]["jueves"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None:
                                        asignar_bloques_horario(materia, inicio, final, "jueves", profesor[0], sec)
                                        dia_exitoso = True
                                        break
                        elif i_dia == 4:
                            dia = seccion[sec]["viernes"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None:
                                        asignar_bloques_horario(materia, inicio, final, "viernes", profesor[0], sec)
                                        dia_exitoso = True
                                        break
                        elif i_dia == 5:
                            dia = seccion[sec]["sabado"]
                            for i_bloque in range(inicio, final):
                                    if dia[i_bloque] is None and dia[i_bloque + 1] is None:
                                        asignar_bloques_horario(materia, inicio, final, "sabado", profesor[0], sec)
                                        dia_exitoso = True
                                        break

                        if dia_exitoso == True:
                            break
                        
                        i_dia = i_dia + 1

                        if i_dia == 5:
                            i_dia = 0
                            inicio = 0
                            final = len(bloques) - 1
    return
        
main()

# print(horario)
creador = Horario()

# creador.guardar_en_excel(horario)