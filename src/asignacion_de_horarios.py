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

horario = dict()

profesores = datos.profesores

materias = datos.materias

aulas = datos.aulas

bloques = datos.bloques

carreras = datos.carreras

def filtro_bloque_por_prioridad(prioridad, lista):
    lista_filtrada = [bloque for bloque in lista if bloque["prioridad"] == prioridad]
    return lista_filtrada

def obtener_bloques_inicio_final_materia(bloques):
    if not bloques:
        return None, None
    return (int(bloques[0]["id"]), int(bloques[-1]["id"]))

def buscar_profesor(cedula):
    for profesor in profesores:
        if profesor["cedula"] == cedula:
            return profesor

def asignar_bloques_horario(materia, inicio, final, dia, prof, seccion, aulas):
    # Buscar información del profesor
    profesor = buscar_profesor(prof)
    
    # Modificar el nombre de la materia si contiene "Laboratorio"
    if "Laboratorio" in materia['nombre']:
        # Seleccionar aleatoriamente uno de los sufijos
        sufijo = random.choice(["LBD", "LCB", "LCS", "LSD"])
        materia['nombre'] += f" {sufijo}"
    
    # print(f"Asignando bloques para materia {materia['nombre']} con profesor {profesor['nombre']} en el día {dia}, sección {seccion}")
    
    # Seleccionar un aula aleatoria
    id_aula = random.choice(aulas)['id_aula']
    
    # Asignar bloques en el horario
    for bloque in range(inicio, final):
        horario[materia["carrera"]]["semestre " + str(materia["semestre"])][seccion][dia][bloque] = {
            "materia": materia["nombre"],
            "profesor": profesor['nombre'],
            "cedula_profesor": prof,
            "bloque_inicial": inicio,
            "bloque_final": final,
            "aula": id_aula,
            "modalidad": materia["modalidad"],
            "codigo": materia["codigo"]
        }

def obtener_cantidad_secciones_semestre(carrera, semestre):
    secciones = 1
    for materia in materias:
        if materia["carrera"] == carrera and materia["semestre"] == semestre and materia["secciones"] > secciones:
            secciones = materia["secciones"]
    return secciones

def verificar_materia_asignada(seccion, materia_nombre):
    for dia in seccion:
        for bloque in seccion[dia]:
            if bloque and bloque.get("materia") == materia_nombre:
                return True
    return False

def main():
    for materia in materias:    
        if materia["carrera"] not in horario:
            horario[materia["carrera"]] = dict()
        
        if ("semestre " + str(materia["semestre"])) not in horario[materia["carrera"]]:
            vacio = { "lunes": [None] * 9, "martes": [None] * 9, "miercoles": [None] * 9, "jueves": [None] * 9, "viernes": [None] * 9 }
            semestre = materia["semestre"]
            cantidad_secciones = obtener_cantidad_secciones_semestre(materia["carrera"], materia["semestre"])
            horario[materia["carrera"]][f"semestre {semestre}"] = [dict(vacio)] * cantidad_secciones
        
        seccion = horario[materia["carrera"]]["semestre " + str(materia["semestre"])]
        
        for profesor in materia["profesor"]:
            for sec in range(int(profesor[1])):
                if verificar_materia_asignada(seccion[sec], materia["nombre"]):
                    continue

                def obtener_bloques_con_prioridad(prioridad):
                    temp = obtener_bloques_inicio_final_materia(filtro_bloque_por_prioridad(prioridad, bloques))

                    if temp[0] is None or temp[1] is None:
                        # print(f"No hay bloques disponibles para {materia['nombre']} con prioridad {prioridad}, intentando con prioridad Baja")
                        temp = obtener_bloques_inicio_final_materia(filtro_bloque_por_prioridad("Alta", bloques))
                    return temp

                if materia["UC"] >= 4:  # 3 bloques
                    i_dia = 0
                    prioridad_interna = materia["prioridad"]
                    
                    inicio, final = obtener_bloques_con_prioridad(prioridad_interna)
                    if inicio is None or final is None:
                        # print(f"No hay bloques disponibles para {materia['nombre']} incluso con prioridad Baja")
                        continue
                    # print(f"Materia: {materia['nombre']}, Sección: {sec}, Inicio: {inicio}, Final: {final}")

                    while i_dia < 5:
                        dia_exitoso = False
                        # print(f"Procesando día: {i_dia} para {materia['nombre']} (3 bloques)")
                        
                        if i_dia == 0:  # Lunes
                            dia = seccion[sec]["lunes"]
                        elif i_dia == 1:  # Martes
                            dia = seccion[sec]["martes"]
                        elif i_dia == 2:  # Miércoles
                            dia = seccion[sec]["miercoles"]
                        elif i_dia == 3:  # Jueves
                            dia = seccion[sec]["jueves"]
                        elif i_dia == 4:  # Viernes
                            dia = seccion[sec]["viernes"]
                        
                        for i_bloque in range(inicio, final - 2):  # Asegúrate de no exceder el rango
                            if dia[i_bloque] is None and dia[i_bloque + 1] is None and dia[i_bloque + 2] is None:
                                asignar_bloques_horario(materia, inicio-1, final-1, list(seccion[sec].keys())[i_dia], profesor[0], sec,aulas)
                                dia_exitoso = True
                                break

                        if dia_exitoso:
                            # print(f"Día exitoso para {materia['nombre']} en {list(seccion[sec].keys())[i_dia]}")
                            break
                        
                        i_dia += 1
                        
                        if i_dia == 5:
                            # print(f"No se pudo asignar {materia['nombre']} en ninguno de los días para sección {sec}")
                            break
                else:  # 2 bloques
                    i_dia = 0
                    prioridad_interna = materia["prioridad"]
                    
                    inicio, final = obtener_bloques_con_prioridad(prioridad_interna)
                    if inicio is None or final is None:
                        # print(f"No hay bloques disponibles para {materia['nombre']} incluso con prioridad Baja")
                        continue
                    # print(f"Materia: {materia['nombre']}, Sección: {sec}, Inicio: {inicio}, Final: {final}")

                    while i_dia < 5:
                        dia_exitoso = False
                        # print(f"Procesando día: {i_dia} para {materia['nombre']} (2 bloques)")
                        
                        if i_dia == 0:  # Lunes
                            dia = seccion[sec]["lunes"]
                        elif i_dia == 1:  # Martes
                            dia = seccion[sec]["martes"]
                        elif i_dia == 2:  # Miércoles
                            dia = seccion[sec]["miercoles"]
                        elif i_dia == 3:  # Jueves
                            dia = seccion[sec]["jueves"]
                        elif i_dia == 4:  # Viernes
                            dia = seccion[sec]["viernes"]
                        
                        for i_bloque in range(inicio, final - 1):  # Asegúrate de no exceder el rango
                            if dia[i_bloque] is None and dia[i_bloque + 1] is None:
                                asignar_bloques_horario(materia, inicio, final, list(seccion[sec].keys())[i_dia], profesor[0], sec,aulas)
                                dia_exitoso = True
                                break

                        if dia_exitoso:
                            # print(f"Día exitoso para {materia['nombre']} en {list(seccion[sec].keys())[i_dia]}")
                            break
                        
                        i_dia += 1
                        
                        if i_dia == 5:
                            print(f"No se pudo asignar {materia['nombre']} en ninguno de los días para sección {sec}")
                            break
    return

        
main()

creador = Horario()

creador.guardar_en_excel(horario)
