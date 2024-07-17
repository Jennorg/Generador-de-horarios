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

def profesor_ocupado(dia, bloque_inicio, bloque_fin, cedula): # se buscara al profesor en la lista de profesores y se verificara si esta ocupado o no
    profesor = buscar_profesor(cedula)
    
    horario_profesor = profesor["horario"]
    for bloque in range(bloque_inicio, bloque_fin + 1):
        if horario_profesor[dia][bloque] is not None:
            return True
    return False

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
    return (int(bloques[0]["id"]), int(bloques[-1]["id"]))

def buscar_profesor(cedula):
    for profesor in profesores:
        if profesor["cedula"] == cedula:
            return profesor

def asignar_bloques_horario(materia, inicio, final, dia, prof, seccion, aulas):
    profesor = buscar_profesor(prof)
    
    if "Laboratorio" in materia['nombre']:
        sufijo = random.choice(["LBD", "LCB", "LCS", "LSD"])
        materia['nombre'] += f" {sufijo}"
    
    aula = None
    while True:
        temp = random.choice(aulas)

        if aula_ocupada(dia, inicio, final, temp["id_aula"]) is False:
            aula = temp
            
            break
    
    for bloque in range(inicio, final + 1):
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

def main():
    for materia in materias:    
        if materia["carrera"] not in horario:
            horario[materia["carrera"]] = dict()
        
        if ("semestre " + str(materia["semestre"])) not in horario[materia["carrera"]]:
            semestre = materia["semestre"]
            cantidad_secciones = obtener_cantidad_secciones_semestre(materia["carrera"], materia["semestre"])
            horario[materia["carrera"]][f"semestre {semestre}"] = [dict({ "lunes": [None] * 9, "martes": [None] * 9, "miercoles": [None] * 9, "jueves": [None] * 9, "viernes": [None] * 9 }) for _ in range(cantidad_secciones)]
        
        seccion = horario[materia["carrera"]]["semestre " + str(materia["semestre"])]
        
        seccion_actual = -1
        # print(materia["nombre"], materia["codigo"])
        # print(materia["profesor"])
        for profesor in materia["profesor"]:
            for sec in range(int(profesor[1])):
                seccion_actual += 1
                print(materia["nombre"], materia["codigo"], seccion_actual)
                if verificar_materia_asignada(seccion[seccion_actual], materia["nombre"]):
                    continue

                def obtener_bloques_con_prioridad(prioridad):
                    temp = obtener_bloques_inicio_final_materia(filtro_bloque_por_prioridad(prioridad, bloques))

                    if temp[0] is None or temp[1] is None:
                        temp = obtener_bloques_inicio_final_materia(filtro_bloque_por_prioridad("Alta", bloques))
                    return temp

                if materia["UC"] >= 4:  # 3 bloques
                    i_dia = 0
                    prioridad_interna = materia["prioridad"]
                    
                    inicio, final = obtener_bloques_con_prioridad(prioridad_interna)
                    if inicio is None or final is None:
                        continue

                    while i_dia < 5:
                        dia_exitoso = False
                        
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
                        
                        for i_bloque in range(inicio, final - 2):  # Ajustar rango para no exceder
                            if profesor_ocupado(list(seccion[seccion_actual].keys())[i_dia], i_bloque, i_bloque + 2, profesor[0]) is True:
                                continue
                            
                            if dia[i_bloque] is None and dia[i_bloque + 1] is None and dia[i_bloque + 2] is None:
                                print(f"Asignando {materia['nombre']} en {list(seccion[seccion_actual].keys())[i_dia]} para bloques {i_bloque} y {i_bloque + 2} de sección {seccion_actual}")
                                asignar_bloques_horario(materia, i_bloque, i_bloque + 2, list(seccion[seccion_actual].keys())[i_dia], profesor[0], seccion_actual, aulas)
                                dia_exitoso = True
                                break

                        if dia_exitoso:
                            break
                        
                        i_dia += 1
                        
                        if i_dia == 5:
                            print(f"No se pudo asignar {materia['nombre']} en ninguno de los días para sección {seccion_actual}")
                            break
                else:  # 2 bloques
                    i_dia = 0
                    prioridad_interna = materia["prioridad"]
                    
                    inicio, final = obtener_bloques_con_prioridad(prioridad_interna)
                    if inicio is None or final is None:
                        continue

                    while i_dia < 5:
                        dia_exitoso = False
                        
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
                        for i_bloque in range(inicio, final - 1):  # Ajustar rango para no exceder
                            if profesor_ocupado(list(seccion[sec].keys())[i_dia], i_bloque, i_bloque + 1, profesor[0]) is True:
                                print(f"No se pudo asignar {materia['nombre']} en ninguno de los días para sección {sec}")
                                continue
                            if dia[i_bloque] is None and dia[i_bloque + 1] is None:
                                asignar_bloques_horario(materia, i_bloque, i_bloque + 1, list(seccion[sec].keys())[i_dia], profesor[0], seccion_actual, aulas)
                                dia_exitoso = True
                                break

                        if dia_exitoso:
                            break
                        
                        i_dia += 1
                        
                        if i_dia == 5:
                            print(f"No se pudo asignar {materia['nombre']} en ninguno de los días para sección {seccion_actual}")
                            break
            
    # imprimir_horarios()  # Llamar a la función para imprimir los horarios
    # print(buscar_profesor("V10929321"))
    return
    
        
main()

creador = Horario()
creador.guardar_en_excel(horario)
