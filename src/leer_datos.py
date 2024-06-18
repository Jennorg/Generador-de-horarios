import pandas as pd
import random

def leer_datos():
    try:
        archivo = "../data/DATOSHORARIOS.xlsx"
        profesores_df = pd.read_excel(archivo, sheet_name="Profesores")
        materias_df = pd.read_excel(archivo, sheet_name="Materias")
        aulas_df = pd.read_excel(archivo, sheet_name="Aulas")
        modalidades_df = pd.read_excel(archivo, sheet_name="Modalidades")
        carreras_df = pd.read_excel(archivo, sheet_name="Carreras")
        bloques_df = pd.read_excel(archivo, sheet_name="Bloques")
        restricciones_df = pd.read_excel(archivo, sheet_name="Restricciones")
        
        return {
            "profesores": profesores_df,
            "materias": materias_df,
            "aulas": aulas_df,
            "modalidades": modalidades_df,
            "carreras": carreras_df,
            "bloques": bloques_df,
            "restricciones": restricciones_df
        }

    except Exception as e:
        print(f"Error leyendo el archivo Excel: {e}")
        return None

def asignar_horario(datos):
    profesores = datos['profesores']
    materias = datos['materias']
    aulas = datos['aulas']
    modalidades = datos['modalidades']
    bloques = datos['bloques']
    
    horarios_asignados = []

    # Filtrar materias que no tienen información completa
    materias = materias.dropna(subset=['Carrera', 'Semestre', 'Nombre', 'Secciones', 'Modalidad'])

    for _, materia in materias.iterrows():
        carrera = materia['Carrera']
        semestre = materia['Semestre']
        asignatura = materia['Nombre']
        seccion = materia['Secciones']
        modalidad = materia['Modalidad']
        
        print(f"Procesando asignatura: {asignatura}, Carrera: {carrera}, Semestre: {semestre}, Sección: {seccion}, Modalidad: {modalidad}")
        
        modalidad_fila = modalidades[modalidades['Modalidades'] == modalidad]
        if modalidad_fila.empty:
            print(f"Advertencia: Modalidad '{modalidad}' no encontrada en la hoja de modalidades.")
            continue

        prioridad = modalidad_fila['Prioridades'].values[0]
        bloques_disponibles = bloques[bloques['Relacion'] == prioridad]

        # Asignar profesor aleatoriamente
        profesor_asignado = random.choice(profesores['Nombre completo'].tolist())
        print(f"Profesor asignado: {profesor_asignado}")

        for _, bloque in bloques_disponibles.iterrows():
            dia = 'Lunes'  # Asignar un día fijo para simplificar
            if not any((h['Día'], h['Hora de Inicio'], h['Hora de Fin']) == (dia, bloque['Hora de inicio'], bloque['Hora de fin']) for h in horarios_asignados if h['Profesor'] == profesor_asignado):
                aula_disponible = None
                for _, aula in aulas.iterrows():
                    if not any((h['Sede'], h['Aula']) == (aula['Sede'], aula['ID aula']) for h in horarios_asignados if (h['Día'], h['Hora de Inicio'], h['Hora de Fin']) == (dia, bloque['Hora de inicio'], bloque['Hora de fin'])):
                        aula_disponible = aula
                        break

                if aula_disponible is not None:
                    print(f"Aula asignada: {aula_disponible['ID aula']} en la sede {aula_disponible['Sede']}")

                    horarios_asignados.append({
                        'Carrera': carrera,
                        'Semestre': semestre,
                        'Asignatura': asignatura,
                        'Sección': seccion,
                        'Profesor': profesor_asignado,
                        'Día': dia,
                        'Hora de Inicio': bloque['Hora de inicio'],
                        'Hora de Fin': bloque['Hora de fin'],
                        'Sede': aula_disponible['Sede'],
                        'Aula': aula_disponible['ID aula']
                    })
                    break
                else:
                    print("No hay aulas disponibles para este bloque.")
            else:
                print("El profesor ya tiene un horario asignado en este bloque.")

    return pd.DataFrame(horarios_asignados)

# # Leer los datos del archivo Excel
# datos = leer_datos()

# # Verificar que se leyeron correctamente los datos antes de proceder
# if datos is not None:
#     horarios = asignar_horario(datos)
#     print(horarios)
# else:
#     print("No se pudieron leer los datos del archivo Excel.")
