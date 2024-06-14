import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import random

# Leer archivo Excel
datos = "DATOSHORARIOS.xlsx"

try:
    profesores = pd.read_excel(datos, sheet_name="Profesores")
    materias = pd.read_excel(datos, sheet_name="Materias")
    modalidades = pd.read_excel(datos, sheet_name="Modalidades")
    aulas = pd.read_excel(datos, sheet_name="Aulas")
    carreras = pd.read_excel(datos, sheet_name="Carreras")
    bloques = pd.read_excel(datos, sheet_name="Bloques")
except Exception as e:
    print(f"Error al leer archivo Excel: {e}")
    exit()

# Convertir 'Dias disponibles (Prioridad)' a string y manejar NaNs
profesores['Dias disponibles (Prioridad)'] = profesores['Dias disponibles (Prioridad)'].fillna('').astype(str)

def asignar_horario(profesores, materias, modalidades, aulas, bloques):
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

# Capturar los horarios asignados
horarios_asignados = asignar_horario(profesores, materias, modalidades, aulas, bloques)

# Verificar si horarios_asignados está vacío
if horarios_asignados.empty:
    print("No se asignaron horarios.")
else:
    print(horarios_asignados.head())

# Verificar las columnas de horarios_asignados
print("Columnas de horarios_asignados:", horarios_asignados.columns.tolist())

# Función para exportar horarios a PDF
def exportar_horarios_a_pdf(horarios_asignados, filename="horarios_universitarios.pdf"):
    c = canvas.Canvas(filename, pagesize=letter)
    c.setFont("Helvetica", 10)

    y_position = 750
    for carrera in horarios_asignados['Carrera'].unique():
        c.drawString(30, y_position, f"Carrera: {carrera}")
        y_position -= 15
        for semestre in horarios_asignados[horarios_asignados['Carrera'] == carrera]['Semestre'].unique():
            c.drawString(30, y_position, f"Semestre: {semestre}")
            y_position -= 15
            for seccion in horarios_asignados[(horarios_asignados['Carrera'] == carrera) & (horarios_asignados['Semestre'] == semestre)]['Sección'].unique():
                c.drawString(30, y_position, f"Sección: {seccion}")
                y_position -= 15
                asignaturas = horarios_asignados[(horarios_asignados['Carrera'] == carrera) & (horarios_asignados['Semestre'] == semestre) & (horarios_asignados['Sección'] == seccion)]
                for _, row in asignaturas.iterrows():
                    c.drawString(30, y_position, f"Asignatura: {row['Asignatura']}")
                    c.drawString(30, y_position - 15, f"Profesor: {row['Profesor']}")
                    c.drawString(30, y_position - 30, f"Día: {row['Día']}")
                    c.drawString(30, y_position - 45, f"Hora: {row['Hora de Inicio']} - {row['Hora de Fin']}")
                    c.drawString(30, y_position - 60, f"Sede: {row['Sede']}, Aula: {row['Aula']}")
                    y_position -= 75
                    if y_position < 100:  # Crear nueva página si no hay espacio suficiente
                        c.showPage()
                        y_position = 750
    c.save()

# Exportar horarios a PDF si no está vacío
if not horarios_asignados.empty:
    exportar_horarios_a_pdf(horarios_asignados)
else:
    print("No se generó ningún horario, no se puede exportar a PDF.")
