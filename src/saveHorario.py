import openpyxl
from openpyxl.styles import Alignment, Border, Side
import copy
from fpdf import FPDF #para poder transformar a pdf para instalarlo pip install fpdf

class Horario:
    def __init__(self):
        # Nivel 1: Detalle de la materia
        nivel1 = dict([("materia", None), ("profesor", None), ("cedula_profesor", None),
                       ("bloque_inicial", None), ("bloque_final", None), ("aula", None), ("modalidad", None)])

        # Nivel 2: Bloques
        cantidad_bloques = 9  # Asumiendo 9 bloques para el ejemplo
        nivel2 = [nivel1.copy() for _ in range(cantidad_bloques)]

        # Nivel 3: Días de la semana (usando copy.deepcopy para asegurarnos de que cada día tenga su propia copia)
        nivel3 = {
            "lunes": copy.deepcopy(nivel2),
            "martes": copy.deepcopy(nivel2),
            "miercoles": copy.deepcopy(nivel2),
            "jueves": copy.deepcopy(nivel2),
            "viernes": copy.deepcopy(nivel2)
        }

        # Nivel 4: Lista de semanas (en este caso solo una semana para simplificar)
        nivel4 = [nivel3]

        # Nivel 5: Semestres
        nivel5 = {"semestre 1": nivel4}

        # Nivel 6: Horario completo
        self.horario = {"null": nivel5}

    def agregar_materia(self, dia, bloque_inicial, bloque_final, materia, aula):
        for bloque in range(bloque_inicial, bloque_final + 1):
            self.horario["null"]["semestre 1"][0][dia][bloque]["materia"] = materia
            self.horario["null"]["semestre 1"][0][dia][bloque]["aula"] = aula

    def guardar_en_excel(self, archivo):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Horario"

        # Definir la estructura del horario
        horas = [
            "7:50 A 8:40",
            "8:45 A 9:35",
            "9:35 A 10:25",
            "10:30 A 11:20",
            "11:20 A 12:10",
            "12:15 A 1:10",
            "1:10 A 2:00",
            "2:00 A 2:50",
            "2:55 A 3:45"
        ]
        dias = ["HORA", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES"]

        # Escribir los encabezados de las columnas (días de la semana)
        ws.append(dias)

        # Escribir las horas en la primera columna
        for hora in horas:
            ws.append([hora])

        # Rellenar el horario con las materias desde la estructura del diccionario
        for semestre, semanas in self.horario.items():
            for semana, dias_semana in semanas.items():
                for dia, bloques in dias_semana[0].items():
                    skip_rows = 0  # Variable para evitar escribir en filas ya combinadas
                    for bloque_idx, bloque in enumerate(bloques):
                        if skip_rows > 0:
                            skip_rows -= 1
                            continue
                        if bloque["materia"] is not None:
                            # Encuentra la fila correspondiente al bloque
                            row_idx_start = bloque_idx + 2  # +2 para ajustar al encabezado y a la base 0
                            row_idx_end = row_idx_start
                            col_idx = dias.index(dia.upper()) + 1  # +1 porque columnas empiezan en 1

                            # Buscar hasta qué bloque se extiende la materia
                            for i in range(bloque_idx + 1, len(bloques)):
                                if bloques[i]["materia"] == bloque["materia"]:
                                    row_idx_end += 1
                                    skip_rows += 1
                                else:
                                    break

                            # Combinar celdas si la materia ocupa varios bloques
                            if row_idx_start != row_idx_end:
                                ws.merge_cells(start_row=row_idx_start, start_column=col_idx, end_row=row_idx_end, end_column=col_idx)

                            # Concatenar nombre de la materia y aula
                            valor_celda = "{} - {}".format(bloque["materia"], bloque["aula"])
                            # Escribe el nombre de la materia en la celda correspondiente
                            ws.cell(row=row_idx_start, column=col_idx, value=valor_celda)

        # Alinear todas las celdas en el centro y aplicar bordes
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
        # Ajustar el ancho de las columnas a 24 pixeles
        for col in ws.columns:
            column = col[0].column_letter  # Obtener la letra de la columna
            ws.column_dimensions[column].width = 24
        
        # Aplicar bordes a todas las celdas
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for row in ws.iter_rows():
            for cell in row:
                cell.border = thin_border
        
        # Guardar el archivo
        wb.save(archivo)
        print('Información guardada en {}'.format(archivo))







    def convertir_a_pdf(self, archivo_excel, archivo_pdf):
        wb = openpyxl.load_workbook(archivo_excel)
        ws = wb.active

        pdf = FPDF('L', 'mm', 'A4')  # Crear un PDF en modo apaisado
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Ajustar anchos de las columnas y alturas de las filas
        col_widths = []
        for col in ws.iter_cols():
            max_length = 0
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            col_widths.append(max_length * 3)  # Ajuste aproximado de la anchura

        row_height = 10  # Altura de fila por defecto

        # Obtener el contenido del archivo Excel y agregarlo al PDF
        for row in ws.iter_rows(values_only=True):
            for col_idx, cell in enumerate(row):
                if cell is None:
                    pdf.cell(col_widths[col_idx], row_height, ' ', border=1, align='C')
                else:
                    pdf.cell(col_widths[col_idx], row_height, str(cell), border=1, align='C')
            pdf.ln(row_height)

        pdf.output(archivo_pdf)
        print('Información guardada en {}'.format(archivo_pdf))


# Crear instancia de Horario
mi_horario = Horario()

# Agregar materias al horario
mi_horario.agregar_materia("lunes", 0, 2, "Matemáticas", "Aula 101")  # Ocupa bloques 1, 2 y 3
mi_horario.agregar_materia("jueves", 3, 4, "Historia", "Aula 104")  # Ocupa bloques 4 y 5

# Guardar el horario en un archivo de Excel
mi_horario.guardar_en_excel("horario.xlsx")

# Convertir el archivo de Excel a PDF
mi_horario.convertir_a_pdf("horario.xlsx", "horario.pdf")
