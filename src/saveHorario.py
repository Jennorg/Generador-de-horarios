# -*- coding: utf-8 -*- 
# Define el encoding del archivo como UTF-8 para asegurar el manejo correcto de caracteres especiales.

import openpyxl # Importa la biblioteca openpyxl para manipular archivos de Exce
from openpyxl.styles import Alignment # Importa la clase Alignment de openpyxl para alinear el contenido de las celdas
from openpyxl import Workbook # Importa la clase Workbook de openpyxl para crear nuevos libros de trabajo de Excel
from openpyxl.styles import PatternFill # Importa la clase PatternFill de openpyxl para aplicar colores de relleno a las celdas
from openpyxl.utils import coordinate_to_tuple # Importa la función coordinate_to_tuple de openpyxl para convertir coordenadas de celdas a tuplas
from openpyxl.worksheet.cell_range import CellRange # Importa la clase CellRange de openpyxl para trabajar con rangos de celdas
import pandas as pd # Importa la biblioteca pandas para trabajar con estructuras de datos, especialmente DataFrames
from reportlab.lib.pagesizes import A4 # Importa la constante A4 de reportlab para establecer el tamaño de las páginas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, PageBreak, Paragraph, Image # Importa varias clases de reportlab.platypus para crear documentos PDF
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle # Importa funciones para obtener estilos predefinidos y crear estilos personalizados
from reportlab.lib.units import inch # Importa la constante inch de reportlab para trabajar con unidades de pulgada
from reportlab.lib import colors # Importa el módulo colors de reportlab para usar colores predefinidos
from reportlab.lib.colors import yellow, blue # Importa los colores amarillo y azul de reportlab.lib.colors
import os  # Importa el módulo 'os' para manipular rutas de archivos y directorios.

class Horario:
    def __init__(self):
        self.horario = None  # Inicializa la variable 'horario' como None.

    def definir_horario(self, horario_dict):
        self.horario = horario_dict  # Define el horario usando un diccionario proporcionado.

    def obtener_nombre_carrera(self, codigo_carrera):
        archivo_datos_horarios = 'DATOSHORARIOS.xlsx'  # Nombre del archivo de Excel con datos de horarios.
        wb = openpyxl.load_workbook(archivo_datos_horarios)  # Carga el libro de trabajo de Excel.
        # Verifica si la hoja 'Carreras' no está presente en el libro de trabajo
        if 'Carreras' not in wb.sheetnames:
            raise KeyError(f"La hoja 'Carreras' no existe en el archivo {archivo_datos_horarios}.")  # Lanza el error.
        
        hoja_carreras = wb['Carreras']  # Obtiene la hoja 'Carreras'.
        for fila in hoja_carreras.iter_rows(min_row=2, values_only=True):  # Itera sobre las filas, empezando desde la segunda.
            vacio, codigo, nombre = fila[:3]  # Toma las primeras tres columnas: vacío, código, y nombre.
             # Compara el código de la carrera con el código proporcionad
            if codigo == codigo_carrera:
                return nombre  # Devuelve el nombre si el código coincide.
        return None  # Devuelve None si no se encuentra el código.

    def guardar_en_excel(self, datos):
        archivos_generados = []  # Lista para almacenar los nombres de archivos generados.
        for carrera, semestres in datos.items():
            nombre_carrera = self.obtener_nombre_carrera(carrera)  # Obtiene el nombre de la carrera basado en el código.
              # Si se encuentra el nombre de la carrera
            if nombre_carrera:
                nombre_archivo = f"{nombre_carrera}.xlsx"  # Usa el nombre de la carrera para el archivo.
              # Si no se encuentra el nombre de la carrera
            else:
                nombre_archivo = f"{carrera}.xlsx"  # Usa el código de la carrera si no se encuentra el nombre.
            archivos_generados.append(nombre_archivo)  # Añade el nombre del archivo a la lista.
            wb = openpyxl.Workbook()  # Crea un nuevo libro de trabajo.
            wb.remove(wb.active)  # Elimina la hoja activa por defecto.
             # Itera sobre los semestres y sus secciones
            for semestre, secciones in semestres.items():
                # Verifica si las secciones son una lista
                if isinstance(secciones, list):
                   # Itera sobre las secciones y sus horarios
                    for idx, horarios in enumerate(secciones):
                        seccion = f"seccion {idx + 1}"
                        self._crear_hoja(wb, semestre, seccion, horarios)  # Crea una hoja para cada sección.
            wb.save(nombre_archivo)  # Guarda el libro de trabajo en el archivo.
        for archivo in archivos_generados:
            self.convertir_a_pdf(archivo)  # Convierte cada archivo Excel a PDF.

    def _crear_hoja(self, wb, semestre, seccion, horarios):
        sheet = wb.create_sheet(title=f"{semestre} {seccion}")  # Crea una nueva hoja con el título especificado.
        # Lista de horas para las filas de la hoja
        horas = [
            "7:50 A 8:40", "8:45 A 9:35", "9:35 A 10:25",
            "10:30 A 11:20", "11:20 A 12:10", "12:15 A 1:10",
            "1:10 A 2:00", "2:00 A 2:50", "2:55 A 3:45"
        ]  
        sheet.cell(row=1, column=1, value="Hora")  # Establece el encabezado de la columna de horas.
         # Itera sobre las horas con su índice
        for i, hora in enumerate(horas):
            sheet.cell(row=i+2, column=1, value=hora)  # Rellena la columna de horas.
            sheet.cell(row=i+2, column=1).alignment = Alignment(horizontal='center', vertical='center')  # Centra el texto en la celda.
        dias = ["lunes", "martes", "miercoles", "jueves", "viernes"]  # Días de la semana.
          # Itera sobre los días con su índice
        for i, dia in enumerate(dias):
            sheet.cell(row=1, column=i+2, value=dia.capitalize())  # Establece los encabezados de los días de la semana.
            sheet.cell(row=1, column=i+2).alignment = Alignment(horizontal='center', vertical='center')  # Centra el texto en la celda.
         # Itera sobre los horarios de cada día
        for dia, bloques in horarios.items():
            # Itera sobre los bloques y las materias de cada día  
            for bloque, materia in enumerate(bloques):
               # Si hay una materia en el bloque
                if materia is not None:
                    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Define el color de fondo amarillo.
                    blue_fill = PatternFill(start_color="00A8FF", end_color="00A8FF", fill_type="solid")  # Define el color de fondo azul.
                    valor_celda = f"{materia['materia']} - {materia['profesor']} - Aula: {materia['aula']} - {materia['modalidad']}"  # Formatea el texto de la celda.
                    cell = sheet.cell(row=bloque + 2, column=dias.index(dia.lower()) + 2, value=valor_celda)  # Rellena la celda con la información.
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)  # Centra el texto y permite el ajuste de línea.
                    #si materia es presencial
                    if 'Presencial' in valor_celda:
                        cell.fill = yellow_fill  # Aplica el color amarillo para clases presenciales.
                    #si materia es virtual
                    elif 'Virtual' in valor_celda:
                        cell.fill = blue_fill  # Aplica el color azul para clases virtuales.
                else:
                    cell = sheet.cell(row=bloque + 2, column=dias.index(dia.lower()) + 2, value="")  # Crea una celda vacía si no hay materia.
                    cell.alignment = Alignment(horizontal='center', vertical='center')  # Centra el texto en la celda vacía.
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)  # Calcula la longitud máxima del contenido de la columna.
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = 25  # Ajusta el ancho de la columna.

    def convertir_a_pdf(self, archivo_excel):
        df = pd.read_excel(archivo_excel, sheet_name=None)  # Lee el archivo Excel y carga todas las hojas en un diccionario.
        nombre_pdf = archivo_excel.replace('.xlsx', '.pdf')  # Define el nombre del archivo PDF.
        doc = SimpleDocTemplate(nombre_pdf, pagesize=A4)  # Crea un documento PDF con tamaño A4.

        elements = []  # Lista para almacenar los elementos del PDF.

        nombre_archivo = os.path.basename(archivo_excel).split('.')[0]  # Obtiene el nombre del archivo sin extensión.

        encabezado = [
            "UNIVERSIDAD NACIONAL EXPERIMENTAL DE GUAYANA",
            "VICERRECTORADO ACADÉMICO",
            "COORDINACIÓN GENERAL DE PREGRADO",
            f"COORDINACIÓN {nombre_archivo}",
            "Presencial : Amarillo, Virtual : Azul "
        ]  # Define el encabezado del PDF.

        current_dir = os.path.dirname(__file__)  # Obtiene el directorio actual del script.
        logo_path = os.path.join(current_dir, 'images', 'logo.png')  # Construye la ruta al archivo de logo.
        
        elements.append(Spacer(1, 0.2 * inch))  # Agrega un espaciador al documento.

        styles = getSampleStyleSheet()  # Obtiene un conjunto de estilos predefinidos.
        styles.add(ParagraphStyle(name='Center', alignment=1))  # Añade un estilo de párrafo para centrar el texto.

        styles['Heading1'].fontSize = 12  # Ajusta el tamaño de la fuente del estilo 'Heading1'.

        for semestre, df_sheet in df.items():
            if logo_path:
                try:
                    elements.append(Image(logo_path, 1 * inch, 1 * inch))  # Añade el logo al documento.
                except OSError:
                    print(f"No se pudo abrir el logo en la ruta: {logo_path}")  # Maneja el error si no se puede abrir el logo.

            for line in encabezado:
                elements.append(Paragraph(line, styles['Center']))  # Añade cada línea del encabezado al documento.

            elements.append(Paragraph(semestre, styles['Heading1']))  # Añade el nombre del semestre como encabezado.
            data = [df_sheet.columns.values.tolist()] + df_sheet.values.tolist()  # Convierte el DataFrame a una lista de listas.
            col_widths = [1 * inch] * len(df_sheet.columns)  # Define el ancho de las columnas.

            adjusted_data = []
            for row in data:
                adjusted_row = []
                for cell in row:
                    if pd.isna(cell):
                        adjusted_row.append("")  # Reemplaza valores NaN por cadenas vacías.
                    else:
                        adjusted_row.append(cell)  # Mantiene otros valores.
                adjusted_data.append(adjusted_row)  # Añade la fila ajustada a la lista de datos.

            table = Table(adjusted_data, colWidths=col_widths)  # Crea una tabla con los datos ajustados.

            style = TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1),1, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ])  # Define el estilo de la tabla.
            adjusted_table = []
            for row in adjusted_data:
                adjusted_row = []
                for cell in row:
                    if isinstance(cell, str):
                        adjusted_cell = self.wrap_text(cell, 1 * inch, 8)  # Ajusta el texto si es una cadena.
                    else:
                        adjusted_cell = cell  # Mantiene el valor de la celda si no es una cadena.
                    adjusted_row.append(adjusted_cell)  # Añade la celda ajustada a la fila.
                adjusted_table.append(adjusted_row)  # Añade la fila ajustada a la tabla.

            for row_idx, row in enumerate(adjusted_table):
                for col_idx, cell in enumerate(row):
                    if cell:
                        if '' in cell:
                            style.add('BACKGROUND', (col_idx, row_idx), (col_idx, row_idx), colors.white)  # Ajusta el color de fondo para celdas no vacías.
                        if 'Presencial' in cell:
                            style.add('BACKGROUND', (col_idx, row_idx), (col_idx, row_idx), colors.yellow)  # Aplica color de fondo amarillo para 'Presencial'.
                            if row_idx > 0 and adjusted_table[row_idx - 1][col_idx] == cell or adjusted_table[row_idx - 2][col_idx] == cell:
                                style.add('LINEBELOW', (col_idx, row_idx - 1), (col_idx, row_idx - 1), 0, colors.yellow)
                                style.add('LINEABOVE', (col_idx, row_idx), (col_idx, row_idx), 0, colors.yellow)  # Aplica bordes a celdas contiguas con 'Presencial'.
                                adjusted_table[row_idx][col_idx] = ""
                        
                        if 'Virtual' in cell:
                            style.add('BACKGROUND', (col_idx, row_idx), (col_idx, row_idx), colors.lightblue)  # Aplica color de fondo azul claro para 'Virtual'.
                            if row_idx > 0 and adjusted_table[row_idx - 1][col_idx] == cell or adjusted_table[row_idx - 2][col_idx] == cell:
                                style.add('LINEBELOW', (col_idx, row_idx - 1), (col_idx, row_idx - 1), 0, colors.lightblue)
                                style.add('LINEABOVE', (col_idx, row_idx), (col_idx, row_idx), 0, colors.lightblue)  # Aplica bordes a celdas contiguas con 'Virtual'.
                                adjusted_table[row_idx][col_idx] = ""

            table = Table(adjusted_table, colWidths=col_widths)  # Crea la tabla ajustada.
            table.setStyle(style)  # Aplica el estilo a la tabla.

            elements.append(table)  # Añade la tabla al documento.
            elements.append(PageBreak())  # Añade un salto de página al documento.

        doc.build(elements)  # Construye el documento PDF con los elementos añadidos.

    def wrap_text(self, text, width, font_size):
        max_chars = int(width / (font_size * 0.5))  # Calcula el número máximo de caracteres por línea basado en el ancho y tamaño de fuente.

        words = text.split()  # Divide el texto en palabras.
        lines = []
        current_line = ""

        for word in words:
            if len(current_line + " " + word) <= max_chars:
                current_line += " " + word
            else:
                lines.append(current_line.strip())  # Añade la línea actual a la lista de líneas.
                current_line = word
        
        lines.append(current_line.strip())  # Añade la última línea a la lista de líneas.
        
        return "\n".join(lines)  # Une las líneas en un solo texto con saltos de línea.