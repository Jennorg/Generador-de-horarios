# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.styles import Alignment
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, PageBreak, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors  # Importar colors desde reportlab.lib

class Horario:
    def __init__(self):
        # Inicialización del horario vacío
        self.horario = None

    def definir_horario(self, horario_dict):
        self.horario = horario_dict

    def guardar_en_excel(self, datos):
        archivos_generados = []  # Lista para almacenar los nombres de los archivos generados
        for carrera, semestres in datos.items():
            nombre_archivo = f"{carrera}.xlsx"
            archivos_generados.append(nombre_archivo)
            
            # Crear un nuevo libro de Excel
            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # Eliminar la hoja predeterminada creada por openpyxl
            
            for semestre, secciones in semestres.items():

                if isinstance(secciones, list):  # Si secciones es una lista
                    for idx, horarios in enumerate(secciones):
                        seccion = f"seccion {idx + 1}"
                        self._crear_hoja(wb, semestre, seccion, horarios)

            # Guardar el archivo Excel
            wb.save(nombre_archivo)

        # Convertir los archivos de Excel a PDF
        for archivo in archivos_generados:
            self.convertir_a_pdf(archivo)

    def _crear_hoja(self, wb, semestre, seccion, horarios):
        # Crear una hoja para el semestre y la sección
        sheet = wb.create_sheet(title=f"{semestre} {seccion}")

        # Definir las horas como primera columna
        horas = [
            "7:50 A 8:40", "8:45 A 9:35", "9:35 A 10:25",
            "10:30 A 11:20", "11:20 A 12:10", "12:15 A 1:10",
            "1:10 A 2:00", "2:00 A 2:50", "2:55 A 3:45"
        ]
        # Escribir las horas en la primera columna
        sheet.cell(row=1, column=1, value="Hora")
        for i, hora in enumerate(horas):
            sheet.cell(row=i+2, column=1, value=hora)
            sheet.cell(row=i+2, column=1).alignment = Alignment(horizontal='center', vertical='center')

        # Definir los días como encabezados de columnas
        dias = ["lunes", "martes", "miercoles", "jueves", "viernes"]
        for i, dia in enumerate(dias):
            sheet.cell(row=1, column=i+2, value=dia.capitalize())
            sheet.cell(row=1, column=i+2).alignment = Alignment(horizontal='center', vertical='center')

        # Obtener los datos para escribir en el archivo Excel
        for dia, bloques in horarios.items():
            for bloque, materia in enumerate(bloques):
                if materia is not None:
                    # Crear el valor de la celda con la información de la materia, el profesor y la sección
                    valor_celda = f"{materia['materia']} - {materia['profesor']} Aula: {materia['aula']}"
                    # Escribir en la celda correspondiente
                    sheet.cell(row=bloque + 2, column=dias.index(dia.lower()) + 2, value=valor_celda)
                    sheet.cell(row=bloque + 2, column=dias.index(dia.lower()) + 2).alignment = Alignment(horizontal='center', vertical='center')
                else:
                    sheet.cell(row=bloque + 2, column=dias.index(dia.lower()) + 2, value="")
                    sheet.cell(row=bloque + 2, column=dias.index(dia.lower()) + 2).alignment = Alignment(horizontal='center', vertical='center')
                    
        # Ajustar el tamaño de las celdas para que se vea todo el contenido
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter  # Obtener el nombre de la columna
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width

    def convertir_a_pdf(self, archivo_excel):
        # Leer el archivo Excel con pandas
        df = pd.read_excel(archivo_excel, sheet_name=None)

        # Crear documento PDF
        nombre_pdf = archivo_excel.replace('.xlsx', '.pdf')
        doc = SimpleDocTemplate(nombre_pdf, pagesize=landscape(A4))  # Orientación horizontal

        # Lista para almacenar los elementos del PDF
        elements = []

        for semestre, df_sheet in df.items():
            # Añadir título del semestre
            elements.append(Paragraph(semestre, getSampleStyleSheet()['Heading1']))

            # Convertir DataFrame de pandas a lista de listas para la tabla de reportlab
            data = [df_sheet.columns.values.tolist()] + df_sheet.values.tolist()

            # Ajustar el tamaño del texto para que quepa en las celdas
            col_widths = [1.5 * inch] * len(df_sheet.columns)

            # Reemplazar None por ""
            adjusted_data = []
            for row in data:
                adjusted_row = []
                for cell in row:
                    if pd.isna(cell):
                        adjusted_row.append("")
                    else:
                        adjusted_row.append(cell)
                adjusted_data.append(adjusted_row)

            # Crear tabla de reportlab
            table = Table(adjusted_data, colWidths=col_widths)

            # Estilo de la tabla
            style = TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Corregir importación de colors
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('FONTSIZE', (0, 0), (-1, -1), 8),  # Tamaño de fuente reducido
            ])

            table.setStyle(style)

            # Ajustar el texto dentro de las celdas usando `para` en lugar de texto simple
            adjusted_table = []
            for row in adjusted_data:
                adjusted_row = []
                for cell in row:
                    if isinstance(cell, str):
                        adjusted_cell = self.wrap_text(cell, 1.5 * inch, 8)
                    else:
                        adjusted_cell = cell
                    adjusted_row.append(adjusted_cell)
                adjusted_table.append(adjusted_row)

            # Crear tabla de reportlab con los datos ajustados
            table = Table(adjusted_table, colWidths=col_widths)
            table.setStyle(style)

            # Añadir la tabla al documento
            elements.append(table)
            elements.append(PageBreak())  # Salto de página después de cada tabla

        # Construir el documento PDF
        doc.build(elements)


    def wrap_text(self, text, width, font_size):
        max_chars = int(width / (font_size * 0.5))  # Máximo número de caracteres por línea

        # Dividir el texto en palabras
        words = text.split()
        lines = []
        current_line = ""

        for word in words:
            if len(current_line + " " + word) <= max_chars:
                current_line += " " + word
            else:
                lines.append(current_line.strip())
                current_line = word
        
        lines.append(current_line.strip())
        
        return "\n".join(lines)