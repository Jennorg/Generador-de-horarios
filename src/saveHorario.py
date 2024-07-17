# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.styles import Alignment
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, PageBreak, Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
import os  # Agrega esta línea al inicio de tu archivo

class Horario:
    def __init__(self):
        self.horario = None

    def definir_horario(self, horario_dict):
        self.horario = horario_dict

    def guardar_en_excel(self, datos):
        archivos_generados = []
        for carrera, semestres in datos.items():
            nombre_archivo = f"{carrera}.xlsx"
            archivos_generados.append(nombre_archivo)
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            for semestre, secciones in semestres.items():
                if isinstance(secciones, list):
                    for idx, horarios in enumerate(secciones):
                        seccion = f"seccion {idx + 1}"
                        self._crear_hoja(wb, semestre, seccion, horarios)
            wb.save(nombre_archivo)
        for archivo in archivos_generados:
            self.convertir_a_pdf(archivo)

    def _crear_hoja(self, wb, semestre, seccion, horarios):
        sheet = wb.create_sheet(title=f"{semestre} {seccion}")
        horas = [
            "7:50 A 8:40", "8:45 A 9:35", "9:35 A 10:25",
            "10:30 A 11:20", "11:20 A 12:10", "12:15 A 1:10",
            "1:10 A 2:00", "2:00 A 2:50", "2:55 A 3:45"
        ]
        sheet.cell(row=1, column=1, value="Hora")
        for i, hora in enumerate(horas):
            sheet.cell(row=i+2, column=1, value=hora)
            sheet.cell(row=i+2, column=1).alignment = Alignment(horizontal='center', vertical='center')
        dias = ["lunes", "martes", "miercoles", "jueves", "viernes"]
        for i, dia in enumerate(dias):
            sheet.cell(row=1, column=i+2, value=dia.capitalize())
            sheet.cell(row=1, column=i+2).alignment = Alignment(horizontal='center', vertical='center')
        for dia, bloques in horarios.items():
            for bloque, materia in enumerate(bloques):
                if materia is not None:
                    valor_celda = f"{materia['materia']} - {materia['profesor']} Aula: {materia['aula']}"
                    sheet.cell(row=bloque + 2, column=dias.index(dia.lower()) + 2, value=valor_celda)
                    sheet.cell(row=bloque + 2, column=dias.index(dia.lower()) + 2).alignment = Alignment(horizontal='center', vertical='center')
                else:
                    sheet.cell(row=bloque + 2, column=dias.index(dia.lower()) + 2, value="")
                    sheet.cell(row=bloque + 2, column=dias.index(dia.lower()) + 2).alignment = Alignment(horizontal='center', vertical='center')
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width

    def convertir_a_pdf(self, archivo_excel):
        df = pd.read_excel(archivo_excel, sheet_name=None)
        nombre_pdf = archivo_excel.replace('.xlsx', '.pdf')
        doc = SimpleDocTemplate(nombre_pdf, pagesize=A4)

        elements = []
        # Mapeo de nombres de carrera a encabezados específicos
        carrera_encabezado = {
            "2072": [
                "UNIVERSIDAD NACIONAL EXPERIMENTAL DE GUAYANA",
                "VICERRECTORADO ACADÉMICO",
                "COORDINACIÓN GENERAL DE PREGRADO",
                "COORDINACIÓN INGENIERÍA EN INFORMÁTICA"
            ],
            "6350": [
                "UNIVERSIDAD NACIONAL EXPERIMENTAL DE GUAYANA",
                "VICERRECTORADO ACADÉMICO",
                "COORDINACIÓN GENERAL DE PREGRADO",
                "COORDINACIÓN CONTADURÍA PÚBLICA"
            ],
            "6176": [
                "UNIVERSIDAD NACIONAL EXPERIMENTAL DE GUAYANA",
                "VICERRECTORADO ACADÉMICO",
                "COORDINACIÓN GENERAL DE PREGRADO",
                "COORDINACIÓN GESTIÓN DE ALOJAMIENTO TURÍSTICO"
            ]
            # Agrega más mappings según sea necesario
        }

        # Obtener el nombre del archivo sin extensión para determinar el encabezado
        nombre_archivo = os.path.basename(archivo_excel).split('.')[0]

        if nombre_archivo in carrera_encabezado:
            encabezado = carrera_encabezado[nombre_archivo]
        else:
            encabezado = [
                "UNIVERSIDAD NACIONAL EXPERIMENTAL DE GUAYANA",
                "VICERRECTORADO ACADÉMICO",
                "COORDINACIÓN GENERAL DE PREGRADO",
                f"COORDINACIÓN {nombre_archivo}"
            ]
        # Construir la ruta relativa al directorio del script
        current_dir = os.path.dirname(__file__)
        logo_path = os.path.join(current_dir, 'images', 'logo.png')
        
        elements.append(Spacer(1, 0.2 * inch))

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name='Center', alignment=1))  # Definir estilo 'Center' aquí

        styles['Heading1'].fontSize = 12

        for semestre, df_sheet in df.items():
            if logo_path:
                try:
                    elements.append(Image(logo_path, 1 * inch, 1 * inch))
                except OSError:
                    print(f"No se pudo abrir el logo en la ruta: {logo_path}")

            for line in encabezado:
                elements.append(Paragraph(line, styles['Center']))  # Usar el estilo 'Center' para centrar

            elements.append(Paragraph(semestre, styles['Heading1']))
            data = [df_sheet.columns.values.tolist()] + df_sheet.values.tolist()
            col_widths = [1 * inch] * len(df_sheet.columns)

            adjusted_data = []
            for row in data:
                adjusted_row = []
                for cell in row:
                    if pd.isna(cell):
                        adjusted_row.append("")
                    else:
                        adjusted_row.append(cell)
                adjusted_data.append(adjusted_row)

            table = Table(adjusted_data, colWidths=col_widths)

            style = TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ])
            adjusted_table = []
            for row in adjusted_data:
                adjusted_row = []
                for cell in row:
                    if isinstance(cell, str):
                        adjusted_cell = self.wrap_text(cell, 1 * inch, 8)
                    else:
                        adjusted_cell = cell
                    adjusted_row.append(adjusted_cell)
                adjusted_table.append(adjusted_row)

            table = Table(adjusted_table, colWidths=col_widths)
            table.setStyle(style)

            elements.append(table)
            elements.append(PageBreak())

        doc.build(elements)

    def wrap_text(self, text, width, font_size):
        max_chars = int(width / (font_size * 0.5))

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
