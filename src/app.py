import pandas as pd

archivo = "DATOSHORARIOS.xlsx"

profesores = pd.read_excel(archivo, sheet_name= "Profesores")
materias = pd.read_excel(archivo, sheet_name= "Materias")
aulas = pd.read_excel(archivo, sheet_name= "Aulas")
modalidades = pd.read_excel(archivo, sheet_name="Modalidades")
carreras = pd.read_excel(archivo, sheet_name= "Carreras")
bloques = pd.read_excel(archivo, sheet_name= "Bloques")
restricciones = pd.read_excel(archivo, sheet_name= "Restricciones")

modalidades_alta_prioridad = modalidades[modalidades["Prioridades"] == "Alta"]
aulas_del_atlantico = aulas[aulas["Sede "] == "Atlántico "]
materias_informatica = materias[materias["Carrera"] == "Ing. Informática "]

print(materias_informatica)