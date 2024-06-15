import pandas as pd

def leer_datos():
    try:
        archivo = "data/DATOSHORARIOS.xlsx"
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
