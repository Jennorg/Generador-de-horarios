import pandas as pd

dirExcel = "./DATOSHORARIOS.xlsx"

class Recuperar_Datos:
  def __init__(self):
    self.profesores = list()
    self.materias = list()
    self.aulas = list()
    self.carreras = list()
    self.bloques = list()

  def insertarProfesores(self):
    self.profesores = list()
    profExcel = pd.DataFrame()
    
    try:
      profExcel = pd.read_excel(dirExcel, sheet_name="Profesores", skiprows=[0])
      profExcel = profExcel.drop(columns=['Unnamed: 0'])
      
    except Exception as excepcion:
      print("Ocurrió un error al leer el archivo excel.")
      print(f"Ocurrió una Excepción de tipo {excepcion}")
      exit()

    for indice, fila in profExcel.iterrows():
      cedulaP = fila['Cédula']
      nombreP = fila['Nombre completo']
      materias = fila['Materias']
      self.asignarTuplaProf(materias, cedulaP)
      self.profesores.append(dict({"cedula": cedulaP, "nombre": nombreP, "materias": materias}))
   
    return self.profesores
  
  def asignarTuplaProf(self, mat: str, cedula):
    temp = mat.split('/')

    for materia_secciones in temp:
      codigo = materia_secciones.split('-')[0]
      secciones = materia_secciones.split('-')[1]
      
      for materia in self.materias:
        if int(materia["codigo"]) == int(codigo):
          materia["profesor"].append((cedula, secciones))

  def insertar_materias(self):
    self.materias = list()

    try:
      matExcel = pd.read_excel(dirExcel, sheet_name="Materias", skiprows=[0])
      matExcel = matExcel.drop(columns=['Unnamed: 0'])
      profExcel = pd.read_excel(dirExcel, sheet_name="Profesores", skiprows=[0])
      profExcel = profExcel.drop(columns=['Unnamed: 0'])

      for indice, fila in matExcel.iterrows():
        codigo = fila['Código']
        nombre = fila['Nombre']
        profesor = ("",0)
        prioridad = fila['Prioridad']
        secciones = fila['Secciones']
        carrera = fila['Código carrera']
        modalidad = fila['Modalidad']
        semestre = fila['Semestre']
        uc = fila['UC']
        self.materias.append({"codigo": codigo, "nombre": nombre, "profesor": [], "prioridad": prioridad, "secciones": secciones,
							"carrera": carrera,"modalidad": modalidad, "semestre": semestre, "UC": uc})

    except Exception as excepcion:
      print("Ocurrió un error al leer el archivo excel.")
      print(f"Ocurrió una Excepción de tipo {excepcion}")
      exit()

    return self.materias

  def insertar_aulas(self):
    self.aulas = list()

    try:
      aulasExcel = pd.read_excel(dirExcel, sheet_name="Aulas", skiprows=[0])
      aulasExcel = aulasExcel.drop(columns=['Unnamed: 0'])

      for indice, fila in aulasExcel.iterrows():
        id_aula = fila['ID aula']
        codigo = fila['Código']
        sede = fila['Sede']
        modulo = fila['Módulo']
        capacidad = fila['Capacidad']
        tipo_aula = fila['Tipo aula']
        self.aulas.append(dict({"id_aula": id_aula, "codigo": codigo, "sede": sede, "modulo": modulo, "capacidad": capacidad, "tipo_aula": tipo_aula}))


    except Exception as excepcion:
      print("Ocurrió un error al leer el archivo excel.")
      print(f"Ocurrió una Excepción de tipo {excepcion}")
      exit()

    return self.aulas

  def insertar_bloque(self):
    self.bloques = list()

    try:
      bloquesExcel = pd.read_excel(dirExcel, sheet_name="Bloques")

      for indice,fila in bloquesExcel.iterrows():
        id = fila['ID']
        hora_inicio = str(fila['Hora inicio '])
        hora_fin = str(fila['Hora fin '])
        prioridad = fila['Prioridad']
        duracion = str(fila['Duración'])
        self.bloques.append(dict({"id": id, "hora_inicio": hora_inicio, "hora_fin": hora_fin, "prioridad": prioridad, "duracion": duracion}))

    except Exception as excepcion:
      print("Ocurrió un error al leer el archivo excel.")
      print(f"Ocurrió una Excepción de tipo {excepcion}")
      exit()
    return self.bloques
  
  def insertar_carreras(self):
    self.carreras = list()

    try:
      carrerasExcel = pd.read_excel(dirExcel, sheet_name="Carreras", skiprows=[0])
      carrerasExcel = carrerasExcel.drop(columns=['Unnamed: 0'])

      for indice,fila in carrerasExcel.iterrows():
        codigo = fila['Código']
        nombre = fila['Nombre']
        sede = fila['Sede']
        self.carreras.append(dict({"codigo": codigo, "nombre": nombre, "sede": sede}))

    except Exception as excepcion:
      print("Ocurrió un error al leer el archivo excel.")
      print(f"Ocurrió una Excepción de tipo {excepcion}")
      exit()
    return self.carreras