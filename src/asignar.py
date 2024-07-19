import pandas as pd

dirExcel = "./DATOS HORARIOS.xlsm"

def weekdays_to_int_array(days_string):
    days_string = days_string.replace(" ", "")
    days_list = days_string.split(",")
    int_array = []

    for day in days_list:
        day = day.strip()
        if day == "Lunes":
            int_array.append(0)
        elif day == "Martes":
            int_array.append(1)
        elif day == "Miercoles":
            int_array.append(2)
        elif day == "Jueves":
            int_array.append(3)
        elif day == "Viernes":
            int_array.append(4)
        elif day == "Sabado":
            int_array.append(5)
        elif day == "Domingo":
            int_array.append(6)
            
    return int_array

class Recuperar_Datos:
  def __init__(self):
    self.profesores = list()
    self.materias = list()
    self.aulas = list()
    self.carreras = list()
    self.bloques = list()
    self.globalVar = dict()

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
      tiempo = fila['Horario']
      dias_disp = weekdays_to_int_array(fila['Días disponibles (Prioridad)'])
      dias_disp_aux = weekdays_to_int_array(fila['Días disponibles (Auxiliar)'])
      dias_disp.extend(dias_disp_aux)
  
      self.asignarTuplaProf(materias, cedulaP)
      self.profesores.append(dict({"cedula": cedulaP, "dias_disp": dias_disp, "tiempo": tiempo, "nombre": nombreP, "materias": materias, "horario": { "lunes": [None] * 9, "martes": [None] * 9, "miercoles": [None] * 9, "jueves": [None] * 9, "viernes": [None] * 9 }}))
   
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
        # eliminar los espacios en blanco de la modalidad
        modalidad = fila['Modalidad'].replace(" ", "")
        semestre = fila['Semestre']
        codigo_Aula = None
        if fila['código aula '] == "NULL" or pd.isna(fila['código aula ']) is True:
          codigo_Aula = None
        else:
          codigo_Aula = fila['código aula ']
          
        uc = fila['UC']
        self.materias.append({"codigo": codigo, "nombre": nombre, "profesor": [], "prioridad": prioridad, "secciones": secciones,
							"carrera": carrera,"modalidad": modalidad, "semestre": semestre, "UC": uc, "aula": codigo_Aula})

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
        self.aulas.append(dict({"id_aula": id_aula, "codigo": codigo, "sede": sede, "modulo": modulo, "capacidad": capacidad, "tipo_aula": tipo_aula, "horario": { "lunes": [None] * 9, "martes": [None] * 9, "miercoles": [None] * 9, "jueves": [None] * 9, "viernes": [None] * 9 }}))


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
  
  def insertar_globarVar(self):
    self.globalVar = dict()

    try:
      carrerasExcel = pd.read_excel(dirExcel, sheet_name="Restricciones")

      for indice, fila in carrerasExcel.iterrows():
        restriccion = fila['Restricción']
        if pd.isna(fila['Valor']) is False:
          nombre = int(fila['Valor'])
          self.globalVar[restriccion] = nombre

    except Exception as excepcion:
      print("Ocurrió un error al leer el archivo excel.")
      print(excepcion)
      print(f"Ocurrió una Excepción de tipo {excepcion}")
      exit()
    return self.globalVar