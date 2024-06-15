import pandas as pd

# Inicialización de variables.
dirExcel= "../data/DATOSHORARIOS.xlsx"

#profesores = None
materias = None
CARRERAS = None
datosCarreras = list()
nombreCarreras = list(["Ing. Informática","Contaduría pública","Gestión de Alojamiento Turístico"])
# El objetivo de la variable 'nombreCarreras' es que esta recibiese
# directamente los valores de las carreras del dataframe, pero al no
# hallar todavía el método, he escogido cargar directamente la variable

try:
	#Cargamos los datos y especificamos que usaremos las columnas del excel 'B, C, D' y omita la fila vacía 1
	# Nota: La aplicación crashea al quitar la limitación de columnas y filas.
	# Es decir los parametros 'usecols=[1,2,3],skiprows=[0]'

	materias = pd.read_excel(dirExcel, sheet_name="Materias", dtype=str)
	CARRERAS = pd.read_excel(dirExcel, sheet_name="Carreras", dtype=str)
	#profesores = pd.read_excel(dirExcel, sheet_name="Profesores", usecols=[1,2,3], skiprows=[0])
	print("El fichero excel se ha leído exitosamente", '\n')

except Exception as excepcion:
	print("Ocurrió un error al leer el archivo excel.")
	print(f"Ocurrió una Excepción de tipo {excepcion}")
	exit()

#Cargamos los diversos datos de las carreras de acuerdo a la cantidad que se halle.
#Con el metodo.loc especificamos el valor

for i in range (len(CARRERAS.columns)):
	datosCarreras.append(
	pd.DataFrame(materias,columns=['Nombre','Carrera']).loc[materias['Carrera'] == nombreCarreras[i]])
	print(datosCarreras[i])
	print('\n')

#Vaciamos la información que no necesitamos.
#materias = Carreras = Profesores = None

'''
Paginas clave consultadas:
https://www.geeksforgeeks.org/ways-to-filter-pandas-dataframe-by-column-values/
https://saturncloud.io/blog/how-to-extract-value-from-a-dataframe-a-comprehensive-guide-for-data-scientists/
'''

'''
Código Inicial Descartado:

materiasIngInf = materias['Nombre '].loc[materias['Carrera'] == 'Ing. Informática']

#Creamos un dataframe de las materias
materiasIngInf = pd.DataFrame(materias,columns=['Nombre ','Carrera']).loc[materias['Carrera'] == 'Ing. Informática']
materiasTurismo = pd.DataFrame(materias,columns=['Nombre ','Carrera']).loc[materias['Carrera'] == 'Gestión de Alojamiento Turístico']
materiasContaduria = pd.DataFrame(materias,columns=['Nombre ','Carrera']).loc[materias['Carrera'] == 'Contaduría pública']

print("Materias de la Carrera 'Ingeniería Informática'",'\n')
print(materiasIngInf)
print('\n')

print("Materias de la Carrera 'Contaduría Pública'",'\n')
print(materiasContaduria)
print('\n')

print('Materias de la Carrera \'Gestión de Alojamiento Turístico\'','\n')
print(materiasTurismo)
'''