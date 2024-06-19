import pandas as pd

# Se emplea la librería pandas para la estructura de datos Dataframe
# como también de la lectura del documento excel que contiene los datos
# de interés para el presente script.

#A continuación con estos parametros que se incluyen en la librería pandas, le especificamos
# a la librería que nos muestre una cantidad de diez columnas, sí es que lo contiene el DataFrame
# y una visualización en la consola de un ancho de 400px para que muestre bastante información,
# de estar en este. De lo contrario pandas mostraría información abreviada, limitándonos la cantidad
# de información de carácter relevante, que deseamos visualizar.
pd.set_option('display.width', 500)
pd.set_option('display.max_columns', 12)

# Variable str que contiene la dirección del excel.
dirExcel= "../data/DATOSHORARIOS.xlsx"

#Inicialización de listas y variables en valor None o vacío.
profesores = None
materias = None
aulas = None
bloqueP_Alto = None
bloqueP_MedBajo = None
listaDesplegable = None
restricciones = None
carreras = None
nombreCarreras = list()
datosCarreras = list()

# El objetivo de la variable 'nombreCarreras' es que esta recibiese
# directamente los valores de las carreras del dataframe, pero al no
# hallar todavía el método, he escogido cargar directamente la variable

try:
	# Cargamos los datos y especificamos que omita la primera fila vacía,
	# ya que si leemos incluyéndola debido a que están vacías no podrá identificar el nombre de
	# las columnas para el dataframe y por ende lanzara una excepción

	materias = pd.read_excel(dirExcel, sheet_name="Materias", skiprows=[0])
	materias = materias.drop(columns=['Unnamed: 0'])

	# Empleando el método '.drop' podremos eliminar la primera
	# columna vacía del excel o cualquier otra, especificándole su nombre.

	CARRERAS = pd.read_excel(dirExcel, sheet_name="Carreras", skiprows=[0])
	CARRERAS = CARRERAS.drop(columns=['Unnamed: 0'])
	nombreCarreras = CARRERAS['Nombre'].values.tolist()

	# Pasamos los nombres de las carreras a una lista para determinar
	# la cantidad de estas y sus valores.

	profesores = pd.read_excel(dirExcel, sheet_name="Profesores", skiprows=[0])
	profesores = profesores.drop(columns=['Unnamed: 0'])
	print(profesores)
	print('\n')

	aulas = pd.read_excel(dirExcel, sheet_name="Aulas", skiprows=[0])
	aulas = aulas.drop(columns=['Unnamed: 0'])
	print(aulas)
	print('\n')

	bloqueP_Alto = pd.read_excel(dirExcel, sheet_name="Bloques", skiprows=[0])
	bloqueP_Alto = bloqueP_Alto.drop(columns=['Unnamed: 0'])
	bloqueP_Alto = bloqueP_Alto.loc[0:2]
	print(bloqueP_Alto)
	print('\n')
	# Con el método loc indicamos a las filas del dataframe, que solo usaremos
	# los registros de las filas que empiezan desde el indice cero hasta dos
	# y los registros de las filas superiores a dos serán descartadas.


	bloqueP_MedBajo = pd.read_excel(dirExcel, sheet_name="Bloques", header = 5 ,skiprows=[0])
	bloqueP_MedBajo = bloqueP_MedBajo.drop(columns=['Unnamed: 0'])
	bloqueP_MedBajo = bloqueP_MedBajo.loc[0:3]
	print(bloqueP_MedBajo)
	print('\n')
	# Caso parecido a las líneas de código anteriores, solo que en
	# esta ocasión se usaran los registros desde la fila cero a tres y
	# se descartaran las restantes. También en la lectura del excel,
	# nuevamente en la hoja llamada 'Bloques' se utiliza el parámetro header,
	# indicándole que empiece la lectura desde la fila 5, tomando en cuenta que se saltó una fila con
	# el parámetro skiprow. Para este tome los elementos de las celdas.

	listaDesplegable = pd.read_excel(dirExcel, sheet_name="ListaDesplegable", header=1)
	listaDesplegable = listaDesplegable.drop(columns=['Unnamed: 0','Unnamed: 3'])
	print(listaDesplegable)
	print('\n')

	restricciones = pd.read_excel(dirExcel, sheet_name="Restricciones", header = 1)
	restricciones = restricciones.drop(columns=['Unnamed: 0'])
	print(restricciones)
	print('\n')

	print("El fichero excel se ha leído exitosamente", '\n')

except Exception as excepcion:
	print("Ocurrió un error al leer el archivo excel.")
	print(f"Ocurrió una Excepción de tipo {excepcion}")
	exit()


for i in range (len(nombreCarreras)):
	datosCarreras.append(materias[['Nombre','Carrera','Semestre','Modalidad','UC']].loc[materias['Carrera'] == nombreCarreras[i]])
	print(nombreCarreras[i])
	print(datosCarreras[i])
	print('\n')
# En este bucle for lo que estamos haciendo es cargar el dataframe materias en la lista datosCarreras
# solamente con las cabeceras o celdas nombre " 'Nombre','Carrera','Semestre','Modalidad','UC' " y filtrando
# con el método loc, usando como condición, de que sí el nombre de carrera coincide con la materia
# se guarde un dataframe en la lista que coincida la materia y la carrera.

# Vaciamos la información que no necesitamos.
#materias = Carreras = Profesores = None

'''
Paginas clave consultadas:
https://www.geeksforgeeks.org/ways-to-filter-pandas-dataframe-by-column-values/
https://saturncloud.io/blog/how-to-extract-value-from-a-dataframe-a-comprehensive-guide-for-data-scientists/
https://sparkbyexamples.com/pandas/pandas-create-new-dataframe-by-selecting-specific-columns/
https://stackoverflow.com/questions/25628496/getting-wider-output-in-pycharms-built-in-console/50947606#50947606
https://www.geeksforgeeks.org/how-to-drop-rows-in-pandas-dataframe-by-index-labels/
'''