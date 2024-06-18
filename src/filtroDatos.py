import pandas as pd

#A continuación con estos parametros que se incluyen en la librería pandas, le especificamos
# a la librería que nos muestre una cantidad de diez columnas, sí es que lo contiene el DataFrame
# y una visualización en la consola de un ancho de 400px para que muestre bastante información,
# de estar en este. De lo contrario pandas mostraría información abreviada, limitándonos la cantidad
# de información de carácter relevante, que deseamos visualizar.
pd.set_option('display.width', 400)
pd.set_option('display.max_columns', 10)

# Variable str que contiene la dirección del excel.
dirExcel= "../data/DATOSHORARIOS.xlsx"

#profesores = None
materias = None
CARRERAS = None
nombreCarreras = None
datosCarreras = list()

# El objetivo de la variable 'nombreCarreras' es que esta recibiese
# directamente los valores de las carreras del dataframe, pero al no
# hallar todavía el método, he escogido cargar directamente la variable

try:
	# Cargamos los datos y especificamos que omita la primera fila vacía,
	# ya que si leemos incluyéndola debido a que están vacías no podrá identificar el nombre de
	# las columnas para el dataframe y por ende lanzara una excepción

	materias = pd.read_excel(dirExcel, sheet_name="Materias",skiprows=[0])
	CARRERAS = pd.read_excel(dirExcel, sheet_name="Carreras",skiprows=[0])
	#profesores = pd.read_excel(dirExcel, sheet_name="Profesores", usecols=[1,2,3], skiprows=[0])
	print("El fichero excel se ha leído exitosamente", '\n')

except Exception as excepcion:
	print("Ocurrió un error al leer el archivo excel.")
	print(f"Ocurrió una Excepción de tipo {excepcion}")
	exit()

# Cargamos los nombres de las carreras del dataframe 'Carreras' en la variable 'nombreCarreras';
# donde la variable 'nombreCarreras' es asignada como una lista, conteniendo los nombres de las carreras.
nombreCarreras = CARRERAS['Nombre'].values.tolist()

# Cargamos los diversos datos de las carreras de acuerdo a la cantidad que se halle menos uno,
# con el fin, que recorra los índices de la lista 'DatosCarreras' y no recorra un índice fuera de rango.
# Con el método .loc de la estructura DataFrame filtramos la las materias que coincidan con la carrera.
for i in range (len(CARRERAS.columns)-1):
	datosCarreras.append(materias[['Nombre','Carrera','Semestre','Modalidad','UC']].loc[materias['Carrera'] == nombreCarreras[i]])
	#datosCarreras[i].to_csv(r'../../../infoMatCarreras.txt', header=None, index=None, sep=' ', mode='a')
	print(nombreCarreras[i])
	print(datosCarreras[i])
	print('\n')

#Vaciamos la información que no necesitamos.
materias = Carreras = Profesores = None

'''
Paginas clave consultadas:
https://www.geeksforgeeks.org/ways-to-filter-pandas-dataframe-by-column-values/
https://saturncloud.io/blog/how-to-extract-value-from-a-dataframe-a-comprehensive-guide-for-data-scientists/
'''