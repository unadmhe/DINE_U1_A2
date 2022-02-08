# Import the xlrd module
import xlrd
import getopt
import sys
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.dates as mdates
import datetime

import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator


try:
    options, remainder = getopt.getopt(
        sys.argv[1:],
        'f:h',
        ['file=',
         'help',
         ])
except getopt.GetoptError as err:
    print('ERROR:', err)
    sys.exit(1)

print('Opciones   :', options)

archivo = ""

for opt, arg in options:
    if opt in ('-f', '--file'):
        archivo = arg
    elif opt in ('-h', '--help'):
        print("uso: "+argv[0]+" -f <archivo> -h")

if archivo == "":
   print ("ERROR: Debe proporcionar un archivo de entrada ( opción -f <archivo xlsx> ). Saliendo")
   exit(1)
   
# Open the Workbook
workbook = xlrd.open_workbook(archivo)

# Open the worksheet
worksheet = workbook.sheet_by_index(0)

num_rows = worksheet.nrows
num_cells = worksheet.ncols

dates=[]
values=[]
criteria=[]

c1="k"
c2="k"


# Iterate the rows and columns
for i in range(0, num_rows):
    for j in range(0, num_cells):
        # Print the cell values with tab space
        print(worksheet.cell_value(i, j), end='\t')
                    
    print('')

# Iterate the rows and columns
for i in range(1, num_rows):
   fecha = xlrd.xldate_as_datetime( int(worksheet.cell_value(rowx=i, colx=0)) , 0)
   string_date = fecha.strftime('%Y/%m')
   dates.append( string_date )
   criteria.append( worksheet.cell_value(rowx=i, colx=1) )
   
   entrada = worksheet.cell_value(rowx=i, colx=2)
   salida = worksheet.cell_value(rowx=i, colx=3)
   
   if str(entrada) == "" and str(salida) == "":
      print("Error: fila "+str(i)+" no tiene entrada ni salida!")
      exit(1)
   if str(entrada) != "" and str(salida) != "":
      print("Error: fila "+str(i)+" tiene ambas entrada y salida!")
      exit(2)
   if str(entrada) != "":
      values.append(entrada)
   if str(salida) != "":
      values.append(salida * -1)
   
   #print ('Entrada fila '+str(i)+' es "'+str(entrada)+'"')
   #print ('Salida file '+str(i)+' es "'+str(salida)+'"')
   

for i in range(0,len(dates)):
   print(dates[i]+", "+criteria[i]+", "+str(values[i])+"\n")


x = np.array(dates)
y = np.array(values)

c = np.array(criteria)


fig, ax = plt.subplots(figsize=(20, 10), constrained_layout=True)

ax.set(title="Diagrama de flujo de efectivo")

plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

ax.set_ylabel('Flujo de efectivo')
ax.tick_params('y', colors=c2)


markerline, stemlines, baseline =  ax.stem(x, y,linefmt='-',markerfmt='D')

for l, v, t in zip(x, y, c):
    ax.annotate(v, xy=(l, v), xytext=(-3, 3),
                textcoords="offset points", ha="right", size=5)

                
fig.savefig(archivo+".png", dpi = 600)

plt.show()