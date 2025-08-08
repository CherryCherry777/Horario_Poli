#Quiero un programa que pueda leer el excel de la facultad y decirme qué horario y en qué aula hay x clase
#Adicionalmente, podré ver las aulas vacías en un horario especificado
#TODO servicio para guardar materias y poder consultar en cualquier momento
#TODO que busque aulas vacias en un tiempo determinado
import openpyxl
from openpyxl import Workbook
from colorama import Fore
import os
from tabulate import tabulate
print(Fore.GREEN + "Cargando..." + Fore.WHITE)
path = "horario.xlsx"
path2 = "horario_guardado.xlsx"

if (os.path.exists(path)):
  try:
    doc = openpyxl.load_workbook(path)
  except Exception as e:
    print("El horario no pudo ser cargado. Por favor verificar su existencia.")
    print("El archivo del horario debe ser de la facultad, nombrado horario.xlsx y estar en la misma carpeta que main.py")
else:
  print("El horario no pudo ser cargado. Por favor verificar su existencia.")
  print("El archivo del horario debe ser de la facultad, nombrado horario.xlsx y estar en la misma carpeta que main.py")

if (os.path.exists(path2)):
  try:
    print(Fore.GREEN + "Horario personalizado cargado exitosamente." + Fore.WHITE)
    doc2 = openpyxl.load_workbook(path2)
    hoja2 = doc2.active
    filas2 = hoja2.max_row
  except Exception as e:
    print("El horario personalizado no pudo ser cargado.")
else:
  print(Fore.GREEN + "Creando archivo de horario personalizado..." + Fore.WHITE)
  doc2 = Workbook()
  doc2.save(path2)
  hoja2 = doc2.active
  filas2 = 0






col_materia = 3 #columna de nombre de materia
col_seccion = 10 #columna de seccion de materia
col_prof_nombre = 14 #columna del nombre del docente
col_prof_apellido = 13 #columna del apellido del docente
col_prof_titulo = 12 #columna del titulo del docente
col_lunes = 36 #columna de las horas del lunes
col_martes = 38 #columna de las horas del martes
col_miercoles = 40 #columna de las horas del miercoles
col_jueves = 42 #columna de las horas del jueves
col_viernes = 44 #columna de las horas del viernes
col_sabado = 46 #columna de las horas del sabado

col_lunes_aula = 35
col_martes_aula = 37
col_miercoles_aula = 39
col_jueves_aula = 41
col_viernes_aula = 43
col_sabado_aula = 45
columnas = 47

carreras = ["IAE","ICM","IEK","IEL","IEN","IIN","IMK","ISP","LCA","LCI","LCIk","LEL","LGH","TSE","Cnel. Oviedo","Villarica"]
carreras_upper = ["IAE","ICM","IEK","IEL","IEN","IIN","IMK","ISP","LCA","LCI","LCIK","LEL","LGH","TSE","CNEL. OVIEDO","VILLARICA"]

aulas_a = ["A50","A51","A52","A53","A54","A55","A56","A57","A58","A59"]
aulas_b = ["B01","B02"]
aulas_c = ["C01","C02","C03","C04"]
aulas_e = ["E01","E02","E03","E04"]
aulas_f = ["F05","F06","F07","F08","F09","F10","F11","F12","F13","F14","F15","F16","F17","F18","F25","F29","F30","F31","F33","F34","F35","F36","F37","F38","F39","F40"]
aulas_h = ["H03","H04","H05","H06","H07","H08"]
aulas_i = ["I01","I02","I03","I04","I05","I06","I07","I08"]

#Una celda vacia se presenta como "None"

#hoja = doc[nombreHoja]

def imprimir_info_materia_2(hoja, fila_materia):
  mat_aux = [[]]
  #especificamente para imprimir UNA materia del horario
  print("Materia: ", hoja.cell(row = fila_materia, column = col_materia).value, " - ", hoja.cell(row = fila_materia, column = col_seccion).value)
  print("Profesor: ",hoja.cell(row = fila_materia, column = col_prof_titulo).value,hoja.cell(row = fila_materia, column= col_prof_nombre).value,hoja.cell(row = fila_materia, column= col_prof_apellido).value)
  print("      Lunes         -     Martes        -     Miércoles     -     Jueves        -     Viernes       -     Sábado")
  if(str(hoja.cell(row = fila_materia, column = col_lunes).value) == "None"):
    mat_aux[0].append("")
  else:
    mat_aux[0].append(hoja.cell(row=fila_materia,column = col_lunes_aula).value + " " + hoja.cell(row = fila_materia, column = col_lunes).value)

  if(str(hoja.cell(row = fila_materia, column = col_martes).value) == "None"):
    mat_aux[0].append("")
  else:
    mat_aux[0].append(hoja.cell(row=fila_materia,column = col_martes_aula).value + " " + hoja.cell(row = fila_materia, column = col_martes).value)

  if(str(hoja.cell(row = fila_materia, column = col_miercoles).value) == "None"):
    mat_aux[0].append("")
  else:
    mat_aux[0].append(hoja.cell(row=fila_materia,column = col_miercoles_aula).value + " " + hoja.cell(row = fila_materia, column = col_miercoles).value)

  if(str(hoja.cell(row = fila_materia, column = col_jueves).value) == "None"):
    mat_aux[0].append("")
  else:
    mat_aux[0].append(hoja.cell(row=fila_materia,column = col_jueves_aula).value + " " + hoja.cell(row = fila_materia, column = col_jueves).value)

  if(str(hoja.cell(row = fila_materia, column = col_viernes).value) == "None"):
    mat_aux[0].append("")
  else:
    mat_aux[0].append(hoja.cell(row=fila_materia,column = col_viernes_aula).value + " " + hoja.cell(row = fila_materia, column = col_viernes).value)

  if(str(hoja.cell(row = fila_materia, column = col_sabado).value) == "None"):
    mat_aux[0].append("")
  else:
    mat_aux[0].append(hoja.cell(row=fila_materia,column = col_sabado_aula).value + " " + hoja.cell(row = fila_materia, column = col_sabado).value)

  print(tabulate(mat_aux,headers=["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"]))



def imprimir_info_materia(hoja, fila_materia):
  #especificamente para imprimir UNA materia del horario
  print("Materia: ", hoja.cell(row = fila_materia, column = col_materia).value, " - ", hoja.cell(row = fila_materia, column = col_seccion).value)
  print("Profesor: ",hoja.cell(row = fila_materia, column = col_prof_titulo).value,hoja.cell(row = fila_materia, column= col_prof_nombre).value,hoja.cell(row = fila_materia, column= col_prof_apellido).value)
  print("      Lunes         -     Martes        -     Miércoles     -     Jueves        -     Viernes       -     Sábado")
  if(str(hoja.cell(row = fila_materia, column = col_lunes).value) == "None"):
    print(" "*20, end="")
  else:
    print("  ",hoja.cell(row=fila_materia,column = col_lunes_aula).value," ",hoja.cell(row = fila_materia, column = col_lunes).value," ",end="",sep="")

  if(str(hoja.cell(row = fila_materia, column = col_martes).value) == "None"):
    print(" "*20, end="")
  else:
    print("  ",hoja.cell(row=fila_materia,column = col_martes_aula).value," ",hoja.cell(row = fila_materia, column = col_martes).value," ",end="",sep="")

  if(str(hoja.cell(row = fila_materia, column = col_miercoles).value) == "None"):
    print(" "*20, end="")
  else:
    print("  ",hoja.cell(row=fila_materia,column = col_miercoles_aula).value," ",hoja.cell(row = fila_materia, column = col_miercoles).value," ",end="",sep="")

  if(str(hoja.cell(row = fila_materia, column = col_jueves).value) == "None"):
    print(" "*20, end="")
  else:
    print("  ",hoja.cell(row=fila_materia,column = col_jueves_aula).value," ",hoja.cell(row = fila_materia, column = col_jueves).value," ",end="",sep="")

  if(str(hoja.cell(row = fila_materia, column = col_viernes).value) == "None"):
    print(" "*20, end="")
  else:
    print("  ",hoja.cell(row=fila_materia,column = col_viernes_aula).value," ",hoja.cell(row = fila_materia, column = col_viernes).value," ",end="",sep="")

  if(str(hoja.cell(row = fila_materia, column = col_sabado).value) != "None"):
    print("  ",hoja.cell(row=fila_materia,column = col_sabado_aula).value," ",hoja.cell(row = fila_materia, column = col_sabado).value," ",end="",sep="")

def ver_horario():
  print("Lista de materias:")
  for i in range (filas2 + 1):
    print(i+1,"-",hoja2.cell(row=i+1,column=col_materia).value,"-",hoja2.cell(row=i+1,column=col_seccion).value,"- Prof:",hoja2.cell(row = i+1, column = col_prof_titulo).value,hoja2.cell(row = i+1, column= col_prof_nombre).value,hoja2.cell(row = i+1, column= col_prof_apellido).value)

  print("Materia              ")



def guardar_materia(hoja, fila_materia):
  for i in range (1, columnas + 1):
    hoja2.cell(row = filas2 + 1, column = i).value = hoja.cell(row = fila_materia, column = i).value
    
  doc2.save(path2)
  print("Materia añadida exitosamente.")

def encontrar_aulas():
  #print("[green]Cargando...")
  #doc = openpyxl.load_workbook(path)
  print("Introduzca el nombre de carrera (en siglas)")

  while(True):
    carrera = input().upper()
    if(carrera in carreras_upper):
      #esto es para asegurar que incluso si no se escribe con las mayusculas o minusculas correctas, igual funcione
      hoja = doc[carreras[carreras_upper.index(carrera)]]
      break
    else:
      print("No se pudo encontrar la carrera. Intente de nuevo")

  # El test imprime la celda F12 que siempre tiene las siglas de la carrera
  # print("Test para ver si se abrio bien: ", hoja['F12'].value)
  while(True):
    print("Introduzca el nombre de la materia (palabras claves, si no esta escrito exactamente no funciona.)")
    busqueda = input()
    resultados = [] #almacena los numeros de filas de resultados de la busqueda
    filas = hoja.max_row
    #print("Una celda vacia se presenta como: ", hoja.cell(row = filas + 1, column=1).value)

    for i in range(12, filas + 1):
      #test
      #print("Test: ",hoja.cell(row = i, column = 3).value)
      if (busqueda.lower() in str(hoja.cell(row = i, column = col_materia).value).lower()):
        resultados.append(i)

    if (len(resultados) == 0):
      print("No se encontro ninguna materia con ese nombre. Intente de nuevo.")
    else:
      break

  for i in range (len(resultados)):
    print(i+1,"-",hoja.cell(row=resultados[i],column=col_materia).value," - ",hoja.cell(row=resultados[i],column=col_seccion).value)

  print("Introduzca la clase de la cual quiere información")
  busqueda = int(input())
  seleccion = resultados[busqueda-1] #es la fila de la materia seleccionada

  imprimir_info_materia_2(hoja,seleccion)

  print("\nDesea guardar esta materia en su horario?")
  print("1- Si\n2- No")
  resp = int(input())
  if (resp == 1):
    guardar_materia(hoja, seleccion)


  return


while (True):
  print("Bienvenido al programa de aulas!")
  print("Qué desea hacer?")
  print("1-Ver horario y aula.\n2-Ver aulas vacías.\n3-Ver horario guardado\n4-Salir")
  accion = int(input())
  if(accion < 1 or accion > 4):
    print("Esa no es una opción válida.")

  if(accion == 4):
    break

  if(accion == 1):
    encontrar_aulas()
    break

  if(accion == 2):
    #encontrar_vacias()
    break

  if(accion == 3):
    #ver_horario()
    break