#Quiero un programa que pueda leer el excel de la facultad y decirme qué horario y en qué aula hay x clase
#Adicionalmente, podré ver las aulas vacías en un horario especificado

#TODO: Abrir el archivo excel al empezar el programa


while (True):
  print("Bienvenido al programa de aulas!")
  print("Qué desea hacer?")
  print("1-Ver horario y aula.\n2-Ver aulas vacías.\n3-Salir")
  accion = int(input())
  if(accion < 1 or accion >3):
    print("Esa no es una opción válida.")

  if(accion == 3):
    break

  if(accion == 1):
    #encontrar_aulas()
    break

  if(accion == 2):
    #encontrar_vacias()
    break
    
