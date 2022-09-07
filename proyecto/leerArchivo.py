import openpyxl

nc = 5

libro = openpyxl.load_workbook('formulario.xlsx' ,data_only=True)

hoja = libro.active

##[fila] [columnas
##las filas no inician de 0 si no de 1

nf = 1
bandera = True
while(bandera):
    contador = 0
    for celda in range(nc):
        if(hoja[nf][celda].value == None):
            contador+=1    
    if(contador == nc):
        bandera = False
    nf+=1
nf-=2


position = 'E'
position += str(nf)

celdas = hoja['A2' : position]

matriz =[]

for fila in celdas:
    persona = [celda.value for celda in fila]
    matriz.append(persona)
    
#for fila in range(nf-1):
 #   for col in range(nc):
  #      print(matriz[fila][col])
   # print("")


def similitudes(palabra1, palabra2):
    conter = 0
    pal1 = ""
    for i in palabra1:
        if(pal1 == "" and i == ' '):
            continue
        if(i != ','):
            pal1 += i
        else:
            pal2 = ""
            for j in palabra2:
                if(pal2 == "" and j == ' '):
                    continue
                if(j != ','):
                    pal2 += j
                else:
                    if(pal1 == pal2):
                        conter +=1
                    pal2 = ""
            if(pal2 != "" and pal1 == pal2):
                conter +=1
            pal1 = ""
    if(pal1 != ""):
        pal2 = ""
        for j in palabra2:
            if(pal2 == "" and j == ' '):
                continue
            if(j != ','):
                pal2 += j
            else:
                if(pal1 == pal2):
                    conter +=1
                pal2 = ""
        if(pal2 != "" and pal1 == pal2):
            conter +=1
        pal1 = ""
    return conter

relaciones = []
for i in range(nf-1):
    numeros = [0 for j in range(nf-1)]
    relaciones.append(numeros)

for colExcel in range(1,4):
    for fila in range(nf-1):
        for col in range(nf-1):
            if(fila != col):
                if(matriz[fila][colExcel] == None or matriz[col][colExcel] == None):
                    numero = 0
                else:
                    numero = similitudes(str(matriz[fila][colExcel]), str(matriz[col][colExcel]))
                relaciones[fila][col] += numero
            else:
                relaciones[fila][col] = 0

for person in range(nf-1):
    print(person+1, ". ", matriz[person][0], sep='')
print()

for i in range(nf+1):
     print(("───").center(3), end= '')
print()


print("   ║", end='')
for i in range(nf-1):
    if(i<10):
        print((' ' + str(i+1)).center(3), end= '')
    else:
        print(str(i+1).center(3), end= '')
print()
        
for i in range(nf+1):
     print(("───").center(3), end= '')
print()

for fila in range(nf-1):
    print(str(fila+1).center(2),"║",end='')
    for col in range(nf-1):
        if(relaciones[fila-1][col]<10):
            print((' ' + str(relaciones[fila][col])).center(3), end= '')
        else:
            print(str(relaciones[fila-1][col]).center(3), end= '')
    print(" ║")

for i in range(nf+1):
     print(("───").center(3), end= '')
print()
 
print("fin")

