# pip install openpyxl
import openpyxl as op

# leer el archivo
book = op.load_workbook('plantilla.xlsx', data_only=True)
# fijar la hoja
hoja = book.active

celdas = hoja['A2' : 'D5']

for fila in celdas:
    
    for celda in fila:
        print(celda.value)
#print('*'*70)    
lista_empleados = []
for fila in celdas:
    """
    for celda in fila:
        print(celda.value)
    """
    #print([celda.value for celda in fila])
    
    empleado = [celda.value for celda in fila]
    lista_empleados.append(empleado)

print('*'*70)
#print(lista_empleados)
for empleado in lista_empleados:
    print(empleado)

print('*'*70)

for empleado in lista_empleados:
    print(f'El empleado {empleado[0]} es un {empleado[1]} y gana ${empleado[2]}')

print('*'*70)