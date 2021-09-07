import openpyxl
from openpyxl.workbook import workbook

excelDoc = openpyxl.load_workbook('Checklist Equipos RRHH.xlsx')
hoja = excelDoc.get_sheet_by_name('Hoja1')

def  tv_planta_baja():
        hoja['b4'].value = input("Indique el estado de la TV de planta baja [Encedido - Apagado - Faltante]: ")
        tv_pb = hoja['b4'].value
        if  tv_pb == "apagado":
            hoja['c4'].value = input("Que solución se le dió?: ")
        elif tv_pb == "faltante":
            hoja['c4'].value = input("Que solución se le dió?: ")
        elif tv_pb == "encendido":
            hoja['c4'].value = " "

def tv_primer_piso():
    hoja['b5'].value = input("Indique el estado de la TV del primer piso [Encedido - Apagado - Faltante]: ")
    tv_piso1 = hoja['b5'].value
    if   tv_piso1 == "apagado":
            hoja['c5'].value = input("Que solución se le dió?: ")
    elif  tv_piso1 == "faltante":
            hoja['c5'].value = input("Que solución se le dió?: ")
    elif  tv_piso1 == "encendido":
            hoja['c5'].value = " "

def totem():
    hoja['b6'].value = input("Indique el estado del totem de planta baja [Encedido - Apagado - Faltante]: ")
    totem_pb = hoja['b6'].value
    if totem_pb == "encendido":
        hoja['c6'].value = " "
    elif totem_pb == "apagado":
        hoja['c6'].value = input("Que solución se le dió?: ") 
    elif totem_pb == "faltante":
        hoja['c6'].value =  input("Que solución se le dió?: ")



# modificar = input("Desea hacer cambios?[Si - No]: ")

# if modificar == "si":
#         tv_planta_baja()
#         tv_primer_piso()
#         totem()

def save():
    excelDoc.save('Checklist Equipos RRHH.xlsx')

save()