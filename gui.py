import PySimpleGUI as sg
from PySimpleGUI.PySimpleGUI import T, InputText
import openpyxl
from openpyxl import workbook

excelDoc = openpyxl.load_workbook('Checklist Equipos RRHH.xlsx')
hoja = excelDoc["Hoja1"]

def save():
    excelDoc.save('Checklist Equipos RRHH.xlsx')

sg.theme('DarkBlue')  #Tema de color


# Botones dentro de la ventana principal

layout = [ [sg.Text('Este es un programa con interfaz visual para modificar el excel de chequeo de equipos.') ],
                        [sg.Text('Ingrese su nombre completo'), sg.InputText()],
                        [sg.Text('Desea hacer cambios en el excel?')],
                        [sg.Button('Si'),  sg.Button('No')],
                        [sg.Button('Cancelar')] ]
            
# Función de la ventana 2, con botones y campos a completar.
def window2(ubicacion):
    layout = [
        [sg.Text(f'Ingrese estado de {ubicacion} : '), sg.InputText()],
        [sg.Text('Que solución se le dió?: '), sg.InputText()],
        [sg.Button('Aceptar'), sg.Button('Cancelar')]
    ]
    window2 = sg.Window('test', layout, modal=True)
    return window2

window = sg.Window('Modificar Excel', layout) #Creamos la ventana y ponemos el titulo

#Si se clickea en cancelar o se cierra la ventana, cancel pasa a ser True y se termina el blucle
close = False
while not close:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancelar':
        close = True
    #Si no se quiere hacer cambios se guarda el archivo con la fecha y hora actualizadas.
    if event == 'No':
        hoja['B17'] = values[0]
        save()
        close = True
    #Si se quiere hacer cambios se cambia la ventana del menú.
    elif event == 'Si':
        hoja['B17'] = values[0]
        finish = False
        while not finish:
            #La segunda ventana toma el valor de la función window2, el parametro toma como valor la ubicación del equipo.
            #Luego de eso, la ventana se cierra.

            window_2 = window2("Planta baja")
            event, values = window_2.read()  
            if event == 'Aceptar':
                hoja['b4'].value = values[0]       #El valor del primer input de la ventana se carga en la celda correspondiente de excel
                hoja['c4'].value = values [1]      #El valor del primer input de la ventana se carga en la celda correspondiente de excel
                window_2.close()
            elif event == "Cancelar":    #Si el usuario da click en cancelar, no se guarda ningun cambio en el excel
                window_2.close()


            #Cuando se da aceptar o cancelar en la ventana anterior, se abre la siguiente, se llama a la función window2 con la ubicación correspondiente
            #Funciona de la misma forma que la anterior.

            window_2 = window2("Primer piso")
            event, values = window_2.read()
            if event == 'Aceptar':
                hoja['b5'].value = values[0]     #El valor del primer input de la ventana se carga en la celda correspondiente de excel
                hoja['c5'].value = values[1]     #El valor del primer input de la ventana se carga en la celda correspondiente de excel
                window_2.close()
            elif event == "Cancelar":    #Si el usuario da click en cancelar, no se guarda ningun cambio en el excel
                window_2.close()

            #La tercera ventana tiene el mismo funcionamiento que las dos anteriores.
            
            window_2 = window2("Totem")
            event, values = window_2.read()
            if event == 'Aceptar':
                hoja['b6'].value = values[0]     #El valor del primer input de la ventana se carga en la celda correspondiente de excel
                hoja['c6'].value = values[1]     #El valor del primer input de la ventana se carga en la celda correspondiente de excel
                window_2.close()
            elif event == "Cancelar":     #Si el usuario da click en cancelar, no se guarda ningun cambio en el excel
                window_2.close()

            finish = True
        window.close()

        sg.popup('Finalizado, muchas gracias')
                


save()    #Se guarda el archivo
