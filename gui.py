import PySimpleGUI as sg
from PySimpleGUI.PySimpleGUI import T, InputText
from Equipos_hr_check import save
from Equipos_hr_check import hoja

sg.theme('DarkBlue')  #Tema de color
# Botones dentro de la ventana


layout = [ [sg.Text('Este es un programa con interfaz visual para modificar el excel de chequeo de equipos.') ],
                        [sg.Text('Desea hacer cambios en el excel?')],
                        [sg.Button('Si'),  sg.Button('No')],
                        [sg.Button('Cancelar')] ]
            
def win_TvPB():
    layout = [
        [sg.Text('Ingrese estado de TV Planta baja: '), sg.InputText()],
        [sg.Text('Que solución se le dió?: '), sg.InputText()],
        [sg.Button('Aceptar'), sg.Button('Cancelar')]
    ]
    window2 = sg.Window('test', layout, modal=True)
    return window2

def win_TvP1():
    layout = [
        [sg.Text('Ingrese estado del TV del primer piso: '), sg.InputText()],
        [sg.Text('Que solución se le dió?: '), sg.InputText()],
        [sg.Button('Aceptar'), sg.Button('Cancelar')]
    ]
    window2 = sg.Window('test', layout, modal=True)
    return window2

def win_Totem():
    layout = [
        [sg.Text('Ingrese estado del Totem de PB: '), sg.InputText()],
        [sg.Text('Que solución se le dió?: '), sg.InputText()],
        [sg.Button('Aceptar'), sg.Button('Cancelar')]
    ]
    window2 = sg.Window('test', layout, modal=True)
    return window2 

window = sg.Window('Modificar Excel', layout) #Creamos la ventana y ponemos el titulo

#Si se clickea en cancelar o se cierra la ventana, cancel pasa a ser True y se termina el blucle
cancel = False
while not cancel:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancelar':
        cancel = True
    #Si no se quiere hacer cambios se guarda el archivo con la fecha y hora actualizadas.
    if event == 'No':
        save()
        cancel = True
    #Si se quiere hacer cambios se cambia la ventana del menú.
    elif event == 'Si':
        finish = False
        while not finish:
            #La segunda ventana toma el valor de la función win_TvPB, mostrando en la misma los valores cargandos de la función.
            #Luego de eso, la ventana se cierra.

            window_2 = win_TvPB()
            event, values = window_2.read()  
            if event == 'Aceptar':
                hoja['b4'].value = values[0]       #El valor del primer input de la ventana se carga en la celda correspondiente de excel
                hoja['c4'].value = values [1]      #El valor del primer input de la ventana se carga en la celda correspondiente de excel
                window_2.close()
            elif event == "Cancelar":    #Si el usuario da click en cancelar, no se guarda ningun cambio en el excel
                window_2.close()


            #Cuando se da aceptar o cancelar en la ventana anterior, se abre la siguiente, que toma el valor de la función win_TvP1.
            #Funciona de la misma forma que la anterior.

            window_2 = win_TvP1()
            event, values = window_2.read()
            if event == 'Aceptar':
                hoja['b5'].value = values[0]     #El valor del primer input de la ventana se carga en la celda correspondiente de excel
                hoja['c5'].value = values[1]     #El valor del primer input de la ventana se carga en la celda correspondiente de excel
                window_2.close()
            elif event == "Cancelar":    #Si el usuario da click en cancelar, no se guarda ningun cambio en el excel
                window_2.close()
            
            window_2 = win_Totem()
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
                


            
            
save()        


# window.close()