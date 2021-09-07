import PySimpleGUI as sg
from PySimpleGUI.PySimpleGUI import T, InputText
from Equipos_hr_check import save
from Equipos_hr_check import tv_planta_baja
from Equipos_hr_check import tv_primer_piso
from Equipos_hr_check import totem
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
        [sg.Button('Aceptar')]
    ]
    window2 = sg.Window('test', layout, modal=True)
    return window2

def win_TvP1():
    layout = [
        [sg.Text('Ingrese estado del TV del primer piso: '), sg.InputText()],
        [sg.Text('Que solución se le dió?: '), sg.InputText()],
        [sg.Button('Aceptar')]
    ]
    window2 = sg.Window('test', layout, modal=True)
    return window2

def win_Totem():
    layout = [
        [sg.Text('Ingrese estado del Totem de PB: '), sg.InputText()],
        [sg.Text('Que solución se le dió?: '), sg.InputText()],
        [sg.Button('Aceptar')]
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
    #Si no se quiere hacer cambios se guarda el archivo.
    if event == 'No':
        save()
        cancel = True
    #Si se quiere hacer cambios se cambia la ventana del menú.
    elif event == 'Si':
        finish = False
        while not finish:
            #
            window_2 = win_TvPB()
            event, values = window_2.read()  
            if event == 'Aceptar':
                hoja['b4'].value = values[0]
                hoja['c4'].value = values [1]
                window_2.close()

            window_2 = win_TvP1()
            event, values = window_2.read()
            if event == 'Aceptar':
                hoja['b5'].value = values[0]
                hoja['c5'].value = values[1]
                window_2.close()
            
            window_2 = win_Totem()
            event, values = window_2.read()
            if event == 'Aceptar':
                hoja['b6'].value = values[0]
                hoja['c6'].value = values[1]
                window_2.close()
            finish = True
        window.close()
        sg.popup('Finalizado, muchas gracias')
                


            
            
save()        


# window.close()