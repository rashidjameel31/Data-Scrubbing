import PySimpleGUI as sg
layout1 = [  [sg.Text('Program Complete')],
            [sg.Button('Ok')] ]
window1 = sg.Window('Automatic Scripts', layout1)
newevent,newvalues = window1.read()
window1.close()
if newevent == "Ok":
    print("Done")