from main import *
import os
import PySimpleGUI as sg
### Imports ####


Layout = [ [sg.Text('Hello From PDF converter, Please Provide the .docx to convert to .PDF')] ,
           [sg.Input(key='file_path'), sg.FileBrowse()],
           [sg.Checkbox('Open the converted file in the default PDF viewer', key='open', default=False)],
           [sg.OK(), sg.Cancel()]]

window = sg.Window('PDF Converter', layout=Layout)

while True:
    event, values = window.read()
    if event in (None, 'Cancel'):
        break
    elif event in ('OK'):
        make_pdf_(values['file_path'], open_=values['open'])
        sg.popup('pdf made !')
        break
    else:
        sg.popup('An Error Occoured')
        break