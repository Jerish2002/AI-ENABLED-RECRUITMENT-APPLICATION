import pandas as pd
from PySimpleGUI import Window
from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime

sg.theme('Dark Amber')
layout=[[sg.Text('Name'),sg.Push(), sg.Input(key='NAME')],
        [sg.Text('Address'),sg.Push(), sg.Input(key='ADDRESS')],
        [sg.Text('Email ID'),sg.Push(), sg.Input(key='EMAIL ID')],
        [sg.Text('Phone number'),sg.Push(), sg.Input(key='PHONE NUMBER')],
        [sg.Text('Skills'),sg.Push(), sg.Input(key='SKILLS')],
        [sg.Text('Date of Birth'),sg.Push(), sg.Input(key='DOB')],
        [sg.Text('LOCATION'),sg.Push(), sg.Input(key='LOCATION')],
        [sg.Text('Current company'),sg.Push(), sg.Input(key='CURRENT COMPANY')],
        [sg.Text('Years of experience'),sg.Push(), sg.Input(key='YEARS OF EXPERIENCE')],
        [sg.Button('Submit'), sg.Button('Cancel')]]


window = sg.Window('Data Entry', layout, element_justification='center')

while True:
    event, values = Window.read(self = window)
    if event == sg.WIN_CLOSED or event == 'Close':
        break
    if event == 'Submit':
        try:
            wb = load_workbook('Candidate_Master.xlsx')
            sheet = wb['Sheet1']
            ID = len(sheet['ID']) + 1
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            data = [ID, values['NAME'], values['ADDRESS'], values['EMAIL ID'], values['PHONE NUMBER'], values['SKILLS'], values['DOB'], values['LOCATION'], values['CURRENT COMPANY'], values['YEARS OF EXPERIENCE'], time_stamp]

            sheet.append(data)
            wb.save('Candidate_Master.xlsx')

            window['NAME'].update(value='')
            window['ADDRESS'].update(value='')
            window['EMAIL ID'].update(value='')
            window['PHONE NUMBER'].update(value='')
            window['SKILLS'].update(value='')
            window['DOB'].update(value='')
            window['LOCATION'].update(value='')
            window['CURRENT COMPANY'].update(value='')
            window['YEARS OF EXPERIENCE'].set_focus()

            sg.popup('Success', 'Data Saved')
        except PermissionError:
            sg.popup('File in use', 'File is being used by another User. \n Please try again later.')

window.close()

