# Import Packages
import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('TealMono')

# Setup Excel Spread Sheet used by this program
pd_writer = pd.ExcelWriter('ContactDatabase.xlsx', engine='xlsxwriter')
pd_writer.save()

EXCEL_FILE = 'ContactDatabase.xlsx'
df = pd.read_excel(EXCEL_FILE)

# Layout for the GUI
layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('First Name', size=(15, 1)), sg.InputText(key='FirstName')],
    [sg.Text('Middle Initial', size=(15, 1)), sg.InputText(size=(1, 1), key='MiddleInitial')],
    [sg.Text('Last Name', size=(15, 1)), sg.InputText(size=(25, 1), key='LastName')],
    [sg.Text('Email Address', size=(15, 1)), sg.InputText(size=(30, 1), key='EmailAddress')],
    [sg.Text('Address', size=(15, 1)), sg.InputText(size=(34, 1), key='Address')],
    [sg.Text('Phone Number', size=(15, 1)), sg.InputText(size=(15, 1), key='PhoneNumber')],

    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

# Create Window
window = sg.Window('Contact Information', layout)


def clear_input():
    for key in values:
        window[key]('')
    return None


# Main Event Loop
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data Saved')
        clear_input()

window.close()
