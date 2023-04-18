from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime

layout = [[sg.Text('ID Мастера'), sg.Push(), sg.Input(key='master')],
          [sg.Text('ID Услуги'), sg.Push(), sg.Input(key='product')],
          [sg.Text('ID Клиента'), sg.Push(), sg.Input(key='client')],
          [sg.Text('Время'), sg.Push(), sg.Input(key='time')],
          [sg.Text('Цена'), sg.Push(), sg.Input(key='price')],
          [sg.Button('Добавить'), sg.Button('Закрыть')]]

window = sg.Window('База данных салона красоты', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Закрыть':
        break
    if event == 'Добавить':
        try:
            wb = load_workbook('lr.xlsx')
            sheet = wb['Лист1']
            sheet['A1']="№"
            sheet['B1']="ID Мастера"
            sheet['C1']="ID Услуги"
            sheet['D1']="ID Клиента"
            sheet['E1']="Время записи на услугу"
            sheet['F1']="Цена"
            sheet['G1']="Время добавления в БД"
            ID = len(sheet['ID']) + 1
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            data = [ID, values['master'], values['product'], values['client'], values['time'], values['price'], time_stamp]

            sheet.append(data)

            wb.save('lr.xlsx')

            window['master'].update(value='')
            window['product'].update(value='')
            window['client'].update(value='')
            window['time'].update(value='')
            window['price'].update(value='')
            window['master'].set_focus()

            sg.popup('Данные сохранены')
        except PermissionError:
            sg.popup('File in use', 'File is being used by another User.\nPlease try again later.')

window.close()