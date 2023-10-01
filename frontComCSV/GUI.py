import PySimpleGUI as sg
import csv

sg.theme('DarkBrown2')
headings = ['N° do Cliente', 'Nome', 'Telefone', 'Endereço']
header = [
        sg.Text('N° do Cliente', pad=(0, 0), size=(15,1), justification='c'),
        sg.Text('Nome', pad=(0, 0), size=(10,1), justification='c'),
        sg.Text('Telefone', pad=(0, 0), size=(83,1), justification='c'),
        sg.Text('Endereço', pad=(0, 0), size=(-500,1), justification='c')
    ]

layout = [header]


for row in range(0, 31):
    layout.append([
        sg.Input(size=(15, 1), pad=(0, 0), key=(row, 0)), 
        sg.Input(size=(46,1), pad=(0, 0), key=(row, 1)), 
        sg.Input(size=(15,1), pad=(0, 0), key=(row, 2)),
        sg.Input(size=(58,1), pad=(0, 0), key=(row, 3))
    ])

layout.append([sg.Button("Teste"), sg.Button("Gera CSV"), sg.Button("Limpa Tabela")])

window = sg.Window('Lab Software', layout, font='Courier 12').finalize()
window.Maximize()

def generate_csv(headings, values):
    headings = ['N° do Cliente', 'Nome', 'Telefone', 'Endereço']

    file = open('arquivo.csv', 'w', encoding='UTF8', newline='')
    writer = csv.writer(file)

    # write the header
    writer.writerow(headings)

    for row in range(31):
        current_row = []
        for column in range(4):
            current_row.append(values[row, column])
        writer.writerow(current_row)

    file.close()

def clear_all(window):
    for row in range(15):
        for column in range(4):
            window[(row, column)].update('')


while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    elif event == 'Teste':
        print("Funciona ")
    elif event == 'Gera CSV':
        generate_csv(headings, values)
    elif event == 'Limpa Tabela':
        values = []
        clear_all(window)
