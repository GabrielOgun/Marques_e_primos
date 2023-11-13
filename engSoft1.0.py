import PySimpleGUI as sg
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

json_content = {
  "type": "service_account",
  "project_id": "projetolab32066724",
  "private_key_id": "60a10bb7f6f8da633ef46ec24c2d2c12aa594608",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDJQ2YdwVJHyvfP\n0KNti28Nz0RdhY94MLKh4fvhP6vcc0fLlFJjS3IVShFxrGzScULCQmX5tpAYoU5B\n/Vic0GOuJQ0upnw/gZmhQvLBJSbmoKBFY6QwRA9fF0pzCt28+dNiBJZKaGGAxFX6\n5zC7UIv7xzTNXDaj8oOE5qCuNDi8r0Tp2EDMDV+xX/iPr9JUcl24wHfStvEZnVMg\njTmVptzGBCwUw0P55VGNzxcyS8D68Og/AMz5MLLR8oOewBvr5twkEw1rHnMtJTDy\nmQX+yGvDiE4h6vNjnrzmjjIL9LuMbAxvc5Dk7J6uYDsEE3fYtd2VrqYwSq43xxdD\nAhfefrDJAgMBAAECggEAFYVNNUfPdElOA8z3uLY05wlfck+ehnfQlhJLxGtyRVWX\nWvuubpsp5QBhSqIpFbySKqL+c/vzPwr6iwBb2hLHYd7o4LDhLso9Ivr9aa0/EmCh\nGqJbs2SHrz+WkKQcD0G4em3qoySzrljwJ67SbWkgusizivT+C7xaF1sw2cfQPE6h\n+H/2HcUjedkZv6u91A3Nh4WayPNA0hsDYNbYpjAFqtIg5TTCiP1WyVEgEY/lq+sE\nEepNrhzb9Rmp6Fxq6apqcTgvxGLHrwu3Fc/HFq7JkAzHsUJ99n9Tej3tYtAPM0yc\n+MFLT7rZgVnNFpQzPcbwiXEeQzoFExuWtnir+0y1/QKBgQDk31EhcwHRjNEbrg2E\nVoiZq/fTSSJ2o8kYoCUlxL8oyYlPKe4JwAxhcfMOTKy3o2wTqkpTQKGCNVg2MhP/\nFwQxr+4SHdjeB6t7dRlKKxvmN1kYpZNb2zs2jCD59rvFE/ZdtDZDBiHu0pV+07Xp\nz5dp5dxk9a+cgnjg9UVDia1LbQKBgQDhHlbXrxD4mozi+N45gWR6v28DCoPDwXae\nU/CtIqVi6VYFz4eLbNu/9uS0tq4f17MBiOiltRG2aG6ccKtx78Tz/FFwumEGIkqy\nqs+/1BSNjcd9FcGkcmVMKCnZ7LYubppnkJIZa4fV9lIMUJydJrF8nSOhFTAMvpZp\nvLIQSfFlTQKBgBQS0hbQ83PhmeWHmn/k5w4zWwUZAQDO1LBoO1nYq7t0EarzzoDk\nazGQwPScHPnuR2hiIyqyHHhDHX2DXuWcqy2AdKz6GS9AFPY7CwDKTyQd7p6OxyHj\nVIowOCQ0U7uxSIZna+rs+sTri1kYUHg1UN5k3rOsKL7dYqS4Xl7SEHTxAoGBAItp\niIplxnLO83UUfjrKoPlLWGpftp4iT11ZynDORfHtYvKSRPTZY3WMZrJrd4YMxLSs\nnrcQXXnDTszfEa2ruSMIHT9cjP2Jew0Orz2zD09igCo8sQEwPv9c2B43c9Npd4Gv\njGrlpuegdctemL7R2ZS8k/YL8wfRd8DftL5VrIL5AoGBAIvFDsIsWv21RkYxrWb0\ntrzjwavPQ4kVLipa25wEKGt0E5PwEjQts/AJlAktx3B/OzNHEInc0JhyRx9jhiCi\nl2DAZbvk7cs4fdWuf3nGGF3DoB35CagwnFtj2cW2v2k61PAq/LYKgmY+Q9SiPYpl\nnhzaa7TSGtcNKVVzrSlpHymx\n-----END PRIVATE KEY-----\n",
  "client_email": "labsoftware@projetolab32066724.iam.gserviceaccount.com",
  "client_id": "109801249843264666130",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/labsoftware%40projetolab32066724.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}


# Define o escopo da API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# Informa as suas credenciais do Google API
# cria as credenciais a partir do conteúdo do arquivo JSON
credentials = ServiceAccountCredentials.from_json_keyfile_dict(json_content, scope)

# autentica as credenciais
gc = gspread.authorize(credentials)

# Abre a planilha pelo seu nome
sheet_name = 'LabSoftware'
planilha = gc.open(sheet_name)

# Seleciona a primeira aba da planilha
abaMaterial = planilha.sheet1
abaCliente = planilha.get_worksheet(1)


# Lê os dados da planilha e converte em um dataframe do pandas
df = pd.DataFrame(abaMaterial.get_all_records())

layout_aba1 = [
    [sg.Text('Nome do material:'), sg.Input(key='inputMaterial')],
    [sg.Button('Buscar', key='buscarMaterial'), sg.Button('Alterar', key='alterarMaterial'), sg.Button('Inserir', key='inserirMaterial'), sg.Button('Remover', key='removerMaterial')],
]
layout_aba2 = [
    [sg.Text('Nome do cliente:'), sg.Input(key='inputCliente')],
    [sg.Button('Buscar', key='buscarCliente'), sg.Button('Inserir', key='inserirCliente'), sg.Button('Remover', key='removerCliente'), sg.Button('Alterar', key='alterarCliente')],
]

layout = [
    [sg.TabGroup([
        [sg.Tab('Estoque', layout_aba1)],
        [sg.Tab('Clientes', layout_aba2)],
    ])]
]

window = sg.Window('SAGACHI', layout, resizable=False) 

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break

    if event == 'buscarMaterial':
        todosMateriais = abaMaterial.get_all_values()
        idMaterial = -1
        for i in range(1, len(todosMateriais)+1):
            if abaMaterial.cell(i,1).value == values['inputMaterial']:
                idMaterial = i

        if idMaterial == -1:
            sg.popup('Material não encontrado')
        else:
        
            layoutBM = [
            [sg.Text('Nome do Material:'), sg.Text(abaMaterial.cell(idMaterial,1).value)],
            [sg.Text('Quantidade no estoque:'), sg.Text(abaMaterial.cell(idMaterial,2).value)],
            [sg.Text('Preço:'), sg.Text(abaMaterial.cell(idMaterial,3).value)],
            [sg.Text('Fabricante:'), sg.Text(abaMaterial.cell(idMaterial,4).value)],
            [sg.Button('Voltar', key='voltaBM')],
            ]
            abaBM = sg.Window('Material Buscado', layoutBM, resizable=False)
            while True:
                event, values = abaBM.read()
                if event == sg.WIN_CLOSED or event == 'voltaBM':
                    abaBM.close()
                    break
        

    elif event == 'buscarCliente':
        todosClientes = abaCliente.get_all_values()
        idCliente = -1
        for i in range(1, len(todosClientes)+1):
            if abaCliente.cell(i,1).value == values['inputCliente']:
                idCliente = i

        if idCliente == -1:
            sg.popup('Cliente não encontrado')
        else:        
            layoutBC = [
            [sg.Text('Nome do Cliente:'), sg.Text(abaCliente.cell(idCliente,1).value)],
            [sg.Text('CPF ou CNPJ:'), sg.Text(abaCliente.cell(idCliente,2).value)],
            [sg.Text('Número do cliente:'), sg.Text(abaCliente.cell(idCliente,3).value)],
            [sg.Text('Email do Cliente:'), sg.Text(abaCliente.cell(idCliente,4).value)],
            [sg.Text('Endereço do Cliente:'), sg.Text(abaCliente.cell(idCliente,5).value)],
            [sg.Text('Data da última compra:'), sg.Text(abaCliente.cell(idCliente,6).value)],
            [sg.Button('Voltar', key='voltaBC')],
            ]
            abaBC = sg.Window('Material Buscado', layoutBC, resizable=False)
            while True:
                event, values = abaBC.read()
                if event == sg.WIN_CLOSED or event == 'voltaBC':
                    abaBC.close()
                    break

    elif event == 'inserirMaterial':
        layoutIM = [
            [sg.Text('Nome do Material:'), sg.Input(key='inputInsereNomeM')],
            [sg.Text('Quantidade no estoque:'), sg.Input(key='inputInsereQtdeM')],
            [sg.Text('Preço:'), sg.Input(key='inputInserePrecoM')],
            [sg.Text('Fabricante:'), sg.Input(key='inputInsereFabricanteM')],
            [sg.Button('Confirmar', key='confirmaIM'), sg.Button('Cancelar', key='cancelaIM')],
        ]
        abaIM = sg.Window('Inserir Material', layoutIM, resizable=False)
        while True:
            event, values = abaIM.read()
            if event == sg.WIN_CLOSED or event == 'cancelaIM' or event == 'confirmaIM':
                if event == 'confirmaIM':
                    todosMateriais = abaMaterial.get_all_values()
                    linha = len(todosMateriais)+1
                    abaMaterial.update_cell(linha, 1, values['inputInsereNomeM'])
                    abaMaterial.update_cell(linha, 2, values['inputInsereQtdeM'])
                    abaMaterial.update_cell(linha, 3, values['inputInserePrecoM'])
                    abaMaterial.update_cell(linha, 4, values['inputInsereFabricanteM'])
                    
                    sg.popup('Inserção realizada com sucesso')                
                abaIM.close()
                break

    elif event == 'inserirCliente':
        layoutIC = [
            [sg.Text('Nome do Cliente:'), sg.Input(key='inputInsereNomeC')],
            [sg.Text('CPF ou CNPJ:'), sg.Input(key='inputInseteCPFC')],
            [sg.Text('Número do Cliente:'), sg.Input(key='inputInsereNumeroC')],
            [sg.Text('Email do Cliente:'), sg.Input(key='inputInsereEmailC')],
            [sg.Text('Endereço do Cliente:'), sg.Input(key='inputInsereEnderecoC')],
            [sg.Text('Data:'), sg.Input(key='inputInsereDataC')],
            [sg.Button('Confirmar', key='confirmaIC'), sg.Button('Cancelar', key='cancelaIC')],
        ]
        abaIC = sg.Window('Inserir Cliente', layoutIC, resizable=False)
        while True:
            event, values = abaIC.read()
            if event == sg.WIN_CLOSED or event == 'cancelaIC' or event == 'confirmaIC':
                if event == 'confirmaIC':
                    todosClientes = abaCliente.get_all_values()
                    linha = len(todosClientes)+1
                    abaCliente.update_cell(linha, 1, values['inputInsereNomeC'])
                    abaCliente.update_cell(linha, 2, values['inputInseteCPFC'])
                    abaCliente.update_cell(linha, 3, values['inputInsereNumeroC'])
                    abaCliente.update_cell(linha, 4, values['inputInsereEmailC'])
                    abaCliente.update_cell(linha, 5, values['inputInsereEnderecoC'])
                    abaCliente.update_cell(linha, 6, values['inputInsereDataC'])                    
                    sg.popup('Inserção realizada com sucesso')                
                abaIC.close()
                break

    elif event == 'alterarMaterial':
        todosMateriais = abaMaterial.get_all_values()
        idMaterial = -1
        for i in range(1, len(todosMateriais)+1):
            if abaMaterial.cell(i,1).value == values['inputMaterial']:
                idMaterial = i

        if idMaterial == -1:
            sg.popup('Material não encontrado')
        else:        
            layoutAM = [
                [sg.Text('Nome do Material:'), sg.InputText(abaMaterial.cell(idMaterial,1).value, key='inputAlteraNomeM')],
                [sg.Text('Quantidade no estoque:'), sg.InputText(abaMaterial.cell(idMaterial,2).value, key='inputAlteraQtdeM')],
                [sg.Text('Preço:'), sg.InputText(abaMaterial.cell(idMaterial,3).value, key='inputAlteraPrecoM')],
                [sg.Text('Fabricante:'), sg.InputText(abaMaterial.cell(idMaterial,4).value, key='inputAlteraFabricanteM')],
                [sg.Button('Confirmar', key='confirmaAM'), sg.Button('Cancelar', key='cancelaAM')],
            ]
            abaAM = sg.Window('Alterar Material', layoutAM, resizable=False)
            while True:
                event, values = abaAM.read()
                if event == sg.WIN_CLOSED or event == 'cancelaAM' or event == 'confirmaAM':
                    if event == 'confirmaAM':
                        abaMaterial.update_cell(idMaterial, 1, values['inputAlteraNomeM'])
                        abaMaterial.update_cell(idMaterial, 2, values['inputAlteraQtdeM'])
                        abaMaterial.update_cell(idMaterial, 3, values['inputAlteraPrecoM'])
                        abaMaterial.update_cell(idMaterial, 4, values['inputAlteraFabricanteM'])
                        sg.popup('Alteração realizada com sucesso')
                    abaAM.close()
                    break


    elif event == 'alterarCliente':
        todosClientes = abaCliente.get_all_values()
        idCliente = -1
        for i in range(1, len(todosClientes)+1):
            if abaCliente.cell(i,1).value == values['inputCliente']:
                idCliente = i

        if idCliente == -1:
            sg.popup('Cliente não encontrado')
        else:        
            layoutAC = [
                [sg.Text('Nome do Cliente:'), sg.InputText(abaCliente.cell(idCliente,1).value, key='inputAlteraNomeC')],
                [sg.Text('CPF ou CNPJ:'), sg.InputText(abaCliente.cell(idCliente,2).value, key='inputAlteraCPFC')],
                [sg.Text('Número do Cliente:'), sg.InputText(abaCliente.cell(idCliente,3).value, key='inputAlteraNumeroC')],
                [sg.Text('Email do Cliente:'), sg.InputText(abaCliente.cell(idCliente,4).value, key='inputAlteraEmailC')],
                [sg.Text('Endereço do Cliente:'), sg.InputText(abaCliente.cell(idCliente,5).value, key='inputAlteraEnderecoC')],
                [sg.Text('Data:'), sg.InputText(abaCliente.cell(idCliente,6).value, key='inputAlteraDataC')],
                [sg.Button('Confirmar', key='confirmaAC'), sg.Button('Cancelar', key='cancelaAC')],
            ]
            abaAC = sg.Window('Alterar Cliente', layoutAC, resizable=False)
            while True:
                event, values = abaAC.read()
                if event == sg.WIN_CLOSED or event == 'cancelaAC' or event == 'confirmaAC':
                    if event == 'confirmaAC':
                        abaCliente.update_cell(idCliente, 1, values['inputAlteraNomeC'])
                        abaCliente.update_cell(idCliente, 2, values['inputAlteraCPFC'])
                        abaCliente.update_cell(idCliente, 3, values['inputAlteraNumeroC'])
                        abaCliente.update_cell(idCliente, 4, values['inputAlteraEmailC'])
                        abaCliente.update_cell(idCliente, 5, values['inputAlteraEnderecoC'])
                        abaCliente.update_cell(idCliente, 6, values['inputAlteraDataC'])
                        sg.popup('Alteração realizada com sucesso')
                    abaAC.close()
                    break

    if event == 'removerMaterial':
        todosMateriais = abaMaterial.get_all_values()
        idMaterial = -1
        for i in range(1, len(todosMateriais)+1):
            if abaMaterial.cell(i,1).value == values['inputMaterial']:
                idMaterial = i

        if idMaterial == -1:
            sg.popup('Material não encontrado')
        else:
            layoutRM = [
            [sg.Text('Deseja remover o seguinte material?')],
            [sg.Text('Nome do Material: ' + values["inputMaterial"])],
            [sg.Button('Confirmar', key='confirmaRM'), sg.Button('Cancelar', key='cancelaRM')],
            ]
            abaRM = sg.Window('Material a remover', layoutRM, resizable=False)
            while True:
                event, values = abaRM.read()
                if event == sg.WIN_CLOSED or event == 'cancelaRM' or event =='confirmaRM':
                    if event == 'confirmaRM':
                        abaMaterial.delete_rows(idMaterial)
                        sg.popup('Material removido com sucesso')
                    abaRM.close()
                    break

    if event == 'removerCliente':
        todosClientes = abaCliente.get_all_values()
        idCliente = -1
        for i in range(1, len(todosClientes)+1):
            if abaCliente.cell(i,1).value == values['inputCliente']:
                idCliente = i

        if idCliente == -1:
            sg.popup('Cliente não encontrado')
        else:
            layoutRC = [
            [sg.Text('Deseja remover o seguinte cliente?')],
            [sg.Text('Nome do cliente: ' + values["inputCliente"])],
            [sg.Button('Confirmar', key='confirmaRC'), sg.Button('Cancelar', key='cancelaRC')],
            ]
            abaRC = sg.Window('Cliente a remover', layoutRC, resizable=False)
            while True:
                event, values = abaRC.read()
                if event == sg.WIN_CLOSED or event == 'cancelaRC' or event == 'confirmaRC':
                    if event == 'confirmaRC':
                        abaCliente.delete_rows(idCliente)
                        sg.popup('Cliente removido com sucesso')
                    abaRC.close()
                    break


window.close()
