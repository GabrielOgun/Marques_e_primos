import PySimpleGUI as sg
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

    
json_content = {
  "type": "service_account",
  "project_id": "teste-273304",
  "private_key_id": "100b385b6b1233ec7c2599b92317c455047b3421",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDtTyarWSM31e4F\nZvlmLnrQGDfoIc4Z+ZKuV5AeLcxa9ORvF6EmIigwiKDCWCGJ87mnf0HeaENPtq2r\nHkBaL9TEMWbjL5nZkqFCiBYuG5ta6cqFOdwaFwHQfceVtD1AV7SngGH2bjbY2rlp\n4WdpNiReTNkHqXh6CAeANGCk33QaOhnrEvC9yKeygIpXxbJ7zLqXQUvTBohxLIpK\niPySizhAs+WwlnJIs6lpLQZnPR9w9XGyYX4opAKGLslfNI/VXD4Kaj+qC/ejHXI3\noy+yiJxKLLACn/7NK1r1+hb0vT+eNQsipzz0f3fgrAE55nweg/iWm6y1kEiQPPi5\nTK0cBRunAgMBAAECggEAHeNAHYiGdPvOlIOZmZL1CMxkDipjyMW0AZ0pm4NtH2+E\nbbFuLF1U7nfmt1NeNf+qPDw80YQUJi/9w3V16WXoyCTormhKWiqrgLOfB4OWl2am\niQz2eZq4McgFoQcoR7hEGmyC6gSLh9hUTc+DtK6K+g13sA1aDRSBzVXLbjhuaPb2\nQYjouoIHouBIJg8VRlECJsPX+a4P7bD/Ew3mTVyudBx/8l4gwmETGx0+yeMni9XN\nLdXVChxqCTl25Wh3ag3A1z4KE5903IZ8RR5v/sM5yy/kFRQUzdKr1fzN1J6GLJP3\nXeYp56iXltoAYucOaAu/vbKfrmdzFKeB7wvMd/iCcQKBgQD5jDuyyJjt/lwq9zs4\nZvm86+Dlfhj1fRqoNBo/aoXH1AQC0ty7Krz2fbFeiajSZXxwtHVBHPzwW1xC6mei\n2ruNwzhFv8VypvMJicDyx+lMmCjppzZW1PPZy+n2vUzGF2o2ks/K1IIERE900ZVK\nGDSamBsALl75Xt96OG7/OFVCFQKBgQDzcej5fn+dwwgtzdzP42CHscQPUVQ/KGUa\nuE1TC4TLJpAf8HTTksW19BgpaYWeBqeOrk5n2S09CqTd9EsCHYPBLnrSOeAHrxT7\n0/LT3JDHZsgP5S1tmhIcWCG/h2K9ybgJaoOGSldrUOIP0nf8ZWXkDy1yvuxJDHBV\nTEW4MZEhywKBgQCub7MvRv93pUziD83KoFjEEZI0eU+TEm820qziVWDMjUx8eM8o\n2jgaiUQZ1Fo5MA2rbsljyZKZpRM6B0aIVSOzdZn5T5MCkObkbPF+A/X1v4shwOvA\nCL2oKd0Sx8JJ2gY5vagYnTGBMArmmrYjhAYJZnfBSajD4eiPM7GLH+Kg/QKBgED9\nsAyrULZ1UsGnq8N0GFkhhA3y0GLsDdHMUhhRguoZKXDdaNLy5AVnXOvxV9KQRDs7\nHYNr3z/kj48RoNS3vGzeU7u756slepygQLt+rsgNEGvt6urPrvYSMTBInHu+Vntt\nDB/VyPDFbxR6Q74F8+Wmh6OShNIbmAGtkw9RbEVxAoGBAIi4ISP95BTcxNpaXDRs\n4fIjdvRDkI/v9uKm8WUvnIpYDrjUC+zZKkWMI7PHVSGl6NHq5Iv8Ad6w2s2pgRWS\n7Gjcbc+XjV08zI33dqsUGKNEz4oV+f8RMOoavpONboJrM0pCnXropgorYPUmkPvQ\neCuK9cywgs8o7qAQr7x4XRUv\n-----END PRIVATE KEY-----\n",
  "client_email": "testeplanilha@teste-273304.iam.gserviceaccount.com",
  "client_id": "115839004229941996790",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/testeplanilha%40teste-273304.iam.gserviceaccount.com"
}





# Define o escopo da API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# Informa as suas credenciais do Google API
# cria as credenciais a partir do conteúdo do arquivo JSON
credentials = ServiceAccountCredentials.from_json_keyfile_dict(json_content, scope)

# autentica as credenciais
gc = gspread.authorize(credentials)

# Abre a planilha pelo seu nome
sheet_name = 'Teste'
planilha = gc.open(sheet_name)

# Seleciona a primeira aba da planilha
aba = planilha.sheet1

# Lê os dados da planilha e converte em um dataframe do pandas
df = pd.DataFrame(aba.get_all_records())


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
        layoutBM = [
        [sg.Text('Nome do Material:'), sg.Text('<Nome Pego no DB>')],
        [sg.Text('Quantidade no estoque:'), sg.Text('<Quantidade pega no DB>')],
        [sg.Text('Preço:'), sg.Text('<Preço Pego no DB>')],
        [sg.Text('Fabricante:'), sg.Text('<Fabricante Pego no DB>')],
        [sg.Button('Voltar', key='voltaBM')],
        ]
        abaBM = sg.Window('Material Buscado', layoutBM, resizable=False)
        while True:
            event, values = abaBM.read()
            if event == sg.WIN_CLOSED or event == 'voltaBM':
                abaBM.close()
                break
        

    elif event == 'buscarCliente':
        layoutBC = [
        [sg.Text('Nome do Cliente:'), sg.Text('<Nome Pego no DB>')],
        [sg.Text('CPF ou CNPJ:'), sg.Text('<CNPJ pego no DB>')],
        [sg.Text('Número do cliente:'), sg.Text('<Número Pego no DB>')],
        [sg.Text('Email do Cliente:'), sg.Text('<Email pego no DB>')],
        [sg.Text('Endereço do Cliente:'), sg.Text('<Endereço Pego no DB>')],
        [sg.Text('Data da última compra:'), sg.Text('<Data no DB>')],
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
                    sg.popup('Inserção realizada com sucesso')                
                abaIC.close()
                break

    elif event == 'alterarMaterial':
        layoutAM = [
            [sg.Text('Nome do Material:'), sg.InputText('<Nome no DB>', key='inputAlteraNomeM')],
            [sg.Text('Quantidade no estoque:'), sg.InputText('<Qtde no DB>', key='inputAlteraQtdeM')],
            [sg.Text('Preço:'), sg.InputText('<Preco no DB>', key='inputAlteraPrecoM')],
            [sg.Text('Fabricante:'), sg.InputText('<Fabricante no DB>', key='inputAlteraFabricanteM')],
            [sg.Button('Confirmar', key='confirmaAM'), sg.Button('Cancelar', key='cancelaAM')],
        ]
        abaAM = sg.Window('Alterar Material', layoutAM, resizable=False)
        while True:
            event, values = abaAM.read()
            if event == sg.WIN_CLOSED or event == 'cancelaAM' or event == 'confirmaAM':
                if event == 'confirmaAM':
                    sg.popup('Aleração realizada com sucesso')
                abaAM.close()
                break


    elif event == 'alterarCliente':
        layoutAC = [
            [sg.Text('Nome do Cliente:'), sg.InputText('<Nome no DB>', key='inputAlteraNomeC')],
            [sg.Text('CPF ou CNPJ:'), sg.InputText('<CPF/CNPJ no DB>', key='inputAlteraCPFC')],
            [sg.Text('Número do Cliente:'), sg.InputText('<Número no DB>', key='inputAlteraNumeroC')],
            [sg.Text('Email do Cliente:'), sg.InputText('<Email no DB>', key='inputAlteraEmailC')],
            [sg.Text('Endereço do Cliente:'), sg.InputText('<Endereço no DB>', key='inputAlteraEnderecoC')],
            [sg.Text('Data:'), sg.InputText('<Data no DB>', key='inputAlteraDataC')],
            [sg.Button('Confirmar', key='confirmaAC'), sg.Button('Cancelar', key='cancelaAC')],
        ]
        abaAC = sg.Window('Alterar Cliente', layoutAC, resizable=False)
        while True:
            event, values = abaAC.read()
            if event == sg.WIN_CLOSED or event == 'cancelaAC' or event == 'confirmaAC':
                if event == 'confirmaAC':
                    sg.popup('Aleração realizada com sucesso')
                abaAC.close()
                break

    if event == 'removerMaterial':
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
                    sg.popup('Material removido com sucesso')
                abaRM.close()
                break

    if event == 'removerCliente':
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
                    sg.popup('Cliente removido com sucesso')
                abaRC.close()
                break


window.close()
