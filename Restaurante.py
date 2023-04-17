import PySimpleGUI as sg
import pandas as pd
import win32print
import getpass
import os
import datetime

data_atual = datetime.datetime.today().strftime('%d-%m-%Y')

sg.theme('Black')
lista= {
    'Marmita':[
        'P',
        'M',
        'G',
    ],
    'Refrigerante':[
        '200 mL',
        '350 mL',
        '500mL',
        '600 mL',
        '1 L',
        '2 L',
    ]}
calc_layout = [
    [sg.InputText('', size=(13, 1), font=('Helvetica', 18), key='input')],
    [sg.Button('7'), sg.Button('8'), sg.Button('9'), sg.Button('+')],
    [sg.Button('4'), sg.Button('5'), sg.Button('6'), sg.Button('-')],
    [sg.Button('1'), sg.Button('2'), sg.Button('3'), sg.Button('*')],
    [sg.Button('0'), sg.Button('.'), sg.Button('C'), sg.Button('/')],
    [sg.Button('('), sg.Button(')'), sg.Button('=', size=(13,1))],
    [sg.Image(filename='logoRes.png')]
]
# Define o layout do frame superior esquerdo
layout_res = [
    [sg.Text('Produto'), sg.Text('Quantidade')],
    [sg.Text('Almoço'), sg.InputText(size=(3,1), key='Almoço'),sg.Text('Conteudo'), sg.InputText(size=(10,1), key='Conteudo_Alm')],
    [sg.Text('Marmita'), sg.Combo(lista['Marmita'], key='Tam_Marmita'), sg.InputText(size=(3,1), key='Marmita'),sg.Text('Conteudo'), sg.InputText(size=(10,1), key='Conteudo_Mar')],
    [sg.Text('Salada'), sg.InputText(size=(3,1), key='Salada')],
    [sg.Text('Refrigerante'), sg.Combo(lista['Refrigerante'], key='Refrigerante'), sg.InputText(size=(3,1), key='Val_Refrigerante')],
    [sg.Text('Outros'), sg.InputText(size=(15,1), key='Outros')],
    [sg.Text('Valor Total:')], [sg.Text('R$'), sg.InputText(size=(5,1), key='Val_Total')],
    [sg.Text('Numero de Contato'), sg.InputText(size=(11,1), key='Nmr_CTT')],
    [sg.Text('Nome'), sg.InputText(size=(7,1),key='Nome')],
    [sg.Text('Rua:')], [sg.InputText(key='Rua')],
    [sg.Text('Numero:')], [sg.InputText(key='Numero')],
    [sg.Text('Complemento:')], [sg.InputText(key='Complemento')],
    [sg.Button('Enviar'), sg.Button('Apagar')]
]

layout= [
    [sg.Column([[sg.Frame('Pedido', layout_res)]], element_justification='left', expand_x=True),
    sg.Column([[sg.Frame('Calculadora', calc_layout)]], element_justification='right', expand_x=True)],
]

window = sg.Window('Sonho Meu').Layout(layout)

while True:
    event, values = window.read()
    if event in (None, 'Exit'):
        break
    
    # Funções da calculadora
    if event in ('0', '1', '2', '3', '4', '5', '6', '7', '8', '9'):
        if values['input'] == '':
            window['input'].update(event)
        else:
            window['input'].update(values['input'] + event)
    elif event in ('+', '-', '*', '/'):
        if values['input'] == '':
            sg.popup('Digite um valor primeiro!')
        else:
            try:
                result = eval(values['input'])
                window['input'].update(str(result) + event)
            except:
                sg.popup('Operação inválida!')
    elif event == '.':
        if '.' in values['input']:
            sg.popup('Já existe um ponto decimal!')
        elif values['input'] == '':
            window['input'].update('0.')
        else:
            window['input'].update(values['input'] + '.')
    elif event == '(':
        window['input'].update(values['input'] + '(')
    elif event == ')':
        window['input'].update(values['input'] + ')')
    elif event== 'C':
        window['input'].update('')
    elif event == '=':
        if values['input'] == '':
            sg.popup('Digite um valor primeiro!')
        else:
            try:
                result = eval(values['input'])
                window['input'].update(str(result))
            except:
                sg.popup('Operação inválida!')
    
    # Funções dos botões de enviar e apagar
    if event == 'Enviar':

            x = []
            dados = {
                'Almoço': values['Almoço'], 'Conteudo_Alm': values['Conteudo_Alm'], 'Tam Marmita': values['Tam_Marmita'],'Salada': values['Salada'],'Outros': values['Outros'],
                'Marmita': values['Marmita'], 'Conteudo_Mar': values['Conteudo_Mar'],'Refrigerante': values['Refrigerante'],'Rua': values['Rua'], 'Numero': values['Numero'],
                'Complemento': values['Complemento'],'Nome':values['Nome'],'Numero de contato': values['Nmr_CTT'],'Valor Total': values['Val_Total']
            }
            colunas = {'Almoço': values['Almoço'], 'Conteudo_Alm': values['Conteudo_Alm'], 'Marmita': values['Marmita'], 'Tam Marmita': values['Tam_Marmita'],'Conteudo_Mar': values['Conteudo_Mar'],
                        'Salada': values['Salada'],'Refrigerante': values['Refrigerante'],'Outros': values['Outros'],
                        'Rua': values['Rua'], 'Numero': values['Numero'],
                        'Complemento': values['Complemento'],'Nome':values['Nome'],'Numero de contato': values['Nmr_CTT'],
                        'Valor Total': values['Val_Total'],
                        }
            df = pd.DataFrame(columns=colunas)
            nova_linha = pd.DataFrame([dados], columns=colunas)
            df = pd.concat([df, nova_linha], ignore_index=True)

            # Salvar em arquivo
            with open(f'Relatorio_dia_{data_atual}.csv', 'a') as f:
                nova_linha.to_csv(f, header=f.tell()==0, index=False)  # Adiciona header só na primeira linha
                f.close()
            # cria uma string com os dados formatados
            output = ''
            for i, row in df.iterrows():
                output += f'Pedido {i+1}:\n'
                for header in df:
                    value = row.get(header, '')
                    if value:
                        output += f'{header}: {value}\n'
                output += '\n'

            # abre a impressora padrão
            printer = win32print.OpenPrinter(win32print.GetDefaultPrinter())

            # inicia um trabalho de impressão
            job = win32print.StartDocPrinter(printer, 1, ("test", None, "RAW"))

            # envia o conteúdo para a impressora
            win32print.StartPagePrinter(printer)
            win32print.WritePrinter(printer, output.encode())
            win32print.EndPagePrinter(printer)

            # finaliza o trabalho de impressão
            win32print.EndDocPrinter(printer)
            win32print.ClosePrinter(printer)
            #apaga todos os inserts
            window['Almoço'].update('')
            window['Tam_Marmita'].update('')
            window['Conteudo_Alm'].update('')
            window['Conteudo_Mar'].update('')
            window['Marmita'].update('')
            window['Salada'].update('')
            window['Refrigerante'].update('')
            window['Val_Refrigerante'].update('')
            window['Outros'].update('')
            window['Val_Total'].update('')
            window['Nmr_CTT'].update('')
            window['Nome'].update('')
            window['Rua'].update('')
            window['Numero'].update('')
            window['Complemento'].update('')

    elif event == 'Apagar':
        window['Almoço'].update('')
        window['Conteudo_Alm'].update('')
        window['Conteudo_Mar'].update('')
        window['Tam_Marmita'].update('')
        window['Marmita'].update('')
        window['Salada'].update('')
        window['Refrigerante'].update('')
        window['Val_Refrigerante'].update('')
        window['Outros'].update('')
        window['Val_Total'].update('')
        window['Nmr_CTT'].update('')
        window['Nome'].update('')
        window['Rua'].update('')
        window['Numero'].update('')
        window['Complemento'].update('')

window.close()

