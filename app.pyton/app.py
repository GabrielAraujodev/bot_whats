import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

# Abrir o WhatsApp Web e aguardar o carregamento
webbrowser.open('https://web.whatsapp.com/')
sleep(30)  # Ajuste o tempo se necessário

# Ler a planilha e armazenar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # Nome, telefone e vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    # Verifica se vencimento não é None antes de formatar a data
    if vencimento:
        data_vencimento = vencimento.strftime("%d/%m/%Y")
    else:
        data_vencimento = "indefinida"  # Define um valor padrão se vencimento for None

    # Formatar a mensagem
    mensagem = f'Olá {nome}, seu boleto vence no dia {data_vencimento}. Favor pagar no link'

    # Criar o link personalizado do WhatsApp e enviar a mensagem
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        
        # Localizar o botão de envio e clicar para enviar a mensagem
        seta = pyautogui.locateCenterOnScreen('seta.png')
        if seta:
            sleep(5)
            pyautogui.click(seta[0], seta[1])
            sleep(5)
            pyautogui.hotkey('ctrl', 'w')
            sleep(5)
        else:
            print(f'Não foi possível encontrar o botão de envio para {nome}')
            with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone}{os.linesep}')
                
    except Exception as e:
        print(f'Erro ao enviar mensagem para {nome}: {e}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
