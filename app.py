import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

#instalar a biblioteca pillow - pip install pollow

#A tabela deve estar  normalizada e os numeros de telefone devem conter o codigo do pais e numeros adicionais
#Exemplo 55889xxxxxxxx

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

workbook = openpyxl.load_workbook('cliente.xlsx')
pagina_clientes = workbook['plan1'] # inserir o nome correto da planilha em que se encontram os contatos

# OBS: inserir poucos contatos para evitar o bloqueio da conta, cuidado com envio em massa!

for linha in pagina_clientes.iter_row(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar no link...'

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possivel envia mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')