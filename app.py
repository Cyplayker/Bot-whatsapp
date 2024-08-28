import openpyxl
from datetime import datetime
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui



workbook = openpyxl.load_workbook('ListaDeContatos.xlsx')

Pagina_clientes = workbook['Plan1']

dia_atual = datetime.now().day

for Linha in Pagina_clientes.iter_rows(min_row=2,max_row=2):

   
    nome = Linha[0].value
    telefone = Linha[1].value
    vencimento = Linha[2].value
    mensagem = f'ola {nome} seu vencimento e {vencimento.strftime("%d/%m/%Y")} '
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(10)
    try:
        
        seta = pyautogui.locateCenterOnScreen('enviar.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)

    except:
        print('erro')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone},{vencimento}\n')
        



