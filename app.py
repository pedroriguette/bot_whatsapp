import openpyxl
import webbrowser
import pyautogui
from time import sleep
from datetime import datetime
from urllib.parse import quote

#to read and save data in excel spreadsheet as name, number phone and due date
workbook = openpyxl.load_workbook('listadenomes.xlsx')
customers_page = workbook['Página1']

for coluna in customers_page.iter_rows(min_row=2):
    present_day = datetime.now()
    formated_date = present_day.strftime('%d/%m/%Y')
    if coluna[2].value.strftime('%d/%m/%Y') == formated_date:
        name = coluna[0].value
        number_phone = int(coluna[1].value)
        due_date = coluna[2].value

        message = f'Olá {name} seu boleto venci no dia {due_date.strftime('%d/%m/%Y')} favor pagar no link...'

    # To create custom link for whatsapp and send menssage for custumers
    # based on spreadsheet data
        try:
            link_menssage_whatsapp = f'https://web.whatsapp.com/send?phone={number_phone}&text={quote(message)}'
            webbrowser.open(link_menssage_whatsapp)
            sleep(15)
            seta = pyautogui.locateCenterOnScreen('seta.png')
            sleep(10)
            pyautogui.click(seta[0],seta[1])
            sleep(5)
            pyautogui.hotkey('ctrl','w')
            sleep(5)
        except:
            print(f'não foi possivel enviar a mensagem para {name}')
            with open('erros.csv', 'a',newline='',encoding='utf-8') as doc:
                doc.write(f'NOME: {name}, NUMERO{number_phone}')
            