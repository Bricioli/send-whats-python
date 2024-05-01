import webbrowser
import openpyxl
import pyautogui
from urllib.parse import quote
from time import sleep


workbook = openpyxl.load_workbook('contatos.xlsx')
client_page = workbook['contatos']
count = 0;
for row in client_page.iter_rows(min_row=2):
    name = row[0].value
    phone = row[1].value
    message = row[2].value
#   message = f"Olá, {name}, tudo bem?! pode ignorar essa msg, só um teste de envio automatico de msg com foto via python, {count}"


   

    try: 
        link_whatsapp = f'https://web.whatsapp.com/send?phone={phone}&text={quote(message)}'
        webbrowser.open(link_whatsapp)
        sleep(10)
        pyautogui.click(759, 981)
        sleep(0.5)
        pyautogui.click(804, 759)
        sleep(0.5)
        pyautogui.hotkey('ctrl', 'v')
        sleep(0.5)
        pyautogui.press('enter')
        sleep(3)
        pyautogui.press('enter')
        sleep(2)
        pyautogui.hotkey('ctrl','w')
    except Exception as error:
        print(f'Não possivel enviar mensagem para {name}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as file: 
            file.write(f'{name}, {phone}')
            
            
            


