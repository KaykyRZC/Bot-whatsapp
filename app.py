import openpyxl
import	webbrowser
import pyautogui
from urllib.parse import quote
from time import sleep

workbook = openpyxl.load_workbook('Dados.xlsx') #In place of 'Dados.xlsx', put your Excel file name.
data_page = workbook['Planilha1'] #In place of 'Planilha1', use the actual name of the sheet in your Excel file.

for row in data_page.iter_rows(min_row = 2):
    name = row[0].value #The 0 is the index of the column number in the Excel file
    telephone = row[1].value
    due_date = row[2].value

    message = f'Ola, {name}. Em que posso ajudar?'

    link_message_whatsapp = f'https://web.whatsapp.com/send?phone={telephone}&text={quote(message)}'
    webbrowser.open(link_message_whatsapp)
    sleep(10)

    try:
        send_button = pyautogui.locateCenterOnScreen('send_button.png')
        sleep(5)
        pyautogui.click(send_button[0], send_button[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Not possible to send a message to {name}, {telephone}')
        with open('fails.csv', 'a', newline='', encoding='utf-8') as file:
            file.write(f' {name} {telephone} ')