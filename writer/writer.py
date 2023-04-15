from __future__ import print_function
import os.path
import tkinter as tk
import pickle

import openpyxl
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials



DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'

def read_property(file: str):
    with open(file, 'r', encoding='utf-8') as file:
        lines = file.readlines()
        # Удаляем символы перевода строки ('\n') из каждой строки и добавляем их в список
        lines = [line.strip() for line in lines]
        property = []
        for p in lines:
            param = p.split(":", 1)
            property.append(param[1])
        for i in range(len(property)):
            property[i] = property[i].strip()
        return property


PROPERTY = tuple(read_property('property.txt'))
FONT = int(PROPERTY[4])
LOVE = 0

class GoogleSheet:
    SPREADSHEET_ID = PROPERTY[2]
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    service = None

    def __init__(self):
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                print('flow')
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.service = build('sheets', 'v4', credentials=creds, discoveryServiceUrl=DISCOVERY_SERVICE_URL)

    def updateRangeValues(self, values, sheet_name):
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.SPREADSHEET_ID, range=sheet_name).execute()
        num_rows = len(result['values'])

        range_ = f'{sheet_name}!A{num_rows + 1}:{chr(ord("A") + len(values) - 1)}{num_rows + 1}'
        value_input_option = 'RAW'
        body = {'range': range_, 'values': [values], 'majorDimension': 'ROWS'}
        result = sheet.values().append(spreadsheetId=self.SPREADSHEET_ID, range=range_, valueInputOption=value_input_option,
                              insertDataOption='INSERT_ROWS', body=body).execute()


        print('{0} cells updated.'.format(result.get('updates').get('updatedCells')))
        #print(result)

def data_list_google_sheet():
    num = num_entry.get()
    out_num = out_num_entry.get()
    out_date = out_date_entry.get()
    agent = agent_entry.get()
    content = content_entry.get()
    data = [num, agent, content + " №" + out_num + " от " + out_date]
    return data

def data_list_excel():
    num = num_entry.get()
    date_input = date_input_entry.get()
    out_num = out_num_entry.get()
    out_date = out_date_entry.get()
    sender = sender_entry.get()
    agent = agent_entry.get()
    content = content_entry.get()
    data = [num, date_input, out_num, out_date, sender, agent, content]
    return data

def clear_entry():
    num_entry.delete(0, tk.END)
    date_input_entry.delete(0, tk.END)
    out_num_entry.delete(0, tk.END)
    out_date_entry.delete(0, tk.END)
    sender_entry.delete(0, tk.END)
    agent_entry.delete(0, tk.END)
    content_entry.delete(0, tk.END)

def result(is_add_google, is_add_excel):
    if (is_add_excel and is_add_google):
        result_label.config(text="Данные успешно добавлены!", fg="#00FF00")
        return True
    elif (is_add_excel != True and is_add_google):
        result_label.config(text="Произошла ошибка с добавление в Excel!", fg="red")
        return False
    elif (is_add_excel  and is_add_google != True):
        result_label.config(text="Произошла ошибка с добавление в Google Sheets!", fg="red")
        return False
    elif (is_add_excel != True and is_add_google != True):
        result_label.config(text="Произошла ошибка с добавление в Google Sheets и Excel!", fg="red")
        return False


def add_google_sheet():
    try:
        gs = GoogleSheet()
        test_values = data_list_google_sheet()
        gs.updateRangeValues(test_values, PROPERTY[3])
        return True
    except Exception as e:
        print("Произошла ошибка: {}".format(e))
        if "404" in str(e):
            error_label.config(text="Не верно указана ссылка", fg="red")
        else:
            error_label.config(text="Произошла ошибка: {}".format(e), fg="red")
        return False

def add_local_sheet():
    try:
        file_path = PROPERTY[0]
        workbook = openpyxl.load_workbook(file_path)
        sheet_name = PROPERTY[1]

        # Получаем объект листа по его имени
        sheet = workbook[sheet_name]
        sheet.append(data_list_excel())

        # Сохраняем изменения в файл Excel
        workbook.save(file_path)

        # Закрываем файл Excel
        workbook.close()
        return True
    except Exception as e:
        #print("Произошла ошибка: {}".format(e))
        if "Supported format" in str(e):
            error_label.config(text="Не верно указан формат файла", fg="red")
        elif "old .xls" in str(e):
            error_label.config(text="Не верно указан формат файла", fg="red")
        else:
            error_label.config(text="Произошла ошибка: {}".format(e), fg="red")
        return False

def add_all():
    global LOVE

    is_google = add_google_sheet()
    is_excel = add_local_sheet()
    if result(is_google, is_excel) != False:
        clear_entry()
        LOVE += 1
        if LOVE == 10:
            love_label.config(text="♡⸜(˃ ᵕ ˂)⸝")
            LOVE = 0
        else:
            love_label.config(text="")


def keys(event):
    if event.keycode==13:
        add_all()
    elif event.keycode==86:
        event.widget.event_generate("<<Paste>>")
    elif event.keycode==67:
        event.widget.event_generate("<<Copy>>")
    elif event.keycode==88:
        event.widget.event_generate("<<Cut>>")
    elif event.keycode==65535:
        event.widget.event_generate("<<Clear>>")
    elif event.keycode==65:
        event.widget.event_generate("<<SelectAll>>")


def test(num):
    print(num)


def main():

# Создаем главное окно
    global num_entry, date_input_entry, out_num_entry, out_date_entry, sender_entry, agent_entry, content_entry, result_label, error_label, love_label
    root = tk.Tk()
    root.title("Для Анечки <3")
    root.minsize(1000, 800)  # Устанавливаем разрешение окна

    # Создаем поле ввода текста


    num_label = tk.Label(root, text="№ п/п", font=("Helvetica", FONT))
    num_label.pack()
    num_entry = tk.Entry(root, width=50, font=("Helvetica", FONT))
    num_entry.pack(pady=10, padx=10)

    # Создаем метку и поле ввода для date_input
    date_input_label = tk.Label(root, text="Дата поступления", font=("Helvetica", FONT))
    date_input_label.pack()
    date_input_entry = tk.Entry(root, width=50, font=("Helvetica", FONT))
    date_input_entry.pack(pady=10, padx=10)

    # Создаем метку и поле ввода для out_num
    out_num_label = tk.Label(root, text="Исходящий №", font=("Helvetica", FONT))
    out_num_label.pack()
    out_num_entry = tk.Entry(root, width=50, font=("Helvetica", FONT))
    out_num_entry.pack(pady=10, padx=10)

    # Создаем метку и поле ввода для out_date
    out_date_label = tk.Label(root, text="Исходящая дата", font=("Helvetica", FONT))
    out_date_label.pack()
    out_date_entry = tk.Entry(root, width=50, font=("Helvetica", FONT))
    out_date_entry.pack(pady=10, padx=10)

    # Создаем метку и поле ввода для sender
    sender_label = tk.Label(root, text="Отправитель", font=("Helvetica", FONT))
    sender_label.pack()
    sender_entry = tk.Entry(root, width=50, font=("Helvetica", FONT))
    sender_entry.pack(pady=10, padx=10)

    # Создаем метку и поле ввода для agent
    agent_label = tk.Label(root, text="Контрагент", font=("Helvetica", FONT))
    agent_label.pack()
    agent_entry = tk.Entry(root, width=50, font=("Helvetica", FONT))
    agent_entry.pack(pady=10, padx=10)

    # Создаем метку и поле ввода для content
    content_label = tk.Label(root, text="Краткое содержание", font=("Helvetica", FONT))
    content_label.pack()
    content_entry = tk.Entry(root, width=50, font=("Helvetica", FONT))
    content_entry.pack(pady=10, padx=10)

    # Создаем кнопку
    save_button = tk.Button(root, text="Отправить", command=add_all, font=("Helvetica", FONT))  # Изменяем шрифт кнопки
    save_button.pack()

    result_label = tk.Label(root, text="", font=("Helvetica", 16))
    result_label.pack(pady=5, padx=10)

    error_label = tk.Label(root, text="", font=("Helvetica", 16))
    error_label.pack(pady=0, padx=10)

    love_label = tk.Label(root, text="", font=("Helvetica", 16))
    love_label.pack(pady=0, padx=10)

    # Запускаем цикл обработки событий
    root.bind("<Control-KeyPress>",keys)
    root.bind('<Return>',keys)
    root.mainloop()






if __name__ == '__main__':
    main()

