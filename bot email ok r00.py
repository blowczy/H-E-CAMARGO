import os
import glob
import pyautogui
import time
import pandas as pd
import pyperclip

def non_empty_cells(df):
    filled_cells = df.dropna(how='all')
    return filled_cells.to_string(index=False)

def write_monitoring_info(buffer, filename):
    with open(filename, 'a') as f:
        f.writelines(buffer)

print("Iniciando o script...")

path = 'C:\\Users\\user\\Desktop\\cotações recebidas\\'
print(f"Monitorando a pasta: {path}")

fornecedores = [
    {'email': 'blowczy@gmail.com', 'message': 'Ola bruno bom dia ! tudo bem?\nEsta é uma requisição enviada de forma automática:'},
    {'email': 'brunolowczy@gmail.com', 'message': 'Ola lowczy bom dia ! tudo bem?\nEsta é uma requisição enviada de forma automática:'}
]

all_files = glob.glob(os.path.join(path, "*.xlsx"))

for file_name in all_files:
    print(f"Processando arquivo: {file_name}")

    xls = pd.ExcelFile(file_name)
    sheet_names = xls.sheet_names

    for sheet_name in sheet_names:
        print(f"Processando a folha: {sheet_name}")
        df = pd.read_excel(file_name, sheet_name=sheet_name)
        pyperclip.copy(non_empty_cells(df))

    monitoring_buffer = []
    process_monitoring_buffer = [f"Iniciando o processamento do arquivo: {file_name}\n"]

    xls.close()

    try:
        os.remove(file_name)
        monitoring_buffer.append(f"Arquivo {file_name} excluído em {time.strftime('%H:%M:%S %Y-%m-%d')}\n")
        process_monitoring_buffer.append(f"Arquivo {file_name} excluído.\n")
    except PermissionError:
        print(f"Não foi possível remover o arquivo {file_name} - arquivo em uso.")
        monitoring_buffer.append(f"Não foi possível remover o arquivo {file_name} - arquivo em uso.\n")
        process_monitoring_buffer.append(f"Não foi possível remover o arquivo {file_name} - arquivo em uso.\n")

    for fornecedor in fornecedores:
        print(f"Enviando e-mail para {fornecedor['email']}...")

        pyautogui.hotkey('ctrl', 'n')
        time.sleep(0.5)
        pyautogui.typewrite(fornecedor['email'])
        pyautogui.press('tab')
        time.sleep(0.5)

        for _ in range(4):
            pyautogui.press('tab')

        pyautogui.typewrite("O assunto do email aqui")
        pyautogui.press('tab')

        pyautogui.typewrite(fornecedor['message'])
        pyautogui.press('enter')
        time.sleep(0.5)

        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

        pyautogui.hotkey('ctrl', 'enter')

        monitoring_buffer.append(f"Enviado para: {fornecedor['email']} no arquivo {file_name} em {time.strftime('%H:%M:%S %Y-%m-%d')}\n")
        process_monitoring_buffer.append(f"Email enviado para {fornecedor['email']}.\n")

    write_monitoring_info(monitoring_buffer, os.path.join(path, os.path.basename(file_name).split('.')[0] + ".txt"))
    write_monitoring_info(process_monitoring_buffer, os.path.join(path, "monitoramento_de_processos.txt"))

print("Script concluído.")
