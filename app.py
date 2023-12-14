import pyautogui
import openpyxl
import time
import pyperclip

planilha = openpyxl.load_workbook('BIANVIRC.xlsx')
info_contatos = planilha['Planilha1']

for linha in info_contatos.iter_rows(min_row=3):
    pyautogui.click(1343, 167, duration=1)
    time.sleep(1)
    nome = linha[0].value
    pyperclip.copy(nome)
    pyautogui.click(234, 237, duration=1)
    pyautogui.hotkey("ctrl", "v")
    pis = linha[3].value
    pyperclip.copy(pis)
    pyautogui.click(692, 235, duration=1)
    pyautogui.hotkey("ctrl", "v")
    cpf = linha[4].value
    pyperclip.copy(cpf)
    pyautogui.click(905, 235, duration=1)
    pyautogui.hotkey("ctrl", "v")
    matricula = linha[2].value
    pyperclip.copy(matricula)
    pyautogui.click(237, 291, duration=1)
    pyautogui.hotkey("ctrl", "v")
    data = linha[5].value
    pyautogui.click(307, 341, duration=1)
    pyautogui.typewrite(str(data))

    pyautogui.click(721, 342)
    time.sleep(1)
    pyautogui.click(689, 359)
    time.sleep(1)
    pyautogui.click(180, 711)
    time.sleep(3)
