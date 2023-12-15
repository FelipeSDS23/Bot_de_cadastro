import pyautogui
from time import sleep
import pandas as pd

# 1 - Clicar e digitar usuário
pyautogui.click(2445,331,duration=1)
pyautogui.write('BOT')
# 2 - Clicar e digitar senha
pyautogui.click(2450,405,duration=1)
pyautogui.write('123')
# 3 - Clicar em entrar
pyautogui.click(2457,458,duration=1)
# 4 - Clicar em cadastrar
pyautogui.click(2657,127,duration=1)

# 5 - Extrair cada aluno da planilha
planilha = pd.read_excel('base.xlsx')

for indice, linha in planilha.iterrows():
    # 5.1 - Clicar e digitar o nome completo
    pyautogui.click(2333,301,duration=1)
    pyautogui.write(linha['Nome'])
    # 5.2 - Clicar e digitar o telefone
    pyautogui.click(2558,371,duration=1)
    pyautogui.write(linha['Telefone'])
    # 5.3 - Clicar e digitar o email
    pyautogui.click(2867,443,duration=1)
    pyautogui.write(linha['E-mail'])
    # 5.4 - Clicar em cadastrar 
    pyautogui.click(2599,493,duration=1)    