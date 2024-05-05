import pyautogui
import time

pyautogui.PAUSE = 0.7
i_arq = 1

# Caminho do PDF e da planilha
caminho_pdf = 'COMPROVANTES.pdf'
caminho_planilha = 'LISTA ALFABETICA.xlsx'

# Abrir o PDF
pyautogui.press('winleft')
pyautogui.write('Adobe Acrobat')
pyautogui.press('enter')
time.sleep(3)  # Espera 2 segundos para o Acrobat Reader abrir completamente
pyautogui.hotkey('ctrl', 'o')
time.sleep(2)
pyautogui.write(caminho_pdf)
pyautogui.press('enter')

# Abrir a planilha
pyautogui.press('winleft')
pyautogui.write('Excel')
pyautogui.press('enter')
time.sleep(5)  # Espera 2 segundos para o Excel abrir completamente
#pyautogui.press('tab')
pyautogui.click(x=68, y=402)
time.sleep(3)
pyautogui.click(x=262, y=498)
time.sleep(2)
pyautogui.write(caminho_planilha)
pyautogui.press('enter')
pyautogui.click()

# Dar um ctrl+home
pyautogui.hotkey('ctrl', 'home')

for j in range(10):
    # Dar um ctrl+c
    pyautogui.hotkey('ctrl', 'c')

    # Dar um alt+tab para mudar de janela
    pyautogui.hotkey('alt', 'tab')

    pyautogui.hotkey('ctrl', 'f')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.hotkey('ctrl', 'p')

    for i in range(7):
        pyautogui.press('tab')

    pyautogui.press('right')
    pyautogui.press('enter')
    time.sleep(2)

    pyautogui.write(str(i_arq) + ' - ')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    pyautogui.sleep(2)

    pyautogui.hotkey('alt', 'tab')
    pyautogui.sleep(1)
    pyautogui.press('down')

    i_arq += 1
