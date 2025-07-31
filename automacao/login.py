import pyautogui
import time
import os

pyautogui.PAUSE = 1

def abrir_sistema():
    erp = r"C:\MDOTech\X36SP04\bin-mdotech\start.bat"
    os.startfile(erp)
    time.sleep(15)
    pyautogui.hotkey('win', 'up')
    
def reiniciar_sistema():
    os.system("taskkill /f /im javaw.exe")
    time.sleep(5)
    abrir_sistema()


def realizar_login(login, senha):
    #pyautogui.click(369, 790)
    pyautogui.write(login)
    pyautogui.press('tab')
    pyautogui.write(senha)
    pyautogui.click(pyautogui.locateCenterOnScreen("imagens/btnEntrar.png", confidence=0.7))
    
    print("Login realizado.")
    time.sleep(5)