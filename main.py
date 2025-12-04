import pyautogui
from dotenv import load_dotenv
from modules.clicar_na_imagem import clicar_imagem
import os

if __name__ == "__main__":

    # Carregar variáveis do arquivo .env
    load_dotenv()

    # abrir o google chrome
    pyautogui.press('win')
    pyautogui.write('chrome')
    pyautogui.press('enter')

    # clicar em qual usuario o chrome deve abrir
    if not clicar_imagem('data/usuario_chrome.png', confidence=0.9, timeout=15, descricao="Usuário do Chrome"):
        print("Não foi possível selecionar o usuário do Chrome.")
        exit(1) # exit com 1 para indicar erro, e automação parar aqui. Se for 0 continua.
    
    # clicar na barra de endereços
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.press('backspace')
    SISTEMA = os.getenv("IP_SISTEMA")
    pyautogui.write(SISTEMA)
    pyautogui.press('enter')

    


    



