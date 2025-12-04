import pyautogui
from dotenv import load_dotenv
from modules.clicar_na_imagem import clicar_imagem
from utils.unidades import UNIDADES
import os
import time
import pyperclip

if __name__ == "__main__":

    # Carregar variáveis do arquivo .env
    load_dotenv()

    # abrir o google chrome
    pyautogui.press('win')
    pyautogui.write('chrome')
    pyautogui.press('enter')

    # aguardar o chrome abrir
    time.sleep(5)

    # clicar em qual usuario o chrome deve abrir
    if not clicar_imagem('data/usuario_chrome.png', confidence=0.9, timeout=15, descricao="Usuário do Chrome"):
        print("Não foi possível selecionar o usuário do Chrome.")
        exit(1) # exit com 1 para indicar erro, e automação parar aqui. Se for 0 continua.

    # aguardar o chrome abrir com o usuario selecionado
    time.sleep(5)
    
    # clicar na barra de endereços
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.press('backspace')
    SISTEMA = os.getenv("IP_SISTEMA")
    pyautogui.write(SISTEMA)
    pyautogui.press('enter')

    # aguardar a página carregar
    time.sleep(5)

    # digitar o email para login
    EMAIL: str = os.getenv("EMAIL")
    pyautogui.write(EMAIL)
    pyautogui.press('tab')  

    # digitar a senha para login
    SENHA: str = os.getenv("SENHA")
    pyautogui.write(SENHA)
    pyautogui.press('enter')

    # aguardar a página carregar
    time.sleep(5)

    # clicar o campo de busca da unidade
    if not clicar_imagem('data/campo_busca_unidade.png', confidence=0.8, timeout=15, descricao="Campo de Busca da Unidade"):
        print("Não foi possível encontrar o campo de busca da unidade.")
        exit(1)

    # digitar a unidade desejada
    for unidade in UNIDADES:
        pyperclip.copy(unidade)  # copia o texto com acentos
        time.sleep(0.1)
        pyautogui.hotkey("ctrl", "v")  # cola tudo corretamente
        time.sleep(0.1)

        time.sleep(2)  # aguardar as sugestões aparecerem
        pyautogui.press('down')  # selecionar a primeira sugestão
        pyautogui.press('enter')
        pyautogui.press('enter')

        time.sleep(3)  # aguardar a unidade carregar

        # caso tenha a tela de recomendações, fechar ela
        clicar_imagem('data/fechar_recomendacoes.png', confidence=0.8, timeout=5, descricao="Fechar Recomendações")
        time.sleep(1)

        # clica na opção buscar
        if not clicar_imagem('data/botao_buscar.png', confidence=0.8, timeout=15, descricao="Botão Buscar"):
            print(f"Não foi possível encontrar o botão buscar para a unidade: {unidade}.")
            exit(1) # sair

        # digitar a opção de competência
        opcion_competencia = 'Abrir nova competência'
        pyperclip.copy(opcion_competencia)  # copia o texto com acentos
        time.sleep(0.1)
        pyautogui.hotkey("ctrl", "v")  # cola tudo corretamente
        time.sleep(0.1)
        pyautogui.press('enter')

        # selecionar a opção de abrir nova competência
        if not clicar_imagem('data/abrir_nova_competencia.png', confidence=0.8, timeout=15, descricao="Abrir Nova Competência"):
            print(f"Não foi possível encontrar a opção de abrir nova competência para a unidade: {unidade}.")
            exit(1) # sair  

        # selecionar o email
        if not clicar_imagem('data/selecionar_email.png', confidence=0.8, timeout=15, descricao="Selecionar Email"):
            print(f"Não foi possível selecionar o email para a unidade: {unidade}.")
            exit(1) # sair

        # clicar na opção trocar unidade
        if not clicar_imagem('data/trocar_unidade.png', confidence=0.8, timeout=15, descricao="Trocar Unidade"):
            print(f"Não foi possível clicar em trocar unidade para a unidade: {unidade}.")
            exit(1) # sair

        time.sleep(3)  # aguardar a troca de unidade

    # finalizar a automação
    print("Saindo do sistema...")
    
    pyautogui.hotkey('ctrl', 'w')

    print("Automação finalizada com sucesso.")

    
        
        
        
        

        



