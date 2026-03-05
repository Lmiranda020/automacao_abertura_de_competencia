import pyautogui
from dotenv import load_dotenv
from api.api_competencia import api_competencia
from modules.clicar_na_imagem import clicar_imagem
from modules.enviar_email import enviar_email_competencia
from modules.log_execucao import iniciar_log, salvar_log
from utils.calcula_competencia import calcular_competencia_anterior
from utils.unidades import UNIDADES
import os
import time
import pyperclip
import pandas as pd

if __name__ == "__main__":

    # Registrar início da execução
    inicio_execucao = iniciar_log()

    ano, mes, competencia_formatada = calcular_competencia_anterior()
    print(f"\n📆 Competência a processar: {competencia_formatada}")

    caminho_base_api = api_competencia()

    # Ler arquivo da API
    df_api = pd.read_excel(caminho_base_api)

    # Filtrar apenas registros da competência atual (ano e mês)
    print(f"\n🔍 Filtrando competências para o ano: {ano} e mês: {mes}...")
    df_api = df_api[(df_api['ano'] == ano) & (df_api['mes'] == mes)]

    # Filtrar apenas as que estão com situação ABERTA
    # (essas são as que NÃO precisam ser processadas)
    df_ja_abertas = df_api[df_api['situacao'] == 'ABERTA']

    # Renomear unidades com nome diferente na API para bater com a lista do projeto
    df_ja_abertas['nome'] = df_ja_abertas['nome'].replace({
        "Filial 40 - HOSPITAL DIA M´BOI MIRIM I - CEJAM": "Filial 40 - HOSPITAL DIA M BOI MIRIM I - CEJAM",
        "Filial 40 - CENTRO REF A DOR CRON PQ MARIA HELENA - CEJAM": "Filial 40 - CENTRO REF A DOR CRON PQ MARIA HELENA - CE",
        "Filial 40 - HOSPITAL DIA M´BOI MIRIM II - CEJAM": "Filial 40 - HOSPITAL DIA M BOI MIRIM II - CEJAM",
    })

    # Montar lista de nomes das unidades que já estão com a competência aberta
    unidades_ja_abertas = df_ja_abertas['nome'].tolist()
    print(f"\n⏭️  Unidades com competência {competencia_formatada} já aberta (serão puladas): {len(unidades_ja_abertas)}")
    if unidades_ja_abertas:
        for u in unidades_ja_abertas:
            print(f"   - {u}")

    # Remover da lista do projeto as unidades que já estão abertas
    unidades_para_processar = [u for u in UNIDADES if u not in unidades_ja_abertas]
    print(f"\n✅ Unidades a processar: {len(unidades_para_processar)}")

    # Rastrear quais unidades foram abertas com sucesso nessa execução
    unidades_abertas_agora = []

    if not unidades_para_processar:
        print("Nenhuma unidade precisa ter a competência aberta.")
        enviar_email_competencia(competencia_formatada, unidades_abertas_agora, unidades_ja_abertas)
        salvar_log(inicio_execucao, competencia_formatada, unidades_abertas_agora, unidades_ja_abertas, status="Sucesso")
        exit(0)

    # Carregar variáveis do arquivo .env
    load_dotenv()

    # Abrir o Google Chrome
    pyautogui.press('win')
    pyautogui.write('chrome')
    pyautogui.press('enter')

    time.sleep(5)

    if not clicar_imagem('data/usuario_chrome.png', confidence=0.9, timeout=15, descricao="Usuário do Chrome"):
        print("Não foi possível selecionar o usuário do Chrome.")
        salvar_log(inicio_execucao, competencia_formatada, unidades_abertas_agora, unidades_ja_abertas, status="Erro")
        exit(1)

    time.sleep(5)

    pyautogui.hotkey('ctrl', 'l')
    pyautogui.press('backspace')
    SISTEMA = os.getenv("IP_SISTEMA")
    pyautogui.write(SISTEMA)
    pyautogui.press('enter')

    time.sleep(5)

    EMAIL: str = os.getenv("EMAIL")
    pyautogui.write(EMAIL)
    pyautogui.press('tab')

    SENHA: str = os.getenv("SENHA")
    pyautogui.write(SENHA)
    pyautogui.press('enter')

    time.sleep(5)

    if not clicar_imagem('data/campo_busca_unidade.png', confidence=0.8, timeout=15, descricao="Campo de Busca da Unidade"):
        print("Não foi possível encontrar o campo de busca da unidade.")
        salvar_log(inicio_execucao, competencia_formatada, unidades_abertas_agora, unidades_ja_abertas, status="Erro")
        exit(1)

    for unidade in unidades_para_processar:
        print(f"\n🔄 Processando unidade: {unidade}")

        pyperclip.copy(unidade)
        time.sleep(0.1)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.1)

        time.sleep(2)
        pyautogui.press('down')
        pyautogui.press('enter')
        pyautogui.press('enter')

        time.sleep(3)

        clicar_imagem('data/fechar_recomendacoes.png', confidence=0.8, timeout=5, descricao="Fechar Recomendações")
        time.sleep(1)

        if not clicar_imagem('data/botao_buscar.png', confidence=0.8, timeout=15, descricao="Botão Buscar"):
            print(f"Não foi possível encontrar o botão buscar para a unidade: {unidade}.")
            salvar_log(inicio_execucao, competencia_formatada, unidades_abertas_agora, unidades_ja_abertas, status="Parcial")
            exit(1)

        opcion_competencia = 'Abrir nova competência'
        pyperclip.copy(opcion_competencia)
        time.sleep(0.1)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.1)
        pyautogui.press('enter')

        if not clicar_imagem('data/abrir_nova_competencia.png', confidence=0.8, timeout=15, descricao="Abrir Nova Competência"):
            print(f"Não foi possível encontrar a opção de abrir nova competência para a unidade: {unidade}.")
            salvar_log(inicio_execucao, competencia_formatada, unidades_abertas_agora, unidades_ja_abertas, status="Parcial")
            exit(1)

        time.sleep(3)

        if not clicar_imagem('data/selecionar_email.png', confidence=0.8, timeout=15, descricao="Selecionar Email"):
            print(f"Não foi possível selecionar o email para a unidade: {unidade}.")
            salvar_log(inicio_execucao, competencia_formatada, unidades_abertas_agora, unidades_ja_abertas, status="Parcial")
            exit(1)

        time.sleep(3)

        if not clicar_imagem('data/trocar_unidade.png', confidence=0.8, timeout=15, descricao="Trocar Unidade"):
            print(f"Não foi possível clicar em trocar unidade para a unidade: {unidade}.")
            salvar_log(inicio_execucao, competencia_formatada, unidades_abertas_agora, unidades_ja_abertas, status="Parcial")
            exit(1)

        time.sleep(3)

        # Registrar unidade como aberta com sucesso nessa execução
        unidades_abertas_agora.append(unidade)

    # Finalizar a automação
    print("\nSaindo do sistema...")
    pyautogui.hotkey('ctrl', 'w')
    print("Automação finalizada com sucesso.")

    # Enviar e-mail com resumo da execução
    enviar_email_competencia(competencia_formatada, unidades_abertas_agora, unidades_ja_abertas)

    # Salvar log da execução
    salvar_log(inicio_execucao, competencia_formatada, unidades_abertas_agora, unidades_ja_abertas, status="Sucesso")