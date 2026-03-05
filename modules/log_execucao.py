import os
import pandas as pd
from datetime import datetime

# Pasta raiz do projeto
PASTA_PROJETO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CAMINHO_LOG = os.path.join(PASTA_PROJETO, "log_execucoes.xlsx")

def iniciar_log() -> datetime:
    """Registra o horário de início e retorna o datetime para cálculo posterior."""
    inicio = datetime.now()
    print(f"\n🕐 Execução iniciada em: {inicio.strftime('%d/%m/%Y %H:%M:%S')}")
    return inicio


def salvar_log(
    inicio: datetime,
    competencia_formatada: str,
    unidades_abertas: list,
    unidades_ja_abertas: list,
    status: str = "Sucesso"
):
    """
    Salva/atualiza o log incremental de execuções.
    Cada chamada adiciona uma nova linha ao arquivo Excel.

    status: "Sucesso" | "Parcial" | "Erro"
    """
    fim = datetime.now()
    duracao_seg = (fim - inicio).seconds
    minutos, segundos = divmod(duracao_seg, 60)

    nova_linha = {
        "Data":                     inicio.strftime("%d/%m/%Y"),
        "Início":                   inicio.strftime("%H:%M:%S"),
        "Término":                  fim.strftime("%H:%M:%S"),
        "Duração":                  f"{minutos}min {segundos}s",
        "Competência":              competencia_formatada,
        "Unidades Abertas Agora":   len(unidades_abertas),
        "Unidades Já Abertas":      len(unidades_ja_abertas),
        "Total Processado":         len(unidades_abertas) + len(unidades_ja_abertas),
        "Detalhe Abertas Agora":    " | ".join(unidades_abertas) if unidades_abertas else "—",
        "Detalhe Já Abertas":       " | ".join(unidades_ja_abertas) if unidades_ja_abertas else "—",
        "Status":                   status,
    }

    # Carrega log existente ou cria novo
    if os.path.exists(CAMINHO_LOG):
        try:
            df_log = pd.read_excel(CAMINHO_LOG)
        except Exception:
            df_log = pd.DataFrame()
    else:
        df_log = pd.DataFrame()

    df_log = pd.concat([df_log, pd.DataFrame([nova_linha])], ignore_index=True)

    try:
        df_log.to_excel(CAMINHO_LOG, index=False)
        print(f"📋 Log salvo em: {CAMINHO_LOG}")
        print(f"⏱️  Duração total: {minutos}min {segundos}s")
        print(f"📊 Resumo: {len(unidades_abertas)} abertas agora | {len(unidades_ja_abertas)} já estavam abertas")
    except Exception as e:
        print(f"❌ Erro ao salvar log: {e}")