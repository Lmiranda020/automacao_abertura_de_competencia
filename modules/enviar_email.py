import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

def enviar_email_competencia(
    competencia_formatada: str,
    unidades_processadas: list,
    unidades_ja_abertas: list
):
    load_dotenv()

    remetente = os.getenv("EMAIL_REMETENTE")
    senha_app = os.getenv("SENHA_APP_MAIL")
    destinatario = os.getenv("EMAIL_FLUXO")

    if not all([remetente, senha_app, destinatario]):
        print("❌ Variáveis de e-mail não configuradas no .env (EMAIL_REMETENTE, SENHA_APP_MAIL, EMAIL_FLUXO).")
        return False

    # Monta competência legível ex: "02_2026" -> "Fevereiro/2026"
    meses_pt = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    try:
        mes_num, ano = competencia_formatada.split("_")
        mes_nome = meses_pt.get(int(mes_num), mes_num)
        competencia_legivel = f"{mes_nome}/{ano}"
    except Exception:
        competencia_legivel = competencia_formatada

    # Monta HTML das unidades processadas
    if unidades_processadas:
        itens_processadas = "".join(
            f"<li style='margin-bottom:4px;'>✅ {u}</li>" for u in unidades_processadas
        )
    else:
        itens_processadas = "<li>Nenhuma unidade foi processada.</li>"

    # Monta HTML das unidades já abertas
    if unidades_ja_abertas:
        itens_ja_abertas = "".join(
            f"<li style='margin-bottom:4px;'>⏭️ {u}</li>" for u in unidades_ja_abertas
        )
        bloco_ja_abertas = f"""
        <h3 style='color:#b45309; margin-top:24px;'>Unidades que já estavam com a competência aberta:</h3>
        <ul style='font-size:14px; color:#555;'>
            {itens_ja_abertas}
        </ul>
        """
    else:
        bloco_ja_abertas = ""

    corpo_html = f"""
    <html>
    <body style='font-family: Arial, sans-serif; color: #333; padding: 20px;'>
        <h2 style='color:#1d4ed8;'>📋 Abertura de Competências — {competencia_legivel}</h2>
        <p>Olá,</p>
        <p>
            A automação foi executada com sucesso. Seguem abaixo as informações 
            sobre a abertura de competências referente a <strong>{competencia_legivel}</strong>.
        </p>

        <h3 style='color:#15803d; margin-top:24px;'>Unidades com competência aberta nesta execução:</h3>
        <ul style='font-size:14px; color:#555;'>
            {itens_processadas}
        </ul>

        {bloco_ja_abertas}

        <hr style='margin-top:32px; border:none; border-top:1px solid #ddd;'/>
        <p style='font-size:12px; color:#999;'>E-mail enviado automaticamente pela automação de abertura de competências.</p>
    </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[Automação] Abertura de Competências — {competencia_legivel}"
    msg["From"] = remetente
    msg["To"] = destinatario
    msg.attach(MIMEText(corpo_html, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as servidor:
            servidor.login(remetente, senha_app)
            servidor.sendmail(remetente, destinatario, msg.as_string())
        print(f"📧 E-mail enviado com sucesso para: {destinatario}")
        return True
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")
        return False