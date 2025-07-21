import streamlit as st
import pandas as pd
import imaplib
import email
from datetime import datetime, timedelta
import ssl
import io

st.set_page_config(page_title="Verificador de E-mails", layout="wide")

st.title("📬 Verificador de E-mails Recebidos (Dia Anterior)")

# Carrega a planilha fixa com os e-mails esperados
df_esperados = pd.read_excel("emails_esperados.xlsx")

# Lê as credenciais seguras do secrets.toml
email_user = st.secrets["email_user"]
email_pass = st.secrets["email_pass"]
imap_server = st.secrets["imap_server"]

# Conecta ao servidor IMAP com segurança relaxada (caso necessário)
try:
    context = ssl.create_default_context()
    context.set_ciphers('DEFAULT@SECLEVEL=1')
    mail = imaplib.IMAP4_SSL(imap_server, ssl_context=context)
    mail.login(email_user, email_pass)
    mail.select("inbox")

    # Data de ontem no formato IMAP
    ontem = (datetime.now() - timedelta(days=1)).strftime("%d-%b-%Y")
    status, dados = mail.search(None, f'(ON "{ontem}")')
    ids = dados[0].split()

    recebidos = []
    for num in ids:
        status, dados = mail.fetch(num, '(RFC822)')
        raw_email = dados[0][1]
        msg = email.message_from_bytes(raw_email)
        remetente = msg["From"]
        assunto = msg["Subject"]
        recebidos.append({"Remetente": remetente, "Assunto": assunto})

    df_recebidos = pd.DataFrame(recebidos)

    # Resumo por remetente
    resumo = df_recebidos.groupby(["Remetente", "Assunto"]).size().reset_index(name="Quantidade")
    st.subheader("📊 Resumo de E-mails Recebidos")
    st.dataframe(resumo, use_container_width=True)

    # Verificação de recebimento esperado
    resultado = []
    for _, row in df_esperados.iterrows():
        esperado_remetente = row["Remetente"]
        palavra_chave = row["Palavra-chave"]
        filtro = df_recebidos[
            df_recebidos["Remetente"].str.contains(esperado_remetente, case=False, na=False) &
            df_recebidos["Assunto"].str.contains(palavra_chave, case=False, na=False)
        ]
        resultado.append({
            "Remetente Esperado": esperado_remetente,
            "Palavra-chave": palavra_chave,
            "Recebido Ontem": "✅ Sim" if not filtro.empty else "❌ Não"
        })

    df_resultado = pd.DataFrame(resultado)
    st.subheader("📥 Status dos E-mails Esperados")
    st.dataframe(df_resultado, use_container_width=True)

    # Exportação do resultado
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_resultado.to_excel(writer, sheet_name='Status', index=False)
        resumo.to_excel(writer, sheet_name='Resumo', index=False)
    st.download_button("📁 Baixar Resultado em Excel", data=buffer.getvalue(), file_name="resultado_emails.xlsx")

except Exception as e:
    st.error(f"Erro ao conectar ou processar e-mails: {str(e)}")
