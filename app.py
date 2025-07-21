import streamlit as st
import pandas as pd
import imaplib
import email
from email.utils import parseaddr
from email.header import decode_header
from datetime import datetime, timedelta
import ssl
import io

st.set_page_config(page_title="Verificador de E-mails", layout="wide")

# Função de autenticação
def autenticar():
    st.sidebar.title("🔐 Acesso Restrito")
    usuario = st.sidebar.text_input("Usuário")
    senha = st.sidebar.text_input("Senha", type="password")
    if usuario == st.secrets["auth_user"] and senha == st.secrets["auth_pass"]:
        return True
    elif usuario and senha:
        st.sidebar.error("Credenciais inválidas.")
        return False
    else:
        return False

# Função para decodificar o campo Assunto
def decodificar_assunto(raw_subject):
    if raw_subject is None:
        return ""
    decoded_parts = decode_header(raw_subject)
    subject = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            subject += part.decode(encoding or "utf-8", errors="ignore")
        else:
            subject += part
    return subject.strip()

# Bloqueia acesso até autenticar
if not autenticar():
    st.stop()

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
        remetente = parseaddr(msg["From"])[1]  # Extrai apenas o e-mail limpo
        assunto = decodificar_assunto(msg["Subject"])
        recebidos.append({"Remetente": remetente, "Assunto": assunto})

    df_recebidos = pd.DataFrame(recebidos)

    if not df_recebidos.empty:
        # Resumo por remetente (sem detalhar por assunto)
        resumo = df_recebidos.groupby("Remetente").size().reset_index(name="Quantidade")
        st.subheader("📊 Resumo de E-mails Recebidos")
        st.dataframe(resumo, use_container_width=True)
    else:
        resumo = pd.DataFrame(columns=["Remetente", "Quantidade"])
        st.warning("Nenhum e-mail encontrado na data de ontem.")

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
            "Recebido On
