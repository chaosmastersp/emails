import streamlit as st
import pandas as pd
import imaplib
import email
from email.utils import parseaddr
from email.header import decode_header
from datetime import datetime, timedelta
import ssl
import io
import os

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

# Menu lateral
aba = st.sidebar.selectbox("📌 Menu", ["Verificação de E-mails", "Registro de Ausências"])

# Variável global usada para exportar '❌ Não'
df_resultado = pd.DataFrame()

# Aba principal de verificação
if aba == "Verificação de E-mails":
    st.title("📬 Verificador de E-mails Recebidos (Dia Anterior)")

    # Carrega a planilha fixa com os e-mails esperados
    df_esperados = pd.read_excel("emails_esperados.xlsx")
    df_esperados.columns = df_esperados.columns.str.strip()

    # Lê as credenciais seguras do secrets.toml
    email_user = st.secrets["email_user"]
    email_pass = st.secrets["email_pass"]
    imap_server = st.secrets["imap_server"]

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
            assunto = decodificar_assunto(msg["Subject"])
            recebidos.append({"Remetente": remetente, "Assunto": assunto})

        if recebidos:
            df_recebidos = pd.DataFrame(recebidos)
        else:
            df_recebidos = pd.DataFrame(columns=["Remetente", "Assunto"])

        if not df_recebidos.empty:
            resumo = df_recebidos.groupby("Remetente").size().reset_index(name="Quantidade")
            st.subheader("📊 Resumo de E-mails Recebidos")
            st.dataframe(resumo, use_container_width=True)
        else:
            resumo = pd.DataFrame(columns=["Remetente", "Quantidade"])
            st.warning("Nenhum e-mail encontrado na data de ontem.")

        # Verificação de recebimento esperado
        resultado = []
        for _, row in df_esperados.iterrows():
            esperado_remetente = row.get("Remetente", "").strip()
            palavra_chave = row.get("Palavra-chave", "").strip()
            filtro = df_recebidos[
                df_recebidos["Remetente"].str.contains(esperado_remetente, case=False, na=False) &
                df_recebidos["Assunto"].str.contains(palavra_chave, case=False, na=False, regex=False)
            ]
            resultado.append({
                "Remetente Esperado": esperado_remetente,
                "Palavra-chave": palavra_chave,
                "Recebido Ontem": "✅ Sim" if not filtro.empty else "❌ Não"
            })

        df_resultado = pd.DataFrame(resultado)
        st.subheader("📥 Status dos E-mails Esperados")
        st.dataframe(df_resultado, use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_resultado.to_excel(writer, sheet_name='Status', index=False)
            resumo.to_excel(writer, sheet_name='Resumo', index=False)
        st.download_button("📁 Baixar Resultado em Excel", data=buffer.getvalue(), file_name="resultado_emails.xlsx")

    except Exception as e:
        st.error(f"Erro ao conectar ou processar e-mails: {str(e)}")

# Aba de registros '❌ Não'
elif aba == "Registro de Ausências":
    st.title("📅 Registro de Ausências (❌ Não)")

    data_ref = st.date_input("Selecionar data de referência", value=datetime(2025, 7, 21))
    nome_arquivo = f"registros_nao/{data_ref.strftime('%Y-%m-%d')}.csv"

    if not os.path.exists("registros_nao"):
        os.makedirs("registros_nao")

    if st.button("📥 Carregar Registros"):
        try:
            df_nao = pd.read_csv(nome_arquivo)
            st.success(f"Registros de {data_ref.strftime('%d/%m/%Y')} carregados com sucesso.")
        except FileNotFoundError:
            df_nao = pd.DataFrame(columns=["Remetente Esperado", "Palavra-chave"])
            st.warning("Nenhum registro encontrado para esta data.")
        st.dataframe(df_nao, use_container_width=True)

    if st.button("📤 Salvar '❌ Não' do último resultado"):
        if not df_resultado.empty:
            try:
                df_nao_salvar = df_resultado[df_resultado["Recebido Ontem"] == "❌ Não"][["Remetente Esperado", "Palavra-chave"]]
                df_nao_salvar.to_csv(nome_arquivo, index=False)
                st.success(f"Ausências salvas para {data_ref.strftime('%d/%m/%Y')}.")
            except Exception as e:
                st.error(f"Erro ao salvar ausências: {str(e)}")
        else:
            st.warning("Nenhum resultado disponível para salvar.")
