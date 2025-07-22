
import streamlit as st
import pandas as pd
import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import ssl
import io
import os

st.set_page_config(page_title="Verificador de E-mails", layout="wide")

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.sidebar.title("🔐 Acesso Restrito")
    usuario = st.sidebar.text_input("Usuário")
    senha = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        if usuario == st.secrets["auth_user"] and senha == st.secrets["auth_pass"]:
            st.session_state.autenticado = True
        else:
            st.sidebar.error("Credenciais inválidas.")
else:
    st.sidebar.success("✅ Acesso autorizado")

if not st.session_state.autenticado:
    st.stop()

if "resultado_nao" not in st.session_state:
    st.session_state["resultado_nao"] = pd.DataFrame()

aba = st.sidebar.selectbox("📌 Menu", ["Verificação de E-mails", "Registro de Ausências"])

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

if aba == "Verificação de E-mails":
    st.title("📬 Verificador de E-mails Recebidos")

    data_ref_verificacao = st.date_input("Selecionar data de verificação", value=datetime.now() - timedelta(days=1))
    data_ref_format_imap = data_ref_verificacao.strftime("%d-%b-%Y")
    nome_arquivo = f"registros_nao/{data_ref_verificacao.strftime('%Y-%m-%d')}.csv"

    df_esperados = pd.read_excel("emails_esperados.xlsx")
    df_esperados.columns = df_esperados.columns.str.strip()

    email_user = st.secrets["email_user"]
    email_pass = st.secrets["email_pass"]
    imap_server = st.secrets["imap_server"]

    try:
        context = ssl.create_default_context()
        context.set_ciphers('DEFAULT@SECLEVEL=1')
        mail = imaplib.IMAP4_SSL(imap_server, ssl_context=context)
        mail.login(email_user, email_pass)
        mail.select("inbox")

        status, dados = mail.search(None, f'(ON "{data_ref_format_imap}")')
        ids = dados[0].split()

        recebidos = []
        for num in ids:
            status, dados = mail.fetch(num, '(RFC822)')
            raw_email = dados[0][1]
            msg = email.message_from_bytes(raw_email)
            remetente = msg["From"] if msg["From"] else "Desconhecido"
            assunto = decodificar_assunto(msg["Subject"])
            recebidos.append({"Remetente": remetente, "Assunto": assunto})

        df_recebidos = pd.DataFrame(recebidos) if recebidos else pd.DataFrame(columns=["Remetente", "Assunto"])

        if "Remetente" in df_recebidos.columns and not df_recebidos.empty:
            resumo = df_recebidos.groupby("Remetente").size().reset_index(name="Quantidade")
            st.subheader("📊 Resumo de E-mails Recebidos")
            st.dataframe(resumo, use_container_width=True)
        else:
            resumo = pd.DataFrame(columns=["Remetente", "Quantidade"])
            st.warning("Nenhum e-mail encontrado ou erro ao processar remetentes.")

        resultado = []
        for _, row in df_esperados.iterrows():
            esperado_remetente = str(row["Remetente"]) if pd.notna(row["Remetente"]) else ""
            esperado_remetente = esperado_remetente.strip()
            palavra_chave = str(row["Palavra-chave"]) if pd.notna(row["Palavra-chave"]) else ""
            palavra_chave = palavra_chave.strip()

            if "Remetente" not in df_recebidos.columns:
                df_recebidos["Remetente"] = ""
            if "Assunto" not in df_recebidos.columns:
                df_recebidos["Assunto"] = ""

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

        df_nao = df_resultado[df_resultado["Recebido Ontem"] == "❌ Não"][["Remetente Esperado", "Palavra-chave"]]
        st.session_state["resultado_nao"] = df_nao

        if st.button("💾 Salvar '❌ Não' para esta data"):
            if not os.path.exists("registros_nao"):
                os.makedirs("registros_nao")
            if not df_nao.empty:
                df_nao.to_csv(nome_arquivo, index=False)
                st.success(f"Registros salvos para {data_ref_verificacao.strftime('%d/%m/%Y')}")
            else:
                st.warning("Nenhum registro ❌ Não encontrado para salvar.")

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_resultado.to_excel(writer, sheet_name='Status', index=False)
            resumo.to_excel(writer, sheet_name='Resumo', index=False)
        st.download_button("📁 Baixar Resultado em Excel", data=buffer.getvalue(), file_name="resultado_emails.xlsx")

    except Exception as e:
        st.error(f"Erro ao conectar ou processar e-mails: {str(e)}")

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
