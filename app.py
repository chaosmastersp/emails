
import streamlit as st
import pandas as pd
import imaplib
import email
from email.header import decode_header
from email.utils import parseaddr
from datetime import datetime, timedelta

st.set_page_config(page_title="Verificador de E-mails Recebidos", layout="wide")
st.title("📬 Verificador de E-mails Recebidos")

# Seleção da data
data_referencia = st.date_input("Selecionar data de verificação", value=datetime.now().date() - timedelta(days=1))

def extrair_remetente(msg):
    raw = msg.get("From", "")
    return parseaddr(raw)[1] if raw else ""

def extrair_assunto(msg):
    assunto, codificacao = decode_header(msg.get("Subject", ""))[0]
    if isinstance(assunto, bytes):
        try:
            return assunto.decode(codificacao or "utf-8")
        except:
            return assunto.decode("latin1", errors="ignore")
    return assunto or ""

def conectar_e_obter_emails():
    try:
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login("seu_email@gmail.com", "sua_senha_segura")
        imap.select("inbox")

        data_formatada = data_referencia.strftime("%d-%b-%Y")
        status, mensagens = imap.search(None, f'SINCE {data_formatada}', f'BEFORE {(data_referencia + timedelta(days=1)).strftime("%d-%b-%Y")}')
        lista_ids = mensagens[0].split()

        registros = []
        for num in lista_ids:
            status, dados = imap.fetch(num, "(RFC822)")
            for response_part in dados:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    remetente = extrair_remetente(msg)
                    assunto = extrair_assunto(msg)
                    registros.append({"Remetente": remetente, "Assunto": assunto})

        imap.logout()
        return pd.DataFrame(registros)
    except Exception as e:
        st.error(f"Erro ao conectar ou processar e-mails: {e}")
        return pd.DataFrame()

# Botão para buscar e-mails
if st.button("📥 Carregar E-mails"):
    df_esperado = pd.read_excel("emails_esperados.xlsx")
    df_esperado["Remetente Esperado"] = df_esperado["Remetente Esperado"].fillna("").str.strip().str.lower()
    df_esperado["Palavra-chave"] = df_esperado["Palavra-chave"].fillna("")

    df_recebido = conectar_e_obter_emails()
    if "Remetente" not in df_recebido.columns:
        st.error("Erro: coluna 'Remetente' não encontrada nos dados recebidos.")
    else:
        df_recebido["Remetente"] = df_recebido["Remetente"].fillna("").str.strip().str.lower()
        df_recebido["Assunto"] = df_recebido["Assunto"].fillna("")

        def verificar(row):
            for _, esperado in df_esperado.iterrows():
                if row["Remetente"] == esperado["Remetente Esperado"] and                    esperado["Palavra-chave"].lower() in row["Assunto"].lower():
                    return "Sim"
            return "Não"

        df_recebido["Recebido"] = df_recebido.apply(verificar, axis=1)

        st.subheader("📊 Resumo de E-mails Recebidos")
        st.dataframe(df_recebido[["Remetente", "Assunto", "Recebido"]])

        faltantes = df_esperado[~df_esperado["Remetente Esperado"].isin(df_recebido[df_recebido["Recebido"] == "Sim"]["Remetente"])]

        if not faltantes.empty:
            st.warning("❌ E-mails esperados não recebidos:")
            st.dataframe(faltantes)
            faltantes.to_excel("nao_recebidos.xlsx", index=False)
            st.download_button("📤 Baixar Não Recebidos", data=open("nao_recebidos.xlsx", "rb"), file_name="nao_recebidos.xlsx")
