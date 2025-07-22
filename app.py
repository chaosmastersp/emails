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
    """Decodifica o assunto do e-mail de diferentes encodings."""
    if raw_subject is None:
        return ""
    try:
        decoded_parts = decode_header(raw_subject)
        subject = ""
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                # Tenta decodificar com o encoding especificado ou utf-8 como fallback
                subject += part.decode(encoding or "utf-8", errors="ignore")
            else:
                subject += str(part) # Garante que seja string para concatenação
        return subject.strip()
    except Exception as e:
        st.warning(f"Não foi possível decodificar o assunto '{raw_subject}': {e}. Retornando string vazia.")
        return ""

if aba == "Verificação de E-mails":
    st.title("📬 Verificador de E-mails Recebidos")

    data_ref_verificacao = st.date_input("Selecionar data de verificação", value=datetime.now() - timedelta(days=1))
    data_ref_format_imap = data_ref_verificacao.strftime("%d-%b-%Y")
    nome_arquivo = f"registros_nao/{data_ref_verificacao.strftime('%Y-%m-%d')}.csv"

    try:
        df_esperados = pd.read_excel("emails_esperados.xlsx")
        df_esperados.columns = df_esperados.columns.str.strip()
        # Garante que as colunas essenciais existam e sejam strings
        for col in ["Remetente", "Palavra-chave"]:
            if col not in df_esperados.columns:
                st.error(f"Coluna '{col}' não encontrada em 'emails_esperados.xlsx'. Verifique o arquivo.")
                st.stop()
            df_esperados[col] = df_esperados[col].astype(str).fillna("")
    except FileNotFoundError:
        st.error("Arquivo 'emails_esperados.xlsx' não encontrado. Por favor, crie o arquivo.")
        st.stop()
    except Exception as e:
        st.error(f"Erro ao ler 'emails_esperados.xlsx': {str(e)}")
        st.stop()

    email_user = st.secrets.get("email_user")
    email_pass = st.secrets.get("email_pass")
    imap_server = st.secrets.get("imap_server")

    if not all([email_user, email_pass, imap_server]):
        st.error("Credenciais de e-mail (email_user, email_pass, imap_server) não configuradas nos secrets do Streamlit.")
        st.stop()

    recebidos = []
    try:
        context = ssl.create_default_context()
        # Tenta com SECLEVEL=2 para maior compatibilidade com servidores modernos
        context.set_ciphers('DEFAULT@SECLEVEL=2')
        mail = imaplib.IMAP4_SSL(imap_server, ssl_context=context)
        mail.login(email_user, email_pass)
        mail.select("inbox")

        status, dados = mail.search(None, f'(ON "{data_ref_format_imap}")')
        ids = dados[0].split()

        if not ids:
            st.info(f"Nenhum e-mail encontrado para a data {data_ref_verificacao.strftime('%d/%m/%Y')}.")
        else:
            progress_bar = st.progress(0)
            total_emails = len(ids)
            for i, num in enumerate(ids):
                try:
                    status, dados = mail.fetch(num, '(RFC822)')
                    raw_email = dados[0][1]
                    msg = email.message_from_bytes(raw_email)

                    remetente_raw = msg.get("From", "Desconhecido").strip()
                    assunto_raw = msg.get("Subject", "").strip()

                    # Decodifica remetente e assunto, garantindo que são strings
                    # O email.utils.parseaddr pode ajudar a extrair o endereço de e-mail do campo 'From'
                    try:
                        _, remetente_email = email.utils.parseaddr(remetente_raw)
                        remetente = remetente_email if remetente_email else remetente_raw
                    except Exception:
                        remetente = remetente_raw # Se falhar, usa o raw

                    assunto = decodificar_assunto(assunto_raw)

                    recebidos.append({"Remetente": remetente, "Assunto": assunto})
                except Exception as e_fetch:
                    st.warning(f"Erro ao processar e-mail ID {num}: {str(e_fetch)}. Pulando este e-mail.")
                progress_bar.progress((i + 1) / total_emails)
            progress_bar.empty() # Remove a barra de progresso após a conclusão

        mail.logout()

        df_recebidos = pd.DataFrame(recebidos)
        # Garante que as colunas Remetente e Assunto sejam strings para evitar erros posteriores
        if not df_recebidos.empty:
            df_recebidos["Remetente"] = df_recebidos["Remetente"].astype(str).fillna("")
            df_recebidos["Assunto"] = df_recebidos["Assunto"].astype(str).fillna("")
        else:
            df_recebidos = pd.DataFrame(columns=["Remetente", "Assunto"]) # Cria um DataFrame vazio com as colunas esperadas

        if not df_recebidos.empty:
            resumo = df_recebidos.groupby("Remetente").size().reset_index(name="Quantidade")
            st.subheader("📊 Resumo de E-mails Recebidos")
            st.dataframe(resumo, use_container_width=True)
        else:
            resumo = pd.DataFrame(columns=["Remetente", "Quantidade"])
            st.warning("Nenhum e-mail recebido ou processado para a data selecionada.")

        resultado = []
        for _, row in df_esperados.iterrows():
            esperado_remetente = str(row["Remetente"]).strip()
            palavra_chave = str(row["Palavra-chave"]).strip()

            if not df_recebidos.empty:
                # Usar regex=False para evitar que a palavra-chave seja interpretada como regex,
                # a menos que seja intencional. Isso previne erros com caracteres especiais.
                filtro = df_recebidos[
                    df_recebidos["Remetente"].str.contains(esperado_remetente, case=False, na=False, regex=False) &
                    df_recebidos["Assunto"].str.contains(palavra_chave, case=False, na=False, regex=False)
                ]
                recebido_ontem = "✅ Sim" if not filtro.empty else "❌ Não"
            else:
                recebido_ontem = "❌ Não (Nenhum e-mail recebido)"


            resultado.append({
                "Remetente Esperado": esperado_remetente,
                "Palavra-chave": palavra_chave,
                "Recebido Ontem": recebido_ontem
            })

        df_resultado = pd.DataFrame(resultado)
        st.subheader("📥 Status dos E-mails Esperados")
        st.dataframe(df_resultado, use_container_width=True)

        df_nao = df_resultado[df_resultado["Recebido Ontem"].str.contains("❌ Não")][["Remetente Esperado", "Palavra-chave"]]
        st.session_state["resultado_nao"] = df_nao

        if st.button("💾 Salvar '❌ Não' para esta data"):
            if not os.path.exists("registros_nao"):
                os.makedirs("registros_nao")
            if not df_nao.empty:
                df_nao.to_csv(nome_arquivo, index=False)
                st.success(f"Registros salvos para {data_ref_verificacao.strftime('%d/%m/%Y')}")
            else:
                st.info("Nenhum registro '❌ Não' encontrado para salvar.")

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_resultado.to_excel(writer, sheet_name='Status', index=False)
            if not resumo.empty: # Garante que o resumo não esteja vazio antes de tentar salvar
                resumo.to_excel(writer, sheet_name='Resumo', index=False)
            else:
                # Opcional: Adicionar uma folha vazia ou mensagem se o resumo estiver vazio
                pd.DataFrame({"Mensagem": ["Nenhum resumo disponível."]}, index=[0]).to_excel(writer, sheet_name='Resumo', index=False)

        st.download_button("📁 Baixar Resultado em Excel", data=buffer.getvalue(), file_name="resultado_emails.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except imaplib.IMAP4.error as e:
        st.error(f"Erro de conexão ou autenticação IMAP: {str(e)}. Verifique suas credenciais e o servidor IMAP.")
    except Exception as e:
        st.error(f"Ocorreu um erro inesperado: {str(e)}. Por favor, contate o suporte.")

elif aba == "Registro de Ausências":
    st.title("📅 Registro de Ausências (❌ Não)")
    data_ref = st.date_input("Selecionar data de referência", value=datetime(2025, 7, 21)) # A data default pode ser o dia atual ou anterior
    nome_arquivo = f"registros_nao/{data_ref.strftime('%Y-%m-%d')}.csv"

    if not os.path.exists("registros_nao"):
        os.makedirs("registros_nao")

    if st.button("📥 Carregar Registros"):
        try:
            df_nao = pd.read_csv(nome_arquivo)
            if not df_nao.empty:
                st.success(f"Registros de {data_ref.strftime('%d/%m/%Y')} carregados com sucesso.")
                st.dataframe(df_nao, use_container_width=True)
            else:
                st.info("Nenhum registro encontrado para esta data ou o arquivo está vazio.")
        except FileNotFoundError:
            st.warning("Nenhum registro de ausência encontrado para esta data.")
            st.dataframe(pd.DataFrame(columns=["Remetente Esperado", "Palavra-chave"]), use_container_width=True) # Exibe DataFrame vazio
        except Exception as e:
            st.error(f"Erro ao carregar registros: {str(e)}")
