import streamlit as st
import smtplib
from email.message import EmailMessage
from pypdf import PdfReader, PdfWriter
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Envio de Holerites", page_icon="ensine icone 2025.png")

# --- CARREGAR CONFIGURAÇÕES DO COFRE (SECRETS) ---
try:
    # Tenta converter os segredos em dicionário para evitar erro de .copy()
    DB_FUNCIONARIOS = dict(st.secrets["funcionarios"])
    
    # Carrega dados do email fixo
    EMAIL_REMETENTE = st.secrets["config_email"]["email_fixo"]
    SENHA_REMETENTE = st.secrets["config_email"]["senha_fixa"]
    EMAIL_BCC = st.secrets["config_email"]["email_copia"]
    
except Exception as e:
    # Se der erro nos segredos, mostra aviso mas não trava totalmente (carrega dummy)
    st.error(f"Erro ao carregar segredos (Secrets): {e}")
    DB_FUNCIONARIOS = {} 
    # Define valores vazios para não quebrar o resto do código
    EMAIL_REMETENTE = ""
    SENHA_REMETENTE = ""
    EMAIL_BCC = ""

# --- INICIALIZA SESSÃO ---
if 'banco_dados' not in st.session_state:
    st.session_state['banco_dados'] = DB_FUNCIONARIOS.copy()

# --- FUNÇÃO DE ENVIO DE EMAIL (COM CÓPIA OCULTA) ---
def enviar_email_fixo(destinatario, assunto, corpo, anexo_bytes, nome_arquivo):
    msg = EmailMessage()
    msg['Subject'] = assunto
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = destinatario
    msg['Bcc'] = EMAIL_BCC # Cópia oculta para o Financeiro
    msg.set_content(corpo)

    msg.add_attachment(anexo_bytes, maintype='application', subtype='pdf', filename=nome_arquivo)

    # Configuração GMAIL (Para o e-mail remetente fixo)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_REMETENTE, SENHA_REMETENTE)
        smtp.send_message(msg)

# --- INTERFACE VISUAL ---
st.title("📧 Sistema de Envio de Holerites")
if EMAIL_REMETENTE:
    st.caption(f"Enviando através de: {EMAIL_REMETENTE}")
else:
    st.error("⚠️ Email remetente não configurado nos Secrets!")

st.markdown("---")

# Barra Lateral
with st.sidebar:
    st.header("ℹ️ Status do Sistema")
    if EMAIL_REMETENTE:
        st.success("✅ Login Automático Ativo")
        st.info(f"Cópia oculta configurada para:\n{EMAIL_BCC}")
    else:
        st.error("❌ Falta configurar Secrets")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload do Arquivo")
    arquivo_pdf = st.file_uploader("Solte o PDF aqui", type="pdf")

with col2:
    st.subheader("2. Mensagem")
    assunto_email = st.text_input("Assunto", value="Holerite - Pagamento")
    corpo_email = st.text_area("Texto", value="Segue em anexo seu holerite.\n\nAtenciosamente,\nFinanceiro - IESPE", height=100)

st.markdown("---")

# --- ÁREA DE ADICIONAR NOVO FUNCIONÁRIO ---
with st.expander("➕ Adicionar alguém fora da lista"):
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        novo_nome = st.text_input("Nome (Como está no PDF)")
    with c2:
        novo_email = st.text_input("E-mail do Funcionário")
    with c3:
        st.write("") 
        st.write("") 
        if st.button("Adicionar"):
            if novo_nome and novo_email:
                st.session_state['banco_dados'][novo_nome] = novo_email
                st.success(f"{novo_nome} adicionado!")
                st.rerun()

st.subheader("3. Seleção de Destinatários")

lista_atualizada = st.session_state['banco_dados']
nomes_selecionados = st.multiselect(
    "Quem vai receber?",
    options=list(lista_atualizada.keys()), 
    default=list(lista_atualizada.keys())
)

st.write(f"Emails serão enviados para **{len(nomes_selecionados)}** pessoas.")

# --- LÓGICA DO BOTÃO DE DISPARO ---
if st.button("🚀 Disparar Holerites", type="primary"):
    if not arquivo_pdf:
        st.error("Falta o arquivo PDF!")
    elif len(nomes_selecionados) == 0:
        st.warning("Selecione alguém.")
    elif not EMAIL_REMETENTE:
        st.error("Configuração de email inválida. Verifique os Secrets.")
    else:
        barra = st.progress(0)
        status = st.empty()
        cont = 0
        
        try:
            # Reseta o ponteiro do arquivo para garantir leitura do início
            arquivo_pdf.seek(0)
            leitor = PdfReader(arquivo_pdf)
            total = len(leitor.pages)
            
            for i, pagina in enumerate(leitor.pages):
                texto = pagina.extract_text()
                
                for nome in nomes_selecionados:
                    # Verifica se o nome está no texto (caixa alta para garantir)
                    if nome.upper() in texto.upper():
                        email_dest = st.session_state['banco_dados'][nome]
                        status.text(f"Enviando para: {nome}...")
                        
                        escritor = PdfWriter()
                        escritor.add_page(pagina)
                        
                        from io import BytesIO
                        pdf_bytes = BytesIO()
                        escritor.write(pdf_bytes)
                        
                        try:
                            enviar_email_fixo(
                                email_dest, 
                                assunto_email, 
                                corpo_email, 
                                pdf_bytes.getvalue(), 
                                f"Holerite_{nome}.pdf"
                            )
                            st.toast(f"✅ Enviado: {nome}")
                            cont += 1
                        except Exception as e:
                            st.error(f"Erro ao enviar para {nome}: {e}")
                
                barra.progress((i + 1) / total)
            
            st.success(f"Finalizado! {cont} holerites enviados.")
            status.empty()
            
        except Exception as e:
            st.error(f"Erro crítico no processamento: {e}")

# --- MODO ESPIÃO (FORA DO BOTÃO DE ENVIO) ---
st.markdown("---")
with st.expander("🔍 Modo Espião (Diagnóstico Completo)"):
    if arquivo_pdf:
        try:
            # 1. Reseta o arquivo para o inicio
            arquivo_pdf.seek(0)
            leitor_debug = PdfReader(arquivo_pdf)
            num_paginas = len(leitor_debug.pages)
            
            st.info(f"📊 O robô detectou **{num_paginas} páginas** neste arquivo.")
            st.info("Abaixo mostro o que consigo ler. Se estiver vazio, o PDF pode ser uma imagem.")

            for i, pagina in enumerate(leitor_debug.pages):
                texto_cru = pagina.extract_text()
                
                st.markdown(f"### 📄 Página {i+1}")
                
                if texto_cru and len(texto_cru.strip()) > 0:
                    # Mostra o texto dentro de uma caixa de texto para facilitar a leitura
                    st.text_area(f"Texto encontrado na Pág {i+1}", value=texto_cru, height=200)
                else:
                    st.warning(f"⚠️ A página {i+1} parece vazia ou é uma imagem escaneada (sem texto selecionável).")
                
                st.divider()

        except Exception as e:
            st.error(f"❌ Erro ao tentar ler o PDF: {e}")
    else:
        st.warning("Faça o upload do PDF lá em cima primeiro.")


