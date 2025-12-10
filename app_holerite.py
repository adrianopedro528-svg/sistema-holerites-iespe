import streamlit as st
import smtplib
from email.message import EmailMessage
from pypdf import PdfReader, PdfWriter
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Envio de Holerites", page_icon="📧")

# --- CARREGAR CONFIGURAÇÕES DO COFRE (SECRETS) ---
try:
    DB_FUNCIONARIOS = dict(st.secrets["funcionarios"])
    
    # Carrega dados do email fixo
    EMAIL_REMETENTE = st.secrets["config_email"]["email_fixo"]
    SENHA_REMETENTE = st.secrets["config_email"]["senha_fixa"]
    EMAIL_BCC = st.secrets["config_email"]["email_copia"]
    
except Exception as e:
    st.error(f"Erro ao carregar segredos: {e}")
    st.stop() # Para o app se não tiver senha configurada

if 'banco_dados' not in st.session_state:
    st.session_state['banco_dados'] = DB_FUNCIONARIOS.copy()

# --- FUNÇÃO DE ENVIO DE EMAIL (COM CÓPIA OCULTA) ---
def enviar_email_fixo(destinatario, assunto, corpo, anexo_bytes, nome_arquivo):
    msg = EmailMessage()
    msg['Subject'] = assunto
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = destinatario
    
    # --- AQUI ESTÁ A MÁGICA DA CÓPIA OCULTA ---
    # O funcionário não vê, mas o financeiro recebe
    msg['Bcc'] = EMAIL_BCC 
    
    msg.set_content(corpo)

    # Anexa o PDF
    msg.add_attachment(anexo_bytes, maintype='application', subtype='pdf', filename=nome_arquivo)

    # Configuração GMAIL (Já que o fixo será Gmail)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_REMETENTE, SENHA_REMETENTE)
        smtp.send_message(msg)

# --- INTERFACE VISUAL ---
st.title("📧 Sistema de Envio de Holerites")
st.caption(f"Enviando através de: {EMAIL_REMETENTE}") # Mostra quem está enviando
st.markdown("---")

# Barra Lateral (Agora só informativa, sem login)
with st.sidebar:
    st.header("ℹ️ Status do Sistema")
    st.success("Login Automático Ativo")
    st.info(f"Cópia oculta configurada para:\n{EMAIL_BCC}")

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

# --- BOTÃO DE AÇÃO ---
if st.button("🚀 Disparar Holerites", type="primary"):
    if not arquivo_pdf:
        st.error("Falta o arquivo PDF!")
    elif len(nomes_selecionados) == 0:
        st.warning("Selecione alguém.")
    else:
        barra = st.progress(0)
        status = st.empty()
        cont = 0
        
        try:
            leitor = PdfReader(arquivo_pdf)
            total = len(leitor.pages)
            
            for i, pagina in enumerate(leitor.pages):
                texto = pagina.extract_text()
                
                for nome in nomes_selecionados:
                    if nome.upper() in texto.upper():
                        email_dest = st.session_state['banco_dados'][nome]
                        status.text(f"Enviando para: {nome}...")
                        
                        escritor = PdfWriter()
                        escritor.add_page(pagina)
                        
                        from io import BytesIO
                        pdf_bytes = BytesIO()
                        escritor.write(pdf_bytes)
                        
                        try:
                            # Chama a nova função sem pedir senha
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
                            st.error(f"Erro {nome}: {e}")
# .Espião
st.markdown("---")
with st.expander("🔍 Modo Espião (Veja como o robô lê o PDF)"):
    if arquivo_pdf:
        leitor_debug = PdfReader(arquivo_pdf)
        for i, pagina in enumerate(leitor_debug.pages):
            st.write(f"--- Página {i+1} ---")
            st.text(pagina.extract_text()) # Mostra o texto cru
    else:
        st.warning("Faça o upload do PDF primeiro.")

                
                barra.progress((i + 1) / total)
            
            st.success(f"Finalizado! {cont} holerites enviados.")
            status.empty()
            
        except Exception as e:
            st.error(f"Erro crítico: {e}")

