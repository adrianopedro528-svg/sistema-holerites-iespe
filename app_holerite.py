import streamlit as st
import smtplib
from email.message import EmailMessage
from pypdf import PdfReader, PdfWriter
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Envio Fácil", page_icon="📧")

# --- LISTA PADRÃO DE FUNCIONÁRIOS (Seu banco de dados fixo) ---
# O sistema sempre vai reiniciar com essa lista completa.
# Em vez de escrever a lista aqui, pegamos dos segredos do Streamlit
# Se estiver rodando local no seu PC, ele vai dar erro se não tiver o arquivo .streamlit/secrets.toml
# Mas na nuvem funciona direto.
if "funcionarios" in st.secrets:
    DB_FUNCIONARIOS = st.secrets["funcionarios"]
else:
    # Fallback apenas para teste local seguro, ou deixe vazio
    DB_FUNCIONARIOS = {}

# --- FUNÇÃO DE ENVIO DE EMAIL ---
def enviar_email(remetente, senha, destinatario, assunto, corpo, anexo_bytes, nome_arquivo):
    msg = EmailMessage()
    msg['Subject'] = assunto
    msg['From'] = remetente
    msg['To'] = destinatario
    msg.set_content(corpo)

    # Anexa o PDF (que está na memória)
    msg.add_attachment(anexo_bytes, maintype='application', subtype='pdf', filename=nome_arquivo)

    # Conexão segura
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(remetente, senha)
        smtp.send_message(msg)

# --- INTERFACE VISUAL (FRONT-END) ---
# ... (Mantenha os imports e a função enviar_email iguais) ...

# --- INICIALIZAÇÃO DA MEMÓRIA (SESSION STATE) ---
# Isso garante que a lista comece com os nomes fixos, mas aceite novos sem esquecer
if 'banco_dados' not in st.session_state:
    st.session_state['banco_dados'] = DB_FUNCIONARIOS.copy()

# --- INTERFACE VISUAL ---
# Colocando a logo se tiver (opcional)
# col_logo, col_titulo = st.columns([1, 4])
# with col_logo: st.image("logo.png", width=100)
# with col_titulo: st.title("Sistema de Holerites")
st.title("📧 Sistema de Envio de Holerites")

st.markdown("---")

# Barra Lateral (Login)
with st.sidebar:
    st.header("🔐 Login do Email")
    email_user = st.text_input("Seu E-mail", value="seu_email@iespe.com.br")
    email_pass = st.text_input("Senha / Senha de App", type="password")
    st.info("Para Outlook/Office365, use sua senha normal ou App Password se tiver autenticação em 2 etapas.")

# Colunas principais
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload do Arquivo")
    arquivo_pdf = st.file_uploader("Solte o PDF aqui", type="pdf")

with col2:
    st.subheader("2. Mensagem")
    assunto_email = st.text_input("Assunto", value="Holerite - Novembro")
    corpo_email = st.text_area("Texto do email", value="Segue em anexo seu holerite.\n\nAtt,\nDP", height=100)

st.markdown("---")

# --- ÁREA DE ADICIONAR NOVO FUNCIONÁRIO (A Mágica acontece aqui) ---
with st.expander("➕ Adicionar alguém que não está na lista (Clique aqui)"):
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        # O nome tem que ser igual ao que está escrito no PDF para o robô achar
        novo_nome = st.text_input("Nome (Como está no PDF)")
    with c2:
        novo_email = st.text_input("E-mail para envio")
    with c3:
        st.write("") # Espaço para alinhar o botão
        st.write("") 
        if st.button("Adicionar"):
            if novo_nome and novo_email:
                # Adiciona na memória do sistema
                st.session_state['banco_dados'][novo_nome] = novo_email
                st.success(f"{novo_nome} adicionado!")
                # O st.rerun() recarrega a página para o nome aparecer na lista abaixo na hora
                st.rerun()
            else:
                st.warning("Preencha nome e email.")

st.subheader("3. Seleção de Destinatários")

# Agora o multiselect lê do 'banco_dados' da memória, que pode ter gente nova
lista_atualizada = st.session_state['banco_dados']

nomes_selecionados = st.multiselect(
    "Quem vai receber o email?",
    options=list(lista_atualizada.keys()), 
    default=list(lista_atualizada.keys())
)

st.write(f"Você enviará emails para **{len(nomes_selecionados)}** pessoas.")

# --- BOTÃO DE AÇÃO (Mantenha o resto igual) ---
if st.button("🚀 Iniciar Disparo", type="primary"):
    # ... O resto do código continua igual, MAS ATENÇÃO:
    # Dentro do loop, onde você usava DB_FUNCIONARIOS[nome], 
    # troque por: st.session_state['banco_dados'][nome]


    # Validações antes de começar
    if not arquivo_pdf:
        st.error("Por favor, faça o upload do PDF primeiro!")
    elif not email_pass:
        st.error("Por favor, preencha a Senha de App na barra lateral!")
    elif len(nomes_selecionados) == 0:
        st.warning("Selecione pelo menos um funcionário.")
    else:
        # INÍCIO DO PROCESSO
        barra_progresso = st.progress(0)
        status_text = st.empty()
        
        try:
            leitor = PdfReader(arquivo_pdf)
            total_paginas = len(leitor.pages)
            enviados_count = 0
            
            status_text.text("Lendo arquivo PDF...")
            
            # Loop pelas páginas
            for i, pagina in enumerate(leitor.pages):
                texto = pagina.extract_text()
                
                # Verifica se a página pertence a alguém SELECIONADO na lista
                for nome in nomes_selecionados:
                    if nome.upper() in texto.upper():
                        email_destino = st.session_state['banco_dados'][nome]
                        
                        status_text.text(f"Processando holerite de: {nome}...")
                        
                        # Recorta o PDF na memória (sem salvar no disco)
                        escritor = PdfWriter()
                        escritor.add_page(pagina)
                        
                        # Salva num "arquivo virtual" na memória RAM
                        from io import BytesIO
                        pdf_bytes = BytesIO()
                        escritor.write(pdf_bytes)
                        pdf_conteudo = pdf_bytes.getvalue()
                        
                        # Envia
                        try:
                            enviar_email(
                                email_user, 
                                email_pass, 
                                email_destino, 
                                assunto_email, 
                                corpo_email, 
                                pdf_conteudo, 
                                f"Holerite_{nome}.pdf"
                            )
                            enviados_count += 1
                            st.toast(f"✅ Enviado para {nome}")
                        except Exception as e:
                            st.error(f"Erro ao enviar para {nome}: {e}")
                
                # Atualiza barra de progresso
                barra_progresso.progress((i + 1) / total_paginas)

            st.success(f"Processo finalizado! Total de emails enviados: {enviados_count}")
            barra_progresso.empty()
            
        except Exception as e:

            st.error(f"Erro crítico no processamento: {e}")
