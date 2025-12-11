import streamlit as st
import smtplib
from email.message import EmailMessage
from pypdf import PdfReader, PdfWriter
import io

# --- CONFIGURAÇÃO DA PÁGINA ---
try:
    st.set_page_config(page_title="Envio de Holerites", page_icon="ensine icone 2025.png")
except:
    st.set_page_config(page_title="Envio de Holerites", page_icon="📧")

# --- CARREGAR CONFIGURAÇÕES DO COFRE (SECRETS) ---
try:
    DB_FUNCIONARIOS = dict(st.secrets["funcionarios"])
    EMAIL_REMETENTE = st.secrets["config_email"]["email_fixo"]
    SENHA_REMETENTE = st.secrets["config_email"]["senha_fixa"]
    EMAIL_BCC = st.secrets["config_email"]["email_copia"]
except Exception as e:
    st.error(f"Erro ao carregar Secrets. Verifique a configuração no Streamlit Cloud.")
    DB_FUNCIONARIOS = {} 
    EMAIL_REMETENTE = ""
    SENHA_REMETENTE = ""
    EMAIL_BCC = ""

# --- INICIALIZA SESSÃO ---
if 'banco_dados' not in st.session_state:
    st.session_state['banco_dados'] = DB_FUNCIONARIOS.copy()

# --- FUNÇÃO DE LIMPEZA DE TEXTO ---
def limpar_texto(texto):
    if not texto: return ""
    return " ".join(texto.split()).upper()

# --- FUNÇÃO DE ENVIO DE EMAIL ---
def enviar_email_fixo(destinatario, assunto, corpo, anexo_bytes, nome_arquivo):
    msg = EmailMessage()
    msg['Subject'] = assunto
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = destinatario
    msg['Bcc'] = EMAIL_BCC 
    msg.set_content(corpo)
    msg.add_attachment(anexo_bytes, maintype='application', subtype='pdf', filename=nome_arquivo)

    if "gmail.com" in EMAIL_REMETENTE.lower():
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_REMETENTE, SENHA_REMETENTE)
            smtp.send_message(msg)
    else:
        with smtplib.SMTP('smtp.office365.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(EMAIL_REMETENTE, SENHA_REMETENTE)
            smtp.send_message(msg)

# --- INTERFACE VISUAL ---
ol_logo, col_texto = st.columns([1, 6]) # Ajuste o 1 e 6 para mudar a proporção
col_logo.image("ensine icone 2025.png", width=80)     # Ajuste o width para o tamanho da sua logo
col_texto.title("Envio Fácil")
if EMAIL_REMETENTE:
    st.caption(f"Enviando através de: {EMAIL_REMETENTE}")
else:
    st.error("⚠️ Configure os Secrets no Streamlit Cloud!")

st.markdown("---")

with st.sidebar:
    st.header("ℹ️ Status")
    if EMAIL_REMETENTE:
        st.success("✅ Sistema Ativo")
    else:
        st.error("❌ Falta Configuração")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Upload")
    arquivo_pdf = st.file_uploader("Solte o PDF aqui", type="pdf")
with col2:
    st.subheader("2. Mensagem")
    assunto_email = st.text_input("Assunto", value="Holerite - Pagamento")
    corpo_email = st.text_area("Texto", value="Segue em anexo seu holerite.\n\nAtt,\nFinanceiro - IESPE", height=100)

st.markdown("---")

# --- ADICIONAR FUNCIONÁRIO ---
with st.expander("➕ Adicionar Novo Funcionário"):
    c1, c2, c3 = st.columns([2, 2, 1])
    novo_nome = c1.text_input("Nome (Trecho único)")
    novo_email = c2.text_input("E-mail")
    if c3.button("Salvar") and novo_nome and novo_email:
        st.session_state['banco_dados'][novo_nome] = novo_email
        st.success("Adicionado!")
        st.rerun()

st.subheader("3. Seleção")
lista_atual = st.session_state['banco_dados']
nomes_selecionados = st.multiselect("Destinatários", options=list(lista_atual.keys()), default=list(lista_atual.keys()))

st.write(f"Selecionados: **{len(nomes_selecionados)}**")

# --- LÓGICA DE DISPARO ---
if st.button("🚀 Disparar Holerites", type="primary"):
    if not arquivo_pdf:
        st.error("Falta o arquivo PDF!")
    elif not nomes_selecionados:
        st.warning("Selecione alguém na lista.")
    elif not EMAIL_REMETENTE:
        st.error("Erro de configuração de e-mail.")
    else:
        paginas_nao_identificadas = []
        funcionarios_encontrados = set()
        erros_envio = []
        
        barra = st.progress(0)
        status = st.empty()
        
        try:
            arquivo_pdf.seek(0)
            leitor = PdfReader(arquivo_pdf)
            total_paginas = len(leitor.pages)
            
            for i, pagina in enumerate(leitor.pages):
                texto_original = pagina.extract_text()
                texto_limpo = limpar_texto(texto_original)
                encontrou_dono = False
                
                for nome in nomes_selecionados:
                    nome_limpo = limpar_texto(nome)
                    
                    if nome_limpo in texto_limpo:
                        encontrou_dono = True
                        funcionarios_encontrados.add(nome)
                        status.text(f"Pág {i+1}: Encontrado {nome}...")
                        
                        escritor = PdfWriter()
                        escritor.add_page(pagina)
                        pdf_bytes = io.BytesIO()
                        escritor.write(pdf_bytes)
                        
                        try:
                            enviar_email_fixo(
                                lista_atual[nome], 
                                assunto_email, 
                                corpo_email, 
                                pdf_bytes.getvalue(), 
                                f"Holerite_{nome}.pdf"
                            )
                            st.toast(f"✅ Enviado: {nome}")
                        except Exception as e:
                            erros_envio.append(f"{nome}: {e}")
                        break 
                
                if not encontrou_dono:
                    preview = texto_limpo[:100] + "..." if texto_limpo else "Página vazia/Imagem"
                    paginas_nao_identificadas.append((i+1, preview))

                barra.progress((i + 1) / total_paginas)
            
            status.empty()
            # st.balloons() REMOVIDO AQUI
            
            # --- RELATÓRIO FINAL ---
            st.divider()
            st.subheader("📊 Relatório do Disparo")
            
            total_enviados = len(funcionarios_encontrados)
            st.success(f"**{total_enviados}** holerites identificados e processados.")

            nao_encontrados = set(nomes_selecionados) - funcionarios_encontrados
            if nao_encontrados:
                st.error(f"❌ **Funcionários não encontrados no arquivo ({len(nao_encontrados)}):**")
                st.write(", ".join(nao_encontrados))
                st.info("Dica: Verifique se o nome no cadastro está idêntico ao PDF (use o Espião abaixo).")
            
            if paginas_nao_identificadas:
                st.warning(f"⚠️ **{len(paginas_nao_identificadas)} Páginas não foram enviadas (sem dono identificado):**")
                for pag, texto in paginas_nao_identificadas:
                    st.text(f"Página {pag}: O robô leu -> {texto}")
            
            if erros_envio:
                with st.expander("Erros de Conexão/Envio"):
                    for erro in erros_envio:
                        st.write(erro)

        except Exception as e:
            st.error(f"Erro crítico: {e}")

# --- MODO ESPIÃO ---
st.markdown("---")
with st.expander("🔍 Modo Espião (Verifique como cadastrar os nomes)"):
    if arquivo_pdf:
        arquivo_pdf.seek(0)
        leitor_debug = PdfReader(arquivo_pdf)
        st.info("Copie o nome EXATAMENTE como aparece abaixo (em maiúsculo e sem acentos se estiver assim).")
        for i, pagina in enumerate(leitor_debug.pages):
            texto = limpar_texto(pagina.extract_text())
            st.text(f"Pág {i+1}: {texto}")
            st.divider()

