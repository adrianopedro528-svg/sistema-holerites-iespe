import streamlit as st
import smtplib
from email.message import EmailMessage
from pypdf import PdfReader, PdfWriter
import pdfplumber
import io

# --- CONFIGURAÇÃO DA PÁGINA ---
try:
    st.set_page_config(page_title="Envio de Holerites", page_icon="logo.png")
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
    # Remove quebras de linha e espaços duplos
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
col_logo, col_texto = st.columns([1, 6])
try:
    col_logo.image("logo.png", width=80)
except:
    col_logo.write("📧") 
col_texto.title("Envio Fácil de Holerites")

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
            # 1. Prepara o PDF para CORTAR (usando pypdf)
            arquivo_pdf.seek(0)
            leitor_corte = PdfReader(arquivo_pdf)
            total_paginas = len(leitor_corte.pages)
            
            # 2. Prepara o PDF para LER TEXTO (usando pdfplumber - MUITO MAIS FORTE)
            arquivo_pdf.seek(0)
            with pdfplumber.open(arquivo_pdf) as pdf_leitura:
                
                # Loop página por página
                for i, pagina_plumber in enumerate(pdf_leitura.pages):
                    
                    # Extração Poderosa do Plumber
                    texto_original = pagina_plumber.extract_text() or ""
                    texto_limpo = limpar_texto(texto_original)
                    
                    encontrou_dono = False
                    
                    for nome in nomes_selecionados:
                        nome_limpo = limpar_texto(nome)
                        
                        # Verifica se o nome está na página
                        if nome_limpo in texto_limpo:
                            encontrou_dono = True
                            funcionarios_encontrados.add(nome)
                            status.text(f"Pág {i+1}: Encontrado {nome}...")
                            
                            # Usa o pypdf para cortar a página correspondente
                            escritor = PdfWriter()
                            # Pega a mesma página 'i' no leitor de corte
                            escritor.add_page(leitor_corte.pages[i])
                            
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
                            
                            # Achou o dono desta página, para de testar nomes e vai pra próxima pág
                            break 
                    
                    if not encontrou_dono:
                        preview = texto_limpo[:100] + "..." if texto_limpo else "Texto ilegível/Imagem"
                        paginas_nao_identificadas.append((i+1, preview))

                    barra.progress((i + 1) / total_paginas)
            
            status.empty()
            
            # --- RELATÓRIO FINAL ---
            st.divider()
            st.subheader("📊 Relatório do Disparo")
            
            total_enviados = len(funcionarios_encontrados)
            if total_enviados > 0:
                st.success(f"**{total_enviados}** holerites identificados e processados com sucesso.")
            else:
                st.warning("Nenhum holerite foi enviado.")

            nao_encontrados = set(nomes_selecionados) - funcionarios_encontrados
            if nao_encontrados:
                st.error(f"❌ **Funcionários não encontrados no arquivo ({len(nao_encontrados)}):**")
                st.write(", ".join(nao_encontrados))
            
            if paginas_nao_identificadas:
                with st.expander(f"⚠️ {len(paginas_nao_identificadas)} Páginas sem dono (Clique para ver o motivo)"):
                    st.write("Isso acontece se o nome no PDF estiver escrito diferente do cadastro.")
                    for pag, texto in paginas_nao_identificadas:
                        st.markdown(f"**Página {pag}:** O robô leu: `{texto}`")
            
            if erros_envio:
                with st.expander("Erros de Conexão/Envio"):
                    for erro in erros_envio:
                        st.write(erro)

        except Exception as e:
            st.error(f"Erro crítico no processamento: {e}")

# --- MODO ESPIÃO ATUALIZADO (COM PDFPLUMBER) ---
st.markdown("---")
with st.expander("🔍 Modo Espião (Verifique como o robô lê AGORA)"):
    if arquivo_pdf:
        st.info("Usando a tecnologia nova (pdfplumber) para ler o arquivo:")
        arquivo_pdf.seek(0)
        with pdfplumber.open(arquivo_pdf) as pdf_debug:
            for i, pagina in enumerate(pdf_debug.pages):
                texto = limpar_texto(pagina.extract_text())
                st.markdown(f"**Página {i+1}**")
                st.code(texto) # Mostra dentro de uma caixa de código para ficar claro
                st.divider()
