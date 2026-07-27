import streamlit as st
import cv2
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from pyzbar.pyzbar import decode
from PIL import Image
import os
import zipfile
import io
import json

# --- CONFIGURAÇÃO DE ADMINISTRADOR ---
USUARIO_ADMIN = "diego.costa"

# --- PALETA DE CORES PERSONALIZADA ---
COR_GRAFITE = "#2A2927"
COR_LARANJA = "#F39200"
COR_FUNDO_CARD = "#333230"
COR_TEXTO = "#FFFFFF"

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Organizador de Planilhas",
    page_icon="📊",
    layout="wide"
)

# --- ESTILIZAÇÃO CSS (OCULTA BARRA SUPERIOR, GITHUB, EDIÇÃO E MENU DE 3 PONTOS) ---
st.markdown(f"""
    <style>
    [data-testid="stToolbar"], 
    [data-testid="stHeader"], 
    header, 
    #MainMenu {{
        display: none !important;
        visibility: hidden !important;
        height: 0px !important;
    }}
    .stApp > header {{ display: none !important; }}
    footer {{ display: none !important; }}

    .stApp {{
        background-color: {COR_GRAFITE};
        color: {COR_TEXTO};
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }}
    div.block-container {{
        padding-top: 1rem !important;
        padding-bottom: 2rem;
        max-width: 92%;
    }}
    h1, h2, h3, h4, h5, h6, p, span, label {{
        color: {COR_TEXTO} !important;
    }}
    .stButton>button {{
        background: linear-gradient(90deg, {COR_LARANJA} 0%, #d88100 100%) !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        font-size: 16px !important;
        border: none !important;
        padding: 0.6rem 1.2rem !important;
        box-shadow: 0 4px 15px rgba(243, 146, 0, 0.3) !important;
    }}
    .stDownloadButton>button {{
        background-color: {COR_LARANJA} !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        border: none !important;
    }}
    .badge-admin {{
        background-color: {COR_LARANJA};
        color: #000000;
        font-size: 11px;
        font-weight: bold;
        padding: 3px 8px;
        border-radius: 12px;
        margin-left: 5px;
    }}
    </style>
""", unsafe_allow_html=True)

# --- GERENCIAMENTO DE USUÁRIOS ---
ARQUIVO_USUARIOS = "usuarios.json"

def carregar_usuarios():
    if not os.path.exists(ARQUIVO_USUARIOS):
        dados_padrao = {
            USUARIO_ADMIN: {"senha": "admin123", "status": "aprovado", "role": "admin"},
            "operador": {"senha": "recorte2026", "status": "aprovado", "role": "user"}
        }
        with open(ARQUIVO_USUARIOS, "w") as f:
            json.dump(dados_padrao, f, indent=4)
        return dados_padrao
    try:
        with open(ARQUIVO_USUARIOS, "r") as f:
            return json.load(f)
    except Exception:
        return {}

def salvar_usuarios_dict(usuarios):
    with open(ARQUIVO_USUARIOS, "w") as f:
        json.dump(usuarios, f, indent=4)

def alterar_status_usuario(usuario, novo_status):
    usuarios = carregar_usuarios()
    if usuario in usuarios:
        if novo_status == "excluir":
            del usuarios[usuario]
        else:
            usuarios[usuario]["status"] = novo_status
        salvar_usuarios_dict(usuarios)

# ESTADO DA SESSÃO
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = ""

# --- TELA DE LOGIN ---
if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.8, 1])

    with col2:
        st.markdown(f"""
            <div style="background-color: {COR_FUNDO_CARD}; padding: 25px; border-radius: 12px; text-align: center;">
                <h2 style="color: {COR_LARANJA}; margin-bottom: 5px;">📊 Organizador de Planilhas</h2>
                <p style="color: #aaaaaa; font-size: 14px;">Acesse com sua conta</p>
            </div>
        """, unsafe_allow_html=True)

        usuarios_cadastrados = carregar_usuarios()
        with st.form("form_login"):
            usuario_input = st.text_input("Usuário").strip().lower()
            senha_input = st.text_input("Senha", type="password")
            btn_entrar = st.form_submit_button("Acessar Plataforma", use_container_width=True)

            if btn_entrar:
                if usuario_input in usuarios_cadastrados:
                    dados_usr = usuarios_cadastrados[usuario_input]
                    if dados_usr.get("senha") == senha_input:
                        if dados_usr.get("status") == "aprovado":
                            st.session_state.autenticado = True
                            st.session_state.usuario_logado = usuario_input
                            st.rerun()
                        else:
                            st.warning("⏳ Conta aguardando aprovação.")
                    else:
                        st.error("Senha incorreta.")
                else:
                    st.error("Usuário não encontrado.")
    st.stop()

# VERIFICA SE É ADMIN
usuarios_db = carregar_usuarios()
dados_logado = usuarios_db.get(st.session_state.usuario_logado, {})
e_admin = (st.session_state.usuario_logado == USUARIO_ADMIN) or (dados_logado.get("role") == "admin")

# --- BARRA LATERAL ---
with st.sidebar:
    if e_admin:
        st.markdown(f"👤 **Usuário:** `{st.session_state.usuario_logado}` <span class='badge-admin'>👑 ADMIN</span>", unsafe_allow_html=True)
    else:
        st.markdown(f"👤 **Usuário:** `{st.session_state.usuario_logado}`")
        
    st.markdown("---")
    if st.button("🚪 Sair da Conta", use_container_width=True):
        st.session_state.autenticado = False
        st.session_state.usuario_logado = ""
        st.rerun()

# --- CABEÇALHO ---
col_logo, col_titulo = st.columns([1.2, 4])
with col_logo:
    st.markdown(f"""
        <div style="display:flex; justify-content:center; align-items:center; height:80px; background-color:{COR_FUNDO_CARD}; border-radius:10px;">
            <h1 style="color:{COR_LARANJA} !important; font-weight:900; margin:0;">LOGO</h1>
        </div>
    """, unsafe_allow_html=True)

with col_titulo:
    st.markdown("""
        <h1 style="margin:0; font-size: 32px;">Organizador de Planilhas</h1>
        <p style="margin:0; color:#bbbbbb !important;">Vincule fotos a planilhas Excel automaticamente através da leitura de código de barras.</p>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- NAVEGAÇÃO ---
if e_admin:
    tab_ferramenta, tab_admin = st.tabs(["⚙️ Ferramenta de Organização", "👑 Painel do Administrador"])
else:
    tab_ferramenta, = st.tabs(["⚙️ Ferramenta de Organização"])
    tab_admin = None

# ==========================================
# ABA 1: FERRAMENTA PRINCIPAL
# ==========================================
with tab_ferramenta:
    col_xlsx, col_imgs = st.columns([1, 1])

    with col_xlsx:
        st.markdown("### 1. Envie a Planilha Excel (.xlsx)")
        arquivo_excel = st.file_uploader("Selecione a planilha", type=["xlsx"])

    with col_imgs:
        st.markdown("### 2. Envie as Fotos")
        arquivos_fotos = st.file_uploader("Selecione as fotos ou arquivo .ZIP", type=["jpg", "jpeg", "png", "zip"], accept_multiple_files=True)

    if arquivo_excel and arquivos_fotos:
        st.markdown("---")
        st.markdown("### 3. Configuração do Cruzamento")

        wb_temp = openpyxl.load_workbook(arquivo_excel)
        sheet_names = wb_temp.sheetnames
        nome_aba = st.selectbox("Selecione a aba da planilha:", sheet_names)
        
        # Lê cabeçalho com pandas para escolher a coluna do código
        df_temp = pd.read_excel(arquivo_excel, sheet_name=nome_aba)
        colunas_planilha = list(df_temp.columns)
        
        col_cod, col_destino = st.columns(2)
        with col_cod:
            coluna_codigo = st.selectbox("Coluna que contém os CÓDIGOS na planilha:", colunas_planilha)
        with col_destino:
            nome_coluna_foto = st.text_input("Nome da nova coluna para INSERIR A FOTO:", value="FOTO")

        if st.button("🚀 INICIAR VÍNCULO DE FOTOS E PLANILHA", use_container_width=True):
            st.markdown("---")
            progresso = st.progress(0)
            status = st.empty()

            # 1. Processar e Extrair Fotos
            lista_fotos_processar = [] # Guarda tuplas: (pil_image, bytes_image, nome)

            status.write("📂 Extraindo e preparando pacote de imagens...")
            for f in arquivos_fotos:
                if f.name.endswith(".zip"):
                    with zipfile.ZipFile(f) as z:
                        for filename in z.namelist():
                            if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                                img_bytes = z.read(filename)
                                pil_img = Image.open(io.BytesIO(img_bytes))
                                lista_fotos_processar.append((pil_img, img_bytes, filename))
                else:
                    img_bytes = f.read()
                    pil_img = Image.open(io.BytesIO(img_bytes))
                    lista_fotos_processar.append((pil_img, img_bytes, f.name))

            status.write(f"🔍 Leitura de códigos de barra em {len(lista_fotos_processar)} fotos...")

            # Dicionario mapeando: { codigo_lido: imagem_pil }
            mapa_codigo_imagem = {}
            nao_identificados = 0

            for idx, (pil_img, img_bytes, nome_f) in enumerate(lista_fotos_processar):
                # Converter para OpenCV para leitura de código de barras
                np_arr = np.frombuffer(img_bytes, np.uint8)
                cv_img = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)
                
                codigos_encontrados = decode(cv_img)
                if codigos_encontrados:
                    for barcode in codigos_encontrados:
                        codigo_texto = barcode.data.decode('utf-8').strip()
                        mapa_codigo_imagem[str(codigo_texto)] = pil_img
                else:
                    nao_identificados += 1

                progresso.progress((idx + 1) / len(lista_fotos_processar) * 0.5)

            status.write(f"✅ {len(mapa_codigo_imagem)} códigos de barra identificados com sucesso! ({nao_identificados} sem código)")

            # 2. Carregar Excel via openpyxl e Inserir Imagens
            wb = openpyxl.load_workbook(arquivo_excel)
            ws = wb[nome_aba]

            # Achar o índice numérico da coluna de código
            col_idx_codigo = None
            for col in range(1, ws.max_column + 1):
                if str(ws.cell(row=1, column=col).value).strip() == str(coluna_codigo).strip():
                    col_idx_codigo = col
                    break

            # Criar coluna para fotos no final
            col_idx_foto = ws.max_column + 1
            ws.cell(row=1, column=col_idx_foto).value = nome_coluna_foto

            vincularam = 0

            # Percorrer as linhas do Excel
            tot_rows = ws.max_row
            for row in range(2, tot_rows + 1):
                valor_celula = str(ws.cell(row=row, column=col_idx_codigo).value).strip()
                
                # Se o código da célula estiver no nosso mapa de fotos
                if valor_celula in mapa_codigo_imagem:
                    pil_img = mapa_codigo_imagem[valor_celula]
                    
                    # Redimensionar mantendo proporção para caber na célula (Ex: altura 80px)
                    pil_img.thumbnail((120, 80))
                    
                    # Salvar temporariamente na memória
                    img_byte_arr = io.BytesIO()
                    pil_img.save(img_byte_arr, format='PNG')
                    
                    # Inserir no Excel
                    img_excel = OpenpyxlImage(img_byte_arr)
                    cell_address = ws.cell(row=row, column=col_idx_foto).coordinate
                    ws.add_image(img_excel, cell_address)
                    
                    # Ajustar altura da linha para caber a foto
                    ws.row_dimensions[row].height = 65
                    vincularam += 1

                progresso.progress(0.5 + (row / tot_rows * 0.5))

            # Ajustar largura da coluna da foto
            col_letter = openpyxl.utils.get_column_letter(col_idx_foto)
            ws.column_dimensions[col_letter].width = 18

            # Salvar Excel final na memória
            output_excel = io.BytesIO()
            wb.save(output_excel)
            output_excel.seek(0)

            status.success(f"🎉 Processamento concluído! {vincularam} fotos foram anexadas às linhas correspondentes.")

            # Botão para Download do Excel
            st.download_button(
                label="📥 BAIXAR PLANILHA ATUALIZADA COM AS FOTOS",
                data=output_excel,
                file_name="Planilha_Com_Fotos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ==========================================
# ABA 2: PAINEL DE ADMIN
# ==========================================
if e_admin and tab_admin:
    with tab_admin:
        st.markdown("## 👑 Gerenciamento de Usuários")
        todos_usuarios = carregar_usuarios()
        
        for usr, dados in todos_usuarios.items():
            col_u, col_r, col_act = st.columns([2, 2, 2])
            col_u.write(f"**`{usr}`**")
            col_r.write(dados.get("role", "user"))
            with col_act:
                if usr != USUARIO_ADMIN:
                    if st.button("Remover", key=f"del_{usr}"):
                        alterar_status_usuario(usr, "excluir")
                        st.rerun()

# --- RODAPÉ ---
st.markdown("<br><br>---", unsafe_allow_html=True)
st.markdown("<div style='text-align:center; color:#888;'>Desenvolvimento e Engenharia por <strong>Diego Costa</strong></div>", unsafe_allow_html=True)
