import streamlit as st
import cv2
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import zxingcpp
from PIL import Image, ImageOps
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

# --- ESTILIZAÇÃO CSS (OCULTA BARRA SUPERIOR, GITHUB, EDIÇÃO E MENU) ---
st.markdown(f"""
    <style>
    [data-testid="stToolbar"], [data-testid="stHeader"], header, #MainMenu {{
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
    h1, h2, h3, h4, h5, h6, p, span, label {{ color: {COR_TEXTO} !important; }}
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

# --- MOTOR ULTRA AVANÇADO DE LEITURA MULTI-ESTÁGIO ---
def extrair_codigos_imagem_extremo(cv_img):
    codigos_encontrados = set()

    def ler_zxing(img):
        results = zxingcpp.read_barcodes(img)
        for r in results:
            if r.valid and r.text:
                codigos_encontrados.add(r.text.strip())

    # Passo 1: Leitura na imagem original e em escala de cinza
    ler_zxing(cv_img)
    if codigos_encontrados:
        return list(codigos_encontrados)

    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY) if len(cv_img.shape) == 3 else cv_img
    ler_zxing(gray)
    if codigos_encontrados:
        return list(codigos_encontrados)

    # Passo 2: Sharpening (Nitidez para fotos desfocadas)
    kernel_sharpen = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
    sharp = cv2.filter2D(gray, -1, kernel_sharpen)
    ler_zxing(sharp)
    if codigos_encontrados:
        return list(codigos_encontrados)

    # Passo 3: CLAHE (Ajuste local de contraste para fotos com sombra)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
    gray_contrast = clahe.apply(gray)
    ler_zxing(gray_contrast)
    if codigos_encontrados:
        return list(codigos_encontrados)

    # Passo 4: Limiarização Adaptativa e Binarização de Otsu
    _, thresh_otsu = cv2.threshold(gray_contrast, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    ler_zxing(thresh_otsu)
    if codigos_encontrados:
        return list(codigos_encontrados)

    thresh_adapt = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    ler_zxing(thresh_adapt)
    if codigos_encontrados:
        return list(codigos_encontrados)

    # Passo 5: Detecção de Região de Interesse (Auto-Crop na área do Código)
    try:
        grad_x = cv2.Sobel(gray, ddepth=cv2.CV_32F, dx=1, dy=0, ksize=-1)
        grad_y = cv2.Sobel(gray, ddepth=cv2.CV_32F, dx=0, dy=1, ksize=-1)
        gradient = cv2.subtract(grad_x, grad_y)
        gradient = cv2.convertScaleAbs(gradient)
        blurred = cv2.blur(gradient, (9, 9))
        _, thresh_crop = cv2.threshold(blurred, 225, 255, cv2.THRESH_BINARY)
        
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (21, 7))
        closed = cv2.morphologyEx(thresh_crop, cv2.MORPH_CLOSE, kernel)
        cnts, _ = cv2.findContours(closed.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        if cnts:
            c = sorted(cnts, key=cv2.contourArea, reverse=True)[0]
            rect = cv2.minAreaRect(c)
            box = cv2.boxPoints(rect)
            box = np.intp(box)
            x, y, w, h = cv2.boundingRect(box)
            # Crop com folga
            pad = 20
            h_img, w_img = gray.shape
            crop = gray[max(0, y-pad):min(h_img, y+h+pad), max(0, x-pad):min(w_img, x+w+pad)]
            if crop.size > 0:
                ler_zxing(crop)
                if codigos_encontrados:
                    return list(codigos_encontrados)
    except Exception:
        pass

    # Passo 6: Testar Várias Escalas (Resize Zoom In / Zoom Out)
    for fx_fy in [1.5, 2.0, 0.5]:
        resized = cv2.resize(gray_contrast, (0, 0), fx=fx_fy, fy=fx_fy, interpolation=cv2.INTER_CUBIC)
        ler_zxing(resized)
        if codigos_encontrados:
            return list(codigos_encontrados)

    # Passo 7: Rotações de 90°, 180° e 270° em todas as tentativas anteriores
    for rot in [cv2.ROTATE_90_CLOCKWISE, cv2.ROTATE_180, cv2.ROTATE_90_COUNTERCLOCKWISE]:
        rot_img = cv2.rotate(gray_contrast, rot)
        ler_zxing(rot_img)
        if codigos_encontrados:
            return list(codigos_encontrados)

    return list(codigos_encontrados)

def normalizar_codigo_9_digitos(val_raw):
    if val_raw is None:
        return ""
    val_bruto = str(val_raw).strip().split('.')[0]
    val_numerico = ''.join(filter(str.isdigit, val_bruto))
    if not val_numerico:
        return ""
    return val_numerico[-9:] if len(val_numerico) >= 9 else val_numerico.zfill(9)


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
        
        df_temp = pd.read_excel(arquivo_excel, sheet_name=nome_aba)
        colunas_planilha = list(df_temp.columns)
        
        col_cod, col_destino = st.columns(2)
        with col_cod:
            coluna_codigo = st.selectbox("Coluna que contém os CÓDIGOS na planilha:", colunas_planilha)
        with col_destino:
            nome_coluna_foto = st.text_input("Nome da nova coluna para INSERIR A FOTO:", value="FOTO")

        if st.button("🚀 INICIAR PROCESSAMENTO AUTOMÁTICO", use_container_width=True):
            st.session_state.mapa_codigo_imagem = {}
            st.session_state.fotos_pendentes = []

            lista_fotos_processar = []
            for f in arquivos_fotos:
                if f.name.endswith(".zip"):
                    with zipfile.ZipFile(f) as z:
                        for filename in z.namelist():
                            if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                                img_bytes = z.read(filename)
                                pil_img = Image.open(io.BytesIO(img_bytes))
                                pil_img = ImageOps.exif_transpose(pil_img)
                                lista_fotos_processar.append((pil_img, img_bytes, filename))
                else:
                    img_bytes = f.read()
                    pil_img = Image.open(io.BytesIO(img_bytes))
                    pil_img = ImageOps.exif_transpose(pil_img)
                    lista_fotos_processar.append((pil_img, img_bytes, f.name))

            progresso = st.progress(0)
            status = st.empty()
            status.write(f"⚡ Processando {len(lista_fotos_processar)} fotos com motor de inteligência visual...")

            for idx, (pil_img, img_bytes, nome_f) in enumerate(lista_fotos_processar):
                np_arr = np.frombuffer(img_bytes, np.uint8)
                cv_img = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)
                
                codigos_lidos = extrair_codigos_imagem_extremo(cv_img) if cv_img is not None else []
                
                if codigos_lidos:
                    for cod in codigos_lidos:
                        cod_9 = normalizar_codigo_9_digitos(cod)
                        if cod_9:
                            st.session_state.mapa_codigo_imagem[cod_9] = pil_img
                else:
                    st.session_state.fotos_pendentes.append((nome_f, pil_img))

                progresso.progress((idx + 1) / len(lista_fotos_processar))

            status.success(f"✅ Leitura concluída! **{len(st.session_state.mapa_codigo_imagem)}** código(s) mapeado(s) automaticamente.")

        # --- SEÇÃO DE EXCEÇÕES (Apenas para emergências) ---
        if "fotos_pendentes" in st.session_state and len(st.session_state.fotos_pendentes) > 0:
            st.markdown("---")
            st.warning(f"⚠️ Apenas **{len(st.session_state.fotos_pendentes)}** foto(s) de um total alto não puderam ser lidas automaticamente.")

            codigos_digitados = {}
            cols = st.columns(min(len(st.session_state.fotos_pendentes), 3))

            for idx, (nome_f, img_obj) in enumerate(st.session_state.fotos_pendentes):
                with cols[idx % 3]:
                    st.image(img_obj, caption=f"📄 {nome_f}", use_container_width=True)
                    ent_code = st.text_input(f"Código ({nome_f}):", key=f"manual_{idx}", placeholder="Últimos 9 dígitos")
                    if ent_code.strip():
                        codigos_digitados[idx] = normalizar_codigo_9_digitos(ent_code)

            if st.button("➕ Vincular exceções manuais", use_container_width=True):
                removidos = []
                for idx, cod_norm in codigos_digitados.items():
                    if cod_norm:
                        nome_f, img_obj = st.session_state.fotos_pendentes[idx]
                        st.session_state.mapa_codigo_imagem[cod_norm] = img_obj
                        removidos.append(idx)
                
                st.session_state.fotos_pendentes = [item for i, item in enumerate(st.session_state.fotos_pendentes) if i not in removidos]
                st.success("Exceções atualizadas!")
                st.rerun()

        # --- GERAÇÃO E DOWNLOAD DA PLANILHA ---
        if "mapa_codigo_imagem" in st.session_state and len(st.session_state.mapa_codigo_imagem) > 0:
            st.markdown("---")
            if st.button("📊 GERAR PLANILHA FINAL COM AS FOTOS", use_container_width=True):
                wb = openpyxl.load_workbook(arquivo_excel)
                ws = wb[nome_aba]

                col_idx_codigo = None
                for col in range(1, ws.max_column + 1):
                    val_cabecalho = str(ws.cell(row=1, column=col).value or "").strip()
                    if val_cabecalho == str(coluna_codigo).strip():
                        col_idx_codigo = col
                        break

                if not col_idx_codigo:
                    st.error(f"❌ Coluna '{coluna_codigo}' não encontrada na planilha.")
                    st.stop()

                col_idx_foto = ws.max_column + 1
                ws.cell(row=1, column=col_idx_foto).value = nome_coluna_foto

                vincularam = 0
                tot_rows = ws.max_row

                for row in range(2, tot_rows + 1):
                    raw_val = ws.cell(row=row, column=col_idx_codigo).value
                    codigo_excel_9 = normalizar_codigo_9_digitos(raw_val)

                    if codigo_excel_9 and codigo_excel_9 in st.session_state.mapa_codigo_imagem:
                        pil_img = st.session_state.mapa_codigo_imagem[codigo_excel_9].copy()
                        pil_img.thumbnail((130, 80))
                        
                        img_byte_arr = io.BytesIO()
                        pil_img.save(img_byte_arr, format='PNG')
                        
                        img_excel = OpenpyxlImage(img_byte_arr)
                        cell_address = ws.cell(row=row, column=col_idx_foto).coordinate
                        ws.add_image(img_excel, cell_address)
                        
                        ws.row_dimensions[row].height = 65
                        vincularam += 1

                col_letter = openpyxl.utils.get_column_letter(col_idx_foto)
                ws.column_dimensions[col_letter].width = 18

                output_excel = io.BytesIO()
                wb.save(output_excel)
                output_excel.seek(0)

                st.success(f"🎉 **{vincularam}** fotos inseridas na planilha com sucesso!")
                st.download_button(
                    label="📥 BAIXAR PLANILHA FINAL COM FOTOS",
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
