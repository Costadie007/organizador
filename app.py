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
import re

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
        <p style="margin:0; color:#bbbbbb !important;">Vincule fotos a planilhas Excel automaticamente através da leitura rápida de dígitos e código de barras.</p>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- NAVEGAÇÃO ---
if e_admin:
    tab_ferramenta, tab_admin = st.tabs(["⚙️ Ferramenta de Organização", "👑 Painel do Administrador"])
else:
    tab_ferramenta, = st.tabs(["⚙️ Ferramenta de Organização"])
    tab_admin = None

# HELPER: EXTRAI APENAS OS DÍGITOS NUMÉRICOS
def extrair_apenas_digitos(texto):
    if texto is None:
        return ""
    return str(re.sub(r'\D', '', str(texto)))

# LEITOR RÁPIDO DE CÓDIGO DA IMAGEM
def extrair_digitos_imagem(cv_img, nome_arquivo):
    digitos_encontrados = set()
    
    # 1. Extrai números do próprio NOME do arquivo
    nums_nome = extrair_apenas_digitos(nome_arquivo)
    if nums_nome:
        digitos_encontrados.add(nums_nome)

    # 2. Tenta leitura direta do código de barras
    try:
        results = zxingcpp.read_barcodes(cv_img)
        for r in results:
            if r.valid and r.text:
                num_bar = extrair_apenas_digitos(r.text)
                if num_bar:
                    digitos_encontrados.add(num_bar)
    except Exception:
        pass

    return list(digitos_encontrados)

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

        if st.button("🚀 INICIAR VERIFICAÇÃO HIERÁRQUICA (4D ➔ 8D)", use_container_width=True):
            st.session_state.mapa_codigo_imagem = {}
            st.session_state.fotos_conflito = []

            # 1. Carregar fotos e extrair identificadores
            lista_fotos = []
            for f in arquivos_fotos:
                if f.name.endswith(".zip"):
                    with zipfile.ZipFile(f) as z:
                        for filename in z.namelist():
                            if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                                img_bytes = z.read(filename)
                                pil_img = Image.open(io.BytesIO(img_bytes))
                                pil_img = ImageOps.exif_transpose(pil_img)
                                lista_fotos.append((filename, pil_img, img_bytes))
                else:
                    img_bytes = f.read()
                    pil_img = Image.open(io.BytesIO(img_bytes))
                    pil_img = ImageOps.exif_transpose(pil_img)
                    lista_fotos.append((f.name, pil_img, img_bytes))

            # Mapeamento de todas as fotos para seus dígitos
            banco_fotos_digitos = []
            for nome_f, pil_img, img_bytes in lista_fotos:
                np_arr = np.frombuffer(img_bytes, np.uint8)
                cv_img = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)
                digitos = extrair_digitos_imagem(cv_img, nome_f)
                banco_fotos_digitos.append({
                    "nome": nome_f,
                    "pil_img": pil_img,
                    "digitos_list": digitos
                })

            # 2. Varre a planilha e aplica as regras de 4 dígitos -> 8 dígitos -> Conflito
            st.session_state.conflitos_pendentes = []
            
            total_linhas = len(df_temp)
            progresso = st.progress(0)
            status = st.empty()

            vincularam_auto = 0

            for idx, row_data in df_temp.iterrows():
                val_raw = row_data[coluna_codigo]
                cod_excel_num = extrair_apenas_digitos(val_raw)

                if not cod_excel_num:
                    continue

                d4_excel = cod_excel_num[-4:] if len(cod_excel_num) >= 4 else cod_excel_num.zfill(4)
                d8_excel = cod_excel_num[-8:] if len(cod_excel_num) >= 8 else cod_excel_num.zfill(8)

                # PASSAGEM 1: Buscar fotos com os mesmos últimos 4 dígitos
                matches_4d = []
                for foto in banco_fotos_digitos:
                    for d in foto["digitos_list"]:
                        if d.endswith(d4_excel):
                            matches_4d.append(foto)
                            break

                # CASO A: Apenas 1 foto bateu nos últimos 4 dígitos -> VÍNCULO DIRETO!
                if len(matches_4d) == 1:
                    st.session_state.mapa_codigo_imagem[cod_excel_num] = matches_4d[0]["pil_img"]
                    vincularam_auto += 1

                # CASO B: Mais de 1 foto com os mesmos 4 dígitos -> Desempate pelos últimos 8 dígitos
                elif len(matches_4d) > 1:
                    matches_8d = []
                    for foto in matches_4d:
                        for d in foto["digitos_list"]:
                            if d.endswith(d8_excel):
                                matches_8d.append(foto)
                                break
                    
                    # Desempate com sucesso nos 8 dígitos!
                    if len(matches_8d) == 1:
                        st.session_state.mapa_codigo_imagem[cod_excel_num] = matches_8d[0]["pil_img"]
                        vincularam_auto += 1
                    else:
                        # Conflito real não resolvido -> Manda para validação manual de desempate
                        st.session_state.conflitos_pendentes.append({
                            "codigo_excel": cod_excel_num,
                            "opcoes_fotos": matches_8d if len(matches_8d) > 0 else matches_4d
                        })
                else:
                    # Nenhuma foto bateu 4 dígitos
                    pass

                progresso.progress((idx + 1) / total_linhas)

            status.success(f"✅ Processamento concluído! **{vincularam_auto}** fotos vinculadas automaticamente. Conflitos para validar: **{len(st.session_state.conflitos_pendentes)}**.")

        # --- SEÇÃO DE VALIDAÇÃO MANUAL APENAS EM CONFLITOS REAIS ---
        if "conflitos_pendentes" in st.session_state and len(st.session_state.conflitos_pendentes) > 0:
            st.markdown("---")
            st.warning(f"⚠️ **{len(st.session_state.conflitos_pendentes)} item(ns)** possuem fotos com os mesmos últimos dígitos. Selecione qual foto pertence a qual código:")

            respostas_conflitos = {}

            for idx, conflito in enumerate(st.session_state.conflitos_pendentes):
                st.markdown(f"#### Código da Planilha: `{conflito['codigo_excel']}`")
                
                opcoes_fotos = conflito["opcoes_fotos"]
                cols = st.columns(min(len(opcoes_fotos), 4))
                
                nomes_opcoes = ["-- Selecionar --"] + [f["nome"] for f in opcoes_fotos]
                
                for f_idx, foto_obj in enumerate(opcoes_fotos):
                    with cols[f_idx % 4]:
                        st.image(foto_obj["pil_img"], caption=foto_obj["nome"], use_container_width=True)

                escolha = st.selectbox(
                    f"Qual foto corresponde ao código {conflito['codigo_excel']}?",
                    options=nomes_opcoes,
                    key=f"conf_{idx}"
                )
                
                if escolha != "-- Selecionar --":
                    # Encontra a imagem da opção
                    for f in opcoes_fotos:
                        if f["nome"] == escolha:
                            respostas_conflitos[conflito['codigo_excel']] = f["pil_img"]

            if st.button("➕ Confirmar Escolhas e Atualizar Planilha", use_container_width=True):
                for cod, img_pil in respostas_conflitos.items():
                    st.session_state.mapa_codigo_imagem[cod] = img_pil
                
                # Remove resolvidos
                st.session_state.conflitos_pendentes = [c for c in st.session_state.conflitos_pendentes if c['codigo_excel'] not in respostas_conflitos]
                st.success("Conflitos resolvidos!")
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
                    cod_num = extrair_apenas_digitos(raw_val)

                    if cod_num and cod_num in st.session_state.mapa_codigo_imagem:
                        pil_img = st.session_state.mapa_codigo_imagem[cod_num].copy()
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
st.markdown("<div style='text-align:center; color:#888;'>Desenvolvido por <strong>Diego Costa</strong></div>", unsafe_allow_html=True)
