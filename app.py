import io
import json
import math
import os
import re
import zipfile
from difflib import SequenceMatcher

import cv2
import openpyxl
import pandas as pd
from openpyxl.drawing.image import Image as OpenpyxlImage
import numpy as np
from PIL import Image, ImageOps
import streamlit as st

# --- 1. MOTORES DE LEITURA E OCR ---
try:
    import zxingcpp
except ImportError:
    import zxing_cpp as zxingcpp

try:
    from pyzbar import pyzbar
    HAS_PYZBAR = True
except Exception:
    HAS_PYZBAR = False

try:
    import easyocr
    @st.cache_resource
    def carregar_ocr():
        # Inicializa o EasyOCR focado em números/letras maiúsculas
        return easyocr.Reader(['en'], gpu=False, verbose=False)
    OCR_READER = carregar_ocr()
    HAS_OCR = True
except Exception:
    HAS_OCR = False

# --- CONFIGURAÇÃO E PALETA DE CORES ---
USUARIO_ADMIN = "diego.costa"
COR_GRAFITE = "#2A2927"
COR_LARANJA = "#F39200"
COR_FUNDO_CARD = "#333230"
COR_TEXTO = "#FFFFFF"

st.set_page_config(
    page_title="Organizador de Planilhas Pro Ultra",
    page_icon="📊",
    layout="wide"
)

# Proteção contra erros DOM/React no Streamlit Cloud
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)

st.markdown(f"""
    <style>
    [data-testid="stToolbar"], [data-testid="stHeader"], header, #MainMenu {{
        display: none !important;
        visibility: hidden !important;
        height: 0px !important;
    }}
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

# --- USUÁRIOS E AUTENTICAÇÃO ---
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

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = ""

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

usuarios_db = carregar_usuarios()
dados_logado = usuarios_db.get(st.session_state.usuario_logado, {})
e_admin = (st.session_state.usuario_logado == USUARIO_ADMIN) or (dados_logado.get("role") == "admin")

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
        <h1 style="margin:0; font-size: 32px;">Organizador de Planilhas Pro Ultra</h1>
        <p style="margin:0; color:#bbbbbb !important;">Leitura de máxima precisão: Pré-processamento avançado, OCR com filtro e Matching por Janela Deslizante.</p>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- TRATAMENTO DE TEXTO E CÓDIGOS ---
def extrair_apenas_digitos(texto):
    if texto is None:
        return ""
    return str(re.sub(r'\D', '', str(texto)))

def limpar_texto_codigo(codigo_bruto):
    if not codigo_bruto:
        return None
    limpo = str(codigo_bruto).replace("(", "").replace(")", "").strip()
    limpo = re.sub(r'[^a-zA-Z0-9]', '', limpo)
    return limpo if len(limpo) >= 3 else None

# --- PIPELINE DE PROCESSAMENTO DE IMAGEM MULTI-FILTRO ---
def gerar_variacoes_imagem(cv_img):
    variacoes = []
    if cv_img is None:
        return variacoes

    # Redimensiona se for gigantesca para acelerar e focar detalhes
    h, w = cv_img.shape[:2]
    if max(h, w) > 1800:
        escala = 1800 / max(h, w)
        cv_img = cv2.resize(cv_img, (int(w * escala), int(h * escala)), interpolation=cv2.INTER_AREA)

    # 1. Original
    variacoes.append(cv_img)

    # 2. Tons de Cinza
    cinza = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY) if len(cv_img.shape) == 3 else cv_img
    variacoes.append(cinza)

    # 3. Aumento de Contraste Adaptativo (CLAHE)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    cinza_clahe = clahe.apply(cinza)
    variacoes.append(cinza_clahe)

    # 4. Nitidez (Sharpening Kernel)
    kernel_sharp = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
    nitida = cv2.filter2D(cinza, -1, kernel_sharp)
    variacoes.append(nitida)

    # 5. Binarização Otsu
    _, otsu = cv2.threshold(cinza, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    variacoes.append(otsu)

    # 6. Binarização Invertida
    _, otsu_inv = cv2.threshold(cinza, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    variacoes.append(otsu_inv)

    return variacoes

def tentar_decodificar_leitores(img_np):
    codigos = set()
    if img_np is None:
        return codigos

    # Engine 1: ZXing
    try:
        resultados = zxingcpp.read_barcodes(img_np)
        for res in resultados:
            cod = limpar_texto_codigo(res.text)
            if cod:
                codigos.add(cod)
    except Exception:
        pass

    # Engine 2: PyZbar
    if HAS_PYZBAR:
        try:
            objs = pyzbar.decode(img_np)
            for obj in objs:
                cod = limpar_texto_codigo(obj.data.decode("utf-8", errors="ignore"))
                if cod:
                    codigos.add(cod)
        except Exception:
            pass

    return codigos

def tentar_ocr_extremo(img_np):
    codigos = set()
    if not HAS_OCR or img_np is None:
        return codigos
    try:
        res = OCR_READER.readtext(img_np, allowlist='0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-')
        for bbox, texto, confianca in res:
            if confianca > 0.15:
                # Trata erros típicos de leitura visual
                texto_corr = (texto.replace('O', '0')
                              .replace('I', '1')
                              .replace('L', '1')
                              .replace('Z', '2')
                              .replace('S', '5')
                              .replace('B', '8')
                              .replace('G', '6'))
                cod = limpar_texto_codigo(texto_corr)
                if cod and len(cod) >= 3:
                    codigos.add(cod)
    except Exception:
        pass
    return codigos

def ler_imagem_todas_camadas(cv_img, nome_arquivo):
    codigos_encontrados = set()

    # 0. Digitos diretos do NOME DO ARQUIVO
    digs_nome = extrair_apenas_digitos(nome_arquivo)
    if digs_nome and len(digs_nome) >= 4:
        codigos_encontrados.add(digs_nome)

    if cv_img is None:
        return list(codigos_encontrados)

    # Testa as 4 rotações da imagem (0, 90, 180, 270 graus)
    orientacoes = [
        cv_img,
        cv2.rotate(cv_img, cv2.ROTATE_90_CLOCKWISE),
        cv2.rotate(cv_img, cv2.ROTATE_180),
        cv2.rotate(cv_img, cv2.ROTATE_90_COUNTERCLOCKWISE)
    ]

    for img_rot in orientacoes:
        variacoes = gerar_variacoes_imagem(img_rot)

        # Passada 1: Leitura Ultrarrápida de Código de Barras nas variações
        for var in variacoes:
            cods = tentar_decodificar_leitores(var)
            if cods:
                codigos_encontrados.update(cods)

        # Se já encontrou código de barras de alta precisão, encerra para economizar tempo
        if codigos_encontrados:
            break

        # Passada 2: Executa OCR caso nenhum código de barras tenha sido lido
        if HAS_OCR:
            for var in variacoes[:3]: # Testa original, cinza e CLAHE
                cods_ocr = tentar_ocr_extremo(var)
                if cods_ocr:
                    codigos_encontrados.update(cods_ocr)
            if codigos_encontrados:
                break

    return list(codigos_encontrados)

# --- ALGORITMO DE MATCHING ULTRA AVANÇADO ---
def calcular_similaridade_avancada(digitos_excel, texto_foto_completo, digitos_foto_lista):
    if not digitos_excel:
        return 0.0

    digs_foto_concat = "".join(digitos_foto_lista) + extrair_apenas_digitos(texto_foto_completo)

    # 1. Contenção Exata de Substring
    if digitos_excel in digs_foto_concat:
        return 1.0

    # 2. Janela Deslizante de Tamanho Semelhante
    tam_ex = len(digitos_excel)
    if tam_ex >= 4 and len(digs_foto_concat) >= tam_ex:
        melhor_ratio_janela = 0.0
        for i in range(len(digs_foto_concat) - tam_ex + 1):
            sub = digs_foto_concat[i:i+tam_ex]
            ratio = SequenceMatcher(None, digitos_excel, sub).ratio()
            if ratio > melhor_ratio_janela:
                melhor_ratio_janela = ratio
        if melhor_ratio_janela >= 0.75:
            return melhor_ratio_janela

    # 3. Sufixos/Prefixos Fortes (Ex: últimos 4 ou 6 dígitos)
    if tam_ex >= 4:
        sufixo_4 = digitos_excel[-4:]
        sufixo_6 = digitos_excel[-6:] if tam_ex >= 6 else sufixo_4
        if sufixo_6 in digs_foto_concat:
            return 0.90
        if sufixo_4 in digs_foto_concat:
            return 0.75

    # 4. Ratio Global do SequenceMatcher
    return SequenceMatcher(None, digitos_excel, digs_foto_concat).ratio()


# --- INTERFACE E NAVEGAÇÃO ---
if e_admin:
    tab_ferramenta, tab_admin = st.tabs(["⚙️ Ferramenta de Organização", "👑 Painel do Administrador"])
else:
    tab_ferramenta = st.tabs(["⚙️ Ferramenta de Organização"])[0]
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
        st.markdown("### 2. Envie as Fotos (ou arquivo .ZIP)")
        arquivos_fotos = st.file_uploader("Selecione as fotos", type=["jpg", "jpeg", "png", "zip"], accept_multiple_files=True)

    if arquivo_excel and arquivos_fotos:
        st.markdown("---")
        st.markdown("### 3. Mapeamento e Parâmetros")

        wb_temp = openpyxl.load_workbook(arquivo_excel)
        sheet_names = wb_temp.sheetnames
        nome_aba = st.selectbox("Selecione a aba da planilha:", sheet_names)

        df_temp = pd.read_excel(arquivo_excel, sheet_name=nome_aba)
        colunas_planilha = list(df_temp.columns)

        col_c1, col_c2 = st.columns(2)
        with col_c1:
            coluna_sgp = st.selectbox("Coluna com o código SGP/Principal:", colunas_planilha)
        with col_c2:
            nome_coluna_foto = st.text_input("Nome da coluna para INSERIR A FOTO:", value="FOTO")

        if st.button("🚀 INICIAR PROCESSAMENTO ULTRA AUTOMÁTICO", use_container_width=True):
            st.session_state.mapa_codigo_imagem = {}

            # Carrega todas as imagens enviadas
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

            banco_fotos = []
            container_status = st.container()

            # Processamento das imagens
            with container_status:
                status_leitura = st.empty()
                prog_bar = st.progress(0)
                status_leitura.write(f"🔍 Processando {len(lista_fotos)} imagens com pré-processamento multi-filtro e OCR...")

                for idx, (nome_f, pil_img, img_bytes) in enumerate(lista_fotos):
                    np_arr = np.frombuffer(img_bytes, np.uint8)
                    cv_img = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)
                    
                    codigos_extraidos = ler_imagem_todas_camadas(cv_img, nome_f)

                    banco_fotos.append({
                        "nome": nome_f,
                        "pil_img": pil_img,
                        "codigos": codigos_extraidos,
                        "digitos_list": [extrair_apenas_digitos(c) for c in codigos_extraidos if extrair_apenas_digitos(c)],
                        "usada": False
                    })
                    prog_bar.progress((idx + 1) / len(lista_fotos))

                prog_bar.empty()

            # --- PROCESSAMENTO DOS CÓDIGOS DA PLANILHA ---
            vincularam_auto = 0
            
            # Mapeamento do Excel
            codigos_excel_dict = {}
            for row_idx, row_data in df_temp.iterrows():
                val_raw = row_data[coluna_sgp]
                digs_excel = extrair_apenas_digitos(val_raw)
                if digs_excel:
                    codigos_excel_dict[digs_excel] = val_raw

            # PASSADA 1: MATCHING DE ALTA CONFIANÇA (Substrings e Sufixos Diretos)
            for cod_excel, val_raw in codigos_excel_dict.items():
                for foto in banco_fotos:
                    if foto["usada"]:
                        continue

                    # Verifica se algum código extraído confere perfeitamente
                    score = calcular_similaridade_avancada(cod_excel, foto["nome"], foto["digitos_list"])
                    
                    if score >= 0.85:
                        st.session_state.mapa_codigo_imagem[cod_excel] = foto["pil_img"]
                        foto["usada"] = True
                        vincularam_auto += 1
                        break

            # PASSADA 2: MATCHING RECURSIVO FUZZY (Para imagens com ruído ou OCR parcial)
            vincularam_fuzzy = 0
            for cod_excel, val_raw in codigos_excel_dict.items():
                if cod_excel in st.session_state.mapa_codigo_imagem:
                    continue

                melhor_score = 0.0
                melhor_foto = None

                for foto in banco_fotos:
                    if foto["usada"]:
                        continue

                    score = calcular_similaridade_avancada(cod_excel, foto["nome"], foto["digitos_list"])
                    if score > melhor_score and score >= 0.45:  # Limiar seguro para segunda passada
                        melhor_score = score
                        melhor_foto = foto

                if melhor_foto is not None:
                    st.session_state.mapa_codigo_imagem[cod_excel] = melhor_foto["pil_img"]
                    melhor_foto["usada"] = True
                    vincularam_fuzzy += 1

            total_vinculados = len(st.session_state.mapa_codigo_imagem)
            st.success(
                f"🎉 **Processamento de Alta Precisão Concluído!**\n\n"
                f"• Vínculos Diretos e Precisos: **{vincularam_auto}**\n"
                f"• Vínculos por Aproximação e Reconhecimento Numérico: **{vincularam_fuzzy}**\n"
                f"• Total de Fotos Vinculadas à Planilha: **{total_vinculados} de {len(lista_fotos)}**"
            )

        # --- GERAÇÃO DA PLANILHA DE SAÍDA ---
        if "mapa_codigo_imagem" in st.session_state and len(st.session_state.mapa_codigo_imagem) > 0:
            st.markdown("---")
            st.markdown("### 4. Configuração da Saída e Divisão do Excel")

            modo_divisao = st.radio(
                "Como deseja salvar/dividir a planilha?",
                ["Planilha Única (Sem divisão)", "Dividir por QUANTIDADE DE PARTES", "Dividir por QUANTIDADE DE LINHAS"],
                horizontal=True
            )

            num_partes = 1
            linhas_por_parte = 0

            if modo_divisao == "Dividir por QUANTIDADE DE PARTES":
                num_partes = st.number_input("Digite a quantidade de partes desejada:", min_value=2, value=2, step=1)
            elif modo_divisao == "Dividir por QUANTIDADE DE LINHAS":
                linhas_por_parte = st.number_input("Digite o número de linhas por arquivo:", min_value=10, value=150, step=10)

            if st.button("📊 APLICAR FOTOS E GERAR ARQUIVO(S)", use_container_width=True):
                arquivo_excel.seek(0)
                wb = openpyxl.load_workbook(arquivo_excel)
                ws = wb[nome_aba]

                col_idx_codigo = None
                for col in range(1, ws.max_column + 1):
                    val_cabecalho = str(ws.cell(row=1, column=col).value or "").strip()
                    if val_cabecalho == str(coluna_sgp).strip():
                        col_idx_codigo = col
                        break

                col_idx_foto = ws.max_column + 1
                ws.cell(row=1, column=col_idx_foto).value = nome_coluna_foto

                mapa_fotos_linha = {}
                vincularam = 0

                for row in range(2, ws.max_row + 1):
                    raw_val = ws.cell(row=row, column=col_idx_codigo).value
                    cod_num = extrair_apenas_digitos(raw_val)

                    if cod_num and cod_num in st.session_state.mapa_codigo_imagem:
                        pil_img = st.session_state.mapa_codigo_imagem[cod_num].copy()
                        pil_img.thumbnail((120, 120))

                        img_byte_arr = io.BytesIO()
                        pil_img.save(img_byte_arr, format='PNG')

                        img_excel = OpenpyxlImage(img_byte_arr)
                        cell_address = ws.cell(row=row, column=col_idx_foto).coordinate
                        ws.add_image(img_excel, cell_address)

                        ws.row_dimensions[row].height = 95
                        vincularam += 1
                        mapa_fotos_linha[row] = img_byte_arr

                col_letter = openpyxl.utils.get_column_letter(col_idx_foto)
                ws.column_dimensions[col_letter].width = 20
                total_dados = ws.max_row - 1

                if modo_divisao == "Planilha Única (Sem divisão)" or total_dados <= 0:
                    out_buffer = io.BytesIO()
                    wb.save(out_buffer)
                    out_buffer.seek(0)

                    st.success(f"🎉 Concluído! **{vincularam}** fotos inseridas com sucesso.")
                    st.download_button(
                        label="📥 BAIXAR PLANILHA COMPLETA (.XLSX)",
                        data=out_buffer,
                        file_name="Planilha_Com_Fotos.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    if modo_divisao == "Dividir por QUANTIDADE DE PARTES":
                        linhas_por_parte = math.ceil(total_dados / num_partes)

                    qtd_partes_calculadas = math.ceil(total_dados / linhas_por_parte)

                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for p in range(qtd_partes_calculadas):
                            dado_inicio = (p * linhas_por_parte) + 1
                            dado_fim = min(total_dados, (p + 1) * linhas_por_parte)

                            linha_orig_inicio = dado_inicio + 1
                            linha_orig_fim = dado_fim + 1

                            arquivo_excel.seek(0)
                            wb_p = openpyxl.load_workbook(arquivo_excel)
                            ws_p = wb_p[nome_aba]
                            ws_p._images.clear()

                            if ws_p.max_row > linha_orig_fim:
                                ws_p.delete_rows(linha_orig_fim + 1, ws_p.max_row - linha_orig_fim)
                            if linha_orig_inicio > 2:
                                ws_p.delete_rows(2, linha_orig_inicio - 2)

                            col_f_p = ws_p.max_column + 1
                            ws_p.cell(row=1, column=col_f_p).value = nome_coluna_foto

                            l_dest = 2
                            for l_orig in range(linha_orig_inicio, linha_orig_fim + 1):
                                if l_orig in mapa_fotos_linha:
                                    img_b = mapa_fotos_linha[l_orig]
                                    img_b.seek(0)
                                    img_ex = OpenpyxlImage(img_b)
                                    ws_p.add_image(img_ex, f"{openpyxl.utils.get_column_letter(col_f_p)}{l_dest}")
                                    ws_p.row_dimensions[l_dest].height = 95
                                l_dest += 1

                            ws_p.column_dimensions[openpyxl.utils.get_column_letter(col_f_p)].width = 20

                            out_p = io.BytesIO()
                            wb_p.save(out_p)
                            out_p.seek(0)

                            nome_sub_arq = f"Planilha_Parte_{p+1}_(Linhas_{dado_inicio}-{dado_fim}).xlsx"
                            zf.writestr(nome_sub_arq, out_p.getvalue())

                    zip_buffer.seek(0)
                    st.success(f"🎉 Divisão concluída! **{qtd_partes_calculadas}** planilhas geradas em um arquivo .ZIP.")
                    st.download_button(
                        label="📥 BAIXAR PACOTE COM AS PLANILHAS DIVIDIDAS (.ZIP)",
                        data=zip_buffer,
                        file_name="Planilhas_Divididas.zip",
                        mime="application/zip",
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

st.markdown("<br><br>---", unsafe_allow_html=True)
st.markdown("<div style='text-align:center; color:#888;'>Desenvolvimento e Engenharia por <strong>Diego Costa</strong></div>", unsafe_allow_html=True)
