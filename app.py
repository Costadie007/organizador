import streamlit as st
import os
import re
import math
import zipfile
import shutil
import cv2
import numpy as np
from PIL import Image
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage

# ==============================================================================
# ✏️ CONFIGURAÇÃO DE TÍTULO, DESCRIÇÃO E LOGO
# ==============================================================================
TITULO_PAGINA = "Organizador de Planilhas"
DESCRICAO_PAGINA = "Envie a pasta compactada com as fotos e a planilha Excel para processamento e vinculo automático."
CAMINHO_LOGO = "logo.png"  # Coloque a imagem logo.png na pasta do projeto (opcional)

# 🎨 PALETA DE CORES
COR_GRAFITE = "#2A2927"
COR_LARANJA = "#F39200"
COR_FUNDO_CARD = "#333230"
COR_TEXTO = "#FFFFFF"
# ==============================================================================

# Configuração da página
st.set_page_config(
    page_title=TITULO_PAGINA,
    page_icon="📊",
    layout="wide"
)

# --- CSS CUSTOMIZADO PARA REPRODUZIR O LAYOUT DO RECORTE ---
css_customizado = f"""
<style>
    /* Fundo Principal */
    .stApp {{
        background-color: {COR_GRAFITE};
        color: {COR_TEXTO};
    }}

    /* Estilização das Abas (Tabs) */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 10px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.1);
    }}

    .stTabs [data-baseweb="tab"] {{
        height: 50px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 4px 4px 0px 0px;
        color: {COR_TEXTO};
        font-weight: 600;
    }}

    .stTabs [aria-selected="true"] {{
        border-bottom: 3px solid {COR_LARANJA} !important;
        color: {COR_TEXTO} !important;
    }}

    /* Cards de Estatísticas no Lado Direito */
    .metric-card {{
        background-color: {COR_FUNDO_CARD};
        border-radius: 8px;
        padding: 20px;
        text-align: center;
        margin-bottom: 15px;
        border: 1px solid rgba(255, 255, 255, 0.05);
        box-shadow: 0px 4px 10px rgba(0,0,0,0.2);
    }}

    .metric-value {{
        font-size: 2.2rem;
        font-weight: bold;
        color: {COR_LARANJA};
        margin: 0;
        line-height: 1;
    }}

    .metric-label {{
        font-size: 0.75rem;
        font-weight: 700;
        color: #A0A0A0;
        letter-spacing: 1px;
        margin-top: 8px;
        text-transform: uppercase;
    }}

    /* Botão Principal em Laranja */
    .stButton > button {{
        background-color: {COR_LARANJA} !important;
        color: #FFFFFF !important;
        border: none !important;
        font-weight: bold !important;
        font-size: 1rem !important;
        border-radius: 6px !important;
        padding: 10px 24px !important;
        transition: all 0.3s ease !important;
    }}

    .stButton > button:hover {{
        background-color: #d88100 !important;
        box-shadow: 0 4px 15px rgba(243, 146, 0, 0.4);
    }}

    /* Botões de Download */
    .stDownloadButton > button {{
        background-color: {COR_FUNDO_CARD} !important;
        color: {COR_LARANJA} !important;
        border: 2px solid {COR_LARANJA} !important;
        font-weight: bold !important;
        border-radius: 6px !important;
    }}

    .stDownloadButton > button:hover {{
        background-color: {COR_LARANJA} !important;
        color: #FFFFFF !important;
    }}

    /* Rodapé Discreto */
    .footer {{
        width: 100%;
        text-align: center;
        padding: 25px 0px 10px 0px;
        margin-top: 60px;
        border-top: 1px solid rgba(255, 255, 255, 0.1);
        color: #888888;
        font-size: 0.85rem;
    }}
</style>
"""
st.markdown(css_customizado, unsafe_allow_html=True)

# --- MOTORES DE LEITURA ---
try:
    import zxingcpp
except ImportError:
    import zxing_cpp as zxingcpp

try:
    from pyzbar import pyzbar
    HAS_PYZBAR = True
except Exception:
    HAS_PYZBAR = False

@st.cache_resource
def carregar_ocr():
    try:
        import easyocr
        return easyocr.Reader(['en'], gpu=False, verbose=False)
    except Exception:
        return None

OCR_READER = carregar_ocr()

# --- FUNÇÕES DE PROCESSAMENTO ---

def limpar_texto_codigo(codigo_bruto):
    if not codigo_bruto:
        return None
    limpo = str(codigo_bruto).replace("(", "").replace(")", "").strip()
    limpo = re.sub(r'[^a-zA-Z0-9]', '', limpo)
    return limpo if len(limpo) >= 3 else None


def tentar_decodificar_engines(img_np):
    if img_np is None:
        return None
    try:
        resultados = zxingcpp.read_barcodes(img_np)
        for res in resultados:
            cod = limpar_texto_codigo(res.text)
            if cod:
                return cod
    except Exception:
        pass

    if HAS_PYZBAR:
        try:
            objetos = pyzbar.decode(img_np)
            for obj in objetos:
                cod = limpar_texto_codigo(obj.data.decode("utf-8", errors="ignore"))
                if cod:
                    return cod
        except Exception:
            pass
    return None


def tentar_ocr_texto(img_np):
    if OCR_READER is None or img_np is None:
        return None
    try:
        cinza = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY) if len(img_np.shape) == 3 else img_np
        _, binaria = cv2.threshold(cinza, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        resultados = OCR_READER.readtext(binaria, allowlist='0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ')
        for bbox, texto, confianca in resultados:
            if confianca > 0.35:
                cod = limpar_texto_codigo(texto)
                if cod and len(cod) >= 4:
                    return cod
    except Exception:
        pass
    return None


def ler_codigo_maxima_precisao(caminho_imagem):
    try:
        stream = open(caminho_imagem, "rb")
        bytes_img = bytearray(stream.read())
        stream.close()
        img_orig = cv2.imdecode(np.frombuffer(bytes_img, np.uint8), cv2.IMREAD_COLOR)
    except Exception:
        return None

    if img_orig is None:
        return None

    cod = tentar_decodificar_engines(img_orig)
    if cod:
        return cod

    cinza = cv2.cvtColor(img_orig, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    cinza_clahe = clahe.apply(cinza)
    _, otsu = cv2.threshold(cinza, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    gaussian = cv2.GaussianBlur(cinza, (0, 0), 3)
    nitida = cv2.addWeighted(cinza, 1.5, gaussian, -0.5, 0)

    for var in [cinza, cinza_clahe, otsu, nitida]:
        cod = tentar_decodificar_engines(var)
        if cod:
            return cod

    for var in [cinza, cinza_clahe]:
        for angulo in [cv2.ROTATE_90_CLOCKWISE, cv2.ROTATE_180, cv2.ROTATE_90_COUNTERCLOCKWISE]:
            rotacionada = cv2.rotate(var, angulo)
            cod = tentar_decodificar_engines(rotacionada)
            if cod:
                return cod

    return tentar_ocr_texto(img_orig)


def criar_cache_da_pasta(caminho_pasta, log_box):
    cache = {}
    if not os.path.exists(caminho_pasta):
        return cache

    for arquivo in os.listdir(caminho_pasta):
        if not arquivo.lower().endswith((".jpg", ".jpeg", ".png", ".bmp", ".webp")):
            continue
        caminho_foto = os.path.join(caminho_pasta, arquivo)
        codigo = ler_codigo_maxima_precisao(caminho_foto)
        if codigo:
            cache[codigo] = caminho_foto
            log_box.text(f"  ✓ [{codigo}] -> {arquivo}")
    return cache


# --- CABEÇALHO (LOGO + TÍTULO ESTILO RECORTE DE ETIQUETAS) ---
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    if os.path.exists(CAMINHO_LOGO):
        st.image(CAMINHO_LOGO, width=160)
    else:
        st.markdown(
            f"""
            <div style="background-color:{COR_FUNDO_CARD}; padding: 25px; border-radius: 8px; text-align: center; font-weight: bold; font-size: 1.5rem; letter-spacing: 2px;">
                LOGO
            </div>
            """,
            unsafe_allow_html=True
        )

with col_titulo:
    st.markdown(f"<h1 style='margin-bottom: 0px;'>{TITULO_PAGINA}</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='color: #B0B0B0; margin-top: 5px;'>{DESCRICAO_PAGINA}</p>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- ABAS DE NAVEGAÇÃO ---
tab_ferramenta, tab_admin = st.tabs(["✂️ Ferramenta de Organização", "👑 Painel do Administrador"])

# --- ESTADOS NA SESSÃO PARA CONTRAR FOTOS E DIVISÕES ---
if "total_fotos_vinculadas" not in st.session_state:
    st.session_state.total_fotos_vinculadas = 0
if "total_partes_geradas" not in st.session_state:
    st.session_state.total_partes_geradas = 0

with tab_ferramenta:
    # Divisão em Colunas: Lado Esquerdo (Uploads e Opções) / Lado Direito (Painel do Lote)
    col_esquerda, col_direita = st.columns([2.5, 1])

    with col_esquerda:
        st.markdown("📁 **Selecione ou arraste os arquivos aqui**")
        file_zip = st.file_uploader("Upload da Pasta de Fotos (.ZIP)", type=["zip"])
        file_excel = st.file_uploader("Upload da Planilha Excel (.XLSX)", type=["xlsx"])

        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("⚙️ Configurações das Colunas e Fatiamento", expanded=False):
            nome_aba = st.text_input("Nome da Aba no Excel", value="Modelo de envio em caso de erro")
            
            c1, c2 = st.columns(2)
            with c1:
                col_sgp = st.text_input("Coluna SGP", value="D").upper()
                col_palete = st.text_input("Coluna Palete", value="B").upper()
            with c2:
                col_codigo = st.text_input("Coluna Código", value="A").upper()
                col_foto = st.text_input("Coluna Destino Foto", value="F").upper()

            modo_divisao = st.radio(
                "Regra de Divisão:",
                options=["Não dividir (Planilha Única)", "Dividir em N partes", "Dividir a cada N linhas por arquivo"]
            )
            
            valor_divisao = 1
            if modo_divisao == "Dividir em N partes":
                valor_divisao = st.number_input("Quantidade de partes desejada:", min_value=2, value=10, step=1)
            elif modo_divisao == "Dividir a cada N linhas por arquivo":
                valor_divisao = st.number_input("Quantidade de linhas por planilha:", min_value=1, value=150, step=10)

        st.markdown("<br>", unsafe_allow_html=True)
        btn_iniciar = st.button("🚀 Iniciar Processamento em Lote", use_container_width=True)

    with col_direita:
        st.markdown("### 📊 Painel do Lote")
        
        # Card 1: Fotos Carregadas/Vinculadas
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-value">{st.session_state.total_fotos_vinculadas}</div>
                <div class="metric-label">Fotos Vinculadas</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # Card 2: Partes Fatiadas
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-value">{st.session_state.total_partes_geradas}</div>
                <div class="metric-label">Partes Geradas</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    # EXECUÇÃO DO PROCESSAMENTO
    if btn_iniciar:
        if not file_zip or not file_excel:
            st.error("Por favor, faça o upload do ZIP de fotos e do Excel antes de iniciar.")
        else:
            temp_dir = "temp_processing"
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)

            zip_path = os.path.join(temp_dir, "fotos.zip")
            excel_path = os.path.join(temp_dir, "base.xlsx")

            with open(zip_path, "wb") as f:
                f.write(file_zip.getbuffer())

            with open(excel_path, "wb") as f:
                f.write(file_excel.getbuffer())

            fotos_dir = os.path.join(temp_dir, "fotos_extraidas")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(fotos_dir)

            progress_bar = st.progress(0)
            status_text = st.empty()
            log_box = st.empty()

            status_text.info("Carregando planilha...")
            
            try:
                wb = load_workbook(excel_path)
                if nome_aba not in wb.sheetnames:
                    st.error(f"A aba '{nome_aba}' não existe no Excel!")
                    st.stop()
                ws = wb[nome_aba]
            except Exception as e:
                st.error(f"Erro ao ler arquivo Excel: {e}")
                st.stop()

            cache_global = {}
            mapa_fotos_por_linha = {}
            fotos_inseridas = 0
            total_linhas = ws.max_row - 1

            if total_linhas <= 0:
                st.error("Planilha vazia!")
                st.stop()

            for idx, linha in enumerate(range(2, ws.max_row + 1), start=1):
                percentual = int((idx / total_linhas) * 70)
                progress_bar.progress(percentual)
                status_text.text(f"Processando linha {idx} de {total_linhas}...")

                sgp = ws[f"{col_sgp}{linha}"].value
                codigo = ws[f"{col_codigo}{linha}"].value
                palete = ws[f"{col_palete}{linha}"].value

                if sgp is None or codigo is None or palete is None:
                    continue

                sgp = str(sgp).strip().split(".")[0]
                codigo = str(codigo).strip().split(".")[0]
                palete = str(palete).strip()

                chave_cache = f"{palete}|{codigo}"

                if chave_cache not in cache_global:
                    caminho_pasta = os.path.join(fotos_dir, palete, codigo)
                    if not os.path.exists(caminho_pasta):
                        for root, dirs, files in os.walk(fotos_dir):
                            if root.endswith(os.path.join(palete, codigo)):
                                caminho_pasta = root
                                break

                    cache_global[chave_cache] = criar_cache_da_pasta(caminho_pasta, log_box)

                cache_pasta = cache_global[chave_cache]

                if sgp in cache_pasta:
                    caminho_foto = cache_pasta[sgp]
                    try:
                        img = OpenpyxlImage(caminho_foto)
                        img.width = 120
                        img.height = 120
                        ws.row_dimensions[linha].height = 95
                        ws.add_image(img, f"{col_foto}{linha}")

                        fotos_inseridas += 1
                        mapa_fotos_por_linha[linha] = caminho_foto
                    except Exception:
                        pass

            ws.column_dimensions[col_foto].width = 20
            caminho_principal_salvo = os.path.join(temp_dir, "Planilha_Preenchida.xlsx")
            wb.save(caminho_principal_salvo)

            # FATIAMENTO
            arquivos_fatiados = []
            if modo_divisao != "Não dividir (Planilha Única)":
                status_text.text("Fatiando planilha...")

                if modo_divisao == "Dividir em N partes":
                    num_partes = int(valor_divisao)
                    linhas_por_parte = math.ceil(total_linhas / num_partes)
                else:
                    linhas_por_parte = int(valor_divisao)
                    num_partes = math.ceil(total_linhas / linhas_por_parte)

                for i in range(num_partes):
                    dado_inicio = (i * linhas_por_parte) + 1
                    dado_fim = min(total_linhas, (i + 1) * linhas_por_parte)

                    if dado_inicio > total_linhas:
                        break

                    linha_orig_inicio = dado_inicio + 1
                    linha_orig_fim = dado_fim + 1

                    wb_parte = load_workbook(caminho_principal_salvo)
                    ws_parte = wb_parte[nome_aba]
                    ws_parte._images.clear()

                    if ws_parte.max_row > linha_orig_fim:
                        ws_parte.delete_rows(linha_orig_fim + 1, ws_parte.max_row - linha_orig_fim)

                    if linha_orig_inicio > 2:
                        ws_parte.delete_rows(2, linha_orig_inicio - 2)

                    linha_destino = 2
                    for linha_orig in range(linha_orig_inicio, linha_orig_fim + 1):
                        caminho_foto = mapa_fotos_por_linha.get(linha_orig)
                        if caminho_foto and os.path.exists(caminho_foto):
                            try:
                                img = OpenpyxlImage(caminho_foto)
                                img.width = 120
                                img.height = 120
                                ws_parte.row_dimensions[linha_destino].height = 95
                                ws_parte.add_image(img, f"{col_foto}{linha_destino}")
                            except Exception:
                                pass
                        linha_destino += 1

                    ws_parte.column_dimensions[col_foto].width = 20
                    nome_saida = f"Parte_{i + 1}_(Linhas_{dado_inicio}-{dado_fim}).xlsx"
                    caminho_saida = os.path.join(temp_dir, nome_saida)
                    wb_parte.save(caminho_saida)
                    arquivos_fatiados.append(caminho_saida)

            # ATUALIZA CONTADORES DO PAINEL
            st.session_state.total_fotos_vinculadas = fotos_inseridas
            st.session_state.total_partes_geradas = len(arquivos_fatiados) if arquivos_fatiados else 1

            progress_bar.progress(100)
            status_text.success(f"🎉 Concluído com sucesso! {fotos_inseridas} fotos inseridas.")

            # BOTÕES DE DOWNLOAD
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                with open(caminho_principal_salvo, "rb") as f:
                    st.download_button(
                        label="📄 Baixar Planilha Completa",
                        data=f.read(),
                        file_name="Planilha_Preenchida_Completa.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            if arquivos_fatiados:
                with col_d2:
                    zip_fatiados_path = os.path.join(temp_dir, "Planilhas_Divididas.zip")
                    with zipfile.ZipFile(zip_fatiados_path, 'w') as zip_f:
                        for arq in arquivos_fatiados:
                            zip_f.write(arq, arcname=os.path.basename(arq))

                    with open(zip_fatiados_path, "rb") as f:
                        st.download_button(
                            label="📦 Baixar Partes Divididas (.ZIP)",
                            data=f.read(),
                            file_name="Planilhas_Divididas.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
            st.rerun()

with tab_admin:
    st.subheader("👑 Painel de Administração")
    st.info("Aqui você pode acompanhar relatórios de uso ou configurações do servidor.")

# --- RODAPÉ DISCRETO ---
st.markdown(
    """
    <div class="footer">
        Desenvolvido por <strong>Diego Costa</strong>
    </div>
    """,
    unsafe_allow_html=True
)
