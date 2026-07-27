import io
import json
import math
import os
import re
import zipfile
from difflib import SequenceMatcher

import cv2
import numpy as np
import openpyxl
import pandas as pd
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image, ImageOps
import streamlit as st

# --- 1. MOTORES DE CÓDIGO DE BARRAS & OCR ---
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
        return easyocr.Reader(['en'], gpu=False, verbose=False)

    OCR_READER = carregar_ocr()
    HAS_OCR = True
except Exception:
    HAS_OCR = False

# --- CONFIGURAÇÃO DE ADMINISTRADOR ---
USUARIO_ADMIN = "diego.costa"

# --- PALETA DE CORES PERSONALIZADA ---
COR_GRAFITE = "#2A2927"
COR_LARANJA = "#F39200"
COR_FUNDO_CARD = "#333230"
COR_TEXTO = "#FFFFFF"

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Organizador de Planilhas Pro Ultra",
    page_icon="📊",
    layout="wide",
)

# Meta tag para evitar que o Google Tradutor quebre o DOM do React
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)

# --- ESTILIZAÇÃO CSS ---
st.markdown(
    f"""
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
""",
    unsafe_allow_html=True,
)

# --- GERENCIAMENTO DE USUÁRIOS ---
ARQUIVO_USUARIOS = "usuarios.json"


def carregar_usuarios():
    if not os.path.exists(ARQUIVO_USUARIOS):
        dados_padrao = {
            USUARIO_ADMIN: {
                "senha": "admin123",
                "status": "aprovado",
                "role": "admin",
            },
            "operador": {
                "senha": "recorte2026",
                "status": "aprovado",
                "role": "user",
            },
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
        st.markdown(
            f"""
            <div style="background-color: {COR_FUNDO_CARD}; padding: 25px; border-radius: 12px; text-align: center;">
                <h2 style="color: {COR_LARANJA}; margin-bottom: 5px;">📊 Organizador de Planilhas</h2>
                <p style="color: #aaaaaa; font-size: 14px;">Acesse com sua conta</p>
            </div>
        """,
            unsafe_allow_html=True,
        )

        usuarios_cadastrados = carregar_usuarios()
        with st.form("form_login"):
            usuario_input = st.text_input("Usuário").strip().lower()
            senha_input = st.text_input("Senha", type="password")
            btn_entrar = st.form_submit_button(
                "Acessar Plataforma", use_container_width=True
            )

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
e_admin = (st.session_state.usuario_logado == USUARIO_ADMIN) or (
    dados_logado.get("role") == "admin"
)

# --- BARRA LATERAL ---
with st.sidebar:
    if e_admin:
        st.markdown(
            f"👤 **Usuário:** `{st.session_state.usuario_logado}` <span"
            " class='badge-admin'>👑 ADMIN</span>",
            unsafe_allow_html=True,
        )
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
    st.markdown(
        f"""
        <div style="display:flex; justify-content:center; align-items:center; height:80px; background-color:{COR_FUNDO_CARD}; border-radius:10px;">
            <h1 style="color:{COR_LARANJA} !important; font-weight:900; margin:0;">LOGO</h1>
        </div>
    """,
        unsafe_allow_html=True,
    )

with col_titulo:
    st.markdown(
        """
        <h1 style="margin:0; font-size: 32px;">Organizador de Planilhas Pro Ultra</h1>
        <p style="margin:0; color:#bbbbbb !important;">Processamento massivo automatizado: Recorte regional, ROI adaptativo, correção OCR avançada e Fuzzy Matching.</p>
    """,
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# --- FUNÇÕES AVANÇADAS DE PROCESSAMENTO DE IMAGEM ---


def limpar_texto_codigo(codigo_bruto):
    if not codigo_bruto:
        return None
    limpo = str(codigo_bruto).replace("(", "").replace(")", "").strip()
    limpo = re.sub(r'[^a-zA-Z0-9]', '', limpo)
    return limpo if len(limpo) >= 3 else None


def extrair_apenas_digitos(texto):
    if texto is None:
        return ""
    return str(re.sub(r'\D', '', str(texto)))


def aplicar_filtro_nitidez(img_np):
    if img_np is None:
        return None
    kernel_nitidez = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
    return cv2.filter2D(img_np, -1, kernel_nitidez)


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
                cod = limpar_texto_codigo(
                    obj.data.decode("utf-8", errors="ignore")
                )
                if cod:
                    return cod
        except Exception:
            pass

    return None


def tentar_ocr_texto(img_np):
    if not HAS_OCR or img_np is None:
        return None
    try:
        cinza = (
            cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)
            if len(img_np.shape) == 3
            else img_np
        )
        _, binaria = cv2.threshold(
            cinza, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )
        resultados = OCR_READER.readtext(
            binaria, allowlist='0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        )

        for bbox, texto, confianca in resultados:
            if confianca > 0.25:
                texto_corrigido = (
                    texto.replace('O', '0')
                    .replace('I', '1')
                    .replace('Z', '2')
                    .replace('S', '5')
                    .replace('B', '8')
                )
                cod = limpar_texto_codigo(texto_corrigido)
                if cod and len(cod) >= 3:
                    return cod
    except Exception:
        pass
    return None


def extrair_rois_recortes(cv_img):
    rois = [cv_img]
    h, w = cv_img.shape[:2]

    if h > 300 and w > 300:
        rois.append(cv_img[int(h * 0.15) : int(h * 0.85), int(w * 0.15) : int(w * 0.85)])
        rois.append(cv_img[int(h * 0.4) : h, 0:w])
        rois.append(cv_img[0 : int(h * 0.6), 0:w])

        cinza = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(cinza, (5, 5), 0)
        _, thresh = cv2.threshold(
            blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )
        contours, _ = cv2.findContours(
            thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )

        contours = sorted(contours, key=cv2.contourArea, reverse=True)[:3]
        for c in contours:
            x, y, bw, bh = cv2.boundingRect(c)
            if bw > 80 and bh > 40 and (bw * bh) < (w * h * 0.9):
                pad = 10
                y1, y2 = max(0, y - pad), min(h, y + bh + pad)
                x1, x2 = max(0, x - pad), min(w, x + bw + pad)
                rois.append(cv_img[y1:y2, x1:x2])

    return rois


def varrer_orientacao_imagem(cv_img):
    cods_encontrados = set()
    if cv_img is None:
        return cods_encontrados

    rois = extrair_rois_recortes(cv_img)

    for roi in rois:
        cod = tentar_decodificar_engines(roi)
        if cod:
            cods_encontrados.add(cod)

        img_nitida = aplicar_filtro_nitidez(roi)
        cinza = (
            cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
            if len(roi.shape) == 3
            else roi
        )
        cinza_nitida = (
            cv2.cvtColor(img_nitida, cv2.COLOR_BGR2GRAY)
            if len(img_nitida.shape) == 3
            else img_nitida
        )

        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        cinza_clahe = clahe.apply(cinza)
        _, otsu = cv2.threshold(
            cinza, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )

        for var in [img_nitida, cinza, cinza_nitida, cinza_clahe, otsu]:
            c = tentar_decodificar_engines(var)
            if c:
                cods_encontrados.add(c)

        if HAS_OCR:
            for var in [roi, cinza_nitida, cinza_clahe]:
                c_ocr = tentar_ocr_texto(var)
                if c_ocr:
                    cods_encontrados.add(c_ocr)

        if cods_encontrados:
            break

    return cods_encontrados


def ler_codigo_multi_camadas(cv_img, nome_arquivo=""):
    cods = set()
    digs_nome = extrair_apenas_digitos(nome_arquivo)
    if digs_nome and len(digs_nome) >= 4:
        cods.add(digs_nome)

    if cv_img is None:
        return list(cods)

    cods.update(varrer_orientacao_imagem(cv_img))

    if not cods:
        rotacoes = [
            cv2.ROTATE_90_CLOCKWISE,
            cv2.ROTATE_180,
            cv2.ROTATE_90_COUNTERCLOCKWISE,
        ]
        for angulo in rotacoes:
            img_rotacionada = cv2.rotate(cv_img, angulo)
            cods_rot = varrer_orientacao_imagem(img_rotacionada)
            if cods_rot:
                cods.update(cods_rot)
                break

    return list(cods)


# --- NAVEGAÇÃO ---
if e_admin:
    tab_ferramenta, tab_admin = st.tabs(
        ["⚙️ Ferramenta de Organização", "👑 Painel do Administrador"]
    )
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
        arquivos_fotos = st.file_uploader(
            "Selecione as fotos",
            type=["jpg", "jpeg", "png", "zip"],
            accept_multiple_files=True,
        )

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
            coluna_sgp = st.selectbox(
                "Coluna com o código SGP/Principal:", colunas_planilha
            )
        with col_c2:
            nome_coluna_foto = st.text_input(
                "Nome da coluna para INSERIR A FOTO:", value="FOTO"
            )

        codigos_excel_disponiveis = [
            extrair_apenas_digitos(val)
            for val in df_temp[coluna_sgp].dropna().unique()
            if extrair_apenas_digitos(val) != ""
        ]

        if st.button(
            "🚀 INICIAR PROCESSAMENTO ULTRA AUTOMÁTICO", use_container_width=True
        ):
            st.session_state.mapa_codigo_imagem = {}

            lista_fotos = []
            for f in arquivos_fotos:
                if f.name.endswith(".zip"):
                    with zipfile.ZipFile(f) as z:
                        for filename in z.namelist():
                            if filename.lower().endswith(
                                ('.png', '.jpg', '.jpeg')
                            ):
                                img_bytes = z.read(filename)
                                pil_img = Image.open(io.BytesIO(img_bytes))
                                pil_img = ImageOps.exif_transpose(pil_img)
                                lista_fotos.append(
                                    (filename, pil_img, img_bytes)
                                )
                else:
                    img_bytes = f.read()
                    pil_img = Image.open(io.BytesIO(img_bytes))
                    pil_img = ImageOps.exif_transpose(pil_img)
                    lista_fotos.append((f.name, pil_img, img_bytes))

            banco_fotos_digitos = []

            # Bloco isolado em container estático para evitar erro de remoção de nó (DOM/React)
            container_status = st.container()
            with container_status:
                status_leitura = st.empty()
                prog_bar = st.progress(0)
                status_leitura.write(
                    f"🔍 Executando varredura regional inteligente em {len(lista_fotos)} imagens..."
                )

                for idx, (nome_f, pil_img, img_bytes) in enumerate(lista_fotos):
                    np_arr = np.frombuffer(img_bytes, np.uint8)
                    cv_img = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)
                    codigos_extraidos = ler_codigo_multi_camadas(cv_img, nome_f)

                    banco_fotos_digitos.append({
                        "nome": nome_f,
                        "pil_img": pil_img,
                        "codigos": codigos_extraidos,
                        "digitos_list": [
                            extrair_apenas_digitos(c)
                            for c in codigos_extraidos
                            if extrair_apenas_digitos(c)
                        ],
                    })
                    prog_bar.progress((idx + 1) / len(lista_fotos))

                prog_bar.empty()

            vincularam_auto = 0
            fotos_usadas = set()

            # --- ETAPA 1: MATCHING EXATO / COMPARAÇÃO DE DÍGITOS ---
            for idx, row_data in df_temp.iterrows():
                val_raw = row_data[coluna_sgp]
                cod_excel_num = extrair_apenas_digitos(val_raw)

                if not cod_excel_num:
                    continue

                d4_excel = (
                    cod_excel_num[-4:]
                    if len(cod_excel_num) >= 4
                    else cod_excel_num
                )
                d8_excel = (
                    cod_excel_num[-8:]
                    if len(cod_excel_num) >= 8
                    else cod_excel_num
                )

                matches = []
                for foto in banco_fotos_digitos:
                    if foto["nome"] in fotos_usadas:
                        continue
                    for d in foto["digitos_list"]:
                        if d.endswith(d4_excel) or d4_excel in d:
                            matches.append(foto)
                            break

                if len(matches) == 1:
                    st.session_state.mapa_codigo_imagem[cod_excel_num] = (
                        matches[0]["pil_img"]
                    )
                    fotos_usadas.add(matches[0]["nome"])
                    vincularam_auto += 1
                elif len(matches) > 1:
                    matches_8d = [
                        f
                        for f in matches
                        if any(
                            d.endswith(d8_excel) or d8_excel in d
                            for d in f["digitos_list"]
                        )
                    ]
                    if len(matches_8d) >= 1:
                        st.session_state.mapa_codigo_imagem[cod_excel_num] = (
                            matches_8d[0]["pil_img"]
                        )
                        fotos_usadas.add(matches_8d[0]["nome"])
                        vincularam_auto += 1

            # --- ETAPA 2: FUZZY MATCHING INTELIGENTE ---
            fotos_restantes = [
                f
                for f in banco_fotos_digitos
                if f["nome"] not in fotos_usadas
            ]
            codigos_excel_sobrando = [
                c
                for c in codigos_excel_disponiveis
                if c not in st.session_state.mapa_codigo_imagem
            ]

            vinculos_fuzzy = 0
            if fotos_restantes and codigos_excel_sobrando:
                for foto in fotos_restantes:
                    melhor_score = 0
                    melhor_cod_excel = None

                    texto_foto = foto["nome"] + " " + " ".join(foto["digitos_list"])
                    digitos_foto = extrair_apenas_digitos(texto_foto)

                    for cod_ex in codigos_excel_sobrando:
                        if (
                            cod_ex
                            in st.session_state.mapa_codigo_imagem.values()
                        ):
                            continue

                        ratio_digitos = SequenceMatcher(
                            None, digitos_foto, cod_ex
                        ).ratio()
                        score_sub = (
                            0.8
                            if (cod_ex in digitos_foto or digitos_foto in cod_ex)
                            else 0.0
                        )

                        score_final = max(ratio_digitos, score_sub)

                        if score_final > melhor_score and score_final >= 0.55:
                            melhor_score = score_final
                            melhor_cod_excel = cod_ex

                    if melhor_cod_excel:
                        st.session_state.mapa_codigo_imagem[melhor_cod_excel] = (
                            foto["pil_img"]
                        )
                        fotos_usadas.add(foto["nome"])
                        codigos_excel_sobrando.remove(melhor_cod_excel)
                        vinculos_fuzzy += 1

            st.success(
                f"🎉 **Processamento concluído com alta taxa de acerto!**\n\n"
                f"• Vínculos Automáticos Diretos: **{vincularam_auto}**\n"
                f"• Vínculos por Reconhecimento Fuzzy/Recorte: **{vinculos_fuzzy}**\n"
                f"• Total de Fotos Vinculadas: **{len(st.session_state.mapa_codigo_imagem)} / {len(lista_fotos)}**"
            )

        # --- GERAÇÃO DA PLANILHA E MENU DE DIVISÃO ---
        if (
            "mapa_codigo_imagem" in st.session_state
            and len(st.session_state.mapa_codigo_imagem) > 0
        ):
            st.markdown("---")
            st.markdown("### 4. Configuração da Saída e Divisão do Excel")

            modo_divisao = st.radio(
                "Como deseja salvar/dividir a planilha?",
                [
                    "Planilha Única (Sem divisão)",
                    "Dividir por QUANTIDADE DE PARTES",
                    "Dividir por QUANTIDADE DE LINHAS",
                ],
                horizontal=True,
            )

            num_partes = 1
            linhas_por_parte = 0

            if modo_divisao == "Dividir por QUANTIDADE DE PARTES":
                num_partes = st.number_input(
                    "Digite a quantidade de partes desejada:",
                    min_value=2,
                    value=2,
                    step=1,
                )
            elif modo_divisao == "Dividir por QUANTIDADE DE LINHAS":
                linhas_por_parte = st.number_input(
                    "Digite o número de linhas por arquivo:",
                    min_value=10,
                    value=150,
                    step=10,
                )

            if st.button(
                "📊 APLICAR FOTOS E GERAR ARQUIVO(S)", use_container_width=True
            ):
                arquivo_excel.seek(0)
                wb = openpyxl.load_workbook(arquivo_excel)
                ws = wb[nome_aba]

                col_idx_codigo = None
                for col in range(1, ws.max_column + 1):
                    val_cabecalho = str(
                        ws.cell(row=1, column=col).value or ""
                    ).strip()
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

                    if (
                        cod_num
                        and cod_num in st.session_state.mapa_codigo_imagem
                    ):
                        pil_img = st.session_state.mapa_codigo_imagem[
                            cod_num
                        ].copy()
                        pil_img.thumbnail((120, 120))

                        img_byte_arr = io.BytesIO()
                        pil_img.save(img_byte_arr, format='PNG')

                        img_excel = OpenpyxlImage(img_byte_arr)
                        cell_address = ws.cell(
                            row=row, column=col_idx_foto
                        ).coordinate
                        ws.add_image(img_excel, cell_address)

                        ws.row_dimensions[row].height = 95
                        vincularam += 1
                        mapa_fotos_linha[row] = img_byte_arr

                col_letter = openpyxl.utils.get_column_letter(col_idx_foto)
                ws.column_dimensions[col_letter].width = 20
                total_dados = ws.max_row - 1

                # CASO A: PLANILHA ÚNICA
                if (
                    modo_divisao == "Planilha Única (Sem divisão)"
                    or total_dados <= 0
                ):
                    out_buffer = io.BytesIO()
                    wb.save(out_buffer)
                    out_buffer.seek(0)

                    st.success(
                        f"🎉 Concluído! **{vincularam}** fotos inseridas com sucesso."
                    )
                    st.download_button(
                        label="📥 BAIXAR PLANILHA COMPLETA (.XLSX)",
                        data=out_buffer,
                        file_name="Planilha_Com_Fotos.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                # CASO B / C: DIVISÃO EM MÚLTIPLOS ARQUIVOS (ZIP)
                else:
                    if modo_divisao == "Dividir por QUANTIDADE DE PARTES":
                        linhas_por_parte = math.ceil(total_dados / num_partes)

                    qtd_partes_calculadas = math.ceil(
                        total_dados / linhas_por_parte
                    )

                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(
                        zip_buffer, "w", zipfile.ZIP_DEFLATED
                    ) as zf:
                        for p in range(qtd_partes_calculadas):
                            dado_inicio = (p * linhas_por_parte) + 1
                            dado_fim = min(
                                total_dados, (p + 1) * linhas_por_parte
                            )

                            linha_orig_inicio = dado_inicio + 1
                            linha_orig_fim = dado_fim + 1

                            arquivo_excel.seek(0)
                            wb_p = openpyxl.load_workbook(arquivo_excel)
                            ws_p = wb_p[nome_aba]
                            ws_p._images.clear()

                            if ws_p.max_row > linha_orig_fim:
                                ws_p.delete_rows(
                                    linha_orig_fim + 1,
                                    ws_p.max_row - linha_orig_fim,
                                )
                            if linha_orig_inicio > 2:
                                ws_p.delete_rows(2, linha_orig_inicio - 2)

                            col_f_p = ws_p.max_column + 1
                            ws_p.cell(row=1, column=col_f_p).value = (
                                nome_coluna_foto
                            )

                            l_dest = 2
                            for l_orig in range(
                                linha_orig_inicio, linha_orig_fim + 1
                            ):
                                if l_orig in mapa_fotos_linha:
                                    img_b = mapa_fotos_linha[l_orig]
                                    img_b.seek(0)
                                    img_ex = OpenpyxlImage(img_b)
                                    ws_p.add_image(
                                        img_ex,
                                        f"{openpyxl.utils.get_column_letter(col_f_p)}{l_dest}",
                                    )
                                    ws_p.row_dimensions[l_dest].height = 95
                                l_dest += 1

                            ws_p.column_dimensions[
                                openpyxl.utils.get_column_letter(col_f_p)
                            ].width = 20

                            out_p = io.BytesIO()
                            wb_p.save(out_p)
                            out_p.seek(0)

                            nome_sub_arq = f"Planilha_Parte_{p+1}_(Linhas_{dado_inicio}-{dado_fim}).xlsx"
                            zf.writestr(nome_sub_arq, out_p.getvalue())

                    zip_buffer.seek(0)
                    st.success(
                        f"🎉 Divisão concluída! **{qtd_partes_calculadas}** planilhas geradas em um arquivo .ZIP."
                    )
                    st.download_button(
                        label="📥 BAIXAR PACOTE COM AS PLANILHAS DIVIDIDAS (.ZIP)",
                        data=zip_buffer,
                        file_name="Planilhas_Divididas.zip",
                        mime="application/zip",
                        use_container_width=True,
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
st.markdown(
    "<div style='text-align:center; color:#888;'>Desenvolvimento e Engenharia"
    " por <strong>Diego Costa</strong></div>",
    unsafe_allow_html=True,
)
