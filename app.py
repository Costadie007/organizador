import io
import json
import math
import os
import re
import sqlite3
import urllib.parse
import zipfile
from difflib import SequenceMatcher

import cv2
import openpyxl
import pandas as pd
from openpyxl.drawing.image import Image as OpenpyxlImage
import numpy as np
from PIL import Image, ImageOps
import streamlit as st

# --- 1. CONFIGURAÇÃO DE BANCO DE DADOS (SQLITE) ---
DB_NAME = "sistema_usuarios.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            usuario TEXT PRIMARY KEY,
            senha TEXT NOT NULL,
            nome_completo TEXT,
            contato TEXT,
            status TEXT DEFAULT 'pendente',
            role TEXT DEFAULT 'user',
            data_cadastro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    # Criar o Administrador Padrão se não existir
    c.execute("SELECT * FROM usuarios WHERE usuario = ?", ("diego.costa",))
    if not c.fetchone():
        c.execute('''
            INSERT INTO usuarios (usuario, senha, nome_completo, contato, status, role)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', ("diego.costa", "admin123", "Diego Costa", "5500000000000", "aprovado", "admin"))
    conn.commit()
    conn.close()

init_db()

def buscar_usuario(usuario):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT usuario, senha, nome_completo, contato, status, role FROM usuarios WHERE usuario = ?", (usuario.lower().strip(),))
    row = c.fetchone()
    conn.close()
    if row:
        return {
            "usuario": row[0],
            "senha": row[1],
            "nome_completo": row[2],
            "contato": row[3],
            "status": row[4],
            "role": row[5]
        }
    return None

def cadastrar_usuario(usuario, senha, nome, contato):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    try:
        c.execute('''
            INSERT INTO usuarios (usuario, senha, nome_completo, contato, status, role)
            VALUES (?, ?, ?, ?, 'pendente', 'user')
        ''', (usuario.lower().strip(), senha, nome, contato))
        conn.commit()
        conn.close()
        return True, "Cadastro realizado com sucesso! Aguarde a aprovação do administrador."
    except sqlite3.IntegrityError:
        conn.close()
        return False, "Este nome de usuário já está cadastrado. Escolha outro."

def atualizar_status_db(usuario, novo_status):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    if novo_status == "excluir":
        c.execute("DELETE FROM usuarios WHERE usuario = ?", (usuario,))
    else:
        c.execute("UPDATE usuarios SET status = ? WHERE usuario = ?", (novo_status, usuario))
    conn.commit()
    conn.close()

def listar_todos_usuarios():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT usuario, nome_completo, contato, status, role, data_cadastro FROM usuarios", conn)
    conn.close()
    return df

# --- 2. MOTORES DE CÓDIGO DE BARRAS & OCR ---
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

# --- ESTILIZAÇÃO E PALETA DE CORES ---
USUARIO_ADMIN = "diego.costa"
COR_GRAFITE = "#2A2927"
COR_LARANJA = "#F39200"
COR_FUNDO_CARD = "#333230"
COR_TEXTO = "#FFFFFF"

st.set_page_config(
    page_title="Organizador de Planilhas",
    page_icon="📊",
    layout="wide"
)

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
        font-size: 15px !important;
        border: none !important;
        padding: 0.5rem 1rem !important;
        box-shadow: 0 4px 15px rgba(243, 146, 0, 0.25) !important;
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
    .card-login {{
        background-color: {COR_FUNDO_CARD};
        padding: 30px;
        border-radius: 16px;
        box-shadow: 0 8px 30px rgba(0,0,0,0.5);
        border: 1px solid #444340;
    }}
    </style>
""", unsafe_allow_html=True)

# --- GERENCIAMENTO DE SESSÃO ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None

# --- TELA DE LOGIN / CADASTRO PROFISSIONAL ---
if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.8, 1])

    with c2:
        st.markdown(f"""
            <div style="text-align: center; margin-bottom: 20px;">
                <h1 style="color: {COR_LARANJA}; font-weight: 900; margin-bottom: 0px;">📊 Organizador Pro</h1>
                <p style="color: #bbbbbb; font-size: 14px;">Gestão Inteligente e Automação de Planilhas</p>
            </div>
        """, unsafe_allow_html=True)

        tab_login, tab_cadastro = st.tabs(["🔒 Acessar Conta", "📝 Solicitar Cadastro"])

        with tab_login:
            with st.form("form_login"):
                usr_login = st.text_input("Usuário").strip().lower()
                senha_login = st.text_input("Senha", type="password")
                btn_entrar = st.form_submit_button("Entrar no Sistema", use_container_width=True)

                if btn_entrar:
                    dados_usr = buscar_usuario(usr_login)
                    if dados_usr and dados_usr["senha"] == senha_login:
                        if dados_usr["status"] == "aprovado":
                            st.session_state.autenticado = True
                            st.session_state.usuario_logado = dados_usr
                            st.rerun()
                        elif dados_usr["status"] == "pendente":
                            st.warning("⏳ Seu cadastro está em análise pelo administrador. Você receberá um aviso assim que for aprovado!")
                        else:
                            st.error("🚫 Acesso bloqueado. Entre em contato com o suporte, (61) 99669-****")
                    else:
                        st.error("Usuário ou senha incorretos.")

        with tab_cadastro:
            with st.form("form_cadastro"):
                novo_usr = st.text_input("Escolha um Usuário (sem espaços)").strip().lower()
                nova_senha = st.text_input("Escolha uma Senha", type="password")
                nome_comp = st.text_input("Nome Completo")
                contato_wa = st.text_input("WhatsApp com DDD (Ex: 11999998888)")
                
                st.caption("📱 Seu contato será usado apenas para enviar a confirmação de aprovação do seu acesso.")
                btn_cadastrar = st.form_submit_button("Enviar Solicitação de Acesso", use_container_width=True)

                if btn_cadastrar:
                    if not novo_usr or not nova_senha or not nome_comp or not contato_wa:
                        st.error("Por favor, preencha todos os campos do formulário.")
                    else:
                        sucesso, msg = cadastrar_usuario(novo_usr, nova_senha, nome_comp, contato_wa)
                        if sucesso:
                            st.success(msg)
                        else:
                            st.error(msg)
    st.stop()

# --- USUÁRIO LOGADO E PERMISSÕES ---
usr_atual = st.session_state.usuario_logado
e_admin = (usr_atual["usuario"] == USUARIO_ADMIN) or (usr_atual.get("role") == "admin")

# --- BARRA LATERAL ---
with st.sidebar:
    if e_admin:
        st.markdown(f"👤 **{usr_atual['nome_completo']}** <span class='badge-admin'>👑 ADMIN</span>", unsafe_allow_html=True)
    else:
        st.markdown(f"👤 **{usr_atual['nome_completo']}**")
    
    st.caption(f"Usuário: `{usr_atual['usuario']}`")
    st.markdown("---")
    if st.button("🚪 Sair da Conta", use_container_width=True):
        st.session_state.autenticado = False
        st.session_state.usuario_logado = None
        st.rerun()

# --- CABEÇALHO DA PLATAFORMA ---
col_logo, col_titulo = st.columns([1.2, 4])
with col_logo:
    st.markdown(f"""
        <div style="display:flex; justify-content:center; align-items:center; height:80px; background-color:{COR_FUNDO_CARD}; border-radius:10px;">
            <h1 style="color:{COR_LARANJA} !important; font-weight:900; margin:0;">PRO</h1>
        </div>
    """, unsafe_allow_html=True)

with col_titulo:
    st.markdown("""
        <h1 style="margin:0; font-size: 32px;">Organizador de Planilhas Pro Ultra</h1>
        <p style="margin:0; color:#bbbbbb !important;">SaaS Profissional: Visão Computacional, Reconhecimento Inteligente e Gestão de Acessos.</p>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- FUNÇÕES DE PROCESSAMENTO VISUAL E MATCHING ---
def extrair_apenas_digitos(texto):
    if texto is None: return ""
    return str(re.sub(r'\D', '', str(texto)))

def limpar_texto_codigo(codigo_bruto):
    if not codigo_bruto: return None
    limpo = str(codigo_bruto).replace("(", "").replace(")", "").strip()
    limpo = re.sub(r'[^a-zA-Z0-9]', '', limpo)
    return limpo if len(limpo) >= 3 else None

def gerar_variacoes_imagem(cv_img):
    variacoes = []
    if cv_img is None: return variacoes
    h, w = cv_img.shape[:2]
    if max(h, w) > 1800:
        escala = 1800 / max(h, w)
        cv_img = cv2.resize(cv_img, (int(w * escala), int(h * escala)), interpolation=cv2.INTER_AREA)

    variacoes.append(cv_img)
    cinza = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY) if len(cv_img.shape) == 3 else cv_img
    variacoes.append(cinza)

    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    variacoes.append(clahe.apply(cinza))

    nitida = cv2.filter2D(cinza, -1, np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]]))
    variacoes.append(nitida)

    _, otsu = cv2.threshold(cinza, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    variacoes.append(otsu)
    return variacoes

def tentar_decodificar_leitores(img_np):
    codigos = set()
    if img_np is None: return codigos
    try:
        for res in zxingcpp.read_barcodes(img_np):
            c = limpar_texto_codigo(res.text)
            if c: codigos.add(c)
    except Exception: pass

    if HAS_PYZBAR:
        try:
            for obj in pyzbar.decode(img_np):
                c = limpar_texto_codigo(obj.data.decode("utf-8", errors="ignore"))
                if c: codigos.add(c)
        except Exception: pass
    return codigos

def tentar_ocr_extremo(img_np):
    codigos = set()
    if not HAS_OCR or img_np is None: return codigos
    try:
        res = OCR_READER.readtext(img_np, allowlist='0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-')
        for bbox, texto, confianca in res:
            if confianca > 0.15:
                texto_corr = (texto.replace('O', '0').replace('I', '1').replace('L', '1')
                              .replace('Z', '2').replace('S', '5').replace('B', '8').replace('G', '6'))
                c = limpar_texto_codigo(texto_corr)
                if c and len(c) >= 3: codigos.add(c)
    except Exception: pass
    return codigos

def ler_imagem_todas_camadas(cv_img, nome_arquivo):
    codigos_encontrados = set()
    digs_nome = extrair_apenas_digitos(nome_arquivo)
    if digs_nome and len(digs_nome) >= 4:
        codigos_encontrados.add(digs_nome)

    if cv_img is None: return list(codigos_encontrados)

    orientacoes = [
        cv_img,
        cv2.rotate(cv_img, cv2.ROTATE_90_CLOCKWISE),
        cv2.rotate(cv_img, cv2.ROTATE_180),
        cv2.rotate(cv_img, cv2.ROTATE_90_COUNTERCLOCKWISE)
    ]

    for img_rot in orientacoes:
        variacoes = gerar_variacoes_imagem(img_rot)
        for var in variacoes:
            cods = tentar_decodificar_leitores(var)
            if cods: codigos_encontrados.update(cods)

        if codigos_encontrados: break

        if HAS_OCR:
            for var in variacoes[:3]:
                cods_ocr = tentar_ocr_extremo(var)
                if cods_ocr: codigos_encontrados.update(cods_ocr)
            if codigos_encontrados: break

    return list(codigos_encontrados)

def calcular_similaridade_avancada(digitos_excel, texto_foto_completo, digitos_foto_lista):
    if not digitos_excel: return 0.0
    digs_foto_concat = "".join(digitos_foto_lista) + extrair_apenas_digitos(texto_foto_completo)

    if digitos_excel in digs_foto_concat: return 1.0

    tam_ex = len(digitos_excel)
    if tam_ex >= 4 and len(digs_foto_concat) >= tam_ex:
        melhor_ratio = 0.0
        for i in range(len(digs_foto_concat) - tam_ex + 1):
            sub = digs_foto_concat[i:i+tam_ex]
            ratio = SequenceMatcher(None, digitos_excel, sub).ratio()
            if ratio > melhor_ratio: melhor_ratio = ratio
        if melhor_ratio >= 0.75: return melhor_ratio

    if tam_ex >= 4:
        sufixo_4 = digitos_excel[-4:]
        sufixo_6 = digitos_excel[-6:] if tam_ex >= 6 else sufixo_4
        if sufixo_6 in digs_foto_concat: return 0.90
        if sufixo_4 in digs_foto_concat: return 0.75

    return SequenceMatcher(None, digitos_excel, digs_foto_concat).ratio()

# --- MÓDULOS DA INTERFACE ---
if e_admin:
    tab_ferramenta, tab_admin = st.tabs(["⚙️ Ferramenta de Organização", "👑 Painel de Usuários e Aprovações"])
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
        nome_aba = st.selectbox("Selecione a aba da planilha:", wb_temp.sheetnames)
        df_temp = pd.read_excel(arquivo_excel, sheet_name=nome_aba)

        col1, col2 = st.columns(2)
        with col1:
            coluna_sgp = st.selectbox("Coluna com o código SGP/Principal:", list(df_temp.columns))
        with col2:
            nome_coluna_foto = st.text_input("Nome da coluna para INSERIR A FOTO:", value="FOTO")

        if st.button("🚀 INICIAR PROCESSAMENTO ULTRA AUTOMÁTICO", use_container_width=True):
            st.session_state.mapa_codigo_imagem = {}

            lista_fotos = []
            for f in arquivos_fotos:
                if f.name.endswith(".zip"):
                    with zipfile.ZipFile(f) as z:
                        for filename in z.namelist():
                            if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                                img_bytes = z.read(filename)
                                pil_img = ImageOps.exif_transpose(Image.open(io.BytesIO(img_bytes)))
                                lista_fotos.append((filename, pil_img, img_bytes))
                else:
                    img_bytes = f.read()
                    pil_img = ImageOps.exif_transpose(Image.open(io.BytesIO(img_bytes)))
                    lista_fotos.append((f.name, pil_img, img_bytes))

            banco_fotos = []
            container_status = st.container()

            with container_status:
                status_leitura = st.empty()
                prog_bar = st.progress(0)
                status_leitura.write(f"🔍 Processando {len(lista_fotos)} imagens com IA e Visão Computacional...")

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

            vincularam_auto = 0
            codigos_excel_dict = {extrair_apenas_digitos(val): val for val in df_temp[coluna_sgp].dropna().unique() if extrair_apenas_digitos(val)}

            for cod_excel, val_raw in codigos_excel_dict.items():
                for foto in banco_fotos:
                    if foto["usada"]: continue
                    if calcular_similaridade_avancada(cod_excel, foto["nome"], foto["digitos_list"]) >= 0.85:
                        st.session_state.mapa_codigo_imagem[cod_excel] = foto["pil_img"]
                        foto["usada"] = True
                        vincularam_auto += 1
                        break

            vincularam_fuzzy = 0
            for cod_excel, val_raw in codigos_excel_dict.items():
                if cod_excel in st.session_state.mapa_codigo_imagem: continue
                melhor_score, melhor_foto = 0.0, None
                for foto in banco_fotos:
                    if foto["usada"]: continue
                    score = calcular_similaridade_avancada(cod_excel, foto["nome"], foto["digitos_list"])
                    if score > melhor_score and score >= 0.45:
                        melhor_score, melhor_foto = score, foto

                if melhor_foto is not None:
                    st.session_state.mapa_codigo_imagem[cod_excel] = melhor_foto["pil_img"]
                    melhor_foto["usada"] = True
                    vincularam_fuzzy += 1

            st.success(
                f"🎉 **Processamento Concluído!**\n\n"
                f"• Vínculos Diretos: **{vincularam_auto}** | Vínculos por Aproximação: **{vincularam_fuzzy}**\n"
                f"• Total de Fotos Vinculadas: **{len(st.session_state.mapa_codigo_imagem)} / {len(lista_fotos)}**"
            )

        if "mapa_codigo_imagem" in st.session_state and len(st.session_state.mapa_codigo_imagem) > 0:
            st.markdown("---")
            st.markdown("### 4. Configuração de Saída e Divisão")
            modo_divisao = st.radio("Como deseja salvar/dividir a planilha?", ["Planilha Única (Sem divisão)", "Dividir por QUANTIDADE DE PARTES", "Dividir por QUANTIDADE DE LINHAS"], horizontal=True)

            num_partes, linhas_por_parte = 1, 0
            if modo_divisao == "Dividir por QUANTIDADE DE PARTES":
                num_partes = st.number_input("Quantidade de partes:", min_value=2, value=2, step=1)
            elif modo_divisao == "Dividir por QUANTIDADE DE LINHAS":
                linhas_por_parte = st.number_input("Número de linhas por arquivo:", min_value=10, value=150, step=10)

            if st.button("📊 APLICAR FOTOS E GERAR ARQUIVO(S)", use_container_width=True):
                arquivo_excel.seek(0)
                wb = openpyxl.load_workbook(arquivo_excel)
                ws = wb[nome_aba]

                col_idx_codigo = next(col for col in range(1, ws.max_column + 1) if str(ws.cell(row=1, column=col).value or "").strip() == str(coluna_sgp).strip())
                col_idx_foto = ws.max_column + 1
                ws.cell(row=1, column=col_idx_foto).value = nome_coluna_foto

                mapa_fotos_linha, vincularam = {}, 0
                for row in range(2, ws.max_row + 1):
                    cod_num = extrair_apenas_digitos(ws.cell(row=row, column=col_idx_codigo).value)
                    if cod_num and cod_num in st.session_state.mapa_codigo_imagem:
                        pil_img = st.session_state.mapa_codigo_imagem[cod_num].copy()
                        pil_img.thumbnail((120, 120))
                        img_byte_arr = io.BytesIO()
                        pil_img.save(img_byte_arr, format='PNG')

                        ws.add_image(OpenpyxlImage(img_byte_arr), ws.cell(row=row, column=col_idx_foto).coordinate)
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
                    st.download_button("📥 BAIXAR PLANILHA COMPLETA (.XLSX)", out_buffer, "Planilha_Com_Fotos.xlsx", use_container_width=True)
                else:
                    if modo_divisao == "Dividir por QUANTIDADE DE PARTES":
                        linhas_por_parte = math.ceil(total_dados / num_partes)
                    qtd_partes_calculadas = math.ceil(total_dados / linhas_por_parte)

                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for p in range(qtd_partes_calculadas):
                            dado_inicio = (p * linhas_por_parte) + 1
                            dado_fim = min(total_dados, (p + 1) * linhas_por_parte)
                            linha_orig_inicio, linha_orig_fim = dado_inicio + 1, dado_fim + 1

                            arquivo_excel.seek(0)
                            wb_p = openpyxl.load_workbook(arquivo_excel)
                            ws_p = wb_p[nome_aba]
                            ws_p._images.clear()

                            if ws_p.max_row > linha_orig_fim: ws_p.delete_rows(linha_orig_fim + 1, ws_p.max_row - linha_orig_fim)
                            if linha_orig_inicio > 2: ws_p.delete_rows(2, linha_orig_inicio - 2)

                            col_f_p = ws_p.max_column + 1
                            ws_p.cell(row=1, column=col_f_p).value = nome_coluna_foto

                            l_dest = 2
                            for l_orig in range(linha_orig_inicio, linha_orig_fim + 1):
                                if l_orig in mapa_fotos_linha:
                                    img_b = mapa_fotos_linha[l_orig]
                                    img_b.seek(0)
                                    ws_p.add_image(OpenpyxlImage(img_b), f"{openpyxl.utils.get_column_letter(col_f_p)}{l_dest}")
                                    ws_p.row_dimensions[l_dest].height = 95
                                l_dest += 1

                            ws_p.column_dimensions[openpyxl.utils.get_column_letter(col_f_p)].width = 20
                            out_p = io.BytesIO()
                            wb_p.save(out_p)
                            zf.writestr(f"Planilha_Parte_{p+1}_(Linhas_{dado_inicio}-{dado_fim}).xlsx", out_p.getvalue())

                    zip_buffer.seek(0)
                    st.download_button("📥 BAIXAR PACOTE DE PLANILHAS (.ZIP)", zip_buffer, "Planilhas_Divididas.zip", use_container_width=True)

# ==========================================
# ABA 2: PAINEL DE ADMINISTRAÇÃO E APROVAÇÃO
# ==========================================
if e_admin and tab_admin:
    with tab_admin:
        st.markdown("## 👑 Painel de Controle de Usuários")
        st.caption("Gerencie solicitações de acesso, aprove cadastros e notifique usuários diretamente via WhatsApp.")

        df_usuarios = listar_todos_usuarios()

        # Separar Usuários Pendentes
        pendentes = df_usuarios[df_usuarios['status'] == 'pendente']
        aprovados = df_usuarios[df_usuarios['status'] == 'aprovado']
        bloqueados = df_usuarios[df_usuarios['status'] == 'bloqueado']

        st.markdown(f"### ⏳ Solicitantes Aguardando Aprovação ({len(pendentes)})")
        
        if len(pendentes) == 0:
            st.info("Nenhuma solicitação de cadastro pendente no momento.")
        else:
            for idx, user in pendentes.iterrows():
                with st.container():
                    c_info, c_wa, c_btn1, c_btn2 = st.columns([2.5, 2, 1.2, 1.2])
                    
                    with c_info:
                        st.markdown(f"**{user['nome_completo']}** (`@{user['usuario']}`)")
                        st.caption(f"Data de Solicitação: {user['data_cadastro']}")

                    with c_wa:
                        st.markdown(f"📱 **WhatsApp:** `{user['contato']}`")

                    # Gerar Link do WhatsApp
                    num_limpo = re.sub(r'\D', '', str(user['contato']))
                    msg = urllib.parse.quote(f"Olá {user['nome_completo']}, seu cadastro no Organizador de Planilhas foi APROVADO! Você já pode acessar a plataforma.")
                    link_wa = f"https://wa.me/{num_limpo}?text={msg}"

                    with c_btn1:
                        if st.button("✅ Aprovar", key=f"aprov_{user['usuario']}"):
                            atualizar_status_db(user['usuario'], 'aprovado')
                            st.success(f"Aprovado! [Clique para Avisar no WhatsApp]({link_wa})")
                            st.rerun()

                    with c_btn2:
                        if st.button("🚫 Bloquear", key=f"rec_{user['usuario']}"):
                            atualizar_status_db(user['usuario'], 'bloqueado')
                            st.rerun()

                st.markdown("---")

        st.markdown(f"### 🟢 Usuários Ativos / Aprovados ({len(aprovados)})")
        for idx, user in aprovados.iterrows():
            c_u, c_c, c_act = st.columns([3, 2, 2])
            c_u.write(f"**{user['nome_completo']}** (`@{user['usuario']}`)")
            c_c.write(f"📱 `{user['contato']}`")
            with c_act:
                if user['usuario'] != USUARIO_ADMIN:
                    if st.button("Remover Acesso", key=f"del_{user['usuario']}"):
                        atualizar_status_db(user['usuario'], 'excluir')
                        st.rerun()

# --- RODAPÉ ---
st.markdown("<br><br>---", unsafe_allow_html=True)
st.markdown("<div style='text-align:center; color:#888;'>Desenvolvido por <strong>Diego Costa</strong></div>", unsafe_allow_html=True)
