import io
import math
import re
import sqlite3
import zipfile
import cv2
import numpy as np
import pandas as pd
from PIL import Image, ImageOps
import streamlit as st

# --- 1. CONFIGURAÇÃO DO BANCO DE DADOS (SQLITE) ---
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


# --- 2. CONFIGURAÇÃO DA PÁGINA E CSS IDENTICO À IMAGEM ---
st.set_page_config(
    page_title="Recorte de Etiquetas",
    page_icon="✂️",
    layout="wide"
)

# Estilização exata do tema escuro
COR_FUNDO = "#262523"
COR_CARD = "#333230"
COR_LARANJA = "#F39200"

st.markdown(f"""
    <style>
    /* Ocultar elementos padrão do Streamlit */
    [data-testid="stToolbar"], [data-testid="stHeader"], header, #MainMenu {{
        display: none !important;
        visibility: hidden !important;
    }}
    .css-1544g2n, .e16nr0p33, a.header-anchor, [data-testid="stHeaderActionElements"], .stMarkdown a[href^="#"] {{
        display: none !important;
        visibility: hidden !important;
    }}
    a[href^="#"] {{
        pointer-events: none !important;
        cursor: default !important;
        text-decoration: none !important;
    }}
    
    .stApp {{
        background-color: {COR_FUNDO};
        color: #FFFFFF;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }}
    div.block-container {{
        padding-top: 2rem !important;
        padding-bottom: 2rem;
        max-width: 90%;
    }}
    
    /* Card de Métricas do Painel do Lote */
    .metric-box {{
        background-color: {COR_CARD};
        border-radius: 8px;
        padding: 18px 10px;
        text-align: center;
        margin-bottom: 12px;
        border: 1px solid rgba(255, 255, 255, 0.05);
    }}
    .metric-number {{
        font-size: 28px;
        font-weight: 800;
        color: {COR_LARANJA};
        line-height: 1;
        margin-bottom: 6px;
    }}
    .metric-label {{
        font-size: 11px;
        font-weight: 700;
        color: #A0A0A0;
        letter-spacing: 0.5px;
    }}
    
    /* Botões Globais */
    .stButton>button {{
        background: linear-gradient(90deg, {COR_LARANJA} 0%, #d88100 100%) !important;
        color: #FFFFFF !important;
        border-radius: 6px !important;
        font-weight: bold !important;
        border: none !important;
        padding: 0.5rem 1rem !important;
    }}
    </style>
""", unsafe_allow_html=True)

# --- 3. GERENCIAMENTO DE SESSÃO ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None

# --- TELA DE LOGIN / CADASTRO ---
if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.6, 1])

    with c2:
        st.markdown(f"""
            <div style="text-align: center; margin-bottom: 20px;">
                <div style="color: {COR_LARANJA}; font-size: 32px; font-weight: 900;">✂️ Recorte de Etiquetas</div>
                <p style="color: #A0A0A0; font-size: 14px; margin-top: 5px;">Processamento e recorte automático em lote</p>
            </div>
        """, unsafe_allow_html=True)

        tab_login, tab_cadastro = st.tabs(["🔒 Acessar Conta", "📝 Solicitar Cadastro"])

        with tab_login:
            with st.form("form_login"):
                usr_login = st.text_input("Usuário").strip().lower()
                senha_login = st.text_input("Senha", type="password")
                btn_entrar = st.form_submit_button("Entrar", use_container_width=True)

                if btn_entrar:
                    dados_usr = buscar_usuario(usr_login)
                    if dados_usr and dados_usr["senha"] == senha_login:
                        if dados_usr["status"] == "aprovado":
                            st.session_state.autenticado = True
                            st.session_state.usuario_logado = dados_usr
                            st.rerun()
                        elif dados_usr["status"] == "pendente":
                            st.warning("⏳ Seu cadastro está aguardando aprovação.")
                        else:
                            st.error("🚫 Acesso bloqueado.")
                    else:
                        st.error("Usuário ou senha incorretos.")

        with tab_cadastro:
            with st.form("form_cadastro"):
                novo_usr = st.text_input("Usuário").strip().lower()
                nova_senha = st.text_input("Senha", type="password")
                nome_comp = st.text_input("Nome Completo")
                contato_wa = st.text_input("WhatsApp com DDD")
                btn_cadastrar = st.form_submit_button("Cadastrar", use_container_width=True)

                if btn_cadastrar:
                    if not novo_usr or not nova_senha or not nome_comp or not contato_wa:
                        st.error("Preencha todos os campos.")
                    else:
                        sucesso, msg = cadastrar_usuario(novo_usr, nova_senha, nome_comp, contato_wa)
                        if sucesso:
                            st.success(msg)
                        else:
                            st.error(msg)
    st.stop()

# --- BARRA LATERAL ---
usr_atual = st.session_state.usuario_logado
e_admin = (usr_atual["usuario"] == "diego.costa") or (usr_atual.get("role") == "admin")

with st.sidebar:
    st.markdown(f"👤 **{usr_atual['nome_completo']}**")
    st.caption(f"Usuário: `{usr_atual['usuario']}`")
    st.markdown("---")
    if st.button("🚪 Sair", use_container_width=True):
        st.session_state.autenticado = False
        st.session_state.usuario_logado = None
        st.rerun()


# --- 4. CABEÇALHO IDÊNTICO À IMAGEM ---
col_logo, col_titulo = st.columns([1.2, 6])

with col_logo:
    st.markdown(f"""
        <div style="display:flex; justify-content:center; align-items:center; height:70px; background-color:{COR_CARD}; border-radius:8px;">
            <div style="color:#FFFFFF; font-size: 26px; font-weight:900; letter-spacing: 2px;">LOGO</div>
        </div>
    """, unsafe_allow_html=True)

with col_titulo:
    st.markdown("""
        <div style="font-size: 32px; font-weight: 800; color: #FFFFFF; margin: 0; line-height: 1.1;">Recorte de Etiquetas</div>
        <div style="margin-top: 8px; color: #B0B0B0; font-size: 14px;">Envie as fotos das etiquetas para processamento e recorte automático em lote.</div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- 5. ABAS DA APLICAÇÃO ---
if e_admin:
    tab_ferramenta, tab_admin = st.tabs(["✂️ Ferramenta de Recorte", "👑 Painel do Administrador"])
else:
    tab_ferramenta = st.tabs(["✂️ Ferramenta de Recorte"])[0]
    tab_admin = None

# ==========================================
# ABA 1: FERRAMENTA DE RECORTE
# ==========================================
with tab_ferramenta:
    col_left, col_right = st.columns([2.3, 1])

    with col_left:
        st.markdown("<div style='font-size: 15px; font-weight: 600; color: #FFFFFF; margin-bottom: 12px;'>📁 Selecione ou arraste o lote de fotos aqui</div>", unsafe_allow_html=True)
        arquivos_fotos = st.file_uploader(
            "Upload de fotos",
            type=["jpg", "jpeg", "png", "zip"],
            accept_multiple_files=True,
            label_visibility="collapsed"
        )

    # Calcular contadores
    qtd_carregadas = len(arquivos_fotos) if arquivos_fotos else 0
    qtd_recortes = 0  # Atualizado após o processamento dos recortes

    with col_right:
        st.markdown("<div style='font-size: 16px; font-weight: 700; color: #FFFFFF; margin-bottom: 12px;'>📊 Painel do Lote</div>", unsafe_allow_html=True)
        
        # Card 1: Fotos Carregadas
        st.markdown(f"""
            <div class="metric-box">
                <div class="metric-number">{qtd_carregadas}</div>
                <div class="metric-label">FOTOS CARREGADAS</div>
            </div>
        """, unsafe_allow_html=True)

        # Card 2: Recortes Prontos
        st.markdown(f"""
            <div class="metric-box">
                <div class="metric-number">{qtd_recortes}</div>
                <div class="metric-label">RECORTES PRONTOS</div>
            </div>
        """, unsafe_allow_html=True)

# ==========================================
# ABA 2: PAINEL DO ADMINISTRADOR
# ==========================================
if e_admin and tab_admin:
    with tab_admin:
        st.markdown("<div style='font-size: 20px; font-weight: bold; color: #FFFFFF; margin-bottom: 10px;'>👑 Gestão de Usuários</div>", unsafe_allow_html=True)
        df_usuarios = listar_todos_usuarios()
        st.dataframe(df_usuarios, use_container_width=True)

# --- 6. RODAPÉ IDÊNTICO À IMAGEM ---
st.markdown("<br><br><hr style='border-color: rgba(255, 255, 255, 0.1);'>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center; color:#777777; font-size: 13px;'>Desenvolvido por <strong>Diego Costa</strong></div>", unsafe_allow_html=True)
