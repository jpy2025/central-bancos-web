# ==========================================================
# app.py — Central de Bancos (Web)
# ==========================================================
# - Login seguro com expiração individual
# - Painel Administrativo (adicionar/editar usuários)
# - Bloqueio remoto via arquivo online (GitHub)
# - Interface web para processamento de bancos
# ==========================================================

import streamlit as st
import streamlit_authenticator as stauth
import importlib
import tempfile
import os
import glob
import requests
import json
from datetime import datetime
from pathlib import Path

# ==========================================================
# CONFIG INICIAL
# ==========================================================
st.set_page_config(page_title="Central de Bancos",
                   page_icon="💰", layout="wide")

USERS_FILE = Path("usuarios.json")
DEFAULT_ICON = "imagens/icone_principal.ico"
STATUS_URL = "https://raw.githubusercontent.com/jpy2025/bloqueio-central/refs/heads/main/status.txt"

# ==========================================================
# FUNÇÕES DE USUÁRIOS
# ==========================================================


def carregar_usuarios():
    if USERS_FILE.exists():
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        return {
            "admin": {
                "email": "josaeloliveira@gmail.com",
                "name": "Admin",
                "password": stauth.Hasher().hash("Jos01600"),
                "expiry_days": 100000000
            },
            "Iasmin": {
                "email": "iasmings@gmail.com",
                "name": "Iasmin",
                "password": stauth.Hasher().hash("12345"),
                "expiry_days": 100000000
            }
        }


def salvar_usuarios(users_dict):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users_dict, f, indent=4, ensure_ascii=False)


usuarios = carregar_usuarios()

# ==========================================================
# AUTENTICAÇÃO
# ==========================================================
login_config = {
    "credentials": {"usernames": usuarios},
    "cookie": {"expiry_days": 1, "key": "centralbancos_cookie", "name": "centralbancos_login"},
    "preauthorized": {"emails": []},
}

authenticator = stauth.Authenticate(
    login_config["credentials"],
    login_config["cookie"]["name"],
    login_config["cookie"]["key"],
    login_config["cookie"]["expiry_days"],
)

nome, auth_status, usuario = authenticator.login(
    location="main", fields={"Form name": "🔐 Login da Central"})

# Recriar cookie com expiração por usuário
if auth_status and usuario in usuarios:
    exp = usuarios[usuario].get("expiry_days", 1)
    authenticator.cookie_expiry_days = exp
    authenticator.cookie_manager.set_cookie(
        key=login_config["cookie"]["name"],
        value=authenticator.token,
        expires_at=datetime.now().timestamp() + (exp * 24 * 60 * 60)
    )

if auth_status is False:
    st.error("Usuário ou senha incorretos.")
    st.stop()
elif auth_status is None:
    st.warning("Digite seu usuário e senha para acessar a Central.")
    st.stop()

# Logout lateral
with st.sidebar:
    authenticator.logout("Sair", "sidebar")
    st.write(
        f"👋 Olá, **{nome}** (expira em {usuarios[usuario]['expiry_days']} dia(s))")

# ==========================================================
# BLOQUEIO REMOTO
# ==========================================================


@st.cache_data(ttl=60)
def check_remote_status(url: str) -> str:
    try:
        r = requests.get(url, timeout=5)
        return (r.text or "").strip().upper()
    except Exception:
        return "ERRO"


status = check_remote_status(STATUS_URL)
if status == "BLOQUEADO":
    st.error("🚫 O programa foi bloqueado pelo administrador.")
    st.stop()

# ==========================================================
# PAINEL ADMINISTRATIVO (somente admin)
# ==========================================================
if usuario == "admin":
    st.sidebar.markdown("---")
    st.sidebar.markdown("🧩 **Painel Administrativo**")
    if st.sidebar.button("Gerenciar Usuários"):
        st.session_state["admin_panel"] = True

if st.session_state.get("admin_panel", False) and usuario == "admin":
    st.title("👑 Painel Administrativo — Gerenciar Usuários")

    st.subheader("Usuários Existentes")
    for user, data in usuarios.items():
        with st.expander(f"👤 {data['name']} ({user})"):
            email = st.text_input(
                f"Email ({user})", data["email"], key=f"email_{user}")
            expiry = st.number_input(
                f"Dias de expiração ({user})", 1, 30, data["expiry_days"], key=f"exp_{user}")
            nova_senha = st.text_input(
                f"Nova senha ({user}) (opcional)", type="password", key=f"senha_{user}")
            col1, col2 = st.columns(2)
            with col1:
                if st.button(f"💾 Atualizar {user}"):
                    if nova_senha:
                        data["password"] = stauth.Hasher(
                            [nova_senha]).generate()[0]
                    data["email"] = email
                    data["expiry_days"] = expiry
                    salvar_usuarios(usuarios)
                    st.success(f"Usuário {user} atualizado com sucesso.")
            with col2:
                if user != "admin" and st.button(f"🗑️ Excluir {user}"):
                    del usuarios[user]
                    salvar_usuarios(usuarios)
                    st.warning(f"Usuário {user} excluído.")
                    st.experimental_rerun()

    st.divider()
    st.subheader("➕ Adicionar Novo Usuário")
    novo_user = st.text_input("Usuário (login)")
    novo_nome = st.text_input("Nome completo")
    novo_email = st.text_input("Email")
    novo_senha = st.text_input("Senha", type="password")
    novo_exp = st.number_input("Dias de expiração", 1, 30, 3)
    if st.button("Adicionar Usuário"):
        if not novo_user or not novo_senha:
            st.error("Preencha o login e senha.")
        elif novo_user in usuarios:
            st.warning("Esse usuário já existe.")
        else:
            usuarios[novo_user] = {
                "email": novo_email,
                "name": novo_nome or novo_user,
                "password": stauth.Hasher([novo_senha]).generate()[0],
                "expiry_days": novo_exp
            }
            salvar_usuarios(usuarios)
            st.success(f"Usuário {novo_user} adicionado com sucesso!")
            st.experimental_rerun()

    st.stop()  # encerra painel aqui

# ==========================================================
# INTERFACE DO SISTEMA
# ==========================================================
if "current_page" not in st.session_state:
    st.session_state.current_page = 0
if "selected_bank" not in st.session_state:
    st.session_state.selected_bank = None


def set_theme(is_dark: bool):
    st.session_state.theme = "dark" if is_dark else "light"


if "theme" not in st.session_state:
    st.session_state.theme = "light"


def inject_theme_css():
    dark = st.session_state.theme == "dark"
    bg = "#000000" if dark else "#ffffff"
    fg = "#ffa500" if dark else "#000000"
    card_bg = "#111111" if dark else "#ffffff"
    border = "#ffa500" if dark else "#ff6600"
    st.markdown(
        f"""
        <style>
            .main {{
                background-color: {bg};
                color: {fg};
            }}
            .stButton>button {{
                border: 2px solid {border};
                border-radius: 12px;
                font-weight: 700;
                padding: 0.5rem 0.9rem;
                background: {card_bg};
                color: {fg};
            }}
            .bank-card {{
                border: 1px solid {border}33;
                border-radius: 16px;
                padding: 12px 8px;
                text-align: center;
                background: {card_bg};
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_theme_css()

# Bancos
BANKS = [
    dict(nome="Asaas", icone="imagens/Asaas1.ico", modulo="Asaas"),
    dict(nome="Itaú Consolidado", icone="imagens/ItauConsolidado1.ico",
         modulo="ItauConsolidado"),
    dict(nome="Santander", icone="imagens/santander-br.ico", modulo="Santander"),
]
BANKS = sorted(BANKS, key=lambda b: b["nome"])
BANKS_PER_PAGE = 20
TOTAL_PAGES = max(1, (len(BANKS) + BANKS_PER_PAGE - 1) // BANKS_PER_PAGE)

# Sidebar
st.sidebar.title("⚙️ Opções")
is_dark = st.sidebar.toggle("🌙 Modo escuro", value=(
    st.session_state.theme == "dark"), on_change=set_theme, args=(True,))
inject_theme_css()

# Cabeçalho
st.title("🏦 Central de Bancos (Web)")
st.caption("Selecione um banco, envie os PDFs e clique em **Processar**.")


def render_grid():
    start = st.session_state.current_page * BANKS_PER_PAGE
    end = min(start + BANKS_PER_PAGE, len(BANKS))
    page_banks = BANKS[start:end]
    cols = st.columns(4)
    for idx, bank in enumerate(page_banks):
        with cols[idx % 4]:
            st.markdown('<div class="bank-card">', unsafe_allow_html=True)
            icon_path = bank["icone"] if os.path.exists(
                bank["icone"]) else DEFAULT_ICON
            st.image(icon_path, width=36)
            if st.button(bank["nome"], key=f"btn-{bank['nome']}"):
                st.session_state.selected_bank = bank
            st.markdown("</div>", unsafe_allow_html=True)


def run_bank_processor(module_name, uploaded_files):
    tmp_dir = tempfile.mkdtemp(prefix="central-bancos-")
    files = []
    for uf in uploaded_files or []:
        safe = uf.name.replace("/", "_").replace("\\", "_")
        path = os.path.join(tmp_dir, safe)
        with open(path, "wb") as f:
            f.write(uf.getbuffer())
        files.append(path)
    if not files:
        st.warning("Envie pelo menos 1 PDF.")
        return

    out_dir = os.path.join(tmp_dir, "output")
    os.makedirs(out_dir, exist_ok=True)
    progress = st.progress(0)
    log = st.empty()

    def progress_cb(p): progress.progress(max(0, min(100, int(p))))
    def log_cb(msg): log.info(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    try:
        mod = importlib.import_module(module_name)
        fn = getattr(mod, "processar_pdf_streamlit", None)
        if not callable(fn):
            st.warning(
                f"O módulo **{module_name}** não possui a função esperada.")
            return
        log_cb("Iniciando processamento...")
        fn(files, out_dir, progress_cb, log_cb)
        progress_cb(100)
        log_cb("Processamento concluído.")
    except Exception as e:
        st.error(f"❌ Erro: {e}")
        return

    excels = glob.glob(os.path.join(out_dir, "*.xlsx"))
    if not excels:
        st.info("Nenhum Excel gerado.")
    else:
        st.success("✅ Processamento finalizado! Baixe os resultados abaixo:")
        for p in excels:
            with open(p, "rb") as f:
                st.download_button(f"📥 {os.path.basename(p)}", f.read(), os.path.basename(p),
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if st.session_state.selected_bank is None:
    render_grid()
else:
    bank = st.session_state.selected_bank
    st.markdown(f"### 🏦 {bank['nome']}")
    uploaded = st.file_uploader("Selecione PDFs", type=[
                                "pdf"], accept_multiple_files=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 Processar", type="primary"):
            run_bank_processor(bank["modulo"], uploaded)
    with col2:
        if st.button("« Voltar"):
            st.session_state.selected_bank = None
            st.experimental_rerun()
