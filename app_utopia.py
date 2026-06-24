import streamlit as st
import base64
from pathlib import Path

# ── Pasta raiz do projeto (onde está este arquivo) ──────────────────────
# Usar Path(__file__).parent garante que os assets são encontrados
# independente de onde o app é rodado — portátil entre máquinas.
ROOT  = Path(__file__).parent

def img_b64(filename: str) -> str:
    p = ROOT / filename

    if not p.exists():
        print(f"Imagem não encontrada: {p}")
        return ""

    ext = p.suffix.lower().lstrip(".")
    mime = {
        "jpg": "jpeg",
        "jpeg": "jpeg",
        "png": "png",
        "gif": "gif",
        "webp": "webp"
    }.get(ext, "png")

    data = base64.b64encode(p.read_bytes()).decode()
    return f"data:image/{mime};base64,{data}"
# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="Planejamento Energético · UTÓPIA",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# SESSION STATE
# =========================================================
if "secao" not in st.session_state: st.session_state["secao"] = None
if "page"  not in st.session_state: st.session_state["page"]  = None

# =========================================================
# QUERY PARAMS — navegação via URL
# =========================================================
params = st.query_params
if "select_secao" in params:
    st.session_state["secao"] = params["select_secao"]
    st.session_state["page"]  = None
    st.query_params.clear()
    st.rerun()
if "reset" in params:
    st.session_state["secao"] = None
    st.session_state["page"]  = None
    st.query_params.clear()
    st.rerun()

# =========================================================
# CSS GLOBAL  (paleta: branco / cinza-claro / sky-blue)
# =========================================================
ACCENT   = "#0ea5e9"   # sky-500
ACCENT_D = "#0284c7"   # sky-600 — hover
BG_PAGE  = "#f0f4f8"
BG_CARD  = "#ffffff"
TEXT_PRI = "#0f172a"
TEXT_SEC = "#64748b"

st.markdown(f"""
<style>
/* ── reset / base ────────────────────────────────── */
html, body, [class*="css"] {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}}
section[data-testid="stSidebar"] {{
    background: {BG_CARD} !important;
    border-right: 1px solid #e2e8f0 !important;
}}

/* ── cards da home ──────────────────────────────── */
.domain-link {{ text-decoration:none; display:block; width:100%; }}
.domain-btn {{
    height:280px; width:100%; border-radius:20px;
    border:1px solid #e2e8f0; overflow:hidden;
    box-shadow:0 4px 20px rgba(14,165,233,0.07), 0 1px 4px rgba(0,0,0,0.04);
    background:{BG_CARD}; transition:all 0.28s cubic-bezier(0.25,0.46,0.45,0.94);
    cursor:pointer; position:relative;
    display:flex; align-items:center; justify-content:center;
}}
.domain-link:hover .domain-btn {{
    transform:translateY(-5px);
    box-shadow:0 16px 48px rgba(14,165,233,0.18), 0 4px 12px rgba(0,0,0,0.06);
    border-color:{ACCENT};
}}
.domain-btn img {{
    width:auto; height:90%; max-width:90%; object-fit:contain;
    display:block;
    transition:transform 0.28s cubic-bezier(0.25,0.46,0.45,0.94);
}}
.domain-link:hover .domain-btn img {{
    transform:scale(1.04);
}}
/* fallback se imagem não carregar */
.domain-icon {{
    font-size:44px; line-height:1;
}}
.domain-title {{
    font-size:17px; font-weight:700; color:{TEXT_PRI};
    letter-spacing:-0.2px; text-align:center; padding:0 12px;
}}
.domain-subtitle {{
    font-size:12px; color:{TEXT_SEC}; text-align:center;
    padding:0 16px; line-height:1.5; margin-top:-8px;
}}

/* ── badge de seção ativa ────────────────────────── */
.secao-chip {{
    display:inline-flex; align-items:center; gap:8px;
    background:rgba(14,165,233,0.08); border:1px solid rgba(14,165,233,0.22);
    border-radius:20px; padding:5px 14px; margin-bottom:18px;
    font-size:13px; font-weight:600; color:{ACCENT};
}}

/* ── botões de nav (sidebar) ─────────────────────── */
section[data-testid="stSidebar"] button {{
    border-radius:10px !important;
    font-size:14px !important;
    letter-spacing:-0.1px !important;
    transition:all 0.15s ease !important;
}}
section[data-testid="stSidebar"] button[kind="secondary"] {{
    background:rgba(14,165,233,0.06) !important;
    border:none !important;
    color:{ACCENT_D} !important;
    font-weight:500 !important;
}}
section[data-testid="stSidebar"] button[kind="secondary"]:hover {{
    background:rgba(14,165,233,0.12) !important;
}}
section[data-testid="stSidebar"] button[kind="primary"] {{
    background:linear-gradient(135deg, {ACCENT}, {ACCENT_D}) !important;
    border:none !important;
    box-shadow:0 3px 12px rgba(14,165,233,0.35) !important;
    font-weight:600 !important; color:white !important;
}}

/* ── sidebar label uppercase ─────────────────────── */
.sb-label {{
    font-size:10px; font-weight:700; letter-spacing:0.12em;
    text-transform:uppercase; color:{TEXT_SEC}; padding:14px 4px 6px;
    display:block;
}}

/* ── título do sidebar ───────────────────────────── */
.sb-title {{
    font-size:16px; font-weight:700; color:{TEXT_PRI};
    padding:10px 4px 4px; letter-spacing:-0.3px;
}}
.sb-subtitle {{
    font-size:11px; color:{TEXT_SEC}; padding:0 4px 4px;
}}
</style>
""", unsafe_allow_html=True)


# =========================================================
# NAV BUTTON (sidebar)
# =========================================================
def nav_button(label: str):
    active = st.session_state.get("page") == label
    if st.sidebar.button(
        label,
        type="primary" if active else "secondary",
        key=f"nav_{st.session_state.get('secao')}_{label}",
        use_container_width=True
    ):
        st.session_state["page"] = label
        st.rerun()


# =========================================================
# SIDEBAR — cabeçalho + botão home
# =========================================================
# Logo no sidebar — portátil via img_b64
_logo_b64 = img_b64("logo_BVM.jpg")

if st.sidebar.button("↩ Início", use_container_width=True):
    st.session_state["secao"] = None
    st.session_state["page"]  = None
    st.rerun()
st.sidebar.markdown("---")



# =========================================================
# SECOES (seções do projeto)
# =========================================================
SECOES = {
    "historico": {
        "icon": "📊",
        "title": "Dados Históricos",
        "subtitle": "1975 – 2025 · 6 localidades",
        "pages": ["Catarina", "Nicodemus", "Saragoça", "Santa Bárbara", "Zona Rural", "TOTAL — UTÓPIA"],
    },
    "decenal": {
        "icon": "🔭",
        "title": "Previsão Decenal (2036)",
        "subtitle": "Projeções e cenários futuros",
        "pages": [
            "Demanda de Energia",
            "Histórico + Projeção",
            "Comparativo 2026 - 2036",
            "Barras por Cidade",
            "Composição por setor",
            "Composição iterativo",
            "Escolha em função do PIB",
            "Ranking de evolução",
            "Decomposição do crescimento",
            "Intensidade Energética",
            "Heatmap do crescimento",
            "Tabela de resultados do Matlab",
            # ── Novas pages ──────────────────────────────────────────
            "População por Ano",
            "Variável Y por Ano",
        ],
    },
    "potencial": {
        "icon": "⚡",
        "title": "Potencial Energético",
        "subtitle": "Recursos disponíveis por região",
        "pages": [],
    },
    "matriz": {
        "icon": "🌐",
        "title": "Matriz Energética",
        "subtitle": "Composição e evolução da matriz",
        "pages": [],
    },
}


# =========================================================
# HOME — grade de 4 botões
# =========================================================
def render_home():
    # ── Logo principal da empresa ──────────────────────────────────────
    _logo_b64 = img_b64("logo_BVM.jpg")
    if _logo_b64:
        # Logo centralizado, largura controlada
        st.markdown(
            f'<div style="display:flex;justify-content:center;margin-bottom:18px;">'
            f'<img src="{_logo_b64}" style="max-height:110px;max-width:480px;'
            f'object-fit:contain;border-radius:8px;" />'
            f'</div>',
            unsafe_allow_html=True,
        )
    # Subtítulo
    st.markdown(
        f'<p style="font-size:15px;color:{TEXT_SEC};text-align:center;'
        f'margin-bottom:32px;letter-spacing:0.01em;">'
        f'Acompanhamento Visual do Planejamento Energético · País de Utópia</p>',
        unsafe_allow_html=True,
    )

    # ── Mapeamento seção → arquivo de imagem ──────────────────────────
    IMG_MAP = {
        "historico": "boton_dados_historicos.png",
        "decenal":   "boton_projecao.png",
        "potencial": "boton_potencial.png",
        "matriz":    "boton_matriz.png",
    }

    keys = list(SECOES.keys())
    row1 = st.columns(2)
    row2 = st.columns(2)
    cols = [row1[0], row1[1], row2[0], row2[1]]

    for col, key in zip(cols, keys):
        s      = SECOES[key]
        img_src = img_b64(IMG_MAP.get(key, ""))

        with col:
            if img_src:
                # Estilo inline na img — não depende de seletor CSS externo
                st.markdown(
                    f'<a href="/?select_secao={key}" target="_self" class="domain-link">'
                    f'<div class="domain-btn">'
                    f'<img src="{img_src}" alt="{s["title"]}" '
                    f'style="width:auto;max-width:35%;'
                    f'object-fit:contain;display:block;margin:0 auto;" />'
                    f'</div></a>',
                    unsafe_allow_html=True,
                )
            else:
                # Fallback texto (imagem não encontrada)
                st.markdown(
                    f'<a href="/?select_secao={key}" target="_self" class="domain-link">'
                    f'<div class="domain-btn" style="flex-direction:column;'
                    f'justify-content:center;align-items:center;gap:18px;">'
                    f'<div class="domain-icon">{s["icon"]}</div>'
                    f'<div class="domain-title">{s["title"]}</div>'
                    f'<div class="domain-subtitle">{s["subtitle"]}</div>'
                    f'</div></a>',
                    unsafe_allow_html=True,
                )

    st.markdown(
        f'<p style="font-size:11px;color:#94a3b8;margin-top:40px;text-align:center;">'
        f'Desenvolvido pelo Eng. Bruno Olmedo.</p>',
        unsafe_allow_html=True,
    )
    st.stop()


# =========================================================
# EM BREVE placeholder
# =========================================================
def render_em_breve(nome: str):
    st.markdown(
        f'<div style="display:flex;flex-direction:column;align-items:center;'
        f'justify-content:center;height:420px;gap:14px;">'
        f'<span style="font-size:64px;">🚧</span>'
        f'<p style="font-size:20px;font-weight:700;color:{TEXT_PRI};">{nome}</p>'
        f'<p style="font-size:14px;color:{TEXT_SEC};">Em desenvolvimento — disponível em breve.</p>'
        f'</div>',
        unsafe_allow_html=True
    )


# =========================================================
# ROTEAMENTO PRINCIPAL
# =========================================================
secao = st.session_state.get("secao")

if secao is None:
    render_home()

# ── DADOS HISTÓRICOS ──────────────────────────────────────
elif secao == "historico":
    s = SECOES["historico"]
    st.sidebar.markdown(f'<span class="sb-label">{s["title"]}</span>', unsafe_allow_html=True)
    for p in s["pages"]:
        nav_button(p)

    # página padrão
    if st.session_state.get("page") not in s["pages"]:
        st.session_state["page"] = s["pages"][0]
        st.rerun()

    st.markdown(
        f'<div class="secao-chip">{s["icon"]} {s["title"]}</div>',
        unsafe_allow_html=True
    )

    page = st.session_state.get("page")
    from dash_historico import run_historico
    run_historico(page)

# ── EM BREVE ──────────────────────────────────────────────
elif secao == "decenal":
    s = SECOES["decenal"]
    st.sidebar.markdown(f'<span class="sb-label">{s["title"]}</span>', unsafe_allow_html=True)
    for p in s["pages"]:
        nav_button(p)

    if st.session_state.get("page") not in s["pages"]:
        st.session_state["page"] = s["pages"][0]
        st.rerun()

    st.markdown(
        f'<div class="secao-chip">{s["icon"]} {s["title"]}</div>',
        unsafe_allow_html=True,
    )

    page = st.session_state.get("page")
    from dash_projecao import run_projecao
    run_projecao(page)

elif secao == "potencial":
    s = SECOES["potencial"]
    st.markdown(f'<div class="secao-chip">{s["icon"]} {s["title"]}</div>', unsafe_allow_html=True)
    render_em_breve(s["title"])

elif secao == "matriz":
    s = SECOES["matriz"]
    st.markdown(f'<div class="secao-chip">{s["icon"]} {s["title"]}</div>', unsafe_allow_html=True)
    render_em_breve(s["title"])

else:
    render_home()
