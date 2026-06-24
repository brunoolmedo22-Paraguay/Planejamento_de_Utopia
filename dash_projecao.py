"""
dash_projecao.py
======================================================================
Dashboard de Projeção de Demanda (2026–2035) — País de Utópia

Navegação via nav_button no sidebar (idêntico ao dash_historico.py).
Integra dados históricos do Excel (1975–2025) com a projeção CSV.

Pages:
  Histórico + Projeção  — série contínua 1975→2035 com toggle isolamento
  Comparativo           — KPIs lado a lado: 2026 vs ano escolhido
  Fig 2 — Séries Setor  — 6 subplots por setor, 3 cenários + banda
  Fig 1 — Barras Cidade — subplots 2×3 empilhados por cidade
  Fig 3 — Composição    — pizza + barras 2035
  Fig 4 — Interativo    — slider ano × cenário
  Fig 5 — Y vs PIB      — elasticidade-renda
  Tabela                — dados brutos + download
======================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from pathlib import Path

ROOT = Path(__file__).parent
CSV_PATH = ROOT / "projecoes_demanda.csv"
EXCEL_PATH = ROOT / "ENTREGA_DEMANDA_utopia.xlsx"

# ── Paleta ───────────────────────────────────────────────────────────
ACCENT   = "#0ea5e9"
ACCENT_D = "#0284c7"
TEXT_PRI = "#0f172a"
TEXT_SEC = "#64748b"
GRID_CLR = "rgba(226,232,240,0.6)"
BG_CHART = "#ffffff"

# Cores por setor — mesma paleta da página Histórico+Projeção (tons pastel/vivos)
SECTOR_COLORS = [
    "rgba(14,  165, 233, 1)",   # RES  — sky-500
    "rgba(245, 158,  11, 1)",   # IND  — amber-500
    "rgba(34,  197,  94, 1)",   # SER  — green-500
    "rgba(239,  68,  68, 1)",   # MIN  — red-500
    "rgba(139,  92, 246, 1)",   # AGR  — violet-500
    "rgba(236,  72, 153, 1)",   # UTIL — pink-500
]
SECTOR_COLORS_ALPHA = [c.replace(", 1)", ", 0.15)") for c in SECTOR_COLORS]

SECTORS     = ["EE_RES", "EE_IND", "EE_SER", "EE_MIN", "EE_AGR", "EE_UTIL"]
SECTOR_LBLS = ["RES",    "IND",    "SER",    "MIN",    "AGR",    "UTIL"]

# Mapeamento coluna histórico → CSV projeção
HIST_COLS = {
    "EE RES":  "EE_RES",
    "EE IND":  "EE_IND",
    "EE SER":  "EE_SER",
    "EE MIN":  "EE_MIN",
    "EE AGR":  "EE_AGR",
    "EE UTIL": "EE_UTIL",
}
HIST_TOTAL = "EE"    # coluna EE total no Excel histórico

CIDADES  = ["Catarina", "Nicodemus", "Saragoça", "Santa Bárbara", "Rural"]
ALL_LOCS = CIDADES + ["Utopia"]

# Mapa cidade → aba Excel (idêntico ao dash_historico)
SHEET_MAP = {
    "Catarina":      "RESUMO CATARINA",
    "Nicodemus":     "RESUMO NICODEMUS",
    "Saragoça":      "RESUMO SARAGOÇA",
    "Santa Bárbara": "RESUMO SANTA BÁRBARA",
    "Rural":         "RESUMO RURAL",
    "Utopia":        "RESUMO UTÓPIA",
}

CEN_NAMES = ["Referencia", "Alto", "Baixo"]
CEN_LBLS  = {
    "Referencia": "Referência",
    "Alto":       "Alto (PIB +3%)",
    "Baixo":      "Baixo (PIB -3%)",
}

# Tema Plotly (sem margin — evita TypeError)
PLOTLY_THEME = dict(
    plot_bgcolor  = BG_CHART,
    paper_bgcolor = BG_CHART,
    font          = dict(
        family="-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
        color=TEXT_PRI, size=12,
    ),
    xaxis      = dict(showgrid=True, gridcolor=GRID_CLR, linecolor="#e2e8f0", tickfont=dict(size=11)),
    yaxis      = dict(showgrid=True, gridcolor=GRID_CLR, linecolor="#e2e8f0", tickfont=dict(size=11)),
    hoverlabel = dict(bgcolor="white", font_size=12),
)


# ═══════════════════════════════════════════════════════════════════════
#  LEITURA DE DADOS
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data
def load_proj(csv_path: str) -> pd.DataFrame:
    """CSV exportado pelo MATLAB — projeções 2026–2035."""
    df = pd.read_csv(csv_path)
    df["Ano"] = df["Ano"].astype(int)
    for col in SECTORS + ["EE_TOTAL", "PIB", "AR_TOTAL", "POP"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


@st.cache_data
def load_hist(cidade: str, excel_path: str) -> pd.DataFrame:
    """
    Lê aba histórica do Excel (mesma lógica do dash_historico).
    Retorna DataFrame com colunas ANO, EE, EE RES, EE IND, EE SER,
    EE MIN, EE AGR, EE UTIL (e mais, mas estas são as que usamos).
    """
    sheet = SHEET_MAP[cidade]
    df = pd.read_excel(excel_path, sheet_name=sheet, header=0)
    df["ANO"] = pd.to_numeric(df["ANO"], errors="coerce")
    df = df[df["ANO"].notna() & (df["ANO"] > 1900)].copy()
    df["ANO"] = df["ANO"].astype(int)
    return df.sort_values("ANO").reset_index(drop=True)


# ═══════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════
def fmt_mwh(v: float) -> str:
    if abs(v) >= 1e9: return f"{v/1e9:,.3f} GWh"
    if abs(v) >= 1e6: return f"{v/1e6:,.2f} ×10⁶ MWh"
    return f"{v:,.0f} MWh"

def fmt_brl(v: float) -> str:
    if abs(v) >= 1e9: return f"Utd$ {v/1e9:,.2f} bi"
    if abs(v) >= 1e6: return f"Utd$ {v/1e6:,.2f} mi"
    return f"Utd$ {v:,.0f}"

def delta_arrow(now, prev):
    if prev == 0: return "—", TEXT_SEC
    pct  = (now - prev) / abs(prev) * 100
    arrow = "▲" if pct > 0 else "▼"
    cor   = "#22c55e" if pct > 0 else "#ef4444"
    return f"{arrow} {abs(pct):.1f}%", cor

CSS_CARD = (
    "background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;"
    "padding:14px 16px;box-shadow:0 2px 10px rgba(14,165,233,0.05);"
    "height:100%;box-sizing:border-box;"
)

def h2(texto: str, sub: str = "") -> None:
    st.markdown(
        f'<h2 style="font-size:22px;font-weight:800;color:{TEXT_PRI};'
        f'letter-spacing:-.4px;margin-bottom:2px;">{texto}</h2>'
        + (f'<p style="font-size:13px;color:{TEXT_SEC};margin-bottom:16px;">{sub}</p>'
           if sub else ""),
        unsafe_allow_html=True,
    )

def sidebar_sep(label: str) -> None:
    st.sidebar.markdown(
        f'<span style="font-size:10px;font-weight:700;text-transform:uppercase;'
        f'color:{TEXT_SEC};letter-spacing:.1em;">{label}</span>',
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: HISTÓRICO + PROJEÇÃO  (nova página principal)
# ═══════════════════════════════════════════════════════════════════════
def page_hist_proj(df_proj: pd.DataFrame):
    """
    Série contínua 1975→2035 integrando histórico (Excel) + projeção (CSV).
    Toggle no sidebar: "Isolar Projeção" oculta o histórico.
    """
    h2("📊 Histórico + Projeção",
       "Série contínua 1975 → 2035 · Histórico real + projeção bottom-up")

    # ── Controles sidebar ───────────────────────────────────────────
    st.sidebar.markdown("---")
    sidebar_sep("Histórico + Projeção")

    local_sel = st.sidebar.selectbox(
        "📍 Localidade", ALL_LOCS, index=len(ALL_LOCS)-1, key="hp_local"
    )
    cenario_sel = st.sidebar.radio(
        "📊 Cenário", CEN_NAMES,
        format_func=lambda c: CEN_LBLS[c], key="hp_cenario",
    )
    isolar = st.sidebar.toggle(
        "🔍 Isolar Projeção", value=False, key="hp_isolar",
        help="Oculta o período histórico e exibe apenas 2026–2035",
    )
    mostrar_cenarios = st.sidebar.toggle(
        "📉 Mostrar Alto/Baixo", value=True, key="hp_cen_toggle",
        help="Exibe as bandas de cenário Alto e Baixo",
    )
    mostrar_setores = st.sidebar.toggle(
        "🏭 Detalhar por Setor", value=False, key="hp_setores",
        help="Exibe linhas individuais de cada setor (somente projeção)",
    )

    # ── Dados históricos ────────────────────────────────────────────
    hist_ok = False
    df_hist = None
    try:
        df_hist = load_hist(local_sel, EXCEL_PATH)
        hist_ok = True
    except Exception:
        if not isolar:
            st.warning(
                f"⚠️ Histórico não carregado (`{EXCEL_PATH}` não encontrado). "
                "Ative **Isolar Projeção** ou coloque o Excel na pasta do app."
            )

    # ── Dados projeção ──────────────────────────────────────────────
    ref  = df_proj[(df_proj["Cenario"] == "Referencia") & (df_proj["Local"] == local_sel)].sort_values("Ano")
    alto = df_proj[(df_proj["Cenario"] == "Alto")        & (df_proj["Local"] == local_sel)].sort_values("Ano")
    baix = df_proj[(df_proj["Cenario"] == "Baixo")       & (df_proj["Local"] == local_sel)].sort_values("Ano")

    anos_f  = ref["Ano"].tolist()
    ano_anc = anos_f[0]   # 2026 — ponto de ancoragem

    # ── FIGURA TOTAL ────────────────────────────────────────────────
    fig = go.Figure()

    # Linha vertical separando histórico de projeção
    fig.add_vline(
        x=ano_anc - 0.5,
        line=dict(color="#94a3b8", width=1, dash="dot"),
        annotation_text="↑ início da projeção",
        annotation_font_size=10,
        annotation_font_color=TEXT_SEC,
    )

    if hist_ok and not isolar:
        # Histórico — linha cinza-azulada sólida
        anos_h = df_hist["ANO"].tolist()

        # Verifica existência da coluna de total
        col_total = HIST_TOTAL if HIST_TOTAL in df_hist.columns else None
        if col_total:
            # Ponto âncora: 2025 conecta histórico com projeção
            anos_h_ext  = anos_h + [ano_anc]
            vals_h_ext  = df_hist[col_total].tolist() + [ref["EE_TOTAL"].iloc[0]]

            fig.add_trace(go.Scatter(
                x=anos_h_ext, y=[v/1e6 for v in vals_h_ext],
                mode="lines",
                name="Histórico (Total)",
                line=dict(color="#475569", width=2, dash="solid"),
                hovertemplate="Hist. %{x}: %{y:,.2f} ×10⁶ MWh<extra></extra>",
            ))

        # Setores históricos (se solicitado)
        if mostrar_setores:
            for hcol, pcol, slbl, sclr in zip(
                HIST_COLS.keys(), HIST_COLS.values(),
                SECTOR_LBLS, SECTOR_COLORS
            ):
                if hcol in df_hist.columns:
                    fig.add_trace(go.Scatter(
                        x=df_hist["ANO"].tolist(),
                        y=(df_hist[hcol]/1e6).tolist(),
                        mode="lines",
                        name=f"Hist. {slbl}",
                        line=dict(color=sclr.replace(", 1)", ", 0.5)"), width=1.2, dash="dot"),
                        showlegend=True,
                        hovertemplate=f"Hist. {slbl} %{{x}}: %{{y:,.2f}}<extra></extra>",
                    ))

    # Projeção — EE Total (referência)
    # âncora: inclui 2025 se histórico disponível, senão começa em 2026
    if hist_ok and not isolar and col_total:
        v_anc = df_hist[col_total].iloc[-1] / 1e6
        anos_ref_plot = [df_hist["ANO"].iloc[-1]] + anos_f
        vals_ref_plot = [v_anc] + (ref["EE_TOTAL"]/1e6).tolist()
    else:
        anos_ref_plot = anos_f
        vals_ref_plot = (ref["EE_TOTAL"]/1e6).tolist()

    fig.add_trace(go.Scatter(
        x=anos_ref_plot, y=vals_ref_plot,
        mode="lines+markers",
        name="Projeção — Referência",
        line=dict(color=ACCENT, width=2.8),
        marker=dict(size=5, color=ACCENT),
        hovertemplate="Ref. %{x}: %{y:,.2f} ×10⁶ MWh<extra></extra>",
    ))

    # Cenários Alto e Baixo
    if mostrar_cenarios:
        # Banda sombreada
        x_band = anos_ref_plot + anos_ref_plot[::-1]
        if hist_ok and not isolar and col_total:
            v_alto_anc = [v_anc] + (alto["EE_TOTAL"]/1e6).tolist()
            v_baix_anc = [v_anc] + (baix["EE_TOTAL"]/1e6).tolist()
        else:
            v_alto_anc = (alto["EE_TOTAL"]/1e6).tolist()
            v_baix_anc = (baix["EE_TOTAL"]/1e6).tolist()

        fig.add_trace(go.Scatter(
            x=x_band, y=v_alto_anc + v_baix_anc[::-1],
            fill="toself",
            fillcolor="rgba(14,165,233,0.08)",
            line=dict(color="rgba(0,0,0,0)"),
            name="Banda ±PIB 3%",
            hoverinfo="skip",
        ))

        # Linhas individuais
        fig.add_trace(go.Scatter(
            x=anos_ref_plot, y=v_alto_anc,
            mode="lines", name="Alto (PIB +3%)",
            line=dict(color=ACCENT, width=1.3, dash="dash"),
            hovertemplate="Alto %{x}: %{y:,.2f}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=anos_ref_plot, y=v_baix_anc,
            mode="lines", name="Baixo (PIB -3%)",
            line=dict(color=ACCENT_D, width=1.3, dash="dot"),
            hovertemplate="Baixo %{x}: %{y:,.2f}<extra></extra>",
        ))

    # Setores da projeção
    if mostrar_setores:
        for sec, slbl, sclr in zip(SECTORS, SECTOR_LBLS, SECTOR_COLORS):
            if hist_ok and not isolar and hcol in df_hist.columns:
                hist_sec_col = {v: k for k, v in HIST_COLS.items()}.get(sec)
                if hist_sec_col and hist_sec_col in df_hist.columns:
                    v_s_anc = df_hist[hist_sec_col].iloc[-1] / 1e6
                    anos_s  = [df_hist["ANO"].iloc[-1]] + anos_f
                    vals_s  = [v_s_anc] + (ref[sec]/1e6).tolist()
                else:
                    anos_s = anos_f
                    vals_s = (ref[sec]/1e6).tolist()
            else:
                anos_s = anos_f
                vals_s = (ref[sec]/1e6).tolist()

            fig.add_trace(go.Scatter(
                x=anos_s, y=vals_s,
                mode="lines+markers",
                name=f"Proj. {slbl}",
                line=dict(color=sclr, width=1.8),
                marker=dict(size=4),
                hovertemplate=f"Proj. {slbl} %{{x}}: %{{y:,.2f}}<extra></extra>",
            ))

    fig.update_layout(
        title=dict(
            text=f"Demanda Total — {local_sel}  ·  {CEN_LBLS[cenario_sel]}",
            font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0,
        ),
        yaxis_title="EE Total (×10⁶ MWh)",
        height=420,
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, font=dict(size=10)),
        margin=dict(l=15, r=15, t=60, b=15),
        **PLOTLY_THEME,
    )
    st.plotly_chart(fig, use_container_width=True, key="hp_total")

    # ── GRADE 2×3: um subplot por setor ─────────────────────────────
    st.markdown("---")
    st.markdown("#### Detalhe por Setor")

    fig2 = make_subplots(
        rows=2, cols=3,
        subplot_titles=[f"Setor {s}" for s in SECTOR_LBLS],
        vertical_spacing=0.14, horizontal_spacing=0.07,
    )

    for idx, (sec, slbl, sclr, sclr_a) in enumerate(
        zip(SECTORS, SECTOR_LBLS, SECTOR_COLORS, SECTOR_COLORS_ALPHA)
    ):
        row = idx // 3 + 1
        col = idx %  3 + 1
        show = (idx == 0)

        hist_col_name = {v: k for k, v in HIST_COLS.items()}.get(sec)

        # Histórico do setor
        if hist_ok and not isolar and hist_col_name and hist_col_name in df_hist.columns:
            # Conecta até o primeiro ponto da projeção
            anos_h_s   = df_hist["ANO"].tolist() + [ano_anc]
            vals_h_s   = (df_hist[hist_col_name]/1e6).tolist() + [ref[sec].iloc[0]/1e6]
            fig2.add_trace(go.Scatter(
                x=anos_h_s, y=vals_h_s,
                mode="lines", name="Histórico",
                line=dict(color="#94a3b8", width=1.5),
                legendgroup="hist_sec",
                showlegend=show,
                hovertemplate=f"Hist. {slbl} %{{x}}: %{{y:,.2f}}<extra></extra>",
            ), row=row, col=col)
            v_anc_s = df_hist[hist_col_name].iloc[-1] / 1e6
            anos_proj_s = [df_hist["ANO"].iloc[-1]] + anos_f
            vals_ref_s  = [v_anc_s] + (ref[sec]/1e6).tolist()
            vals_alto_s = [v_anc_s] + (alto[sec]/1e6).tolist()
            vals_baix_s = [v_anc_s] + (baix[sec]/1e6).tolist()
        else:
            anos_proj_s = anos_f
            vals_ref_s  = (ref[sec]/1e6).tolist()
            vals_alto_s = (alto[sec]/1e6).tolist()
            vals_baix_s = (baix[sec]/1e6).tolist()

        # Banda ±
        if mostrar_cenarios:
            x_b = anos_proj_s + anos_proj_s[::-1]
            y_b = vals_alto_s + vals_baix_s[::-1]
            fig2.add_trace(go.Scatter(
                x=x_b, y=y_b, fill="toself",
                fillcolor=sclr_a, line=dict(color="rgba(0,0,0,0)"),
                name="Banda ±3%", legendgroup="banda_sec",
                showlegend=show, hoverinfo="skip",
            ), row=row, col=col)

        # Linha referência
        fig2.add_trace(go.Scatter(
            x=anos_proj_s, y=vals_ref_s,
            mode="lines+markers", name="Projeção Ref.",
            line=dict(color=sclr, width=2.4), marker=dict(size=4),
            legendgroup="ref_sec", showlegend=show,
            hovertemplate=f"Ref. {slbl} %{{x}}: %{{y:,.2f}}<extra></extra>",
        ), row=row, col=col)

        # Alto / Baixo
        if mostrar_cenarios:
            fig2.add_trace(go.Scatter(
                x=anos_proj_s, y=vals_alto_s, mode="lines",
                name="Alto", line=dict(color=sclr, width=1.2, dash="dash"),
                legendgroup="alto_sec", showlegend=show,
                hovertemplate=f"Alto {slbl} %{{x}}: %{{y:,.2f}}<extra></extra>",
            ), row=row, col=col)
            fig2.add_trace(go.Scatter(
                x=anos_proj_s, y=vals_baix_s, mode="lines",
                name="Baixo", line=dict(color=sclr, width=1.2, dash="dot"),
                legendgroup="baix_sec", showlegend=show,
                hovertemplate=f"Baixo {slbl} %{{x}}: %{{y:,.2f}}<extra></extra>",
            ), row=row, col=col)

    fig2.update_layout(
        title=dict(
            text=f"Histórico + Projeção por Setor — {local_sel}",
            font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0,
        ),
        height=680,
        legend=dict(orientation="h", yanchor="bottom", y=-0.10,
                    xanchor="center", x=0.5, font=dict(size=10)),
        **PLOTLY_THEME,
    )
    fig2.update_yaxes(title_text="EE (×10⁶ MWh)", col=1)
    st.plotly_chart(fig2, use_container_width=True, key="hp_setores_grid")


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: COMPARATIVO  (KPIs 2026 vs ano escolhido)
# ═══════════════════════════════════════════════════════════════════════
def page_comparativo(df: pd.DataFrame):
    h2("📊 Comparativo de Projeção",
       "Indicadores lado a lado — base 2026 vs ano selecionado")

    st.sidebar.markdown("---")
    sidebar_sep("Comparativo")

    local_sel = st.sidebar.selectbox(
        "📍 Localidade", ALL_LOCS, index=len(ALL_LOCS)-1, key="cmp_local"
    )
    cenario_sel = st.sidebar.radio(
        "📊 Cenário", CEN_NAMES,
        format_func=lambda c: CEN_LBLS[c], key="cmp_cenario",
    )

    sub = df[(df["Cenario"] == cenario_sel) & (df["Local"] == local_sel)].sort_values("Ano")
    if sub.empty:
        st.warning("Sem dados para esta seleção.")
        return

    anos = sorted(sub["Ano"].unique().tolist())
    ano_ini = anos[0]
    ano_fim = st.sidebar.slider(
        "📅 Ano de comparação",
        min_value=anos[0], max_value=anos[-1], value=anos[-1],
        step=1, key="cmp_ano",
    )

    r_ini = sub[sub["Ano"] == ano_ini].iloc[0]
    r_fim = sub[sub["Ano"] == ano_fim].iloc[0]

    # Cabeçalho da tabela de comparação
    st.markdown("---")
    col_lbl, col_ini, col_delta, col_fim = st.columns([1.8, 2.5, 1.4, 2.5])
    with col_lbl:
        st.markdown(f'<div style="font-size:11px;font-weight:700;color:{TEXT_SEC};'
                    f'text-transform:uppercase;padding-bottom:6px;">Indicador</div>',
                    unsafe_allow_html=True)
    with col_ini:
        st.markdown(f'<div style="font-size:15px;font-weight:800;color:{TEXT_PRI};'
                    f'padding-bottom:6px;">◀ {ano_ini} (Base)</div>',
                    unsafe_allow_html=True)
    with col_delta:
        st.markdown(f'<div style="font-size:11px;font-weight:700;color:{TEXT_SEC};'
                    f'text-transform:uppercase;padding-bottom:6px;">Variação</div>',
                    unsafe_allow_html=True)
    with col_fim:
        st.markdown(f'<div style="font-size:15px;font-weight:800;color:{ACCENT};'
                    f'padding-bottom:6px;">{ano_fim} ▶ ({CEN_LBLS[cenario_sel]})</div>',
                    unsafe_allow_html=True)

    indicadores = [
        ("⚡ EE Total",        "EE_TOTAL", fmt_mwh),
        ("🏠 EE Residencial",  "EE_RES",   fmt_mwh),
        ("🏭 EE Industrial",   "EE_IND",   fmt_mwh),
        ("🏪 EE Serviços",     "EE_SER",   fmt_mwh),
        ("⛏️ EE Mineral",      "EE_MIN",   fmt_mwh),
        ("🌾 EE Agropecuário", "EE_AGR",   fmt_mwh),
        ("💡 EE Utilidades",   "EE_UTIL",  fmt_mwh),
        ("💰 PIB",             "PIB",      fmt_brl),
        ("👥 População",       "POP",      lambda v: f"{v:,.0f} hab"),
        ("🗺️ Área",           "AR_TOTAL", lambda v: f"{v:,.2f} km²"),
    ]

    def row_css(i):
        bg = "#f8fafc" if i % 2 == 0 else "#ffffff"
        return (f"background:{bg};border-radius:10px;padding:10px 14px;"
                "margin-bottom:4px;border:1px solid #f1f5f9;")

    for i, (lbl, campo, fmt) in enumerate(indicadores):
        v_ini = r_ini[campo]
        v_fim = r_fim[campo]
        dlt_txt, dlt_cor = delta_arrow(v_fim, v_ini)
        col_lbl, col_ini_c, col_delta_c, col_fim_c = st.columns([1.8, 2.5, 1.4, 2.5])
        with col_lbl:
            st.markdown(f'<div style="{row_css(i)}font-size:12px;font-weight:600;'
                        f'color:{TEXT_PRI};">{lbl}</div>', unsafe_allow_html=True)
        with col_ini_c:
            st.markdown(f'<div style="{row_css(i)}font-size:14px;font-weight:700;'
                        f'color:{TEXT_PRI};">{fmt(v_ini)}</div>', unsafe_allow_html=True)
        with col_delta_c:
            st.markdown(f'<div style="{row_css(i)}font-size:13px;font-weight:700;'
                        f'color:{dlt_cor};text-align:center;">{dlt_txt}</div>',
                        unsafe_allow_html=True)
        with col_fim_c:
            st.markdown(f'<div style="{row_css(i)}font-size:14px;font-weight:700;'
                        f'color:{ACCENT};">{fmt(v_fim)}</div>', unsafe_allow_html=True)

    # Mini gráfico trajetória todos cenários
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### Trajetória da EE Total — todos os cenários")
    fig_traj = go.Figure()
    cen_styles = {
        "Referencia": dict(color=ACCENT,    width=2.8, dash="solid"),
        "Alto":        dict(color="#22c55e", width=1.5, dash="dash"),
        "Baixo":       dict(color="#f59e0b", width=1.5, dash="dot"),
    }
    for cen in CEN_NAMES:
        s_cen = df[(df["Cenario"] == cen) & (df["Local"] == local_sel)].sort_values("Ano")
        fig_traj.add_trace(go.Scatter(
            x=s_cen["Ano"], y=s_cen["EE_TOTAL"]/1e6,
            mode="lines+markers", name=CEN_LBLS[cen],
            line=cen_styles[cen], marker=dict(size=4),
            hovertemplate=f"{CEN_LBLS[cen]}: %{{y:,.3f}} ×10⁶ MWh<extra></extra>",
        ))
    v_sel = sub[sub["Ano"] == ano_fim]["EE_TOTAL"].values[0] / 1e6
    fig_traj.add_trace(go.Scatter(
        x=[ano_fim], y=[v_sel], mode="markers",
        name=f"Selecionado ({ano_fim})",
        marker=dict(size=12, color=ACCENT, symbol="diamond",
                    line=dict(color="white", width=2)),
        hovertemplate=f"Ano {ano_fim}: %{{y:,.3f}}<extra></extra>",
    ))
    fig_traj.update_layout(
        yaxis_title="EE Total (×10⁶ MWh)", height=280,
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, font=dict(size=10)),
        margin=dict(l=15, r=15, t=30, b=15),
        **PLOTLY_THEME,
    )
    st.plotly_chart(fig_traj, use_container_width=True, key="traj_cmp")




# ═══════════════════════════════════════════════════════════════════════
#  PAGE: FIG 1 — barras por cidade
# ═══════════════════════════════════════════════════════════════════════
def page_fig1(df: pd.DataFrame):
    h2("🏙️ Barras por Cidade",
       "Barras empilhadas por setor · Banda = intervalo Alto/Baixo")

    st.sidebar.markdown("---")
    sidebar_sep("Fig 1")
    cenario_sel = st.sidebar.radio(
        "📊 Cenário", CEN_NAMES,
        format_func=lambda c: CEN_LBLS[c], key="f1_cenario",
    )

    locs    = CIDADES + ["Utopia"]
    ref_df  = df[df["Cenario"] == "Referencia"]
    alto_df = df[df["Cenario"] == "Alto"]
    baix_df = df[df["Cenario"] == "Baixo"]
    sel_df  = df[df["Cenario"] == cenario_sel]
    show_lg = True

    fig = make_subplots(
        rows=2, cols=3,
        subplot_titles=["Catarina","Nicodemus","Saragoça",
                        "Santa Bárbara","Rural","UTÓPIA (Total Nacional)"],
        vertical_spacing=0.14, horizontal_spacing=0.07,
    )
    for i, loc in enumerate(locs):
        row = i // 3 + 1
        col = i %  3 + 1
        r_loc   = ref_df [ref_df ["Local"]==loc].sort_values("Ano")
        al_loc  = alto_df[alto_df["Local"]==loc].sort_values("Ano")
        bx_loc  = baix_df[baix_df["Local"]==loc].sort_values("Ano")
        sel_loc = sel_df [sel_df ["Local"]==loc].sort_values("Ano")
        anos = sel_loc["Ano"].tolist()
        for sec, slbl, sclr in zip(SECTORS, SECTOR_LBLS, SECTOR_COLORS):
            fig.add_trace(go.Bar(
                x=anos, y=sel_loc[sec]/1e6, name=slbl,
                marker_color=sclr, marker_line_width=0,
                legendgroup=slbl, showlegend=show_lg,
                hovertemplate=f"<b>{slbl}</b> %{{x}}: %{{y:,.2f}}<extra></extra>",
            ), row=row, col=col)
        show_lg = False
        x_band = anos + anos[::-1]
        y_band = (al_loc["EE_TOTAL"]/1e6).tolist() + (bx_loc["EE_TOTAL"]/1e6).tolist()[::-1]
        fig.add_trace(go.Scatter(x=x_band, y=y_band, fill="toself",
            fillcolor="rgba(100,100,100,0.09)", line=dict(color="rgba(0,0,0,0)"),
            name="Banda ±3%", legendgroup="banda",
            showlegend=(i==0), hoverinfo="skip",
        ), row=row, col=col)
        fig.add_trace(go.Scatter(
            x=anos, y=r_loc["EE_TOTAL"]/1e6, mode="lines+markers",
            name="Total Ref.", line=dict(color="black", width=1.6, dash="dash"),
            marker=dict(size=3), legendgroup="total_ref", showlegend=(i==0),
            hovertemplate="Total Ref. %{x}: %{y:,.2f}<extra></extra>",
        ), row=row, col=col)

    fig.update_layout(
        barmode="stack",
        title=dict(text=f"Demanda por Cidade · {CEN_LBLS[cenario_sel]}",
                   font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0),
        height=700, **PLOTLY_THEME,
        legend=dict(orientation="h", yanchor="bottom", y=-0.10,
                    xanchor="center", x=0.5, font=dict(size=10)),
    )
    fig.update_yaxes(title_text="EE (×10⁶ MWh)", col=1)
    st.plotly_chart(fig, use_container_width=True, key="fig1_main")


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: FIG 3 — composição setorial
# ═══════════════════════════════════════════════════════════════════════
def page_fig3(df: pd.DataFrame):
    h2("🥧 Composição Setorial",
       "Pizza: participação média 2026–2035 · Barras: demanda em 2035 por cidade")

    fig = make_subplots(
        rows=1, cols=2,
        specs=[[{"type": "pie"}, {"type": "bar"}]],
        subplot_titles=["Participação setorial média 2026–2035 (Utópia)",
                        "Demanda por cidade e setor — 2035"],
        column_widths=[0.42, 0.58],
    )
    ut_ref = df[(df["Cenario"]=="Referencia") & (df["Local"]=="Utopia")]
    medias = [ut_ref[s].mean() for s in SECTORS]
    valid  = [(l,m,c) for l,m,c in zip(SECTOR_LBLS,medias,SECTOR_COLORS) if m>0]
    lv,mv,cv = zip(*valid)
    fig.add_trace(go.Pie(
        labels=lv, values=mv, marker_colors=cv,
        textinfo="label+percent", hole=0.32,
        hovertemplate="%{label}: %{value:,.0f} MWh<br>%{percent}<extra></extra>",
    ), row=1, col=1)

    ref35 = df[(df["Cenario"]=="Referencia")&(df["Ano"]==2035)&(df["Local"].isin(CIDADES))]
    for sec, slbl, sclr in zip(SECTORS, SECTOR_LBLS, SECTOR_COLORS):
        vals = [ref35[ref35["Local"]==c][sec].values[0]/1e6
                if len(ref35[ref35["Local"]==c]) else 0 for c in CIDADES]
        fig.add_trace(go.Bar(
            x=vals, y=CIDADES, orientation="h", name=slbl,
            marker_color=sclr, marker_line_width=0, legendgroup=slbl,
            hovertemplate=f"<b>{slbl}</b>: %{{x:,.2f}}<extra></extra>",
        ), row=1, col=2)

    fig.update_layout(
        barmode="stack",
        title=dict(text="Composição Setorial — Utópia",
                   font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0),
        height=440, **PLOTLY_THEME,
        legend=dict(orientation="h", yanchor="bottom", y=-0.22,
                    xanchor="center", x=0.5, font=dict(size=10)),
    )
    fig.update_xaxes(title_text="EE (×10⁶ MWh)", row=1, col=2)
    st.plotly_chart(fig, use_container_width=True, key="fig3_main")


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: FIG 4 — dashboard interativo
# ═══════════════════════════════════════════════════════════════════════
def page_fig4(df: pd.DataFrame):
    h2("🎛️ Dashboard Interativo — Ano × Cenário",
       "Replica o slider + dropdown do modelo MATLAB")

    anos_disp = sorted(df["Ano"].unique().tolist())
    c1, c2 = st.columns([3, 2])
    with c1:
        ano_sel = st.slider("📅 Ano", min_value=anos_disp[0], max_value=anos_disp[-1],
                            value=anos_disp[0], step=1, key="f4_ano")
    with c2:
        cen_sel = st.selectbox("📊 Cenário", CEN_NAMES,
                               format_func=lambda c: CEN_LBLS[c], key="f4_cen")

    subset = df[(df["Cenario"]==cen_sel) & (df["Ano"]==ano_sel) & (df["Local"].isin(CIDADES))]
    fig = go.Figure()
    for sec, slbl, sclr in zip(SECTORS, SECTOR_LBLS, SECTOR_COLORS):
        vals = [subset[subset["Local"]==c][sec].values[0]/1e6
                if len(subset[subset["Local"]==c]) else 0 for c in CIDADES]
        fig.add_trace(go.Bar(x=CIDADES, y=vals, name=slbl,
            marker_color=sclr, marker_line_width=0,
            hovertemplate=f"<b>{slbl}</b><br>%{{x}}: %{{y:,.2f}}<extra></extra>",
        ))
    totais = [subset[subset["Local"]==c]["EE_TOTAL"].values[0]/1e6
              if len(subset[subset["Local"]==c]) else 0 for c in CIDADES]
    fig.add_trace(go.Scatter(x=CIDADES, y=totais, mode="markers+lines", name="Total",
        line=dict(color="black", width=1.8, dash="dot"),
        marker=dict(size=7, symbol="diamond"),
        hovertemplate="Total %{x}: %{y:,.2f}<extra></extra>",
    ))
    fig.update_layout(
        barmode="stack",
        title=dict(text=f"Demanda por Cidade — {ano_sel}  ·  {CEN_LBLS[cen_sel]}",
                   font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0),
        yaxis_title="EE (×10⁶ MWh)", height=420,
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, font=dict(size=10)),
        margin=dict(l=15, r=15, t=60, b=15),
        **PLOTLY_THEME,
    )
    st.plotly_chart(fig, use_container_width=True, key="fig4_main")

    st.markdown(f"##### Valores — {ano_sel} · {CEN_LBLS[cen_sel]}")
    disp = subset[["Local"]+SECTORS+["EE_TOTAL"]].copy()
    for s in SECTORS+["EE_TOTAL"]:
        disp[s] = disp[s].apply(lambda v: f"{v:,.0f}")
    disp.columns = ["Cidade","Res","Ind","Ser","Min","Agr","Util","TOTAL"]
    st.dataframe(disp, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: FIG 5 — Y vs PIB
# ═══════════════════════════════════════════════════════════════════════
def page_fig5(df: pd.DataFrame):
    h2("🔗 Elasticidade-Renda · Y vs PIB",
       "Eixo X = PIB (variável livre do modelo) · cada curva = um cenário")

    st.sidebar.markdown("---")
    sidebar_sep("Y vs PIB")
    local_sel = st.sidebar.selectbox(
        "📍 Localidade", ALL_LOCS, index=len(ALL_LOCS)-1, key="f5_local"
    )
    var_opts = {
        "EE Total": "EE_TOTAL", "EE Residencial": "EE_RES",
        "EE Industrial": "EE_IND", "EE Serviços": "EE_SER",
        "EE Mineral": "EE_MIN", "EE Agropecuário": "EE_AGR",
        "EE Utilidades": "EE_UTIL", "Área Total (km²)": "AR_TOTAL",
        "População": "POP",
    }
    var_lbl = st.selectbox("Variável do Eixo Y", list(var_opts.keys()), key="f5_var")
    var_col = var_opts[var_lbl]

    cen_colors = {"Referencia": ACCENT, "Alto": "#22c55e", "Baixo": "#f59e0b"}
    fig = go.Figure()
    for cen in CEN_NAMES:
        sub = df[(df["Cenario"]==cen) & (df["Local"]==local_sel)].sort_values("Ano")
        if sub.empty: continue
        y_vals = sub[var_col]/1e6 if "EE" in var_col else sub[var_col]
        fig.add_trace(go.Scatter(
            x=sub["PIB"]/1e6, y=y_vals, mode="lines+markers",
            name=CEN_LBLS[cen], line=dict(color=cen_colors[cen], width=2.2),
            marker=dict(size=5), customdata=sub["Ano"],
            hovertemplate=(f"<b>{CEN_LBLS[cen]}</b><br>PIB: %{{x:,.2f}}<br>"
                           f"Valor: %{{y:,.2f}}<br>Ano: %{{customdata}}<extra></extra>"),
        ))
    y_unit = "×10⁶ MWh" if "EE" in var_col else ("hab" if var_col=="POP" else "km²")
    fig.update_layout(
        title=dict(text=f"{var_lbl}  vs  PIB — {local_sel}",
                   font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0),
        xaxis_title="PIB (Utd$ ×10⁶)", yaxis_title=f"{var_lbl} ({y_unit})",
        height=420, legend=dict(font=dict(size=11)),
        margin=dict(l=15, r=15, t=50, b=15),
        **PLOTLY_THEME,
    )
    st.plotly_chart(fig, use_container_width=True, key="fig5_main")


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: DEMANDA DE ENERGIA  — gráfico único e enorme, uma linha por cidade
# ═══════════════════════════════════════════════════════════════════════
def page_demanda_energia(df: pd.DataFrame):
    h2("⚡ Demanda de Energia",
       "Série completa 1975 → 2035 · histórico real + projeção · uma curva por cidade")

    # ── Controles compactos no topo ─────────────────────────────────
    ctrl1, ctrl2, ctrl3 = st.columns([2, 2, 2])
    with ctrl1:
        cidades_opcoes = ["Todas"] + ALL_LOCS
        cidade_filtro = st.selectbox(
            "🏙️ Filtrar cidade / Utópia",
            options=cidades_opcoes, index=0, key="de_cidade",
        )
    with ctrl2:
        cenario_sel = st.selectbox(
            "📊 Cenário projeção",
            options=CEN_NAMES, format_func=lambda c: CEN_LBLS[c],
            key="de_cenario",
        )
    with ctrl3:
        mostrar_banda = st.toggle("📉 Banda Alto/Baixo", value=False, key="de_banda")

    # Cores por localidade (6 cidades + Utópia)
    LOC_COLORS = [
        "#0ea5e9",  # Catarina    — sky
        "#f59e0b",  # Nicodemus   — amber
        "#22c55e",  # Saragoça    — green
        "#ef4444",  # Santa Bárbara — red
        "#8b5cf6",  # Rural       — violet
        "#ec4899",  # Utopia      — pink
    ]
    loc_color = {loc: clr for loc, clr in zip(ALL_LOCS, LOC_COLORS)}

    # Localidades a plotar
    locs_plot = ALL_LOCS if cidade_filtro == "Todas" else [cidade_filtro]

    fig = go.Figure()

    for loc in locs_plot:
        clr = loc_color[loc]
        clr_lite = clr + "22"   # ~13% alpha para banda

        # ── Histórico ─────────────────────────────────────────────
        df_hist_loc = None
        try:
            df_hist_loc = load_hist(loc, EXCEL_PATH)
        except Exception:
            pass

        ref_loc  = df[(df["Cenario"]=="Referencia") & (df["Local"]==loc)].sort_values("Ano")
        alto_loc = df[(df["Cenario"]=="Alto")        & (df["Local"]==loc)].sort_values("Ano")
        baix_loc = df[(df["Cenario"]=="Baixo")       & (df["Local"]==loc)].sort_values("Ano")
        anos_f   = ref_loc["Ano"].tolist()
        ano_anc  = anos_f[0]

        if df_hist_loc is not None and HIST_TOTAL in df_hist_loc.columns:
            anos_h = df_hist_loc["ANO"].tolist()
            vals_h = (df_hist_loc[HIST_TOTAL] / 1e6).tolist()
            # Ponto âncora (conecta)
            anos_h_ext = anos_h + [ano_anc]
            vals_h_ext = vals_h + [ref_loc["EE_TOTAL"].iloc[0] / 1e6]
            fig.add_trace(go.Scatter(
                x=anos_h_ext, y=vals_h_ext,
                mode="lines", name=f"{loc} (hist.)",
                line=dict(color=clr, width=1.6, dash="dot"),
                legendgroup=loc,
                hovertemplate=f"<b>{loc}</b> %{{x}}: %{{y:,.3f}} ×10⁶ MWh<extra>Histórico</extra>",
            ))
            v_anc = df_hist_loc[HIST_TOTAL].iloc[-1] / 1e6
            anos_proj = [df_hist_loc["ANO"].iloc[-1]] + anos_f
            vals_ref  = [v_anc] + (ref_loc["EE_TOTAL"] / 1e6).tolist()
            vals_alto = [v_anc] + (alto_loc["EE_TOTAL"] / 1e6).tolist()
            vals_baix = [v_anc] + (baix_loc["EE_TOTAL"] / 1e6).tolist()
        else:
            anos_proj = anos_f
            vals_ref  = (ref_loc ["EE_TOTAL"] / 1e6).tolist()
            vals_alto = (alto_loc["EE_TOTAL"] / 1e6).tolist()
            vals_baix = (baix_loc["EE_TOTAL"] / 1e6).tolist()

        # Banda Alto/Baixo
        if mostrar_banda:
            x_b = anos_proj + anos_proj[::-1]
            y_b = vals_alto + vals_baix[::-1]
            fig.add_trace(go.Scatter(
                x=x_b, y=y_b, fill="toself",
                fillcolor=clr + "20",
                line=dict(color="rgba(0,0,0,0)"),
                name=f"{loc} ±3%", legendgroup=loc,
                showlegend=False, hoverinfo="skip",
            ))

        # Projeção — linha sólida mais grossa
        fig.add_trace(go.Scatter(
            x=anos_proj, y=vals_ref,
            mode="lines+markers", name=f"{loc}",
            line=dict(color=clr, width=2.8),
            marker=dict(size=5, color=clr),
            legendgroup=loc,
            hovertemplate=f"<b>{loc}</b> %{{x}}: %{{y:,.3f}} ×10⁶ MWh<extra>Projeção {CEN_LBLS[cenario_sel]}</extra>",
        ))

    # Linha vertical de corte
    ref_anos = df[(df["Cenario"]=="Referencia") & (df["Local"]=="Utopia")]["Ano"]
    if not ref_anos.empty:
        fig.add_vline(
            x=ref_anos.min() - 0.5,
            line=dict(color="#94a3b8", width=1.2, dash="dot"),
            annotation_text="início da projeção",
            annotation_font_size=10, annotation_font_color=TEXT_SEC,
        )

    fig.update_layout(
        title=dict(
            text="Demanda de Energia Elétrica Total — Utópia",
            font=dict(size=16, color=TEXT_PRI, weight="bold"), x=0,
        ),
        yaxis_title="EE Total (×10⁶ MWh)",
        xaxis_title="Ano",
        height=580,
        legend=dict(
            orientation="v", yanchor="top", y=1,
            xanchor="left", x=1.01, font=dict(size=11),
            bordercolor="#e2e8f0", borderwidth=1,
        ),
        margin=dict(l=15, r=160, t=60, b=15),
        **PLOTLY_THEME,
    )
    st.plotly_chart(fig, use_container_width=True, key="de_main")


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: RANKING DE CRESCIMENTO
# ═══════════════════════════════════════════════════════════════════════
def page_ranking(df: pd.DataFrame):
    h2("🔥 Ranking de Crescimento 2026 → 2035",
       "Quais cidades e setores crescem mais? Δ% entre o primeiro e o último ano de projeção")

    st.sidebar.markdown("---")
    sidebar_sep("Ranking")
    cenario_sel = st.sidebar.radio(
        "📊 Cenário", CEN_NAMES,
        format_func=lambda c: CEN_LBLS[c], key="rk_cenario",
    )
    modo = st.sidebar.radio(
        "📐 Modo", ["Por Setor (Utópia)", "Por Cidade (Total)"],
        key="rk_modo",
    )

    sub = df[df["Cenario"] == cenario_sel]
    ano_ini = sub["Ano"].min()
    ano_fim = sub["Ano"].max()

    if modo == "Por Setor (Utópia)":
        ut = sub[sub["Local"] == "Utopia"]
        r0 = ut[ut["Ano"] == ano_ini].iloc[0]
        r1 = ut[ut["Ano"] == ano_fim].iloc[0]

        items = []
        for sec, slbl in zip(SECTORS, SECTOR_LBLS):
            v0, v1 = r0[sec], r1[sec]
            if v0 > 0:
                items.append({
                    "label": slbl,
                    "delta_pct": (v1 - v0) / v0 * 100,
                    "delta_abs": (v1 - v0) / 1e6,
                    "v2026": v0 / 1e6,
                    "v2035": v1 / 1e6,
                })
        items.sort(key=lambda x: x["delta_pct"], reverse=True)
        labels = [i["label"] for i in items]
        deltas = [i["delta_pct"] for i in items]
        cores  = ["#22c55e" if d >= 0 else "#ef4444" for d in deltas]
        customdata = [[i["v2026"], i["v2035"], i["delta_abs"]] for i in items]
        title_txt = f"Crescimento por Setor — Utópia · {CEN_LBLS[cenario_sel]}"

    else:  # Por Cidade
        items = []
        for loc in CIDADES:
            s = sub[sub["Local"] == loc]
            r0 = s[s["Ano"] == ano_ini]
            r1 = s[s["Ano"] == ano_fim]
            if r0.empty or r1.empty: continue
            v0 = r0["EE_TOTAL"].values[0]
            v1 = r1["EE_TOTAL"].values[0]
            if v0 > 0:
                items.append({
                    "label": loc,
                    "delta_pct": (v1 - v0) / v0 * 100,
                    "delta_abs": (v1 - v0) / 1e6,
                    "v2026": v0 / 1e6,
                    "v2035": v1 / 1e6,
                })
        items.sort(key=lambda x: x["delta_pct"], reverse=True)
        labels = [i["label"] for i in items]
        deltas = [i["delta_pct"] for i in items]
        cores  = ["#22c55e" if d >= 0 else "#ef4444" for d in deltas]
        customdata = [[i["v2026"], i["v2035"], i["delta_abs"]] for i in items]
        title_txt = f"Crescimento por Cidade (EE Total) · {CEN_LBLS[cenario_sel]}"

    fig = go.Figure(go.Bar(
        x=deltas, y=labels,
        orientation="h",
        marker_color=cores,
        marker_line_width=0,
        customdata=customdata,
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Δ%: <b>%{x:+.1f}%</b><br>"
            f"{ano_ini}: %{{customdata[0]:,.2f}} ×10⁶ MWh<br>"
            f"{ano_fim}: %{{customdata[1]:,.2f}} ×10⁶ MWh<br>"
            "Δ abs: %{customdata[2]:+.2f} ×10⁶ MWh"
            "<extra></extra>"
        ),
    ))
    # Linha de zero
    fig.add_vline(x=0, line=dict(color="#94a3b8", width=1))
    # Anotações de valor nas barras
    for i, (lbl, d) in enumerate(zip(labels, deltas)):
        fig.add_annotation(
            x=d, y=i,
            text=f" {d:+.1f}%",
            xanchor="left" if d >= 0 else "right",
            showarrow=False,
            font=dict(size=12, color="#0f172a", weight="bold"),
        )

    fig.update_layout(
        title=dict(text=title_txt,
                   font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0),
        xaxis_title="Variação % (2026 → 2035)",
        height=max(320, len(labels) * 60 + 80),
        margin=dict(l=15, r=80, t=50, b=15),
        **PLOTLY_THEME,
    )
    st.plotly_chart(fig, use_container_width=True, key="rk_main")

    # KPI cards de resumo
    if items:
        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        maior = items[0]
        menor = items[-1]
        medio = sum(i["delta_pct"] for i in items) / len(items)
        css = CSS_CARD
        with col1:
            st.markdown(
                f'<div style="{css}"><div style="font-size:10px;font-weight:700;'
                f'color:{TEXT_SEC};text-transform:uppercase;">🏆 Maior crescimento</div>'
                f'<div style="font-size:20px;font-weight:800;color:#22c55e;">'
                f'{maior["label"]}</div>'
                f'<div style="font-size:14px;color:#22c55e;font-weight:600;">'
                f'+{maior["delta_pct"]:.1f}%</div></div>',
                unsafe_allow_html=True,
            )
        with col2:
            st.markdown(
                f'<div style="{css}"><div style="font-size:10px;font-weight:700;'
                f'color:{TEXT_SEC};text-transform:uppercase;">📊 Crescimento médio</div>'
                f'<div style="font-size:20px;font-weight:800;color:{ACCENT};">'
                f'{medio:+.1f}%</div>'
                f'<div style="font-size:12px;color:{TEXT_SEC};">média do grupo</div></div>',
                unsafe_allow_html=True,
            )
        with col3:
            cor_m = "#ef4444" if menor["delta_pct"] < 0 else "#f59e0b"
            st.markdown(
                f'<div style="{css}"><div style="font-size:10px;font-weight:700;'
                f'color:{TEXT_SEC};text-transform:uppercase;">📉 Menor crescimento</div>'
                f'<div style="font-size:20px;font-weight:800;color:{cor_m};">'
                f'{menor["label"]}</div>'
                f'<div style="font-size:14px;color:{cor_m};font-weight:600;">'
                f'{menor["delta_pct"]:+.1f}%</div></div>',
                unsafe_allow_html=True,
            )


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: DECOMPOSIÇÃO DO CRESCIMENTO (waterfall)
# ═══════════════════════════════════════════════════════════════════════
def page_waterfall(df: pd.DataFrame):
    h2("📉 Decomposição do Crescimento",
       "Quanto cada setor contribui (em MWh) para o crescimento total de Utópia")

    st.sidebar.markdown("---")
    sidebar_sep("Waterfall")
    cenario_sel = st.sidebar.radio(
        "📊 Cenário", CEN_NAMES,
        format_func=lambda c: CEN_LBLS[c], key="wf_cenario",
    )
    local_sel = st.sidebar.selectbox(
        "📍 Localidade", ALL_LOCS, index=len(ALL_LOCS)-1, key="wf_local",
    )

    sub = df[(df["Cenario"] == cenario_sel) & (df["Local"] == local_sel)]
    ano_ini = sub["Ano"].min()
    ano_fim = sub["Ano"].max()
    r0 = sub[sub["Ano"] == ano_ini].iloc[0]
    r1 = sub[sub["Ano"] == ano_fim].iloc[0]

    total_ini = r0["EE_TOTAL"] / 1e6
    total_fim = r1["EE_TOTAL"] / 1e6

    # Contribuições por setor
    contribs = []
    for sec, slbl in zip(SECTORS, SECTOR_LBLS):
        delta = (r1[sec] - r0[sec]) / 1e6
        contribs.append((slbl, delta))

    # Monta waterfall: ponto inicial + setores + ponto final
    measure = ["absolute"] + ["relative"] * len(contribs) + ["total"]
    x_labels = [f"{ano_ini}"] + [c[0] for c in contribs] + [f"{ano_fim}"]
    y_values = [total_ini]    + [c[1] for c in contribs] + [total_fim]

    # Cores: verde = positivo, vermelho = negativo, cinza = totais
    marker_colors = ["#64748b"]  # ponto inicial
    for _, d in contribs:
        marker_colors.append("#22c55e" if d >= 0 else "#ef4444")
    marker_colors.append(ACCENT)  # ponto final

    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=measure,
        x=x_labels,
        y=y_values,
        connector=dict(line=dict(color="#cbd5e1", width=1.5, dash="dot")),
        decreasing=dict(marker_color="#ef4444"),
        increasing=dict(marker_color="#22c55e"),
        totals=dict(marker_color=ACCENT),
        text=[f"{v:+.2f}" if m == "relative" else f"{v:.2f}"
              for v, m in zip(y_values, measure)],
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>%{y:+.3f} ×10⁶ MWh<extra></extra>",
    ))

    fig.update_layout(
        title=dict(
            text=f"Decomposição do Crescimento — {local_sel}  ·  {CEN_LBLS[cenario_sel]}  ·  {ano_ini} → {ano_fim}",
            font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0,
        ),
        yaxis_title="EE (×10⁶ MWh)",
        height=480,
        showlegend=False,
        margin=dict(l=15, r=15, t=60, b=15),
        **PLOTLY_THEME,
    )
    st.plotly_chart(fig, use_container_width=True, key="wf_main")

    # Mini tabela de contribuições absolutas e relativas
    st.markdown("##### Contribuições detalhadas")
    rows = []
    for sec, slbl in zip(SECTORS, SECTOR_LBLS):
        v0 = r0[sec] / 1e6
        v1 = r1[sec] / 1e6
        d  = v1 - v0
        pct_total = d / (total_fim - total_ini) * 100 if (total_fim - total_ini) else 0
        rows.append({
            "Setor": slbl,
            f"{ano_ini} (×10⁶ MWh)": f"{v0:,.3f}",
            f"{ano_fim} (×10⁶ MWh)": f"{v1:,.3f}",
            "Δ (×10⁶ MWh)": f"{d:+.3f}",
            "% do crescimento total": f"{pct_total:+.1f}%",
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: INTENSIDADE ENERGÉTICA
# ═══════════════════════════════════════════════════════════════════════
def page_intensidade(df: pd.DataFrame):
    h2("🧮 Intensidade Energética",
       "EE/PIB e EE/hab ao longo do tempo — histórico + projeção · eficiência energética")

    st.sidebar.markdown("---")
    sidebar_sep("Intensidade")
    local_sel = st.sidebar.selectbox(
        "📍 Localidade", ALL_LOCS, index=len(ALL_LOCS)-1, key="int_local",
    )
    cenario_sel = st.sidebar.radio(
        "📊 Cenário", CEN_NAMES,
        format_func=lambda c: CEN_LBLS[c], key="int_cenario",
    )

    # ── Dados projeção ──────────────────────────────────────────────
    ref = df[(df["Cenario"] == cenario_sel) & (df["Local"] == local_sel)].sort_values("Ano")
    anos_f  = ref["Ano"].tolist()
    ano_anc = anos_f[0]

    # Intensidade projetada: MWh / Utd$ (×1000) e MWh/hab
    ee_pib_f = (ref["EE_TOTAL"] / ref["PIB"] * 1e3).tolist()   # MWh por 1000 Utd$
    ee_hab_f = (ref["EE_TOTAL"] / ref["POP"]).tolist()          # MWh por habitante

    # ── Dados históricos ────────────────────────────────────────────
    df_hist = None
    anos_h, ee_pib_h, ee_hab_h = [], [], []
    try:
        df_hist = load_hist(local_sel, EXCEL_PATH)
        anos_h   = df_hist["ANO"].tolist()
        # Excel histórico: coluna EE (total), PIB, POPULAÇÃO
        if all(c in df_hist.columns for c in [HIST_TOTAL, "PIB", "POPULAÇÃO"]):
            ee_pib_h = (df_hist[HIST_TOTAL] / df_hist["PIB"] * 1e3).tolist()
            ee_hab_h = (df_hist[HIST_TOTAL] / df_hist["POPULAÇÃO"]).tolist()
    except Exception:
        pass

    # ── Figura com dois subplots ────────────────────────────────────
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=[
            "Intensidade Energética / PIB  (MWh por 1000 Utd$)",
            "Intensidade Energética / Hab  (MWh por habitante)",
        ],
        horizontal_spacing=0.09,
    )

    # Ponto âncora: conecta histórico → projeção
    def build_series(anos_h, vals_h, anos_f, vals_f, ano_anc):
        """Retorna (anos_hist_ext, vals_hist_ext, anos_proj_ext, vals_proj_ext)."""
        if anos_h and vals_h:
            # âncora = primeiro ponto de projeção
            a_h = anos_h + [ano_anc]
            v_h = vals_h + [vals_f[0]]
            a_p = [anos_h[-1]] + anos_f
            v_p = [vals_h[-1]] + vals_f
        else:
            a_h, v_h = [], []
            a_p, v_p = anos_f, vals_f
        return a_h, v_h, a_p, v_p

    for col_idx, (vals_h, vals_f, ytitle, key_sfx) in enumerate([
        (ee_pib_h, ee_pib_f, "MWh / 1 000 Utd$", "pib"),
        (ee_hab_h, ee_hab_f, "MWh / habitante",   "hab"),
    ], start=1):
        a_h, v_h, a_p, v_p = build_series(anos_h, vals_h, anos_f, vals_f, ano_anc)

        if a_h:
            fig.add_trace(go.Scatter(
                x=a_h, y=v_h, mode="lines",
                name="Histórico",
                line=dict(color="#94a3b8", width=1.8),
                legendgroup="hist_int",
                showlegend=(col_idx == 1),
                hovertemplate=f"Hist. %{{x}}: %{{y:,.3f}} {ytitle}<extra></extra>",
            ), row=1, col=col_idx)

        fig.add_trace(go.Scatter(
            x=a_p, y=v_p, mode="lines+markers",
            name=f"Projeção — {CEN_LBLS[cenario_sel]}",
            line=dict(color=ACCENT, width=2.5),
            marker=dict(size=5),
            legendgroup="proj_int",
            showlegend=(col_idx == 1),
            hovertemplate=f"Proj. %{{x}}: %{{y:,.3f}} {ytitle}<extra></extra>",
        ), row=1, col=col_idx)

        # Linha de tendência (linear)
        if a_p:
            x_arr = np.array(a_p)
            y_arr = np.array(v_p)
            p = np.polyfit(x_arr, y_arr, 1)
            fig.add_trace(go.Scatter(
                x=a_p, y=np.polyval(p, x_arr).tolist(),
                mode="lines", name="Tendência linear",
                line=dict(color="#f59e0b", width=1.2, dash="dash"),
                legendgroup="trend_int",
                showlegend=(col_idx == 1),
                hoverinfo="skip",
            ), row=1, col=col_idx)

        # Linha de corte
        fig.add_vline(x=ano_anc - 0.5,
                      line=dict(color="#cbd5e1", width=1, dash="dot"),
                      row=1, col=col_idx)

    fig.update_layout(
        title=dict(
            text=f"Intensidade Energética — {local_sel}  ·  {CEN_LBLS[cenario_sel]}",
            font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0,
        ),
        height=420,
        legend=dict(orientation="h", yanchor="bottom", y=-0.20,
                    xanchor="center", x=0.5, font=dict(size=10)),
        margin=dict(l=15, r=15, t=60, b=15),
        **PLOTLY_THEME,
    )
    fig.update_yaxes(title_text="MWh / 1 000 Utd$", row=1, col=1)
    fig.update_yaxes(title_text="MWh / habitante",   row=1, col=2)
    st.plotly_chart(fig, use_container_width=True, key="int_main")

    # ── Insight textual ─────────────────────────────────────────────
    if ee_pib_f:
        tend_pib = "↓ caindo" if ee_pib_f[-1] < ee_pib_f[0] else "↑ subindo"
        tend_hab = "↓ caindo" if ee_hab_f[-1] < ee_hab_f[0] else "↑ subindo"
        d_pib = (ee_pib_f[-1] - ee_pib_f[0]) / ee_pib_f[0] * 100
        d_hab = (ee_hab_f[-1] - ee_hab_f[0]) / ee_hab_f[0] * 100
        cor_pib = "#22c55e" if d_pib < 0 else "#ef4444"
        cor_hab = "#22c55e" if d_hab < 0 else "#ef4444"
        st.markdown(
            f'<div style="{CSS_CARD}margin-top:12px;">'
            f'<b>Leitura da intensidade ({CEN_LBLS[cenario_sel]}):</b><br>'
            f'• EE/PIB {tend_pib} '
            f'<span style="color:{cor_pib};font-weight:700;">{d_pib:+.1f}%</span> '
            f'no período — economia {"mais" if d_pib < 0 else "menos"} eficiente por unidade de PIB.<br>'
            f'• EE/hab {tend_hab} '
            f'<span style="color:{cor_hab};font-weight:700;">{d_hab:+.1f}%</span> '
            f'— consumo per capita {"cai" if d_hab < 0 else "sobe"}.'
            f'</div>',
            unsafe_allow_html=True,
        )


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: MAPA DE CALOR  Ano × Cidade
# ═══════════════════════════════════════════════════════════════════════
def page_heatmap(df: pd.DataFrame):
    h2("🗺️ Mapa de Calor — Ano × Cidade",
       "Cor = EE Total ou variação % · visão compacta de toda a projeção")

    st.sidebar.markdown("---")
    sidebar_sep("Heatmap")
    cenario_sel = st.sidebar.radio(
        "📊 Cenário", CEN_NAMES,
        format_func=lambda c: CEN_LBLS[c], key="hm_cenario",
    )
    metrica = st.sidebar.radio(
        "📐 Métrica", ["EE Total (×10⁶ MWh)", "Variação % acum. vs 2026"],
        key="hm_metrica",
    )

    sub = df[(df["Cenario"] == cenario_sel) & (df["Local"].isin(CIDADES))]
    anos = sorted(sub["Ano"].unique().tolist())
    # Pivô: linhas = cidades, colunas = anos
    pivot = sub.pivot_table(index="Local", columns="Ano", values="EE_TOTAL", aggfunc="sum")
    # Garante a ordem de cidades
    pivot = pivot.reindex(CIDADES)

    if metrica == "Variação % acum. vs 2026":
        base = pivot[anos[0]]
        z = ((pivot.subtract(base, axis=0)).divide(base, axis=0) * 100).values
        colorscale = "RdYlGn"
        zmid = 0
        colorbar_title = "Δ% vs 2026"
        text_fmt = ".1f"
    else:
        z = (pivot / 1e6).values
        colorscale = [
            [0.0, "#eff6ff"], [0.25, "#bae6fd"],
            [0.5, "#0ea5e9"], [0.75, "#0369a1"], [1.0, "#0c4a6e"],
        ]
        zmid = None
        colorbar_title = "×10⁶ MWh"
        text_fmt = ".2f"

    # Texto nas células
    text_matrix = [[f"{v:{text_fmt}}" for v in row] for row in z]

    fig = go.Figure(go.Heatmap(
        z=z,
        x=[str(a) for a in anos],
        y=CIDADES,
        text=text_matrix,
        texttemplate="%{text}",
        textfont=dict(size=11, color="#0f172a"),
        colorscale=colorscale,
        zmid=zmid,
        colorbar=dict(title=colorbar_title, thickness=14, len=0.8),
        hovertemplate="<b>%{y}</b>  %{x}<br>%{z:.3f}<extra></extra>",
    ))

    fig.update_layout(
        title=dict(
            text=f"Heatmap — {metrica}  ·  {CEN_LBLS[cenario_sel]}",
            font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0,
        ),
        height=380,
        margin=dict(l=15, r=30, t=60, b=15),
        **PLOTLY_THEME,
    )
    fig.update_xaxes(title_text="Ano", tickfont=dict(size=11))
    fig.update_yaxes(title_text="Cidade", tickfont=dict(size=11))
    st.plotly_chart(fig, use_container_width=True, key="hm_main")

    # Segundo heatmap: por setor × ano (Utópia)
    st.markdown("---")
    st.markdown("#### Heatmap Setor × Ano — Utópia")
    ut = df[(df["Cenario"] == cenario_sel) & (df["Local"] == "Utopia")].sort_values("Ano")
    if not ut.empty:
        z2 = np.array([(ut[s] / 1e6).tolist() for s in SECTORS])
        if metrica == "Variação % acum. vs 2026":
            base2 = z2[:, 0:1]
            z2 = (z2 - base2) / base2 * 100
            cs2 = "RdYlGn"
            zm2 = 0
            fmt2 = ".1f"
            cb2  = "Δ% vs 2026"
        else:
            cs2 = [[0,"#f0fdf4"],[0.5,"#22c55e"],[1,"#14532d"]]
            zm2 = None
            fmt2 = ".2f"
            cb2  = "×10⁶ MWh"

        text2 = [[f"{v:{fmt2}}" for v in row] for row in z2]
        fig2 = go.Figure(go.Heatmap(
            z=z2, x=[str(a) for a in anos], y=SECTOR_LBLS,
            text=text2, texttemplate="%{text}",
            textfont=dict(size=11), colorscale=cs2, zmid=zm2,
            colorbar=dict(title=cb2, thickness=14, len=0.8),
            hovertemplate="<b>%{y}</b>  %{x}<br>%{z:.3f}<extra></extra>",
        ))
        fig2.update_layout(
            height=300,
            margin=dict(l=15, r=30, t=30, b=15),
            **PLOTLY_THEME,
        )
        fig2.update_xaxes(title_text="Ano")
        fig2.update_yaxes(title_text="Setor")
        st.plotly_chart(fig2, use_container_width=True, key="hm_setor")


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: TABELA
# ═══════════════════════════════════════════════════════════════════════
def page_tabela(df: pd.DataFrame):
    h2("📋 Tabela de Projeções")

    st.sidebar.markdown("---")
    sidebar_sep("Tabela")
    local_sel   = st.sidebar.selectbox(
        "📍 Localidade", ALL_LOCS, index=len(ALL_LOCS)-1, key="tab_local"
    )
    cenario_sel = st.sidebar.radio(
        "📊 Cenário", CEN_NAMES,
        format_func=lambda c: CEN_LBLS[c], key="tab_cenario",
    )

    sub = df[(df["Cenario"]==cenario_sel) & (df["Local"]==local_sel)].sort_values("Ano")
    if sub.empty:
        st.warning("Sem dados.")
        return
    cols_show = ["Ano","PIB","EE_RES","EE_IND","EE_SER","EE_MIN","EE_AGR","EE_UTIL","EE_TOTAL","POP"]
    labels    = {"Ano":"Ano","PIB":"PIB (Utd$)","EE_RES":"Res","EE_IND":"Ind",
                 "EE_SER":"Ser","EE_MIN":"Min","EE_AGR":"Agr","EE_UTIL":"Util",
                 "EE_TOTAL":"TOTAL (MWh)","POP":"Pop."}
    disp = sub[cols_show].copy().rename(columns=labels)
    for c in disp.columns:
        if c == "Ano": disp[c] = disp[c].astype(int)
        else: disp[c] = disp[c].apply(lambda v: f"{v:,.0f}")
    st.dataframe(disp, use_container_width=True, hide_index=True)
    csv_bytes = sub.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Baixar CSV", data=csv_bytes,
                       file_name=f"projecao_{local_sel}_{cenario_sel}.csv",
                       mime="text/csv")


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: POPULAÇÃO POR ANO
# ═══════════════════════════════════════════════════════════════════════
def page_populacao(df: pd.DataFrame):
    """
    Série temporal de população (histórico + projeção) por localidade.
    Sidebar: localidade, cenário, toggle histórico, toggle todas as cidades.
    Inclui KPI cards (pop. inicial / final / CAGR) e subplots por cidade.
    """
    h2("👥 População por Ano",
       "Evolução demográfica 1975 → 2035 · histórico real + projeção por localidade")

    # ── Controles sidebar ───────────────────────────────────────────
    st.sidebar.markdown("---")
    sidebar_sep("População")

    modo = st.sidebar.radio(
        "🗺️ Visão", ["Uma localidade", "Todas as cidades"],
        key="pop_modo",
    )
    cenario_sel = st.sidebar.radio(
        "📊 Cenário", CEN_NAMES,
        format_func=lambda c: CEN_LBLS[c], key="pop_cenario",
    )
    mostrar_hist = st.sidebar.toggle(
        "📜 Incluir Histórico", value=True, key="pop_hist",
        help="Sobrepõe a série histórica (1975–2025) do Excel",
    )
    mostrar_cenarios = st.sidebar.toggle(
        "📉 Mostrar Alto/Baixo", value=False, key="pop_cen_toggle",
    )

    # Cores por localidade (igual à page_demanda_energia)
    LOC_COLORS = {
        "Catarina":      "#0ea5e9",
        "Nicodemus":     "#f59e0b",
        "Saragoça":      "#22c55e",
        "Santa Bárbara": "#ef4444",
        "Rural":         "#8b5cf6",
        "Utopia":        "#ec4899",
    }

    # ── MODO: uma localidade ─────────────────────────────────────────
    if modo == "Uma localidade":
        local_sel = st.sidebar.selectbox(
            "📍 Localidade", ALL_LOCS, index=len(ALL_LOCS) - 1, key="pop_local",
        )

        ref  = df[(df["Cenario"] == "Referencia") & (df["Local"] == local_sel)].sort_values("Ano")
        alto = df[(df["Cenario"] == "Alto")        & (df["Local"] == local_sel)].sort_values("Ano")
        baix = df[(df["Cenario"] == "Baixo")       & (df["Local"] == local_sel)].sort_values("Ano")

        if ref.empty:
            st.warning("Sem dados de projeção para esta localidade.")
            return

        anos_f  = ref["Ano"].tolist()
        ano_anc = anos_f[0]
        pop_ref  = ref["POP"].tolist()
        pop_alto = alto["POP"].tolist() if not alto.empty else pop_ref
        pop_baix = baix["POP"].tolist() if not baix.empty else pop_ref

        # ── Histórico demográfico ────────────────────────────────────
        anos_h, pop_h = [], []
        try:
            df_hist = load_hist(local_sel, EXCEL_PATH)
            if "POPULAÇÃO" in df_hist.columns:
                anos_h = df_hist["ANO"].tolist()
                pop_h  = df_hist["POPULAÇÃO"].tolist()
        except Exception:
            pass

        clr = LOC_COLORS.get(local_sel, ACCENT)
        fig = go.Figure()

        # Linha divisória histórico / projeção
        fig.add_vline(
            x=ano_anc - 0.5,
            line=dict(color="#94a3b8", width=1, dash="dot"),
            annotation_text="↑ início da projeção",
            annotation_font_size=10,
            annotation_font_color=TEXT_SEC,
        )

        # Histórico
        if mostrar_hist and anos_h:
            anos_h_ext = anos_h + [ano_anc]
            pop_h_ext  = pop_h  + [pop_ref[0]]
            fig.add_trace(go.Scatter(
                x=anos_h_ext, y=pop_h_ext,
                mode="lines",
                name="Histórico",
                line=dict(color="#475569", width=2, dash="solid"),
                hovertemplate="Hist. %{x}: %{y:,.0f} hab<extra></extra>",
            ))

        # Banda Alto / Baixo
        if mostrar_cenarios:
            if anos_h and mostrar_hist:
                anc_v  = pop_h[-1]
                anos_b = [anos_h[-1]] + anos_f
                v_alto = [anc_v] + pop_alto
                v_baix = [anc_v] + pop_baix
            else:
                anos_b = anos_f
                v_alto = pop_alto
                v_baix = pop_baix
            x_band = anos_b + anos_b[::-1]
            fig.add_trace(go.Scatter(
                x=x_band, y=v_alto + v_baix[::-1],
                fill="toself",
                fillcolor=clr.replace("#", "rgba(") + ",0.08)" if clr.startswith("#") else "rgba(14,165,233,0.08)",
                line=dict(color="rgba(0,0,0,0)"),
                name="Banda ±PIB 3%", hoverinfo="skip",
            ))
            fig.add_trace(go.Scatter(
                x=anos_b, y=v_alto, mode="lines",
                name="Alto (PIB +3%)",
                line=dict(color=clr, width=1.3, dash="dash"),
                hovertemplate="Alto %{x}: %{y:,.0f} hab<extra></extra>",
            ))
            fig.add_trace(go.Scatter(
                x=anos_b, y=v_baix, mode="lines",
                name="Baixo (PIB -3%)",
                line=dict(color=clr, width=1.3, dash="dot"),
                hovertemplate="Baixo %{x}: %{y:,.0f} hab<extra></extra>",
            ))

        # Projeção referência
        if anos_h and mostrar_hist:
            anc_v   = pop_h[-1]
            anos_r  = [anos_h[-1]] + anos_f
            pop_r   = [anc_v] + pop_ref
        else:
            anos_r = anos_f
            pop_r  = pop_ref

        fig.add_trace(go.Scatter(
            x=anos_r, y=pop_r,
            mode="lines+markers",
            name=f"Projeção — {CEN_LBLS[cenario_sel]}",
            line=dict(color=clr, width=2.8),
            marker=dict(size=5, color=clr),
            hovertemplate="Proj. %{x}: %{y:,.0f} hab<extra></extra>",
        ))

        fig.update_layout(
            title=dict(
                text=f"População — {local_sel}  ·  {CEN_LBLS[cenario_sel]}",
                font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0,
            ),
            yaxis_title="Habitantes",
            xaxis_title="Ano",
            height=420,
            legend=dict(orientation="h", yanchor="bottom", y=1.02,
                        xanchor="right", x=1, font=dict(size=10)),
            margin=dict(l=15, r=15, t=60, b=15),
            **PLOTLY_THEME,
        )
        st.plotly_chart(fig, use_container_width=True, key="pop_main")

        # ── KPI cards ───────────────────────────────────────────────
        st.markdown("---")
        pop_ini = pop_ref[0]
        pop_fin = pop_ref[-1]
        n_anos  = anos_f[-1] - anos_f[0]
        cagr    = ((pop_fin / pop_ini) ** (1 / n_anos) - 1) * 100 if pop_ini and n_anos else 0
        delta   = pop_fin - pop_ini

        col1, col2, col3, col4 = st.columns(4)
        kpi_data = [
            (col1, f"Pop. {anos_f[0]}",  f"{pop_ini:,.0f}", "hab", ACCENT),
            (col2, f"Pop. {anos_f[-1]}", f"{pop_fin:,.0f}", "hab", ACCENT),
            (col3, "Crescimento Total",  f"{delta:+,.0f}",  "hab",
             "#22c55e" if delta >= 0 else "#ef4444"),
            (col4, "CAGR Demográfico",  f"{cagr:+.2f}",    "% a.a.",
             "#22c55e" if cagr >= 0 else "#ef4444"),
        ]
        for col, lbl, val, unit, cor in kpi_data:
            with col:
                st.markdown(
                    f'<div style="{CSS_CARD}">'
                    f'<div style="font-size:10px;font-weight:700;color:{TEXT_SEC};'
                    f'text-transform:uppercase;">{lbl}</div>'
                    f'<div style="font-size:24px;font-weight:800;color:{cor};'
                    f'letter-spacing:-.5px;">{val}</div>'
                    f'<div style="font-size:11px;color:{TEXT_SEC};">{unit}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    # ── MODO: todas as cidades ───────────────────────────────────────
    else:
        fig = go.Figure()
        for loc in ALL_LOCS:
            clr = LOC_COLORS.get(loc, ACCENT)
            ref_loc = df[(df["Cenario"] == cenario_sel) & (df["Local"] == loc)].sort_values("Ano")
            if ref_loc.empty:
                continue
            anos_f  = ref_loc["Ano"].tolist()
            ano_anc = anos_f[0]
            pop_ref = ref_loc["POP"].tolist()

            # Histórico
            anos_h, pop_h = [], []
            try:
                df_hist = load_hist(loc, EXCEL_PATH)
                if "POPULAÇÃO" in df_hist.columns:
                    anos_h = df_hist["ANO"].tolist()
                    pop_h  = df_hist["POPULAÇÃO"].tolist()
            except Exception:
                pass

            # Divisória (só na primeira localidade para não poluir)
            if loc == ALL_LOCS[0]:
                fig.add_vline(
                    x=ano_anc - 0.5,
                    line=dict(color="#94a3b8", width=1, dash="dot"),
                    annotation_text="↑ projeção",
                    annotation_font_size=9, annotation_font_color=TEXT_SEC,
                )

            if mostrar_hist and anos_h:
                anos_h_ext = anos_h + [ano_anc]
                pop_h_ext  = pop_h  + [pop_ref[0]]
                fig.add_trace(go.Scatter(
                    x=anos_h_ext, y=pop_h_ext,
                    mode="lines", name=f"{loc} hist.",
                    line=dict(color=clr, width=1.4, dash="dot"),
                    legendgroup=loc,
                    hovertemplate=f"<b>{loc}</b> hist. %{{x}}: %{{y:,.0f}} hab<extra></extra>",
                ))
                anos_r = [anos_h[-1]] + anos_f
                pop_r  = [pop_h[-1]] + pop_ref
            else:
                anos_r = anos_f
                pop_r  = pop_ref

            fig.add_trace(go.Scatter(
                x=anos_r, y=pop_r,
                mode="lines+markers", name=loc,
                line=dict(color=clr, width=2.4),
                marker=dict(size=4),
                legendgroup=loc,
                hovertemplate=f"<b>{loc}</b> %{{x}}: %{{y:,.0f}} hab<extra></extra>",
            ))

        fig.update_layout(
            title=dict(
                text=f"População por Localidade — {CEN_LBLS[cenario_sel]}",
                font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0,
            ),
            yaxis_title="Habitantes",
            xaxis_title="Ano",
            height=480,
            legend=dict(orientation="h", yanchor="bottom", y=-0.22,
                        xanchor="center", x=0.5, font=dict(size=10)),
            margin=dict(l=15, r=15, t=60, b=15),
            **PLOTLY_THEME,
        )
        st.plotly_chart(fig, use_container_width=True, key="pop_multi")

        # Subplots 2×3 por cidade (sem Utópia)
        st.markdown("---")
        st.markdown("#### Subplots por Cidade")
        fig2 = make_subplots(
            rows=2, cols=3,
            subplot_titles=CIDADES,
            vertical_spacing=0.16, horizontal_spacing=0.07,
        )
        for idx, loc in enumerate(CIDADES):
            row = idx // 3 + 1
            col = idx %  3 + 1
            clr = LOC_COLORS.get(loc, ACCENT)
            ref_loc = df[(df["Cenario"] == cenario_sel) & (df["Local"] == loc)].sort_values("Ano")
            if ref_loc.empty:
                continue
            anos_f  = ref_loc["Ano"].tolist()
            pop_ref = ref_loc["POP"].tolist()
            anos_h, pop_h = [], []
            try:
                df_hist = load_hist(loc, EXCEL_PATH)
                if "POPULAÇÃO" in df_hist.columns:
                    anos_h = df_hist["ANO"].tolist()
                    pop_h  = df_hist["POPULAÇÃO"].tolist()
            except Exception:
                pass

            if mostrar_hist and anos_h:
                anos_r = [anos_h[-1]] + anos_f
                pop_r  = [pop_h[-1]] + pop_ref
                fig2.add_trace(go.Scatter(
                    x=anos_h, y=pop_h, mode="lines",
                    name="Histórico", legendgroup="hist_pop",
                    showlegend=(idx == 0),
                    line=dict(color="#94a3b8", width=1.5),
                    hovertemplate=f"Hist. %{{x}}: %{{y:,.0f}}<extra></extra>",
                ), row=row, col=col)
            else:
                anos_r = anos_f
                pop_r  = pop_ref

            fig2.add_trace(go.Scatter(
                x=anos_r, y=pop_r, mode="lines+markers",
                name="Projeção", legendgroup="proj_pop",
                showlegend=(idx == 0),
                line=dict(color=clr, width=2.2), marker=dict(size=4),
                hovertemplate=f"%{{x}}: %{{y:,.0f}} hab<extra></extra>",
            ), row=row, col=col)

        fig2.update_layout(
            height=560,
            legend=dict(orientation="h", yanchor="bottom", y=-0.12,
                        xanchor="center", x=0.5, font=dict(size=10)),
            margin=dict(l=15, r=15, t=40, b=15),
            **PLOTLY_THEME,
        )
        fig2.update_yaxes(title_text="Hab.", col=1)
        st.plotly_chart(fig2, use_container_width=True, key="pop_subplots")


# ═══════════════════════════════════════════════════════════════════════
#  PAGE: VARIÁVEL Y POR ANO
# ═══════════════════════════════════════════════════════════════════════
def page_variavel_y(df: pd.DataFrame):
    """
    Gráfico de linha genérico: qualquer coluna do CSV (Y) ao longo do tempo.
    Sidebar: variável Y, localidade, cenário, toggle linha de tendência.
    Permite comparar cidades lado a lado para a mesma variável.
    """
    h2("📈 Variável Y por Ano",
       "Evolução temporal de qualquer variável do modelo · projeção 2026 → 2035")

    st.sidebar.markdown("---")
    sidebar_sep("Variável Y × Ano")

    # Mapa completo de variáveis disponíveis no CSV
    VAR_OPTS = {
        "EE Total (MWh)":           ("EE_TOTAL", "×10⁶ MWh",    1e6),
        "EE Residencial (MWh)":     ("EE_RES",   "×10⁶ MWh",    1e6),
        "EE Industrial (MWh)":      ("EE_IND",   "×10⁶ MWh",    1e6),
        "EE Serviços (MWh)":        ("EE_SER",   "×10⁶ MWh",    1e6),
        "EE Mineral (MWh)":         ("EE_MIN",   "×10⁶ MWh",    1e6),
        "EE Agropecuário (MWh)":    ("EE_AGR",   "×10⁶ MWh",    1e6),
        "EE Utilidades (MWh)":      ("EE_UTIL",  "×10⁶ MWh",    1e6),
        "PIB (Utd$)":               ("PIB",      "×10⁶ Utd$",   1e6),
        "População (hab)":          ("POP",      "habitantes",   1),
        "Área Total (km²)":         ("AR_TOTAL", "km²",          1),
    }

    var_lbl     = st.sidebar.selectbox(
        "📐 Variável Y", list(VAR_OPTS.keys()), key="vy_var",
    )
    var_col, var_unit, var_div = VAR_OPTS[var_lbl]

    modo = st.sidebar.radio(
        "🗺️ Visão", ["Uma localidade (3 cenários)", "Todas as cidades (cenário fixo)"],
        key="vy_modo",
    )
    mostrar_tend = st.sidebar.toggle(
        "📏 Linha de tendência", value=False, key="vy_tend",
        help="Adiciona regressão linear sobre a projeção de referência",
    )

    LOC_COLORS = {
        "Catarina":      "#0ea5e9",
        "Nicodemus":     "#f59e0b",
        "Saragoça":      "#22c55e",
        "Santa Bárbara": "#ef4444",
        "Rural":         "#8b5cf6",
        "Utopia":        "#ec4899",
    }
    CEN_COLORS = {"Referencia": ACCENT, "Alto": "#22c55e", "Baixo": "#f59e0b"}

    fig = go.Figure()

    if modo == "Uma localidade (3 cenários)":
        local_sel = st.sidebar.selectbox(
            "📍 Localidade", ALL_LOCS, index=len(ALL_LOCS) - 1, key="vy_local",
        )

        # Banda Alto–Baixo como área preenchida
        alto_s = df[(df["Cenario"] == "Alto")  & (df["Local"] == local_sel)].sort_values("Ano")
        baix_s = df[(df["Cenario"] == "Baixo") & (df["Local"] == local_sel)].sort_values("Ano")
        if not alto_s.empty and not baix_s.empty:
            anos_b = alto_s["Ano"].tolist()
            v_alto = (alto_s[var_col] / var_div).tolist()
            v_baix = (baix_s[var_col] / var_div).tolist()
            x_band = anos_b + anos_b[::-1]
            fig.add_trace(go.Scatter(
                x=x_band, y=v_alto + v_baix[::-1],
                fill="toself",
                fillcolor="rgba(14,165,233,0.07)",
                line=dict(color="rgba(0,0,0,0)"),
                name="Banda ±PIB 3%", hoverinfo="skip",
            ))
            for cen_key, y_vals, dash_style in [
                ("Alto",  v_alto, "dash"),
                ("Baixo", v_baix, "dot"),
            ]:
                fig.add_trace(go.Scatter(
                    x=anos_b, y=y_vals, mode="lines",
                    name=CEN_LBLS[cen_key],
                    line=dict(color=CEN_COLORS[cen_key], width=1.4, dash=dash_style),
                    hovertemplate=f"<b>{CEN_LBLS[cen_key]}</b> %{{x}}: %{{y:,.2f}} {var_unit}<extra></extra>",
                ))

        # Linha referência (destaque)
        ref_s = df[(df["Cenario"] == "Referencia") & (df["Local"] == local_sel)].sort_values("Ano")
        if not ref_s.empty:
            anos_f = ref_s["Ano"].tolist()
            y_ref  = (ref_s[var_col] / var_div).tolist()
            fig.add_trace(go.Scatter(
                x=anos_f, y=y_ref, mode="lines+markers",
                name="Referência",
                line=dict(color=ACCENT, width=2.8),
                marker=dict(size=5),
                hovertemplate=f"<b>Ref.</b> %{{x}}: %{{y:,.3f}} {var_unit}<extra></extra>",
            ))
            # Tendência linear
            if mostrar_tend and len(anos_f) >= 2:
                x_arr = np.array(anos_f)
                y_arr = np.array(y_ref)
                p = np.polyfit(x_arr, y_arr, 1)
                fig.add_trace(go.Scatter(
                    x=anos_f, y=np.polyval(p, x_arr).tolist(),
                    mode="lines", name="Tendência linear",
                    line=dict(color="#f59e0b", width=1.4, dash="dash"),
                    hoverinfo="skip",
                ))

        title_txt = f"{var_lbl} — {local_sel} · 3 Cenários"

    else:
        # Todas as cidades com cenário fixo
        cenario_sel = st.sidebar.radio(
            "📊 Cenário", CEN_NAMES,
            format_func=lambda c: CEN_LBLS[c], key="vy_cenario",
        )
        for loc in ALL_LOCS:
            clr  = LOC_COLORS.get(loc, ACCENT)
            sub  = df[(df["Cenario"] == cenario_sel) & (df["Local"] == loc)].sort_values("Ano")
            if sub.empty:
                continue
            anos_f = sub["Ano"].tolist()
            y_vals = (sub[var_col] / var_div).tolist()
            fig.add_trace(go.Scatter(
                x=anos_f, y=y_vals, mode="lines+markers", name=loc,
                line=dict(color=clr, width=2.2), marker=dict(size=4),
                hovertemplate=f"<b>{loc}</b> %{{x}}: %{{y:,.3f}} {var_unit}<extra></extra>",
            ))
            if mostrar_tend and len(anos_f) >= 2:
                x_arr = np.array(anos_f)
                y_arr = np.array(y_vals)
                p = np.polyfit(x_arr, y_arr, 1)
                fig.add_trace(go.Scatter(
                    x=anos_f, y=np.polyval(p, x_arr).tolist(),
                    mode="lines", name=f"{loc} tend.",
                    line=dict(color=clr, width=1, dash="dot"),
                    showlegend=False, hoverinfo="skip",
                ))
        title_txt = f"{var_lbl} — Todas as Cidades · {CEN_LBLS[cenario_sel]}"

    fig.update_layout(
        title=dict(
            text=title_txt,
            font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0,
        ),
        yaxis_title=f"{var_lbl.split('(')[0].strip()} ({var_unit})",
        xaxis_title="Ano",
        height=460,
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, font=dict(size=10)),
        margin=dict(l=15, r=15, t=60, b=15),
        **PLOTLY_THEME,
    )
    st.plotly_chart(fig, use_container_width=True, key="vy_main")

    # ── Tabela resumo: primeiro, último ano, Δ% ─────────────────────
    st.markdown("---")
    st.markdown(f"#### Resumo — {var_lbl}")

    if modo == "Uma localidade (3 cenários)":
        rows = []
        for cen in CEN_NAMES:
            sub = df[(df["Cenario"] == cen) & (df["Local"] == local_sel)].sort_values("Ano")
            if sub.empty:
                continue
            v0 = sub[var_col].iloc[0]  / var_div
            v1 = sub[var_col].iloc[-1] / var_div
            d  = (v1 - v0) / v0 * 100 if v0 else 0
            rows.append({
                "Cenário":        CEN_LBLS[cen],
                f"{sub['Ano'].iloc[0]}":  f"{v0:,.3f}",
                f"{sub['Ano'].iloc[-1]}": f"{v1:,.3f}",
                "Δ% total":       f"{d:+.1f}%",
                "Unidade":        var_unit,
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        rows = []
        for loc in ALL_LOCS:
            sub = df[(df["Cenario"] == cenario_sel) & (df["Local"] == loc)].sort_values("Ano")
            if sub.empty:
                continue
            v0 = sub[var_col].iloc[0]  / var_div
            v1 = sub[var_col].iloc[-1] / var_div
            d  = (v1 - v0) / v0 * 100 if v0 else 0
            rows.append({
                "Localidade":     loc,
                f"{sub['Ano'].iloc[0]}":  f"{v0:,.3f}",
                f"{sub['Ano'].iloc[-1]}": f"{v1:,.3f}",
                "Δ% total":       f"{d:+.1f}%",
                "Unidade":        var_unit,
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════════
#  ENTRY POINT — chamado pelo app_utopia.py
# ═══════════════════════════════════════════════════════════════════════
def run_projecao(page: str):
    """Recebe a page já resolvida pelo app_utopia.py e despacha."""
    try:
        df = load_proj(CSV_PATH)
    except FileNotFoundError:
        st.error(f"❌ CSV não encontrado:\n`{CSV_PATH}`\n\n"
                 "Confirme que o MATLAB exportou o arquivo e que o caminho está correto.")
        return
    except Exception as e:
        st.error(f"Erro ao ler o CSV: {e}")
        return

    dispatch = {
        "Demanda de Energia":              page_demanda_energia,
        "Histórico + Projeção":            page_hist_proj,
        "Comparativo 2026 - 2036":         page_comparativo,
        "Barras por Cidade":               page_fig1,
        "Composição por setor":            page_fig3,
        "Composição iterativo":            page_fig4,
        "Escolha em função do PIB":        page_fig5,
        "Ranking de evolução":             page_ranking,
        "Decomposição do crescimento":     page_waterfall,
        "Intensidade Energética":          page_intensidade,
        "Heatmap do crescimento":          page_heatmap,
        "Tabela de resultados do Matlab":  page_tabela,
        # ── Novas pages ──────────────────────────────────────────────
        "População por Ano":               page_populacao,
        "Variável Y por Ano":              page_variavel_y,
    }
    fn = dispatch.get(page)
    if fn:
        fn(df)
    else:
        st.warning(f"Página desconhecida: {page}")
