"""
dash_historico.py
======================================================================
Dashboard de Dados Históricos (1975–2025) — País de Utópia
Inclui modo "Construa seu Próprio Gráfico" integrado no mesmo fluxo.
======================================================================
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# ── Caminho base ────────────────────────────────────────────────────
EXCEL_PATH = "ENTREGA_DEMANDA_utopia.xlsx"

# ── Mapa cidade → aba do Excel ──────────────────────────────────────
SHEET_MAP = {
    "Catarina":        "RESUMO CATARINA",
    "Nicodemus":       "RESUMO NICODEMUS",
    "Saragoça":        "RESUMO SARAGOÇA",
    "Santa Bárbara":   "RESUMO SANTA BÁRBARA",
    "Zona Rural":      "RESUMO RURAL",
    "TOTAL — UTÓPIA":  "RESUMO UTÓPIA",
}

# Dicionário amigável para o construtor customizado (Label -> Coluna Real)
VARIABLE_DICTIONARY = {
    "Ano Histórico": "ANO",
    "População Total": "POPULAÇÃO",
    "PEA (Pop. Economicamente Ativa)": "PEA",
    "PIB Absoluto (Utd$)": "PIB",
    "PIB per Capita (Utd$)": "PIB PC",
    "Consumo de Energia Elétrica Total (MWh)": "EE",
    "Índice de Desenvolvimento Humano (IDH)": "IDH",
    "IDH — Renda": "IDH_R",
    "IDH — Educação": "IDH_E",
    "IDH — Saúde": "IDH_S",
    "Anos Médios de Estudo": "ANOS ESTUDO",
    "Domicílios Classe Alta": "No RES Alta",
    "Domicílios Classe Média": "No RES Média",
    "Domicílios Classe Baixa": "No RES Baixa",
    "Consumo EE Residencial Total": "EE RES",
    "Área Residencial Total (km²)": "AR RES",
    "PIB Industrial": "PIB IND",
    "Consumo EE Industrial": "EE IND",
    "Área Ocupada Industrial (km²)": "AR IND",
    "PIB do Setor de Serviços": "PIB SER",
    "Consumo EE Serviços": "EE SER",
    "Área Ocupada Serviços (km²)": "AR SER",
    "PIB Mineral": "PIB MIN",
    "Consumo EE Mineral": "EE MIN",
    "PIB Agropecuário": "PIB AGR",
    "Consumo EE Agropecuário": "EE AGR",
    "Área Territorial Total (km²)": "AR TOTAL",
    "Área de Domínio Público (km²)": "AR PUB"
}

# ── Paleta de Cores e Estilo ────────────────────────────────────────
ACCENT    = "#0ea5e9"
ACCENT_D  = "#0284c7"
TEXT_PRI  = "#0f172a"
TEXT_SEC  = "#64748b"
GRID_CLR  = "rgba(226,232,240,0.6)"
BG_CHART  = "#ffffff"

THEME = dict(
    plot_bgcolor  = BG_CHART,
    paper_bgcolor = BG_CHART,
    font          = dict(family="-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif", color=TEXT_PRI, size=12),
    margin        = dict(l=15, r=15, t=40, b=15),
    xaxis         = dict(showgrid=True, gridcolor=GRID_CLR, linecolor="#e2e8f0", tickfont=dict(size=11)),
    yaxis         = dict(showgrid=True, gridcolor=GRID_CLR, linecolor="#e2e8f0", tickfont=dict(size=11)),
    hoverlabel    = dict(bgcolor="white", font_size=12),
)

# =======================================================================
#  LEITURA DE DADOS
# =======================================================================
@st.cache_data
def load_data(cidade: str, excel_path: str) -> pd.DataFrame:
    sheet = SHEET_MAP[cidade]
    df = pd.read_excel(excel_path, sheet_name=sheet, header=0)
    df["ANO"] = df["ANO"].astype(int)
    return df.sort_values("ANO").reset_index(drop=True)


# =======================================================================
#  HELPERS DE GRÁFICOS
# =======================================================================
def base_fig(title: str, height=280) -> go.Figure:
    fig = go.Figure()
    fig.update_layout(
        title      = dict(text=title, font=dict(size=13, color=TEXT_PRI, weight="bold"), x=0, xanchor="left", pad=dict(l=4, t=4)),
        height     = height,
        showlegend = True,
        legend     = dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        **THEME,
    )
    return fig


def line(fig: go.Figure, x, y, name: str, color: str, dash="solid", width=2.5, fill=None) -> go.Figure:
    fill_color = None
    if fill == "tozeroy":
        if "rgba" in color:
            fill_color = color.replace(",1)", ",0.06)").replace(", 1)", ",0.06)")
        elif "rgb" in color:
            fill_color = color.replace("rgb", "rgba").replace(")", ", 0.06)")
        else:
            fill_color = "rgba(14, 165, 233, 0.06)"

    fig.add_trace(go.Scatter(
        x=x, y=y, mode="lines", name=name,
        line=dict(color=color, width=width, dash=dash),
        fill=fill, fillcolor=fill_color,
        hovertemplate="%{y:,.2f}<extra>" + name + "</extra>",
    ))
    return fig


# =======================================================================
#  DICIONÁRIO MAESTRO DE GRÁFICOS FIXOS
# =======================================================================
def get_chart_builders(df: pd.DataFrame):
    return {
        "Evolução do PIB Absoluto": lambda h=280: (
            line(base_fig("Evolução do PIB Absoluto", h), df["ANO"], df["PIB"], "PIB Total", "rgba(14, 165, 233, 1)", fill="tozeroy")
            .update_layout(yaxis_title="PIB (Utd$)")
        ),
        "Evolução do PIB per Capita": lambda h=280: (
            line(base_fig("Evolução do PIB per Capita", h), df["ANO"], df["PIB PC"], "PIB per Capita", "rgba(2, 132, 199, 1)", fill="tozeroy")
            .update_layout(yaxis_title="PIB PC (Utd$)")
        ),
        "Consumo de Energia Elétrica Total Consolidado": lambda h=280: (
            line(base_fig("Consumo de Energia Elétrica Total Consolidado", h), df["ANO"], df["EE"], "EE Total Consolidada", "rgba(3, 105, 161, 1)", width=3, fill="tozeroy")
            .update_layout(yaxis_title="MWh")
        ),
        "Crescimento de População e PEA": lambda h=280: (
            line(line(base_fig("Crescimento de População e PEA", h), df["ANO"], df["POPULAÇÃO"], "População", "rgba(14, 165, 233, 1)"),
                 df["ANO"], df["PEA"], "PEA (Pop. Economicamente Ativa)", "rgba(124, 58, 237, 1)", dash="dash")
            .update_layout(yaxis_title="Habitantes")
        ),
        "IDH e Componentes (Renda, Educação, Saúde)": lambda h=280: (
            line(line(line(line(base_fig("IDH e Componentes (Renda, Educação, Saúde)", h), df["ANO"], df["IDH"], "IDH Total", ACCENT, width=3),
                           df["ANO"], df["IDH_R"], "IDH Renda", "rgba(245, 158, 11, 1)", width=1.5, dash="dot"),
                      df["ANO"], df["IDH_E"], "IDH Educação", "rgba(16, 185, 129, 1)", width=1.5, dash="dot"),
                 df["ANO"], df["IDH_S"], "IDH Saúde", "rgba(239, 68, 68, 1)", width=1.5, dash="dot")
            .update_layout(yaxis=dict(range=[0, 1.05]))
        ),
        "Anos Médios de Estudo da População": lambda h=280: (
            line(base_fig("Anos Médios de Estudo da População", h), df["ANO"], df["ANOS ESTUDO"], "Anos Médios de Estudo", "rgba(2, 132, 199, 1)", fill="tozeroy")
            .update_layout(yaxis_title="Anos")
        ),
        "Domicílios por Categoria Social": lambda h=280: (
            line(line(line(base_fig("Domicílios por Categoria Social", h), df["ANO"], df["No RES Alta"], "Classe Alta", "rgba(14, 165, 233, 1)"),
                      df["ANO"], df["No RES Média"], "Classe Média", "rgba(56, 189, 248, 1)"),
                 df["ANO"], df["No RES Baixa"], "Classe Baixa", "rgba(186, 230, 253, 1)", dash="dash")
            .update_layout(yaxis_title="Domicílios")
        ),
        "Consumo EE Residencial por Faixa": lambda h=280: (
            line(line(line(line(base_fig("Consumo EE Residencial por Faixa", h), df["ANO"], df["EE RES"], "EE Residencial Total", "rgba(2, 132, 199, 1)", width=3),
                           df["ANO"], df["EE RES A"], "Consumo Classe Alta", "rgba(14, 165, 233, 1)", width=1.5, dash="dot"),
                      df["ANO"], df["EE RES M"], "Consumo Classe Média", "rgba(56, 189, 248, 1)", width=1.5, dash="dot"),
                 df["ANO"], df["EE RES B"], "Consumo Classe Baixa", "rgba(186, 230, 253, 1)", width=1.5, dash="dot")
            .update_layout(yaxis_title="MWh")
        ),
        "Área Residencial por Categoria": lambda h=280: (
            line(line(line(base_fig("Área Residencial por Categoria (km²)", h), df["ANO"], df["AR RES A"], "Área Alta", "rgba(14, 165, 233, 1)"),
                      df["ANO"], df["AR RES M"], "Área Média", "rgba(56, 189, 248, 1)"),
                 df["ANO"], df["AR RES B"], "Área Baixa", "rgba(186, 230, 253, 1)", dash="dash")
            .update_layout(yaxis_title="km²")
        ),
        "Área Residencial Total": lambda h=280: (
            line(base_fig("Área Residencial Total", h), df["ANO"], df["AR RES"], "AR RES Total", "rgba(2, 132, 199, 1)", fill="tozeroy")
            .update_layout(yaxis_title="km²")
        ),
        "Estabelecimentos Industriais por Porte": lambda h=280: (
            line(line(line(base_fig("Estabelecimentos Industriais por Porte", h), df["ANO"], df["No IND G"], "Grandes Indústrias", "rgba(3, 105, 161, 1)"),
                      df["ANO"], df["No IND M"], "Médias Indústrias", "rgba(14, 165, 233, 1)"),
                 df["ANO"], df["No IND P"], "Pequenas Indústrias", "rgba(125, 211, 252, 1)", dash="dash")
            .update_layout(yaxis_title="Unidades")
        ),
        "PIB Industrial": lambda h=280: (
            line(base_fig("PIB Industrial", h), df["ANO"], df["PIB IND"], "PIB Setor Industrial", "rgba(124, 58, 237, 1)", fill="tozeroy")
            .update_layout(yaxis_title="Utd$")
        ),
        "Consumo EE Industrial": lambda h=280: (
            line(line(line(line(base_fig("Consumo EE Industrial", h), df["ANO"], df["EE IND"], "EE Industrial Total", "rgba(2, 132, 199, 1)", width=3),
                           df["ANO"], df["EE IND G"], "Grande Porte", "rgba(3, 105, 161, 1)", width=1.5, dash="dot"),
                      df["ANO"], df["EE IND M"], "Médio Porte", "rgba(14, 165, 233, 1)", width=1.5, dash="dot"),
                 df["ANO"], df["EE IND P"], "Pequeno Porte", "rgba(125, 211, 252, 1)", width=1.5, dash="dot")
            .update_layout(yaxis_title="MWh")
        ),
        "Área Ocupada Industrial": lambda h=280: (
            line(base_fig("Área Ocupada Industrial", h), df["ANO"], df["AR IND"], "AR IND Total", "rgba(124, 58, 237, 1)", fill="tozeroy")
            .update_layout(yaxis_title="km²")
        ),
        "Estabelecimentos de Serviços": lambda h=280: (
            line(line(line(base_fig("Estabelecimentos de Serviços", h), df["ANO"], df["No SER G"], "Grandes Serviços", "rgba(6, 95, 70, 1)"),
                      df["ANO"], df["No SER M"], "Médios Serviços", "rgba(16, 185, 129, 1)"),
                 df["ANO"], df["No SER P"], "Pequenos Serviços", "rgba(167, 243, 208, 1)", dash="dash")
            .update_layout(yaxis_title="Unidades")
        ),
        "PIB do Setor de Serviços": lambda h=280: (
            line(base_fig("PIB do Setor de Serviços", h), df["ANO"], df["PIB SER"], "PIB Setor Serviços", "rgba(16, 185, 129, 1)", fill="tozeroy")
            .update_layout(yaxis_title="Utd$")
        ),
        "Consumo EE Serviços": lambda h=280: (
            line(line(line(line(base_fig("Consumo EE Serviços", h), df["ANO"], df["EE SER"], "EE Serviços Total", "rgba(4, 120, 87, 1)", width=3),
                           df["ANO"], df["EE SER G"], "Grande Porte", "rgba(6, 95, 70, 1)", width=1.5, dash="dot"),
                      df["ANO"], df["EE SER M"], "Médio Porte", "rgba(16, 185, 129, 1)", width=1.5, dash="dot"),
                 df["ANO"], df["EE SER P"], "Pequeno Porte", "rgba(167, 243, 208, 1)", width=1.5, dash="dot")
            .update_layout(yaxis_title="MWh")
        ),
        "Área Ocupada Serviços": lambda h=280: (
            line(base_fig("Área Ocupada Serviços", h), df["ANO"], df["AR SER"], "AR SER Total", "rgba(16, 185, 129, 1)", fill="tozeroy")
            .update_layout(yaxis_title="km²")
        ),
        "Extratoras Minerais": lambda h=280: (
            line(line(line(base_fig("Extratoras Minerais por Porte", h), df["ANO"], df["No MIN G"], "Grandes Minerações", "rgba(180, 83, 9, 1)"),
                      df["ANO"], df["No MIN M"], "Médias Minerações", "rgba(245, 158, 11, 1)"),
                 df["ANO"], df["No MIN P"], "Pequenas Minerações", "rgba(253, 230, 138, 1)", dash="dash")
            .update_layout(yaxis_title="Unidades")
        ),
        "PIB Mineral": lambda h=280: (
            line(base_fig("PIB Mineral", h), df["ANO"], df["PIB MIN"], "PIB Setor Mineral", "rgba(245, 158, 11, 1)", fill="tozeroy")
            .update_layout(yaxis_title="Utd$")
        ),
        "Consumo EE Mineral": lambda h=280: (
            line(line(line(line(base_fig("Consumo EE Mineral", h), df["ANO"], df["EE MIN"], "EE Mineral Total", "rgba(146, 64, 14, 1)", width=3),
                           df["ANO"], df["EE MIN G"], "Grande Porte", "rgba(180, 83, 9, 1)", width=1.5, dash="dot"),
                      df["ANO"], df["EE MIN M"], "Médio Porte", "rgba(245, 158, 11, 1)", width=1.5, dash="dot"),
                 df["ANO"], df["EE MIN P"], "Pequeno Porte", "rgba(253, 230, 138, 1)", width=1.5, dash="dot")
            .update_layout(yaxis_title="MWh")
        ),
        "Área Mineral": lambda h=280: (
            line(base_fig("Área Mineral", h), df["ANO"], df["AR MIN"], "AR MIN Total", "rgba(245, 158, 11, 1)", fill="tozeroy")
            .update_layout(yaxis_title="km²")
        ),
        "Propriedades Rurais Agropecuárias": lambda h=280: (
            line(line(line(base_fig("Propriedades Rurais Agropecuárias", h), df["ANO"], df["No AGR G"], "Grandes Propriedades", "rgba(21, 128, 61, 1)"),
                      df["ANO"], df["No AGR M"], "Médias Propriedades", "rgba(34, 197, 95, 1)"),
                 df["ANO"], df["No AGR P"], "Pequenas Propriedades", "rgba(187, 247, 208, 1)", dash="dash")
            .update_layout(yaxis_title="Unidades")
        ),
        "PIB Agropecuário": lambda h=280: (
            line(base_fig("PIB Agropecuário", h), df["ANO"], df["PIB AGR"], "PIB Setor Agropecuário", "rgba(34, 197, 95, 1)", fill="tozeroy")
            .update_layout(yaxis_title="Utd$")
        ),
        "Consumo EE Agropecuário": lambda h=280: (
            line(line(line(line(base_fig("Consumo EE Agropecuário", h), df["ANO"], df["EE AGR"], "EE Agropecuária Total", "rgba(22, 101, 52, 1)", width=3),
                           df["ANO"], df["EE AGR G"], "Grande Porte", "rgba(21, 128, 61, 1)", width=1.5, dash="dot"),
                      df["ANO"], df["EE AGR M"], "Médio Porte", "rgba(34, 197, 95, 1)", width=1.5, dash="dot"),
                 df["ANO"], df["EE AGR P"], "Pequeno Porte", "rgba(187, 247, 208, 1)", width=1.5, dash="dot")
            .update_layout(yaxis_title="MWh")
        ),
        "Área Agropecuária": lambda h=280: (
            line(base_fig("Área Agropecuária", h), df["ANO"], df["AR AGR"], "AR AGR Total", "rgba(34, 197, 95, 1)", fill="tozeroy")
            .update_layout(yaxis_title="km²")
        ),
        "EE Utilidades Públicas": lambda h=280: (
            line(line(line(line(base_fig("EE Utilidades Públicas", h), df["ANO"], df["EE UTIL"], "EE Utilidades Total", "rgba(76, 29, 149, 1)", width=3),
                           df["ANO"], df["EE UTIL G"], "Grande Porte", "rgba(109, 40, 217, 1)", width=1.5, dash="dot"),
                      df["ANO"], df["EE UTIL M"], "Médio Porte", "rgba(139, 92, 246, 1)", width=1.5, dash="dot"),
                 df["ANO"], df["EE UTIL P"], "Pequeno Porte", "rgba(196, 181, 253, 1)", width=1.5, dash="dot")
            .update_layout(yaxis_title="MWh")
        ),
        "Área Utilidades": lambda h=280: (
            line(base_fig("Área Utilidades", h), df["ANO"], df["AR UTIL"], "AR UTIL Total", "rgba(139, 92, 246, 1)", fill="tozeroy")
            .update_layout(yaxis_title="km²")
        ),
        "Comparativo Área Total vs Área Pública": lambda h=280: (
            line(line(base_fig("Comparativo Área Total vs Área Pública", h), df["ANO"], df["AR TOTAL"], "AR Total Territorial", ACCENT, width=2.5),
                 df["ANO"], df["AR PUB"], "Área de Domínio Público", "rgba(248, 113, 113, 1)", width=2, dash="dash")
            .update_layout(yaxis_title="km²")
        )
    }


# =======================================================================
#  KPI CARDS
# =======================================================================
def render_kpis(df: pd.DataFrame, cidade: str):
    last = df.iloc[-1]
    first = df.iloc[0]
    ano_last = int(last["ANO"])
    ano_first = int(first["ANO"])

    def delta_pct(now, prev):
        if prev == 0: return "—"
        return f"{((now - prev) / abs(prev)) * 100:+.1f}%"

    kpis = [
        ("👥 População",   f"{last['POPULAÇÃO']:,.0f}",       delta_pct(last["POPULAÇÃO"], first["POPULAÇÃO"])),
        ("💰 PIB Absoluto", f"Utd$ {last['PIB']:,.0f}",        delta_pct(last["PIB"], first["PIB"])),
        ("📊 PIB PC",      f"Utd$ {last['PIB PC']:,.2f}",     delta_pct(last["PIB PC"], first["PIB PC"])),
        ("🌟 IDH",         f"{last['IDH']:.4f}",              delta_pct(last["IDH"], first["IDH"])),
        ("⚡ EE Total",    f"{last['EE']:,.0f} MWh",          delta_pct(last["EE"], first["EE"])),
        ("🗺️ Área Total",  f"{last['AR TOTAL']:,.2f} km²",   delta_pct(last["AR TOTAL"], first["AR TOTAL"])),
    ]

    cols = st.columns(len(kpis))
    css_card = "background:#ffffff; border:1px solid #e2e8f0; border-radius:14px; padding:12px 14px; box-shadow:0 2px 10px rgba(14,165,233,0.05);"
    
    for col, (label, val, dlt) in zip(cols, kpis):
        color = "#22c55e" if "+" in dlt else "#ef4444" if "-" in dlt else TEXT_SEC
        with col:
            st.markdown(
                f'<div style="{css_card}">'
                f'<div style="font-size:10px; color:{TEXT_SEC}; font-weight:600; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:2px;">{label}</div>'
                f'<div style="font-size:16px; font-weight:700; color:{TEXT_PRI}; letter-spacing:-0.3px;">{val}</div>'
                f'<div style="font-size:11px; color:{color}; margin-top:2px; font-weight:500;">{ano_first}→{ano_last}: {dlt}</div>'
                f'</div>', unsafe_allow_html=True
            )
    st.markdown("<br>", unsafe_allow_html=True)


# =======================================================================
#  LAYOUT EM GRADE FIXA DE 29 GRÁFICOS (RETORNADO AO PADRÃO ORIGINAL)
# =======================================================================
def render_grid_layout(df: pd.DataFrame, cidade_key: str):
    builders = get_chart_builders(df)

    st.markdown("### 📈 Indicadores Gerais e Macroeconômicos")
    r1_c1, r1_c2, r1_c3 = st.columns(3)
    with r1_c1: st.plotly_chart(builders["Evolução do PIB Absoluto"](), use_container_width=True, key=f"pib_abs_{cidade_key}")
    with r1_c2: st.plotly_chart(builders["Evolução do PIB per Capita"](), use_container_width=True, key=f"pib_pc_{cidade_key}")
    with r1_c3: st.plotly_chart(builders["Consumo de Energia Elétrica Total Consolidado"](), use_container_width=True, key=f"ee_tot_{cidade_key}")

    r2_c1, r2_c2, r2_c3 = st.columns(3)
    with r2_c1: st.plotly_chart(builders["Crescimento de População e PEA"](), use_container_width=True, key=f"pop_pea_{cidade_key}")
    with r2_c2: st.plotly_chart(builders["IDH e Componentes (Renda, Educação, Saúde)"](), use_container_width=True, key=f"idh_comp_{cidade_key}")
    with r2_c3: st.plotly_chart(builders["Anos Médios de Estudo da População"](), use_container_width=True, key=f"estudo_{cidade_key}")

    st.markdown("---")
    st.markdown("### 🏠 Setor Residencial")
    r3_c1, r3_c2, r3_c3 = st.columns(3)
    with r3_c1: st.plotly_chart(builders["Domicílios por Categoria Social"](), use_container_width=True, key=f"dom_cat_{cidade_key}")
    with r3_c2: st.plotly_chart(builders["Consumo EE Residencial por Faixa"](), use_container_width=True, key=f"ee_res_{cidade_key}")
    with r3_c3: st.plotly_chart(builders["Área Residencial por Categoria"](), use_container_width=True, key=f"ar_res_{cidade_key}")

    st.markdown("---")
    st.markdown("### 🏭 Indústria & 🏪 Serviços")
    r4_c1, r4_c2, r4_c3 = st.columns(3)
    with r4_c1: st.plotly_chart(builders["Estabelecimentos Industriais por Porte"](), use_container_width=True, key=f"ind_porte_{cidade_key}")
    with r4_c2: st.plotly_chart(builders["PIB Industrial"](), use_container_width=True, key=f"pib_ind_{cidade_key}")
    with r4_c3: st.plotly_chart(builders["Consumo EE Industrial"](), use_container_width=True, key=f"ee_ind_{cidade_key}")

    r5_c1, r5_c2, r5_c3 = st.columns(3)
    with r5_c1: st.plotly_chart(builders["Estabelecimentos de Serviços"](), use_container_width=True, key=f"ser_porte_{cidade_key}")
    with r5_c2: st.plotly_chart(builders["PIB do Setor de Serviços"](), use_container_width=True, key=f"pib_ser_{cidade_key}")
    with r5_c3: st.plotly_chart(builders["Consumo EE Serviços"](), use_container_width=True, key=f"ee_ser_{cidade_key}")

    st.markdown("---")
    st.markdown("### ⛏️ Mineração, 🌾 Agropecuária & ⚡ Infraestrutura")
    r6_c1, r6_c2, r6_c3 = st.columns(3)
    with r6_c1: st.plotly_chart(builders["Extratoras Minerais"](), use_container_width=True, key=f"min_porte_{cidade_key}")
    with r6_c2: st.plotly_chart(builders["PIB Mineral"](), use_container_width=True, key=f"pib_min_{cidade_key}")
    with r6_c3: st.plotly_chart(builders["Consumo EE Mineral"](), use_container_width=True, key=f"ee_min_{cidade_key}")

    r7_c1, r7_c2, r7_c3 = st.columns(3)
    with r7_c1: st.plotly_chart(builders["Propriedades Rurais Agropecuárias"](), use_container_width=True, key=f"agr_porte_{cidade_key}")
    with r7_c2: st.plotly_chart(builders["PIB Agropecuário"](), use_container_width=True, key=f"pib_agr_{cidade_key}")
    with r7_c3: st.plotly_chart(builders["Consumo EE Agropecuário"](), use_container_width=True, key=f"ee_agr_{cidade_key}")

    r8_c1, r8_c2, r8_c3 = st.columns(3)
    with r8_c1: st.plotly_chart(builders["EE Utilidades Públicas"](), use_container_width=True, key=f"ee_util_{cidade_key}")
    with r8_c2: st.plotly_chart(builders["Área Ocupada Industrial"](), use_container_width=True, key=f"ar_ind_tot_{cidade_key}")
    with r8_c3: st.plotly_chart(builders["Comparativo Área Total vs Área Pública"](), use_container_width=True, key=f"ar_total_pub_{cidade_key}")


# =======================================================================
#  CONSTRUTOR CUSTOMIZADO CORRIGIDO (PROPRIO GRAFICO)
# =======================================================================
def render_custom_builder_layout(df: pd.DataFrame, cidade_sel: str):
    """Renderiza o espaço de trabalho dinâmico se o botão do sidebar estiver ativo."""
    st.markdown("### 🧪 Laboratório Customizado Ativo")
    st.info("Abaixo você pode cruzar livremente qualquer par de dados da tabela. O período de anos filtrado na barra lateral será respeitado.")

    # Selectores na zona de trabalho principal
    ctrl_c1, ctrl_c2 = st.columns(2)
    with ctrl_c1:
        x_label = st.selectbox("Selecione a Variável do Eixo X (Horizontal):", list(VARIABLE_DICTIONARY.keys()), index=0, key=f"custom_x_{cidade_sel}")
    with ctrl_c2:
        y_label = st.selectbox("Selecione a Variável do Eixo Y (Vertical):", list(VARIABLE_DICTIONARY.keys()), index=3, key=f"custom_y_{cidade_sel}")

    col_x = VARIABLE_DICTIONARY[x_label]
    col_y = VARIABLE_DICTIONARY[y_label]

    # Limpeza absoluta de duplicatas de colunas
    df_base = df[[col_x, col_y, "ANO"]].dropna()
    df_plot = df_base.loc[:, ~df_base.columns.duplicated()].copy()
    
    # Ordenamento pelo eixo X para fluxo perfeito da linha
    df_plot = df_plot.sort_values(by=col_x).reset_index(drop=True)

    fig = go.Figure()

    color_line = "rgba(14, 165, 233, 1)"
    color_fill = "rgba(14, 165, 233, 0.06)"

    if "IDH" in y_label:
        color_line = "rgba(16, 185, 129, 1)"
        color_fill = "rgba(16, 185, 129, 0.06)"
    elif "Consumo" in y_label or "EE" in col_y:
        color_line = "rgba(2, 132, 199, 1)"
        color_fill = "rgba(2, 132, 199, 0.06)"

    # Linha com marcadores históricos discretos
    fig.add_trace(go.Scatter(
        x=df_plot[col_x], y=df_plot[col_y],
        mode="lines+markers",
        marker=dict(size=4),
        name=y_label,
        line=dict(color=color_line, width=2.5),
        fill="tozeroy", fillcolor=color_fill,
        customdata=df_plot["ANO"],
        hovertemplate=(
            f"<b>{x_label}:</b> %{{x:,.2f}}<br>"
            f"<b>{y_label}:</b> %{{y:,.2f}}<br>"
            "<b>Ano de referência:</b> %{customdata}<extra></extra>"
        )
    ))

    titulo_grafico = f"Correlation: {y_label} vs {x_label} ({cidade_sel})"
    if x_label == "Ano Histórico":
        titulo_grafico = f"Evolução Temporal de: {y_label} ({cidade_sel})"

    fig.update_layout(
        title=dict(text=titulo_grafico, font=dict(size=14, color=TEXT_PRI, weight="bold"), x=0),
        xaxis_title=x_label, yaxis_title=y_label,
        height=500, **THEME
    )

    st.plotly_chart(fig, use_container_width=True, key=f"chart_custom_build_{cidade_sel}")


# =======================================================================
#  ENTRY POINT (CHAMADO PELO APP_UTOPIA.PY)
# =======================================================================
def run_historico(cidade: str):
    """Executado pelo arquivo central do app."""
    
    # 1. TÍTULO GRANDE E ADAPTATIVO
    if "TOTAL" in cidade or "UTÓPIA" in cidade:
        texto_titulo = "Histórico do País: UTÓPIA"
    else:
        texto_titulo = f"Histórico da Cidade de {cidade}"

    st.markdown(
        f'<h1 style="font-size:32px; font-weight:800; color:{TEXT_PRI}; letter-spacing:-0.6px; margin-bottom:2px;">{texto_titulo}</h1>'
        f'<p style="font-size:14px; color:{TEXT_SEC}; margin-bottom:20px;">Série histórica integrada · Planejamento Energético Integrado (1975 – 2025)</p>',
        unsafe_allow_html=True
    )

    # ── CARGA BÁSICA DE DADOS ──────────────────────────────────────────
    try:
        df_ori = load_data(cidade, EXCEL_PATH)
    except Exception as e:
        st.error(f"Erro ao ler o Excel: {e}")
        return

    ano_min, ano_max = int(df_ori["ANO"].min()), int(df_ori["ANO"].max())

    # ── CONFIGURAÇÃO DO SIDEBAR (INTEGRAÇÃO COMPLETA) ─────────────────
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🛠️ Modos Extras de Visualização")
    
    # Checkbox premium integrado na barra lateral
    modo_custom = st.sidebar.checkbox(
        "✨ Construa seu Próprio Gráfico", 
        value=False, 
        key=f"chk_custom_{cidade}",
        help="Ative para ocultar os blocos fixos e cruzar quaisquer variáveis socioeconômicas livres."
    )

    # Layout de inputs de ano na barra lateral
    st.sidebar.markdown("**Filtrar Período Histórico:**")
    side_col1, side_col2 = st.sidebar.columns(2)
    with side_col1:
        a1 = side_col2.number_input("Ano início", min_value=ano_min, max_value=ano_max, value=ano_min, step=1, key=f"a1_{cidade}")
    with side_col2:
        a2 = side_col1.number_input("Ano fim",    min_value=ano_min, max_value=ano_max, value=ano_max, step=1, key=f"a2_{cidade}")

    # Processar filtros de ano
    df = df_ori[(df_ori["ANO"] >= a1) & (df_ori["ANO"] <= a2)].reset_index(drop=True)
    if df.empty:
        st.warning("Nenhum dado no intervalo selecionado.")
        return

    # Chave única para controle de concorrência de chaves de componentes no Streamlit
    cidade_key = cidade.replace(" ", "_").replace("—", "").strip()

    # ── RENDERIZADO CONDICIONAL DA PÁGINA PRINCIPAL ───────────────────
    if modo_custom:
        # CASO A: Usuário ativou o laboratório livre customizado
        render_kpis(df, cidade)
        render_custom_builder_layout(df, cidade)
    else:
        # CASO B: Dashboard Tradicional Completo em Grade com 29 gráficos fluindo direto
        render_kpis(df, cidade)
        st.markdown("---")
        render_grid_layout(df, cidade_key)
