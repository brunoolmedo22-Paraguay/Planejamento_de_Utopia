# -*- coding: utf-8 -*-
"""
dash_matriz.py · Seção "Matriz Energética" do app UTÓPIA
==========================================================
Plano Decenal de Expansão (2025 → 2035) — parte econômica e resultados finais.

Abas
----
1. Atualidade (2025)      → matriz atual, equivalências de energia, socioeconômico,
                            estimativa de potência instalada a partir do FC de cada fonte.
2. Configuração Econômica → tabela-modelo (CAPEX/OPEX/FC) — fonte de verdade do cálculo.
3. Análise Econômica      → calculadora VPL + tarifa (LCOE), replicando a planilha
                            CAPEXOPEXYESO.xlsx (Sheet2 = cálculos, Sheet3 = VPL no tempo).

Modelo econômico (idêntico à planilha)
--------------------------------------
    EE_anual   = P · 8760 · (FC/100)                      [MWh/ano]
    CAPEX      = A · (P · 1000)^B                          [US$]   (P em MW → kW)
    OPEX_F/ano = (OPEX_FIXO% / 100) · CAPEX                [US$/ano]
    OPEX_V/ano = OPEX_VAR · EE_anual                       [US$/ano]  (OPEX_VAR em US$/MWh)
    Parcela    = CAPEX / FA(LFSP, WACC)                    [US$/ano]  (amortização em LFSP anos)
                 FA(n,i) = ((1+i)^n - 1) / ((1+i)^n · i)   (fator de valor presente da anuidade)
    VPL_x      = Σ_{t=0..ΔOP}  R_x / (1+WACC)^t            (só os anos contados até 2035)
    Tarifa     = VPL_custo / VPL_energia                   [US$/MWh]  (LCOE)

Arquivos de dados (na raiz do projeto)
--------------------------------------
    - ENTREGA_DEMANDA_utopia.xlsx  (aba "RESUMO UTÓPIA" → POPULAÇÃO / PIB / EE de 2025)
"""

import base64
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# =======================================================================
#  RAIZ / ARQUIVOS
# =======================================================================
ROOT       = Path(__file__).parent
EXCEL_HIST = str(ROOT / "ENTREGA_DEMANDA_utopia.xlsx")

# Fallback caso o Excel socioeconômico não esteja disponível (ex.: deploy parcial)
EE_TOTAL_FALLBACK = 1_450_000.0   # MWh (≈ 1,45 TWh — último valor de 2025)

# =======================================================================
#  PALETA / TEMA  (idêntico aos demais dashes p/ consistência visual)
# =======================================================================
ACCENT   = "#0ea5e9"
ACCENT_D = "#0284c7"
TEXT_PRI = "#0f172a"
TEXT_SEC = "#64748b"
GRID_CLR = "rgba(226,232,240,0.6)"
BG_CHART = "#ffffff"

SOL = "#f59e0b"   # solar (âmbar)
WIN = "#10b981"   # eólica (verde)
HYD = "#0ea5e9"   # hidro (azul)
THR = "#ef4444"   # termo (vermelho)

# cores dos componentes de custo
C_CAPEX = "#0284c7"   # sky-600
C_OPEXF = "#7c3aed"   # violeta
C_OPEXV = "#f59e0b"   # âmbar

THEME = dict(
    plot_bgcolor=BG_CHART,
    paper_bgcolor=BG_CHART,
    font=dict(family="-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif", color=TEXT_PRI, size=12),
    margin=dict(l=15, r=15, t=46, b=15),
    xaxis=dict(showgrid=True, gridcolor=GRID_CLR, linecolor="#e2e8f0", tickfont=dict(size=11)),
    yaxis=dict(showgrid=True, gridcolor=GRID_CLR, linecolor="#e2e8f0", tickfont=dict(size=11)),
    hoverlabel=dict(bgcolor="white", font_size=12),
)

# =======================================================================
#  TABELA-MODELO (ENTRADA FIXA) — exatamente a planilha CAPEXOPEXYESO.xlsx
#  nome, limite, A, B, FC(%), OPEX_FIXO(%CAPEX/ano), OPEX_VAR(US$/MWh), grupo
# =======================================================================
FONTES = [
    ("Biogás (resíduos)",          "≤ 10 MW",   3800, 0.82, 80, 5.0,   8, "Térmica"),
    ("Biomassa (cana/madeira)",    "≤ 50 MW",   2800, 0.88, 65, 4.0,  24, "Térmica"),
    ("Gás natural (ciclo comb.)",  "≥ 50 MW",   1900, 0.85, 75, 3.0,  88, "Térmica"),
    ("Nuclear (urânio)",           "≥ 500 MW",  9000, 0.92, 92, 6.0,   6, "Térmica"),
    ("Óleo diesel",                "≤ 100 MW",  1000, 0.95, 15, 2.0, 220, "Térmica"),
    ("Solar FV",                   "—",         1700, 0.94, 26, 1.8,   0, "Solar"),
    ("Eólica onshore",             "—",         2400, 0.92, 42, 2.5,   0, "Eólica"),
    ("Eólica offshore",            "≥ 50 MW",   4400, 0.90, 52, 4.0,   0, "Eólica"),
    ("CGH",                        "≤ 5 MW",    3600, 0.85, 60, 2.0,   0, "Hídrica"),
    ("PCH",                        "5 – 30 MW", 3000, 0.88, 55, 1.5,   0, "Hídrica"),
    ("UHE",                        "> 30 MW",   2600, 0.86, 50, 1.2,   0, "Hídrica"),
]
COLS_FONTE = ["Fonte", "Limite", "A", "B", "FC", "OPEX_FIXO", "OPEX_VAR", "Grupo"]
WACC_PADRAO = 7.0   # % a.a. (Sheet1!N9)


def fontes_df() -> pd.DataFrame:
    return pd.DataFrame(FONTES, columns=COLS_FONTE)


def fonte_lookup(nome: str) -> dict:
    df = fontes_df()
    r = df[df["Fonte"] == nome].iloc[0]
    return dict(A=float(r["A"]), B=float(r["B"]), fc=float(r["FC"]),
               opex_fix=float(r["OPEX_FIXO"]), opex_var=float(r["OPEX_VAR"]),
               limite=str(r["Limite"]), grupo=str(r["Grupo"]))


# =======================================================================
#  HELPERS GERAIS
# =======================================================================
def _fmt(v, dec=0, suf=""):
    """Número com separador de milhar estilo BR (1.234.567,8)."""
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return "—"
    s = f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s}{suf}"


def _us(v, dec=0):
    """Valor monetário em US$ (formatação BR de milhares)."""
    return f"US$ {_fmt(v, dec)}"


def kpi_card(label, value, sub="", color=TEXT_PRI):
    css = ("background:#ffffff; border:1px solid #e2e8f0; border-radius:14px;"
           "padding:12px 14px; box-shadow:0 2px 10px rgba(14,165,233,0.05);"
           "display:flex; flex-direction:column; justify-content:center; height:100%;")
    sub_html = (f'<div style="font-size:11px;color:{TEXT_SEC};margin-top:3px;font-weight:500;">{sub}</div>'
                if sub else "")
    return (
        f'<div style="height:100%;display:flex;flex-direction:column;">'
        f'<div style="{css}">'
        f'<div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;'
        f'letter-spacing:0.05em;margin-bottom:2px;">{label}</div>'
        f'<div style="font-size:18px;font-weight:700;color:{color};letter-spacing:-0.3px;">{value}</div>'
        f'{sub_html}</div></div>'
    )


def section_title(txt, sub=""):
    sub_html = f'<p style="font-size:14px;color:{TEXT_SEC};margin:2px 0 18px;">{sub}</p>' if sub else ""
    st.markdown(
        f'<h2 style="font-size:24px;font-weight:800;color:{TEXT_PRI};letter-spacing:-0.4px;'
        f'margin:6px 0 2px;">{txt}</h2>{sub_html}',
        unsafe_allow_html=True,
    )


def base_fig(title, height=320):
    fig = go.Figure()
    fig.update_layout(
        title=dict(text=title, font=dict(size=13, color=TEXT_PRI, weight="bold"),
                   x=0, xanchor="left", pad=dict(l=4, t=4)),
        height=height, showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        **THEME,
    )
    return fig


def _gap(px=10):
    st.markdown(f"<div style='height:{px}px'></div>", unsafe_allow_html=True)


# =======================================================================
#  CARGA DE DADOS SOCIOECONÔMICOS (2025)
# =======================================================================
@st.cache_data(ttl=300)
def load_socio_2025(excel_path: str) -> dict:
    """População, PIB e consumo EE de 2025 — aba RESUMO UTÓPIA."""
    df = pd.read_excel(excel_path, sheet_name="RESUMO UTÓPIA", header=0)
    df["ANO"] = df["ANO"].astype(int)
    row = df[df["ANO"] == 2025]
    row = row.iloc[0] if not row.empty else df.sort_values("ANO").iloc[-1]
    return {
        "ano": int(row["ANO"]),
        "pop": float(row["POPULAÇÃO"]),
        "pib": float(row["PIB"]),
        "pib_pc": float(row["PIB PC"]) if "PIB PC" in row else float(row["PIB"]) / float(row["POPULAÇÃO"]),
        "ee": float(row["EE"]),   # consumo total em MWh
    }


# =======================================================================
#  NÚCLEO ECONÔMICO  (replica a planilha — funções puras, testáveis)
# =======================================================================
def annuity_pv_factor(n: int, i: float) -> float:
    """Fator de valor presente de uma anuidade: FA = ((1+i)^n - 1)/((1+i)^n · i)."""
    if i == 0:
        return float(n)
    return ((1 + i) ** n - 1) / (((1 + i) ** n) * i)


def annuity_payment(capex: float, lfsp: int, i: float) -> float:
    """Parcela anual que amortiza o CAPEX em `lfsp` anos à taxa `i` (Sheet3!D17)."""
    return capex / annuity_pv_factor(lfsp, i)


def compute_economics(A, B, fc_pct, opex_fix_pct, opex_var,
                      potencia_mw, delta_op, lfsp, wacc_pct) -> dict:
    """
    Reproduz CAPEXOPEXYESO.xlsx para UMA usina.
    `delta_op` = nº de anos contados além do ano 0 (série t = 0..delta_op, inclusiva).
    Retorna grandezas anuais, VPLs e a tarifa (LCOE).
    """
    w = wacc_pct / 100.0

    ee_anual     = potencia_mw * 8760.0 * (fc_pct / 100.0)        # MWh/ano
    capex        = A * (potencia_mw * 1000.0) ** B                # US$
    opex_f_anual = (opex_fix_pct / 100.0) * capex                 # US$/ano
    opex_v_anual = opex_var * ee_anual                            # US$/ano
    pmt_capex    = annuity_payment(capex, lfsp, w)                # US$/ano

    anos = list(range(0, int(delta_op) + 1))
    disc = [1.0 / (1 + w) ** t for t in anos]                     # fatores de desconto

    serie_capex   = [pmt_capex    * d for d in disc]
    serie_opex_f  = [opex_f_anual * d for d in disc]
    serie_opex_v  = [opex_v_anual * d for d in disc]
    serie_energia = [ee_anual     * d for d in disc]

    npv_capex   = float(np.sum(serie_capex))
    npv_opex_f  = float(np.sum(serie_opex_f))
    npv_opex_v  = float(np.sum(serie_opex_v))
    npv_total   = npv_capex + npv_opex_f + npv_opex_v
    npv_energia = float(np.sum(serie_energia))

    lcoe = npv_total / npv_energia if npv_energia > 0 else float("nan")   # US$/MWh

    return dict(
        ee_anual=ee_anual, capex=capex, opex_f_anual=opex_f_anual,
        opex_v_anual=opex_v_anual, pmt_capex=pmt_capex,
        anos=anos, disc=disc,
        serie_capex=serie_capex, serie_opex_f=serie_opex_f,
        serie_opex_v=serie_opex_v, serie_energia=serie_energia,
        npv_capex=npv_capex, npv_opex_f=npv_opex_f, npv_opex_v=npv_opex_v,
        npv_total=npv_total, npv_energia=npv_energia,
        lcoe=lcoe,
        # decomposição da tarifa (US$/MWh)
        tar_capex=(npv_capex / npv_energia if npv_energia > 0 else float("nan")),
        tar_opex_f=(npv_opex_f / npv_energia if npv_energia > 0 else float("nan")),
        tar_opex_v=(npv_opex_v / npv_energia if npv_energia > 0 else float("nan")),
    )


# =======================================================================
#  ABA 1 · ATUALIDADE (2025)
# =======================================================================
# Matriz atual (premissas do usuário)
MATRIZ_2025 = {"Hidro": 40.0, "Termo": 55.0, "Eólica": 5.0, "Solar": 0.0}
MATRIZ_COR  = {"Hidro": HYD, "Termo": THR, "Eólica": WIN, "Solar": SOL}

# opções de FC por tecnologia (rótulo exibido → fonte da tabela)
FC_OPCOES = {
    "Hidro":  ["UHE", "PCH", "CGH"],
    "Termo":  ["Gás natural (ciclo comb.)", "Biomassa (cana/madeira)", "Biogás (resíduos)",
               "Óleo diesel", "Nuclear (urânio)"],
    "Eólica": ["Eólica onshore", "Eólica offshore"],
    "Solar":  ["Solar FV"],
}


def tab_atualidade():
    section_title("Matriz Energética Atual · 2025",
                 "Composição da geração, equivalências de energia e indicadores socioeconômicos do País de Utópia")

    # ── dados socioeconômicos (com fallback robusto) ──────────────────
    socio, fonte_ok = None, True
    try:
        socio = load_socio_2025(EXCEL_HIST)
    except Exception:
        fonte_ok = False

    ee_total = socio["ee"] if socio else EE_TOTAL_FALLBACK     # MWh
    pop      = socio["pop"] if socio else None
    pib      = socio["pib"] if socio else None
    pib_pc   = socio["pib_pc"] if socio else None

    if not fonte_ok:
        st.caption("⚠️ `ENTREGA_DEMANDA_utopia.xlsx` não encontrado — usando EE total ≈ 1,45 TWh "
                  "(fallback). População e PIB indisponíveis.")

    # ── KPIs socioeconômicos ──────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(kpi_card("População 2025", _fmt(pop) if pop else "—", "habitantes", ACCENT_D),
                unsafe_allow_html=True)
    c2.markdown(kpi_card("PIB 2025", f"Utd$ {_fmt(pib)}" if pib else "—", "absoluto"),
                unsafe_allow_html=True)
    c3.markdown(kpi_card("PIB per capita", f"Utd$ {_fmt(pib_pc)}" if pib_pc else "—",
                        "por habitante", "#7c3aed"), unsafe_allow_html=True)
    c4.markdown(kpi_card("Energia Elétrica Total", _fmt(ee_total / 1e6, 3, " TWh"),
                        "consumo/geração 2025", ACCENT), unsafe_allow_html=True)

    _gap(8)

    # ── Equivalências de energia (o KPI em kWh que você pediu) ─────────
    st.markdown(f'<div style="font-size:11px;font-weight:700;letter-spacing:0.08em;'
               f'text-transform:uppercase;color:{TEXT_SEC};margin:6px 0 6px;">'
               f'Equivalências da energia total de 2025</div>', unsafe_allow_html=True)
    e1, e2, e3, e4 = st.columns(4)
    e1.markdown(kpi_card("Em TWh", _fmt(ee_total / 1e6, 3), "terawatt-hora"), unsafe_allow_html=True)
    e2.markdown(kpi_card("Em GWh", _fmt(ee_total / 1e3, 0), "gigawatt-hora"), unsafe_allow_html=True)
    e3.markdown(kpi_card("Em MWh", _fmt(ee_total, 0), "megawatt-hora"), unsafe_allow_html=True)
    e4.markdown(kpi_card("Em kWh", _fmt(ee_total * 1e3, 0), "quilowatt-hora", ACCENT_D),
                unsafe_allow_html=True)

    st.markdown("---")

    # ── Composição da matriz (premissas editáveis) ────────────────────
    cL, cR = st.columns([1, 1.15], gap="large")

    with cL:
        st.markdown("##### Composição da matriz (2025)")
        st.caption("Participações de referência — ajustáveis para análise.")
        g = st.columns(4)
        shares = {}
        for col, (fonte, val) in zip(g, MATRIZ_2025.items()):
            shares[fonte] = col.number_input(fonte, min_value=0.0, max_value=100.0,
                                            value=float(val), step=1.0, key=f"mz_{fonte}")
        soma = sum(shares.values())
        if abs(soma - 100.0) > 0.01:
            st.caption(f"Σ = {_fmt(soma,1)} % — os valores serão normalizados para 100 % nos cálculos.")
        # normaliza p/ cálculo
        shares_n = {k: (v / soma * 100.0 if soma > 0 else 0.0) for k, v in shares.items()}

        # donut
        fig = base_fig("Participação por fonte (%)", height=300)
        labels = list(shares_n.keys())
        vals   = [shares_n[k] for k in labels]
        cores  = [MATRIZ_COR[k] for k in labels]
        fig.add_trace(go.Pie(
            labels=labels, values=vals, hole=0.58,
            marker=dict(colors=cores, line=dict(color="#ffffff", width=2)),
            textinfo="label+percent", textfont=dict(size=12),
            hovertemplate="%{label}<br>%{percent}<extra></extra>",
            sort=False,
        ))
        fig.update_layout(showlegend=False,
                          annotations=[dict(text="Matriz<br>2025", x=0.5, y=0.5,
                                            font=dict(size=14, color=TEXT_PRI, weight="bold"),
                                            showarrow=False)])
        st.plotly_chart(fig, use_container_width=True, key="mz_donut")

    with cR:
        st.markdown("##### Estimativa de potência instalada")
        st.caption("Potência ≈ EE da fonte ÷ (8760 · FC). Selecione o fator de capacidade de referência "
                  "(da tabela econômica) para cada tecnologia.")

        # seleção de FC por tecnologia
        fc_sel = {}
        gsel = st.columns(2)
        techs = list(MATRIZ_2025.keys())
        for idx, tech in enumerate(techs):
            col = gsel[idx % 2]
            opt = FC_OPCOES[tech]
            esc = col.selectbox(f"FC · {tech}", opt, index=0, key=f"fc_{tech}")
            fc_sel[tech] = fonte_lookup(esc)["fc"]

        # tabela de resultados
        linhas, p_total = [], 0.0
        for tech in techs:
            part = shares_n[tech]                       # %
            ee_f = ee_total * part / 100.0              # MWh/ano
            fc   = fc_sel[tech]                         # %
            p    = ee_f / (8760.0 * fc / 100.0) if fc > 0 else 0.0   # MW
            p_total += p
            linhas.append({
                "Fonte": tech,
                "Participação (%)": part,
                "EE (MWh/ano)": ee_f,
                "FC adotado (%)": fc,
                "Pot. instalada (MW)": p,
            })
        df_pot = pd.DataFrame(linhas)

        sty = (df_pot.style
               .format({"Participação (%)": "{:.1f}", "EE (MWh/ano)": "{:,.0f}",
                        "FC adotado (%)": "{:.0f}", "Pot. instalada (MW)": "{:,.1f}"})
               .set_properties(**{"font-size": "12px"})
               .hide(axis="index"))
        st.dataframe(sty, use_container_width=True)

        cc1, cc2 = st.columns(2)
        cc1.markdown(kpi_card("Potência instalada total", _fmt(p_total, 1, " MW"),
                             "soma das fontes (estimativa)", ACCENT_D), unsafe_allow_html=True)
        # fator de capacidade médio do sistema
        fc_medio = (ee_total / (8760.0 * p_total) * 100.0) if p_total > 0 else 0.0
        cc2.markdown(kpi_card("FC médio do sistema", _fmt(fc_medio, 1, " %"),
                             "ponderado pela geração", "#047857"), unsafe_allow_html=True)

    _gap(6)

    # ── barras de potência instalada por fonte ────────────────────────
    fig2 = base_fig("Potência instalada estimada por fonte (MW)", height=300)
    fig2.add_trace(go.Bar(
        x=df_pot["Fonte"], y=df_pot["Pot. instalada (MW)"],
        marker_color=[MATRIZ_COR[t] for t in df_pot["Fonte"]],
        text=[_fmt(v, 1) for v in df_pot["Pot. instalada (MW)"]],
        textposition="outside",
        hovertemplate="%{x}<br>%{y:,.1f} MW<extra></extra>",
    ))
    fig2.update_layout(showlegend=False, yaxis_title="MW")
    st.plotly_chart(fig2, use_container_width=True, key="mz_bar_pot")

    st.info("As potências são **estimativas de equivalência energética** — assumem que cada fonte opera "
           "exatamente no fator de capacidade selecionado. Servem para dimensionar a ordem de grandeza do "
           "parque atual, não substituem o cadastro real de usinas.")


# =======================================================================
#  ABA 2 · CONFIGURAÇÃO ECONÔMICA (tabela-modelo)
# =======================================================================
def tab_config_economico():
    section_title("Configuração Econômica",
                 "Parâmetros técnico-econômicos por fonte — base de todo o cálculo de VPL e tarifa")

    # legenda das fórmulas
    st.markdown(
        f'<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:14px;'
        f'padding:14px 18px;margin-bottom:14px;">'
        f'<div style="font-size:13px;font-weight:700;color:#1d4ed8;margin-bottom:8px;">'
        f'📐 Modelo de custos</div>'
        f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px 24px;font-size:12.5px;color:{TEXT_PRI};">'
        f'<div><code>CAPEX = A · (P·1000)<sup>B</sup></code> &nbsp;<span style="color:{TEXT_SEC};">— P em MW; resultado em US$</span></div>'
        f'<div><code>EE = P · 8760 · (FC/100)</code> &nbsp;<span style="color:{TEXT_SEC};">— energia anual em MWh</span></div>'
        f'<div><code>OPEX_F = (OPEX_FIXO%/100) · CAPEX</code> &nbsp;<span style="color:{TEXT_SEC};">— US$/ano</span></div>'
        f'<div><code>OPEX_V = OPEX_VAR · EE</code> &nbsp;<span style="color:{TEXT_SEC};">— US$/ano (OPEX_VAR em US$/MWh)</span></div>'
        f'</div>'
        f'<div style="font-size:12px;color:{TEXT_SEC};margin-top:8px;">'
        f'WACC = <b>{_fmt(WACC_PADRAO,0)} % a.a.</b> · sem realimentação · A em US$ · B adimensional · FC em %.</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    df = fontes_df()
    df_show = df.rename(columns={
        "Limite": "Limite (MW)", "A": "A (US$)", "B": "B",
        "FC": "FC (%)", "OPEX_FIXO": "OPEX fixo (%CAPEX/ano)",
        "OPEX_VAR": "OPEX var. (US$/MWh)",
    })[["Fonte", "Grupo", "Limite (MW)", "A (US$)", "B", "FC (%)",
        "OPEX fixo (%CAPEX/ano)", "OPEX var. (US$/MWh)"]]

    grp_cor = {"Térmica": "#fee2e2", "Solar": "#fef3c7", "Eólica": "#d1fae5", "Hídrica": "#e0f2fe"}
    grp_txt = {"Térmica": "#b91c1c", "Solar": "#b45309", "Eólica": "#047857", "Hídrica": "#0369a1"}

    def _row_style(row):
        bg = grp_cor.get(row["Grupo"], "#ffffff")
        return [f"background-color:{bg}"] * len(row)

    sty = (df_show.style
           .apply(_row_style, axis=1)
           .format({"A (US$)": "{:,.0f}", "B": "{:.2f}", "FC (%)": "{:.0f}",
                    "OPEX fixo (%CAPEX/ano)": "{:.1f}", "OPEX var. (US$/MWh)": "{:.0f}"})
           .set_properties(**{"font-size": "13px"})
           .set_properties(subset=["Fonte"], **{"font-weight": "700"})
           .hide(axis="index"))
    st.dataframe(sty, use_container_width=True, height=430)

    # chips-resumo por grupo
    _gap(6)
    cols = st.columns(4)
    for col, g in zip(cols, ["Hídrica", "Eólica", "Solar", "Térmica"]):
        sub = df[df["Grupo"] == g]
        fc_min, fc_max = sub["FC"].min(), sub["FC"].max()
        a_min, a_max = sub["A"].min(), sub["A"].max()
        col.markdown(
            f'<div style="background:{grp_cor[g]};border:1px solid #e2e8f0;border-radius:12px;padding:12px 14px;">'
            f'<div style="font-size:12px;font-weight:700;color:{grp_txt[g]};">{g}</div>'
            f'<div style="font-size:11px;color:{TEXT_SEC};margin-top:6px;">FC: {_fmt(fc_min,0)}–{_fmt(fc_max,0)} %</div>'
            f'<div style="font-size:11px;color:{TEXT_SEC};">A: {_fmt(a_min,0)}–{_fmt(a_max,0)} US$</div>'
            f'<div style="font-size:11px;color:{TEXT_SEC};">{len(sub)} tecnologia(s)</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.caption("As fontes renováveis (hídrica, eólica e solar) têm OPEX variável nulo — não há custo de "
              "combustível. Esta tabela é a **entrada fixa** consumida pela aba de Análise Econômica.")


# =======================================================================
#  ABA 3 · ANÁLISE ECONÔMICA (VPL + TARIFA)
# =======================================================================
def tab_analise_economica():
    section_title("Análise Econômica · VPL & Tarifa",
                 "Custo presente de implantar e operar uma usina até 2035 e a tarifa (LCOE) resultante")

    df = fontes_df()

    # ── Entradas ──────────────────────────────────────────────────────
    st.markdown(f'<div style="font-size:11px;font-weight:700;letter-spacing:0.08em;'
               f'text-transform:uppercase;color:{TEXT_SEC};margin-bottom:6px;">Entradas variáveis</div>',
               unsafe_allow_html=True)
    i1, i2, i3, i4 = st.columns([1.4, 1, 1, 1])

    with i1:
        fonte = st.selectbox("Fonte", df["Fonte"].tolist(), index=6, key="ae_fonte")  # default Eólica onshore
    f = fonte_lookup(fonte)

    with i2:
        potencia = st.number_input("Potência (MW)", min_value=0.1, value=100.0, step=10.0,
                                  format="%.1f", key="ae_pot")
    with i3:
        ano_entrada = st.selectbox("Ano de entrada", list(range(2025, 2036)), index=0, key="ae_ano")
        delta_op = 2035 - int(ano_entrada)
    with i4:
        lfsp = st.number_input("Vida útil / LFSP (anos)", min_value=1, value=20, step=1, key="ae_lfsp")

    j1, j2, j3 = st.columns([1, 1, 2])
    with j1:
        wacc = st.number_input("WACC (% a.a.)", min_value=0.0, value=float(WACC_PADRAO),
                              step=0.5, format="%.1f", key="ae_wacc")
    with j2:
        st.markdown(kpi_card("Período de operação", f"{ano_entrada}–2035",
                            f"{delta_op + 1} anos descontados (t = 0…{delta_op})", ACCENT_D),
                    unsafe_allow_html=True)
    with j3:
        st.markdown(
            f'<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;'
            f'padding:10px 14px;height:100%;display:flex;flex-direction:column;justify-content:center;">'
            f'<div style="font-size:11px;color:{TEXT_SEC};font-weight:600;">Parâmetros da fonte selecionada</div>'
            f'<div style="font-size:12.5px;color:{TEXT_PRI};margin-top:4px;">'
            f'A = <b>{_fmt(f["A"],0)}</b> · B = <b>{f["B"]:.2f}</b> · FC = <b>{_fmt(f["fc"],0)} %</b> · '
            f'OPEX_F = <b>{_fmt(f["opex_fix"],1)} %</b> · OPEX_V = <b>{_fmt(f["opex_var"],0)} US$/MWh</b> · '
            f'Limite: <b>{f["limite"]}</b></div></div>',
            unsafe_allow_html=True,
        )

    # alerta leve de limite
    _limite_warn(potencia, f["limite"])

    # ── Cálculo ───────────────────────────────────────────────────────
    R = compute_economics(f["A"], f["B"], f["fc"], f["opex_fix"], f["opex_var"],
                         potencia, delta_op, lfsp, wacc)

    st.markdown("---")

    # ── Grandezas anuais ──────────────────────────────────────────────
    st.markdown("##### Grandezas anuais")
    a1, a2, a3, a4 = st.columns(4)
    a1.markdown(kpi_card("Energia anual", _fmt(R["ee_anual"], 0, " MWh"),
                        f"{_fmt(R['ee_anual']/1000, 3)} GWh/ano", ACCENT), unsafe_allow_html=True)
    a2.markdown(kpi_card("CAPEX total", _us(R["capex"]),
                        f"{_us(R['capex']/potencia/1000, 0)}/kW", C_CAPEX), unsafe_allow_html=True)
    a3.markdown(kpi_card("OPEX fixo / ano", _us(R["opex_f_anual"]), "manutenção", C_OPEXF),
                unsafe_allow_html=True)
    a4.markdown(kpi_card("OPEX variável / ano", _us(R["opex_v_anual"]), "combustível", C_OPEXV),
                unsafe_allow_html=True)

    _gap(8)
    b1, b2, b3 = st.columns(3)
    b1.markdown(kpi_card("Parcela anual do CAPEX", _us(R["pmt_capex"]),
                        f"amortização em {lfsp} anos @ {_fmt(wacc,1)} %", C_CAPEX), unsafe_allow_html=True)
    b2.markdown(kpi_card("Desembolso anual total", _us(R["pmt_capex"] + R["opex_f_anual"] + R["opex_v_anual"]),
                        "parcela CAPEX + OPEX (ano cheio)", TEXT_PRI), unsafe_allow_html=True)
    b3.markdown(kpi_card("Energia descontada (VPL)", _fmt(R["npv_energia"], 0, " MWh"),
                        f"anos 0…{delta_op} ao WACC", "#047857"), unsafe_allow_html=True)

    st.markdown("---")

    # ── Resultado: VPL do custo + Tarifa ──────────────────────────────
    cVPL, cTAR = st.columns([1.25, 1], gap="large")

    with cVPL:
        st.markdown("##### Valor Presente Líquido do custo")
        v1, v2, v3 = st.columns(3)
        v1.markdown(kpi_card("VPL CAPEX", _us(R["npv_capex"]),
                            f"{_fmt(R['npv_capex']/R['npv_total']*100,1)} % do total", C_CAPEX),
                    unsafe_allow_html=True)
        v2.markdown(kpi_card("VPL OPEX fixo", _us(R["npv_opex_f"]),
                            f"{_fmt(R['npv_opex_f']/R['npv_total']*100,1)} % do total", C_OPEXF),
                    unsafe_allow_html=True)
        v3.markdown(kpi_card("VPL OPEX variável", _us(R["npv_opex_v"]),
                            f"{_fmt(R['npv_opex_v']/R['npv_total']*100,1)} % do total", C_OPEXV),
                    unsafe_allow_html=True)
        _gap(8)

        # barra empilhada da composição do VPL
        fig = base_fig("Composição do VPL do custo (US$)", height=240)
        fig.add_trace(go.Bar(y=["VPL"], x=[R["npv_capex"]], name="CAPEX", orientation="h",
                            marker_color=C_CAPEX, hovertemplate="CAPEX<br>%{x:,.0f} US$<extra></extra>"))
        fig.add_trace(go.Bar(y=["VPL"], x=[R["npv_opex_f"]], name="OPEX fixo", orientation="h",
                            marker_color=C_OPEXF, hovertemplate="OPEX fixo<br>%{x:,.0f} US$<extra></extra>"))
        fig.add_trace(go.Bar(y=["VPL"], x=[R["npv_opex_v"]], name="OPEX var.", orientation="h",
                            marker_color=C_OPEXV, hovertemplate="OPEX var.<br>%{x:,.0f} US$<extra></extra>"))
        fig.update_layout(barmode="stack", height=200, xaxis_title="US$",
                          yaxis=dict(showticklabels=False),
                          margin=dict(l=10, r=10, t=20, b=30))
        st.plotly_chart(fig, use_container_width=True, key="ae_vpl_stack")

    with cTAR:
        st.markdown("##### Tarifa (LCOE)")
        # card grande da tarifa
        st.markdown(
            f'<div style="background:linear-gradient(135deg,{ACCENT},{ACCENT_D});border-radius:18px;'
            f'padding:22px 24px;color:#fff;box-shadow:0 8px 28px rgba(14,165,233,0.30);">'
            f'<div style="font-size:12px;font-weight:600;opacity:0.92;text-transform:uppercase;'
            f'letter-spacing:0.06em;">Custo nivelado da energia</div>'
            f'<div style="font-size:38px;font-weight:800;letter-spacing:-1px;margin:6px 0 0;">'
            f'{_fmt(R["lcoe"],2)} <span style="font-size:18px;font-weight:600;">US$/MWh</span></div>'
            f'<div style="font-size:14px;font-weight:600;opacity:0.92;margin-top:2px;">'
            f'{_fmt(R["lcoe"]/1000,4)} US$/kWh</div>'
            f'<div style="font-size:11.5px;opacity:0.85;margin-top:12px;line-height:1.5;">'
            f'VPL do custo ÷ VPL da energia. É a tarifa que, cobrada por MWh ao longo da operação, '
            f'recupera exatamente o custo a valor presente.</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        _gap(10)
        # decomposição da tarifa
        fig2 = base_fig("Composição da tarifa (US$/MWh)", height=210)
        fig2.add_trace(go.Bar(
            x=["CAPEX", "OPEX fixo", "OPEX var."],
            y=[R["tar_capex"], R["tar_opex_f"], R["tar_opex_v"]],
            marker_color=[C_CAPEX, C_OPEXF, C_OPEXV],
            text=[_fmt(R["tar_capex"], 2), _fmt(R["tar_opex_f"], 2), _fmt(R["tar_opex_v"], 2)],
            textposition="outside",
            hovertemplate="%{x}<br>%{y:,.2f} US$/MWh<extra></extra>",
        ))
        fig2.update_layout(showlegend=False, height=210, yaxis_title="US$/MWh",
                          margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig2, use_container_width=True, key="ae_tar_decomp")

    st.markdown("---")

    # ── Fluxo de caixa ano a ano (espelho da Sheet3) ──────────────────
    with st.expander("📋 Fluxo de caixa descontado, ano a ano (espelho da planilha · Sheet3)", expanded=False):
        anos = R["anos"]
        df_fc = pd.DataFrame({
            "Ano (t)": anos,
            "Ano calendário": [int(ano_entrada) + t for t in anos],
            "Fator desconto": R["disc"],
            "CAPEX desc. (US$)": R["serie_capex"],
            "OPEX fixo desc. (US$)": R["serie_opex_f"],
            "OPEX var. desc. (US$)": R["serie_opex_v"],
            "Energia desc. (MWh)": R["serie_energia"],
        })
        df_fc["Custo total desc. (US$)"] = (df_fc["CAPEX desc. (US$)"]
                                            + df_fc["OPEX fixo desc. (US$)"]
                                            + df_fc["OPEX var. desc. (US$)"])
        # linha de totais (VPL)
        tot = {
            "Ano (t)": "VPL", "Ano calendário": "—", "Fator desconto": np.nan,
            "CAPEX desc. (US$)": R["npv_capex"], "OPEX fixo desc. (US$)": R["npv_opex_f"],
            "OPEX var. desc. (US$)": R["npv_opex_v"], "Energia desc. (MWh)": R["npv_energia"],
            "Custo total desc. (US$)": R["npv_total"],
        }
        df_fc = pd.concat([df_fc, pd.DataFrame([tot])], ignore_index=True)

        def _bold_last(row):
            return ["font-weight:700;background-color:#eff6ff" if row["Ano (t)"] == "VPL"
                    else "" for _ in row]

        sty = (df_fc.style
               .apply(_bold_last, axis=1)
               .format({"Fator desconto": "{:.4f}", "CAPEX desc. (US$)": "{:,.0f}",
                        "OPEX fixo desc. (US$)": "{:,.0f}", "OPEX var. desc. (US$)": "{:,.0f}",
                        "Energia desc. (MWh)": "{:,.0f}", "Custo total desc. (US$)": "{:,.0f}"},
                       na_rep="—")
               .set_properties(**{"font-size": "12px"})
               .hide(axis="index"))
        st.dataframe(sty, use_container_width=True)

        st.caption(f"VPL total do custo = **{_us(R['npv_total'])}** · VPL da energia = "
                  f"**{_fmt(R['npv_energia'],0)} MWh** · Tarifa = VPL custo ÷ VPL energia = "
                  f"**{_fmt(R['lcoe'],2)} US$/MWh**. CAPEX amortizado como anuidade de "
                  f"{lfsp} anos, mas só as {delta_op + 1} parcelas até 2035 entram no VPL.")

        # download CSV
        csv = df_fc.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Baixar fluxo de caixa (CSV)", csv,
                          file_name=f"fluxo_caixa_{fonte.split(' ')[0].lower()}_{potencia:.0f}MW.csv",
                          mime="text/csv", key="ae_dl")


def _limite_warn(potencia: float, limite: str):
    """Aviso leve se a potência viola o limite declarado da fonte."""
    try:
        lim = limite.replace(" ", "")
        if lim in ("—", ""):
            return
        if lim.startswith("≤"):
            v = float(lim[1:].replace("MW", ""))
            if potencia > v:
                st.warning(f"⚠️ Potência {potencia:.0f} MW acima do limite típico desta fonte ({limite}).")
        elif lim.startswith("≥"):
            v = float(lim[1:].replace("MW", ""))
            if potencia < v:
                st.warning(f"⚠️ Potência {potencia:.0f} MW abaixo do limite típico desta fonte ({limite}).")
        elif lim.startswith(">"):
            v = float(lim[1:].replace("MW", ""))
            if potencia <= v:
                st.warning(f"⚠️ Potência {potencia:.0f} MW abaixo do limite típico desta fonte ({limite}).")
        elif "–" in lim:
            lo, hi = lim.replace("MW", "").split("–")
            if not (float(lo) <= potencia <= float(hi)):
                st.warning(f"⚠️ Potência {potencia:.0f} MW fora da faixa típica desta fonte ({limite}).")
    except Exception:
        pass


# =======================================================================
#  ENTRADA — chamado pelo app central (app_utopia.py)
# =======================================================================
def run_matriz(page=None):
    st.markdown(
        f'<h1 style="font-size:32px;font-weight:800;color:{TEXT_PRI};letter-spacing:-0.6px;'
        f'margin-bottom:2px;">Matriz Energética · UTÓPIA</h1>'
        f'<p style="font-size:14px;color:{TEXT_SEC};margin-bottom:18px;">'
        f'Plano Decenal de Expansão (2025 → 2035) — análise econômica e tarifária</p>',
        unsafe_allow_html=True,
    )

    t_atual, t_cfg, t_econ = st.tabs(
        ["📅 Atualidade (2025)", "🧮 Configuração Econômica", "💰 Análise Econômica · VPL & Tarifa"]
    )
    with t_atual:
        tab_atualidade()
    with t_cfg:
        tab_config_economico()
    with t_econ:
        tab_analise_economica()


if __name__ == "__main__":
    run_matriz()