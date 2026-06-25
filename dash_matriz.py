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


def formula_box(titulo: str, corpo_html: str, cor: str = ACCENT):
    """Caixa clara para explicar uma fórmula/conceito dentro do dash."""
    st.markdown(
        f'<div style="background:#f8fafc;border:1px solid #e2e8f0;border-left:4px solid {cor};'
        f'border-radius:10px;padding:12px 16px;margin:4px 0;">'
        f'<div style="font-size:11px;font-weight:700;letter-spacing:0.06em;text-transform:uppercase;'
        f'color:{cor};margin-bottom:6px;">{titulo}</div>'
        f'<div style="font-size:13px;color:{TEXT_PRI};line-height:1.65;">{corpo_html}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


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

    # cronograma COMPLETO de amortização do CAPEX (toda a vida útil = LFSP anos).
    # A parcela é a mesma todo ano; o plano só "conta" as que caem até 2035.
    anos_full = list(range(0, int(lfsp)))
    disc_full = [1.0 / (1 + w) ** t for t in anos_full]
    serie_capex_nom_full  = [pmt_capex for _ in anos_full]                  # nominal (constante)
    serie_capex_disc_full = [pmt_capex * d for d in disc_full]              # trazido ao presente

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
        # meta + cronograma completo (vida útil = LFSP)
        wacc_frac=w, lfsp=int(lfsp), delta_op=int(delta_op),
        anos_full=anos_full, disc_full=disc_full,
        serie_capex_nom_full=serie_capex_nom_full,
        serie_capex_disc_full=serie_capex_disc_full,
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
    e1.markdown(kpi_card("Em TWh", _fmt(ee_total / 1e6, 4), "terawatt-hora"), unsafe_allow_html=True)
    e2.markdown(kpi_card("Em GWh", _fmt(ee_total / 1e3, 1), "gigawatt-hora"), unsafe_allow_html=True)
    e3.markdown(kpi_card("Em MWh", _fmt(ee_total, 1), "megawatt-hora"), unsafe_allow_html=True)
    e4.markdown(kpi_card("Em kWh", _fmt(ee_total * 1e3, 1), "quilowatt-hora", ACCENT_D),
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

        formula_box(
            "Como a energia vira potência instalada",
            f'A energia gerada por uma usina em um ano é '
            f'<span style="font-family:monospace;color:{ACCENT_D};">EE = P &times; 8760 h &times; FC</span>. '
            f'Como aqui conhecemos a energia de cada fonte (a fatia da matriz) e queremos a potência, '
            f'invertemos a fórmula:<br>'
            f'<span style="font-family:monospace;font-size:13.5px;color:{ACCENT_D};">'
            f'P = EE<sub>fonte</sub> &divide; (8760 h &times; FC)</span>.<br>'
            f'<span style="color:{TEXT_SEC};font-size:12px;">O FC (fator de capacidade) é a fração média do '
            f'tempo em que a usina gera à potência plena — hídricas e térmicas têm FC alto; solar e eólica, '
            f'mais baixo. Por isso, para a mesma energia, uma fonte de FC baixo precisa de mais potência instalada.</span>',
            ACCENT,
        )

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

    formula_box(
        "Como se calcula a energia gerada",
        f'A energia elétrica gerada por ano é a potência instalada multiplicada pelas horas do ano '
        f'e pelo <b>fator de capacidade</b> (FC) — a fração do tempo que a usina gera na média:<br>'
        f'<span style="font-family:monospace;font-size:13.5px;color:{ACCENT_D};">'
        f'EE<sub>ano</sub> = P &times; 8760 h &times; FC</span> &nbsp;=&nbsp; '
        f'<b>{_fmt(potencia,1)} MW</b> &times; 8760 h &times; <b>{_fmt(f["fc"],0)} %</b> '
        f'= <b>{_fmt(R["ee_anual"],0)} MWh/ano</b>.<br>'
        f'<span style="color:{TEXT_SEC};font-size:12px;">O FC é o da fonte selecionada (tabela da aba '
        f'<i>Config Econômico</i>). Um FC de {_fmt(f["fc"],0)} % equivale a gerar à potência plena durante '
        f'{_fmt(8760*f["fc"]/100,0)} h por ano.</span>',
        ACCENT,
    )

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

    _gap(8)
    formula_box(
        f"CAPEX: distribuído em {lfsp} anos, contado só até 2035",
        f'O CAPEX total (<b>{_us(R["capex"])}</b>) <b>não</b> é pago à vista nem amortizado no prazo do plano. '
        f'Ele é distribuído como uma anuidade ao longo de toda a vida útil — <b>{lfsp} anos</b> — gerando uma '
        f'parcela constante de <b>{_us(R["pmt_capex"])}/ano</b>:<br>'
        f'<span style="font-family:monospace;font-size:13.5px;color:{C_CAPEX};">'
        f'parcela = CAPEX &divide; FA({lfsp}; {_fmt(wacc,1)} %)</span>, onde '
        f'<span style="font-family:monospace;">FA(n;i) = [(1+i)<sup>n</sup>&minus;1] &divide; [i&middot;(1+i)<sup>n</sup>]</span>.<br>'
        f'Entrar em <b>{ano_entrada}</b> não significa amortizar o CAPEX em {delta_op + 1} anos: significa que, '
        f'das {lfsp} parcelas dessa amortização, <b>só as {delta_op + 1} que caem entre {ano_entrada} e 2035</b> '
        f'entram neste plano (e no VPL). As demais parcelas seguem sendo pagas pela usina depois de 2035, '
        f'fora do horizonte do plano. Veja isso nas figuras de fluxo de caixa abaixo. 👇',
        C_CAPEX,
    )

    st.markdown("---")

    # ── Resultado: VPL do custo + Tarifa ──────────────────────────────
    cVPL, cTAR = st.columns([1.25, 1], gap="large")

    with cVPL:
        st.markdown("##### Valor Presente Líquido do custo")
        v1, v2, v3 = st.columns(3)
        custo_nom = R["pmt_capex"] + R["opex_f_anual"] + R["opex_v_anual"]  # desembolso nominal/ano
        v1.markdown(kpi_card("VPL CAPEX", _us(R["npv_capex"]),
                            f"{_fmt(R['pmt_capex']/custo_nom*100,1)} % do desembolso anual", C_CAPEX),
                    unsafe_allow_html=True)
        v2.markdown(kpi_card("VPL OPEX fixo", _us(R["npv_opex_f"]),
                            f"{_fmt(R['opex_f_anual']/custo_nom*100,1)} % do desembolso anual", C_OPEXF),
                    unsafe_allow_html=True)
        v3.markdown(kpi_card("VPL OPEX variável", _us(R["npv_opex_v"]),
                            f"{_fmt(R['opex_v_anual']/custo_nom*100,1)} % do desembolso anual", C_OPEXV),
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

        # ── equação LCOE como SVG inline ──────────────────────────────
        lcoe_svg = """<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;
padding:12px 16px;margin-bottom:10px;text-align:center;">
<svg viewBox="0 0 340 72" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-width:340px;">
  <!-- LCOE = -->
  <text x="4" y="40" font-family="Georgia,serif" font-size="15" fill="#0f172a" font-style="italic">LCOE</text>
  <text x="52" y="40" font-family="Georgia,serif" font-size="15" fill="#0f172a">=</text>
  <!-- numerador -->
  <text x="72" y="22" font-family="Georgia,serif" font-size="11" fill="#0284c7">
    <tspan font-size="13">&#x2211;</tspan>
    <tspan baseline-shift="super" font-size="9">n</tspan>
  </text>
  <text x="88" y="22" font-family="Georgia,serif" font-size="11" fill="#0284c7">
    <tspan font-style="italic">C</tspan><tspan baseline-shift="sub" font-size="9">t</tspan>
  </text>
  <text x="104" y="22" font-family="Georgia,serif" font-size="11" fill="#0284c7"> / (1+r)</text>
  <text x="144" y="18" font-family="Georgia,serif" font-size="9" fill="#0284c7">t</text>
  <!-- barra de fracción -->
  <line x1="68" y1="32" x2="200" y2="32" stroke="#64748b" stroke-width="1.2"/>
  <!-- denominador -->
  <text x="72" y="52" font-family="Georgia,serif" font-size="11" fill="#047857">
    <tspan font-size="13">&#x2211;</tspan>
    <tspan baseline-shift="super" font-size="9">n</tspan>
  </text>
  <text x="88" y="52" font-family="Georgia,serif" font-size="11" fill="#047857">
    <tspan font-style="italic">E</tspan><tspan baseline-shift="sub" font-size="9">t</tspan>
  </text>
  <text x="104" y="52" font-family="Georgia,serif" font-size="11" fill="#047857"> / (1+r)</text>
  <text x="144" y="48" font-family="Georgia,serif" font-size="9" fill="#047857">t</text>
  <!-- legenda -->
  <text x="210" y="20" font-family="Arial,sans-serif" font-size="9" fill="#0284c7">C&#x209C; = CAPEX + OPEX</text>
  <text x="210" y="34" font-family="Arial,sans-serif" font-size="9" fill="#047857">E&#x209C; = energia gerada</text>
  <text x="210" y="48" font-family="Arial,sans-serif" font-size="9" fill="#64748b">r = WACC,  t = 0…n</text>
  <text x="210" y="62" font-family="Arial,sans-serif" font-size="9" fill="#64748b">n = vida útil (LFSP)</text>
</svg>
<div style="font-size:10.5px;color:#64748b;margin-top:4px;">Fonte: IRENA · IEA · literatura técnica</div>
</div>"""
        st.markdown(lcoe_svg, unsafe_allow_html=True)

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

    # ── Fluxo de caixa em figuras ─────────────────────────────────────
    st.markdown("##### Fluxo de caixa")
    st.caption("O desembolso ano a ano (CAPEX distribuído na vida útil + OPEX), os mesmos valores "
               "trazidos ao presente e a formação acumulada do VPL.")

    ano0      = int(ano_entrada)
    anos_full = R["anos_full"]                       # 0 … LFSP-1
    cal_full  = [ano0 + t for t in anos_full]
    dop       = R["delta_op"]
    pmt, opf, opv = R["pmt_capex"], R["opex_f_anual"], R["opex_v_anual"]
    x_lim     = 2035.5                               # fronteira do horizonte do plano
    GRIS      = "#cbd5e1"

    # máscaras dentro / fora do horizonte (até 2035)
    capex_in  = [pmt if t <= dop else 0 for t in anos_full]
    capex_out = [pmt if t >  dop else 0 for t in anos_full]
    opexf_in  = [opf if t <= dop else 0 for t in anos_full]
    opexv_in  = [opv if t <= dop else 0 for t in anos_full]

    # ---- Figura 1 · fluxo de caixa NOMINAL (vida útil inteira) -------------
    st.markdown(f'<div style="font-size:13px;font-weight:600;color:{TEXT_PRI};margin:6px 0 2px;">'
                f'1 · Fluxo de caixa nominal — desembolso a cada ano</div>', unsafe_allow_html=True)
    fig_nom = base_fig("", height=340)
    fig_nom.add_trace(go.Bar(x=cal_full, y=capex_in, name="Parcela CAPEX (no plano)",
                             marker_color=C_CAPEX,
                             hovertemplate="%{x}<br>Parcela CAPEX: %{y:,.0f} US$<extra></extra>"))
    fig_nom.add_trace(go.Bar(x=cal_full, y=opexf_in, name="OPEX fixo",
                             marker_color=C_OPEXF,
                             hovertemplate="%{x}<br>OPEX fixo: %{y:,.0f} US$<extra></extra>"))
    fig_nom.add_trace(go.Bar(x=cal_full, y=opexv_in, name="OPEX variável",
                             marker_color=C_OPEXV,
                             hovertemplate="%{x}<br>OPEX var.: %{y:,.0f} US$<extra></extra>"))
    fig_nom.add_trace(go.Bar(x=cal_full, y=capex_out, name="Parcela CAPEX (além de 2035)",
                             marker_color=GRIS, marker_line_width=0,
                             hovertemplate="%{x}<br>Parcela CAPEX (fora do plano): %{y:,.0f} US$<extra></extra>"))
    fig_nom.update_layout(barmode="stack", height=340, yaxis_title="US$/ano",
                          xaxis=dict(title="ano", dtick=1),
                          legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                          margin=dict(l=10, r=10, t=10, b=30))
    fig_nom.add_vline(x=x_lim, line_dash="dot", line_color=TEXT_SEC, line_width=1.5)
    fig_nom.add_annotation(x=x_lim, y=1.0, yref="paper", text="fim do plano (2035)",
                           showarrow=False, font=dict(size=10, color=TEXT_SEC),
                           xanchor="left", xshift=4)
    st.plotly_chart(fig_nom, use_container_width=True, key="ae_cf_nom")
    st.caption(f"A parcela do CAPEX (**{_us(pmt)}/ano**) se repete por toda a vida útil de **{lfsp} anos**. "
               f"As barras coloridas (até 2035) entram no plano e no VPL; as **cinzas** são parcelas que a "
               f"usina segue pagando depois de 2035, fora do horizonte.")
    _gap(6)

    # ---- Figura 2 · valores TRAZIDOS AO PRESENTE -------------------------
    st.markdown(f'<div style="font-size:13px;font-weight:600;color:{TEXT_PRI};margin:6px 0 2px;">'
                f'2 · Valores trazidos ao presente — fluxo descontado ao WACC</div>',
                unsafe_allow_html=True)
    cal_in = [ano0 + t for t in R["anos"]]
    nom_tot = pmt + opf + opv
    fig_pv = base_fig("", height=320)
    fig_pv.add_trace(go.Bar(x=cal_in, y=R["serie_capex"], name="CAPEX (VP)", marker_color=C_CAPEX,
                            hovertemplate="%{x}<br>CAPEX VP: %{y:,.0f} US$<extra></extra>"))
    fig_pv.add_trace(go.Bar(x=cal_in, y=R["serie_opex_f"], name="OPEX fixo (VP)", marker_color=C_OPEXF,
                            hovertemplate="%{x}<br>OPEX fixo VP: %{y:,.0f} US$<extra></extra>"))
    fig_pv.add_trace(go.Bar(x=cal_in, y=R["serie_opex_v"], name="OPEX var. (VP)", marker_color=C_OPEXV,
                            hovertemplate="%{x}<br>OPEX var. VP: %{y:,.0f} US$<extra></extra>"))
    fig_pv.add_trace(go.Scatter(x=cal_in, y=[nom_tot] * len(cal_in), name="desembolso nominal",
                                mode="lines", line=dict(color=TEXT_SEC, dash="dash", width=1.5),
                                hovertemplate="%{x}<br>nominal: %{y:,.0f} US$<extra></extra>"))
    fig_pv.update_layout(barmode="stack", height=320, yaxis_title="US$ (valor presente)",
                         xaxis=dict(title="ano", dtick=1),
                         legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                         margin=dict(l=10, r=10, t=10, b=30))
    st.plotly_chart(fig_pv, use_container_width=True, key="ae_cf_pv")
    st.caption(f"Cada barra é o desembolso daquele ano dividido por (1+WACC)^t. A linha tracejada é o "
               f"desembolso nominal constante (**{_us(nom_tot)}/ano**): quanto mais distante o ano, mais o "
               f"desconto encolhe a barra. A soma de todas as barras é o **VPL do custo = {_us(R['npv_total'])}**.")
    _gap(6)

    # ---- Figura 3 · DISTRIBUIÇÃO acumulada do VPL ------------------------
    st.markdown(f'<div style="font-size:13px;font-weight:600;color:{TEXT_PRI};margin:6px 0 2px;">'
                f'3 · Distribuição acumulada — como o VPL se forma ano a ano</div>',
                unsafe_allow_html=True)
    tot_pv_year = [c + ff + vv for c, ff, vv in
                   zip(R["serie_capex"], R["serie_opex_f"], R["serie_opex_v"])]
    acum = list(np.cumsum(tot_pv_year))
    fig_cum = base_fig("", height=300)
    fig_cum.add_trace(go.Bar(x=cal_in, y=tot_pv_year, name="custo do ano (VP)",
                             marker_color="rgba(2,132,199,0.30)",
                             hovertemplate="%{x}<br>custo VP do ano: %{y:,.0f} US$<extra></extra>"))
    fig_cum.add_trace(go.Scatter(x=cal_in, y=acum, name="VPL acumulado", mode="lines+markers",
                                 line=dict(color=ACCENT_D, width=2.5),
                                 marker=dict(size=6, color=ACCENT_D),
                                 fill="tozeroy", fillcolor="rgba(14,165,233,0.12)",
                                 hovertemplate="%{x}<br>acumulado: %{y:,.0f} US$<extra></extra>"))
    fig_cum.add_hline(y=R["npv_total"], line_dash="dot", line_color="#047857", line_width=1.5)
    fig_cum.add_annotation(x=cal_in[0], y=R["npv_total"], text=f"VPL total = {_us(R['npv_total'])}",
                           showarrow=False, font=dict(size=11, color="#047857"),
                           xanchor="left", yanchor="bottom", yshift=2)
    fig_cum.update_layout(height=300, yaxis_title="US$",
                          xaxis=dict(title="ano", dtick=1),
                          legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                          margin=dict(l=10, r=10, t=10, b=30))
    st.plotly_chart(fig_cum, use_container_width=True, key="ae_cf_cum")
    st.caption("A curva mostra o custo a valor presente se somando ano a ano até atingir o VPL total — "
               "é a contribuição de cada ano para o custo do plano.")

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
# =======================================================================
#  PLANO DECENAL DE EXPANSÃO — NÚCLEO + 2 ABAS
# =======================================================================

CSV_PROJ = ROOT / "projecoes_demanda.csv"

# Matriz base 2025 (fixa, conforme aba Atualidade)
MATRIZ_BASE_2025 = {"Hidro": 0.40, "Termo": 0.55, "Eólica": 0.05, "Solar": 0.0}

# Mapeamento tipo → cor (paleta da aba Atualidade)
TIPO_COR = {"Hidro": HYD, "Termo": THR, "Eólica": WIN, "Solar": SOL}
TIPO_ORDEM = ["Hidro", "Termo", "Eólica", "Solar"]


@st.cache_data(ttl=300)
def load_demanda_2025_2035(csv_path, cenario: str = "Referencia",
                           ee_2025_fallback: float = 1_450_000.0) -> dict:
    """
    Retorna dict {ano: EE_total_MWh} de 2025 a 2035.
    - 2025: do Excel histórico (ENTREGA_DEMANDA), com fallback.
    - 2026-2035: do CSV de projeção (cenario escolhido, Local=Utopia).
      cenario ∈ {"Referencia", "Alto", "Baixo"}.
    """
    # taxa de fallback por cenário (a.a.)
    taxa = {"Referencia": 0.025, "Alto": 0.035, "Baixo": 0.015}.get(cenario, 0.025)

    out = {}
    # 2025 do histórico (independe do cenário)
    try:
        socio = load_socio_2025(str(EXCEL_HIST))
        out[2025] = float(socio["ee"])
    except Exception:
        out[2025] = ee_2025_fallback

    # 2026-2035 do CSV
    try:
        df = pd.read_csv(csv_path)
        df["Ano"] = df["Ano"].astype(int)
        sub = df[(df["Cenario"] == cenario) & (df["Local"] == "Utopia")].sort_values("Ano")
        for _, row in sub.iterrows():
            ano = int(row["Ano"])
            if 2026 <= ano <= 2035:
                out[ano] = float(row["EE_TOTAL"])
    except Exception:
        for ano in range(2026, 2036):
            out[ano] = out[2025] * ((1 + taxa) ** (ano - 2025))

    # Garante anos faltantes via crescimento do cenário
    for ano in range(2025, 2036):
        if ano not in out:
            out[ano] = out[2025] * ((1 + taxa) ** (ano - 2025))

    return out


# ── Definição das 5 Propostas ────────────────────────────────────────
PROPOSTAS = {
    1: dict(
        nome="Alta Penetração Solar",
        icon="☀️",
        cor=SOL,
        descricao=("Térmica mantém sua geração de 2025 constante (0,80 TWh). "
                  "Hidro recebe +10 % da energia de 2025 → 1 PCH em 2027. "
                  "Solar FV absorve todo o crescimento residual em 5 etapas. "
                  "Eólica onshore mantém o volume de 2025."),
        tech_hidro="PCH",
        tech_eolica=[],
        tech_solar=["Solar FV"],
        tech_termo=["Gás natural (ciclo comb.)"],
        ano_hidro=2027,
        anos_eolica=[],
        anos_solar=[2026, 2028, 2030, 2032, 2034],
        frac_eolica_nova=0.0,
        frac_solar_nova=1.0,
    ),
    2: dict(
        nome="Alta Penetração Eólica",
        icon="💨",
        cor=WIN,
        descricao=("Térmica mantém sua geração de 2025 constante. "
                  "Hidro recebe +10 % (PCH em 2027). "
                  "Eólica (on+offshore) cobre 70 % do crescimento residual. "
                  "Solar FV cobre os 30 % restantes."),
        tech_hidro="PCH",
        tech_eolica=["Eólica onshore", "Eólica offshore",
                     "Eólica onshore", "Eólica offshore", "Eólica onshore"],
        tech_solar=["Solar FV"],
        tech_termo=["Gás natural (ciclo comb.)"],
        ano_hidro=2027,
        anos_eolica=[2026, 2028, 2030, 2032, 2034],
        anos_solar=[2028, 2032],
        frac_eolica_nova=0.70,
        frac_solar_nova=0.30,
    ),
    3: dict(
        nome="Termo Dominante",
        icon="🔥",
        cor=THR,
        descricao=("Sem expansão hídrica. A térmica EXPANDE cobrindo 80 % do "
                  "crescimento com mix gás natural + biomassa. Eólica e Solar "
                  "cobrem os 20 % restantes."),
        tech_hidro=None,
        tech_eolica=["Eólica onshore"],
        tech_solar=["Solar FV"],
        tech_termo=["Gás natural (ciclo comb.)", "Biomassa (cana/madeira)"],
        ano_hidro=None,
        anos_eolica=[2030],
        anos_solar=[2030],
        frac_eolica_nova=0.10,
        frac_solar_nova=0.10,
    ),
    4: dict(
        nome="Mix Renovável Equilibrado",
        icon="🌱",
        cor="#22c55e",
        descricao=("Térmica mantém geração constante. Hidro recebe +10 % (PCH). "
                  "Crescimento residual dividido 50/50 entre Eólica (on+offshore) "
                  "e Solar FV."),
        tech_hidro="PCH",
        tech_eolica=["Eólica onshore", "Eólica offshore",
                     "Eólica onshore", "Eólica offshore"],
        tech_solar=["Solar FV"],
        tech_termo=["Gás natural (ciclo comb.)"],
        ano_hidro=2027,
        anos_eolica=[2027, 2029, 2031, 2033],
        anos_solar=[2026, 2028, 2030, 2032, 2034],
        frac_eolica_nova=0.50,
        frac_solar_nova=0.50,
    ),
    5: dict(
        nome="Otimizada (mínimo LCOE)",
        icon="🎯",
        cor=ACCENT,
        descricao=("Térmica mantém geração constante. Hidro recebe +10 % (PCH). "
                  "Crescimento residual: 80 % para Eólica onshore (menor LCOE), "
                  "20 % para Solar FV."),
        tech_hidro="PCH",
        tech_eolica=["Eólica onshore"],
        tech_solar=["Solar FV"],
        tech_termo=["Gás natural (ciclo comb.)"],
        ano_hidro=2027,
        anos_eolica=[2026, 2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034],
        anos_solar=[2030, 2033],
        frac_eolica_nova=0.80,
        frac_solar_nova=0.20,
    ),
}


def build_plantas(prop_id: int, demanda: dict) -> list:
    """
    Constrói o portfólio de usinas NOVAS.

    Regra: a térmica de 2025 (0.80 TWh) é CONSTANTE — nunca diminui.
    O CRESCIMENTO (EE2035 - EE2025) é coberto por:
      1. Hidro: +10% da energia hídrica 2025 → 1 PCH
      2. Eólica: frac_eolica_nova × crescimento_residual
      3. Solar:  frac_solar_nova  × crescimento_residual
      4. Térmica nova (P3): quando eólica+solar não chegam (ou para P3 que quer expandir termo)
    """
    p = PROPOSTAS[prop_id]
    plantas = []
    ee_2025 = demanda[2025]
    ee_2035 = demanda[2035]
    ee_base = {k: v * ee_2025 for k, v in MATRIZ_BASE_2025.items()}
    crescimento_total = ee_2035 - ee_2025

    def _add(tipo, fonte, ee_alvo, ano):
        f = fonte_lookup(fonte)
        pot = ee_alvo / (8760.0 * f["fc"] / 100.0)
        plantas.append(dict(tipo=tipo, fonte=fonte, potencia=pot,
                           ano_entrada=int(ano), ee_anual=ee_alvo, fc=f["fc"]))

    # 1) HIDRO: +10% da energia hídrica 2025, 1 PCH
    delta_hidro = 0.0
    if p.get("tech_hidro") and p.get("ano_hidro"):
        delta_hidro = 0.10 * ee_base["Hidro"]
        _add("Hidro", p["tech_hidro"], delta_hidro, p["ano_hidro"])

    # crescimento residual após a hidro
    crescimento_res = crescimento_total - delta_hidro

    # 2) EÓLICA
    delta_eol = crescimento_res * p.get("frac_eolica_nova", 0.0)
    if delta_eol > 1e-3 and p.get("anos_eolica"):
        n = len(p["anos_eolica"])
        ee_por = delta_eol / n
        techs = p["tech_eolica"]
        for i, ano in enumerate(p["anos_eolica"]):
            _add("Eólica", techs[i % len(techs)], ee_por, ano)

    # 3) SOLAR
    delta_sol = crescimento_res * p.get("frac_solar_nova", 0.0)
    if delta_sol > 1e-3 and p.get("anos_solar"):
        n = len(p["anos_solar"])
        ee_por = delta_sol / n
        techs = p["tech_solar"]
        for i, ano in enumerate(p["anos_solar"]):
            _add("Solar", techs[i % len(techs)], ee_por, ano)

    # 4) TÉRMICA NOVA — SOMENTE em P3 (Termo Dominante).
    # Nos outros planos a térmica base é mantida constante; não expande.
    if p.get("tech_termo") and prop_id == 3:
        techs_t = p["tech_termo"]
        cnt = 0
        for ano in range(2026, 2036):
            h = ee_base["Hidro"]; e = ee_base["Eólica"]; s = 0.0; tn = 0.0
            for pl in plantas:
                if pl["ano_entrada"] <= ano:
                    if   pl["tipo"] == "Hidro":  h  += pl["ee_anual"]
                    elif pl["tipo"] == "Eólica": e  += pl["ee_anual"]
                    elif pl["tipo"] == "Solar":  s  += pl["ee_anual"]
                    elif pl["tipo"] == "Termo":  tn += pl["ee_anual"]
            termo_cap = ee_base["Termo"] + tn
            gap = demanda[ano] - (h + e + s + termo_cap)
            if gap > 1e-3:
                _add("Termo", techs_t[cnt % len(techs_t)], gap, ano)
                cnt += 1

    return plantas


def simular_plano(plantas: list, demanda: dict) -> dict:
    """
    Despacho ano a ano — termo BASE é CONSTANTE, nunca modula para baixo.
    Hidro/Eólica/Solar: operam sempre ao ee_anual integral.
    Termo nova (P3): soma-se à base.
    """
    ee_base = {k: v * demanda[2025] for k, v in MATRIZ_BASE_2025.items()}
    anos = list(range(2025, 2036))
    por_tipo = {tipo: [] for tipo in TIPO_ORDEM}
    total, dem, gap, exced, termo_cap_list = [], [], [], [], []

    for ano in anos:
        h = ee_base["Hidro"]; e = ee_base["Eólica"]; s = 0.0; t_novo = 0.0
        for pl in plantas:
            if pl["ano_entrada"] <= ano:
                if   pl["tipo"] == "Hidro":  h      += pl["ee_anual"]
                elif pl["tipo"] == "Eólica": e      += pl["ee_anual"]
                elif pl["tipo"] == "Solar":  s      += pl["ee_anual"]
                elif pl["tipo"] == "Termo":  t_novo += pl["ee_anual"]
        t_total = ee_base["Termo"] + t_novo   # CONSTANTE (+ eventual expansão P3)
        oferta = h + e + s + t_total
        por_tipo["Hidro"].append(h)
        por_tipo["Eólica"].append(e)
        por_tipo["Solar"].append(s)
        por_tipo["Termo"].append(t_total)
        total.append(oferta)
        dem.append(demanda[ano])
        gap.append(demanda[ano] - oferta)
        exced.append(max(0.0, oferta - demanda[ano]))
        termo_cap_list.append(t_total)

    return dict(anos=anos, por_tipo=por_tipo, total=total,
                demanda=dem, gap=gap, excedente=exced, termo_cap=termo_cap_list)


def economics_plano(plantas: list, sim: dict, demanda: dict,
                    wacc_pct: float = 7.0, lfsp: int = 20) -> dict:
    """
    Agrega economics de TODAS as usinas. LCOE do plano = ΣVPL custos / ΣVPL energia.

    Tratamento da TÉRMICA BASE de 2025 (121 MW já existentes):
      • NÃO entra CAPEX (já amortizada antes de 2025).
      • Entra apenas OPEX da energia que efetivamente despacha ano a ano
        (OPEX fixo da capacidade existente + OPEX variável ∝ energia gerada).
      • A energia despachada vem de sim["por_tipo"]["Termo"] (modula com o tempo).
    Térmica NOVA (quando a proposta expande) entra com CAPEX + OPEX normalmente.
    """
    tot = dict(
        capex_total=0.0, opex_f_total=0.0, opex_v_total=0.0,
        npv_capex=0.0, npv_opex_f=0.0, npv_opex_v=0.0, npv_energia=0.0,
        potencia_total=0.0, ee_anual_total=0.0,
        detalhe=[],
        serie_capex=np.zeros(11), serie_opex_f=np.zeros(11),
        serie_opex_v=np.zeros(11), serie_energia=np.zeros(11),
        serie_capex_nom=np.zeros(11), serie_opex_f_nom=np.zeros(11),
        serie_opex_v_nom=np.zeros(11),
    )
    anos_idx = {2025 + i: i for i in range(11)}  # 2025..2035 → 0..10
    w = wacc_pct / 100.0

    for pl in plantas:
        f = fonte_lookup(pl["fonte"])
        delta_op = 2035 - pl["ano_entrada"]
        if delta_op < 0:
            continue
        R = compute_economics(f["A"], f["B"], f["fc"], f["opex_fix"], f["opex_var"],
                             pl["potencia"], delta_op, lfsp, wacc_pct)
        tot["capex_total"]   += R["capex"]
        tot["opex_f_total"]  += R["opex_f_anual"]
        tot["opex_v_total"]  += R["opex_v_anual"]
        tot["npv_capex"]     += R["npv_capex"]
        tot["npv_opex_f"]    += R["npv_opex_f"]
        tot["npv_opex_v"]    += R["npv_opex_v"]
        tot["npv_energia"]   += R["npv_energia"]
        tot["potencia_total"] += pl["potencia"]
        tot["ee_anual_total"] += R["ee_anual"]

        # série anual da usina mapeada no eixo 2025..2035
        for t, (sc, sf, sv, se) in enumerate(zip(R["serie_capex"], R["serie_opex_f"],
                                                  R["serie_opex_v"], R["serie_energia"])):
            ano = pl["ano_entrada"] + t
            if ano in anos_idx:
                idx = anos_idx[ano]
                tot["serie_capex"][idx]   += sc
                tot["serie_opex_f"][idx]  += sf
                tot["serie_opex_v"][idx]  += sv
                tot["serie_energia"][idx] += se
                tot["serie_capex_nom"][idx]  += R["pmt_capex"]
                tot["serie_opex_f_nom"][idx] += R["opex_f_anual"]
                tot["serie_opex_v_nom"][idx] += R["opex_v_anual"]

        tot["detalhe"].append(dict(
            tipo=pl["tipo"], fonte=pl["fonte"], ano_entrada=pl["ano_entrada"],
            potencia=pl["potencia"], ee_anual=R["ee_anual"],
            capex=R["capex"], parcela_capex=R["pmt_capex"],
            opex_f_anual=R["opex_f_anual"], opex_v_anual=R["opex_v_anual"],
            npv_custo=R["npv_total"], npv_energia=R["npv_energia"],
            lcoe=R["lcoe"], delta_op=delta_op,
        ))

    # ─── TÉRMICA BASE 2025 (existente) — só OPEX, sem CAPEX ──────────────
    # Capacidade existente: gás natural, 121 MW (energia base = 55% de 2025).
    ee_termo_base_2025 = MATRIZ_BASE_2025["Termo"] * demanda[2025]
    f_base = fonte_lookup("Gás natural (ciclo comb.)")
    pot_termo_base = ee_termo_base_2025 / (8760.0 * f_base["fc"] / 100.0)
    # OPEX fixo anual da capacidade existente (sobre o CAPEX de referência, não pago)
    capex_ref_base = f_base["A"] * (pot_termo_base * 1000.0) ** f_base["B"]
    opex_f_base_ano = (f_base["opex_fix"] / 100.0) * capex_ref_base

    # energia térmica base despachada ano a ano = min(despacho total, capacidade base)
    npv_opex_f_base = 0.0
    npv_opex_v_base = 0.0
    npv_energia_base = 0.0
    opex_f_base_2035 = 0.0
    opex_v_base_2035 = 0.0
    for i, ano in enumerate(sim["anos"]):
        termo_desp = sim["por_tipo"]["Termo"][i]
        # parte despachada pela base (a base tem prioridade; novas térmicas completam)
        ee_base_desp = min(termo_desp, ee_termo_base_2025)
        opex_v_ano = f_base["opex_var"] * ee_base_desp     # US$/MWh × MWh
        d = 1.0 / (1 + w) ** i                              # desconto (t=0 em 2025)
        npv_opex_f_base += opex_f_base_ano * d
        npv_opex_v_base += opex_v_ano * d
        npv_energia_base += ee_base_desp * d
        # séries nominais/descontadas no eixo 2025..2035
        tot["serie_opex_f_nom"][i] += opex_f_base_ano
        tot["serie_opex_v_nom"][i] += opex_v_ano
        tot["serie_opex_f"][i] += opex_f_base_ano * d
        tot["serie_opex_v"][i] += opex_v_ano * d
        tot["serie_energia"][i] += ee_base_desp * d
        if ano == 2035:
            opex_f_base_2035 = opex_f_base_ano
            opex_v_base_2035 = opex_v_ano

    tot["opex_f_total"]  += opex_f_base_ano       # anual (referência 2035)
    tot["opex_v_total"]  += opex_v_base_2035
    tot["npv_opex_f"]    += npv_opex_f_base
    tot["npv_opex_v"]    += npv_opex_v_base
    tot["npv_energia"]   += npv_energia_base
    tot["ee_anual_total"] += sim["por_tipo"]["Termo"][-1] if False else 0  # já contado no despacho

    tot["detalhe"].append(dict(
        tipo="Termo", fonte="Gás natural (base 2025, só OPEX)", ano_entrada=2025,
        potencia=pot_termo_base, ee_anual=sim["por_tipo"]["Termo"][-1],
        capex=0.0, parcela_capex=0.0,
        opex_f_anual=opex_f_base_ano, opex_v_anual=opex_v_base_2035,
        npv_custo=npv_opex_f_base + npv_opex_v_base, npv_energia=npv_energia_base,
        lcoe=((npv_opex_f_base + npv_opex_v_base) / npv_energia_base
              if npv_energia_base > 0 else 0.0),
        delta_op=10,
    ))

    tot["npv_total"] = tot["npv_capex"] + tot["npv_opex_f"] + tot["npv_opex_v"]
    tot["lcoe_plano"] = (tot["npv_total"] / tot["npv_energia"]
                        if tot["npv_energia"] > 0 else float("nan"))
    return tot


def _matriz_2035_fracoes(sim: dict) -> dict:
    """Fração realizada de cada tipo em 2035."""
    total_2035 = sim["total"][-1]
    if total_2035 <= 0:
        return {t: 0.0 for t in TIPO_ORDEM}
    return {t: sim["por_tipo"][t][-1] / total_2035 for t in TIPO_ORDEM}


# =======================================================================
#  ABA 4 · PLANO DE EXPANSÃO  (selector de proposta + visualizações)
# =======================================================================
def tab_plano_expansao():
    section_title("Plano Decenal de Expansão · 2025 → 2035",
                 "Selecione uma proposta para ver o portfólio de usinas, a oferta "
                 "ano a ano, a evolução da matriz e o LCOE agregado")

    # Seletor de cenário de demanda + proposta
    csel1, csel2 = st.columns([1, 2.4])
    with csel1:
        cen_lbl = st.selectbox(
            "📊 Cenário de demanda",
            options=["Referencia", "Alto", "Baixo"],
            format_func=lambda c: {"Referencia": "Referência",
                                   "Alto": "Otimista (Alto +PIB)",
                                   "Baixo": "Pessimista (Baixo −PIB)"}[c],
            key="plano_cen",
        )
    with csel2:
        prop_id = st.radio(
            "Proposta de expansão:",
            options=[1, 2, 3, 4, 5],
            format_func=lambda i: f"{PROPOSTAS[i]['icon']}  P{i} · {PROPOSTAS[i]['nome']}",
            horizontal=True, key="plano_prop",
        )
    p = PROPOSTAS[prop_id]

    # Carregar demanda do cenário escolhido
    demanda = load_demanda_2025_2035(str(CSV_PROJ), cenario=cen_lbl)

    # Banner com descrição
    st.markdown(
        f'<div style="background:linear-gradient(135deg,{p["cor"]}15,{p["cor"]}05);'
        f'border-left:4px solid {p["cor"]};border-radius:10px;padding:12px 18px;margin:8px 0 16px;">'
        f'<div style="font-size:11px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;'
        f'color:{p["cor"]};margin-bottom:4px;">Proposta {prop_id} · {p["nome"]}</div>'
        f'<div style="font-size:13.5px;color:{TEXT_PRI};line-height:1.55;">{p["descricao"]}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Parâmetros econômicos editáveis
    cpar1, cpar2 = st.columns([1, 1])
    with cpar1:
        wacc_p = st.number_input("WACC (% a.a.)", min_value=0.0,
                                value=float(WACC_PADRAO), step=0.5, format="%.1f",
                                key=f"plano_wacc_{prop_id}")
    with cpar2:
        lfsp_p = st.number_input("LFSP / Vida útil (anos)", min_value=1,
                                value=20, step=1, key=f"plano_lfsp_{prop_id}")

    # Construir plantas, simular e calcular economics
    plantas = build_plantas(prop_id, demanda)
    sim = simular_plano(plantas, demanda)
    eco = economics_plano(plantas, sim, demanda, wacc_pct=wacc_p, lfsp=lfsp_p)

    # ── KPIs grandes ──────────────────────────────────────────────────
    st.markdown("##### Resultados-chave do plano")
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(
        f'<div style="background:linear-gradient(135deg,{ACCENT},{ACCENT_D});border-radius:14px;'
        f'padding:18px 20px;color:#fff;box-shadow:0 6px 20px rgba(14,165,233,0.25);">'
        f'<div style="font-size:10.5px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;'
        f'opacity:0.92;">LCOE do plano</div>'
        f'<div style="font-size:30px;font-weight:800;margin:4px 0 0;letter-spacing:-0.6px;">'
        f'{_fmt(eco["lcoe_plano"],2)} <span style="font-size:14px;">US$/MWh</span></div>'
        f'<div style="font-size:11.5px;opacity:0.88;margin-top:2px;">'
        f'{_fmt(eco["lcoe_plano"]/1000,4)} US$/kWh</div></div>',
        unsafe_allow_html=True,
    )
    k2.markdown(kpi_card("CAPEX total (nominal)", _us(eco["capex_total"]),
                        f"{len(plantas)} usinas no plano", C_CAPEX),
                unsafe_allow_html=True)
    k3.markdown(kpi_card("VPL do custo (CAPEX+OPEX)", _us(eco["npv_total"]),
                        "trazido ao presente @ WACC", TEXT_PRI),
                unsafe_allow_html=True)
    k4.markdown(kpi_card("Potência adicionada", _fmt(eco["potencia_total"], 1, " MW"),
                        f"{_fmt(eco['ee_anual_total']/1000,2)} GWh/ano em 2035", "#047857"),
                unsafe_allow_html=True)

    _gap(8)
    # KPIs secundários: matriz 2035 realizada
    frac35 = _matriz_2035_fracoes(sim)
    st.markdown(f'<div style="font-size:11px;font-weight:700;letter-spacing:0.06em;'
                f'text-transform:uppercase;color:{TEXT_SEC};margin:8px 0 6px;">'
                f'Matriz realizada em 2035 (e meta da proposta)</div>',
                unsafe_allow_html=True)
    cm1, cm2, cm3, cm4 = st.columns(4)
    for col, tipo in zip([cm1, cm2, cm3, cm4], TIPO_ORDEM):
        real = frac35[tipo] * 100
        if tipo == "Eólica":
            frac_meta = p.get("frac_eolica_nova", 0.0)
            sub = f"cobre {_fmt(frac_meta*100,0)} % do crescimento"
        elif tipo == "Solar":
            frac_meta = p.get("frac_solar_nova", 0.0)
            sub = f"cobre {_fmt(frac_meta*100,0)} % do crescimento"
        elif tipo == "Hidro":
            sub = "resultante (+10 % en. 2025)" if p.get("ano_hidro") else "sem expansão"
        else:  # Termo
            sub = "constante (base 2025)" if prop_id != 3 else "expande (P3)"
        col.markdown(kpi_card(f"{tipo} 2035", f"{_fmt(real,1)} %", sub,
                             TIPO_COR[tipo]), unsafe_allow_html=True)

    st.markdown("---")

    # ── Gráfico 0: SOMENTE A EXPANSÃO (usinas novas, sem base 2025) ──────
    st.markdown("##### 0 · Expansão ao longo do tempo — somente usinas novas")
    st.caption("Mostra APENAS a geração das usinas construídas neste plano, "
               "acumulando ano a ano conforme entram em operação. "
               "Permite verificar: hidro sobe +10 %, térmica não decresce, renováveis crescem.")

    # calcula a geração INCREMENTAL por tipo (descontando a base 2025)
    ee_base_25 = {k: v * demanda[2025] for k, v in MATRIZ_BASE_2025.items()}

    fig0 = base_fig("", height=360)

    # linhas de referência horizontais (base 2025 de cada tipo)
    for tipo in TIPO_ORDEM:
        if ee_base_25[tipo] > 0:
            fig0.add_hline(
                y=ee_base_25[tipo] / 1e6,
                line_dash="dot", line_color=TIPO_COR[tipo], line_width=1.2,
                annotation_text=f"{tipo} base 2025",
                annotation_font_size=9,
                annotation_font_color=TIPO_COR[tipo],
                annotation_position="right",
            )

    # curvas de geração TOTAL por tipo (base + novo) para ver a trajetória real
    for tipo in TIPO_ORDEM:
        vals = [v / 1e6 for v in sim["por_tipo"][tipo]]
        fig0.add_trace(go.Scatter(
            x=sim["anos"], y=vals,
            mode="lines+markers",
            name=tipo,
            line=dict(color=TIPO_COR[tipo], width=2.8),
            marker=dict(size=7, color=TIPO_COR[tipo],
                       line=dict(color="white", width=1.5)),
            fill="tozeroy",
            fillcolor=TIPO_COR[tipo].replace("#", "rgba(") + "12)" if False else "rgba(0,0,0,0)",
            hovertemplate=f"<b>{tipo}</b> %{{x}}: %{{y:.4f}} TWh<extra></extra>",
        ))

    # marca as entradas das usinas novas como linhas verticais
    for pl in plantas:
        cor_tipo = TIPO_COR.get(pl["tipo"], TEXT_SEC)
        fig0.add_vline(
            x=pl["ano_entrada"] - 0.3,
            line_dash="dash", line_color=cor_tipo, line_width=1,
            annotation_text=f"{pl['tipo'][:3]} {pl['ano_entrada']}",
            annotation_font_size=8,
            annotation_font_color=cor_tipo,
            annotation_position="top",
        )

    fig0.update_layout(
        yaxis_title="Geração anual (TWh)",
        xaxis=dict(title="Ano", dtick=1),
        height=360,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        margin=dict(l=10, r=120, t=20, b=30),
    )
    st.plotly_chart(fig0, use_container_width=True, key=f"plano_expansao_{prop_id}")

    # tabela resumo da expansão
    ee_hidro_2025  = ee_base_25["Hidro"]
    ee_termo_2025  = ee_base_25["Termo"]
    ee_eol_2025    = ee_base_25["Eólica"]
    ee_sol_2025    = ee_base_25["Solar"]
    ee_hidro_2035  = sim["por_tipo"]["Hidro"][-1]
    ee_termo_2035  = sim["por_tipo"]["Termo"][-1]
    ee_eol_2035    = sim["por_tipo"]["Eólica"][-1]
    ee_sol_2035    = sim["por_tipo"]["Solar"][-1]

    linhas_res = []
    for tipo, v25, v35 in [("Hidro", ee_hidro_2025, ee_hidro_2035),
                            ("Termo", ee_termo_2025, ee_termo_2035),
                            ("Eólica", ee_eol_2025, ee_eol_2035),
                            ("Solar", ee_sol_2025, ee_sol_2035)]:
        delta_v = v35 - v25
        pct_v = delta_v / v25 * 100 if v25 > 0 else float("inf")
        ok = "✓" if tipo == "Termo" and abs(delta_v) < 1 else \
             ("✓" if tipo == "Hidro" and abs(v35 - v25*1.10) < 1000 else "–")
        linhas_res.append({
            "Tipo": tipo,
            "2025 (TWh)": f"{v25/1e6:.4f}",
            "2035 (TWh)": f"{v35/1e6:.4f}",
            "Δ (TWh)": f"{delta_v/1e6:+.4f}",
            "Δ (%)": f"{pct_v:+.1f} %",
            "Verificação": ok,
        })
    df_res = pd.DataFrame(linhas_res)
    st.dataframe(df_res, use_container_width=True, hide_index=True)
    _gap(12)

    # ── Gráfico 1: Demanda vs Oferta proposta ──────────────────────────
    st.markdown("##### 1 · Demanda × Oferta proposta (2025 → 2035)")
    fig1 = base_fig("", height=340)
    fig1.add_trace(go.Scatter(
        x=sim["anos"], y=[v/1e6 for v in sim["demanda"]],
        mode="lines+markers", name="Demanda (CSV)",
        line=dict(color=TEXT_PRI, width=2.6, dash="dash"),
        marker=dict(size=7, symbol="diamond", color=TEXT_PRI),
        hovertemplate="Demanda %{x}: %{y:.3f} ×10⁶ MWh<extra></extra>",
    ))
    fig1.add_trace(go.Scatter(
        x=sim["anos"], y=[v/1e6 for v in sim["total"]],
        mode="lines+markers", name="Oferta da proposta",
        line=dict(color=p["cor"], width=3.2),
        marker=dict(size=8, color=p["cor"], line=dict(color="white", width=1.5)),
        fill="tozeroy",
        fillcolor=p["cor"].replace("#", "rgba(") + "0.10)" if False else "rgba(14,165,233,0.05)",
        hovertemplate="Oferta %{x}: %{y:.3f} ×10⁶ MWh<extra></extra>",
    ))
    fig1.update_layout(yaxis_title="EE (×10⁶ MWh)", xaxis_title="Ano",
                       xaxis=dict(dtick=1), height=340,
                       legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                       margin=dict(l=10, r=10, t=20, b=30))
    st.plotly_chart(fig1, use_container_width=True, key=f"plano_dem_of_{prop_id}")

    max_def = max(sim["gap"]) if sim["gap"] else 0
    max_exc = max(sim["excedente"]) if sim["excedente"] else 0
    if max_def > 1.0:
        st.error(f"⚠️ Plano não atende demanda em algum ano (déficit máximo: "
                 f"{_fmt(max_def,0)} MWh). Reveja as etapas / metas.")
    elif max_exc > 1.0:
        st.caption(f"Curtailment máximo (excedente renovável): **{_fmt(max_exc,0)} MWh** "
                  f"em algum ano (renováveis geram mais que a demanda; a térmica desliga). "
                  f"Em todos os outros anos, a térmica modula para fechar o balanço exatamente.")
    else:
        st.caption("A oferta acompanha a demanda em todos os anos do horizonte. "
                  "A térmica do plano modula sua geração ano a ano (despacho merit-order) "
                  "para fechar o balanço sem excedente.")
    _gap(10)

    # ── Gráfico 2: Matriz energética ano a ano (stacked bars) ──────────
    st.markdown("##### 2 · Evolução da matriz energética (MWh por fonte)")
    fig2 = base_fig("", height=400)
    for tipo in TIPO_ORDEM:
        fig2.add_trace(go.Bar(
            x=sim["anos"], y=[v/1e6 for v in sim["por_tipo"][tipo]],
            name=tipo, marker_color=TIPO_COR[tipo], marker_line_width=0,
            hovertemplate=f"<b>{tipo}</b> %{{x}}: %{{y:.3f}} ×10⁶ MWh<extra></extra>",
        ))
    fig2.add_trace(go.Scatter(
        x=sim["anos"], y=[v/1e6 for v in sim["demanda"]],
        mode="lines+markers", name="Demanda",
        line=dict(color=TEXT_PRI, width=2, dash="dot"),
        marker=dict(size=5, symbol="diamond"),
        hovertemplate="Demanda %{x}: %{y:.3f} ×10⁶ MWh<extra></extra>",
    ))
    fig2.update_layout(barmode="stack", yaxis_title="EE (×10⁶ MWh)", xaxis_title="Ano",
                       xaxis=dict(dtick=1), height=400,
                       legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                       margin=dict(l=10, r=10, t=20, b=30))
    st.plotly_chart(fig2, use_container_width=True, key=f"plano_matriz_{prop_id}")
    _gap(10)

    # ── Gráfico 3: Evolução das frações da matriz (área 100%) ──────────
    st.markdown("##### 3 · Composição percentual da matriz ano a ano")
    fig3 = base_fig("", height=320)
    fracoes_anuais = {tipo: [] for tipo in TIPO_ORDEM}
    for i, _ano in enumerate(sim["anos"]):
        tot_a = sim["total"][i]
        for tipo in TIPO_ORDEM:
            fracoes_anuais[tipo].append(sim["por_tipo"][tipo][i] / tot_a * 100 if tot_a else 0)
    for tipo in TIPO_ORDEM:
        fig3.add_trace(go.Scatter(
            x=sim["anos"], y=fracoes_anuais[tipo],
            name=tipo, mode="lines",
            line=dict(color=TIPO_COR[tipo], width=0.5),
            stackgroup="one", groupnorm="percent",
            fillcolor=TIPO_COR[tipo],
            hovertemplate=f"<b>{tipo}</b> %{{x}}: %{{y:.1f}} %<extra></extra>",
        ))
    fig3.update_layout(yaxis_title="% da matriz", xaxis_title="Ano",
                       xaxis=dict(dtick=1), height=320,
                       legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                       margin=dict(l=10, r=10, t=20, b=30))
    st.plotly_chart(fig3, use_container_width=True, key=f"plano_pct_{prop_id}")

    st.markdown("---")

    # ── Tabela de usinas do plano ──────────────────────────────────────
    st.markdown("##### Portfólio de usinas do plano")
    df_p = pd.DataFrame(eco["detalhe"])
    # exclui a térmica base 2025 da tabela de portfólio (é infra existente, não expansão)
    df_p_new = df_p[~df_p["fonte"].str.contains("base 2025", na=False)].copy()
    if not df_p_new.empty:
        df_show = df_p_new.copy()
        df_show = df_show.sort_values(["ano_entrada", "tipo"]).reset_index(drop=True)
        df_show.insert(0, "#", range(1, len(df_show) + 1))
        df_show = df_show.rename(columns={
            "tipo": "Tipo", "fonte": "Fonte", "ano_entrada": "Ano entrada",
            "potencia": "Potência (MW)", "ee_anual": "EE/ano (MWh)",
            "capex": "CAPEX (US$)", "parcela_capex": "Parcela CAPEX/ano (US$)",
            "opex_f_anual": "OPEX fixo/ano (US$)", "opex_v_anual": "OPEX var./ano (US$)",
            "npv_custo": "VPL custo (US$)", "lcoe": "LCOE (US$/MWh)",
            "delta_op": "Δop (anos)",
        })

        def _color_tipo(row):
            cor = TIPO_COR.get(row["Tipo"], TEXT_SEC)
            return [f"border-left:3px solid {cor};" if c == "Tipo" else "" for c in row.index]

        sty = (df_show.drop(columns=["npv_energia"], errors="ignore").style
               .apply(_color_tipo, axis=1)
               .format({"Potência (MW)": "{:,.1f}", "EE/ano (MWh)": "{:,.0f}",
                        "CAPEX (US$)": "{:,.0f}", "Parcela CAPEX/ano (US$)": "{:,.0f}",
                        "OPEX fixo/ano (US$)": "{:,.0f}", "OPEX var./ano (US$)": "{:,.0f}",
                        "VPL custo (US$)": "{:,.0f}", "LCOE (US$/MWh)": "{:,.2f}"})
               .set_properties(**{"font-size": "12px"})
               .hide(axis="index"))
        st.dataframe(sty, use_container_width=True)
    else:
        st.info("Sem usinas no plano.")
    _gap(10)

    # ── Composição econômica do plano ──────────────────────────────────
    cc1, cc2 = st.columns([1.2, 1])
    with cc1:
        st.markdown("##### Composição do VPL do custo")
        fig4 = base_fig("", height=240)
        fig4.add_trace(go.Bar(y=["VPL"], x=[eco["npv_capex"]], name="CAPEX",
                              orientation="h", marker_color=C_CAPEX,
                              hovertemplate="CAPEX VPL: %{x:,.0f} US$<extra></extra>"))
        fig4.add_trace(go.Bar(y=["VPL"], x=[eco["npv_opex_f"]], name="OPEX fixo",
                              orientation="h", marker_color=C_OPEXF,
                              hovertemplate="OPEX fixo VPL: %{x:,.0f} US$<extra></extra>"))
        fig4.add_trace(go.Bar(y=["VPL"], x=[eco["npv_opex_v"]], name="OPEX var.",
                              orientation="h", marker_color=C_OPEXV,
                              hovertemplate="OPEX var. VPL: %{x:,.0f} US$<extra></extra>"))
        fig4.update_layout(barmode="stack", height=200, xaxis_title="US$",
                           yaxis=dict(showticklabels=False),
                           legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                           margin=dict(l=10, r=10, t=10, b=30))
        st.plotly_chart(fig4, use_container_width=True, key=f"plano_vpl_{prop_id}")

    with cc2:
        st.markdown("##### LCOE por fonte do plano")
        if not df_p.empty:
            agg = df_p.groupby("tipo").apply(
                lambda g: (g["npv_custo"].sum() / g["npv_energia"].sum())
                if g["npv_energia"].sum() > 0 else 0
            ).reindex(TIPO_ORDEM).fillna(0)
            fig5 = base_fig("", height=240)
            fig5.add_trace(go.Bar(
                x=agg.index.tolist(), y=agg.values.tolist(),
                marker_color=[TIPO_COR[t] for t in agg.index],
                text=[f"{v:.1f}" for v in agg.values], textposition="outside",
                hovertemplate="%{x}: %{y:,.2f} US$/MWh<extra></extra>",
            ))
            fig5.update_layout(showlegend=False, height=240, yaxis_title="US$/MWh",
                               margin=dict(l=10, r=10, t=10, b=30))
            st.plotly_chart(fig5, use_container_width=True, key=f"plano_lcoe_t_{prop_id}")

    st.markdown("---")

    # ── Fluxo de caixa agregado do plano ───────────────────────────────
    st.markdown("##### Fluxo de caixa do plano (agregado)")
    anos_eixo = list(range(2025, 2036))
    # nominal: barras por componente
    fig6 = base_fig("", height=320)
    fig6.add_trace(go.Bar(x=anos_eixo, y=eco["serie_capex_nom"], name="Parcela CAPEX",
                          marker_color=C_CAPEX,
                          hovertemplate="%{x}<br>CAPEX nom.: %{y:,.0f} US$<extra></extra>"))
    fig6.add_trace(go.Bar(x=anos_eixo, y=eco["serie_opex_f_nom"], name="OPEX fixo",
                          marker_color=C_OPEXF,
                          hovertemplate="%{x}<br>OPEX fixo: %{y:,.0f} US$<extra></extra>"))
    fig6.add_trace(go.Bar(x=anos_eixo, y=eco["serie_opex_v_nom"], name="OPEX variável",
                          marker_color=C_OPEXV,
                          hovertemplate="%{x}<br>OPEX var.: %{y:,.0f} US$<extra></extra>"))
    fig6.update_layout(barmode="stack", height=320, yaxis_title="US$/ano (nominal)",
                       xaxis=dict(title="ano", dtick=1),
                       legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                       margin=dict(l=10, r=10, t=10, b=30))
    st.plotly_chart(fig6, use_container_width=True, key=f"plano_cf_nom_{prop_id}")
    st.caption("Desembolso nominal anual do plano: parcelas de CAPEX (todas as usinas em "
               "operação naquele ano) + OPEX fixo + OPEX variável.")
    _gap(6)

    # descontado + acumulado
    tot_pv = eco["serie_capex"] + eco["serie_opex_f"] + eco["serie_opex_v"]
    acum = np.cumsum(tot_pv)
    fig7 = base_fig("", height=320)
    fig7.add_trace(go.Bar(x=anos_eixo, y=tot_pv, name="Custo VP do ano",
                          marker_color="rgba(2,132,199,0.30)",
                          hovertemplate="%{x}<br>VP do ano: %{y:,.0f} US$<extra></extra>"))
    fig7.add_trace(go.Scatter(x=anos_eixo, y=acum, name="VPL acumulado",
                              mode="lines+markers",
                              line=dict(color=ACCENT_D, width=2.5),
                              marker=dict(size=6, color=ACCENT_D),
                              fill="tozeroy", fillcolor="rgba(14,165,233,0.12)",
                              hovertemplate="%{x}<br>acumulado: %{y:,.0f} US$<extra></extra>"))
    fig7.add_hline(y=eco["npv_total"], line_dash="dot", line_color="#047857")
    fig7.add_annotation(x=anos_eixo[0], y=eco["npv_total"],
                        text=f"VPL total = {_us(eco['npv_total'])}",
                        showarrow=False, font=dict(size=11, color="#047857"),
                        xanchor="left", yanchor="bottom", yshift=2)
    fig7.update_layout(height=320, yaxis_title="US$",
                       xaxis=dict(title="ano", dtick=1),
                       legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                       margin=dict(l=10, r=10, t=10, b=30))
    st.plotly_chart(fig7, use_container_width=True, key=f"plano_cf_pv_{prop_id}")

    st.markdown("---")

    # ── Download CSV ─────────────────────────────────────────────────
    st.markdown("##### 📥 Downloads")
    cd1, cd2 = st.columns(2)
    with cd1:
        if not df_p.empty:
            csv_p = df_show.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Portfólio de usinas (CSV)", csv_p,
                              file_name=f"plano_P{prop_id}_usinas.csv",
                              mime="text/csv", key=f"dl_usinas_{prop_id}")
    with cd2:
        df_anual = pd.DataFrame({
            "Ano": sim["anos"],
            "Demanda (MWh)": sim["demanda"],
            "Oferta total (MWh)": sim["total"],
            "Gap (MWh)": sim["gap"],
            **{f"EE {t} (MWh)": sim["por_tipo"][t] for t in TIPO_ORDEM},
            "CAPEX nominal/ano (US$)": eco["serie_capex_nom"],
            "OPEX fixo/ano (US$)":     eco["serie_opex_f_nom"],
            "OPEX var./ano (US$)":     eco["serie_opex_v_nom"],
            "Custo VP do ano (US$)":   tot_pv,
            "VPL acumulado (US$)":     acum,
        })
        csv_a = df_anual.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Resultado ano a ano (CSV)", csv_a,
                          file_name=f"plano_P{prop_id}_anual.csv",
                          mime="text/csv", key=f"dl_anual_{prop_id}")


# =======================================================================
#  ABA 5 · COMPARAÇÃO DAS 5 PROPOSTAS
# =======================================================================
@st.cache_data(ttl=60)
def _cache_all_propostas(wacc: float, lfsp: int, cenario: str = "Referencia") -> dict:
    """Roda as 5 propostas de uma vez (cacheado)."""
    demanda = load_demanda_2025_2035(str(CSV_PROJ), cenario=cenario)
    out = {}
    for pid in [1, 2, 3, 4, 5]:
        plantas = build_plantas(pid, demanda)
        sim = simular_plano(plantas, demanda)
        eco = economics_plano(plantas, sim, demanda, wacc_pct=wacc, lfsp=lfsp)
        out[pid] = dict(plantas=plantas, sim=sim, eco=eco,
                       frac35=_matriz_2035_fracoes(sim))
    return out


def tab_comparacao_propostas():
    section_title("Comparação das 5 Propostas",
                 "Lado a lado: LCOE, custos, matriz 2035 e ranking das estratégias")

    cpar0, cpar1, cpar2 = st.columns([1.4, 1, 1])
    with cpar0:
        cen_c = st.selectbox(
            "📊 Cenário de demanda",
            options=["Referencia", "Alto", "Baixo"],
            format_func=lambda c: {"Referencia": "Referência",
                                   "Alto": "Otimista (Alto +PIB)",
                                   "Baixo": "Pessimista (Baixo −PIB)"}[c],
            key="comp_cen",
        )
    with cpar1:
        wacc_c = st.number_input("WACC (% a.a.)", min_value=0.0,
                                value=float(WACC_PADRAO), step=0.5, format="%.1f",
                                key="comp_wacc")
    with cpar2:
        lfsp_c = st.number_input("LFSP (anos)", min_value=1,
                                value=20, step=1, key="comp_lfsp")

    R = _cache_all_propostas(wacc_c, lfsp_c, cen_c)

    # Identifica o melhor LCOE
    lcoes = {pid: R[pid]["eco"]["lcoe_plano"] for pid in [1, 2, 3, 4, 5]}
    best_id = min(lcoes, key=lcoes.get)
    worst_id = max(lcoes, key=lcoes.get)

    # ── Tabela comparativa ────────────────────────────────────────────
    st.markdown("##### Tabela comparativa")
    linhas = []
    for pid in [1, 2, 3, 4, 5]:
        p = PROPOSTAS[pid]
        eco = R[pid]["eco"]
        sim = R[pid]["sim"]
        frac = R[pid]["frac35"]
        linhas.append({
            "Proposta": f"{p['icon']} P{pid} · {p['nome']}",
            "LCOE (US$/MWh)": eco["lcoe_plano"],
            "CAPEX total (US$)": eco["capex_total"],
            "VPL custo (US$)": eco["npv_total"],
            "Potência (MW)": eco["potencia_total"],
            "Usinas": len(R[pid]["plantas"]),
            "Hidro 2035": frac["Hidro"] * 100,
            "Termo 2035": frac["Termo"] * 100,
            "Eólica 2035": frac["Eólica"] * 100,
            "Solar 2035": frac["Solar"] * 100,
        })
    df_cmp = pd.DataFrame(linhas)

    def _hl(row):
        if "P{}".format(best_id) in row["Proposta"]:
            return ["background:#dcfce7;font-weight:700;" for _ in row.index]
        if "P{}".format(worst_id) in row["Proposta"]:
            return ["background:#fee2e2;" for _ in row.index]
        return ["" for _ in row.index]

    sty = (df_cmp.style.apply(_hl, axis=1)
           .format({"LCOE (US$/MWh)": "{:,.2f}",
                    "CAPEX total (US$)": "{:,.0f}",
                    "VPL custo (US$)": "{:,.0f}",
                    "Potência (MW)": "{:,.1f}",
                    "Hidro 2035": "{:,.1f} %", "Termo 2035": "{:,.1f} %",
                    "Eólica 2035": "{:,.1f} %", "Solar 2035": "{:,.1f} %"})
           .set_properties(**{"font-size": "12.5px"})
           .hide(axis="index"))
    st.dataframe(sty, use_container_width=True)
    st.caption(f"🏆 **Melhor LCOE**: P{best_id} · {PROPOSTAS[best_id]['nome']} "
               f"({_fmt(lcoes[best_id],2)} US$/MWh) · "
               f"⛔ Pior: P{worst_id} ({_fmt(lcoes[worst_id],2)} US$/MWh)")

    st.markdown("---")

    # ── Ranking LCOE ──────────────────────────────────────────────────
    st.markdown("##### Ranking de LCOE")
    ranked = sorted([1, 2, 3, 4, 5], key=lambda i: lcoes[i])
    fig_r = base_fig("", height=320)
    fig_r.add_trace(go.Bar(
        x=[lcoes[i] for i in ranked],
        y=[f"P{i} · {PROPOSTAS[i]['nome']}" for i in ranked],
        orientation="h",
        marker_color=[("#16a34a" if i == best_id else
                      "#dc2626" if i == worst_id else PROPOSTAS[i]["cor"]) for i in ranked],
        text=[f"{lcoes[i]:.2f}" for i in ranked], textposition="outside",
        hovertemplate="%{y}<br>LCOE: %{x:,.2f} US$/MWh<extra></extra>",
    ))
    fig_r.update_layout(showlegend=False, height=320, xaxis_title="LCOE (US$/MWh)",
                        margin=dict(l=10, r=80, t=10, b=30))
    st.plotly_chart(fig_r, use_container_width=True, key="cmp_rank")
    _gap(10)

    # ── Donuts da matriz 2035 ─────────────────────────────────────────
    st.markdown("##### Matriz energética em 2035 por proposta")
    cols_d = st.columns(5)
    for i, pid in enumerate([1, 2, 3, 4, 5]):
        with cols_d[i]:
            frac = R[pid]["frac35"]
            labs = [t for t in TIPO_ORDEM if frac[t] > 0.001]
            vals = [frac[t] * 100 for t in labs]
            cores = [TIPO_COR[t] for t in labs]
            fig_d = base_fig("", height=200)
            fig_d.add_trace(go.Pie(labels=labs, values=vals, hole=0.55,
                                   marker=dict(colors=cores),
                                   textinfo="label+percent",
                                   textfont=dict(size=9),
                                   hovertemplate="%{label}: %{value:.1f} %<extra></extra>"))
            star = " 🏆" if pid == best_id else ""
            fig_d.update_layout(height=200, showlegend=False,
                                margin=dict(l=5, r=5, t=30, b=5),
                                title=dict(text=f"P{pid}{star}",
                                          font=dict(size=12, color=TEXT_PRI, weight="bold"),
                                          x=0.5, xanchor="center"))
            st.plotly_chart(fig_d, use_container_width=True, key=f"cmp_donut_{pid}")
    _gap(10)

    # ── Composição econômica empilhada por proposta ────────────────────
    st.markdown("##### Composição do VPL do custo por proposta")
    fig_cmp = base_fig("", height=340)
    nomes = [f"P{i}" for i in [1, 2, 3, 4, 5]]
    fig_cmp.add_trace(go.Bar(name="CAPEX", x=nomes,
                             y=[R[i]["eco"]["npv_capex"] for i in [1, 2, 3, 4, 5]],
                             marker_color=C_CAPEX,
                             hovertemplate="%{x} CAPEX: %{y:,.0f} US$<extra></extra>"))
    fig_cmp.add_trace(go.Bar(name="OPEX fixo", x=nomes,
                             y=[R[i]["eco"]["npv_opex_f"] for i in [1, 2, 3, 4, 5]],
                             marker_color=C_OPEXF,
                             hovertemplate="%{x} OPEX fixo: %{y:,.0f} US$<extra></extra>"))
    fig_cmp.add_trace(go.Bar(name="OPEX var.", x=nomes,
                             y=[R[i]["eco"]["npv_opex_v"] for i in [1, 2, 3, 4, 5]],
                             marker_color=C_OPEXV,
                             hovertemplate="%{x} OPEX var.: %{y:,.0f} US$<extra></extra>"))
    fig_cmp.update_layout(barmode="stack", height=340, yaxis_title="VPL (US$)",
                          legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                          margin=dict(l=10, r=10, t=10, b=30))
    st.plotly_chart(fig_cmp, use_container_width=True, key="cmp_vpl")
    _gap(10)

    # ── Demanda × oferta de todas (linhas sobrepostas) ────────────────
    st.markdown("##### Oferta de cada proposta vs Demanda")
    fig_of = base_fig("", height=340)
    fig_of.add_trace(go.Scatter(
        x=R[1]["sim"]["anos"], y=[v/1e6 for v in R[1]["sim"]["demanda"]],
        mode="lines+markers", name="Demanda",
        line=dict(color=TEXT_PRI, width=3, dash="dash"),
        marker=dict(size=7, symbol="diamond"),
        hovertemplate="Demanda %{x}: %{y:.3f}<extra></extra>",
    ))
    for pid in [1, 2, 3, 4, 5]:
        p = PROPOSTAS[pid]
        fig_of.add_trace(go.Scatter(
            x=R[pid]["sim"]["anos"], y=[v/1e6 for v in R[pid]["sim"]["total"]],
            mode="lines+markers", name=f"P{pid} · {p['nome']}",
            line=dict(color=p["cor"], width=1.8),
            marker=dict(size=5),
            hovertemplate=f"P{pid} %{{x}}: %{{y:.3f}}<extra></extra>",
        ))
    fig_of.update_layout(yaxis_title="EE (×10⁶ MWh)", xaxis_title="Ano",
                         xaxis=dict(dtick=1), height=340,
                         legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                         margin=dict(l=10, r=10, t=10, b=30))
    st.plotly_chart(fig_of, use_container_width=True, key="cmp_of")
    st.caption("Todas as propostas cobrem a demanda ano a ano por construção — a térmica "
               "do plano fecha o balanço onde falta. As diferenças entre curvas refletem "
               "leves excedentes de geração renovável escalonada.")

    # ── Download tabela comparativa ───────────────────────────────────
    _gap(10)
    csv_cmp = df_cmp.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Baixar tabela comparativa (CSV)", csv_cmp,
                      file_name="comparacao_propostas.csv", mime="text/csv",
                      key="dl_cmp")


def run_matriz(page=None):
    st.markdown(
        f'<h1 style="font-size:32px;font-weight:800;color:{TEXT_PRI};letter-spacing:-0.6px;'
        f'margin-bottom:2px;">Matriz Energética · UTÓPIA</h1>'
        f'<p style="font-size:14px;color:{TEXT_SEC};margin-bottom:18px;">'
        f'Plano Decenal de Expansão (2025 → 2035) — análise econômica e tarifária</p>',
        unsafe_allow_html=True,
    )

    t_atual, t_cfg, t_econ, t_plano, t_comp = st.tabs(
        ["📅 Atualidade (2025)",
         "🧮 Configuração Econômica",
         "🧪 TESTE da Análise Econômica",
         "🗺️ Plano de Expansão",
         "⚖️ Comparação das Propostas"]
    )
    with t_atual:
        tab_atualidade()
    with t_cfg:
        tab_config_economico()
    with t_econ:
        tab_analise_economica()
    with t_plano:
        tab_plano_expansao()
    with t_comp:
        tab_comparacao_propostas()


if __name__ == "__main__":
    run_matriz()
