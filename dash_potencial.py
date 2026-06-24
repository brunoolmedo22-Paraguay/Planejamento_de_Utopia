"""
dash_potencial.py
======================================================================
Seção "Potencial Energético" — País de Utópia
Abas: Geolocalização · Eólica · Solar · Hidro · Termo

Arquivos esperados na MESMA pasta deste módulo (raiz do projeto):
    - ENTREGA_DEMANDA.xlsx      (aba "RESUMO UTÓPIA" -> pop/PIB/EE 2025)
    - GSA_Report_UTOPIA.xlsx    (relatório do Global Solar Atlas)
    - mapa.jpeg                 (mapa do país)

Entry-point chamado pelo app central:
    from dash_potencial import run_potencial
    run_potencial()
======================================================================
"""

import base64
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ── Raiz portátil (assets ao lado deste arquivo) ────────────────────
ROOT = Path(__file__).parent

EXCEL_HIST = str(ROOT / "ENTREGA_DEMANDA.xlsx")
EXCEL_GSA  = str(ROOT / "GSA_Report_UTOPIA.xlsx")
MAPA_FILE  = "mapa.jpeg"

# ── Constantes do país (valores fixos informados) ───────────────────
COORD_TXT      = "15° 00' 00\" N  ·  75° 00' 00\" W"
MAPS_LINK      = ("https://www.google.com/maps/place/15%C2%B000'00.0%22N+75%C2%B000'00.0%22W/"
                  "@17.0533229,-87.2473738,5.2z/data=!4m4!3m3!8m2!3d15!4d-75")
VIZINHO_NOME   = "Jamaica"
VIZINHO_DIST   = 420       # km
COMPRIMENTO_KM = 200       # comprimento da ilha
LARGURA_KM     = 60        # largura da ilha
AREA_KM2       = 12000     # área total

# Participações da matriz (BEN) informadas
HIDRO_SHARE_2025 = 0.40    # BEN 2025
TERMO_SHARE_2026 = 0.55    # BEN 2026

# ── Paleta / tema (idêntico ao dash_historico p/ consistência) ──────
ACCENT   = "#0ea5e9"
ACCENT_D = "#0284c7"
TEXT_PRI = "#0f172a"
TEXT_SEC = "#64748b"
GRID_CLR = "rgba(226,232,240,0.6)"
BG_CHART = "#ffffff"

SOL = "#f59e0b"    # solar (âmbar)
WIN = "#10b981"    # eólica (verde)
HYD = "#0ea5e9"    # hidro (azul)
THR = "#ef4444"    # termo (vermelho)

THEME = dict(
    plot_bgcolor=BG_CHART,
    paper_bgcolor=BG_CHART,
    font=dict(family="-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif", color=TEXT_PRI, size=12),
    margin=dict(l=15, r=15, t=46, b=15),
    xaxis=dict(showgrid=True, gridcolor=GRID_CLR, linecolor="#e2e8f0", tickfont=dict(size=11)),
    yaxis=dict(showgrid=True, gridcolor=GRID_CLR, linecolor="#e2e8f0", tickfont=dict(size=11)),
    hoverlabel=dict(bgcolor="white", font_size=12),
)

# Mapa de rótulos amigáveis das estatísticas do Global Solar Atlas
GSA_META = {
    "PVOUT": ("Produção fotovoltaica específica (PVOUT)", "kWh/kWp/dia"),
    "GHI":   ("Irradiação horizontal global (GHI)",       "kWh/m²/dia"),
    "DNI":   ("Irradiação normal direta (DNI)",           "kWh/m²/dia"),
    "DIF":   ("Irradiação difusa horizontal (DIF)",       "kWh/m²/dia"),
    "GTI":   ("Irradiação global inclinada (GTI)",        "kWh/m²/dia"),
    "TEMP":  ("Temperatura do ar",                        "°C"),
    "OPTA":  ("Inclinação ótima dos módulos",             "°"),
    "ELE":   ("Elevação do terreno",                      "m"),
}


# =======================================================================
#  HELPERS GERAIS
# =======================================================================
def _img_b64(filename: str) -> str:
    p = ROOT / filename
    if not p.exists():
        return ""
    ext = p.suffix.lower().lstrip(".")
    mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "gif": "gif", "webp": "webp"}.get(ext, "png")
    return f"data:image/{mime};base64,{base64.b64encode(p.read_bytes()).decode()}"


def _fmt(v, dec=0, suf=""):
    """Formata número com separador de milhar (estilo BR)."""
    if v is None:
        return "—"
    s = f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s}{suf}"


def kpi_card(label, value, sub="", color=TEXT_PRI):
    css = ("background:#ffffff; border:1px solid #e2e8f0; border-radius:14px;"
           "padding:12px 14px; box-shadow:0 2px 10px rgba(14,165,233,0.05); height:100%;")
    sub_html = (f'<div style="font-size:11px;color:{TEXT_SEC};margin-top:3px;font-weight:500;">{sub}</div>'
                if sub else "")
    return (
        f'<div style="{css}">'
        f'<div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;'
        f'letter-spacing:0.05em;margin-bottom:2px;">{label}</div>'
        f'<div style="font-size:18px;font-weight:700;color:{color};letter-spacing:-0.3px;">{value}</div>'
        f'{sub_html}</div>'
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
        title=dict(text=title, font=dict(size=13, color=TEXT_PRI, weight="bold"), x=0, xanchor="left", pad=dict(l=4, t=4)),
        height=height, showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        **THEME,
    )
    return fig


# =======================================================================
#  CARGA DE DADOS
# =======================================================================
@st.cache_data
def load_socio_2025(excel_path: str) -> dict:
    """Lê população, PIB e consumo EE de 2025 da aba RESUMO UTÓPIA."""
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


@st.cache_data
def load_gsa_stats(excel_path: str) -> dict:
    """Lê as estatísticas (médias e percentis) do relatório Global Solar Atlas."""
    out = {}
    for key in GSA_META:
        df = pd.read_excel(excel_path, sheet_name=f"{key}_stats", header=None)
        stats, unit = {}, ""
        for _, r in df.iterrows():
            lbl = r[0]
            if isinstance(lbl, str) and (lbl in ("Average", "Maximum", "Minimum") or lbl.startswith("Percentile")):
                try:
                    stats[lbl] = float(r[1])
                except (TypeError, ValueError):
                    pass
                if isinstance(r[2], str):
                    unit = r[2]
        out[key] = {"unit": unit, "stats": stats}
    return out


@st.cache_data
def load_gsa_dist(excel_path: str, key: str) -> pd.DataFrame:
    """Lê a distribuição (histograma) de uma variável do GSA."""
    raw = pd.read_excel(excel_path, sheet_name=f"{key}_dist", header=None)
    rows = []
    started = False
    for _, r in raw.iterrows():
        c0 = r[0]
        if isinstance(c0, str) and c0.startswith("From"):
            started = True
            continue
        if started:
            if c0 is None or c0 == "" or not isinstance(c0, (int, float)):
                break
            rows.append((float(r[0]), float(r[1]), float(r[2])))
    return pd.DataFrame(rows, columns=["de", "ate", "pct"])


# =======================================================================
#  ABA 1 — GEOLOCALIZAÇÃO
# =======================================================================
def tab_geolocalizacao(socio: dict):
    section_title("Geolocalização", "Posição, dimensões e perfil socioeconômico do país (referência: 2025)")

    col_map, col_kpi = st.columns([1.15, 1], gap="large")

    # ── Mapa ───────────────────────────────────────────────────────
    with col_map:
        img = _img_b64(MAPA_FILE)
        if img:
            st.markdown(
                f'<div style="border:1px solid #e2e8f0;border-radius:18px;overflow:hidden;'
                f'box-shadow:0 4px 20px rgba(14,165,233,0.08);">'
                f'<img src="{img}" style="width:100%;display:block;object-fit:cover;" /></div>',
                unsafe_allow_html=True,
            )
        else:
            st.warning(f"Arquivo `{MAPA_FILE}` não encontrado na raiz do projeto.")
        st.markdown(
            f'<a href="{MAPS_LINK}" target="_blank" style="text-decoration:none;">'
            f'<div style="margin-top:14px;background:linear-gradient(135deg,{ACCENT},{ACCENT_D});'
            f'color:white;text-align:center;padding:12px;border-radius:12px;font-weight:600;'
            f'box-shadow:0 4px 14px rgba(14,165,233,0.35);">📍 Abrir no Google Maps</div></a>',
            unsafe_allow_html=True,
        )

    # ── KPIs geográficos + socioeconômicos ────────────────────────
    with col_kpi:
        g1, g2 = st.columns(2)
        g1.markdown(kpi_card("Coordenadas centrais", "15° N", "75° W", ACCENT), unsafe_allow_html=True)
        g2.markdown(kpi_card("Vizinho mais próximo", VIZINHO_NOME, f"≈ {VIZINHO_DIST} km de distância"), unsafe_allow_html=True)
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

        d1, d2, d3 = st.columns(3)
        d1.markdown(kpi_card("Comprimento", f"{COMPRIMENTO_KM} km"), unsafe_allow_html=True)
        d2.markdown(kpi_card("Largura", f"{LARGURA_KM} km"), unsafe_allow_html=True)
        d3.markdown(kpi_card("Área total", f"{_fmt(AREA_KM2)}", "km²"), unsafe_allow_html=True)
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

        s1, s2, s3 = st.columns(3)
        s1.markdown(kpi_card("População 2025", _fmt(socio["pop"]), "habitantes", ACCENT_D), unsafe_allow_html=True)
        s2.markdown(kpi_card("PIB 2025", f"Utd$ {_fmt(socio['pib'])}", "absoluto"), unsafe_allow_html=True)
        s3.markdown(kpi_card("Consumo EE 2025", f"{_fmt(socio['ee'])}", "MWh/ano", "#7c3aed"), unsafe_allow_html=True)

    st.markdown("---")
    st.caption("As dimensões definem a área de referência (12.000 km²) usada nas estimativas de potencial "
               "solar e eólico das abas seguintes.")


# =======================================================================
#  ABA 2 — EÓLICA
# =======================================================================
def tab_eolica(socio: dict):
    section_title("Potencial Eólico",
                  "Estimativa a partir das leituras do Global Wind Atlas para a área do país")

    st.markdown(
        '<a href="https://globalwindatlas.info/" target="_blank" style="text-decoration:none;">'
        f'<div style="display:inline-block;background:rgba(16,185,129,0.10);border:1px solid rgba(16,185,129,0.3);'
        f'color:#047857;padding:8px 16px;border-radius:10px;font-weight:600;font-size:13px;margin-bottom:6px;">'
        f'🌬️ Abrir o Global Wind Atlas</div></a>',
        unsafe_allow_html=True,
    )
    st.info("Os campos abaixo vêm **pré-preenchidos com valores típicos da região**. "
            "Substitua pelos números que o Global Wind Atlas devolve para o retângulo do país "
            "(velocidade média, densidade de potência e fator de capacidade na altura do rotor).")

    # ── Entradas (leituras do GWA + premissas de implantação) ─────
    with st.expander("⚙️  Parâmetros do recurso e da implantação", expanded=True):
        c1, c2, c3 = st.columns(3)
        v100 = c1.number_input("Velocidade média @100 m (m/s)", 3.0, 12.0, 6.5, 0.1, key="w_v100")
        pd100 = c2.number_input("Densidade de potência @100 m (W/m²)", 50.0, 1500.0, 320.0, 10.0, key="w_pd")
        cf = c3.number_input("Fator de capacidade (0–1)", 0.10, 0.65, 0.32, 0.01, key="w_cf")

        c4, c5, c6 = st.columns(3)
        frac = c4.slider("Fração da área aproveitável (%)", 1, 30, 5, key="w_frac") / 100
        dens = c5.number_input("Densidade de instalação (MW/km²)", 1.0, 12.0, 4.0, 0.5, key="w_dens")
        perdas = c6.slider("Perdas de sistema (%)", 0, 25, 10, key="w_loss") / 100

        st.markdown("**Perfil de velocidade por altura (opcional, para o gráfico):**")
        h1, h2, h3 = st.columns(3)
        v150 = h1.number_input("Vel. @150 m (m/s)", 3.0, 13.0, 7.0, 0.1, key="w_v150")
        v200 = h2.number_input("Vel. @200 m (m/s)", 3.0, 14.0, 7.4, 0.1, key="w_v200")
        v50  = h3.number_input("Vel. @50 m (m/s)",  3.0, 12.0, 5.9, 0.1, key="w_v50")

    # ── Cálculos ──────────────────────────────────────────────────
    area_util = AREA_KM2 * frac                       # km²
    cap_mw = area_util * dens                          # MW
    energia_mwh = cap_mw * cf * 8760 * (1 - perdas)    # MWh/ano
    cobertura = energia_mwh / socio["ee"] * 100 if socio["ee"] else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi_card("Área aproveitável", f"{_fmt(area_util)}", "km²", WIN), unsafe_allow_html=True)
    k2.markdown(kpi_card("Capacidade instalável", f"{_fmt(cap_mw/1000, 2)} GW", f"{_fmt(cap_mw)} MW", WIN), unsafe_allow_html=True)
    k3.markdown(kpi_card("Geração anual estimada", f"{_fmt(energia_mwh/1e6, 2)} TWh", f"{_fmt(energia_mwh)} MWh", WIN), unsafe_allow_html=True)
    k4.markdown(kpi_card("Cobertura da demanda 2025", f"{_fmt(cobertura, 1)} %", "do consumo de 2025", ACCENT_D), unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    g1, g2 = st.columns(2)

    # Perfil de velocidade
    with g1:
        fig = base_fig("Perfil de velocidade do vento por altura")
        alturas = [50, 100, 150, 200]
        vels = [v50, v100, v150, v200]
        fig.add_trace(go.Scatter(
            x=vels, y=alturas, mode="lines+markers+text",
            text=[f"{v:.1f}" for v in vels], textposition="top right",
            line=dict(color=WIN, width=3), marker=dict(size=9, color=WIN),
            name="Velocidade média", hovertemplate="%{x:.1f} m/s @ %{y} m<extra></extra>",
        ))
        fig.update_layout(xaxis_title="Velocidade (m/s)", yaxis_title="Altura (m)", showlegend=False)
        st.plotly_chart(fig, use_container_width=True, key="w_perfil")

    # Potencial vs demanda
    with g2:
        fig = base_fig("Geração potencial vs. consumo atual")
        fig.add_trace(go.Bar(
            x=["Consumo 2025", "Potencial eólico"],
            y=[socio["ee"], energia_mwh],
            marker_color=["#cbd5e1", WIN],
            text=[f"{_fmt(socio['ee']/1e6,2)} TWh", f"{_fmt(energia_mwh/1e6,2)} TWh"],
            textposition="outside",
            hovertemplate="%{y:,.0f} MWh<extra></extra>",
        ))
        fig.update_layout(yaxis_title="MWh/ano", showlegend=False)
        st.plotly_chart(fig, use_container_width=True, key="w_bar")

    st.caption("Capacidade = área × fração aproveitável × densidade (MW/km²). "
               "Geração = capacidade × fator de capacidade × 8.760 h × (1 − perdas). "
               "É um potencial técnico: ajuste a fração aproveitável para um cenário realista de implantação.")


# =======================================================================
#  ABA 3 — SOLAR
# =======================================================================
def tab_solar(socio: dict, gsa: dict):
    section_title("Potencial Solar",
                  "Recurso medido pelo Global Solar Atlas para a área do país (12.000 km²)")

    st.markdown(
        '<a href="https://globalsolaratlas.info/" target="_blank" style="text-decoration:none;">'
        f'<div style="display:inline-block;background:rgba(245,158,11,0.12);border:1px solid rgba(245,158,11,0.35);'
        f'color:#b45309;padding:8px 16px;border-radius:10px;font-weight:600;font-size:13px;margin-bottom:10px;">'
        f'☀️ Dados oficiais do Global Solar Atlas (GSA_Report_UTOPIA.xlsx)</div></a>',
        unsafe_allow_html=True,
    )

    # ── KPIs do recurso (médias do GSA) ───────────────────────────
    def avg(k):
        return gsa.get(k, {}).get("stats", {}).get("Average")

    r1 = st.columns(4)
    r1[0].markdown(kpi_card("GHI (horizontal global)", f"{_fmt(avg('GHI'),2)}", "kWh/m²/dia", SOL), unsafe_allow_html=True)
    r1[1].markdown(kpi_card("DNI (normal direta)", f"{_fmt(avg('DNI'),2)}", "kWh/m²/dia", SOL), unsafe_allow_html=True)
    r1[2].markdown(kpi_card("GTI (inclinada ótima)", f"{_fmt(avg('GTI'),2)}", "kWh/m²/dia", SOL), unsafe_allow_html=True)
    r1[3].markdown(kpi_card("PVOUT (produção FV)", f"{_fmt(avg('PVOUT'),2)}", "kWh/kWp/dia", "#d97706"), unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    r2 = st.columns(4)
    r2[0].markdown(kpi_card("Inclinação ótima", f"{_fmt(avg('OPTA'),0)}°", "dos módulos"), unsafe_allow_html=True)
    r2[1].markdown(kpi_card("Temperatura média", f"{_fmt(avg('TEMP'),1)} °C", "do ar"), unsafe_allow_html=True)
    r2[2].markdown(kpi_card("Elevação média", f"{_fmt(avg('ELE'),0)} m", "do terreno"), unsafe_allow_html=True)
    r2[3].markdown(kpi_card("DIF (difusa)", f"{_fmt(avg('DIF'),2)}", "kWh/m²/dia"), unsafe_allow_html=True)

    st.markdown("---")

    # ── Estimativa de potencial fotovoltaico ──────────────────────
    pvout_dia = avg("PVOUT") or 4.0
    pvout_ano = pvout_dia * 365          # kWh/kWp/ano (já inclui PR do GSA)

    with st.expander("⚙️  Premissas de implantação fotovoltaica", expanded=True):
        c1, c2 = st.columns(2)
        frac = c1.slider("Fração da área aproveitável (%)", 1, 20, 2, key="s_frac") / 100
        dens = c2.number_input("Densidade de instalação (MWp/km²)", 20.0, 80.0, 45.0, 5.0, key="s_dens")

    area_util = AREA_KM2 * frac
    cap_mwp = area_util * dens
    energia_mwh = cap_mwp * pvout_ano       # MWh/ano  (MWp × kWh/kWp/ano = MWh/ano)
    cobertura = energia_mwh / socio["ee"] * 100 if socio["ee"] else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi_card("Rendimento específico", f"{_fmt(pvout_ano,0)}", "kWh/kWp/ano", SOL), unsafe_allow_html=True)
    k2.markdown(kpi_card("Capacidade instalável", f"{_fmt(cap_mwp/1000,2)} GWp", f"{_fmt(cap_mwp)} MWp", SOL), unsafe_allow_html=True)
    k3.markdown(kpi_card("Geração anual estimada", f"{_fmt(energia_mwh/1e6,2)} TWh", f"{_fmt(energia_mwh)} MWh", "#d97706"), unsafe_allow_html=True)
    k4.markdown(kpi_card("Cobertura da demanda 2025", f"{_fmt(cobertura,1)} %", "do consumo de 2025", ACCENT_D), unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    g1, g2 = st.columns(2)

    # Comparativo de irradiação
    with g1:
        fig = base_fig("Componentes da irradiação (média anual)")
        comp = [("GHI", avg("GHI")), ("DNI", avg("DNI")), ("DIF", avg("DIF")), ("GTI", avg("GTI"))]
        fig.add_trace(go.Bar(
            x=[c[0] for c in comp], y=[c[1] for c in comp],
            marker_color=[SOL, "#fbbf24", "#fde68a", "#d97706"],
            text=[f"{c[1]:.2f}" for c in comp], textposition="outside",
            hovertemplate="%{x}: %{y:.2f} kWh/m²/dia<extra></extra>",
        ))
        fig.update_layout(yaxis_title="kWh/m²/dia", showlegend=False)
        st.plotly_chart(fig, use_container_width=True, key="s_irr")

    # Distribuição PVOUT
    with g2:
        try:
            dist = load_gsa_dist(EXCEL_GSA, "PVOUT")
            fig = base_fig("Distribuição do PVOUT na área")
            centros = ((dist["de"] + dist["ate"]) / 2).round(2)
            fig.add_trace(go.Bar(
                x=centros, y=dist["pct"] * 100,
                marker_color=SOL, width=0.18,
                text=[f"{p*100:.1f}%" for p in dist["pct"]], textposition="outside",
                hovertemplate="%{x:.2f} kWh/kWp/dia<br>%{y:.1f}% da área<extra></extra>",
            ))
            fig.update_layout(xaxis_title="PVOUT (kWh/kWp/dia)", yaxis_title="% da área", showlegend=False)
            st.plotly_chart(fig, use_container_width=True, key="s_dist")
        except Exception as e:
            st.caption(f"Distribuição PVOUT indisponível: {e}")

    st.caption("O PVOUT do GSA já incorpora o desempenho típico do sistema (perdas térmicas, inversor, etc.). "
               "Geração = capacidade (MWp) × PVOUT anual (kWh/kWp/ano).")


# =======================================================================
#  ABA 4 — HIDRO
# =======================================================================
def tab_hidro(socio: dict):
    section_title("Potencial Hidrelétrico",
                  "Participação atual na matriz e margem de expansão permitida para o PDE 2035")

    st.info(f"No **BEN 2025** a hidroeletricidade respondeu por **{HIDRO_SHARE_2025*100:.0f}%** da matriz. "
            f"O PDE 2035 permite um acréscimo de até **+10 pontos percentuais** "
            f"(de 40% para 50%) ao longo de 10 anos.")

    with st.expander("⚙️  Premissas do cenário", expanded=True):
        c1, c2 = st.columns(2)
        teto = c1.slider("Teto de participação da hidro em 2035 (%)", 40, 60, 50, key="h_teto") / 100
        dem_2035 = c2.number_input(
            "Demanda projetada para 2035 (MWh)",
            min_value=float(socio["ee"]) * 0.5,
            value=float(socio["ee"]),
            step=float(socio["ee"]) * 0.01,
            key="h_dem",
            help="Por padrão usa o consumo de 2025. Substitua pela demanda projetada da seção Previsão Decenal.",
        )

    hidro_2025 = HIDRO_SHARE_2025 * socio["ee"]       # MWh hoje
    hidro_2035 = teto * dem_2035                        # MWh teto
    margem = hidro_2035 - hidro_2025                     # MWh adicionais
    margem_pct = (margem / hidro_2025 * 100) if hidro_2025 else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi_card("Participação 2025", "40 %", "do BEN", HYD), unsafe_allow_html=True)
    k2.markdown(kpi_card("Geração hidro 2025", f"{_fmt(hidro_2025/1e6,2)} TWh", f"{_fmt(hidro_2025)} MWh", HYD), unsafe_allow_html=True)
    k3.markdown(kpi_card("Teto hidro 2035", f"{_fmt(hidro_2035/1e6,2)} TWh", f"a {teto*100:.0f}% da demanda", ACCENT_D), unsafe_allow_html=True)
    k4.markdown(kpi_card("Margem de expansão", f"+{_fmt(margem/1e6,2)} TWh", f"+{_fmt(margem_pct,1)}% vs. 2025", WIN if margem >= 0 else THR), unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    g1, g2 = st.columns(2)

    with g1:
        fig = base_fig("Geração hidrelétrica: hoje vs. teto 2035")
        fig.add_trace(go.Bar(
            x=["Hidro 2025", "Teto hidro 2035"], y=[hidro_2025, hidro_2035],
            marker_color=["#93c5fd", HYD],
            text=[f"{hidro_2025/1e6:.2f} TWh", f"{hidro_2035/1e6:.2f} TWh"],
            textposition="outside", hovertemplate="%{y:,.0f} MWh<extra></extra>",
        ))
        fig.update_layout(yaxis_title="MWh/ano", showlegend=False)
        st.plotly_chart(fig, use_container_width=True, key="h_bar")

    with g2:
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=teto * 100,
            delta={"reference": HIDRO_SHARE_2025 * 100, "suffix": " pp", "increasing": {"color": WIN}},
            number={"suffix": " %", "font": {"size": 40}},
            title={"text": "Participação da hidro na matriz", "font": {"size": 13}},
            gauge={
                "axis": {"range": [0, 100], "tickwidth": 1},
                "bar": {"color": HYD},
                "steps": [
                    {"range": [0, 40], "color": "#e0f2fe"},
                    {"range": [40, 50], "color": "#bae6fd"},
                    {"range": [50, 100], "color": "#f1f5f9"},
                ],
                "threshold": {"line": {"color": THR, "width": 3}, "value": 50},
            },
        ))
        fig.update_layout(height=320, margin=dict(l=20, r=20, t=50, b=10),
                          paper_bgcolor=BG_CHART, font=dict(color=TEXT_PRI))
        st.plotly_chart(fig, use_container_width=True, key="h_gauge")

    st.caption("Interpretação: a regra de **+10% em 10 anos** é tratada aqui como +10 pontos percentuais "
               "(40% → 50%). Se a leitura correta for +10% relativo, o teto seria 44%; ajuste o slider acima.")


# =======================================================================
#  ABA 5 — TERMO
# =======================================================================
def tab_termo(socio: dict):
    section_title("Potencial Termelétrico",
                  "Participação atual e como tratar a expansão térmica no planejamento")

    st.info(f"No **BEN 2026** a geração térmica respondeu por **{TERMO_SHARE_2026*100:.0f}%** da matriz. "
            "Diferente de sol, vento e hidro, a térmica não é um recurso natural com fluxo fixo: "
            "seu 'potencial' é definido por **disponibilidade de combustível, custo e metas de emissão**, "
            "não pela geografia.")

    with st.expander("⚙️  Cenário de participação térmica em 2035", expanded=True):
        c1, c2 = st.columns(2)
        alvo = c1.slider("Participação térmica alvo em 2035 (%)", 20, 55, 40, key="t_alvo") / 100
        dem_2035 = c2.number_input(
            "Demanda projetada para 2035 (MWh)",
            min_value=float(socio["ee"]) * 0.5,
            value=float(socio["ee"]),
            step=float(socio["ee"]) * 0.01,
            key="t_dem",
            help="Por padrão usa o consumo de 2025. Substitua pela demanda projetada da seção Previsão Decenal.",
        )

    termo_2026 = TERMO_SHARE_2026 * socio["ee"]
    termo_2035 = alvo * dem_2035
    var = termo_2035 - termo_2026
    var_pct = (var / termo_2026 * 100) if termo_2026 else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi_card("Participação 2026", "55 %", "do BEN", THR), unsafe_allow_html=True)
    k2.markdown(kpi_card("Geração térmica 2026", f"{_fmt(termo_2026/1e6,2)} TWh", f"{_fmt(termo_2026)} MWh", THR), unsafe_allow_html=True)
    k3.markdown(kpi_card("Cenário térmico 2035", f"{_fmt(termo_2035/1e6,2)} TWh", f"a {alvo*100:.0f}% da demanda", ACCENT_D), unsafe_allow_html=True)
    k4.markdown(kpi_card("Variação vs. 2026", f"{'+' if var>=0 else ''}{_fmt(var/1e6,2)} TWh",
                         f"{'+' if var_pct>=0 else ''}{_fmt(var_pct,1)}%", THR if var > 0 else WIN), unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    g1, g2 = st.columns([1, 1.1])

    with g1:
        fig = base_fig("Geração térmica: 2026 vs. cenário 2035")
        fig.add_trace(go.Bar(
            x=["Térmica 2026", "Cenário 2035"], y=[termo_2026, termo_2035],
            marker_color=["#fca5a5", THR],
            text=[f"{termo_2026/1e6:.2f} TWh", f"{termo_2035/1e6:.2f} TWh"],
            textposition="outside", hovertemplate="%{y:,.0f} MWh<extra></extra>",
        ))
        fig.update_layout(yaxis_title="MWh/ano", showlegend=False)
        st.plotly_chart(fig, use_container_width=True, key="t_bar")

    with g2:
        st.markdown("##### 📝 Como descrever o potencial térmico no relatório")
        st.markdown(
            f"""
- **Papel de segurança e firmeza:** a térmica é a capacidade **despachável** que garante o
  atendimento quando sol e vento não estão disponíveis. Seu valor no PDE é dado pela
  **potência firme** que oferece, não por um limite geográfico.
- **Limitado por combustível e emissões:** o teto da térmica vem da **disponibilidade e custo
  do combustível** (gás, óleo, biomassa) e das **metas de descarbonização**, não do território.
- **Tendência esperada:** num plano decenal com expansão de solar e eólica, a participação
  térmica de **55% tende a recuar** em termos relativos, mesmo que a potência instalada se
  mantenha como reserva. Ela passa de "base" para "complemento/backup".
- **Métrica recomendada:** em vez de "potencial técnico" (como sol/vento), reporte a
  térmica pela **capacidade firme disponível** e pelo **fator de capacidade de operação**
  (tipicamente baixo, pois opera na ponta e em complementação).
            """
        )

    st.caption("Diferente das renováveis, o eixo da térmica é decisão de política e economia de combustível. "
               "Use o slider para mostrar o cenário de transição que o PDE 2035 adota.")


# =======================================================================
#  ENTRY POINT
# =======================================================================
def run_potencial(page=None):
    """Chamado pelo app central. Renderiza a seção Potencial em abas."""
    st.markdown(
        f'<h1 style="font-size:32px;font-weight:800;color:{TEXT_PRI};letter-spacing:-0.6px;'
        f'margin-bottom:2px;">Potencial Energético · UTÓPIA</h1>'
        f'<p style="font-size:14px;color:{TEXT_SEC};margin-bottom:18px;">'
        f'Recursos disponíveis por fonte para o Planejamento Energético Integrado</p>',
        unsafe_allow_html=True,
    )

    # ── Cargas com tratamento de erro ─────────────────────────────
    try:
        socio = load_socio_2025(EXCEL_HIST)
    except Exception as e:
        st.error(f"Não foi possível ler os dados socioeconômicos de 2025 "
                 f"(`ENTREGA_DEMANDA.xlsx` / aba RESUMO UTÓPIA): {e}")
        return

    try:
        gsa = load_gsa_stats(EXCEL_GSA)
    except Exception as e:
        st.warning(f"Relatório do Global Solar Atlas indisponível (`GSA_Report_UTOPIA.xlsx`): {e}")
        gsa = {}

    # ── Abas ──────────────────────────────────────────────────────
    t_geo, t_eol, t_sol, t_hid, t_ter = st.tabs(
        ["🗺️ Geolocalização", "🌬️ Eólica", "☀️ Solar", "💧 Hidro", "🔥 Termo"]
    )

    with t_geo:
        tab_geolocalizacao(socio)
    with t_eol:
        tab_eolica(socio)
    with t_sol:
        tab_solar(socio, gsa)
    with t_hid:
        tab_hidro(socio)
    with t_ter:
        tab_termo(socio)