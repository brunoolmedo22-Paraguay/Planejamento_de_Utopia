"""
dash_potencial.py
======================================================================
Seção "Potencial Energético" — País de Utópia
Abas: Geolocalização · Eólica · Solar · Hidro · Termo

Arquivos esperados na MESMA pasta deste módulo (raiz do projeto):
    - ENTREGA_DEMANDA.xlsx      (aba "RESUMO UTÓPIA" -> pop/PIB/EE 2025)
    - GSA_Report_UTOPIA.xlsx    (relatório do Global Solar Atlas)
    - mapa.jpeg                 (mapa do país)
    - meanwindspeed.png         (curva % áreas mais ventosas, do Global Wind Atlas)
    - EYC_area_1_*_Annual-Energy-Production.tif  (4 rasters AEP, um por turbina)

Dependência extra para ler os .tif (adicionar ao requirements.txt):
    rasterio        (preferida)  ou  tifffile  (alternativa leve)

Entry-point chamado pelo app central:
    from dash_potencial import run_potencial
    run_potencial()
======================================================================
"""

import base64
from pathlib import Path

import numpy as np
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

# ── Recurso eólico (Global Wind Atlas — 10% das áreas mais ventosas) ─
WIND_PCT       = 10        # % das áreas mais ventosas consideradas
WIND_HEIGHT    = 100       # m (altura de referência das leituras)
WIND_PD_10PCT  = 712       # W/m²  · densidade de potência (10% mais ventoso)
WIND_VMEAN     = 9.98      # m/s   · velocidade média (10% mais ventoso)
WIND_IMG       = "meanwindspeed.png"

# Espaçamento típico de parque eólico (em diâmetros de rotor)
SEP_CROSS = 5              # transversal ao vento
SEP_DOWN  = 8              # na direção do vento

# Turbinas selecionadas no EYC + raster .tif de cada uma (AEP em GWh/ano)
TURBINES = [
    {"nome": "Generic 4.0 MW · IEC Class 1", "p_kw": 4000,  "rotor": 117, "hub": 100, "v_design": 10.0, "rho": 1.225,
     "tif": "EYC_area_1_Generic-4.0-MW---IEC-Class-1_100m_10%_Annual-Energy-Production.tif"},
    {"nome": "Generic 4.5 MW · IEC Class 2", "p_kw": 4500,  "rotor": 136, "hub": 100, "v_design": 8.5,  "rho": 1.225,
     "tif": "EYC_area_1_Generic-4.5-MW---IEC-Class-2_100m_10%_Annual-Energy-Production.tif"},
    {"nome": "Generic 4.5 MW · IEC Class 3", "p_kw": 4500,  "rotor": 150, "hub": 100, "v_design": 7.5,  "rho": 1.225,
     "tif": "EYC_area_1_Generic-4.5-MW---IEC-Class-3_100m_10%_Annual-Energy-Production.tif"},
    {"nome": "Generic 15.0 MW · Offshore",   "p_kw": 15000, "rotor": 240, "hub": 150, "v_design": 10.0, "rho": 1.225,
     "tif": "EYC_area_1_Generic-15.0-MW---Offshore_150m_10%_Annual-Energy-Production.tif"},
]

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
        "ee": float(row["EE"]),   # consumo total em kWh
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


@st.cache_data
def read_aep_tif(tif_name: str) -> dict:
    """
    Lê um raster .tif de Annual Energy Production (GWh/ano por turbina) do GWA.
    Tenta rasterio; se ausente, cai para tifffile. Devolve estatísticas e um
    array 2D reduzido para o heatmap. Retorna {} se não conseguir ler.
    """
    path = ROOT / tif_name
    if not path.exists():
        return {"erro": "arquivo não encontrado"}

    arr = None
    # 1) rasterio (trata nodata/máscara corretamente)
    try:
        import rasterio
        with rasterio.open(str(path)) as ds:
            band = ds.read(1, masked=True)
            arr = np.ma.filled(band.astype("float64"), np.nan)
            if ds.nodata is not None:
                arr = np.where(arr == ds.nodata, np.nan, arr)
    except Exception:
        # 2) tifffile (lê como matriz bruta)
        try:
            import tifffile
            arr = np.asarray(tifffile.imread(str(path)), dtype="float64")
            if arr.ndim == 3:
                arr = arr[0]
        except Exception as e:
            return {"erro": f"sem leitor de .tif ({e})"}

    arr = np.asarray(arr, dtype="float64")
    # limpa nodata/sentinelas: mantém apenas valores físicos plausíveis
    arr = np.where(np.isfinite(arr), arr, np.nan)
    arr = np.where((arr > 0) & (arr < 1e6), arr, np.nan)

    valid = arr[np.isfinite(arr)]
    if valid.size == 0:
        return {"erro": "raster vazio (use retângulo, não polígono, no EYC)"}

    stats = {
        "mean": float(np.nanmean(valid)),
        "max": float(np.nanmax(valid)),
        "min": float(np.nanmin(valid)),
        "p90": float(np.nanpercentile(valid, 90)),
        "n_pix": int(valid.size),
    }

    # reduz para heatmap (~120 px no maior lado)
    step_r = max(1, arr.shape[0] // 120)
    step_c = max(1, arr.shape[1] // 120)
    heat = arr[::step_r, ::step_c]
    return {"stats": stats, "heat": heat.tolist()}


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
        s3.markdown(kpi_card("Consumo EE 2025", f"{_fmt(socio['ee'])}", "kWh/ano", "#7c3aed"), unsafe_allow_html=True)

    # ── RESUMO DO ESTUDO DE POTENCIAL ────────────────────────────
    st.markdown("---")
    st.markdown(
        f'<h3 style="font-size:18px;font-weight:700;color:{TEXT_PRI};letter-spacing:-0.2px;'
        f'margin-bottom:4px;">🔋 Resumo do Estudo de Potencial Energético</h3>'
        f'<p style="font-size:13px;color:{TEXT_SEC};margin-bottom:18px;">'
        f'Principais indicadores de cada fonte · análise detalhada nas abas seguintes</p>',
        unsafe_allow_html=True,
    )

    css_bloco = (
        "background:#ffffff;border:1px solid #e2e8f0;border-radius:16px;"
        "padding:18px 20px;box-shadow:0 2px 12px rgba(14,165,233,0.05);height:100%;"
    )

    def _badge(cor, emoji, txt):
        return (f'<span style="display:inline-block;background:{cor}22;color:{cor};'
                f'border:1px solid {cor}44;border-radius:20px;padding:2px 10px;'
                f'font-size:11px;font-weight:700;margin-bottom:10px;">{emoji} {txt}</span>')

    def _linha(label, valor, sub=""):
        s = f'<div style="margin:6px 0;"><span style="font-size:11px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.04em;">{label}</span><br>'
        s += f'<span style="font-size:20px;font-weight:800;color:{TEXT_PRI};letter-spacing:-0.4px;">{valor}</span>'
        if sub:
            s += f'<span style="font-size:11px;color:{TEXT_SEC};margin-left:6px;">{sub}</span>'
        s += '</div>'
        return s

    hidro_2025_twh = HIDRO_SHARE_2025 * socio["ee"] / 1e6
    hidro_44_twh   = 0.44 * socio["ee"] / 1e6
    margem_twh      = hidro_44_twh - hidro_2025_twh

    # Tenta ler AEP dos tifs para resumo eólico
    aep_vals = []
    for t in TURBINES:
        r = read_aep_tif(t["tif"])
        if "stats" in r:
            aep_vals.append(r["stats"]["mean"])

    # Row 1: Eólica | Solar
    col_e, col_s = st.columns(2, gap="large")
    with col_e:
        # monta lista de AEPs por turbina
        if aep_vals:
            linhas_aep = "".join(
                f'<div style="display:flex;justify-content:space-between;align-items:center;'
                f'border-bottom:1px solid #f1f5f9;padding:5px 0;">'
                f'<span style="font-size:11px;color:{TEXT_SEC};">{t["nome"].split("·")[0].strip()}</span>'
                f'<span style="font-size:13px;font-weight:700;color:{WIN};">{_fmt(v,1)} GWh/ano</span></div>'
                for t, v in zip(TURBINES, aep_vals)
            )
        else:
            linhas_aep = f'<p style="font-size:12px;color:{TEXT_SEC};">Arquivos .tif não encontrados</p>'

        st.markdown(
            f'<div style="{css_bloco}">'
            f'{_badge(WIN, "🌬️", "Eólica")}'
            f'{_linha("Densidade de potência", f"{_fmt(WIND_PD_10PCT)} W/m²", "10% mais ventoso")}'
            f'{_linha("Velocidade média @100 m", f"{_fmt(WIND_VMEAN,2)} m/s")}'
            f'<div style="margin-top:10px;font-size:11px;font-weight:700;color:{TEXT_SEC};text-transform:uppercase;letter-spacing:0.04em;margin-bottom:4px;">AEP médio por turbina</div>'
            f'{linhas_aep}'
            f'</div>',
            unsafe_allow_html=True,
        )

    with col_s:
        st.markdown(
            f'<div style="{css_bloco}">'
            f'{_badge(SOL, "☀️", "Solar")}'
            f'{_linha("GTI — Irradiação inclinada ótima", "5,03", "kWh/m²/dia")}'
            f'{_linha("PVOUT — Produção fotovoltaica", "4,00", "kWh/kWp/dia")}'
            f'<div style="margin-top:12px;font-size:11px;color:{TEXT_SEC};line-height:1.5;">'
            f'Inclinação ótima dos módulos: <b>15°</b> · Temperatura média: <b>24,4 °C</b><br>'
            f'GHI: 4,90 · DNI: 3,61 · DIF: 2,40 kWh/m²/dia'
            f'</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # Row 2: Hidro | Termo
    col_h, col_t = st.columns(2, gap="large")
    with col_h:
        _hidro_kWh_str = _fmt(HIDRO_SHARE_2025 * socio["ee"])
        st.markdown(
            f'<div style="{css_bloco}">'
            f'{_badge(HYD, "💧", "Hidro")}'
            f'{_linha("Geração hidro 2025", f"{_fmt(hidro_2025_twh,2)} GWh", f"40% do BEN · {_hidro_kWh_str} kWh")}'
            f'{_linha("Projeção com +10 pp (2035)", f"{_fmt(hidro_44_twh,2)} GWh", "44% da demanda futura")}'
            f'<div style="display:flex;align-items:center;gap:8px;margin-top:10px;">'
            f'<div style="background:{HYD}22;border:1px solid {HYD}44;border-radius:10px;padding:8px 14px;flex:1;text-align:center;">'
            f'<div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;">Margem de expansão</div>'
            f'<div style="font-size:18px;font-weight:800;color:{HYD};">+{_fmt(margem_twh,2)} TWh</div>'
            f'<div style="font-size:11px;color:{TEXT_SEC};">ao longo de 10 anos</div>'
            f'</div></div></div>',
            unsafe_allow_html=True,
        )

    with col_t:
        st.markdown(
            f'<div style="{css_bloco}">'
            f'{_badge(THR, "🔥", "Termo")}'
            f'{_linha("Participação atual (BEN 2026)", "55 %", "da matriz")}'
            f'<div style="margin-top:12px;padding:12px;background:#fff7f7;border:1px solid #fecaca;'
            f'border-radius:10px;font-size:12px;color:#7f1d1d;line-height:1.65;">'
            f'O potencial termelétrico <b>depende da política energética</b>. '
            f'Reduzir, pressionar, abandonar ou aumentar a participação da térmica são '
            f'<b>decisões estratégicas</b>, uma vez que consideramos a entrada de combustível '
            f'como estável. Não há limite geográfico — o teto é político e econômico.'
            f'</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.caption("Análise detalhada de cada fonte disponível nas abas 🌬️ Eólica · ☀️ Solar · 💧 Hidro · 🔥 Termo")


# =======================================================================
#  ABA 2 — EÓLICA
# =======================================================================
def _turbine_table_html() -> str:
    """Tabela de especificações das 4 turbinas em HTML estilizado."""
    linhas = [
        ("Potência nominal", "p_kw", "kW", 0),
        ("Diâmetro do rotor", "rotor", "m", 0),
        ("Altura do cubo", "hub", "m", 0),
        ("Vel. média de projeto", "v_design", "m/s", 1),
        ("Densidade do ar", "rho", "kg/m³", 3),
    ]
    th = ("padding:8px 10px;font-size:11px;font-weight:700;color:#fff;"
          "background:linear-gradient(135deg,#059669,#10b981);text-align:center;")
    td = "padding:7px 10px;font-size:12px;color:#0f172a;text-align:center;border-bottom:1px solid #e2e8f0;"
    tdl = ("padding:7px 10px;font-size:11px;font-weight:600;color:#475569;text-align:left;"
           "border-bottom:1px solid #e2e8f0;background:#f8fafc;text-transform:uppercase;letter-spacing:0.03em;")

    head = f'<th style="{th};text-align:left;border-top-left-radius:10px;">Especificação</th>'
    for i, t in enumerate(TURBINES):
        radius = "border-top-right-radius:10px;" if i == len(TURBINES) - 1 else ""
        head += f'<th style="{th}{radius}">{t["nome"]}</th>'

    body = ""
    for lbl, key, unit, dec in linhas:
        body += f'<tr><td style="{tdl}">{lbl} ({unit})</td>'
        for t in TURBINES:
            body += f'<td style="{td}">{_fmt(t[key], dec)}</td>'
        body += "</tr>"

    return (f'<div style="border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;'
            f'box-shadow:0 2px 10px rgba(16,185,129,0.06);"><table style="border-collapse:collapse;width:100%;">'
            f"<tr>{head}</tr>{body}</table></div>")


def _turbine_cell(t: dict, socio: dict):
    """Renderiza a análise do .tif de UMA turbina dentro de uma célula do grid."""
    st.markdown(
        f'<div style="font-size:15px;font-weight:700;color:{TEXT_PRI};margin-bottom:1px;">{t["nome"]}</div>'
        f'<div style="font-size:11px;color:{TEXT_SEC};margin-bottom:8px;">'
        f'{_fmt(t["p_kw"]/1000,1)} MW · rotor {t["rotor"]} m · cubo {t["hub"]} m</div>',
        unsafe_allow_html=True,
    )

    res = read_aep_tif(t["tif"])
    if "erro" in res:
        st.warning(f"`.tif` indisponível: {res['erro']}\n\nArquivo: `{t['tif']}`")
        return

    s = res["stats"]
    aep_med = s["mean"]                       # GWh/ano por turbina
    aep_max = s["max"]

    # Fator de capacidade derivado do AEP médio
    cf = aep_med * 1e6 / (t["p_kw"] * 8760)   # GWh->kWh / (kW * h)
    nota_unidade = ""
    if cf > 0.75:                              # AEP provavelmente em MWh -> corrige
        aep_med, aep_max = aep_med / 1000, aep_max / 1000
        cf = cf / 1000
        nota_unidade = " (valores reescalados de MWh→GWh)"

    # Quantas turbinas cabem nos 10% mais ventosos (1.200 km²) e potencial total
    usable_km2 = AREA_KM2 * WIND_PCT / 100
    area_por_turb = (SEP_CROSS * t["rotor"]) * (SEP_DOWN * t["rotor"]) / 1e6   # km²
    n_turb = usable_km2 / area_por_turb
    total_gwh = n_turb * aep_med
    cobertura = total_gwh * 1e3 / socio["ee"] * 100 if socio["ee"] else 0

    m1, m2, m3 = st.columns(3)
    m1.markdown(kpi_card("AEP médio/turbina", f"{_fmt(aep_med,1)}", "GWh/ano", WIN), unsafe_allow_html=True)
    m2.markdown(kpi_card("Fator de capacidade", f"{_fmt(cf*100,1)} %", "no recurso 10%", "#047857"), unsafe_allow_html=True)
    m3.markdown(kpi_card("Potencial da área", f"{_fmt(total_gwh/1000,2)} TWh", f"≈ {_fmt(n_turb,0)} turbinas", ACCENT_D), unsafe_allow_html=True)

    # Heatmap do raster
    heat = np.array(res["heat"], dtype="float64")
    fig = go.Figure(go.Heatmap(
        z=heat, colorscale="YlGn", showscale=True,
        colorbar=dict(title="GWh", thickness=10, len=0.85),
        hovertemplate="AEP: %{z:.1f} GWh/ano<extra></extra>",
    ))
    fig.update_layout(
        height=240, margin=dict(l=6, r=6, t=28, b=6),
        title=dict(text="AEP por posição (raster .tif)", font=dict(size=12, color=TEXT_PRI, weight="bold"), x=0),
        paper_bgcolor=BG_CHART, plot_bgcolor=BG_CHART,
        xaxis=dict(showticklabels=False, showgrid=False, zeroline=False),
        yaxis=dict(showticklabels=False, showgrid=False, zeroline=False, autorange="reversed", scaleanchor="x"),
    )
    st.plotly_chart(fig, use_container_width=True, key=f"tif_{t['p_kw']}_{t['rotor']}")
    st.caption(f"AEP máx {_fmt(aep_max,1)} GWh · {_fmt(s['n_pix'],0)} px válidos · "
               f"espaçamento {SEP_CROSS}D×{SEP_DOWN}D{nota_unidade}")


def tab_eolica(socio: dict):
    section_title("Potencial Eólico",
                  "Recurso e produção por turbina · Global Wind Atlas (10% das áreas mais ventosas)")

    # ── KPIs de AEP por turbina (resultado principal, topo) ───────
    st.markdown(
        f'<div style="font-size:11px;font-weight:700;color:{TEXT_SEC};text-transform:uppercase;'
        f'letter-spacing:0.08em;margin-bottom:8px;">⚡ AEP médio/turbina · 10% da área mais ventosa</div>',
        unsafe_allow_html=True,
    )
    aep_cols = st.columns(4)
    for col, t in zip(aep_cols, TURBINES):
        r = read_aep_tif(t["tif"])
        if "stats" in r:
            aep = r["stats"]["mean"]
            cf  = aep * 1e6 / (t["p_kw"] * 8760)
            if cf > 0.75:
                aep = aep / 1000
            label_curto = t["nome"].split("·")[0].strip()
            col.markdown(
                kpi_card(label_curto, f"{_fmt(aep,1)}", "GWh/ano", WIN),
                unsafe_allow_html=True,
            )
        else:
            col.markdown(
                kpi_card(t["nome"].split("·")[0].strip(), "—", ".tif não encontrado"),
                unsafe_allow_html=True,
            )

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    st.markdown("---")

    # ── KPIs do recurso (números reais do GWA) ────────────────────
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi_card("Densidade de potência", f"{_fmt(WIND_PD_10PCT)}", f"W/m² · 10% mais ventoso", WIN), unsafe_allow_html=True)
    k2.markdown(kpi_card("Velocidade média", f"{_fmt(WIND_VMEAN,2)}", "m/s", "#047857"), unsafe_allow_html=True)
    k3.markdown(kpi_card("Altura de referência", f"{WIND_HEIGHT} m", "altura do cubo"), unsafe_allow_html=True)
    k4.markdown(kpi_card("Área avaliada", f"{_fmt(AREA_KM2*WIND_PCT/100)}", f"km² · {WIND_PCT}% de {_fmt(AREA_KM2)}", ACCENT_D), unsafe_allow_html=True)

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

    # ── Tabela de turbinas + curva % áreas mais ventosas ──────────
    col_tab, col_img = st.columns([1.4, 1], gap="large")
    with col_tab:
        st.markdown("##### Turbinas avaliadas no Energy Yield Calculator")
        st.markdown(_turbine_table_html(), unsafe_allow_html=True)
    with col_img:
        st.markdown("##### Velocidade vs. % de área mais ventosa")
        img = _img_b64(WIND_IMG)
        if img:
            st.markdown(
                f'<div style="border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;padding:8px;background:#fff;">'
                f'<img src="{img}" style="width:100%;display:block;" /></div>',
                unsafe_allow_html=True,
            )
        else:
            st.warning(f"Arquivo `{WIND_IMG}` não encontrado na raiz do projeto.")

    st.markdown("---")
    st.markdown("##### Análise dos rasters de produção anual (AEP) por turbina")

    # ── Grid 2×2: uma turbina por célula ──────────────────────────
    r1c1, r1c2 = st.columns(2, gap="large")
    with r1c1:
        _turbine_cell(TURBINES[0], socio)
    with r1c2:
        _turbine_cell(TURBINES[1], socio)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    r2c1, r2c2 = st.columns(2, gap="large")
    with r2c1:
        _turbine_cell(TURBINES[2], socio)
    with r2c2:
        _turbine_cell(TURBINES[3], socio)

    st.markdown("---")
    st.caption("AEP lido diretamente dos .tif do GWA (GWh/ano por turbina). "
               f"Potencial da área = AEP médio × nº de turbinas que cabem em {_fmt(AREA_KM2*WIND_PCT/100)} km² "
               f"(espaçamento {SEP_CROSS}D×{SEP_DOWN}D). É potencial técnico bruto, sem restrições de uso do solo.")


# =======================================================================
#  ABA 3 — SOLAR
# =======================================================================
def tab_solar(socio: dict, gsa: dict):
    section_title("Potencial Solar",
                  "Recurso medido pelo Global Solar Atlas para a área do país (12.000 km²)")

    # ── def avg local ──────────────────────────────────────────────
    def avg(k):
        return gsa.get(k, {}).get("stats", {}).get("Average")

    # ── Resultados principais em destaque (topo) ──────────────────
    st.markdown(
        f'<div style="font-size:11px;font-weight:700;color:{TEXT_SEC};text-transform:uppercase;'
        f'letter-spacing:0.08em;margin-bottom:8px;">☀️ Indicadores principais do recurso solar</div>',
        unsafe_allow_html=True,
    )
    hero1, hero2, spacer1, spacer2 = st.columns([1, 1, 0.6, 0.6])

    # GTI — hero card maior
    gti_val = avg("GTI") or 5.03
    pvout_val = avg("PVOUT") or 4.00
    hero_css = ("background:linear-gradient(135deg,#fffbeb,#fef3c7);border:1.5px solid #fcd34d;"
                "border-radius:16px;padding:18px 22px;box-shadow:0 4px 16px rgba(245,158,11,0.12);")
    hero1.markdown(
        f'<div style="{hero_css}">'
        f'<div style="font-size:10px;font-weight:700;color:#92400e;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:4px;">GTI — Irradiação inclinada ótima</div>'
        f'<div style="font-size:36px;font-weight:900;color:{SOL};letter-spacing:-1px;line-height:1.1;">{gti_val:.2f}</div>'
        f'<div style="font-size:13px;color:#78350f;font-weight:600;">kWh/m²/dia</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    hero2.markdown(
        f'<div style="{hero_css}">'
        f'<div style="font-size:10px;font-weight:700;color:#92400e;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:4px;">PVOUT — Produção fotovoltaica</div>'
        f'<div style="font-size:36px;font-weight:900;color:#d97706;letter-spacing:-1px;line-height:1.1;">{pvout_val:.2f}</div>'
        f'<div style="font-size:13px;color:#78350f;font-weight:600;">kWh/kWp/dia</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    spacer1.markdown(kpi_card("Inclinação ótima", f"{_fmt(avg('OPTA') or 15,0)}°", "dos módulos"), unsafe_allow_html=True)
    spacer2.markdown(kpi_card("Temperatura média", f"{_fmt(avg('TEMP') or 24.4,1)} °C", "do ar"), unsafe_allow_html=True)

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    st.markdown("---")

    # ── KPIs do recurso (médias do GSA) ───────────────────────────
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
                  "Participação atual na matriz e margem de expansão para o PDE 2035")

    hidro_2025 = HIDRO_SHARE_2025 * socio["ee"]    # kWh hoje      (40%)
    TETO_HIDRO = 0.44                               # teto fixo: +10% relativo sobre 40%
    hidro_2035 = TETO_HIDRO * socio["ee"]           # kWh teto 2035 (projeção p/ mesma demanda)
    margem     = hidro_2035 - hidro_2025
    margem_pct = (margem / hidro_2025 * 100) if hidro_2025 else 0
    outros_2025 = socio["ee"] * 0.05               # 5% restante da matriz 2025

    # ── KPIs principais ───────────────────────────────────────────
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi_card("Participação atual (2025)", "40 %", "do BEN · base para projeção", HYD), unsafe_allow_html=True)
    k2.markdown(kpi_card("Geração hidro 2025", f"{_fmt(hidro_2025/1e6,2)} GWh", f"{_fmt(hidro_2025)} KWh/ano", HYD), unsafe_allow_html=True)
    k3.markdown(kpi_card("Projeção 2035 (+10%)", f"{_fmt(hidro_2035/1e6,2)} GWh", f"44% da demanda futura · {_fmt(hidro_2035)} kWh", ACCENT_D), unsafe_allow_html=True)
    k4.markdown(kpi_card("Margem de expansão", f"+{_fmt(margem/1e6,3)} GWh", f"+{_fmt(margem_pct,1)}% em relação a 2025", WIN), unsafe_allow_html=True)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    g1, g2, g3 = st.columns(3)

    # Gráfico 1 — barras 2025 vs 2035
    with g1:
        fig = base_fig("Geração hidrelétrica: 2025 → 2035", height=300)
        fig.add_trace(go.Bar(
            x=["Hidro 2025", "Projeção 2035"],
            y=[hidro_2025, hidro_2035],
            marker_color=["#93c5fd", HYD],
            text=[f"{hidro_2025/1e6:.2f} GWh", f"{hidro_2035/1e6:.2f} GWh"],
            textposition="outside",
            hovertemplate="%{y:,.0f} kWh<extra></extra>",
        ))
        fig.update_layout(yaxis_title="kWh/ano", showlegend=False)
        st.plotly_chart(fig, use_container_width=True, key="h_bar")

    # Gráfico 2 — gauge participação
    with g2:
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=TETO_HIDRO * 100,
            delta={"reference": HIDRO_SHARE_2025 * 100, "suffix": " pp",
                   "increasing": {"color": HYD}, "font": {"size": 14}},
            number={"suffix": " %", "font": {"size": 44, "color": HYD}},
            title={"text": "Participação hidro na matriz", "font": {"size": 13, "color": TEXT_PRI}},
            gauge={
                "axis": {"range": [0, 60], "tickwidth": 1,
                         "tickvals": [0, 10, 20, 30, 40, 44, 50, 60],
                         "ticktext": ["0", "10", "20", "30", "40", "<b>44</b>", "50", "60"]},
                "bar": {"color": HYD, "thickness": 0.28},
                "bgcolor": "#f0f9ff",
                "steps": [
                    {"range": [0, 40], "color": "#e0f2fe"},
                    {"range": [40, 44], "color": "#bae6fd"},
                    {"range": [44, 60], "color": "#f1f5f9"},
                ],
                "threshold": {"line": {"color": ACCENT_D, "width": 3}, "value": 44},
            },
        ))
        fig.update_layout(height=300, margin=dict(l=20, r=20, t=60, b=10),
                          paper_bgcolor=BG_CHART, font=dict(color=TEXT_PRI))
        st.plotly_chart(fig, use_container_width=True, key="h_gauge")

    # Gráfico 3 — composição da matriz 2025 (pizza)
    with g3:
        fig = go.Figure(go.Pie(
            labels=["Hidro (40%)", "Termo (55%)", "Outros (5%)"],
            values=[hidro_2025, socio["ee"] * 0.55, outros_2025],
            marker_colors=[HYD, THR, "#94a3b8"],
            hole=0.52,
            textinfo="label+percent",
            textfont=dict(size=11),
            hovertemplate="%{label}<br>%{value:,.0f} kWh<extra></extra>",
            pull=[0.04, 0, 0],
        ))
        fig.update_layout(
            title=dict(text="Composição da matriz 2025", font=dict(size=13, color=TEXT_PRI, weight="bold"), x=0),
            height=300, showlegend=False,
            margin=dict(l=10, r=10, t=46, b=10),
            paper_bgcolor=BG_CHART,
        )
        st.plotly_chart(fig, use_container_width=True, key="h_pie")

    # Linha adicional — variação anual necessária
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    cr1, cr2 = st.columns(2)

    with cr1:
        # Progressão da meta 40→44% em 10 anos
        anos = list(range(2025, 2036))
        participacoes = [40 + i * 0.4 for i in range(11)]   # linear 40→44%
        kwh_anual = [p / 100 * socio["ee"] for p in participacoes]
        fig = base_fig("Trajetória de expansão da hidro (2025–2035)", height=280)
        fig.add_trace(go.Scatter(
            x=anos, y=kwh_anual, mode="lines+markers",
            line=dict(color=HYD, width=2.5, dash="dot"),
            marker=dict(size=6, color=HYD),
            fill="tozeroy", fillcolor="rgba(14,165,233,0.07)",
            hovertemplate="<b>%{x}</b>: %{y:,.0f} kWh<extra></extra>",
            name="Geração hidro"
        ))
        fig.add_trace(go.Scatter(
            x=[2025, 2035], y=[hidro_2025, hidro_2025],
            mode="lines", line=dict(color="#94a3b8", dash="dash", width=1),
            name="Base 2025", hoverinfo="skip",
        ))
        fig.update_layout(yaxis_title="kWh/ano", xaxis=dict(tickvals=anos, tickfont=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True, key="h_traj")

    with cr2:
        # Incremento anual médio necessário
        incremento_anual = margem / 10
        incremento_pct = margem_pct / 10
        css_info = ("background:#f0f9ff;border:1px solid #bae6fd;border-radius:14px;"
                    "padding:20px 22px;height:100%;box-sizing:border-box;")
        st.markdown(
            f'<div style="{css_info}">'
            f'<div style="font-size:12px;font-weight:700;color:{HYD};text-transform:uppercase;'
            f'letter-spacing:0.05em;margin-bottom:14px;">📐 Detalhamento da meta</div>'
            f'<div style="display:grid;gap:10px;">'

            f'<div style="background:#fff;border:1px solid #e0f2fe;border-radius:10px;padding:10px 14px;">'
            f'<div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.04em;">Base de cálculo</div>'
            f'<div style="font-size:15px;font-weight:700;color:{TEXT_PRI};">40% → 44% da demanda</div>'
            f'<div style="font-size:11px;color:{TEXT_SEC};">+10% relativo sobre participação atual</div></div>'

            f'<div style="background:#fff;border:1px solid #e0f2fe;border-radius:10px;padding:10px 14px;">'
            f'<div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.04em;">Incremento anual médio</div>'
            f'<div style="font-size:15px;font-weight:700;color:{HYD};">+{_fmt(incremento_anual/1e3,1)} GWh/ano</div>'
            f'<div style="font-size:11px;color:{TEXT_SEC};">≈ +{_fmt(incremento_pct,2)}% da geração hidro/ano</div></div>'

            f'<div style="background:#fff;border:1px solid #e0f2fe;border-radius:10px;padding:10px 14px;">'
            f'<div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.04em;">Expansão total 2025–2035</div>'
            f'<div style="font-size:15px;font-weight:700;color:{ACCENT_D};">+{_fmt(margem/1e6,3)} TWh</div>'
            f'<div style="font-size:11px;color:{TEXT_SEC};">{_fmt(margem)} kWh adicionais em 10 anos</div></div>'

            f'</div></div>',
            unsafe_allow_html=True,
        )

    st.caption("Teto de 44% = +10% relativo sobre os 40% do BEN 2025, projetado para a mesma demanda base de 2025. "
               "Se a demanda crescer (ver Previsão Decenal), a geração hidro necessária sobe proporcionalmente.")


# =======================================================================
#  ABA 5 — TERMO
# =======================================================================
def tab_termo(socio: dict):
    section_title("Potencial Termelétrico",
                  "Participação atual e cenário de transição para o PDE 2035")

    with st.expander("⚙️  Cenário de participação térmica em 2035", expanded=True):
        c1, c2 = st.columns(2)
        alvo = c1.slider("Participação térmica alvo em 2035 (%)", 20, 55, 40, key="t_alvo") / 100
        dem_2035 = c2.number_input(
            "Demanda projetada para 2035 (kWh)",
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
    k2.markdown(kpi_card("Geração térmica 2026", f"{_fmt(termo_2026/1e6,2)} TWh", f"{_fmt(termo_2026)} kWh", THR), unsafe_allow_html=True)
    k3.markdown(kpi_card("Cenário térmico 2035", f"{_fmt(termo_2035/1e6,2)} TWh", f"a {alvo*100:.0f}% da demanda", ACCENT_D), unsafe_allow_html=True)
    k4.markdown(kpi_card("Variação vs. 2026", f"{'+' if var>=0 else ''}{_fmt(var/1e6,2)} TWh",
                         f"{'+' if var_pct>=0 else ''}{_fmt(var_pct,1)}%", THR if var > 0 else WIN), unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    g1, g2 = st.columns([1.1, 1], gap="large")

    with g1:
        fig = base_fig("Geração térmica: 2026 vs. cenário 2035")
        fig.add_trace(go.Bar(
            x=["Térmica 2026", "Cenário 2035"], y=[termo_2026, termo_2035],
            marker_color=["#fca5a5", THR],
            text=[f"{termo_2026/1e6:.2f} TWh", f"{termo_2035/1e6:.2f} TWh"],
            textposition="outside", hovertemplate="%{y:,.0f} kWh<extra></extra>",
        ))
        fig.update_layout(yaxis_title="kWh/ano", showlegend=False)
        st.plotly_chart(fig, use_container_width=True, key="t_bar")

    with g2:
        direcoes = [
            ("⬇️ Reduzir",   "#10b981", "Priorizar renováveis; térmica recua para backup de pico."),
            ("⚖️ Pressionar", "#f59e0b", "Manter participação como segurança de suprimento."),
            ("🚫 Abandonar",  "#0ea5e9", "Eliminar gradual via fontes limpas (>10 anos)."),
            ("⬆️ Aumentar",   "#ef4444", "Atender crescimento rápido de carga com segurança imediata."),
        ]
        css_dir = ("background:#fff7ed;border:1px solid #fed7aa;border-radius:14px;"
                   "padding:16px 18px;")
        st.markdown(
            f'<div style="{css_dir}">'
            f'<div style="font-size:11px;font-weight:700;color:#7c2d12;text-transform:uppercase;'
            f'letter-spacing:0.05em;margin-bottom:12px;">🏛️ O potencial depende da política energética</div>'
            + "".join(
                f'<div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:9px;">'
                f'<span style="min-width:90px;font-size:12px;font-weight:700;color:{cor};">{icone}</span>'
                f'<span style="font-size:12px;color:{TEXT_PRI};line-height:1.5;">{desc}</span></div>'
                for icone, cor, desc in direcoes
            )
            + f'<div style="margin-top:10px;padding-top:10px;border-top:1px solid #fed7aa;'
            f'font-size:11px;color:#92400e;line-height:1.5;">'
            f'Considerando o <b>suprimento de combustível estável</b>, reduzir, pressionar, '
            f'abandonar ou aumentar são <b>decisões estratégicas</b>, não limites físicos.'
            f'</div></div>',
            unsafe_allow_html=True,
        )

    st.caption(f"Participação de 55% no BEN 2026. Use o slider para projetar o papel da térmica no PDE 2035.")


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
