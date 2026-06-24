"""
dash_potencial.py
======================================================================
Seção "Potencial Energético" — País de Utópia
Abas: Geolocalização · Eólica · Solar · Hidro · Termo

Arquivos esperados na MESMA pasta deste módulo (raiz do projeto):
    - ENTREGA_DEMANDA_utopia.xlsx      (aba "RESUMO UTÓPIA" -> pop/PIB/EE 2025)
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

EXCEL_HIST = str(ROOT / "ENTREGA_DEMANDA_utopia.xlsx")
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
    # Wrapper height:100%+flex faz todos os cards da row terem mesma altura
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
        title=dict(text=title, font=dict(size=13, color=TEXT_PRI, weight="bold"), x=0, xanchor="left", pad=dict(l=4, t=4)),
        height=height, showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        **THEME,
    )
    return fig


# =======================================================================
#  CARGA DE DADOS
# =======================================================================
@st.cache_data(ttl=300)
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


@st.cache_data(ttl=300)
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


@st.cache_data(ttl=300)
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


@st.cache_data(ttl=300)
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
        s3.markdown(kpi_card("Consumo EE 2025", f"{_fmt(socio['ee'])}", "MWh/ano", "#7c3aed"), unsafe_allow_html=True)

    st.markdown("---")

    # ── Resumo do potencial por fonte ─────────────────────────────
    st.markdown(
        f'<h3 style="font-size:16px;font-weight:700;color:{TEXT_PRI};margin:0 0 14px;">'        'Resumo do Estudo de Potencial Energético</h3>',
        unsafe_allow_html=True,
    )

    # Row 1 — Eólica + Solar
    r1c1, r1c2 = st.columns(2, gap="large")

    with r1c1:
        st.markdown(
            f'<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:16px;padding:18px 20px;">'            f'<div style="font-size:13px;font-weight:700;color:#047857;margin-bottom:10px;">🌬️ Eólica</div>'            f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">'            f'<div><div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.05em;">AEP médio (4,0 MW Class 1)</div><div style="font-size:16px;font-weight:700;color:#10b981;">18,8 GWh/ano</div></div>'            f'<div><div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.05em;">AEP médio (4,5 MW Class 2)</div><div style="font-size:16px;font-weight:700;color:#10b981;">~AEP tif</div></div>'            f'<div><div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.05em;">Densidade de potência</div><div style="font-size:16px;font-weight:700;color:#047857;">712 W/m²</div></div>'            f'<div><div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.05em;">Velocidade média</div><div style="font-size:16px;font-weight:700;color:#047857;">9,98 m/s</div></div>'            f'</div></div>',
            unsafe_allow_html=True,
        )

    with r1c2:
        st.markdown(
            f'<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:16px;padding:18px 20px;">'            f'<div style="font-size:13px;font-weight:700;color:#b45309;margin-bottom:10px;">☀️ Solar</div>'            f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">'            f'<div><div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.05em;">GTI (inclinada ótima)</div><div style="font-size:16px;font-weight:700;color:{SOL};">5,03 kWh/m²/dia</div></div>'            f'<div><div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.05em;">PVOUT (produção FV)</div><div style="font-size:16px;font-weight:700;color:#d97706;">4,00 kWh/kWp/dia</div></div>'            f'</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # Row 2 — Hidro + Termo
    r2c1, r2c2 = st.columns(2, gap="large")

    with r2c1:
        st.markdown(
            f'<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:16px;padding:18px 20px;">'            f'<div style="font-size:13px;font-weight:700;color:#1d4ed8;margin-bottom:10px;">💧 Hidro</div>'            f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">'            f'<div><div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.05em;">Participação 2025</div><div style="font-size:16px;font-weight:700;color:{HYD};">44 %</div></div>'            f'<div><div style="font-size:10px;color:{TEXT_SEC};font-weight:600;text-transform:uppercase;letter-spacing:0.05em;">Projeção máxima 2035</div><div style="font-size:16px;font-weight:700;color:{ACCENT_D};">+ 10 pp</div></div>'            f'<div style="grid-column:1/-1;font-size:11px;color:{TEXT_SEC};margin-top:4px;">Margem de expansão: até +10 pontos percentuais ao longo da década.</div>'            f'</div></div>',
            unsafe_allow_html=True,
        )

    with r2c2:
        st.markdown(
            f'<div style="background:#fff1f2;border:1px solid #fecdd3;border-radius:16px;padding:18px 20px;">'            f'<div style="font-size:13px;font-weight:700;color:#b91c1c;margin-bottom:10px;">🔥 Térmica</div>'            f'<div style="font-size:12px;color:{TEXT_SEC};line-height:1.6;">'            f'O potencial termelétrico depende da <strong>política energética</strong>. '            f'O interesse em <em>reduzir, pressionar, abandonar, aumentar</em> são decisões estratégicas — '            f'uma vez que consideramos a entrada de combustível como estável, '            f'a térmica é definida por escolha, não por recurso.'            f'</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.caption("As dimensões definem a área de referência (12.000 km²) usada nas estimativas de potencial "
               "solar e eólico das abas seguintes.")


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

    # ── AEP médio das 4 turbinas — lido dos .tif (resultado principal) ──
    _aep_vals = []
    for _t in TURBINES:
        _r = read_aep_tif(_t["tif"])
        if "stats" in _r:
            _v = _r["stats"]["mean"]
            _cf_chk = _v * 1e6 / (_t["p_kw"] * 8760)
            if _cf_chk > 0.75:
                _v = _v / 1000
            _aep_vals.append(_v)
        else:
            _aep_vals.append(None)

    _cells = ""
    for _t, _v in zip(TURBINES, _aep_vals):
        _nome_curto = _t["nome"].replace("Generic ", "").replace(" · ", " ")
        _val_str = f"{_fmt(_v, 1)}" if _v is not None else "—"
        _cells += (
            f'<div style="text-align:center;">'
            f'<div style="font-size:11px;color:{TEXT_SEC};font-weight:600;margin-bottom:2px;">{_nome_curto}</div>'
            f'<div style="font-size:22px;font-weight:800;color:#10b981;">{_val_str}</div>'
            f'<div style="font-size:10px;color:{TEXT_SEC};">GWh/ano</div></div>'
        )

    st.markdown(
        f'<div style="background:linear-gradient(135deg,#f0fdf4,#dcfce7);border:1px solid #bbf7d0;'
        f'border-radius:14px;padding:12px 18px;margin-bottom:14px;">'
        f'<div style="font-size:11px;font-weight:700;color:#047857;text-transform:uppercase;'
        f'letter-spacing:0.06em;margin-bottom:8px;">⚡ AEP Médio por Turbina — resultado principal</div>'
        f'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;">'
        f'{_cells}'
        f'</div></div>',
        unsafe_allow_html=True,
    )

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

    # ── Banner resultado principal ────────────────────────────────
    st.markdown(
        f'<div style="background:linear-gradient(135deg,#fffbeb,#fef9c3);border:1px solid #fde68a;'
        f'border-radius:14px;padding:12px 18px;margin-bottom:14px;">'
        f'<div style="font-size:11px;font-weight:700;color:#b45309;text-transform:uppercase;'
        f'letter-spacing:0.06em;margin-bottom:8px;">☀️ Resultado Principal — Global Solar Atlas</div>'
        f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;">'
        f'<div style="text-align:center;">'
        f'<div style="font-size:11px;color:{TEXT_SEC};font-weight:600;">GTI · Irradiação inclinada ótima</div>'
        f'<div style="font-size:28px;font-weight:800;color:{SOL};">5,03</div>'
        f'<div style="font-size:12px;color:{TEXT_SEC};font-weight:600;">kWh/m²/dia</div></div>'
        f'<div style="text-align:center;">'
        f'<div style="font-size:11px;color:{TEXT_SEC};font-weight:600;">PVOUT · Produção fotovoltaica específica</div>'
        f'<div style="font-size:28px;font-weight:800;color:#d97706;">4,00</div>'
        f'<div style="font-size:12px;color:{TEXT_SEC};font-weight:600;">kWh/kWp/dia</div></div>'
        f'</div></div>',
        unsafe_allow_html=True,
    )

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

    # ── Simulador de usina FV fictícia ───────────────────────────
    pvout_dia = avg("PVOUT") or 4.0
    gti_dia   = avg("GTI")   or 5.03
    pvout_ano = pvout_dia * 365     # kWh/kWp/ano
    EFF_MOD   = 0.21                # eficiência do módulo assumida (21%)

    st.markdown(
        f'<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:14px;'        f'padding:14px 18px;margin-bottom:14px;">'        f'<div style="font-size:12px;font-weight:700;color:#b45309;text-transform:uppercase;'        f'letter-spacing:0.05em;margin-bottom:10px;">⚙️  Simulador de Usina Fotovoltaica Fictícia</div>'        f'<div style="font-size:12px;color:{TEXT_SEC};margin-bottom:4px;">'        f'Insira a capacidade da usina — os resultados são calculados com o recurso solar local '        f'(PVOUT = {pvout_dia:.2f} kWh/kWp/dia · GTI = {gti_dia:.2f} kWh/m²/dia · η = {EFF_MOD*100:.0f}%).'        f'</div></div>',
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns([1, 2])
    with c1:
        kwp_usina = st.number_input(
            "Capacidade da usina (kWp)",
            min_value=1.0,
            max_value=1_000_000.0,
            value=1_000.0,
            step=100.0,
            format="%.0f",
            key="s_kwp",
            help="Digite o pico de potência da usina em kWp (ex: 1000 = 1 MWp)",
        )

    # ── Cálculos da usina ─────────────────────────────────────────
    mwp_usina    = kwp_usina / 1_000                          # MWp
    e_dia_kwh    = kwp_usina * pvout_dia                      # kWh/dia
    e_ano_mwh    = kwp_usina * pvout_ano / 1_000              # MWh/ano
    e_ano_gwh    = e_ano_mwh / 1_000                          # GWh/ano
    # Área: P_pico = GTI_dia * A * EFF  →  A = kWp/(GTI_dia_kW * EFF)
    # GTI em kWh/m²/dia → energia/dia por m² em kWh; kWp instalado = GTI*η*A
    area_m2      = (kwp_usina * 1_000) / (gti_dia * 1_000 * EFF_MOD)  # m²
    area_ha      = area_m2 / 10_000
    # Fator de capacidade: E_ano / (kWp * 8760 h)
    fc           = (kwp_usina * pvout_ano) / (kwp_usina * 8_760) * 100  # %
    cobertura    = e_ano_mwh / socio["ee"] * 100 if socio["ee"] else 0

    with c2:
        st.markdown(
            f'<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;">'            f'<div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:10px 12px;">'            f'<div style="font-size:9px;font-weight:700;color:{TEXT_SEC};text-transform:uppercase;letter-spacing:0.05em;">Geração diária</div>'            f'<div style="font-size:18px;font-weight:800;color:{SOL};">{e_dia_kwh:,.1f}</div>'            f'<div style="font-size:10px;color:{TEXT_SEC};">kWh/dia</div></div>'            f'<div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:10px 12px;">'            f'<div style="font-size:9px;font-weight:700;color:{TEXT_SEC};text-transform:uppercase;letter-spacing:0.05em;">Geração anual</div>'            f'<div style="font-size:18px;font-weight:800;color:#d97706;">{e_ano_gwh:,.3f}</div>'            f'<div style="font-size:10px;color:{TEXT_SEC};">GWh/ano</div></div>'            f'<div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:10px 12px;">'            f'<div style="font-size:9px;font-weight:700;color:{TEXT_SEC};text-transform:uppercase;letter-spacing:0.05em;">Área necessária</div>'            f'<div style="font-size:18px;font-weight:800;color:{ACCENT_D};">{area_ha:,.2f}</div>'            f'<div style="font-size:10px;color:{TEXT_SEC};">hectares (η=21%)</div></div>'            f'<div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:10px 12px;">'            f'<div style="font-size:9px;font-weight:700;color:{TEXT_SEC};text-transform:uppercase;letter-spacing:0.05em;">Fator de capacidade</div>'            f'<div style="font-size:18px;font-weight:800;color:#7c3aed;">{fc:.1f}</div>'            f'<div style="font-size:10px;color:{TEXT_SEC};">%</div></div>'            f'<div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:10px 12px;">'            f'<div style="font-size:9px;font-weight:700;color:{TEXT_SEC};text-transform:uppercase;letter-spacing:0.05em;">Capacidade</div>'            f'<div style="font-size:18px;font-weight:800;color:{SOL};">{mwp_usina:,.2f}</div>'            f'<div style="font-size:10px;color:{TEXT_SEC};">MWp</div></div>'            f'<div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:10px 12px;">'            f'<div style="font-size:9px;font-weight:700;color:{TEXT_SEC};text-transform:uppercase;letter-spacing:0.05em;">Cobertura demanda</div>'            f'<div style="font-size:18px;font-weight:800;color:{ACCENT};">{cobertura:.3f}</div>'            f'<div style="font-size:10px;color:{TEXT_SEC};">% do consumo 2025</div></div>'            f'</div>',
            unsafe_allow_html=True,
        )

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
                  "Participação atual e projeção da geração hidrelétrica até 2035")

    HIDRO_2025_PCT = 0.40
    HIDRO_2035_PCT = 0.44

    CSV_PROJ = str(ROOT / "projecoes_demanda.csv")

    def _dem2035(cenario_key):
        try:
            df_proj = pd.read_csv(CSV_PROJ)
            df_proj["Ano"] = df_proj["Ano"].astype(int)
            row = df_proj[
                (df_proj["Ano"] == 2035) &
                (df_proj["Cenario"] == cenario_key) &
                (df_proj["Local"] == "Utopia")
            ]
            if not row.empty:
                return float(row["EE_TOTAL"].iloc[0])
        except Exception:
            pass
        return socio["ee"]

    CEN_OPTS = {
        "Referência":     "Referencia",
        "Alto (PIB +3%)": "Alto",
        "Baixo (PIB -3%)":"Baixo",
    }
    CEN_CORES = {
        "Referencia": ACCENT,
        "Alto":       "#22c55e",
        "Baixo":      "#f59e0b",
    }

    cen_lbl = st.radio(
        "📊 Cenário de demanda 2035",
        list(CEN_OPTS.keys()),
        horizontal=True,
        key="h_cen_sel",
    )
    cen_key = CEN_OPTS[cen_lbl]
    cen_cor = CEN_CORES[cen_key]

    dem_2035   = _dem2035(cen_key)
    hidro_2025 = HIDRO_2025_PCT * socio["ee"]
    hidro_2035 = HIDRO_2035_PCT * dem_2035
    margem     = hidro_2035 - hidro_2025
    margem_pct = (margem / hidro_2025 * 100) if hidro_2025 else 0

    # ── Row 1: KPIs ───────────────────────────────────────────────
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi_card("Participação 2025",    "40 %",
                         "do BEN 2025", HYD), unsafe_allow_html=True)
    k2.markdown(kpi_card("Geração hidro 2025",   f"{_fmt(hidro_2025/1e6,2)} TWh",
                         f"{_fmt(hidro_2025)} MWh", HYD), unsafe_allow_html=True)
    k3.markdown(kpi_card("Participação 2035",    "44 %",
                         "teto fixo", ACCENT_D), unsafe_allow_html=True)
    k4.markdown(kpi_card("Geração hidro 2035",   f"{_fmt(hidro_2035/1e6,2)} TWh",
                         "44% × demanda 2035", ACCENT_D), unsafe_allow_html=True)

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    k5, k6, k7, k8 = st.columns(4)
    k5.markdown(kpi_card("Expansão necessária",  f"+{_fmt(margem/1e6,2)} TWh",
                         "acréscimo 2025 → 2035",
                         WIN if margem >= 0 else THR), unsafe_allow_html=True)
    k6.markdown(kpi_card("Crescimento relativo", f"+{_fmt(margem_pct,1)} %",
                         "vs. geração hidro 2025",
                         WIN if margem >= 0 else THR), unsafe_allow_html=True)
    k7.markdown(kpi_card("Demanda total 2035",   f"{_fmt(dem_2035/1e6,2)} TWh",
                         f"cenário {cen_lbl}", cen_cor), unsafe_allow_html=True)
    k8.markdown(kpi_card("Fator de firmeza",     "Alto",
                         "fonte despachável 24/7", HYD), unsafe_allow_html=True)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        fig = base_fig("Geração hidrelétrica: 2025 vs. 2035", height=300)
        fig.add_trace(go.Bar(
            x=["Hidro 2025", "Hidro 2035"],
            y=[hidro_2025, hidro_2035],
            marker_color=["#93c5fd", HYD],
            text=[f"{hidro_2025/1e6:.2f} TWh", f"{hidro_2035/1e6:.2f} TWh"],
            textposition="outside",
            hovertemplate="%{x}: %{y:,.0f} MWh<extra></extra>",
        ))
        fig.update_layout(yaxis_title="MWh/ano", showlegend=False,
                          yaxis=dict(range=[0, hidro_2035 * 1.18]))
        st.plotly_chart(fig, use_container_width=True, key="h_bar")

    with g2:
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=HIDRO_2035_PCT * 100,
            delta={"reference": HIDRO_2025_PCT * 100, "suffix": " pp",
                   "increasing": {"color": WIN}},
            number={"suffix": " %", "font": {"size": 40}},
            title={"text": "Participação da hidro na matriz (2035)", "font": {"size": 13}},
            gauge={
                "axis": {"range": [0, 70], "tickwidth": 1},
                "bar":  {"color": HYD},
                "steps": [
                    {"range": [0,  40], "color": "#e0f2fe"},
                    {"range": [40, 44], "color": "#bae6fd"},
                    {"range": [44, 70], "color": "#f1f5f9"},
                ],
                "threshold": {"line": {"color": THR, "width": 3}, "value": 44},
            },
        ))
        fig.update_layout(height=300, margin=dict(l=20, r=20, t=50, b=10),
                          paper_bgcolor=BG_CHART, font=dict(color=TEXT_PRI))
        st.plotly_chart(fig, use_container_width=True, key="h_gauge")

    g3, g4 = st.columns(2)
    with g3:
        resto_2025 = socio["ee"] - hidro_2025
        resto_2035 = dem_2035    - hidro_2035
        fig = base_fig("Composição da matriz: hidro × outras fontes", height=300)
        fig.add_trace(go.Bar(name="Hidro",
            x=["2025", "2035"], y=[hidro_2025, hidro_2035],
            marker_color=HYD,
            text=[f"{HIDRO_2025_PCT*100:.0f}%", f"{HIDRO_2035_PCT*100:.0f}%"],
            textposition="inside", textfont=dict(color="white", size=12),
        ))
        fig.add_trace(go.Bar(name="Outras fontes",
            x=["2025", "2035"], y=[resto_2025, resto_2035],
            marker_color="#cbd5e1",
            text=[f"{(1-HIDRO_2025_PCT)*100:.0f}%", f"{(1-HIDRO_2035_PCT)*100:.0f}%"],
            textposition="inside", textfont=dict(color="#475569", size=12),
        ))
        fig.update_layout(barmode="stack", yaxis_title="MWh/ano",
                          legend=dict(orientation="h", y=1.08))
        st.plotly_chart(fig, use_container_width=True, key="h_comp")

    with g4:
        anos  = list(range(2025, 2036))
        vals  = [hidro_2025 + (hidro_2035 - hidro_2025) * (a - 2025) / 10 for a in anos]
        fig = base_fig("Trajetória de crescimento hidro 2025–2035", height=300)
        fig.add_trace(go.Scatter(
            x=anos, y=[v/1e6 for v in vals],
            mode="lines+markers",
            line=dict(color=HYD, width=2.5),
            marker=dict(size=6, color=HYD),
            fill="tozeroy", fillcolor="rgba(14,165,233,0.08)",
            hovertemplate="Ano %{x}: %{y:.2f} TWh<extra></extra>",
            name="Hidro",
        ))
        fig.add_hline(y=hidro_2025/1e6, line_dash="dot", line_color="#94a3b8",
                      annotation_text="Base 2025", annotation_position="bottom right")
        fig.update_layout(yaxis_title="TWh/ano", showlegend=False,
                          xaxis=dict(tickvals=anos, tickangle=-45))
        st.plotly_chart(fig, use_container_width=True, key="h_traj")

    st.caption(
        f"Participação base: 40% (BEN 2025). Teto 2035: 44% × demanda projetada {cen_lbl} "
        f"({_fmt(dem_2035/1e6,2)} TWh). Expansão: +{_fmt(margem/1e6,2)} TWh "
        f"(+{_fmt(margem_pct,1)}%). A hidro é a única renovável firmemente despachável do país."
    )

# =======================================================================
#  ABA 5 — TERMO
# =======================================================================
def tab_termo(socio: dict):
    section_title("Potencial Termelétrico",
                  "Participação atual e projeção da geração termelétrica até 2035")

    TERMO_2026_PCT = 0.55
    CSV_PROJ = str(ROOT / "projecoes_demanda.csv")

    def _dem2035(cenario_key):
        try:
            df_proj = pd.read_csv(CSV_PROJ)
            df_proj["Ano"] = df_proj["Ano"].astype(int)
            row = df_proj[
                (df_proj["Ano"] == 2035) &
                (df_proj["Cenario"] == cenario_key) &
                (df_proj["Local"] == "Utopia")
            ]
            if not row.empty:
                return float(row["EE_TOTAL"].iloc[0])
        except Exception:
            pass
        return socio["ee"]

    CEN_OPTS = {
        "Referência":      "Referencia",
        "Alto (PIB +3%)":  "Alto",
        "Baixo (PIB -3%)": "Baixo",
    }
    CEN_CORES = {
        "Referencia": ACCENT,
        "Alto":       "#22c55e",
        "Baixo":      "#f59e0b",
    }

    # ── Controles ─────────────────────────────────────────────────
    ctrl1, ctrl2 = st.columns([1.2, 1])
    with ctrl1:
        cen_lbl = st.radio(
            "📊 Cenário de demanda 2035",
            list(CEN_OPTS.keys()),
            horizontal=True,
            key="t_cen_sel",
        )
    cen_key = CEN_OPTS[cen_lbl]
    cen_cor = CEN_CORES[cen_key]

    dem_2035 = _dem2035(cen_key)

    with ctrl2:
        termo_2035_pct = st.slider(
            "🔥 Participação térmica em 2035 (%)",
            min_value=10,
            max_value=70,
            value=40,
            step=1,
            key="t_pct_slider",
        ) / 100

    # ── Cálculos ──────────────────────────────────────────────────
    termo_2026 = TERMO_2026_PCT * socio["ee"]
    termo_2035 = termo_2035_pct * dem_2035
    var        = termo_2035 - termo_2026
    var_pct    = (var / termo_2026 * 100) if termo_2026 else 0
    pp_var     = termo_2035_pct * 100 - TERMO_2026_PCT * 100

    # ── Row 1: KPIs ───────────────────────────────────────────────
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi_card("Participação 2026",    "55 %",
                         "do BEN 2026", THR), unsafe_allow_html=True)
    k2.markdown(kpi_card("Geração térmica 2026", f"{_fmt(termo_2026/1e6,2)} TWh",
                         f"{_fmt(termo_2026)} MWh", THR), unsafe_allow_html=True)
    k3.markdown(kpi_card("Participação 2035",    f"{termo_2035_pct*100:.0f} %",
                         f"{pp_var:+.0f} pp vs. 2026",
                         WIN if pp_var < 0 else THR), unsafe_allow_html=True)
    k4.markdown(kpi_card("Geração térmica 2035", f"{_fmt(termo_2035/1e6,2)} TWh",
                         f"{termo_2035_pct*100:.0f}% × demanda 2035", ACCENT_D), unsafe_allow_html=True)

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    k5, k6, k7, k8 = st.columns(4)
    k5.markdown(kpi_card("Variação absoluta",  f"{_fmt(var/1e6,2)} TWh",
                         "2026 → 2035",
                         WIN if var < 0 else THR), unsafe_allow_html=True)
    k6.markdown(kpi_card("Variação relativa",  f"{_fmt(var_pct,1)} %",
                         "vs. geração térmica 2026",
                         WIN if var < 0 else THR), unsafe_allow_html=True)
    k7.markdown(kpi_card("Demanda total 2035", f"{_fmt(dem_2035/1e6,2)} TWh",
                         f"cenário {cen_lbl}", cen_cor), unsafe_allow_html=True)
    k8.markdown(kpi_card("Variação pp",        f"{pp_var:+.0f} pp",
                         "pontos percentuais na participação",
                         WIN if pp_var < 0 else THR), unsafe_allow_html=True)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        fig = base_fig("Geração térmica: 2026 vs. 2035", height=300)
        fig.add_trace(go.Bar(
            x=["Térmica 2026", "Térmica 2035"],
            y=[termo_2026, termo_2035],
            marker_color=["#fca5a5", THR],
            text=[f"{termo_2026/1e6:.2f} TWh", f"{termo_2035/1e6:.2f} TWh"],
            textposition="outside",
            hovertemplate="%{x}: %{y:,.0f} MWh<extra></extra>",
        ))
        fig.update_layout(yaxis_title="MWh/ano", showlegend=False,
                          yaxis=dict(range=[0, max(termo_2026, termo_2035) * 1.18]))
        st.plotly_chart(fig, use_container_width=True, key="t_bar")

    with g2:
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=termo_2035_pct * 100,
            delta={"reference": TERMO_2026_PCT * 100, "suffix": " pp",
                   "decreasing": {"color": WIN}, "increasing": {"color": THR}},
            number={"suffix": " %", "font": {"size": 40}},
            title={"text": "Participação da térmica na matriz (2035)", "font": {"size": 13}},
            gauge={
                "axis": {"range": [0, 80], "tickwidth": 1},
                "bar":  {"color": THR},
                "steps": [
                    {"range": [0,  40], "color": "#fef2f2"},
                    {"range": [40, 55], "color": "#fecdd3"},
                    {"range": [55, 80], "color": "#f1f5f9"},
                ],
                "threshold": {"line": {"color": ACCENT, "width": 3},
                              "value": termo_2035_pct * 100},
            },
        ))
        fig.update_layout(height=300, margin=dict(l=20, r=20, t=50, b=10),
                          paper_bgcolor=BG_CHART, font=dict(color=TEXT_PRI))
        st.plotly_chart(fig, use_container_width=True, key="t_gauge")

    g3, g4 = st.columns(2)
    with g3:
        resto_2026 = socio["ee"] - termo_2026
        resto_2035 = dem_2035    - termo_2035
        fig = base_fig("Composição da matriz: térmica × outras fontes", height=300)
        fig.add_trace(go.Bar(name="Térmica",
            x=["2026", "2035"], y=[termo_2026, termo_2035],
            marker_color=THR,
            text=[f"{TERMO_2026_PCT*100:.0f}%", f"{termo_2035_pct*100:.0f}%"],
            textposition="inside", textfont=dict(color="white", size=12),
        ))
        fig.add_trace(go.Bar(name="Outras fontes",
            x=["2026", "2035"], y=[resto_2026, resto_2035],
            marker_color="#cbd5e1",
            text=[f"{(1-TERMO_2026_PCT)*100:.0f}%", f"{(1-termo_2035_pct)*100:.0f}%"],
            textposition="inside", textfont=dict(color="#475569", size=12),
        ))
        fig.update_layout(barmode="stack", yaxis_title="MWh/ano",
                          legend=dict(orientation="h", y=1.08))
        st.plotly_chart(fig, use_container_width=True, key="t_comp")

    with g4:
        anos = list(range(2026, 2036))
        vals = [termo_2026 + (termo_2035 - termo_2026) * (a - 2026) / 9 for a in anos]
        fig = base_fig("Trajetória de transição térmica 2026–2035", height=300)
        fig.add_trace(go.Scatter(
            x=anos, y=[v/1e6 for v in vals],
            mode="lines+markers",
            line=dict(color=THR, width=2.5),
            marker=dict(size=6, color=THR),
            fill="tozeroy", fillcolor="rgba(239,68,68,0.08)",
            hovertemplate="Ano %{x}: %{y:.2f} TWh<extra></extra>",
            name="Térmica",
        ))
        fig.add_hline(y=termo_2026/1e6, line_dash="dot", line_color="#94a3b8",
                      annotation_text="Base 2026", annotation_position="bottom right")
        fig.update_layout(yaxis_title="TWh/ano", showlegend=False,
                          xaxis=dict(tickvals=anos, tickangle=-45))
        st.plotly_chart(fig, use_container_width=True, key="t_traj")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.markdown(
        f'<div style="background:#fff1f2;border:1px solid #fecdd3;border-radius:14px;padding:18px 20px;">'        f'<div style="font-size:13px;font-weight:700;color:#b91c1c;margin-bottom:10px;">🔥 Natureza do Potencial Térmico</div>'        f'<div style="font-size:13px;color:{TEXT_SEC};line-height:1.7;">'        f'O potencial termelétrico depende da <strong style="color:{TEXT_PRI};">política energética</strong>. '        f'O interesse em <em>reduzir, pressionar, abandonar, aumentar</em> são <strong style="color:{TEXT_PRI};">decisões estratégicas</strong> — '        f'uma vez que consideramos a entrada de combustível como estável, a térmica é definida por escolha, não por recurso geográfico. '        f'No planejamento decenal, a tendência é que a participação relativa <em>recue</em> com a expansão de renováveis.'        f'</div></div>',
        unsafe_allow_html=True,
    )
    st.caption(
        f"Base: 55% (BEN 2026) = {_fmt(termo_2026/1e6,2)} TWh. "
        f"Alvo 2035: {termo_2035_pct*100:.0f}% × {_fmt(dem_2035/1e6,2)} TWh ({cen_lbl}) "
        f"= {_fmt(termo_2035/1e6,2)} TWh. Variação: {_fmt(var/1e6,2)} TWh ({_fmt(var_pct,1)}%)."
    )

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
                 f"(`ENTREGA_DEMANDA_utopia.xlsx` / aba RESUMO UTÓPIA): {e}")
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
