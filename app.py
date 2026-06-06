# =============================================================================
# 📊 Analyse et Visualisation Interactive de Données — Version Pro
# =============================================================================
# Application Streamlit professionnelle orientée Data Analytics.
# Mini Power BI / Tableau développé en Python.
#
# Onglets :
#   1. Données          — Aperçu, info, types détectés
#   2. Profiling         — Analyse automatique colonne par colonne
#   3. Qualité           — Valeurs manquantes, heatmap, insights
#   4. Corrélations      — Heatmap, top +/−, classification, insights
#   5. Visualisation     — 16 types de graphiques interactifs
#   6. Statistiques      — Résumé descriptif complet
#   7. Export            — CSV, Excel, PNG, SVG, PDF, rapport
# =============================================================================

# ---------------------------------------------------------------------------
# Imports
# ---------------------------------------------------------------------------
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import base64
import tempfile
import os

# ---------------------------------------------------------------------------
# Configuration de la page
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Data Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# SECTION 1 — CSS PERSONNALISÉ
# =============================================================================

def inject_custom_css() -> None:
    """Injecte les styles CSS pour un rendu moderne et premium."""
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    /* --- Titre principal --- */
    .main-title {
        font-size: 2.4rem; font-weight: 800;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        text-align: center; padding: 0.5rem 0 0.2rem;
    }
    .main-subtitle {
        text-align: center; color: #8892b0; font-size: 1.05rem; margin-bottom: 1.2rem;
    }

    /* --- Cartes KPI --- */
    .kpi-card {
        background: linear-gradient(135deg, #1e1e2f 0%, #2d2d44 100%);
        border: 1px solid rgba(102,126,234,0.2); border-radius: 16px;
        padding: 1.3rem 1rem; text-align: center;
        box-shadow: 0 4px 24px rgba(0,0,0,0.18);
        transition: transform 0.25s ease, box-shadow 0.25s ease;
    }
    .kpi-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 32px rgba(102,126,234,0.3);
    }
    .kpi-value { font-size: 1.9rem; font-weight: 700; color: #667eea; }
    .kpi-label {
        font-size: 0.78rem; color: #8892b0; margin-top: 0.25rem;
        text-transform: uppercase; letter-spacing: 0.06em;
    }

    /* --- Variantes KPI --- */
    .kpi-card-purple .kpi-value { color: #a78bfa; }
    .kpi-card-green  .kpi-value { color: #10b981; }
    .kpi-card-rose   .kpi-value { color: #f093fb; }
    .kpi-card-orange .kpi-value { color: #f59e0b; }
    .kpi-card-red    .kpi-value { color: #f5576c; }

    /* --- Séparateur --- */
    .custom-divider {
        height: 3px; border: none; border-radius: 2px; margin: 1.2rem 0;
        background: linear-gradient(90deg, #667eea, #764ba2, #f093fb, #667eea);
    }

    /* --- Section headers --- */
    .section-header {
        font-size: 1.3rem; font-weight: 700; color: #e2e8f0;
        border-left: 4px solid #667eea; padding-left: 12px; margin: 1.2rem 0 0.8rem;
    }

    /* --- Insight card --- */
    .insight-card {
        background: rgba(102,126,234,0.08); border-left: 4px solid #667eea;
        border-radius: 0 12px 12px 0; padding: 0.9rem 1.2rem; margin: 0.5rem 0;
        color: #c9d1d9; font-size: 0.92rem; line-height: 1.5;
    }
    .insight-card-warning {
        background: rgba(245,87,108,0.08); border-left-color: #f5576c;
    }
    .insight-card-success {
        background: rgba(16,185,129,0.08); border-left-color: #10b981;
    }

    /* --- Sidebar --- */
    [data-testid="stSidebar"] { background: linear-gradient(180deg, #0f0f1a 0%, #1a1a2e 100%); }

    /* --- Tabs --- */
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] { border-radius: 10px 10px 0 0; padding: 10px 18px; font-weight: 600; }

    /* --- Download buttons --- */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white; border: none; border-radius: 10px; font-weight: 600;
    }
    .stDownloadButton > button:hover { opacity: 0.85; }

    /* --- Expander --- */
    .streamlit-expanderHeader { font-weight: 600; }
    </style>
    """, unsafe_allow_html=True)


# =============================================================================
# SECTION 2 — CHARGEMENT DES DONNÉES
# =============================================================================

@st.cache_data(show_spinner=False)
def load_data(uploaded_file) -> pd.DataFrame:
    """Charge un fichier CSV ou Excel et renvoie un DataFrame."""
    try:
        name = uploaded_file.name.lower()
        if name.endswith(".csv"):
            return pd.read_csv(uploaded_file, sep=None, engine="python")
        elif name.endswith((".xlsx", ".xls")):
            return pd.read_excel(uploaded_file, engine="openpyxl")
        else:
            st.error("❌ Format non supporté.")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erreur de chargement : {e}")
        return pd.DataFrame()


# =============================================================================
# SECTION 3 — DÉTECTION AUTOMATIQUE DES TYPES
# =============================================================================

def detect_variable_type(df: pd.DataFrame) -> dict:
    """
    Classe chaque colonne en Quantitative / Qualitative / Temporelle.
    Tente la conversion automatique des dates.
    """
    types = {}
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            types[col] = "Temporelle"
            continue
        if df[col].dtype == object:
            try:
                converted = pd.to_datetime(
                    df[col], format="mixed", dayfirst=False, errors="coerce"
                )
                ratio = converted.notna().sum() / max(df[col].notna().sum(), 1)
                if ratio >= 0.7:
                    df[col] = converted
                    types[col] = "Temporelle"
                    continue
            except Exception:
                pass
        if pd.api.types.is_numeric_dtype(df[col]):
            types[col] = "Quantitative"
        else:
            types[col] = "Qualitative"
    return types


# =============================================================================
# SECTION 4 — CONFIGURATION GRAPHIQUE & THÈMES
# =============================================================================

THEME_MAP = {
    "Plotly Dark": "plotly_dark",
    "Plotly": "plotly",
    "Plotly White": "plotly_white",
    "GGPlot2": "ggplot2",
    "Seaborn": "seaborn",
    "Presentation": "presentation",
    "Simple White": "simple_white",
}

PALETTE_MAP = {
    "Prism": px.colors.qualitative.Prism,
    "Vivid": px.colors.qualitative.Vivid,
    "Bold": px.colors.qualitative.Bold,
    "Pastel": px.colors.qualitative.Pastel,
    "Safe": px.colors.qualitative.Safe,
    "D3": px.colors.qualitative.D3,
    "Set2": px.colors.qualitative.Set2,
    "Plotly": px.colors.qualitative.Plotly,
}

LEGEND_POS = {
    "Top": dict(yanchor="bottom", y=1.02, xanchor="center", x=0.5, orientation="h"),
    "Bottom": dict(yanchor="top", y=-0.15, xanchor="center", x=0.5, orientation="h"),
    "Left": dict(yanchor="middle", y=0.5, xanchor="right", x=-0.05),
    "Right": dict(yanchor="middle", y=0.5, xanchor="left", x=1.05),
}


def default_config() -> dict:
    """Retourne la configuration graphique par défaut."""
    return dict(
        template="plotly_dark",
        width=None,  # None = responsive
        height=520,
        font_size=13,
        title_font_size=18,
        palette=px.colors.qualitative.Prism,
        primary="#667eea",
        secondary="#764ba2",
        x_title="",
        y_title="",
        label_rotation=0,
        show_legend=True,
        legend_pos="Top",
        show_grid=True,
        show_values=False,
        show_pct=False,
        export_format="PNG",
    )


def apply_config(fig: go.Figure, cfg: dict) -> go.Figure:
    """Applique la configuration personnalisée à une figure Plotly."""
    layout_updates = dict(
        template=cfg["template"],
        height=cfg["height"],
        font=dict(family="Inter", size=cfg["font_size"]),
        title_font_size=cfg["title_font_size"],
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    if cfg["width"]:
        layout_updates["width"] = cfg["width"]
    if cfg["x_title"]:
        layout_updates["xaxis_title"] = cfg["x_title"]
    if cfg["y_title"]:
        layout_updates["yaxis_title"] = cfg["y_title"]
    if cfg["label_rotation"]:
        layout_updates["xaxis_tickangle"] = -cfg["label_rotation"]
    if not cfg["show_grid"]:
        layout_updates["xaxis_showgrid"] = False
        layout_updates["yaxis_showgrid"] = False
    if cfg["show_legend"]:
        layout_updates["showlegend"] = True
        layout_updates["legend"] = LEGEND_POS.get(cfg["legend_pos"], {})
    else:
        layout_updates["showlegend"] = False
    fig.update_layout(**layout_updates)
    return fig


def export_figure(fig: go.Figure, fmt: str = "png") -> bytes | None:
    """Exporte une figure Plotly dans le format demandé."""
    try:
        return fig.to_image(format=fmt.lower(), width=1200, height=700, scale=2)
    except Exception:
        return None


# =============================================================================
# SECTION 5 — KDE (Kernel Density Estimation) MANUELLE
# =============================================================================

def gaussian_kde(data: np.ndarray, n_points: int = 200) -> tuple:
    """Estime la densité par noyau gaussien (sans scipy)."""
    data = data[~np.isnan(data)]
    n = len(data)
    if n < 2:
        return np.array([]), np.array([])
    std = np.std(data, ddof=1)
    if std == 0:
        return np.array([]), np.array([])
    bw = 1.06 * std * n ** (-1 / 5)  # Règle de Scott
    x_grid = np.linspace(data.min() - 3 * bw, data.max() + 3 * bw, n_points)
    kde = np.zeros(n_points)
    for xi in data:
        kde += np.exp(-0.5 * ((x_grid - xi) / bw) ** 2)
    kde /= n * bw * np.sqrt(2 * np.pi)
    return x_grid, kde


# =============================================================================
# SECTION 6 — FONCTIONS DE CRÉATION DE GRAPHIQUES
# =============================================================================

def _base(fig, cfg, title):
    """Raccourci : applique config + titre."""
    fig.update_layout(title=title)
    return apply_config(fig, cfg)


# ---- Bar Chart ----
def chart_bar(df, x, cfg, title):
    counts = df[x].value_counts().reset_index()
    counts.columns = [x, "Effectif"]
    fig = px.bar(counts, x=x, y="Effectif", color=x,
                 color_discrete_sequence=cfg["palette"], text_auto=cfg["show_values"])
    if cfg["show_pct"]:
        total = counts["Effectif"].sum()
        counts["pct"] = (counts["Effectif"] / total * 100).round(1).astype(str) + "%"
        fig.update_traces(text=counts["pct"], textposition="outside")
    fig.update_traces(marker_line_width=0)
    return _base(fig, cfg, title)


# ---- Histogramme ----
def chart_histogram(df, x, cfg, title):
    fig = px.histogram(df, x=x, nbins=30, color_discrete_sequence=[cfg["primary"]],
                       marginal="box")
    if cfg["show_values"]:
        fig.update_traces(texttemplate="%{y}", textposition="outside")
    return _base(fig, cfg, title)


# ---- Histogramme + KDE ----
def chart_histogram_kde(df, x, cfg, title):
    fig = go.Figure()
    vals = df[x].dropna().values
    fig.add_trace(go.Histogram(
        x=vals, nbinsx=30, name="Histogramme",
        marker_color=cfg["primary"], opacity=0.7,
        histnorm="probability density",
    ))
    xk, yk = gaussian_kde(vals)
    if len(xk) > 0:
        fig.add_trace(go.Scatter(
            x=xk, y=yk, mode="lines", name="KDE",
            line=dict(color=cfg["secondary"], width=3),
        ))
    fig.update_layout(barmode="overlay", title=title)
    return apply_config(fig, cfg)


# ---- Scatter Plot ----
def chart_scatter(df, x, y, color, cfg, title):
    fig = px.scatter(df, x=x, y=y, color=color,
                     color_discrete_sequence=cfg["palette"],
                     opacity=0.75, hover_data=df.columns.tolist()[:6])
    fig.update_traces(marker=dict(size=8, line=dict(width=0.5, color="white")))
    return _base(fig, cfg, title)


# ---- Bubble Chart ----
def chart_bubble(df, x, y, size, color, cfg, title):
    fig = px.scatter(df, x=x, y=y, size=size, color=color,
                     color_discrete_sequence=cfg["palette"],
                     hover_data=df.columns.tolist()[:6],
                     size_max=50)
    return _base(fig, cfg, title)


# ---- Box Plot ----
def chart_box(df, x, y, cfg, title):
    fig = px.box(df, x=x, y=y, color=x, color_discrete_sequence=cfg["palette"])
    return _base(fig, cfg, title)


# ---- Violin Plot ----
def chart_violin(df, x, y, cfg, title):
    fig = px.violin(df, x=x, y=y, color=x, box=True, points="all",
                    color_discrete_sequence=cfg["palette"])
    return _base(fig, cfg, title)


# ---- Strip Plot ----
def chart_strip(df, x, y, cfg, title):
    fig = px.strip(df, x=x, y=y, color=x, color_discrete_sequence=cfg["palette"])
    return _base(fig, cfg, title)


# ---- Line Chart ----
def chart_line(df, x, y, cfg, title):
    df_s = df.sort_values(by=x)
    fig = px.line(df_s, x=x, y=y, color_discrete_sequence=[cfg["primary"]], markers=True)
    fig.update_traces(line=dict(width=2.5), marker=dict(size=5))
    return _base(fig, cfg, title)


# ---- Area Chart ----
def chart_area(df, x, y, cfg, title):
    df_s = df.sort_values(by=x)
    fig = px.area(df_s, x=x, y=y, color_discrete_sequence=[cfg["primary"]])
    return _base(fig, cfg, title)


# ---- Pie Chart ----
def chart_pie(df, x, cfg, title):
    counts = df[x].value_counts().reset_index()
    counts.columns = [x, "Effectif"]
    fig = px.pie(counts, names=x, values="Effectif",
                 color_discrete_sequence=cfg["palette"])
    if cfg["show_pct"]:
        fig.update_traces(textinfo="label+percent")
    else:
        fig.update_traces(textinfo="label+value")
    return _base(fig, cfg, title)


# ---- Donut Chart ----
def chart_donut(df, x, cfg, title):
    counts = df[x].value_counts().reset_index()
    counts.columns = [x, "Effectif"]
    fig = px.pie(counts, names=x, values="Effectif",
                 color_discrete_sequence=cfg["palette"], hole=0.45)
    if cfg["show_pct"]:
        fig.update_traces(textinfo="label+percent")
    else:
        fig.update_traces(textinfo="label+value")
    return _base(fig, cfg, title)


# ---- Scatter Matrix ----
def chart_scatter_matrix(df, cols, color, cfg, title):
    fig = px.scatter_matrix(df, dimensions=cols, color=color,
                            color_discrete_sequence=cfg["palette"])
    fig.update_traces(diagonal_visible=True, marker=dict(size=4))
    fig.update_layout(height=max(cfg["height"], 700))
    return _base(fig, cfg, title)


# ---- Treemap ----
def chart_treemap(df, path_cols, cfg, title):
    fig = px.treemap(df, path=path_cols, color_discrete_sequence=cfg["palette"])
    return _base(fig, cfg, title)


# ---- Sunburst ----
def chart_sunburst(df, path_cols, cfg, title):
    fig = px.sunburst(df, path=path_cols, color_discrete_sequence=cfg["palette"])
    return _base(fig, cfg, title)


# ---- Heatmap de corrélation ----
def chart_corr_heatmap(corr, cfg, title="🔗 Matrice de Corrélation"):
    fig = go.Figure(data=go.Heatmap(
        z=corr.values, x=corr.columns.tolist(), y=corr.index.tolist(),
        colorscale=[[0, cfg["secondary"]], [0.5, "#1a1a2e"], [1, cfg["primary"]]],
        text=np.round(corr.values, 2), texttemplate="%{text}",
        textfont=dict(size=11),
        hovertemplate="<b>%{x}</b> vs <b>%{y}</b><br>r = %{z:.3f}<extra></extra>",
    ))
    fig.update_layout(height=max(cfg["height"], 550))
    return _base(fig, cfg, title)


# ---- Bar Chart valeurs manquantes ----
def chart_missing_bar(df, cfg):
    missing = df.isnull().sum()
    missing = missing[missing > 0].sort_values(ascending=True)
    if missing.empty:
        return None
    pct = (missing / len(df) * 100).round(2)
    fig = px.bar(x=pct.values, y=pct.index, orientation="h",
                 labels={"x": "% manquant", "y": "Colonne"},
                 color=pct.values,
                 color_continuous_scale=[cfg["primary"], "#f5576c"],
                 text_auto=True)
    fig.update_layout(coloraxis_showscale=False, showlegend=False)
    return _base(fig, cfg, "⚠️ Pourcentage de Valeurs Manquantes par Colonne")


# ---- Heatmap valeurs manquantes (pattern) ----
def chart_missing_heatmap(df, cfg):
    """Heatmap binaire : 1 = manquant, 0 = présent."""
    binary = df.isnull().astype(int)
    # Limiter à 200 lignes pour lisibilité
    sample = binary.head(200)
    fig = go.Figure(data=go.Heatmap(
        z=sample.values, x=sample.columns.tolist(),
        y=list(range(len(sample))),
        colorscale=[[0, "#1a1a2e"], [1, "#f5576c"]],
        showscale=False,
        hovertemplate="Ligne %{y} — %{x}<br>Manquant : %{z}<extra></extra>",
    ))
    fig.update_layout(yaxis=dict(autorange="reversed"),
                      xaxis_tickangle=-45)
    return _base(fig, cfg, "🔍 Motifs de Données Manquantes (200 premières lignes)")


# =============================================================================
# SECTION 7 — PROFILING
# =============================================================================

def create_profiling_table(df: pd.DataFrame, var_types: dict) -> pd.DataFrame:
    """Génère le tableau de profiling colonne par colonne."""
    rows = []
    for col in df.columns:
        vtype = var_types.get(col, "Inconnu")
        row = {
            "Nom": col,
            "Type": vtype,
            "Valeurs uniques": int(df[col].nunique()),
            "Valeurs manquantes": int(df[col].isnull().sum()),
            "% manquantes": round(df[col].isnull().sum() / max(len(df), 1) * 100, 2),
        }
        if vtype == "Quantitative":
            row["Minimum"] = round(float(df[col].min()), 2) if pd.notna(df[col].min()) else None
            row["Maximum"] = round(float(df[col].max()), 2) if pd.notna(df[col].max()) else None
            row["Moyenne"] = round(float(df[col].mean()), 2) if pd.notna(df[col].mean()) else None
            row["Médiane"] = round(float(df[col].median()), 2) if pd.notna(df[col].median()) else None
            row["Écart-type"] = round(float(df[col].std()), 2) if pd.notna(df[col].std()) else None
        else:
            for k in ["Minimum", "Maximum", "Moyenne", "Médiane", "Écart-type"]:
                row[k] = "—"
        rows.append(row)
    return pd.DataFrame(rows)


# =============================================================================
# SECTION 8 — CORRÉLATIONS AVANCÉES
# =============================================================================

def get_top_correlations(corr: pd.DataFrame, n: int = 10):
    """Retourne les top corrélations positives et négatives."""
    pairs = []
    cols = corr.columns
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            pairs.append((cols[i], cols[j], round(corr.iloc[i, j], 4)))
    pairs_sorted = sorted(pairs, key=lambda x: x[2], reverse=True)
    top_pos = [(a, b, r) for a, b, r in pairs_sorted if r > 0][:n]
    top_neg = [(a, b, r) for a, b, r in pairs_sorted if r < 0][-n:]
    top_neg.reverse()
    return top_pos, top_neg


def classify_correlations(corr: pd.DataFrame):
    """Classifie les corrélations en forte / modérée / faible."""
    strong, moderate, weak = [], [], []
    cols = corr.columns
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            r = corr.iloc[i, j]
            entry = {"Variable A": cols[i], "Variable B": cols[j], "r": round(r, 4)}
            if abs(r) > 0.7:
                strong.append(entry)
            elif abs(r) > 0.3:
                moderate.append(entry)
            else:
                weak.append(entry)
    return (
        pd.DataFrame(strong).sort_values("r", key=abs, ascending=False) if strong else pd.DataFrame(),
        pd.DataFrame(moderate).sort_values("r", key=abs, ascending=False) if moderate else pd.DataFrame(),
        pd.DataFrame(weak) if weak else pd.DataFrame(),
    )


def generate_corr_insights(corr: pd.DataFrame) -> list[str]:
    """Génère des insights automatiques sur les corrélations."""
    insights = []
    cols = corr.columns
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            r = corr.iloc[i, j]
            if r > 0.7:
                insights.append(
                    f"📈 Une **forte corrélation positive** est observée entre "
                    f"**{cols[i]}** et **{cols[j]}** (r = {r:.2f})."
                )
            elif r < -0.7:
                insights.append(
                    f"📉 Une **forte corrélation négative** existe entre "
                    f"**{cols[i]}** et **{cols[j]}** (r = {r:.2f})."
                )
    return insights


# =============================================================================
# SECTION 9 — INSIGHTS VALEURS MANQUANTES
# =============================================================================

def generate_missing_insights(df: pd.DataFrame) -> list[str]:
    """Génère des observations automatiques sur les valeurs manquantes."""
    insights = []
    missing = df.isnull().sum()
    total_pct = round(df.isnull().sum().sum() / max(df.size, 1) * 100, 2)
    insights.append(
        f"📊 Le dataset contient globalement **{total_pct} %** de données manquantes."
    )
    high = missing[missing / len(df) > 0.1]
    for col in high.index:
        pct = round(missing[col] / len(df) * 100, 1)
        insights.append(
            f"⚠️ La colonne **\"{col}\"** contient **{pct} %** de données manquantes."
        )
    critical = missing[missing / len(df) > 0.5]
    if not critical.empty:
        names = ", ".join([f"**\"{c}\"**" for c in critical.index])
        insights.append(
            f"🔴 Les colonnes {names} présentent un taux très élevé de valeurs absentes "
            f"(> 50 %) et pourraient nécessiter un traitement spécifique."
        )
    if missing.sum() == 0:
        insights = ["✅ Aucune valeur manquante détectée — le dataset est complet !"]
    return insights


# =============================================================================
# SECTION 10 — EXPORT & RAPPORT PDF
# =============================================================================

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Données")
    return buf.getvalue()


def generate_pdf_report(
    df: pd.DataFrame,
    var_types: dict,
    quant_cols: list,
    qual_cols: list,
    temp_cols: list,
    profiling_table: pd.DataFrame,
    corr_matrix: pd.DataFrame | None,
) -> bytes | None:
    """Génère un rapport PDF complet avec fpdf2."""
    try:
        from fpdf import FPDF
    except ImportError:
        return None

    def safe_txt(txt):
        """Encode en latin-1 pour éviter les crashs FPDFUnicodeEncodingException avec Helvetica."""
        return str(txt).encode("latin-1", "replace").decode("latin-1")

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    # --- Page de titre ---
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 24)
    pdf.ln(40)
    pdf.cell(0, 15, "Rapport d'Analyse de Donnees", ln=True, align="C")
    pdf.set_font("Helvetica", "", 12)
    pdf.ln(10)
    pdf.cell(0, 8, f"Genere automatiquement | {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align="C")
    pdf.ln(10)
    pdf.set_draw_color(102, 126, 234)
    pdf.set_line_width(1)
    pdf.line(30, pdf.get_y(), 180, pdf.get_y())

    # --- Resume du dataset ---
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 12, "1. Resume du Dataset", ln=True)
    pdf.ln(4)
    pdf.set_font("Helvetica", "", 11)
    info_lines = [
        f"Nombre de lignes : {len(df)}",
        f"Nombre de colonnes : {len(df.columns)}",
        f"Variables quantitatives : {len(quant_cols)}",
        f"Variables qualitatives : {len(qual_cols)}",
        f"Variables temporelles : {len(temp_cols)}",
        f"Valeurs manquantes totales : {int(df.isnull().sum().sum())}",
        f"Completude : {100 * (1 - df.isnull().sum().sum() / max(df.size, 1)):.1f} %",
    ]
    for line in info_lines:
        pdf.cell(0, 7, safe_txt(line), ln=True)

    # --- Profiling ---
    pdf.ln(6)
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 12, "2. Profiling des Colonnes", ln=True)
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 8)
    headers = ["Nom", "Type", "Uniques", "Manq.", "% Manq.", "Min", "Max", "Moy", "Med", "Std"]
    col_widths = [30, 20, 14, 14, 14, 18, 18, 18, 18, 18]
    for h, w in zip(headers, col_widths):
        pdf.cell(w, 7, safe_txt(h), border=1, align="C")
    pdf.ln()
    pdf.set_font("Helvetica", "", 7)
    for _, row in profiling_table.iterrows():
        vals = [
            str(row["Nom"])[:18],
            str(row["Type"])[:12],
            str(row["Valeurs uniques"]),
            str(row["Valeurs manquantes"]),
            str(row["% manquantes"]),
            str(row["Minimum"]),
            str(row["Maximum"]),
            str(row["Moyenne"]),
            str(row["Médiane"]),
            str(row["Écart-type"]),
        ]
        for v, w in zip(vals, col_widths):
            pdf.cell(w, 6, safe_txt(v), border=1, align="C")
        pdf.ln()

    # --- Statistiques ---
    if quant_cols:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 16)
        pdf.cell(0, 12, "3. Statistiques Descriptives", ln=True)
        pdf.ln(2)
        desc = df[quant_cols].describe().T
        pdf.set_font("Helvetica", "B", 8)
        desc_headers = ["Variable"] + [str(c) for c in desc.columns]
        dw = [30] + [20] * len(desc.columns)
        for h, w in zip(desc_headers, dw):
            pdf.cell(w, 7, safe_txt(h), border=1, align="C")
        pdf.ln()
        pdf.set_font("Helvetica", "", 7)
        for idx, row_data in desc.iterrows():
            pdf.cell(30, 6, safe_txt(str(idx)[:18]), border=1, align="C")
            for c in desc.columns:
                pdf.cell(20, 6, safe_txt(f"{row_data[c]:.2f}"), border=1, align="C")
            pdf.ln()

    # --- Valeurs manquantes ---
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 12, "4. Valeurs Manquantes", ln=True)
    pdf.ln(4)
    pdf.set_font("Helvetica", "", 11)
    for insight in generate_missing_insights(df):
        clean = insight.replace("**", "").replace("⚠️", "").replace("📊", "").replace("🔴", "").replace("✅", "").replace("—", "-")
        pdf.multi_cell(0, 7, safe_txt(clean))
        pdf.ln(2)

    # --- Correlations ---
    if corr_matrix is not None and not corr_matrix.empty:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 16)
        pdf.cell(0, 12, "5. Correlations", ln=True)
        pdf.ln(4)
        pdf.set_font("Helvetica", "", 11)
        for insight in generate_corr_insights(corr_matrix):
            clean = insight.replace("**", "").replace("📈", "").replace("📉", "").replace("—", "-")
            pdf.multi_cell(0, 7, safe_txt(clean))
            pdf.ln(2)

    return bytes(pdf.output())


# =============================================================================
# SECTION 11 — FILTRES INTERACTIFS
# =============================================================================

def apply_filters(df: pd.DataFrame, var_types: dict) -> pd.DataFrame:
    """Crée des widgets de filtrage dans la sidebar et retourne le df filtré."""
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔎 Filtres")
    filtered = df.copy()
    with st.sidebar.expander("Configurer les filtres", expanded=False):
        for col, vtype in var_types.items():
            if vtype == "Qualitative":
                uniques = df[col].dropna().unique().tolist()
                if 1 < len(uniques) <= 50:
                    sel = st.multiselect(f"🏷️ {col}", uniques, default=uniques, key=f"f_{col}")
                    filtered = filtered[filtered[col].isin(sel)]
            elif vtype == "Quantitative":
                cmin = float(df[col].min()) if pd.notna(df[col].min()) else 0.0
                cmax = float(df[col].max()) if pd.notna(df[col].max()) else 1.0
                if cmin < cmax:
                    lo, hi = st.slider(f"🔢 {col}", cmin, cmax, (cmin, cmax), key=f"f_{col}")
                    filtered = filtered[(filtered[col] >= lo) & (filtered[col] <= hi)]
            elif vtype == "Temporelle" and pd.api.types.is_datetime64_any_dtype(df[col]):
                dmin, dmax = df[col].min().date(), df[col].max().date()
                if dmin < dmax:
                    d1 = st.date_input(f"📅 {col} début", dmin, dmin, dmax, key=f"f_{col}_s")
                    d2 = st.date_input(f"📅 {col} fin", dmax, dmin, dmax, key=f"f_{col}_e")
                    filtered = filtered[(filtered[col].dt.date >= d1) & (filtered[col].dt.date <= d2)]
    return filtered


# =============================================================================
# SECTION 12 — SIDEBAR : THÈME & PERSONNALISATION
# =============================================================================

def sidebar_chart_config() -> dict:
    """Construit la configuration graphique depuis les widgets sidebar."""
    cfg = default_config()

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎨 Thème graphique")
    theme_name = st.sidebar.selectbox("Thème", list(THEME_MAP.keys()), index=0)
    cfg["template"] = THEME_MAP[theme_name]

    st.sidebar.markdown("---")
    st.sidebar.markdown("### ⚙️ Personnalisation")
    with st.sidebar.expander("Apparence", expanded=False):
        cfg["height"] = st.slider("Hauteur (px)", 300, 1000, 520, 20)
        use_fixed = st.checkbox("Largeur fixe", False)
        if use_fixed:
            cfg["width"] = st.slider("Largeur (px)", 400, 1400, 800, 50)
        cfg["font_size"] = st.slider("Taille police", 8, 22, 13)
        cfg["title_font_size"] = st.slider("Taille titre", 12, 32, 18)

    with st.sidebar.expander("Couleurs", expanded=False):
        pal_name = st.selectbox("Palette", list(PALETTE_MAP.keys()), index=0)
        cfg["palette"] = PALETTE_MAP[pal_name]
        cfg["primary"] = st.color_picker("Couleur principale", "#667eea")
        cfg["secondary"] = st.color_picker("Couleur secondaire", "#764ba2")

    with st.sidebar.expander("Axes & Légende", expanded=False):
        cfg["x_title"] = st.text_input("Titre axe X", "")
        cfg["y_title"] = st.text_input("Titre axe Y", "")
        cfg["label_rotation"] = st.slider("Rotation labels (°)", 0, 90, 0, 5)
        cfg["show_legend"] = st.checkbox("Afficher légende", True)
        cfg["legend_pos"] = st.selectbox("Position légende", ["Top", "Bottom", "Left", "Right"])
        cfg["show_grid"] = st.checkbox("Afficher grille", True)

    with st.sidebar.expander("Étiquettes", expanded=False):
        cfg["show_values"] = st.checkbox("Afficher les valeurs", False)
        cfg["show_pct"] = st.checkbox("Afficher les pourcentages", False)

    with st.sidebar.expander("Export visuel", expanded=False):
        cfg["export_format"] = st.selectbox("Format d'export", ["PNG", "SVG", "JPEG", "PDF"])

    return cfg


# =============================================================================
# SECTION 13 — TITRES DYNAMIQUES
# =============================================================================

def dynamic_title(chart_type: str, x: str, y: str | None) -> str:
    titles = {
        "bar": f"📊 Distribution de {x}",
        "histogram": f"📈 Distribution de {x}",
        "histogram_kde": f"📈 Distribution de {x} avec densité",
        "scatter": f"🔵 Relation entre {x} et {y}",
        "bubble": f"🫧 Relation entre {x} et {y}",
        "box": f"📦 Répartition de {y} selon {x}" if y else f"📦 {x}",
        "violin": f"🎻 Distribution de {y} selon {x}" if y else f"🎻 {x}",
        "strip": f"🔘 Dispersion de {y} selon {x}" if y else f"🔘 {x}",
        "line": f"📉 Évolution de {y} dans le temps" if y else f"📉 Évolution de {x}",
        "area": f"📐 Évolution de {y} dans le temps" if y else f"📐 {x}",
        "pie": f"🥧 Répartition de {x}",
        "donut": f"🍩 Répartition de {x}",
        "scatter_matrix": f"🔢 Matrice de dispersion",
        "treemap": f"🌳 Treemap",
        "sunburst": f"☀️ Sunburst",
    }
    return titles.get(chart_type, f"Graphique — {x}")


def suggest_chart(type_x, type_y):
    """Suggère le type de graphique le plus approprié."""
    if type_x == "Temporelle" or type_y == "Temporelle":
        return "line"
    if type_y is None:
        return "bar" if type_x == "Qualitative" else "histogram"
    if type_x == "Quantitative" and type_y == "Quantitative":
        return "scatter"
    if (type_x == "Qualitative") != (type_y == "Qualitative"):
        return "box"
    return "bar"


# =============================================================================
# SECTION 14 — RENDU DES ONGLETS
# =============================================================================

def render_kpis(df, var_types):
    """Affiche les KPI en haut de page."""
    quant = sum(1 for v in var_types.values() if v == "Quantitative")
    qual = sum(1 for v in var_types.values() if v == "Qualitative")
    temp = sum(1 for v in var_types.values() if v == "Temporelle")
    miss = int(df.isnull().sum().sum())
    pct = 100 * (1 - miss / max(df.size, 1))
    cols = st.columns(6)
    cards = [
        (f"{df.shape[0]:,}", "Lignes", ""),
        (f"{df.shape[1]}", "Colonnes", "kpi-card-purple"),
        (f"{quant}", "Quantitatives", "kpi-card-green"),
        (f"{qual}", "Qualitatives", "kpi-card-rose"),
        (f"{miss:,}", "Val. manquantes", "kpi-card-red" if miss > 0 else "kpi-card-green"),
        (f"{pct:.1f}%", "Complétude", "kpi-card-green" if pct > 95 else "kpi-card-orange"),
    ]
    for col, (val, label, cls) in zip(cols, cards):
        col.markdown(
            f'<div class="kpi-card {cls}">'
            f'<div class="kpi-value">{val}</div>'
            f'<div class="kpi-label">{label}</div></div>',
            unsafe_allow_html=True,
        )


def render_tab_data(df, var_types):
    """Onglet Données."""
    st.markdown('<p class="section-header">🗂️ Aperçu des données</p>', unsafe_allow_html=True)
    st.dataframe(df.head(100), use_container_width=True, height=420)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📐 Dimensions")
        st.info(f"**{df.shape[0]:,}** lignes × **{df.shape[1]}** colonnes")
    with c2:
        st.markdown("#### ℹ️ Types détectés")
        type_df = pd.DataFrame([
            {"Variable": k, "Type": v} for k, v in var_types.items()
        ])
        st.dataframe(type_df, use_container_width=True, hide_index=True)

    # Info technique
    st.markdown("#### 🔧 Informations techniques")
    buf = io.StringIO()
    df.info(buf=buf)
    st.code(buf.getvalue(), language="text")


def render_tab_profiling(df, var_types, quant_cols, qual_cols, temp_cols):
    """Onglet Profiling Dataset."""
    st.markdown('<p class="section-header">📋 Profiling automatique du dataset</p>', unsafe_allow_html=True)

    # KPI résumé
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("📏 Lignes", f"{len(df):,}")
    with c2:
        st.metric("📊 Colonnes", len(df.columns))
    with c3:
        st.metric("🔢 Quantitatives", len(quant_cols))
    with c4:
        st.metric("🏷️ Qualitatives", len(qual_cols))
    with c5:
        st.metric("📅 Temporelles", len(temp_cols))

    st.markdown("---")

    prof = create_profiling_table(df, var_types)
    st.markdown("#### 🔬 Analyse colonne par colonne")
    st.dataframe(
        prof,
        use_container_width=True,
        hide_index=True,
        height=min(len(prof) * 38 + 40, 600),
    )

    # Téléchargement du profiling
    st.download_button(
        "⬇️ Télécharger le profiling (CSV)", to_csv_bytes(prof),
        "profiling.csv", "text/csv",
    )


def render_tab_quality(df, cfg):
    """Onglet Qualité des données."""
    st.markdown('<p class="section-header">🧹 Qualité des données</p>', unsafe_allow_html=True)

    missing = df.isnull().sum()
    missing_nz = missing[missing > 0].sort_values(ascending=False)

    # Tableau
    if missing_nz.empty:
        st.success("✅ Aucune valeur manquante — dataset parfaitement complet !")
    else:
        st.markdown("#### 📋 Tableau des valeurs manquantes")
        miss_df = pd.DataFrame({
            "Colonne": missing_nz.index,
            "Nombre manquant": missing_nz.values,
            "% manquant": (missing_nz.values / len(df) * 100).round(2),
        })
        st.dataframe(miss_df, use_container_width=True, hide_index=True)

        # Graphiques
        c1, c2 = st.columns(2)
        with c1:
            fig_bar = chart_missing_bar(df, cfg)
            if fig_bar:
                st.plotly_chart(fig_bar, use_container_width=True)
        with c2:
            fig_heat = chart_missing_heatmap(df, cfg)
            if fig_heat:
                st.plotly_chart(fig_heat, use_container_width=True)

    # Insights
    st.markdown("---")
    st.markdown("#### 💡 Synthèse automatique")
    for insight in generate_missing_insights(df):
        css = "insight-card"
        if "⚠️" in insight or "🔴" in insight:
            css = "insight-card insight-card-warning"
        elif "✅" in insight:
            css = "insight-card insight-card-success"
        st.markdown(f'<div class="{css}">{insight}</div>', unsafe_allow_html=True)


def render_tab_correlations(df, quant_cols, cfg):
    """Onglet Corrélations."""
    st.markdown('<p class="section-header">🔗 Corrélations avancées</p>', unsafe_allow_html=True)

    if len(quant_cols) < 2:
        st.info("Il faut au moins 2 variables quantitatives pour calculer les corrélations.")
        return

    corr = df[quant_cols].corr()

    # Heatmap
    fig = chart_corr_heatmap(corr, cfg)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # Top corrélations
    top_pos, top_neg = get_top_correlations(corr)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📈 Top corrélations positives")
        if top_pos:
            st.dataframe(
                pd.DataFrame(top_pos, columns=["Variable A", "Variable B", "Corrélation"]),
                use_container_width=True, hide_index=True,
            )
        else:
            st.info("Aucune corrélation positive.")
    with c2:
        st.markdown("#### 📉 Top corrélations négatives")
        if top_neg:
            st.dataframe(
                pd.DataFrame(top_neg, columns=["Variable A", "Variable B", "Corrélation"]),
                use_container_width=True, hide_index=True,
            )
        else:
            st.info("Aucune corrélation négative.")

    # Classification
    st.markdown("---")
    st.markdown("#### 🎯 Classification des corrélations")
    strong, moderate, weak = classify_correlations(corr)
    t1, t2, t3 = st.tabs(["🔴 Fortes (|r| > 0.7)", "🟡 Modérées (0.3 < |r| < 0.7)", "🟢 Faibles (|r| < 0.3)"])
    with t1:
        if not strong.empty:
            st.dataframe(strong, use_container_width=True, hide_index=True)
        else:
            st.info("Aucune corrélation forte détectée.")
    with t2:
        if not moderate.empty:
            st.dataframe(moderate, use_container_width=True, hide_index=True)
        else:
            st.info("Aucune corrélation modérée.")
    with t3:
        if not weak.empty:
            st.dataframe(weak, use_container_width=True, hide_index=True)
        else:
            st.info("Aucune corrélation faible.")

    # Insights
    st.markdown("---")
    st.markdown("#### 💡 Insights automatiques")
    insights = generate_corr_insights(corr)
    if insights:
        for ins in insights:
            st.markdown(f'<div class="insight-card">{ins}</div>', unsafe_allow_html=True)
    else:
        st.markdown(
            '<div class="insight-card insight-card-success">'
            '✅ Aucune corrélation forte détectée — les variables semblent indépendantes.'
            '</div>', unsafe_allow_html=True,
        )

    return corr


def render_tab_visualization(df, var_types, quant_cols, qual_cols, temp_cols, cfg):
    """Onglet Visualisation — 16 types de graphiques."""
    st.markdown('<p class="section-header">🎨 Visualisation interactive</p>', unsafe_allow_html=True)

    # --- Paramètres du graphique (sidebar) ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📊 Paramètres du graphique")

    all_cols = list(var_types.keys())
    CHART_TYPES = {
        "auto": "🤖 Automatique",
        "histogram": "📈 Histogramme",
        "histogram_kde": "📈 Histogramme + KDE",
        "bar": "📊 Bar Chart",
        "pie": "🥧 Pie Chart",
        "donut": "🍩 Donut Chart",
        "scatter": "🔵 Scatter Plot",
        "bubble": "🫧 Bubble Chart",
        "box": "📦 Box Plot",
        "violin": "🎻 Violin Plot",
        "strip": "🔘 Strip Plot",
        "line": "📉 Line Chart",
        "area": "📐 Area Chart",
        "scatter_matrix": "🔢 Scatter Matrix",
        "treemap": "🌳 Treemap",
        "sunburst": "☀️ Sunburst",
    }

    selected_chart = st.sidebar.selectbox(
        "Type de graphique", list(CHART_TYPES.keys()),
        format_func=lambda k: CHART_TYPES[k], index=0,
    )

    # --- Variables ---
    var_x = st.sidebar.selectbox("Variable X", all_cols, index=0)
    var_y_options = ["— Aucune —"] + all_cols
    var_y = st.sidebar.selectbox("Variable Y (optionnel)", var_y_options, index=0)
    y_label = var_y if var_y != "— Aucune —" else None

    color_col = None
    size_col = None

    # Couleur
    color_options = ["— Aucune —"] + qual_cols
    color_choice = st.sidebar.selectbox("Variable couleur", color_options, index=0)
    if color_choice != "— Aucune —":
        color_col = color_choice

    # Taille (pour bubble chart)
    if selected_chart == "bubble":
        size_options = quant_cols
        if size_options:
            size_col = st.sidebar.selectbox("Variable taille (bulles)", size_options, index=0)

    # Variables pour scatter matrix
    sm_cols = None
    if selected_chart == "scatter_matrix" and quant_cols:
        sm_cols = st.sidebar.multiselect(
            "Variables (Scatter Matrix)", quant_cols, default=quant_cols[:4]
        )

    # Variables pour treemap / sunburst
    path_cols = None
    if selected_chart in ("treemap", "sunburst") and qual_cols:
        path_cols = st.sidebar.multiselect(
            "Colonnes hiérarchiques", qual_cols, default=qual_cols[:2]
        )

    # --- Déterminer le type de graphique ---
    type_x = var_types.get(var_x)
    type_y = var_types.get(var_y) if y_label else None

    if selected_chart == "auto":
        chart_type = suggest_chart(type_x, type_y)
    else:
        chart_type = selected_chart

    title = dynamic_title(chart_type, var_x, y_label)

    # --- Suggestion ---
    suggested = suggest_chart(type_x, type_y)
    st.sidebar.markdown("---")
    st.sidebar.info(f"💡 Suggestion auto : **{CHART_TYPES.get(suggested, suggested)}**")

    # --- Construction du graphique ---
    fig = None
    try:
        if chart_type == "bar":
            fig = chart_bar(df, var_x, cfg, title)

        elif chart_type == "histogram":
            fig = chart_histogram(df, var_x, cfg, title)

        elif chart_type == "histogram_kde":
            if type_x == "Quantitative":
                fig = chart_histogram_kde(df, var_x, cfg, title)
            else:
                st.warning("⚠️ L'histogramme + KDE nécessite une variable quantitative en X.")

        elif chart_type == "scatter":
            if y_label:
                fig = chart_scatter(df, var_x, y_label, color_col, cfg, title)
            else:
                st.warning("⚠️ Sélectionnez une variable Y pour le Scatter Plot.")

        elif chart_type == "bubble":
            if y_label and size_col:
                fig = chart_bubble(df, var_x, y_label, size_col, color_col, cfg, title)
            else:
                st.warning("⚠️ Le Bubble Chart nécessite X, Y et une variable taille.")

        elif chart_type == "box":
            if y_label:
                if type_x == "Qualitative":
                    fig = chart_box(df, var_x, y_label, cfg, title)
                else:
                    fig = chart_box(df, y_label, var_x, cfg, title)
            else:
                st.warning("⚠️ Sélectionnez une variable Y pour le Box Plot.")

        elif chart_type == "violin":
            if y_label:
                if type_x == "Qualitative":
                    fig = chart_violin(df, var_x, y_label, cfg, title)
                else:
                    fig = chart_violin(df, y_label, var_x, cfg, title)
            else:
                st.warning("⚠️ Sélectionnez une variable Y pour le Violin Plot.")

        elif chart_type == "strip":
            if y_label:
                if type_x == "Qualitative":
                    fig = chart_strip(df, var_x, y_label, cfg, title)
                else:
                    fig = chart_strip(df, y_label, var_x, cfg, title)
            else:
                st.warning("⚠️ Sélectionnez une variable Y pour le Strip Plot.")

        elif chart_type == "line":
            if y_label:
                fig = chart_line(df, var_x, y_label, cfg, title)
            elif type_x == "Temporelle":
                grp = df.groupby(df[var_x].dt.date).size().reset_index(name="Nombre")
                grp.columns = [var_x, "Nombre"]
                fig = px.line(grp, x=var_x, y="Nombre",
                              color_discrete_sequence=[cfg["primary"]], markers=True)
                fig = _base(fig, cfg, title)
            else:
                st.warning("⚠️ Sélectionnez une variable Y ou une variable temporelle en X.")

        elif chart_type == "area":
            if y_label:
                fig = chart_area(df, var_x, y_label, cfg, title)
            else:
                st.warning("⚠️ Sélectionnez une variable Y pour l'Area Chart.")

        elif chart_type == "pie":
            fig = chart_pie(df, var_x, cfg, title)

        elif chart_type == "donut":
            fig = chart_donut(df, var_x, cfg, title)

        elif chart_type == "scatter_matrix":
            if sm_cols and len(sm_cols) >= 2:
                fig = chart_scatter_matrix(df, sm_cols, color_col, cfg, title)
            else:
                st.warning("⚠️ Sélectionnez au moins 2 variables quantitatives.")

        elif chart_type == "treemap":
            if path_cols and len(path_cols) >= 1:
                fig = chart_treemap(df, path_cols, cfg, title)
            else:
                st.warning("⚠️ Sélectionnez au moins 1 colonne catégorielle.")

        elif chart_type == "sunburst":
            if path_cols and len(path_cols) >= 1:
                fig = chart_sunburst(df, path_cols, cfg, title)
            else:
                st.warning("⚠️ Sélectionnez au moins 1 colonne catégorielle.")

    except Exception as e:
        st.error(f"❌ Erreur : {e}")

    # --- Affichage du graphique ---
    if fig:
        st.plotly_chart(fig, use_container_width=True, key="main_viz")

        # Bouton d'export du graphique
        fmt = cfg["export_format"].lower()
        img_bytes = export_figure(fig, fmt)
        if img_bytes:
            st.download_button(
                f"🖼️ Télécharger ({cfg['export_format']})",
                img_bytes, f"graphique.{fmt}",
                f"image/{fmt}" if fmt != "pdf" else "application/pdf",
            )
        else:
            st.caption("💡 Installez `kaleido` pour activer l'export d'images.")


def render_tab_statistics(df, quant_cols, qual_cols, temp_cols):
    """Onglet Statistiques descriptives."""
    st.markdown('<p class="section-header">📈 Résumé statistique complet</p>', unsafe_allow_html=True)

    if quant_cols:
        st.markdown("#### 🔢 Variables quantitatives")
        desc_quant = df[quant_cols].describe().T
        st.dataframe(
            desc_quant.style.format("{:.2f}"),
            use_container_width=True,
        )

    if qual_cols:
        st.markdown("#### 🏷️ Variables qualitatives")
        st.dataframe(df[qual_cols].describe().T, use_container_width=True)

    if temp_cols:
        st.markdown("#### 📅 Variables temporelles")
        for col in temp_cols:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                c1, c2, c3 = st.columns(3)
                c1.metric("Début", str(df[col].min().date()))
                c2.metric("Fin", str(df[col].max().date()))
                c3.metric("Durée", f"{(df[col].max() - df[col].min()).days} jours")


def render_tab_export(df, var_types, quant_cols, qual_cols, temp_cols, profiling_table, corr_matrix, cfg):
    """Onglet Export complet."""
    st.markdown('<p class="section-header">📥 Centre d\'export</p>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### 📄 Données filtrées")
        st.download_button(
            "⬇️ Télécharger CSV", to_csv_bytes(df),
            "donnees_filtrees.csv", "text/csv",
        )
        st.download_button(
            "⬇️ Télécharger Excel", to_excel_bytes(df),
            "donnees_filtrees.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with c2:
        st.markdown("#### 📊 Statistiques & Profiling")
        st.download_button(
            "⬇️ Profiling (CSV)", to_csv_bytes(profiling_table),
            "profiling.csv", "text/csv",
        )
        if quant_cols:
            desc = df[quant_cols].describe().T.reset_index()
            desc.columns = ["Variable"] + list(df[quant_cols].describe().T.columns)
            st.download_button(
                "⬇️ Statistiques (CSV)", to_csv_bytes(desc),
                "statistiques.csv", "text/csv",
            )

    st.markdown("---")

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("#### 🔗 Corrélations")
        if corr_matrix is not None and not corr_matrix.empty:
            st.download_button(
                "⬇️ Matrice de corrélation (CSV)",
                to_csv_bytes(corr_matrix.reset_index()),
                "correlations.csv", "text/csv",
            )
        else:
            st.info("Pas de corrélations disponibles.")

    with c4:
        st.markdown("#### 📄 Rapport complet (PDF)")
        pdf_data = generate_pdf_report(
            df, var_types, quant_cols, qual_cols, temp_cols,
            profiling_table, corr_matrix,
        )
        if pdf_data:
            st.download_button(
                "📄 Générer & Télécharger le Rapport PDF",
                pdf_data, "rapport_analyse.pdf", "application/pdf",
            )
        else:
            st.warning(
                "💡 Installez `fpdf2` pour activer la génération de rapport PDF : "
                "`pip install fpdf2`"
            )


# =============================================================================
# SECTION 15 — APPLICATION PRINCIPALE
# =============================================================================

def main():
    """Point d'entrée de l'application."""
    inject_custom_css()

    # --- En-tête ---
    st.markdown(
        '<h1 class="main-title">📊 Analyse et Visualisation Interactive de Données</h1>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="main-subtitle">'
        'Importez, explorez, visualisez et exportez vos données — Data Analytics professionnel.'
        '</p>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # --- Sidebar : Upload ---
    st.sidebar.markdown("## 📂 Import de données")
    uploaded_file = st.sidebar.file_uploader(
        "Importer votre fichier", type=["csv", "xlsx"],
        help="Formats : CSV, XLSX",
    )

    if uploaded_file is None:
        st.markdown(
            '<div style="text-align:center;padding:5rem 1rem;">'
            '<p style="font-size:4.5rem;">📂</p>'
            '<h2 style="color:#667eea;">Bienvenue !</h2>'
            '<p style="color:#8892b0;font-size:1.1rem;max-width:500px;margin:0 auto;">'
            'Commencez par importer un fichier <b>CSV</b> ou <b>Excel</b> '
            'depuis la barre latérale pour démarrer l\'analyse.'
            '</p></div>',
            unsafe_allow_html=True,
        )
        return

    # --- Chargement ---
    with st.spinner("⏳ Chargement en cours…"):
        df_raw = load_data(uploaded_file)
    if df_raw.empty:
        st.warning("Le fichier est vide ou n'a pas pu être lu.")
        return

    # --- Détection des types ---
    var_types = detect_variable_type(df_raw)
    quant_cols = [c for c, t in var_types.items() if t == "Quantitative"]
    qual_cols = [c for c, t in var_types.items() if t == "Qualitative"]
    temp_cols = [c for c, t in var_types.items() if t == "Temporelle"]

    # --- Configuration graphique (sidebar) ---
    cfg = sidebar_chart_config()

    # --- Filtres (sidebar) ---
    df = apply_filters(df_raw, var_types)

    # --- KPI ---
    render_kpis(df, var_types)
    st.markdown("<br>", unsafe_allow_html=True)

    # --- Onglets ---
    tabs = st.tabs([
        "📋 Données",
        "📋 Profiling",
        "🧹 Qualité",
        "🔗 Corrélations",
        "📊 Visualisation",
        "📈 Statistiques",
        "📥 Export",
    ])

    with tabs[0]:
        render_tab_data(df, var_types)

    with tabs[1]:
        render_tab_profiling(df, var_types, quant_cols, qual_cols, temp_cols)

    with tabs[2]:
        render_tab_quality(df, cfg)

    corr_matrix = None
    with tabs[3]:
        result = render_tab_correlations(df, quant_cols, cfg)
        if result is not None:
            corr_matrix = result

    with tabs[4]:
        render_tab_visualization(df, var_types, quant_cols, qual_cols, temp_cols, cfg)

    with tabs[5]:
        render_tab_statistics(df, quant_cols, qual_cols, temp_cols)

    profiling_table = create_profiling_table(df, var_types)
    with tabs[6]:
        render_tab_export(df, var_types, quant_cols, qual_cols, temp_cols,
                          profiling_table, corr_matrix, cfg)

    # --- Footer ---
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.markdown(
        '<div style="text-align:center;color:#8892b0;font-size:0.8rem;padding-bottom:1rem;">'
        'Done by BoyWonder 🐦</div>',
        unsafe_allow_html=True,
    )


# =============================================================================
# Point d'entrée
# =============================================================================
if __name__ == "__main__":
    main()
