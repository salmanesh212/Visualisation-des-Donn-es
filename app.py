# =============================================================================
# 📊 Analyse et Visualisation Interactive de Données
# =============================================================================
# Application Streamlit professionnelle pour l'analyse exploratoire de données.
# Développée dans le cadre d'un projet universitaire.
#
# Fonctionnalités :
#   - Import CSV / Excel
#   - Détection automatique des types de variables
#   - Visualisations interactives avec Plotly
#   - Statistiques descriptives et corrélations
#   - Filtres dynamiques et export des résultats
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

# ---------------------------------------------------------------------------
# Configuration de la page Streamlit
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Data Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =============================================================================
# 🎨 Styles CSS personnalisés
# =============================================================================
def inject_custom_css() -> None:
    """Injecte du CSS personnalisé pour un rendu moderne et professionnel."""
    st.markdown(
        """
        <style>
        /* ---------- Police Google ---------- */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

        /* ---------- Base ---------- */
        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }

        /* ---------- Titre principal ---------- */
        .main-title {
            font-size: 2.4rem;
            font-weight: 700;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-align: center;
            padding: 0.5rem 0 0.2rem 0;
        }
        .main-subtitle {
            text-align: center;
            color: #8892b0;
            font-size: 1.05rem;
            margin-bottom: 1.5rem;
        }

        /* ---------- Cartes KPI ---------- */
        .kpi-card {
            background: linear-gradient(135deg, #1e1e2f 0%, #2d2d44 100%);
            border: 1px solid rgba(102, 126, 234, 0.25);
            border-radius: 16px;
            padding: 1.4rem 1.2rem;
            text-align: center;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .kpi-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 8px 30px rgba(102, 126, 234, 0.25);
        }
        .kpi-value {
            font-size: 2rem;
            font-weight: 700;
            color: #667eea;
        }
        .kpi-label {
            font-size: 0.85rem;
            color: #8892b0;
            margin-top: 0.3rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        /* ---------- Séparateur ---------- */
        .custom-divider {
            height: 3px;
            background: linear-gradient(90deg, #667eea, #764ba2, #667eea);
            border: none;
            border-radius: 2px;
            margin: 1.5rem 0;
        }

        /* ---------- Tableau des types ---------- */
        .type-badge-quant {
            background: rgba(102, 126, 234, 0.15);
            color: #667eea;
            padding: 3px 10px;
            border-radius: 8px;
            font-weight: 600;
            font-size: 0.82rem;
        }
        .type-badge-qual {
            background: rgba(118, 75, 162, 0.15);
            color: #a78bfa;
            padding: 3px 10px;
            border-radius: 8px;
            font-weight: 600;
            font-size: 0.82rem;
        }
        .type-badge-temp {
            background: rgba(16, 185, 129, 0.15);
            color: #10b981;
            padding: 3px 10px;
            border-radius: 8px;
            font-weight: 600;
            font-size: 0.82rem;
        }

        /* ---------- Sidebar ---------- */
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0f0f1a 0%, #1a1a2e 100%);
        }
        [data-testid="stSidebar"] .stMarkdown h1,
        [data-testid="stSidebar"] .stMarkdown h2,
        [data-testid="stSidebar"] .stMarkdown h3 {
            color: #e2e8f0;
        }

        /* ---------- Onglets ---------- */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 8px;
            padding: 8px 20px;
            font-weight: 500;
        }

        /* ---------- Bouton de téléchargement ---------- */
        .stDownloadButton > button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 10px;
            font-weight: 600;
            transition: opacity 0.2s ease;
        }
        .stDownloadButton > button:hover {
            opacity: 0.85;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# 📂 Chargement des données
# =============================================================================
@st.cache_data(show_spinner=False)
def load_data(file) -> pd.DataFrame:
    """
    Charge un fichier CSV ou Excel et retourne un DataFrame Pandas.

    Parameters
    ----------
    file : UploadedFile
        Fichier uploadé via st.file_uploader.

    Returns
    -------
    pd.DataFrame
        DataFrame contenant les données importées.
    """
    try:
        if file.name.endswith(".csv"):
            # Tenter plusieurs séparateurs courants
            df = pd.read_csv(file, sep=None, engine="python")
        elif file.name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file, engine="openpyxl")
        else:
            st.error("❌ Format de fichier non supporté. Veuillez importer un CSV ou XLSX.")
            return pd.DataFrame()
        return df
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement : {e}")
        return pd.DataFrame()


# =============================================================================
# 🔍 Détection automatique des types de variables
# =============================================================================
def detect_variable_type(df: pd.DataFrame) -> dict:
    """
    Analyse chaque colonne du DataFrame et la classe en :
      - 'Quantitative'  (int, float)
      - 'Qualitative'   (object, category, string, bool)
      - 'Temporelle'    (datetime, date, timestamp)

    La fonction tente également de convertir automatiquement les colonnes
    textuelles susceptibles de contenir des dates.

    Parameters
    ----------
    df : pd.DataFrame
        Le DataFrame à analyser.

    Returns
    -------
    dict
        Dictionnaire {nom_colonne: type_détecté}.
    """
    types = {}
    for col in df.columns:
        # --- Vérifier si c'est déjà un datetime ---
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            types[col] = "Temporelle"
            continue

        # --- Tenter la conversion en datetime pour les colonnes objet ---
        if df[col].dtype == object:
            try:
                converted = pd.to_datetime(df[col], infer_datetime_format=True, errors="coerce")
                # Si au moins 70 % des valeurs non-nulles sont converties → Temporelle
                if converted.notna().sum() / max(df[col].notna().sum(), 1) >= 0.7:
                    df[col] = converted
                    types[col] = "Temporelle"
                    continue
            except Exception:
                pass

        # --- Types numériques ---
        if pd.api.types.is_numeric_dtype(df[col]):
            types[col] = "Quantitative"
        else:
            types[col] = "Qualitative"

    return types


def build_type_table(var_types: dict) -> pd.DataFrame:
    """Construit un DataFrame récapitulatif des types détectés."""
    rows = []
    for var, vtype in var_types.items():
        rows.append({"Variable": var, "Type": vtype})
    return pd.DataFrame(rows)


# =============================================================================
# 📊 Fonctions de visualisation
# =============================================================================
# Palette de couleurs professionnelle
COLOR_PALETTE = px.colors.qualitative.Prism
GRADIENT_COLORS = ["#667eea", "#764ba2", "#f093fb", "#f5576c", "#4facfe", "#00f2fe"]


def create_bar_chart(df: pd.DataFrame, x: str, title: str) -> go.Figure:
    """Crée un Bar Chart des effectifs pour une variable qualitative."""
    counts = df[x].value_counts().reset_index()
    counts.columns = [x, "Effectif"]
    fig = px.bar(
        counts,
        x=x,
        y="Effectif",
        color=x,
        color_discrete_sequence=COLOR_PALETTE,
        title=title,
        text_auto=True,
    )
    fig.update_layout(
        template="plotly_dark",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", size=13),
        title_font_size=18,
        showlegend=False,
    )
    fig.update_traces(
        marker_line_width=0,
        textposition="outside",
    )
    return fig


def create_histogram(df: pd.DataFrame, x: str, title: str) -> go.Figure:
    """Crée un histogramme pour une variable quantitative."""
    fig = px.histogram(
        df,
        x=x,
        nbins=30,
        color_discrete_sequence=["#667eea"],
        title=title,
        marginal="box",
    )
    fig.update_layout(
        template="plotly_dark",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", size=13),
        title_font_size=18,
        bargap=0.05,
    )
    return fig


def create_scatter(df: pd.DataFrame, x: str, y: str, color: str | None, title: str) -> go.Figure:
    """Crée un Scatter Plot entre deux variables quantitatives."""
    fig = px.scatter(
        df,
        x=x,
        y=y,
        color=color,
        color_discrete_sequence=COLOR_PALETTE,
        title=title,
        hover_data=df.columns.tolist()[:5],
        opacity=0.75,
    )
    fig.update_layout(
        template="plotly_dark",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", size=13),
        title_font_size=18,
    )
    fig.update_traces(marker=dict(size=8, line=dict(width=0.5, color="white")))
    return fig


def create_box_plot(df: pd.DataFrame, x: str, y: str, title: str) -> go.Figure:
    """Crée un Box Plot (qualitative × quantitative)."""
    fig = px.box(
        df,
        x=x,
        y=y,
        color=x,
        color_discrete_sequence=COLOR_PALETTE,
        title=title,
    )
    fig.update_layout(
        template="plotly_dark",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", size=13),
        title_font_size=18,
        showlegend=False,
    )
    return fig


def create_line_chart(df: pd.DataFrame, x: str, y: str, title: str) -> go.Figure:
    """Crée un Line Chart pour les séries temporelles."""
    df_sorted = df.sort_values(by=x)
    fig = px.line(
        df_sorted,
        x=x,
        y=y,
        title=title,
        color_discrete_sequence=["#667eea"],
        markers=True,
    )
    fig.update_layout(
        template="plotly_dark",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", size=13),
        title_font_size=18,
    )
    fig.update_traces(line=dict(width=2.5), marker=dict(size=5))
    return fig


def create_correlation_heatmap(df: pd.DataFrame, quant_cols: list) -> go.Figure:
    """Crée une Heatmap de corrélation pour les variables quantitatives."""
    corr = df[quant_cols].corr()
    fig = go.Figure(
        data=go.Heatmap(
            z=corr.values,
            x=corr.columns.tolist(),
            y=corr.index.tolist(),
            colorscale=[[0, "#764ba2"], [0.5, "#1a1a2e"], [1, "#667eea"]],
            text=np.round(corr.values, 2),
            texttemplate="%{text}",
            textfont=dict(size=11),
            hovertemplate="<b>%{x}</b> vs <b>%{y}</b><br>Corrélation: %{z:.3f}<extra></extra>",
        )
    )
    fig.update_layout(
        title="🔗 Matrice de Corrélation",
        template="plotly_dark",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", size=13),
        title_font_size=18,
        height=550,
    )
    return fig


def create_missing_values_chart(df: pd.DataFrame) -> go.Figure:
    """Crée un Bar Chart des valeurs manquantes par colonne."""
    missing = df.isnull().sum()
    missing = missing[missing > 0].sort_values(ascending=True)
    if missing.empty:
        return None
    fig = px.bar(
        x=missing.values,
        y=missing.index,
        orientation="h",
        labels={"x": "Nombre de valeurs manquantes", "y": "Colonne"},
        title="⚠️ Valeurs Manquantes par Colonne",
        color=missing.values,
        color_continuous_scale=["#667eea", "#f5576c"],
        text_auto=True,
    )
    fig.update_layout(
        template="plotly_dark",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", size=13),
        title_font_size=18,
        showlegend=False,
        coloraxis_showscale=False,
    )
    return fig


# =============================================================================
# 🎯 Suggestion automatique du type de graphique
# =============================================================================
def suggest_chart_type(
    type_x: str | None,
    type_y: str | None,
) -> str:
    """
    Suggère automatiquement le type de graphique approprié en fonction
    des types de variables sélectionnées.

    Returns
    -------
    str
        L'un de : 'bar', 'histogram', 'scatter', 'box', 'line'
    """
    if type_x == "Temporelle" or type_y == "Temporelle":
        return "line"
    if type_y is None:
        # Une seule variable sélectionnée
        if type_x == "Qualitative":
            return "bar"
        return "histogram"
    # Deux variables
    if type_x == "Quantitative" and type_y == "Quantitative":
        return "scatter"
    if (type_x == "Qualitative" and type_y == "Quantitative") or \
       (type_x == "Quantitative" and type_y == "Qualitative"):
        return "box"
    if type_x == "Qualitative" and type_y == "Qualitative":
        return "bar"
    return "scatter"


def generate_dynamic_title(
    chart_type: str, x: str, y: str | None
) -> str:
    """Génère un titre dynamique adapté au graphique."""
    titles = {
        "bar": f"📊 Distribution de {x}",
        "histogram": f"📈 Distribution de {x}",
        "scatter": f"🔵 Relation entre {x} et {y}",
        "box": f"📦 Répartition de {y} selon {x}" if y else f"📦 Répartition de {x}",
        "line": f"📉 Évolution de {y} dans le temps" if y else f"📉 Évolution de {x} dans le temps",
    }
    return titles.get(chart_type, f"Graphique : {x}")


# =============================================================================
# 🔧 Filtrage interactif des données
# =============================================================================
def apply_filters(df: pd.DataFrame, var_types: dict) -> pd.DataFrame:
    """
    Affiche des widgets de filtrage dans la sidebar et retourne
    le DataFrame filtré.
    """
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔎 Filtres interactifs")

    filtered_df = df.copy()

    for col, vtype in var_types.items():
        if vtype == "Qualitative":
            unique_vals = df[col].dropna().unique().tolist()
            if len(unique_vals) <= 50:  # Limiter pour l'UX
                selected = st.sidebar.multiselect(
                    f"🏷️ {col}",
                    options=unique_vals,
                    default=unique_vals,
                    key=f"filter_{col}",
                )
                filtered_df = filtered_df[filtered_df[col].isin(selected)]

        elif vtype == "Quantitative":
            col_min = float(df[col].min()) if pd.notna(df[col].min()) else 0.0
            col_max = float(df[col].max()) if pd.notna(df[col].max()) else 1.0
            if col_min < col_max:
                min_val, max_val = st.sidebar.slider(
                    f"🔢 {col}",
                    min_value=col_min,
                    max_value=col_max,
                    value=(col_min, col_max),
                    key=f"filter_{col}",
                )
                filtered_df = filtered_df[
                    (filtered_df[col] >= min_val) & (filtered_df[col] <= max_val)
                ]

        elif vtype == "Temporelle":
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                date_min = df[col].min().date()
                date_max = df[col].max().date()
                if date_min < date_max:
                    start_date = st.sidebar.date_input(
                        f"📅 {col} — début",
                        value=date_min,
                        min_value=date_min,
                        max_value=date_max,
                        key=f"filter_{col}_start",
                    )
                    end_date = st.sidebar.date_input(
                        f"📅 {col} — fin",
                        value=date_max,
                        min_value=date_min,
                        max_value=date_max,
                        key=f"filter_{col}_end",
                    )
                    filtered_df = filtered_df[
                        (filtered_df[col].dt.date >= start_date)
                        & (filtered_df[col].dt.date <= end_date)
                    ]

    return filtered_df


# =============================================================================
# 🖥️ Fonction de rendu des informations du DataFrame
# =============================================================================
def render_dataframe_info(df: pd.DataFrame) -> str:
    """
    Génère un résumé textuel équivalent à df.info() pour un affichage
    propre dans Streamlit (évite la sortie console brute).
    """
    buf = io.StringIO()
    df.info(buf=buf)
    return buf.getvalue()


# =============================================================================
# 📥 Export des données et graphiques
# =============================================================================
def download_filtered_data(df: pd.DataFrame) -> bytes:
    """Convertit le DataFrame filtré en CSV pour téléchargement."""
    return df.to_csv(index=False).encode("utf-8")


def download_figure_png(fig: go.Figure) -> bytes:
    """Convertit une figure Plotly en image PNG pour téléchargement."""
    try:
        return fig.to_image(format="png", width=1200, height=700, scale=2)
    except Exception:
        # kaleido peut ne pas être installé — retour silencieux
        return b""


# =============================================================================
# 🚀 APPLICATION PRINCIPALE
# =============================================================================
def main() -> None:
    """Point d'entrée principal de l'application Streamlit."""

    # --- Injection des styles ---
    inject_custom_css()

    # --- En-tête ---
    st.markdown('<h1 class="main-title">📊 Analyse et Visualisation Interactive de Données</h1>', unsafe_allow_html=True)
    st.markdown(
        '<p class="main-subtitle">Importez vos données, explorez, visualisez et exportez — en quelques clics.</p>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    # --- Sidebar : upload ---
    st.sidebar.image(
        "https://img.icons8.com/fluency/96/combo-chart.png",
        width=64,
    )
    st.sidebar.markdown("## 📂 Import de données")
    uploaded_file = st.sidebar.file_uploader(
        "Importer votre fichier",
        type=["csv", "xlsx"],
        help="Formats supportés : CSV, XLSX",
    )

    if uploaded_file is None:
        # --- Écran d'accueil ---
        st.markdown(
            """
            <div style="text-align:center; padding:4rem 1rem;">
                <p style="font-size:4rem;">📂</p>
                <h2 style="color:#667eea;">Bienvenue !</h2>
                <p style="color:#8892b0; font-size:1.1rem;">
                    Commencez par importer un fichier <b>CSV</b> ou <b>Excel</b>
                    depuis la barre latérale pour démarrer l'analyse.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    # --- Chargement ---
    with st.spinner("⏳ Chargement des données en cours…"):
        df = load_data(uploaded_file)

    if df.empty:
        st.warning("Le fichier importé est vide ou n'a pas pu être lu.")
        return

    # --- Détection des types ---
    var_types = detect_variable_type(df)
    type_table = build_type_table(var_types)

    # Listes pratiques par type
    quant_cols = [c for c, t in var_types.items() if t == "Quantitative"]
    qual_cols = [c for c, t in var_types.items() if t == "Qualitative"]
    temp_cols = [c for c, t in var_types.items() if t == "Temporelle"]

    # --- Filtres interactifs (sidebar) ---
    filtered_df = apply_filters(df, var_types)

    # --- KPI ---
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(
            f"""<div class="kpi-card">
                <div class="kpi-value">{filtered_df.shape[0]:,}</div>
                <div class="kpi-label">Lignes</div>
            </div>""",
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            f"""<div class="kpi-card">
                <div class="kpi-value">{filtered_df.shape[1]}</div>
                <div class="kpi-label">Colonnes</div>
            </div>""",
            unsafe_allow_html=True,
        )
    with k3:
        total_missing = int(filtered_df.isnull().sum().sum())
        st.markdown(
            f"""<div class="kpi-card">
                <div class="kpi-value">{total_missing:,}</div>
                <div class="kpi-label">Valeurs manquantes</div>
            </div>""",
            unsafe_allow_html=True,
        )
    with k4:
        pct_complete = (
            100 * (1 - filtered_df.isnull().sum().sum() / max(filtered_df.size, 1))
        )
        st.markdown(
            f"""<div class="kpi-card">
                <div class="kpi-value">{pct_complete:.1f}%</div>
                <div class="kpi-label">Complétude</div>
            </div>""",
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # =================================================================
    # Onglets principaux
    # =================================================================
    tab_data, tab_analysis, tab_viz, tab_stats = st.tabs(
        ["📋 Données", "🔬 Analyse", "📊 Visualisation", "📈 Statistiques"]
    )

    # -----------------------------------------------------------------
    # Onglet 1 — Données
    # -----------------------------------------------------------------
    with tab_data:
        st.markdown("### 🗂️ Aperçu des données")
        st.dataframe(filtered_df.head(50), use_container_width=True, height=420)

        col_info1, col_info2 = st.columns(2)
        with col_info1:
            st.markdown("#### 📐 Dimensions")
            st.info(
                f"**{filtered_df.shape[0]}** lignes  ×  **{filtered_df.shape[1]}** colonnes"
            )
        with col_info2:
            st.markdown("#### ℹ️ Informations")
            st.code(render_dataframe_info(filtered_df), language="text")

        st.markdown("#### 🏷️ Types de variables détectés")
        # Affichage stylisé du tableau des types
        st.dataframe(
            type_table,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Variable": st.column_config.TextColumn("Variable", width="medium"),
                "Type": st.column_config.TextColumn("Type", width="small"),
            },
        )

        # Téléchargement des données filtrées
        st.markdown("---")
        st.download_button(
            label="⬇️ Télécharger les données filtrées (CSV)",
            data=download_filtered_data(filtered_df),
            file_name="donnees_filtrees.csv",
            mime="text/csv",
        )

    # -----------------------------------------------------------------
    # Onglet 2 — Analyse
    # -----------------------------------------------------------------
    with tab_analysis:
        st.markdown("### 🔍 Analyse exploratoire")

        # -- Valeurs manquantes --
        st.markdown("#### ⚠️ Valeurs manquantes")
        missing_series = filtered_df.isnull().sum()
        missing_non_zero = missing_series[missing_series > 0]

        if missing_non_zero.empty:
            st.success("✅ Aucune valeur manquante détectée !")
        else:
            col_m1, col_m2 = st.columns([1, 2])
            with col_m1:
                missing_display = pd.DataFrame({
                    "Colonne": missing_non_zero.index,
                    "Manquantes": missing_non_zero.values,
                    "% Manquant": (
                        missing_non_zero.values / len(filtered_df) * 100
                    ).round(2),
                })
                st.dataframe(missing_display, hide_index=True, use_container_width=True)
            with col_m2:
                fig_missing = create_missing_values_chart(filtered_df)
                if fig_missing:
                    st.plotly_chart(fig_missing, use_container_width=True)

        # -- Corrélations --
        if len(quant_cols) >= 2:
            st.markdown("---")
            st.markdown("#### 🔗 Matrice de corrélation")
            fig_corr = create_correlation_heatmap(filtered_df, quant_cols)
            st.plotly_chart(fig_corr, use_container_width=True)

    # -----------------------------------------------------------------
    # Onglet 3 — Visualisation
    # -----------------------------------------------------------------
    with tab_viz:
        st.markdown("### 🎨 Visualisation interactive")

        # -- Contrôles dans la sidebar --
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📊 Paramètres du graphique")

        all_cols = list(var_types.keys())
        var_x = st.sidebar.selectbox("Variable X", options=all_cols, index=0)
        var_y = st.sidebar.selectbox(
            "Variable Y (optionnel)",
            options=["— Aucune —"] + all_cols,
            index=0,
        )

        # Colonne couleur pour scatter
        color_col = None
        if var_y != "— Aucune —":
            color_options = ["— Aucune —"] + qual_cols
            color_choice = st.sidebar.selectbox(
                "Variable couleur (optionnel)", options=color_options
            )
            if color_choice != "— Aucune —":
                color_col = color_choice

        # Type de graphique
        type_x = var_types.get(var_x)
        type_y = var_types.get(var_y) if var_y != "— Aucune —" else None

        suggested = suggest_chart_type(type_x, type_y)

        chart_labels = {
            "auto": "🤖 Automatique",
            "histogram": "📈 Histogramme",
            "bar": "📊 Bar Chart",
            "scatter": "🔵 Scatter Plot",
            "box": "📦 Box Plot",
            "line": "📉 Line Chart",
        }

        selected_chart = st.sidebar.selectbox(
            "Type de graphique",
            options=list(chart_labels.keys()),
            format_func=lambda k: chart_labels[k],
            index=0,
        )

        chart_type = suggested if selected_chart == "auto" else selected_chart

        # Titre dynamique
        y_label = var_y if var_y != "— Aucune —" else None
        title = generate_dynamic_title(chart_type, var_x, y_label)

        # -- Construction du graphique --
        fig = None
        try:
            if chart_type == "bar":
                fig = create_bar_chart(filtered_df, var_x, title)

            elif chart_type == "histogram":
                fig = create_histogram(filtered_df, var_x, title)

            elif chart_type == "scatter":
                if y_label is None:
                    st.warning("⚠️ Veuillez sélectionner une variable Y pour le Scatter Plot.")
                else:
                    fig = create_scatter(filtered_df, var_x, y_label, color_col, title)

            elif chart_type == "box":
                if y_label is None:
                    st.warning("⚠️ Veuillez sélectionner une variable Y pour le Box Plot.")
                else:
                    # S'assurer que X = qualitative, Y = quantitative
                    if type_x == "Qualitative":
                        fig = create_box_plot(filtered_df, var_x, y_label, title)
                    else:
                        fig = create_box_plot(filtered_df, y_label, var_x, title)

            elif chart_type == "line":
                if y_label is None:
                    # Si une seule variable temporelle, afficher le count par date
                    if type_x == "Temporelle":
                        temp_group = (
                            filtered_df.groupby(filtered_df[var_x].dt.date)
                            .size()
                            .reset_index(name="Nombre")
                        )
                        temp_group.columns = [var_x, "Nombre"]
                        fig = px.line(
                            temp_group,
                            x=var_x,
                            y="Nombre",
                            title=title,
                            color_discrete_sequence=["#667eea"],
                            markers=True,
                        )
                        fig.update_layout(
                            template="plotly_dark",
                            plot_bgcolor="rgba(0,0,0,0)",
                            paper_bgcolor="rgba(0,0,0,0)",
                            font=dict(family="Inter", size=13),
                            title_font_size=18,
                        )
                    else:
                        st.warning("⚠️ Veuillez sélectionner une variable Y pour le Line Chart.")
                else:
                    fig = create_line_chart(filtered_df, var_x, y_label, title)

        except Exception as e:
            st.error(f"❌ Erreur lors de la création du graphique : {e}")

        if fig:
            st.plotly_chart(fig, use_container_width=True, key="main_chart")

            # Bouton de téléchargement PNG (nécessite kaleido)
            png_bytes = download_figure_png(fig)
            if png_bytes:
                st.download_button(
                    label="🖼️ Télécharger le graphique (PNG)",
                    data=png_bytes,
                    file_name="graphique.png",
                    mime="image/png",
                )

        # Suggestion affichée
        st.sidebar.markdown("---")
        st.sidebar.info(f"💡 Suggestion : **{chart_labels.get(suggested, suggested)}**")

    # -----------------------------------------------------------------
    # Onglet 4 — Statistiques descriptives
    # -----------------------------------------------------------------
    with tab_stats:
        st.markdown("### 📈 Résumé statistique")

        if quant_cols:
            st.markdown("#### 🔢 Variables quantitatives")
            st.dataframe(
                filtered_df[quant_cols].describe().T.style.format("{:.2f}"),
                use_container_width=True,
            )
        else:
            st.info("Aucune variable quantitative détectée.")

        if qual_cols:
            st.markdown("#### 🏷️ Variables qualitatives")
            st.dataframe(
                filtered_df[qual_cols].describe().T,
                use_container_width=True,
            )

        if temp_cols:
            st.markdown("#### 📅 Variables temporelles")
            for col in temp_cols:
                if pd.api.types.is_datetime64_any_dtype(filtered_df[col]):
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Début", str(filtered_df[col].min().date()))
                    c2.metric("Fin", str(filtered_df[col].max().date()))
                    c3.metric(
                        "Durée",
                        f"{(filtered_df[col].max() - filtered_df[col].min()).days} jours",
                    )

    # --- Footer ---
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div style="text-align:center; color:#8892b0; font-size:0.82rem; padding-bottom:1rem;">
            Développé avec ❤️ en Python · Streamlit · Plotly · Pandas &nbsp;|&nbsp;
            Projet universitaire — © 2026
        </div>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# Point d'entrée
# =============================================================================
if __name__ == "__main__":
    main()
