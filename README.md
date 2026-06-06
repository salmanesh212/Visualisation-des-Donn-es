# 📊 Analyse et Visualisation Interactive de Données — Version Pro

Application web professionnelle de **Data Analytics** développée avec **Streamlit**.  
Véritable mini **Power BI / Tableau** en Python, adaptée à une soutenance universitaire.

> Projet universitaire — ESITH

---

## ✨ Fonctionnalités

### 📋 Import & Détection
| Fonctionnalité | Description |
|---|---|
| 📂 Import | CSV et Excel (`.xlsx`) |
| 🔍 Détection auto | Quantitative / Qualitative / Temporelle |
| 🏷️ KPI dynamiques | Lignes, colonnes, complétude, manquantes |

### 📋 Profiling Dataset
- Analyse automatique colonne par colonne
- Nom, type, uniques, manquantes, min, max, moyenne, médiane, écart-type
- Tableau interactif avec tri et recherche

### 🧹 Qualité des Données
- Tableau des valeurs manquantes trié par % décroissant
- Bar Chart interactif des % manquants
- Heatmap des motifs de données manquantes
- Synthèse automatique avec insights en langage naturel

### 🔗 Corrélations Avancées
- Heatmap interactive avec zoom et hover
- Top corrélations positives et négatives
- Classification : forte (|r|>0.7), modérée, faible
- Insights automatiques

### 📊 16 Types de Graphiques
| Graphique | Usage |
|---|---|
| Histogramme | Variable quantitative |
| Histogramme + KDE | Quantitative + densité |
| Bar Chart | Variable qualitative |
| Pie Chart | Répartition qualitative |
| Donut Chart | Répartition moderne |
| Scatter Plot | Quant × Quant |
| Bubble Chart | Quant × Quant + taille + couleur |
| Box Plot | Qual × Quant |
| Violin Plot | Distribution + densité |
| Strip Plot | Dispersion individuelle |
| Line Chart | Séries temporelles |
| Area Chart | Évolution temporelle |
| Scatter Matrix | Multi-variables |
| Treemap | Hiérarchie catégorielle |
| Sunburst | Hiérarchie circulaire |

### 🎨 Personnalisation Complète
- **7 thèmes** : Plotly Dark, Plotly, Plotly White, GGPlot2, Seaborn, Presentation, Simple White
- **Dimensions** : hauteur, largeur fixe optionnelle
- **Couleurs** : 8 palettes + color pickers
- **Axes** : titres personnalisés, rotation des labels
- **Légende** : afficher/masquer, position (Top/Bottom/Left/Right)
- **Grille** : afficher/masquer
- **Étiquettes** : valeurs, pourcentages

### 📥 Export Complet
| Format | Contenu |
|---|---|
| CSV | Données filtrées, profiling, statistiques, corrélations |
| Excel | Données filtrées |
| PNG / SVG / JPEG / PDF | Graphiques individuels |
| Rapport PDF | Analyse complète automatique |

---

## 🛠️ Technologies

- **Python 3.9+**
- [Streamlit](https://streamlit.io/)
- [Pandas](https://pandas.pydata.org/) & [NumPy](https://numpy.org/)
- [Plotly Express](https://plotly.com/python/plotly-express/)
- [Openpyxl](https://openpyxl.readthedocs.io/)
- [Kaleido](https://github.com/plotly/Kaleido) (export images)
- [fpdf2](https://py-pdf.github.io/fpdf2/) (rapport PDF)

---

## 🚀 Installation & Exécution

### 1. Cloner le dépôt

```bash
git clone https://github.com/<votre-utilisateur>/<votre-repo>.git
cd <votre-repo>
```

### 2. Environnement virtuel (recommandé)

```bash
python -m venv venv
venv\Scripts\activate        # Windows
source venv/bin/activate     # Linux / macOS
```

### 3. Installer les dépendances

```bash
pip install -r requirements.txt
```

### 4. Lancer

```bash
streamlit run app.py
```

→ Ouvrir **http://localhost:8501**

---

## ☁️ Déploiement sur Streamlit Cloud

1. Pousser sur GitHub : `app.py`, `requirements.txt`
2. Aller sur [share.streamlit.io](https://share.streamlit.io/)
3. **New app** → sélectionner le repo, branche `main`, fichier `app.py`
4. **Deploy** 🚀

---

## 📁 Structure

```
├── app.py               # Application principale (~900 lignes)
├── requirements.txt     # Dépendances
└── README.md            # Documentation
```

---

## 📸 Onglets

| Onglet | Contenu |
|---|---|
| 📋 Données | Aperçu, dimensions, types |
| 📋 Profiling | Analyse colonne par colonne |
| 🧹 Qualité | Valeurs manquantes, heatmap, insights |
| 🔗 Corrélations | Heatmap, top +/−, classification |
| 📊 Visualisation | 16 graphiques interactifs |
| 📈 Statistiques | Résumé descriptif complet |
| 📥 Export | CSV, Excel, images, rapport PDF |

---

## 📝 Licence

Projet académique — usage éducatif uniquement.

---

<p align="center">
  Développé avec ❤️ · Python · Streamlit · Plotly · Pandas
</p>
