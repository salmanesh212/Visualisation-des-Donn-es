# 📊 Analyse et Visualisation Interactive de Données 

Application web professionnelle de **Data Analytics** développée avec **Streamlit**.  
Véritable mini **Power BI / Tableau** en Python .

> Projet universitaire — ESITH | Done by BoyWonder 🐦

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

### 🔎 Filtres Avancés
- Filtrage dynamique par colonnes
- Multiselect pour variables qualitatives
- Range sliders pour variables quantitatives
- Date pickers pour séries temporelles

---

## 🛠️ Technologies

- **Python 3.9+**
- [Streamlit](https://streamlit.io/) - Framework web
- [Pandas](https://pandas.pydata.org/) & [NumPy](https://numpy.org/) - Manipulation données
- [Plotly Express](https://plotly.com/python/plotly-express/) - Visualisation interactive
- [Openpyxl](https://openpyxl.readthedocs.io/) - Lecture/écriture Excel
- [Kaleido](https://github.com/plotly/Kaleido) - Export graphiques (PNG, SVG, PDF)
- [fpdf2](https://py-pdf.github.io/fpdf2/) - Génération rapport PDF

---

## 📁 Structure

```
├── app.py               # Application principale (~1500 lignes)
├── requirements.txt     # Dépendances
└── README.md            # Documentation
```

---

## 🚀 Installation

### Prérequis
- Python 3.9 ou supérieur
- pip ou conda

### Étapes

1. **Cloner / Télécharger le projet**
```bash
cd "App Web"
```

2. **Créer un environnement virtuel** (optionnel mais recommandé)
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# macOS/Linux
source venv/bin/activate
```

3. **Installer les dépendances**
```bash
pip install -r requirements.txt
```

4. **Lancer l'application**
```bash
streamlit run app.py
```

L'application s'ouvre automatiquement dans votre navigateur à `http://localhost:8501`

---

## 📖 Utilisation

### Workflow Principal

1. **Importer des données** 📂
   - Cliquez sur "Importer votre fichier" dans la barre latérale
   - Sélectionnez un fichier CSV ou Excel
   - L'app détecte automatiquement les types de colonnes

2. **Explorer les données** 📊
   - **Onglet Données** : Aperçu du dataset, infos techniques
   - **Onglet Profiling** : Analyse détaillée de chaque colonne
   - **Onglet Qualité** : Valeurs manquantes et insights
   - **Onglet Corrélations** : Analyse des relations entre variables

3. **Visualiser** 📈
   - **Onglet Visualisation** : 16 types de graphiques
   - Sélectionnez les variables à représenter
   - Personnalisez le style (thème, couleurs, légende, etc.)

4. **Exporter** 💾
   - **Onglet Export** : Téléchargez les données, statistiques, graphiques
   - Générez un rapport PDF complet avec un seul clic

### Sidebar (Barre latérale)

**🎨 Thème graphique**
- Choisissez entre 7 thèmes prédéfinis

**⚙️ Personnalisation**
- Ajustez hauteur, largeur, tailles de police
- Sélectionnez une palette de couleurs
- Configurez axes, légende et grille

**🔎 Filtres**
- Filtrez les données avant visualisation
- Les graphiques s'adaptent automatiquement

---

## 📸 Onglets Détaillés

| Onglet | Description |
|---|---|
| **📋 Données** | Aperçu brut du dataset avec info technique |
| **📋 Profiling** | Statistiques détaillées colonne par colonne |
| **🧹 Qualité** | Analyse des valeurs manquantes + insights auto |
| **🔗 Corrélations** | Heatmap, top corrélations, classification |
| **📊 Visualisation** | 16 graphiques interactifs avec personnalisation |
| **📈 Statistiques** | Résumé descriptif complet (min, max, moy, etc.) |
| **📥 Export** | Téléchargement données, stats, graphiques, PDF |

---

## ⚙️ Configuration

### Variables d'environnement (optionnel)
```bash
# Activer le mode debug
STREAMLIT_DEBUG=true
```

### Fichiers de données supportés
- ✅ CSV (détection auto du séparateur)
- ✅ Excel (.xlsx, .xls)
- ❌ JSON, Parquet, etc. (non supportés actuellement)

---

## 🔧 Dépannage

### Erreur : "Module not found"
```bash
pip install -r requirements.txt
```

### Erreur : "FPDFUnicodeEncodingException" lors de l'export PDF
- Cette erreur est corrigée. Relancez l'app après le pull des dernières mises à jour.
- Assurez-vous que `fpdf2` est à jour : `pip install --upgrade fpdf2`

### Erreur : "Kaleido not found"
- Pour exporter en PNG/SVG/PDF :
```bash
pip install kaleido
```

### L'app est lente
- Les grandes datasets peuvent ralentir le chargement
- Utilisez les filtres pour réduire la taille
- Considérez un pré-traitement des données (supprimer colonnes inutiles)

---

## 📝 Exemple d'utilisation

1. Téléchargez un dataset (ex: `sales_data.csv`)
2. Importez-le via l'interface
3. Explorez les données dans l'onglet **Profiling**
4. Créez un graphique scatter : X = Âge, Y = Revenu
5. Personalisez : thème Dark, palette Vivid
6. Exportez en PNG
7. Générez un rapport PDF complet

---

## 🎯 Cas d'usage

✅ Soutenance universitaire / Présentations data  
✅ Exploration rapide de datasets  
✅ Dashboard interactif léger  
✅ Analyse EDA (Exploratory Data Analysis)  
✅ Visualisations ad-hoc pour rapports  

---

## 📜 Licence

Projet universitaire — ESITH 2026

---

## 👤 Développé par

**BoyWonder** 🐦

---

## 📞 Support

Pour des bugs ou suggestions :
- Consultez le code source (`app.py`)
- Vérifiez les dépendances (`requirements.txt`)
- Relancez l'application en mode debug

---

**Version** : 2.0 Pro  
**Dernière mise à jour** : Juin 2026  
**Statut** : Stable ✅
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
  Développé by BoyWonder 🐦· Python · Streamlit · Plotly · Pandas
</p>
