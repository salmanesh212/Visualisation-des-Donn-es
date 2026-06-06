# 📊 Analyse et Visualisation Interactive de Données

Application web professionnelle de **Data Science** développée avec **Streamlit**, permettant l'analyse exploratoire et la visualisation interactive de jeux de données (CSV / Excel).

> Projet universitaire — ESITH

---

## ✨ Fonctionnalités

| Fonctionnalité | Description |
|---|---|
| 📂 Import de fichiers | CSV et Excel (`.xlsx`) |
| 🔍 Détection automatique | Classification des colonnes en *Quantitative*, *Qualitative* ou *Temporelle* |
| 📊 Visualisations interactives | Histogramme, Bar Chart, Scatter Plot, Box Plot, Line Chart |
| 🤖 Suggestion automatique | Le graphique le plus adapté est proposé selon les types de variables |
| 🔎 Filtres dynamiques | Multiselect, sliders, sélection de dates |
| 📈 Statistiques descriptives | `describe()`, valeurs manquantes, corrélations |
| 🔗 Matrice de corrélation | Heatmap interactive (Plotly) |
| ⬇️ Export | Téléchargement des données filtrées (CSV) et du graphique (PNG) |
| 🎨 Interface moderne | Thème sombre, KPI, onglets, animations CSS |

---

## 🛠️ Technologies

- **Python 3.9+**
- [Streamlit](https://streamlit.io/)
- [Pandas](https://pandas.pydata.org/)
- [NumPy](https://numpy.org/)
- [Plotly Express](https://plotly.com/python/plotly-express/)
- [Openpyxl](https://openpyxl.readthedocs.io/)

---

## 🚀 Installation & Exécution locale

### 1. Cloner le dépôt

```bash
git clone https://github.com/<votre-utilisateur>/<votre-repo>.git
cd <votre-repo>
```

### 2. Créer un environnement virtuel (recommandé)

```bash
python -m venv venv
source venv/bin/activate   # Linux / macOS
venv\Scripts\activate      # Windows
```

### 3. Installer les dépendances

```bash
pip install -r requirements.txt
```

### 4. Lancer l'application

```bash
streamlit run app.py
```

L'application s'ouvrira automatiquement dans votre navigateur à l'adresse :  
`http://localhost:8501`

---

## ☁️ Déploiement sur Streamlit Cloud

### Étapes

1. **Pousser le code** sur un dépôt GitHub public contenant :
   - `app.py`
   - `requirements.txt`

2. Se rendre sur [share.streamlit.io](https://share.streamlit.io/)

3. Cliquer sur **« New app »**

4. Renseigner :
   - **Repository** : `<votre-utilisateur>/<votre-repo>`
   - **Branch** : `main`
   - **Main file path** : `app.py`

5. Cliquer sur **Deploy** 🚀

> L'application sera accessible via une URL publique de type :  
> `https://<votre-app>.streamlit.app`

---

## 📁 Structure du projet

```
.
├── app.py               # Application Streamlit principale
├── requirements.txt     # Dépendances Python
└── README.md            # Ce fichier
```

---

## 📸 Aperçu

Après import d'un fichier :

- **Onglet Données** — Aperçu du dataset, dimensions, types détectés
- **Onglet Analyse** — Valeurs manquantes, matrice de corrélation
- **Onglet Visualisation** — Graphiques interactifs avec choix automatique ou manuel
- **Onglet Statistiques** — Résumé statistique complet

---

## 📝 Licence

Projet académique — usage éducatif uniquement.

---

<p align="center">
  Développé avec ❤️ en Python · Streamlit · Plotly · Pandas
</p>
