# ◈ Patrimonial v2.0

> Suivi de patrimoine personnel — interface sombre style VS Code

---

## Aperçu

Patrimonial est une application de bureau Python qui te permet de suivre
l'évolution de ton patrimoine mois par mois, avec des colonnes de revenus
entièrement personnalisables, un graphique d'évolution, et des exports CSV/Excel.

---

## Prérequis

- **Python 3.8+**
- **tkinter** (inclus par défaut avec Python sur Windows/macOS)
- **openpyxl** *(optionnel, pour l'export Excel)*

```bash
pip install openpyxl
```

---

## Lancement

```bash
python patrimonial.py
```

---

## Fonctionnalités

### 📋 Tableau de données
| Action | Comment |
|---|---|
| **Ajouter une ligne** | Bouton `＋ Ajouter ligne` ou `Ctrl+N` |
| **Supprimer une ligne** | Sélectionner puis `－ Supprimer` ou `Suppr` |
| **Éditer une cellule** | **Double-clic** sur la cellule |
| **Choisir le mois** | Double-clic sur la colonne "Mois" → menu déroulant |
| **Valider une saisie** | `Entrée` |

### ⚙ Gestion des colonnes de revenus
Via le bouton **`⚙ Colonnes`** :
- ➕ **Ajouter** une nouvelle source de revenu
- ✎ **Renommer** une colonne existante
- ✕ **Supprimer** une colonne
- ↑↓ **Réordonner** les colonnes

> Les colonnes configurées sont sauvegardées dans `config.json`.

### 📈 Graphique d'évolution
- Bouton **`📈 Graphique`** ou `Ctrl+G`
- Affiche la courbe du patrimoine total
- Barres colorées pour visualiser les gains (vert) et pertes (rouge)
- Valeurs affichées sur chaque point

### ⬇ Export
| Format | Raccourci |
|---|---|
| **CSV** | Bouton `⬇ Export CSV` ou `Ctrl+E` |
| **Excel (.xlsx)** | Bouton `⬇ Export Excel` *(nécessite openpyxl)* |

---

## Raccourcis clavier

| Raccourci | Action |
|---|---|
| `Ctrl+N` | Ajouter une ligne |
| `Suppr` | Supprimer la ligne sélectionnée |
| `Ctrl+G` | Ouvrir le graphique |
| `Ctrl+E` | Exporter en CSV |

---

## Fichiers générés

| Fichier | Contenu |
|---|---|
| `donnees.json` | Données du tableau (auto-sauvegarde) |
| `config.json` | Configuration des colonnes personnalisées |

---

## Calculs automatiques

- **Total** = somme de toutes les colonnes de revenus de la ligne
- **% Gain/Perte** = variation en % par rapport au mois précédent
- **€ Gain/Perte** = différence en euros par rapport au mois précédent

---

## Structure du projet

```
main.py   ← application principale
donnees.json     ← données (créé automatiquement)
config.json      ← colonnes personnalisées (créé automatiquement)
README.md        ← ce fichier
```

---

*Patrimonial v2.0 — Python + tkinter*
