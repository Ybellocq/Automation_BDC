# 🎯 Automatisation Bons de Commande CIRKUS

Un outil Python pour automatiser la génération de bons de commande à partir des exports de ventes Cirkus, gérant les formats 10ML et 50ML (PAB - Prêt à Booster).

## 📋 Table des matières

- [Fonctionnalités](#-fonctionnalités)
- [Prérequis](#-prérequis)
- [Installation](#-installation)
- [Structure des fichiers](#-structure-des-fichiers)
- [Utilisation](#-utilisation)
- [Formats supportés](#-formats-supportés)
- [Mapping des produits](#-mapping-des-produits)
- [Exemples](#-exemples)
- [Dépannage](#-dépannage)
- [Contribution](#-contribution)

## ✨ Fonctionnalités

- **Traitement automatique** des exports de ventes CSV/XLSX
- **Support double format** : flacons 10ML et 50ML (PAB)
- **Génération Excel** avec conservation du formatage (via template)
- **Génération CSV** de secours si pas de template
- **Sélection flexible** des clients (individuel, multiple, ou tous)
- **Mapping intelligent** des noms de produits entre export et bon de commande
- **Fusion automatique** des données 10ML et 50ML par client
- **Calcul automatique** des totaux par format

## 🔧 Prérequis

- **Python 3.6+**
- **Système d'exploitation** : Windows, macOS, Linux
- **Fichiers requis** :
  - Export des ventes (CSV ou XLSX)
  - Template de bon de commande Excel (optionnel mais recommandé)

## 📦 Installation

### 1. Cloner ou télécharger le projet

```bash
git clone [URL_DU_REPO]
cd cirkus-order-automation
```

### 2. Installer les dépendances

#### Installation complète (recommandée)
```bash
pip install openpyxl
```

#### Installation minimale (CSV uniquement)
```bash
# Aucune dépendance externe requise
# Les modules utilisés sont inclus dans Python standard
```

### 3. Préparer vos fichiers

- **Template Excel** : Préparez votre modèle de bon de commande (.xlsx)
- **Exports de ventes** : Exportez vos données depuis votre système de vente

## 📁 Structure des fichiers

```
projet/
├── cirkus_automation.py    # Script principal
├── bons_commande/         # Dossier de sortie (créé automatiquement)
│   ├── BON_COMMANDE_CLIENT1_20241215.xlsx
│   └── BON_COMMANDE_CLIENT2_20241215.xlsx
├── templates/
│   └── template_bon_commande.xlsx
└── exports/
    ├── export_10ml.xlsx
    └── export_50ml.xlsx
```

## 🚀 Utilisation

### Lancement du script

```bash
python cirkus_automation.py
```

### Étapes d'exécution

1. **Template Excel** : Indiquez le chemin vers votre template
   ```
   📁 Chemin vers votre template Excel : ./templates/template.xlsx
   ```

2. **Export 10ML** : Indiquez le fichier d'export des ventes 10ML
   ```
   📁 Chemin vers votre fichier d'export 10ML : ./exports/ventes_10ml.xlsx
   ```

3. **Export 50ML** : Indiquez le fichier d'export des ventes 50ML (optionnel)
   ```
   📁 Chemin vers votre fichier d'export 50ML : ./exports/ventes_50ml.xlsx
   ```

4. **Sélection des clients** :
   - Un client : `3`
   - Plusieurs clients : `1,2,3` ou `1,4,6`
   - Tous les clients : `all` ou `tous`
   - Par nom : `Tabac Martin`

### Options de sélection

| Saisie | Description |
|--------|-------------|
| `3` | Sélectionne le client n°3 |
| `1,2,3` | Sélectionne les clients 1, 2 et 3 |
| `all` | Sélectionne tous les clients |
| `Tabac` | Recherche les clients contenant "Tabac" |
| `exit` | Quitte le programme |

## 📊 Formats supportés

### Fichiers d'entrée
- **CSV** : Séparateur virgule, encodage UTF-8
- **XLSX** : Excel moderne
- **XLS** : Excel ancien (avec openpyxl)

### Fichiers de sortie
- **XLSX** : Format Excel avec formatage conservé (recommandé)
- **CSV** : Format de secours si pas de template Excel

## 🔄 Mapping des produits

### Catégories supportées

| Catégorie | Exemples de produits |
|-----------|---------------------|
| **CLASSICS** | CLASSIC FR, CLASSIC RY4, CLASSIC BLEND |
| **FRUITÉS** | MANGUE FRAMBOISE, FRUITS ROUGES, TROPICAL |
| **FRAIS** | MENTHE POLAIRE, CASSIS FRAIS, LEMON ICE |
| **GIVRÉS** | HANS LÉGEL, AL K'POMME, MÛRE A POINT |
| **GOURMANDS** | CARAMEL, CAFE EXPRESSO, NOUGAT |

### Conversion automatique

Le script convertit automatiquement :
- `CLASSIC FR - 10ML` → `CLASSIC FR` (colonne TOTAL)
- `CLASSIC FR - 50ML` → `CLASSIC FR` (colonne PAB 50ML)

## 💡 Exemples

### Exemple 1 : Client unique avec 10ML seulement

```bash
python cirkus_automation.py
```

```
📁 Template : ./template.xlsx
📁 Export 10ML : ./ventes_10ml.csv
📁 Export 50ML : [Entrée vide]

Clients disponibles :
1. Tabac Martin (10ML: 5, 50ML: 0)

🎯 Sélection : 1
✅ BON_COMMANDE_Tabac_Martin_20241215.xlsx créé
```

### Exemple 2 : Plusieurs clients avec 10ML + 50ML

```bash
📁 Template : ./template.xlsx
📁 Export 10ML : ./ventes_10ml.xlsx
📁 Export 50ML : ./ventes_50ml.xlsx

Clients disponibles :
1. Tabac Central (10ML: 8, 50ML: 3)
2. Buraliste Dupont (10ML: 12, 50ML: 5)
3. Shop Nicotine (10ML: 6, 50ML: 2)

🎯 Sélection : 1,3
✅ 2 bons de commande créés
```

## 🛠️ Dépannage

### Problèmes courants

#### ❌ "openpyxl non installé"
```bash
pip install openpyxl
```

#### ❌ "Template Excel non trouvé"
- Vérifiez le chemin du fichier
- Utilisez des guillemets si le chemin contient des espaces
- Le script générera un CSV en cas d'échec

#### ❌ "Impossible de trouver les en-têtes"
- Vérifiez que votre export contient bien des colonnes avec "- 10ML" ou "- 50ML"
- La première ligne doit contenir les noms des clients
- La ligne suivante doit contenir les noms des produits

#### ❌ "Fichier vide ou invalide"
- Vérifiez que le fichier n'est pas vide
- Assurez-vous qu'il contient au moins 2 lignes (en-têtes + données)
- Vérifiez l'encodage (UTF-8 recommandé)

### Structure attendue des exports

#### Format CSV/XLSX d'export
```
Client          | CLASSIC FR - 10ML | MANGUE FRAMBOISE - 10ML | CLASSIC FR - 50ML
Tabac Martin    | 12               | 6                       | 2
Buraliste Dubois| 8                | 0                       | 1
```

#### Structure du template Excel
Le template doit contenir :
- Une cellule "CLIENT" pour le nom
- Des lignes avec les noms de produits (ex: "CLASSIC FR")
- Les colonnes appropriées pour les quantités

### Messages d'erreur et solutions

| Message | Cause probable | Solution |
|---------|---------------|----------|
| `❌ Format non supporté` | Extension de fichier invalide | Utilisez .csv, .xlsx ou .xls |
| `❌ Erreur lors du chargement` | Fichier corrompu ou verrouillé | Vérifiez que le fichier n'est pas ouvert |
| `❌ Aucun client sélectionné` | Sélection invalide | Respectez le format de sélection |

## 🔧 Configuration avancée

### Ajouter de nouveaux produits

Pour ajouter des produits au mapping, modifiez les dictionnaires dans la classe :

```python
self.product_mapping_10ml = {
    'NOUVEAU PRODUIT - 10ML': 'NOUVEAU PRODUIT',
    # ... autres produits
}

self.product_mapping_50ml = {
    'NOUVEAU PRODUIT - 50ML': 'NOUVEAU PRODUIT',
    # ... autres produits
}
```

### Personnaliser le format de sortie

Modifiez la méthode `create_order_form_csv_enhanced()` pour adapter :
- Les colonnes du bon de commande
- Les catégories de produits
- Le format des noms de fichiers

## 🤝 Contribution

### Signaler un bug
1. Vérifiez que le bug n'est pas déjà signalé
2. Fournissez un exemple de fichier d'export (anonymisé)
3. Indiquez la version de Python utilisée
4. Décrivez les étapes pour reproduire le problème

### Proposer une amélioration
1. Décrivez le besoin
2. Proposez une solution
3. Testez votre modification
4. Documentez les changements

## 📝 Notes techniques

### Sécurité
- Les fichiers sont traités localement
- Aucune donnée n'est envoyée sur internet
- Les fichiers temporaires sont automatiquement nettoyés

### Performance
- Traitement optimisé pour des milliers de lignes
- Utilisation de la mémoire maîtrisée
- Génération parallèle possible (modification requise)

### Compatibilité
- Testé sur Windows 10/11
- Compatible macOS et Linux
- Python 3.6 minimum requis


---

**Version** : 2.0  
**Dernière mise à jour** : Juin 2025 
**Auteur** : [Yoann]  
**Support** : [bellocq.yoann@gmail.com]