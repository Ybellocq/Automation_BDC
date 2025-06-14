# 🎯 Automatisation Bon de Commande Cirkus

## 📋 Description

Cet outil automatise la création de bons de commande pour les produits Cirkus à partir d'un fichier d'export de ventes Excel ou CSV. Il transforme automatiquement les données de ventes clients en bons de commande formatés selon le template standard Cirkus, **sans nécessiter de bibliothèques externes**.

## ✨ Fonctionnalités

- **Import polyvalent** : Support des fichiers Excel (.xlsx) et CSV
- **Zéro dépendance** : Fonctionne uniquement avec Python standard
- **Extraction XLSX native** : Lecture des fichiers Excel sans openpyxl/pandas
- **Mapping intelligent** des produits entre l'export et le bon de commande
- **Génération automatique** de bons de commande au format CSV
- **Classification par catégories** : Classics, Fruités, Frais, Givrés, Gourmands
- **Interface utilisateur** simple et intuitive
- **Sauvegarde automatique** avec horodatage
- **Compatible Excel** : Les fichiers CSV générés s'ouvrent parfaitement dans Excel

## 🛠️ Prérequis

### Logiciels requis
- **Python 3.6 ou plus récent** (aucune bibliothèque externe requise)
- Excel ou LibreOffice (pour visualiser les bons de commande générés)

### Avantages
- ✅ **Aucune installation** de bibliothèques supplémentaires
- ✅ **Portable** : Fonctionne sur tous les systèmes avec Python
- ✅ **Léger** : Utilise uniquement les modules Python standard
- ✅ **Fiable** : Moins de risques de conflits de dépendances

## 📦 Installation

1. **Télécharger le script**
   ```bash
   # Téléchargez directement le fichier cirkus_automation.py
   # Aucune installation supplémentaire requise !
   ```

2. **Vérifier Python**
   ```bash
   python --version
   # ou
   python3 --version
   ```

3. **Tester le script**
   ```bash
   python cirkus_automation.py
   ```

## 🚀 Utilisation

### Formats de fichiers supportés
- **Excel (.xlsx)** : Traitement natif sans bibliothèques externes
- **CSV (.csv)** : Lecture directe et optimisée

### Préparation des données

#### Option 1 : Fichier Excel (.xlsx)
- Utilisez directement votre fichier d'export Excel
- Structure : première colonne = noms clients, colonnes suivantes = produits avec quantités

#### Option 2 : Fichier CSV (Recommandé)
- Exportez votre fichier Excel en CSV pour de meilleures performances
- Dans Excel : Fichier → Enregistrer sous → CSV UTF-8

### Lancement du programme
```bash
python cirkus_automation.py
```

### Étapes d'utilisation

1. **Lancer le script**
   ```bash
   python cirkus_automation.py
   ```

2. **Entrer le chemin du fichier**
   - Supporté : `.xlsx` et `.csv`
   - Le script détecte automatiquement le format

3. **Sélectionner le client**
   - Liste automatique des clients disponibles
   - Choix par numéro ou nom (recherche partielle)

4. **Récupérer le bon de commande**
   - Fichier CSV généré dans `bons_commande/`
   - Format : `BON_COMMANDE_[CLIENT]_[DATE].csv`
   - Compatible Excel et LibreOffice

### Exemple d'utilisation
```
🎯 AUTOMATISATION BON DE COMMANDE CIRKUS
==================================================
💡 Conseil : Exportez votre fichier Excel en CSV pour de meilleurs résultats

📁 Chemin vers votre fichier d'export (.xlsx ou .csv) : export_ventes.xlsx
🔄 Chargement des données...
✅ Données chargées : 25 clients trouvés

👥 25 clients disponibles :
 1. TABAC DUPONT
 2. BUREAU DE TABAC MARTIN
 3. SMOKE SHOP PARIS
 ...

🎯 Entrez le numéro ou le nom du client : 1

🚀 Traitement en cours pour : TABAC DUPONT
📦 12 produits trouvés pour TABAC DUPONT
✅ Bon de commande CSV créé : bons_commande/BON_COMMANDE_TABAC_DUPONT_20250614.csv

✅ TERMINÉ ! Bon de commande créé
📂 Le fichier CSV peut être ouvert dans Excel ou Numbers
💡 Pour convertir en Excel : ouvrez le CSV et sauvegardez-le en .xlsx
```

## 📊 Structure des données

### Format d'entrée
- **Colonne A** : Noms des clients
- **Colonnes B+** : Produits avec quantités (format "NOM PRODUIT - 10ML")
- **Formats acceptés** : Excel (.xlsx) ou CSV (.csv)

### Format de sortie (CSV)
- **En-tête** : Informations client et commande
- **Colonnes** : SAVEUR, 0mg, 3mg, 6mg, 9mg, New Taux, 12mg, 16mg, SDN 10mg, SDN 20mg, PAB 50ML, AROMES 10mL, AROMES 30ml
- **Catégories** : Classification automatique des produits
- **Total** : Calcul automatique du nombre total de flacons
- **Compatible Excel** : S'ouvre directement dans Excel avec le bon formatage

## 🏷️ Produits supportés

### CLASSICS (11 produits)
- CLASSIC FR, CLASSIC RY4, CLASSIC BLEND, CLASSIC US
- CLASSIC ORIGINAL, CLASSIC MENTHE, CLASSIC BLOND
- CLASSIC MENTHOL, CLASSIC CERISE, CLASSIC GOLD, CLASSIC WHITE

### FRUITÉS (15 produits)
- MANGUE FRAMBOISE, FRUITS ROUGES, PASTEQUE MIX
- FRAMBOISE BLEUE, FRAISE KIWI, FRAMBOISE LITCHI
- BONBON FRAISE, TROPICAL, FRUIT DU DRAGON
- BONBON CERISE, MANGUE PASSION VANILLE, PINA FRAISE, BONBON BANANE

### FRAIS (6 produits)
- MENTHE POLAIRE, CASSIS FRAIS, ABSINTHE ROUGE
- LEMON ICE, MENTHE CHLOROPHYLLE, FRAISE MENTHE

### GIVRÉS (9 produits)
- HANS LÉGEL (XTRA GIVRÉE), AL K'POMME, MÛRE A POINT
- INST'AGRUMES, GARDE LA PÊCHE, MANGUE DE SOLEIL
- PRENDS LE MELON, CASSIS CLAY, SODA RYAN

### GOURMANDS (15 produits)
- CARAMEL, CAFE EXPRESSO, NOUGAT, SWEET, GOURMET
- BRAVE, RESERVE, LOFTY, CHEESECAKE CITRON YUZU
- CACAHUETE CRUNCHY, NOISETTE GOURMANDE, CLASSIC SAVAGE

## 📁 Structure des fichiers

```
votre-dossier/
├── cirkus_automation.py          # Script principal (aucune dépendance !)
├── export_ventes.xlsx            # Votre fichier Excel
├── export_ventes.csv             # Ou votre fichier CSV
└── bons_commande/               # Dossier généré automatiquement
    ├── BON_COMMANDE_CLIENT1_20250614.csv
    ├── BON_COMMANDE_CLIENT2_20250614.csv
    └── ...
```

## 🔧 Configuration avancée

### Modifier le mapping des produits
Pour ajouter ou modifier des produits, éditez le dictionnaire `product_mapping` dans la classe `CirkusOrderAutomation` :

```python
self.product_mapping = {
    'NOUVEAU PRODUIT EXPORT - 10ML': 'NOUVEAU PRODUIT BON COMMANDE',
    # ... autres mappings
}
```

### Personnaliser les catégories
Modifiez le dictionnaire `categories` dans la méthode `create_order_form_csv()` pour ajuster l'organisation des produits.

## 🔄 Conversion Excel ↔ CSV

### Excel vers CSV (Recommandé)
```
Dans Excel :
1. Ouvrir votre fichier .xlsx
2. Fichier → Enregistrer sous
3. Choisir "CSV UTF-8 (délimité par des virgules)"
4. Utiliser ce fichier CSV avec le script
```

### CSV vers Excel (Après génération)
```
1. Ouvrir le fichier CSV généré dans Excel
2. Fichier → Enregistrer sous
3. Choisir "Classeur Excel (.xlsx)"
```

## ❌ Résolution de problèmes

### Erreurs communes

**"Fichier non trouvé"**
- Vérifiez le chemin complet vers votre fichier
- Utilisez des guillemets si le chemin contient des espaces
- Exemple : `"C:\Documents\mon fichier.xlsx"`

**"Impossible de trouver les en-têtes de produits"**
- Vérifiez que vos produits ont le format "NOM - 10ML"
- Assurez-vous que la première colonne contient les noms de clients
- Essayez d'exporter en CSV depuis Excel

**"Erreur lors de l'extraction XLSX"**
- Le fichier Excel est peut-être corrompu ou dans un format non standard
- **Solution** : Exportez le fichier en CSV et réessayez
- Fermez Excel avant d'exécuter le script

**"Client non trouvé"**
- Vérifiez l'orthographe exacte du nom du client
- Utilisez le numéro dans la liste plutôt que le nom
- La recherche est sensible à la casse

### Conseils de performance
- **Préférez le CSV** : Plus rapide et plus fiable que l'extraction XLSX
- **Fermez Excel** : Assurez-vous que le fichier n'est pas ouvert ailleurs
- **Vérifiez l'encodage** : Utilisez UTF-8 pour les caractères spéciaux

### Avantages du format CSV
- ✅ **Plus rapide** : Lecture directe sans extraction ZIP
- ✅ **Plus fiable** : Moins de risques d'erreurs de parsing
- ✅ **Plus compatible** : Fonctionne avec tous les tableurs
- ✅ **Plus léger** : Fichiers plus petits et plus faciles à traiter

## 🛠️ Fonctionnalités techniques

### Extraction XLSX native
- **Lecture ZIP** : Traite les fichiers .xlsx comme des archives ZIP
- **Parsing XML** : Extrait les données des feuilles de calcul et chaînes partagées
- **Conversion automatique** : Transforme les références de cellules en données utilisables
- **Gestion d'erreurs** : Fallback vers CSV en cas de problème

### Traitement CSV optimisé
- **Détection d'encodage** : Support UTF-8 pour les caractères spéciaux
- **Parsing intelligent** : Détection automatique des en-têtes et données
- **Nettoyage automatique** : Suppression des espaces et normalisation

## 🔄 Mises à jour

### Version actuelle : 2.0
- ✅ Support natif des fichiers Excel (.xlsx) sans dépendances
- ✅ Support optimisé des fichiers CSV
- ✅ Interface utilisateur améliorée
- ✅ Gestion d'erreurs renforcée
- ✅ Output au format CSV compatible Excel
- ✅ Extraction XLSX basée sur les standards OpenXML

### Fonctionnalités prévues
- Interface graphique (GUI)
- Traitement par lots (plusieurs clients)
- Export direct en Excel (.xlsx)
- Personnalisation avancée des templates

## 🎯 Avantages de cette version

### Simplicité d'installation
- **Zéro configuration** : Fonctionne immédiatement avec Python
- **Portable** : Un seul fichier Python à déployer
- **Compatible** : Fonctionne sur Windows, Mac et Linux

### Robustesse
- **Gestion d'erreurs** : Messages d'aide clairs en cas de problème
- **Fallback intelligent** : Suggestion d'export CSV si l'Excel pose problème
- **Compatibilité étendue** : Support des différents formats d'export

### Performance
- **Extraction native** : Pas de dépendances lourdes
- **Optimisé CSV** : Traitement rapide des gros fichiers
- **Mémoire efficace** : Lecture streaming pour les gros volumes

## 📞 Support

### Ordre de diagnostic
1. **Vérifiez le format** : Préférez CSV pour éviter les problèmes
2. **Contrôlez les données** : Première colonne = clients, format produits = "NOM - 10ML"
3. **Testez avec un petit fichier** : Validez le fonctionnement sur un échantillon
4. **Consultez les messages d'erreur** : Le script donne des indications précises

### Messages d'aide intégrés
- Le script affiche des conseils contextuel
- Suggestions d'export CSV en cas de problème Excel
- Instructions de conversion détaillées

## 📜 Licence

Ce script est fourni tel quel pour automatiser les tâches répétitives de création de bons de commande Cirkus. Il ne nécessite aucune bibliothèque externe et fonctionne uniquement avec Python standard.

---

**Dernière mise à jour** : Juin 2025  
**Compatibilité** : Python 3.6+, Windows/Mac/Linux  
**Dépendances** : Aucune (Python standard uniquement) ✅