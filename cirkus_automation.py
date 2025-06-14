import csv
import json
import os
from datetime import datetime
import zipfile
import xml.etree.ElementTree as ET
import shutil

# Installation requise : pip install openpyxl
try:
    from openpyxl import load_workbook, Workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("⚠️ openpyxl non installé. Installez-le avec : pip install openpyxl")

class CirkusOrderAutomation:
    def __init__(self):
        # Mapping entre les noms de produits de l'export et du bon de commande
        self.product_mapping = {
            # CLASSICS
            'CLASSIC FR - 10ML': 'CLASSIC FR',
            'CLASSIC RY4 - 10ML': 'CLASSIC RY4',
            'CLASSIC BLEND - 10ML': 'CLASSIC BLEND',
            'CLASSIC US - 10ML': 'CLASSIC US',
            'CLASSIC ORIGINAL - 10ML': 'CLASSIC ORIGINAL',
            'CLASSIC MENTHE - 10ML': 'CLASSIC MENTHE',
            'CLASSIC BLOND - 10ML': 'CLASSIC BLOND',
            'CLASSIC MENTHOL - 10ML': 'CLASSIC MENTHOL',
            'CLASSIC CERISE - 10ML': 'CLASSIC CERISE',
            'CLASSIC GOLD - 10ML': 'CLASSIC GOLD',
            'CLASSIC WHITE - 10ML': 'CLASSIC WHITE',
            
            # FRUITÉS
            'MANGUE FRAMBOISE - 10ML': 'MANGUE FRAMBOISE',
            'FRUITS ROUGES - 10ML': 'FRUITS ROUGES',
            'PASTEQUE MIX - 10ML': 'PASTEQUE MIX',
            'FRAMBOISE BLEUE - 10ML': 'FRAMBOISE BLEUE',
            'FRAISE KIWI - 10ML': 'FRAISE KIWI',
            'FRAMBOISE LITCHI - 10ML': 'FRAMBOISE LITCHI',
            'BONBON FRAISE - 10ML': 'BONBON FRAISE',
            'TROPICAL - 10ML': 'TROPICAL',
            'FRUIT DU DRAGON - 10ML': 'FRUIT DU DRAGON',
            'BONBON CERISE - 10ML': 'BONBON CERISE',
            'MANGUE PASSION VANILLE - 10ML': 'MANGUE PASSION VANILLE',
            'PINA FRAISE - 10ML': 'PINA FRAISE',
            'BONBON BANANE - 10ML': 'BONBON BANANE',
            
            # FRAIS
            'MENTHE POLAIRE - 10ML': 'MENTHE POLAIRE',
            'CASSIS FRAIS - 10ML': 'CASSIS FRAIS',
            'ABSINTHE ROUGE - 10ML': 'ABSINTHE ROUGE',
            'LEMON ICE - 10ML': 'LEMON ICE',
            'MENTHE CHLOROPHYLLE - 10ML': 'MENTHE CHLOROPHYLLE',
            'FRAISE MENTHE - 10ML': 'FRAISE MENTHE',
            
            # GIVRÉS
            'HANS LEGEL - 10ML': 'HANS LÉGEL (XTRA GIVRÉE)',
            'AL K\'POMME - 10ML': 'AL K\'POMME',
            'MURE A POINT - 10ML': 'MÛRE A POINT',
            'INST\'AGRUMES - 10ML': 'INST\'AGRUMES',
            'GARDE LA PECHE - 10ML': 'GARDE LA PÊCHE',
            'MANGUE DE SOLEIL - 10ML': 'MANGUE DE SOLEIL',
            'PRENDS LE MELON - 10ML': 'PRENDS LE MELON',
            'CASSIS CLAY - 10ML': 'CASSIS CLAY',
            'SODA RYAN - 10ML': 'SODA RYAN',
            
            # GOURMANDS
            'CARAMEL - 10ML': 'CARAMEL',
            'CAFE EXPRESSO - 10ML': 'CAFE EXPRESSO',
            'NOUGAT - 10ML': 'NOUGAT',
            'SWEET - 10ML': 'SWEET',
            'GOURMET - 10ML': 'GOURMET',
            'BRAVE - 10ML': 'BRAVE',
            'RESERVE - 10ML': 'RESERVE',
            'LOFTY - 10ML': 'LOFTY',
            'CHEESECAKE CITRON YUZU - 10ML': 'CHEESECAKE CITRON YUZU',
            'CACAHUETE CRUNCHY - 10ML': 'CACAHUETE CRUNCHY',
            'NOISETTE GOURMANDE - 10ML': 'NOISETTE GOURMANDE',
            'SAVAGE - 10ML': 'CLASSIC SAVAGE'
        }
        
        self.template_path = None

    def set_template_path(self, template_path):
        """Définit le chemin vers le template Excel"""
        if os.path.exists(template_path):
            self.template_path = template_path
            print(f"✅ Template défini : {template_path}")
            return True
        else:
            print(f"❌ Template non trouvé : {template_path}")
            return False

    def read_excel_as_csv(self, file_path):
        """Lit un fichier Excel en le convertissant d'abord en CSV"""
        try:
            # Vérifier si c'est un fichier .xlsx
            if not file_path.lower().endswith('.xlsx'):
                print("❌ Le fichier doit être au format .xlsx")
                return None
            
            # Essayer d'abord avec openpyxl si disponible
            if OPENPYXL_AVAILABLE:
                return self.read_excel_with_openpyxl(file_path)
            else:
                # Méthode de fallback
                return self.extract_xlsx_data(file_path)
            
        except Exception as e:
            print(f"❌ Erreur lors de la lecture : {e}")
            return None

    def read_excel_with_openpyxl(self, file_path):
        """Lit un fichier Excel avec openpyxl"""
        try:
            workbook = load_workbook(file_path, data_only=True)
            sheet = workbook.active
            
            data = []
            for row in sheet.iter_rows(values_only=True):
                data.append([str(cell) if cell is not None else '' for cell in row])
            
            workbook.close()
            return data
            
        except Exception as e:
            print(f"❌ Erreur openpyxl : {e}")
            return None

    def extract_xlsx_data(self, file_path):
        """Extrait les données d'un fichier XLSX sans openpyxl (méthode de fallback)"""
        try:
            data = []
            
            # Ouvrir le fichier XLSX comme un ZIP
            with zipfile.ZipFile(file_path, 'r') as zip_file:
                # Lire le fichier des chaînes partagées
                shared_strings = []
                try:
                    with zip_file.open('xl/sharedStrings.xml') as f:
                        tree = ET.parse(f)
                        root = tree.getroot()
                        ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                        for si in root.findall('.//si', ns):
                            t = si.find('.//t', ns)
                            if t is not None:
                                shared_strings.append(t.text or '')
                            else:
                                shared_strings.append('')
                except:
                    pass
                
                # Lire la feuille de calcul
                with zip_file.open('xl/worksheets/sheet1.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    
                    rows = {}
                    for row in root.findall('.//row', ns):
                        row_num = int(row.get('r', 0))
                        rows[row_num] = {}
                        
                        for cell in row.findall('.//c', ns):
                            cell_ref = cell.get('r', '')
                            col = ''.join([c for c in cell_ref if c.isalpha()])
                            
                            value_elem = cell.find('.//v', ns)
                            if value_elem is not None:
                                cell_type = cell.get('t', '')
                                if cell_type == 's':  # Chaîne partagée
                                    try:
                                        idx = int(value_elem.text)
                                        if idx < len(shared_strings):
                                            rows[row_num][col] = shared_strings[idx]
                                        else:
                                            rows[row_num][col] = ''
                                    except:
                                        rows[row_num][col] = value_elem.text or ''
                                else:
                                    rows[row_num][col] = value_elem.text or ''
                            else:
                                rows[row_num][col] = ''
                    
                    # Convertir en format lisible
                    if rows:
                        max_row = max(rows.keys())
                        # Déterminer le nombre de colonnes
                        all_cols = set()
                        for row_data in rows.values():
                            all_cols.update(row_data.keys())
                        
                        # Convertir les lettres de colonnes en nombres
                        def col_letter_to_num(letter):
                            num = 0
                            for c in letter:
                                num = num * 26 + (ord(c) - ord('A') + 1)
                            return num
                        
                        sorted_cols = sorted(all_cols, key=col_letter_to_num)
                        
                        # Extraire les données
                        for row_num in sorted(rows.keys()):
                            row_data = []
                            for col in sorted_cols:
                                row_data.append(rows[row_num].get(col, ''))
                            data.append(row_data)
            
            return data
            
        except Exception as e:
            print(f"❌ Erreur lors de l'extraction XLSX : {e}")
            print("💡 Essayez d'exporter votre fichier Excel en CSV et utilisez le fichier CSV à la place")
            return None

    def load_sales_data(self, file_path):
        """Charge les données de ventes depuis le fichier"""
        try:
            if file_path.lower().endswith('.csv'):
                return self.load_csv_data(file_path)
            else:
                return self.read_excel_as_csv(file_path)
        except Exception as e:
            print(f"❌ Erreur lors du chargement : {e}")
            return None

    def load_csv_data(self, file_path):
        """Charge un fichier CSV"""
        try:
            data = []
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for row in reader:
                    data.append(row)
            return data
        except Exception as e:
            print(f"❌ Erreur CSV : {e}")
            return None

    def parse_sales_data(self, raw_data):
        """Parse les données de ventes"""
        try:
            if not raw_data or len(raw_data) < 2:
                print("❌ Fichier vide ou invalide")
                return None
            
            # Trouver la ligne d'en-tête (celle qui contient les noms de produits)
            header_row = None
            data_start_row = None
            
            for i, row in enumerate(raw_data):
                if len(row) > 0 and any('- 10ML' in str(cell) for cell in row[1:]):
                    header_row = i
                    data_start_row = i + 1
                    break
            
            if header_row is None:
                print("❌ Impossible de trouver les en-têtes de produits")
                return None
            
            headers = [str(cell).strip() for cell in raw_data[header_row]]
            
            # Extraire les données clients
            clients_data = {}
            for row in raw_data[data_start_row:]:
                if len(row) > 0 and str(row[0]).strip():
                    client_name = str(row[0]).strip()
                    if client_name:
                        client_orders = {}
                        for j, value in enumerate(row[1:], 1):
                            if j < len(headers) and str(value).strip() and str(value).strip() != '0':
                                try:
                                    qty = int(float(str(value)))
                                    if qty > 0:
                                        product_name = headers[j]
                                        client_orders[product_name] = qty
                                except:
                                    pass
                        
                        if client_orders:  # Seulement si le client a des commandes
                            clients_data[client_name] = client_orders
            
            print(f"✅ Données chargées : {len(clients_data)} clients trouvés")
            return clients_data
            
        except Exception as e:
            print(f"❌ Erreur lors du parsing : {e}")
            return None

    def get_client_list(self, clients_data):
        """Retourne la liste des clients"""
        return list(clients_data.keys())

    def create_order_form_with_openpyxl(self, client_name, client_orders, output_dir="bons_commande"):
        """Crée le bon de commande avec openpyxl (conserve le formatage)"""
        if not OPENPYXL_AVAILABLE:
            print("❌ openpyxl requis pour conserver le formatage Excel")
            return self.create_order_form_csv(client_name, client_orders, output_dir)
        
        if not self.template_path or not os.path.exists(self.template_path):
            print("❌ Template Excel non défini ou introuvable")
            return self.create_order_form_csv(client_name, client_orders, output_dir)
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Créer le nom du fichier
        date_str = datetime.now().strftime("%Y%m%d")
        filename = f"BON_COMMANDE_{client_name.replace(' ', '_').replace(',', '').replace('/', '_')}_{date_str}.xlsx"
        filepath = os.path.join(output_dir, filename)
        
        try:
            # Copier le template
            shutil.copy2(self.template_path, filepath)
            
            # Ouvrir et modifier le fichier
            workbook = load_workbook(filepath)
            sheet = workbook.active
            
            # Variables pour le suivi
            total_bottles = 0
            client_updated = False
            
            # Parcourir toutes les cellules pour trouver les éléments à modifier
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    
                    cell_value = str(cell.value).strip()
                    
                    # 1. Mise à jour du nom du client
                    if not client_updated and (
                        cell_value.upper() == 'CLIENT' or 
                        'CLIENT' in cell_value.upper() or
                        cell_value.upper() == 'NOM CLIENT' or
                        'NOM' in cell_value.upper() and 'CLIENT' in cell_value.upper()
                    ):
                        # Chercher la cellule adjacente pour y mettre le nom
                        next_cell = sheet.cell(row=cell.row, column=cell.column + 1)
                        if next_cell.value is None or str(next_cell.value).strip() == '':
                            next_cell.value = client_name
                            client_updated = True
                            print(f"✅ Client mis à jour : {client_name}")
                    
                    # 2. Mise à jour des quantités de produits
                    for export_product, quantity in client_orders.items():
                        if export_product in self.product_mapping:
                            template_product = self.product_mapping[export_product]
                            
                            # Vérification flexible du nom du produit
                            if (template_product.upper() in cell_value.upper() or 
                                cell_value.upper() in template_product.upper() or
                                self.fuzzy_match(cell_value, template_product)):
                                
                                # Chercher la cellule de quantité (généralement la suivante)
                                qty_cell = sheet.cell(row=cell.row, column=cell.column + 1)
                                if qty_cell.value is None or str(qty_cell.value).strip() in ['', '0']:
                                    qty_cell.value = quantity
                                    total_bottles += quantity
                                    print(f"✅ {template_product}: {quantity}")
                                break
                    
                    # 3. Mise à jour du total
                    if ('TOTAL' in cell_value.upper() and 
                        ('FLACON' in cell_value.upper() or 'BOUTEILLE' in cell_value.upper())):
                        total_cell = sheet.cell(row=cell.row, column=cell.column + 1)
                        total_cell.value = total_bottles
                        print(f"✅ Total mis à jour : {total_bottles}")
            
            # Sauvegarder le fichier
            workbook.save(filepath)
            workbook.close()
            
            print(f"✅ Bon de commande Excel créé : {filepath}")
            print(f"📦 Total flacons : {total_bottles}")
            
            return filepath
            
        except Exception as e:
            print(f"❌ Erreur lors de la création avec openpyxl : {e}")
            # Fallback vers CSV
            return self.create_order_form_csv(client_name, client_orders, output_dir)

    def fuzzy_match(self, str1, str2, threshold=0.8):
        """Matching approximatif entre deux chaînes"""
        str1 = str1.upper().strip()
        str2 = str2.upper().strip()
        
        # Correspondance exacte
        if str1 == str2:
            return True
        
        # L'une contient l'autre
        if str1 in str2 or str2 in str1:
            return True
        
        # Calcul simple de similarité (Jaccard sur les mots)
        words1 = set(str1.split())
        words2 = set(str2.split())
        
        intersection = words1.intersection(words2)
        union = words1.union(words2)
        
        if len(union) == 0:
            return False
        
        similarity = len(intersection) / len(union)
        return similarity >= threshold

    def create_order_form_csv(self, client_name, client_orders, output_dir="bons_commande"):
        """Crée le bon de commande en CSV (méthode de fallback)"""
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Créer le nom du fichier
        date_str = datetime.now().strftime("%Y%m%d")
        filename = f"BON_COMMANDE_{client_name.replace(' ', '_').replace(',', '').replace('/', '_')}_{date_str}.csv"
        filepath = os.path.join(output_dir, filename)
        
        # Catégories et produits
        categories = {
            'CLASSICS': ['CLASSIC FR', 'CLASSIC RY4', 'CLASSIC BLEND', 'CLASSIC US',
                        'CLASSIC ORIGINAL', 'CLASSIC MENTHE', 'CLASSIC BLOND',
                        'CLASSIC MENTHOL', 'CLASSIC CERISE', 'CLASSIC GOLD', 'CLASSIC WHITE'],
            
            'FRUITÉS': ['MANGUE FRAMBOISE', 'FRUITS ROUGES', 'PASTEQUE MIX', 'FRAMBOISE BLEUE',
                       'FRAISE KIWI', 'FRAMBOISE LITCHI', 'PASSION', 'FRUITY PAMP\'',
                       'BONBON FRAISE', 'TROPICAL', 'FRUIT DU DRAGON', 'BONBON CERISE',
                       'MANGUE PASSION VANILLE', 'PINA FRAISE', 'BONBON BANANE'],
            
            'FRAIS': ['MENTHE POLAIRE', 'CASSIS FRAIS', 'ABSINTHE ROUGE', 'LEMON ICE',
                     'MENTHE CHLOROPHYLLE', 'FRAISE MENTHE', 'ABSINTHE POMME'],
            
            'GIVRÉS': ['HANS LÉGEL (XTRA GIVRÉE)', 'AL K\'POMME', 'MÛRE A POINT',
                      'INST\'AGRUMES', 'GARDE LA PÊCHE', 'MANGUE DE SOLEIL',
                      'PRENDS LE MELON', 'CASSIS CLAY', 'SODA RYAN'],
            
            'GOURMANDS': ['CARAMEL', 'CAFE EXPRESSO', 'NOUGAT', 'SWEET', 'GOURMET',
                         'BRAVE', 'RESERVE', 'LOFTY', 'CHEESECAKE CITRON YUZU',
                         'PECHE GOURMANDE', 'CACAHUETE CRUNCHY', 'VANILLE CUSTARD',
                         'CAFFE LATTE', 'NOISETTE GOURMANDE', 'CLASSIC SAVAGE']
        }
        
        total_bottles = 0
        
        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                
                # En-tête
                writer.writerow(['BON DE COMMANDE CIRKUS'])
                writer.writerow(['CLIENT', client_name])
                writer.writerow(['DIVALTO', ''])
                writer.writerow(['VILLE DE LIVRAISON', ''])
                writer.writerow([''])
                
                # Headers
                headers = ['SAVEUR', '0mg', '3mg', '6mg', '9mg', 'New Taux', '12mg', 
                    '16mg', 'TOTAL', 'SDN 10 mg', 'SDN 20 mg', 'PAB 50ML', 'AROMES 10mL', 'AROMES 30ml']
                writer.writerow(headers)
                
                # Catégories et produits
                for category, products in categories.items():
                    # Ajouter la catégorie
                    writer.writerow([category] + [''] * (len(headers) - 1))
                    
                    # Ajouter les produits
                    for product in products:
                        row = [product] + [0] * (len(headers) - 1)
                        
                        # Chercher le produit dans les commandes
                        for export_product, quantity in client_orders.items():
                            if export_product in self.product_mapping:
                                if self.product_mapping[export_product] == product:
                                    row[8] = quantity  # Colonne TOTAL (position 8)
                                    total_bottles += quantity
                                    break
                                
                        
                        writer.writerow(row)
                
                # Total
                writer.writerow(['TOTAL FLACONS', total_bottles] + [0] * (len(headers) - 2))
            
            print(f"✅ Bon de commande CSV créé : {filepath}")
            return filepath
            
        except Exception as e:
            print(f"❌ Erreur lors de la création du CSV : {e}")
            return None

    def process_client_order(self, sales_file, client_name, clients_data=None):
        """Traite la commande complète pour un client"""
        print(f"🔄 Traitement de la commande pour : {client_name}")
        
        # Si les données ne sont pas fournies, les charger
        if clients_data is None:
            # 1. Charger les données
            raw_data = self.load_sales_data(sales_file)
            if raw_data is None:
                return None
            
            # 2. Parser les données
            clients_data = self.parse_sales_data(raw_data)
            if clients_data is None:
                return None
        
        # 3. Vérifier si le client existe
        if client_name not in clients_data:
            print(f"❌ Client '{client_name}' non trouvé")
            return None
        
        client_orders = clients_data[client_name]
        print(f"📦 {len(client_orders)} produits trouvés pour {client_name}")
        
        # 4. Créer le bon de commande avec le template Excel si disponible
        output_file = self.create_order_form_with_openpyxl(client_name, client_orders)
        
        return output_file

    def parse_client_selection(self, selection, clients):
        """Parse la sélection de clients (ex: "1,2,3" ou "1,4,6")"""
        selected_clients = []
        
        # Diviser par les virgules et traiter chaque élément
        parts = [part.strip() for part in selection.split(',')]
        
        for part in parts:
            if part.isdigit():
                idx = int(part) - 1
                if 0 <= idx < len(clients):
                    selected_clients.append(clients[idx])
                else:
                    print(f"⚠️ Numéro {part} invalide (max: {len(clients)})")
            else:
                # Recherche par nom
                matches = [c for c in clients if part.upper() in c.upper()]
                if len(matches) == 1:
                    selected_clients.append(matches[0])
                elif len(matches) > 1:
                    print(f"🔍 Plusieurs clients trouvés pour '{part}' : {matches}")
                else:
                    print(f"❌ Client '{part}' non trouvé")
        
        return selected_clients

    def process_multiple_orders(self, sales_file, selected_clients, clients_data):
        """Traite les commandes pour plusieurs clients"""
        print(f"\n🚀 Traitement de {len(selected_clients)} commandes...")
        
        created_files = []
        errors = []
        
        for i, client_name in enumerate(selected_clients, 1):
            print(f"\n--- [{i}/{len(selected_clients)}] {client_name} ---")
            try:
                output_file = self.process_client_order(sales_file, client_name, clients_data)
                if output_file:
                    created_files.append(output_file)
                else:
                    errors.append(client_name)
            except Exception as e:
                print(f"❌ Erreur pour {client_name}: {e}")
                errors.append(client_name)
        
        return created_files, errors

def main():
    """Fonction principale"""
    automation = CirkusOrderAutomation()
    
    print("🎯 AUTOMATISATION BON DE COMMANDE CIRKUS")
    print("=" * 50)
    
    # Vérifier openpyxl
    if not OPENPYXL_AVAILABLE:
        print("⚠️ ATTENTION : openpyxl n'est pas installé !")
        print("   Pour conserver le formatage Excel, installez-le avec :")
        print("   pip install openpyxl")
        print("")
    
    # Demander le template Excel
    print("📋 ÉTAPE 1 : Template de bon de commande")
    template_path = input("📁 Chemin vers votre template Excel (.xlsx) : ").strip().strip('"')
    
    if template_path and os.path.exists(template_path):
        automation.set_template_path(template_path)
    else:
        print("⚠️ Template non trouvé, utilisation du format CSV par défaut")
    
    print("\n📊 ÉTAPE 2 : Fichier d'export des ventes")
    print("💡 Conseil : Exportez votre fichier Excel en CSV pour de meilleurs résultats")
    
    # Demander le fichier
    sales_file = input("📁 Chemin vers votre fichier d'export (.xlsx ou .csv) : ").strip().strip('"')
    
    if not os.path.exists(sales_file):
        print("❌ Fichier non trouvé !")
        return
    
    # Charger les données
    print("🔄 Chargement des données...")
    raw_data = automation.load_sales_data(sales_file)
    if raw_data is None:
        return
    
    clients_data = automation.parse_sales_data(raw_data)
    if clients_data is None:
        return
    
    clients = automation.get_client_list(clients_data)
    
    print(f"\n👥 {len(clients)} clients disponibles :")
    for i, client in enumerate(clients, 1):
        print(f"{i:2d}. {client}")
    
    # Instructions pour la sélection
    print("\n" + "=" * 50)
    print("🎯 SÉLECTION DES CLIENTS :")
    print("   • Un seul client : entrez le numéro (ex: 3)")
    print("   • Plusieurs clients : séparez par des virgules (ex: 1,2,3 ou 1,4,6)")
    print("   • Vous pouvez aussi utiliser les noms de clients")
    print("   • Pour tous les clients : tapez 'all' ou 'tous'")
    print("   • Pour quitter : tapez 'exit' ou 'quitter'")
    
    # SUITE DE LA FONCTION main() :
    
    selection = input("\n🎯 Votre sélection : ").strip()
    
    # Traitement de la sélection
    if selection.lower() in ['all', 'tous', 'tout']:
        selected_clients = clients
    else:
        selected_clients = automation.parse_client_selection(selection, clients)
    
    if not selected_clients:
        print("❌ Aucun client sélectionné")
        return
    
    # Traitement des commandes
    created_files, errors = automation.process_multiple_orders(sales_file, selected_clients, clients_data)
    
    # Résultats
    print(f"\n✅ {len(created_files)} bons de commande créés")
    if errors:
        print(f"❌ {len(errors)} erreurs : {errors}")

if __name__ == "__main__":
    main()