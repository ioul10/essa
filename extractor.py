import pdfplumber
import pandas as pd
from thefuzz import process
import re
from datetime import datetime

class PDFExtractor:
    def __init__(self):
        # Dictionnaire de référence pour le mapping métier
        self.field_references = {
            'chiffre_affaires': ['chiffre affaires', 'ca', 'ventes', 'revenus', 'turnover'],
            'resultat_net': ['résultat net', 'benefice net', 'perte nette', 'net income'],
            'total_actif': ['total actif', 'actif total', 'total assets'],
            'total_passif': ['total passif', 'passif total', 'total liabilities'],
            'capitaux_propres': ['capitaux propres', 'equity', 'fonds propres'],
            'date_cloture': ['date clôture', 'date arreté', 'closing date', 'exercice'],
            'societe': ['société', 'entreprise', 'company', 'dénomination'],
            'exercice': ['exercice', 'année', 'year', 'period']
        }
        
    def extract_text(self, pdf_file):
        """Étape 1 & 2 : Extraction texte brut et tableaux"""
        full_text = ""
        tables = []
        
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                # Extraction texte
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
                
                # Extraction tableaux
                page_tables = page.extract_tables()
                for table in page_tables:
                    if table:
                        tables.append(table)
        
        return full_text, tables
    
    def clean_data(self, tables):
        """Étape 3 : Nettoyage des données"""
        cleaned_tables = []
        
        for table in tables:
            cleaned_rows = []
            for row in table:
                # Nettoyer chaque cellule
                cleaned_row = []
                for cell in row:
                    if cell is None:
                        cleaned_row.append("")
                    else:
                        # Supprimer espaces multiples, newlines
                        cell_str = str(cell).strip()
                        cell_str = re.sub(r'\s+', ' ', cell_str)
                        cleaned_row.append(cell_str)
                
                # Ignorer lignes vides
                if any(cell != "" for cell in cleaned_row):
                    cleaned_rows.append(cleaned_row)
            
            if len(cleaned_rows) > 0:
                cleaned_tables.append(cleaned_rows)
        
        return cleaned_tables
    
    def filter_useful_lines(self, cleaned_tables, keywords=None):
        """Étape 4 : Filtrage des lignes utiles"""
        if keywords is None:
            keywords = ['total', 'montant', 'prix', 'quantité', 'désignation', 
                       'actif', 'passif', 'résultat', 'chiffre']
        
        useful_tables = []
        
        for table in cleaned_tables:
            # Garder le tableau si au moins une ligne contient un mot-clé
            has_useful_line = False
            for row in table:
                row_text = ' '.join(row).lower()
                if any(kw in row_text for kw in keywords):
                    has_useful_line = True
                    break
            
            if has_useful_line or len(table) > 2: # Ou si tableau substantiel
                useful_tables.append(table)
        
        return useful_tables
    
    def extract_amounts(self, text):
        """Étape 5 : Extraction des montants (Regex)"""
        # Patterns pour différents formats de nombres
        patterns = [
            r'(\d{1,3}(?:[\s\.]\d{3})*(?:,\d{2})?)',  # 1.000,00 ou 1 000,00
            r'(\d{1,3}(?:[,\s]\d{3})*(?:\.\d{2})?)',  # 1,000.00 ou 1 000.00
            r'(\d+,\d{2})',  # 123,45
        ]
        
        amounts = []
        for pattern in patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                # Normaliser le format (enlever espaces, convertir virgule)
                normalized = match.replace(' ', '').replace('.', '').replace(',', '.')
                try:
                    amount = float(normalized)
                    amounts.append(amount)
                except:
                    pass
        
        return amounts
    
    def fuzzy_map_field(self, text_line):
        """Étape 6 : Mapping métier avec Fuzzy Matching"""
        text_lower = text_line.lower()
        
        best_match = None
        best_score = 0
        
        for field, variations in self.field_references.items():
            for variation in variations:
                score = process.extractOne(variation, [text_lower])[1]
                if score > best_score and score >= 70:  # Seuil de confiance 70%
                    best_score = score
                    best_match = field
        
        return best_match, best_score
    
    def extract_key_values(self, full_text):
        """Extraire les paires Clé-Valeur du texte"""
        key_values = {}
        
        lines = full_text.split('\n')
        for line in lines:
            if ':' in line or '  ' in line:
                field, confidence = self.fuzzy_map_field(line)
                if field:
                    # Extraire la valeur (après : ou en fin de ligne)
                    parts = re.split(r':|\s{2,}', line, 1)
                    if len(parts) > 1:
                        value = parts[1].strip()
                        # Essayer de convertir en nombre
                        amounts = self.extract_amounts(value)
                        if amounts:
                            key_values[field] = amounts[0]
                        else:
                            key_values[field] = value
        
        return key_values
    
    def create_template_df(self):
        """Créer un DataFrame template vide"""
        template = {
            'Champ': list(self.field_references.keys()),
            'Valeur Extraite': [''] * len(self.field_references),
            'Confiance (%)': [''] * len(self.field_references),
            'Source': [''] * len(self.field_references)
        }
        return pd.DataFrame(template)
    
    def fill_template(self, key_values, template_df):
        """Étape 7 : Remplir le template avec les données extraites"""
        for idx, row in template_df.iterrows():
            field = row['Champ']
            if field in key_values:
                template_df.at[idx, 'Valeur Extraite'] = key_values[field]
                template_df.at[idx, 'Source'] = 'Texte PDF'
        
        return template_df
