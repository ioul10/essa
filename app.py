import streamlit as st
import pdfplumber
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Configuration de la page
st.set_page_config(page_title="PDF Pro vers Excel", page_icon="📊", layout="wide")

st.title("📄 Convertisseur PDF vers Excel (Structure Optimisée)")
st.markdown("""
Cette version améliore la structure des tableaux et applique un style Excel professionnel 
(bordures, en-têtes gras, largeur auto).
""")

# Fonction pour styliser la feuille Excel
def style_excel_sheet(ws, df):
    """
    Applique un style professionnel à la feuille openpyxl
    """
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center_alignment = Alignment(horizontal='center', vertical='center')

    # 1. Styliser les en-têtes (Ligne 1)
    for col in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = center_alignment

    # 2. Styliser le corps du tableau et ajuster les largeurs
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            # Ajustement largeur de colonne (basé sur la longueur du texte)
            col_letter = get_column_letter(col)
            max_length = 0
            # On calcule la largeur max pour cette colonne
            for r in range(1, ws.max_row + 1):
                c_val = str(ws.cell(row=r, column=col).value)
                if len(c_val) > max_length:
                    max_length = len(c_val)
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[col_letter].width = min(adjusted_width, 50) # Max 50 chars

# Fonction de nettoyage du tableau brut
def clean_table_data(table):
    """
    Nettoie les données brutes de pdfplumber (None -> '', strip spaces)
    """
    cleaned_data = []
    for row in table:
        cleaned_row = [str(cell).strip() if cell is not None else "" for cell in row]
        # Ignorer les lignes entièrement vides
        if any(cell != "" for cell in cleaned_row):
            cleaned_data.append(cleaned_row)
    return cleaned_data

uploaded_file = st.file_uploader("Choisissez un fichier PDF", type="pdf")

if uploaded_file is not None:
    st.info("Analyse du PDF en cours...")
    
    try:
        excel_buffer = io.BytesIO()
        wb = Workbook()
        # Supprimer la feuille par défaut
        wb.remove(wb.active)
        
        tables_created = 0
        
        with pdfplumber.open(uploaded_file) as pdf:
            progress_bar = st.progress(0)
            
            for i, page in enumerate(pdf.pages):
                # Extraction avec tolérance pour mieux grouper les mots
                # x_tolerance: espace horizontal pour considérer que c'est la même cellule
                # y_tolerance: espace vertical pour considérer que c'est la même ligne
                tables = page.extract_tables(x_tolerance=5, y_tolerance=5)
                
                for j, table in enumerate(tables):
                    if table:
                        cleaned_data = clean_table_data(table)
                        
                        if len(cleaned_data) > 1: # Au moins un header + 1 ligne
                            # Création DataFrame
                            df = pd.DataFrame(cleaned_data[1:], columns=cleaned_data[0])
                            
                            # Nom de la feuille
                            sheet_name = f"Page{i+1}_T{j+1}"[:31]
                            ws = wb.create_sheet(title=sheet_name)
                            
                            # Conversion DataFrame -> Liste pour openpyxl
                            for r_idx, row in enumerate(df.values, 2): # Commence à la ligne 2 (après header)
                                for c_idx, value in enumerate(row, 1):
                                    ws.cell(row=r_idx, column=c_idx, value=value)
                            
                            # Ajout des headers manuellement pour le style
                            for c_idx, col_name in enumerate(df.columns, 1):
                                ws.cell(row=1, column=c_idx, value=col_name)
                            
                            # Application du style
                            style_excel_sheet(ws, df)
                            tables_created += 1
                
                progress_bar.progress((i + 1) / len(pdf.pages))
            
            progress_bar.empty()
            
            if tables_created == 0:
                st.warning("Aucun tableau structuré détecté. Essayez un PDF avec des lignes de grille visibles.")
            else:
                st.success(f"{tables_created} tableaux extraits et formatés !")
                
                # Sauvegarde du workbook
                wb.save(excel_buffer)
                excel_bytes = excel_buffer.getvalue()
                
                st.download_button(
                    label="📥 Télécharger l'Excel Stylisé",
                    data=excel_bytes,
                    file_name="tableaux_structures.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                with st.expander("Aperçu des données brutes"):
                    st.write("Voici comment les données ont été interprétées (avant style):")
                    # Juste pour l'exemple, on affiche le premier tableau trouvé
                    # (Note: dans cette version on n'a pas gardé les DF en mémoire pour l'affichage pour économiser RAM)
                    st.info("Le téléchargement contient le résultat final formaté.")

    except Exception as e:
        st.error(f"Erreur : {e}")
        st.write("Vérifiez que le PDF n'est pas protégé par un mot de passe.")
