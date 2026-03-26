import streamlit as st
import pdfplumber
import pandas as pd
import io

# Configuration de la page
st.set_page_config(page_title="PDF vers Excel", page_icon="📊")

st.title("📄 Convertisseur PDF (Tableaux) vers Excel")
st.markdown("""
Cette application extrait les tableaux d'un fichier PDF et les place 
dans des feuilles séparées d'un fichier Excel.
""")

# Widget d'upload de fichier
uploaded_file = st.file_uploader("Choisissez un fichier PDF", type="pdf")

if uploaded_file is not None:
    st.info("Fichier chargé avec succès. Traitement en cours...")
    
    try:
        # Buffer pour le fichier Excel en mémoire
        excel_buffer = io.BytesIO()
        
        # Ouverture du PDF
        with pdfplumber.open(uploaded_file) as pdf:
            all_tables = []
            sheet_names = []
            
            progress_bar = st.progress(0)
            
            # Extraction des tableaux page par page
            for i, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                
                for j, table in enumerate(tables):
                    if table:
                        # Création d'un DataFrame pandas
                        df = pd.DataFrame(table)
                        
                        # Nettoyage basique (remplacer les None par vide)
                        df = df.fillna("")
                        
                        all_tables.append(df)
                        
                        # Nom de la feuille (Max 31 caractères pour Excel)
                        sheet_name = f"Page{i+1}_Tab{j+1}"[:31]
                        sheet_names.append(sheet_name)
                
                # Mise à jour de la barre de progression
                progress_bar.progress((i + 1) / len(pdf.pages))
            
            progress_bar.empty()
            
            if len(all_tables) == 0:
                st.warning("Aucun tableau n'a été détecté dans ce PDF.")
            else:
                st.success(f"{len(all_tables)} tableaux trouvés !")
                
                # Création du fichier Excel
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    for df, name in zip(all_tables, sheet_names):
                        df.to_excel(writer, sheet_name=name, index=False)
                
                # Préparation du téléchargement
                excel_bytes = excel_buffer.getvalue()
                
                st.download_button(
                    label="📥 Télécharger le fichier Excel",
                    data=excel_bytes,
                    file_name="tableaux_extraits.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Aperçu des données (optionnel)
                with st.expander("Voir un aperçu des tableaux"):
                    for idx, df in enumerate(all_tables):
                        st.write(f"**Feuille : {sheet_names[idx]}**")
                        st.dataframe(df.head())

    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")
        st.write("Conseil : Assurez-vous que le PDF contient des tableaux sélectionnables (pas des images).")
