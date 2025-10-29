# frontend/Home.py
import streamlit as st

st.set_page_config(page_title="KAIVAA", page_icon="🧩", layout="wide")

# Router : on enregistre toutes les pages, mais on cache le menu natif
nav = st.navigation(
    {
        "Main": [
            st.Page("pages/1_🧱_Gabarits.py",              title="Gabarits",       icon="🧱", default=True),
            st.Page("pages/2_🔀_Transformations.py",       title="Transformation", icon="🔀"),
            st.Page("pages/3_📚_Bibliotheque.py",          title="Templates",      icon="📝"),
            st.Page("pages/4_📁_Projets.py",               title="Projets",        icon="📁"),
            st.Page("pages/5_📚_Bibliotheque_Livrables.py",title="Livrables",      icon="📦"),
        ],
        # Pages secondaires "cachées" (ne s'affichent pas dans la nav)
        "Hidden": [
            # 1. Gabarits (Construction)
            st.Page("pages/_1a_🧱_Detail_Gabarit.py",           title=None),
            st.Page("pages/_1b_🧱_Structure_Gabarit.py",        title=None),
            st.Page("pages/_1c_🔗_Enrichissements_Gabarit.py",  title=None),
            st.Page("pages/_1d_📁_Donnee_Par_Defaut.py",        title=None),
            st.Page("pages/_1e_⚙️_Methodes_Gabarit.py",         title=None),
            
            # 2. Transformations
            st.Page("pages/_2a_🔀_Detail_Transformation.py",    title=None),
            st.Page("pages/_2b_🔀_Builder_Transformation.py",   title=None),
            
            # 3. Templates/Livrables (Développement)
            st.Page("pages/_3a_📊_Detail_Livrable.py",          title=None),
            st.Page("pages/_3b_➕_Form_Template.py",            title=None),
            st.Page("pages/_3c_📑_Tables_Template.py",          title=None),
            st.Page("pages/_3d_🧾_Ajustement_Table.py",         title=None),
            
            # 4. Projets (Paramétrage)
            st.Page("pages/_4a_🗂️_Hub_Projet.py",              title=None),
            st.Page("pages/_4b_💾_Data_Projet.py",              title=None),
            st.Page("pages/_4c_⚙️_Config_Deliverable.py",       title=None),
            
            # 5. Livrables (Finalisation)
            st.Page("pages/_5a_📁_Livrables_Projet.py",         title=None),
        ],
    },
    position="hidden",  # masque la nav Streamlit par défaut
)

# Ta sidebar "maison" : n'affiche que les pages principales
with st.sidebar:
    st.page_link("pages/1_🧱_Gabarits.py",              label="Gabarits",       icon="🧱")
    st.page_link("pages/2_🔀_Transformations.py",       label="Transformation", icon="🔀")
    st.page_link("pages/3_📚_Bibliotheque.py",          label="Templates",      icon="📝")
    st.page_link("pages/4_📁_Projets.py",               label="Projets",        icon="📁")
    st.page_link("pages/5_📚_Bibliotheque_Livrables.py",label="Livrables",      icon="📦")

nav.run()