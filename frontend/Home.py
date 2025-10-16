# frontend/Home.py
import streamlit as st

st.set_page_config(page_title="KAIVAA", page_icon="🧩", layout="wide")

# Router : on enregistre toutes les pages, mais on cache le menu natif
nav = st.navigation(
    {
        "Main": [
            st.Page("pages/1_📁_Projets.py",               title="Projets",      icon="📁", default=True),
            st.Page("pages/2_📚_Bibliotheque.py",          title="Templates", icon="📚"),
            st.Page("pages/3d_🔀_Transformations.py",      title="Transformation",     icon="🔀"),
            st.Page("pages/3_🧱_Gabarits.py",              title="Gabarits",     icon="🧱"),
            st.Page("pages/4_🕘_Historique.py",            title="Historique",   icon="🕘"),
        ],
        # Pages secondaires “cachées” (ne s’affichent pas dans la nav)
        "Hidden": [
            st.Page("pages/_1a_🗂️_Hub_Projet.py",    title=None),
            st.Page("pages/_1b_💾_Data_Projet.py", title=None),
            st.Page("pages/_1c_⚙️_Config_Deliverable.py", title=None),
            st.Page("pages/_2a_📊_Detail_Livrable.py",  title=None),
            st.Page("pages/_2b_➕_Form_Template.py",     title=None),
            st.Page("pages/_2b3_📑_Tables_Template.py",     title=None),
            st.Page("pages/_2b4_🧾_Ajustement_Table.py",     title=None),
            st.Page("pages/_3a_🧱_Detail_Gabarit.py",    title=None),
            st.Page("pages//_3c_⚙️_Methodes_Gabarit.py", title=None),
            st.Page("pages//_3b1_🧱_Structure_Gabarit.py", title=None),
            st.Page("pages//_3b2_🔗_Enrichissements_Gabarit.py", title=None),
            st.Page("pages//_3b3_📁_Donnee_Par_Defaut.py", title=None),
            st.Page("pages//_3d1_🔀_Detail_Transformation.py", title=None),
            st.Page("pages//_3d2_🔀_Builder_Transformation.py", title=None),
        ],
    },
    position="hidden",  # masque la nav Streamlit par défaut
)

# Ta sidebar “maison” : n’affiche que les pages principales
with st.sidebar:
    st.page_link("pages/1_📁_Projets.py",               label="Projets",        icon="📁")
    st.page_link("pages/2_📚_Bibliotheque.py",          label="Templates",      icon="📚")
    st.page_link("pages/3d_🔀_Transformations.py",      label="Transformation", icon="🔀")
    st.page_link("pages/3_🧱_Gabarits.py",              label="Gabarits",       icon="🧱")
    st.page_link("pages/4_🕘_Historique.py",            label="Historique",     icon="🕘")

nav.run()

