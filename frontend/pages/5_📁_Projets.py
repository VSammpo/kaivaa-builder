import streamlit as st
from zoneinfo import ZoneInfo
from loguru import logger

from backend.services.database_service import DatabaseService
from backend.services.project_service import ProjectService
from frontend.utils.ui_helpers import page_header

PARIS = ZoneInfo("Europe/Paris")
st.set_page_config(page_title="Projets", page_icon="📁", layout="wide")
st.session_state.setdefault("selected_project_id", None)

def _goto(page_path: str):
    try:
        st.switch_page(page_path)
    except Exception:
        st.info("Utilise le menu de gauche si la redirection ne fonctionne pas.")

page_header("Projets", icon=None, crumbs=[("Catalogue", "1_📚_Bibliotheque"), ("Projets", "")])

DatabaseService.initialize()
with DatabaseService.get_session() as db:
    ps = ProjectService(db)

    colL, colR = st.columns([1, 2], gap="large")

    with colL:
        st.subheader("Créer un projet")
        name = st.text_input("Nom du projet")
        desc = st.text_area("Description", height=80)
        st.caption("Un ID stable sera généré automatiquement.")
        if st.button("Créer", type="primary", use_container_width=True, disabled=not name.strip()):
            proj = ps.create_project(name=name.strip(), description=desc.strip())
            st.session_state["selected_project_id"] = proj["project_id"]
            st.success(f"Projet créé : `{proj['project_id']}`")
            _goto("pages/5a_📁_Projet_Detail.py")

    with colR:
        st.subheader("Mes projets")
        items = ps.list_projects()
        if not items:
            st.info("Aucun projet pour l’instant.")
        else:
            for p in items:
                with st.container(border=True):
                    st.markdown(f"**{p.get('name','(sans nom)')}**  \nID : `{p.get('project_id')}`  \nMis à jour : {p.get('updated_at','')}")
                    c1, c2 = st.columns(2)
                    if c1.button("Ouvrir", key=f"open_{p['project_id']}", use_container_width=True):
                        st.session_state["selected_project_id"] = p["project_id"]
                        _goto("pages/5a_📁_Projet_Detail.py")
                    c2.button("Supprimer (à venir)", key=f"del_{p['project_id']}", disabled=True, use_container_width=True)
