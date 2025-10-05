# frontend/pages/1_📁_Projets.py
import streamlit as st
from zoneinfo import ZoneInfo

from backend.services.database_service import DatabaseService
from backend.services.project_service import ProjectService

PARIS = ZoneInfo("Europe/Paris")
st.set_page_config(page_title="Projets", page_icon="📁", layout="wide")
st.session_state.setdefault("selected_project_id", None)

st.title("📁 Projets")
st.caption("Orchestrez plusieurs livrables avec des pipelines data configurés")

st.divider()

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
            st.switch_page("pages/_1a_📁_Projet_Detail.py")  # ← AJOUTER UNDERSCORE

    with colR:
        st.subheader("Mes projets")
        items = ps.list_projects()
        if not items:
            st.info("Aucun projet pour l'instant.")
        else:
            for p in items:
                with st.container(border=True):
                    st.markdown(f"**{p.get('name','(sans nom)')}**")
                    st.caption(f"ID : `{p.get('project_id')}`  ·  Mis à jour : {p.get('updated_at','')}")
                    
                    c1, c2 = st.columns(2)
                    if c1.button("Ouvrir", key=f"open_{p['project_id']}", use_container_width=True):
                        st.session_state["selected_project_id"] = p["project_id"]
                        st.switch_page("pages/_1a_📁_Projet_Detail.py")  # ← AJOUTER UNDERSCORE
                    
                    with c2.popover("Supprimer", use_container_width=True):
                        st.warning(
                            "Cette action retire le projet de la liste et déplace son fichier JSON dans "
                            "`assets/projects/00_Trash/`. Vous pourrez le restaurer plus tard.",
                            icon="🗑️",
                        )
                        cc1, cc2 = st.columns(2)
                        if cc1.button("Confirmer", key=f"confirm_del_{p['project_id']}", type="primary", use_container_width=True):
                            ps.soft_delete(p["project_id"])  # ← déplace vers assets/projects/00_Trash/
                            st.toast(f"Projet « {p.get('name','(sans nom)')} » déplacé dans la corbeille.", icon="✅")
                            st.rerun()
                        if cc2.button("Annuler", key=f"cancel_del_{p['project_id']}", use_container_width=True):
                            pass
