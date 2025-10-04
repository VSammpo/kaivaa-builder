import streamlit as st
import pandas as pd
from loguru import logger

from backend.services.database_service import DatabaseService
from backend.services.project_service import ProjectService
from frontend.utils.ui_helpers import page_header
# Bandeau d'actions
st.markdown(
    """
    <div style="display:flex; gap:.5rem; margin:.5rem 0 1rem 0;">
      <a href="?page=5_📁_Projets" style="text-decoration:none;">
        <button>⟵ Retour Projets</button>
      </a>
      <a href="?page=5a_📁_Projet_Detail" style="text-decoration:none;">
        <button>⟵ Retour Détail projet</button>
      </a>
    </div>
    """,
    unsafe_allow_html=True
)

st.set_page_config(page_title="Pipeline Gabarit", page_icon="🧪", layout="wide")
st.session_state.setdefault("selected_project_id", None)
st.session_state.setdefault("selected_pipeline_gab", None)

pid = st.session_state.get("selected_project_id")
gab = st.session_state.get("selected_pipeline_gab")

page_header("Pipeline gabarit", icon=None, crumbs=[
    ("Projets", "5_📁_Projets"),
    ("Détail du projet", "5a_📁_Projet_Detail"),
    ("Pipeline gabarit", "")
])

if not pid or not gab:
    st.warning("Projet ou gabarit non défini. Ouvre d’abord « Détail du projet ».") 
    st.stop()

gname, gver = gab
st.caption(f"Projet : `{pid}` · Gabarit : **{gname}** · Version `{gver}`")

DatabaseService.initialize()
with DatabaseService.get_session() as db:
    ps = ProjectService(db)

    pipe = ps.get_pipeline(pid, gname, gver) or {}
    src = pipe.get("source") or {}
    code = pipe.get("python") or ""

    st.subheader("Source CSV")
    c1, c2, c3 = st.columns([3, 1, 1])
    path = c1.text_input("Chemin CSV", value=src.get("path", ""), placeholder="C:/data/source.csv")
    sep = c2.text_input("Séparateur", value=src.get("sep", ";"))
    enc = c3.text_input("Encodage", value=src.get("encoding", "utf-8-sig"))

    st.subheader("Transformation Python (df -> df)")
    code_new = st.text_area("Code pandas", value=code, height=200,
                            placeholder="# df = df.rename(columns={'A':'B'})\n# df = df[df['col'] > 0]\n")

    cL, cR = st.columns(2)
    if cL.button("💾 Enregistrer", type="primary", use_container_width=True):
        ps.set_pipeline(
            project_id=pid,
            gabarit_name=gname,
            gabarit_version=gver,
            source={"type": "csv", "path": path, "sep": sep, "encoding": enc},
            python_code=code_new
        )
        st.success("Pipeline enregistré.")

    if cR.button("🔍 Preview (1000)", use_container_width=True):
        try:
            df, profile = ps.preview(pid, gname, gver, head=1000)
            st.caption("Aperçu des 1000 premières lignes")
            st.dataframe(df, use_container_width=True)
            st.caption("Profiling simple")
            st.json(profile)
        except Exception as e:
            st.error(f"Erreur preview : {e}")

    st.divider()
    if st.button("✅ Valider (non bloquant)", use_container_width=True):
        try:
            res = ps.validate(pid, gname, gver)
            st.success("Validation effectuée.")
            st.json(res)
        except Exception as e:
            st.error(f"Erreur validation : {e}")

# Barre flottante
st.markdown(
    """
    <div style="
      position:fixed; bottom:12px; left:12px; 
      display:flex; gap:.5rem; 
      background:rgba(255,255,255,.85);
      border:1px solid #e5e5e5; 
      padding:.5rem .75rem; border-radius:.75rem;">
      <a href="?page=5_📁_Projets" style="text-decoration:none;">
        <button>⟵ Projets</button>
      </a>
      <a href="?page=5a_📁_Projet_Detail" style="text-decoration:none;">
        <button>⟵ Détail projet</button>
      </a>
    </div>
    """,
    unsafe_allow_html=True
)
