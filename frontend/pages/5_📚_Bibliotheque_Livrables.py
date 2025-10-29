# frontend/pages/4_📚_Bibliotheque_Livrables.py
import streamlit as st
from math import ceil
from backend.services.database_service import DatabaseService
from backend.services.project_service import ProjectService

st.set_page_config(page_title="Bibliothèque — Navigation", page_icon="📚", layout="wide")

# === STATE ===
ss = st.session_state
ss.setdefault("nav_client_query", "")
ss.setdefault("nav_project_query", "")
ss.setdefault("nav_clients_page", 1)
ss.setdefault("nav_projects_page", 1)
ss.setdefault("nav_view_mode", "grid")  # grid ou list

# === CONST ===
PAGE_SIZE_CLIENTS = 20
PAGE_SIZE_PROJECTS = 20

# === STYLES PERSONNALISÉS ===
st.markdown("""
<style>
    .client-card {
        padding: 1.5rem;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
        background: white;
        transition: all 0.2s;
        cursor: pointer;
        height: 100%;
    }
    .client-card:hover {
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        border-color: #4CAF50;
        transform: translateY(-2px);
    }
    .client-name {
        font-size: 1.2rem;
        font-weight: 600;
        color: #1f1f1f;
        margin-bottom: 0.5rem;
    }
    .project-count {
        color: #666;
        font-size: 0.9rem;
    }
    .search-box {
        margin-bottom: 2rem;
    }
    .stat-badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 12px;
        background: #e3f2fd;
        color: #1976d2;
        font-size: 0.85rem;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)

st.title("📚 Bibliothèque de Livrables")
st.caption("Parcourez vos projets et accédez à leurs livrables")

with DatabaseService.get_session() as db:
    ps = ProjectService(db)
    projects = ps.list_projects()

# ---------------------------------------
# SECTION 1: FILTRES ET RECHERCHE
# ---------------------------------------
with st.container():
    col1, col2, col3 = st.columns([3, 1, 1])
    
    with col1:
        ss.nav_client_query = st.text_input(
            "🔍 Rechercher un client",
            value=ss.nav_client_query,
            placeholder="Tapez le nom du client...",
            label_visibility="collapsed"
        )
    
    with col2:
        if st.button("🔄 Réinitialiser", use_container_width=True):
            ss.nav_client_query = ""
            ss.nav_clients_page = 1
            ss.nav_project_query = ""
            ss.nav_projects_page = 1
            if "selected_client_for_projects" in ss:
                del ss["selected_client_for_projects"]
            st.rerun()
    
    with col3:
        view_mode = st.segmented_control(
            "Affichage",
            options=["grid", "list"],
            default="grid",
            label_visibility="collapsed"
        )
        if view_mode:
            ss.nav_view_mode = view_mode

st.divider()

# ---------------------------------------
# SECTION 2: LISTE DES CLIENTS
# ---------------------------------------
# Extraire et filtrer les clients
clients_all = sorted({p.get("client_name", "") for p in projects if p.get("client_name")})
clients_filtered = [c for c in clients_all if ss.nav_client_query.strip().lower() in c.lower()]

# Pagination
total_clients = len(clients_filtered)
pages_clients = max(1, ceil(total_clients / PAGE_SIZE_CLIENTS))
ss.nav_clients_page = min(ss.nav_clients_page, pages_clients)
start = (ss.nav_clients_page - 1) * PAGE_SIZE_CLIENTS
end = start + PAGE_SIZE_CLIENTS
clients_page = clients_filtered[start:end]

# En-tête de section
col_head1, col_head2 = st.columns([3, 1])
with col_head1:
    st.subheader(f"Clients ({total_clients})")
with col_head2:
    if pages_clients > 1:
        st.caption(f"Page {ss.nav_clients_page} / {pages_clients}")

# Contrôles de pagination
if pages_clients > 1:
    pcol1, pcol2, pcol3 = st.columns([1, 3, 1])
    with pcol1:
        if st.button("⬅️ Préc", disabled=ss.nav_clients_page <= 1, use_container_width=True):
            ss.nav_clients_page -= 1
            st.rerun()
    with pcol2:
        # Sélecteur de page
        page_num = st.selectbox(
            "Page",
            options=list(range(1, pages_clients + 1)),
            index=ss.nav_clients_page - 1,
            label_visibility="collapsed"
        )
        if page_num != ss.nav_clients_page:
            ss.nav_clients_page = page_num
            st.rerun()
    with pcol3:
        if st.button("Suiv ➡️", disabled=ss.nav_clients_page >= pages_clients, use_container_width=True):
            ss.nav_clients_page += 1
            st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

# Affichage des clients
if not clients_page:
    st.info("Aucun client trouvé avec ces critères de recherche.")
else:
    if ss.nav_view_mode == "grid":
        # Vue en grille (3 colonnes)
        cols_per_row = 3
        for i in range(0, len(clients_page), cols_per_row):
            cols = st.columns(cols_per_row)
            for j, col in enumerate(cols):
                if i + j < len(clients_page):
                    client = clients_page[i + j]
                    projs = [p for p in projects if p.get("client_name") == client]
                    
                    with col:
                        with st.container(border=True):
                            st.markdown(f"### {client}")
                            st.markdown(f'<span class="stat-badge">{len(projs)} projet{"s" if len(projs) > 1 else ""}</span>', unsafe_allow_html=True)
                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button("📂 Voir les projets", key=f"see_{client}", use_container_width=True):
                                ss.selected_client_for_projects = client
                                ss.nav_projects_page = 1
                                st.rerun()
    else:
        # Vue en liste
        for client in clients_page:
            projs = [p for p in projects if p.get("client_name") == client]
            
            col1, col2, col3 = st.columns([5, 2, 2])
            with col1:
                st.markdown(f"**{client}**")
            with col2:
                st.caption(f"📊 {len(projs)} projet(s)")
            with col3:
                if st.button("📂 Ouvrir", key=f"see_list_{client}", use_container_width=True):
                    ss.selected_client_for_projects = client
                    ss.nav_projects_page = 1
                    st.rerun()
            st.divider()

# ---------------------------------------
# SECTION 3: PROJETS DU CLIENT SÉLECTIONNÉ
# ---------------------------------------
if ss.get("selected_client_for_projects"):
    st.markdown("---")
    
    # En-tête avec retour
    col_back, col_title = st.columns([1, 5])
    with col_back:
        if st.button("⬅️ Retour", use_container_width=True):
            del ss["selected_client_for_projects"]
            ss.nav_project_query = ""
            ss.nav_projects_page = 1
            st.rerun()
    with col_title:
        st.subheader(f"Projets : {ss.selected_client_for_projects}")
    
    # Filtrer les projets
    projs_all = [p for p in projects if p.get("client_name") == ss.selected_client_for_projects]
    
    # Recherche projet
    col_search, col_reset = st.columns([4, 1])
    with col_search:
        ss.nav_project_query = st.text_input(
            "🔍 Filtrer les projets",
            value=ss.nav_project_query,
            placeholder="Tapez le nom du projet...",
            label_visibility="collapsed"
        )
    with col_reset:
        if st.button("🔄 Reset", use_container_width=True, key="reset_projects"):
            ss.nav_project_query = ""
            ss.nav_projects_page = 1
            st.rerun()
    
    # Filtrage
    projs_filtered = [p for p in projs_all if ss.nav_project_query.strip().lower() in p.get("name", "").lower()]
    
    # Pagination projets
    total_projs = len(projs_filtered)
    pages_projs = max(1, ceil(total_projs / PAGE_SIZE_PROJECTS))
    ss.nav_projects_page = min(ss.nav_projects_page, pages_projs)
    pstart = (ss.nav_projects_page - 1) * PAGE_SIZE_PROJECTS
    pend = pstart + PAGE_SIZE_PROJECTS
    projs_page = projs_filtered[pstart:pend]
    
    # En-tête projets
    st.markdown(f"**{total_projs} projet(s)** trouvé(s)")
    
    # Contrôles pagination projets
    if pages_projs > 1:
        p1, p2, p3 = st.columns([1, 3, 1])
        with p1:
            if st.button("⬅️", disabled=ss.nav_projects_page <= 1, use_container_width=True, key="proj_prev"):
                ss.nav_projects_page -= 1
                st.rerun()
        with p2:
            proj_page_num = st.selectbox(
                "Page projets",
                options=list(range(1, pages_projs + 1)),
                index=ss.nav_projects_page - 1,
                label_visibility="collapsed",
                key="proj_page_select"
            )
            if proj_page_num != ss.nav_projects_page:
                ss.nav_projects_page = proj_page_num
                st.rerun()
        with p3:
            if st.button("➡️", disabled=ss.nav_projects_page >= pages_projs, use_container_width=True, key="proj_next"):
                ss.nav_projects_page += 1
                st.rerun()
    
    st.divider()
    
    # Affichage des projets
    if not projs_page:
        st.info("Aucun projet trouvé.")
    else:
        for p in projs_page:
            with st.container(border=True):
                col1, col2, col3 = st.columns([4, 2, 2])
                
                with col1:
                    st.markdown(f"### {p.get('name', '?')}")
                    st.caption(f"ID: `{p.get('project_id', '?')}`")
                
                with col2:
                    st.caption(f"📅 Mis à jour")
                    st.caption(p.get('updated_at', '—'))
                
                with col3:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("➡️ Accéder aux livrables", key=f"open_proj_{p['project_id']}", use_container_width=True, type="primary"):
                        ss.selected_project_id = p["project_id"]
                        ss.selected_project_name = p.get("name", "")
                        ss.selected_project_client = p.get("client_name", "")
                        st.switch_page("pages/_5a_📁_Livrables_Projet.py")