# frontend/pages/4b_📁_Livrables_Projet.py
import streamlit as st
import os
from pathlib import Path
from datetime import datetime
from backend.services.database_service import DatabaseService
from backend.services.project_service import ProjectService
from backend.services.iteration_service import IterationService

st.set_page_config(page_title="📁 Livrables — Projet", page_icon="📁", layout="wide")

# === STYLES PERSONNALISÉS ===
st.markdown("""
<style>
    .deliverable-card {
        padding: 1.25rem;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
        background: #fafafa;
        margin-bottom: 0.75rem;
        transition: all 0.2s;
    }
    .deliverable-card:hover {
        background: #f5f5f5;
        border-color: #4CAF50;
    }
    .deliverable-card.selected {
        background: #e8f5e9;
        border: 2px solid #4CAF50;
    }
    .badge {
        display: inline-block;
        padding: 0.25rem 0.5rem;
        border-radius: 4px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-right: 0.5rem;
    }
    .badge-wip { background: #fff3cd; color: #856404; }
    .badge-final { background: #d1ecf1; color: #0c5460; }
    .badge-presented { background: #d4edda; color: #155724; }
    .badge-sent { background: #cce5ff; color: #004085; }
    
    .section-header {
        padding: 1rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 8px;
        margin-bottom: 1.5rem;
    }
    .consolidation-zone {
        background: #fff8e1;
        border: 2px dashed #ffa000;
        border-radius: 8px;
        padding: 1.5rem;
        margin: 1.5rem 0;
    }
    .order-indicator {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 28px;
        height: 28px;
        background: #4CAF50;
        color: white;
        border-radius: 50%;
        font-weight: 700;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

ss = st.session_state

# === GUARDS ===
if "selected_project_id" not in ss:
    st.warning("⚠️ Aucun projet sélectionné. Retournez à la page Bibliothèque.")
    if st.button("🔙 Retour à la bibliothèque"):
        st.switch_page("pages/5_📚_Bibliotheque_Livrables.py")
    st.stop()

project_id = ss.selected_project_id
project_name = ss.get("selected_project_name", "?")
client_name = ss.get("selected_project_client", "?")

# === STATE MANAGEMENT ===
ss.setdefault("sel_iters", [])
ss.setdefault("show_conso_modal", False)
ss.setdefault("show_state_modal", False)
ss.setdefault("selected_iter_for_state", None)
ss.setdefault("show_workdoc_actions_modal", False)
ss.setdefault("selected_workdoc_for_actions", None)
ss.setdefault("show_delete_iter_modal", False)
ss.setdefault("selected_iter_for_delete", None)
ss.setdefault("filter_template", "Tous")
ss.setdefault("sort_by", "date_desc")
ss.setdefault("show_all_raw_versions", "latest")

# === HEADER ===
col_back, col_title = st.columns([1, 5])
with col_back:
    if st.button("⬅️ Retour", use_container_width=True):
        st.switch_page("pages/5_📚_Bibliotheque_Livrables.py")

with col_title:
    st.markdown(f"<div class='section-header'><h1>📁 {project_name}</h1><p style='margin:0; opacity:0.9;'>Client: {client_name} | ID: {project_id}</p></div>", unsafe_allow_html=True)

# === LOAD DATA ===
with DatabaseService.get_session() as db:
    ps = ProjectService(db)
    iserv = IterationService(db)
    
    try:
        _ = ps.load_project(project_id)
    except Exception:
        st.error("❌ Projet introuvable.")
        st.stop()

iterations = iserv.list_iterations(project_id)

# Séparation bruts vs workdocs
unit_iterations = [it for it in iterations if "template_name" in it]
workdocs = [it for it in iterations if "label" in it]

# Grouper bruts par template
templates = {}
for it in unit_iterations:
    tmpl = it.get("template_name", "?")
    templates.setdefault(tmpl, []).append(it)

# === TABS NAVIGATION ===
tab1, tab2, tab3 = st.tabs(["🎨 Livrables Bruts", "🧩 Consolidation", "🧱 Documents de Travail"])

# ============================================================================
# TAB 1: LIVRABLES BRUTS
# ============================================================================
with tab1:
    st.subheader("🎨 Livrables Bruts")
    st.caption("Livrables générés individuellement, groupés par template")
    
    # Filtres et tri
    col_filter, col_sort, col_version = st.columns([2, 2, 2])
    with col_filter:
        template_options = ["Tous"] + list(templates.keys())
        ss.filter_template = st.selectbox(
            "Filtrer par template",
            options=template_options,
            index=template_options.index(ss.filter_template) if ss.filter_template in template_options else 0
        )
    
    with col_sort:
        ss.sort_by = st.selectbox(
            "Trier par",
            options=[
                ("date_desc", "📅 Plus récent d'abord"),
                ("date_asc", "📅 Plus ancien d'abord"),
                ("template", "📂 Par template")
            ],
            format_func=lambda x: x[1],
            index=0
        )[0]
    
    with col_version:
        ss.show_all_raw_versions = st.selectbox(
            "Affichage",
            options=[("latest", "📌 Dernière version"), ("all", "📜 Toutes")],
            format_func=lambda x: x[1],
            index=0 if ss.show_all_raw_versions == "latest" else 1
        )[0]
    
    st.divider()
    
    # Application des filtres
    filtered_templates = templates if ss.filter_template == "Tous" else {ss.filter_template: templates.get(ss.filter_template, [])}
    
    if not filtered_templates or all(len(items) == 0 for items in filtered_templates.values()):
        st.info("📭 Aucun livrable brut généré pour ce projet.")
    else:
        for tmpl, items in filtered_templates.items():
            if not items:
                continue
            
            # Tri
            if ss.sort_by == "date_desc":
                items = sorted(items, key=lambda x: x.get("generated_at", ""), reverse=True)
            elif ss.sort_by == "date_asc":
                items = sorted(items, key=lambda x: x.get("generated_at", ""))
            
            # Filtrer dernière version si demandé
            if ss.show_all_raw_versions == "latest" and items:
                items = [items[0]]
            
            with st.expander(f"📦 **{tmpl}** ({len(items)} version{' s' if len(items) > 1 else ''})", expanded=True):
                for idx, it in enumerate(items):
                    iter_id = it["id"]
                    is_selected = iter_id in ss.sel_iters
                    order_num = ss.sel_iters.index(iter_id) + 1 if is_selected else None
                    
                    with st.container(border=True):
                        # Ligne principale
                        col_sel, col_info, col_actions = st.columns([0.5, 2.5, 2.5])
                        
                        with col_sel:
                            # Checkbox avec indicateur d'ordre
                            if order_num:
                                st.markdown(f"<div class='order-indicator'>{order_num}</div>", unsafe_allow_html=True)
                            checked = st.checkbox(
                                "Sél.",
                                key=f"chk_{iter_id}",
                                value=is_selected,
                                label_visibility="collapsed"
                            )
                            if checked and not is_selected:
                                ss.sel_iters.append(iter_id)
                            elif not checked and is_selected:
                                ss.sel_iters.remove(iter_id)
                        
                        with col_info:
                            gen_at = it.get("generated_at", "—")
                            try:
                                dt = datetime.fromisoformat(gen_at)
                                gen_at_formatted = dt.strftime("%d/%m/%Y %H:%M")
                            except:
                                gen_at_formatted = gen_at
                            
                            st.markdown(f"**Généré le:** {gen_at_formatted}")
                            st.caption(it.get("ppt_path", "Chemin non disponible"))
                        
                        with col_actions:
                            acol1, acol2, acol3, acol4 = st.columns(4)
                            
                            with acol1:
                                ppt_path = it.get("ppt_path")
                                if ppt_path and st.button("📂 PPT", key=f"open_ppt_{iter_id}", use_container_width=True):
                                    try:
                                        os.startfile(ppt_path)
                                    except Exception as e:
                                        st.error(f"Erreur: {e}")
                            
                            with acol2:
                                xls_path = it.get("excel_path")
                                if xls_path and st.button("📊 XLS", key=f"open_xls_{iter_id}", use_container_width=True):
                                    try:
                                        os.startfile(xls_path)
                                    except Exception as e:
                                        st.error(f"Erreur: {e}")
                            
                            with acol3:
                                if st.button("⚙️", key=f"state_{iter_id}", use_container_width=True, help="Changer l'état"):
                                    ss.selected_iter_for_state = iter_id
                                    ss.show_state_modal = True
                                    st.rerun()
                            
                            with acol4:
                                if st.button("🗑️", key=f"del_{iter_id}", use_container_width=True, help="Supprimer"):
                                    ss.selected_iter_for_delete = iter_id
                                    ss.show_delete_iter_modal = True
                                    st.rerun()

# ============================================================================
# TAB 2: CONSOLIDATION
# ============================================================================
with tab2:
    st.subheader("🧩 Zone de Consolidation")
    st.caption("Réorganisez l'ordre des livrables")
    
    sel_count = len(ss.sel_iters)
    
    st.markdown('<div class="consolidation-zone">', unsafe_allow_html=True)
    
    if sel_count == 0:
        st.info("🔍 Sélectionnez des livrables dans l'onglet 'Livrables Bruts'")
    else:
        st.success(f"✅ **{sel_count} livrable(s) sélectionné(s)**")
        st.divider()
        
        for i, iter_id in enumerate(ss.sel_iters):
            iter_data = next((it for it in iterations if it["id"] == iter_id), None)
            if not iter_data:
                continue
            
            col_num, col_info, col_btns = st.columns([0.5, 4, 1.5])
            
            with col_num:
                st.markdown(f"<div class='order-indicator'>{i+1}</div>", unsafe_allow_html=True)
            
            with col_info:
                st.markdown(f"**{iter_data.get('template_name', '?')}**")
                st.caption(f"📅 {iter_data.get('generated_at', '?')}")
            
            with col_btns:
                b1, b2, b3 = st.columns(3)
                with b1:
                    if st.button("⬆️", key=f"up_{iter_id}", disabled=i==0, use_container_width=True):
                        ss.sel_iters[i], ss.sel_iters[i-1] = ss.sel_iters[i-1], ss.sel_iters[i]
                        st.rerun()
                with b2:
                    if st.button("⬇️", key=f"down_{iter_id}", disabled=i==len(ss.sel_iters)-1, use_container_width=True):
                        ss.sel_iters[i], ss.sel_iters[i+1] = ss.sel_iters[i+1], ss.sel_iters[i]
                        st.rerun()
                with b3:
                    if st.button("❌", key=f"rm_{iter_id}", use_container_width=True):
                        ss.sel_iters.remove(iter_id)
                        st.rerun()
        
        st.divider()
        
        col_reset, col_create = st.columns(2)
        with col_reset:
            if st.button("🔄 Réinitialiser", use_container_width=True):
                ss.sel_iters = []
                st.rerun()
        
        with col_create:
            if st.button("🚀 Créer document", type="primary", use_container_width=True):
                ss.show_conso_modal = True
                st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)

# ============================================================================
# TAB 3: DOCUMENTS DE TRAVAIL
# ============================================================================
with tab3:
    st.subheader("🧱 Documents de Travail Consolidés")
    st.caption("Documents créés à partir de plusieurs livrables")
    
    # Option d'affichage
    col_opt1, col_opt2 = st.columns([3, 1])
    with col_opt1:
        show_all_versions = st.checkbox("📜 Afficher toutes les versions", value=False)
    
    if not workdocs:
        st.info("📭 Aucun document consolidé pour l'instant.")
    else:
        # Grouper par label si on affiche seulement la dernière version
        if not show_all_versions:
            # Garder seulement la dernière version de chaque document
            latest_docs = {}
            for wd in workdocs:
                label = wd.get("label", "")
                if label not in latest_docs or wd.get("version", 0) > latest_docs[label].get("version", 0):
                    latest_docs[label] = wd
            display_docs = list(latest_docs.values())
        else:
            display_docs = workdocs
        
        for wd in display_docs:
            state = wd.get("state", "WIP")
            version = wd.get("version", 1)
            
            # Badge de statut
            badge_class = f"badge-{state.lower()}"
            
            with st.container(border=True):
                col1, col2, col3 = st.columns([3, 2, 2])
                
                with col1:
                    st.markdown(f"### {wd.get('label', '(sans nom)')}")
                    st.markdown(f'<span class="badge {badge_class}">{state}</span><span class="badge">V{version}</span>', unsafe_allow_html=True)
                
                with col2:
                    created = wd.get("created_at", "—")
                    try:
                        dt = datetime.fromisoformat(created)
                        created_fmt = dt.strftime("%d/%m/%Y %H:%M")
                    except:
                        created_fmt = created
                    st.caption(f"📅 Créé le {created_fmt}")
                    st.caption(f"📁 {wd.get('ppt_path', 'N/A')}")
                
                with col3:
                    # Boutons d'action
                    acol1, acol2 = st.columns(2)
                    with acol1:
                        if st.button("📂 Ouvrir", key=f"open_wd_{wd['id']}", use_container_width=True):
                            try:
                                os.startfile(wd["ppt_path"])
                            except Exception as e:
                                st.error(f"Erreur: {e}")
                    
                    with acol2:
                        if st.button("⚙️ Plus", key=f"more_wd_{wd['id']}", use_container_width=True):
                            ss.selected_workdoc_for_actions = wd["id"]
                            ss.show_workdoc_actions_modal = True
                            st.rerun()

# ============================================================================
# MODALES
# ============================================================================

# Modale de consolidation
if ss.show_conso_modal:
    @st.dialog("🚀 Créer un Document Consolidé", width="large")
    def _conso_modal():
        st.markdown(f"**{len(ss.sel_iters)} livrables** vont être consolidés dans l'ordre de sélection.")
        
        wd_name = st.text_input(
            "Nom du document de travail",
            placeholder="Ex : Atelier 1 - Octobre 2025",
            help="Choisissez un nom descriptif"
        )
        
        with st.expander("📋 Ordre de consolidation"):
            for i, iter_id in enumerate(ss.sel_iters, 1):
                iter_data = next((it for it in iterations if it["id"] == iter_id), None)
                if iter_data:
                    st.markdown(f"{i}. {iter_data.get('template_name', '?')}")
        
        col_cancel, col_confirm = st.columns(2)
        with col_cancel:
            if st.button("❌ Annuler", use_container_width=True):
                ss.show_conso_modal = False
                st.rerun()
        
        with col_confirm:
            if st.button("✅ Consolider", type="primary", disabled=not wd_name.strip(), use_container_width=True):
                with st.spinner("🔄 Consolidation en cours..."):
                    try:
                        with DatabaseService.get_session() as db:
                            meta = IterationService(db).consolidate(
                                project_id=project_id,
                                iteration_ids=ss.sel_iters,
                                workdoc_name=wd_name.strip()
                            )
                        ss.sel_iters = []
                        ss.show_conso_modal = False
                        st.success(f"✅ Document créé : {meta['ppt_path']}")
                        st.balloons()
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Erreur lors de la consolidation: {e}")
    
    _conso_modal()

# Modale de changement d'état
if ss.show_state_modal and ss.selected_iter_for_state:
    @st.dialog("⚙️ Changer l'État", width="medium")
    def _state_modal():
        st.markdown("Sélectionnez le nouvel état du document:")
        
        states = {
            "WIP": "🟡 En cours (Work In Progress)",
            "FINAL": "🔵 Finalisé",
            "PRESENTED": "🟢 Présenté",
            "SENT": "🟣 Envoyé"
        }
        
        new_state = st.radio(
            "État",
            options=list(states.keys()),
            format_func=lambda x: states[x],
            label_visibility="collapsed"
        )
        
        col_cancel, col_save = st.columns(2)
        with col_cancel:
            if st.button("❌ Annuler", use_container_width=True):
                ss.show_state_modal = False
                ss.selected_iter_for_state = None
                st.rerun()
        
        with col_save:
            if st.button("✅ Enregistrer", type="primary", use_container_width=True):
                try:
                    with DatabaseService.get_session() as db:
                        IterationService(db).update_state(
                            project_id=project_id,
                            iteration_id=ss.selected_iter_for_state,
                            new_state=new_state
                        )
                    ss.show_state_modal = False
                    ss.selected_iter_for_state = None
                    st.success(f"✅ État mis à jour : {states[new_state]}")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur: {e}")
    
    _state_modal()

# Modale d'actions pour workdocs
if ss.get("show_workdoc_actions_modal") and ss.get("selected_workdoc_for_actions"):
    @st.dialog("⚙️ Actions sur le Document", width="medium")
    def _workdoc_actions_modal():
        # Recharger les workdocs
        with DatabaseService.get_session() as db:
            itersv = IterationService(db)
            all_iters = itersv.list_iterations(project_id)
        wd_list = [it for it in all_iters if "label" in it]
        
        workdoc = next((wd for wd in wd_list if wd["id"] == ss.selected_workdoc_for_actions), None)
        if not workdoc:
            ss.show_workdoc_actions_modal = False
            ss.selected_workdoc_for_actions = None
            st.rerun()
            return
        
        st.markdown(f"**{workdoc.get('label', '?')}** - V{workdoc.get('version', 1)}")
        st.divider()
        
        # Action 1: Nouvelle version
        with st.expander("📋 Créer une nouvelle version", expanded=True):
            st.caption("Crée une copie du fichier actuel avec un numéro de version incrémenté")
            if st.button("🆕 Nouvelle Version", use_container_width=True, type="primary"):
                try:
                    with DatabaseService.get_session() as db:
                        new_meta = IterationService(db).create_new_version(
                            project_id=project_id,
                            workdoc_id=ss.selected_workdoc_for_actions
                        )
                    ss.show_workdoc_actions_modal = False
                    ss.selected_workdoc_for_actions = None
                    st.success(f"✅ Version {new_meta['version']} créée!")
                    st.balloons()
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur: {e}")
        
        # Action 2: Changer l'état
        with st.expander("🏷️ Changer l'état"):
            states = {
                "WIP": "🟡 En cours",
                "FINAL": "🔵 Finalisé",
                "PRESENTED": "🟢 Présenté",
                "SENT": "🟣 Envoyé"
            }
            current_state = workdoc.get("state", "WIP")
            new_state = st.radio(
                "Nouvel état",
                options=list(states.keys()),
                format_func=lambda x: states[x],
                index=list(states.keys()).index(current_state)
            )
            if st.button("💾 Enregistrer l'état", use_container_width=True):
                try:
                    with DatabaseService.get_session() as db:
                        IterationService(db).update_state(
                            project_id=project_id,
                            iteration_id=ss.selected_workdoc_for_actions,
                            new_state=new_state
                        )
                    ss.show_workdoc_actions_modal = False
                    ss.selected_workdoc_for_actions = None
                    st.success(f"✅ État mis à jour : {states[new_state]}")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur: {e}")
        
        # Action 3: Supprimer
        with st.expander("🗑️ Supprimer le document"):
            st.warning("⚠️ Cette action déplacera le document dans la corbeille")
            confirm_name = st.text_input(
                "Tapez le nom du document pour confirmer",
                placeholder=workdoc.get('label', '')
            )
            if st.button("🗑️ Supprimer", use_container_width=True, type="secondary"):
                if confirm_name == workdoc.get('label', ''):
                    try:
                        with DatabaseService.get_session() as db:
                            IterationService(db).delete_workdoc(
                                project_id=project_id,
                                workdoc_id=ss.selected_workdoc_for_actions
                            )
                        ss.show_workdoc_actions_modal = False
                        ss.selected_workdoc_for_actions = None
                        st.success("✅ Document supprimé (déplacé dans la corbeille)")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Erreur: {e}")
                else:
                    st.error("❌ Le nom ne correspond pas")
        
        st.divider()
        if st.button("❌ Fermer", use_container_width=True):
            ss.show_workdoc_actions_modal = False
            ss.selected_workdoc_for_actions = None
            st.rerun()
    
    _workdoc_actions_modal()

# Modale suppression livrable brut
if ss.get("show_delete_iter_modal") and ss.get("selected_iter_for_delete"):
    @st.dialog("🗑️ Supprimer Livrable")
    def _delete_iter():
        iter_data = next((it for it in unit_iterations if it["id"] == ss.selected_iter_for_delete), None)
        if not iter_data:
            ss.show_delete_iter_modal = False
            ss.selected_iter_for_delete = None
            st.rerun()
            return
        st.warning("⚠️ Déplacement vers corbeille")
        st.markdown(f"**{iter_data.get('template_name', '?')}**")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("❌ Annuler", use_container_width=True):
                ss.show_delete_iter_modal = False
                ss.selected_iter_for_delete = None
                st.rerun()
        with c2:
            if st.button("🗑️ Confirmer", use_container_width=True, type="secondary"):
                try:
                    with DatabaseService.get_session() as db:
                        IterationService(db).delete_iteration(project_id, ss.selected_iter_for_delete)
                    if ss.selected_iter_for_delete in ss.sel_iters:
                        ss.sel_iters.remove(ss.selected_iter_for_delete)
                    ss.show_delete_iter_modal = False
                    ss.selected_iter_for_delete = None
                    st.success("✅ Supprimé")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {e}")
    _delete_iter()