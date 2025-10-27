# frontend/pages/2_📚_Bibliotheque.py
# Bibliothèque de templates — cartes sans image + actions intégrées (dans la carte)

import streamlit as st
from pathlib import Path
import sys
import hashlib
import re

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService

st.set_page_config(page_title="Bibliothèque", page_icon="📚", layout="wide")

# =============== CSS (léger) ===============
st.markdown("""
<style>
:root {
  --card-radius: 14px;
  --card-shadow: 0 6px 24px rgba(0,0,0,.06), 0 1px 3px rgba(0,0,0,.05);
  --card-shadow-hover: 0 10px 30px rgba(0,0,0,.10), 0 2px 6px rgba(0,0,0,.06);
  --border: 1px solid rgba(0,0,0,.08);
}

/* On stylise le CONTENEUR du formulaire comme une carte */
div[data-testid="stForm"] {
  border: var(--border);
  border-radius: var(--card-radius);
  background: #fff;
  overflow: hidden;
  box-shadow: var(--card-shadow);
  transition: transform .12s ease-out, box-shadow .12s ease-out;
  padding: 0;                 /* on gère nos sections */
  margin-bottom: 16px;
}
div[data-testid="stForm"]:hover {
  transform: translateY(-2px);
  box-shadow: var(--card-shadow-hover);
}

/* Sections internes */
.ka-card__header {
  padding: 14px 16px;
  color: #fff;
  display: flex; align-items: center; gap: 12px;
}
.ka-avatar {
  width: 44px; height: 44px; flex: 0 0 44px;
  border-radius: 999px;
  background: rgba(255,255,255,.22);
  backdrop-filter: saturate(120%) blur(1px);
  display: flex; align-items: center; justify-content: center;
  font-weight: 800; letter-spacing: .5px;
}
.ka-meta { line-height: 1.2; }
.ka-title { margin: 0; font-size: 1.05rem; font-weight: 700; letter-spacing: .2px; }
.ka-sub { opacity: .9; font-size: .85rem; }

.ka-pill {
  margin-left: auto; padding: 4px 10px; border-radius: 999px;
  background: rgba(255,255,255,.2); color: #fff; font-weight: 600; font-size: .8rem;
  border: 1px solid rgba(255,255,255,.22);
}

.ka-card__body { padding: 12px 16px 8px; color: #111827; }
.ka-desc { margin: 0; opacity: .82; font-size: .92rem; }

.ka-card__actions {
  padding: 12px 0 16px 0;  /* Espace en haut et en bas, MAIS PAS sur les côtés */
  border-top: 1px solid rgba(0,0,0,.06);
}

.ka-card__actions .stButton>button {
  width: 100%;
  margin: 0;
}
</style>
""", unsafe_allow_html=True)

# =============== Helpers ===============
def _initials(name: str) -> str:
    if not name:
        return "T"
    parts = [p for p in re.split(r"\s+", name.strip()) if p]
    if len(parts) >= 2:
        return (parts[0][0] + parts[1][0]).upper()
    return name[:2].upper()

def _gradient_from_text(text: str) -> tuple[str, str]:
    base = int(hashlib.sha1((text or "kaivaa").encode("utf-8")).hexdigest(), 16)
    hue = base % 360
    c1 = f"hsl({hue}, 72%, 44%)"
    c2 = f"hsl({(hue + 18) % 360}, 78%, 36%)"
    return c1, c2

def _compact_desc(text: str, max_len: int = 140) -> str:
    if not text:
        return "Aucune description."
    t = text.strip()
    return (t[: max_len - 1] + "…") if len(t) > max_len else t

# =============== Flash éventuel ===============
if msg := st.session_state.pop("_flash_success", None):
    st.success(msg); st.toast(msg)

# =============== En-tête & filtres ===============
st.title("📚 Bibliothèque de Templates")

# Barre de recherche (nom)
search = st.text_input("🔍 Rechercher", placeholder="Nom du template…")

# Préparer les options de filtres (scan rapide)
with DatabaseService.get_session() as db:
    svc = TemplateService(db)
    all_templates = svc.list_templates(active_only=True)
    # Construire les options pour gabarits / thématiques / familles
    all_gabarits, all_themes, all_families = set(), set(), set()
    for t in all_templates:
        facets = svc.extract_template_facets(t.id)
        all_gabarits |= facets["gabarits"]
        all_themes   |= facets["themes"]
        all_families |= facets["families"]
    gabarit_options = sorted(all_gabarits, key=str.lower)
    theme_options   = sorted(all_themes, key=str.lower)
    family_options  = sorted(all_families, key=str.lower)

# Filtres avancés
fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 1])
with fc1:
    selected_gabarits = st.multiselect("Sources (gabarits)", options=gabarit_options, placeholder="Sélectionner…")
with fc2:
    selected_themes = st.multiselect("Thématiques", options=theme_options, placeholder="Sélectionner…")
with fc3:
    selected_families = st.multiselect("Familles", options=family_options, placeholder="Sélectionner…")
with fc4:
    strict_sources = st.toggle("Recherche stricte", value=False, help="ON : n’affiche que les templates dont les sources sont exclusivement dans la sélection")

st.divider()


colr1, _, _ = st.columns([1,1,6])
with colr1:
    if st.button("🔄 Rafraîchir"):
        st.rerun()

# =============== Données ===============
with DatabaseService.get_session() as db:
    service = TemplateService(db)
    templates = service.list_templates(active_only=True)

    # Construire une table enrichie avec facettes
    enriched = []
    for t in templates:
        facets = service.extract_template_facets(t.id)
        enriched.append({
            "id": t.id,
            "name": t.name,
            "version": t.version,
            "description": t.description,
            "ppt_path": t.ppt_template_path,
            "excel_path": t.excel_template_path,  
            "is_active": t.is_active,
            "gabarits": facets["gabarits"],
            "themes": facets["themes"],
            "families": facets["families"],
        })


# Filtres (nom + facettes)
def _pass_sources(item) -> bool:
    if not selected_gabarits:
        return True
    G = set(item["gabarits"])
    S = set(selected_gabarits)
    if not G:
        return False
    if strict_sources:
        # Contient au moins un sélectionné ET aucun gabarit hors sélection
        return (len(G & S) > 0) and (G.issubset(S))
    else:
        # OU logique
        return len(G & S) > 0

def _pass_themes(item) -> bool:
    if not selected_themes:
        return True
    return len(set(item["themes"]) & set(selected_themes)) > 0

def _pass_families(item) -> bool:
    if not selected_families:
        return True
    return len(set(item["families"]) & set(selected_families)) > 0

templates_data = [
    {
        "id": e["id"],
        "name": e["name"],
        "version": e["version"],
        "description": e["description"],
        "ppt_path": e["ppt_path"],
        "excel_path": e.get("excel_path"), 
        "is_active": e["is_active"],
    }

    for e in enriched
    if (search.lower() in (e["name"] or "").lower() if search else True)
    and _pass_sources(e)
    and _pass_themes(e)
    and _pass_families(e)
]


# =============== Affichage ===============
if not templates_data:
    st.info("Aucun template trouvé. Créez-en un pour démarrer.")
    if st.button("➕ Nouveau template", type="primary", use_container_width=True):
        st.session_state.selected_template = None
        st.switch_page("pages/_2b_➕_Form_Template.py")
else:
    top1, top2 = st.columns([3, 1])
    with top1:
        st.markdown(f"**{len(templates_data)} template(s) trouvé(s)**")
    with top2:
        if st.button("➕ Nouveau template", type="primary", use_container_width=True):
            st.session_state.selected_template = None
            st.switch_page("pages/_2b_➕_Form_Template.py")

    st.markdown("")

    cols_per_row = 2 if len(templates_data) <= 2 else 3
    for i in range(0, len(templates_data), cols_per_row):
        cols = st.columns(cols_per_row)
        for j, col in enumerate(cols):
            idx = i + j
            if idx >= len(templates_data):
                continue
            t = templates_data[idx]

            c1, c2 = _gradient_from_text(f"{t['name']}-{t['version']}")
            init = _initials(t["name"])
            status = "Actif" if t["is_active"] else "Inactif"

            with col:
                # ✅ La carte est le formulaire lui-même : boutons "dans" la carte
                with st.form(f"card_{t['id']}", clear_on_submit=False):
                    st.markdown(
                        f"""
                        <div class="ka-card__header" style="background: linear-gradient(135deg, {c1}, {c2});">
                          <div class="ka-avatar">{init}</div>
                          <div class="ka-meta">
                            <h3 class="ka-title">{t['name']}</h3>
                            <div class="ka-sub">Version {t['version']}</div>
                          </div>
                          <div class="ka-pill">{status}</div>
                        </div>
                        <div class="ka-card__body">
                          <p class="ka-desc">{_compact_desc(t['description'])}</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    st.markdown('<div class="ka-card__actions">', unsafe_allow_html=True)
                    spacer1, a1, a2, a3, spacer2 = st.columns([0.5, 4, 4, 4, 0.5])
                    open_clicked      = a1.form_submit_button("📊 Ouvrir", use_container_width=True)
                    duplicate_clicked = a2.form_submit_button("📄 Dupliquer", use_container_width=True)   # 👈 nouveau
                    delete_clicked    = a3.form_submit_button("🗑️ Supprimer", use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)

                    if open_clicked:
                        st.session_state.selected_template_detail = t["id"]
                        st.switch_page("pages/_2a_📊_Detail_Livrable.py")

                    if duplicate_clicked:
                        # Ouvrir la modale de duplication
                        st.session_state._dup_template = {
                            "id": t["id"],
                            "name": t["name"],
                            "version": t["version"],
                            "ppt_path": t.get("ppt_path"),
                            "excel_path": t.get("excel_path"),
                        }
                        st.session_state.show_duplicate_modal = True
                        st.rerun()

                    if delete_clicked:
                        st.session_state.delete_template_id = t["id"]
                        st.session_state.show_delete_modal = True
                        st.rerun()


st.divider()

# =============== Modal suppression ===============
if st.session_state.get("show_delete_modal"):
    @st.dialog("⚠️ Confirmer la suppression")
    def confirm_delete():
        template_id = st.session_state.get("delete_template_id")
        template_to_delete = None
        try:
            local_list = templates_data
        except NameError:
            local_list = []
        for _t in local_list:
            if _t["id"] == template_id:
                template_to_delete = _t
                break

        if template_to_delete:
            st.warning("**Vous êtes sur le point de supprimer le template :**")
            st.markdown(f"### {template_to_delete['name']} (v{template_to_delete['version']})")
            st.divider()
            st.markdown("**Action irréversible.** Tapez le nom exact du template pour confirmer :")

            confirmation = st.text_input("Nom du template", key="delete_confirm_input",
                                         placeholder=template_to_delete["name"])
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Annuler", use_container_width=True):
                    st.session_state.show_delete_modal = False
                    st.session_state.pop("delete_template_id", None)
                    st.rerun()
            with c2:
                if st.button("Supprimer définitivement", type="primary", use_container_width=True):
                    try:
                        if confirmation != template_to_delete["name"]:
                            st.error("❌ Le nom ne correspond pas")
                        else:
                            with DatabaseService.get_session() as db:
                                TemplateService(db).delete_template(template_id)
                            st.success(f"✅ Template '{template_to_delete['name']}' supprimé")
                            st.session_state.show_delete_modal = False
                            st.session_state.pop("delete_template_id", None)
                            st.rerun()
                    except Exception as e:
                        st.error(f"❌ Erreur lors de la suppression : {e}")
    confirm_delete()

# =============== Modal duplication ===============
if st.session_state.get("show_duplicate_modal"):
    @st.dialog("📄 Dupliquer le template")
    def duplicate_template_dialog():
        from pathlib import Path
        from backend.services.template_service import TemplateService
        from backend.services.database_service import DatabaseService

        src = st.session_state.get("_dup_template") or {}
        src_id      = src.get("id")
        src_name    = src.get("name") or ""
        src_version = src.get("version") or "1.0"
        src_ppt     = src.get("ppt_path")
        src_excel   = src.get("excel_path")

        st.markdown("**Source :**")
        st.markdown(f"- Nom : `{src_name}`  \n- Version : `{src_version}`")
        st.divider()

        new_name = st.text_input("Nouveau nom du template", key="dup_new_name",
                                 placeholder=f"{src_name} (copie)")
        new_version = st.text_input("Version", key="dup_new_version",
                                    value=src_version)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("Annuler", use_container_width=True, key="dup_cancel"):
                st.session_state.show_duplicate_modal = False
                st.session_state.pop("_dup_template", None)
                st.rerun()

        with c2:
            can_go = bool((new_name or "").strip())
            if st.button("Dupliquer", type="primary", disabled=not can_go,
                         use_container_width=True, key="dup_confirm"):
                try:
                    with DatabaseService.get_session() as db:
                        svc = TemplateService(db)
                        # 1) Charger la config d'origine
                        cfg = svc.load_template_config(src_id)
                        # 2) Muter le nom + version
                        cfg.name = new_name.strip()
                        cfg.version = (new_version or src_version).strip()
                        # 3) Construire les chemins PPT/Excel (s'ils existent)
                        ppt_src = Path(src_ppt) if src_ppt else None
                        if ppt_src and not ppt_src.exists():
                            ppt_src = None
                        excel_src = Path(src_excel) if src_excel else None
                        if excel_src and not excel_src.exists():
                            excel_src = None
                        # 4) Créer le nouveau template (duplication physique + en base)
                        svc.duplicate_template(
                            src_template_id=src_id,
                            new_name=new_name.strip(),
                            new_version=(new_version or src_version).strip(),
                            user_id=1,   # adapte si tu gères l'utilisateur courant
                        )


                    st.session_state["_flash_success"] = f"✅ Template '{new_name}' créé (duplication de '{src_name}')"
                    st.session_state.show_duplicate_modal = False
                    st.session_state.pop("_dup_template", None)
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Erreur lors de la duplication : {e}")

    duplicate_template_dialog()
