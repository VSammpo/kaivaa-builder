# frontend/pages/_3b_➕_Form_Gabarit.py
import streamlit as st
import pandas as pd
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.gabarit_registry import (
    get_gabarit, upsert_gabarit, list_gabarits,
    get_relations, add_relation, delete_relation,
    set_role, get_role,
    get_default_source, set_default_source, clear_default_source,
    set_default_preview
)
from backend.models.gabarits import TableGabarit, GabaritColumn

try:
    from backend.services.dataset_service import align_df_to_expected_columns
except Exception:
    def align_df_to_expected_columns(df: pd.DataFrame, expected_columns):
        expected = [c for c in (expected_columns or []) if isinstance(c, str) and c.strip()]
        cur_cols = list(df.columns)
        missing = [c for c in expected if c not in cur_cols]
        for c in missing:
            df[c] = pd.NA
        ordered = expected + [c for c in df.columns if c not in expected]
        return df[ordered], {"missing": missing, "extra": [c for c in cur_cols if c not in expected]}

st.set_page_config(page_title="Formulaire Gabarit", page_icon="➕", layout="wide")

# CSS
st.markdown("""
<style>
.success-box {
    background: #d4edda;
    border-left: 4px solid #28a745;
    padding: 1rem;
    border-radius: 4px;
    margin: 1rem 0;
}
.warning-box {
    background: #fff3cd;
    border-left: 4px solid #ffc107;
    padding: 1rem;
    border-radius: 4px;
    margin: 1rem 0;
}
</style>
""", unsafe_allow_html=True)

# ============= HELPERS =============
def _try_load_source(fmt: str, path: str, sep: str | None, enc: str | None, head: int = 20):
    try:
        p = Path(path)
        if not p.exists() or not p.is_file():
            return None, f"Fichier introuvable : {path}"
        if fmt == "csv":
            df = pd.read_csv(path, sep=sep or ";", encoding=enc or "utf-8-sig")
        elif fmt == "parquet":
            df = pd.read_parquet(path)
        else:
            return None, f"Format non supporté : {fmt}"
        return df.head(head), None
    except Exception as e:
        return None, str(e)

def _apply_python(df: pd.DataFrame, code: str | None):
    if not code or not isinstance(code, str) or not code.strip():
        return df, None
    try:
        loc = {"df": df, "pd": pd}
        exec(code, {}, loc)
        new_df = loc.get("df")
        if isinstance(new_df, pd.DataFrame):
            return new_df, None
        return df, None
    except Exception as e:
        return None, f"Erreur script Python : {e}"

# ============= MODE =============
edit_mode = False
gabarit = None
if "selected_gabarit" in st.session_state and st.session_state.selected_gabarit:
    edit_mode = True
    gab_name, gab_version = st.session_state.selected_gabarit
    gabarit = get_gabarit(gab_name, gab_version)

# ============= EN-TÊTE =============
col_back, col_title = st.columns([1, 5])

with col_back:
    if st.button("← Retour", use_container_width=True):
        if edit_mode:
            st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")
        else:
            st.switch_page("pages/3_🧱_Gabarits.py")

with col_title:
    if edit_mode:
        st.title(f"✏️ Modifier : {gabarit.name}")
    else:
        st.title("➕ Créer un gabarit")

st.divider()

# ============= VALEURS PAR DÉFAUT =============
if edit_mode:
    name_init = gabarit.name
    version_init = gabarit.version
    desc_init = gabarit.description or ""
    cols_init = pd.DataFrame([c.model_dump() for c in gabarit.columns])
    existing_role = get_role(gabarit.name, gabarit.version) or "mixed"
else:
    name_init = ""
    version_init = "v1"
    desc_init = ""
    cols_init = pd.DataFrame([{"name": "", "type": "text", "is_key": False}])
    existing_role = "mixed"

# ============= FORMULAIRE PRINCIPAL =============
st.subheader("1️⃣ Informations générales")

with st.form("gabarit_form"):
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        name = st.text_input(
            "Nom du gabarit *",
            value=name_init,
            disabled=edit_mode,
            placeholder="Ex: SELL_OUT, PRICING_SOURCE",
        )
    with col2:
        version = st.text_input("Version", value=version_init, placeholder="Ex: v1")
    with col3:
        table_role = st.selectbox(
            "Rôle",
            ["fact", "dimension", "mixed"],
            index=["fact", "dimension", "mixed"].index(existing_role),
            help="Fact = événements • Dimension = référentiel • Mixed = hybride"
        )

    description = st.text_area("Description", value=desc_init, height=80,
                              placeholder="Décrivez l'utilité et le contenu de ce gabarit...")

    st.divider()
    st.subheader("2️⃣ Structure des colonnes")
    st.caption("Types disponibles : text | number | integer | date | boolean")

    edited = st.data_editor(
        cols_init,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "name": st.column_config.TextColumn("Nom *", help="Nom de la colonne", width="medium"),
            "type": st.column_config.SelectboxColumn("Type", 
                options=["text", "number", "integer", "date", "boolean"], width="small"),
            "is_key": st.column_config.CheckboxColumn("Clé ?", 
                help="Fait partie de la clé composite", width="small"),
        },
        key="columns_editor",
    )

    st.divider()
    
    col_submit, col_cancel = st.columns([1, 1])
    
    with col_submit:
        submitted = st.form_submit_button(
            "💾 Enregistrer le gabarit" if edit_mode else "🚀 Créer le gabarit",
            type="primary",
            use_container_width=True,
        )
    
    with col_cancel:
        if st.form_submit_button("Annuler", use_container_width=True):
            if edit_mode:
                st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")
            else:
                st.switch_page("pages/3_🧱_Gabarits.py")

    if submitted:
        if not name.strip():
            st.error("❌ Le nom est obligatoire")
        else:
            edited = edited.fillna("")
            cols = []
            seen = set()
            for r in edited.to_dict(orient="records"):
                n = (r.get("name") or "").strip()
                if not n or n in seen:
                    continue
                t = (r.get("type") or "text").strip().lower()
                is_key = bool(r.get("is_key", False))
                cols.append(GabaritColumn(name=n, type=t, is_key=is_key))
                seen.add(n)

            if not cols:
                st.error("❌ Au moins une colonne est requise")
            else:
                try:
                    g = TableGabarit(
                        name=name.strip(),
                        version=version.strip(),
                        description=(description or "").strip(),
                        columns=cols,
                    )
                    upsert_gabarit(g)
                    set_role(g.name, g.version, table_role)
                    st.success(f"✅ Gabarit enregistré : {g.name} v{g.version}")

                    st.session_state.selected_gabarit = (g.name, g.version)
                    st.switch_page("pages/_3a_🧱_Detail_Gabarit.py")
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")

# ============= ENRICHISSEMENTS (si édition) =============
if edit_mode:
    st.divider()
    st.subheader("3️⃣ Enrichissements (optionnel)")
    
    with st.expander("💡 Principe des enrichissements", expanded=False):
        st.markdown("""
        Vous enrichissez **cette table** (source) avec des données provenant d'une **table de référence** (dimension).
        
        **Exemple :** enrichir une table de CA par `siren` avec le nom d'entreprise depuis une table SIREN.
        
        - **Clé locale** : la colonne dans cette table (ex: `siren`)
        - **Clé d'enrichissement** : la colonne dans la table de référence (ex: `siren`)
        """)
    
    gab = get_gabarit(gabarit.name, gabarit.version)
    local_cols = [c.name for c in (gab.columns or [])]

    # Liste des enrichissements existants
    existing_relations = get_relations(gab.name, gab.version)
    if existing_relations:
        st.caption(f"📋 {len(existing_relations)} enrichissement(s) configuré(s)")
        for r in existing_relations:
            with st.container(border=True):
                col_info, col_del = st.columns([4, 1])
                with col_info:
                    st.markdown(f"**Depuis** `{r['to_gabarit']}` [{r.get('to_version','v1')}]")
                    st.caption(f"Jointure : `{r['left_key']}` = `{r['right_key']}`")
                with col_del:
                    if st.button("🗑️", key=f"del_enrich_{r['relation_id']}", use_container_width=True):
                        try:
                            delete_relation(
                                r["from_gabarit"], r.get("from_version", "v1"),
                                r["to_gabarit"], r.get("to_version", "v1"),
                                r["left_key"], r["right_key"],
                            )
                            st.rerun()
                        except Exception as e:
                            st.error(f"Suppression impossible : {e}")
    
    # Formulaire d'ajout
    with st.form("enrichment_form"):
        st.caption("Ajouter un nouvel enrichissement")
        
        all_gabs = list_gabarits()
        options = [
            f"{g.name}|{g.version}"
            for g in all_gabs
            if not (g.name == gab.name and g.version == gab.version)
        ]
        
        target = st.selectbox(
            "Table de référence",
            options=options,
            index=0 if options else None,
            format_func=lambda s: f"{s.split('|')[0]} [{s.split('|')[1]}]" if '|' in s else s,
            help="Choisissez la table qui contient les données d'enrichissement"
        )
        
        target_cols = []
        if target:
            tgt_name, tgt_ver = target.split("|", 1)
            tgt_gab = get_gabarit(tgt_name, tgt_ver)
            if tgt_gab:
                target_cols = [c.name for c in (tgt_gab.columns or [])]
        
        col1, col2 = st.columns(2)
        with col1:
            left_key = st.selectbox(
                "Clé dans cette table",
                options=local_cols,
                index=0 if local_cols else None,
            )
        with col2:
            if target_cols:
                right_key = st.selectbox(
                    "Clé dans la table de référence",
                    options=target_cols,
                    index=0,
                )
            else:
                right_key = st.text_input("Clé dans la table de référence", placeholder="ex: siren")
        
        add_ok = st.form_submit_button("➕ Ajouter l'enrichissement", use_container_width=True)

        if add_ok:
            if not target or not left_key or not right_key:
                st.error("❌ Veuillez remplir tous les champs")
            else:
                try:
                    tgt_name, tgt_ver = target.split("|", 1)
                    add_relation(
                        from_gabarit=gab.name, from_version=gab.version,
                        to_gabarit=tgt_name, to_version=tgt_ver,
                        left_key=left_key,
                        right_key=right_key,
                    )
                    st.success("✅ Enrichissement ajouté")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")

# ============= DONNÉES PAR DÉFAUT (si édition) =============
if edit_mode:
    st.divider()
    st.subheader("4️⃣ Donnée par défaut (optionnel)")
    
    gab = get_gabarit(gabarit.name, gabarit.version)
    expected_cols = [c.name for c in (gab.columns or [])]
    current_default = get_default_source(gab.name, gab.version) or {}

    with st.expander("💡 À quoi sert la donnée par défaut ?", expanded=False):
        st.markdown("""
        La donnée par défaut permet de :
        - Tester rapidement les méthodes de calcul
        - Valider la structure du gabarit
        - Servir d'exemple pour les utilisateurs
        
        **Aucun fichier n'est uploadé**, vous fournissez simplement un chemin accessible par le serveur.
        """)
    
    use_default = st.checkbox("Activer une donnée par défaut", value=bool(current_default))
    
    if use_default:
        # Configuration source
        col1, col2 = st.columns([1, 3])
        with col1:
            fmt = st.selectbox("Format", ["csv", "parquet"], 
                             index=0 if current_default.get("type") == "csv" else 1)
        with col2:
            path = st.text_input(
                "Chemin du fichier",
                value=current_default.get("path", ""),
                placeholder=r"Ex: C:\data\sources\fichier.csv",
                help="Chemin accessible par le serveur"
            )
        
        if fmt == "csv":
            col_sep, col_enc = st.columns(2)
            with col_sep:
                sep = st.text_input("Séparateur", value=current_default.get("sep", ";"))
            with col_enc:
                enc = st.text_input("Encodage", value=current_default.get("encoding", "utf-8-sig"))
        else:
            sep, enc = None, None
        
        # Code Python
        with st.expander("Transformation Python (optionnel)", expanded=bool(current_default.get("python"))):
            python_code = st.text_area(
                "Code de transformation (df → df)",
                value=current_default.get("python", ""),
                height=150,
                placeholder="# Exemples :\n# df = df.rename(columns={'A': 'B'})\n# df = df[df['col'] > 0]",
                help="Transformations pandas appliquées après le chargement"
            )
        
        st.divider()
        
        # Actions
        col_preview, col_validate, col_save, col_clear = st.columns(4)
        
        with col_preview:
            if st.button("👁️ Aperçu", use_container_width=True, disabled=not path):
                with st.spinner("Chargement..."):
                    df, err = _try_load_source(fmt, path, sep, enc, head=20)
                    if err:
                        st.error(f"❌ {err}")
                    else:
                        df2, perr = _apply_python(df, python_code)
                        if perr:
                            st.error(f"❌ {perr}")
                        else:
                            st.success(f"✅ Aperçu chargé : {df2.shape[0]} lignes × {df2.shape[1]} colonnes")
                            st.dataframe(df2, use_container_width=True, height=300)
        
        with col_validate:
            if st.button("✅ Valider", use_container_width=True, disabled=not path):
                with st.spinner("Validation..."):
                    df, err = _try_load_source(fmt, path, sep, enc, head=100)
                    if err:
                        st.error(f"❌ {err}")
                    else:
                        df2, perr = _apply_python(df, python_code)
                        if perr:
                            st.error(f"❌ {perr}")
                        else:
                            aligned, warns = align_df_to_expected_columns(df2.copy(), expected_cols)
                            
                            if warns.get("missing"):
                                st.warning(f"⚠️ Colonnes manquantes : {', '.join(warns['missing'])}")
                            if warns.get("extra"):
                                st.info(f"ℹ️ Colonnes supplémentaires : {', '.join(warns['extra'])}")
                            if not warns.get("missing") and not warns.get("extra"):
                                st.success("✅ Structure parfaitement alignée")
        
        with col_save:
            if st.button("💾 Enregistrer", use_container_width=True, disabled=not path, type="primary"):
                with st.spinner("Enregistrement..."):
                    # Charger un échantillon pour l'aperçu
                    df20, err = _try_load_source(fmt, path, sep, enc, head=20)
                    if err:
                        st.error(f"❌ {err}")
                    else:
                        df20, perr = _apply_python(df20, python_code)
                        if perr:
                            st.error(f"❌ {perr}")
                        else:
                            # Sauvegarder la config
                            src = {"type": fmt, "path": str(Path(path).resolve())}
                            if fmt == "csv":
                                src.update({"sep": sep or ";", "encoding": enc or "utf-8-sig"})
                            if python_code and python_code.strip():
                                src["python"] = python_code
                            
                            set_default_source(gab.name, gab.version, src)
                            
                            # Sauvegarder l'aperçu
                            try:
                                sample = df20.head(20)
                                set_default_preview(
                                    gab.name, gab.version,
                                    rows=sample.to_dict(orient="records"),
                                    columns=list(sample.columns)
                                )
                                st.success("✅ Donnée par défaut enregistrée avec aperçu")
                                st.rerun()
                            except Exception as e:
                                st.warning(f"Source enregistrée mais aperçu non sauvegardé : {e}")
        
        with col_clear:
            if st.button("🗑️ Retirer", use_container_width=True, disabled=not current_default):
                clear_default_source(gab.name, gab.version)
                st.success("✅ Donnée par défaut retirée")
                st.rerun()
    
    elif current_default:
        st.info("La donnée par défaut a été désactivée. Activez la case ci-dessus pour la reconfigurer.")