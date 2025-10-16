"""
Page de détail d'une transformation
"""

import streamlit as st
import pandas as pd
import json
from datetime import datetime

from backend.services.transformation_service import (
    get_transformation,
    validate_transformation,
    get_transformation_input_columns,
    get_transformation_output_columns
)
from backend.services.gabarit_registry import get_gabarit
from backend.services.table_builder_service import build_table_from_transformation

st.set_page_config(
    page_title="Détail Transformation",
    page_icon="🔀",
    layout="wide"
)

# Vérifier qu'une transformation est sélectionnée
if "selected_transformation" not in st.session_state:
    st.error("Aucune transformation sélectionnée")
    if st.button("🔙 Retour à la liste"):
        st.switch_page("pages/_3d_🔀_Transformations.py")
    st.stop()

name, version = st.session_state.selected_transformation

# Charger la transformation
transformation = get_transformation(name, version)
if not transformation:
    st.error(f"Transformation '{name}' v{version} introuvable")
    st.stop()

# En-tête
st.title(f"🔀 {transformation['name']}")
st.caption(f"Version : {transformation['version']}")

if transformation.get("description"):
    st.info(transformation["description"])

# Barre d'actions
col1, col2, col3, col4 = st.columns([1, 1, 1, 3])
with col1:
    if st.button("✏️ Éditer", type="primary", use_container_width=True):
        st.session_state.edit_mode = True
        st.switch_page("pages/_3d2_🔀_Builder_Transformation.py")
with col2:
    if st.button("📋 Dupliquer", use_container_width=True):
        st.info("Fonctionnalité à venir")
with col3:
    if st.button("🔙 Retour", use_container_width=True):
        st.switch_page("pages/_3d_🔀_Transformations.py")

st.markdown("---")

# Tabs principaux
tab_config, tab_preview, tab_usage, tab_json = st.tabs([
    "⚙️ Configuration",
    "👁️ Prévisualisation",
    "📊 Utilisation",
    "💾 Export JSON"
])

# ==================== TAB CONFIGURATION ====================
with tab_config:
    # Validation
    validation = validate_transformation(name, version)
    
    if validation["valid"]:
        st.success("✅ Configuration valide")
    else:
        st.error("❌ Configuration invalide")
        for error in validation["errors"]:
            st.error(f"• {error}")
    
    if validation["warnings"]:
        for warning in validation["warnings"]:
            st.warning(f"• {warning}")
    
    # Détails de configuration
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📊 Gabarit de base")
        gabarit_base = transformation.get("gabarit_base", {})
        gab_name = gabarit_base.get("name", "N/A")
        gab_version = gabarit_base.get("version", "v1")
        
        st.markdown(f"**Nom :** {gab_name}")
        st.markdown(f"**Version :** {gab_version}")
        
        # Vérifier que le gabarit existe
        gabarit = get_gabarit(gab_name, gab_version)
        if gabarit:
            st.success("Gabarit disponible")
            n_cols_base = len(gabarit.columns)
            st.metric("Colonnes disponibles", n_cols_base)
        else:
            st.error("Gabarit introuvable")
        
        st.markdown("### 📝 Colonnes sélectionnées")
        columns_enabled = transformation.get("columns_enabled", [])
        if columns_enabled:
            st.metric("Colonnes conservées", len(columns_enabled))
            
            # Afficher les colonnes
            cols = st.columns(min(len(columns_enabled), 3))
            for idx, col_name in enumerate(columns_enabled[:9]):  # Max 9 pour lisibilité
                with cols[idx % len(cols)]:
                    st.caption(f"• {col_name}")
            
            if len(columns_enabled) > 9:
                st.caption(f"... et {len(columns_enabled) - 9} autres")
        else:
            st.info("Toutes les colonnes du gabarit")
    
    with col2:
        st.markdown("### 🔗 Enrichissements")
        enrichments = transformation.get("enrichments", [])
        if enrichments:
            st.metric("Nombre d'enrichissements", len(enrichments))
            
            for idx, enrich in enumerate(enrichments):
                with st.expander(f"Enrichissement #{idx + 1}"):
                    st.markdown(f"**Type :** {enrich.get('join', 'left')}")
                    
                    path = enrich.get("path", [])
                    if path:
                        for step in path:
                            if len(step) >= 4:
                                st.caption(f"{step[0]}.{step[1]} → {step[2]}.{step[3]}")
                    
                    cols_to_add = enrich.get("columns", [])
                    if cols_to_add:
                        st.markdown(f"**Colonnes ajoutées :** {', '.join(cols_to_add)}")
        else:
            st.info("Aucun enrichissement")
        
        st.markdown("### ⚙️ Méthodes")
        methods = transformation.get("methods", [])
        if methods:
            st.metric("Méthodes appliquées", len(methods))
            for method in methods:
                st.caption(f"• {method}")
        else:
            st.info("Aucune méthode")
    
    # Script Python
    st.markdown("### 🐍 Script Python")
    overlay_python = transformation.get("overlay_python", "")
    if overlay_python:
        with st.expander("Voir le script", expanded=False):
            st.code(overlay_python, language="python")
        
        # Compter les lignes
        lines = overlay_python.strip().split('\n')
        st.metric("Lignes de code", len(lines))
    else:
        st.info("Aucun script Python")
    
    # Finalisation
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 📋 Ordre final")
        final_order = transformation.get("final_order", [])
        if final_order:
            st.metric("Colonnes ordonnées", len(final_order))
        else:
            st.info("Ordre par défaut")
    
    with col2:
        st.markdown("### 🔄 Renommages")
        final_renames = transformation.get("final_renames", {})
        if final_renames:
            st.metric("Colonnes renommées", len(final_renames))
            for old, new in list(final_renames.items())[:3]:
                st.caption(f"{old} → {new}")
            if len(final_renames) > 3:
                st.caption(f"... et {len(final_renames) - 3} autres")
        else:
            st.info("Aucun renommage")
    
    with col3:
        st.markdown("### ❌ Exclusions")
        final_excludes = transformation.get("final_excludes", [])
        if final_excludes:
            st.metric("Colonnes exclues", len(final_excludes))
            for col in final_excludes[:3]:
                st.caption(f"• {col}")
            if len(final_excludes) > 3:
                st.caption(f"... et {len(final_excludes) - 3} autres")
        else:
            st.info("Aucune exclusion")
    
    # Métadonnées
    st.markdown("### 📅 Métadonnées")
    col1, col2 = st.columns(2)
    
    with col1:
        created_at = transformation.get("created_at", "")
        if created_at:
            try:
                dt = datetime.fromisoformat(created_at)
                st.markdown(f"**Créé le :** {dt.strftime('%d/%m/%Y à %H:%M')}")
            except:
                st.markdown(f"**Créé le :** {created_at}")
    
    with col2:
        updated_at = transformation.get("updated_at", "")
        if updated_at:
            try:
                dt = datetime.fromisoformat(updated_at)
                st.markdown(f"**Modifié le :** {dt.strftime('%d/%m/%Y à %H:%M')}")
            except:
                st.markdown(f"**Modifié le :** {updated_at}")

# ==================== TAB PREVIEW ====================
with tab_preview:
    st.markdown("### Prévisualisation des données transformées")
    
    col1, col2, col3 = st.columns([1, 1, 3])
    
    with col1:
        mode = st.radio("Mode", ["Preview", "Complet"], index=0)
        full = (mode == "Complet")
    
    with col2:
        if st.button("🔄 Générer", type="primary", use_container_width=True):
            with st.spinner("Construction en cours..."):
                df, error = build_table_from_transformation(
                    name,
                    version,
                    full=full,
                    log_kpis=True
                )
                
                if error:
                    st.error(f"Erreur : {error}")
                    st.session_state.preview_result = None
                else:
                    st.session_state.preview_result = df
                    st.success(f"✅ Généré : {len(df)} lignes × {len(df.columns)} colonnes")
    
    # Afficher le résultat
    if "preview_result" in st.session_state and st.session_state.preview_result is not None:
        df = st.session_state.preview_result
        
        # Statistiques
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Lignes", f"{len(df):,}")
        with col2:
            st.metric("Colonnes", len(df.columns))
        with col3:
            st.metric("Cellules", f"{len(df) * len(df.columns):,}")
        with col4:
            # Taille mémoire approximative
            memory_mb = df.memory_usage(deep=True).sum() / 1024 / 1024
            st.metric("Mémoire", f"{memory_mb:.2f} MB")
        
        # Tabs pour différentes vues
        subtab1, subtab2, subtab3, subtab4 = st.tabs([
            "📊 Données",
            "📈 Statistiques",
            "🏷️ Types",
            "💾 Export"
        ])
        
        with subtab1:
            # Limiter l'affichage pour les grosses tables
            display_limit = 1000
            if len(df) > display_limit:
                st.warning(f"Affichage limité aux {display_limit} premières lignes")
                display_df = df.head(display_limit)
            else:
                display_df = df
            
            st.dataframe(display_df, use_container_width=True)
        
        with subtab2:
            st.markdown("### Statistiques descriptives")
            
            # Sélectionner uniquement les colonnes numériques
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            if numeric_cols:
                stats_df = df[numeric_cols].describe().T
                stats_df['missing'] = df[numeric_cols].isnull().sum()
                stats_df['missing_pct'] = (stats_df['missing'] / len(df) * 100).round(2)
                
                st.dataframe(stats_df, use_container_width=True)
            else:
                st.info("Aucune colonne numérique")
        
        with subtab3:
            st.markdown("### Types de données")
            
            types_df = pd.DataFrame({
                'Colonne': df.columns,
                'Type': [str(df[col].dtype) for col in df.columns],
                'Non-nulls': [df[col].count() for col in df.columns],
                'Nulls': [df[col].isnull().sum() for col in df.columns],
                'Unique': [df[col].nunique() for col in df.columns],
                'Premier exemple': [df[col].dropna().iloc[0] if not df[col].dropna().empty else None 
                                   for col in df.columns]
            })
            
            st.dataframe(types_df, use_container_width=True)
        
        with subtab4:
            st.markdown("### Export des données")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Export CSV
                csv = df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 Télécharger CSV",
                    data=csv,
                    file_name=f"{name}_{version}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            with col2:
                # Export Excel
                from io import BytesIO
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Data', index=False)
                
                st.download_button(
                    label="📥 Télécharger Excel",
                    data=buffer.getvalue(),
                    file_name=f"{name}_{version}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# ==================== TAB USAGE ====================
with tab_usage:
    st.markdown("### 📊 Analyse d'utilisation")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Colonnes d'entrée requises")
        input_cols = get_transformation_input_columns(name, version)
        
        if input_cols:
            st.info(f"{len(input_cols)} colonnes requises en entrée")
            
            for col in input_cols:
                st.caption(f"• {col}")
        else:
            st.warning("Impossible de déterminer les colonnes d'entrée")
    
    with col2:
        st.markdown("#### Colonnes de sortie")
        output_cols = get_transformation_output_columns(name, version)
        
        if output_cols:
            st.success(f"{len(output_cols)} colonnes en sortie")
            
            for col in output_cols[:10]:  # Limiter l'affichage
                st.caption(f"• {col}")
            
            if len(output_cols) > 10:
                st.caption(f"... et {len(output_cols) - 10} autres")
        else:
            # Fallback : utiliser la preview si disponible
            if "preview_result" in st.session_state and st.session_state.preview_result is not None:
                df = st.session_state.preview_result
                st.success(f"{len(df.columns)} colonnes en sortie")
                
                for col in list(df.columns)[:10]:
                    st.caption(f"• {col}")
                
                if len(df.columns) > 10:
                    st.caption(f"... et {len(df.columns) - 10} autres")
            else:
                st.info("Générez une preview pour voir les colonnes de sortie")
    
    # Diagramme de flux
    st.markdown("### 🔄 Flux de transformation")
    
    with st.expander("Voir le pipeline", expanded=True):
        steps = []
        
        # 1. Gabarit de base
        gabarit_base = transformation.get("gabarit_base", {})
        steps.append(f"1️⃣ **Gabarit source :** {gabarit_base.get('name', 'N/A')}")
        
        # 2. Sélection colonnes
        if transformation.get("columns_enabled"):
            steps.append(f"2️⃣ **Sélection :** {len(transformation['columns_enabled'])} colonnes")
        
        # 3. Enrichissements
        if transformation.get("enrichments"):
            steps.append(f"3️⃣ **Enrichissements :** {len(transformation['enrichments'])} jointure(s)")
        
        # 4. Méthodes
        if transformation.get("methods"):
            steps.append(f"4️⃣ **Méthodes :** {len(transformation['methods'])} colonne(s) calculée(s)")
        
        # 5. Script
        if transformation.get("overlay_python"):
            lines = len(transformation['overlay_python'].strip().split('\n'))
            steps.append(f"5️⃣ **Script Python :** {lines} ligne(s)")
        
        # 6. Finalisation
        finalization = []
        if transformation.get("final_renames"):
            finalization.append(f"{len(transformation['final_renames'])} renommages")
        if transformation.get("final_excludes"):
            finalization.append(f"{len(transformation['final_excludes'])} exclusions")
        if transformation.get("final_order"):
            finalization.append("ordre personnalisé")
        
        if finalization:
            steps.append(f"6️⃣ **Finalisation :** {', '.join(finalization)}")
        
        # Afficher le pipeline
        for step in steps:
            st.markdown(step)

# ==================== TAB JSON ====================
with tab_json:
    st.markdown("### 💾 Configuration JSON")
    st.info("Copiez cette configuration pour la réutiliser ou la partager")
    
    # Formatter le JSON
    json_str = json.dumps(transformation, ensure_ascii=False, indent=2)
    
    # Afficher avec coloration syntaxique
    st.code(json_str, language="json")
    
    # Bouton de copie
    st.download_button(
        label="📥 Télécharger JSON",
        data=json_str,
        file_name=f"{name}_{version}.json",
        mime="application/json",
        use_container_width=True
    )
    
    # Statistiques sur le JSON
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Taille", f"{len(json_str):,} caractères")
    with col2:
        st.metric("Lignes", len(json_str.split('\n')))
    with col3:
        st.metric("Ko", f"{len(json_str.encode('utf-8')) / 1024:.2f}")