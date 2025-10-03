"""
Page de génération de rapports
"""

import streamlit as st
from pathlib import Path
import sys

project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from backend.services.database_service import DatabaseService
from backend.services.template_service import TemplateService
from backend.services.report_service import ReportService
from backend.database.models import ExecutionJob

st.set_page_config(page_title="Générer Rapport", page_icon="▶️", layout="wide")

st.title("▶️ Générer un Rapport")

# Sélection du template
with DatabaseService.get_session() as db:
    service = TemplateService(db)
    templates = service.list_templates(active_only=True)
    
    template_options = {t.name: t.id for t in templates}

if not template_options:
    st.error("Aucun template disponible. Créez-en un d'abord.")
    st.stop()

# Pré-sélectionner si venant de la bibliothèque
preselected_id = st.session_state.get('selected_template_for_generation')
if preselected_id:
    # Trouver le nom correspondant à l'ID
    preselected_name = next((name for name, tid in template_options.items() if tid == preselected_id), None)
    default_index = list(template_options.keys()).index(preselected_name) if preselected_name else 0

else:
    default_index = 0

selected_template_name = st.selectbox(
    "Template à utiliser",
    options=list(template_options.keys()),
    index=default_index
)

st.session_state['selected_template_for_generation'] = template_options[selected_template_name]
template_id = template_options[selected_template_name]

st.divider()

# Charger la config du template
with DatabaseService.get_session() as db:
    service = TemplateService(db)
    template_config = service.load_template_config(template_id)

# Afficher les infos du template
with st.expander("ℹ️ Informations sur le template", expanded=False):
    st.markdown(f"**Description :** {template_config.description or 'Aucune'}")
    st.markdown(f"**Version :** {template_config.version}")
    st.markdown(f"**Source de données :** {template_config.data_source.type}")
    if template_config.data_source.required_tables:
        st.markdown(f"**Tables utilisées :** {', '.join(template_config.data_source.required_tables)}")

st.divider()

# Formulaire des paramètres
st.header("📝 Paramètres du rapport")

parameters = {}

if template_config.parameters:
    for param in template_config.parameters:
        label = f"{param.name}" + (" *" if param.required else "")
        
        if param.type == "string":
            if param.allowed_values:
                value = st.selectbox(
                    label,
                    options=param.allowed_values,
                    help=param.description
                )
            else:
                value = st.text_input(
                    label,
                    value=str(param.default) if param.default else "",
                    help=param.description
                )
        
        elif param.type == "integer":
            value = st.number_input(
                label,
                value=int(param.default) if param.default else 0,
                help=param.description
            )
        
        elif param.type == "date":
            value = st.date_input(
                label,
                help=param.description
            )
        
        else:
            value = st.text_input(label, help=param.description)
        
        parameters[param.name] = value
else:
    st.info("Ce template ne nécessite aucun paramètre.")

st.divider()

# Configuration avancée
with st.expander("⚙️ Configuration avancée", expanded=False):
    custom_output_name = st.text_input(
        "Nom personnalisé du rapport (optionnel)",
        placeholder="Laissez vide pour un nom automatique"
    )

st.divider()

# Bouton de génération
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    if st.button("🚀 Générer le rapport", type="primary", use_container_width=True):
        # Validation
        missing_params = []
        for param in template_config.parameters:
            if param.required and not parameters.get(param.name):
                missing_params.append(param.name)
        
        if missing_params:
            st.error(f"Paramètres obligatoires manquants : {', '.join(missing_params)}")
        else:
            # Créer le job d'exécution
            with DatabaseService.get_session() as db:
                job = ExecutionJob(
                    template_id=template_id,
                    parameters=parameters,
                    status='running'
                )
                db.add(job)
                db.commit()
                job_id = job.id
            
            # Lancer la génération
            with st.spinner("Génération en cours..."):
                try:
                    report_service = ReportService(template_config)
                    result = report_service.generate_report(
                        parameters=parameters,
                        output_name=custom_output_name if custom_output_name else None
                    )
                    
                    # Mettre à jour le job
                    with DatabaseService.get_session() as db:
                        job = db.query(ExecutionJob).filter_by(id=job_id).first()
                        if result['success']:
                            job.status = 'completed'
                            job.output_file_path = result['pptx_path']
                            job.execution_time_seconds = result['execution_time_seconds']
                            
                            st.success(f"✅ Rapport généré avec succès en {result['execution_time_seconds']:.1f}s")
                            
                            # Afficher les fichiers générés
                            st.markdown("**Fichiers générés :**")
                            st.code(result['pptx_path'])
                            st.code(result['excel_path'])
                            
                            # Boutons de téléchargement
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                with open(result['pptx_path'], 'rb') as f:
                                    st.download_button(
                                        "📥 Télécharger PowerPoint",
                                        data=f.read(),
                                        file_name=Path(result['pptx_path']).name,
                                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                                    )
                            
                            with col2:
                                with open(result['excel_path'], 'rb') as f:
                                    st.download_button(
                                        "📥 Télécharger Excel",
                                        data=f.read(),
                                        file_name=Path(result['excel_path']).name,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                            
                            st.balloons()
                        else:
                            job.status = 'failed'
                            job.error_message = result['error']
                            job.execution_time_seconds = result['execution_time_seconds']
                            
                            st.error(f"❌ Erreur lors de la génération : {result['error']}")
                        
                        db.commit()
                
                except Exception as e:
                    st.error(f"❌ Erreur critique : {e}")
                    
                    # Mettre à jour le job
                    with DatabaseService.get_session() as db:
                        job = db.query(ExecutionJob).filter_by(id=job_id).first()
                        job.status = 'failed'
                        job.error_message = str(e)
                        db.commit()

st.divider()

# Historique des dernières exécutions
st.header("📊 Dernières exécutions")

with DatabaseService.get_session() as db:
    recent_jobs = db.query(ExecutionJob).filter_by(
        template_id=template_id
    ).order_by(ExecutionJob.created_at.desc()).limit(5).all()
    
    if recent_jobs:
        for job in recent_jobs:
            status_icon = {
                'completed': '✅',
                'failed': '❌',
                'running': '⏳'
            }.get(job.status, '❓')
            
            with st.expander(f"{status_icon} {job.created_at.strftime('%d/%m/%Y %H:%M')} - {job.status}"):
                st.json(job.parameters)
                if job.execution_time_seconds:
                    st.text(f"Durée : {job.execution_time_seconds:.1f}s")
                if job.error_message:
                    st.error(job.error_message)
    else:
        st.info("Aucune exécution pour ce template")