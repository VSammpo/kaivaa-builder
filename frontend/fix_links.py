"""
Script de correction des liens - Version intelligente
À exécuter depuis le dossier frontend/
Usage: python fix_links_smart.py
"""

import os
import re
from pathlib import Path

def main():
    print("=== Correction intelligente des liens ===\n")
    
    if not os.path.exists("pages"):
        print("ERREUR : Le dossier 'pages' n'existe pas.")
        print("Veuillez exécuter ce script depuis le dossier frontend/")
        return
    
    # Mapping complet : ce qui est dans le CODE -> ce qui existe sur le DISQUE
    code_to_disk = {
        # Les fichiers ont été renommés de 3->1, 3d->2, 2->3, 1->4, 4->5
        # Mais le code pointe encore vers 3, 3d, 2, 1, 4
        
        # Dans le code c'est "3_..." mais sur disque c'est "1_..."
        "pages/3_🧱_Gabarits.py": "pages/1_🧱_Gabarits.py",
        "pages/_3a_🧱_Detail_Gabarit.py": "pages/_1a_🧱_Detail_Gabarit.py",
        "pages/_3b1_🧱_Structure_Gabarit.py": "pages/_1b_🧱_Structure_Gabarit.py",
        "pages/_3b2_🔗_Enrichissements_Gabarit.py": "pages/_1c_🔗_Enrichissements_Gabarit.py",
        "pages/_3b3_📁_Donnee_Par_Defaut.py": "pages/_1d_📁_Donnee_Par_Defaut.py",
        "pages/_3c_⚙️_Methodes_Gabarit.py": "pages/_1e_⚙️_Methodes_Gabarit.py",
        
        # Dans le code c'est "3d_..." mais sur disque c'est "2_..."
        "pages/3d_🔀_Transformations.py": "pages/2_🔀_Transformations.py",
        "pages/_3d1_🔀_Detail_Transformation.py": "pages/_2a_🔀_Detail_Transformation.py",
        "pages/_3d2_🔀_Builder_Transformation.py": "pages/_2b_🔀_Builder_Transformation.py",
        
        # Dans le code c'est "2_..." mais sur disque c'est "3_..."
        "pages/2_📚_Bibliotheque.py": "pages/3_📚_Bibliotheque.py",
        "pages/_2a_📊_Detail_Livrable.py": "pages/_3a_📊_Detail_Livrable.py",
        "pages/_2b_➕_Form_Template.py": "pages/_3b_➕_Form_Template.py",
        "pages/_2b3_📑_Tables_Template.py": "pages/_3c_📑_Tables_Template.py",
        "pages/_2b4_🧾_Ajustement_Table.py": "pages/_3d_🧾_Ajustement_Table.py",
        
        # Dans le code c'est "1_..." mais sur disque c'est "4_..."
        "pages/1_📁_Projets.py": "pages/4_📁_Projets.py",
        "pages/_1a_🗂️_Hub_Projet.py": "pages/_4a_🗂️_Hub_Projet.py",
        "pages/_1b_💾_Data_Projet.py": "pages/_4b_💾_Data_Projet.py",
        "pages/_1c_⚙️_Config_Deliverable.py": "pages/_4c_⚙️_Config_Deliverable.py",
        
        # Dans le code c'est "4_..." mais sur disque c'est "5_..."
        "pages/4_📚_Bibliotheque_Livrables.py": "pages/5_📚_Bibliotheque_Livrables.py",
        "pages/4b_📁_Livrables_Projet.py": "pages/_5a_📁_Livrables_Projet.py",
    }
    
    print("Étape 1/2 : Analyse des fichiers")
    print("----------------------------------")
    
    # Lister tous les fichiers à corriger
    files_to_fix = ["Home.py"]
    for file in Path("pages").glob("*.py"):
        files_to_fix.append(str(file))
    
    print(f"Fichiers à analyser : {len(files_to_fix)}")
    
    print("\nÉtape 2/2 : Correction des liens")
    print("----------------------------------")
    
    total_replacements = 0
    files_modified = 0
    
    for filepath in files_to_fix:
        if not os.path.exists(filepath):
            continue
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            print(f"⚠ Erreur lecture {filepath}: {e}")
            continue
        
        original_content = content
        replacements_in_file = 0
        
        # Pour chaque mapping, remplacer dans le contenu
        for old_path, new_path in code_to_disk.items():
            # Compter combien de fois l'ancien chemin apparaît
            count = content.count(f'"{old_path}"') + content.count(f"'{old_path}'")
            
            if count > 0:
                # Remplacer avec guillemets doubles
                content = content.replace(f'"{old_path}"', f'"{new_path}"')
                # Remplacer avec guillemets simples
                content = content.replace(f"'{old_path}'", f"'{new_path}'")
                
                replacements_in_file += count
        
        # Si le contenu a changé, écrire le fichier
        if content != original_content:
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(content)
                print(f"✓ {filepath} : {replacements_in_file} correction(s)")
                total_replacements += replacements_in_file
                files_modified += 1
            except Exception as e:
                print(f"✗ Erreur écriture {filepath}: {e}")
    
    print(f"\n{'='*50}")
    print(f"Résumé :")
    print(f"  • {files_modified} fichier(s) modifié(s)")
    print(f"  • {total_replacements} lien(s) corrigé(s)")
    print(f"{'='*50}\n")
    
    if total_replacements > 0:
        print("✅ Correction terminée avec succès !")
        print("Relancez votre application Streamlit.")
    else:
        print("⚠ Aucune correction effectuée.")
        print("\nVérification de la situation actuelle :")
        
        # Lister les fichiers qui existent vraiment
        print("\nFichiers présents dans pages/ :")
        for f in sorted(Path("pages").glob("*.py")):
            print(f"  • {f.name}")
        
        # Chercher un exemple de lien cassé
        print("\nRecherche d'exemples de liens dans Home.py :")
        if os.path.exists("Home.py"):
            with open("Home.py", 'r', encoding='utf-8') as f:
                for i, line in enumerate(f, 1):
                    if 'st.Page("pages/' in line or "st.Page('pages/" in line:
                        print(f"  Ligne {i}: {line.strip()}")
                        if i > 5:
                            break

if __name__ == "__main__":
    main()