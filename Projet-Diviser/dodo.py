import os
import yaml
import fnmatch
from doit.tools import create_folder

BASE_DIR = os.path.dirname(__file__)
SRC_DIR = os.path.join(BASE_DIR, 'src')
FILTERS_DIR = os.path.join(SRC_DIR, 'filters')
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')
INPUT_DIR = os.path.join(BASE_DIR, 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
CONFIG_FILE = os.path.join(BASE_DIR, 'rapport_config.yaml')

def read_config():
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except Exception as e:
        print(f"Erreur lors de la lecture de {CONFIG_FILE} : {e}")
        return {}

def task_generate():
    """Générer les scripts à partir de rapport_config.yaml."""
    config = read_config()
    fichiers = config.get('fichiers', {})

    # Générer filters/<rapport_type>.py
    for fichier_type, info in fichiers.items():
        for rapport_type, rapport_info in info.get('rapports', {}).items():
            target_file = os.path.join(FILTERS_DIR, f"{rapport_type}.py")
            template_file = os.path.join(TEMPLATES_DIR, 'filter_template.py')
            
            def create_filter_file(target=target_file, rt=rapport_type, info_=rapport_info):
                try:
                    with open(template_file, 'r', encoding='utf-8') as tf:
                        template = tf.read()
                    dialog_param = ", dialog_result" if info_.get('dialog', False) else ""
                    dialog_arg = ", dialog_result" if info_.get('dialog', False) else ""
                    with open(target, 'w', encoding='utf-8') as f:
                        f.write(template.format(
                            type=rt,
                            nom_feuille=info_.get('feuille', rt.capitalize()),
                            dialog_param=dialog_param,
                            dialog_arg=dialog_arg
                        ))
                except Exception as e:
                    print(f"Erreur lors de la création de {target} : {e}")
            
            yield {
                'name': f'filter_{rapport_type}',
                'actions': [create_filter_file],
                'file_dep': [CONFIG_FILE, template_file],
                'targets': [target_file],
            }

def task_process_excel():
    """Traiter les fichiers Excel dans input/."""
    from src.report_generator import generer_rapport_individuel
    from src.config import get_config, DOSSIER_ENTREE, DOSSIER_SORTIE
    import glob
    import os
    config = get_config()
    for fichier_type, fichier_info in config['fichiers'].items():
        for fichier in glob.glob(os.path.join(DOSSIER_ENTREE, fichier_info['nom_fichier'])):
            for rapport_name, rapport_info in fichier_info['rapports'].items():
                if 'feuille' not in rapport_info:
                    print(f"Erreur : Clé 'feuille' manquante pour le rapport {rapport_name} dans {fichier_type}")
                    continue
                try:
                    date_part = os.path.basename(fichier).split(' ')[1].replace('.xlsx', '')
                    output_file = os.path.join(DOSSIER_SORTIE, f"Rapport Journalier {rapport_info['feuille'].strip().upper()} {date_part}.xlsx")
                    # Inclure le nom du fichier dans le nom de la tâche pour éviter les doublons
                    fichier_base = os.path.basename(fichier).replace('.xlsx', '').replace(' ', '_')
                    task_name = f"process_excel:{fichier_type}:{rapport_name}:{fichier_base}"
                except IndexError:
                    print(f"Erreur : Nom de fichier {os.path.basename(fichier)} ne contient pas de date valide")
                    continue
                yield {
                    'name': task_name,
                    'actions': [(generer_rapport_individuel, [fichier_type, fichier, rapport_name, None])],
                    'file_dep': [fichier],
                    'targets': [output_file],
                    'clean': True,
                }