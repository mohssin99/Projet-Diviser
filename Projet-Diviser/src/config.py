import os
import yaml
import fnmatch

DOSSIER_ENTREE = os.path.join(os.path.dirname(__file__), '..', 'input')
DOSSIER_SORTIE = os.path.join(os.path.dirname(__file__), '..', 'output')

CONFIG_FILE = os.path.join(os.path.dirname(__file__), '..', 'rapport_config.yaml')

def get_config():
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except FileNotFoundError:
        print(f"Fichier de configuration introuvable : {CONFIG_FILE}")
        return {}
    except yaml.YAMLError as e:
        print(f"Erreur dans le fichier YAML {CONFIG_FILE} : {e}")
        return {}

CONFIGS_EXCEL = {
    'SuiviVehicule': {
        'colonnes': [
            'Zone', "Date d'Affectation", 'Shift', 'Secteur',
            'Equipment', 'Genre', 'Num de parc', 'Matricule'
        ]
    }
}

def get_fichier_type(nom_fichier):
    config = get_config()
    for fichier_type, info in config.get('fichiers', {}).items():
        if fnmatch.fnmatch(nom_fichier, info['nom_fichier']):
            return fichier_type
    return None

def get_rapport_config(fichier_type):
    config = get_config()
    return config.get('fichiers', {}).get(fichier_type, {})

def get_rapports():
    config = get_config()
    return config.get('fichiers', {}).get('SuiviVehicule', {}).get('rapports', {})