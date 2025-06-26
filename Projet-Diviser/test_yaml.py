import yaml
import os

config_file = os.path.join(os.path.dirname(__file__), 'rapport_config.yaml')  # Dans Projet Diviser/

try:
    with open(config_file, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
        print(config)
except yaml.YAMLError as e:
    print(f"Erreur YAML dans {config_file} : {e}")
except FileNotFoundError:
    print(f"Fichier introuvable : {config_file}")
except Exception as e:
    print(f"Erreur inattendue : {e}")