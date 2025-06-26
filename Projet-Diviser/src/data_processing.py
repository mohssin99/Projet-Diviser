import pandas as pd
import re
from datetime import datetime
from .config import get_config

def charger_fichier_excel(chemin_fichier, nom_feuille):
    try:
        return pd.read_excel(chemin_fichier, sheet_name=nom_feuille)
    except Exception as e:
        print(f"Erreur lors du chargement du fichier Excel {chemin_fichier} : {e}")
        return None

def extraire_date(chemin_fichier, df=None):
    try:
        date_pattern = r'(\d{2}-\d{2}-\d{4})'
        match = re.search(date_pattern, chemin_fichier)
        if match:
            return datetime.strptime(match.group(1), '%d-%m-%Y')
        if df is not None and "Date d'Affectation" in df.columns:
            date_str = df["Date d'Affectation"].iloc[0]
            return pd.to_datetime(date_str, dayfirst=True)
        print(f"Impossible d'extraire la date du fichier {chemin_fichier}")
        return None
    except Exception as e:
        print(f"Erreur lors de l'extraction de la date : {e}")
        return None

def filtrer_donnees(df, fichier_type, rapport_type):
    if df is None or df.empty:
        return None
    try:
        config = get_config()
        if fichier_type not in config['fichiers'] or rapport_type not in config['fichiers'][fichier_type]['rapports']:
            print(f"Configuration ou rapport inconnu : {rapport_type} pour {fichier_type}")
            return None

        if 'Genre' in df.columns:
            df['Genre'] = df['Genre'].str.upper().str.strip()
            filtre = config['fichiers'][fichier_type]['rapports'][rapport_type]['filtre']
            if isinstance(filtre, list):
                return df[df['Genre'].isin(filtre)]
            return df[df['Genre'] == filtre]
        return df
    except KeyError as e:
        print(f"Colonne manquante dans le DataFrame : {e}")
        return None
    except Exception as e:
        print(f"Erreur lors du filtrage des données : {e}")
        return None