```python
import pandas as pd
from .config import CONFIGS_EXCEL

def charger_fichier_excel(chemin_fichier, nom_feuille):
    try:
        df = pd.read_excel(chemin_fichier, sheet_name=nom_feuille, engine='openpyxl')
        colonnes_attendues = CONFIGS_EXCEL.get(nom_feuille, {}).get('colonnes', [])
        if not all(col in df.columns for col in colonnes_attendues):
            print(f"Colonnes manquantes dans {{chemin_fichier}}")
            return None
        return df
    except Exception as e:
        print(f"Erreur lors du chargement de {{chemin_fichier}}: {{e}}")
        return None

def filtrer_donnees(df, fichier_type, rapport_type):
    if df is None or df.empty:
        return None
    
    try:
        {normalize_columns}
        
        {filtre_cases}
        
        else:
            print(f"Filtre inconnu : {{fichier_type}}/{{rapport_type}}")
            return None
    except KeyError as e:
        print(f"Colonne manquante dans le DataFrame : {{e}}")
        return None
    except Exception as e:
        print(f"Erreur lors du filtrage des données : {{e}}")
        return None
```