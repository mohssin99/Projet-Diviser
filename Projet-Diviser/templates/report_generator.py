```python
import os
import pandas as pd
from openpyxl import Workbook
from datetime import datetime
from .config import DOSSIER_SORTIE, get_rapport_config
from .data_processing import charger_fichier_excel, filtrer_donnees
from .utils import configurer_styles, creer_rapport_journalier
{imports_filters}

from tkinter import messagebox

def extraire_date(chemin_fichier, df):
    try:
        date_str = df['Date'].iloc[0] if 'Date' in df.columns else df["Date d'Affectation"].iloc[0]
        date_rapport = pd.to_datetime(date_str, format='%m/%d/%y').date()
    except (ValueError, KeyError, IndexError):
        nom_fichier = os.path.basename(chemin_fichier)
        try:
            date_str = nom_fichier.split(' ')[1].split('.xlsx')[0]
            date_rapport = datetime.strptime(date_str, '%d-%m-%Y').date()
        except (IndexError, ValueError):
            messagebox.showerror("Erreur", "Impossible d'extraire la date du fichier.")
            return None
    return date_rapport

def generer_rapport_combine(fichier_type, chemin_fichier, root):
    config = get_rapport_config(fichier_type)
    df = charger_fichier_excel(chemin_fichier, config['feuille'])
    if df is None:
        messagebox.showerror("Erreur", "Erreur lors du chargement du fichier.")
        return None

    date_rapport = extraire_date(chemin_fichier, df)
    if date_rapport is None:
        return None

    nom_fichier_sortie = f"Rapport Journalier {{fichier_type}} {{date_rapport.strftime('%d-%m-%Y')}}.xlsx"
    chemin_sortie = os.path.join(DOSSIER_SORTIE, nom_fichier_sortie)
    wb = Workbook()
    styles = configurer_styles()

    {combine_blocks}

    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    try:
        wb.save(chemin_sortie)
        return nom_fichier_sortie
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de l'enregistrement du rapport : {{str(e)}}")
        return None

def configurer_impression(ws):
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.75
    ws.page_margins.bottom = 0.75

{individual_functions}
```