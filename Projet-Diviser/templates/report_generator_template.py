import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Border, Side, Alignment
from datetime import datetime
import re
import tkinter.messagebox as messagebox

from .config import get_rapport_config, DOSSIER_SORTIE
from .data_processing import charger_fichier_excel, filtrer_donnees

{imports_filters}

def extraire_date(nom_fichier, df):
    try:
        date_str = re.search(r'(\d{2}[-/]\d{2}[-/]\d{2,4})', nom_fichier)
        if date_str:
            date = datetime.strptime(date_str.group(1), '%d-%m-%y')
            return date
        if 'Date d'Affectation' in df.columns:
            date_str = df['Date d'Affectation'].iloc[0]
            return pd.to_datetime(date_str, dayfirst=True).to_pydatetime()
        return None
    except Exception as e:
        print(f"Erreur lors de l'extraction de la date : {e}")
        return None

def configurer_styles():
    return {
        'titre': Font(name='Arial', size=12, bold=True),
        'entete': Font(name='Arial', size=10, bold=True),
        'cellule': Font(name='Arial', size=10),
        'bordure': Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        ),
        'align_center': Alignment(horizontal='center', vertical='center'),
        'align_left': Alignment(horizontal='left', vertical='center')
    }

def configurer_impression(ws):
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToHeight = False
    ws.page_setup.fitToWidth = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.75
    ws.page_margins.bottom = 0.75

def generer_rapport_combine(fichier_type, chemin_fichier):
    try:
        config = get_rapport_config(fichier_type)
        if not config:
            messagebox.showerror("Erreur", f"Configuration introuvable pour {fichier_type}")
            return None

        df = charger_fichier_excel(chemin_fichier, config['feuille'])
        if df is None:
            return None

        date_rapport = extraire_date(chemin_fichier, df)
        if date_rapport is None:
            messagebox.showerror("Erreur", "Impossible d'extraire la date du fichier")
            return None

        nom_fichier_sortie = f"Rapport Journalier COMBINE {date_rapport.strftime('%d-%m-%Y')}.xlsx"
        chemin_sortie = os.path.join(DOSSIER_SORTIE, nom_fichier_sortie)
        wb = Workbook()
        wb.remove(wb.active)
        styles = configurer_styles()

        {combine_blocks}

        try:
            wb.save(chemin_sortie)
            return nom_fichier_sortie
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'enregistrement du rapport : {e}")
            return None
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur inattendue : {e}")
        return None

{individual_functions}