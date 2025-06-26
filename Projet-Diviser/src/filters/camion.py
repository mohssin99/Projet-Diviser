from ..utils import creer_rapport_journalier

def generer_rapport_camion_feuille(ws, df_camion, date_rapport, styles):
    creer_rapport_journalier(ws, df_camion, date_rapport, "Camion", styles)