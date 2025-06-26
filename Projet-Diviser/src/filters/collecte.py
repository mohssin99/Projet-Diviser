
from ..utils import creer_rapport_journalier

def generer_rapport_collecte_feuille(ws, df_collecte, date_rapport, styles):
    creer_rapport_journalier(ws, df_collecte, date_rapport, "Collecte", styles)