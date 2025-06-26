import argparse
import os
import tkinter as tk
from .gui import Application

def main():
    parser = argparse.ArgumentParser(description="Générateur de rapports Excel")
    parser.add_argument('--input', type=str, help="Chemin du fichier Excel à traiter")
    args = parser.parse_args()

    app = Application()
    
    if args.input:
        try:
            if os.path.isfile(args.input):
                from .config import get_fichier_type
                fichier_type = get_fichier_type(os.path.basename(args.input))
                if fichier_type:
                    app.chemins_fichiers[fichier_type] = args.input
                    app.label_fichiers.config(text=f"1 fichier sélectionné pour {fichier_type}")
                    app.ajouter_log(f"Ajouté {os.path.basename(args.input)} pour {fichier_type}")
                else:
                    app.ajouter_log(f"Fichier {os.path.basename(args.input)} non reconnu")
            else:
                app.ajouter_log(f"Fichier introuvable : {args.input}")
        except Exception as e:
            app.ajouter_log(f"Erreur lors du chargement du fichier {args.input} : {e}")
    
    print("Démarrage de l'application...")
    app.mainloop()

if __name__ == "__main__":
    main()