import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from .report_generator import generer_rapport_combine, generer_rapport_individuel
from .config import get_config, DOSSIER_ENTREE

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Générateur de rapports")
        self.geometry("800x600")
        self.chemins_fichiers = {}
        self.config = get_config()
        self.creer_interface()

    def creer_interface(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(padx=10, pady=10, fill='both', expand=True)

        self.label_fichiers = ttk.Label(main_frame, text="Aucun fichier sélectionné")
        self.label_fichiers.pack(anchor='w')

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill='both', expand=True, pady=5)

        for fichier_type, fichier_info in self.config['fichiers'].items():
            tab = ttk.Frame(notebook)
            notebook.add(tab, text=fichier_type)

            ttk.Button(tab, text=f"Sélectionner Fichier {fichier_type}", command=lambda ft=fichier_type: self.selectionner_fichier(ft)).pack(anchor='w', pady=2)

            ttk.Label(tab, text=f"Rapports disponibles pour {fichier_type} :").pack(anchor='w', pady=5)
            for rapport_type, rapport_info in fichier_info['rapports'].items():
                ttk.Button(tab, text=f"Générer Rapport {rapport_info['feuille']}", 
                           command=lambda ft=fichier_type, rt=rapport_type: self.generer_rapport_individuel(ft, rt)).pack(anchor='w', pady=2)

        ttk.Button(main_frame, text="Générer Rapport Combiné", command=self.generer_rapport_combine).pack(anchor='w', pady=10)

        self.log_text = tk.Text(main_frame, height=10, width=80)
        self.log_text.pack(fill='x', pady=5)
        self.log_text.config(state='disabled')

    def selectionner_fichier(self, fichier_type):
        chemin_fichier = filedialog.askopenfilename(initialdir=DOSSIER_ENTREE, 
                                                   filetypes=[("Fichiers Excel", "*.xlsx *.xls")])
        if chemin_fichier:
            self.chemins_fichiers[fichier_type] = chemin_fichier
            self.label_fichiers.config(text=f"Fichiers sélectionnés : {', '.join([os.path.basename(f) for f in self.chemins_fichiers.values()])}")
            self.ajouter_log(f"Ajouté {os.path.basename(chemin_fichier)} pour {fichier_type}")

    def generer_rapport_individuel(self, fichier_type, rapport_type):
        if fichier_type not in self.chemins_fichiers:
            messagebox.showerror("Erreur", f"Aucun fichier sélectionné pour {fichier_type}")
            return
        chemin_fichier = self.chemins_fichiers[fichier_type]
        resultat = generer_rapport_individuel(fichier_type, chemin_fichier, rapport_type, self)
        if resultat:
            self.ajouter_log(f"Rapport généré : {resultat}")
            messagebox.showinfo("Succès", f"Rapport généré : {resultat}")
        else:
            self.ajouter_log(f"Échec de la génération du rapport {rapport_type} pour {fichier_type}")

    def generer_rapport_combine(self):
        if not self.chemins_fichiers:
            messagebox.showerror("Erreur", "Aucun fichier sélectionné")
            return
        resultat = generer_rapport_combine(self.chemins_fichiers, self)
        if resultat:
            self.ajouter_log(f"Rapport combiné généré : {resultat}")
            messagebox.showinfo("Succès", f"Rapport combiné généré : {resultat}")
        else:
            self.ajouter_log("Échec de la génération du rapport combiné")

    def ajouter_log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)