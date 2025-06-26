# Template file for doit: placeholders like {import_functions} are replaced during generation
from tkinter import ttk, filedialog, messagebox
import tkinter as tk
import os
from .report_generator import {import_functions}
from .config import DOSSIER_SORTIE

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Générateur de Rapports")
        self.geometry("800x600")
        self.chemins_fichiers = {}
        self.creer_gui()

    def creer_gui(self):
        frame_principal = ttk.Frame(self, padding="10")
        frame_principal.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Button(frame_principal, text="Sélectionner Fichiers Excel", command=self.selectionner_fichiers).grid(row=0, column=0, pady=5)
        self.label_fichiers = ttk.Label(frame_principal, text="Aucun fichier sélectionné")
        self.label_fichiers.grid(row=1, column=0, pady=5)

        {buttons}

        ttk.Button(frame_principal, text="Ouvrir Dossier de Sortie", command=self.ouvrir_dossier_sortie).grid(row={row_ouvrir_dossier}, column=0, pady=5)

        self.log_text = tk.Text(frame_principal, height=10, width=80)
        self.log_text.grid(row={row_log}, column=0, pady=5)
        scrollbar = ttk.Scrollbar(frame_principal, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row={row_log}, column=1, sticky=(tk.N, tk.S))
        self.log_text['yscrollcommand'] = scrollbar.set

    def selectionner_fichiers(self):
        fichiers = filedialog.askopenfilenames(filetypes=[("Fichiers Excel", "*.xlsx")])
        if fichiers:
            from .config import get_fichier_type
            for fichier in fichiers:
                fichier_type = get_fichier_type(os.path.basename(fichier))
                if fichier_type:
                    self.chemins_fichiers[fichier_type] = fichier
                else:
                    messagebox.showwarning("Avertissement", f"Fichier non reconnu : {os.path.basename(fichier)}")
            self.label_fichiers.config(text=f"Fichiers sélectionnés : {', '.join([os.path.basename(f) for f in fichiers])}")
            self.ajouter_log(f"Fichiers sélectionnés : {', '.join([os.path.basename(f) for f in fichiers])}")
        else:
            self.ajouter_log("Aucun fichier sélectionné.")

    def ajouter_log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)

    def ouvrir_dossier_sortie(self):
        import subprocess
        try:
            subprocess.Popen(f'explorer "{DOSSIER_SORTIE}"')
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le dossier : {e}")
            self.ajouter_log(f"Erreur : {e}")

    {methods}

if __name__ == "__main__":
    app = Application()
    app.mainloop()