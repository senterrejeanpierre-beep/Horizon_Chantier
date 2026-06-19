from __future__ import annotations

import json
import subprocess
import sys
import os
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

APP_NAME = "Horizon Chantier"


# -----------------------------
# Utils PyInstaller
# -----------------------------
def resource_path(relative_path: str) -> Path:
    """Compatibilité PyInstaller"""
    try:
        base_path = Path(sys._MEIPASS)  # type: ignore
    except Exception:
        base_path = Path(__file__).parent
    return base_path / relative_path


# -----------------------------
# Dossiers sur le Bureau
# -----------------------------
def dossier_base() -> Path:
    d = Path.home() / "Desktop" / APP_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def dossier_chantiers() -> Path:
    d = dossier_base() / "Chantiers"
    d.mkdir(parents=True, exist_ok=True)
    return d


def ouvrir_dossier(path: Path) -> None:
    try:
        if sys.platform.startswith("darwin"):
            subprocess.run(["open", str(path)], check=False)
        elif os.name == "nt":
            os.startfile(str(path))
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        messagebox.showerror("Erreur", f"Impossible d’ouvrir le dossier :\n{path}")


def nom_fichier_chantier(nom_chantier: str, client: str) -> str:
    # nom propre pour un fichier (sans caractères bizarres)
    def clean(s: str) -> str:
        s = s.strip().replace(" ", "_")
        keep = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-"
        return "".join(c for c in s if c in keep) or "chantier"

    date = datetime.now().strftime("%Y-%m-%d")
    return f"{date}_{clean(nom_chantier)}_{clean(client)}.json"


def enregistrer_chantier(data: dict) -> Path:
    fichier = dossier_chantiers() / nom_fichier_chantier(data["nom"], data["client"])
    with open(fichier, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return fichier


def lire_chantiers() -> list[dict]:
    chantiers = []
    for p in sorted(dossier_chantiers().glob("*.json")):
        try:
            with open(p, "r", encoding="utf-8") as f:
                d = json.load(f)
            d["_fichier"] = str(p)
            chantiers.append(d)
        except Exception:
            # si un fichier est abîmé, on l’ignore
            pass
    return chantiers


# -----------------------------
# Application
# -----------------------------
class HorizonChantierApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1100x650")
        self.minsize(900, 550)

        self._build_ui()
        self.refresh_liste()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)

        # TITRE
        ttk.Label(root, text=APP_NAME, font=("Helvetica", 18, "bold")).pack(anchor="w")

        # BARRE BOUTONS
        bar = ttk.Frame(root)
        bar.pack(fill="x", pady=(12, 10))

        ttk.Button(bar, text="➕ Nouveau chantier", command=self.ouvrir_fenetre_nouveau).pack(side="left", padx=(0, 10))
        ttk.Button(bar, text="📂 Ouvrir dossier Chantiers", command=lambda: ouvrir_dossier(dossier_chantiers())).pack(side="left")

        info = ttk.Label(root, text=f"Emplacement: {dossier_chantiers()}", foreground="#444")
        info.pack(anchor="w", pady=(0, 12))

        # LISTE DES CHANTIERS
        cols = ("nom", "client", "etat", "avancement", "date_debut", "date_fin")
        self.table = ttk.Treeview(root, columns=cols, show="headings", height=18)

        self.table.heading("nom", text="Chantier")
        self.table.heading("client", text="Client")
        self.table.heading("etat", text="État")
        self.table.heading("avancement", text="Avancement")
        self.table.heading("date_debut", text="Début")
        self.table.heading("date_fin", text="Fin prévue")

        self.table.column("nom", width=260)
        self.table.column("client", width=220)
        self.table.column("etat", width=120)
        self.table.column("avancement", width=120)
        self.table.column("date_debut", width=120)
        self.table.column("date_fin", width=120)

        self.table.pack(fill="both", expand=True)

        bas = ttk.Frame(root)
        bas.pack(fill="x", pady=(10, 0))

        ttk.Button(bas, text="🔄 Rafraîchir", command=self.refresh_liste).pack(side="left")

    def refresh_liste(self) -> None:
        for item in self.table.get_children():
            self.table.delete(item)

        for c in lire_chantiers():
            self.table.insert(
                "", "end",
                values=(
                    c.get("nom", ""),
                    c.get("client", ""),
                    c.get("etat", "Devis"),
                    f'{c.get("avancement", 0)}%',
                    c.get("date_debut", ""),
                    c.get("date_fin", ""),
                ),
            )

    def ouvrir_fenetre_nouveau(self) -> None:
        win = tk.Toplevel(self)
        win.title("Nouveau chantier")
        win.geometry("650x360")
        win.transient(self)

        frame = ttk.Frame(win, padding=18)
        frame.pack(fill="both", expand=True)

        # Champs
        def row(label: str, var: tk.StringVar, r: int):
            ttk.Label(frame, text=label).grid(row=r, column=0, sticky="w", padx=(0, 10), pady=8)
            ttk.Entry(frame, textvariable=var, width=60).grid(row=r, column=1, sticky="ew", pady=8)

        self.var_nom = tk.StringVar()
        self.var_client = tk.StringVar()
        self.var_adresse = tk.StringVar()
        self.var_type = tk.StringVar()
        self.var_debut = tk.StringVar()
        self.var_fin = tk.StringVar()

        row("Nom du chantier", self.var_nom, 0)
        row("Client", self.var_client, 1)
        row("Adresse", self.var_adresse, 2)
        row("Type de travaux", self.var_type, 3)
        row("Date début (ex: 2026-01-23)", self.var_debut, 4)
        row("Date fin prévue (ex: 2026-02-10)", self.var_fin, 5)

        frame.columnconfigure(1, weight=1)

             # Boutons
        btns = ttk.Frame(frame)
        btns.grid(row=6, column=0, columnspan=2, sticky="w", pady=(18, 0))

        ttk.Button(
            btns,
            text="💾 Enregistrer",
            command=lambda: self.enregistrer_depuis_fenetre(win)
        ).pack(side="left", padx=(0, 10))
        ttk.Button(btns, text="Annuler", command=win.destroy).pack(side="left")

    def enregistrer_depuis_fenetre(self, win: tk.Toplevel) -> None:
        nom = self.var_nom.get().strip()
        client = self.var_client.get().strip()

        if not nom:
            messagebox.showwarning("Manque une info", "Écris le nom du chantier.")
            return
        if not client:
            messagebox.showwarning("Manque une info", "Écris le client.")
            return

        data = {
            "nom": nom,
            "client": client,
            "adresse": self.var_adresse.get().strip(),
            "type_travaux": self.var_type.get().strip(),
            "date_debut": self.var_debut.get().strip(),
            "date_fin": self.var_fin.get().strip(),
            "etat": "Devis",
            "avancement": 0,
            "materiaux": [],
            "main_oeuvre": [],
            "frais": [],
            "taches": [],
        }

        fichier = enregistrer_chantier(data)
        messagebox.showinfo("OK", f"Chantier enregistré ✅\n\n{fichier}")
        win.destroy()
        self.refresh_liste()


if __name__ == "__main__":
    app = HorizonChantierApp()
    app.mainloop()

    
