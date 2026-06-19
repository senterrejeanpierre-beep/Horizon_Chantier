import pandas as pd
from datetime import datetime

OUT = f"PRIX_REVIENT_COURT_{datetime.now().strftime('%Y_%m_%d_%H%M')}.xlsx"

COLS = [
    "Poste", "Désignation", "Type", "Sous-poste",
    "Unité", "Quantité", "PU net", "Marge %", "PU vente", "Total"
]

def ajouter_poste(poste, designation,
                  matieres, mo,
                  marge_matiere=10.0, marge_mo=10.0):
    """
    Crée 3 lignes Matière + 3 lignes Main-d’œuvre + 1 ligne TOTAL POSTE
    """
    rows = []

    # 3 lignes Matière
    for i, lib in enumerate(matieres):
        rows.append([
            poste if i == 0 else "",
            designation if i == 0 else "",
            "Matière",
            lib,
            "forfait",
            0.0,
            0.0,
            marge_matiere,
            None,
            None
        ])

    # 3 lignes Main-d’œuvre
    for lib in mo:
        rows.append([
            "",
            "",
            "Main-d’œuvre",
            lib,
            "h",
            0.0,
            0.0,
            marge_mo,
            None,
            None
        ])

    # Ligne TOTAL POSTE (sera calculée après)
    rows.append(["", "", "TOTAL POSTE", "", "", "", "", "", "", None])

    return rows

# -------------------------
# 🧱 Construire le tableau
# -------------------------
data = []

# Poste 3.01
data += ajouter_poste(
    "3.01",
    "Container décombre 10 m³",
    matieres=["Container", "Transport", "Petite fourniture"],
    mo=["Démolition", "Étançonnement", "Calepinage"],
    marge_matiere=10,
    marge_mo=10
)

# 👉 Ajoute d'autres postes ici si besoin
# data += ajouter_poste("3.02", "Autre poste", [...], [...], 10, 10)

df = pd.DataFrame(data, columns=COLS)

# -------------------------
# 🔒 Sécuriser les nombres
# -------------------------
for c in ["Quantité", "PU net", "Marge %"]:
    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

# -------------------------
# 🧮 Calculs ligne à ligne
# -------------------------
df["PU vente"] = df["PU net"] * (1 + df["Marge %"] / 100.0)
df["Total"] = df["Quantité"] * df["PU vente"]

# -------------------------
# 🧮 TOTAL POSTE (somme des 6 lignes au-dessus)
# -------------------------
idx_total = df.index[df["Type"] == "TOTAL POSTE"].tolist()
for i in idx_total:
    df.at[i, "Total"] = df.loc[i-6:i-1, "Total"].sum()

# -------------------------
# 💾 Écrire l’Excel
# -------------------------
df.to_excel(OUT, index=False)

print(f"✅ Fichier créé : {OUT}")
print("👉 Renseigne Quantité, PU net, Marge %. Les totaux se calculent automatiquement.")
