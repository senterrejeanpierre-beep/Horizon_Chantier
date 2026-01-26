import os, sys, shutil

def resource_path(*parts):
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base, *parts)

def user_root():
    return os.path.join(os.path.expanduser("~"), "Documents", "Devis")

def ensure_user_folders_and_data():
    root = user_root()
    data_dir = os.path.join(root, "DATA")
    sorties_dir = os.path.join(root, "SORTIES")

    os.makedirs(data_dir, exist_ok=True)
    os.makedirs(sorties_dir, exist_ok=True)

    for name in ("compteur.json", "historique.json"):
        src = resource_path("DATA", name)
        dst = os.path.join(data_dir, name)
        if not os.path.exists(dst) and os.path.exists(src):
            shutil.copy2(src, dst)

    return data_dir, sorties_dir
