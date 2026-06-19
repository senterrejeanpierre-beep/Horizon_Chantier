import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from main import (
    _chemin_fichier_chantier,
    _lire_revision_globale_pilotage,
    _lire_synthese_avenants_pilotage,
    _montant_diminution_avenant,
)


def nombre_pilotage(val):
    if val in (None, "", "-", "—"):
        return 0
    try:
        texte_nombre = str(val)
        texte_nombre = texte_nombre.replace("€", "")
        texte_nombre = texte_nombre.replace("−", "-")
        texte_nombre = texte_nombre.replace(" ", "")
        texte_nombre = texte_nombre.replace("\xa0", "")
        texte_nombre = texte_nombre.replace("\u202f", "")
        texte_nombre = texte_nombre.replace(",", ".")
        return float(texte_nombre)
    except:
        return 0


class TestPilotageAvenantsMoins(unittest.TestCase):
    def test_avenants_moins_negatif_ou_positif_reste_une_diminution(self):
        attendu = 20439.02

        self.assertAlmostEqual(
            _montant_diminution_avenant(nombre_pilotage("-20 439,02 €")),
            attendu,
            places=2,
        )
        self.assertAlmostEqual(
            _montant_diminution_avenant(nombre_pilotage("20 439,02 €")),
            attendu,
            places=2,
        )

    def test_pilotage_lit_la_synthese_du_fichier_avenants(self):
        class Cellule:
            def __init__(self, value):
                self.value = value

        class FeuilleAvenants:
            max_row = 62
            valeurs = {
                "E62": Cellule("12 000,00 €"),
                "G62": Cellule("-20 439,02 €"),
            }

            def __getitem__(self, cellule):
                return self.valeurs.get(cellule, Cellule(None))

        avenants_plus, avenants_moins = _lire_synthese_avenants_pilotage(
            FeuilleAvenants(),
            nombre_pilotage,
        )

        self.assertAlmostEqual(avenants_plus, 12000.00, places=2)
        self.assertAlmostEqual(avenants_moins, 20439.02, places=2)

    def test_pilotage_lit_la_revision_globale_en_colonne_e(self):
        class Cellule:
            def __init__(self, value):
                self.value = value

        class FeuilleRevisionGlobale:
            max_row = 62
            valeurs = {
                "E62": Cellule("3 250,50 €"),
            }

            def __getitem__(self, cellule):
                return self.valeurs.get(cellule, Cellule(None))

        revision_globale = _lire_revision_globale_pilotage(
            FeuilleRevisionGlobale(),
            nombre_pilotage,
        )

        self.assertAlmostEqual(revision_globale, 3250.50, places=2)

    def test_ouverture_revision_retrouve_nom_unicode_existant(self):
        with TemporaryDirectory() as tmp:
            dossier = Path(tmp)
            fichier_existant = dossier / "Formule_révision.xlsm"
            fichier_existant.touch()

            chemin = _chemin_fichier_chantier(dossier, "Formule_révision.xlsm")

            self.assertEqual(chemin, fichier_existant)


if __name__ == "__main__":
    unittest.main()
