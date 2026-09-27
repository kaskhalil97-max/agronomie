# Automatisation CONTROLE_CONSIGNE_PRELEVEMENT

## Prérequis (une seule fois)
Le script utilise Python 3 et la librairie `openpyxl` pour lire les fichiers Excel. Si vous avez ce message au lancement :
```
ModuleNotFoundError: No module named 'openpyxl'
```
installez-la avec :
```
pip install openpyxl
```
(sous Windows, si `pip` seul ne fonctionne pas : `py -m pip install openpyxl` ou `python -m pip install openpyxl`)

## Fichiers
- `ods_reader.py` : lit un .ods (valeurs + formules) sans dépendance externe (odfpy indisponible dans cet environnement).
- `ods_writer.py` : modifie des cellules précises d'un .ods en place (gère l'éclatement des `number-rows/columns-repeated`), refuse d'écraser une formule.
- `automate.py` : le script principal. Contient `CONFIG` (mapping commune -> colonnes sources) et `fill_month(avant, nord, sud, sortie, annee, mois)`.
- `validate.py` : compare un fichier généré à un fichier de référence (utilisé pour valider mars 2026).

## Comment ça marche
1. Pour chaque commune, le script repère dans `CONTROLE_CONSIGNE_PRELEVEMENT.ods` la ligne du mois demandé dans le grand tableau historique.
2. Il regarde si la cellule à remplir est une formule simple (`=[.B65]`, `=[.D47]+[.C47]`, ...). Si oui, il **remonte automatiquement** jusqu'à la vraie cellule de saisie (mini-tableau "année en cours"), au lieu d'écraser une formule.
3. Il va chercher la valeur source dans l'onglet `Data` / `Data_C1_R0` des fichiers "Volume Production Mensuelle" Nord/Sud (le flux brut de télégestion), en utilisant exactement les mêmes colonnes que les formules déjà présentes dans les onglets `PROD ... MENSUEL`.
4. Il écrit uniquement les cellules de saisie manuelle. Toutes les formules (totaux annuels, ratios Vp/Va, limites DUP) se recalculent **automatiquement à l'ouverture du fichier dans LibreOffice/Excel** — le script n'a pas besoin de les toucher.

## Validation faite
`fill_month('avant.ods', 'nord.xlsx', 'sud.xlsx', 'reproduced.ods', 2026, 3)` reproduit `apres.ods` :
- **13 communes sur 13 : correspondance à 100% sur toutes les cellules de saisie manuelle.**
- BEZIERS : correspondance à 100% sur mars ; l'unique écart observé pendant les tests concernait une correction rétroactive de février dans `apres.ods` (hors périmètre : on ne corrige que le mois demandé).
- LA BAUME - SERVIAN : le script produit 33 576 m³ (= `Data!G` du mois, Forage 2 la Baume) — confirmé correct ; le fichier `apres.ods` fourni initialement contenait une erreur de saisie manuelle (975) sur ce point.

## Pour lancer un nouveau mois (le plus simple)
Dans un terminal, dans le dossier contenant les 3 fichiers :
```
python3 automate.py CONTROLE_CONSIGNE_PRELEVEMENT.ods Volume_Nord.xlsx Volume_Sud.xlsx
```
Le script devine tout seul le mois manquant (celui qui suit le dernier mois déjà rempli) et crée un fichier `CONTROLE_CONSIGNE_PRELEVEMENT_rempli.ods` à côté. Il ne reste plus qu'à l'ouvrir dans LibreOffice Calc (il recalcule tout automatiquement) puis l'enregistrer/renommer comme d'habitude.

Options si besoin :
- `-o mon_fichier.ods` : choisir le nom du fichier de sortie.
- `--annee 2026 --mois 4` : forcer un mois précis au lieu de la détection automatique.

## Limites connues / à surveiller
- SAUVIAN a un 2e champ (toujours à 0 dans nos données) dont la source n'a pas été identifiée — probablement sans usage réel actuellement.
- L'ordre de répartition entre 2 puits d'une même somme (ex. VILLENEUVE, VALRAS) n'a pas d'incidence sur le total affiché (addition), mais peut différer légèrement de la ventilation détaillée d'origine.
- Le script suppose que la ligne du mois cible existe déjà dans le modèle (ce qui est le cas pour les mois à venir de l'année en cours, comme observé dans vos fichiers).
