#!/usr/bin/env python3
"""Classeur de paramétrage du club — saison 2026-2027.

Code couleur des cellules, appliqué partout :
  blanc  — saisie libre
  vert   — saisie assistée (liste déroulante ou sélection de personne)
  gris   — valeur calculée, à ne pas modifier

Chaque table de données réelles est une Table Excel nommée (mêmes noms que
les onglets d'avant la fusion du paramétrage, ou que l'onglet lui-même) :
une ligne tapée juste en dessous de la table reprend automatiquement ses
formules, ses formats et ses listes déroulantes. Les tables ne portent donc
qu'une seule ligne libre en réserve, pas des dizaines.
"""
import json
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from aa_common import *          # noqa: F403
from aa_common import (sheet, section, header, dv_list, dv_range, style_cell,
                       lookup, write_readme, make_table, TAB_SAISIE,
                       TAB_TECHNIQUE)
import aa_data as D

# Chemins relatifs au script, portables d'un environnement à l'autre — pas
# un chemin absolu propre à une seule machine.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.join(SCRIPT_DIR, "..", "classeurs")
os.makedirs(OUTDIR, exist_ok=True)
OUT = os.path.join(OUTDIR, "AA - Parametrage 2026-2027.xlsx")
SAISON = D.SAISON

wb = Workbook()

# ---- personnes du club, chargées tôt : leur nombre fixe la taille des
# plages Personnes et Selection, utilisées par les colonnes de sélection
# d'onglets écrits avant l'onglet Personnes lui-même.
data = json.load(open(os.path.join(SCRIPT_DIR, "personnes.json")))
PERS_ROWS, _seen = [], set()
for nom, pre, _h in data["apneistes_L3"]:
    _seen.add(nom.upper()); _seen.add(pre.upper())
    PERS_ROWS.append((nom, pre, "actif",
                      "Classeur Suivi Apnéistes L3 2025-2026", ""))
for grp, nom in data["encadrants_planning"]:
    if nom.upper() in _seen:
        continue
    _seen.add(nom.upper())
    PERS_ROWS.append(("", nom, "actif",
                      "Planning encadrement 2026-27, groupe %s" % (grp or "—"),
                      "Prénom seul ou surnom : nom de famille à compléter"))
N_REMPLIES = len(PERS_ROWS)
N_CREN = len(D.CRENEAUX)
CODES_AFFECTATIONS = [x[0] for x in D.GROUPES] + [x[0] for x in D.LIGNES_SEANCES]

# ---- positions figées : les données réelles, puis une seule ligne libre
# reprise par la Table Excel dès qu'on tape dedans.
# Exception : la sélection de personne (Selection/Personnes) est partagée par
# cinq onglets (Creneaux_Affectations, Qualifications_Personnes, Inscriptions,
# Responsables_Ligne, Indisponibilites) ; elle garde une capacité fixe pour
# tout le club, plus large qu'une simple ligne libre.
SEL_FIRST, SEL_LAST = 2, 150
N_CLUB = SEL_LAST - SEL_FIRST + 1   # 149 personnes
P_FIRST, P_LAST = 5, 5 + N_CLUB - 1
CR_FIRST, CR_LAST = 5, 5 + N_CREN

R_ID = "Personnes!$A$%d:$A$%d" % (P_FIRST, P_LAST)
R_NOM = "Personnes!$B$%d:$B$%d" % (P_FIRST, P_LAST)
R_PRE = "Personnes!$C$%d:$C$%d" % (P_FIRST, P_LAST)
R_LIB = "Personnes!$J$%d:$J$%d" % (P_FIRST, P_LAST)
R_RANG = "Personnes!$K$%d:$K$%d" % (P_FIRST, P_LAST)
SEL_PERS = "=Selection!$A$%d:$A$%d" % (SEL_FIRST, SEL_LAST)
SEL_CREN = "=Creneaux!$I$%d:$I$%d" % (CR_FIRST, CR_LAST)
R_CR_ID = "Creneaux!$A$%d:$A$%d" % (CR_FIRST, CR_LAST)

MSG_SEL = ("Choisissez la personne dans la liste, triée par prénom. Tapez les "
           "premières lettres du prénom pour y aller plus vite. L'identifiant, "
           "le nom et le prénom se remplissent ensuite tout seuls.")
MSG_CREN = "Choisissez le créneau de piscine dans la liste."


def extract_id(col, row):
    c = "$%s%d" % (col, row)
    return ('=IFERROR(MID({c},FIND("[",{c})+1,FIND("]",{c})-FIND("[",{c})-1),"")'
            .format(c=c))


def bloc_personne(ws, first, last, c_sel, c_id, c_nom, c_pre):
    """Une colonne de sélection, trois colonnes déduites."""
    cs, ci = get_column_letter(c_sel), get_column_letter(c_id)
    for row in range(first, last + 1):
        style_cell(ws.cell(row, c_sel), "select")
        style_cell(ws.cell(row, c_id, extract_id(cs, row)), "calc")
        style_cell(ws.cell(row, c_nom,
                           lookup(R_NOM, "$%s%d" % (ci, row), R_ID)), "calc")
        style_cell(ws.cell(row, c_pre,
                           lookup(R_PRE, "$%s%d" % (ci, row), R_ID)), "calc")
    dv_range(ws, SEL_PERS, "%s%d:%s%d" % (cs, first, cs, last), MSG_SEL)


def remplir(ws, first, last, cols, kind="input", wrap=(), fmt=None):
    for row in range(first, last + 1):
        for col in cols:
            style_cell(ws.cell(row, col), kind, wrap=(col in wrap),
                       fmt=fmt.get(col) if fmt else None)


# ============================================================ Lisez-moi
ws = wb.active
ws.title = "Lisez-moi"
write_readme(ws, "Classeur de paramétrage — Apnée Attitude — saison %s" % SAISON, [
    ("", ""),
    ("titre", "À quoi sert ce classeur"),
    ("p", "Ce classeur porte le paramétrage du suivi des entraînements pour tout "
          "le club : référentiels, cadre d'organisation, personnes, calendrier "
          "perpétuel. Il ne contient aucune donnée de suivi — celles-ci vivent "
          "dans un classeur par ligne d'eau."),
    ("p", "Il sert aussi de contrat de données à la future application : un "
          "onglet correspond à une table, une colonne à un champ, une ligne à un "
          "enregistrement. Sa structure ne doit donc pas être modifiée."),
    ("", ""),
    ("titre", "Trois couleurs de fond, trois natures de cellule"),
    ("li", "Blanc — vous saisissez librement."),
    ("li", "Vert clair — vous saisissez avec une aide : liste déroulante, ou "
           "sélection d'une personne."),
    ("li", "Gris — valeur calculée. N'y touchez pas, elle se met à jour seule."),
    ("p", "C'est le seul code couleur du classeur. Aucune information métier "
          "n'est portée par une couleur : si une information manque quelque "
          "part, il faut lui ajouter une colonne."),
    ("", ""),
    ("titre", "Des tables Excel, pas des plages figées"),
    ("p", "Chaque table de données est une Table Excel nommée, avec une seule "
          "ligne libre en réserve. Tapez dans cette ligne : la table s'agrandit "
          "d'elle-même et reprend formules, formats et listes déroulantes de la "
          "ligne du dessus. Inutile d'insérer une ligne ou de recopier quoi que "
          "ce soit."),
    ("", ""),
    ("titre", "Une personne, plusieurs rôles"),
    ("p", "Dans Inscriptions, une ligne porte un seul rôle pour un seul groupe. "
          "Une personne encadrante en L3 et élève en L4 a donc **deux lignes** : "
          "une (L3, encadrant) et une (L4, élève). Le contrôle de doublon porte "
          "sur le quadruplet saison, personne, groupe, rôle — deux lignes de "
          "rôles différents ne sont donc pas des doublons."),
    ("", ""),
    ("titre", "Sélectionner une personne"),
    ("p", "Vous ne saisissez jamais un identifiant. Une colonne « Sélection par "
          "nom ou prénom » vous propose la liste des personnes du club, "
          "**triée par prénom**, au format « Prénom NOM [identifiant] ». Tapez "
          "les premières lettres du prénom pour y aller plus vite. "
          "L'identifiant, le nom et le prénom se remplissent ensuite tout seuls "
          "dans les trois colonnes grises qui suivent — un coup d'œil suffit à "
          "vérifier que vous avez désigné la bonne personne."),
    ("p", "La liste est alimentée par l'onglet Personnes. Une personne ajoutée "
          "là devient immédiatement sélectionnable partout ailleurs, sans rien "
          "d'autre à faire."),
    ("", ""),
    ("titre", "Les créneaux et lignes de séances"),
    ("li", "Creneaux — les créneaux de piscine : jour, horaires, bassin, nombre "
           "de lignes. Stables sur la saison."),
    ("li", "Creneaux_Affectations — le calendrier perpétuel : quels groupes ou lignes de séances occupent "
           "quel créneau, **à partir de quelle date**, et **avec quel "
           "encadrant**. Un créneau peut accueillir plusieurs groupes ; un "
           "couple créneau / groupe peut avoir plusieurs encadrants, à raison "
           "d'une ligne par encadrant. La répartition fine se règle ensuite "
           "séance par séance dans le classeur de ligne."),
    ("p", "Une évolution en cours d'année n'écrase rien : on ajoute une ligne "
          "avec une nouvelle date d'effet."),
    ("li", "Lignes_Seances — une ligne DNF ou STA, sa capacité et sa période "
           "d'ouverture. Ce n'est pas un groupe d'inscription."),
    ("li", "Regles_Acces_Seances — les groupes autorisés à participer à une "
           "ligne de séances, avec leurs dates de début et de fin d'effet."),
    ("", ""),
    ("titre", "Six règles"),
    ("li", "1. Ne renommez pas les onglets et ne déplacez pas les colonnes."),
    ("li", "2. Une ligne = un enregistrement. Jamais de sous-total ni de ligne "
           "vide au milieu d'un tableau."),
    ("li", "3. Aucune cellule fusionnée dans les zones de données."),
    ("li", "4. Les codes de la première colonne des référentiels sont des "
           "identifiants : ne les modifiez jamais après la première saisie."),
    ("li", "5. L'onglet Controles indique ce qui reste à compléter."),
    ("li", "6. Après toute modification, relancez la synchronisation des "
           "classeurs de ligne."),
    ("", ""),
    ("titre", "Points de vigilance connus"),
    ("li", "Le nombre de répétitions exigé pour valider un objectif en maîtrise "
           "varie selon le niveau : 7 en L1, 5 en L2, 3 en L3. Le L4 relève d'un "
           "critère qualitatif non calculable en l'état."),
    ("li", "L'intervalle maximal de 7 minutes entre départs ne figure dans aucun "
           "plan d'entraînement : il est appliqué à L1, L2 et L3 et reste à "
           "valider."),
    ("li", "Les critères observables des compétences et les niveaux attendus "
           "sont des propositions, à valider collégialement."),
    ("li", "Les prérogatives par qualification sont vides : elles conditionnent "
           "le contrôle de validité des affectations et doivent être renseignées "
           "d'après les textes fédéraux."),
    ("li", "La statique, le DNF et l'apnée en profondeur sont hors périmètre du "
           "MVP : disciplines et groupes déclarés, sans référentiel d'objectifs."),
    ("", ""),
    ("titre", "Version"),
    ("p", "Version 1.3 — 21 août 2026. Généré depuis les documents de cadrage "
          "1.2 à 1.6 (dossier site_apnee/Cadrage_v2)."),
])

# =========================================================== Referentiels
# Regroupe sur un seul onglet les petits référentiels qui ne comportent que
# quelques lignes chacun : les répartir sur sept onglets presque vides les
# rendait difficiles à retrouver. Le sommaire en tête renvoie à chaque section.
ws_ref, r_ref = sheet(wb, "Referentiels", "Référentiels généraux",
    "Huit petits référentiels, regroupés ici plutôt que dispersés sur autant "
    "d'onglets presque vides. Chaque section est une Table Excel à part entière.")
ws_ref.cell(r_ref, 1, "Sommaire").font = SEC_FONT
r_ref += 1
SOMMAIRE_REF = ["Disciplines", "Groupes de niveaux", "Familles d'objectifs",
                "Familles de compétences", "Qualifications", "Bassins",
                "Régimes d'ouverture", "Périodes d'application des régimes"]
SOM_ROW = {}
for titre in SOMMAIRE_REF:
    SOM_ROW[titre] = r_ref
    style_cell(ws_ref.cell(r_ref, 1, titre), "link")
    r_ref += 1
r_ref += 1
REF_SOMMAIRE_END = r_ref - 1
ws_ref.freeze_panes = ws_ref.cell(r_ref, 1)

# ---- Disciplines
r_ref = section(ws_ref, r_ref, "Disciplines",
    "Liste paramétrable. Le code est l'identifiant : ne jamais le modifier.",
    [("Code", 10), ("Libellé", 34), ("Ordre", 8), ("Dans le MVP", 13),
     ("Commentaire", 60)])
DISC_FIRST = r_ref
for d in D.DISCIPLINES:
    for i, v in enumerate(d, 1):
        style_cell(ws_ref.cell(r_ref, i, v), "select" if i == 4 else "input",
                   wrap=(i == 5))
    r_ref += 1
DISC_LAST = DISC_FIRST + len(D.DISCIPLINES)
remplir(ws_ref, r_ref, DISC_LAST, range(1, 6), wrap=(5,))
for row in range(DISC_FIRST, DISC_LAST + 1):
    style_cell(ws_ref.cell(row, 4), "select")
dv_list(ws_ref, ["oui", "non"], "D%d:D%d" % (DISC_FIRST, DISC_LAST))
make_table(ws_ref, "Disciplines", DISC_FIRST - 1, DISC_LAST, 5)
r_ref = DISC_LAST + 2

# ---- Groupes_Niveaux
r_ref = section(ws_ref, r_ref, "Groupes de niveaux",
    "Paramétrables par discipline. « Groupe précédent » rend la progression "
    "explicite d'un niveau au suivant.",
    [("Code", 10), ("Libellé", 24), ("Discipline", 12), ("Rang", 7),
     ("Groupe précédent", 16), ("Dans le MVP", 12), ("Commentaire", 56)])
GR_FIRST = r_ref
for g in D.GROUPES:
    for i, v in enumerate(g, 1):
        style_cell(ws_ref.cell(r_ref, i, v),
                   "select" if i in (3, 5, 6) else "input", wrap=(i == 7))
    r_ref += 1
GR_LAST = GR_FIRST + len(D.GROUPES)
remplir(ws_ref, r_ref, GR_LAST, range(1, 8), wrap=(7,))
for row in range(GR_FIRST, GR_LAST + 1):
    for col in (3, 5, 6):
        style_cell(ws_ref.cell(row, col), "select")
dv_list(ws_ref, [x[0] for x in D.DISCIPLINES], "C%d:C%d" % (GR_FIRST, GR_LAST))
dv_list(ws_ref, [x[0] for x in D.GROUPES], "E%d:E%d" % (GR_FIRST, GR_LAST))
dv_list(ws_ref, ["oui", "non"], "F%d:F%d" % (GR_FIRST, GR_LAST))
make_table(ws_ref, "Groupes_Niveaux", GR_FIRST - 1, GR_LAST, 7)
r_ref = GR_LAST + 2

# ---- Familles_Objectifs
r_ref = section(ws_ref, r_ref, "Familles d'objectifs",
    "Sept familles couvrent les objectifs des quatre niveaux. Chaque groupe "
    "prélève deux seuils sur l'échelle d'une famille.",
    [("Code", 12), ("Libellé", 34), ("Discipline", 12), ("Unité", 24),
     ("Ce qui varie d'un niveau à l'autre", 42), ("Commentaire", 46)])
FO_FIRST = r_ref
for f in D.FAMILLES:
    for i, v in enumerate(f, 1):
        style_cell(ws_ref.cell(r_ref, i, v), "select" if i == 3 else "input",
                   wrap=(i in (5, 6)))
    r_ref += 1
FO_LAST = FO_FIRST + len(D.FAMILLES)
remplir(ws_ref, r_ref, FO_LAST, range(1, 7), wrap=(5, 6))
for row in range(FO_FIRST, FO_LAST + 1):
    style_cell(ws_ref.cell(row, 3), "select")
dv_list(ws_ref, [x[0] for x in D.DISCIPLINES], "C%d:C%d" % (FO_FIRST, FO_LAST))
make_table(ws_ref, "Familles_Objectifs", FO_FIRST - 1, FO_LAST, 6)
r_ref = FO_LAST + 2

# =================================================== Objectifs_Niveaux
ws, r = sheet(wb, "Objectifs_Niveaux", "Objectifs par niveau",
              "Intitulés reproduits littéralement des plans d'entraînement. "
              "Le critère est porté par l'objectif, jamais par le logiciel.")
r = header(ws, r, [("ID", 9), ("Groupe", 8), ("Famille", 11), ("Type", 12),
                   ("Intitulé", 25), ("Distance (m)", 11), ("Durée STA", 10),
                   ("Temps cible", 11), ("Intervalle départ", 12),
                   ("Récupération", 12), ("Nb séries", 9),
                   ("Nb répétitions exigé", 12),
                   ("Intervalle max entre départs", 14),
                   ("Contrainte requise", 26),
                   ("Critère de validation (texte littéral)", 60),
                   ("Source", 46)])
OBJ_FIRST = r
for o in D.OBJECTIFS:
    grp, typ = o[1], o[3]
    nb, itv, contr, txt = D.CRITERES[(grp, typ)]
    for i, v in enumerate(list(o) + [nb, itv, contr, txt, D.SOURCES[grp]], 1):
        style_cell(ws.cell(r, i, v), "select" if i in (2, 3, 4, 14) else "input",
                   wrap=(i in (5, 15, 16)))
    r += 1
OBJ_LAST = OBJ_FIRST + len(D.OBJECTIFS)
remplir(ws, r, OBJ_LAST, range(1, 17), wrap=(5, 15, 16))
for row in range(OBJ_FIRST, OBJ_LAST + 1):
    for col in (2, 3, 4, 14):
        style_cell(ws.cell(row, col), "select")
dv_list(ws, [x[0] for x in D.GROUPES], "B%d:B%d" % (OBJ_FIRST, OBJ_LAST))
dv_list(ws, [x[0] for x in D.FAMILLES], "C%d:C%d" % (OBJ_FIRST, OBJ_LAST))
dv_list(ws, TYPES_OBJ, "D%d:D%d" % (OBJ_FIRST, OBJ_LAST))
dv_list(ws, CONTRAINTES, "N%d:N%d" % (OBJ_FIRST, OBJ_LAST))
make_table(ws, "Objectifs_Niveaux", OBJ_FIRST - 1, OBJ_LAST, 16)
ws.freeze_panes = ws.cell(OBJ_FIRST, 6)

# ---- Familles_Competences
r_ref = section(ws_ref, r_ref, "Familles de compétences",
    "Les 8 compétences réellement évaluées. Critères observables proposés à "
    "partir des plans d'entraînement : à valider.",
    [("Code", 12), ("Libellé", 22), ("Critère observable (proposition)", 72),
     ("Appui documentaire", 56)])
CMP_FIRST = r_ref
for c in D.COMPETENCES:
    for i, v in enumerate(c, 1):
        style_cell(ws_ref.cell(r_ref, i, v), "input", wrap=(i in (3, 4)))
    r_ref += 1
CMP_LAST = CMP_FIRST + len(D.COMPETENCES)
remplir(ws_ref, r_ref, CMP_LAST, range(1, 5), wrap=(3, 4))
make_table(ws_ref, "Familles_Competences", CMP_FIRST - 1, CMP_LAST, 4)
r_ref = CMP_LAST + 2

# ================================================= Competences_Niveaux
ws, r = sheet(wb, "Competences_Niveaux",
              "Niveau attendu par compétence et par groupe",
              "Une seule table pour toutes les lignes : c'est ce qui garantit "
              "la cohérence d'ensemble de la progression.")
r = header(ws, r, [("Compétence", 12), ("Libellé", 22), ("Groupe", 8),
                   ("Niveau attendu", 14), ("Commentaire", 50)])
ATT_FIRST = r
for code, lib, _c, _s in D.COMPETENCES:
    for g in ("L1", "L2", "L3", "L4", "LC"):
        for i, v in enumerate([code, lib, g, D.ATTENDU[code][g], ""], 1):
            style_cell(ws.cell(r, i, v),
                       "select" if i in (1, 3, 4) else "input", wrap=(i == 5))
        r += 1
ATT_LAST = ATT_FIRST + len(D.COMPETENCES) * 5
remplir(ws, r, ATT_LAST, range(1, 6), wrap=(5,))
for row in range(ATT_FIRST, ATT_LAST + 1):
    for col in (1, 3, 4):
        style_cell(ws.cell(row, col), "select")
dv_list(ws, NIVEAUX_COMP, "D%d:D%d" % (ATT_FIRST, ATT_LAST))
dv_list(ws, [x[0] for x in D.COMPETENCES], "A%d:A%d" % (ATT_FIRST, ATT_LAST))
dv_list(ws, [x[0] for x in D.GROUPES], "C%d:C%d" % (ATT_FIRST, ATT_LAST))
make_table(ws, "Competences_Niveaux", ATT_FIRST - 1, ATT_LAST, 5)
ws.freeze_panes = ws.cell(ATT_FIRST, 1)

# ---- Qualifications
r_ref = section(ws_ref, r_ref, "Qualifications",
    "Qualifications d'encadrement et d'apnéiste. Les prérogatives "
    "conditionnent le contrôle de validité des affectations : à renseigner "
    "d'après les textes fédéraux.",
    [("Code", 10), ("Libellé", 44), ("Type", 14), ("Recyclage", 24),
     ("Prérogatives", 46)])
QU_FIRST = r_ref
for code, lib, cat, recy, _st in D.QUALIFS:
    for i, v in enumerate([code, lib, cat, recy, ""], 1):
        style_cell(ws_ref.cell(r_ref, i, v), "select" if i == 3 else "input",
                   wrap=(i in (2, 5)))
    r_ref += 1
QU_LAST = QU_FIRST + len(D.QUALIFS)
remplir(ws_ref, r_ref, QU_LAST, range(1, 6), wrap=(2, 5))
for row in range(QU_FIRST, QU_LAST + 1):
    style_cell(ws_ref.cell(row, 3), "select")
dv_list(ws_ref, ["Encadrement", "Niveau apnée", "Sécurité"],
        "C%d:C%d" % (QU_FIRST, QU_LAST))
make_table(ws_ref, "Qualifications", QU_FIRST - 1, QU_LAST, 5)
r_ref = QU_LAST + 2

# ---- Bassins
r_ref = section(ws_ref, r_ref, "Bassins",
    "La longueur du bassin change la nature des exercices et "
    "l'interprétation d'une performance en dynamique.",
    [("Code", 10), ("Libellé", 26), ("Longueur (m)", 12),
     ("Nb de lignes", 12), ("Commentaire", 60)])
BAS_FIRST = r_ref
for b in D.BASSINS:
    for i, v in enumerate(b, 1):
        style_cell(ws_ref.cell(r_ref, i, v), "input", wrap=(i == 5))
    r_ref += 1
BAS_LAST = BAS_FIRST + len(D.BASSINS)
remplir(ws_ref, r_ref, BAS_LAST, range(1, 6), wrap=(5,))
make_table(ws_ref, "Bassins", BAS_FIRST - 1, BAS_LAST, 5)
r_ref = BAS_LAST + 2

# ---- Regimes
r_ref = section(ws_ref, r_ref, "Régimes d'ouverture",
    "Un régime dérogatoire daté se substitue au régime nominal.",
    [("Code", 14), ("Libellé", 30), ("Commentaire", 66)])
REG_FIRST = r_ref
for g in D.REGIMES:
    for i, v in enumerate(g, 1):
        style_cell(ws_ref.cell(r_ref, i, v), "input", wrap=(i == 3))
    r_ref += 1
REG_LAST = REG_FIRST + len(D.REGIMES)
remplir(ws_ref, r_ref, REG_LAST, range(1, 4), wrap=(3,))
make_table(ws_ref, "Regimes", REG_FIRST - 1, REG_LAST, 3)
r_ref = REG_LAST + 2

# ---- Periodes
r_ref = section(ws_ref, r_ref, "Périodes d'application des régimes",
    "À compléter avec le calendrier scolaire et les jours fériés.",
    [("Régime", 14), ("Date de début", 14), ("Date de fin", 14),
     ("Libellé", 30), ("Commentaire", 30)])
PER_FIRST = r_ref
for p in D.PERIODES:
    for i, v in enumerate(p[:4] + ("",), 1):
        style_cell(ws_ref.cell(r_ref, i, v), "select" if i == 1 else "input",
                   fmt="DD/MM/YYYY" if i in (2, 3) else None)
    r_ref += 1
PER_LAST = PER_FIRST + len(D.PERIODES)
remplir(ws_ref, r_ref, PER_LAST, range(1, 6),
        fmt={2: "DD/MM/YYYY", 3: "DD/MM/YYYY"})
for row in range(PER_FIRST, PER_LAST + 1):
    style_cell(ws_ref.cell(row, 1), "select")
dv_list(ws_ref, [x[0] for x in D.REGIMES], "A%d:A%d" % (PER_FIRST, PER_LAST))
make_table(ws_ref, "Periodes", PER_FIRST - 1, PER_LAST, 5)
r_ref = PER_LAST + 2

# ---- le sommaire renvoie vers les sections effectivement écrites
for titre, first_row in (("Disciplines", DISC_FIRST),
                         ("Groupes de niveaux", GR_FIRST),
                         ("Familles d'objectifs", FO_FIRST),
                         ("Familles de compétences", CMP_FIRST),
                         ("Qualifications", QU_FIRST), ("Bassins", BAS_FIRST),
                         ("Régimes d'ouverture", REG_FIRST),
                         ("Périodes d'application des régimes", PER_FIRST)):
    style_cell(ws_ref.cell(SOM_ROW[titre], 1,
        '=HYPERLINK("#Referentiels!A%d","%s")' % (first_row, titre)), "link")

# largeurs de colonnes communes : un compromis, chaque section a ses propres
# en-têtes juste au-dessus de ses données
for _col, _w in ((1, 14), (2, 34), (3, 20), (4, 22), (5, 26), (6, 48), (7, 56)):
    ws_ref.column_dimensions[get_column_letter(_col)].width = _w

# ============================================================= Creneaux
ws, r = sheet(wb, "Creneaux", "Créneaux de piscine",
              "La structure hebdomadaire récurrente. Les groupes de niveaux et "
              "lignes de séances accueillis sont dans l'onglet "
              "Creneaux_Affectations, avec leur date d'effet.")
r = header(ws, r, [("ID", 10), ("Jour", 11), ("Heure début", 11),
                   ("Heure fin", 11), ("Bassin", 9),
                   ("Nb de lignes ouvertes", 12), ("Régime", 12),
                   ("Actif", 8), ("Libellé (calculé)", 42),
                   ("Commentaire", 50)])
assert r == CR_FIRST, r
for i in range(N_CREN + 1):
    row = CR_FIRST + i
    src = D.CRENEAUX[i] if i < N_CREN else None
    vals = ([src[0], src[2], src[3], src[4], src[5], src[7], src[1], src[8]]
            if src else [""] * 8)
    for col, v in enumerate(vals, 1):
        style_cell(ws.cell(row, col, v),
                   "select" if col in (2, 5, 7, 8) else "input")
    style_cell(ws.cell(row, 9,
        '=IF($A{r}="","",$A{r}&" — "&$B{r}'
        '&IF($C{r}=""," (horaire à saisir)"," "&$C{r}&"-"&$D{r})'
        '&IF($E{r}=""," "," · "&$E{r})&" ["&$A{r}&"]")'.format(r=row)), "calc")
    style_cell(ws.cell(row, 10, src[9] if src else ""), "input", wrap=True)
dv_list(ws, JOURS, "B%d:B%d" % (CR_FIRST, CR_LAST))
dv_list(ws, [x[0] for x in D.BASSINS], "E%d:E%d" % (CR_FIRST, CR_LAST))
dv_list(ws, [x[0] for x in D.REGIMES], "G%d:G%d" % (CR_FIRST, CR_LAST))
dv_list(ws, ["oui", "non"], "H%d:H%d" % (CR_FIRST, CR_LAST))
make_table(ws, "Creneaux", CR_FIRST - 1, CR_LAST, 10)
ws.freeze_panes = ws.cell(CR_FIRST, 2)

# ================================================= Creneaux_Affectations
ws, r = sheet(wb, "Creneaux_Affectations",
              "Groupes et lignes de séances, par créneau",
              "Le calendrier perpétuel. Un créneau peut accueillir plusieurs "
              "groupes ou lignes de séances, et un couple créneau / affectation plusieurs encadrants : "
              "une ligne par encadrant. Pour une ligne de séances, désignez un "
              "unique encadrant référent. La répartition fine se règle ensuite "
              "séance par séance dans le classeur de ligne. La périodicité est "
              "hebdomadaire si elle est laissée vide.")
r = header(ws, r, [("Créneau", 42), ("Créneau (ID)", 11),
                   ("Date d'effet", 13), ("Groupe ou ligne de séance", 22),
                   ("Sélection par nom ou prénom", 34), ("Encadrant (ID)", 12),
                   ("Nom", 20), ("Prénom", 14),
                   ("Nb de lignes allouées", 12), ("Commentaire", 40),
                   ("Périodicité (jours)", 18), ("Contrôle", 30),
                   ("Encadrant référent", 18)])
CG_FIRST = r
N_CG = 5   # créneaux « à confirmer » déduits du planning 2026-27
CG_LAST = CG_FIRST + N_CG
bloc_personne(ws, CG_FIRST, CG_LAST, 5, 6, 7, 8)
for row in range(CG_FIRST, CG_LAST + 1):
    style_cell(ws.cell(row, 1), "select")
    style_cell(ws.cell(row, 2, extract_id("A", row)), "calc")
    style_cell(ws.cell(row, 3), "input", fmt="DD/MM/YYYY")
    style_cell(ws.cell(row, 4), "select")
    style_cell(ws.cell(row, 9), "input")
    style_cell(ws.cell(row, 10), "input", wrap=True)
    style_cell(ws.cell(row, 11), "input")
    style_cell(ws.cell(row, 12,
        '=IF($A{r}="","",'
        'IF($B{r}="","créneau non reconnu",'
        'IF($C{r}="","date d’effet manquante",'
        'IF($D{r}="","groupe manquant",'
        'IF(COUNTIFS($B${f}:$B${l},$B{r},$C${f}:$C${l},$C{r},'
        '$D${f}:$D${l},$D{r},$F${f}:$F${l},$F{r})>1,'
        '"doublon créneau/date/groupe/encadrant",'
        'IF($F{r}="","encadrant à désigner","ok"))))))'
        .format(r=row, f=CG_FIRST, l=CG_LAST)), "calc")
    style_cell(ws.cell(row, 13), "select")
dv_range(ws, SEL_CREN, "A%d:A%d" % (CG_FIRST, CG_LAST), MSG_CREN)
dv_list(ws, CODES_AFFECTATIONS, "D%d:D%d" % (CG_FIRST, CG_LAST))
dv_list(ws, ["oui", "non"], "M%d:M%d" % (CG_FIRST, CG_LAST))
for i_, (cid, grp) in enumerate([("CR-LUN", "L1"), ("CR-MAR", "L3"),
                                 ("CR-MER", "L2"), ("CR-VEN", "L3"),
                                 ("CR-SAM", "L4")]):
    ws.cell(CG_FIRST + i_, 1, "%s — à confirmer [%s]" % (cid, cid))
    ws.cell(CG_FIRST + i_, 4, grp)
    ws.cell(CG_FIRST + i_, 9, 1)
    ws.cell(CG_FIRST + i_, 10, "Déduit du planning 2026-27 : à confirmer, "
                               "à dater, et à affecter")
make_table(ws, "Creneaux_Affectations", CG_FIRST - 1, CG_LAST, 13)
ws.freeze_panes = ws.cell(CG_FIRST, 3)

# ======================================================== Lignes_Seances
ws, r = sheet(wb, "Lignes_Seances", "Lignes de séances",
              "Table datée des lignes DNF, STA ou équivalentes. Leur capacité "
              "ne se déduit pas des affectations de créneau : elle est définie ici.")
r = header(ws, r, [("Code ligne", 16), ("Libellé", 34),
                   ("Nombre de lignes", 16), ("Début d’effet", 13),
                   ("Fin d’effet", 13), ("Actif", 10)])
LSE_FIRST = r
LSE_LAST = LSE_FIRST + 19
for row in range(LSE_FIRST, LSE_LAST + 1):
    style_cell(ws.cell(row, 1), "select")
    style_cell(ws.cell(row, 2), "input")
    style_cell(ws.cell(row, 3), "input")
    style_cell(ws.cell(row, 4), "input", fmt="DD/MM/YYYY")
    style_cell(ws.cell(row, 5), "input", fmt="DD/MM/YYYY")
    style_cell(ws.cell(row, 6), "select")
dv_list(ws, [x[0] for x in D.LIGNES_SEANCES], "A%d:A%d" % (LSE_FIRST, LSE_LAST))
dv_list(ws, ["oui", "non"], "F%d:F%d" % (LSE_FIRST, LSE_LAST))
make_table(ws, "Lignes_Seances", LSE_FIRST - 1, LSE_LAST, 6)
ws.freeze_panes = ws.cell(LSE_FIRST, 3)

# =================================================== Regles_Acces_Seances
ws, r = sheet(wb, "Regles_Acces_Seances", "Règles d'accès aux séances",
              "Table datée des groupes autorisés à participer à chaque ligne de "
              "séances. Elle évite une affectation nominative des participants.")
r = header(ws, r, [("Ligne de séance", 18), ("Groupe autorisé", 18),
                   ("Début d’effet", 13), ("Fin d’effet", 13),
                   ("Actif", 10)])
RAS_FIRST = r
RAS_LAST = RAS_FIRST + 39
for row in range(RAS_FIRST, RAS_LAST + 1):
    style_cell(ws.cell(row, 1), "select")
    style_cell(ws.cell(row, 2), "select")
    style_cell(ws.cell(row, 3), "input", fmt="DD/MM/YYYY")
    style_cell(ws.cell(row, 4), "input", fmt="DD/MM/YYYY")
    style_cell(ws.cell(row, 5), "select")
dv_list(ws, [x[0] for x in D.LIGNES_SEANCES], "A%d:A%d" % (RAS_FIRST, RAS_LAST))
dv_list(ws, [x[0] for x in D.GROUPES], "B%d:B%d" % (RAS_FIRST, RAS_LAST))
dv_list(ws, ["oui", "non"], "E%d:E%d" % (RAS_FIRST, RAS_LAST))
make_table(ws, "Regles_Acces_Seances", RAS_FIRST - 1, RAS_LAST, 5)
ws.freeze_panes = ws.cell(RAS_FIRST, 3)

# ============================================================ Personnes
ws, r = sheet(wb, "Personnes", "Personnes — référentiel unique du club",
              "Élèves et encadrants dans un seul référentiel. Les deux "
              "dernières colonnes alimentent la liste de sélection, triée par "
              "prénom.")
r = header(ws, r, [("ID", 10), ("Nom", 22), ("Prénom", 16), ("Genre", 8),
                   ("Date de naissance", 14), ("Courriel", 30),
                   ("Statut", 22), ("Origine de la donnée", 30),
                   ("Commentaire", 34), ("Libellé (calculé)", 34),
                   ("Rang par prénom (calculé)", 12)])
assert r == P_FIRST, r
for i in range(N_CLUB):
    row = P_FIRST + i
    if i < N_REMPLIES:
        nom, pre, st, org, comm = PERS_ROWS[i]
        vals = ["P%03d" % (i + 1), nom, pre, "", "", "", st, org, comm]
    else:
        vals = [""] * 9
    for col, v in enumerate(vals, 1):
        style_cell(ws.cell(row, col, v),
                   "select" if col in (4, 7) else "input",
                   wrap=(col == 9), fmt="DD/MM/YYYY" if col == 5 else None)
    style_cell(ws.cell(row, 10,
        '=IF($A{r}="","",TRIM(IF($C{r}="",$B{r},$C{r}&" "&$B{r}))&" ["&$A{r}&"]")'
        .format(r=row)), "calc")
    # rang alphabétique sur le prénom ; le second terme rend le rang unique
    # en cas de prénoms identiques. Le prénom est une valeur saisie, donc les
    # lignes vides sont réellement vides : COUNTIF les ignore.
    style_cell(ws.cell(row, 11,
        '=IF($A{r}="","",COUNTIF({pre},"<"&$C{r})+COUNTIF($C${f}:$C{r},$C{r}))'
        .format(r=row, pre=R_PRE, f=P_FIRST)), "calc")
dv_list(ws, GENRES, "D%d:D%d" % (P_FIRST, P_LAST))
dv_list(ws, STATUTS_PERS, "G%d:G%d" % (P_FIRST, P_LAST))
make_table(ws, "Personnes", P_FIRST - 1, P_LAST, 11)
ws.freeze_panes = ws.cell(P_FIRST, 4)

# ============================================= Qualifications_Personnes
ws, r = sheet(wb, "Qualifications_Personnes", "Qualifications détenues",
              "Sélectionnez la personne : l'identifiant, le nom et le prénom "
              "se remplissent tout seuls.")
r = header(ws, r, [("Sélection par nom ou prénom", 34), ("Personne (ID)", 12),
                   ("Nom", 20), ("Prénom", 14), ("Qualification", 13),
                   ("Type (calculé)", 16), ("Libellé (calculé)", 40),
                   ("Date d'obtention", 14), ("Date de recyclage", 14),
                   ("Valide", 9), ("Commentaire", 40)])
QP_FIRST = r
QP_LAST = QP_FIRST      # aucune qualification pré-saisie : une ligne libre
bloc_personne(ws, QP_FIRST, QP_LAST, 1, 2, 3, 4)
R_QU_CODE = "Referentiels!$A$%d:$A$%d" % (QU_FIRST, QU_LAST)
for row in range(QP_FIRST, QP_LAST + 1):
    style_cell(ws.cell(row, 5), "select")
    style_cell(ws.cell(row, 6,
        '=IFERROR(INDEX(Referentiels!$C${f}:$C${l},'
        'MATCH($E{r},{code},0))&"","")'
        .format(r=row, f=QU_FIRST, l=QU_LAST, code=R_QU_CODE)), "calc")
    style_cell(ws.cell(row, 7,
        '=IFERROR(INDEX(Referentiels!$B${f}:$B${l},'
        'MATCH($E{r},{code},0))&"","")'
        .format(r=row, f=QU_FIRST, l=QU_LAST, code=R_QU_CODE)), "calc")
    for col in (8, 9):
        style_cell(ws.cell(row, col), "input", fmt="DD/MM/YYYY")
    style_cell(ws.cell(row, 10), "select")
    style_cell(ws.cell(row, 11), "input", wrap=True)
dv_list(ws, [x[0] for x in D.QUALIFS], "E%d:E%d" % (QP_FIRST, QP_LAST))
dv_list(ws, ["oui", "non"], "J%d:J%d" % (QP_FIRST, QP_LAST))
make_table(ws, "Qualifications_Personnes", QP_FIRST - 1, QP_LAST, 11)
ws.freeze_panes = ws.cell(QP_FIRST, 5)

# ========================================================== Inscriptions
ws, r = sheet(wb, "Inscriptions", "Inscriptions de la saison",
              "La colonne Contrôle signale les doublons et les informations "
              "manquantes.")
r = header(ws, r, [("Saison", 11), ("Sélection par nom ou prénom", 34),
                   ("Personne (ID)", 12), ("Nom", 20), ("Prénom", 14),
                   ("Groupe de niveau", 15), ("Rôle", 18),
                   ("Support de saisie", 15), ("Date de début", 13),
                   ("Date de fin", 13), ("Objectif de la saison", 40),
                   ("Commentaire", 28), ("Contrôle", 20)])
INS_FIRST = r
INS_LAST = INS_FIRST      # aucune inscription pré-saisie : une ligne libre
bloc_personne(ws, INS_FIRST, INS_LAST, 2, 3, 4, 5)
for row in range(INS_FIRST, INS_LAST + 1):
    style_cell(ws.cell(row, 1, SAISON if row == INS_FIRST else ""), "input")
    for col in (6, 7, 8):
        style_cell(ws.cell(row, col), "select")
    for col in (9, 10):
        style_cell(ws.cell(row, col), "input", fmt="DD/MM/YYYY")
    for col in (11, 12):
        style_cell(ws.cell(row, col), "input", wrap=True)
    style_cell(ws.cell(row, 13,
        '=IF($B{r}="","",'
        'IF($C{r}="","personne non reconnue",'
        'IF(ISERROR(SEARCH(TRIM($E{r}&" "&$D{r}),$B{r})),'
        '"libellé incohérent avec l’identifiant",'
        'IF(COUNTIFS($A${f}:$A${l},$A{r},$C${f}:$C${l},$C{r},'
        '$F${f}:$F${l},$F{r},$G${f}:$G${l},$G{r})>1,"doublon",'
        'IF($F{r}="","groupe manquant",'
        'IF($G{r}="","rôle manquant",'
        'IF($H{r}="","support de saisie manquant",'
        'IF(AND($I{r}<>"",$J{r}<>"",$J{r}<$I{r}),'
        '"date de fin antérieure à la date de début","ok"))))))))'
        .format(r=row, f=INS_FIRST, l=INS_LAST)), "calc")
dv_list(ws, [x[0] for x in D.GROUPES], "F%d:F%d" % (INS_FIRST, INS_LAST))
dv_list(ws, ["élève", "encadrant"], "G%d:G%d" % (INS_FIRST, INS_LAST))
dv_list(ws, SUPPORTS, "H%d:H%d" % (INS_FIRST, INS_LAST))
make_table(ws, "Inscriptions", INS_FIRST - 1, INS_LAST, 13)
ws.freeze_panes = ws.cell(INS_FIRST, 6)

# ==================================================== Responsables_Ligne
ws, r = sheet(wb, "Responsables_Ligne", "Responsable de ligne, par groupe",
              "Table datée : le responsable en vigueur est celui de la dernière "
              "date d'effet antérieure à la séance. Un changement en cours de "
              "saison ajoute une ligne, il n'écrase rien.")
r = header(ws, r, [("Groupe de niveau", 15), ("Date d'effet", 13),
                   ("Sélection par nom ou prénom", 34), ("Personne (ID)", 12),
                   ("Nom", 20), ("Prénom", 14), ("Commentaire", 46)])
RES_FIRST = r
RES_LAST = RES_FIRST      # aucun responsable pré-saisi : une ligne libre
bloc_personne(ws, RES_FIRST, RES_LAST, 3, 4, 5, 6)
for row in range(RES_FIRST, RES_LAST + 1):
    style_cell(ws.cell(row, 1), "select")
    style_cell(ws.cell(row, 2), "input", fmt="DD/MM/YYYY")
    style_cell(ws.cell(row, 7), "input", wrap=True)
dv_list(ws, [x[0] for x in D.GROUPES], "A%d:A%d" % (RES_FIRST, RES_LAST))
make_table(ws, "Responsables_Ligne", RES_FIRST - 1, RES_LAST, 7)
ws.freeze_panes = ws.cell(RES_FIRST, 4)

# ====================================================== Indisponibilites
ws, r = sheet(wb, "Indisponibilites", "Exceptions d'indisponibilité",
              "Les exceptions au calendrier perpétuel. Une indisponibilité "
              "couvrant une séance fait apparaître un défaut de couverture.")
r = header(ws, r, [("Sélection par nom ou prénom", 34), ("Personne (ID)", 12),
                   ("Nom", 20), ("Prénom", 14), ("Date de début", 13),
                   ("Date de fin", 13), ("Motif", 16),
                   ("Créneau concerné (facultatif)", 42), ("Commentaire", 40)])
IND_FIRST = r
IND_LAST = IND_FIRST      # aucune indisponibilité pré-saisie : une ligne libre
bloc_personne(ws, IND_FIRST, IND_LAST, 1, 2, 3, 4)
for row in range(IND_FIRST, IND_LAST + 1):
    for col in (5, 6):
        style_cell(ws.cell(row, col), "input", fmt="DD/MM/YYYY")
    for col in (7, 8):
        style_cell(ws.cell(row, col), "select")
    style_cell(ws.cell(row, 9), "input", wrap=True)
dv_list(ws, MOTIFS_INDISPO, "G%d:G%d" % (IND_FIRST, IND_LAST))
dv_range(ws, SEL_CREN, "H%d:H%d" % (IND_FIRST, IND_LAST), MSG_CREN)
make_table(ws, "Indisponibilites", IND_FIRST - 1, IND_LAST, 9)
ws.freeze_panes = ws.cell(IND_FIRST, 5)

# ============================================================= Controles
ws, r = sheet(wb, "Controles", "Contrôles de cohérence",
              "Tout est calculé. Une valeur non nulle sur une ligne « dont … » "
              "demande une action avant la reprise du 7 septembre 2026.")
for L, w in (("A", 58), ("B", 12), ("C", 62)):
    ws.column_dimensions[L].width = w
CTRL_HEADER_ROW = r
r = header(ws, r, [("Contrôle", 58), ("Valeur", 12), ("Signification", 62)])
CTRL = [
    ("Personnes enregistrées",
     '=COUNTIF(Personnes!$A$%d:$A$%d,"<>")' % (P_FIRST, P_LAST),
     "Référentiel unique du club."),
    ("dont sans nom de famille",
     '=COUNTIFS(Personnes!$A$%d:$A$%d,"<>",Personnes!$B$%d:$B$%d,"")'
     % (P_FIRST, P_LAST, P_FIRST, P_LAST),
     "Issues du planning, où seul le prénom figure."),
    ("dont sans courriel",
     '=COUNTIFS(Personnes!$A$%d:$A$%d,"<>",Personnes!$F$%d:$F$%d,"")'
     % (P_FIRST, P_LAST, P_FIRST, P_LAST),
     "Nécessaire pour l'auto-inscription de l'application."),
    ("Inscriptions saisies",
     '=SUMPRODUCT(--(Inscriptions!$C$%d:$C$%d<>""))' % (INS_FIRST, INS_LAST),
     "Une ligne par personne et par groupe de niveau."),
    ("dont en anomalie",
     '=SUMPRODUCT((Inscriptions!$C$%d:$C$%d<>"")*'
     '(Inscriptions!$M$%d:$M$%d<>"ok"))'
     % (INS_FIRST, INS_LAST, INS_FIRST, INS_LAST),
     "Voir la colonne Contrôle de l'onglet Inscriptions."),
    ("Créneaux de piscine actifs",
     '=COUNTIFS(Creneaux!$A$%d:$A$%d,"<>",Creneaux!$H$%d:$H$%d,"oui")'
     % (CR_FIRST, CR_LAST, CR_FIRST, CR_LAST),
     "Structure hebdomadaire récurrente."),
    ("dont sans horaire",
     '=COUNTIFS(Creneaux!$A$%d:$A$%d,"<>",Creneaux!$C$%d:$C$%d,"")'
     % (CR_FIRST, CR_LAST, CR_FIRST, CR_LAST),
     "À saisir avant la reprise."),
    ("Associations créneau / groupe",
     '=SUMPRODUCT(--(Creneaux_Affectations!$B$%d:$B$%d<>""))' % (CG_FIRST, CG_LAST),
     "Un créneau peut accueillir plusieurs groupes."),
    ("dont en anomalie",
     '=SUMPRODUCT((Creneaux_Affectations!$A$%d:$A$%d<>"")*'
     '(Creneaux_Affectations!$L$%d:$L$%d<>"ok")*'
     '(Creneaux_Affectations!$L$%d:$L$%d<>"encadrant à désigner"))'
     % (CG_FIRST, CG_LAST, CG_FIRST, CG_LAST, CG_FIRST, CG_LAST),
     "Date d'effet ou groupe manquant, créneau non reconnu, doublon."),
    ("Créneaux actifs sans aucun groupe",
     '=SUMPRODUCT((Creneaux!$A$%d:$A$%d<>"")*(Creneaux!$H$%d:$H$%d="oui")*'
     '(COUNTIF(Creneaux_Affectations!$B$%d:$B$%d,Creneaux!$A$%d:$A$%d)=0))'
     % (CR_FIRST, CR_LAST, CR_FIRST, CR_LAST, CG_FIRST, CG_LAST,
        CR_FIRST, CR_LAST),
     "Un créneau ouvert doit accueillir au moins un groupe."),
    ("dont sans encadrant désigné",
     '=COUNTIF(Creneaux_Affectations!$L$%d:$L$%d,"encadrant à désigner")'
     % (CG_FIRST, CG_LAST),
     "Une ligne par encadrant : plusieurs encadrants sur un même couple "
     "créneau / groupe sont attendus."),
    ("Responsables de ligne déclarés",
     '=SUMPRODUCT(--(Responsables_Ligne!$D$%d:$D$%d<>""))'
     % (RES_FIRST, RES_LAST),
     "Une ligne par groupe et par changement en cours de saison."),
    ("Groupes du MVP sans responsable déclaré",
     '=SUMPRODUCT((Referentiels!$F$%d:$F$%d="oui")*'
     '(COUNTIF(Responsables_Ligne!$A$%d:$A$%d,Referentiels!$A$%d:$A$%d)=0))'
     % (GR_FIRST, GR_FIRST + 9, RES_FIRST, RES_LAST, GR_FIRST, GR_FIRST + 9),
     "Sans responsable, la couverture des séances reste incomplète."),
    ("Qualifications déclarées",
     '=SUMPRODUCT(--(Qualifications_Personnes!$B$%d:$B$%d<>""))'
     % (QP_FIRST, QP_LAST),
     "Conditionne le contrôle de validité des affectations."),
    ("dont d'encadrement",
     '=COUNTIF(Qualifications_Personnes!$F$%d:$F$%d,"Encadrement")'
     % (QP_FIRST, QP_LAST),
     "AEL, ACEL, AEEL, IE, MEF1, MEF2."),
    ("dont de sécurité",
     '=COUNTIF(Qualifications_Personnes!$F$%d:$F$%d,"Sécurité")'
     % (QP_FIRST, QP_LAST),
     "RIFAA, PSM."),
    ("Qualifications sans prérogatives renseignées",
     '=COUNTIFS(Referentiels!$A$%d:$A$%d,"<>",Referentiels!$E$%d:$E$%d,"")'
     % (QU_FIRST, QU_LAST, QU_FIRST, QU_LAST),
     "À renseigner d'après les textes fédéraux."),
    ("Périodes dérogatoires sans dates",
     '=COUNTIFS(Referentiels!$A$%d:$A$%d,"<>",Referentiels!$B$%d:$B$%d,"")'
     % (PER_FIRST, PER_LAST, PER_FIRST, PER_LAST),
     "Calendrier scolaire et jours fériés à saisir."),
    ("Objectifs au référentiel",
     '=COUNTIF(Objectifs_Niveaux!$A$%d:$A$%d,"<>")' % (OBJ_FIRST, OBJ_LAST),
     "46 attendus pour L1 à L4."),
    ("Objectifs sans critère de validation chiffré",
     '=COUNTIFS(Objectifs_Niveaux!$A$%d:$A$%d,"<>",'
     'Objectifs_Niveaux!$L$%d:$L$%d,"")'
     % (OBJ_FIRST, OBJ_LAST, OBJ_FIRST, OBJ_LAST),
     "Les 5 objectifs de maîtrise du L4 : critère qualitatif à trancher."),
    ("Compétences sans critère observable",
     '=COUNTIFS(Referentiels!$A$%d:$A$%d,"<>",'
     'Referentiels!$C$%d:$C$%d,"")'
     % (CMP_FIRST, CMP_LAST, CMP_FIRST, CMP_LAST),
     "Les critères proposés restent à valider collégialement."),
]
for label, f, note in CTRL:
    style_cell(ws.cell(r, 1, label), "body", wrap=True)
    style_cell(ws.cell(r, 2, f), "calc")
    style_cell(ws.cell(r, 3, note), "body", wrap=True)
    r += 1
CTRL_LAST = r - 1
make_table(ws, "Controles", CTRL_HEADER_ROW, CTRL_LAST, 3)
r += 1
ws.cell(r, 1, "Synchronisation des classeurs de ligne").font = SEC_FONT
r += 1
for txt in [
    "Les classeurs de ligne ne sont pas liés à celui-ci par des formules "
    "externes : ce type de lien casse dès qu'un fichier est déplacé, renommé ou "
    "synchronisé par OneDrive.",
    "Ils reçoivent une copie du paramétrage, rafraîchie par le script "
    "« sync_referentiel.py ». Chaque classeur de ligne affiche dans son onglet "
    "Synchro la version et la date du paramétrage dont il est issu.",
]:
    c = ws.cell(r, 1, txt)
    c.font = BODY
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 14 * (1 + len(txt) // 58)
    r += 1

# ============================================================= Selection
ws, r = sheet(wb, "Selection", "Liste de sélection des personnes",
              "Onglet technique. Les personnes de l'onglet Personnes, triées "
              "par prénom. Alimente toutes les colonnes « Sélection par nom ou "
              "prénom ». Ne rien modifier ici.", tab_color=TAB_TECHNIQUE)
ws.column_dimensions["A"].width = 40
c = ws.cell(1, 1, "Personnes triées par prénom")
c.fill, c.font, c.border = H_FILL, H_FONT, BOX
for i in range(N_CLUB):
    row = SEL_FIRST + i
    style_cell(ws.cell(row, 1,
        '=IFERROR(INDEX({lib},MATCH({n},{rang},0)),"")'
        .format(lib=R_LIB, n=i + 1, rang=R_RANG)), "calc")
make_table(ws, "Selection", 1, SEL_LAST, 1)
ws.freeze_panes = ws.cell(SEL_FIRST, 1)

# =============================================================== Listes
ws, r = sheet(wb, "Listes", "Listes de valeurs",
              "Onglet technique : valeurs autorisées, reprises par "
              "l'application. Ne pas modifier sans revoir les listes déroulantes.",
              tab_color=TAB_TECHNIQUE)
LISTES = [("Niveaux de compétence", NIVEAUX_COMP),
          ("Statut d'acquisition", STATUT_ACQ),
          ("Zone de confort", ZONE_CONFORT),
          ("Rôles d'encadrement", ROLES_ENC),
          ("Rôles de personne", ROLES_PERS),
          ("Supports de saisie", SUPPORTS),
          ("Périodicités", PERIODICITES),
          ("Motifs d'indisponibilité", MOTIFS_INDISPO),
          ("Motifs d'absence", MOTIFS_ABS),
          ("Jours", JOURS), ("Genres", GENRES),
          ("Statuts de personne", STATUTS_PERS),
          ("Types d'objectif", TYPES_OBJ),
          ("Contraintes de validation", CONTRAINTES),
          ("Natures de séance", NATURES),
          ("Statuts de séance", STATUTS_SEANCE),
          ("Avis d'encadrant", AVIS),
          ("Types d'événement", TYPES_EVT), ("Gravités", GRAVITES)]
LISTES_HEADER_ROW = r
for i, (label, vals) in enumerate(LISTES, start=1):
    col = get_column_letter(i)
    ws.column_dimensions[col].width = max(16, min(28, len(label) + 2))
    c = ws.cell(r, i, label)
    c.fill, c.font, c.border = H_FILL, H_FONT, BOX
    c.alignment = Alignment(wrap_text=True, vertical="center")
    for j, v in enumerate(vals, start=1):
        style_cell(ws.cell(r + j, i, v), "body")
LISTES_LAST = r + max(len(vals) for _, vals in LISTES)
make_table(ws, "Listes", LISTES_HEADER_ROW, LISTES_LAST, len(LISTES))
ws.row_dimensions[r].height = 32
ws.freeze_panes = ws.cell(r + 1, 1)

# Tables normalisées utilisées par les prochaines migrations : un identifiant
# stable, le libellé présenté à l'utilisateur et un indicateur d'activation.
ws.merge_cells(start_row=2, start_column=21, end_row=2, end_column=23)
ws.cell(2, 21, "Tables normalisées : ID stable, libellé affiché, statut actif.").font = NOTE_FONT
for start_col, table_name, values in [
    (21, "Zone de confort", ZONE_CONFORT),
    (25, "Statuts de séance", STATUTS_SEANCE),
]:
    headers = [f"{table_name} - ID", f"{table_name} - libellé", "Actif"]
    for offset, value in enumerate(headers):
        c = ws.cell(4, start_col + offset, value)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for index, value in enumerate(values, start=1):
        for offset, cell_value in enumerate((index, value, "oui")):
            style_cell(ws.cell(4 + index, start_col + offset, cell_value), "body")
    for offset, width in enumerate((20, 26, 10)):
        ws.column_dimensions[get_column_letter(start_col + offset)].width = width
ws.row_dimensions[4].height = 32

wb.save(OUT)
print("écrit :", OUT)
print("onglets :", len(wb.sheetnames), wb.sheetnames)
print("personnes pré-remplies :", N_REMPLIES)
