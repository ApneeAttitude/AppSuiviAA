#!/usr/bin/env python3
"""Classeurs de suivi par ligne d'eau — saison 2026-2027.

Usage : python3 build_suivi.py [CODE_GROUPE ...]   (défaut : toutes les lignes)
"""
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from aa_common import *          # noqa: F403
from aa_common import (sheet, header, dv_list, dv_range, style_cell,
                       write_readme, make_table)
import aa_data as D

# Chemin relatif au script, portable d'un environnement à l'autre — pas un
# chemin absolu propre à une seule machine.
OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                      "..", "classeurs")
os.makedirs(OUTDIR, exist_ok=True)

# ---- positions figées, connues du script de synchronisation. Creneaux_Ref
# et Responsables sont des sections 3 et 4 de l'onglet Referentiel (fusion
# dans le même esprit que Referentiels côté paramétrage) : ce ne sont pas des
# onglets alimentés en direct par la saisie mais des copies recopiées par
# sync_referentiel.py, qui a donc besoin d'une capacité déjà présente — d'où
# une capacité fixe modeste plutôt qu'une simple ligne libre.
REF_OBJ_TITRE, REF_OBJ_HDR, REF_OBJ_FIRST, REF_OBJ_N = 4, 5, 6, 20
REF_CMP_TITRE, REF_CMP_HDR, REF_CMP_FIRST, REF_CMP_N = 28, 29, 30, 20
CREF_TITRE, CREF_HDR, CREF_FIRST, CREF_N = 52, 53, 54, 10
RESP_TITRE, RESP_HDR, RESP_FIRST, RESP_N = 66, 67, 68, 10
# Section 5 : personnes du club détenant une qualification de type Sécurité
# valide (RIFAA, PSM...) — club entier, pas de notion de ligne. Alimente le
# contrôle « DP sans qualification Sécurité valide » du Calendrier
# (1.8_matrice_rbac_v0.md F3, COMPTE-07/D-15).
QUAL_TITRE, QUAL_HDR, QUAL_FIRST, QUAL_N = 79, 80, 81, 30
PERS_FIRST, PERS_N = 5, 150     # aligné sur la capacité club du paramétrage
ROSTER_N = 150                    # lignes des matrices de suivi (1 ligne)
# Relevé à 40 le 30/08/2026 : L3 a atteint 31 membres réels (inscriptions
# restaurées + ajouts de saison), au-delà de l'ancienne capacité de 30.
# Relevé à 150 le 13/09/2026 : DNF1 regroupe 110 inscriptions actives des
# niveaux L1 à L4. Les lignes transverses DNF/STA dépassent donc largement
# la capacité d'une ligne d'eau classique.
# Calendrier, Zones_Physiologiques : aucune donnée n'est reprise du
# paramétrage, tout y est saisi à la main au fil de la saison — une seule
# ligne libre suffit, la Table Excel s'agrandit d'elle-même à chaque
# nouvelle séance saisie.
# Presences (500) et Evenements_Securite (100) : pré-provisionnées, formules
# et mise en forme déjà en place dès la génération, plutôt qu'une seule
# ligne qui grandit au fil de la saisie — une fois converti en Google Sheet,
# une ligne ajoutée en bas d'un tableau ne récupère pas automatiquement
# formules ni couleurs comme le ferait une Table Excel (06/09/2026).
CAL_FIRST, CAL_N = 5, 1
PRE_FIRST, PRE_N = 5, 500
EV_FIRST, EV_N = 5, 100
ZO_FIRST, ZO_N = 5, 1
AV_CAMP, AV_ENC = 2, 6           # campagnes et colonnes d'encadrant

REF_OBJ_LAST = REF_OBJ_FIRST + REF_OBJ_N - 1
REF_CMP_LAST = REF_CMP_FIRST + REF_CMP_N - 1
CREF_LAST = CREF_FIRST + CREF_N - 1
RESP_LAST = RESP_FIRST + RESP_N - 1
QUAL_LAST = QUAL_FIRST + QUAL_N - 1
PERS_LAST = PERS_FIRST + PERS_N - 1
CAL_LAST = CAL_FIRST + CAL_N - 1
PRE_LAST = PRE_FIRST + PRE_N - 1
EV_LAST = EV_FIRST + EV_N - 1
ZO_LAST = ZO_FIRST + ZO_N - 1

# Personnes : A code · B id · C nom · D prénom · E rôle · F groupes
#             G objectif · H membre · I séances suivies · J tenues · K taux
#             L code élève · M code encadrant · N libellé · O rang · P courriel
R_CODE = "Personnes!$A$%d:$A$%d" % (PERS_FIRST, PERS_LAST)
R_ID = "Personnes!$B$%d:$B$%d" % (PERS_FIRST, PERS_LAST)
R_NOM = "Personnes!$C$%d:$C$%d" % (PERS_FIRST, PERS_LAST)
R_PRE = "Personnes!$D$%d:$D$%d" % (PERS_FIRST, PERS_LAST)
R_OBJ = "Personnes!$G$%d:$G$%d" % (PERS_FIRST, PERS_LAST)
R_MEMBRE = "Personnes!$H$%d:$H$%d" % (PERS_FIRST, PERS_LAST)
R_CODE_ELEVE = "Personnes!$L$%d:$L$%d" % (PERS_FIRST, PERS_LAST)
R_CODE_ENC = "Personnes!$M$%d:$M$%d" % (PERS_FIRST, PERS_LAST)
R_LIB = "Personnes!$N$%d:$N$%d" % (PERS_FIRST, PERS_LAST)
R_RANG = "Personnes!$O$%d:$O$%d" % (PERS_FIRST, PERS_LAST)
R_EMAIL = "Personnes!$P$%d:$P$%d" % (PERS_FIRST, PERS_LAST)
SEL_FIRST, SEL_LAST = 2, 2 + PERS_N - 1
SEL_PERS = "=Selection!$A$%d:$A$%d" % (SEL_FIRST, SEL_LAST)
SEL_CR = "=Referentiel!$G$%d:$G$%d" % (CREF_FIRST, CREF_LAST)
R_CRID = "Referentiel!$A$%d:$A$%d" % (CREF_FIRST, CREF_LAST)
R_CREF_ENC_ID = "Referentiel!$H$%d:$H$%d" % (CREF_FIRST, CREF_LAST)
R_QUAL_ID = "Referentiel!$A$%d:$A$%d" % (QUAL_FIRST, QUAL_LAST)
R_SEANCE = "Calendrier!$A$%d:$A$%d" % (CAL_FIRST, CAL_LAST)

MSG_PERS = ("Choisissez la personne dans la liste, triée par prénom. Tapez "
            "les premières lettres du prénom pour y aller plus vite. "
            "L'identifiant, le nom et le prénom se remplissent ensuite tout "
            "seuls. Le code numérique de l'onglet Personnes reste aussi "
            "accepté.")
MSG_ENC = MSG_PERS


# --------------------------------------------------------------- formules
def f_resolve(inp):
    """ID depuis une sélection de la liste (texte entre crochets), un code
    numérique court, ou un début de nom / prénom tapé directement."""
    return ('=IF({i}="","",IFERROR(IF(ISNUMBER(FIND("[",{i})),'
            'MID({i},FIND("[",{i})+1,FIND("]",{i})-FIND("[",{i})-1),'
            'IF(ISNUMBER({i}),INDEX({id},MATCH({i},{code},0)),'
            'IFERROR(INDEX({id},MATCH({i}&"*",{nom},0)),'
            'INDEX({id},MATCH({i}&"*",{pre},0))))),""))'
            ).format(i=inp, code=R_CODE, id=R_ID, nom=R_NOM, pre=R_PRE)


def f_lookup(target, key):
    return '=IFERROR(INDEX(%s,MATCH(%s,%s,0))&"","")' % (target, key, R_ID)


def bloc_saisie(ws, first, last, c_inp, c_id, c_nom, c_pre, msg=MSG_PERS):
    """Une colonne de sélection assistée par liste (même principe que le
    paramétrage), trois colonnes de confirmation déduites."""
    ci = get_column_letter(c_id)
    li = get_column_letter(c_inp)
    for row in range(first, last + 1):
        style_cell(ws.cell(row, c_inp), "select")
        style_cell(ws.cell(row, c_id, f_resolve("$%s%d" % (li, row))), "calc")
        style_cell(ws.cell(row, c_nom, f_lookup(R_NOM, "$%s%d" % (ci, row))),
                   "calc")
        style_cell(ws.cell(row, c_pre, f_lookup(R_PRE, "$%s%d" % (ci, row))),
                   "calc")
    dv_range(ws, SEL_PERS, "%s%d:%s%d" % (li, first, li, last), msg)


def write_calendrier_row(ws, row, groupe, statut=None):
    """Écrit les formules d'une ligne de l'onglet Calendrier. Réutilisée par
    sync_referentiel.py pour l'initialisation automatique des séances de la
    saison : seules les colonnes Date (et Statut, optionnel) portent une
    valeur saisie, tout le reste est déduit."""
    # La date (et le jour) sont intégrés directement dans l'identifiant de
    # séance : c'est ce même texte qui alimente la liste déroulante de
    # sélection de séance dans Presences/Evenements_Securite, et un simple
    # « S001 » n'y était pas assez parlant (06/09/2026). TEXT($B,"DD/MM/YYYY")
    # a été évité : les jetons de format de date sont dépendants de la
    # locale (Excel FR attend "JJ/MM/AAAA"), ce qui a produit #VALEUR! une
    # fois le classeur ouvert dans Google Sheets — DAY/MONTH/YEAR avec un
    # simple format numérique "00" ne dépend d'aucune locale.
    style_cell(ws.cell(row, 1,
        '=IF($B{r}="","","S"&TEXT(COUNTIF($B${f}:$B{r},"<>"),"000")'
        '&" — "&TEXT(DAY($B{r}),"00")&"/"&TEXT(MONTH($B{r}),"00")'
        '&"/"&YEAR($B{r})&" ("&$C{r}&")")'
        .format(r=row, f=CAL_FIRST)), "calc")
    style_cell(ws.cell(row, 2), "input", fmt="DD/MM/YYYY")
    style_cell(ws.cell(row, 3),
               "calc").value = ('=IF($B{r}="","",IFERROR(INDEX('
                                'Listes!$A$2:$A$8,WEEKDAY($B{r},2)),""))'
                                .format(r=row))
    style_cell(ws.cell(row, 4), "select")
    style_cell(ws.cell(row, 5,
        '=IFERROR(MID($D{r},FIND("[",$D{r})+1,'
        'FIND("]",$D{r})-FIND("[",$D{r})-1),"")'.format(r=row)), "calc")
    for col, src in ((6, "E"), (7, "F")):
        style_cell(ws.cell(row, col,
            '=IFERROR(IF($E{r}="","",INDEX(Referentiel!${s}${f}:${s}${l},'
            'MATCH($E{r},{rid},0))&""),"")'
            .format(s=src, f=CREF_FIRST, l=CREF_LAST, r=row, rid=R_CRID)),
            "calc")
    style_cell(ws.cell(row, 8), "select")
    style_cell(ws.cell(row, 9, statut or ""), "select")
    # directeur de bassin : saisie assistée, comme bloc_saisie (colonnes J-M),
    # jamais répliquée ici avant le 06/09/2026 — d'où l'absence de couleurs,
    # listes déroulantes et déduction du nom/prénom sur les séances générées
    # par sync_referentiel.py (seule la ligne-modèle initiale les avait).
    style_cell(ws.cell(row, 10), "select")
    style_cell(ws.cell(row, 11, f_resolve("$J%d" % row)), "calc")
    style_cell(ws.cell(row, 12, f_lookup(R_NOM, "$K%d" % row)), "calc")
    style_cell(ws.cell(row, 13, f_lookup(R_PRE, "$K%d" % row)), "calc")
    # responsable : remplaçant saisi, sinon encadrant du créneau (Referentiel §3)
    style_cell(ws.cell(row, 14), "select")
    style_cell(ws.cell(row, 15,
        '=IF($N{r}<>"",{res},IF($E{r}="","",'
        'IFERROR(INDEX({encid},MATCH($E{r},{rid},0)),"")))'
        .format(r=row, res=f_resolve("$N%d" % row)[1:],
                encid=R_CREF_ENC_ID, rid=R_CRID)), "calc")
    style_cell(ws.cell(row, 16, f_lookup(R_NOM, "$O%d" % row)), "calc")
    style_cell(ws.cell(row, 17, f_lookup(R_PRE, "$O%d" % row)), "calc")
    for col in (18, 19):
        style_cell(ws.cell(row, col), "input", wrap=True)
    style_cell(ws.cell(row, 20,
        '=IF($A{r}="","",COUNTIFS(Presences!$A${pf}:$A${pl},$A{r},'
        'Presences!$C${pf}:$C${pl},"<>"))'
        .format(r=row, pf=PRE_FIRST, pl=PRE_LAST)), "calc")
    style_cell(ws.cell(row, 21,
        '=IF($A{r}="","",'
        'IF(AND($J{r}<>"",$K{r}=""),"DP non reconnu",'
        'IF(AND($K{r}<>"",COUNTIF({qualid},$K{r})=0),'
        '"DP sans qualification Sécurité valide",'
        'IF(AND($N{r}<>"",$O{r}=""),"responsable non reconnu",'
        'IF(AND($G{r}<>"",$G{r}<>"{g}"),"créneau rattaché à un autre groupe",'
        'IF($I{r}<>"tenue","",'
        'IF(OR($K{r}="",$O{r}=""),"couverture incomplète","ok")))))))'
        .format(r=row, g=groupe, qualid=R_QUAL_ID)), "calc")


def bloc_roster(ws, first, c_code, c_id, c_nom, c_pre, n=ROSTER_N,
                code_range=None, id_range=None):
    """Les n premiers membres de la ligne, construits automatiquement.
    code_range/id_range permettent de restreindre le roster (élèves ou
    encadrants) au lieu de tous les membres."""
    code_range = code_range or R_CODE
    id_range = id_range or R_ID
    for i in range(n):
        row, rang = first + i, i + 1
        style_cell(ws.cell(row, c_code, rang), "body")
        style_cell(ws.cell(row, c_id,
            '=IFERROR(INDEX(%s,MATCH(%d,%s,0)),"")' % (id_range, rang, code_range)),
            "calc")
        ci = get_column_letter(c_id)
        style_cell(ws.cell(row, c_nom, f_lookup(R_NOM, "$%s%d" % (ci, row))),
                   "calc")
        style_cell(ws.cell(row, c_pre, f_lookup(R_PRE, "$%s%d" % (ci, row))),
                   "calc")


def build(groupe):
    lib = dict(D.LIGNES)[groupe]
    # Les deux lignes de compétition reprennent le référentiel L4 : elles
    # n'ont pas de plan d'entraînement autonome dans le MVP.
    ref_groupe = "L4" if groupe in ("LC", "STAC") else groupe
    obj_m = [o for o in D.OBJECTIFS if o[1] == ref_groupe and o[3] == "maîtrise"]
    obj_a = [o for o in D.OBJECTIFS if o[1] == ref_groupe and o[3] == "acquisition"]
    out = "%s/AA - Suivi %s %s.xlsx" % (OUTDIR, groupe, D.SAISON)
    wb = Workbook()
    # test d'appartenance robuste aux listes de groupes (« L2 / L3 »)
    membre = ('ISNUMBER(SEARCH("/{g}/","/"&SUBSTITUTE($F{{r}}," / ","/")&"/"))'
              .format(g=groupe))

    # =================================================== Lisez-moi
    ws = wb.active
    ws.title = "Lisez-moi"
    herite = ("", "")
    if groupe in ("LC", "STAC"):
        herite = ("li", "La ligne compétition n'a pas de plan d'entraînement "
                        "propre : le référentiel repris est celui du L4. "
                        "À valider.")
    write_readme(ws, "Suivi des entraînements — ligne %s — saison %s"
                 % (lib, D.SAISON), [
        ("", ""),
        ("titre", "Un classeur par ligne d'eau"),
        ("p", "Ce classeur suit une seule ligne d'eau, pour une seule saison. "
              "Le paramétrage du club vit dans « AA - Parametrage %s.xlsx » ; "
              "ce classeur en reçoit une copie de travail. Voir l'onglet "
              "Synchro." % D.SAISON),
        ("p", "Une ligne d'eau a un seul support de saisie déclaré pour la "
              "saison : ce classeur, ou l'application. Jamais les deux."),
        ("", ""),
        ("titre", "Quatre couleurs, quatre natures de cellule"),
        ("li", "Blanc — saisie libre, à remplir à la main."),
        ("li", "Vert clair — saisie assistée par une liste déroulante : "
               "choisissez une valeur proposée."),
        ("li", "Gris clair — calculé par une formule de ce classeur : ça se "
               "met à jour tout seul, n'y touchez jamais."),
        ("li", "Bleu pâle — copié automatiquement depuis le classeur de "
               "Paramétrage à chaque synchronisation (onglets Referentiel, "
               "Personnes, Creneaux_Ref, Responsables, Qualifs_Securite) : "
               "une saisie ici serait silencieusement écrasée au prochain "
               "Synchro, sans message d'erreur."),
        ("", ""),
        ("titre", "Sélectionner une personne"),
        ("p", "Même principe que le classeur de paramétrage : une liste "
              "déroulante propose les personnes du club, triée par prénom. "
              "Tapez les premières lettres du prénom pour y aller plus vite, "
              "ou choisissez dans la liste. Le code numérique de l'onglet "
              "Personnes reste aussi accepté. Les colonnes suivantes affichent "
              "aussitôt l'identifiant, le nom et le prénom : un coup d'œil "
              "suffit à vérifier."),
        ("", ""),
        ("titre", "Les listes se construisent toutes seules"),
        ("p", "L'onglet Personnes porte l'ensemble du club, avec le groupe de "
              "niveau de chacun. Les membres de cette ligne y reçoivent "
              "automatiquement un code, de 1 à n. Les onglets Suivi_Objectifs, "
              "Suivi_Competences et Avis_Encadrants construisent leur liste "
              "d'apnéistes à partir de ce code : vous n'avez jamais de nom à "
              "recopier."),
        ("", ""),
        ("titre", "Où saisir"),
        ("li", "Calendrier — une ligne par séance, numérotée automatiquement "
               "(S001, S002…). Le responsable de ligne est déduit de l'onglet "
               "Responsables ; la colonne « responsable » ne sert qu'aux "
               "remplacements."),
        ("li", "Presences — une ligne par participant présent. Pas de colonne "
               "« présent » : l'existence de la ligne vaut présence."),
        ("li", "Suivi_Objectifs — le meilleur nombre de répétitions réussies au "
               "cours d'une même séance, et le statut des objectifs en "
               "acquisition."),
        ("li", "Suivi_Competences — le niveau atteint sur les 8 compétences."),
        ("li", "Avis_Encadrants — un tableau apnéistes × encadrants, deux "
               "campagnes. Cette donnée n'est plus une couleur de cellule."),
        ("li", "Evenements_Securite — tout incident, daté et qualifié."),
        ("li", "Zones_Physiologiques — les bornes de confort et d'inconfort."),
        herite,
        ("", ""),
        ("titre", "Cinq règles"),
        ("li", "1. Ne saisissez que dans les cellules blanches ou vert clair "
               "— jamais dans le gris ou le bleu pâle."),
        ("li", "2. Une ligne = un enregistrement. Jamais de ligne vide au milieu."),
        ("li", "3. Ne renommez pas les onglets, ne déplacez pas les colonnes."),
        ("li", "4. Aucune information portée par une couleur de fond."),
        ("li", "5. Les onglets Referentiel, Personnes, Creneaux_Ref, "
               "Responsables et Qualifs_Securite sont écrasés à chaque "
               "synchronisation : n'y saisissez rien."),
        ("", ""),
        ("titre", "Version"),
        ("p", "Version %s — %s." % (VERSION, DATE_VERSION)),
    ])

    # ===================================================== Synchro
    ws, r = sheet(wb, "Synchro", "Synchronisation avec le paramétrage du club",
                  "Renseigné automatiquement par « sync_referentiel.py ». "
                  "Ne pas modifier à la main.")
    for L, w in (("A", 40), ("B", 46), ("C", 60)):
        ws.column_dimensions[L].width = w
    r = header(ws, r, [("Information", 40), ("Valeur", 46), ("Remarque", 60)])
    SYNC_FIRST = r
    for label, note in [
        ("Ligne d'eau", "Groupe de niveau suivi par ce classeur"),
        ("Saison", ""),
        ("Classeur de paramétrage source", ""),
        ("Version du paramétrage", ""),
        ("Date de la dernière synchronisation",
         "À relancer après toute modification du paramétrage"),
        ("Objectifs repris", ""),
        ("Compétences reprises", ""),
        ("Personnes reprises", "Tout le club : un encadrant n'est pas "
         "nécessairement inscrit dans la ligne qu'il encadre"),
        ("Créneaux repris", ""),
        ("Responsables de ligne repris", "Table datée : une ligne par "
         "changement en cours de saison"),
    ]:
        style_cell(ws.cell(r, 1, label), "body")
        style_cell(ws.cell(r, 2, ""), "link")
        style_cell(ws.cell(r, 3, note), "body", wrap=True)
        r += 1
    ws.cell(SYNC_FIRST, 2, groupe)
    ws.cell(SYNC_FIRST + 1, 2, D.SAISON)
    ws.cell(SYNC_FIRST + 2, 2, "non synchronisé")
    r += 1
    ws.cell(r, 1, "Membres de la ligne, calculés en direct").font = SEC_FONT
    r += 1
    style_cell(ws.cell(r, 1, "Apnéistes rattachés au groupe %s" % groupe), "body")
    style_cell(ws.cell(r, 2,
        '=COUNTIF(%s,"oui")' % R_MEMBRE), "calc")
    style_cell(ws.cell(r, 3, "Déduit de la colonne « Groupe(s) » de l'onglet "
                             "Personnes, sans intervention", ), "body", wrap=True)
    r += 2
    ws.cell(r, 1, "Pourquoi pas de lien Excel direct ?").font = SEC_FONT
    r += 1
    for txt in [
        "Une formule pointant vers un autre fichier dépend du chemin sur le "
        "disque. Elle casse dès que le classeur est déplacé, renommé, ouvert "
        "depuis un autre poste ou synchronisé par OneDrive.",
        "Le paramétrage est donc recopié ici, en lecture seule, par un script. "
        "C'est exactement ce que fera l'application plus tard.",
    ]:
        c = ws.cell(r, 1, txt)
        c.font = BODY
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 14 * (1 + len(txt) // 60)
        r += 1

    # ================================================= Referentiel
    # Quatre sections copiées du paramétrage, dans le même esprit que
    # l'onglet Referentiels du paramétrage : ce sont des copies en lecture
    # seule, jamais saisies à la main, donc autant les regrouper plutôt que
    # les disperser sur quatre onglets.
    ws, r = sheet(wb, "Referentiel", "Référentiel de la ligne %s" % lib,
                  "Copie du paramétrage du club. Écrasé à chaque "
                  "synchronisation : ne rien modifier ici.")
    assert r == REF_OBJ_TITRE, r
    ws.cell(REF_OBJ_TITRE, 1,
            "1. OBJECTIFS DE PROGRESSION — maîtrise et acquisition").font = SEC_FONT
    r = header(ws, REF_OBJ_HDR,
               [("ID", 10), ("Type", 12), ("Famille", 11), ("Intitulé", 26),
                ("Seuil de répétitions", 12),
                ("Intervalle max entre départs", 13),
                ("Critère de validation", 74)])
    assert r == REF_OBJ_FIRST, r
    for i in range(REF_OBJ_N):
        for col in range(1, 8):
            style_cell(ws.cell(REF_OBJ_FIRST + i, col), "link",
                       wrap=(col in (4, 7)))
    make_table(ws, "Objectifs", REF_OBJ_HDR, REF_OBJ_LAST, 7)
    ws.cell(REF_CMP_TITRE, 1,
            "2. COMPÉTENCES TECHNIQUES — niveau attendu en %s" % groupe
            ).font = SEC_FONT
    for i, (label, w) in enumerate([("Code", 12), ("Libellé", 26),
                                    ("Niveau attendu", 14),
                                    ("Critère observable", 74)], 1):
        c = ws.cell(REF_CMP_HDR, i, label)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[REF_CMP_HDR].height = 32
    for i in range(REF_CMP_N):
        for col in range(1, 5):
            style_cell(ws.cell(REF_CMP_FIRST + i, col), "link", wrap=(col == 4))
    make_table(ws, "Competences", REF_CMP_HDR, REF_CMP_LAST, 4)
    ws.cell(CREF_TITRE, 1,
            "3. CRÉNEAUX DE LA LIGNE — copie du paramétrage").font = SEC_FONT
    for i, (label, w) in enumerate([("ID", 10), ("Jour", 11),
                                    ("Heure début", 11), ("Heure fin", 11),
                                    ("Bassin", 9), ("Groupe de niveau", 15),
                                    ("Libellé (calculé)", 44),
                                    ("Encadrant (ID)", 12),
                                    ("Encadrant — nom", 18),
                                    ("Encadrant — prénom", 14)], 1):
        c = ws.cell(CREF_HDR, i, label)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)
    assert CREF_HDR + 1 == CREF_FIRST, CREF_FIRST
    for i in range(CREF_N):
        row = CREF_FIRST + i
        for col in range(1, 10):
            style_cell(ws.cell(row, col), "link")
        style_cell(ws.cell(row, 7,
            '=IF($A{r}="","",$A{r}&" — "&$B{r}'
            '&IF($C{r}=""," (horaire à saisir)"," "&$C{r}&"-"&$D{r})'
            '&IF($E{r}=""," "," · "&$E{r})'
            '&IF($F{r}=""," (groupe à saisir)"," · "&$F{r})'
            '&" ["&$A{r}&"]")'.format(r=row)), "calc")
    make_table(ws, "Creneaux_Ref", CREF_HDR, CREF_LAST, 10)
    ws.cell(RESP_TITRE, 1,
            "4. RESPONSABLE DE LIGNE — copie du paramétrage, table datée"
            ).font = SEC_FONT
    for i, (label, w) in enumerate([("Date d'effet", 13),
                                    ("Personne (ID)", 12), ("Nom", 22),
                                    ("Prénom", 16), ("Commentaire", 50)], 1):
        c = ws.cell(RESP_HDR, i, label)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)
    assert RESP_HDR + 1 == RESP_FIRST, RESP_FIRST
    for i in range(RESP_N):
        row = RESP_FIRST + i
        for col in range(1, 6):
            style_cell(ws.cell(row, col), "link", wrap=(col == 5),
                       fmt="DD/MM/YYYY" if col == 1 else None)
    make_table(ws, "Responsables", RESP_HDR, RESP_LAST, 5)
    ws.cell(QUAL_TITRE, 1,
            "5. PERSONNES QUALIFIÉES SÉCURITÉ — copie du paramétrage, "
            "club entier").font = SEC_FONT
    for i, (label, w) in enumerate([("Personne (ID)", 12), ("Nom", 22),
                                    ("Prénom", 16)], 1):
        c = ws.cell(QUAL_HDR, i, label)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)
    assert QUAL_HDR + 1 == QUAL_FIRST, QUAL_FIRST
    for i in range(QUAL_N):
        row = QUAL_FIRST + i
        for col in range(1, 4):
            style_cell(ws.cell(row, col), "link")
    make_table(ws, "Qualifs_Securite", QUAL_HDR, QUAL_LAST, 3)
    ws.freeze_panes = ws.cell(REF_OBJ_FIRST, 1)

    # =================================================== Personnes
    ws, r = sheet(wb, "Personnes", "Personnes du club et effectif de la ligne",
                  "Copie du référentiel du club. Le code de la colonne A est "
                  "attribué automatiquement aux membres du groupe %s : c'est le "
                  "raccourci de saisie." % groupe)
    assert r == PERS_FIRST - 1, r
    r = header(ws, r, [("Code", 8), ("ID", 10), ("Nom", 22), ("Prénom", 16),
                       ("Rôle déclaré", 18), ("Groupe(s)", 14),
                       ("Objectif de la saison", 38),
                       ("Membre de la ligne", 12),
                       ("Séances suivies", 11), ("Séances tenues", 11),
                       ("Taux de participation", 12),
                       ("Code élève", 10), ("Code encadrant", 12),
                       ("Libellé (calculé)", 34),
                       ("Rang par prénom (calculé)", 12),
                       ("Courriel", 30)])
    for i in range(PERS_N):
        row = PERS_FIRST + i
        for col in range(2, 8):
            style_cell(ws.cell(row, col), "link", wrap=(col == 7))
        style_cell(ws.cell(row, 16), "link")
        style_cell(ws.cell(row, 8,
            '=IF($B{r}="","",IF({m},"oui",""))'
            .format(r=row, m=membre.format(r=row))), "calc")
        style_cell(ws.cell(row, 1,
            '=IF($H{r}="","",COUNTIF($H${f}:$H{r},"oui"))'
            .format(r=row, f=PERS_FIRST)), "calc")
        style_cell(ws.cell(row, 9,
            '=IF($H{r}="","",COUNTIF(Presences!$C${pf}:$C${pl},$B{r}))'
            .format(r=row, pf=PRE_FIRST, pl=PRE_LAST)), "calc")
        style_cell(ws.cell(row, 10,
            '=IF($H%d="","",Tableau_de_bord!$B$5)' % row), "calc")
        # rangs séparés élèves / encadrants : rosters filtrés par rôle,
        # « élève et encadrant » compte dans les deux (D-20)
        style_cell(ws.cell(row, 12,
            '=IF($H{r}<>"oui","",IF(OR($E{r}="élève",$E{r}="élève et encadrant"),'
            'COUNTIFS($H${f}:$H{r},"oui",$E${f}:$E{r},"élève")'
            '+COUNTIFS($H${f}:$H{r},"oui",$E${f}:$E{r},"élève et encadrant"),""))'
            .format(r=row, f=PERS_FIRST)), "calc")
        style_cell(ws.cell(row, 13,
            '=IF($H{r}<>"oui","",IF(OR($E{r}="encadrant",$E{r}="élève et encadrant"),'
            'COUNTIFS($H${f}:$H{r},"oui",$E${f}:$E{r},"encadrant")'
            '+COUNTIFS($H${f}:$H{r},"oui",$E${f}:$E{r},"élève et encadrant"),""))'
            .format(r=row, f=PERS_FIRST)), "calc")
        style_cell(ws.cell(row, 11,
            '=IFERROR(IF($J{r}=0,"",$I{r}/$J{r}),"")'.format(r=row)),
            "calc", fmt="0.0%")
        # libellé et rang : alimentent la liste de sélection (onglet Selection)
        style_cell(ws.cell(row, 14,
            '=IF($B{r}="","",TRIM(IF($D{r}="",$C{r},$D{r}&" "&$C{r}))'
            '&" ["&$B{r}&"]")'.format(r=row)), "calc")
        style_cell(ws.cell(row, 15,
            '=IF($B{r}="","",COUNTIF({pre},"<"&$D{r})+COUNTIF($D${f}:$D{r},$D{r}))'
            .format(r=row, pre=R_PRE, f=PERS_FIRST)), "calc")
    make_table(ws, "Personnes", PERS_FIRST - 1, PERS_LAST, 16)
    ws.freeze_panes = ws.cell(PERS_FIRST, 5)

    # ================================================== Calendrier
    ws, r = sheet(wb, "Calendrier", "Calendrier des séances — ligne %s" % lib,
                  "Choisissez le créneau : jour, bassin et groupe se "
                  "remplissent seuls. Le responsable de ligne est déduit de "
                  "l'onglet Responsables.")
    r = header(ws, r, [("ID séance", 30), ("Date", 12), ("Jour", 11),
                       ("Créneau", 40), ("Créneau (ID)", 11), ("Bassin", 9),
                       ("Groupe", 9), ("Nature", 14), ("Statut", 11),
                       ("Directeur de bassin — début du nom", 18),
                       ("DP (ID)", 10), ("DP — nom", 18), ("DP — prénom", 14),
                       ("Responsable remplaçant — début du nom", 18),
                       ("Responsable (ID)", 12), ("Resp. — nom", 18),
                       ("Resp. — prénom", 14),
                       ("Plan de séance (texte ou lien)", 38),
                       ("Commentaire", 28), ("Nb de présents", 11),
                       ("Contrôle", 24)])
    assert r == CAL_FIRST, r
    bloc_saisie(ws, CAL_FIRST, CAL_LAST, 10, 11, 12, 13, MSG_ENC)
    for i in range(CAL_N):
        write_calendrier_row(ws, CAL_FIRST + i, groupe)
    dv_range(ws, SEL_PERS, "N%d:N%d" % (CAL_FIRST, CAL_LAST),
             "À remplir seulement si le responsable de cette séance n'est pas "
             "l'encadrant habituel du créneau. Laissez vide sinon.")
    dv_range(ws, SEL_CR, "D%d:D%d" % (CAL_FIRST, CAL_LAST),
             "Choisissez le créneau. La liste vient du paramétrage du club.")
    dv_list(ws, NATURES, "H%d:H%d" % (CAL_FIRST, CAL_LAST))
    dv_list(ws, STATUTS_SEANCE, "I%d:I%d" % (CAL_FIRST, CAL_LAST))
    make_table(ws, "Calendrier", CAL_FIRST - 1, CAL_LAST, 21)
    ws.freeze_panes = ws.cell(CAL_FIRST, 3)

    # =================================================== Presences
    ws, r = sheet(wb, "Presences", "Présences et observations",
                  "Une ligne par participant présent : l'existence de la ligne "
                  "vaut présence. Les encadrants sont dans l'onglet Calendrier.")
    r = header(ws, r, [("ID séance", 30),
                       ("Apnéiste — code ou début du nom", 26),
                       ("Apnéiste (ID)", 12), ("Nom", 20), ("Prénom", 14),
                       ("Zone de confort", 15),
                       ("Observation de l'encadrant", 62), ("Contrôle", 30)])
    assert r == PRE_FIRST, r
    bloc_saisie(ws, PRE_FIRST, PRE_LAST, 2, 3, 4, 5)
    for i in range(PRE_N):
        row = PRE_FIRST + i
        style_cell(ws.cell(row, 1), "select")
        style_cell(ws.cell(row, 6), "select")
        style_cell(ws.cell(row, 7), "input", wrap=True)
        style_cell(ws.cell(row, 8,
            '=IF($B{r}="","",'
            'IF($C{r}="","apnéiste non reconnu",'
            'IF($A{r}="","séance non renseignée",'
            'IF(COUNTIFS($A${f}:$A${l},$A{r},$C${f}:$C${l},$C{r})>1,'
            '"doublon séance / apnéiste",'
            'IF(OR($C{r}=IFERROR(INDEX(Calendrier!$K${cf}:$K${cl},'
            'MATCH($A{r},{seance},0)),""),'
            '$C{r}=IFERROR(INDEX(Calendrier!$O${cf}:$O${cl},'
            'MATCH($A{r},{seance},0)),"")),'
            '"encadrant et élève sur la même séance",'
            'IF(IFERROR(INDEX({mem},MATCH($C{r},{id},0)),"")<>"oui",'
            '"invité (hors ligne)","ok"))))))'
            .format(r=row, f=PRE_FIRST, l=PRE_LAST, cf=CAL_FIRST, cl=CAL_LAST,
                    seance=R_SEANCE, mem=R_MEMBRE, id=R_ID)), "calc")
    dv_range(ws, "=%s" % R_SEANCE, "A%d:A%d" % (PRE_FIRST, PRE_LAST),
             "Choisissez la séance dans l'onglet Calendrier.")
    dv_list(ws, ZONE_CONFORT, "F%d:F%d" % (PRE_FIRST, PRE_LAST))
    make_table(ws, "Presences", PRE_FIRST - 1, PRE_LAST, 8)
    ws.freeze_panes = ws.cell(PRE_FIRST, 4)

    # ============================================= Suivi_Objectifs
    ws, r = sheet(wb, "Suivi_Objectifs", "Suivi des objectifs — ligne %s" % lib,
                  "Liste construite automatiquement depuis l'onglet Personnes. "
                  "Saisissez le meilleur nombre de répétitions réussies au "
                  "cours d'une même séance.")
    ID_ROW, LIB_ROW, SEU_ROW, HDR_ROW = r, r + 1, r + 2, r + 3
    OBJ_FIRST = HDR_ROW + 1
    OBJ_LAST = OBJ_FIRST + ROSTER_N - 1
    ws.cell(ID_ROW, 5, "OBJECTIFS EN MAÎTRISE").font = SUB_FONT
    col, pct_cols = 5, []
    for o in obj_m:
        ws.cell(LIB_ROW, col, o[4]).font = SUB_FONT
        ws.cell(LIB_ROW, col).alignment = Alignment(wrap_text=True,
                                                   vertical="bottom")
        ws.cell(HDR_ROW, col, "Nb rép. %s" % o[0])
        ws.cell(HDR_ROW, col + 1, "%% %s" % o[0])
        style_cell(ws.cell(SEU_ROW, col,
            '=IFERROR(INDEX(Referentiel!$E${f}:$E${l},'
            'MATCH("{oid}",Referentiel!$A${f}:$A${l},0)),"")'
            .format(f=REF_OBJ_FIRST, l=REF_OBJ_LAST, oid=o[0])), "calc")
        ws.cell(SEU_ROW, col + 1, "seuil →").font = NOTE_FONT
        ws.column_dimensions[get_column_letter(col)].width = 9
        ws.column_dimensions[get_column_letter(col + 1)].width = 8
        pct_cols.append(get_column_letter(col + 1))
        col += 2
    if obj_a:
        ws.cell(ID_ROW, col, "OBJECTIFS EN ACQUISITION").font = SUB_FONT
    acq_first = col
    for o in obj_a:
        ws.cell(LIB_ROW, col, o[4]).font = SUB_FONT
        ws.cell(LIB_ROW, col).alignment = Alignment(wrap_text=True,
                                                   vertical="bottom")
        ws.cell(HDR_ROW, col, "Statut %s" % o[0])
        ws.column_dimensions[get_column_letter(col)].width = 13
        col += 1
    NCOL_OBJ = col - 1
    MAJ_PAR, MAJ_LE = NCOL_OBJ + 1, NCOL_OBJ + 2
    NCOL = MAJ_LE
    ws.cell(HDR_ROW, MAJ_PAR, "Mis à jour par")
    ws.cell(HDR_ROW, MAJ_LE, "le")
    ws.column_dimensions[get_column_letter(MAJ_PAR)].width = 24
    ws.column_dimensions[get_column_letter(MAJ_LE)].width = 12
    for i, (label, w) in enumerate([("Code", 7), ("ID", 10), ("Nom", 20),
                                    ("Prénom", 14)], 1):
        ws.cell(HDR_ROW, i, label)
        ws.column_dimensions[get_column_letter(i)].width = w
    for i in range(1, NCOL + 1):
        c = ws.cell(HDR_ROW, i)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
    ws.row_dimensions[LIB_ROW].height = 44
    bloc_roster(ws, OBJ_FIRST, 1, 2, 3, 4, code_range=R_CODE_ELEVE)
    for i in range(ROSTER_N):
        row = OBJ_FIRST + i
        col = 5
        for _ in obj_m:
            L = get_column_letter(col)
            style_cell(ws.cell(row, col), "input")
            style_cell(ws.cell(row, col + 1,
                '=IFERROR(IF($%s%d="","",MIN(1,$%s%d/%s$%d)),"")'
                % (L, row, L, row, L, SEU_ROW)), "calc", fmt="0%")
            col += 2
        for _ in obj_a:
            style_cell(ws.cell(row, col), "select")
            col += 1
        style_cell(ws.cell(row, MAJ_PAR), "select")
        style_cell(ws.cell(row, MAJ_LE), "input", fmt="DD/MM/YYYY")
    if obj_a:
        dv_list(ws, STATUT_ACQ, "%s%d:%s%d" % (get_column_letter(acq_first),
                                               OBJ_FIRST,
                                               get_column_letter(NCOL_OBJ), OBJ_LAST))
    dv_range(ws, SEL_PERS, "%s%d:%s%d" % (get_column_letter(MAJ_PAR), OBJ_FIRST,
                                          get_column_letter(MAJ_PAR), OBJ_LAST),
             "Qui vient de mettre à jour cette ligne ? À renseigner après "
             "chaque saisie ou correction.")
    make_table(ws, "Suivi_Objectifs", HDR_ROW, OBJ_LAST, NCOL)
    ws.freeze_panes = ws.cell(OBJ_FIRST, 5)

    # =========================================== Suivi_Competences
    ws, r = sheet(wb, "Suivi_Competences",
                  "Suivi des compétences — ligne %s" % lib,
                  "Liste construite automatiquement. Échelle à quatre paliers ; "
                  "la ligne « niveau attendu » vient du référentiel.")
    ATT_ROW, HDR_ROW = r, r + 1
    CMP_FIRST = HDR_ROW + 1
    CMP_LAST = CMP_FIRST + ROSTER_N - 1
    ws.cell(ATT_ROW, 4, "Niveau attendu →").font = SUB_FONT
    ws.cell(ATT_ROW, 4).alignment = Alignment(horizontal="right")
    cols = [("Code", 7), ("ID", 10), ("Nom", 20), ("Prénom", 14)]
    cols += [(c[1], 13) for c in D.COMPETENCES]
    cols += [("Acquises ou expertes", 12), ("Hors niveau attendu", 12),
             ("Mis à jour par", 24), ("le", 12)]
    MAJ_PAR, MAJ_LE = len(cols) - 1, len(cols)
    for i, (label, w) in enumerate(cols, 1):
        c = ws.cell(HDR_ROW, i, label)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[HDR_ROW].height = 36
    NC = len(D.COMPETENCES)
    for j, (code, _l, _c, _s) in enumerate(D.COMPETENCES):
        c = ws.cell(ATT_ROW, 5 + j,
            '=IFERROR(INDEX(Referentiel!$C${f}:$C${l},'
            'MATCH("{code}",Referentiel!$A${f}:$A${l},0)),"")'
            .format(f=REF_CMP_FIRST, l=REF_CMP_LAST, code=code))
        c.font, c.fill, c.border = CALC_FONT, SUB_FILL, BOX
        c.alignment = Alignment(horizontal="center")
    a, b = get_column_letter(5), get_column_letter(4 + NC)
    bloc_roster(ws, CMP_FIRST, 1, 2, 3, 4, code_range=R_CODE_ELEVE)
    for i in range(ROSTER_N):
        row = CMP_FIRST + i
        for j in range(NC):
            style_cell(ws.cell(row, 5 + j), "select").alignment = Alignment(
                horizontal="center")
        style_cell(ws.cell(row, 5 + NC,
            '=IF($B{r}="","",COUNTIF($%s{r}:$%s{r},"acquis")'
            '+COUNTIF($%s{r}:$%s{r},"expert"))'
            .replace("{r}", str(row)) % (a, b, a, b)), "calc")
        style_cell(ws.cell(row, 6 + NC,
            ('=IF($B{r}="","",SUMPRODUCT(($X1<>"")*($X1<>$X2)))'
             .replace("$X1", "${a}{r}:${b}{r}")
             .replace("$X2", "${a}${at}:${b}${at}")
             .format(r=row, at=ATT_ROW, a=a, b=b))), "calc")
        style_cell(ws.cell(row, MAJ_PAR), "select")
        style_cell(ws.cell(row, MAJ_LE), "input", fmt="DD/MM/YYYY")
    dv_list(ws, NIVEAUX_COMP, "%s%d:%s%d" % (a, CMP_FIRST, b, CMP_LAST))
    dv_range(ws, SEL_PERS, "%s%d:%s%d" % (get_column_letter(MAJ_PAR), CMP_FIRST,
                                          get_column_letter(MAJ_PAR), CMP_LAST),
             "Qui vient de mettre à jour cette ligne ? À renseigner après "
             "chaque saisie ou correction.")
    make_table(ws, "Suivi_Competences", HDR_ROW, CMP_LAST, len(cols))
    ws.freeze_panes = ws.cell(CMP_FIRST, 5)

    # ============================================= Avis_Encadrants
    ws, r = sheet(wb, "Avis_Encadrants", "Avis des encadrants",
                  "Tableau apnéistes × encadrants, une colonne par encadrant "
                  "et deux campagnes. Liste des apnéistes construite "
                  "automatiquement.")
    CAMP_ROW, ENC_ROW, EID_ROW, ENOM_ROW, EPRE_ROW = r, r + 1, r + 2, r + 3, r + 4
    AV_HDR = r + 5
    AV_FIRST = AV_HDR + 1
    AV_LAST = AV_FIRST + ROSTER_N - 1
    for i, (label, w) in enumerate([("Code", 7), ("ID", 10), ("Nom", 20),
                                    ("Prénom", 14)], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        ws.cell(AV_HDR, i, label)
    style_cell(ws.cell(CAMP_ROW, 1, "Campagne →"), "body")
    style_cell(ws.cell(ENC_ROW, 1, "Encadrant (sélection) →"), "body")
    style_cell(ws.cell(EID_ROW, 1, "ID →"), "body")
    style_cell(ws.cell(ENOM_ROW, 1, "Nom →"), "body")
    style_cell(ws.cell(EPRE_ROW, 1, "Prénom →"), "body")
    col = 5
    blocs = []
    for k in range(AV_CAMP):
        style_cell(ws.cell(CAMP_ROW, col, "Campagne %d" % (k + 1)), "input")
        first_enc = col
        for _ in range(AV_ENC):
            L = get_column_letter(col)
            style_cell(ws.cell(ENC_ROW, col), "select")
            style_cell(ws.cell(EID_ROW, col, f_resolve("$%s$%d" % (L, ENC_ROW))),
                       "calc")
            style_cell(ws.cell(ENOM_ROW, col,
                               f_lookup(R_NOM, "$%s$%d" % (L, EID_ROW))), "calc")
            style_cell(ws.cell(EPRE_ROW, col,
                               f_lookup(R_PRE, "$%s$%d" % (L, EID_ROW))), "calc")
            ws.cell(AV_HDR, col, "Avis")
            ws.column_dimensions[L].width = 15
            col += 1
        ws.cell(AV_HDR, col, "Commentaires")
        ws.column_dimensions[get_column_letter(col)].width = 46
        blocs.append((first_enc, col))
        col += 1
    NCOL = col - 1
    for i in range(1, NCOL + 1):
        c = ws.cell(AV_HDR, i)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
    # roster : élèves seulement, comme Suivi_Objectifs et Suivi_Competences —
    # les encadrants sont en colonnes, jamais en ligne
    bloc_roster(ws, AV_FIRST, 1, 2, 3, 4, code_range=R_CODE_ELEVE)
    for i in range(ROSTER_N):
        row = AV_FIRST + i
        for first_enc, com in blocs:
            for c_ in range(first_enc, com):
                style_cell(ws.cell(row, c_), "select")
            style_cell(ws.cell(row, com), "input", wrap=True)
    for first_enc, com in blocs:
        dv_list(ws, AVIS, "%s%d:%s%d" % (get_column_letter(first_enc), AV_FIRST,
                                         get_column_letter(com - 1), AV_LAST))
        L1, L2 = get_column_letter(first_enc), get_column_letter(com - 1)
        dv_range(ws, SEL_PERS, "%s%d:%s%d" % (L1, ENC_ROW, L2, ENC_ROW), MSG_ENC)
    make_table(ws, "Avis_Encadrants", AV_HDR, AV_LAST, NCOL)
    ws.freeze_panes = ws.cell(AV_FIRST, 5)

    # ========================================= Evenements_Securite
    ws, r = sheet(wb, "Evenements_Securite",
                  "Registre des événements de sécurité",
                  "Tout incident est daté, qualifié et rattaché à une séance. "
                  "La dernière colonne compte les antécédents de l'apnéiste.")
    r = header(ws, r, [("ID", 9), ("Date", 12), ("ID séance", 30),
                       ("Apnéiste — code ou début du nom", 26),
                       ("Apnéiste (ID)", 12), ("Nom", 20), ("Prénom", 14),
                       ("Type", 26), ("Gravité", 13), ("Contexte", 52),
                       ("Suites données", 40),
                       ("Déclaré par — début du nom", 18),
                       ("Déclarant (ID)", 12), ("Déclarant — nom", 18),
                       ("Déclarant — prénom", 14), ("Antécédents", 12)])
    assert r == EV_FIRST, r
    bloc_saisie(ws, EV_FIRST, EV_LAST, 4, 5, 6, 7)
    bloc_saisie(ws, EV_FIRST, EV_LAST, 12, 13, 14, 15, MSG_ENC)
    for i in range(EV_N):
        row = EV_FIRST + i
        style_cell(ws.cell(row, 1), "input")
        style_cell(ws.cell(row, 2), "input", fmt="DD/MM/YYYY")
        style_cell(ws.cell(row, 3), "select")
        for col in (8, 9):
            style_cell(ws.cell(row, col), "select")
        for col in (10, 11):
            style_cell(ws.cell(row, col), "input", wrap=True)
        style_cell(ws.cell(row, 16,
            '=IF($E{r}="","",COUNTIF($E${f}:$E${l},$E{r}))'
            .format(r=row, f=EV_FIRST, l=EV_LAST)), "calc")
    dv_range(ws, "=%s" % R_SEANCE, "C%d:C%d" % (EV_FIRST, EV_LAST))
    dv_list(ws, TYPES_EVT, "H%d:H%d" % (EV_FIRST, EV_LAST))
    dv_list(ws, GRAVITES, "I%d:I%d" % (EV_FIRST, EV_LAST))
    make_table(ws, "Evenements_Securite", EV_FIRST - 1, EV_LAST, 16)
    ws.freeze_panes = ws.cell(EV_FIRST, 6)

    # ======================================== Zones_Physiologiques
    ws, r = sheet(wb, "Zones_Physiologiques",
                  "Zones de confort et d'inconfort",
                  "Indicateur du plan d'entraînement : confort de 0-50 m à "
                  "0-75 m, inconfort de 50-85 m à 75-90 m en fin de formation.")
    r = header(ws, r, [("Apnéiste — code ou début du nom", 18),
                       ("Apnéiste (ID)", 12), ("Nom", 20), ("Prénom", 14),
                       ("Date", 12), ("Confort de (m)", 12),
                       ("Confort à (m)", 12), ("Inconfort de (m)", 12),
                       ("Inconfort à (m)", 12),
                       ("Évalué par — début du nom", 18),
                       ("Évaluateur (ID)", 12), ("Évaluateur — nom", 18),
                       ("Évaluateur — prénom", 14), ("Commentaire", 40)])
    assert r == ZO_FIRST, r
    bloc_saisie(ws, ZO_FIRST, ZO_LAST, 1, 2, 3, 4)
    bloc_saisie(ws, ZO_FIRST, ZO_LAST, 10, 11, 12, 13, MSG_ENC)
    for i in range(ZO_N):
        row = ZO_FIRST + i
        style_cell(ws.cell(row, 5), "input", fmt="DD/MM/YYYY")
        for col in (6, 7, 8, 9, 14):
            style_cell(ws.cell(row, col), "input", wrap=(col == 14))
    make_table(ws, "Zones_Physiologiques", ZO_FIRST - 1, ZO_LAST, 14)
    ws.freeze_panes = ws.cell(ZO_FIRST, 5)

    # ============================================ Tableau_de_bord
    ws, r = sheet(wb, "Tableau_de_bord", "Tableau de bord — ligne %s" % lib,
                  "Tout est calculé, par une seule méthode, visible et "
                  "reproductible.")
    ws.column_dimensions["A"].width = 46
    for L in "BCD":
        ws.column_dimensions[L].width = 15
    r = header(ws, r, [("Indicateur", 46), ("Valeur", 15), ("", 15), ("", 15)])
    ROW_TENUES = r
    for label, f, fmt in [
        ("Séances tenues",
         '=COUNTIF(Calendrier!$I$%d:$I$%d,"tenue")' % (CAL_FIRST, CAL_LAST), None),
        ("Séances planifiées non encore tenues",
         '=COUNTIF(Calendrier!$I$%d:$I$%d,"planifiée")' % (CAL_FIRST, CAL_LAST), None),
        ("Séances annulées ou fermées",
         '=COUNTIF(Calendrier!$I$%d:$I$%d,"annulée")'
         '+COUNTIF(Calendrier!$I$%d:$I$%d,"fermée")'
         % (CAL_FIRST, CAL_LAST, CAL_FIRST, CAL_LAST), None),
        ("Séances tenues sans couverture complète",
         '=COUNTIF(Calendrier!$U$%d:$U$%d,"couverture incomplète")'
         % (CAL_FIRST, CAL_LAST), None),
        ("Fréquentation moyenne d'une séance tenue",
         '=IFERROR(AVERAGEIFS(Calendrier!$T$%d:$T$%d,'
         'Calendrier!$I$%d:$I$%d,"tenue"),"")'
         % (CAL_FIRST, CAL_LAST, CAL_FIRST, CAL_LAST), "0.0"),
        ("Fréquentation la plus faible",
         '=IFERROR(_xlfn.MINIFS(Calendrier!$T$%d:$T$%d,'
         'Calendrier!$I$%d:$I$%d,"tenue"),"")'
         % (CAL_FIRST, CAL_LAST, CAL_FIRST, CAL_LAST), "0.0"),
        ("Fréquentation la plus forte",
         '=IFERROR(_xlfn.MAXIFS(Calendrier!$T$%d:$T$%d,'
         'Calendrier!$I$%d:$I$%d,"tenue"),"")'
         % (CAL_FIRST, CAL_LAST, CAL_FIRST, CAL_LAST), "0.0"),
    ]:
        style_cell(ws.cell(r, 1, label), "body", wrap=True)
        style_cell(ws.cell(r, 2, f), "calc", fmt=fmt)
        r += 1
    assert ROW_TENUES == 5, ROW_TENUES

    r += 1
    ws.cell(r, 1, "Fréquentation par jour de la semaine").font = SEC_FONT
    r += 1
    for i, label in enumerate(["Jour", "Séances tenues",
                               "Fréquentation moyenne"]):
        c = ws.cell(r, i + 1, label)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
    r += 1
    for j in JOURS:
        style_cell(ws.cell(r, 1, j), "body")
        style_cell(ws.cell(r, 2,
            '=COUNTIFS(Calendrier!$C$%d:$C$%d,$A%d,'
            'Calendrier!$I$%d:$I$%d,"tenue")'
            % (CAL_FIRST, CAL_LAST, r, CAL_FIRST, CAL_LAST)), "calc")
        style_cell(ws.cell(r, 3,
            '=IFERROR(AVERAGEIFS(Calendrier!$T$%d:$T$%d,'
            'Calendrier!$C$%d:$C$%d,$A%d,Calendrier!$I$%d:$I$%d,"tenue"),"")'
            % (CAL_FIRST, CAL_LAST, CAL_FIRST, CAL_LAST, r,
               CAL_FIRST, CAL_LAST)), "calc", fmt="0.0")
        r += 1

    r += 1
    ws.cell(r, 1, "Validation des objectifs en maîtrise").font = SEC_FONT
    r += 1
    for i, label in enumerate(["Objectif", "Ont validé", "En progression",
                               "Taux de validation"]):
        c = ws.cell(r, i + 1, label)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    r += 1
    for k, o in enumerate(obj_m):
        pc = pct_cols[k]
        style_cell(ws.cell(r, 1, o[4]), "body", wrap=True)
        style_cell(ws.cell(r, 2, '=COUNTIF(Suivi_Objectifs!$%s$%d:$%s$%d,1)'
                   % (pc, OBJ_FIRST, pc, OBJ_LAST)), "calc")
        style_cell(ws.cell(r, 3,
            '=COUNTIFS(Suivi_Objectifs!$%s$%d:$%s$%d,">0",'
            'Suivi_Objectifs!$%s$%d:$%s$%d,"<1")'
            % (pc, OBJ_FIRST, pc, OBJ_LAST, pc, OBJ_FIRST, pc, OBJ_LAST)),
            "calc")
        style_cell(ws.cell(r, 4,
            '=IFERROR($B%d/COUNTIF(%s,"oui"),"")' % (r, R_MEMBRE)),
            "calc", fmt="0.0%")
        r += 1

    r += 1
    ws.cell(r, 1, "Effectif, saisie et sécurité").font = SEC_FONT
    r += 1
    for label, f in [
        ("Apnéistes rattachés à la ligne", '=COUNTIF(%s,"oui")' % R_MEMBRE),
        ("Participations enregistrées",
         '=SUMPRODUCT(--(Presences!$C$%d:$C$%d<>""))' % (PRE_FIRST, PRE_LAST)),
        ("Lignes de présence en anomalie",
         '=SUMPRODUCT((Presences!$C$%d:$C$%d<>"")*'
         '(Presences!$H$%d:$H$%d<>"ok"))'
         % (PRE_FIRST, PRE_LAST, PRE_FIRST, PRE_LAST)),
        ("dont invités hors ligne",
         '=COUNTIF(Presences!$H$%d:$H$%d,"invité (hors ligne)")'
         % (PRE_FIRST, PRE_LAST)),
        ("Événements de sécurité déclarés",
         '=SUMPRODUCT(--(Evenements_Securite!$E$%d:$E$%d<>""))'
         % (EV_FIRST, EV_LAST)),
        ("dont graves",
         '=COUNTIF(Evenements_Securite!$I$%d:$I$%d,"grave")'
         % (EV_FIRST, EV_LAST)),
    ]:
        style_cell(ws.cell(r, 1, label), "body", wrap=True)
        style_cell(ws.cell(r, 2, f), "calc")
        r += 1

    # =================================================== Selection
    ws, r = sheet(wb, "Selection", "Liste de sélection des personnes",
                  "Onglet technique. Les personnes de l'onglet Personnes, "
                  "triées par prénom. Alimente toutes les colonnes de "
                  "sélection assistée. Ne rien modifier ici.",
                  tab_color=TAB_TECHNIQUE)
    ws.column_dimensions["A"].width = 40
    c = ws.cell(1, 1, "Personnes triées par prénom")
    c.fill, c.font, c.border = H_FILL, H_FONT, BOX
    for i in range(PERS_N):
        row = SEL_FIRST + i
        style_cell(ws.cell(row, 1,
            '=IFERROR(INDEX({lib},MATCH({n},{rang},0)),"")'
            .format(lib=R_LIB, n=i + 1, rang=R_RANG)), "calc")
    make_table(ws, "Selection", 1, SEL_LAST, 1)
    ws.freeze_panes = ws.cell(SEL_FIRST, 1)

    # ====================================================== Listes
    ws = wb.create_sheet("Listes")
    ws.sheet_properties.tabColor = TAB_TECHNIQUE
    ws["A1"] = "Jours"
    ws["A1"].fill, ws["A1"].font = H_FILL, H_FONT
    for i, j in enumerate(JOURS, start=2):
        style_cell(ws.cell(i, 1, j), "body")
    ws.column_dimensions["A"].width = 14
    COLS = [("Niveaux de compétence", NIVEAUX_COMP),
            ("Statut d'acquisition", STATUT_ACQ),
            ("Zone de confort", ZONE_CONFORT),
            ("Natures de séance", NATURES),
            ("Statuts de séance", STATUTS_SEANCE),
            ("Avis d'encadrant", AVIS),
            ("Types d'événement", TYPES_EVT), ("Gravités", GRAVITES)]
    for k, (label, vals) in enumerate(COLS, start=2):
        L = get_column_letter(k)
        ws.column_dimensions[L].width = max(16, min(28, len(label) + 2))
        c = ws.cell(1, k, label)
        c.fill, c.font = H_FILL, H_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
        for i, v in enumerate(vals, start=2):
            style_cell(ws.cell(i, k, v), "body")
    ws.row_dimensions[1].height = 32

    wb.save(out)
    return out


if __name__ == "__main__":
    codes = sys.argv[1:] or [g for g, _ in D.LIGNES]
    for code in codes:
        print("écrit :", build(code))
