#!/usr/bin/env python3
"""Synchronise les classeurs de ligne d'eau depuis le classeur de paramétrage.

Remplace les liens Excel entre fichiers, qui cassent dès qu'un fichier est
déplacé, renommé ou synchronisé par un service de stockage en ligne.

Usage :
    python3 sync_referentiel.py [dossier] [CODE_GROUPE ...]

Sans argument : dossier courant et toutes les lignes déclarées.

Règle importante : l'ordre des personnes déjà présentes dans un classeur de
ligne est préservé. Le code court de saisie (colonne A de l'onglet Personnes)
est un rang calculé parmi les membres de la ligne ; si l'ordre changeait, un
« 7 » saisi en octobre désignerait quelqu'un d'autre en mars. Les nouveaux
arrivants sont donc ajoutés à la fin, jamais insérés.

Initialisation du calendrier : si l'onglet Calendrier est encore vide, il est
rempli une fois pour toutes à partir des créneaux du groupe (Creneaux_Affectations)
et des dates de la saison (Périodes, régime NOMINAL). La périodicité est
hebdomadaire par défaut, ou suit la valeur en jours de l'affectation (par
exemple 14 jours pour une séance sur deux). N'exclut pas encore les vacances ni les
jours fériés : leurs dates ne sont pas encore renseignées au paramétrage.
Statut, directeur de bassin et plan de séance restent à compléter séance par
séance. Une fois une date saisie à la main, l'onglet n'est plus jamais
regénéré automatiquement.
"""
import os
import re
import sys
import glob
import datetime
import subprocess
import openpyxl

import aa_data as D
from aa_common import dv_range, dv_list, NATURES, STATUTS_SEANCE, JOURS
from build_suivi import (REF_OBJ_FIRST, REF_OBJ_LAST, REF_CMP_FIRST,
                         REF_CMP_LAST, PERS_FIRST, PERS_LAST, CREF_FIRST,
                         CREF_LAST, RESP_FIRST, RESP_LAST, QUAL_FIRST,
                         QUAL_LAST, ROSTER_N,
                         CAL_FIRST, SEL_CR, SEL_PERS, MSG_ENC,
                         PRE_FIRST, PRE_LAST, EV_FIRST, EV_LAST,
                         write_calendrier_row)

SAISON = D.SAISON
RECALC = "/root/.claude/skills/synced/xlsx/scripts/recalc.py"
JOUR_INDEX = {j: i for i, j in enumerate(JOURS)}   # lundi=0 … dimanche=6

# --- positions dans le classeur de paramétrage
# P_CMP pointe dans l'onglet Referentiels (section « Familles de compétences »,
# regroupée avec les autres petits référentiels) : la ligne de départ dépend
# de l'ordre des sections dans build_parametrage.py, à revérifier si cet
# ordre change.
P_PERS = (5, 205)      # bornes de range(), donc dernière ligne + 1
P_CREN = (5, 45)
P_OBJ = (5, 51)
P_CMP = (53, 62)
P_ATT = (5, 205)
P_INS = (5, 505)      # Inscriptions provisionné à 500 lignes (06/09/2026)
P_RES = (5, 105)
P_QUA = (5, 305)


def cellv(ws, row, col):
    v = ws.cell(row, col).value
    return "" if v is None else v


def extract_bracket_id(text):
    """Repli quand la valeur calculée d'une formule de résolution
    (identifiant entre crochets) n'est pas dans le cache — cas d'un
    classeur enregistré par openpyxl, qui ne recalcule jamais les
    formules. La saisie brute (« Prénom NOM [ID] ») reste lisible telle
    quelle, donc ce repli n'a besoin d'aucun recalcul Excel/LibreOffice."""
    m = re.search(r"\[([^\]]+)\]", str(text or ""))
    return m.group(1) if m else ""


def as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str) and len(v) >= 10:
        try:
            return datetime.date.fromisoformat(v[:10])
        except ValueError:
            return None
    return None


def actif(v):
    """Les tables métier utilisent « oui » / « non » ; une valeur vide
    n'est jamais considérée comme active."""
    return str(v or "").strip().lower() == "oui"


def periodicite_jours(v):
    """Lit une périodicité positive ; une ancienne affectation vide reste
    hebdomadaire afin de préserver les lignes déjà paramétrées."""
    try:
        n = int(v)
    except (TypeError, ValueError):
        return 7
    return n if n > 0 else 7


def recouvrent(debut_a, fin_a, debut_b, fin_b):
    """Deux périodes se recouvrent, bornes incluses. Une borne vide est
    ouverte : le modèle reste ainsi utilisable avant que toutes les dates
    d'une saison soient renseignées."""
    debut = max(x for x in (debut_a, debut_b) if x is not None) \
        if debut_a or debut_b else None
    fin = min(x for x in (fin_a, fin_b) if x is not None) \
        if fin_a or fin_b else None
    return not (debut and fin and debut > fin)


def read_parametrage(path):
    wv = openpyxl.load_workbook(path, data_only=True)
    par = {}

    ws = wv["Personnes"]
    par["personnes"] = {}
    par["ordre_club"] = []
    for r in range(*P_PERS):
        pid = cellv(ws, r, 1)
        if not pid:
            continue
        par["personnes"][pid] = {"id": pid, "nom": cellv(ws, r, 2),
                                 "prenom": cellv(ws, r, 3),
                                 "email": cellv(ws, r, 6)}
        par["ordre_club"].append(pid)

    ws = wv["Creneaux"]
    par["creneaux"] = {}
    for r in range(*P_CREN):
        cid = cellv(ws, r, 1)
        if cid:
            par["creneaux"][cid] = [cid, cellv(ws, r, 2), cellv(ws, r, 3),
                                    cellv(ws, r, 4), cellv(ws, r, 5)]

    ws = wv["Creneaux_Affectations"]
    par["affectations"] = []
    for r in range(5, 205):
        cid = cellv(ws, r, 2) or extract_bracket_id(cellv(ws, r, 1))
        grp = cellv(ws, r, 4)
        if cid and grp:
            encadrant = cellv(ws, r, 6) or extract_bracket_id(cellv(ws, r, 5))
            par["affectations"].append(
                {"creneau": cid, "date": as_date(ws.cell(r, 3).value),
                 "affectation": grp, "encadrant": encadrant,
                 # K : périodicité ; L : contrôle calculé ; M : référent.
                 "periodicite": periodicite_jours(cellv(ws, r, 11)),
                 "referent": actif(cellv(ws, r, 13))})

    # Les lignes DNF/STA ont leur propre durée de vie, distincte de celle des
    # groupes de niveau. Leur capacité ne figure volontairement pas dans les
    # affectations : une ligne par encadrant y serait ambiguë.
    par["lignes_seances"] = {}
    if "Lignes_Seances" in wv.sheetnames:
        ws = wv["Lignes_Seances"]
        for r in range(5, 205):
            code = cellv(ws, r, 1)
            if code:
                par["lignes_seances"][code] = {
                    "code": code, "libelle": cellv(ws, r, 2),
                    "nb_lignes": cellv(ws, r, 3),
                    "debut": as_date(ws.cell(r, 4).value),
                    "fin": as_date(ws.cell(r, 5).value),
                    "actif": actif(cellv(ws, r, 6)),
                }

    par["regles_acces_seances"] = []
    if "Regles_Acces_Seances" in wv.sheetnames:
        ws = wv["Regles_Acces_Seances"]
        for r in range(5, 405):
            ligne, groupe = cellv(ws, r, 1), cellv(ws, r, 2)
            if ligne and groupe:
                par["regles_acces_seances"].append({
                    "ligne": ligne, "groupe": groupe,
                    "debut": as_date(ws.cell(r, 3).value),
                    "fin": as_date(ws.cell(r, 4).value),
                    "actif": actif(cellv(ws, r, 5)),
                })

    ws = wv["Objectifs_Niveaux"]
    par["objectifs"] = []
    for r in range(*P_OBJ):
        oid = cellv(ws, r, 1)
        if oid:
            par["objectifs"].append({
                "id": oid, "groupe": cellv(ws, r, 2), "famille": cellv(ws, r, 3),
                "type": cellv(ws, r, 4), "intitule": cellv(ws, r, 5),
                "seuil": cellv(ws, r, 12), "intervalle": cellv(ws, r, 13),
                "critere": cellv(ws, r, 15)})

    ws = wv["Referentiels"]
    par["competences"] = []
    for r in range(*P_CMP):
        code = cellv(ws, r, 1)
        if code:
            par["competences"].append({"code": code, "libelle": cellv(ws, r, 2),
                                       "critere": cellv(ws, r, 3)})

    # Dates de la saison : première ligne « NOMINAL » de la section Périodes
    # qui porte de vraies dates (la section Régimes a aussi un code NOMINAL,
    # mais sa colonne B est un libellé texte, pas une date).
    par["saison_debut"] = par["saison_fin"] = None
    for r in range(1, 200):
        if cellv(ws, r, 1) == "NOMINAL":
            debut, fin = as_date(ws.cell(r, 2).value), as_date(ws.cell(r, 3).value)
            if debut and fin:
                par["saison_debut"], par["saison_fin"] = debut, fin
                break

    ws = wv["Competences_Niveaux"]
    par["attendus"] = {}
    for r in range(*P_ATT):
        code, grp = cellv(ws, r, 1), cellv(ws, r, 3)
        if code and grp:
            par["attendus"][(code, grp)] = cellv(ws, r, 4)

    ws = wv["Inscriptions"]
    par["inscriptions"] = []
    for r in range(*P_INS):
        pid = cellv(ws, r, 3) or extract_bracket_id(cellv(ws, r, 2))
        if pid and cellv(ws, r, 1) in ("", SAISON):
            par["inscriptions"].append({
                "id": pid, "groupe": cellv(ws, r, 6), "role": cellv(ws, r, 7),
                "support": cellv(ws, r, 8), "objectif": cellv(ws, r, 11),
                "fin": as_date(ws.cell(r, 10).value)})

    # Les membres de LC (groupe compétition élite) sont aussi membres de L4 :
    # on synthétise une inscription L4/élève pour quiconque est actif en LC
    # sans déjà avoir d'inscription L4 active — rien n'est écrit dans le
    # classeur de paramétrage, seulement dans cette liste en mémoire, pour
    # que le roster de L4 les intègre automatiquement (06/09/2026).
    aujourdhui_rp = datetime.date.today()
    lc_actifs = {x["id"] for x in par["inscriptions"]
                if x["groupe"] == "LC" and (x["fin"] is None or x["fin"] >= aujourdhui_rp)}
    l4_actifs = {x["id"] for x in par["inscriptions"]
                if x["groupe"] == "L4" and (x["fin"] is None or x["fin"] >= aujourdhui_rp)}
    for pid in sorted(lc_actifs - l4_actifs):
        par["inscriptions"].append({"id": pid, "groupe": "L4", "role": "élève",
                                    "support": "dérivé de LC", "objectif": "",
                                    "fin": None})

    ws = wv["Qualifications_Personnes"]
    par["dp_valides"] = []
    vus = set()
    for r in range(*P_QUA):
        pid = cellv(ws, r, 2) or extract_bracket_id(cellv(ws, r, 1))
        if pid and cellv(ws, r, 6) == "Sécurité" and cellv(ws, r, 10) == "oui" \
                and pid not in vus:
            vus.add(pid)
            pers = par["personnes"].get(pid, {})
            par["dp_valides"].append({"id": pid, "nom": pers.get("nom", ""),
                                      "prenom": pers.get("prenom", "")})

    ws = wv["Responsables_Ligne"]
    par["responsables"] = []
    for r in range(*P_RES):
        pid = cellv(ws, r, 4) or extract_bracket_id(cellv(ws, r, 3))
        d = as_date(ws.cell(r, 2).value)
        if pid and d:
            pers = par["personnes"].get(pid, {})
            par["responsables"].append({
                "groupe": cellv(ws, r, 1), "date": d, "id": pid,
                "nom": cellv(ws, r, 5) or pers.get("nom", ""),
                "prenom": cellv(ws, r, 6) or pers.get("prenom", ""),
                "commentaire": cellv(ws, r, 7)})

    ws = wv["Lisez-moi"]
    par["version"] = "inconnue"
    for r in range(1, 80):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v.startswith("Version "):
            par["version"] = v.split("—")[0].replace("Version", "").strip()
            break
    return par


def est_ligne_seances(par, code):
    ligne = par["lignes_seances"].get(code)
    return bool(ligne and ligne["actif"])


def groupes_autorises(par, code):
    """Groupes autorisés au moins pendant une partie de la période de vie
    de la ligne. Cette liste construit le roster du classeur ; la validation
    fine à la date d'une séance est appliquée par l'application."""
    ligne = par["lignes_seances"].get(code)
    if not ligne or not ligne["actif"]:
        return set()
    return {r["groupe"] for r in par["regles_acces_seances"]
            if r["ligne"] == code and r["actif"]
            and recouvrent(ligne["debut"], ligne["fin"],
                            r["debut"], r["fin"])}


def clear(ws, first, last, cols):
    for r in range(first, last + 1):
        for c in cols:
            ws.cell(r, c).value = None


def generer_occurrences(jour, debut, fin, periodicite=7, ancre=None):
    """Dates du créneau entre deux bornes incluses.

    L'ancre conserve la phase d'une alternance : avec 14 jours et une ancre
    au vendredi 18 septembre, le vendredi 25 septembre n'est pas produit.
    """
    if jour not in JOUR_INDEX or debut > fin:
        return []
    idx = JOUR_INDEX[jour]
    d = debut + datetime.timedelta(days=(idx - debut.weekday()) % 7)
    periodicite = periodicite_jours(periodicite)
    # Une date d'effet peut tomber un autre jour que celui du créneau (par
    # exemple STAC ouvre un lundi pour une séance du mardi). L'ancre de la
    # périodicité est alors la première occurrence du créneau, jamais la date
    # d'effet brute ; autrement une cadence de sept jours ne produit rien.
    if ancre:
        ancre = ancre + datetime.timedelta(days=(idx - ancre.weekday()) % 7)
    else:
        ancre = d
    out = []
    while d <= fin:
        if (d - ancre).days % periodicite == 0:
            out.append(d)
        d += datetime.timedelta(days=7)
    return out


def init_calendrier(wb, par, groupe):
    """Remplit le calendrier des séances à partir des créneaux du groupe et
    des dates de saison — seulement si l'onglet est encore vierge, pour ne
    jamais écraser une saisie réelle. Une seule ligne libre est laissée après
    la dernière séance générée, comme pour les autres tables."""
    ws = wb["Calendrier"]
    if ws.cell(CAL_FIRST, 2).value not in (None, ""):
        return None    # déjà initialisé, ou déjà saisi à la main
    if not (par["saison_debut"] and par["saison_fin"]):
        return None    # dates de saison pas encore renseignées au paramétrage

    par_creneau = {}
    for a in par["affectations"]:
        if a["affectation"] == groupe and a["date"]:
            # Plusieurs encadrants peuvent intervenir sur une même séance.
            # La séance ne doit être générée qu'une fois.
            par_creneau.setdefault(a["creneau"], set()).add(
                (a["date"], a["periodicite"]))

    occurrences = []
    for cid, configurations in par_creneau.items():
        cr = par["creneaux"].get(cid)
        if not cr:
            continue
        _cid, jour, hdebut, hfin, bassin = cr
        bornes = sorted(configurations)
        for i, (debut_effet, periodicite) in enumerate(bornes):
            fin_bloc = (bornes[i + 1][0] - datetime.timedelta(days=1)
                        if i + 1 < len(bornes) else par["saison_fin"])
            debut = max(debut_effet, par["saison_debut"])
            fin = min(fin_bloc, par["saison_fin"])
            # Une ligne de séances peut ouvrir ou fermer en cours de saison.
            # Les affectations de créneau restent néanmoins datées de façon
            # indépendante, pour pouvoir évoluer sans modifier la ligne.
            ligne = par["lignes_seances"].get(groupe)
            if ligne:
                if not ligne["actif"]:
                    continue
                if ligne["debut"]:
                    debut = max(debut, ligne["debut"])
                if ligne["fin"]:
                    fin = min(fin, ligne["fin"])
            label = "%s — %s %s-%s%s [%s]" % (
                cid, jour, hdebut or "?", hfin or "?",
                " · %s" % bassin if bassin else "", cid)
            for d in generer_occurrences(jour, debut, fin, periodicite,
                                         ancre=debut_effet):
                occurrences.append((d, label))
    if not occurrences:
        return None
    occurrences.sort(key=lambda x: x[0])

    last = CAL_FIRST + len(occurrences)   # + une ligne libre à la fin
    for i, (d, label) in enumerate(occurrences):
        row = CAL_FIRST + i
        write_calendrier_row(ws, row, groupe, statut="planifiée")
        c = ws.cell(row, 2, d)
        c.number_format = "DD/MM/YYYY"
        ws.cell(row, 4, label)
    write_calendrier_row(ws, last, groupe)   # ligne libre : formules seules

    if "Calendrier" in ws.tables:
        ws.tables["Calendrier"].ref = "A%d:U%d" % (CAL_FIRST - 1, last)
    ws.data_validations.dataValidation = []
    dv_range(ws, SEL_CR, "D%d:D%d" % (CAL_FIRST, last),
             "Choisissez le créneau. La liste vient du paramétrage du club.")
    dv_list(ws, NATURES, "H%d:H%d" % (CAL_FIRST, last))
    dv_list(ws, STATUTS_SEANCE, "I%d:I%d" % (CAL_FIRST, last))
    dv_range(ws, SEL_PERS, "J%d:J%d" % (CAL_FIRST, last), MSG_ENC)
    dv_range(ws, SEL_PERS, "N%d:N%d" % (CAL_FIRST, last),
             "À remplir seulement si le responsable de cette séance n'est "
             "pas l'encadrant habituel du créneau. Laissez vide sinon.")

    # La liste déroulante de séance dans Presences/Evenements_Securite pointe
    # vers Calendrier!$A$5:$A$5 telle que figée à la génération (build_suivi.py,
    # une seule séance de capacité initiale) : il faut l'étendre jusqu'à la
    # dernière séance réellement générée ci-dessus, sinon seule S001 reste
    # sélectionnable (06/09/2026).
    seance_range = "Calendrier!$A$%d:$A$%d" % (CAL_FIRST, last)
    for nom_onglet, col, first, dlast in (
            ("Presences", "A", PRE_FIRST, PRE_LAST),
            ("Evenements_Securite", "C", EV_FIRST, EV_LAST)):
        wsx = wb[nom_onglet]
        cible = "%s%d:%s%d" % (col, first, col, dlast)
        for dv in list(wsx.data_validations.dataValidation):
            if str(dv.sqref) == cible:
                wsx.data_validations.dataValidation.remove(dv)
        dv_range(wsx, "=%s" % seance_range, cible)

    return len(occurrences)


def sync_one(par, path, groupe, par_name):
    wb = openpyxl.load_workbook(path)          # formules conservées
    ref_groupe = "L4" if groupe == "LC" else groupe
    ligne_seances = est_ligne_seances(par, groupe)
    groupes_ligne = groupes_autorises(par, groupe) if ligne_seances else {groupe}
    # STAC est une ligne de séances à effectif nominatif. L'absence de règle
    # d'accès n'est donc pas une absence de membre : ses inscriptions portent
    # directement le code STAC et constituent la source de son roster.
    if ligne_seances and not groupes_ligne:
        groupes_ligne = {groupe}

    # ---------------------------------------------------- Referentiel
    ws = wb["Referentiel"]
    clear(ws, REF_OBJ_FIRST, REF_OBJ_LAST, range(1, 8))
    objs = [o for o in par["objectifs"] if o["groupe"] == ref_groupe]
    objs.sort(key=lambda o: (0 if o["type"] == "maîtrise" else 1, o["id"]))
    if len(objs) > REF_OBJ_LAST - REF_OBJ_FIRST + 1:
        raise SystemExit("trop d'objectifs pour %s : %d" % (groupe, len(objs)))
    for i, o in enumerate(objs):
        for c, v in enumerate([o["id"], o["type"], o["famille"], o["intitule"],
                               o["seuil"], o["intervalle"], o["critere"]], 1):
            ws.cell(REF_OBJ_FIRST + i, c, v)
    clear(ws, REF_CMP_FIRST, REF_CMP_LAST, range(1, 5))
    for i, c in enumerate(par["competences"]):
        att = par["attendus"].get((c["code"], ref_groupe), "")
        for col, v in enumerate([c["code"], c["libelle"], att, c["critere"]], 1):
            ws.cell(REF_CMP_FIRST + i, col, v)

    # ------------------------------------------------------ Personnes
    ws = wb["Personnes"]
    ordre_existant = []
    for r in range(PERS_FIRST, PERS_LAST + 1):
        pid = ws.cell(r, 2).value
        if pid:
            ordre_existant.append(pid)

    ins = {}
    for x in par["inscriptions"]:
        ins.setdefault(x["id"], []).append(x)

    # Une inscription avec une date de fin passée ne compte plus comme
    # active : la personne sort du roster de la ligne (RGM-07,
    # 1.5_regles_de_gestion_v0.md), sans que rien ne soit supprimé — son
    # historique de séances, suivis et avis déjà saisis reste inchangé
    # ailleurs dans le classeur.
    aujourdhui = datetime.date.today()

    def actives(pid):
        return [x for x in ins.get(pid, [])
               if x["fin"] is None or x["fin"] >= aujourdhui]

    def groupes(pid):
        return sorted({x["groupe"] for x in actives(pid) if x["groupe"]})

    # Une ligne par inscription active dans CE groupe, pas par personne : une
    # personne avec deux inscriptions dans la même ligne (élève ET encadrant,
    # ex. Sandry en L3) apparaît deux fois, une fois par rôle — jamais de
    # rôle agrégé du type « élève / encadrant » sur une même ligne (D-2026-09-06).
    inscriptions_groupe = [x for x in par["inscriptions"]
                           if x["groupe"] in groupes_ligne
                           and x["id"] in par["personnes"]
                           and (x["fin"] is None or x["fin"] >= aujourdhui)]
    par_cle = {}
    for x in inscriptions_groupe:
        par_cle.setdefault((x["id"], x["role"]), x)  # doublon : on garde le 1er

    ordre_existant = []
    for r in range(PERS_FIRST, PERS_LAST + 1):
        pid = ws.cell(r, 2).value
        role_c = ws.cell(r, 5).value
        if pid:
            ordre_existant.append((pid, role_c))

    cles_actuelles = set(par_cle)
    # ordre : inscriptions déjà connues (ordre préservé), puis les nouvelles
    connues = [k for k in ordre_existant if k in cles_actuelles]
    dejavu = set(connues)
    nouvelles = sorted([k for k in cles_actuelles if k not in dejavu],
                       key=lambda k: (str(par["personnes"][k[0]]["nom"]).upper(),
                                      str(par["personnes"][k[0]]["prenom"]).upper(),
                                      k[1] or ""))
    membres = connues + nouvelles
    membres_pids = {k[0] for k in membres}
    autres = [p for p in par["ordre_club"] if p not in membres_pids]
    if len(membres) > ROSTER_N:
        raise SystemExit("trop de membres pour %s : %d (max %d)"
                         % (groupe, len(membres), ROSTER_N))
    ordre = membres + [(p, None) for p in autres]
    clear(ws, PERS_FIRST, PERS_LAST, range(2, 8))
    clear(ws, PERS_FIRST, PERS_LAST, [16])
    if len(ordre) > PERS_LAST - PERS_FIRST + 1:
        ordre = ordre[:PERS_LAST - PERS_FIRST + 1]
    for i, (pid, role_c) in enumerate(ordre):
        p = par["personnes"][pid]
        r = PERS_FIRST + i
        x = par_cle.get((pid, role_c))
        role_val = x["role"] if x else ""
        objectif_val = x["objectif"] if x else ""
        for c, v in enumerate([pid, p["nom"], p["prenom"], role_val,
                               " / ".join(groupes(pid)), objectif_val], 2):
            ws.cell(r, c, v)
        ws.cell(r, 16, p.get("email", ""))

        # Dans une ligne de séances, l'appartenance est issue des règles
        # d'accès datées et non du texte « Groupe(s) » de la personne. On
        # inscrit donc explicitement le résultat de la synchronisation dans
        # la colonne Membre, afin que les codes et rosters du classeur restent
        # cohérents. Les groupes de niveau gardent leur formule historique.
        if ligne_seances:
            ws.cell(r, 8, "oui" if (pid, role_c) in par_cle else "")
    if ligne_seances:
        for r in range(PERS_FIRST + len(ordre), PERS_LAST + 1):
            ws.cell(r, 8, "")

    # --------------------------------------------------- Creneaux_Ref
    # Section 3 de l'onglet Referentiel (fusionné avec Responsables, §4).
    # L'encadrant (colonnes 8-10) est celui de la date d'effet la plus
    # récente dans Creneaux_Affectations : c'est lui qui alimente la colonne
    # Responsable du Calendrier, pas le référent pédagogique de la ligne.
    ws = wb["Referentiel"]
    clear(ws, CREF_FIRST, CREF_LAST, range(1, 11))
    ids = []
    for a in par["affectations"]:
        if a["affectation"] == groupe and a["creneau"] not in ids:
            ids.append(a["creneau"])
    crens = [par["creneaux"][c] + [groupe] for c in ids
             if c in par["creneaux"]]
    if len(crens) > CREF_LAST - CREF_FIRST + 1:
        raise SystemExit("trop de créneaux pour %s : %d" % (groupe, len(crens)))
    enc_par_creneau = {}
    for cid in ids:
        candidats = [a for a in par["affectations"]
                     if a["affectation"] == groupe and a["creneau"] == cid
                     and a["encadrant"]]
        referents = [a for a in candidats if a["referent"]]
        # Une ligne de séances peut avoir plusieurs co-encadrants, mais le
        # calendrier doit connaître un seul responsable habituel, notamment
        # pour le remplacement. Le référent est donc obligatoire dans ce cas.
        if ligne_seances and not referents:
            raise SystemExit("encadrant référent manquant pour %s / %s"
                             % (groupe, cid))
        choix = referents if referents else candidats
        if choix:
            enc_par_creneau[cid] = max(
                choix, key=lambda a: a["date"] or datetime.date.min)
    for i, cr in enumerate(crens):
        for c, v in enumerate(cr, 1):
            ws.cell(CREF_FIRST + i, c, v)
        enc = enc_par_creneau.get(cr[0])
        if enc:
            p = par["personnes"].get(enc["encadrant"], {})
            ws.cell(CREF_FIRST + i, 8, enc["encadrant"])
            ws.cell(CREF_FIRST + i, 9, p.get("nom", ""))
            ws.cell(CREF_FIRST + i, 10, p.get("prenom", ""))

    # -------------------------------------------------- Responsables
    clear(ws, RESP_FIRST, RESP_LAST, range(1, 6))
    resp = sorted([x for x in par["responsables"] if x["groupe"] == groupe],
                  key=lambda x: x["date"])
    if len(resp) > RESP_LAST - RESP_FIRST + 1:
        raise SystemExit("trop de responsables pour %s" % groupe)
    for i, x in enumerate(resp):
        r = RESP_FIRST + i
        for c, v in enumerate([x["date"], x["id"], x["nom"], x["prenom"],
                               x["commentaire"]], 1):
            ws.cell(r, c, v)
        ws.cell(r, 1).number_format = "DD/MM/YYYY"

    # ------------------------------------------- Qualifications Sécurité
    # Section 5 de l'onglet Referentiel : club entier, pas de notion de
    # ligne — alimente le contrôle « DP sans qualification Sécurité valide »
    # du Calendrier (1.8_matrice_rbac_v0.md F3, COMPTE-07/D-15).
    clear(ws, QUAL_FIRST, QUAL_LAST, range(1, 4))
    quals = par["dp_valides"]
    if len(quals) > QUAL_LAST - QUAL_FIRST + 1:
        raise SystemExit("trop de personnes qualifiées Sécurité : %d"
                         % len(quals))
    for i, x in enumerate(quals):
        for c, v in enumerate([x["id"], x["nom"], x["prenom"]], 1):
            ws.cell(QUAL_FIRST + i, c, v)

    # -------------------------------------------------------- Synchro
    ws = wb["Synchro"]
    vals = [groupe, SAISON, par_name, par["version"],
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            len(objs), len(par["competences"]), len(ordre),
            len(crens), len(resp)]
    for i, v in enumerate(vals):
        ws.cell(5 + i, 2, v)

    # ------------------------------------------------- Calendrier (init)
    seances = init_calendrier(wb, par, groupe)

    wb.save(path)
    stats = {"objectifs": len(objs), "personnes": len(ordre),
             "membres": len(membres), "créneaux": len(crens),
             "responsables": len(resp)}
    if seances is not None:
        stats["séances initialisées"] = seances
    return stats


def main():
    args = list(sys.argv[1:])
    folder = "."
    if args and os.path.isdir(args[0]):
        folder = args.pop(0)
    codes = args or [g for g, _ in D.LIGNES]

    par_files = glob.glob(os.path.join(folder, "AA - Parametrage *.xlsx"))
    if not par_files:
        raise SystemExit("classeur de paramétrage introuvable dans %s" % folder)
    par_path = sorted(par_files)[-1]
    par = read_parametrage(par_path)
    print("paramétrage lu : %s (version %s)"
          % (os.path.basename(par_path), par["version"]))
    print("  %d personnes · %d créneaux · %d objectifs · %d compétences · "
          "%d inscriptions · %d responsables · %d qualifiés Sécurité valides"
          % (len(par["personnes"]), len(par["creneaux"]), len(par["objectifs"]),
             len(par["competences"]), len(par["inscriptions"]),
             len(par["responsables"]), len(par["dp_valides"])))
    print("  %d associations créneau/affectation · %d lignes de séances · "
          "%d règles d'accès" % (len(par["affectations"]),
                                  len(par["lignes_seances"]),
                                  len(par["regles_acces_seances"])))

    for code in codes:
        path = os.path.join(folder, "AA - Suivi %s %s.xlsx" % (code, SAISON))
        if not os.path.exists(path):
            print("  ignoré : %s absent" % os.path.basename(path))
            continue
        stats = sync_one(par, path, code, os.path.basename(par_path))
        if os.path.exists(RECALC):
            subprocess.run(["python3", RECALC, path, "300"],
                           capture_output=True)
        print("  %s : %s" % (os.path.basename(path),
                             " · ".join("%s %s" % (v, k)
                                        for k, v in stats.items())))


if __name__ == "__main__":
    main()
