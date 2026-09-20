"""Contenu métier : référentiels, taxonomies, créneaux."""

SAISON = "2026-2027"

DISCIPLINES = [
    ("DYN", "Apnée dynamique avec palmes", 1, "oui",
     "Seule discipline retenue au MVP (décision D-23)."),
    ("DNF", "Apnée dynamique sans palme", 2, "non",
     "Hors MVP. Aucun plan d'entraînement dédié à ce jour."),
    ("STA", "Apnée statique", 3, "non",
     "Hors MVP. Aucun plan d'entraînement dédié à ce jour."),
    ("PRO", "Apnée en profondeur", 4, "non",
     "Hors MVP. Organisation en fosse et milieu naturel, à cadrer."),
]

GROUPES = [
    ("L1", "L1 — Initiation", "DYN", 1, "", "oui", "Débutants"),
    ("L2", "L2 — Intermédiaire", "DYN", 2, "L1", "oui", "Intermédiaires"),
    ("L3", "L3 — Confirmé", "DYN", 3, "L2", "oui", "Confirmés"),
    ("L4", "L4 — Expert", "DYN", 4, "L3", "oui", "Experts"),
    ("LC", "LC — Compétition", "DYN", 5, "L4", "oui",
     "Pas de plan propre : séances calquées sur le L4, logique de records."),
]

# Les lignes de séances ne sont pas des groupes d'inscription. Leur capacité,
# période d'effet et règles d'accès vivent dans les tables Lignes_Seances et
# Regles_Acces_Seances du paramétrage du club.
LIGNES_SEANCES = [
    ("DNF1", "Apnée dynamique sans palme 1"),
    ("DNF2", "Apnée dynamique sans palme 2"),
    ("STA1", "Apnée statique 1"),
    ("STA2", "Apnée statique 2"),
    ("STA3", "Apnée statique 3"),
    ("STAC", "Apnée statique compétition"),
]

FAMILLES = [
    ("DIST", "Distance maximale", "DYN", "mètre", "la distance",
     "Objectif structurant de chaque niveau."),
    ("SERIE_HC", "Série hypercapnique", "DYN", "temps de départ",
     "l'intervalle de départ, à 4 × 50 m constants", ""),
    ("SERIE_CT", "Série courte à récupération brève", "DYN",
     "nombre × distance + secondes",
     "le nombre de répétitions et la récupération", "Absente au-delà du L2."),
    ("STA_DYN", "Enchaînement statique-dynamique", "DYN", "temps + mètre",
     "la durée de statique et la distance",
     "Exercice de la dynamique, pas une discipline distincte."),
    ("LENT", "Nage lente à durée imposée", "DYN", "temps",
     "la durée imposée pour 50 m", ""),
    ("BRASSE", "Brasse (sans palme)", "DYN", "mètre", "la distance",
     "Exercice de la dynamique, pas une discipline distincte."),
    ("VOL_PUL", "Volume pulmonaire de départ", "DYN", "modalité",
     "le volume de départ (FRC, PV)", "Présente au L2 uniquement."),
]

# ID, groupe, famille, type, intitulé, distance, durée STA, temps cible,
# intervalle départ, récup, nb séries
OBJECTIFS = [
    ("L1-M1", "L1", "DIST", "maîtrise", "50m dyn bi-palmes", 50, "", "", "", "", ""),
    ("L1-M2", "L1", "SERIE_HC", "maîtrise", "4 x 50m départ 2'", 50, "", "", "2'", "", 4),
    ("L1-M3", "L1", "SERIE_CT", "maîtrise", "4 x 25m récup 15''", 25, "", "", "", "15''", 4),
    ("L1-M4", "L1", "STA_DYN", "maîtrise", "30'' de STA + 50m", 50, "30''", "", "", "", ""),
    ("L1-M5", "L1", "LENT", "maîtrise", "50m lent (1'15'')", 50, "", "1'15''", "", "", ""),
    ("L1-M6", "L1", "BRASSE", "maîtrise", "37,5m en brasse", 37.5, "", "", "", "", ""),
    ("L1-A1", "L1", "DIST", "acquisition", "65m dyn bi-palmes", 65, "", "", "", "", ""),
    ("L1-A2", "L1", "SERIE_HC", "acquisition", "4 x 50m départ 1'45''", 50, "", "", "1'45''", "", 4),
    ("L1-A3", "L1", "SERIE_CT", "acquisition", "8 x 25m récup 15''", 25, "", "", "", "15''", 8),
    ("L1-A4", "L1", "STA_DYN", "acquisition", "1' de STA + 50m", 50, "1'", "", "", "", ""),
    ("L1-A5", "L1", "LENT", "acquisition", "50m lent (1'30'')", 50, "", "1'30''", "", "", ""),
    ("L1-A6", "L1", "BRASSE", "acquisition", "50m en brasse", 50, "", "", "", "", ""),
    ("L2-M1", "L2", "DIST", "maîtrise", "75m dyn bi-palmes", 75, "", "", "", "", ""),
    ("L2-M2", "L2", "SERIE_HC", "maîtrise", "4 x 50m départ 1'30''", 50, "", "", "1'30''", "", 4),
    ("L2-M3", "L2", "SERIE_CT", "maîtrise", "4 x 25m récup 10''", 25, "", "", "", "10''", 4),
    ("L2-M4", "L2", "STA_DYN", "maîtrise", "1' de STA + 50m", 50, "1'", "", "", "", ""),
    ("L2-M5", "L2", "LENT", "maîtrise", "50m lent (1'30'')", 50, "", "1'30''", "", "", ""),
    ("L2-M6", "L2", "BRASSE", "maîtrise", "50m en brasse", 50, "", "", "", "", ""),
    ("L2-M7", "L2", "VOL_PUL", "maîtrise", "50m FRC", 50, "", "", "", "", ""),
    ("L2-A1", "L2", "DIST", "acquisition", "85m dyn bi-palmes", 85, "", "", "", "", ""),
    ("L2-A2", "L2", "SERIE_HC", "acquisition", "4 x 50m départ 1'20''", 50, "", "", "1'20''", "", 4),
    ("L2-A3", "L2", "SERIE_CT", "acquisition", "8 x 25m récup 10''", 25, "", "", "", "10''", 8),
    ("L2-A4", "L2", "STA_DYN", "acquisition", "30'' de STA + 75m", 75, "30''", "", "", "", ""),
    ("L2-A5", "L2", "LENT", "acquisition", "50m lent (1'45'')", 50, "", "1'45''", "", "", ""),
    ("L2-A6", "L2", "BRASSE", "acquisition", "65m en brasse", 65, "", "", "", "", ""),
    ("L2-A7", "L2", "VOL_PUL", "acquisition", "50m PV", 50, "", "", "", "", ""),
    ("L3-M1", "L3", "DIST", "maîtrise", "100m dyn", 100, "", "", "", "", ""),
    ("L3-M2", "L3", "SERIE_HC", "maîtrise", "4 x 50m départ 1'10''", 50, "", "", "1'10''", "", 4),
    ("L3-M3", "L3", "STA_DYN", "maîtrise", "1' de STA + 75m", 75, "1'", "", "", "", ""),
    ("L3-M4", "L3", "LENT", "maîtrise", "50m lent (2')", 50, "", "2'", "", "", ""),
    ("L3-M5", "L3", "BRASSE", "maîtrise", "65m en brasse", 65, "", "", "", "", ""),
    ("L3-A1", "L3", "DIST", "acquisition", "125m dyn", 125, "", "", "", "", ""),
    ("L3-A2", "L3", "SERIE_HC", "acquisition", "4 x 50m départ 1'", 50, "", "", "1'", "", 4),
    ("L3-A3", "L3", "STA_DYN", "acquisition", "1'30'' de STA + 75m", 75, "1'30''", "", "", "", ""),
    ("L3-A4", "L3", "LENT", "acquisition", "50m lent (2'30'')", 50, "", "2'30''", "", "", ""),
    ("L3-A5", "L3", "BRASSE", "acquisition", "75m en brasse", 75, "", "", "", "", ""),
    ("L4-M1", "L4", "DIST", "maîtrise", "125m dyn", 125, "", "", "", "", ""),
    ("L4-M2", "L4", "SERIE_HC", "maîtrise", "4 x 50m départ 1'", 50, "", "", "1'", "", 4),
    ("L4-M3", "L4", "STA_DYN", "maîtrise", "1'30 de STA + 75m", 75, "1'30", "", "", "", ""),
    ("L4-M4", "L4", "LENT", "maîtrise", "50m lent (2'30)", 50, "", "2'30", "", "", ""),
    ("L4-M5", "L4", "BRASSE", "maîtrise", "75m en brasse", 75, "", "", "", "", ""),
    ("L4-A1", "L4", "DIST", "acquisition", "150m dyn", 150, "", "", "", "", ""),
    ("L4-A2", "L4", "SERIE_HC", "acquisition", "4 x 50m départ 55''", 50, "", "", "55''", "", 4),
    ("L4-A3", "L4", "STA_DYN", "acquisition", "2' de STA + 75m", 75, "2'", "", "", "", ""),
    ("L4-A4", "L4", "LENT", "acquisition", "50m lent (3')", 50, "", "3'", "", "", ""),
    ("L4-A5", "L4", "BRASSE", "acquisition", "100m en brasse", 100, "", "", "", "", ""),
]

CONTR_HYPO = "hypercapnique et/ou hypoxique"
CONTR_OPTI = "meilleures conditions"
CONTR_AUCUNE = "aucune contrainte particulière"

CRITERES = {
    ("L1", "maîtrise"): (7, "7'", CONTR_HYPO,
        "Pouvoir réaliser la distance 7 fois dans la même séance avec "
        "contraintes hypercapniques et/ou hypoxiques."),
    ("L2", "maîtrise"): (5, "7'", CONTR_HYPO,
        "Pouvoir réaliser la distance 5 fois dans la même séance avec "
        "contraintes hypercapniques et/ou hypoxiques."),
    ("L3", "maîtrise"): (3, "7'", CONTR_HYPO,
        "Pouvoir réaliser la distance 3 fois dans la même séance avec "
        "contraintes hypercapniques et/ou hypoxiques."),
    ("L4", "maîtrise"): ("", "", CONTR_AUCUNE,
        "Pouvoir réaliser la distance sans contrainte particulière, "
        "régulièrement, sans incident. — critère qualitatif, non calculable."),
    ("L1", "acquisition"): (1, "", CONTR_OPTI,
        "Pouvoir réaliser la distance ou l'exercice une fois en étant mis "
        "dans les meilleures conditions."),
    ("L2", "acquisition"): (1, "", CONTR_OPTI,
        "Pouvoir réaliser la distance ou l'exercice une fois en étant mis "
        "dans les meilleures conditions."),
    ("L3", "acquisition"): (1, "", CONTR_OPTI,
        "Pouvoir réaliser la distance ou l'exercice une fois en étant mis "
        "dans les meilleures conditions."),
    ("L4", "acquisition"): (1, "", CONTR_OPTI,
        "Pouvoir réaliser la distance ou l'exercice une fois de temps en "
        "temps en étant mis dans les meilleures conditions."),
}

SOURCES = {
    "L1": "AA - Plan d'entrainement débutants (L1) V.2.2024, p. 3",
    "L2": "AA - Plan d'entrainement intermédiaires (L2) V.2.2024, p. 3",
    "L3": "AA - Plan d'entrainement confirmés (L3) V.2.2024, p. 3",
    "L4": "AA - Plan d'entrainement experts (L4) V.1.2025, p. 3",
}

COMPETENCES = [
    ("LEST", "Lestage",
     "Assiette horizontale tenue sans compensation musculaire ; flottabilité "
     "neutre en fin d'apnée.",
     "L2 : « il apporte l'hydrodynamisme en corrigeant l'assiette »"),
    ("HYDRO", "Hydrodynamisme",
     "Position hydrodynamique conservée sur toute la distance ; aucun "
     "mouvement parasite.", "L2 / L3 : « hydrodynamisme optimal »"),
    ("PALM_BI", "Palmage bi-palmes",
     "Ondulation propre, amplitude régulière, propulsion issue des hanches "
     "et non des cuisses.",
     "L3 : facteur limitant « propulsion, sursollicitation des cuisses »"),
    ("PALM_MONO", "Palmage monopalme",
     "Ondulation continue tête-bassin-palme, bras en flèche tenus.",
     "À rédiger"),
    ("VIRAGE", "Virage",
     "Virage fluide, sans perte de vitesse notable, sans apnée "
     "supplémentaire ; poussée d'au moins 10 m.",
     "L3 : « l'objectif : faire 10 m »"),
    ("VENTIL", "Ventilation",
     "Préparation ventilatoire abaissant la fréquence cardiaque, dernière "
     "inspiration maîtrisée, protocole de sécurité en sortie, récupération "
     "codée respectée.",
     "L2 : « préparation ventilatoire / dernière inspiration / protocole »"),
    ("SAUV", "Sauvetage",
     "Sait repérer les signaux hypoxiques chez son binôme et conduire une "
     "remontée et un protocole de sauvetage.",
     "L1 : « savoir pratiquer un sauvetage »"),
    ("BRASSE", "Brasse DNF",
     "Coordination traction-glisse-ondulation, temps de glisse tenu, "
     "trajectoire rectiligne.", "À rédiger"),
]

ATTENDU = {
    "LEST":      {"L1": "en cours",   "L2": "acquis",     "L3": "acquis",   "L4": "expert", "LC": "expert"},
    "HYDRO":     {"L1": "en cours",   "L2": "acquis",     "L3": "acquis",   "L4": "expert", "LC": "expert"},
    "PALM_BI":   {"L1": "en cours",   "L2": "acquis",     "L3": "acquis",   "L4": "expert", "LC": "expert"},
    "PALM_MONO": {"L1": "non acquis", "L2": "non acquis", "L3": "en cours", "L4": "acquis", "LC": "expert"},
    "VIRAGE":    {"L1": "en cours",   "L2": "acquis",     "L3": "acquis",   "L4": "expert", "LC": "expert"},
    "VENTIL":    {"L1": "en cours",   "L2": "en cours",   "L3": "acquis",   "L4": "expert", "LC": "expert"},
    "SAUV":      {"L1": "non acquis", "L2": "en cours",   "L3": "acquis",   "L4": "expert", "LC": "expert"},
    "BRASSE":    {"L1": "en cours",   "L2": "acquis",     "L3": "acquis",   "L4": "expert", "LC": "expert"},
}

QUALIFS = [
    ("AP", "Apnéiste Piscine", "Niveau apnée", "", "à valider"),
    ("ACP", "Apnéiste Confirmé Piscine", "Niveau apnée", "", "à valider"),
    ("AEL", "Apnéiste Eau Libre", "Niveau apnée", "", "à valider"),
    ("ACEL", "Apnéiste Confirmé Eau Libre", "Niveau apnée", "", "à valider"),
    ("AEEL", "Apnéiste Expert Eau Libre", "Niveau apnée", "", "à valider"),
    ("IE", "Initiateur Entraîneur", "Encadrement", "", "à valider"),
    ("MEF1", "Moniteur Entraîneur Fédéral 1er degré", "Encadrement", "", "à valider"),
    ("MEF2", "Moniteur Entraîneur Fédéral 2e degré", "Encadrement", "", "à valider"),
    ("RIFAA", "Réactions et Interventions Face à un Accident d'Apnée",
     "Sécurité", "recyclage recommandé", "à valider"),
    ("PSM", "PSM — libellé à préciser", "Sécurité", "", "à préciser"),
]

BASSINS = [
    ("B50", "Grand bassin 50 m", 50, 8, "Configuration nominale de la saison."),
    ("B25", "Petit bassin 25 m", 25, 3,
     "Utilisé en été et sur certains créneaux ; « tout le bassin » = 3 lignes."),
]

REGIMES = [
    ("NOMINAL", "Régime nominal de saison", "Semaine type de septembre à juin."),
    ("VACANCES", "Vacances scolaires", "Horaires et lignes réduits."),
    ("ETE", "Régime d'été", "Organisation dérogatoire de juillet et août."),
    ("FERIE", "Jour férié", "Aucune séance."),
    ("FERMETURE", "Fermeture technique", "Piscine indisponible."),
]

PERIODES = [
    ("NOMINAL", "2026-09-07", "2027-07-04", "Saison 2026-2027", "à ajuster"),
    ("VACANCES", "", "", "Vacances de la Toussaint", "dates à saisir"),
    ("VACANCES", "", "", "Vacances de Noël", "dates à saisir"),
    ("VACANCES", "", "", "Vacances d'hiver", "dates à saisir"),
    ("VACANCES", "", "", "Vacances de printemps", "dates à saisir"),
    ("ETE", "", "", "Été 2027", "dates à saisir"),
]

# ID, régime, jour, h début, h fin, bassin, groupe, nb lignes, actif, commentaire
CRENEAUX = [
    ("CR-LUN", "NOMINAL", "lundi", "", "", "B50", "", 1, "oui",
     "Jour confirmé par le planning 2026-27. Horaires et groupe à saisir."),
    ("CR-MAR", "NOMINAL", "mardi", "", "", "B50", "", 1, "oui",
     "Jour confirmé par le planning 2026-27. Horaires et groupe à saisir."),
    ("CR-MER", "NOMINAL", "mercredi", "", "", "B50", "", 1, "oui",
     "Jour confirmé par le planning 2026-27. Horaires et groupe à saisir."),
    ("CR-VEN", "NOMINAL", "vendredi", "", "", "B50", "", 1, "oui",
     "Jour confirmé par le planning 2026-27. Horaires et groupe à saisir."),
    ("CR-SAM", "NOMINAL", "samedi", "", "", "B50", "", 1, "oui",
     "Jour confirmé par le planning 2026-27. Horaires et groupe à saisir."),
    ("CR-DIM", "NOMINAL", "dimanche", "", "", "B50", "", 1, "non",
     "Colonne présente dans le planning mais sans directeur de bassin."),
]

AFFECT = [
    ("CR-LUN", "directeur de bassin", "toutes les semaines",
     "Planning 2026-27 : Etienne puis François en alternance — à trancher."),
    ("CR-MAR", "directeur de bassin", "toutes les semaines",
     "Planning 2026-27 : Sandry sur les 4 premières semaines."),
    ("CR-MER", "directeur de bassin", "toutes les semaines",
     "Planning 2026-27 : PA sur les 4 premières semaines."),
    ("CR-VEN", "directeur de bassin", "toutes les semaines",
     "Planning 2026-27 : François sur les 4 premières semaines."),
    ("CR-SAM", "directeur de bassin", "toutes les semaines",
     "Planning 2026-27 : Danny à partir de la 2e semaine."),
]

# lignes d'eau à générer : (code groupe, libellé)
LIGNES = [("L1", "L1 — Initiation"), ("L2", "L2 — Intermédiaire"),
          ("L3", "L3 — Confirmé"), ("L4", "L4 — Expert"),
          ("LC", "LC — Compétition")] + LIGNES_SEANCES
