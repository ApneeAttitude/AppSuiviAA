# Proposition — listes de valeurs sous forme ID + libellé

Statut : **proposition, rien n'a été implémenté**. Ce document répond à la demande de Fred du 09/09/2026 : identifier les cas concernés, valider le modèle de données, proposer une ergonomie côté appli. Il attend sa validation avant tout développement.

---

## 1. Inventaire des cas identifiés

Toutes les listes de valeurs actuelles du classeur sont déjà centralisées dans un seul fichier Python, `generateurs/aa_common.py`, qui les fournit à `build_suivi.py` (génération des classeurs PROD par ligne) et `build_parametrage.py` (génération du classeur de Paramétrage). Aujourd'hui, chaque liste y est une simple liste de chaînes (les libellés seuls, sans identifiant), posée en liste déroulante Excel **à valeur littérale** — la cellule stocke directement le texte affiché.

Il existe même déjà un onglet technique nommé **"Listes"**, présent (a) dans le classeur de Paramétrage avec 19 listes, et (b) en version réduite (9 listes) dans chacun des 7 classeurs PROD — mais dans les deux cas il ne contient que les libellés, pas d'identifiant, et il n'est **consulté par aucune formule ni par `Code.gs`** : c'est un onglet de référence pour les listes déroulantes Excel/Sheets, découplé du fonctionnement réel de l'appli.

Je distingue trois catégories, selon leur usage réel aujourd'hui :

### 1.1 Utilisé aujourd'hui par l'appli (le seul cas actif)

| Liste | Valeurs actuelles | Où | Remarque |
|---|---|---|---|
| **Zone de confort** | Hors zone / En limite / Dans la zone | `Presences` colonne F, saisi depuis l'écran de saisie de présence | C'est l'exemple que tu donnes. Le libellé est aujourd'hui codé en dur **7 fois** (une fois par ligne, dans chaque `index.html`), en plus d'être dans `aa_common.py` et dans l'onglet Listes de chaque classeur — trois copies indépendantes du même texte. |

### 1.2 Présent dans les classeurs, pas encore exposé par l'appli

Ces colonnes existent déjà (structure posée par `build_suivi.py`), `Code.gs` les lit parfois, mais rien dans l'écran de saisie ne les affiche ou ne les modifie aujourd'hui — elles sont éditées à la main dans le Sheet. Même mécanique, même bénéfice à en faire des ID+libellé, mais sans impact visible sur l'appli tant qu'elles n'y sont pas branchées :

| Liste | Valeurs actuelles | Où |
|---|---|---|
| **Statut de séance** | planifiée / tenue / annulée / fermée | `Calendrier` colonne I — lu par `Code.gs` (`statut: row[8] \|\| 'planifiée'`), jamais affiché |
| **Nature de séance** | ordinaire, technique, hypercapnique, mixte, hypoxique, test maximal, RIFAA, PSM, examen, formation, tutorat, multi-lignes | `Calendrier` colonne H — pas lu par `Code.gs` actuellement |
| **Rôle de personne** | élève / encadrant / élève et encadrant | `Personnes` colonne E — lu par `Code.gs` (`role: ...toLowerCase()`), jamais affiché ni modifiable depuis l'appli |

### 1.3 Prévu uniquement pour le futur module Suivi v2 (pas encore de colonne réelle)

Ces listes existent déjà dans `aa_common.py` et dans l'onglet Listes, en anticipation du cahier des charges v2 (voir `analyse_excel_suivi_apneistes.md`), mais aucune colonne du classeur actuel ne les porte encore — donc pas de migration de données à faire pour elles, seulement à garder le même principe en tête quand ce module sera construit :

- **Statut d'acquisition** (non tenté / non réussi / réussi) et **Niveaux de compétence** (non acquis / en cours / acquis / expert) — futur suivi des objectifs et compétences
- **Avis d'encadrant** (absentéiste / dernier tiers / dans la moyenne / premier tiers / prêt pour le niveau supérieur) — remplace le codage couleur actuel du classeur Excel historique
- **Types d'événement** (PCM, samba, syncope, arrêt de séance, malaise, barotraumatisme, autre) et **Gravités** (mineur / significatif / grave) — futur registre d'événements de sécurité

### 1.4 Côté Paramétrage uniquement (rarement modifiées, hors écran de saisie)

Pour mémoire, des listes similaires existent aussi dans le classeur de Paramétrage (Genres, Statuts de personne, Motifs d'absence, Motifs d'indisponibilité, Rôles d'encadrement, Supports de saisie, Périodicités, Jours, Types d'objectif, Contraintes de validation). Même mécanique possible, mais ces colonnes sont éditées ponctuellement à la main par toi dans le Paramétrage, jamais via l'appli — donc priorité plus faible. Je ne les détaille pas ici sauf si tu veux les inclure.

**Constat annexe** (sans rapport avec ta demande, je le signale au passage) : la liste "Rôle" utilisée dans l'onglet `Inscriptions` du Paramétrage ne propose que 2 valeurs (élève/encadrant) alors que `ROLES_PERS` en propose 3 (élève, encadrant, élève **et** encadrant) — une incohérence pré-existante, à corriger un jour si besoin, indépendamment du sujet ID/libellé.

---

## 2. Modèle de données — validation de ta proposition

Ta proposition (ID entier + libellé, l'ID stocké dans la donnée, le libellé résolu par formule) est saine et correspond exactement au pattern qu'un tableur permet bien de faire. Je la valide, avec quatre ajustements pratiques :

**a) Où stocker les tables ID+libellé.** L'onglet "Listes" existe déjà, une colonne par liste (aujourd'hui, seulement le libellé). Le plus simple, sans changer l'organisation actuelle, est de faire de chaque liste **un couple de colonnes adjacentes** (ID | Libellé) plutôt qu'une colonne unique — donc un léger réaménagement de cet onglet, pas une refonte. Je recommande de garder cette table à la fois dans le Paramétrage (source de référence, que tu édites) et une copie dans chaque classeur de ligne (comme aujourd'hui), avec une synchronisation identique à celle déjà en place pour les autres colonnes "recopiées automatiquement" (bleu pâle) — pour que la formule de résolution du libellé dans `Presences` reste locale au classeur, sans dépendre d'un autre classeur ouvert.

**b) La formule de résolution.** Dans `Presences`, la colonne F contiendrait l'ID (1, 2 ou 3) au lieu du libellé ; une formule dans une colonne dédiée (par exemple juste à côté, en `INDEX`/`MATCH` ou `VLOOKUP` sur la table `Listes` du même classeur) afficherait le libellé correspondant — exactement ta proposition. Cette formule sert surtout à la lecture humaine directe du Sheet (un encadrant qui ouvre le classeur sans passer par l'appli) ; côté appli, `Code.gs` peut résoudre le libellé lui-même à partir de la même table, sans dépendre de la formule.

**c) Stabilité des ID — la règle à respecter.** Une fois un ID attribué à une valeur, il ne doit **jamais être réutilisé pour autre chose**, même si la valeur d'origine devient obsolète — sinon d'anciennes saisies de présence changeraient de sens rétroactivement. Je propose d'ajouter une colonne **"Actif" (oui/non)** à chaque table ID+libellé : pour retirer une valeur du menu déroulant sans casser les présences déjà enregistrées avec cet ID, on la marque inactive plutôt que de supprimer la ligne. L'appli n'affiche alors que les valeurs actives dans le menu déroulant, mais reste capable d'afficher correctement le libellé d'une ancienne saisie qui pointe vers une valeur devenue inactive.

**d) Portée de l'ID.** L'ID n'a besoin d'être unique qu'à l'intérieur d'une même liste (l'ID 1 de "Zone de confort" n'a rien à voir avec l'ID 1 de "Statut de séance") — pas de table d'ID globale nécessaire, chaque colonne du classeur sait déjà à quelle liste elle se réfère.

Avec ces quatre précisions, ta proposition est confirmée : simple, réversible, et elle règle exactement le problème que tu poses (changer un libellé sans toucher aux données déjà saisies).

---

## 3. Ergonomie proposée pour l'écran de saisie de présence

Côté utilisateur (l'encadrant qui saisit les présences), **rien ne change visuellement** : le menu déroulant "Zone de confort" affiche toujours les libellés lisibles (Hors zone / En limite / Dans la zone), dans le même ordre qu'aujourd'hui — jamais l'ID brut à l'écran, qui n'aurait aucun sens pour l'utilisateur.

Ce qui change, en coulisses :

- Le menu déroulant n'est plus codé en dur dans chaque `index.html` (7 copies identiques aujourd'hui) : il est construit dynamiquement à partir d'une liste ID+libellé fournie par `Code.gs` (par exemple ajoutée à la réponse de l'appel `action=data`, déjà utilisé pour charger la séance). Une seule source, plus de duplication entre les 7 lignes.
- Quand l'encadrant sélectionne une valeur, c'est l'**ID** qui est envoyé et stocké lors de l'enregistrement — pas le texte du libellé.
- Si tu changes un libellé dans la table de référence (ex. "En limite" → "Zone limite"), le changement apparaît **immédiatement** dans les 7 lignes et sur toutes les présences déjà enregistrées, sans redéploiement ni retouche des fichiers front-end.
- L'ordre d'affichage des options suit l'ordre des lignes dans la table de référence (donc facilement réordonnable par toi dans le Sheet, sans dépendre de l'ordre alphabétique ni de la valeur numérique de l'ID).
- Seules les valeurs marquées "Actif = oui" apparaissent dans le menu déroulant ; une ancienne présence référençant une valeur retirée continue de s'afficher correctement (avec son libellé d'origine), simplement non resélectionnable pour une nouvelle saisie.

Aucun autre changement d'écran n'est nécessaire : pas de nouveau champ, pas de nouvelle étape de saisie.

---

## 4. Points à valider avant toute implémentation

1. Périmètre : uniquement "Zone de confort" pour l'instant (le seul cas actif dans l'appli), ou veux-tu qu'on prépare en même temps "Statut de séance", "Nature de séance" et "Rôle de personne" (§1.2), même s'ils ne sont pas encore affichés dans l'appli ?
2. Migration : les valeurs déjà saisies dans `Presences` colonne F (texte libre aujourd'hui) devront être converties en ID lors du passage — à faire sur PROD (7 classeurs) et sur les copies TEST-L2/TEST-L3. Es-tu d'accord pour qu'on fasse cette conversion automatiquement (correspondance texte → ID à partir du libellé actuel), avec vérification avant/après ?
3. Confirmes-tu la colonne "Actif" proposée au §2c, ou préfères-tu une autre façon de gérer le retrait d'une valeur ?
4. Le format de l'ID : nombre entier simple (1, 2, 3…) comme tu le proposes — je confirme que c'est le plus simple, sauf si tu préfères un code court plus lisible dans le Sheet brut (ex. `HZ`, `EL`, `DZ`) plutôt qu'un chiffre.

Dès que tu valides (ou ajustes) ces points, je peux démarrer l'implémentation en environnement de test, comme pour les évolutions précédentes.
