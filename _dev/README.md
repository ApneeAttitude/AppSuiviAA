# MVP de secours — suivi des présences

Stopgap pour le 07/09/2026, hors périmètre du cadrage principal (`1.1_prompt_maitre.md` §5).
Ne remplace pas l'architecture cible (FastAPI/PostgreSQL/Vue3) — voir `1.7_exigences_non_fonctionnelles_v0.md`.

## Environnements

Un **seul** projet Apps Script **standalone** (un seul `Code.gs`, un seul
codebase à maintenir — décision du 07/09/2026), mais **trois
déploiements distincts** de ce même projet, un par palier — c'est ainsi
qu'on organise DEV/TEST/PROD dans script.google.com : pas trois projets,
un seul projet avec plusieurs entrées dans **Déployer ▸ Gérer les
déploiements**. Chaque déploiement a sa propre URL `.../exec`, et peut
être figé sur une version différente du code (voir « Promotion d'une
version » plus bas) — c'est ce qui isole PROD des changements en cours de
test en DEV, sans dupliquer le code. Chaque page `index.html` envoie aussi
une **cible** (ex. `"DEV-L2"`, `"PROD-L3"`) qui dit au script quel classeur
Google Sheet ouvrir, à l'intérieur du palier déjà déterminé par l'URL
appelée.

| Palier | Dossiers de ce repo | Déploiement Apps Script | Qui l'utilise |
|---|---|---|---|
| **DEV** | `dev/L2/`, `dev/L3/` | 1 déploiement "DEV", figé sur la dernière version (ou « Nouvelle version » à chaque itération) | Moi (Claude) + toi |
| **TEST** | `test/L2/`, `test/L3/` | 1 déploiement "TEST", figé sur une version validée | 1-2 encadrants volontaires |
| **PROD** | `L1/` … `STA/` | 1 déploiement "PROD", figé sur une version validée en TEST | Tous les encadrants |

Chaque déploiement a sa propre URL, donc son propre `WEB_APP_URL` :
identique entre `dev/L2/index.html` et `dev/L3/index.html` (même
déploiement DEV), mais différent de celui de `test/L2/index.html` ou de
`L3/index.html` (déploiement PROD).

**Le paramétrage cible → classeur se trouve dans `Code.gs`, objet
`CLASSEURS`** — c'est le seul et unique endroit à modifier pour brancher un
nouveau classeur (y coller son identifiant Google Sheet, visible dans son
URL : `.../spreadsheets/d/`**`IDENTIFIANT`**`/edit`). Ce fichier étant
partagé par les 3 déploiements, y ajouter une entrée pour TEST ou PROD
n'affecte PAS automatiquement les autres paliers tant qu'on ne fait pas
« Nouvelle version ▸ Déployer » sur LEUR déploiement respectif.

Depuis le 13/09/2026, les lignes DNF de test utilisent les cibles
`TEST-DNF1` et `TEST-DNF2`, chacune reliée à son propre classeur Google
Sheet et à une page sous `_test/DNF1/` ou `_test/DNF2/`. Elles utilisent
le déploiement Apps Script TEST ; elles ne doivent jamais pointer vers une
cible `PROD-*`. Avant leur première utilisation, lancer une fois `preparerDnf1Test` puis `preparerDnf2Test` dans l’éditeur Apps Script : ces migrations créent les tables ID/libellé de Zone de confort et Statut de séance sans modifier les autres lignes.

Depuis le 14/09/2026, une affectation de `Creneaux_Affectations` peut porter
une **Périodicité (jours)**. Une cellule vide signifie une séance chaque
semaine ; la valeur `14` génère une séance sur deux, ancrée sur la date
d’effet. Les co-encadrants d’une même affectation n’ajoutent pas de séance
supplémentaire. Cette règle est utilisée par `sync_referentiel.py` uniquement
lors de l’initialisation d’un calendrier encore vide. Les copies Google Sheets
TEST ont été créées le 14/09/2026 : STA1
(`1QGzitEk7jM7undte_y_EmaQYhsK_faKcFHTyni9folo`) et STA2
(`1PwMAxWjeOEE8Qpe0DRNLrVwv-KqmHZMlsTqFKloqBPg`). Leurs calendriers ont été
contrôlés avant toute configuration Apps Script. Le même jour, les cibles isolées
`TEST-STA1` et `TEST-STA2` ont été ajoutées dans `CLASSEURS`, puis les migrations
manuelles `preparerSta1Test` et `preparerSta2Test` ont été exécutées. Elles ne
peuvent agir que sur ces deux cibles TEST et créent les listes normalisées de
Zone de confort et Statut de séance. Le déploiement Apps Script TEST a ensuite
été mis à jour en **version 47** ; le déploiement PROD reste en version 46.
Les appels de lecture confirment STA1 (François Memheld, 18/09/2026) et STA2
(Guillaume Boulant, 25/09/2026), avec les listes normalisées attendues. Les pages `_test/STA1/` et `_test/STA2/` sont publiées et reliées au
déploiement TEST (commit `c8cdda5`) ; leur validation fonctionnelle a été
confirmée le 15/09/2026. Toute préparation puis publication en production
doit rester une opération séparée et explicitement validée.


Les copies de production STA1 et STA2 possèdent depuis le 15/09/2026 les
cibles distinctes `PROD-STA1` et `PROD-STA2`. Elles sont des copies natives
séparées des classeurs TEST, placées dans leurs dossiers Drive respectifs et
partagées avec les mêmes 24 personnes. Les migrations manuelles
`preparerSta1Prod` et `preparerSta2Prod` sont strictement limitées à ces deux
cibles : elles créent les listes normalisées de Zone de confort et de Statut de
séance. La version Apps Script PROD 48 a été contrôlée par lecture API avant la
publication des pages `STA1/` et `STA2/` (commit `d87d178`) ; les deux URLs
GitHub Pages ont ensuite été vérifiées.

Les copies de production DNF1 et DNF2 possèdent depuis le 13/09/2026 les
cibles distinctes `PROD-DNF1` et `PROD-DNF2`. Avant la première publication,
lancer une fois `preparerDnf1Prod`, puis `preparerDnf2Prod` : ces fonctions
sont strictement limitées à ces deux cibles et créent les listes normalisées de
Zone de confort et de Statut de séance. Elles ont été publiées avec les pages
`DNF1/` et `DNF2/` le 13/09/2026, via Apps Script PROD version 46. Pour une
nouvelle ligne, conserver la même séquence : migration du classeur, page qui
pointe vers la cible `PROD-*`, puis promotion explicite d’une version validée
vers le déploiement PROD.

### Promotion d'une version (DEV → TEST → PROD)

1. Modifier `Code.gs`, enregistrer (Cmd+S) — ça crée une nouvelle version
   disponible, sans rien déployer nulle part.
2. **Déployer ▸ Gérer les déploiements ▸ déploiement "DEV" ▸ crayon ▸
   Version : Nouvelle version ▸ Déployer.** DEV reflète le changement,
   TEST et PROD restent sur leur version précédente.
3. Une fois validé en DEV : **déploiement "TEST" ▸ crayon ▸ Version :**
   choisir **le même numéro de version** que celui utilisé par DEV (pas
   « Nouvelle version », qui prendrait la toute dernière — ici on veut
   figer TEST sur EXACTEMENT ce qui a été validé) **▸ Déployer.**
4. Une fois validé par un encadrant en TEST : répéter à l'identique sur le
   déploiement **"PROD"**.

Coût accepté pour ce choix : `SpreadsheetApp.openById()` (script
standalone) est mesuré 1 à 2,5 s par appel, contre quasi gratuit pour un
script *lié* à un seul classeur (l'architecture d'origine de ce MVP,
abandonnée ici au profit d'un développement unique).

OAuth : **un seul** Client ID Google partagé entre tous les paliers et
toutes les lignes (un seul écran de consentement à maintenir). Comme les
pages sont toutes hébergées sur le même domaine GitHub Pages, l'origine
JavaScript autorisée est la même quel que soit le chemin (`/dev/L2/`,
`/L1/`, ...) — pas besoin d'ajouter une origine par palier ou par ligne.

### Page unique TEST

Depuis le 19/09/2026, `_test/index.html` constitue la première version de
l'application à URL unique. Le choix d'une ligne se fait avec le paramètre
`?ligne=L2`, `?ligne=DNF1`, `?ligne=DNF2`, `?ligne=STA1` ou `?ligne=STA2` ;
l'absence de paramètre ouvre L2. La page ne référence que les cibles TEST
isolées correspondantes. L1, L3, L4, LC et STAC restent visibles dans le menu
mais désactivées tant que leur classeur TEST n'existe pas. Elle ne doit jamais
être étendue à une cible `PROD-*` pour des essais. Après connexion, elle
affiche la séance du jour, sinon la prochaine séance. Le menu hamburger ouvre
la sélection de ligne, les statistiques de la ligne, les statistiques
générales et le guide. La palette retenue est **Océan profond**. Les anciennes
pages `_test/L2/`, `_test/DNF1/`, etc. restent disponibles pendant la
validation ; leurs redirections vers la page unique seront ajoutées après test
fonctionnel.

La feuille de style commune est `assets/appsuivi.css` (19/09/2026). La page
unique TEST la charge directement. Les anciennes pages conservent encore leur
style embarqué afin de ne pas modifier leur apparence pendant la validation ;
elles basculeront vers cette feuille lors de leur remplacement par l'URL
unique. Toute nouvelle page AppSuiviAA doit charger cette feuille au lieu de
créer une nouvelle palette locale. Le sélecteur de séance de la page unique
reste un sélecteur natif afin d'être fiable sur les téléphones ; sa carte et
son chevron explicitent l'ouverture de la liste sans ajouter de comportement
spécifique au navigateur.

### Page unique PROD

Après validation du parcours TEST, `index.html` fournit l'URL unique de
production : `https://apneeattitude.github.io/AppSuiviAA/?ligne=L2`. Le même
paramètre ouvre L1, L2, L3, L4, LC, DNF1, DNF2, STA1 ou STA2 ; STAC demeure
visible mais indisponible tant que sa ligne n'est pas créée. Les anciennes URL
par ligne sont maintenues pendant la transition. La page `statistiques/` est
la vue globale de production et n'est pas mentionnée dans le guide.

Depuis la version Apps Script PROD 61, le même contrôle est actif en
production. Pour les lignes DNF et STA, les encadrants sont lus dans
`Creneaux_Affectations` afin de compléter les rôles présents dans
`Inscriptions`.

Depuis la version Apps Script TEST 59, les statistiques de ligne de la page
unique sont demandées par `POST` avec le jeton Google. Le serveur lit le
courriel associé à l'ID dans `AA - Parametrage 2026-2027` / `Personnes`,
colonne F, puis autorise seulement un rôle actif `encadrant` ou `prépa
encadrant` inscrit sur la ligne demandée. La liste des statistiques générales
reste séparée et réservée à Frédéric ainsi qu'aux responsables explicitement
configurés.

## PROD : un dossier Drive par ligne, avec accès dédié

Chaque ligne a son propre classeur, dans son propre dossier Google Drive,
partagé uniquement avec ses encadrants (+ le responsable de ligne). Un
encadrant de L2 n'a ainsi pas accès aux classeurs des autres lignes.

Arborescence Drive à créer (une fois) :

```
Apnée Attitude — Suivi des présences/
├── L1/   → AA - Suivi L1 2026-2027
├── L2/   → AA - Suivi L2 2026-2027
├── L3/   → AA - Suivi L3 2026-2027
├── L4/   → AA - Suivi L4 2026-2027
├── LC/   → AA - Suivi LC 2026-2027
├── DNF/  → AA - Suivi DNF 2026-2027
└── STA/  → AA - Suivi STA 2026-2027
```

Pour chaque dossier : clic droit → Partager → ajouter les emails ci-dessous
en **Lecteur** (le web app écrit via le script, exécuté en tant que
propriétaire — pas besoin d'accès Éditeur pour que l'app fonctionne ; donner
Éditeur seulement à qui doit pouvoir corriger le classeur à la main).

### Accès par ligne (encadrants actifs, emails déjà connus)

**L1** — Dominique Beauvallet (dombeauvallet@free.fr), Victorine Chardonnet
(chardonnetvictorine@gmail.com), Michael Grenier (micgren@gmail.com)

**L2** — Charline Danseux (charline.danseux@gmail.com), Cyrille Hannou
(hannou_c@yahoo.fr), Nelly Jacquemot (nellyjacquemot@yahoo.fr), Sophie
Landeau (sophiemattioli@yahoo.fr), Edward Lichtner
(edwardlichtner515@gmail.com), François Memheld
(francois.memheld@orange.fr)

**L3** — Alain Guevel (aguevel94@gmail.com), Etienne Colin de Verdiere
(e.cdv@free.fr), Gabrielle Beyl (g.beyl@outlook.fr), Sandry Wallon
(sandry.wallon@gmail.com), Fred Le Brigand (flebrigand@gmail.com)

**L4** — Philippe Fredon (philippe.fredon@orange.fr), Anthony Guglielmo
(anthony.guglielmo.fra@gmail.com), Pierre-Alain Hoyau
(pierre-alain.hoyau@univ-eiffel.fr), David Inserguet
(valdav92@gmail.com), Céline Dumont (celibouteiller@gmail.com), Alain
Pascal (linux.apa@gmail.com)

**LC** — Danny Segui (d.segui@psm-ffessm.fr)

**DNF** — Benjamin Frasca (benjamin.frasca@gmail.com), Alain Pascal
(linux.apa@gmail.com), Yvan Measson (ymeasson@isybot.com), Katarina
Stankiewicz (kasia.milena@gmail.com)

**STA** — Victorine Chardonnet (chardonnetvictorine@gmail.com), François
Memheld (francois.memheld@orange.fr), Guillaume Boulant
(gboulant@gmail.com)

Cette liste vient de l'onglet Inscriptions du paramétrage (rôle
"encadrant", actifs à ce jour) — à régénérer si la composition change.

**Séparément** : chaque encadrant listé ci-dessus doit aussi être ajouté
comme **utilisateur test** sur l'écran de consentement OAuth (une seule
liste, club entier, cf. étape 2 ci-dessous) — sinon Google refuse la
connexion côté web app, même si le classeur Drive est bien partagé.

## Étapes de déploiement

### 1. Créer un Google Sheet par classeur (à répéter pour chaque cible)

Convertir le `.xlsx` correspondant en Google Sheet (Drive : Nouveau ▸
Importer un fichier, dans le bon dossier — voir plus bas pour PROD).
Noter son identifiant (dans l'URL, entre `/d/` et `/edit`).

### 2. Créer le projet standalone, puis un déploiement par palier

Sur [script.google.com](https://script.google.com) (pas depuis un Sheet —
c'est justement le principe du standalone) :

1. **Nouveau projet** — une seule fois, jamais par classeur ni par palier.
   Coller le contenu de `Code.gs` (ce dossier).
2. Renseigner l'objet `CLASSEURS` en haut du fichier : remplacer chaque
   `REMPLACER_PAR_ID_CLASSEUR_*` par l'identifiant Google Sheet noté à
   l'étape 1, pour chaque cible déjà créée (les autres peuvent rester en
   placeholder — elles échoueront proprement avec un message clair tant
   qu'elles ne sont pas renseignées).
3. Enregistrer (Cmd+S).
4. **Déployer ▸ Nouveau déploiement** — répéter cette étape **trois fois**,
   une par palier (aujourd'hui, seul le premier est utile) :
   - Type : **Application web**
   - Description : **DEV** (puis, plus tard, **TEST** et **PROD**) — pour
     s'y retrouver dans « Gérer les déploiements »
   - Exécuter en tant que : **Moi**
   - Qui a accès : **Tout le monde**
5. Autoriser l'accès quand demandé — l'écran « Google n'a pas vérifié
   cette application » est normal (c'est ton propre script). Déployer.
6. Copier l'URL `.../exec` du déploiement **DEV** — à coller dans
   `WEB_APP_URL` de `dev/L2/index.html` **et** `dev/L3/index.html` (ces
   deux fichiers partagent la même URL, celle du palier DEV). Les
   déploiements TEST et PROD attendront d'avoir une version validée à
   servir (cf. « Promotion d'une version » plus haut).

Après toute modification de `Code.gs` (y compris juste ajouter une cible à
`CLASSEURS`) : éditer et enregistrer le fichier ne suffit pas, chaque
déploiement continue de servir la version sur laquelle il est figé tant
qu'on ne le redéploie pas explicitement — voir « Promotion d'une version »
ci-dessus pour savoir quel déploiement mettre à jour (DEV seul, pour
itérer vite, ou DEV puis TEST puis PROD, pour propager un changement déjà
validé).

### 3. Créer l'identifiant OAuth (une seule fois, partagé)

Dans [console.cloud.google.com](https://console.cloud.google.com), sur le
projet associé au script :

1. **APIs et services ▸ Écran de consentement OAuth** :
   - Type : Externe, **Statut : Test** (jusqu'à 100 utilisateurs sans
     validation Google). Ajouter les emails de **tous** les encadrants
     (toutes lignes confondues, cf. liste ci-dessus) dans « Utilisateurs
     test ».
2. **APIs et services ▸ Identifiants ▸ Créer des identifiants ▸ ID client
   OAuth ▸ Application Web** :
   - Origines JavaScript autorisées : l'origine GitHub Pages (bare, sans
     chemin — ex. `https://apneeattitude.github.io`, pas
     `.../mvp-secours-presences/L3`).
3. Copier le **Client ID** — déjà renseigné dans tous les `index.html`
   générés (le même partout).

### 4. Héberger

GitHub Pages, un chemin par environnement/ligne dans le même dépôt :
`https://<utilisateur>.github.io/<repo>/dev/L2/`, `.../dev/L3/`,
`.../L1/`, … `.../STA/`.

### 5. Tester

Ouvrir l'URL sur un smartphone, se connecter avec un compte **ajouté comme
utilisateur test**, vérifier : la liste des séances se charge (avec la
date, ex. « S001 — 07/09/2026 (lundi) »), la recherche incrémentale trouve
un apnéiste par nom/prénom, l'enregistrement écrit bien dans l'onglet
`Presences` du bon classeur — **et pas celui d'une autre cible** (premier
risque introduit par le script partagé : vérifier que `dev/L2/` écrit bien
dans le classeur L2, jamais L3).

## Flux de promotion

Les générateurs Python (`Cadrage_v2/generateurs/`) restent la seule source
de vérité pour la structure et les données des classeurs. À chaque
évolution :

1. Régénérer localement (`build_suivi.py` + `sync_referentiel.py`).
2. Réimporter le fichier dans le Sheet **DEV**, valider.
3. Réimporter le même fichier dans le Sheet **TEST**, faire valider par un
   encadrant volontaire.
4. Réimporter dans le(s) Sheet(s) **PROD** concerné(s) (Drive : Nouveau →
   Importer un fichier, puis dans le Sheet déjà ouvert : Fichier → Importer
   → Remplacer la feuille de calcul — conserve l'identifiant du classeur,
   donc l'entrée correspondante dans `CLASSEURS` reste valide sans rien
   retoucher côté script).

## Synchronisation des rôles et remplacements

Le classeur `AA - Parametrage 2026-2027` est la source de référence des rôles
par ligne : l'onglet `Inscriptions` porte le couple personne, ligne et rôle
(`élève`, `encadrant` ou `prépa encadrant`). Le rôle `prépa encadrant` est
traité exactement comme `encadrant` par l'application, notamment pour les
remplacements. La migration ponctuelle `preparerRolesPrepaEncadrant` ajoute ce
choix au référentiel et l'a attribué le 19/09/2026 à Francis Wang, Corinne Le
Brigand et Magali Cavatore, sans retirer leur rôle `élève`. La fonction Apps Script `synchroniserEffectifs`
reporte les inscriptions actives dans l'onglet `Personnes` de chaque classeur
de ligne, avec une ligne par couple personne/rôle et les groupes associés à ce
rôle. Les personnes sans inscription active restent présentes dans le
référentiel, sans rôle.

La liste des remplaçants est construite depuis cet onglet `Personnes`. Elle
inclut tous les encadrants actifs du club, y compris ceux dont la ligne
habituelle est différente de la ligne de la séance. Une personne cumulant les
rôles élève et encadrant reste proposée comme encadrant. Le serveur vérifie
également ce rôle avant d'enregistrer le remplacement ; seul l'ID est écrit
dans `Calendrier`.

Depuis le 19/09/2026, `LIGNES_SYNC` inclut aussi DNF1, DNF2, STA1 et STA2.
Les entrées `synchroniserEffectifsTestDnf1`, `synchroniserEffectifsTestDnf2`,
`synchroniserEffectifsTestSta1` et `synchroniserEffectifsTestSta2` permettent
de vérifier chaque classeur de test isolément. Après validation, la fonction
`synchroniserEffectifsLignesSeancesProd` synchronise uniquement ces quatre
lignes en production, sans modifier L1 à LC. La version 50 du déploiement
Apps Script TEST puis PROD a été validée ainsi : chaque classeur expose les
25 encadrants actifs du club dans la liste des remplaçants. Le 19/09/2026,
la version 51 a été validée en TEST puis promue en PROD pour le rôle `prépa
encadrant` ; la synchronisation complète a contrôlé les neuf classeurs. STAC
rejoindra cette synchronisation lors de la création de son classeur et de ses cibles.

## Statistiques de fréquentation

L'action Apps Script `stats` lit l'historique complet du `Calendrier` et de
`Presences` d'une cible. Elle retient exclusivement les séances antérieures au
jour courant, au statut `tenue`, et ayant au moins une présence. Elle renvoie
la moyenne de participants depuis le début de saison, les moyennes mensuelles
et celles par jour de semaine. Elle expose aussi les séances passées sans
présence qui ne sont ni `tenue`, ni `annulée`, ni `fermée`, avec l'encadrant
résolu dans le calendrier qui doit les renseigner. Le cache est invalidé à
chaque enregistrement de séance. Première livraison isolée sur `_test/L2/`
avec le déploiement TEST version 55 du 19/09/2026 ; elle doit être validée
avant généralisation.

## Statuts de séance

Le classeur central `AA - Parametrage 2026-2027` contient les valeurs de
référence des statuts dans une table normalisée : `ID`, `Libellé`, `Actif`.
Les valeurs de la saison 2026-2027 sont `1 / planifiée`, `2 / tenue`,
`3 / annulée` et `4 / fermée`.

Le pilote TEST-L2 (13/09/2026) reproduit cette table dans son onglet
`Listes`, sous la plage nommée `ListeStatutsSeance`. La colonne I
`Calendrier!Statut` reste le libellé lisible et sa validation est une liste
« depuis une plage » pointant vers `Listes`. Une colonne `Statut (ID)` est
ajoutée en fin du tableau Calendrier : sa formule déduit l'ID depuis le
libellé. Elle est volontairement ajoutée à la fin, jamais près de la colonne
I, afin de ne pas déplacer les colonnes utilisées par le remplaçant, le plan
de séance et les formules existantes.

`migrerStatutsSeanceTestL2` est une opération manuelle, limitée à TEST-L2.
Elle vérifie les statuts déjà présents avant toute écriture, crée la table et
la validation, puis contrôle que chaque libellé historique reçoit un ID.

Depuis la version TEST 42, `savePresences_` applique aussi une règle métier
côté serveur : si l'enregistrement final contient au moins une présence et
que la date de la séance est strictement antérieure à la date courante, le
statut est forcé à `tenue`. La comparaison utilise la date civile dans le
fuseau Apps Script ; une séance du jour n'est donc pas considérée comme
passée. La règle est ignorée lorsqu'un classeur ne possède pas la plage
`ListeStatutsSeance`, afin de ne pas affecter un environnement non migré.

Après validation fonctionnelle, créer une migration équivalente pour les
classeurs PROD concernés, à lancer d'abord sur une copie de sauvegarde. Ne
promouvoir le déploiement Apps Script TEST vers PROD qu'après cette migration
et une validation explicite.

## Limites assumées de ce MVP (à ne pas découvrir en prod)

- **Pas de contrôle de couverture** (D-02 : directeur de bassin, responsable
  de ligne) — hors périmètre « présences uniquement ».
- **Colonne Contrôle de `Presences` non maintenue** par l'API (formule du
  classeur d'origine laissée telle quelle, purement informative).
- **Pas de vérification que l'encadrant connecté appartient bien à la ligne
  qu'il modifie** : l'accès Drive et la liste de testeurs OAuth filtrent
  qui peut se connecter globalement, mais rien n'empêche aujourd'hui un
  encadrant de L2 de modifier la `CIBLE` envoyée par son navigateur (les
  outils de développement du navigateur suffisent) pour écrire dans L3.
  Le script standalone rend ce risque un peu plus direct qu'avant (une
  seule URL pour tout le monde, `CIBLE` n'est qu'une valeur envoyée par le
  client, jamais vérifiée côté serveur contre l'identité de l'appelant) —
  acceptable pour un stopgap à diffusion restreinte, mais premier point à
  traiter avant toute diffusion plus large (vérifier côté script que
  l'email de l'encadrant a bien un rôle actif dans la ligne demandée,
  via l'onglet Personnes du classeur ciblé).
## Vue globale des statistiques

La page TEST `_test/statistiques/` appelle l’action Apps Script `statsClub`.
Cette action est une requête POST authentifiée par le jeton Google de la
personne connectée. Elle limite l’accès aux adresses de
`ACCES_STATS_GLOBALES_TEST` et ne retourne que des agrégats par ligne :
fréquentation moyenne, séances tenues, moyenne par jour de la semaine et
séances à compléter. Le tableau déplie chaque ligne pour afficher ces détails et propose « Tout déplier / Tout replier ». STAC est affichée comme à configurer tant que sa cible n’existe pas. Les noms des participants ne sont jamais retournés. En
TEST, la vue lit les neuf classeurs PROD uniquement en lecture, car tous les
classeurs n'ont pas de copie TEST. Le déploiement TEST correspondant est la version 58 (19/09/2026). Avant toute publication en PROD,
créer une liste d’autorisation PROD contenant Frédéric et les responsables du
club confirmés, sans confondre ce droit avec le rôle d’encadrant.
