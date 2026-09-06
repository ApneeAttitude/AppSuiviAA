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
