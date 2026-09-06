# Git & GitHub — ce qui a été mis en place

Ce document résume la configuration Git/GitHub du projet AppSuiviAA : pourquoi ces choix, ce qui a été fait, et comment refaire les opérations courantes.

## 1. Le dépôt GitHub

- Dépôt : `https://github.com/ApneeAttitude/AppSuiviAA`
- Anciennement nommé `AppAA`, renommé en `AppSuiviAA` (GitHub redirige automatiquement l'ancienne URL vers la nouvelle, et l'historique des commits est conservé — un renommage de dépôt sur GitHub ne perd rien).
- Branche unique : `main`.
- Avant ce travail, le dépôt ne contenait qu'un seul fichier à la racine : un ancien prototype `index.html` (page L3, avant le refactoring de `Code.gs`). Ce fichier a été conservé (archivé, pas supprimé) en le renommant `indexold.html`.

## 2. Structure du dépôt (racine du dépôt = racine du site)

```
AppSuiviAA/
├── index.html          (n'existe plus à la racine — voir §4)
├── indexold.html        ancien prototype L3, conservé en archive
├── Code.gs               script Apps Script backend (référence/historique)
├── git.md                ce fichier
├── dev/
│   ├── README.md         spécification complète du projet
│   ├── L2/index.html     front DEV, ligne L2 (WEB_APP_URL déjà renseignée)
│   └── L3/index.html     front DEV, ligne L3 (WEB_APP_URL déjà renseignée)
├── test/
│   ├── L2/index.html     front TEST, ligne L2 (placeholder WEB_APP_URL)
│   └── L3/index.html     front TEST, ligne L3 (placeholder WEB_APP_URL)
├── L1/, L2/, L3/, L4/, LC/, DNF/, STA/
│   └── index.html        front PROD, une page par ligne (placeholder WEB_APP_URL)
└── .gitignore            exclut les .DS_Store
```

Chaque page `index.html` porte un `CIBLE` différent (ex. `DEV-L2`, `TEST-L3`, `PROD-L1`) qui correspond à une entrée de l'objet `CLASSEURS` dans `Code.gs` (le mapping cible → ID de classeur Google Sheets).

## 3. GitHub Pages

- Activé, source = branche `main`, dossier `/ (root)`.
- Site en ligne : `https://apneeattitude.github.io/AppSuiviAA/`
- Comme il n'y a plus d'`index.html` à la racine, cette URL nue renvoie un 404 — c'est normal, il faut viser directement le sous-dossier de la ligne/environnement voulu, par exemple :
  - DEV : `https://apneeattitude.github.io/AppSuiviAA/dev/L2/`
  - TEST : `https://apneeattitude.github.io/AppSuiviAA/test/L2/`
  - PROD : `https://apneeattitude.github.io/AppSuiviAA/L2/`
- Origine JavaScript autorisée côté Google (OAuth) : `https://apneeattitude.github.io` (déjà configurée, correspond).

## 4. Les commandes exécutées (pour référence / cas où il faut refaire une mise à jour)

Depuis le dossier local du projet (Terminal, sur ta machine — jamais depuis l'environnement Claude) :

```bash
cd "/Users/flebrigand/Library/CloudStorage/OneDrive-Personal/Dev/Apnee/SiteApnee/AppSuiviAA"

# Première mise en place uniquement (déjà fait) :
git init
git remote add origin https://github.com/ApneeAttitude/AppSuiviAA.git
git fetch origin
git checkout -b main origin/main     # récupère l'historique existant
mv index.html indexold.html          # archive l'ancien prototype

# Pour toute mise à jour ultérieure :
git add -A
git status                           # vérifier ce qui va être commité avant de valider
git commit -m "message décrivant le changement"
git push
```

## 5. Authentification GitHub

Depuis 2021, GitHub n'accepte plus le mot de passe du compte pour les opérations `git` (push/pull en HTTPS) — seul un **token d'accès personnel (PAT)** fonctionne à la place du mot de passe.

- Le username demandé est le vrai identifiant GitHub (visible en haut à droite sur github.com), pas le nom complet configuré dans `git config user.name`.
- Le token se génère sur `https://github.com/settings/tokens` (ou `.../personal-access-tokens/new` pour un token "fine-grained"), avec :
  - Resource owner : `ApneeAttitude` si c'est une organisation à laquelle ton compte a accès.
  - Repository access : "Only select repositories" → `AppSuiviAA`.
  - Permissions : `Contents: Read and write` au minimum.
- Une fois collé une première fois comme mot de passe, macOS le mémorise en général dans le trousseau (Keychain) — pas besoin de le retaper à chaque push.
- Le token expire après la durée choisie à sa création : en cas d'échec d'authentification plus tard, il faudra en régénérer un.

## 6. Ce qui reste à faire

- Configurer les classeurs `TEST-L2`, `TEST-L3` et tous les `PROD-*` dans `CLASSEURS` (`Code.gs`), aujourd'hui encore des placeholders `REMPLACER_PAR_ID_CLASSEUR_*`.
- Créer les déploiements Apps Script TEST et PROD ("Gérer les déploiements"), puis reporter leurs URLs `.../exec` dans les `index.html` correspondants (`WEB_APP_URL`, encore des placeholders `REMPLACER_PAR_URL_EXEC_*`).
- Éventuellement ajouter une page d'accueil à la racine du site (liste de liens vers chaque ligne/environnement), pour que l'URL nue `https://apneeattitude.github.io/AppSuiviAA/` affiche quelque chose plutôt qu'un 404.
- À terme, migrer l'hébergement vers `https://apneeattitude.com` une fois les accès admin de ce domaine obtenus.
