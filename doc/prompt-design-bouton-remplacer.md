# Prompt à envoyer à une IA — refonte du bouton « Remplacer »

Tu es un designer UI/UX spécialisé dans les interfaces web mobile-first, avec une expertise en design système et micro-interactions. J'ai besoin de tes propositions pour améliorer un bouton précis dans une application web.

## Contexte de l'application

- Application de suivi de séances d'apnée, utilisée sur le terrain (souvent sur mobile, parfois en plein soleil ou avec les mains humides).
- Palette : fond photo sombre (plongeurs) en en-tête, cartes « verre dépoli » (glassmorphism : fond blanc translucide + flou d'arrière-plan), coins arrondis 14 à 18px.
- Couleur d'accent principale : vert-bleu teal `#0f6e6a` (texte foncé `#0f5754`), couleur de focus/accessibilité : ambre `#e3a331`.
- Typographies : titres en **Space Grotesk** (gras), texte courant en **Manrope**.
- Largeur de contenu max 480px, cible mobile en priorité, desktop en secondaire.

## L'élément à retravailler

Un bouton « ⇄ Remplacer » placé à droite du texte « Encadrant : Prénom Nom », qui ouvre une boîte de dialogue (élément natif `<dialog>`) permettant de choisir un encadrant remplaçant (recherche + liste de choix). Actuellement, ce bouton est un simple rectangle à bord teal, fond blanc, texte teal, 13px — peu intégré visuellement au reste de la carte. Je le trouve moche et je veux le retravailler.

Contraintes à respecter obligatoirement :

- Le bouton doit rester à côté du texte « Encadrant : Prénom Nom » (ne pas casser la mise en page en ligne).
- Il doit changer de libellé une fois un remplaçant choisi (passer de « Remplacer » à « Modifier »).
- Zone de tap ≥ 44px de hauteur (usage tactile).
- La boîte de dialogue doit rester : centrée sur desktop, en feuille remontant du bas (« bottom sheet ») sur mobile — ce comportement existe déjà et fonctionne bien ; il ne s'agit pas de le changer, seulement le déclencheur (bouton) et éventuellement son harmonisation visuelle avec la fenêtre.

## Ce que j'attends de toi

Propose-moi **3 pistes de design différentes** pour ce bouton (pas une seule variante déclinée trois fois). Pour chacune, précise comment elle répond à ces trois critères :

1. **Simplicité d'usage** — en quoi elle rend l'action plus évidente et plus rapide à comprendre et à déclencher, y compris pour un utilisateur peu à l'aise avec le numérique.
2. **Harmonie visuelle** — en quoi elle s'intègre mieux à l'identité visuelle existante (teal/ambre, glassmorphism, coins arrondis) plutôt que de trancher comme un élément rapporté.
3. **Responsive design** — comment elle se comporte sur petit écran (mobile, priorité) par rapport au desktop, y compris si sa position, sa taille ou sa forme doivent s'adapter.

Pour chaque piste, fournis :

- Un nom court (ex. « Icône seule + tooltip », « Chip pilule avec icône »…).
- Une description en 2-3 phrases.
- Le code HTML + CSS correspondant, prêt à intégrer, cohérent avec les couleurs/polices ci-dessus (CSS pur, sans framework externe).
- Un avantage et une limite.

Termine par une recommandation : parmi les 3 pistes, laquelle tu choisirais en priorité, et pourquoi.
