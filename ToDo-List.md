# ToDo-List — Projet Apnée

Dernière mise à jour : 07/09/2026

## Accès et partage — dossier Drive "App Suivi"

- [ ] Réinviter comme Éditeurs (adresses désormais confirmées/corrigées le 06/09/2026) : Gabrielle Beyl (Gabrielle.beyl@gmail.com), Pierre-Alain Hoyau (pa.hoyau@gmail.com), Sophie Landeau (sophtou91@gmail.com)
- [ ] Yvan Measson (ymeasson@isybot.com) : obtenir une adresse reliée à un compte Google (la sienne ne l'est pas), puis l'inviter comme Éditeur
- [ ] Revoir les droits d'accès des répertoires Drive des lignes L1, L2, L3, L4, LC, DNF, STA (qui a accès à chaque dossier, en lecture ou en édition)

## Application de suivi des présences (AppSuiviAA)

- [x] Conversion d'`AA - Parametrage 2026-2027.xlsx` en Google Sheet natif, à la racine de Mon Drive (fait par Fred le 06/09/2026) : https://docs.google.com/spreadsheets/d/18vMX5fqFgCN7NrSkPsSr1ftVoonPFC5xEbjf73iM8Lw
- [x] (partiel) Portage en Apps Script (`Code.gs`, fonction `synchroniserEffectifs`) de la logique déjà écrite et testée hors ligne dans `generateurs/sync_referentiel.py` : synchronise l'onglet Personnes (Nom/Prénom/Rôle déclaré/Groupe(s)/Objectif de la saison) de chaque classeur PROD (L1, L2, L3, L4, LC) depuis le Google Sheet de paramétrage — sans ressaisie. Ne touche jamais aux colonnes A (Code) et H (Membre), qui restent des formules propres à chaque classeur.
- [x] Renommage du champ « Qualité » en « Zone de confort » (valeurs : Dans la zone / En limite / Hors zone), suite à la remarque d'un encadrant sur le caractère normatif de « conforme » — mis à jour dans les 7 front-ends de ligne (L1, L2, L3, L4, LC, DNF, STA + copies _test/_dev L2/L3) et dans `generateurs/aa_common.py` (`QUALITE` → `ZONE_CONFORT`), `build_suivi.py`, `build_parametrage.py` (07/09/2026)
- [ ] Mettre à jour manuellement les classeurs PROD déjà déployés (L1-STA) suite à ce renommage : en-tête de la colonne F de l'onglet Presences (« Qualité de réalisation » → « Zone de confort ») et liste de validation associée (nouvelles valeurs Dans la zone / En limite / Hors zone) — la génération Python ne touche pas les classeurs déjà créés
- [ ] Tester `synchroniserEffectifs` en manuel depuis l'éditeur Apps Script (menu Exécuter) et relire les onglets Personnes des classeurs PROD avant de faire confiance au résultat (Fred, après sa pause du 06/09/2026)
- [ ] Analyser les écarts entre les membres par ligne consignés dans le fichier de paramétrage (onglet Inscriptions) et la composition des groupes décrite dans le document de référence du club : https://sides-carry-3kj.craft.me/j2Gng2vJkRJtVa (« École d'apnée » et « Groupe compétition », composition par niveau L1-L4, DNF, STA, groupe compétition)
- [ ] Tester quelques cas unitaires de la synchronisation (cas simples ciblés) avant de passer à un cas réel complet
- [ ] Une fois le développement du fichier de paramétrage centralisé terminé : tester la synchronisation sur l'exemple réel du document Craft.me ci-dessus (composition des groupes), après les cas unitaires
- [ ] Une fois validé : ajouter un menu "Synchroniser" sur le classeur de Paramétrage + un déclencheur `onOpen` (et un filet de sécurité en tâche planifiée nocturne), pour ne plus avoir à lancer la synchronisation depuis l'éditeur Apps Script
- [ ] Porter aussi la synchronisation du Référentiel (Objectifs, Compétences, Créneaux, Responsables, Qualifications Sécurité) et l'initialisation du Calendrier — `sync_referentiel.py` le fait déjà, mais la mise en page (positions figées) des classeurs PROD actuellement déployés n'a pas encore été vérifiée ligne à ligne comme elle l'a été pour Personnes
- [ ] Une seule URL pour toutes les lignes (au lieu d'une URL par ligne)
- [ ] Authentification : gérer des tokens pour ne pas la redemander à chaque fois
- [ ] Améliorer les temps de réponse au chargement
- [ ] Automatiser la chaîne CI/CD (déploiement du script Apps Script et publication GitHub Pages)
- [ ] Adapter les scripts du dossier `generateurs/` (`build_parametrage.py`, `build_suivi.py`, `sync_referentiel.py`, `check_classeurs.py`) : ils supposent une arborescence `generateurs/` + `classeurs/` (chemins relatifs du type `../classeurs/...`), alors qu'AppSuiviAA range chaque ligne dans son propre dossier (`L1/`, `L2/`, `L3/`, `L4/`, `LC/`, `DNF/`, `STA/`)
- [ ] Committer/pousser les fichiers rapatriés depuis Cadrage_v2 (`AA - Parametrage 2026-2027.xlsx` et `generateurs/`) si ce n'est pas déjà fait
- [ ] Vérifier le bon fonctionnement des lignes PROD L1, L3, L4 et LC sur le site GitHub Pages, suite au renommage du déploiement Apps Script et à l'ajout de `cache: 'no-store'` sur les appels front
- [ ] Supprimer les répertoires OneDrive inutiles : `mvp-secours-presences`, `App AA DEV`, `App AA MVP` (déjà supprimés deux fois par Fred mais réapparus — vérifier sur onedrive.com, dans la corbeille, et sur les autres appareils synchronisés avant de resupprimer)

## Fonctionnalités à venir

Idées de fonctionnalités futures, classées par priorité (P1 = prioritaire, P3 = à explorer sans urgence) :

- [x] **P1** — Saisie de la description de la séance (texte et/ou lien) depuis l'appli, ergonomie retenue : accordéon repliable « 📋 Plan de séance » sous le sélecteur de séance (option B), replié par défaut, prérempli si une valeur existe déjà, coché « ✓ » quand rempli. Enregistré avec le même bouton « Enregistrer la séance » que les présences (pas de bouton séparé). Implémenté le 07/09/2026 : `Code.gs` (lecture/écriture colonne R de Calendrier, invalidation du cache « data ») + les 7 `index.html` de ligne + copies _test/_dev L2/L3 — non testé en conditions réelles, à valider par Fred.
- [ ] **P1** — Saisie du remplaçant en cas d'absence de l'encadrant habituel, depuis l'appli. La colonne « Responsable remplaçant » existe déjà dans Calendrier et l'appli affiche déjà le nom résolu (encadrant ou remplaçant), mais la saisie du remplaçant se fait uniquement en éditant directement le Google Sheet — pas possible depuis l'appli elle-même. Reste à faire.
- [ ] **P3** — Interface dédiée aux apnéistes (et non plus seulement aux encadrants) pour qu'ils renseignent eux-mêmes leur ressenti de séance et leur forme du jour. Développement distinct de l'existant car public différent (apnéistes plutôt qu'encadrants), mais qui permettrait aussi, potentiellement, de corréler les avis encadrants et les avis apnéistes.

## Refonte du module Suivi — cahier des charges v2

Arbitrages nécessaires avant rédaction (issus de l'analyse du classeur "Suivi Apnéistes L3") :

- [ ] Reprise de l'historique 2025-2026 (525 cellules à éclater, en partie manuellement) ou démarrage à neuf en septembre 2026 avec seulement synthèse + référentiel ?
- [ ] Avis d'encadrants (couleurs `T1`/`T1+T2`) : repris comme données historiques ou reconstitués dès la première campagne dans le nouvel outil ?
- [ ] Échelle souhaitée pour la qualité de réalisation, en remplacement de `ok/x/-` (binaire, trois niveaux, ou l'échelle à 4 niveaux déjà utilisée pour les techniques) ?
- [ ] Objectif "en maîtrise" : les 3 répétitions doivent-elles être faites dans une même séance, ou cumulées sur la saison ?
- [ ] Périmètre de visibilité des commentaires : un encadrant voit-il les commentaires des autres encadrants sur un apnéiste ? Et l'apnéiste lui-même ?
- [ ] Registre dédié aux événements de sécurité, avec alerte en cas de cumul d'incidents pour un même apnéiste ?
- [ ] Les invités deviennent-ils des personnes en base, ou restent une simple mention textuelle sur la séance ?
- [ ] Référentiels d'objectifs/compétences par ligne (L1 à L4, groupe compétition) : existent-ils déjà rédigés au-delà de L3 ?
- [ ] Conserve-t-on la logique de deux campagnes statistiques annuelles (décembre, avril), à quelles dates de référence ?
- [ ] Seuils de fréquentation (vert/jaune/orange/rouge) par ligne d'eau : quelles valeurs retenir ?
