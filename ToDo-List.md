# ToDo-List — Projet Apnée

Dernière mise à jour : 06/09/2026

## Accès et partage — dossier Drive "App Suivi"

- [ ] Réinviter comme Éditeurs (adresses désormais confirmées/corrigées le 06/09/2026) : Gabrielle Beyl (Gabrielle.beyl@gmail.com), Pierre-Alain Hoyau (pa.hoyau@gmail.com), Sophie Landeau (sophtou91@gmail.com)
- [ ] Yvan Measson (ymeasson@isybot.com) : obtenir une adresse reliée à un compte Google (la sienne ne l'est pas), puis l'inviter comme Éditeur

## Application de suivi des présences (AppSuiviAA)

- [ ] Adapter les scripts du dossier `generateurs/` (`build_parametrage.py`, `build_suivi.py`, `sync_referentiel.py`, `check_classeurs.py`) : ils supposent une arborescence `generateurs/` + `classeurs/` (chemins relatifs du type `../classeurs/...`), alors qu'AppSuiviAA range chaque ligne dans son propre dossier (`L1/`, `L2/`, `L3/`, `L4/`, `LC/`, `DNF/`, `STA/`)
- [ ] Committer/pousser les fichiers rapatriés depuis Cadrage_v2 (`AA - Parametrage 2026-2027.xlsx` et `generateurs/`) si ce n'est pas déjà fait
- [ ] Vérifier le bon fonctionnement des lignes PROD L1, L3, L4 et LC sur le site GitHub Pages, suite au renommage du déploiement Apps Script et à l'ajout de `cache: 'no-store'` sur les appels front
- [ ] Confirmer la disparition définitive des dossiers OneDrive `mvp-secours-presences`, `App AA DEV`, `App AA MVP` (réapparus à deux reprises après suppression locale) : vérifier sur onedrive.com, dans la corbeille, et sur les autres appareils synchronisés

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
