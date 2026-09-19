// Code.gs — Suivi des présences, MVP de secours (jalon 07/09/2026)
//
// Script STANDALONE, déployé une seule fois pour toutes les lignes et tous
// les environnements — décision du 07/09/2026 : un seul développement à
// maintenir plutôt qu'une copie de ce fichier par classeur.
//
// Coût accepté en connaissance de cause : SpreadsheetApp.openById() est
// mesuré 1 à 2,5 s par appel (contre quasi gratuit pour un script LIÉ avec
// getActiveSpreadsheet(), l'architecture d'origine de ce MVP) — c'est ce
// qui rendait lent le prototype AppSuivi initial. Le gain de maintenance
// (une seule URL, un seul code) l'emporte ici sur la vitesse perçue.
// Pour limiter l'impact sur la vitesse perçue, les réponses de "data" et
// "fiche" sont mises en cache (CacheService) — voir plus bas (06/09/2026).
//
// « cible » identifie quel classeur ouvrir : un couple environnement/ligne
// (ex. "DEV-L2", "PROD-L3"), envoyé par le front à chaque appel. Le
// paramétrage — la correspondance cible → identifiant de classeur Google
// Sheet — est CLASSEURS ci-dessous : c'est le seul et unique endroit à
// mettre à jour quand un nouveau classeur est créé ou remplacé.
//
// Le front (hébergé ailleurs, ex. GitHub Pages) appelle ce script en
// fetch() classique sur son URL /exec — pas de souci de sandbox tant que
// le front n'est pas lui-même servi par ce script via HtmlService.

// Identifiants des GoogleSheet de chaque ligne
var CLASSEURS = {
  'DEV-L2':   'REMPLACER_PAR_ID_CLASSEUR_DEV_L2',
  'DEV-L3':   'REMPLACER_PAR_ID_CLASSEUR_DEV_L3',
  'TEST-L2':  '1jy8hW_2haO3kolxPwgNnT5qANA1kI91_2Kutvn2FQs8',
  'TEST-L3':  'REMPLACER_PAR_ID_CLASSEUR_TEST_L3',
  'TEST-DNF1': '1T9l-Vmy6v0-GtRuOjrq8aZAVXsC6474FL0RxhnaZgdI',
  'TEST-DNF2': '1E3OoYTEWQ07Ykmbydwkf08Hns1e1YVR1ofOThezeZ94',
  'TEST-STA1': '1QGzitEk7jM7undte_y_EmaQYhsK_faKcFHTyni9folo',
  'TEST-STA2': '1PwMAxWjeOEE8Qpe0DRNLrVwv-KqmHZMlsTqFKloqBPg',
  'PROD-DNF1': '1kc_G0xQx2SjPIHWUdGB_tt_HKN4cZQUWsFMAPttX21g',
  'PROD-DNF2': '1X8E-TMjNh2nvH71MLeJ2DJrzwUvun0eNsD0lTMtqa_c',
  'PROD-STA1': '1J5rI7VnXsQ2ixXgjEvzamtFdys2NrwjcUAgWxzInA3c',
  'PROD-STA2': '1TE7hiZEJPuvTOF6BVct-1Jp0opxpbqIL3kwp9aSqGXo',
  'PROD-L1':  '1K2h_E7NaJwGuGiAZn_qqhoMHZGxbW39v-j1_j8x9Pf8',
  'PROD-L2':  '1ANbbV-lc9GZeVHQH8X4mFOaMKo4s8wB5TrnmX96WEko',
  'PROD-L3':  '1gVjJxXIXzvfJUObElSu8D3dnFRHTNcCWemQH-JwDsRY',
  'PROD-L4':  '1_MSHANd4z8XDeh8d8CD7402AP_caSkBIaPbiwHIZLjQ',
  'PROD-LC':  '1XCrHI4MRfhWnGljkul7XIlHteI5hjyCuvUYLi-HzOqg',
  'PROD-DNF': 'REMPLACER_PAR_ID_CLASSEUR_PROD_DNF',
  'PROD-STA': 'REMPLACER_PAR_ID_CLASSEUR_PROD_STA'
};

var SHEET_CALENDRIER = 'Calendrier';
var SHEET_PRESENCES = 'Presences';
var SHEET_PERSONNES = 'Personnes';
var SHEET_LISTES = 'Listes';

// Durées de cache (secondes) — CacheService, partagé entre toutes les
// exécutions du script (donc entre tous les appareils). "data" (roster +
// calendrier) ne change que par une modification manuelle du classeur
// (saisie de saison, ajout d'un adhérent) : un léger délai d'affichage
// après une telle modif est un compromis acceptable pour un chargement
// quasi instantané le reste du temps (demande du 06/09/2026 : "le
// chargement des data est lent"). "fiche" (présences d'une séance) est
// recalculée et remise en cache immédiatement après chaque enregistrement
// (cf. savePresences_) : le TTL ne sert donc qu'à couvrir une modif faite
// directement dans le Sheet, jamais un enregistrement fait depuis l'appli.
var CACHE_TTL_DONNEES = 120;
var CACHE_TTL_FICHE = 60;

// Vue de synthèse réservée aux responsables. La liste est volontairement
// séparée des rôles d'encadrant : consulter les statistiques de toutes les
// lignes n'est pas nécessaire pour renseigner une séance.
var ACCES_STATS_GLOBALES_TEST = ['flebrigand@gmail.com'];
// La page reste en TEST, mais compare les données réelles en lecture seule :
// tous les classeurs n'ont pas de copie TEST. Aucun de ces classeurs n'est
// modifié par le calcul.
var CIBLES_STATS_GLOBALES_TEST = ['PROD-L1', 'PROD-L2', 'PROD-L3', 'PROD-L4', 'PROD-LC', 'PROD-DNF1', 'PROD-DNF2', 'PROD-STA1', 'PROD-STA2'];

function ouvrirClasseur_(cible) {
  if (!cible) throw new Error('cible manquante (environnement/ligne)');
  var id = CLASSEURS[cible];
  if (!id) throw new Error('cible inconnue : ' + cible);
  if (id.indexOf('REMPLACER_PAR_ID_') === 0) {
    throw new Error('cible pas encore configurée (CLASSEURS) : ' + cible);
  }
  return SpreadsheetApp.openById(id);
}

function doGet(e) {
  var action = e.parameter.action;
  var cible = e.parameter.cible;
  try {
    if (action === 'data') {
      return jsonOut_(getData_(cible));
    }
    if (action === 'fiche') {
      return jsonOut_(lirePresences_(cible, e.parameter.seance));
    }
    if (action === 'stats') {
      return jsonOut_(getStats_(cible));
    }
    return jsonOut_({ ok: false, error: 'action inconnue : ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    // Content-Type text/plain côté client, volontairement : un
    // Content-Type application/json déclenche un préflight CORS OPTIONS
    // qu'Apps Script Web App ne sait pas traiter. On parse le JSON
    // nous-mêmes depuis le corps texte brut.
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'saveSeance') return jsonOut_(savePresences_(body));
    if (body.action === 'statsClub') return jsonOut_(getStatsClub_(body));
    return jsonOut_({ ok: false, error: 'action inconnue : ' + body.action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- Identification de l'encadrant -----------------------------------------
// Le front envoie le jeton ID obtenu par Google Identity Services (Sign In
// With Google) ; on le vérifie auprès de Google plutôt que de faire
// confiance à un email envoyé tel quel — même principe que le prototype
// AppSuivi (1.7_architecture_web_v0.md Cadrage_v2, §1).
function verifierJeton_(idToken) {
  if (!idToken) throw new Error('jeton manquant — connexion Google requise');
  var resp = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('jeton invalide ou expiré');
  var info = JSON.parse(resp.getContentText());
  if (!info.email || info.email_verified !== 'true') {
    throw new Error('email non vérifié par Google');
  }
  return info.email;
}

// --- Lecture ----------------------------------------------------------------
// Réponse mise en cache (roster + séances) : évite de rouvrir le classeur
// (SpreadsheetApp.openById(), 1 à 2,5 s) à chaque rechargement de la page
// pendant la durée du cache (demande du 06/09/2026, cf. CACHE_TTL_DONNEES).
function getData_(cible) {
  var cache = CacheService.getScriptCache();
  var cle = 'v2|' + cible + '|data';
  var brut = cache.get(cle);
  if (brut) return JSON.parse(brut);

  var ss = ouvrirClasseur_(cible);
  var resultat = {
    ok: true, roster: lirePersonnes_(ss), seances: lireCalendrier_(ss),
    listesZoneConfort: lireListeZoneConfort_(ss),
    listesStatutsSeance: lireListeStatutsSeance_(ss)
  };
  cache.put(cle, JSON.stringify(resultat), CACHE_TTL_DONNEES);
  return resultat;
}

// Statistiques de fréquentation : seules les séances passées, marquées
// « tenue » et possédant au moins une présence sont prises en compte. Ce
// choix évite d'interpréter une séance planifiée ou une saisie incomplète
// comme une séance sans participant.
function getStats_(cible) {
  var cache = CacheService.getScriptCache();
  var cle = 'v2|' + cible + '|stats';
  var brut = cache.get(cle);
  if (brut) return JSON.parse(brut);

  var ss = ouvrirClasseur_(cible);
  var shCalendrier = ss.getSheetByName(SHEET_CALENDRIER);
  var shPresences = ss.getSheetByName(SHEET_PRESENCES);
  if (!shCalendrier || !shPresences) throw new Error('onglet Calendrier ou Presences introuvable');

  var presenceParSeance = {};
  var presences = shPresences.getDataRange().getValues();
  for (var p = 4; p < presences.length; p++) {
    var idPresence = presences[p][0];
    // Une présence est valide quand l'identifiant de l'apnéiste est résolu.
    if (idPresence && presences[p][2]) {
      var clePresence = String(idPresence);
      presenceParSeance[clePresence] = (presenceParSeance[clePresence] || 0) + 1;
    }
  }

  var maintenant = new Date();
  maintenant.setHours(0, 0, 0, 0);
  var calendrier = shCalendrier.getDataRange().getValues();
  var totalParticipants = 0;
  var totalSeances = 0;
  var seancesACompleter = [];
  var mois = {};
  var jours = [
    { cle: 1, libelle: 'Lundi', participants: 0, seances: 0 },
    { cle: 2, libelle: 'Mardi', participants: 0, seances: 0 },
    { cle: 3, libelle: 'Mercredi', participants: 0, seances: 0 },
    { cle: 4, libelle: 'Jeudi', participants: 0, seances: 0 },
    { cle: 5, libelle: 'Vendredi', participants: 0, seances: 0 },
    { cle: 6, libelle: 'Samedi', participants: 0, seances: 0 },
    { cle: 0, libelle: 'Dimanche', participants: 0, seances: 0 }
  ];
  for (var r = 4; r < calendrier.length; r++) {
    var row = calendrier[r];
    var date = row[1];
    var idSeance = row[0];
    var statut = String(row[8] || '').toLowerCase().trim();
    var estDate = Object.prototype.toString.call(date) === '[object Date]';
    var nbParticipants = presenceParSeance[String(idSeance)] || 0;
    if (!idSeance || !estDate) continue;
    var statutClos = statut === 'annulée' || statut === 'annulee' || statut === 'fermée' || statut === 'fermee';
    if (date < maintenant && statut !== 'tenue' && !statutClos && !nbParticipants) {
      seancesACompleter.push({
        date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        encadrant: (String(row[16] || '') + ' ' + String(row[15] || '')).trim() || 'Encadrant non renseigné'
      });
    }
    if (date >= maintenant || statut !== 'tenue' || !nbParticipants) continue;

    totalSeances++;
    totalParticipants += nbParticipants;
    var cleMois = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
    if (!mois[cleMois]) {
      mois[cleMois] = {
        cle: cleMois,
        libelle: ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'][date.getMonth()] + ' ' + date.getFullYear(),
        participants: 0,
        seances: 0
      };
    }
    mois[cleMois].participants += nbParticipants;
    mois[cleMois].seances++;
    var indexJour = jours.map(function (j) { return j.cle; }).indexOf(date.getDay());
    jours[indexJour].participants += nbParticipants;
    jours[indexJour].seances++;
  }

  function avecMoyenne_(x) {
    return {
      libelle: x.libelle,
      seances: x.seances,
      moyenne: x.seances ? Math.round((x.participants / x.seances) * 10) / 10 : 0
    };
  }
  var resultat = {
    ok: true,
    regle: 'Séances passées, tenues et avec au moins une présence',
    seances: totalSeances,
    participants: totalParticipants,
    moyenne: totalSeances ? Math.round((totalParticipants / totalSeances) * 10) / 10 : 0,
    mensuel: Object.keys(mois).sort().map(function (cleMois) { return avecMoyenne_(mois[cleMois]); }),
    hebdomadaire: jours.filter(function (jour) { return jour.seances > 0; }).map(avecMoyenne_),
    seances_a_completer: seancesACompleter.sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); })
  };
  cache.put(cle, JSON.stringify(resultat), CACHE_TTL_DONNEES);
  return resultat;
}

// La vue globale ne transmet que des agrégats par ligne. Les statistiques
// individuelles et les noms des participants restent dans leur classeur.
function getStatsClub_(body) {
  if (body.environnement !== 'TEST') {
    throw new Error('vue globale indisponible pour cet environnement');
  }
  var email = verifierJeton_(body.idToken).toLowerCase();
  if (ACCES_STATS_GLOBALES_TEST.indexOf(email) === -1) {
    throw new Error('accès réservé à Frédéric et aux responsables du club');
  }

  var cache = CacheService.getScriptCache();
  var cle = 'v1|stats-club|TEST';
  var brut = cache.get(cle);
  if (brut) return JSON.parse(brut);

  var totalParticipants = 0;
  var totalSeances = 0;
  var totalACompleter = 0;
  var lignes = CIBLES_STATS_GLOBALES_TEST.map(function (cible) {
    var stats = getStats_(cible);
    totalParticipants += stats.participants || 0;
    totalSeances += stats.seances || 0;
    totalACompleter += (stats.seances_a_completer || []).length;
    return {
      code: cible.replace(/^(TEST|PROD)-/, ''),
      moyenne: stats.moyenne,
      seances: stats.seances,
      a_completer: (stats.seances_a_completer || []).length,
      hebdomadaire: stats.hebdomadaire || [],
      seances_a_completer: stats.seances_a_completer || []
    };
  });
  var resultat = {
    ok: true,
    regle: 'Séances passées, tenues et avec au moins une présence',
    moyenne: totalSeances ? Math.round((totalParticipants / totalSeances) * 10) / 10 : 0,
    seances: totalSeances,
    a_completer: totalACompleter,
    lignes: lignes
  };
  cache.put(cle, JSON.stringify(resultat), CACHE_TTL_DONNEES);
  return resultat;
}

// Colonnes de l'onglet Personnes (1-based, cf. build_suivi.py) :
// A code, B id, C nom, D prenom, E role declare, F groupe(s), G objectif,
// H membre de la ligne ("oui"/"")...
// Renvoie TOUT le monde (pas seulement les membres de la ligne), avec
// l'indicateur "membre" : permet au front de proposer un "invité"
// (personne connue du classeur mais pas de cette ligne) et de l'afficher
// differemment (demande du 07/09/2026).
// Une même personne peut apparaître plusieurs fois dans le classeur lorsqu'elle
// cumule des rôles. L'entrée exposée au front reste unique par ID : `role`
// conserve le rôle élève pour les listes de participants, tandis que `roles`
// et `est_encadrant` conservent tous les rôles pour la liste des remplaçants.
function lirePersonnes_(ss) {
  var sh = ss.getSheetByName(SHEET_PERSONNES);
  var values = sh.getDataRange().getValues();
  var parId = {};
  var ordre = [];
  for (var r = 4; r < values.length; r++) {          // ligne 5 = 1ere donnee
    var row = values[r];
    if (!row[1]) continue;
    var id = row[1];
    var candidat = {
      id: id, nom: row[2], prenom: row[3],
      role: String(row[4] || '').toLowerCase().trim(),
      membre: row[7] === 'oui'
    };
    var existant = parId[id];
    if (!existant) {
      candidat.roles = {};
      if (candidat.role) candidat.roles[candidat.role] = true;
      parId[id] = candidat;
      ordre.push(id);
    } else {
      if (candidat.role) existant.roles[candidat.role] = true;
      existant.membre = existant.membre || candidat.membre;
      if (existant.role !== 'élève' && candidat.role === 'élève') {
        existant.role = candidat.role;
      }
    }
  }
  return ordre.map(function (id) {
    var p = parId[id];
    var roles = Object.keys(p.roles);
    return {
      id: p.id, nom: p.nom, prenom: p.prenom, membre: p.membre, role: p.role,
      roles: roles,
      est_encadrant: roles.indexOf('encadrant') !== -1 || roles.indexOf('prépa encadrant') !== -1
    };
  });
}

// Listes de valeurs sous forme ID + libellé (proposition validée par Fred le
// 09/09/2026, cf. claude/proposition-listes-de-valeurs.md) — Zone de confort
// est le premier cas migré, uniquement sur TEST-L2 pour l'instant. Table de
// référence dans l'onglet "Listes", repérée par la plage nommée
// "ListeZoneConfort" (colonnes ID | Libellé | Actif "oui"/"non") plutôt que
// par des coordonnées fixes (K:M) — depuis la réorganisation du 10/09/2026
// de cet onglet (titre / en-têtes / valeurs, une colonne d'écart entre
// chaque liste, cf. reorganiserListesEtPresences), la position de ce
// tableau varie d'un classeur à l'autre. Un classeur sans cette plage
// nommée n'est simplement pas migré vers ce modèle (ex. tous les PROD
// actuellement) : zone de confort y reste un texte libre, comme avant.
// L'ID est ce qui est stocké dans Presences colonne "Zone de confort
// (ID)", jamais le libellé. Une valeur inactive reste lisible (une
// ancienne présence peut encore y pointer) mais n'est plus proposée dans
// le menu déroulant de saisie (filtrage fait côté front).
function lireListeZoneConfort_(ss) {
  var rng = ss.getRangeByName('ListeZoneConfort');
  if (!rng) return [];
  var values = rng.getValues();
  var out = [];
  for (var r = 0; r < values.length; r++) {
    var id = values[r][0];
    if (id === '' || id === null) continue;
    out.push({ id: id, libelle: values[r][1], actif: values[r][2] === 'oui' });
  }
  return out;
}

// Même convention que pour la zone de confort : l'application manipule un
// identifiant stable, tandis que le classeur et l'interface affichent le
// libellé. L'absence de plage nommée signifie simplement que le classeur
// n'est pas encore migré vers ce modèle.
function lireListeStatutsSeance_(ss) {
  var rng = ss.getRangeByName('ListeStatutsSeance');
  if (!rng) return [];
  var values = rng.getValues();
  var out = [];
  for (var r = 0; r < values.length; r++) {
    var id = values[r][0];
    if (id === '' || id === null) continue;
    out.push({ id: id, libelle: values[r][1], actif: values[r][2] === 'oui' });
  }
  return out;
}

// Fenêtre glissante de séances chargées par l'app : pas la peine de charger
// tout le calendrier, la saisie se fait sur smartphone séance par séance —
// 2 semaines passées + semaine en cours + semaine suivante suffisent
// (décision du 05/09/2026). Recalculée à chaque appel par rapport à
// aujourd'hui ; semaine = lundi à dimanche, au sens du fuseau du script.
function debutSemaine_(date) {
  var d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  var jour = d.getDay(); // 0 = dimanche ... 6 = samedi
  var decalageLundi = (jour === 0) ? -6 : (1 - jour);
  d.setDate(d.getDate() + decalageLundi);
  d.setHours(0, 0, 0, 0);
  return d;
}

// Colonnes de l'onglet Calendrier (cf. build_suivi.py, onglet Calendrier) :
// A id seance, B date, C jour, D creneau (libelle), ... I statut,
// ... N responsable remplacant (debut du nom), O responsable (id),
// P resp. - nom, Q resp. - prenom. Les colonnes O/P/Q resolvent deja,
// par formule dans le classeur, le remplacant s'il est saisi en N, sinon
// l'encadrant habituel du creneau — Code.gs n'a rien a arbitrer ici.
function extraireHoraire_(creneau) {
  var m = /(\d{1,2}h\d{2})\s*-\s*(\d{1,2}h\d{2})/.exec(creneau || '');
  return m ? { debut: m[1], fin: m[2] } : { debut: '', fin: '' };
}

function lireCalendrier_(ss) {
  var debutSemaineEnCours = debutSemaine_(new Date());
  var debutFenetre = new Date(debutSemaineEnCours);
  debutFenetre.setDate(debutFenetre.getDate() - 14);  // lundi, 2 semaines avant
  var finFenetre = new Date(debutSemaineEnCours);
  finFenetre.setDate(finFenetre.getDate() + 14);       // lundi, 2 semaines après (exclusif)

  var tz = Session.getScriptTimeZone();
  var sh = ss.getSheetByName(SHEET_CALENDRIER);
  var responsablesHabituels = {};
  var ref = ss.getSheetByName('Referentiel').getDataRange().getValues();
  for (var iRef = 0; iRef < ref.length; iRef++) {
    if (ref[iRef][0] !== '' && ref[iRef][0] !== null &&
        ref[iRef][7] !== '' && ref[iRef][7] !== null) {
      responsablesHabituels[String(ref[iRef][0])] = ref[iRef][7];
    }
  }
  var values = sh.getDataRange().getValues();
  var header = values[3] || [];
  var colStatutId = header.indexOf('Statut (ID)');
  var out = [];
  for (var r = 4; r < values.length; r++) {
    var row = values[r];
    if (!row[0]) continue;
    var d = row[1];
    var estDate = Object.prototype.toString.call(d) === '[object Date]';
    if (estDate && (d < debutFenetre || d >= finFenetre)) {
      continue; // hors fenêtre : on ne charge pas cette séance
    }
    var horaire = extraireHoraire_(row[3]);
    out.push({
      id: row[0],
      numero: String(row[0] || '').split(' — ')[0],
      date: formatDate_(row[1]),
      date_iso: estDate ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '',
      jour: row[2],
      creneau: row[3],
      heure_debut: horaire.debut,
      heure_fin: horaire.fin,
      // colonnes P/Q : deja resolues cote classeur (remplacant saisi en N,
      // sinon encadrant habituel du creneau) — cf. build_suivi.py write_calendrier_row.
      encadrant_nom: row[15] || '',
      encadrant_prenom: row[16] || '',
      // colonne N (brute) et O (ID resolu) : exposees pour permettre a
      // l'appli de proposer/pre-selectionner un remplacant deja saisi
      // (accordeon "Remplacement", TEST uniquement — demande du 12/09/2026).
      remplacant_actif: !!row[13],
      responsable_id: row[14] || '',
      // Responsable défini dans Referentiel pour le créneau, indépendamment
      // du remplaçant éventuellement saisi en N. Repli sur l'ID résolu en O
      // quand aucun remplacement n'est actif, pour rester compatible avec
      // un ancien classeur dont le référentiel serait incomplet.
      responsable_habituel_id: responsablesHabituels[String(row[4])] ||
        (!row[13] ? (row[14] || '') : ''),
      // colonne R : Plan de séance (texte ou lien), saisi depuis l'appli
      // (accordéon repliable — demande du 07/09/2026).
      plan: row[17] || '',
      // La colonne I conserve le libellé pour la lecture humaine dans le
      // classeur. La colonne ajoutée en fin de tableau contient l'ID dérivé
      // par formule et constitue la donnée envoyée par l'application.
      statut: row[8] || 'planifiée',
      statut_id: colStatutId === -1 ? '' : (row[colStatutId] || '')
    });
  }
  // tri chronologique croissant : la plus ancienne en premier (demande du
  // 07/09/2026) — le front n'a plus besoin d'inverser l'ordre lui-meme.
  out.sort(function (a, b) {
    return a.date_iso < b.date_iso ? -1 : (a.date_iso > b.date_iso ? 1 : 0);
  });
  return out;
}

// Réponse mise en cache (présences d'une séance) : voir CACHE_TTL_FICHE.
// Toujours rafraîchie explicitement par savePresences_ juste après un
// enregistrement, donc jamais périmée pour ce que fait l'appli elle-même.
function lirePresences_(cible, seanceId) {
  if (!seanceId) throw new Error('seance manquante');
  var cache = CacheService.getScriptCache();
  var cle = 'v2|' + cible + '|fiche|' + seanceId;
  var brut = cache.get(cle);
  if (brut) return JSON.parse(brut);

  var ss = ouvrirClasseur_(cible);
  var resultat = calculerFiche_(ss, seanceId);
  cache.put(cle, JSON.stringify(resultat), CACHE_TTL_FICHE);
  return resultat;
}

// Colonnes de l'onglet Presences : A id seance, B selection (saisie
// manuelle, non utilisee par l'app), C apneiste id, D nom, E prenom —
// toujours à ces positions sur tous les classeurs. En revanche, à partir
// de F, la mise en page peut varier d'un classeur à l'autre (ex. TEST-L2
// a une colonne "Zone de confort (libellé)" entre l'ID et l'Observation,
// depuis la réorganisation du 10/09/2026 — pas les classeurs PROD) : on
// repère donc qualité/observation/contrôle par leur en-tête (ligne 4),
// jamais par un numéro de colonne fixe, pour qu'un même Code.gs partagé
// par tous les environnements reste correct quelle que soit la mise en
// page réelle du classeur ouvert (leçon de la régression du 10/09/2026,
// cf. savePresences_ V-02).
function presencesCols_(sh) {
  var header = sh.getRange(4, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = {};
  header.forEach(function (h, i) {
    var cle = String(h || '').trim();
    if (cle) col[cle] = i + 1; // 1-based, comme getRange(...)
  });
  return col;
}

function calculerFiche_(ss, seanceId) {
  var sh = ss.getSheetByName(SHEET_PRESENCES);
  var col = presencesCols_(sh);
  var iQualite = (col['Zone de confort (ID)'] || col['Zone de confort']) - 1;
  var iObservation = col['Observation de l\'encadrant'] - 1;
  var values = sh.getDataRange().getValues();
  var presences = [];
  for (var r = 4; r < values.length; r++) {
    var row = values[r];
    if (row[0] === seanceId) {
      presences.push({
        apneiste_id: row[2], nom: row[3], prenom: row[4],
        qualite: row[iQualite] || '', observation: row[iObservation] || ''
      });
    }
  }
  return { ok: true, seance_id: seanceId, presences: presences };
}

// --- Écriture -----------------------------------------------------------
// L'onglet Presences est pré-provisionné à 500 lignes, formules et mise en
// forme déjà en place (colonnes C/D/E — Apnéiste ID/Nom/Prénom déduits de
// la colonne B — et H — Contrôle) : on n'écrit jamais dans ces colonnes,
// seulement dans A (séance), B (sélection, au format qui alimente C/D/E),
// F (qualité) et G (observation), pour ne jamais abîmer une formule ou une
// couleur du modèle (06/09/2026 — avant, une réécriture complète du
// classeur en valeurs plates finissait par écraser les formules).
// Écrit le plan de séance (colonne R de Calendrier) pour la séance
// seanceId. Appelée depuis savePresences_ : un seul bouton "Enregistrer
// la séance" côté front, qui couvre présences + plan de séance (demande
// du 07/09/2026 — pas de bouton d'enregistrement séparé pour le plan).
// undefined/null : le front n'a rien envoyé (accordéon non touché) — on
// ne touche pas la cellule pour ne jamais écraser une valeur existante
// avec une chaîne vide par accident.
function ecrirePlanSeance_(ss, seanceId, plan) {
  if (plan === undefined || plan === null) return;
  var sh = ss.getSheetByName(SHEET_CALENDRIER);
  var nRows = sh.getLastRow() - 4;
  var idCol = sh.getRange(5, 1, nRows, 1).getValues().map(function (r) { return r[0]; });
  var idx = idCol.indexOf(seanceId);
  if (idx === -1) throw new Error('séance introuvable dans Calendrier : ' + seanceId);
  sh.getRange(5 + idx, 18).setValue(plan); // colonne R
}

// Écrit le remplaçant (colonne N de Calendrier, accordéon "Remplacement" côté
// front, TEST uniquement — demande du 12/09/2026). Même convention que
// ecrirePlanSeance_ : undefined/null = accordéon non touché, on ne modifie
// rien. '' = "Aucun (encadrant habituel)" sélectionné : on efface N, la
// formule existante (colonne O) retombe alors sur l'encadrant habituel du
// créneau. Un id : on écrit "Prénom Nom [id]" (même format que
// selectionTexte_, déjà utilisé pour Presences colonne B), que la formule de
// la colonne O sait déjà extraire directement entre crochets — aucune
// formule du classeur à modifier pour ce chantier.
function ecrireRemplacant_(ss, seanceId, remplacantId, roster) {
  if (remplacantId === undefined || remplacantId === null) return;
  var sh = ss.getSheetByName(SHEET_CALENDRIER);
  var nRows = sh.getLastRow() - 4;
  var idCol = sh.getRange(5, 1, nRows, 1).getValues().map(function (r) { return r[0]; });
  var idx = idCol.indexOf(seanceId);
  if (idx === -1) throw new Error('séance introuvable dans Calendrier : ' + seanceId);
  if (remplacantId === '') {
    sh.getRange(5 + idx, 14).clearContent(); // colonne N
    return;
  }
  var pers = roster[remplacantId];
  if (!pers) throw new Error('remplaçant inconnu du référentiel : ' + remplacantId);
  // V-09 : le remplaçant doit être qualifié encadrant, comme pour toute
  // affectation d'encadrement (même principe que D-15/COMPTE-07).
  if (!pers.est_encadrant) {
    throw new Error('remplaçant non qualifié encadrant : ' + remplacantId);
  }
  sh.getRange(5 + idx, 14).setValue(selectionTexte_(pers));
}

// Écrit le libellé du statut dans Calendrier. La colonne "Statut (ID)" est
// une formule : elle convertit ce libellé en ID stable depuis la plage nommée
// ListeStatutsSeance, exactement comme pour la zone de confort. Les colonnes
// de remplacement et de plan ne sont donc jamais déplacées.
function ecrireStatutSeance_(ss, seanceId, statutId) {
  if (statutId === undefined || statutId === null || statutId === '') return;
  var statuts = lireListeStatutsSeance_(ss);
  if (!statuts.length) {
    throw new Error('liste des statuts de séance introuvable : classeur non migré');
  }
  var statut = statuts.filter(function (x) { return String(x.id) === String(statutId); })[0];
  if (!statut) throw new Error('statut de séance inconnu du référentiel : ' + statutId);
  if (!statut.actif) throw new Error('statut de séance inactif : ' + statut.libelle);

  var sh = ss.getSheetByName(SHEET_CALENDRIER);
  var header = sh.getRange(4, 1, 1, sh.getLastColumn()).getValues()[0];
  var colStatut = header.indexOf('Statut') + 1;
  var colStatutId = header.indexOf('Statut (ID)') + 1;
  if (!colStatut || !colStatutId) {
    throw new Error('colonnes Statut / Statut (ID) introuvables dans Calendrier');
  }
  var nRows = sh.getLastRow() - 4;
  var ids = sh.getRange(5, 1, nRows, 1).getValues().map(function (r) { return r[0]; });
  var idx = ids.indexOf(seanceId);
  if (idx === -1) throw new Error('séance introuvable dans Calendrier : ' + seanceId);
  var ligne = 5 + idx;
  sh.getRange(ligne, colStatut).setValue(statut.libelle);
  SpreadsheetApp.flush();
  if (String(sh.getRange(ligne, colStatutId).getValue()) !== String(statut.id)) {
    throw new Error('ID du statut non calculé pour la séance : ' + seanceId);
  }
}

// Règle métier validée le 13/09/2026 : une séance passée pour laquelle au
// moins une présence vient d'être enregistrée est nécessairement "tenue".
// La comparaison porte sur la date civile dans le fuseau du script : une
// séance du jour n'est pas considérée comme passée, même si son horaire est
// terminé. Un classeur sans ListeStatutsSeance n'est pas encore migré et
// conserve son comportement historique, ce qui protège les PROD actuelles.
function forcerStatutTenueSiNecessaire_(ss, seanceId, nombrePresences) {
  if (!nombrePresences) return false;
  var statuts = lireListeStatutsSeance_(ss);
  if (!statuts.length) return false;

  var statutTenue = statuts.filter(function (x) {
    return String(x.libelle || '').toLowerCase().trim() === 'tenue';
  })[0];
  if (!statutTenue) throw new Error('statut "tenue" introuvable dans le référentiel');

  var sh = ss.getSheetByName(SHEET_CALENDRIER);
  var nRows = sh.getLastRow() - 4;
  var values = sh.getRange(5, 1, nRows, Math.max(9, sh.getLastColumn())).getValues();
  var idx = -1;
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === seanceId) { idx = i; break; }
  }
  if (idx === -1) throw new Error('séance introuvable dans Calendrier : ' + seanceId);

  var dateSeance = values[idx][1];
  if (Object.prototype.toString.call(dateSeance) !== '[object Date]') {
    Logger.log('Statut non automatisé : date invalide pour ' + seanceId);
    return false;
  }
  var tz = Session.getScriptTimeZone();
  var cleDateSeance = Utilities.formatDate(dateSeance, tz, 'yyyyMMdd');
  var cleAujourdhui = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  if (cleDateSeance >= cleAujourdhui) return false;

  var libelleActuel = String(values[idx][8] || '').toLowerCase().trim();
  if (libelleActuel === 'tenue') return false;
  ecrireStatutSeance_(ss, seanceId, statutTenue.id);
  return true;
}

function savePresences_(body) {
  var email = verifierJeton_(body.idToken);
  var seanceId = body.seance_id;
  if (!seanceId) throw new Error('seance_id manquant');
  var ss = ouvrirClasseur_(body.cible);

  var roster = {};
  lirePersonnes_(ss).forEach(function (p) { roster[p.id] = p; });

  var zonesValides = {};
  var zonesParId = {};
  lireListeZoneConfort_(ss).forEach(function (z) {
    zonesValides[z.id] = true;
    zonesParId[z.id] = z.libelle;
  });
  // Un classeur n'est « migré » vers le modèle ID+libellé que s'il a une
  // table Listes!K:M non vide (aujourd'hui, uniquement TEST-L2) — sur tout
  // autre classeur, la zone de confort reste un texte libre historique et
  // V-02 ne doit pas s'appliquer. RÉGRESSION du 09/09/2026, corrigée le
  // 10/09/2026 : comme Code.gs est un script partagé par tous les
  // environnements, la validation V-02 bloquait par erreur l'enregistrement
  // de TOUTE présence avec une zone de confort sur les classeurs PROD (leur
  // Listes!K:M est vide, donc zonesValides était vide et rejetait tout
  // texte non-ID) — signalé par Fred sur PROD-L3.
  var zoneConfortMigree = Object.keys(zonesValides).length > 0;

  var presences = (body.presences || []).filter(function (p) { return p.present; });
  presences.forEach(function (p) {
    // V-01 : la clé étrangère doit exister dans le référentiel — jamais de
    // confiance aveugle dans ce que le client envoie.
    if (!roster[p.apneiste_id]) {
      throw new Error('apnéiste inconnu du référentiel : ' + p.apneiste_id);
    }
    // V-02 : idem pour la zone de confort — désormais un ID (cf.
    // lireListeZoneConfort_) sur les classeurs migrés uniquement
    // (09/09/2026, TEST-L2 pour l'instant) ; ailleurs, texte libre non
    // validé ici, comme avant ce chantier.
    if (zoneConfortMigree && p.qualite !== '' && p.qualite !== undefined && p.qualite !== null &&
        !zonesValides[p.qualite]) {
      throw new Error('zone de confort inconnue du référentiel : ' + p.qualite);
    }
  });

  var sh = ss.getSheetByName(SHEET_PRESENCES);
  // Qualité et observation ne sont pas forcément dans des colonnes
  // adjacentes ni à une position fixe (cf. presencesCols_ / calculerFiche_
  // ci-dessus) : sur TEST-L2, "Zone de confort (libellé)" (une formule,
  // jamais écrite ici) s'intercale désormais entre les deux — on écrit
  // donc chaque colonne séparément plutôt que sur une plage de 2 colonnes
  // contiguës comme avant le 10/09/2026.
  var colonnesPresences = presencesCols_(sh);
  var colQualite = colonnesPresences['Zone de confort'] || colonnesPresences['Zone de confort (ID)'];
  var colObservation = colonnesPresences['Observation de l\'encadrant'];
  var nRows = sh.getLastRow() - 4;
  var idCol = sh.getRange(5, 1, nRows, 1).getValues().map(function (r) { return r[0]; });
  var selCol = sh.getRange(5, 2, nRows, 1).getValues().map(function (r) { return r[0]; });

  // lignes déjà liées à cette séance (réutilisables en priorité), puis
  // lignes totalement libres.
  var deSeance = [], vides = [];
  for (var i = 0; i < nRows; i++) {
    if (idCol[i] === seanceId) deSeance.push(i);
    else if (!idCol[i] && !selCol[i]) vides.push(i);
  }
  var disponibles = deSeance.concat(vides);
  if (disponibles.length < presences.length) {
    throw new Error('capacité de présences atteinte (' + nRows +
                     ' lignes) — contactez l’administrateur');
  }

  presences.forEach(function (p, i) {
    var row = 5 + disponibles[i];
    var pers = roster[p.apneiste_id];
    sh.getRange(row, 1, 1, 2).setValues([[seanceId, selectionTexte_(pers)]]);
    // Sur TEST-L2 migré vers la saisie par libellé, l'application écrit le
    // libellé visible en colonne de saisie. La colonne "Zone de confort
    // (ID)" contient une formule et reste la donnée technique de référence.
    // Les autres classeurs conservent leur comportement historique.
    var qualiteEcrite = colonnesPresences['Zone de confort'] && zoneConfortMigree
      ? (p.qualite === '' || p.qualite === undefined || p.qualite === null ? '' : zonesParId[p.qualite])
      : (p.qualite || '');
    sh.getRange(row, colQualite).setValue(qualiteEcrite);
    sh.getRange(row, colObservation).setValue(p.observation || '');
  });
  // lignes qui appartenaient à la séance mais ne sont plus nécessaires —
  // vidées (jamais les colonnes C/D/E/H, formules).
  for (var j = presences.length; j < deSeance.length; j++) {
    var r2 = 5 + deSeance[j];
    sh.getRange(r2, 1, 1, 2).clearContent();
    sh.getRange(r2, colQualite).clearContent();
    sh.getRange(r2, colObservation).clearContent();
  }

  ecrirePlanSeance_(ss, seanceId, body.plan_seance);
  ecrireRemplacant_(ss, seanceId, body.remplacant_id, roster);
  ecrireStatutSeance_(ss, seanceId, body.statut_id);
  var statutForceTenue = forcerStatutTenueSiNecessaire_(ss, seanceId, presences.length);

  // rafraîchit immédiatement le cache "fiche" de cette séance : la
  // prochaine lecture (même appareil ou un autre) voit tout de suite le
  // résultat de cet enregistrement, sans attendre l'expiration du TTL
  // (demande du 06/09/2026).
  var cache = CacheService.getScriptCache();
  cache.put('v2|' + body.cible + '|fiche|' + seanceId,
            JSON.stringify(calculerFiche_(ss, seanceId)), CACHE_TTL_FICHE);
  // le plan de séance et le remplaçant font partie de la réponse "data"
  // (lireCalendrier_) : on invalide ce cache aussi, sinon le prochain
  // chargement de la page renverrait encore l'ancienne valeur pendant
  // CACHE_TTL_DONNEES.
  cache.remove('v2|' + body.cible + '|data');
  cache.remove('v2|' + body.cible + '|stats');

  return {
    ok: true, saved: presences.length, par: email,
    statut_force_tenue: statutForceTenue
  };
}

function selectionTexte_(pers) {
  return pers.nom ? (pers.prenom + ' ' + pers.nom + ' [' + pers.id + ']')
                  : (pers.prenom + ' [' + pers.id + ']');
}

function formatDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(v || '');
}
// ============================================================================
// Synchronisation des effectifs depuis le classeur de Paramétrage centralisé
// ----------------------------------------------------------------------------
// Principe validé le 06/09/2026 avec Fred : les affectations aux lignes
// (Inscriptions) sont saisies UNE SEULE FOIS dans le classeur de Paramétrage
// (Google Sheet), puis propagées ici vers l'onglet Personnes de chaque
// classeur de ligne — sans ressaisie, sans lien Excel fragile entre fichiers.
//
// Portée de cette première version : uniquement l'onglet Personnes (colonnes
// C à G : Nom, Prénom, Rôle déclaré, Groupe(s), Objectif de la saison). Les
// colonnes A (Code) et H (Membre de la ligne) restent des formules propres à
// chaque classeur, jamais touchées ici — elles se recalculent seules.
//
// Cette même logique existe déjà, testée, dans generateurs/sync_referentiel.py
// (variante hors ligne, sur fichiers .xlsx locaux) : ce bloc en est le
// portage fidèle pour des classeurs qui sont maintenant des Google Sheets.
// La synchronisation du Référentiel (Objectifs, Compétences, Créneaux,
// Responsables, Qualifications Sécurité) et l'initialisation du Calendrier,
// que sync_referentiel.py fait aussi, ne sont PAS encore portées ici : leur
// mise en page dans les classeurs de ligne actuellement déployés n'a pas été
// vérifiée ligne à ligne, contrairement à Personnes. À faire dans un second
// temps, après vérification.
//
// Identifiant du classeur de Paramétrage centralisé (Google Sheet, racine de
// Mon Drive). À mettre à jour si ce classeur est un jour recréé ailleurs.
var ID_PARAMETRAGE = '18vMX5fqFgCN7NrSkPsSr1ftVoonPFC5xEbjf73iM8Lw';

// Lignes à synchroniser : toutes les lignes réellement configurées dans
// CLASSEURS (PROD-*) ci-dessus. Chaque classeur reçoit aussi les encadrants
// des autres lignes : ils sont ainsi disponibles comme remplaçants.
// STAC sera ajoutée lors de la création de son classeur et de ses cibles.
var LIGNES_SYNC = ['L1', 'L2', 'L3', 'L4', 'LC', 'DNF1', 'DNF2', 'STA1', 'STA2'];

// ---------------------------------------------------------- lecture Parametrage
// Lit tout ce qu'il faut du classeur de Paramétrage : le référentiel des
// personnes du club (Personnes) et les affectations aux lignes (Inscriptions,
// filtrées sur les inscriptions actives de la saison en cours). Contrairement
// à sync_referentiel.py (qui lit un fichier .xlsx figé avec openpyxl et doit
// donc composer avec des formules non recalculées), Apps Script recalcule
// toujours les formules à la lecture : pas besoin de repli sur l'identifiant
// entre crochets d'une sélection.
function lireParametrage_() {
  var ssp = SpreadsheetApp.openById(ID_PARAMETRAGE);

  // Saison courante déduite du nom du classeur (« AA - Parametrage 2026-2027 »)
  // plutôt que codée en dur : pas de modification de script à chaque saison.
  var mSaison = /(\d{4}-\d{4})/.exec(ssp.getName());
  var saison = mSaison ? mSaison[1] : '';

  var personnes = {};
  var ordreClub = [];
  var vp = ssp.getSheetByName('Personnes').getDataRange().getValues();
  for (var r = 4; r < vp.length; r++) {           // ligne 5 = première donnée
    var pid = vp[r][0];
    if (!pid) continue;
    personnes[pid] = { id: pid, nom: vp[r][1], prenom: vp[r][2] };
    ordreClub.push(pid);
  }

  var aujourdhui = new Date();
  aujourdhui.setHours(0, 0, 0, 0);

  var vi = ssp.getSheetByName('Inscriptions').getDataRange().getValues();
  var inscriptions = [];
  for (var r2 = 4; r2 < vi.length; r2++) {
    var row = vi[r2];
    var pidI = row[2];                            // colonne C : Personne (ID)
    if (!pidI) continue;
    var saisonRow = row[0];                       // colonne A : Saison
    if (saisonRow && saison && saisonRow !== saison) continue;
    var fin = (row[9] instanceof Date) ? row[9] : null;   // colonne J : Date de fin
    inscriptions.push({
      id: pidI,
      groupe: row[5],                              // colonne F : Groupe de niveau
      role: String(row[6] || '').toLowerCase().trim(), // colonne G : Rôle
      objectif: row[10] || '',                     // colonne K : Objectif de la saison
      fin: fin
    });
  }

  function estActive(x) { return !x.fin || x.fin >= aujourdhui; }

  // Les membres de LC (compétition) sont aussi membres de L4 : une personne
  // active en LC sans inscription L4 active reçoit une inscription L4/élève
  // synthétique, en mémoire seulement (rien n'est écrit dans Paramétrage),
  // pour que le roster de L4 les intègre automatiquement — même règle que
  // sync_referentiel.py (décision du 06/09/2026).
  var lcActifs = {}, l4Actifs = {};
  inscriptions.forEach(function (x) {
    if (!estActive(x)) return;
    if (x.groupe === 'LC') lcActifs[x.id] = true;
    if (x.groupe === 'L4') l4Actifs[x.id] = true;
  });
  Object.keys(lcActifs).sort().forEach(function (pid) {
    if (!l4Actifs[pid]) {
      inscriptions.push({ id: pid, groupe: 'L4', role: 'élève', objectif: '', fin: null });
    }
  });

  return { saison: saison, personnes: personnes, ordreClub: ordreClub,
           inscriptions: inscriptions, estActive: estActive };
}

// Toutes les inscriptions actives (aujourd'hui) d'une personne, tous groupes
// confondus — alimente la colonne « Groupe(s) », club entier.
function groupesActifs_(par, pid) {
  var groupes = {};
  par.inscriptions.forEach(function (x) {
    if (x.id === pid && x.groupe && par.estActive(x)) groupes[x.groupe] = true;
  });
  return Object.keys(groupes).sort().join(' / ');
}

// Groupes dans lesquels une personne exerce un rôle donné. La colonne
// Groupe(s) doit suivre le rôle de la ligne Personnes : cela permet à la
// formule "Membre de la ligne" de distinguer, par exemple, un élève L2 qui
// est encadrant dans une autre ligne.
function groupesActifsRole_(par, pid, role) {
  var groupes = {};
  par.inscriptions.forEach(function (x) {
    if (x.id === pid && x.role === role && x.groupe && par.estActive(x)) {
      groupes[x.groupe] = true;
    }
  });
  return Object.keys(groupes).sort().join(' / ');
}

// Dernière ligne, à partir de `first`, où la colonne `col` porte une formule
// (et non une simple valeur ou une cellule vide) — sert à savoir jusqu'où les
// formules Code/Membre sont déjà provisionnées dans l'onglet Personnes d'un
// classeur de ligne, sans supposer une capacité fixe.
function derniereLigneFormule_(sh, col, first, maxScan) {
  var n = Math.min(maxScan, sh.getMaxRows() - first + 1);
  if (n <= 0) return first - 1;
  var formules = sh.getRange(first, col, n, 1).getFormulas();
  var last = first - 1;
  for (var i = 0; i < formules.length; i++) {
    if (formules[i][0]) last = first + i;
  }
  return last;
}

// -------------------------------------------------- synchronisation d'une ligne
function synchroniserPersonnesLigne_(par, ligne, cible) {
  var ss = ouvrirClasseur_(cible || ('PROD-' + ligne));
  var sh = ss.getSheetByName(SHEET_PERSONNES);
  var PERS_FIRST_L = 5;

  // 1. Ordre déjà en place (colonnes B ID, E Rôle), avant toute écriture —
  //    c'est ce qui permet de préserver la position (donc le Code, colonne A)
  //    des personnes déjà connues. Jamais d'insertion au milieu : les
  //    nouvelles arrivées sont ajoutées à la fin (même règle que
  //    sync_referentiel.py — un « 7 » tapé en octobre ne doit pas désigner
  //    quelqu'un d'autre en mars).
  var lastRowActuelle = sh.getLastRow();
  var ordreExistant = [];
  if (lastRowActuelle >= PERS_FIRST_L) {
    var existant = sh.getRange(PERS_FIRST_L, 2, lastRowActuelle - PERS_FIRST_L + 1, 4).getValues();
    existant.forEach(function (row) {
      var pid = row[0];                            // colonne B
      if (pid) ordreExistant.push([pid, row[3] || '']);  // colonne E (index 3 de la plage)
    });
  }

  // 2. Toutes les inscriptions actives du club : une entrée par
  //    (personne, rôle). Les groupes sont regroupés dans la colonne F pour
  //    ce même rôle. Ainsi chaque classeur connaît tous les encadrants, même
  //    lorsqu'ils encadrent habituellement une autre ligne.
  var parCle = {};      // clé "id|role" -> inscription
  par.inscriptions.forEach(function (x) {
    if (!par.estActive(x) || !par.personnes[x.id] || !x.role) return;
    var cle = x.id + '|' + x.role;
    if (!parCle[cle]) parCle[cle] = x;
  });
  var clesActuelles = Object.keys(parCle);
  var clesSet = {};
  clesActuelles.forEach(function (c) { clesSet[c] = true; });

  var connues = [], dejaVu = {};
  ordreExistant.forEach(function (pair) {
    var cle = pair[0] + '|' + pair[1];
    if (clesSet[cle] && !dejaVu[cle]) { connues.push(cle); dejaVu[cle] = true; }
  });
  var nouvelles = clesActuelles.filter(function (c) { return !dejaVu[c]; });
  nouvelles.sort(function (a, b) {
    var xa = parCle[a], xb = parCle[b];
    var na = (par.personnes[xa.id].nom + par.personnes[xa.id].prenom + xa.role).toUpperCase();
    var nb = (par.personnes[xb.id].nom + par.personnes[xb.id].prenom + xb.role).toUpperCase();
    return na < nb ? -1 : (na > nb ? 1 : 0);
  });
  var inscriptionsActives = connues.concat(nouvelles);
  var personnesAvecRole = {};
  inscriptionsActives.forEach(function (c) { personnesAvecRole[parCle[c].id] = true; });

  // 3. Le reste du club, sans rôle dans cette ligne : présent quand même
  //    (un encadrant n'est pas nécessairement inscrit dans la ligne qu'il
  //    encadre — l'onglet Personnes porte tout le club).
  var autres = par.ordreClub.filter(function (pid) { return !personnesAvecRole[pid]; });

  var ordre = inscriptionsActives.map(function (c) { return { pid: parCle[c].id, x: parCle[c] }; })
    .concat(autres.map(function (pid) { return { pid: pid, x: null }; }));

  // 4. S'assurer que les formules Code/Membre (colonnes A et H) couvrent
  //    assez de lignes pour tout le monde : si le club a grandi au-delà de
  //    la capacité déjà provisionnée dans ce classeur, on prolonge les
  //    formules par recopie de la dernière ligne formulée — jamais de
  //    capacité supposée à l'avance.
  var maxScan = 1000;
  var derniereA = derniereLigneFormule_(sh, 1, PERS_FIRST_L, maxScan);
  var derniereH = derniereLigneFormule_(sh, 8, PERS_FIRST_L, maxScan);
  var capaciteActuelle = Math.max(derniereA, derniereH, PERS_FIRST_L - 1) - PERS_FIRST_L + 1;
  if (ordre.length > capaciteActuelle && capaciteActuelle > 0) {
    var manquantes = ordre.length - capaciteActuelle;
    var deLigne = PERS_FIRST_L + capaciteActuelle - 1;   // dernière ligne déjà formulée
    sh.getRange(deLigne, 1).copyTo(
      sh.getRange(deLigne + 1, 1, manquantes, 1));
    sh.getRange(deLigne, 8).copyTo(
      sh.getRange(deLigne + 1, 8, manquantes, 1));
    capaciteActuelle = ordre.length;
  }

  // 5. Réécrire les colonnes B à G (jamais A ni H) sur toute la
  //    plage couverte par les formules (au moins ordre.length, au moins
  //    l'ancienne étendue, pour ne pas laisser de lignes fantômes si le club
  //    a rétréci).
  var nLignes = Math.max(capaciteActuelle, ordre.length, lastRowActuelle - PERS_FIRST_L + 1);
  if (nLignes > 0) {
    var valeurs = [];
    for (var i = 0; i < nLignes; i++) {
      if (i < ordre.length) {
        var o = ordre[i];
        var p = par.personnes[o.pid];
        var roleVal = o.x ? o.x.role : '';
        var objectifVal = o.x ? o.x.objectif : '';
        var groupesVal = o.x ? groupesActifsRole_(par, o.pid, o.x.role) : groupesActifs_(par, o.pid);
        valeurs.push([o.pid, p.nom, p.prenom, roleVal, groupesVal, objectifVal]);
      } else {
        valeurs.push(['', '', '', '', '', '']);
      }
    }
    sh.getRange(PERS_FIRST_L, 2, nLignes, 6).setValues(valeurs);
  }

  var membres = inscriptionsActives.filter(function (c) {
    return groupesActifsRole_(par, parCle[c].id, parCle[c].role).split(' / ').indexOf(ligne) !== -1;
  }).length;
  return { ligne: ligne, personnes: ordre.length, membres: membres,
           formulesProlongees: (capaciteActuelle > (derniereA - PERS_FIRST_L + 1)) };
}

// Point d'entrée manuel : à lancer depuis l'éditeur Apps Script (menu
// « Exécuter »), en sélectionnant synchroniserEffectifs. Recommandé de
// relire les onglets Personnes des classeurs de ligne après une première
// exécution avant d'envisager un menu ou un déclencheur automatique.
function synchroniserEffectifs() {
  var par = lireParametrage_();
  var resultats = [];
  LIGNES_SYNC.forEach(function (ligne) {
    try {
      resultats.push(synchroniserPersonnesLigne_(par, ligne));
    } catch (e) {
      resultats.push({ ligne: ligne, erreur: String(e) });
    }
  });
  Logger.log(JSON.stringify(resultats, null, 2));
  return resultats;
}

// Point d'entrée de validation : synchronise uniquement le classeur TEST-L2.
// La synchronisation de la PROD reste une action distincte et manuelle.
function synchroniserEffectifsTestL2() {
  var par = lireParametrage_();
  var resultat = synchroniserPersonnesLigne_(par, 'L2', 'TEST-L2');
  Logger.log(JSON.stringify(resultat, null, 2));
  return resultat;
}

// Points d'entrée ciblés pour valider la synchronisation des lignes de
// séances sans toucher aux classeurs de production.
function synchroniserEffectifsTestDnf1() {
  var par = lireParametrage_();
  var resultat = synchroniserPersonnesLigne_(par, 'DNF1', 'TEST-DNF1');
  Logger.log(JSON.stringify(resultat, null, 2));
  return resultat;
}

function synchroniserEffectifsTestDnf2() {
  var par = lireParametrage_();
  var resultat = synchroniserPersonnesLigne_(par, 'DNF2', 'TEST-DNF2');
  Logger.log(JSON.stringify(resultat, null, 2));
  return resultat;
}

function synchroniserEffectifsTestSta1() {
  var par = lireParametrage_();
  var resultat = synchroniserPersonnesLigne_(par, 'STA1', 'TEST-STA1');
  Logger.log(JSON.stringify(resultat, null, 2));
  return resultat;
}

function synchroniserEffectifsTestSta2() {
  var par = lireParametrage_();
  var resultat = synchroniserPersonnesLigne_(par, 'STA2', 'TEST-STA2');
  Logger.log(JSON.stringify(resultat, null, 2));
  return resultat;
}

// Synchronise uniquement les lignes de séances déjà livrées en production.
// Cette entrée évite de relancer les autres lignes lorsque le besoin porte
// exclusivement sur la disponibilité des remplaçants DNF/STA.
function synchroniserEffectifsLignesSeancesProd() {
  var par = lireParametrage_();
  var resultats = [];
  ['DNF1', 'DNF2', 'STA1', 'STA2'].forEach(function (ligne) {
    try {
      resultats.push(synchroniserPersonnesLigne_(par, ligne));
    } catch (e) {
      resultats.push({ ligne: ligne, erreur: String(e) });
    }
  });
  Logger.log(JSON.stringify(resultats, null, 2));
  return resultats;
}

// Migration ponctuelle du référentiel central : le rôle « prépa encadrant »
// est proposé dans Inscriptions et remplace le rôle encadrant pour les trois
// personnes désignées. Le serveur traite ce rôle comme un encadrant pour les
// remplacements et les listes associées.
function preparerRolesPrepaEncadrant() {
  var ssp = SpreadsheetApp.openById(ID_PARAMETRAGE);
  var listes = ssp.getSheetByName(SHEET_LISTES);
  var inscriptions = ssp.getSheetByName('Inscriptions');
  if (!listes || !inscriptions) throw new Error('onglet Listes ou Inscriptions introuvable');

  listes.getRange('E8').setValue('prépa encadrant');
  var premiereLigne = 5;
  var n = inscriptions.getMaxRows() - premiereLigne + 1;
  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['élève', 'encadrant', 'prépa encadrant'], true)
    .setAllowInvalid(false)
    .build();
  inscriptions.getRange(premiereLigne, 7, n, 1).setDataValidation(validation);

  var cibles = {
    'P013|L4': 'Francis Wang',
    'P052|L2': 'Corinne Le Brigand',
    'P069|L3': 'Magali Cavatore'
  };
  var valeurs = inscriptions.getRange(premiereLigne, 3, n, 5).getValues();
  var lignes = {};
  valeurs.forEach(function (r, i) {
    var cle = String(r[0] || '') + '|' + String(r[3] || '');
    if (cibles[cle] && (r[4] === 'encadrant' || r[4] === 'prépa encadrant')) {
      lignes[cle] = premiereLigne + i;
    }
  });
  var manquantes = Object.keys(cibles).filter(function (cle) { return !lignes[cle]; });
  if (manquantes.length) throw new Error('inscription encadrant introuvable : ' + manquantes.join(', '));
  Object.keys(lignes).forEach(function (cle) {
    inscriptions.getRange(lignes[cle], 7).setValue('prépa encadrant');
  });
  var resultat = { role: 'prépa encadrant', personnes: Object.keys(cibles).map(function (cle) { return cibles[cle]; }) };
  Logger.log(JSON.stringify(resultat, null, 2));
  return resultat;
}

// --- Migration ponctuelle : mise en page de l'onglet Listes + colonne ----
// libellé dans Presences (demande de Fred du 10/09/2026) ------------------
//
// Réorganise l'onglet "Listes" du classeur `cible` : chaque liste devient
// un bloc titre / en-tête(s) de colonne / valeurs, séparé du suivant par
// une colonne vide (au lieu des 9 listes historiques posées côte à côte
// sans séparation, et de la table Zone de confort en K:M sans distinction
// entre titre et en-tête). Crée/rafraîchit la plage nommée
// "ListeZoneConfort" sur la nouvelle position du tableau ID/Libellé/Actif
// (cf. lireListeZoneConfort_, qui s'y réfère par nom et non par
// coordonnées). Ajoute aussi, dans l'onglet "Presences", une colonne
// "Zone de confort (libellé)" juste après la colonne ID (au lieu d'une
// colonne isolée en fin de tableau) — formule VLOOKUP sur cette plage
// nommée, pour les 500 lignes pré-provisionnées.
//
// Fonction de maintenance, à lancer à la main depuis l'éditeur Apps Script
// (jamais appelée par doGet/doPost), comme synchroniserEffectifs. Conçue
// pour être rejouée sur d'autres classeurs le jour où Fred validera une
// extension du périmètre (TEST-L3, PROD) — mais à ce jour, lancée
// uniquement sur TEST-L2 (périmètre validé par Fred le 10/09/2026).
function reorganiserListesEtPresences(cible) {
  if (!cible) throw new Error('cible manquante, ex. reorganiserListesEtPresences("TEST-L2")');
  var ss = ouvrirClasseur_(cible);
  // Garde-fou (10/09/2026) : cette fonction relit l'onglet Listes en supposant
  // la mise en page D'ORIGINE (une seule ligne de titre par liste, valeurs
  // juste en-dessous). Une fois la migration faite, l'onglet a une ligne de
  // titre PUIS une ligne d'en-tête "Valeur"/"ID"... : la relire avec les
  // mêmes hypothèses décale tout et corrompt les listes (incident constaté
  // sur TEST-L2, corrigé par restauration d'une version antérieure du
  // classeur). Donc : jamais une deuxième fois sur un classeur déjà migré.
  if (ss.getRangeByName('ListeZoneConfort')) {
    throw new Error('classeur ' + cible + ' déjà migré (plage ListeZoneConfort ' +
      'existante) - reorganiserListesEtPresences ne doit être lancée qu\'une ' +
      'seule fois par classeur. Pour corriger uniquement les formules de la ' +
      'colonne "Zone de confort (libellé)", utiliser ' +
      'corrigerFormulesLibellePresences(cible) à la place.');
  }
  var shL = ss.getSheetByName(SHEET_LISTES);

  // 1) Capture des listes existantes (colonnes A à I) avant de tout effacer.
  var definitions = [
    { titre: 'Jours', col: 1 },
    { titre: 'Niveaux de compétence', col: 2 },
    { titre: 'Statut d\'acquisition', col: 3 },
    { titre: 'Qualité de réalisation', col: 4 },
    { titre: 'Natures de séance', col: 5 },
    { titre: 'Statuts de séance', col: 6 },
    { titre: 'Avis d\'encadrant', col: 7 },
    { titre: 'Types d\'événement', col: 8 },
    { titre: 'Gravités', col: 9 }
  ];
  var anciennesListes = definitions.map(function (l) {
    var lastRow = derniereLigneNonVide_(shL, l.col);
    var valeurs = lastRow > 1
      ? shL.getRange(2, l.col, lastRow - 1, 1).getValues().map(function (r) { return r[0]; })
      : [];
    return { titre: l.titre, valeurs: valeurs };
  });
  var lastRowZC = derniereLigneNonVide_(shL, 11); // colonne K = ID (ancienne position)
  var zoneConfort = lastRowZC > 1 ? shL.getRange(2, 11, lastRowZC - 1, 3).getValues() : [];

  var fondTitre = shL.getRange('A1').getBackground();
  var couleurTitre = shL.getRange('A1').getFontColor();

  // 2) Reconstruction : un bloc par liste, séparé du suivant par une
  // colonne vide (titre en ligne 1, en-tête(s) de colonne en ligne 2,
  // valeurs à partir de la ligne 3).
  shL.clear();
  var blocs = [];
  var col = 1;
  anciennesListes.forEach(function (l) {
    shL.getRange(1, col).setValue(l.titre);
    shL.getRange(2, col).setValue('Valeur');
    if (l.valeurs.length) {
      shL.getRange(3, col, l.valeurs.length, 1).setValues(l.valeurs.map(function (v) { return [v]; }));
    }
    blocs.push({ colDebut: col, largeur: 1 });
    col += 2; // 1 colonne de liste + 1 colonne d'écart
  });
  var colDebutZC = col;
  shL.getRange(1, colDebutZC).setValue('Zone de confort');
  shL.getRange(2, colDebutZC, 1, 3).setValues([['ID', 'Libellé', 'Actif']]);
  if (zoneConfort.length) {
    shL.getRange(3, colDebutZC, zoneConfort.length, 3).setValues(zoneConfort);
  }
  blocs.push({ colDebut: colDebutZC, largeur: 3 });

  // Titre et en-têtes repris du style d'origine (fond/couleur de l'ancienne
  // ligne 1), pour rester visuellement cohérent avec le reste du classeur.
  blocs.forEach(function (b) {
    shL.getRange(1, b.colDebut, 1, b.largeur)
      .setBackground(fondTitre).setFontColor(couleurTitre).setFontWeight('bold');
    shL.getRange(2, b.colDebut, 1, b.largeur)
      .setFontWeight('bold').setBackground('#f3f3f3');
  });

  // Plage nommée : Code.gs (lireListeZoneConfort_) s'y réfère par nom,
  // jamais par coordonnées — ne dépend donc plus de la mise en page de cet
  // onglet, contrairement à avant le 10/09/2026.
  var nomPlage = 'ListeZoneConfort';
  if (ss.getRangeByName(nomPlage)) ss.removeNamedRange(nomPlage);
  if (zoneConfort.length) {
    ss.setNamedRange(nomPlage, shL.getRange(3, colDebutZC, zoneConfort.length, 3));
  }

  // 3) Onglet Presences : colonne "Zone de confort (libellé)" juste après
  // la colonne ID — et nettoyage d'une éventuelle colonne du même nom déjà
  // présente ailleurs (tentative précédente, jamais alimentée).
  var shP = ss.getSheetByName(SHEET_PRESENCES);
  var header = shP.getRange(4, 1, 1, shP.getLastColumn()).getValues()[0];
  var colIdActuelle = header.indexOf('Zone de confort (ID)') + 1;
  if (!colIdActuelle) throw new Error('colonne "Zone de confort (ID)" introuvable dans Presences');
  // Idempotent : si la colonne existe déjà juste après l'ID (ex. deuxième
  // exécution de cette fonction), on la réutilise au lieu d'en re-insérer
  // une (qui aurait sinon décalé une seconde fois Observation/Contrôle).
  var colLibelleExistante = header.indexOf('Zone de confort (libellé)') + 1;
  var colLibelle;
  if (colLibelleExistante && colLibelleExistante !== colIdActuelle + 1) {
    shP.deleteColumn(colLibelleExistante); // toujours après colIdActuelle : ne la décale pas
    colLibelleExistante = 0;
  }
  if (colLibelleExistante === colIdActuelle + 1) {
    colLibelle = colLibelleExistante;
  } else {
    shP.insertColumnAfter(colIdActuelle);
    colLibelle = colIdActuelle + 1;
    var refFormat = shP.getRange(4, colIdActuelle);
    shP.getRange(4, colLibelle).setValue('Zone de confort (libellé)')
      .setFontWeight(refFormat.getFontWeight())
      .setBackground(refFormat.getBackground())
      .setFontColor(refFormat.getFontColor());
  }

  // Séparateur point-virgule : classeur en locale fr_FR (confirmé par
  // diagnostic le 10/09/2026) - la virgule, bien qu'acceptée sans erreur par
  // Range.setFormulas/getFormula (qui restent en syntaxe US), produit une
  // #ERROR! ("Erreur d'analyse de formule") à l'évaluation par le moteur de
  // calcul dans cette locale. Le point-virgule fonctionne dans les deux cas.
  var nRows = shP.getMaxRows() - 4;
  var formules = [];
  for (var r = 0; r < nRows; r++) {
    var celluleId = shP.getRange(5 + r, colIdActuelle).getA1Notation();
    formules.push(['=IFERROR(VLOOKUP(' + celluleId + ';' + nomPlage + ';2;FALSE);"")']);
  }
  shP.getRange(5, colLibelle, nRows, 1).setFormulas(formules);

  var resultat = {
    listes: anciennesListes.map(function (l) { return l.titre + ' (' + l.valeurs.length + ' valeurs)'; }),
    zoneConfort: zoneConfort.length + ' valeurs, plage ' +
      (zoneConfort.length ? shL.getRange(3, colDebutZC, zoneConfort.length, 3).getA1Notation() : '(aucune)'),
    colonneLibellePresences: colLibelle,
    presencesLignesFormulees: nRows
  };
  Logger.log(JSON.stringify(resultat, null, 2));
  return resultat;
}

// Dernière ligne non vide d'une colonne donnée (1-based), 0 si la colonne
// est entièrement vide. Utilisé par reorganiserListesEtPresences pour ne
// recopier que les valeurs réellement saisies de chaque liste, sans
// supposer une hauteur fixe (les listes n'ont pas toutes le même nombre
// de valeurs).
function derniereLigneNonVide_(sh, col) {
  var nRows = sh.getLastRow();
  if (nRows < 1) return 0;
  var values = sh.getRange(1, col, nRows, 1).getValues();
  var last = 0;
  for (var r = 0; r < values.length; r++) {
    if (values[r][0] !== '' && values[r][0] !== null) last = r + 1;
  }
  return last;
}

// Corrige uniquement les formules de la colonne "Zone de confort (libellé)"
// dans Presences, sans toucher à l'onglet Listes ni à la position de cette
// colonne (contrairement à reorganiserListesEtPresences, rejouable sans
// risque). À utiliser si les formules doivent être régénérées (ex. après une
// restauration de version, ou une correction du séparateur ',' vs ';' -
// cf. 10/09/2026).
function corrigerFormulesLibellePresences(cible) {
  if (!cible) throw new Error('cible manquante, ex. corrigerFormulesLibellePresences("TEST-L2")');
  var ss = ouvrirClasseur_(cible);
  var nomPlage = 'ListeZoneConfort';
  if (!ss.getRangeByName(nomPlage)) throw new Error('plage nommée ' + nomPlage + ' introuvable sur ' + cible);
  var shP = ss.getSheetByName(SHEET_PRESENCES);
  var header = shP.getRange(4, 1, 1, shP.getLastColumn()).getValues()[0];
  var colIdActuelle = header.indexOf('Zone de confort (ID)') + 1;
  var colLibelle = header.indexOf('Zone de confort (libellé)') + 1;
  if (!colIdActuelle || !colLibelle) throw new Error('colonnes Zone de confort introuvables dans Presences');
  var nRows = shP.getMaxRows() - 4;
  var formules = [];
  for (var r = 0; r < nRows; r++) {
    var celluleId = shP.getRange(5 + r, colIdActuelle).getA1Notation();
    formules.push(['=IFERROR(VLOOKUP(' + celluleId + ';' + nomPlage + ';2;FALSE);"")']);
  }
  shP.getRange(5, colLibelle, nRows, 1).setFormulas(formules);
  var resultat = { colonneLibellePresences: colLibelle, presencesLignesFormulees: nRows };
  Logger.log(JSON.stringify(resultat));
  return resultat;
}

// Migration ciblée TEST-L2 vers une saisie humaine par libellé : la colonne
// F affiche une liste déroulante (Confort / Challenge / Limite) et la colonne
// G calcule l'ID stable correspondant. Les ID déjà enregistrés en F servent
// de source de vérité pendant la migration, afin de préserver l'historique.
function migrerSaisieZoneConfortParLibelle(cible) {
  if (cible !== 'TEST-L2') {
    throw new Error('migration autorisée uniquement pour TEST-L2');
  }
  var ss = ouvrirClasseur_(cible);
  var liste = ss.getRangeByName('ListeZoneConfort');
  if (!liste) throw new Error('plage ListeZoneConfort introuvable');
  var sh = ss.getSheetByName(SHEET_PRESENCES);
  var col = presencesCols_(sh);
  if (col['Zone de confort'] && col['Zone de confort (ID)']) {
    return {
      cible: cible,
      colonneSaisie: col['Zone de confort'],
      colonneId: col['Zone de confort (ID)'],
      dejaMigree: true
    };
  }
  var colId = col['Zone de confort (ID)'];
  var colLibelle = col['Zone de confort (libellé)'];
  if (!colId || !colLibelle) {
    throw new Error('structure initiale ID/libellé introuvable');
  }

  var valeursListe = liste.getValues();
  var libelleParId = {};
  var libellesActifs = [];
  valeursListe.forEach(function (row) {
    if (row[0] === '' || row[0] === null) return;
    libelleParId[String(row[0])] = row[1];
    if (row[2] === 'oui') libellesActifs.push(row[1]);
  });
  if (!libellesActifs.length) throw new Error('aucune zone de confort active');

  var nRows = sh.getMaxRows() - 4;
  var idsExistants = sh.getRange(5, colId, nRows, 1).getValues();
  var libelles = idsExistants.map(function (row) {
    var id = row[0];
    if (id === '' || id === null) return [''];
    var libelle = libelleParId[String(id)];
    if (!libelle) throw new Error('ID de zone inconnu dans Presences : ' + id);
    return [libelle];
  });

  sh.getRange(4, colId).setValue('Zone de confort');
  sh.getRange(4, colLibelle).setValue('Zone de confort (ID)');
  sh.getRange(5, colId, nRows, 1).setValues(libelles);

  var formules = [];
  for (var r = 0; r < nRows; r++) {
    var celluleLibelle = sh.getRange(5 + r, colId).getA1Notation();
    formules.push(['=IFERROR(INDEX(ListeZoneConfort;MATCH(' + celluleLibelle +
      ';INDEX(ListeZoneConfort;0;2);0);1);"")']);
  }
  sh.getRange(5, colLibelle, nRows, 1).setFormulas(formules);

  return {
    cible: cible,
    colonneSaisie: colId,
    colonneId: colLibelle,
    lignesMigrees: nRows,
    valeurs: libellesActifs
  };
}

function migrerSaisieZoneConfortTestL2() {
  return migrerSaisieZoneConfortParLibelle('TEST-L2');
}

// Migration des modèles DNF/STA fraîchement générés vers le même principe que
// TEST-L2 : le libellé reste lisible dans Presences, l'ID stable est calculé
// dans la colonne voisine. Limité explicitement aux deux classeurs de TEST.
function migrerZoneConfortLignesTest_(cible) {
  if (['TEST-DNF1', 'TEST-DNF2', 'TEST-STA1', 'TEST-STA2'].indexOf(cible) === -1) {
    throw new Error('migration Zone de confort autorisée uniquement pour les lignes DNF/STA de TEST');
  }
  var ss = ouvrirClasseur_(cible);
  var shP = ss.getSheetByName(SHEET_PRESENCES);
  var cols = presencesCols_(shP);
  if (ss.getRangeByName('ListeZoneConfort') && cols['Zone de confort'] && cols['Zone de confort (ID)']) {
    return { cible: cible, dejaMigre: true };
  }
  if (!cols['Zone de confort'] || cols['Zone de confort (ID)']) {
    throw new Error('structure Presences inattendue pour ' + cible);
  }

  var shL = ss.getSheetByName(SHEET_LISTES);
  var titres = shL.getRange(1, 1, 1, shL.getLastColumn()).getValues()[0];
  var colZone = titres.indexOf('Zone de confort') + 1;
  if (!colZone) throw new Error('liste Zone de confort introuvable');
  var definitions = [[1, 'Confort', 'oui'], [2, 'Challenge', 'oui'], [3, 'Limite', 'oui']];
  // La liste historique occupe une seule colonne : on insère deux colonnes
  // afin de préserver les listes Nature et Statut déjà présentes à droite.
  shL.insertColumnsAfter(colZone, 2);
  shL.getRange(1, colZone).setValue('Zone de confort');
  shL.getRange(2, colZone, 1, 3).setValues([['ID', 'Libellé', 'Actif']]);
  shL.getRange(3, colZone, definitions.length, 3).setValues(definitions);
  ss.setNamedRange('ListeZoneConfort', shL.getRange(3, colZone, definitions.length, 3));

  var colLibelle = cols['Zone de confort'];
  shP.insertColumnAfter(colLibelle);
  var colId = colLibelle + 1;
  var ref = shP.getRange(4, colLibelle);
  shP.getRange(4, colId).setValue('Zone de confort (ID)')
    .setBackground(ref.getBackground()).setFontColor(ref.getFontColor()).setFontWeight(ref.getFontWeight());
  var nRows = shP.getMaxRows() - 4;
  var formules = [];
  for (var r = 5; r < 5 + nRows; r++) {
    formules.push(['=IFERROR(INDEX(ListeZoneConfort;MATCH(' + shP.getRange(r, colLibelle).getA1Notation() + ';INDEX(ListeZoneConfort;0;2);0);1);"")']);
  }
  shP.getRange(5, colId, nRows, 1).setFormulas(formules);
  shP.getRange(5, colLibelle, nRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(shL.getRange(3, colZone + 1, definitions.length, 1), true).setAllowInvalid(false).build()
  );
  SpreadsheetApp.flush();
  return { cible: cible, colonneLibelle: colLibelle, colonneId: colId, valeurs: definitions.length };
}

// Prépare exclusivement les deux lignes DNF de TEST. À lancer une fois après
// leur création ; les fonctions sont idempotentes en cas de second passage.
function preparerDnfTest_(cible) {
  return { cible: cible, zoneConfort: migrerZoneConfortLignesTest_(cible), statuts: migrerStatutsSeance(cible) };
}

function preparerDnf1Test() { return preparerDnfTest_('TEST-DNF1'); }
function preparerDnf2Test() { return preparerDnfTest_('TEST-DNF2'); }
function preparerDnfTests() {
  return [preparerDnf1Test(), preparerDnf2Test()];
}

// Prépare exclusivement les deux lignes STA de TEST. Les contrôles de nom
// empêchent d'appliquer cette préparation à un classeur DNF ou de production.
function preparerStaTest_(cible) {
  if (['TEST-STA1', 'TEST-STA2'].indexOf(cible) === -1) {
    throw new Error('préparation STA TEST autorisée uniquement pour TEST-STA1 et TEST-STA2');
  }
  var ss = ouvrirClasseur_(cible);
  if (ss.getName().indexOf('AA - Suivi STA') !== 0 || ss.getName().indexOf('TEST') !== -1) {
    throw new Error('le classeur STA de TEST attendu est introuvable : ' + ss.getName());
  }
  return { cible: cible, zoneConfort: migrerZoneConfortLignesTest_(cible), statuts: migrerStatutsSeance(cible) };
}

function preparerSta1Test() { return preparerStaTest_('TEST-STA1'); }
function preparerSta2Test() { return preparerStaTest_('TEST-STA2'); }
function preparerStaTests() {
  return [preparerSta1Test(), preparerSta2Test()];
}

// Migration des copies DNF/STA de production. La liste fermée empêche toute
// écriture sur les autres classeurs de production.
function migrerZoneConfortLignesProd_(cible) {
  if (['PROD-DNF1', 'PROD-DNF2', 'PROD-STA1', 'PROD-STA2'].indexOf(cible) === -1) {
    throw new Error('migration Zone de confort autorisée uniquement pour les lignes DNF/STA de PROD');
  }
  var ss = ouvrirClasseur_(cible);
  var shP = ss.getSheetByName(SHEET_PRESENCES);
  var cols = presencesCols_(shP);
  if (ss.getRangeByName('ListeZoneConfort') && cols['Zone de confort'] && cols['Zone de confort (ID)']) {
    return { cible: cible, dejaMigre: true };
  }
  if (!cols['Zone de confort'] || cols['Zone de confort (ID)']) {
    throw new Error('structure Presences inattendue pour ' + cible);
  }

  var shL = ss.getSheetByName(SHEET_LISTES);
  var titres = shL.getRange(1, 1, 1, shL.getLastColumn()).getValues()[0];
  var colZone = titres.indexOf('Zone de confort') + 1;
  if (!colZone) throw new Error('liste Zone de confort introuvable');
  var definitions = [[1, 'Confort', 'oui'], [2, 'Challenge', 'oui'], [3, 'Limite', 'oui']];
  // La liste historique occupe une seule colonne : on insère deux colonnes
  // afin de préserver les listes Nature et Statut déjà présentes à droite.
  shL.insertColumnsAfter(colZone, 2);
  shL.getRange(1, colZone).setValue('Zone de confort');
  shL.getRange(2, colZone, 1, 3).setValues([['ID', 'Libellé', 'Actif']]);
  shL.getRange(3, colZone, definitions.length, 3).setValues(definitions);
  ss.setNamedRange('ListeZoneConfort', shL.getRange(3, colZone, definitions.length, 3));

  var colLibelle = cols['Zone de confort'];
  shP.insertColumnAfter(colLibelle);
  var colId = colLibelle + 1;
  var ref = shP.getRange(4, colLibelle);
  shP.getRange(4, colId).setValue('Zone de confort (ID)')
    .setBackground(ref.getBackground()).setFontColor(ref.getFontColor()).setFontWeight(ref.getFontWeight());
  var nRows = shP.getMaxRows() - 4;
  var formules = [];
  for (var r = 5; r < 5 + nRows; r++) {
    formules.push(['=IFERROR(INDEX(ListeZoneConfort;MATCH(' + shP.getRange(r, colLibelle).getA1Notation() + ';INDEX(ListeZoneConfort;0;2);0);1);"")']);
  }
  shP.getRange(5, colId, nRows, 1).setFormulas(formules);
  shP.getRange(5, colLibelle, nRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(shL.getRange(3, colZone + 1, definitions.length, 1), true).setAllowInvalid(false).build()
  );
  SpreadsheetApp.flush();
  return { cible: cible, colonneLibelle: colLibelle, colonneId: colId, valeurs: definitions.length };
}


// Prépare exclusivement les deux copies DNF de PROD. Elle est séparée de la
// procédure TEST afin qu'une mauvaise cible ne puisse pas modifier un autre
// classeur. À exécuter manuellement avant toute publication des pages DNF.
function preparerDnfProd_(cible) {
  if (['PROD-DNF1', 'PROD-DNF2'].indexOf(cible) === -1) {
    throw new Error('préparation DNF PROD autorisée uniquement pour PROD-DNF1 et PROD-DNF2');
  }
  var ss = ouvrirClasseur_(cible);
  if (ss.getName().indexOf('AA - Suivi DNF') !== 0 || ss.getName().indexOf('TEST') !== -1) {
    throw new Error('nom de classeur DNF PROD inattendu : ' + ss.getName());
  }
  return {
    cible: cible,
    zoneConfort: migrerZoneConfortLignesProd_(cible),
    statuts: migrerStatutsSeance(cible)
  };
}

function preparerDnf1Prod() { return preparerDnfProd_('PROD-DNF1'); }
function preparerDnf2Prod() { return preparerDnfProd_('PROD-DNF2'); }

// Prépare exclusivement les deux copies STA de PROD. La liste fermée et le
// contrôle de nom évitent toute migration sur une ligne de TEST ou DNF.
function preparerStaProd_(cible) {
  if (['PROD-STA1', 'PROD-STA2'].indexOf(cible) === -1) {
    throw new Error('préparation STA PROD autorisée uniquement pour PROD-STA1 et PROD-STA2');
  }
  var ss = ouvrirClasseur_(cible);
  if (ss.getName().indexOf('AA - Suivi STA') !== 0 || ss.getName().indexOf('TEST') !== -1) {
    throw new Error('nom de classeur STA PROD inattendu : ' + ss.getName());
  }
  return {
    cible: cible,
    zoneConfort: migrerZoneConfortLignesProd_(cible),
    statuts: migrerStatutsSeance(cible)
  };
}

function preparerSta1Prod() { return preparerStaProd_('PROD-STA1'); }
function preparerSta2Prod() { return preparerStaProd_('PROD-STA2'); }
function preparerStaProds() {
  return [preparerSta1Prod(), preparerSta2Prod()];
}

// Migration des statuts de séance pour un classeur ciblé. Elle conserve la
// colonne I "Statut" telle qu'elle est aujourd'hui : celle-ci reste lisible
// et éditable dans Google Sheets. L'ID stable est ajouté à la fin du tableau
// afin de ne pas décaler les colonnes N à R, déjà utilisées par le
// remplacement d'encadrant et le plan de séance.
//
// Fonction de maintenance à lancer une fois dans l'éditeur Apps Script,
// jamais depuis doGet/doPost. Les contrôles sont réalisés avant l'écriture.
function migrerStatutsSeance(cible) {
  if (!cible) throw new Error('cible manquante, ex. migrerStatutsSeance("TEST-L2")');
  var ss = ouvrirClasseur_(cible);
  var shCalendrier = ss.getSheetByName(SHEET_CALENDRIER);
  var headers = shCalendrier.getRange(4, 1, 1, shCalendrier.getLastColumn()).getValues()[0];
  var colStatut = headers.indexOf('Statut') + 1;
  var colStatutId = headers.indexOf('Statut (ID)') + 1;
  var plageExistante = ss.getRangeByName('ListeStatutsSeance');

  if (plageExistante && colStatutId) {
    return { cible: cible, dejaMigre: true, colonneStatut: colStatut, colonneId: colStatutId };
  }
  if (plageExistante || colStatutId) {
    throw new Error('migration incomplète détectée : ListeStatutsSeance et Statut (ID) doivent être présents ensemble');
  }
  if (colStatut !== 9) {
    throw new Error('structure Calendrier inattendue : Statut attendu en colonne I');
  }

  // On valide les valeurs réellement présentes avant de créer quoi que ce
  // soit : aucun statut historique ne doit être transformé à l'aveugle.
  var definitions = [
    [1, 'planifiée', 'oui'],
    [2, 'tenue', 'oui'],
    [3, 'annulée', 'oui'],
    [4, 'fermée', 'oui']
  ];
  var libellesValides = {};
  definitions.forEach(function (x) { libellesValides[x[1]] = true; });
  var nRows = shCalendrier.getMaxRows() - 4;
  var statutsExistants = shCalendrier.getRange(5, colStatut, nRows, 1).getValues();
  var inconnus = {};
  statutsExistants.forEach(function (row) {
    var v = String(row[0] || '').trim();
    if (v && !libellesValides[v]) inconnus[v] = true;
  });
  if (Object.keys(inconnus).length) {
    throw new Error('statut(s) inconnu(s), aucune écriture effectuée : ' + Object.keys(inconnus).join(', '));
  }

  var shListes = ss.getSheetByName(SHEET_LISTES);
  var colDebut = shListes.getLastColumn() + 2; // une colonne vide de séparation
  var refTitre = shListes.getRange(1, 1);
  var refEntete = shListes.getRange(2, 1);
  shListes.getRange(1, colDebut).setValue('Statuts de séance');
  shListes.getRange(2, colDebut, 1, 3).setValues([['ID', 'Libellé', 'Actif']]);
  shListes.getRange(3, colDebut, definitions.length, 3).setValues(definitions);
  shListes.getRange(1, colDebut, 1, 3)
    .setBackground(refTitre.getBackground()).setFontColor(refTitre.getFontColor()).setFontWeight('bold');
  shListes.getRange(2, colDebut, 1, 3)
    .setBackground(refEntete.getBackground()).setFontColor(refEntete.getFontColor()).setFontWeight('bold');
  ss.setNamedRange('ListeStatutsSeance', shListes.getRange(3, colDebut, definitions.length, 3));

  // Ajout en fin de Calendrier : aucune colonne existante n'est déplacée.
  var derniereColonne = shCalendrier.getLastColumn();
  shCalendrier.insertColumnAfter(derniereColonne);
  var colId = derniereColonne + 1;
  shCalendrier.getRange(4, colId).setValue('Statut (ID)')
    .setBackground(shCalendrier.getRange(4, colStatut).getBackground())
    .setFontColor(shCalendrier.getRange(4, colStatut).getFontColor())
    .setFontWeight(shCalendrier.getRange(4, colStatut).getFontWeight());
  var formules = [];
  for (var r = 5; r < 5 + nRows; r++) {
    formules.push(['=IFERROR(INDEX(ListeStatutsSeance;MATCH(I' + r + ';INDEX(ListeStatutsSeance;0;2);0);1);"")']);
  }
  shCalendrier.getRange(5, colId, nRows, 1).setFormulas(formules);
  shCalendrier.getRange(5, colStatut, nRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(shListes.getRange(3, colDebut + 1, definitions.length, 1), true)
      .setAllowInvalid(false)
      .build()
  );

  SpreadsheetApp.flush();
  var ids = shCalendrier.getRange(5, colId, nRows, 1).getValues();
  var erreurs = [];
  for (var i = 0; i < statutsExistants.length; i++) {
    if (statutsExistants[i][0] && !ids[i][0]) erreurs.push(i + 5);
  }
  if (erreurs.length) {
    throw new Error('ID de statut manquant après migration, lignes : ' + erreurs.slice(0, 10).join(', '));
  }
  return {
    cible: cible,
    liste: shListes.getRange(3, colDebut, definitions.length, 3).getA1Notation(),
    colonneStatut: colStatut,
    colonneId: colId,
    lignesFormulees: nRows
  };
}

// Compatibilité avec la première procédure de test.
function migrerStatutsSeanceTestL2() {
  return migrerStatutsSeance('TEST-L2');
}

// Prépare les classeurs déjà configurés dans CLASSEURS. Cette fonction ne
// touche pas aux cibles encore placeholders (DNF/STA, TEST-L3, DEV).
function migrerStatutsSeanceLignesConfigurees() {
  var cibles = ['TEST-L2', 'PROD-L1', 'PROD-L2', 'PROD-L3', 'PROD-L4', 'PROD-LC'];
  var resultats = [];
  cibles.forEach(function (cible) {
    try {
      resultats.push(migrerStatutsSeance(cible));
    } catch (e) {
      resultats.push({ cible: cible, erreur: String(e) });
    }
  });
  Logger.log(JSON.stringify(resultats, null, 2));
  return resultats;
}

// Migration PROD validée sur les cinq copies TEST MIGRATION le 12/09/2026.
// Chaque contrôle est effectué avant la première écriture du classeur visé.
function migrerZonesConfortProd_(cible) {
  var autorisees = { 'PROD-L1': true, 'PROD-L2': true, 'PROD-L3': true, 'PROD-L4': true, 'PROD-LC': true };
  if (!autorisees[cible]) throw new Error('cible PROD non autorisée : ' + cible);
  var ss = ouvrirClasseur_(cible);
  if (ss.getName().indexOf('AA - Suivi ') !== 0 || ss.getName().indexOf('TEST') !== -1) {
    throw new Error('nom de classeur PROD inattendu : ' + ss.getName());
  }
  var sh = ss.getSheetByName(SHEET_PRESENCES);
  var col = presencesCols_(sh);
  if (col['Zone de confort'] !== 6) throw new Error(cible + ' : Zone de confort attendue en F');
  if (col['Zone de confort (ID)']) return { cible: cible, dejaMigre: true };
  if (col['Observation de l\'encadrant'] !== 7 || col['Contrôle'] !== 8) {
    throw new Error(cible + ' : structure Presences inattendue');
  }
  var nRows = sh.getMaxRows() - 4;
  var source = sh.getRange(5, 6, nRows, 1).getValues();
  var compte = { vides: 0, confort: 0, limite: 0 };
  var inconnues = {};
  var libelles = source.map(function (row) {
    if (row[0] === '' || row[0] === null) { compte.vides++; return ['']; }
    var cle = String(row[0]).trim().toLowerCase();
    if (cle === 'dans la zone') { compte.confort++; return ['Confort']; }
    if (cle === 'en limite' || cle === 'hors zone' || cle === 'hors zoone') {
      compte.limite++; return ['Limite'];
    }
    inconnues[String(row[0])] = true;
    return [row[0]];
  });
  if (Object.keys(inconnues).length) {
    throw new Error(cible + ' : valeurs inconnues, aucune écriture effectuée : ' + Object.keys(inconnues).join(', '));
  }
  var listes = ss.getSheetByName(SHEET_LISTES);
  listes.getRange('S2:U5').setValues([
    ['Zone de confort — ID', 'Zone de confort — libellé', 'Actif'],
    [1, 'Confort', 'oui'], [2, 'Challenge', 'oui'], [3, 'Limite', 'oui']
  ]);
  ss.setNamedRange('ListeZoneConfort', listes.getRange('S3:U5'));
  sh.insertColumnAfter(6);
  sh.getRange(4, 7).setValue('Zone de confort (ID)');
  sh.getRange(5, 6, nRows, 1).setValues(libelles);
  var formules = [];
  for (var r = 5; r < 5 + nRows; r++) {
    formules.push(['=IFERROR(INDEX(ListeZoneConfort;MATCH(F' + r + ';INDEX(ListeZoneConfort;0;2);0);1);"")']);
  }
  sh.getRange(5, 7, nRows, 1).setFormulas(formules);
  SpreadsheetApp.flush();
  var ids = sh.getRange(5, 7, nRows, 1).getValues();
  var apres = { vides: 0, id1: 0, id2: 0, id3: 0, autres: 0 };
  ids.forEach(function (row) {
    var id = row[0];
    if (id === '' || id === null) apres.vides++;
    else if (id === 1) apres.id1++;
    else if (id === 2) apres.id2++;
    else if (id === 3) apres.id3++;
    else apres.autres++;
  });
  var colApres = presencesCols_(sh);
  if (colApres['Observation de l\'encadrant'] !== 8 || colApres['Contrôle'] !== 9 ||
      apres.id1 !== compte.confort || apres.id3 !== compte.limite || apres.id2 !== 0 || apres.autres !== 0) {
    throw new Error(cible + ' : contrôle final incohérent ' + JSON.stringify(apres));
  }
  return { cible: cible, avant: compte, apres: apres };
}

function migrerZonesConfortProd20260912() {
  return ['PROD-L1', 'PROD-L2', 'PROD-L3', 'PROD-L4', 'PROD-LC'].map(migrerZonesConfortProd_);
}
