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
  'DEV-L2':   '1ANbbV-lc9GZeVHQH8X4mFOaMKo4s8wB5TrnmX96WEko',
  'DEV-L3':   '1gVjJxXIXzvfJUObElSu8D3dnFRHTNcCWemQH-JwDsRY',
  'TEST-L2':  'REMPLACER_PAR_ID_CLASSEUR_TEST_L2',
  'TEST-L3':  'REMPLACER_PAR_ID_CLASSEUR_TEST_L3',
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
  var resultat = { ok: true, roster: lirePersonnes_(ss), seances: lireCalendrier_(ss) };
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
// Une même personne peut apparaître sur plusieurs lignes du même classeur
// (même id en colonne B) quand elle a plusieurs rôles dans la ligne, par
// exemple élève ET encadrant (cas Sandry Wallon en L3) : on ne garde
// qu'une seule entrée par id, en préférant la ligne "élève" quand elle
// existe (demande du 06/09/2026 — l'encadrant n'a pas à apparaître comme
// participant, il est déjà identifié comme tel dans le calendrier).
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
      parId[id] = candidat;
      ordre.push(id);
    } else if (existant.role !== 'élève' && candidat.role === 'élève') {
      parId[id] = candidat;
    }
  }
  return ordre.map(function (id) {
    var p = parId[id];
    return { id: p.id, nom: p.nom, prenom: p.prenom, membre: p.membre };
  });
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
  var values = sh.getDataRange().getValues();
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
      statut: row[8] || 'planifiée'
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
// manuelle, non utilisee par l'app), C apneiste id, D nom, E prenom,
// F qualite, G observation, H controle (formule, non touchee par l'app).
function calculerFiche_(ss, seanceId) {
  var values = ss.getSheetByName(SHEET_PRESENCES).getDataRange().getValues();
  var presences = [];
  for (var r = 4; r < values.length; r++) {
    var row = values[r];
    if (row[0] === seanceId) {
      presences.push({
        apneiste_id: row[2], nom: row[3], prenom: row[4],
        qualite: row[5] || '', observation: row[6] || ''
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
function savePresences_(body) {
  var email = verifierJeton_(body.idToken);
  var seanceId = body.seance_id;
  if (!seanceId) throw new Error('seance_id manquant');
  var ss = ouvrirClasseur_(body.cible);

  var roster = {};
  lirePersonnes_(ss).forEach(function (p) { roster[p.id] = p; });

  var presences = (body.presences || []).filter(function (p) { return p.present; });
  presences.forEach(function (p) {
    // V-01 : la clé étrangère doit exister dans le référentiel — jamais de
    // confiance aveugle dans ce que le client envoie.
    if (!roster[p.apneiste_id]) {
      throw new Error('apnéiste inconnu du référentiel : ' + p.apneiste_id);
    }
  });

  var sh = ss.getSheetByName(SHEET_PRESENCES);
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
    sh.getRange(row, 6, 1, 2).setValues([[p.qualite || '', p.observation || '']]);
  });
  // lignes qui appartenaient à la séance mais ne sont plus nécessaires —
  // vidées (jamais les colonnes C/D/E/H, formules).
  for (var j = presences.length; j < deSeance.length; j++) {
    var r2 = 5 + deSeance[j];
    sh.getRange(r2, 1, 1, 2).clearContent();
    sh.getRange(r2, 6, 1, 2).clearContent();
  }

  // rafraîchit immédiatement le cache "fiche" de cette séance : la
  // prochaine lecture (même appareil ou un autre) voit tout de suite le
  // résultat de cet enregistrement, sans attendre l'expiration du TTL
  // (demande du 06/09/2026).
  var cache = CacheService.getScriptCache();
  cache.put('v2|' + body.cible + '|fiche|' + seanceId,
            JSON.stringify(calculerFiche_(ss, seanceId)), CACHE_TTL_FICHE);

  return { ok: true, saved: presences.length, par: email };
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