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

var CLASSEURS = {
  'DEV-L2':   '1wlmAwZwR0z4KNSnYiadXiLqQUv05IwRzIi8ERjmf0NI',
  'DEV-L3':   '1Q1ZRnsUfbG0aW2KiSY7B3zI11H2TMQA9cbRm-S90acQ',
  'TEST-L2':  'REMPLACER_PAR_ID_CLASSEUR_TEST_L2',
  'TEST-L3':  'REMPLACER_PAR_ID_CLASSEUR_TEST_L3',
  'PROD-L1':  'REMPLACER_PAR_ID_CLASSEUR_PROD_L1',
  'PROD-L2':  'REMPLACER_PAR_ID_CLASSEUR_PROD_L2',
  'PROD-L3':  'REMPLACER_PAR_ID_CLASSEUR_PROD_L3',
  'PROD-L4':  'REMPLACER_PAR_ID_CLASSEUR_PROD_L4',
  'PROD-LC':  'REMPLACER_PAR_ID_CLASSEUR_PROD_LC',
  'PROD-DNF': 'REMPLACER_PAR_ID_CLASSEUR_PROD_DNF',
  'PROD-STA': 'REMPLACER_PAR_ID_CLASSEUR_PROD_STA'
};

var SHEET_CALENDRIER = 'Calendrier';
var SHEET_PRESENCES = 'Presences';
var SHEET_PERSONNES = 'Personnes';

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
  try {
    var ss = ouvrirClasseur_(e.parameter.cible);
    if (action === 'data') return jsonOut_(getData_(ss));
    if (action === 'fiche') return jsonOut_(lirePresences_(ss, e.parameter.seance));
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
    var ss = ouvrirClasseur_(body.cible);
    if (body.action === 'saveSeance') return jsonOut_(savePresences_(ss, body));
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
function getData_(ss) {
  return { ok: true, roster: lirePersonnes_(ss), seances: lireCalendrier_(ss) };
}

// Colonnes de l'onglet Personnes (1-based, cf. build_suivi.py) :
// A code, B id, C nom, D prenom, E role declare, F groupe(s), G objectif,
// H membre de la ligne ("oui"/"")...
// Renvoie TOUT le monde (pas seulement les membres de la ligne), avec
// l'indicateur "membre" : permet au front de proposer un "invité"
// (personne connue du classeur mais pas de cette ligne) et de l'afficher
// differemment (demande du 07/09/2026).
function lirePersonnes_(ss) {
  var sh = ss.getSheetByName(SHEET_PERSONNES);
  var values = sh.getDataRange().getValues();
  var out = [];
  for (var r = 4; r < values.length; r++) {          // ligne 5 = 1ere donnee
    var row = values[r];
    if (row[1]) {
      out.push({ id: row[1], nom: row[2], prenom: row[3], membre: row[7] === 'oui' });
    }
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

// Colonnes de l'onglet Calendrier : A id seance, B date, C jour,
// D creneau (libelle), ... I statut.
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
    out.push({
      id: row[0],
      date: formatDate_(row[1]),
      date_iso: estDate ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '',
      jour: row[2],
      creneau: row[3],
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

// Colonnes de l'onglet Presences : A id seance, B selection (saisie
// manuelle, non utilisee par l'app), C apneiste id, D nom, E prenom,
// F qualite, G observation, H controle (formule, non touchee par l'app).
function lirePresences_(ss, seanceId) {
  if (!seanceId) throw new Error('seance manquante');
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
function savePresences_(ss, body) {
  var email = verifierJeton_(body.idToken);
  var seanceId = body.seance_id;
  if (!seanceId) throw new Error('seance_id manquant');

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
