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
  var resultat = {
    ok: true, roster: lirePersonnes_(ss), seances: lireCalendrier_(ss),
    listesZoneConfort: lireListeZoneConfort_(ss)
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

// Listes de valeurs sous forme ID + libellé (proposition validée par Fred le
// 09/09/2026, cf. claude/proposition-listes-de-valeurs.md) — Zone de confort
// est le premier cas migré, uniquement sur TEST-L2 pour l'instant. Table de
// référence dans l'onglet "Listes", colonnes K (id), L (libellé), M (actif
// "oui"/"non") ; l'ID est ce qui est stocké dans Presences colonne F, jamais
// le libellé. Une valeur inactive reste lisible (une ancienne présence peut
// encore y pointer) mais n'est plus proposée dans le menu déroulant de
// saisie (filtrage fait côté front, cf. _test/L2/index.html).
function lireListeZoneConfort_(ss) {
  var sh = ss.getSheetByName(SHEET_LISTES);
  var nRows = sh.getLastRow() - 1;   // ligne 1 = en-tetes
  if (nRows < 1) return [];
  var values = sh.getRange(2, 11, nRows, 3).getValues();  // K:M
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
      // colonne R : Plan de séance (texte ou lien), saisi depuis l'appli
      // (accordéon repliable — demande du 07/09/2026).
      plan: row[17] || '',
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

function savePresences_(body) {
  var email = verifierJeton_(body.idToken);
  var seanceId = body.seance_id;
  if (!seanceId) throw new Error('seance_id manquant');
  var ss = ouvrirClasseur_(body.cible);

  var roster = {};
  lirePersonnes_(ss).forEach(function (p) { roster[p.id] = p; });

  var zonesValides = {};
  lireListeZoneConfort_(ss).forEach(function (z) { zonesValides[z.id] = true; });

  var presences = (body.presences || []).filter(function (p) { return p.present; });
  presences.forEach(function (p) {
    // V-01 : la clé étrangère doit exister dans le référentiel — jamais de
    // confiance aveugle dans ce que le client envoie.
    if (!roster[p.apneiste_id]) {
      throw new Error('apnéiste inconnu du référentiel : ' + p.apneiste_id);
    }
    // V-02 : idem pour la zone de confort — désormais un ID (cf.
    // lireListeZoneConfort_), pas un libellé libre (09/09/2026, TEST-L2).
    if (p.qualite !== '' && p.qualite !== undefined && p.qualite !== null &&
        !zonesValides[p.qualite]) {
      throw new Error('zone de confort inconnue du référentiel : ' + p.qualite);
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

  ecrirePlanSeance_(ss, seanceId, body.plan_seance);

  // rafraîchit immédiatement le cache "fiche" de cette séance : la
  // prochaine lecture (même appareil ou un autre) voit tout de suite le
  // résultat de cet enregistrement, sans attendre l'expiration du TTL
  // (demande du 06/09/2026).
  var cache = CacheService.getScriptCache();
  cache.put('v2|' + body.cible + '|fiche|' + seanceId,
            JSON.stringify(calculerFiche_(ss, seanceId)), CACHE_TTL_FICHE);
  // le plan de séance fait partie de la réponse "data" (lireCalendrier_) :
  // on invalide ce cache aussi, sinon le prochain chargement de la page
  // renverrait encore l'ancien plan pendant CACHE_TTL_DONNEES.
  cache.remove('v2|' + body.cible + '|data');

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

// Lignes à synchroniser : seules celles réellement configurées dans
// CLASSEURS (PROD-*) ci-dessus. DNF et STA ne le sont pas encore.
var LIGNES_SYNC = ['L1', 'L2', 'L3', 'L4', 'LC'];

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
      role: row[6],                                // colonne G : Rôle
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
function synchroniserPersonnesLigne_(par, ligne) {
  var ss = ouvrirClasseur_('PROD-' + ligne);
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

  // 2. Inscriptions actives dans CETTE ligne : une entrée par (personne, rôle)
  //    — une personne à la fois élève et encadrante dans la même ligne y
  //    apparaît deux fois, jamais un rôle agrégé (même règle que
  //    sync_referentiel.py, décision du 06/09/2026).
  var parCle = {};      // clé "id|role" -> inscription
  par.inscriptions.forEach(function (x) {
    if (x.groupe !== ligne || !par.estActive(x) || !par.personnes[x.id]) return;
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
  var membres = connues.concat(nouvelles);
  var membresPids = {};
  membres.forEach(function (c) { membresPids[parCle[c].id] = true; });

  // 3. Le reste du club, sans rôle dans cette ligne : présent quand même
  //    (un encadrant n'est pas nécessairement inscrit dans la ligne qu'il
  //    encadre — l'onglet Personnes porte tout le club).
  var autres = par.ordreClub.filter(function (pid) { return !membresPids[pid]; });

  var ordre = membres.map(function (c) { return { pid: parCle[c].id, x: parCle[c] }; })
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

  // 5. Effacer puis réécrire les colonnes C à G (jamais A ni H) sur toute la
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
        valeurs.push([p.nom, p.prenom, roleVal, groupesActifs_(par, o.pid), objectifVal]);
      } else {
        valeurs.push(['', '', '', '', '']);
      }
    }
    sh.getRange(PERS_FIRST_L, 3, nLignes, 5).setValues(valeurs);
  }

  return { ligne: ligne, personnes: ordre.length, membres: membres.length,
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
