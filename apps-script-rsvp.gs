/**
 * Web app — Réception RSVP mariage S&G
 * À déployer : Extensions > Apps Script > Déployer > Nouvelle déploiement > Web app
 * Exécuter en tant que : Moi
 * Accès : Tout le monde
 */
const SHEET_NAME = 'Réponses';

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'RSVP endpoint up' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" introuvable');

    // Anti-doublon simple : si email + présence + nombre invités identiques dans les 60 dernières secondes
    const lastRows = sheet.getRange(Math.max(2, sheet.getLastRow() - 20), 1, Math.min(20, sheet.getLastRow() - 1), 11).getValues();
    const now = new Date();
    for (const r of lastRows) {
      const ts = r[0];
      if (ts && (now - ts) < 60000 && r[1] === data.email && r[5] === data.nb_invites) {
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, dedup: true }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    const jours = (data.jours || data.days || {});
    const joursStr = ['ven', 'sam', 'dim'].filter(k => jours[k]).map(k => ({ven: 'Ven 4', sam: 'Sam 5', dim: 'Dim 6'}[k])).join(', ');

    const row = [
      now,
      data.email || '',
      data.attending === true ? 'OUI' : (data.attending === false ? 'NON' : '?'),
      joursStr,
      (data.guests && data.guests[0] && data.guests[0].nom) || '',
      (data.guests || []).length,
      JSON.stringify(data.guests || []),
      data.chanson || '',
      data.message || '',
      (e && e.parameter && e.parameter.userAgent) || '',
      'ghelleu.github.io/mariage-sg'
    ];
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, row: sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
