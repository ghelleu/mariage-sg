/**
 * Web app — Réception RSVP mariage S&G
 * v3 — Ajout du jour "lundi 7 déc" pour le brunch prolongé
 *
 * À déployer : Extensions > Apps Script > Déployer > Gérer les déploiements > ✏️ > "Nouvelle version" > Déployer
 * Exécuter en tant que : Moi
 * Accès : Tout le monde
 */
const SHEET_NAME = 'Réponses';
const RSVP_WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbyPHna7JcCjBBOLhjDE2tyT76ta2dzoVnLAxejkfLLxCGrje4WnC41sEtuEBEGwMtk/exec';

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'RSVP endpoint up' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const now = new Date();
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" introuvable');

    // Anti-doublon simple : si email + nombre invités identiques dans les 60 dernières secondes
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const startRow = Math.max(2, lastRow - 20);
      const numRows = Math.min(20, lastRow - 1);
      const lastRows = sheet.getRange(startRow, 1, numRows, 11).getValues();
      for (const r of lastRows) {
        const ts = r[0];
        if (ts && (now - ts) < 60000 && r[1] === data.email && r[5] === (data.guests || []).length) {
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, dedup: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    // v3 : mapping jours (inclut 'lun' = Lundi 7 déc)
    const jours = (data.jours || data.days || {});
    const joursMap = { ven: 'Ven 4', sam: 'Sam 5', dim: 'Dim 6', lun: 'Lun 7' };
    const ordre = ['ven', 'sam', 'dim', 'lun'];
    const joursStr = ordre.filter(k => jours[k]).map(k => joursMap[k]).join(', ');

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
      e.parameter && e.parameter.userAgent ? e.parameter.userAgent : '',
      'mariage-au-chateau.xyz'
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
