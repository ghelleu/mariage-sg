/**
 * Web app — Réception RSVP mariage S&G
 * v5 — Suppression colonne Menu (formulaire n'a plus de choix viande/poisson/enfant)
 *
 * À déployer : Extensions > Apps Script > Déployer > Gérer les déploiements > ✏️ > "Nouvelle version" > Déployer
 * Exécuter en tant que : Moi
 * Accès : Tout le monde
 */
const SHEET_NAME = 'Réponses';
const SHEET_ID = 0; // ID de la feuille "Réponses" (vérifier avec SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getSheetId())
const RSVP_WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbyPHna7JcCjBBOLhjDE2tyT76ta2dzoVnLAxejkfLLxCGrje4WnC41sEtuEBEGwMtk/exec';

// En-têtes (dans l'ordre exact du Sheet) — v5 : 12 colonnes (Menu supprimé)
const HEADERS = [
  'Date',                  // A
  'Email',                 // B
  'Responsable réservation', // C
  'Prénom & Nom',          // D
  'Présence',              // E
  'Ven 4',                 // F
  'Sam 5',                 // G
  'Dim 6',                 // H
  'Lun 7',                 // I
  'Allergies',             // J
  'Chanson',               // K
  'Message'                // L
];

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'RSVP endpoint up' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const now = new Date();
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" introuvable');

    // Anti-doublon simple : si email + N invités identiques dans les 60 dernières secondes
    const guests = data.guests || [];
    const nbInvites = guests.length;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const startRow = Math.max(2, lastRow - 20);
      const numRows = Math.min(20, lastRow - 1);
      const lastRows = sheet.getRange(startRow, 1, numRows, HEADERS.length).getValues();
      for (const r of lastRows) {
        const ts = r[0];
        if (ts && (now - ts) < 60000 && r[1] === data.email && r[3] === (guests[0] && guests[0].nom)) {
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, dedup: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    // S'assurer que les en-têtes existent (v5 : 12 colonnes)
    const existingHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    const headersUpToDate = HEADERS.every((h, i) => existingHeaders[i] === h);
    if (!headersUpToDate) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      // Mise en forme des en-têtes
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f2ecde');
      sheet.setFrozenRows(1);
    }

    // v4 : jours éclatés en 4 colonnes
    const jours = (data.jours || data.days || {});
    const joursCols = {
      ven: jours.ven ? 'OUI' : 'NON',
      sam: jours.sam ? 'OUI' : 'NON',
      dim: jours.dim ? 'OUI' : 'NON',
      lun: jours.lun ? 'OUI' : 'NON'
    };
    const presence = data.attending === true ? 'OUI' : (data.attending === false ? 'NON' : '?');

    // Responsable = nom du guest principal (1er guest)
    const responsable = (guests[0] && guests[0].nom) || '';

    // v5 : une ligne par invité (Menu supprimé)
    const rows = guests.map((g, i) => [
      now,                          // A: Date (identique pour tout le groupe)
      data.email || '',             // B: Email (identique pour tout le groupe)
      responsable,                  // C: Responsable réservation (identique pour tout le groupe)
      g.nom || '',                  // D: Prénom & Nom (cet invité)
      presence,                     // E: Présence (identique pour tout le groupe)
      joursCols.ven,                // F
      joursCols.sam,                // G
      joursCols.dim,                // H
      joursCols.lun,                // I
      g.allergies || '',            // J: Allergies / précisions repas
      i === 0 ? (data.chanson || '') : '',  // K: Chanson (1ère ligne seulement)
      i === 0 ? (data.message || '') : ''   // L: Message (1ère ligne seulement)
    ]);

    // Append en bloc (plus rapide que N appendRow)
    if (rows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);

      // Dropdowns OUI/NON sur les colonnes E, F, G, H, I (Présence + 4 jours)
      const validationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['OUI', 'NON'], true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(startRow, 5, rows.length, 5).setDataValidation(validationRule);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        rowsAdded: rows.length,
        firstRow: sheet.getLastRow() - rows.length + 1,
        lastRow: sheet.getLastRow()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
