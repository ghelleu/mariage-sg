/**
 * Web app — Réception RSVP mariage S&G
 * v6 — Standalone script lié au GSheet via openById. 14 colonnes, 1 ligne par invité.
 *
 * Déployé via l'API Apps Script (Aug 2026) :
 *   - Projet standalone : scriptId 1axEmdTtIPbeBPmZSRb3DKJsZKv82J46B9tFFbTAMvi4Vz7-QCLSQaIeL
 *   - Accès : ANYONE_ANONYMOUS (webhook public appelé depuis le site, sans login Google)
 *   - Execute as : USER_DEPLOYING (Guillaume)
 *   - URL : https://script.google.com/macros/s/AKfycbzU-i6rvPiEYDccK9fMwoVLHYRn-U0NSLF8QSziKVPJ1D66aYp20ldY88SuF0xStFJQVw/exec
 *
 * PIEGE CLE : un déploiement en accès "ANYONE" (pas _ANONYMOUS) renvoie 401 sur
 * un POST depuis un navigateur non-loggué. Il FAUT "ANYONE_ANONYMOUS" pour un
 * webhook public. C'est la cause du bug initial (données en vrac).
 */
const SHEET_ID = '1OA4_GAmPu3crn5g6ovQWzvrs3yVnYaJseZjgisX1_8s';
const SHEET_NAME = 'Réponses';

// En-têtes (dans l'ordre exact du Sheet) — v6 : 14 colonnes
const HEADERS = [
  'Date',                    // A
  'Email',                   // B
  'Responsable reservation', // C
  'Prenom & Nom',            // D
  'Presence',                // E
  'Ven 4',                   // F
  'Sam 5',                   // G
  'Dim 6',                   // H
  'Lun 7',                   // I
  'Allergies',               // J
  'Maquillage',              // K
  'Coiffure',                // L
  'Chanson',                 // M
  'Message'                  // N
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
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" introuvable');

    // Anti-doublon simple : si email + nom responsable identiques dans les 60 dernières secondes
    const guests = data.guests || [];
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

    // S'assurer que les en-têtes existent (v6 : 14 colonnes)
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

    // Jours éclatés en 4 colonnes OUI/NON
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

    // v6 : une ligne par invité (14 colonnes)
    const rows = guests.map((g, i) => [
      now,                          // A: Date (identique pour tout le groupe)
      data.email || '',             // B: Email (identique pour tout le groupe)
      responsable,                  // C: Responsable réservation (identique pour tout le groupe)
      g.nom || '',                  // D: Prénom & Nom (cet invité)
      presence,                     // E: Présence
      joursCols.ven,                // F: Ven 4
      joursCols.sam,                // G: Sam 5
      joursCols.dim,                // H: Dim 6
      joursCols.lun,                // I: Lun 7
      g.allergies || '',            // J: Allergies / précisions repas
      g.maquille ? 'OUI' : 'NON',   // K: Maquillage
      g.coiffe ? 'OUI' : 'NON',     // L: Coiffure
      i === 0 ? (data.chanson || '') : '',  // M: Chanson (1ère ligne seulement)
      i === 0 ? (data.message || '') : ''   // N: Message (1ère ligne seulement)
    ]);

    // Append en bloc (plus rapide que N appendRow)
    if (rows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);

      // Dropdowns OUI/NON sur les colonnes E, F, G, H, I (Présence + 4 jours) + K, L (Maquillage, Coiffure)
      const validationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['OUI', 'NON'], true)
        .setAllowInvalid(false)
        .build();
      
      // Colonnes E-I (Présence + 4 jours) : range contiguë
      sheet.getRange(startRow, 5, rows.length, 5).setDataValidation(validationRule);
      
      // Colonnes K-L (Maquillage, Coiffure) : range contiguë
      sheet.getRange(startRow, 11, rows.length, 2).setDataValidation(validationRule);
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
