import { sheets_v4 } from "googleapis";
import { NewRegistration } from "../bocir/fetchResponses";
import { formatPhone, isoDateToDDMMYY, msToDDMMYY } from "../utils/format";
import { ensureCounterRangesCover, getInsertionPlan } from "./locateCounters";

export interface SheetWriteResult {
  row: number;
}

/**
 * Ecrit une inscription dans l'onglet de suivi :
 * - trouve une ligne (vide existante, ou insere une ligne au-dessus des compteurs)
 * - remplit uniquement les colonnes A-D, F-H (jamais E, I-N qui sont geres manuellement
 *   ou par formule)
 * - met a jour les formules NBVAL des compteurs pour couvrir la nouvelle derniere ligne
 *
 * Ne modifie et ne supprime jamais une ligne existante : n'ajoute que des lignes neuves.
 */
export async function appendRegistrationToSheet(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  sheetNumericId: number,
  tabName: string,
  registration: NewRegistration
): Promise<SheetWriteResult> {
  const plan = await getInsertionPlan(sheets, spreadsheetId, tabName);

  if (!plan.useExistingBlankRow) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            insertDimension: {
              range: {
                sheetId: sheetNumericId,
                dimension: "ROWS",
                startIndex: plan.targetRow - 1, // 0-indexed
                endIndex: plan.targetRow,
              },
              inheritFromBefore: true,
            },
          },
        ],
      },
    });
  }

  const row = plan.targetRow;

  const values = [
    msToDDMMYY(registration.submissionDate), // A
    registration.prenom, // B
    registration.nom, // C
    registration.dateNaissance ? isoDateToDDMMYY(registration.dateNaissance) : "", // D
    `=ARRONDI.INF((AUJOURDHUI()-D${row})/365)`, // E (formule, jamais une valeur en dur)
    registration.email, // F
    registration.telephone ? `'${formatPhone(registration.telephone)}` : "", // G
    registration.question ?? "", // H
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${tabName}'!A${row}:H${row}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [values] },
  });

  // Apres insertion, les compteurs peuvent avoir shifte d'une ligne (insertDimension
  // deplace les references automatiquement), on relit leur position reelle pour etendre
  // leur plage jusqu'a la nouvelle derniere ligne de donnees.
  const refreshedPlan = await getInsertionPlan(sheets, spreadsheetId, tabName);
  await ensureCounterRangesCover(
    sheets,
    spreadsheetId,
    tabName,
    refreshedPlan.counterRows,
    row
  );

  return { row };
}
