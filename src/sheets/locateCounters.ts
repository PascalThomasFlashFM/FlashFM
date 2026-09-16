import { sheets_v4 } from "googleapis";

// La ligne 2 contient les en-tetes, les donnees commencent en ligne 3.
export const DATA_START_ROW = 3;

export interface InsertionPlan {
  // Ligne (1-indexee) ou ecrire la prochaine inscription.
  targetRow: number;
  // true si targetRow est une ligne vide deja presente (pas besoin d'inserer),
  // false s'il faut inserer une nouvelle ligne au-dessus de firstCounterRow.
  useExistingBlankRow: boolean;
  // Toutes les lignes (1-indexees) contenant une formule NBVAL en colonne B.
  counterRows: number[];
  firstCounterRow: number;
  lastDataRowBeforeWrite: number;
}

async function getColumnBFormulas(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  tabName: string
): Promise<string[]> {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${tabName}'!B1:B2000`,
    valueRenderOption: "FORMULA",
  });
  const values = res.data.values ?? [];
  return values.map((row) => (row && row[0] != null ? String(row[0]) : ""));
}

/**
 * Localise dynamiquement les cellules contenant une formule NBVAL (les compteurs
 * "Pour la saison :" / "Inscrits :") en colonne B. Leur position bouge dans le temps,
 * donc on ne les code jamais en dur.
 */
export async function getInsertionPlan(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  tabName: string
): Promise<InsertionPlan> {
  const columnB = await getColumnBFormulas(sheets, spreadsheetId, tabName);

  const counterRows: number[] = [];
  for (let i = 0; i < columnB.length; i++) {
    if (columnB[i].toUpperCase().includes("NBVAL(")) {
      counterRows.push(i + 1); // 1-indexed
    }
  }

  if (counterRows.length === 0) {
    throw new Error(
      `Aucune formule NBVAL trouvee en colonne B de l'onglet "${tabName}". ` +
        `Impossible de localiser la zone des compteurs d'inscrits en toute securite : abandon.`
    );
  }

  const firstCounterRow = Math.min(...counterRows);

  let lastDataRow = DATA_START_ROW - 1; // = 2, la ligne d'en-tete
  for (let row = DATA_START_ROW; row < firstCounterRow; row++) {
    const value = columnB[row - 1] ?? "";
    if (value.trim() !== "") {
      lastDataRow = row;
    }
  }

  const gapRow = lastDataRow + 1;
  if (gapRow < firstCounterRow) {
    return {
      targetRow: gapRow,
      useExistingBlankRow: true,
      counterRows,
      firstCounterRow,
      lastDataRowBeforeWrite: lastDataRow,
    };
  }

  return {
    targetRow: firstCounterRow,
    useExistingBlankRow: false,
    counterRows,
    firstCounterRow,
    lastDataRowBeforeWrite: lastDataRow,
  };
}

const RANGE_PATTERN = /NBVAL\(\s*([A-Za-z]+)(\d+)\s*:\s*([A-Za-z]+)(\d+)\s*\)/i;

/**
 * Verifie que chaque formule NBVAL couvre bien jusqu'a newLastDataRow, sans jamais
 * inclure la ligne du compteur lui-meme (dependance circulaire). Corrige si besoin.
 */
export async function ensureCounterRangesCover(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  tabName: string,
  counterRows: number[],
  newLastDataRow: number
): Promise<void> {
  const columnB = await getColumnBFormulas(sheets, spreadsheetId, tabName);
  const minCounterRow = Math.min(...counterRows);

  if (newLastDataRow >= minCounterRow) {
    throw new Error(
      `Incoherence : la derniere ligne de donnees (${newLastDataRow}) atteindrait la ligne ` +
        `des compteurs (${minCounterRow}). Ecriture annulee pour eviter une reference circulaire.`
    );
  }

  const updates: sheets_v4.Schema$ValueRange[] = [];

  for (const row of counterRows) {
    const formula = columnB[row - 1] ?? "";
    const match = formula.match(RANGE_PATTERN);
    if (!match) continue; // formule NBVAL dans une forme inattendue : on ne touche pas

    const [, startCol, startRowStr, endCol, endRowStr] = match;
    const endRow = parseInt(endRowStr, 10);
    if (endRow >= newLastDataRow) continue; // deja a jour

    const newFormula = formula.replace(
      RANGE_PATTERN,
      `NBVAL(${startCol}${startRowStr}:${endCol}${newLastDataRow})`
    );

    updates.push({
      range: `'${tabName}'!B${row}`,
      values: [[newFormula]],
    });
  }

  if (updates.length === 0) return;

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: updates,
    },
  });
}
