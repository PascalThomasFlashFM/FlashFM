import * as fs from "fs";
import { google, sheets_v4 } from "googleapis";
import { AppConfig } from "../config";

export async function getSheetsClient(config: AppConfig): Promise<sheets_v4.Sheets> {
  const credentials = JSON.parse(
    fs.readFileSync(config.google.serviceAccountJsonPath, "utf-8")
  );

  const auth = new google.auth.JWT({
    email: credentials.client_email,
    key: credentials.private_key,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  await auth.authorize();

  return google.sheets({ version: "v4", auth });
}

export async function getSheetNumericId(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  tabName: string
): Promise<number> {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = meta.data.sheets?.find((s) => s.properties?.title === tabName);
  if (!sheet || sheet.properties?.sheetId == null) {
    throw new Error(
      `Onglet "${tabName}" introuvable dans le Google Sheet. Verifie GOOGLE_SHEET_TAB.`
    );
  }
  return sheet.properties.sheetId;
}
