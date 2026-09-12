import { fetchFormFields } from "../bocir/fetchFormFields";
import { fetchNewResponses } from "../bocir/fetchResponses";
import { AppConfig } from "../config";
import * as db from "../db";
import { logger } from "../logger";
import { appendRegistrationToSheet } from "../sheets/appendRegistrations";
import { getSheetNumericId, getSheetsClient } from "../sheets/client";

export interface SyncSummary {
  status: "success" | "partial" | "error" | "no_new";
  nbProcessed: number;
  nbErrors: number;
  message: string;
}

export async function runSync(config: AppConfig): Promise<SyncSummary> {
  const runId = db.startSyncRun();
  logger.info("Debut de synchro BOCIR -> Sheet");

  try {
    const fieldMap = await fetchFormFields(config);
    const lastProcessed = db.getLastSubmissionDate();
    const newRegistrations = await fetchNewResponses(config, fieldMap, lastProcessed);

    if (newRegistrations.length === 0) {
      logger.info("Aucune nouvelle inscription.");
      db.finishSyncRun(runId, "success", 0, 0, null);
      return { status: "no_new", nbProcessed: 0, nbErrors: 0, message: "Aucune nouvelle inscription." };
    }

    logger.info(`${newRegistrations.length} nouvelle(s) inscription(s) a traiter.`);

    const sheets = await getSheetsClient(config);
    const sheetNumericId = await getSheetNumericId(sheets, config.google.sheetId, config.google.sheetTab);

    let processed = 0;

    for (const registration of newRegistrations) {
      if (db.registrationExists(registration.bocirId)) {
        // Deja traite (reprise apres un run partiel) : on avance juste le seuil.
        db.setLastSubmissionDate(registration.submissionDate);
        continue;
      }

      try {
        const result = await appendRegistrationToSheet(
          sheets,
          config.google.sheetId,
          sheetNumericId,
          config.google.sheetTab,
          registration
        );

        db.insertRegistration({
          bocir_id: registration.bocirId,
          submission_date: registration.submissionDate,
          prenom: registration.prenom,
          nom: registration.nom,
          email: registration.email,
          telephone: registration.telephone || null,
          date_naissance: registration.dateNaissance || null,
          question: registration.question || null,
          sheet_row: result.row,
        });

        db.setLastSubmissionDate(registration.submissionDate);
        processed += 1;
        logger.info(
          `Inscription ${registration.bocirId} (${registration.prenom} ${registration.nom}) ecrite ligne ${result.row}.`
        );
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        logger.error(
          `Echec ecriture Sheet pour l'inscription ${registration.bocirId} : ${message}. ` +
            `Le seuil de synchro n'avance pas au-dela de cette inscription, elle sera reprise au prochain passage.`
        );
        db.finishSyncRun(runId, processed > 0 ? "partial" : "error", processed, 1, message);
        return {
          status: processed > 0 ? "partial" : "error",
          nbProcessed: processed,
          nbErrors: 1,
          message,
        };
      }
    }

    db.finishSyncRun(runId, "success", processed, 0, null);
    return {
      status: "success",
      nbProcessed: processed,
      nbErrors: 0,
      message: `${processed} inscription(s) synchronisee(s).`,
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    logger.error(`Echec de la synchro : ${message}`);
    db.finishSyncRun(runId, "error", 0, 1, message);
    return { status: "error", nbProcessed: 0, nbErrors: 1, message };
  }
}
