import { AppConfig } from "./config";
import * as db from "./db";
import { logger } from "./logger";
import { buildConfirmationEmail } from "./mail/template";
import { sendConfirmationMail } from "./mail/smtp";

export function excludeRegistration(bocirId: string, reason: string | null): db.Registration {
  const reg = db.getRegistration(bocirId);
  if (!reg) throw new Error(`Inscription ${bocirId} introuvable.`);
  if (reg.mail_status !== "en_attente") {
    throw new Error(`Inscription ${bocirId} n'est pas en attente (statut actuel: ${reg.mail_status}).`);
  }
  db.markExcluded(bocirId, reason);
  logger.info(`Inscription ${bocirId} exclue${reason ? ` (raison: ${reason})` : ""}.`);
  return db.getRegistration(bocirId)!;
}

export async function sendRegistrationMail(
  config: AppConfig,
  bocirId: string
): Promise<{ ok: true } | { ok: false; error: string }> {
  const reg = db.getRegistration(bocirId);
  if (!reg) throw new Error(`Inscription ${bocirId} introuvable.`);
  if (reg.mail_status !== "en_attente") {
    throw new Error(`Inscription ${bocirId} n'est pas en attente (statut actuel: ${reg.mail_status}).`);
  }

  const email = buildConfirmationEmail(config.smtp.mailFrom, reg.prenom);

  try {
    await sendConfirmationMail(config, reg.email, email.subject, email.html);
    db.markSent(bocirId);
    logger.info(`Mail de confirmation envoye a ${reg.email} (inscription ${bocirId}).`);
    return { ok: true };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    db.markSendFailed(bocirId);
    logger.error(`Echec envoi mail pour ${bocirId} : ${message}`);
    return { ok: false, error: message };
  }
}

export async function sendAllPending(
  config: AppConfig
): Promise<{ sent: string[]; failed: Array<{ bocirId: string; error: string }> }> {
  const pending = db.listRegistrations("en_attente");
  const sent: string[] = [];
  const failed: Array<{ bocirId: string; error: string }> = [];

  for (const reg of pending) {
    const result = await sendRegistrationMail(config, reg.bocir_id);
    if (result.ok) {
      sent.push(reg.bocir_id);
    } else {
      failed.push({ bocirId: reg.bocir_id, error: result.error });
    }
  }

  return { sent, failed };
}

export function buildPreview(reg: db.Registration, mailFrom: string) {
  const email = buildConfirmationEmail(mailFrom, reg.prenom);
  return {
    ...reg,
    mail_preview: {
      to: reg.email,
      subject: email.subject,
      html: email.html,
    },
  };
}
