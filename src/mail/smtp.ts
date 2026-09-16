import nodemailer, { Transporter } from "nodemailer";
import { AppConfig } from "../config";

let transporter: Transporter | null = null;

function getTransporter(config: AppConfig): Transporter {
  if (!transporter) {
    transporter = nodemailer.createTransport({
      host: config.smtp.host,
      port: config.smtp.port,
      secure: true, // SSL/TLS direct, port 465
      auth: {
        user: config.smtp.user,
        pass: config.smtp.password,
      },
    });
  }
  return transporter;
}

export async function sendConfirmationMail(
  config: AppConfig,
  to: string,
  subject: string,
  html: string
): Promise<void> {
  try {
    await getTransporter(config).sendMail({
      from: config.smtp.mailFrom,
      to,
      subject,
      html,
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    throw new Error(`Echec envoi SMTP vers ${to} : ${message}`);
  }
}
