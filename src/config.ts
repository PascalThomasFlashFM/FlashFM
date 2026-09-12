import * as fs from "fs";
import * as path from "path";
import * as dotenv from "dotenv";

dotenv.config();

function required(name: string): string {
  const value = process.env[name];
  if (!value || value.trim() === "") {
    throw new Error(
      `Variable d'environnement manquante : ${name}. Copie .env.example vers .env et renseigne-la.`
    );
  }
  return value;
}

function optional(name: string, fallback: string): string {
  const value = process.env[name];
  return value && value.trim() !== "" ? value : fallback;
}

export interface AppConfig {
  bocir: {
    apiKey: string;
    bearerToken: string;
    formId: string;
    endpoint: string;
  };
  google: {
    sheetId: string;
    sheetTab: string;
    serviceAccountJsonPath: string;
  };
  smtp: {
    host: string;
    port: number;
    user: string;
    password: string;
    mailFrom: string;
  };
  app: {
    dbPath: string;
    logPath: string;
    webPort: number;
  };
}

let cached: AppConfig | null = null;

export function loadConfig(): AppConfig {
  if (cached) return cached;

  const serviceAccountJsonPath = required("GOOGLE_SERVICE_ACCOUNT_JSON");
  const resolvedServiceAccountPath = path.resolve(serviceAccountJsonPath);
  if (!fs.existsSync(resolvedServiceAccountPath)) {
    throw new Error(
      `Fichier de credentials Google introuvable : ${resolvedServiceAccountPath}. ` +
        `Verifie la variable GOOGLE_SERVICE_ACCOUNT_JSON.`
    );
  }

  cached = {
    bocir: {
      apiKey: required("BOCIR_API_KEY"),
      bearerToken: required("BOCIR_BEARER_TOKEN"),
      formId: required("BOCIR_FORM_ID"),
      endpoint: "https://api.bocir.fr/graphql",
    },
    google: {
      sheetId: required("GOOGLE_SHEET_ID"),
      sheetTab: required("GOOGLE_SHEET_TAB"),
      serviceAccountJsonPath: resolvedServiceAccountPath,
    },
    smtp: {
      host: optional("SMTP_HOST", "ssl0.ovh.net"),
      port: parseInt(optional("SMTP_PORT", "465"), 10),
      user: required("SMTP_USER"),
      password: required("SMTP_PASSWORD"),
      mailFrom: optional("MAIL_FROM", "pascal@flashfm.fr"),
    },
    app: {
      dbPath: optional("DB_PATH", "./data/voyance.sqlite"),
      logPath: optional("LOG_PATH", "./data/app.log"),
      webPort: parseInt(optional("WEB_PORT", "3000"), 10),
    },
  };

  return cached;
}
