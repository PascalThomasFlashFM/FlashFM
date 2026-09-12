import { sendAllPending, sendRegistrationMail } from "./actions";
import { loadConfig } from "./config";
import * as db from "./db";
import { initLogger, logger } from "./logger";
import { runSync } from "./sync/runSync";
import { startWebServer } from "./web/server";

function printUsage(): void {
  console.log(`Usage :
  node dist/cli.js sync            Synchronise BOCIR -> Google Sheet (jamais d'envoi de mail)
  node dist/cli.js list            Liste les inscriptions en attente de validation
  node dist/cli.js send <id>       Envoie le mail de confirmation pour une inscription
  node dist/cli.js send-all        Envoie tous les mails en attente
  node dist/cli.js serve           Demarre l'interface web de validation
`);
}

const KNOWN_COMMANDS = ["sync", "list", "send", "send-all", "serve"];

async function main(): Promise<void> {
  const [, , command, arg] = process.argv;

  if (!command || !KNOWN_COMMANDS.includes(command)) {
    printUsage();
    process.exit(command ? 1 : 0);
  }

  const config = loadConfig();
  initLogger(config.app.logPath);
  db.openDb(config.app.dbPath);

  switch (command) {
    case "sync": {
      const summary = await runSync(config);
      console.log(summary.message);
      process.exit(summary.status === "error" ? 1 : 0);
      break;
    }
    case "list": {
      const pending = db.listRegistrations("en_attente");
      if (pending.length === 0) {
        console.log("Aucune inscription en attente.");
      } else {
        for (const r of pending) {
          console.log(`${r.bocir_id}\t${r.prenom} ${r.nom}\t${r.email}`);
        }
      }
      break;
    }
    case "send": {
      if (!arg) {
        console.error("Usage : node dist/cli.js send <id>");
        process.exit(1);
      }
      const result = await sendRegistrationMail(config, arg);
      if (result.ok) {
        console.log(`Mail envoye pour ${arg}.`);
      } else {
        console.error(`Echec : ${result.error}`);
        process.exit(1);
      }
      break;
    }
    case "send-all": {
      const result = await sendAllPending(config);
      console.log(`${result.sent.length} mail(s) envoye(s), ${result.failed.length} echec(s).`);
      if (result.failed.length > 0) {
        for (const f of result.failed) {
          console.error(`  ${f.bocirId} : ${f.error}`);
        }
        process.exit(1);
      }
      break;
    }
    case "serve": {
      startWebServer(config);
      break;
    }
    default: {
      printUsage();
      process.exit(command ? 1 : 0);
    }
  }
}

main().catch((err) => {
  logger.error(err instanceof Error ? err.stack ?? err.message : String(err));
  process.exit(1);
});
