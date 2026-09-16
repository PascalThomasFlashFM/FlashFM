import express from "express";
import * as path from "path";
import { AppConfig } from "../config";
import { logger } from "../logger";
import { createApiRouter } from "./routes";

export function startWebServer(config: AppConfig): void {
  const app = express();
  app.use(express.json());
  app.use("/api", createApiRouter(config));
  app.use(express.static(path.join(__dirname, "public")));

  app.listen(config.app.webPort, () => {
    logger.info(`Interface de validation disponible sur http://localhost:${config.app.webPort}`);
  });
}
